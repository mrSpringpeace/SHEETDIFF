"""HTML report: jeden soubor bez externích zdrojů, filtr a hledání v JavaScriptu."""
from __future__ import annotations

import difflib
import html
import json
import os
import re
from datetime import datetime

from . import __version__
from .compare import CATEGORIES, NOISE_CATS, Result, cat_name
from .i18n import _, get_lang

_TOKEN = re.compile(r"\s+|\w+|[^\w\s]", re.UNICODE)
_DIFF_CATS = {"text", "shape_text", "formula", "cache", "struct", "chart", "rules", "link", "comment"}


def _marker(kind: str, text: str) -> str:
    """Značka ‸ na místě, kde druhá verze text vložila (ins) nebo smazala (del)."""
    label = _("vloženo") if kind == "ins" else _("smazáno")
    tip = text if len(text) <= 200 else text[:200] + "…"
    return f'<span class="mk mk-{kind}" title="{label}: {html.escape(tip)}">‸</span>'


def word_diff(a: str, b: str) -> tuple[str, str]:
    """Zvýraznění rozdílu po slovech: vrací (staré, nové) jako HTML.

    Čisté vložení dostane ve starém textu značku ‸ na místě vložení, čisté smazání
    totéž v novém textu - aby bylo vidět, kde se změna stala i na druhé straně.
    """
    ta, tb = _TOKEN.findall(a), _TOKEN.findall(b)
    if len(ta) + len(tb) > 6000:
        return html.escape(a), html.escape(b)
    sm = difflib.SequenceMatcher(None, ta, tb, autojunk=False)
    oa, ob = [], []
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        ra, rb = "".join(ta[i1:i2]), "".join(tb[j1:j2])
        sa, sb = html.escape(ra), html.escape(rb)
        if op == "equal":
            oa.append(sa)
            ob.append(sb)
            continue
        if sa:
            oa.append(f"<del>{sa}</del>")
        elif rb.strip():
            oa.append(_marker("ins", rb))
        if sb:
            ob.append(f"<ins>{sb}</ins>")
        elif ra.strip():
            ob.append(_marker("del", ra))
    return "".join(oa), "".join(ob)


def _fmt_time(ts: float) -> str:
    fmt = "%d. %m. %Y %H:%M" if get_lang() == "cs" else "%Y-%m-%d %H:%M"
    return datetime.fromtimestamp(ts).strftime(fmt)


def _fmt_size(n: int) -> str:
    return f"{n / 1024 / 1024:.2f} MB" if n > 1024 * 1024 else f"{n / 1024:.1f} kB"


def render(res: Result) -> str:
    rows = []
    for c in res.changes:
        if c.cat in _DIFF_CATS and c.old and c.new:
            oh, nh = word_diff(c.old, c.new)
        else:
            oh, nh = html.escape(c.old), html.escape(c.new)
        rows.append({"l": c.layer, "s": c.sheet, "a": c.loc, "k": c.kind, "c": c.cat,
                     "o": oh, "n": nh, "t": c.note, "z": c.cat in NOISE_CATS})
    cat_names = {k: cat_name(k) for k in list(CATEGORIES) + list(NOISE_CATS)}
    wb_label = _("(sešit)")

    # přehled podle listů
    sheets = [s for s in res.sheets]
    by_sheet: dict[str, dict[str, int]] = {}
    for c in res.changes:
        if c.cat in NOISE_CATS:
            continue
        by_sheet.setdefault(c.sheet or wb_label, {}).setdefault(c.cat, 0)
        by_sheet[c.sheet or wb_label][c.cat] += 1
    used_cats = [k for k in CATEGORIES if any(k in d for d in by_sheet.values())]
    order = [wb_label] + sheets
    sheet_rows = []
    for s in order:
        if s in by_sheet:
            d = by_sheet[s]
            cells = "".join(f"<td class=num>{d.get(k, '') or ''}</td>" for k in used_cats)
            sheet_rows.append(
                f"<tr data-sheet=\"{html.escape(s)}\"><th scope=row>{html.escape(s)}</th>"
                f"{cells}<td class=num><b>{sum(d.values())}</b></td></tr>")
    sheet_table = ""
    if sheet_rows:
        head = "".join(f"<th class=num>{html.escape(cat_name(k))}</th>" for k in used_cats)
        sheet_table = (f"<div class=scroll><table class=summary><thead><tr><th>{_('List')}</th>{head}"
                       f"<th class=num>{_('Celkem')}</th></tr></thead><tbody>{''.join(sheet_rows)}"
                       f"</tbody></table></div><p class=hint>{_('Klikněte na list pro filtr.')}</p>")
    unchanged = [s for s in sheets if s not in by_sheet]

    tone = {0: "same", 1: "noise", 2: "diff"}[res.exit_code]
    facts = "".join(f"<li>{html.escape(f)}</li>" for f in res.facts)
    noise = "".join(f"<li>{html.escape(k)}{(' – ' + html.escape(v)) if v not in ('', '1') else ''}"
                    f"</li>" for k, v in res.noise) or f"<li>{_('žádný')}</li>"
    status_name = {k: _(k) for k in ("vysvětleno", "šum", "nepřiřazeno")}
    parts = "".join(
        f"<tr><td class=mono>{html.escape(p.part)}</td>"
        f"<td><span class='badge {('b-' + {'vysvětleno': 'ok', 'šum': 'noise', 'nepřiřazeno': 'warn'}[p.status])}'>"
        f"{status_name[p.status]}</span></td><td>{html.escape(p.detail)}</td></tr>"
        for p in res.parts) or f"<tr><td colspan=3>{_('Žádná část balíčku se neliší.')}</td></tr>"

    data = json.dumps(rows, ensure_ascii=False).replace("</", "<\\/")
    cats_json = json.dumps(cat_names, ensure_ascii=False)
    texts = {k: _(v) for k, v in _TEXTS.items()}
    js_texts = json.dumps({k: _(v) for k, v in _JS_TEXTS.items()}, ensure_ascii=False)
    title = f"{os.path.basename(res.a.path)} ↔ {os.path.basename(res.b.path)}"
    return _TEMPLATE.format(
        title=html.escape(title), version=__version__,
        a_name=html.escape(os.path.basename(res.a.path)), a_path=html.escape(res.a.path),
        a_meta=_("{s} · uloženo {t}", s=_fmt_size(res.a.size), t=_fmt_time(res.a.mtime)),
        b_name=html.escape(os.path.basename(res.b.path)), b_path=html.escape(res.b.path),
        b_meta=_("{s} · uloženo {t}", s=_fmt_size(res.b.size), t=_fmt_time(res.b.mtime)),
        created=_fmt_time(datetime.now().timestamp()),
        tone=tone, headline=html.escape(res.headline), facts=facts,
        sheet_table=sheet_table,
        unchanged=("<p class=hint>" + _("Listy beze změny: ") + html.escape(", ".join(unchanged)) + "</p>")
        if unchanged and res.changes else "",
        noise=noise, parts=parts, data=data, cats=cats_json, L=texts, js_texts=js_texts,
        lang=get_lang(), wb_label=json.dumps(wb_label, ensure_ascii=False),
        tol=f"{res.tolerance:g}", exit_code=res.exit_code,
        n_changes=len([r for r in rows if not r["z"]]),
        has_changes="" if rows else "hidden",
    )


_TEXTS = {
    "sub": "porovnání sešitů", "old": "Starý (A)", "new": "Nový (B)", "exit": "Návratový kód",
    "tol": "tolerance čísel", "rel": "(relativní)", "by_sheet": "Přehled podle listů",
    "changes": "Změny", "layer": "Vrstva", "sheet": "List", "kind": "Druh změny",
    "all_layers": "Všechny vrstvy", "all_sheets": "Všechny listy", "all_kinds": "Všechny druhy",
    "search": "Hledat v adrese, hodnotách, poznámce…", "search_label": "Hledat",
    "show_noise": "zobrazit šum", "place": "Místo", "change": "Změna", "more": "Zobrazit další",
    "noise_h": "Technické rozdíly (šum)", "ignored": "Ignorované rozdíly",
    "pkg_h": "Balíček (ZIP) – části, které se liší",
    "pkg_hint": "Záchranná síť: každá část souboru, která se liší, je tu uvedená. „Nepřiřazeno“ "
                "znamená, že rozdíl žádná vrstva zatím neumí vyložit – je vidět jen tady.",
    "part": "Část", "status": "Stav", "details": "Podrobnosti",
    "footer": "soubory byly jen čteny, nic se do nich nezapisovalo.",
}
_JS_TEXTS = {
    "none": "Žádná změna neodpovídá filtru.",
    "shown": "zobrazeno {a} z {b}",
    "count": "{n} změn",
}


_TEMPLATE = """<!doctype html>
<html lang="{lang}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SHEETDIFF – {title}</title>
<style>
:root {{
  --bg: #f6f7f9; --panel: #ffffff; --text: #1d2330; --muted: #5d6677; --line: #dfe3ea;
  --accent: #2357c6; --del-bg: #fde4e4; --del-text: #8a1c1c; --ins-bg: #ddf3e3; --ins-text: #145c2c;
  --same: #e5f4ea; --same-line: #2f8a4f; --noise: #eef1f6; --noise-line: #7a8597;
  --diff: #fff3dd; --diff-line: #c27a00; --warn: #fde8d2; --warn-text: #8a4a00;
  --row-hover: #f0f4fb;
}}
@media (prefers-color-scheme: dark) {{
  :root:not([data-theme="light"]) {{
    --bg: #14171c; --panel: #1c2027; --text: #e3e7ee; --muted: #9aa3b2; --line: #2e343e;
    --accent: #7ea6ff; --del-bg: #4a2224; --del-text: #ffb8b8; --ins-bg: #1f3d2a; --ins-text: #a9e8bb;
    --same: #1c3325; --same-line: #57c27f; --noise: #232831; --noise-line: #8b95a6;
    --diff: #3a2e17; --diff-line: #e6a63c; --warn: #43301b; --warn-text: #ffcf99;
    --row-hover: #232a35;
  }}
}}
:root[data-theme="dark"] {{
  --bg: #14171c; --panel: #1c2027; --text: #e3e7ee; --muted: #9aa3b2; --line: #2e343e;
  --accent: #7ea6ff; --del-bg: #4a2224; --del-text: #ffb8b8; --ins-bg: #1f3d2a; --ins-text: #a9e8bb;
  --same: #1c3325; --same-line: #57c27f; --noise: #232831; --noise-line: #8b95a6;
  --diff: #3a2e17; --diff-line: #e6a63c; --warn: #43301b; --warn-text: #ffcf99;
  --row-hover: #232a35;
}}
* {{ box-sizing: border-box; }}
body {{ margin: 0; background: var(--bg); color: var(--text);
  font: 14px/1.45 "Segoe UI", system-ui, sans-serif; }}
main {{ max-width: 1400px; margin: 0 auto; padding: 20px 16px 60px; }}
h1 {{ font-size: 20px; margin: 0 0 14px; }}
h1 small {{ font-weight: 400; color: var(--muted); font-size: 13px; margin-left: 8px; }}
h2 {{ font-size: 16px; margin: 28px 0 10px; }}
.files {{ display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }}
.file {{ background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 10px 14px; min-width: 0; }}
.file .lbl {{ font-size: 12px; text-transform: uppercase; letter-spacing: .04em; color: var(--muted); }}
.file .name {{ font-weight: 600; overflow-wrap: anywhere; }}
.file .path, .file .meta {{ color: var(--muted); font-size: 12px; overflow-wrap: anywhere; }}
.verdict {{ margin-top: 14px; border-radius: 8px; padding: 14px 18px; border-left: 5px solid; }}
.verdict.same {{ background: var(--same); border-color: var(--same-line); }}
.verdict.noise {{ background: var(--noise); border-color: var(--noise-line); }}
.verdict.diff {{ background: var(--diff); border-color: var(--diff-line); }}
.verdict .head {{ font-size: 17px; font-weight: 600; }}
.verdict ul {{ margin: 6px 0 0; padding-left: 20px; }}
.verdict .code {{ color: var(--muted); font-size: 12px; margin-top: 6px; }}
.scroll {{ overflow-x: auto; }}
table {{ border-collapse: collapse; width: 100%; background: var(--panel); }}
th, td {{ border-bottom: 1px solid var(--line); padding: 6px 8px; text-align: left; vertical-align: top; }}
thead th {{ position: sticky; top: 0; background: var(--panel); font-size: 12px; color: var(--muted);
  font-weight: 600; z-index: 1; }}
.num {{ text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }}
.summary tbody tr {{ cursor: pointer; }}
.summary tbody tr:hover, #changes tbody tr:hover {{ background: var(--row-hover); }}
.hint {{ color: var(--muted); font-size: 12px; margin: 6px 0 0; }}
.filters {{ display: flex; flex-wrap: wrap; gap: 8px; align-items: center; margin-bottom: 10px; }}
.filters select, .filters input[type=search] {{ font: inherit; padding: 5px 8px; border: 1px solid var(--line);
  border-radius: 6px; background: var(--panel); color: var(--text); }}
.filters input[type=search] {{ flex: 1 1 220px; min-width: 0; }}
.filters label {{ color: var(--muted); font-size: 13px; }}
.filters .count {{ color: var(--muted); font-size: 13px; margin-left: auto; }}
#changes td.v {{ white-space: pre-wrap; overflow-wrap: anywhere; max-width: 460px;
  font-family: Consolas, "Cascadia Mono", monospace; font-size: 12.5px; }}
#changes td.old {{ background: color-mix(in srgb, var(--del-bg) 35%, transparent); }}
#changes td.new {{ background: color-mix(in srgb, var(--ins-bg) 35%, transparent); }}
#changes td.loc {{ white-space: nowrap; font-family: Consolas, monospace; font-size: 12.5px; }}
#changes .note {{ display: block; color: var(--muted); font-size: 12px; margin-top: 2px; }}
#changes tr.noise {{ opacity: .6; }}
del {{ background: var(--del-bg); color: var(--del-text); text-decoration: line-through; border-radius: 2px; }}
ins {{ background: var(--ins-bg); color: var(--ins-text); text-decoration: none; border-radius: 2px; }}
.mk {{ font-weight: 700; padding: 0 2px; margin: 0 1px; border-radius: 3px; cursor: help;
  font-family: "Segoe UI", system-ui, sans-serif; }}
.mk-ins {{ background: var(--ins-bg); color: var(--ins-text); }}
.mk-del {{ background: var(--del-bg); color: var(--del-text); }}
.tag {{ display: inline-block; font-size: 11px; padding: 1px 6px; border-radius: 10px;
  border: 1px solid var(--line); color: var(--muted); white-space: nowrap; }}
.badge {{ font-size: 12px; padding: 1px 7px; border-radius: 10px; white-space: nowrap; }}
.b-ok {{ background: var(--same); }}
.b-noise {{ background: var(--noise); color: var(--muted); }}
.b-warn {{ background: var(--warn); color: var(--warn-text); font-weight: 600; }}
.mono {{ font-family: Consolas, monospace; font-size: 12.5px; overflow-wrap: anywhere; }}
button.more {{ margin-top: 10px; font: inherit; padding: 6px 14px; border-radius: 6px;
  border: 1px solid var(--line); background: var(--panel); color: var(--accent); cursor: pointer; }}
details {{ background: var(--panel); border: 1px solid var(--line); border-radius: 8px; padding: 8px 14px; margin-top: 10px; }}
summary {{ cursor: pointer; font-weight: 600; }}
footer {{ color: var(--muted); font-size: 12px; margin-top: 30px; }}
[hidden] {{ display: none !important; }}
@media (max-width: 700px) {{ .files {{ grid-template-columns: 1fr; }} }}
</style>
</head>
<body>
<main>
<h1>SHEETDIFF <small>{L[sub]} · {created}</small></h1>
<div class="files">
  <div class="file"><div class="lbl">{L[old]}</div><div class="name">{a_name}</div>
    <div class="path">{a_path}</div><div class="meta">{a_meta}</div></div>
  <div class="file"><div class="lbl">{L[new]}</div><div class="name">{b_name}</div>
    <div class="path">{b_path}</div><div class="meta">{b_meta}</div></div>
</div>
<section class="verdict {tone}">
  <div class="head">{headline}</div>
  <ul>{facts}</ul>
  <div class="code">{L[exit]} {exit_code} · {L[tol]} {tol} {L[rel]}</div>
</section>

<section {has_changes}>
<h2>{L[by_sheet]}</h2>
{sheet_table}
{unchanged}

<h2>{L[changes]} <span class="tag" id="total">{n_changes}</span></h2>
<div class="filters">
  <select id="f-layer" aria-label="{L[layer]}"><option value="">{L[all_layers]}</option></select>
  <select id="f-sheet" aria-label="{L[sheet]}"><option value="">{L[all_sheets]}</option></select>
  <select id="f-cat" aria-label="{L[kind]}"><option value="">{L[all_kinds]}</option></select>
  <input type="search" id="f-q" placeholder="{L[search]}" aria-label="{L[search_label]}">
  <label><input type="checkbox" id="f-noise"> {L[show_noise]}</label>
  <span class="count" id="count"></span>
</div>
<div class="scroll">
<table id="changes">
  <thead><tr><th>{L[layer]}</th><th>{L[sheet]}</th><th>{L[place]}</th><th>{L[change]}</th><th>{L[old]}</th><th>{L[new]}</th></tr></thead>
  <tbody></tbody>
</table>
</div>
<button class="more" id="more" hidden>{L[more]}</button>
</section>

<h2>{L[noise_h]}</h2>
<details><summary>{L[ignored]}</summary><ul>{noise}</ul></details>

<h2>{L[pkg_h]}</h2>
<p class="hint">{L[pkg_hint]}</p>
<div class="scroll"><table>
  <thead><tr><th>{L[part]}</th><th>{L[status]}</th><th>{L[details]}</th></tr></thead>
  <tbody>{parts}</tbody>
</table></div>

<footer>SHEETDIFF {version} · {L[footer]}</footer>
</main>
<script>
const ROWS = {data};
const CATS = {cats};
const T = {js_texts};
const WB = {wb_label};
const PAGE = 500;
let shown = PAGE;
const $ = id => document.getElementById(id);
const tbody = document.querySelector('#changes tbody');

function fill(sel, values, label) {{
  for (const v of values) {{
    const o = document.createElement('option');
    o.value = v; o.textContent = label ? label(v) : v;
    sel.appendChild(o);
  }}
}}
const uniq = key => [...new Set(ROWS.map(r => r[key]))];
fill($('f-layer'), uniq('l'));
fill($('f-sheet'), uniq('s').filter(Boolean));
fill($('f-cat'), uniq('c'), v => CATS[v] || v);

const strip = s => s.replace(/<[^>]+>/g, '').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
for (const r of ROWS) r.q = (r.s + ' ' + r.a + ' ' + r.k + ' ' + strip(r.o) + ' ' + strip(r.n) + ' ' + r.t).toLowerCase();

function esc(s) {{ return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }}

function render() {{
  const layer = $('f-layer').value, sheet = $('f-sheet').value, cat = $('f-cat').value;
  const q = $('f-q').value.trim().toLowerCase(), noise = $('f-noise').checked;
  const rows = ROWS.filter(r => (noise || !r.z) && (!layer || r.l === layer) && (!sheet || r.s === sheet)
    && (!cat || r.c === cat) && (!q || r.q.includes(q)));
  const html = [];
  for (const r of rows.slice(0, shown)) {{
    html.push(`<tr class="${{r.z ? 'noise' : ''}}"><td><span class="tag">${{esc(r.l)}}</span></td>`
      + `<td>${{esc(r.s)}}</td><td class="loc">${{esc(r.a)}}</td>`
      + `<td>${{esc(r.k)}}${{r.t ? `<span class="note">${{esc(r.t)}}</span>` : ''}}</td>`
      + `<td class="v old">${{r.o}}</td><td class="v new">${{r.n}}</td></tr>`);
  }}
  tbody.innerHTML = html.join('') || `<tr><td colspan="6">${{T.none}}</td></tr>`;
  $('count').textContent = rows.length > shown ? T.shown.replace('{{a}}', shown).replace('{{b}}', rows.length) : T.count.replace('{{n}}', rows.length);
  $('more').hidden = rows.length <= shown;
}}
for (const id of ['f-layer', 'f-sheet', 'f-cat', 'f-noise']) $(id).addEventListener('change', () => {{ shown = PAGE; render(); }});
$('f-q').addEventListener('input', () => {{ shown = PAGE; render(); }});
$('more').addEventListener('click', () => {{ shown += PAGE; render(); }});
document.querySelectorAll('.summary tbody tr').forEach(tr => tr.addEventListener('click', () => {{
  const s = tr.dataset.sheet;
  $('f-sheet').value = s === WB ? '' : s;
  shown = PAGE; render();
  $('changes').scrollIntoView({{behavior: 'smooth'}});
}}));
render();
</script>
</body>
</html>
"""
