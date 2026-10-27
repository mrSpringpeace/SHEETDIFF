"""Porovnání dvou sešitů po vrstvách: struktura, buňky, formát, tvary, komentáře, balíček."""
from __future__ import annotations

import dataclasses
import os
import posixpath
from collections import Counter
from dataclasses import dataclass, field

from lxml import etree

from . import ooxml
from .align import SheetMap, adjust_refs, sheet_map
from .i18n import _
from .ooxml import Package, SheetInfo, Workbook, addr_key, col_letter, local

LAYERS = ("struktura", "bunky", "format", "pravidla", "tvary", "komentare", "balicek")

# Kategorie změn (popisky se překládají při použití přes cat_name).
# Šumové kategorie se nepočítají jako obsahový rozdíl.
CATEGORIES = {
    "struct": "Struktura (listy, řádky, sloupce)",
    "text": "Texty buněk",
    "number": "Čísla a konstanty",
    "type": "Typ hodnoty",
    "formula": "Vzorce",
    "cache": "Vypočtené hodnoty",
    "format": "Formát buněk",
    "rules": "Podmíněné formátování a ověření dat",
    "link": "Hypertextové odkazy",
    "comment": "Komentáře",
    "shape_text": "Texty tvarů a textových polí",
    "shape": "Tvary (poloha, formát)",
    "image": "Obrázky",
    "chart": "Grafy",
    "vba": "Makra VBA",
    "other": "Nepřiřazené změny",
}
NOISE_CATS = {"chart_cache"}
_CONTENT = {"text", "number", "type", "formula", "cache"}


def cat_name(cat: str) -> str:
    if cat in NOISE_CATS:
        return _("Data grafu (šum)")
    return _(CATEGORIES.get(cat, cat))


@dataclass
class Change:
    layer: str
    sheet: str
    loc: str
    kind: str
    cat: str
    old: str = ""
    new: str = ""
    note: str = ""
    sort: tuple = ()


@dataclass
class PartDiff:
    part: str
    status: str   # šum | vysvětleno | nepřiřazeno (vnitřní kód, překládá se v reportu)
    detail: str = ""


@dataclass
class FileInfo:
    path: str
    size: int
    mtime: float


@dataclass
class Result:
    a: FileInfo
    b: FileInfo
    changes: list[Change] = field(default_factory=list)
    noise: list[tuple[str, str]] = field(default_factory=list)
    parts: list[PartDiff] = field(default_factory=list)
    sheets: list[str] = field(default_factory=list)
    headline: str = ""
    facts: list[str] = field(default_factory=list)
    exit_code: int = 0
    tolerance: float = 1e-9

    @property
    def counts(self) -> Counter:
        return Counter(c.cat for c in self.changes if c.cat not in NOISE_CATS)


class _Ctx:
    def __init__(self, a: Workbook, b: Workbook, tol: float, layers):
        self.a, self.b, self.tol, self.layers = a, b, tol, set(layers)
        self.changes: list[Change] = []
        self.noise: Counter = Counter()
        self.explained: set[str] = set()      # části balíčku, jejichž změnu ukázala vrstva
        self.sheets_with_cells: set[str] = set()
        self.sheet_order: dict[str, int] = {}
        self.maps: dict[str, SheetMap] = {}   # starý název listu -> posun řádků/sloupců
        self.renames: dict[str, str] = {}     # starý název -> nový název
        self.format_changed = False
        self.rules_changed = False
        n = min(len(a.styles), len(b.styles))
        self.same_styles = {str(i) for i in range(n) if a.styles[i] == b.styles[i]}

    def add(self, layer, sheet, loc, kind, cat, old="", new="", note="", sub=(0, 0)):
        self.changes.append(Change(_(layer), sheet, loc, kind, cat, old, new, note,
                                   (self.sheet_order.get(sheet, -1), *sub)))

    def adjust(self, text: str, sheet_a: str | None) -> str:
        if not self.maps and not self.renames:
            return text
        return adjust_refs(text, sheet_a, self.maps, self.renames)


# ---------------------------------------------------------------- pomůcky

def _ranges(addrs, limit: int = 15) -> str:
    """Seznam adres jako souvislé oblasti: A1:C5, E7 …"""
    rows: dict[int, list[int]] = {}
    for r, c in sorted(addr_key(a) for a in addrs):
        rows.setdefault(r, []).append(c)
    runs = []
    for r in sorted(rows):
        cs = rows[r]
        start = prev = cs[0]
        for c in cs[1:]:
            if c != prev + 1:
                runs.append((r, start, prev))
                start = c
            prev = c
        runs.append((r, start, prev))
    rects, open_ = [], {}
    for r, c1, c2 in runs:
        rect = open_.get((c1, c2))
        if rect and rect[3] == r - 1:
            rect[3] = r
        else:
            rect = [c1, c2, r, r]
            open_[(c1, c2)] = rect
            rects.append(rect)
    out = [f"{col_letter(c1)}{r1}" + ("" if (c1, r1) == (c2, r2) else f":{col_letter(c2)}{r2}")
           for c1, c2, r1, r2 in rects]
    return ", ".join(out[:limit]) + (f" … (+{len(out) - limit})" if len(out) > limit else "")


def _spans(idx) -> list[tuple[int, int]]:
    out: list[tuple[int, int]] = []
    for i in sorted(idx):
        if out and i == out[-1][1] + 1:
            out[-1] = (out[-1][0], i)
        else:
            out.append((i, i))
    return out


def _span_text(spans, fmt, sep=":", limit: int = 20) -> str:
    txt = [fmt(a) if a == b else f"{fmt(a)}{sep}{fmt(b)}" for a, b in spans]
    return ", ".join(txt[:limit]) + (f" … (+{len(txt) - limit})" if len(txt) > limit else "")


def _letter(c: int) -> str:
    return col_letter(c)


# ---------------------------------------------------------------- struktura

_STATE = {"visible": "viditelný", "hidden": "skrytý", "veryHidden": "velmi skrytý"}


def _compare_sheets(ctx: _Ctx) -> list[tuple[SheetInfo, SheetInfo]]:
    a, b = ctx.a, ctx.b
    an = [s.name for s in a.sheets]
    bn = [s.name for s in b.sheets]
    removed = [n for n in an if n not in bn]
    added = [n for n in bn if n not in an]
    pairs = [(a.sheet(n), b.sheet(n)) for n in an if n in bn]

    # přejmenování: odebraný a přidaný list se stejným obsahem, nebo jediný na stejném místě
    for rn in list(removed):
        sa = a.sheet(rn)
        da = a.pkg.parts.get(sa.part or "")
        for nn in added:
            sb = b.sheet(nn)
            if da is not None and da == b.pkg.parts.get(sb.part or ""):
                ctx.add("Struktura", nn, nn, _("List přejmenován"), "struct", rn, nn,
                        _("obsah listu beze změny"))
                break
        else:
            if len(removed) == 1 and len(added) == 1 and \
                    a.sheet(rn).index == b.sheet(added[0]).index:
                nn = added[0]
                ctx.add("Struktura", nn, nn, _("List pravděpodobně přejmenován"), "struct", rn, nn,
                        _("stejné pořadí listu; obsah porovnán níže"))
            else:
                continue
        removed.remove(rn)
        added.remove(nn)
        pairs.append((sa, b.sheet(nn)))
        ctx.renames[rn] = nn
    for n in removed:
        ctx.add("Struktura", n, n, _("List odebrán"), "struct", n, "")
    for n in added:
        ctx.add("Struktura", n, n, _("List přidán"), "struct", "", n)
        sb = b.sheet(n)
        if sb.part:
            ctx.explained.add(sb.part)

    common_a = [ctx.renames.get(n, n) for n in an if n in bn or n in ctx.renames]
    common_b = [n for n in bn if n in an or n in ctx.renames.values()]
    if common_a != common_b:
        ctx.add("Struktura", "", _("pořadí listů"), _("Pořadí listů změněno"), "struct",
                ", ".join(common_a), ", ".join(common_b))
    for sa, sb in pairs:
        if sa.state != sb.state:
            ctx.add("Struktura", sb.name, sb.name, _("Viditelnost listu změněna"), "struct",
                    _(_STATE.get(sa.state, sa.state)), _(_STATE.get(sb.state, sb.state)))
    return pairs


def _compare_names(ctx: _Ctx):
    a, b = ctx.a, ctx.b
    names_a = {}
    for key, val in a.names.items():
        name, _sep, scope = key.partition(" [")
        if scope:
            scope = scope[:-1]
            key = f"{name} [{ctx.renames.get(scope, scope)}]"
        names_a[key] = (val, ctx.adjust(val, scope or None))
    for key in sorted(set(names_a) | set(b.names)):
        va, va_adj = names_a.get(key, (None, None))
        vb = b.names.get(key)
        if va_adj == vb:
            continue
        kind = (_("Pojmenovaná oblast přidána") if va is None else
                _("Pojmenovaná oblast odebrána") if vb is None else
                _("Pojmenovaná oblast změněna"))
        ctx.add("Struktura", "", key, kind, "struct", va or "", vb or "")

    ea, eb = Counter(a.external), Counter(b.external)
    for x in sorted((ea - eb).elements()):
        ctx.add("Struktura", "", _("externí odkaz"), _("Externí odkaz odebrán"), "struct", x, "")
    for x in sorted((eb - ea).elements()):
        ctx.add("Struktura", "", _("externí odkaz"), _("Externí odkaz přidán"), "struct", "", x)
    if ea != eb:
        for r in a.pkg.rels(a.part) + b.pkg.rels(b.part):
            if r.kind == "externalLink":
                ctx.explained.add(r.target)


_SHIFT_KINDS = {
    # (osa, vloženo?, jeden?) -> text
    ("r", True, True): "Vložen řádek", ("r", True, False): "Vloženy řádky ({n})",
    ("r", False, True): "Smazán řádek", ("r", False, False): "Smazány řádky ({n})",
    ("c", True, True): "Vložen sloupec", ("c", True, False): "Vloženy sloupce ({n})",
    ("c", False, True): "Smazán sloupec", ("c", False, False): "Smazány sloupce ({n})",
}


def _report_shifts(ctx: _Ctx, sheet: str, m: SheetMap, ca, cb):
    def preview(cells: ooxml.SheetCells, keep) -> tuple[str, int]:
        hits = sorted((a for a, c in cells.cells.items()
                       if (c.kind is not None or c.formula is not None) and keep(*addr_key(a))),
                      key=addr_key)
        lines = []
        for a in hits[:6]:
            f = cells.formula_text(a)
            lines.append(f"{a}: " + (f"={f}" if f else _show(cells.cells[a])))
        if len(hits) > 6:
            lines.append(_("… a dalších {n}", n=len(hits) - 6))
        return "\n".join(lines), len(hits)

    for axis, amap in (("r", m.rows), ("c", m.cols)):
        if amap is None:
            continue
        for idx, cells, inserted in ((amap.inserted, cb, True), (amap.deleted, ca, False)):
            for lo, hi in _spans(idx):
                n = hi - lo + 1
                kind = _(_SHIFT_KINDS[(axis, inserted, n == 1)], n=n)
                if axis == "r":
                    loc = str(lo) if n == 1 else f"{lo}:{hi}"
                    text, cnt = preview(cells, lambda r, c: lo <= r <= hi)
                    sub = (lo, 0)
                else:
                    loc = _letter(lo) if n == 1 else f"{_letter(lo)}:{_letter(hi)}"
                    text, cnt = preview(cells, lambda r, c: lo <= c <= hi)
                    sub = (0, lo)
                note = _("{n} buněk s obsahem", n=cnt) if cnt else _("prázdné")
                if inserted:
                    ctx.add("Struktura", sheet, loc, kind, "struct", "", text, note, sub)
                else:
                    ctx.add("Struktura", sheet, loc, kind, "struct", text, "",
                            note + _(" (adresa ve starém souboru)"), sub)


# ---------------------------------------------------------------- buňky

_BOOL = {"0": "NEPRAVDA", "1": "PRAVDA"}
_KIND_NAME = {"s": "text", "n": "číslo", "b": "logická hodnota", "e": "chyba", "d": "datum"}


def _show(c: ooxml.Cell | None) -> str:
    if c is None or c.kind is None:
        return ""
    if c.kind == "b":
        return _(_BOOL.get(c.value or "", c.value or ""))
    return c.value or ""


def _value_class(kind: str | None) -> str:
    return {"s": "text", "n": "number", "d": "number", "b": "number", "e": "number"}.get(kind or "", "")


def _same_value(x: ooxml.Cell | None, y: ooxml.Cell | None, ctx: _Ctx) -> bool:
    kx = x.kind if x else None
    ky = y.kind if y else None
    vx = x.value if x else None
    vy = y.value if y else None
    if kx != ky:
        return False
    if vx == vy:
        return True
    if kx == "n":
        try:
            fx, fy = float(vx), float(vy)
        except (TypeError, ValueError):
            return False
        if abs(fx - fy) <= ctx.tol * max(abs(fx), abs(fy)):
            ctx.noise[_("Číselné rozdíly pod tolerancí")] += 1
            return True
    return False


def _compare_cells(ctx: _Ctx, sa: SheetInfo, sb: SheetInfo):
    ca, cb = ctx.a.cells(sa), ctx.b.cells(sb)
    sheet = sb.name
    m = ctx.maps.get(sa.name)
    do_values = "bunky" in ctx.layers
    do_format = "format" in ctx.layers
    n_before = len(ctx.changes)
    if m:
        _report_shifts(ctx, sheet, m, ca, cb)
        a_at = {}
        for a_addr in ca.cells:
            b_addr = m.a2b(a_addr)
            if b_addr:
                a_at[b_addr] = a_addr
        b_cells = {k: v for k, v in cb.cells.items() if m.kept_b(k)}
        keys = a_at.keys() | b_cells.keys()
        same_masters: set = set()
    else:
        a_at = None
        b_cells = cb.cells
        keys = ca.cells.keys() | b_cells.keys()
        same_masters = {si for si, mm in ca.masters.items() if cb.masters.get(si) == mm}
    adjust = bool(ctx.maps or ctx.renames)
    fmt: dict[tuple, list[str]] = {}

    for addr in keys:
        a_addr = a_at.get(addr) if a_at is not None else addr
        x = ca.cells.get(a_addr) if a_addr else None
        y = b_cells.get(addr)
        if do_format:
            sx, sy = (x.style if x else "0"), (y.style if y else "0")
            if sx != sy or sx not in ctx.same_styles:
                dx, dy = ctx.a.style(sx), ctx.b.style(sy)
                if dx != dy:
                    fmt.setdefault((dx, dy), []).append(addr)
        if not do_values:
            continue
        if x == y and (x.formula is None or (m is None and (
                x.formula[0] != "s" or x.formula[1] in same_masters) and not ctx.renames)):
            continue
        empty_x = x is None or (x.kind is None and x.formula is None)
        empty_y = y is None or (y.kind is None and y.formula is None)
        if empty_x and empty_y:
            continue
        sub = addr_key(addr)
        fa = ca.formula_text(a_addr) if not empty_x else None
        fb = cb.formula_text(addr) if not empty_y else None
        va, vb = _show(x), _show(y)
        loc = addr if a_addr == addr or a_addr is None else _("{b} (dříve {a})", b=addr, a=a_addr)
        if empty_x:
            cat = "formula" if fb else _value_class(y.kind)
            ctx.add("Buňky", sheet, loc, _("Buňka vyplněna"), cat, "",
                    f"={fb}" if fb else vb, _("hodnota: {v}", v=vb) if fb else "", sub)
        elif empty_y:
            cat = "formula" if fa else _value_class(x.kind)
            ctx.add("Buňky", sheet, loc, _("Buňka vymazána"), cat,
                    f"={fa}" if fa else va, "", _("hodnota: {v}", v=va) if fa else "", sub)
        elif fa or fb:
            fa_cmp = ctx.adjust(fa, sa.name) if (fa and adjust) else fa
            if fa_cmp != fb:
                kind = (_("Vzorec změněn") if fa and fb else
                        _("Konstanta nahrazena vzorcem") if fb else _("Vzorec nahrazen konstantou"))
                note = _("hodnota: {a} → {b}", a=va, b=vb) if va != vb else \
                    _("hodnota beze změny: {v}", v=va) if va else ""
                ctx.add("Buňky", sheet, loc, kind, "formula",
                        f"={fa}" if fa else va, f"={fb}" if fb else vb, note, sub)
            elif not _same_value(x, y, ctx):
                ctx.add("Buňky", sheet, loc, _("Vypočtená hodnota změněna"), "cache",
                        va, vb, _("vzorec beze změny: ={f}", f=fb), sub)
        elif not _same_value(x, y, ctx):
            if x.kind != y.kind:
                ctx.add("Buňky", sheet, loc, _("Typ hodnoty změněn"), "type", va, vb,
                        f"{_(_KIND_NAME.get(x.kind, x.kind))} → {_(_KIND_NAME.get(y.kind, y.kind))}", sub)
            elif _value_class(x.kind) == "text":
                ctx.add("Buňky", sheet, loc, _("Text změněn"), "text", va, vb, "", sub)
            else:
                ctx.add("Buňky", sheet, loc, _("Hodnota změněna"), "number", va, vb, "", sub)

    if do_format:
        for (dx, dy), addrs in sorted(fmt.items(), key=lambda kv: min(map(addr_key, kv[1]))):
            da, db = dict(dx), dict(dy)
            keys_ = [k for k in dict.fromkeys(list(da) + list(db)) if da.get(k) != db.get(k)]
            ctx.add("Formát", sheet, _ranges(addrs), _("Formát změněn"), "format",
                    "\n".join(f"{k}: {da.get(k, '?')}" for k in keys_),
                    "\n".join(f"{k}: {db.get(k, '?')}" for k in keys_),
                    _("{n} buněk", n=len(addrs)) if len(addrs) > 1 else "",
                    min(map(addr_key, addrs)))
        _compare_layout(ctx, sheet, m, ca, cb)
    if len(ctx.changes) > n_before:
        for p in (sa.part, sb.part):
            if p:
                ctx.sheets_with_cells.add(p)
        if any(c.cat == "format" for c in ctx.changes[n_before:]):
            ctx.format_changed = True


def _describe_dim(v, size_label) -> dict:
    size, hidden, level = v
    return {size_label: size if size is not None else _("výchozí"),
            _("skrytí"): _("skrytý") if hidden else _("viditelný"),
            _("seskupení"): _("úroveň {n}", n=level)}


_DIM_KINDS = {
    ("r", "hide"): "Řádky skryty", ("r", "show"): "Řádky zobrazeny", ("r", "other"): "Řádky: změna ({what})",
    ("c", "hide"): "Sloupce skryty", ("c", "show"): "Sloupce zobrazeny", ("c", "other"): "Sloupce: změna ({what})",
}


def _compare_layout(ctx: _Ctx, sheet: str, m: SheetMap | None, ca, cb):
    # sloučené buňky
    def map_range(ref: str) -> str | None:
        if m is None:
            return ref
        parts = [m.a2b(p) for p in ref.split(":")]
        return None if None in parts else ":".join(parts)

    ma = {map_range(r) for r in ca.merges} - {None}
    mb = set(cb.merges)
    for ref in sorted(ma - mb, key=lambda r: addr_key(r.split(":")[0])):
        ctx.add("Formát", sheet, ref, _("Sloučení buněk zrušeno"), "format", _("sloučeno"), "",
                "", addr_key(ref.split(":")[0]))
    for ref in sorted(mb - ma, key=lambda r: addr_key(r.split(":")[0])):
        ctx.add("Formát", sheet, ref, _("Buňky sloučeny"), "format", "", _("sloučeno"),
                "", addr_key(ref.split(":")[0]))

    # výšky a skrytí řádků, šířky a skrytí sloupců
    default = (None, False, "0")
    for axis, da, db, size_label, fmt in (
            ("r", ca.rows, cb.rows, _("výška"), str),
            ("c", ca.cols, cb.cols, _("šířka"), _letter)):
        amap = (m.rows if axis == "r" else m.cols) if m else None
        if amap:
            da = {amap.map_a(k): v for k, v in da.items() if amap.kept_a(k)}
            db = {k: v for k, v in db.items() if amap.kept_b(k)}
        groups: dict[tuple, list[int]] = {}
        for k in da.keys() | db.keys():
            va, vb = da.get(k, default), db.get(k, default)
            if axis == "c" and va[0] is not None and vb[0] is not None:
                va = (round(va[0], 1), *va[1:])
                vb = (round(vb[0], 1), *vb[1:])
            if va != vb:
                groups.setdefault((va, vb), []).append(k)
        for (va, vb), idx in sorted(groups.items(), key=lambda kv: min(kv[1])):
            xa, xb = _describe_dim(va, size_label), _describe_dim(vb, size_label)
            diff = [k for k in xa if xa[k] != xb[k]]
            if diff == [_("skrytí")]:
                kind = _(_DIM_KINDS[(axis, "hide" if vb[1] else "show")])
            else:
                kind = _(_DIM_KINDS[(axis, "other")], what=", ".join(diff))
            spans = _spans(idx)
            sub = (spans[0][0], 0) if axis == "r" else (0, spans[0][0])
            ctx.add("Formát", sheet, _span_text(spans, fmt), kind, "format",
                    ", ".join(f"{k} {xa[k]}" for k in diff), ", ".join(f"{k} {xb[k]}" for k in diff),
                    "", sub)


# ---------------------------------------------------------------- pravidla a odkazy

_CF_TYPES = {
    "expression": "vzorec", "cellIs": "hodnota buňky", "colorScale": "barevná škála",
    "dataBar": "datový pruh", "iconSet": "sada ikon", "top10": "nejvyšší / nejnižší",
    "uniqueValues": "jedinečné hodnoty", "duplicateValues": "duplicitní hodnoty",
    "containsText": "obsahuje text", "notContainsText": "neobsahuje text",
    "beginsWith": "začíná na", "endsWith": "končí na", "containsBlanks": "prázdné",
    "notContainsBlanks": "neprázdné", "containsErrors": "chyby", "notContainsErrors": "bez chyb",
    "timePeriod": "datum v období", "aboveAverage": "nad / pod průměrem",
}


def _map_ref(ref: str, m: SheetMap | None) -> str:
    """Oblast (i více oblastí oddělených mezerou) posunutá vloženými řádky a sloupci."""
    if m is None:
        return ref
    out = []
    for rng in ref.split():
        pts = []
        for pnt in rng.split(":"):
            if not ooxml._ADDR_RE.match(pnt):
                pts.append(pnt)
                continue
            r, c = addr_key(pnt)
            pts.append(f"{col_letter(m.col(c))}{m.row(r)}")
        out.append(":".join(pts))
    return " ".join(out)


_RULE_KEYS = {
    "operator": "operátor", "allowBlank": "povolit prázdné", "showErrorMessage": "chybové hlášení",
    "showInputMessage": "vstupní zpráva", "errorStyle": "druh chyby", "error": "text chyby",
    "errorTitle": "nadpis chyby", "prompt": "zpráva", "promptTitle": "nadpis zprávy",
    "stopIfTrue": "zastavit, je-li splněno", "text": "text", "rank": "počet", "percent": "procenta",
    "bottom": "nejnižší", "timePeriod": "období", "aboveAverage": "nad průměrem",
}


def _rule_text(sqref: str, desc: tuple, formulas, fmt: str) -> str:
    shown = []
    for x in desc:
        k, eq, v = x.partition("=")
        if not eq:
            shown.append(x)
        elif k == "type":
            shown.insert(0, _("typ {t}", t=_(_CF_TYPES.get(v, v))))
        elif k == "showDropDown":
            if v in ("1", "true"):
                shown.append(_("bez rozbalovacího seznamu"))
        elif v in ("0", "false"):
            continue
        else:
            shown.append(f"{_(_RULE_KEYS.get(k, k))}: {_('ano') if v in ('1', 'true') else v}")
    lines = [_("oblast {r}", r=sqref), ", ".join(shown)]
    lines += [_("vzorec: ={f}", f=f) if not f.startswith('"') else _("hodnota: {v}", v=f)
              for f in formulas if f]
    if fmt:
        lines.append(_("formát: {f}", f=fmt))
    return "\n".join(x for x in lines if x)


def _pair_rules(ra: list[tuple], rb: list[tuple]):
    """Párování pravidel (oblast, popis, vzorce, formát): shodná, stejná oblast, stejný obsah."""
    rest_b = list(rb)
    rest_a = []
    for x in ra:
        if x in rest_b:
            rest_b.remove(x)
        else:
            rest_a.append(x)
    pairs = []
    for key in (lambda r: r[0], lambda r: r[1:]):
        for x in list(rest_a):
            y = next((y for y in rest_b if key(y) == key(x)), None)
            if y is not None:
                pairs.append((x, y))
                rest_a.remove(x)
                rest_b.remove(y)
    return pairs, rest_a, rest_b


_RULE_KINDS = {
    ("cf", "chg"): "Podmíněné formátování: pravidlo změněno",
    ("cf", "del"): "Podmíněné formátování: pravidlo odebráno",
    ("cf", "add"): "Podmíněné formátování: pravidlo přidáno",
    ("dv", "chg"): "Ověření dat: pravidlo změněno",
    ("dv", "del"): "Ověření dat: pravidlo odebráno",
    ("dv", "add"): "Ověření dat: pravidlo přidáno",
}


def _compare_rules(ctx: _Ctx, sa: SheetInfo, sb: SheetInfo):
    ca, cb = ctx.a.cells(sa), ctx.b.cells(sb)
    m = ctx.maps.get(sa.name)
    sheet = sb.name
    n0 = len(ctx.changes)
    for what, la, lb, with_fmt in (("cf", ca.cf, cb.cf, True), ("dv", ca.dv, cb.dv, False)):
        if not la and not lb:
            continue
        ra = [(_map_ref(r.sqref, m), r.desc, tuple(ctx.adjust(f, sa.name) for f in r.formulas),
               ctx.a.dxf(r.dxf) if with_fmt else "") for r in la]
        rb = [(r.sqref, r.desc, r.formulas, ctx.b.dxf(r.dxf) if with_fmt else "") for r in lb]
        pairs, gone, new = _pair_rules(ra, rb)
        for x, y in pairs:
            ctx.add("Pravidla", sheet, y[0], _(_RULE_KINDS[(what, "chg")]), "rules",
                    _rule_text(*x), _rule_text(*y), "", addr_key(y[0].split()[0].split(":")[0]))
        for x in gone:
            ctx.add("Pravidla", sheet, x[0], _(_RULE_KINDS[(what, "del")]), "rules",
                    _rule_text(*x), "", "", addr_key(x[0].split()[0].split(":")[0]))
        for y in new:
            ctx.add("Pravidla", sheet, y[0], _(_RULE_KINDS[(what, "add")]), "rules",
                    "", _rule_text(*y), "", addr_key(y[0].split()[0].split(":")[0]))

    def targets(wb: Workbook, sh: SheetInfo, links) -> dict[str, tuple[str, str]]:
        rels = wb.pkg.rel_map(sh.part) if sh.part else {}
        out = {}
        for ref, rid, loc, disp in links:
            r = rels.get(rid) if rid else None
            tgt = (r.target if r else "") + (f"#{loc}" if loc else "")
            out[ref] = (tgt, disp)
        return out

    ta = targets(ctx.a, sa, ca.links)
    tb = targets(ctx.b, sb, cb.links)
    if ta or tb:
        if m:
            ta = {_map_ref(k, m): (ctx.adjust(t, sa.name) if t.startswith("#") else t, d)
                  for k, (t, d) in ta.items()}
        for ref in sorted(ta.keys() | tb.keys(), key=lambda r: addr_key(r.split(":")[0])):
            x, y = ta.get(ref), tb.get(ref)
            if x == y:
                continue
            sub = addr_key(ref.split(":")[0])
            if x is None:
                ctx.add("Odkazy", sheet, ref, _("Odkaz přidán"), "link", "", y[0], y[1], sub)
            elif y is None:
                ctx.add("Odkazy", sheet, ref, _("Odkaz odebrán"), "link", x[0], "", x[1], sub)
            else:
                ctx.add("Odkazy", sheet, ref, _("Odkaz změněn"), "link", f"{x[0]}  [{x[1]}]".strip(),
                        f"{y[0]}  [{y[1]}]".strip(), "", sub)
    if len(ctx.changes) > n0:
        ctx.rules_changed = True
        for sh in (sa, sb):
            if sh.part:
                ctx.sheets_with_cells.add(sh.part)
                ctx.explained.add(Package.rels_name(sh.part))


# ---------------------------------------------------------------- tvary

def _map_shape(s: ooxml.Shape, m: SheetMap) -> ooxml.Shape:
    g = list(s.geom)
    for ci, ri in ((0, 2), (4, 6)):
        if g[ci] >= 0:
            g[ci] = m.col(g[ci])
        if g[ri] >= 0:
            g[ri] = m.row(g[ri] + 1) - 1
    g = tuple(g)
    return dataclasses.replace(s, geom=g, anchor=ooxml.anchor_addr(g))


def _shape_loc(s: ooxml.Shape) -> str:
    anchor = _(s.anchor) if s.anchor == "absolutně" else s.anchor
    return f"{s.name} ({anchor})"


def _compare_shapes(ctx: _Ctx, sa: SheetInfo, sb: SheetInfo):
    la, lb = ctx.a.shapes(sa), ctx.b.shapes(sb)
    if not la and not lb:
        return
    m = ctx.maps.get(sa.name)
    orig_anchor = {s.key: s.anchor for s in la}   # adresa ve starém souboru (před posunem)
    if m:
        la = [_map_shape(s, m) for s in la]
    sheet = sb.name
    ma = {s.key: s for s in la}
    mb = {s.key: s for s in lb}
    pairs = [(ma[k], mb[k]) for k in ma if k in mb]
    rest_a = [s for s in la if s.key not in mb]
    rest_b = [s for s in lb if s.key not in ma]
    # přejmenované tvary: stejný druh a ukotvení
    for s in list(rest_a):
        t = next((t for t in rest_b if t.kind == s.kind and t.anchor == s.anchor), None)
        if t is not None:
            pairs.append((s, t))
            rest_a.remove(s)
            rest_b.remove(t)

    def note(sh):
        ctx.explained.add(sh.drawing)
        for p in (sh.media, sh.chart):
            if p:
                ctx.explained.add(p)

    for s in rest_a:
        ctx.add("Tvary", sheet, _shape_loc(s), _("Odebráno: {k}", k=_(s.kind).lower()),
                _shape_cat(s), s.text, "", "", addr_key(s.anchor))
        note(s)
    for t in rest_b:
        ctx.add("Tvary", sheet, _shape_loc(t), _("Přidáno: {k}", k=_(t.kind).lower()),
                _shape_cat(t), "", t.text, "", addr_key(t.anchor))
        note(t)
    for s, t in pairs:
        loc = _shape_loc(t)
        sub = addr_key(t.anchor)
        n0 = len(ctx.changes)
        if s.name != t.name:
            ctx.add("Tvary", sheet, loc, _("Přejmenováno"), "shape", s.name, t.name, "", sub)
        if s.text != t.text:
            ctx.add("Tvary", sheet, loc, _("Text změněn"), "shape_text", s.text, t.text,
                    _(t.kind).lower(), sub)
        if s.anchor != t.anchor:
            ctx.add("Tvary", sheet, loc, _("Přesunut na jinou buňku"), "shape",
                    orig_anchor.get(s.key, s.anchor), t.anchor, "", sub)
        elif s.geom != t.geom:
            ctx.add("Tvary", sheet, loc, _("Posunut nebo změněna velikost"), "shape", "", "",
                    _("ukotvení stále v {a}", a=t.anchor), sub)
        if s.media_hash != t.media_hash:
            ctx.add("Tvary", sheet, loc, _("Obrázek vyměněn"), "image", s.media or "", t.media or "",
                    _("obsah obrázku se liší"), sub)
        if s.chart or t.chart:
            _compare_chart(ctx, sheet, loc, s, t, sub)
        if s.fmt != t.fmt:
            ctx.add("Tvary", sheet, loc, _("Formát nebo vlastnosti změněny"), "shape", "", "",
                    _("výplň, čára, písmo, propojení s buňkou apod."), sub)
        if len(ctx.changes) > n0:
            note(s)
            note(t)


def _shape_cat(s) -> str:
    return {"Obrázek": "image", "Graf": "chart"}.get(s.kind, "shape_text" if s.text else "shape")


def _compare_chart(ctx: _Ctx, sheet, loc, s, t, sub):
    pa, pb = ctx.a.pkg, ctx.b.pkg
    da = pa.parts.get(s.chart or "")
    db = pb.parts.get(t.chart or "")
    if da == db and not (ctx.maps or ctx.renames):
        return
    if da is None or db is None:
        ctx.add("Tvary", sheet, loc, _("Graf změněn"), "chart", s.chart or "", t.chart or "", "", sub)
        return
    ia = {k: ctx.adjust(v, None) for k, v in ooxml.chart_info(pa, s.chart).items()}
    ib = ooxml.chart_info(pb, t.chart)
    for k in sorted(set(ia) | set(ib)):
        if ia.get(k) != ib.get(k):
            ctx.add("Tvary", sheet, loc, _("Graf: {k}", k=k), "chart", ia.get(k, ""), ib.get(k, ""),
                    "", sub)
    if ia == ib and da != db:
        if ooxml.chart_without_cache(pa, s.chart) == ooxml.chart_without_cache(pb, t.chart):
            ctx.add("Tvary", sheet, loc, _("Graf: uložená data (cache)"), "chart_cache", "", "",
                    _("graf ukazuje změněná data buněk"), sub)
        elif not ctx.maps:
            ctx.add("Tvary", sheet, loc, _("Graf: formát nebo nastavení"), "chart", "", "",
                    _("řady a osy beze změny"), sub)
        else:
            ctx.noise[_("Graf: odkazy posunuté vložením řádků/sloupců")] += 1
    if da != db:
        ctx.explained.add(s.chart)
        ctx.explained.add(t.chart)
        for r in pa.rels(s.chart) + pb.rels(t.chart):
            ctx.explained.add(r.target)


# ---------------------------------------------------------------- komentáře

def _compare_comments(ctx: _Ctx, sa: SheetInfo, sb: SheetInfo):
    ka, parts_a = ctx.a.comments(sa)
    kb, parts_b = ctx.b.comments(sb)
    if not ka and not kb:
        return
    m = ctx.maps.get(sa.name)
    if m:
        ka = {b: v for b, v in ((m.a2b(k), v) for k, v in ka.items()) if b}
    sheet = sb.name
    n0 = len(ctx.changes)
    for addr in sorted(ka.keys() | kb.keys(), key=addr_key):
        x, y = ka.get(addr), kb.get(addr)
        if x == y:
            continue
        sub = addr_key(addr)
        if x is None:
            ctx.add("Komentáře", sheet, addr, _("Komentář přidán"), "comment", "", y[1],
                    _("autor: {a}", a=y[0]), sub)
        elif y is None:
            ctx.add("Komentáře", sheet, addr, _("Komentář odebrán"), "comment", x[1], "",
                    _("autor: {a}", a=x[0]), sub)
        elif x[1] != y[1]:
            ctx.add("Komentáře", sheet, addr, _("Komentář změněn"), "comment", x[1], y[1],
                    _("autor: {a}", a=x[0]) + (f" → {y[0]}" if x[0] != y[0] else ""), sub)
        else:
            ctx.add("Komentáře", sheet, addr, _("Autor komentáře změněn"), "comment", x[0], y[0], "", sub)
    if len(ctx.changes) > n0:
        ctx.explained.update(parts_a + parts_b)


# ---------------------------------------------------------------- balíček

_SHEET_LABELS = {
    "sheetPr": "vlastnosti listu (barva záložky…)", "cols": "šířky sloupců",
    "mergeCells": "sloučené buňky", "conditionalFormatting": "podmíněné formátování",
    "dataValidations": "ověření dat", "hyperlinks": "hypertextové odkazy",
    "pageMargins": "okraje tisku", "pageSetup": "nastavení tisku", "printOptions": "tisk",
    "headerFooter": "záhlaví a zápatí", "rowBreaks": "konce stránek", "colBreaks": "konce stránek",
    "autoFilter": "automatický filtr", "sheetProtection": "zámek listu",
    "sheetFormatPr": "výchozí výška řádků", "tableParts": "tabulky",
    "extLst": "rozšíření (např. podmíněné formátování, sparkliny)",
    "sheetData": "data listu – jiné rozdíly (výšky řádků, skryté řádky…)",
}
_SHEET_NOISE = {"sheetViews", "dimension", "drawing", "legacyDrawing"}
_WB_COVERED = {"sheets", "definedNames", "externalReferences"}
_WB_NOISE = {"bookViews", "fileVersion", "revisionPtr", "AlternateContent", "calcPr"}
_STYLES_COVERED = {"numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs"}
_STYLES_LABELS = {"dxfs": "styly podmíněného formátu", "colors": "vlastní paleta barev",
                  "tableStyles": "styly tabulek"}
_EXT_RULES = {"{78C0D931-6437-407D-A8EE-F0AAD7539E65}", "{CCE6A557-97BC-4B89-ADB6-D9C93CAAB3DF}"}
_EXT_LABELS = {
    "{05C60535-1F16-4FD2-B633-F4F36F0B64E0}": "sparkliny",
    "{3A4CF648-6AED-40F4-86FF-DC5316D8AED3}": "průřezy (slicery)",
    "{A8765BA9-456A-4DAB-B4F3-ACF838C121DE}": "průřezy (slicery)",
    "{7E03D99C-DC04-49D9-9315-930204A7B6E9}": "časové osy",
    "{EBBEBD26-6C4B-4B03-8E9A-7A5F3C1E4F3B}": "ovládací prvky",
}


def _ext_diff(da: bytes, db: bytes) -> list[str]:
    """URI rozšíření (extLst listu), která se liší."""
    def groups(data):
        root = etree.fromstring(data, ooxml._PARSER)
        ext = ooxml.child(root, "extLst")
        out = {}
        for e in (ext if ext is not None else []):
            if local(e) == "ext":
                out[e.get("uri", "").upper()] = etree.tostring(e, method="c14n")
        return out
    ga, gb = groups(da), groups(db)
    return [u for u in dict.fromkeys(list(ga) + list(gb)) if ga.get(u) != gb.get(u)]


def _child_groups(data: bytes) -> dict[str, list[bytes]]:
    root = etree.fromstring(data, ooxml._PARSER)
    out: dict[str, list[bytes]] = {}
    for c in root:
        n = local(c)
        if n:
            out.setdefault(n, []).append(etree.tostring(c, method="c14n"))
    return out


def _differing_children(da: bytes, db: bytes) -> list[str]:
    ga, gb = _child_groups(da), _child_groups(db)
    return [n for n in dict.fromkeys(list(ga) + list(gb)) if ga.get(n) != gb.get(n)]


def _rel_set(pkg: Package, name: str) -> Counter:
    root = pkg.xml(name)
    return Counter((r.get("Type", "").rsplit("/", 1)[-1], r.get("Target", ""), r.get("TargetMode", ""))
                   for r in root if local(r) == "Relationship")


def _compare_package(ctx: _Ctx, res: Result):
    pa, pb = ctx.a.pkg, ctx.b.pkg
    names = list(dict.fromkeys(list(pa.parts) + list(pb.parts)))
    any_diff = False
    sheet_parts = {s.part for s in ctx.a.sheets + ctx.b.sheets if s.part}
    wb_parts = {ctx.a.part, ctx.b.part}
    shared_parts = {ctx.a.shared_part, ctx.b.shared_part}
    styles_parts = {ctx.a.styles_part, ctx.b.styles_part} - {None}
    covered = set()
    if "format" in ctx.layers:
        covered |= {"cols", "mergeCells"}
    if "pravidla" in ctx.layers:
        covered |= {"conditionalFormatting", "dataValidations", "hyperlinks"}
    view_noise = _("Zobrazení listu (aktivní buňka, posun, lupa)")
    props_noise = _("Vlastnosti dokumentu (čas uložení, autor)")

    def put(part, status, detail=""):
        res.parts.append(PartDiff(part, status, detail))
        if status == "nepřiřazeno":
            ctx.add("Balíček", "", part, _("Nepřiřazená změna"), "other", "", "", detail, (0, 0))
        elif status == "šum":
            ctx.noise[detail or part] += 1

    for n in names:
        da, db = pa.parts.get(n), pb.parts.get(n)
        if da == db:
            continue
        any_diff = True
        if "vba" in n.lower() and n.endswith(".bin"):
            kind = (_("Makra VBA přidána") if da is None else _("Makra VBA odebrána") if db is None
                    else _("Makra VBA změněna"))
            ctx.add("VBA", "", n, kind, "vba", "", "", _("porovnává se jen otisk obsahu (hash)"), (0, 0))
            put(n, "vysvětleno", kind)
            continue
        if da is None or db is None:
            what = _("část přidána") if da is None else _("část odebrána")
            base = n.replace("/_rels/", "/").removesuffix(".rels") if n.endswith(".rels") else n
            if n in ctx.explained or base in ctx.explained:
                put(n, "vysvětleno", what)
            elif n.startswith("docProps/") and n != "docProps/custom.xml":
                put(n, "šum", props_noise)
            else:
                put(n, "nepřiřazeno", what)
            continue
        if n.startswith("docProps/") and n != "docProps/custom.xml":
            put(n, "šum", props_noise)
        elif n.endswith("calcChain.xml"):
            put(n, "šum", _("Pořadí výpočtu (calcChain)"))
        elif n == "[Content_Types].xml":
            put(n, "šum", _("Seznam typů částí"))
        elif n.endswith(".rels"):
            sa, sb = _rel_set(pa, n), _rel_set(pb, n)
            if sa == sb:
                put(n, "šum", _("Přečíslování vazeb mezi částmi"))
            elif n in ctx.explained:
                put(n, "vysvětleno", _("vazby změněných odkazů"))
            else:
                diff = [f"− {t[1]}" for t in (sa - sb).elements()] + \
                       [f"+ {t[1]}" for t in (sb - sa).elements()]
                targets = {posixpath.normpath(posixpath.join(posixpath.dirname(
                    n.replace("/_rels/", "/").removesuffix(".rels")), t[1]))
                           for t in ((sa - sb) + (sb - sa)) if not t[2]}
                if targets and all(t in ctx.explained for t in targets):
                    put(n, "vysvětleno", "; ".join(diff))
                else:
                    put(n, "nepřiřazeno", _("změněné vazby: {d}", d="; ".join(diff)))
        elif n in shared_parts:
            if ctx.sheets_with_cells:
                put(n, "vysvětleno", _("texty porovnány v buňkách"))
            else:
                put(n, "šum", _("Přečíslování sdílených řetězců (texty v buňkách stejné)"))
        elif n in sheet_parts:
            details, unexplained = [], []
            for c in _differing_children(da, db):
                if c in _SHEET_NOISE:
                    ctx.noise[view_noise] += 1
                elif c == "sheetData" and n in ctx.sheets_with_cells:
                    details.append(_("buňky"))
                elif c == "sheetData" and {"bunky", "format"} <= ctx.layers:
                    ctx.noise[_("Technické atributy řádků a buněk (přečíslování stylů…)")] += 1
                elif c in covered:
                    details.append(_(_SHEET_LABELS.get(c, c)))
                elif c == "extLst":
                    for uri in _ext_diff(da, db):
                        if uri in _EXT_RULES and "pravidla" in ctx.layers:
                            details.append(_("rozšířená pravidla (x14)"))
                        else:
                            unexplained.append(_("rozšíření: {x}", x=_(_EXT_LABELS.get(uri, uri))))
                else:
                    unexplained.append(_(_SHEET_LABELS.get(c, c)))
            if unexplained:
                put(n, "nepřiřazeno", ", ".join(unexplained))
            elif details:
                put(n, "vysvětleno", ", ".join(details))
            else:
                put(n, "šum", _("Zobrazení listu nebo technické atributy"))
        elif n in wb_parts:
            unexplained = [c for c in _differing_children(da, db)
                           if c not in _WB_COVERED and c not in _WB_NOISE]
            if unexplained:
                put(n, "nepřiřazeno", _("sešit: {x}", x=", ".join(unexplained)))
            else:
                put(n, "šum", _("Nastavení okna a verze Excelu ve workbook.xml"))
        elif n in styles_parts:
            cov = (_STYLES_COVERED if "format" in ctx.layers else set()) | \
                ({"dxfs"} if "pravidla" in ctx.layers else set())
            unexplained = [_(_STYLES_LABELS.get(c, c)) for c in _differing_children(da, db)
                           if c not in cov and c != "cellStyles"]
            if unexplained:
                put(n, "nepřiřazeno", _("styly: {x}", x=", ".join(unexplained)))
            elif ctx.format_changed or ctx.rules_changed:
                put(n, "vysvětleno", _("formát porovnán v buňkách a pravidlech"))
            else:
                put(n, "šum", _("Styly – změny bez vlivu na použité buňky"))
        elif n in ctx.explained:
            put(n, "vysvětleno")
        elif "/theme/" in n:
            put(n, "nepřiřazeno", _("motiv (barvy a písma)"))
        elif "/drawings/" in n and n.endswith(".xml"):
            put(n, "nepřiřazeno", _("kresba – jiné rozdíly"))
        elif "vmlDrawing" in n:
            put(n, "nepřiřazeno", _("VML (okna komentářů, ovládací prvky)"))
        elif "comments" in n.lower():
            if "komentare" in ctx.layers:
                put(n, "šum", _("Formát komentářů (text a autoři stejní)"))
            else:
                put(n, "nepřiřazeno", _("komentáře"))
        elif "/persons/" in n:
            put(n, "šum", _("Seznam osob vláknových komentářů"))
        else:
            put(n, "nepřiřazeno", _("obsah nepřiřazen"))

    if not any_diff and pa.file_hash != pb.file_hash:
        ctx.noise[_("Jen komprese nebo pořadí částí v ZIP")] += 1


# ---------------------------------------------------------------- celkově

def compare(path_a: str, path_b: str, tolerance: float = 1e-9,
            layers: tuple[str, ...] = LAYERS) -> Result:
    res = Result(_info(path_a), _info(path_b), tolerance=tolerance)
    pa, pb = Package(path_a), Package(path_b)
    if pa.file_hash == pb.file_hash:
        res.headline = _("Bez rozdílů – soubory jsou totožné.")
        res.exit_code = 0
        return res
    a, b = Workbook.load(pa), Workbook.load(pb)
    ctx = _Ctx(a, b, tolerance, layers)
    order = [s.name for s in b.sheets] + [s.name for s in a.sheets if not b.sheet(s.name)]
    ctx.sheet_order = {n: i for i, n in enumerate(order)}
    res.sheets = order

    pairs = _compare_sheets(ctx)
    if "struktura" not in layers:
        ctx.changes.clear()
    if "bunky" in layers:
        for sa, sb in pairs:
            m = sheet_map(a.cells(sa), b.cells(sb))
            if m:
                ctx.maps[sa.name] = m
    if "struktura" in layers:
        _compare_names(ctx)
    for sa, sb in pairs:
        if layers_any(layers, "bunky", "format"):
            _compare_cells(ctx, sa, sb)
        if "tvary" in layers:
            _compare_shapes(ctx, sa, sb)
        if "pravidla" in layers:
            _compare_rules(ctx, sa, sb)
        if "komentare" in layers:
            _compare_comments(ctx, sa, sb)
    if "balicek" in layers:
        _compare_package(ctx, res)

    first, last = _("Struktura"), _("Balíček")
    ctx.changes.sort(key=lambda c: (c.layer != first, c.layer == last, c.sort))
    res.changes = ctx.changes
    res.noise = sorted(((k, str(v)) for k, v in ctx.noise.items()), key=lambda x: x[0])
    for c in ctx.changes:
        if c.cat in NOISE_CATS:
            res.noise.append((f"{c.kind} – {c.sheet} {c.loc}", c.note))
    _verdict(res, layers)
    return res


def layers_any(layers, *names) -> bool:
    return any(n in layers for n in names)


def _info(path: str) -> FileInfo:
    st = os.stat(path)
    return FileInfo(os.path.abspath(path), st.st_size, st.st_mtime)


def _verdict(res: Result, layers):
    cnt = res.counts
    cats = set(cnt)
    if not cats:
        if res.noise:
            res.headline = _("Jen technické rozdíly (šum) – obsah sešitu je stejný.")
            res.exit_code = 1
        else:
            res.headline = _("Bez rozdílů.")
            res.exit_code = 0
        return
    res.exit_code = 2
    if cats == {"text"}:
        res.headline = _("Změněny jen texty buněk – vzorce a čísla beze změny.")
    elif cats == {"shape_text"}:
        res.headline = _("Změněny jen texty v textových polích a tvarech – buňky beze změny.")
    elif cats <= {"text", "shape_text"}:
        res.headline = _("Změněny jen texty (v buňkách a v tvarech) – vzorce a čísla beze změny.")
    elif cats == {"struct"}:
        res.headline = _("Změněna jen struktura (listy, vložené či smazané řádky a sloupce) – "
                         "ostatní obsah beze změny.")
    elif cats == {"format"}:
        res.headline = _("Změněn jen formát – hodnoty, texty a vzorce beze změny.")
    elif not cats & (_CONTENT | {"struct"}):
        names = [cat_name(c) for c in CATEGORIES if c in cats]
        res.headline = _("Obsah buněk beze změny – liší se jen: {x}.",
                         x=", ".join(n[0].lower() + n[1:] for n in names))
    else:
        res.headline = _("Sešity se liší obsahem.")
    res.facts = [f"{cat_name(k)}: {v}" for k, v in
                 sorted(cnt.items(), key=lambda kv: list(CATEGORIES).index(kv[0])
                        if kv[0] in CATEGORIES else 99)]
    if "bunky" in layers:
        same = []
        if not cats & {"formula"}:
            same.append(_("Vzorce beze změny."))
        if not cats & {"number", "type", "cache"}:
            same.append(_("Čísla a vypočtené hodnoty beze změny."))
        if not cats & (_CONTENT | {"format"}):
            same.append(_("Buňky beze změny."))
        if same and "struct" in cats:
            same.insert(0, _("Mimo vložené a smazané řádky a sloupce:"))
        res.facts += same
    if "format" in layers and "format" not in cats and cats & _CONTENT:
        res.facts.append(_("Formát beze změny."))
    if "other" in cats:
        res.facts.append(_("Pozor: některé části souboru se liší a žádná vrstva je zatím neumí "
                           "vyložit – viz sekce Balíček."))
