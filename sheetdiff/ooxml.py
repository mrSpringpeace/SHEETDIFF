"""Čtení sešitu .xlsx / .xlsm přímo z balíčku (ZIP + XML), bez Excelu.

Soubor se načte celý do paměti a jen se čte. Prvky se rozlišují podle
lokálního jména, takže funguje přechodný (transitional) i strict OOXML.
"""
from __future__ import annotations

import copy
import hashlib
import io
import posixpath
import re
import zipfile
from dataclasses import dataclass, field

from lxml import etree

from .i18n import _

_PARSER = etree.XMLParser(huge_tree=True, resolve_entities=False, no_network=True)
_ADDR_RE = re.compile(r"^\$?([A-Z]{1,3})\$?(\d+)$")


def local(el) -> str:
    tag = el.tag
    if not isinstance(tag, str):  # komentář, instrukce
        return ""
    return tag[tag.rfind("}") + 1:]


def children(el, name: str):
    return [c for c in el if local(c) == name]


def child(el, name: str):
    for c in el:
        if local(c) == name:
            return c
    return None


def rel_attr(el, name: str = "id"):
    """Atribut r:id / r:embed bez ohledu na jmenný prostor vazeb."""
    for k, v in el.attrib.items():
        if k.endswith("}" + name) and "relationships" in k:
            return v
    return None


def sha1(data: bytes) -> str:
    return hashlib.sha1(data).hexdigest()


def col_letter(idx: int) -> str:
    """0 -> A, 25 -> Z, 26 -> AA"""
    s = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def addr_key(addr: str) -> tuple[int, int]:
    m = _ADDR_RE.match(addr)
    if not m:
        return (0, 0)
    return (int(m.group(2)), col_index(m.group(1)))


# ---------------------------------------------------------------- balíček

@dataclass
class Rel:
    id: str
    type: str
    target: str  # úplná cesta části v balíčku, nebo externí cíl
    external: bool

    @property
    def kind(self) -> str:
        return self.type.rsplit("/", 1)[-1]


class Package:
    def __init__(self, path: str):
        self.path = path
        with open(path, "rb") as f:
            self.raw = f.read()
        self.file_hash = sha1(self.raw)
        zf = zipfile.ZipFile(io.BytesIO(self.raw))
        self.parts: dict[str, bytes] = {}
        for info in zf.infolist():
            if not info.is_dir():
                self.parts[info.filename] = zf.read(info)
        self._xml: dict[str, etree._Element] = {}
        self._rels: dict[str, list[Rel]] = {}

    def xml(self, part: str):
        if part not in self._xml:
            self._xml[part] = etree.fromstring(self.parts[part], _PARSER)
        return self._xml[part]

    @staticmethod
    def rels_name(part: str) -> str:
        d, f = posixpath.split(part)
        return posixpath.join(d, "_rels", f + ".rels")

    def rels(self, part: str) -> list[Rel]:
        """Vazby části; part='' znamená kořen balíčku."""
        if part in self._rels:
            return self._rels[part]
        name = "_rels/.rels" if part == "" else self.rels_name(part)
        out: list[Rel] = []
        if name in self.parts:
            base = posixpath.dirname(part)
            for r in self.xml(name):
                if local(r) != "Relationship":
                    continue
                target = r.get("Target", "")
                external = r.get("TargetMode") == "External"
                if not external:
                    if target.startswith("/"):
                        target = target[1:]
                    else:
                        target = posixpath.normpath(posixpath.join(base, target))
                out.append(Rel(r.get("Id", ""), r.get("Type", ""), target, external))
        self._rels[part] = out
        return out

    def rel_map(self, part: str) -> dict[str, Rel]:
        return {r.id: r for r in self.rels(part)}


# ---------------------------------------------------------------- buňky

@dataclass(slots=True, eq=True)
class Cell:
    kind: str | None      # s = text, n = číslo, b = logická, e = chyba, d = datum, None = prázdná
    value: str | None     # text, nebo číslo tak, jak je uložené v XML
    formula: tuple | None  # ('t', text) | ('a', text) maticový | ('s', si) závislý sdílený
    style: str


class SheetCells:
    def __init__(self, cells: dict[str, Cell], masters: dict[str, tuple[str, str]],
                 rows=None, cols=None, merges=None, cf=None, dv=None, links=None):
        self.cells = cells
        self.masters = masters  # si -> (adresa hlavní buňky, text vzorce)
        self.rows: dict[int, tuple] = rows or {}   # řádek -> (ruční výška, skrytý, úroveň osnovy)
        self.cols: dict[int, tuple] = cols or {}   # sloupec (0..) -> (šířka, skrytý, úroveň osnovy)
        self.merges: list[str] = merges or []
        self.cf: list[Rule] = cf or []        # podmíněné formátování
        self.dv: list[Rule] = dv or []        # ověření dat
        self.links: list[tuple] = links or []  # (buňka, r:id, místo v sešitu, zobrazený text)
        self._cache: dict[str, str] = {}

    def formula_text(self, addr: str) -> str | None:
        c = self.cells.get(addr)
        if c is None or c.formula is None:
            return None
        ftype, text = c.formula
        if ftype == "t":
            return text
        if ftype == "a":
            return "{" + text + "}"
        if addr in self._cache:
            return self._cache[addr]
        master = self.masters.get(text)
        if master is None:
            res = _("(sdílený vzorec {si})", si=text)
        else:
            m_addr, m_text = master
            res = _translate(m_text, m_addr, addr)
        self._cache[addr] = res
        return res


@dataclass
class Rule:
    """Pravidlo podmíněného formátování nebo ověření dat."""
    sqref: str
    desc: tuple               # vlastnosti "klíč=hodnota" (typ, operátor…) bez vzorců
    formulas: tuple = ()
    dxf: tuple | None = None  # ('id', index do dxfs) | ('inline', popis)


_SKIP_ATTRS = {"priority", "dxfId", "id", "sqref", "uid"}


def compact(el) -> str:
    """Krátký čitelný popis prvku a jeho potomků (bez rozšíření a identifikátorů)."""
    attrs = [f"{k}={v}" for k, v in sorted(el.attrib.items())
             if "}" not in k and k not in _SKIP_ATTRS]
    kids = [compact(c) for c in el if local(c) and local(c) not in ("extLst", "ext")]
    text = (el.text or "").strip() if not len(el) else ""
    inner = " ".join(attrs + kids + ([text] if text else []))
    return f"{local(el)}({inner})" if inner else local(el)


def _sqref(el) -> str:
    if el.get("sqref"):
        return el.get("sqref")
    c = child(el, "sqref")
    return (c.text or "").strip() if c is not None else ""


def _cf_rules(el) -> list[Rule]:
    out = []
    sqref = _sqref(el)
    for rule in el:
        if local(rule) != "cfRule":
            continue
        attrs = [f"{k}={v}" for k, v in sorted(rule.attrib.items())
                 if "}" not in k and k not in _SKIP_ATTRS]
        extra = [compact(c) for c in rule if local(c) not in ("formula", "f", "dxf", "extLst", "")]
        formulas = tuple((c.text or "") for c in rule if local(c) in ("formula", "f"))
        if rule.get("dxfId") is not None:
            dxf = ("id", rule.get("dxfId"))
        elif child(rule, "dxf") is not None:
            dxf = ("inline", dxf_desc(child(rule, "dxf")))
        else:
            dxf = None
        out.append(Rule(sqref, tuple(attrs + extra), formulas, dxf))
    return out


def _dv_rule(el) -> Rule:
    attrs = [f"{k}={v}" for k, v in sorted(el.attrib.items())
             if "}" not in k and k not in _SKIP_ATTRS]
    formulas = []
    for name in ("formula1", "formula2"):
        c = child(el, name)
        if c is not None:
            f = child(c, "f")
            formulas.append(((f.text if f is not None else c.text) or "").strip())
    return Rule(_sqref(el), tuple(attrs), tuple(formulas))


def _translate(formula: str, origin: str, dest: str) -> str:
    try:
        from openpyxl.formula.translate import Translator
        return Translator("=" + formula, origin=origin).translate_formula(dest)[1:]
    except Exception:
        return _("{f} (sdílený z {a})", f=formula, a=origin)


def _inline_text(el) -> str:
    """Text z <si> nebo <is>: <t> přímo, nebo v bězích <r>; fonetika <rPh> se vynechá."""
    parts = []
    for c in el:
        n = local(c)
        if n == "t":
            parts.append(c.text or "")
        elif n == "r":
            t = child(c, "t")
            if t is not None:
                parts.append(t.text or "")
    return "".join(parts)


def parse_shared_strings(data: bytes) -> list[str]:
    root = etree.fromstring(data, _PARSER)
    return [_inline_text(si) for si in root if local(si) == "si"]


def parse_sheet(data: bytes, shared: list[str]) -> SheetCells:
    cells: dict[str, Cell] = {}
    masters: dict[str, tuple[str, str]] = {}
    rows: dict[int, tuple] = {}
    cols: dict[int, tuple] = {}
    merges: list[str] = []
    cf: list[Rule] = []
    dv: list[Rule] = []
    links: list[tuple] = []
    row_no = 0
    col_no = -1
    for _ev, el in etree.iterparse(io.BytesIO(data), events=("end",), huge_tree=True,
                                 resolve_entities=False):
        n = local(el)
        if n == "row":
            r = el.get("r")
            info = (el.get("ht") if el.get("customHeight") in ("1", "true") else None,
                    el.get("hidden") in ("1", "true"), el.get("outlineLevel", "0"))
            if r and info != (None, False, "0"):
                rows[int(r)] = info
            el.clear()
            continue
        if n == "col":
            try:
                lo, hi = int(el.get("min")), min(int(el.get("max")), 16384)
                w = round(float(el.get("width", "0")), 2)
            except (TypeError, ValueError):
                continue
            info = (w, el.get("hidden") in ("1", "true"), el.get("outlineLevel", "0"))
            for i in range(lo - 1, hi):
                cols[i] = info
            continue
        if n == "mergeCell":
            if el.get("ref"):
                merges.append(el.get("ref"))
            continue
        if n == "conditionalFormatting":
            cf.extend(_cf_rules(el))
            continue
        if n == "dataValidation":
            dv.append(_dv_rule(el))
            continue
        if n == "hyperlink":
            links.append((el.get("ref", ""), rel_attr(el, "id"), el.get("location", ""),
                          el.get("display", "")))
            continue
        if n != "c":
            continue
        addr = el.get("r")
        if addr:
            row_no, col_no = addr_key(addr)
        else:  # adresa může chybět, pak buňky jdou za sebou
            parent = el.getparent()
            r = parent.get("r") if parent is not None else None
            if r and int(r) != row_no:
                row_no, col_no = int(r), -1
            col_no += 1
            addr = f"{col_letter(col_no)}{row_no}"
        t = el.get("t", "n")
        style = el.get("s", "0")
        v_text = None
        formula = None
        inline = None
        for c in el:
            cn = local(c)
            if cn == "v":
                v_text = c.text
            elif cn == "f":
                ft = c.get("t")
                text = c.text or ""
                if ft == "shared":
                    si = c.get("si")
                    if text:
                        masters[si] = (addr, text)
                        formula = ("t", text)
                    else:
                        formula = ("s", si)
                elif ft == "array":
                    formula = ("a", text)
                elif ft == "dataTable":
                    formula = ("t", f"TABLE({c.get('r1', '')},{c.get('r2', '')})")
                else:
                    formula = ("t", text)
            elif cn == "is":
                inline = _inline_text(c)
        if t == "s":
            kind = "s"
            value = None
            if v_text is not None:
                try:
                    value = shared[int(v_text)]
                except (ValueError, IndexError):
                    value = _("(neplatný řetězec {v})", v=v_text)
        elif t == "inlineStr":
            kind, value = "s", inline
        elif t == "str":
            kind, value = "s", v_text
        else:
            kind, value = t, v_text
        if value is None:
            kind = None
        cells[addr] = Cell(kind, value, formula, style)
        el.clear()
    return SheetCells(cells, masters, rows, cols, merges, cf, dv, links)


# ---------------------------------------------------------------- kresby

@dataclass
class Shape:
    key: str          # jméno#pořadí, pro párování mezi verzemi
    name: str
    kind: str         # Textové pole, Tvar, Obrázek, Graf, Skupina, Spojnice, Objekt
    anchor: str       # buňka levého horního rohu, např. F30
    geom: tuple       # (sl, odsaz, ř, odsaz) od, (…) do, (x, y, cx, cy); -1 = nezadáno
    text: str
    fmt: str          # otisk ostatních vlastností (formát, výplň, propojení…)
    media: str | None = None       # cesta obrázku
    media_hash: str | None = None
    chart: str | None = None       # cesta grafu
    drawing: str = ""


_ANCHOR_PARTS = {"from", "to", "pos", "ext", "clientData"}
_KINDS = {"sp": "Tvar", "pic": "Obrázek", "graphicFrame": "Objekt", "grpSp": "Skupina",
          "cxnSp": "Spojnice", "contentPart": "Objekt"}


def _shape_text(el) -> str:
    body = None
    for d in el.iter():
        if local(d) == "txBody":
            body = d
            break
    if body is None:
        return ""
    paras = []
    for p in body:
        if local(p) != "p":
            continue
        buf = []
        for r in p:
            rn = local(r)
            if rn in ("r", "fld"):
                t = child(r, "t")
                if t is not None:
                    buf.append(t.text or "")
            elif rn == "br":
                buf.append("\n")
        paras.append("".join(buf))
    return "\n".join(paras).strip("\n")


def _fmt_signature(el) -> str:
    """Otisk tvaru bez textu, bez id/vazeb a bez polohy - zbytek je formát a vlastnosti.

    Odstavce textu se vynechají celé: Excel při úpravě textu běžně přeskupí běhy
    a jejich formát, což by se jinak hlásilo jako změna formátu tvaru.
    """
    c = copy.deepcopy(el)
    for d in list(c.iter()):
        n = local(d)
        if n == "p" and d.getparent() is not None and local(d.getparent()) == "txBody":
            d.getparent().remove(d)
            continue
        if n == "cNvPr":
            d.attrib.pop("id", None)
        elif n == "xfrm":
            for x in list(d):
                if local(x) in ("off", "ext", "chOff", "chExt"):
                    d.remove(x)
        for k in [k for k in d.attrib if "relationships" in k]:
            del d.attrib[k]
    return sha1(etree.tostring(c, method="c14n"))


def _int(el, name) -> int:
    c = child(el, name) if el is not None else None
    try:
        return int(c.text)
    except (AttributeError, TypeError, ValueError):
        return -1


def _anchor_info(anchor) -> tuple[str, tuple]:
    frm, to = child(anchor, "from"), child(anchor, "to")
    pos, ext = child(anchor, "pos"), child(anchor, "ext")
    geom = tuple(_int(frm, n) for n in ("col", "colOff", "row", "rowOff")) +         tuple(_int(to, n) for n in ("col", "colOff", "row", "rowOff")) +         tuple(int(e.get(a, "-1")) if e is not None else -1
              for e, a in ((pos, "x"), (pos, "y"), (ext, "cx"), (ext, "cy")))
    return anchor_addr(geom), geom


def anchor_addr(geom: tuple) -> str:
    if geom[0] < 0 or geom[2] < 0:
        return "absolutně"
    return f"{col_letter(geom[0])}{geom[2] + 1}"


def parse_drawing(pkg: Package, part: str) -> list[Shape]:
    root = pkg.xml(part)
    rels = pkg.rel_map(part)
    shapes: list[Shape] = []
    seen: dict[str, int] = {}

    def add(el, anchor_addr, geom, prefix):
        n = local(el)
        if n == "AlternateContent":
            choice = child(el, "Choice")
            if choice is None:
                choice = child(el, "Fallback")
            if choice is not None:
                for c in choice:
                    add(c, anchor_addr, geom, prefix)
            return
        if n not in _KINDS:
            return
        nv = next((c for c in el if local(c).startswith("nv")), None)
        cnv = child(nv, "cNvPr") if nv is not None else None
        name = cnv.get("name", "") if cnv is not None else ""
        kind = _KINDS[n]
        if n == "sp":
            cnvsp = child(nv, "cNvSpPr") if nv is not None else None
            if cnvsp is not None and cnvsp.get("txBox") == "1":
                kind = "Textové pole"
        full = prefix + name
        seen[full] = seen.get(full, 0) + 1
        sh = Shape(key=f"{full}#{seen[full]}", name=full, kind=kind, anchor=anchor_addr,
                   geom=geom, text=_shape_text(el) if n == "sp" else "", fmt="", drawing=part)
        if n == "grpSp":
            # formát skupiny bez členů; členy jdou jako samostatné tvary
            g = copy.deepcopy(el)
            for c in list(g):
                if local(c) not in ("nvGrpSpPr", "grpSpPr"):
                    g.remove(c)
            sh.fmt = _fmt_signature(g)
            shapes.append(sh)
            for c in el:
                add(c, anchor_addr, geom, full + " / ")
            return
        sh.fmt = _fmt_signature(el)
        for d in el.iter():
            dn = local(d)
            if dn == "blip":
                rid = rel_attr(d, "embed")
                r = rels.get(rid) if rid else None
                if r and not r.external:
                    sh.media = r.target
                    data = pkg.parts.get(r.target)
                    sh.media_hash = sha1(data) if data is not None else None
            elif dn == "chart" and n == "graphicFrame":
                rid = rel_attr(d, "id")
                r = rels.get(rid) if rid else None
                if r and not r.external:
                    sh.chart = r.target
                    sh.kind = "Graf"
        shapes.append(sh)

    for anchor in root:
        if local(anchor) not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
            continue
        addr, geom = _anchor_info(anchor)
        for el in anchor:
            if local(el) not in _ANCHOR_PARTS:
                add(el, addr, geom, "")
    return shapes


def _texts(el) -> str:
    return "".join(t.text or "" for t in el.iter() if local(t) == "t")


def chart_info(pkg: Package, part: str) -> dict[str, str]:
    """Obsahové vlastnosti grafu: typ, názvy, osy, řady (odkazy na oblasti)."""
    root = pkg.xml(part)
    info: dict[str, str] = {}
    chart = child(root, "chart")
    if chart is None:
        return info
    title = child(chart, "title")
    if title is not None:
        info[_("Název")] = _texts(title) or _refs(title) or _("(automatický)")
    plot = child(chart, "plotArea")
    if plot is None:
        return info
    types = []
    n_ser = 0
    for grp in plot:
        g = local(grp)
        if g.endswith("Chart"):
            types.append(g)
            for ser in children(grp, "ser"):
                n_ser += 1
                desc = []
                for sub in ser:
                    sn = local(sub)
                    if sn in ("tx", "cat", "val", "xVal", "yVal", "bubbleSize"):
                        ref = _refs(sub)
                        if not ref and sn == "tx":
                            ref = _texts(sub) or "".join(v.text or "" for v in sub.iter()
                                                         if local(v) == "v")
                        desc.append(f"{sn}: {ref}")
                info[_("Řada {n}", n=n_ser)] = " | ".join(desc)
        elif g.endswith("Ax"):
            ax_id = child(grp, "axId")
            t = child(grp, "title")
            label = _("Osa") + f" {g} {ax_id.get('val') if ax_id is not None else ''}".strip()
            info[label] = (_texts(t) or _refs(t)) if t is not None else _("(bez názvu)")
    info[_("Typ")] = ", ".join(types)
    return info


def _refs(el) -> str:
    return ", ".join(f.text or "" for f in el.iter() if local(f) == "f")


def chart_without_cache(pkg: Package, part: str) -> str:
    c = copy.deepcopy(pkg.xml(part))
    for d in list(c.iter()):
        if local(d) in ("numCache", "strCache", "multiLvlStrCache"):
            d.getparent().remove(d)
    return sha1(etree.tostring(c, method="c14n"))


# ---------------------------------------------------------------- styly

_BORDER_NAMES = {"left": "vlevo", "right": "vpravo", "top": "nahoře", "bottom": "dole",
                 "diagonal": "úhlopříčka", "start": "vlevo", "end": "vpravo"}


def _color(el) -> str:
    if el is None:
        return ""
    if el.get("rgb"):
        v = el.get("rgb")[-6:]
    elif el.get("theme") is not None:
        v = f"motiv {el.get('theme')}"
    elif el.get("indexed") is not None:
        v = f"index {el.get('indexed')}"
    elif el.get("auto") in ("1", "true"):
        v = "auto"
    else:
        return ""
    try:
        tint = float(el.get("tint", "0"))
    except ValueError:
        tint = 0.0
    return v + (f" {tint:+.0%}" if tint else "")


def _flag(el, name) -> bool:
    c = child(el, name)
    return c is not None and c.get("val", "1") not in ("0", "false", "none")


def _font(el) -> tuple[str, str]:
    if el is None:
        return "", ""
    name = child(el, "name")
    size = child(el, "sz")
    desc = [name.get("val", "") if name is not None else "",
            size.get("val", "") if size is not None else ""]
    for tag, label in (("b", "tučné"), ("i", "kurzíva"), ("u", "podtržené"),
                       ("strike", "přeškrtnuté")):
        if _flag(el, tag):
            desc.append(_(label))
    va = child(el, "vertAlign")
    if va is not None and va.get("val") not in (None, "baseline"):
        desc.append(va.get("val"))
    return " ".join(d for d in desc if d), _color(child(el, "color"))


def _fill(el) -> str:
    if el is None:
        return _("žádná")
    pf = child(el, "patternFill")
    if pf is not None:
        pt = pf.get("patternType", "none")
        if pt == "none":
            return _("žádná")
        fg, bg = _color(child(pf, "fgColor")), _color(child(pf, "bgColor"))
        return _("plná {c}", c=fg).strip() if pt == "solid" else f"{pt} {fg}/{bg}"
    gf = child(el, "gradientFill")
    if gf is not None:
        return _("přechod") + " " + ", ".join(_color(c) for c in gf.iter() if local(c) == "color")
    return _("žádná")


def _border(el) -> str:
    if el is None:
        return _("žádné")
    out = []
    for side in el:
        n = local(side)
        if n in _BORDER_NAMES and side.get("style"):
            out.append(f"{_(_BORDER_NAMES[n])} {side.get('style')} {_color(child(side, 'color'))}".strip())
    up = el.get("diagonalUp") in ("1", "true")
    down = el.get("diagonalDown") in ("1", "true")
    if up or down:
        out.append(_("úhlopříčka") + " " + ("↗" if up else "") + ("↘" if down else ""))
    return ", ".join(out) or _("žádné")


def dxf_desc(el) -> str:
    """Formát zvýraznění podmíněného formátování (dxf): jen to, co pravidlo nastavuje."""
    if el is None:
        return ""
    out = []
    font = child(el, "font")
    if font is not None:
        f, color = _font(font)
        out.append(_("písmo") + " " + " ".join(x for x in (f, color) if x))
    nf = child(el, "numFmt")
    if nf is not None:
        out.append(_("formát") + f" {nf.get('formatCode', '')}")
    fill = child(el, "fill")
    if fill is not None:
        pf = child(fill, "patternFill")
        if pf is not None:
            color = _color(child(pf, "bgColor")) or _color(child(pf, "fgColor"))
            out.append((_("výplň") + f" {pf.get('patternType', '')} {color}").replace("  ", " ").strip())
        else:
            out.append(_fill(fill))
    border = child(el, "border")
    if border is not None:
        out.append(_("ohraničení") + " " + _border(border))
    return "; ".join(out) or _("bez formátu")


def parse_dxfs(pkg: "Package", part: str | None) -> list[str]:
    if not part or part not in pkg.parts:
        return []
    d = child(pkg.xml(part), "dxfs")
    return [dxf_desc(x) for x in d if local(x) == "dxf"] if d is not None else []


def parse_styles(pkg: "Package", part: str | None) -> list[tuple]:
    """Popis formátu pro každý index stylu buňky (cellXfs) jako n-tice (vlastnost, hodnota)."""
    if not part or part not in pkg.parts:
        return []
    try:
        from openpyxl.styles.numbers import BUILTIN_FORMATS
    except Exception:
        BUILTIN_FORMATS = {0: "General"}
    root = pkg.xml(part)

    def items(name, sub):
        c = child(root, name)
        return [x for x in c if local(x) == sub] if c is not None else []

    numfmts = {int(n.get("numFmtId", "-1")): n.get("formatCode", "")
               for n in items("numFmts", "numFmt")}
    fonts = [_font(f) for f in items("fonts", "font")]
    fills = [_fill(f) for f in items("fills", "fill")]
    borders = [_border(b) for b in items("borders", "border")]

    def pick(lst, idx, default):
        try:
            return lst[int(idx)]
        except (TypeError, ValueError, IndexError):
            return default

    out = []
    for xf in items("cellXfs", "xf"):
        nid = int(xf.get("numFmtId", "0"))
        font, font_color = pick(fonts, xf.get("fontId", 0), ("", ""))
        al = child(xf, "alignment")
        align = " ".join(f"{k}={v}" for k, v in sorted(al.attrib.items())) if al is not None else ""
        pr = child(xf, "protection")
        locked = pr.get("locked", "1") if pr is not None else "1"
        hidden = pr.get("hidden", "0") if pr is not None else "0"
        out.append((
            (_("Číselný formát"), numfmts.get(nid, BUILTIN_FORMATS.get(nid, f"#{nid}"))),
            (_("Písmo"), font),
            (_("Barva písma"), font_color or _("automatická")),
            (_("Výplň"), pick(fills, xf.get("fillId", 0), _("žádná"))),
            (_("Ohraničení"), pick(borders, xf.get("borderId", 0), _("žádné"))),
            (_("Zarovnání"), align or _("výchozí")),
            (_("Zámek"), (_("zamčeno") if locked in ("1", "true") else _("odemčeno")) +
             (_(", skrytý vzorec") if hidden in ("1", "true") else "")),
        ))
    return out


# ---------------------------------------------------------------- komentáře

def parse_comments(pkg: "Package", sheet_part: str, persons: dict[str, str]) -> tuple[dict, list[str]]:
    """Komentáře listu: adresa -> (autor, text) a seznam částí, které je nesou."""
    legacy: dict[str, tuple[str, str]] = {}
    threaded: dict[str, list[tuple[str, str]]] = {}
    parts: list[str] = []
    for r in pkg.rels(sheet_part):
        if r.external or r.target not in pkg.parts:
            continue
        if r.kind == "comments":
            parts.append(r.target)
            root = pkg.xml(r.target)
            au = child(root, "authors")
            authors = [a.text or "" for a in au] if au is not None else []
            cl = child(root, "commentList")
            for c in (cl if cl is not None else []):
                if local(c) != "comment":
                    continue
                try:
                    author = authors[int(c.get("authorId", "0"))]
                except (ValueError, IndexError):
                    author = ""
                t = child(c, "text")
                legacy[c.get("ref", "")] = (author, _inline_text(t) if t is not None else "")
        elif r.kind == "threadedComment":
            parts.append(r.target)
            for c in pkg.xml(r.target):
                if local(c) != "threadedComment":
                    continue
                t = child(c, "text")
                text = (t.text or "") if t is not None else ""
                threaded.setdefault(c.get("ref", ""), []).append(
                    (persons.get(c.get("personId", ""), ""), text))
        elif r.kind == "vmlDrawing":
            parts.append(r.target)
    out = dict(legacy)
    for ref, items in threaded.items():
        author = items[0][0]
        text = items[0][1] + "".join(f"\n↳ {a}: {t}" for a, t in items[1:])
        out[ref] = (author, text)
    return out, parts


# ---------------------------------------------------------------- sešit

@dataclass
class SheetInfo:
    name: str
    index: int
    state: str        # visible, hidden, veryHidden
    part: str | None
    kind: str         # worksheet, chartsheet, dialogsheet …


@dataclass
class Workbook:
    pkg: Package
    part: str = ""
    sheets: list[SheetInfo] = field(default_factory=list)
    names: dict[str, str] = field(default_factory=dict)    # "název [list]" -> odkaz
    external: list[str] = field(default_factory=list)     # cesty externích sešitů
    shared: list[str] = field(default_factory=list)
    shared_part: str | None = None
    styles_part: str | None = None
    persons: dict = field(default_factory=dict)   # id osoby -> jméno (vláknové komentáře)
    _styles: list | None = None
    _dxfs: list | None = None
    _cells: dict = field(default_factory=dict)
    _shapes: dict = field(default_factory=dict)

    @classmethod
    def load(cls, pkg: Package) -> "Workbook":
        wb = cls(pkg)
        doc = next((r for r in pkg.rels("") if r.kind == "officeDocument"), None)
        if doc is None:
            raise ValueError(_("Soubor není sešit Excel (chybí hlavní část workbook)."))
        wb.part = doc.target
        root = pkg.xml(wb.part)
        rels = pkg.rel_map(wb.part)
        sheets_el = child(root, "sheets")
        for i, s in enumerate(sheets_el if sheets_el is not None else []):
            if local(s) != "sheet":
                continue
            r = rels.get(rel_attr(s, "id") or "")
            wb.sheets.append(SheetInfo(
                name=s.get("name", ""), index=i, state=s.get("state", "visible"),
                part=r.target if r else None, kind=r.kind if r else "?"))
        by_index = {i: s.name for i, s in enumerate(wb.sheets)}
        dn = child(root, "definedNames")
        for d in (dn if dn is not None else []):
            if local(d) != "definedName":
                continue
            scope = d.get("localSheetId")
            key = d.get("name", "")
            if scope is not None:
                key += f" [{by_index.get(int(scope), scope)}]"
            wb.names[key] = d.text or ""
        for r in pkg.rels(wb.part):
            if r.kind == "sharedStrings":
                wb.shared_part = r.target
            elif r.kind == "styles":
                wb.styles_part = r.target
            elif r.kind == "person" and r.target in pkg.parts:
                for p in pkg.xml(r.target):
                    if local(p) == "person":
                        wb.persons[p.get("id", "")] = p.get("displayName", "")
            elif r.kind == "externalLink":
                tgt = [x.target for x in pkg.rels(r.target) if x.external]
                wb.external.append(", ".join(tgt) or r.target)
        if wb.shared_part and wb.shared_part in pkg.parts:
            wb.shared = parse_shared_strings(pkg.parts[wb.shared_part])
        return wb

    def sheet(self, name: str) -> SheetInfo | None:
        return next((s for s in self.sheets if s.name == name), None)

    def cells(self, sh: SheetInfo) -> SheetCells:
        if sh.name not in self._cells:
            data = self.pkg.parts.get(sh.part or "")
            if data is None or sh.kind != "worksheet":
                self._cells[sh.name] = SheetCells({}, {})
            else:
                self._cells[sh.name] = parse_sheet(data, self.shared)
        return self._cells[sh.name]

    def drawing_parts(self, sh: SheetInfo) -> list[str]:
        if not sh.part:
            return []
        return [r.target for r in self.pkg.rels(sh.part)
                if r.kind == "drawing" and r.target in self.pkg.parts]

    def shapes(self, sh: SheetInfo) -> list[Shape]:
        if sh.name not in self._shapes:
            out = []
            for d in self.drawing_parts(sh):
                out.extend(parse_drawing(self.pkg, d))
            self._shapes[sh.name] = out
        return self._shapes[sh.name]

    @property
    def styles(self) -> list[tuple]:
        if self._styles is None:
            self._styles = parse_styles(self.pkg, self.styles_part)
        return self._styles

    def dxf(self, ref: tuple | None) -> str:
        if ref is None:
            return ""
        if ref[0] == "inline":
            return ref[1]
        if self._dxfs is None:
            self._dxfs = parse_dxfs(self.pkg, self.styles_part)
        try:
            return self._dxfs[int(ref[1])]
        except (ValueError, IndexError):
            return f"dxf {ref[1]}"

    def style(self, idx: str) -> tuple:
        try:
            return self.styles[int(idx)]
        except (ValueError, IndexError):
            return ()

    def comments(self, sh: SheetInfo) -> tuple[dict, list[str]]:
        if not sh.part:
            return {}, []
        return parse_comments(self.pkg, sh.part, self.persons)
