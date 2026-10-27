"""Detekce vložených a smazaných řádků / sloupců a přepočet odkazů starého souboru.

Řádky (sloupce) se porovnávají podle obsahu bez polohy: podpis řádku je posloupnost
hodnot jeho neprázdných buněk. Vložení prázdného sloupce tak nemění podpisy řádků
a naopak. Zarovnání se přijme jen tehdy, když vysvětlí víc shodných řádků než
prosté párování podle adresy.
"""
from __future__ import annotations

import bisect
import re
from collections import Counter
from difflib import SequenceMatcher

from .ooxml import SheetCells, addr_key, col_index, col_letter

_MAX_ALIGN = 40000  # nad tuto délku rozdílného úseku se zarovnání nezkouší


class AxisMap:
    """Párování indexů jedné osy (řádky 1.., nebo sloupce 0..) starý -> nový."""

    def __init__(self, pairs: list[tuple[int, int]], inserted: list[int], deleted: list[int]):
        self.a2b = dict(pairs)
        self.b2a = {b: a for a, b in pairs}
        self.inserted = inserted   # indexy v novém souboru
        self.deleted = deleted     # indexy ve starém souboru
        self._inserted_set = set(inserted)
        self._deleted_set = set(deleted)
        self._keys = sorted(self.a2b)

    def map_a(self, a: int) -> int:
        """Kam se posunul index starého souboru; smazaný index padne na následující."""
        if a in self.a2b:
            return self.a2b[a]
        i = bisect.bisect_left(self._keys, a)
        if i < len(self._keys):
            return self.a2b[self._keys[i]]
        if not self._keys:
            return a
        last = self._keys[-1]
        return self.a2b[last] + (a - last)

    def kept_a(self, a: int) -> bool:
        return a not in self._deleted_set

    def kept_b(self, b: int) -> bool:
        return b not in self._inserted_set


class SheetMap:
    def __init__(self, rows: AxisMap | None, cols: AxisMap | None):
        self.rows, self.cols = rows, cols

    def a2b(self, addr: str) -> str | None:
        r, c = addr_key(addr)
        if self.rows:
            if not self.rows.kept_a(r):
                return None
            r = self.rows.map_a(r)
        if self.cols:
            if not self.cols.kept_a(c):
                return None
            c = self.cols.map_a(c)
        return f"{col_letter(c)}{r}"

    def kept_b(self, addr: str) -> bool:
        r, c = addr_key(addr)
        return (not self.rows or self.rows.kept_b(r)) and (not self.cols or self.cols.kept_b(c))

    def row(self, r: int) -> int:
        return self.rows.map_a(r) if self.rows else r

    def col(self, c: int) -> int:
        return self.cols.map_a(c) if self.cols else c


def _signatures(cells: SheetCells):
    rows: dict[int, list] = {}
    cols: dict[int, list] = {}
    for addr, c in cells.cells.items():
        if c.kind is None and c.formula is None:
            continue
        r, k = addr_key(addr)
        v = (c.kind, c.value) if c.kind is not None else ("f",)
        rows.setdefault(r, []).append((k, v))
        cols.setdefault(k, []).append((r, v))
    n_rows = max(rows, default=0)
    n_cols = max(cols, default=-1) + 1
    sig_r = [tuple(v for _, v in sorted(rows.get(i, ()))) for i in range(1, n_rows + 1)]
    sig_c = [tuple(v for _, v in sorted(cols.get(i, ()))) for i in range(n_cols)]
    return sig_r, sig_c


def _similarity(x: tuple, y: tuple) -> float:
    if not x and not y:
        return 1.0
    common = sum((Counter(x) & Counter(y)).values())
    return common / max(len(x), len(y))


def _pair_block(ba: list, bb: list) -> list[tuple[int, int]]:
    """Úsek, kde se zároveň vkládalo i upravovalo: řádky se spárují podle podobnosti obsahu
    (při zachování pořadí), zbytek jsou vložené / smazané řádky."""
    m, n = len(ba), len(bb)
    if m * n > 40000:
        k = min(m, n)
        return [(i, i) for i in range(k)]
    sim = [[_similarity(ba[i], bb[j]) for j in range(n)] for i in range(m)]
    best = [[0.0] * (n + 1) for _ in range(m + 1)]
    for i in range(1, m + 1):
        for j in range(1, n + 1):
            best[i][j] = max(best[i - 1][j], best[i][j - 1], best[i - 1][j - 1] + sim[i - 1][j - 1])
    out = []
    i, j = m, n
    while i and j:
        s = sim[i - 1][j - 1]
        if s > 0 and best[i][j] == best[i - 1][j - 1] + s:
            out.append((i - 1, j - 1))
            i, j = i - 1, j - 1
        elif best[i][j] == best[i - 1][j]:
            i -= 1
        else:
            j -= 1
    return out[::-1]


def _align(sa: list, sb: list, base: int) -> AxisMap | None:
    if sa == sb:
        return None
    na, nb = len(sa), len(sb)
    eq_addr = sum(1 for i in range(min(na, nb)) if sa[i] and sa[i] == sb[i])
    p = 0
    while p < na and p < nb and sa[p] == sb[p]:
        p += 1
    s = 0
    while s < na - p and s < nb - p and sa[na - 1 - s] == sb[nb - 1 - s]:
        s += 1
    mid_a, mid_b = sa[p:na - s], sb[p:nb - s]
    if len(mid_a) + len(mid_b) > _MAX_ALIGN:
        return None
    pairs = [(i, i) for i in range(p)]
    eq_map = sum(1 for i in range(p) if sa[i])
    shifted = False
    sm = SequenceMatcher(None, mid_a, mid_b, autojunk=False)
    for op, i1, i2, j1, j2 in sm.get_opcodes():
        if op == "equal":
            pairs += [(p + i1 + k, p + j1 + k) for k in range(i2 - i1)]
            eq_map += sum(1 for k in range(i1, i2) if mid_a[k])
        elif op == "replace":
            if i2 - i1 == j2 - j1:   # upravené řádky na svém místě
                pairs += [(p + i1 + x, p + j1 + x) for x in range(i2 - i1)]
            else:
                pairs += [(p + i1 + x, p + j1 + y) for x, y in _pair_block(mid_a[i1:i2], mid_b[j1:j2])]
                shifted = True
        else:
            shifted = True
    pairs += [(na - s + k, nb - s + k) for k in range(s)]
    eq_map += sum(1 for k in range(s) if sa[na - s + k])
    if not shifted or eq_map <= eq_addr:
        return None
    matched_a = {a for a, _ in pairs}
    matched_b = {b for _, b in pairs}
    return AxisMap([(a + base, b + base) for a, b in pairs],
                   [b + base for b in range(nb) if b not in matched_b],
                   [a + base for a in range(na) if a not in matched_a])


def sheet_map(ca: SheetCells, cb: SheetCells) -> SheetMap | None:
    ra, ka = _signatures(ca)
    rb, kb = _signatures(cb)
    rows = _align(ra, rb, 1)
    cols = _align(ka, kb, 0)
    if rows is None and cols is None:
        return None
    return SheetMap(rows, cols)


# ---------------------------------------------------------------- odkazy ve vzorcích

_STR = re.compile(r'"(?:[^"]|"")*"')
_REF = re.compile(
    r"(?<![A-Za-z0-9_.\]])"
    r"((?:\[\d+\])?(?:'(?:[^']|'')+'|[A-Za-z0-9_.À-ɏ]+)!)?"
    r"(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?\d+:\$?\d+|\$?[A-Z]{1,3}:\$?[A-Z]{1,3})"
    r"(?![A-Za-z0-9_(!])")
_CELL = re.compile(r"^(\$?)([A-Z]{1,3})(\$?)(\d+)$")


def _adj_part(tok: str, m: SheetMap) -> str:
    c = _CELL.match(tok)
    if c:
        return f"{c.group(1)}{col_letter(m.col(col_index(c.group(2))))}{c.group(3)}{m.row(int(c.group(4)))}"
    if tok.lstrip("$").isdigit():
        return ("$" if tok.startswith("$") else "") + str(m.row(int(tok.lstrip("$"))))
    letters = tok.lstrip("$")
    return ("$" if tok.startswith("$") else "") + col_letter(m.col(col_index(letters)))


def adjust_refs(formula: str, sheet: str | None, maps: dict[str, SheetMap],
                renames: dict[str, str]) -> str:
    """Přepíše odkazy starého vzorce tak, jak by je Excel upravil po vložení / smazání
    řádků a sloupců a po přejmenování listů. sheet = list, na kterém vzorec leží."""
    if not maps and not renames:
        return formula

    def fix(mo):
        prefix, ref = mo.group(1) or "", mo.group(2)
        if prefix.startswith("["):
            return mo.group(0)
        name = None
        if prefix:
            name = prefix[:-1]
            if name.startswith("'"):
                name = name[1:-1].replace("''", "'")
        target = name if name is not None else sheet
        m = maps.get(target) if target is not None else None
        if m is not None:
            ref = ":".join(_adj_part(p, m) for p in ref.split(":"))
        if name is not None and name in renames:
            new = renames[name]
            quoted = not re.fullmatch(r"[A-Za-z_À-ɏ][A-Za-z0-9_.À-ɏ]*", new)
            prefix = ("'" + new.replace("'", "''") + "'!") if quoted else new + "!"
        return prefix + ref

    out = []
    pos = 0
    for s in _STR.finditer(formula):
        out.append(_REF.sub(fix, formula[pos:s.start()]))
        out.append(s.group(0))
        pos = s.end()
    out.append(_REF.sub(fix, formula[pos:]))
    return "".join(out)
