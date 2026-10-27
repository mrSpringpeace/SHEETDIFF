"""Příkazová řádka: python -m sheetdiff stary.xlsx novy.xlsx [-o report.html] [--lang en]"""
from __future__ import annotations

import argparse
import os
import sys
import tempfile

from . import __version__
from .compare import LAYERS, compare
from .i18n import _, set_lang
from .report import render


def default_report_path(a: str, b: str) -> str:
    folder = os.path.join(tempfile.gettempdir(), "SHEETDIFF")
    os.makedirs(folder, exist_ok=True)
    sa = os.path.splitext(os.path.basename(a))[0]
    sb = os.path.splitext(os.path.basename(b))[0]
    return os.path.join(folder, f"{sa}__vs__{sb}.html")


def run(a: str, b: str, out: str | None = None, tolerance: float = 1e-9,
        layers=LAYERS, excel: bool = False):
    """Porovná, zapíše report(y) a vrátí (výsledek, cesta HTML, cesta xlsx nebo None)."""
    res = compare(a, b, tolerance=tolerance, layers=tuple(layers))
    out = out or default_report_path(a, b)
    with open(out, "w", encoding="utf-8") as f:
        f.write(render(res))
    xlsx = None
    if excel:
        from .report_xlsx import write
        xlsx = os.path.splitext(out)[0] + ".xlsx"
        write(res, xlsx)
    return res, out, xlsx


def _peek_lang(argv) -> str:
    """Jazyk je potřeba znát už pro nápovědu, proto se přečte před argparse."""
    for i, a in enumerate(argv):
        if a in ("--lang", "--jazyk") and i + 1 < len(argv):
            return argv[i + 1]
        if a.startswith(("--lang=", "--jazyk=")):
            return a.split("=", 1)[1]
    return "cs"


def main(argv=None) -> int:
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(errors="replace")
        except AttributeError:
            pass
    argv = list(sys.argv[1:] if argv is None else argv)
    set_lang(_peek_lang(argv))
    p = argparse.ArgumentParser(prog="sheetdiff",
                                description=_("Porovná dva sešity Excel (.xlsx, .xlsm) a vytvoří HTML report."))
    p.add_argument("stary", metavar=_("STARY"), help=_("starší verze (A)"))
    p.add_argument("novy", metavar=_("NOVY"), help=_("novější verze (B)"))
    p.add_argument("-o", "--output", help=_("cesta HTML reportu (výchozí: %%TEMP%%\\SHEETDIFF)"))
    p.add_argument("-t", "--tolerance", type=float, default=1e-9,
                   help=_("relativní tolerance čísel (výchozí 1e-9)"))
    en = _peek_lang(argv) == "en"   # v nápovědě napřed názvy v jazyce uživatele

    def names(cs, eng):
        return (eng, cs) if en else (cs, eng)

    p.add_argument(*names("--bez", "--skip"), dest="bez", action="append", default=[], choices=LAYERS,
                   metavar=_("VRSTVA"),
                   help=_("vypnout vrstvu: {x} (lze opakovat)", x=", ".join(LAYERS)))
    p.add_argument("-V", "--version", action="version", version=f"SHEETDIFF {__version__}")
    p.add_argument("--excel", action="store_true",
                   help=_("vytvořit i Excel report (stejný název, přípona .xlsx)"))
    p.add_argument(*names("--jazyk", "--lang"), dest="lang", choices=("cs", "en"), default="cs",
                   help=_("jazyk reportu: cs nebo en"))
    p.add_argument(*names("--otevrit", "--open"), dest="otevrit", action="store_true",
                   help=_("otevřít report v prohlížeči"))
    args = p.parse_args(argv)
    set_lang(args.lang)
    for f in (args.stary, args.novy):
        if not os.path.isfile(f):
            print(_("Soubor neexistuje: {f}", f=f), file=sys.stderr)
            return 3
    try:
        res, out, xlsx = run(args.stary, args.novy, args.output, args.tolerance,
                             [x for x in LAYERS if x not in args.bez], args.excel)
    except Exception as e:  # poškozený soubor, nejde o xlsx …
        print(_("Chyba: {e}", e=e), file=sys.stderr)
        return 3
    print(res.headline)
    for f in res.facts:
        print("  " + f)
    if res.noise:
        print("  " + _("Šum: ") + "; ".join(k for k, _v in res.noise))
    print(_("Report: {p}", p=out))
    if xlsx:
        print(_("Excel:  {p}", p=xlsx))
    if args.otevrit:
        os.startfile(out)
    return res.exit_code
