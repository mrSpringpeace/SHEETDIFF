"""Vstupní bod SHEETDIFF.exe.

Bez přepínačů otevře okno (0–2 soubory se předvyplní, se dvěma se rovnou porovná).
S přepínačem (--cli, -o, --lang …) běží jako příkazová řádka:
    SHEETDIFF.exe --cli stary.xlsx novy.xlsx -o report.html
"""
import sys


def main():
    args = sys.argv[1:]
    if args and any(a.startswith("-") for a in args):
        from sheetdiff.cli import main as cli
        sys.exit(cli([a for a in args if a != "--cli"]))
    from sheetdiff.gui import main as gui
    gui(args)


if __name__ == "__main__":
    main()
