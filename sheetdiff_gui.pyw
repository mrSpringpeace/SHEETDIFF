"""Spouštěč okna SHEETDIFF (pythonw, bez konzole)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from sheetdiff.gui import main
except ImportError as e:  # chybí lxml / openpyxl
    import tkinter.messagebox as mb
    mb.showerror("SHEETDIFF", f"Chybí knihovna: {e.name}\n\nNainstalujte ji příkazem:\n"
                              f"python -m pip install lxml openpyxl")
    sys.exit(1)

main()
