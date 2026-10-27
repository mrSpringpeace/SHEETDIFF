"""Okno a přetažení souborů (jen Windows). Spouští ho acceptance.py v samostatném procesu,
protože chyba v proceduře okna shodí celý Python a test by ji jinak neviděl.

Zprávy WM_DROPFILES jdou frontou (PostMessage) jako při skutečném přetažení z Průzkumníka.
Na konci vypíše GUI OK, jinak popis chyby.
"""
import os
import sys
import tempfile
import time
import tkinter as tk

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

import demo  # noqa: E402
from sheetdiff import gui, win_drop  # noqa: E402
from sheetdiff.i18n import set_lang  # noqa: E402

TMP = tempfile.mkdtemp(prefix="sheetdiff_gui_")
gui._settings_path = lambda: os.path.join(TMP, "settings.json")
gui.os.startfile = lambda p: None   # report neotvírat
A = demo.build(os.path.join(TMP, "a.xlsx"))
time.sleep(1.1)   # novější datum uložení pro B
B = demo.build(os.path.join(TMP, "b.xlsx"), ["text"])
set_lang("cs")

root = tk.Tk()
app = gui.App(root, [], {})
errors = []


def expect(cond, msg):
    if not cond:
        errors.append(msg)


def client_xy(entry):
    root.update()
    return (entry.winfo_rootx() - root.winfo_rootx() + 10, entry.winfo_rooty() - root.winfo_rooty() + 5)


steps = [
    lambda: win_drop.simulate_drop(root, [A], client_xy(app.new_entry)),
    lambda: expect(app.new.get() == A and not app.old.get(), f"jeden soubor na B: {app.old.get()!r} {app.new.get()!r}"),
    lambda: win_drop.simulate_drop(root, [B, A], (5, 5)),
    lambda: expect((app.old.get(), app.new.get()) == (A, B), "dva soubory: starší má být A"),
    lambda: (app.lang.set("English"), app.change_lang()),
    lambda: win_drop.simulate_drop(root, [B], client_xy(app.old_entry)),
    lambda: expect(app.old.get() == B and app.status.cget("text").startswith("File set"),
                   f"po přepnutí jazyka: {app.status.cget('text')!r}"),
    lambda: (app.old.set(A), app.new.set(B), app.start()),
    lambda: win_drop.simulate_drop(root, [A], (5, 5)),   # přetažení během porovnání
]


def run(i=0):
    if i < len(steps):
        try:
            steps[i]()
        except Exception as e:
            errors.append(f"krok {i}: {e!r}")
        root.after(400, run, i + 1)
    elif app.btn.instate(["disabled"]):
        root.after(200, run, i)
    else:
        expect(app.last_report is not None, "porovnání nedoběhlo")
        root.destroy()


root.after(300, run)
root.after(30000, root.destroy)   # pojistka
root.mainloop()
print("GUI OK" if not errors else "GUI CHYBA: " + "; ".join(errors))
