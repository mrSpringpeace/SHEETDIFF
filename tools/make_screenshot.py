"""Screenshot anglického reportu z ukázkového rozpočtu pro README (docs/screenshot_report*.png).

Cesty k souborům se v reportu nahradí neutrálními, aby screenshot neukazoval místní disk.
Potřebuje Microsoft Edge nebo Chrome (headless).
    python tools/make_screenshot.py
"""
import os
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "tools"))

import demo  # noqa: E402
from sheetdiff.compare import compare  # noqa: E402
from sheetdiff.i18n import set_lang  # noqa: E402
from sheetdiff.report import render  # noqa: E402

BROWSERS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
]


def main():
    set_lang("en")
    tmp = tempfile.mkdtemp(prefix="sheetdiff_shot_")
    a = demo.build(os.path.join(tmp, "office_budget_v1.xlsx"))
    b = demo.build(os.path.join(tmp, "office_budget_v2.xlsx"), demo.ALL)
    res = compare(a, b)
    res.a.path = r"D:\Projects\Office\office_budget_v1.xlsx"
    res.b.path = r"D:\Projects\Office\office_budget_v2.xlsx"
    page = os.path.join(tmp, "report.html")
    with open(page, "w", encoding="utf-8") as f:
        f.write(render(res))
    browser = next((p for p in BROWSERS if os.path.isfile(p)), None)
    if browser is None:
        sys.exit("Edge ani Chrome nenalezen.")
    out_dir = os.path.join(ROOT, "docs")
    os.makedirs(out_dir, exist_ok=True)
    for name, extra in (("screenshot_report.png", []),
                        ("screenshot_report_dark.png", ["--force-dark-mode",
                                                        "--blink-settings=preferredColorScheme=0"])):
        out = os.path.join(out_dir, name)
        subprocess.run([browser, "--headless=new", "--disable-gpu", "--hide-scrollbars",
                        "--force-device-scale-factor=1", "--window-size=1320,1560",
                        f"--user-data-dir={os.path.join(tmp, 'profile')}",
                        f"--screenshot={out}", *extra, "file:///" + page.replace("\\", "/")],
                       check=True, capture_output=True, timeout=120)
        print(out)
    gui_shot(res, os.path.join(out_dir, "screenshot_window.png"))


def gui_shot(res, out: str):
    """Snímek okna (spuštěného z Pythonu) s vyplněnými neutrálními cestami a verdiktem."""
    import tkinter as tk
    from tkinter import ttk

    from PIL import ImageGrab

    from sheetdiff import gui
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    if "vista" in ttk.Style(root).theme_names():
        ttk.Style(root).theme_use("vista")
    app = gui.App(root, [res.a.path, res.b.path], {"excel": True})
    app.set_status(res.headline + "\n" + " · ".join(res.facts[:4]) + " …", gui.TONES[2])
    app.last_report = "report.html"
    app.open_btn.grid(row=0, column=1, padx=12)
    root.attributes("-topmost", True)
    root.geometry("+80+80")

    def grab():
        root.update()
        x, y = root.winfo_rootx(), root.winfo_rooty()
        ImageGrab.grab((x, y, x + root.winfo_width(), y + root.winfo_height())).save(out)
        print(out)
        root.destroy()

    root.after(800, grab)
    root.mainloop()


if __name__ == "__main__":
    main()
