"""Jednoduché okno: výběr dvou sešitů (i přetažením myší), Porovnat, report v prohlížeči."""
from __future__ import annotations

import json
import os
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import win_drop
from .cli import run
from .i18n import _, get_lang, set_lang, system_lang

TONES = {0: "#2f8a4f", 1: "#5d6677", 2: "#b06a00"}
LANG_NAMES = {"cs": "Čeština", "en": "English"}
EXCEL_EXT = (".xlsx", ".xlsm")


def _settings_path() -> str:
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "SHEETDIFF", "settings.json")


def load_settings() -> dict:
    try:
        with open(_settings_path(), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def save_settings(data: dict):
    try:
        os.makedirs(os.path.dirname(_settings_path()), exist_ok=True)
        with open(_settings_path(), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=1)
    except OSError:
        pass


def order_by_age(files: list[str]) -> list[str]:
    """Dva soubory seřadí tak, aby starší (podle data uložení) byl A."""
    try:
        return sorted(files, key=os.path.getmtime)
    except OSError:
        return files


class App(ttk.Frame):
    def __init__(self, root: tk.Tk, files: list[str], settings: dict):
        super().__init__(root, padding=16)
        self.root = root
        self.settings = settings
        root.minsize(680, 0)
        self.grid(sticky="nsew")
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)

        self.old = tk.StringVar(value=files[0] if len(files) > 0 else "")
        self.new = tk.StringVar(value=files[1] if len(files) > 1 else "")
        self.beside = tk.BooleanVar(value=settings.get("beside", False))
        self.excel = tk.BooleanVar(value=settings.get("excel", False))
        self.lang = tk.StringVar(value=LANG_NAMES[get_lang()])
        self.last_report: str | None = None
        self.status_text = ("", "")
        self.build()

        root.bind("<Return>", lambda e: self.start())
        root.bind("<Escape>", lambda e: root.destroy())
        self.drop_ok = win_drop.enable(root, self.on_drop)

    # ------------------------------------------------------------ rozvržení

    def build(self):
        for w in self.winfo_children():
            w.destroy()
        self.root.title(_("SHEETDIFF – porovnání dvou sešitů"))

        top = ttk.Frame(self)
        top.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        top.columnconfigure(0, weight=1)
        ttk.Label(top, text=_("Vyberte dva sešity nebo je sem přetáhněte myší."),
                  foreground="#5d6677").grid(row=0, column=0, sticky="w")
        ttk.Label(top, text=_("Jazyk:")).grid(row=0, column=1, padx=(8, 4))
        lang = ttk.Combobox(top, textvariable=self.lang, values=list(LANG_NAMES.values()),
                            state="readonly", width=10)
        lang.grid(row=0, column=2)
        lang.bind("<<ComboboxSelected>>", self.change_lang)

        ttk.Label(self, text=_("Starý soubor (A):")).grid(row=1, column=0, sticky="w", pady=4)
        self.old_entry = ttk.Entry(self, textvariable=self.old)
        self.old_entry.grid(row=1, column=1, sticky="ew", padx=8)
        ttk.Button(self, text=_("Vybrat…"), command=lambda: self.pick(self.old)).grid(row=1, column=2)

        ttk.Label(self, text=_("Nový soubor (B):")).grid(row=2, column=0, sticky="w", pady=4)
        self.new_entry = ttk.Entry(self, textvariable=self.new)
        self.new_entry.grid(row=2, column=1, sticky="ew", padx=8)
        ttk.Button(self, text=_("Vybrat…"), command=lambda: self.pick(self.new)).grid(row=2, column=2)

        opts = ttk.Frame(self)
        opts.grid(row=3, column=1, sticky="w", padx=8, pady=(4, 0))
        ttk.Button(opts, text=_("⇅ Prohodit A a B"), command=self.swap).pack(side="left")
        ttk.Checkbutton(opts, text=_("Uložit report vedle nového souboru"),
                        variable=self.beside).pack(side="left", padx=16)
        ttk.Checkbutton(opts, text=_("Také Excel report"), variable=self.excel).pack(side="left")

        ttk.Separator(self).grid(row=4, column=0, columnspan=3, sticky="ew", pady=12)

        bottom = ttk.Frame(self)
        bottom.grid(row=5, column=0, columnspan=3, sticky="ew")
        bottom.columnconfigure(0, weight=1)
        self.status = ttk.Label(bottom, wraplength=540, justify="left")
        self.status.grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(bottom, mode="indeterminate", length=120)
        self.open_btn = ttk.Button(bottom, text=_("Otevřít report znovu"), command=self.reopen)
        self.btn = ttk.Button(bottom, text=_("Porovnat"), command=self.start, default="active")
        self.btn.grid(row=0, column=2, sticky="e")
        if self.last_report:
            self.open_btn.grid(row=0, column=1, padx=12)
        text, color = self.status_text
        self.status.configure(text=text or _("Vyberte dva soubory a stiskněte Porovnat."),
                              foreground=color)

    def set_status(self, text: str, color: str = ""):
        self.status_text = (text, color)
        self.status.configure(text=text, foreground=color)

    def change_lang(self, _event=None):
        code = next(k for k, v in LANG_NAMES.items() if v == self.lang.get())
        if code == get_lang():
            return
        set_lang(code)
        self.settings["lang"] = code
        self.status_text = ("", "")
        self.save()
        self.build()

    def save(self):
        self.settings.update(excel=self.excel.get(), beside=self.beside.get())
        save_settings(self.settings)

    # ------------------------------------------------------------ soubory

    def pick(self, var: tk.StringVar):
        other = self.new.get() if var is self.old else self.old.get()
        start = os.path.dirname(var.get() or other) or self.settings.get("last_dir") or os.getcwd()
        path = filedialog.askopenfilename(
            parent=self.root, initialdir=start, title=_("Vyberte sešit"),
            filetypes=[(_("Sešity Excel"), "*.xlsx *.xlsm"), (_("Všechny soubory"), "*.*")])
        if path:
            var.set(os.path.normpath(path))
            self.settings["last_dir"] = os.path.dirname(path)

    def swap(self):
        a, b = self.old.get(), self.new.get()
        self.old.set(b)
        self.new.set(a)

    def on_drop(self, files: list[str], screen_xy):
        files = [f for f in files if os.path.isfile(f)]
        if not files:
            return
        bad = [f for f in files if not f.lower().endswith(EXCEL_EXT)]
        if len(files) >= 2:
            a, b = order_by_age(files[:2])
            self.old.set(a)
            self.new.set(b)
            msg = _("Starší soubor je A, novější B (podle data uložení). Lze prohodit.")
        else:
            target = self._drop_target(screen_xy)
            target.set(files[0])
            msg = _("Soubor vložen jako {x}.", x="A" if target is self.old else "B")
        if bad:
            msg += " " + _("Pozor: {f} není sešit .xlsx / .xlsm.", f=os.path.basename(bad[0]))
        self.settings["last_dir"] = os.path.dirname(files[0])
        self.set_status(msg)

    def _drop_target(self, screen_xy) -> tk.StringVar:
        y = screen_xy[1]
        for entry, var in ((self.old_entry, self.old), (self.new_entry, self.new)):
            top = entry.winfo_rooty() - 6
            if top <= y <= top + entry.winfo_height() + 12:
                return var
        if not self.old.get():
            return self.old
        return self.new

    # ------------------------------------------------------------ porovnání

    def start(self):
        if self.btn.instate(["disabled"]):
            return
        a, b = self.old.get().strip().strip('"'), self.new.get().strip().strip('"')
        for label, f in ((_("Starý"), a), (_("Nový"), b)):
            if not f:
                messagebox.showwarning("SHEETDIFF", _("{x} soubor není vybraný.", x=label),
                                       parent=self.root)
                return
            if not os.path.isfile(f):
                messagebox.showwarning("SHEETDIFF", _("{x} soubor neexistuje:\n{f}", x=label, f=f),
                                       parent=self.root)
                return
        self.save()
        out = None
        if self.beside.get():
            sa = os.path.splitext(os.path.basename(a))[0]
            sb = os.path.splitext(os.path.basename(b))[0]
            out = os.path.join(os.path.dirname(b), f"{sa}__vs__{sb}_diff.html")
        self.btn.state(["disabled"])
        self.open_btn.grid_remove()
        self.progress.grid(row=0, column=1, padx=12)
        self.progress.start(12)
        self.set_status(_("Porovnávám…"))
        threading.Thread(target=self.work, args=(a, b, out, self.excel.get()), daemon=True).start()

    def work(self, a, b, out, excel):
        try:
            res, path, xlsx = run(a, b, out, excel=excel)
            self.root.after(0, self.done, res, path, xlsx, None)
        except Exception as e:
            self.root.after(0, self.done, None, None, None, (e, traceback.format_exc()))

    def done(self, res, path, xlsx, err):
        self.progress.stop()
        self.progress.grid_remove()
        self.btn.state(["!disabled"])
        if err:
            e, tb = err
            self.set_status(_("Chyba: {e}", e=e), "#b3261e")
            messagebox.showerror("SHEETDIFF", _("Porovnání selhalo:\n{e}", e=e) + f"\n\n{tb[-1500:]}",
                                 parent=self.root)
            return
        text = res.headline
        if res.facts:
            text += "\n" + " · ".join(res.facts)
        self.set_status(text, TONES.get(res.exit_code, ""))
        self.last_report = path
        self.open_btn.grid(row=0, column=1, padx=12)
        os.startfile(path)
        if xlsx:
            os.startfile(xlsx)

    def reopen(self):
        if self.last_report and os.path.isfile(self.last_report):
            os.startfile(self.last_report)


def main(argv=None):
    args = [a for a in (argv if argv is not None else sys.argv[1:]) if not a.startswith("-")]
    files = [os.path.abspath(f) for f in args][:2]
    if len(files) == 2:
        files = order_by_age(files)
    settings = load_settings()
    set_lang(settings.get("lang") or system_lang())
    try:  # ostré písmo na monitorech se zvětšením
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = App(root, files, settings)
    if len(files) == 2:   # dva soubory přetažené na .exe / .bat: rovnou porovnat
        root.after(300, app.start)
    root.mainloop()


if __name__ == "__main__":
    main()
