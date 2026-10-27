"""Přetažení souborů z Průzkumníka do okna Tk (Windows, bez dalších knihoven).

Okno se zaregistruje přes DragAcceptFiles a jeho procedura se obalí, aby zachytila
WM_DROPFILES. Ostatní zprávy jdou beze změny původní proceduře Tk.
"""
from __future__ import annotations

import ctypes
import sys
from ctypes import wintypes

WM_DROPFILES = 0x0233
GWLP_WNDPROC = -4
LRESULT = ctypes.c_ssize_t
WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)


def _api():
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    shell32 = ctypes.WinDLL("shell32")
    user32.SetWindowLongPtrW.restype = ctypes.c_void_p
    user32.SetWindowLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
    user32.CallWindowProcW.restype = LRESULT
    user32.CallWindowProcW.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT,
                                       wintypes.WPARAM, wintypes.LPARAM]
    user32.DefWindowProcW.restype = LRESULT
    user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user32.ClientToScreen.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.POINT)]
    shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
    shell32.DragQueryFileW.restype = wintypes.UINT
    shell32.DragQueryFileW.argtypes = [ctypes.c_void_p, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
    shell32.DragQueryPoint.argtypes = [ctypes.c_void_p, ctypes.POINTER(wintypes.POINT)]
    shell32.DragFinish.argtypes = [ctypes.c_void_p]
    return user32, shell32


def enable(widget, on_drop) -> bool:
    """on_drop(seznam_cest, (x, y) na obrazovce) se zavolá ve vlákně Tk. Vrací, zda se povedlo."""
    if sys.platform != "win32":
        return False
    try:
        user32, shell32 = _api()
        hwnd = widget.winfo_id()
        state = {"old": None}

        def proc(h, msg, wparam, lparam):
            if msg == WM_DROPFILES:
                try:
                    n = shell32.DragQueryFileW(wparam, 0xFFFFFFFF, None, 0)
                    files = []
                    for i in range(n):
                        size = shell32.DragQueryFileW(wparam, i, None, 0) + 1
                        buf = ctypes.create_unicode_buffer(size)
                        shell32.DragQueryFileW(wparam, i, buf, size)
                        files.append(buf.value)
                    pt = wintypes.POINT()
                    shell32.DragQueryPoint(wparam, ctypes.byref(pt))
                    user32.ClientToScreen(h, ctypes.byref(pt))
                finally:
                    shell32.DragFinish(wparam)
                widget.after(0, on_drop, files, (pt.x, pt.y))
                return 0
            if state["old"] is None:
                return user32.DefWindowProcW(h, msg, wparam, lparam)
            return user32.CallWindowProcW(state["old"], h, msg, wparam, lparam)

        new_proc = WNDPROC(proc)
        state["old"] = user32.SetWindowLongPtrW(hwnd, GWLP_WNDPROC,
                                                ctypes.cast(new_proc, ctypes.c_void_p))
        if not state["old"]:
            return False
        widget._drop_proc = new_proc   # bez reference by procedura zanikla
        shell32.DragAcceptFiles(hwnd, True)
        return True
    except Exception:
        return False


def simulate_drop(widget, files: list[str], client_xy=(0, 0)):
    """Pošle oknu WM_DROPFILES jako Průzkumník (pro testy)."""
    kernel32 = ctypes.WinDLL("kernel32")
    user32 = ctypes.WinDLL("user32")
    kernel32.GlobalAlloc.restype = ctypes.c_void_p
    kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    user32.SendMessageW.restype = LRESULT
    user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]

    class DROPFILES(ctypes.Structure):
        _fields_ = [("pFiles", wintypes.DWORD), ("pt", wintypes.POINT),
                    ("fNC", wintypes.BOOL), ("fWide", wintypes.BOOL)]

    names = ("\0".join(files) + "\0\0").encode("utf-16-le")
    head = DROPFILES(ctypes.sizeof(DROPFILES), wintypes.POINT(*client_xy), False, True)
    size = ctypes.sizeof(DROPFILES) + len(names)
    mem = kernel32.GlobalAlloc(0x0040, size)   # GMEM_FIXED | GMEM_ZEROINIT
    ctypes.memmove(mem, ctypes.byref(head), ctypes.sizeof(DROPFILES))
    ctypes.memmove(mem + ctypes.sizeof(DROPFILES), names, len(names))
    user32.SendMessageW(widget.winfo_id(), WM_DROPFILES, mem, 0)
