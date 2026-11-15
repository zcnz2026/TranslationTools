"""Native Windows Explorer file & folder drag-and-drop for Tkinter (100% stdlib ctypes).

Hooks WM_DROPFILES on the top-level HWND using 64-bit-safe SetWindowLongPtrW / CallWindowProcW,
extracts full UTF-16 Unicode paths via DragQueryFileW, resolves the exact drop target widget via
DragQueryPoint + ClientToScreen + root.winfo_containing, and cleanly unhooks on window close.
"""

import ctypes
from ctypes import wintypes
from pathlib import Path
import sys
import tkinter as tk
from typing import Callable, List, Optional, Tuple


WM_DROPFILES = 0x0233
WM_COPYDATA = 0x004A
WM_COPYGLOBALDATA = 0x0049
MSGFLT_ALLOW = 1
GWLP_WNDPROC = -4

LRESULT = ctypes.c_ssize_t
if sys.platform == "win32":
    WNDPROC = ctypes.WINFUNCTYPE(
        LRESULT,
        wintypes.HWND,
        wintypes.UINT,
        wintypes.WPARAM,
        wintypes.LPARAM,
    )
else:
    WNDPROC = None


class WinFileDropManager:
    """Manages native Windows file/folder drag-and-drop across a Tkinter root window."""

    def __init__(
        self,
        root: tk.Tk,
        fallback_cb: Optional[Callable[[List[Path]], None]] = None,
    ):
        self.root = root
        self.fallback_cb = fallback_cb
        self._targets: List[Tuple[tk.Misc, Callable[[List[Path]], None]]] = []
        self._hwnd: Optional[int] = None
        self._old_wndproc: Optional[int] = None
        self._new_wndproc_ref = None
        self._hooked = False

        if sys.platform == "win32":
            self._init_win32_apis()

    def _init_win32_apis(self):
        self._user32 = ctypes.windll.user32
        self._shell32 = ctypes.windll.shell32

        self._user32.GetParent.argtypes = [wintypes.HWND]
        self._user32.GetParent.restype = wintypes.HWND

        if ctypes.sizeof(ctypes.c_void_p) == 8 and hasattr(self._user32, "SetWindowLongPtrW"):
            self._set_wnd_long_ptr = self._user32.SetWindowLongPtrW
        else:
            self._set_wnd_long_ptr = self._user32.SetWindowLongW
        self._set_wnd_long_ptr.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]
        self._set_wnd_long_ptr.restype = ctypes.c_void_p

        self._user32.CallWindowProcW.argtypes = [
            ctypes.c_void_p,
            wintypes.HWND,
            wintypes.UINT,
            wintypes.WPARAM,
            wintypes.LPARAM,
        ]
        self._user32.CallWindowProcW.restype = LRESULT

        self._user32.ClientToScreen.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.POINT)]
        self._user32.ClientToScreen.restype = wintypes.BOOL

        self._shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
        self._shell32.DragAcceptFiles.restype = None

        self._shell32.DragQueryFileW.argtypes = [
            ctypes.c_void_p,
            wintypes.UINT,
            wintypes.LPWSTR,
            wintypes.UINT,
        ]
        self._shell32.DragQueryFileW.restype = wintypes.UINT

        self._shell32.DragQueryPoint.argtypes = [ctypes.c_void_p, ctypes.POINTER(wintypes.POINT)]
        self._shell32.DragQueryPoint.restype = wintypes.BOOL

        self._shell32.DragFinish.argtypes = [ctypes.c_void_p]
        self._shell32.DragFinish.restype = None

        self._change_msg_filter = getattr(self._user32, "ChangeWindowMessageFilterEx", None)
        if self._change_msg_filter is not None:
            self._change_msg_filter.argtypes = [
                wintypes.HWND,
                wintypes.UINT,
                wintypes.DWORD,
                ctypes.c_void_p,
            ]
            self._change_msg_filter.restype = wintypes.BOOL

    def register_target(self, widget: tk.Misc, callback: Callable[[List[Path]], None]):
        """Register a widget (and all its visual children) as a specific drop zone."""
        self._targets.append((widget, callback))

    def install(self) -> bool:
        """Install WM_DROPFILES hook on the top-level window. Call after widgets are created."""
        if sys.platform != "win32" or self._hooked:
            return self._hooked

        self.root.update_idletasks()
        child_hwnd = self.root.winfo_id()
        top_hwnd = self._user32.GetParent(child_hwnd) or child_hwnd
        if not top_hwnd:
            return False

        if self._change_msg_filter is not None:
            for msg_id in (WM_DROPFILES, WM_COPYDATA, WM_COPYGLOBALDATA):
                try:
                    self._change_msg_filter(top_hwnd, msg_id, MSGFLT_ALLOW, None)
                except Exception:
                    pass

        self._shell32.DragAcceptFiles(top_hwnd, True)

        def _wnd_proc(hwnd, msg, wparam, lparam):
            if msg == WM_DROPFILES:
                try:
                    self._handle_wm_dropfiles(hwnd, wparam)
                except Exception:
                    pass
                return 0
            return self._user32.CallWindowProcW(self._old_wndproc, hwnd, msg, wparam, lparam)

        self._new_wndproc_ref = WNDPROC(_wnd_proc)
        proc_ptr = ctypes.cast(self._new_wndproc_ref, ctypes.c_void_p).value
        prev_ptr = self._set_wnd_long_ptr(top_hwnd, GWLP_WNDPROC, proc_ptr)
        if not prev_ptr:
            self._shell32.DragAcceptFiles(top_hwnd, False)
            self._new_wndproc_ref = None
            return False

        self._hwnd = top_hwnd
        self._old_wndproc = prev_ptr
        self._hooked = True
        return True

    def unhook(self):
        """Restore original WndProc cleanly before window destruction."""
        if not self._hooked or sys.platform != "win32":
            return
        try:
            if self._hwnd:
                self._shell32.DragAcceptFiles(self._hwnd, False)
                if self._old_wndproc:
                    self._set_wnd_long_ptr(self._hwnd, GWLP_WNDPROC, self._old_wndproc)
        except Exception:
            pass
        finally:
            self._hooked = False
            self._hwnd = None
            self._old_wndproc = None
            self._new_wndproc_ref = None

    def _handle_wm_dropfiles(self, hwnd: int, hdrop: int):
        pt = wintypes.POINT(0, 0)
        paths: List[str] = []
        try:
            self._shell32.DragQueryPoint(hdrop, ctypes.byref(pt))
            self._user32.ClientToScreen(hwnd, ctypes.byref(pt))

            count = self._shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
            for idx in range(count):
                char_len = self._shell32.DragQueryFileW(hdrop, idx, None, 0)
                if char_len <= 0:
                    continue
                buf = ctypes.create_unicode_buffer(char_len + 2)
                self._shell32.DragQueryFileW(hdrop, idx, buf, char_len + 2)
                if buf.value:
                    paths.append(buf.value)
        finally:
            self._shell32.DragFinish(hdrop)

        if paths:
            screen_x, screen_y = int(pt.x), int(pt.y)
            self.root.after(0, lambda: self.dispatch_drop(paths, screen_x, screen_y))

    def resolve_callback_for_widget(
        self, widget: Optional[tk.Misc]
    ) -> Optional[Callable[[List[Path]], None]]:
        """Find the most specific registered callback owning `widget` (walking up `.master`)."""
        if widget is None:
            return self.fallback_cb

        reg_map = {w: cb for w, cb in self._targets}
        cur: Optional[tk.Misc] = widget
        visited = set()
        while cur is not None and id(cur) not in visited:
            visited.add(id(cur))
            if cur in reg_map:
                return reg_map[cur]
            cur = getattr(cur, "master", None)
        return self.fallback_cb

    def dispatch_drop(self, raw_paths: List[str], screen_x: int, screen_y: int):
        resolved_paths = []
        seen = set()
        for rp in raw_paths:
            s = str(rp).strip().strip('"')
            if not s:
                continue
            p = Path(s)
            key = str(p).lower()
            if key not in seen:
                seen.add(key)
                resolved_paths.append(p)

        if not resolved_paths:
            return

        hit_widget = None
        try:
            hit_widget = self.root.winfo_containing(screen_x, screen_y)
        except Exception:
            hit_widget = None

        cb = self.resolve_callback_for_widget(hit_widget)
        if cb is not None:
            cb(resolved_paths)
