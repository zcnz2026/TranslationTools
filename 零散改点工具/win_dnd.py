"""Native Windows Explorer file & folder drag-and-drop for Tkinter (100% stdlib ctypes).

Hooks WM_DROPFILES on the top-level HWND using 64-bit-safe GetWindowLongPtrW / SetWindowLongPtrW /
CallWindowProcW, extracts full UTF-16 Unicode paths via DragQueryFileW, and enqueues drop events
into a lock-free SimpleQueue polled by Tkinter's normal event loop (strictly avoiding any _tkinter
calls inside the raw Win32 WndProc callback, which would corrupt _tkinter's GIL thread state).
"""

import ctypes
from pathlib import Path
import queue
import sys
import tkinter as tk
from typing import Callable, List, Optional, Tuple

if sys.platform == "win32":
    from ctypes import wintypes
else:
    wintypes = None  # type: ignore[assignment]


def create_tk_root() -> tk.Tk:
    """Create a Tk root window, enabling native Finder/X11 drag-and-drop via tkinterdnd2 when available."""
    if sys.platform != "win32":
        try:
            from tkinterdnd2 import TkinterDnD  # type: ignore[import-not-found]

            return TkinterDnD.Tk()
        except Exception:
            pass
    return tk.Tk()


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
        self._drop_queue: "queue.SimpleQueue[Tuple[List[str], int, int]]" = queue.SimpleQueue()
        self._hwnd: Optional[int] = None
        self._old_wndproc: Optional[int] = None
        self._new_wndproc_ref = None
        self._hooked = False
        self._after_id: Optional[str] = None

        if sys.platform == "win32":
            self._init_win32_apis()

    def _init_win32_apis(self):
        self._user32 = ctypes.windll.user32
        self._shell32 = ctypes.windll.shell32

        self._user32.GetParent.argtypes = [wintypes.HWND]
        self._user32.GetParent.restype = wintypes.HWND

        if ctypes.sizeof(ctypes.c_void_p) == 8 and hasattr(self._user32, "SetWindowLongPtrW"):
            self._get_wnd_long_ptr = self._user32.GetWindowLongPtrW
            self._set_wnd_long_ptr = self._user32.SetWindowLongPtrW
        else:
            self._get_wnd_long_ptr = self._user32.GetWindowLongW
            self._set_wnd_long_ptr = self._user32.SetWindowLongW

        self._get_wnd_long_ptr.argtypes = [wintypes.HWND, ctypes.c_int]
        self._get_wnd_long_ptr.restype = ctypes.c_void_p
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

        self._user32.DefWindowProcW.argtypes = [
            wintypes.HWND,
            wintypes.UINT,
            wintypes.WPARAM,
            wintypes.LPARAM,
        ]
        self._user32.DefWindowProcW.restype = LRESULT

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
        """Install native file drop hook on the top-level window. Call before root.mainloop()."""
        if self._hooked:
            return True
        if sys.platform != "win32":
            if hasattr(self.root, "drop_target_register") and hasattr(self.root, "dnd_bind"):
                try:
                    self.root.drop_target_register("DND_Files")  # type: ignore[attr-defined]
                    self.root.dnd_bind("<<Drop>>", self._on_tkdnd_drop)  # type: ignore[attr-defined]
                    self._hooked = True
                    return True
                except Exception:
                    return False
            return False

        self.root.update_idletasks()
        child_hwnd = self.root.winfo_id()
        top_hwnd = self._user32.GetParent(child_hwnd) or child_hwnd
        if not top_hwnd:
            return False

        orig_ptr = self._get_wnd_long_ptr(top_hwnd, GWLP_WNDPROC)
        if not orig_ptr:
            return False

        if self._change_msg_filter is not None:
            for msg_id in (WM_DROPFILES, WM_COPYDATA, WM_COPYGLOBALDATA):
                try:
                    self._change_msg_filter(top_hwnd, msg_id, MSGFLT_ALLOW, None)
                except Exception:
                    pass

        self._hwnd = top_hwnd
        self._old_wndproc = orig_ptr

        # IMPORTANT: Never call any self.root.* (_tkinter) method inside _wnd_proc!
        # _tkinter releases the GIL in Tcl_DoOneEvent while holding its internal Tcl state;
        # calling back into _tkinter from a raw ctypes WNDPROC corrupts _tkinter's saved
        # PyThreadState (causing Fatal Python error: PyEval_RestoreThread -> 0xc0000409).
        def _wnd_proc(hwnd, msg, wparam, lparam):
            if msg == WM_DROPFILES:
                try:
                    self._enqueue_wm_dropfiles(hwnd, wparam)
                except Exception:
                    pass
                return 0
            old_p = self._old_wndproc
            if old_p:
                return self._user32.CallWindowProcW(old_p, hwnd, msg, wparam, lparam)
            return self._user32.DefWindowProcW(hwnd, msg, wparam, lparam)

        self._new_wndproc_ref = WNDPROC(_wnd_proc)
        proc_ptr = ctypes.cast(self._new_wndproc_ref, ctypes.c_void_p).value
        prev_ptr = self._set_wnd_long_ptr(top_hwnd, GWLP_WNDPROC, proc_ptr)
        if not prev_ptr:
            self._hwnd = None
            self._old_wndproc = None
            self._new_wndproc_ref = None
            return False

        self._old_wndproc = prev_ptr
        self._shell32.DragAcceptFiles(top_hwnd, True)
        self._hooked = True
        self._schedule_poll()
        return True

    def _schedule_poll(self):
        if not self._hooked:
            return
        try:
            if self.root.winfo_exists():
                self._after_id = self.root.after(60, self._poll_drop_queue)
        except Exception:
            pass

    def _poll_drop_queue(self):
        self._after_id = None
        if not self._hooked:
            return
        try:
            while not self._drop_queue.empty():
                paths, screen_x, screen_y = self._drop_queue.get_nowait()
                try:
                    self.dispatch_drop(paths, screen_x, screen_y)
                except Exception:
                    pass
        finally:
            self._schedule_poll()

    def unhook(self):
        """Restore original WndProc cleanly before window destruction."""
        if not self._hooked:
            return
        if sys.platform != "win32":
            self._hooked = False
            return
        self._hooked = False
        if self._after_id is not None:
            try:
                self.root.after_cancel(self._after_id)
            except Exception:
                pass
            self._after_id = None
        try:
            if self._hwnd:
                self._shell32.DragAcceptFiles(self._hwnd, False)
                if self._old_wndproc:
                    self._set_wnd_long_ptr(self._hwnd, GWLP_WNDPROC, self._old_wndproc)
        except Exception:
            pass
        finally:
            self._hwnd = None
            self._old_wndproc = None
            # Keep _new_wndproc_ref alive briefly in case an in-flight message unwinds

    def _enqueue_wm_dropfiles(self, hwnd: int, hdrop: int):
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
            self._drop_queue.put((paths, int(pt.x), int(pt.y)))

    def _on_tkdnd_drop(self, event):
        try:
            raw_data = getattr(event, "data", "") or ""
            if not raw_data:
                return
            raw_paths = [str(x) for x in self.root.tk.splitlist(raw_data) if str(x).strip()]
            sx = int(getattr(event, "x_root", 0) or 0)
            sy = int(getattr(event, "y_root", 0) or 0)
            self.dispatch_drop(raw_paths, sx, sy)
        except Exception:
            pass

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
