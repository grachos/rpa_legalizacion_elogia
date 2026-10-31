"""Auto-accepts SAP GUI security prompts for SAP's own temp files (Office integration of GR55).

Only dialogs titled 'Seguridad SAP GUI' (or 'SAP GUI Security') are handled, and only while the
GR55 run is active. Uses plain Win32 (pywin32) first, which works with SAP's classic dialogs;
UIA (pywinauto) is a fallback to read text / click. Events and errors are written to
output/security_watcher.log to diagnose problems on other machines.
"""
from __future__ import annotations

import threading
import time
from datetime import datetime

TITLES = ("seguridad sap gui", "sap gui security")
ALLOW = ("permitir", "allow")


def _log(msg: str) -> None:
    try:
        from .config import settings

        settings.output_dir.mkdir(parents=True, exist_ok=True)
        with open(settings.output_dir / "security_watcher.log", "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {msg}\n")
    except Exception:  # noqa: BLE001 - logging must never break the run
        pass


class SecurityPromptWatcher:
    def __init__(self, interval: float = 0.4):
        self.interval = interval
        self.accepted = 0
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._tries: dict[int, int] = {}

    def __enter__(self) -> "SecurityPromptWatcher":
        self._thread.start()
        return self

    def __exit__(self, *exc) -> None:
        self._stop.set()
        self._thread.join(timeout=5)
        _log(f"watcher stopped, accepted={self.accepted}")

    # --- Win32 helpers -------------------------------------------------
    @staticmethod
    def _top_windows() -> list[int]:
        import win32gui

        found: list[int] = []

        def cb(h, _):
            if win32gui.IsWindowVisible(h) and win32gui.GetWindowText(h).strip().lower() in TITLES:
                found.append(h)
            return True

        win32gui.EnumWindows(cb, None)
        return found

    @staticmethod
    def _children(hwnd: int) -> list[tuple[int, str, str]]:
        import win32gui

        out: list[tuple[int, str, str]] = []

        def cb(h, _):
            out.append((h, win32gui.GetClassName(h), win32gui.GetWindowText(h)))
            return True

        win32gui.EnumChildWindows(hwnd, cb, None)
        return out

    @staticmethod
    def _uia_texts(hwnd: int) -> list[str]:
        from pywinauto.controls.uiawrapper import UIAWrapper
        from pywinauto.uia_element_info import UIAElementInfo

        w = UIAWrapper(UIAElementInfo(handle_or_elem=hwnd))
        return [t.window_text() for t in w.descendants() if t.window_text()]

    def _handle(self, hwnd: int) -> None:
        import win32con
        import win32gui

        kids = self._children(hwnd)
        texts = [t for _, _, t in kids if t]
        if not any("~sap{" in t.lower() for t in texts):
            try:
                texts += self._uia_texts(hwnd)
            except Exception as e:  # noqa: BLE001
                _log(f"UIA text read failed: {e}")
        joined = " ".join(texts).lower()
        if joined.strip() and "~sap{" not in joined:
            return  # a different prompt: leave it for the user
        if not joined.strip():
            _log("dialog text unreadable; accepting (only active during GR55)")

        n = self._tries.get(hwnd, 0)
        self._tries[hwnd] = n + 1
        btn = next((h for h, cls, t in kids
                    if t.replace("&", "").strip().lower() in ALLOW and "button" in cls.lower()), None)
        if btn and n < 5:
            win32gui.PostMessage(btn, win32con.BM_CLICK, 0, 0)
            _log(f"clicked Permitir via Win32 (hwnd={hwnd}, try={n + 1})")
        else:  # Win32 click didn't close it (or no button found): UIA invoke, then Alt+P
            try:
                from pywinauto.controls.uiawrapper import UIAWrapper
                from pywinauto.uia_element_info import UIAElementInfo

                w = UIAWrapper(UIAElementInfo(handle_or_elem=hwnd))
                for name in ("Permitir", "Allow"):
                    b = w.descendants(title=name, control_type="Button")
                    if b:
                        b[0].invoke()
                        _log(f"clicked Permitir via UIA (hwnd={hwnd})")
                        break
                else:
                    win32gui.SetForegroundWindow(hwnd)
                    from pywinauto.keyboard import send_keys

                    send_keys("%p")
                    _log(f"sent Alt+P (hwnd={hwnd})")
            except Exception as e:  # noqa: BLE001
                _log(f"fallback click failed: {e}")
        self.accepted += 1
        time.sleep(0.3)

    def _loop(self) -> None:
        try:
            import pythoncom

            pythoncom.CoInitialize()  # UIA fallback needs COM in this thread
        except Exception as e:  # noqa: BLE001
            _log(f"CoInitialize failed: {e}")
        _log("watcher started")
        while not self._stop.is_set():
            try:
                for hwnd in self._top_windows():
                    self._handle(hwnd)
            except Exception as e:  # noqa: BLE001 - window may vanish mid-scan; retry next tick
                _log(f"scan error: {e}")
            self._stop.wait(self.interval)
