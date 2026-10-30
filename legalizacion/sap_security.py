"""Auto-accepts SAP GUI security prompts for SAP's own temp files (Office integration of GR55).

Only dialogs titled 'Seguridad SAP GUI' (or 'SAP GUI Security') whose text points to
%TEMP%\\~SAP{...} are accepted; any other prompt is left for the user.
"""
from __future__ import annotations

import os
import threading

TITLES = ("Seguridad SAP GUI", "SAP GUI Security")
ALLOW = ("Permitir", "Allow")


class SecurityPromptWatcher:
    def __init__(self, interval: float = 0.5):
        self.interval = interval
        self.accepted = 0
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._temp = os.path.normcase(os.environ.get("TEMP", ""))

    def __enter__(self) -> "SecurityPromptWatcher":
        self._thread.start()
        return self

    def __exit__(self, *exc) -> None:
        self._stop.set()
        self._thread.join(timeout=5)

    def _is_sap_temp_prompt(self, dlg) -> bool:
        text = " ".join(t.window_text() for t in dlg.descendants() if t.window_text())
        norm = os.path.normcase(text)
        return "~sap{" in norm and (not self._temp or self._temp in norm or "\\temp\\" in norm)

    def _loop(self) -> None:
        import comtypes  # pywinauto UIA backend needs COM initialized in this thread
        from pywinauto import Desktop

        comtypes.CoInitialize()
        try:
            desk = Desktop(backend="uia")
            while not self._stop.is_set():
                try:
                    for title in TITLES:
                        for dlg in desk.windows(title=title, visible_only=True):
                            if not self._is_sap_temp_prompt(dlg):
                                continue
                            for name in ALLOW:
                                btns = dlg.descendants(title=name, control_type="Button")
                                if btns:
                                    btns[0].invoke()
                                    self.accepted += 1
                                    break
                except Exception:  # noqa: BLE001 - window may vanish mid-scan; retry next tick
                    pass
                self._stop.wait(self.interval)
        finally:
            comtypes.CoUninitialize()
