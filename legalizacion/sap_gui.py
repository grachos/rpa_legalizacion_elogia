"""SAP GUI Scripting helpers: connect via SAP Logon, log in, run transactions (Windows only)."""
from __future__ import annotations

from pathlib import Path
import subprocess
import time

from .config import SapConnection, sap_connection

SAPLOGON_PATHS = [
    Path(r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe"),
    Path(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe"),
]


class SapError(RuntimeError):
    pass


def _scripting_engine(wait_seconds: int = 30):
    import win32com.client  # pywin32, imported lazily so the rest of the app runs without it

    try:
        return win32com.client.GetObject("SAPGUI").GetScriptingEngine
    except Exception:  # noqa: BLE001 - SAP Logon not running yet
        exe = next((p for p in SAPLOGON_PATHS if p.exists()), None)
        if not exe:
            raise SapError("SAP Logon no está abierto y no se encontró saplogon.exe.")
        subprocess.Popen([str(exe)])
        for _ in range(wait_seconds):
            time.sleep(1)
            try:
                return win32com.client.GetObject("SAPGUI").GetScriptingEngine
            except Exception:  # noqa: BLE001
                continue
    raise SapError("No fue posible conectarse a SAP GUI Scripting. ¿Está habilitado el scripting?")


class SapSession:
    """Context manager: opens (or reuses) a logged-in session for the configured environment."""

    def __init__(self, env: str | None = None):
        self.conf: SapConnection = sap_connection(env)
        self.connection = None
        self.session = None
        self._opened = False

    def __enter__(self) -> "SapSession":
        engine = _scripting_engine()
        self.session = self._find_existing(engine)
        if self.session is None:
            self.connection = engine.OpenConnection(self.conf.connection, True)
            self._opened = True
            self.session = self.connection.Children(0)
            self._login()
        return self

    def __exit__(self, *exc) -> None:
        if self._opened and self.connection is not None:
            try:
                self.connection.CloseConnection()
            except Exception:  # noqa: BLE001
                pass

    def _find_existing(self, engine):
        for i in range(engine.Children.Count):
            conn = engine.Children(i)
            if conn.Description.strip().upper() != self.conf.connection.strip().upper():
                continue
            for j in range(conn.Children.Count):
                s = conn.Children(j)
                if s.Info.Client == self.conf.client and s.Info.User.upper() == self.conf.user.upper() and not s.Busy:
                    self.connection = conn
                    return s
        return None

    def _login(self) -> None:
        if not self.conf.password:
            raise SapError(f"No hay sesión abierta en {self.conf.connection} y falta SAP_{self.conf.env}_PASSWORD en .env.")
        s = self.session
        s.findById("wnd[0]/usr/txtRSYST-MANDT").text = self.conf.client
        s.findById("wnd[0]/usr/txtRSYST-BNAME").text = self.conf.user
        s.findById("wnd[0]/usr/pwdRSYST-BCODE").text = self.conf.password
        s.findById("wnd[0]/usr/txtRSYST-LANGU").text = self.conf.language
        s.findById("wnd[0]").sendVKey(0)
        # Multiple logon popup: continue keeping the other sessions.
        if s.Children.Count > 1:
            try:
                s.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select()
                s.findById("wnd[1]/tbar[0]/btn[0]").press()
            except Exception:  # noqa: BLE001 - another popup (e.g. system message): just close it
                s.findById("wnd[1]").sendVKey(0)
        self.raise_on_error("Inicio de sesión")
        if s.Info.User == "":
            raise SapError("No se pudo iniciar sesión en SAP.")

    # ---------- helpers ----------
    def status(self) -> tuple[str, str]:
        bar = self.session.findById("wnd[0]/sbar")
        return bar.MessageType, bar.Text

    def raise_on_error(self, step: str) -> None:
        kind, text = self.status()
        if kind in ("E", "A"):
            raise SapError(f"{step}: {text}")

    def transaction(self, tcode: str) -> None:
        self.session.StartTransaction(tcode)
        self.raise_on_error(f"Transacción {tcode}")
