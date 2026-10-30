"""Application settings loaded from .env (next to the .exe, or project root in dev)."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import os
import shutil
import sys

from dotenv import load_dotenv

if getattr(sys, "frozen", False):  # PyInstaller .exe: user files live next to the executable
    ROOT = Path(sys.executable).resolve().parent
    BUNDLE = Path(getattr(sys, "_MEIPASS", ROOT))
else:
    ROOT = Path(__file__).resolve().parent.parent
    BUNDLE = ROOT

ENV_FILE = ROOT / ".env"
ENV_CREATED = False
if not ENV_FILE.exists() and (BUNDLE / ".env.example").exists():
    shutil.copy(BUNDLE / ".env.example", ENV_FILE)
    ENV_CREATED = True
load_dotenv(ENV_FILE)


def _env(key: str, default: str = "") -> str:
    return os.getenv(key, default)


@dataclass(frozen=True)
class Settings:
    api_base_url: str = _env("API_BASE_URL", "https://apis-legalizacion-silogtran.onrender.com").rstrip("/")
    api_key: str = _env("API_KEY")
    api_timeout: int = int(_env("API_TIMEOUT_SECONDS", "300"))

    gl_travel: str = _env("GL_TRAVEL_EXPENSES", "6130105595")
    gl_tolls: str = _env("GL_CASH_TOLLS", "6130105530")
    pospre_travel: str = _env("POSPRE_TRAVEL_EXPENSES", "61301055")
    pospre_tolls: str = _env("POSPRE_CASH_TOLLS", "6130105530")

    sap_indicator: str = _env("SAP_INDICATOR", "V7")
    sap_fixed_account: str = _env("SAP_FIXED_ACCOUNT", "222222222")

    drivers_file: Path = ROOT / _env("DRIVERS_FILE", "Asignación Conductores.xlsx")
    # Sheet name has a trailing space in the source workbook.
    drivers_sheet: str = _env("DRIVERS_SHEET", "CENTROS DE COSTOS FLOTA ")

    output_dir: Path = ROOT / _env("OUTPUT_DIR", "output")

    gr55_group: str = _env("GR55_REPORT_GROUP", "ZFM1")
    gr55_cc_low: str = _env("GR55_CC_LOW", "25000000")
    gr55_cc_high: str = _env("GR55_CC_HIGH", "29000000")


@dataclass(frozen=True)
class SapConnection:
    env: str
    connection: str
    client: str
    user: str
    password: str
    language: str


def sap_connection(env: str | None = None) -> SapConnection:
    """SAP Logon credentials for QAS or PRD (defaults to SAP_ENV)."""
    env = (env or _env("SAP_ENV", "QAS")).strip().upper()
    if env not in {"QAS", "PRD"}:
        raise ValueError(f"SAP_ENV must be QAS or PRD, got '{env}'.")
    conn = SapConnection(
        env=env,
        connection=_env(f"SAP_{env}_CONNECTION"),
        client=_env(f"SAP_{env}_CLIENT"),
        user=_env(f"SAP_{env}_USER"),
        password=_env(f"SAP_{env}_PASSWORD"),
        language=_env("SAP_LANGUAGE", "ES"),
    )
    missing = [k for k in ("connection", "client", "user") if not getattr(conn, k)]
    if missing:
        raise ValueError(f"Missing SAP_{env}_* values in .env: {', '.join(m.upper() for m in missing)}")
    return conn


settings = Settings()
