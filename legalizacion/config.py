"""Application settings stored in .env (next to the .exe, or project root in dev).

FIELDS is the single source of truth: it drives defaults, the .env file layout and the
settings dialog in the UI. `settings.reload()` re-reads values after the dialog saves.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import os
import sys

from dotenv import dotenv_values

if getattr(sys, "frozen", False):  # PyInstaller .exe: user files live next to the executable
    ROOT = Path(sys.executable).resolve().parent
else:
    ROOT = Path(__file__).resolve().parent.parent

ENV_FILE = ROOT / ".env"


@dataclass(frozen=True)
class Field:
    key: str
    label: str
    default: str = ""
    secret: bool = False
    kind: str = "text"  # text | file


# (section, fields) in display / file order
FIELDS: list[tuple[str, list[Field]]] = [
    ("API Silogtran", [
        Field("API_BASE_URL", "URL de la API", "https://apis-legalizacion-silogtran.onrender.com"),
        Field("API_KEY", "API key", secret=True),
        Field("API_TIMEOUT_SECONDS", "Tiempo máximo (segundos)", "300"),
    ]),
    ("SAP Productivo (PRD)", [
        Field("SAP_PRD_CONNECTION", "Entrada en SAP Logon", "Productivo"),
        Field("SAP_PRD_CLIENT", "Mandante", "300"),
        Field("SAP_PRD_USER", "Usuario"),
        Field("SAP_PRD_PASSWORD", "Contraseña (opcional)", secret=True),
    ]),
    ("SAP Calidad (QAS)", [
        Field("SAP_QAS_CONNECTION", "Entrada en SAP Logon", "Calidad"),
        Field("SAP_QAS_CLIENT", "Mandante", "300"),
        Field("SAP_QAS_USER", "Usuario"),
        Field("SAP_QAS_PASSWORD", "Contraseña (opcional)", secret=True),
    ]),
    ("Archivos", [
        Field("DRIVERS_FILE", "Asignación Conductores (.xlsx)", "Asignación Conductores.xlsx", kind="file"),
        Field("DRIVERS_SHEET", "Hoja de centros de costo", "CENTROS DE COSTOS FLOTA "),
        Field("OUTPUT_DIR", "Carpeta de resultados", "output"),
    ]),
    ("Cuentas y GR55", [
        Field("GL_TRAVEL_EXPENSES", "Cuenta FV50 gastos de viaje", "6130105595"),
        Field("GL_CASH_TOLLS", "Cuenta FV50 peajes", "6130105530"),
        Field("POSPRE_TRAVEL_EXPENSES", "Pospre gastos de viaje", "61301055"),
        Field("POSPRE_CASH_TOLLS", "Pospre peajes", "6130105530"),
        Field("SAP_INDICATOR", "Indicador", "V7"),
        Field("SAP_FIXED_ACCOUNT", "Cuenta fija", "222222222"),
        Field("GR55_REPORT_GROUP", "Grupo de informes GR55", "ZFM1"),
        Field("GR55_CC_LOW", "Centro gestor desde", "25000000"),
        Field("GR55_CC_HIGH", "Centro gestor hasta", "29000000"),
        Field("SAP_LANGUAGE", "Idioma SAP", "ES"),
    ]),
]
ALL_FIELDS = {f.key: f for _, fs in FIELDS for f in fs}


def read_env() -> dict[str, str]:
    """Current values (file over defaults)."""
    values = {k: f.default for k, f in ALL_FIELDS.items()}
    if ENV_FILE.exists():
        values.update({k: (v or "") for k, v in dotenv_values(ENV_FILE).items() if k in ALL_FIELDS})
    return values


def _quote(v: str) -> str:
    return f'"{v}"' if (v != v.strip() or any(c in v for c in " #'\"")) else v


def write_env(values: dict[str, str]) -> None:
    lines = ["# RPA Legalización Elogia - configuración (editable desde la app: Configuración)"]
    for section, fields in FIELDS:
        lines += ["", f"# {section}"]
        lines += [f"{f.key}={_quote(values.get(f.key, f.default).replace(chr(34), ''))}" for f in fields]
    ENV_FILE.write_text("\n".join(lines) + "\n", encoding="utf-8")
    settings.reload()


def is_configured() -> bool:
    v = read_env()
    return bool(v["API_KEY"] and v["API_KEY"] != "change-me" and v["SAP_PRD_USER"])


class Settings:
    def __init__(self) -> None:
        self.reload()

    def reload(self) -> None:
        v = read_env()
        self.raw = v
        self.api_base_url = v["API_BASE_URL"].rstrip("/")
        self.api_key = v["API_KEY"]
        self.api_timeout = int(v["API_TIMEOUT_SECONDS"] or 300)
        self.gl_travel, self.gl_tolls = v["GL_TRAVEL_EXPENSES"], v["GL_CASH_TOLLS"]
        self.pospre_travel, self.pospre_tolls = v["POSPRE_TRAVEL_EXPENSES"], v["POSPRE_CASH_TOLLS"]
        self.sap_indicator, self.sap_fixed_account = v["SAP_INDICATOR"], v["SAP_FIXED_ACCOUNT"]
        self.drivers_file = ROOT / v["DRIVERS_FILE"]  # absolute paths stay absolute
        self.drivers_sheet = v["DRIVERS_SHEET"]
        self.output_dir = ROOT / (v["OUTPUT_DIR"] or "output")
        self.gr55_group, self.gr55_cc_low, self.gr55_cc_high = (
            v["GR55_REPORT_GROUP"], v["GR55_CC_LOW"], v["GR55_CC_HIGH"])
        self.sap_language = v["SAP_LANGUAGE"] or "ES"


@dataclass(frozen=True)
class SapConnection:
    env: str
    connection: str
    client: str
    user: str
    password: str
    language: str


def sap_connection(env: str | None = None) -> SapConnection:
    """SAP Logon credentials for PRD (default) or QAS."""
    env = (env or "PRD").strip().upper()
    if env not in {"QAS", "PRD"}:
        raise ValueError(f"Ambiente SAP inválido: '{env}'.")
    v = settings.raw
    conn = SapConnection(env, v[f"SAP_{env}_CONNECTION"], v[f"SAP_{env}_CLIENT"], v[f"SAP_{env}_USER"],
                         v[f"SAP_{env}_PASSWORD"], settings.sap_language)
    missing = [k for k in ("connection", "client", "user") if not getattr(conn, k)]
    if missing:
        raise ValueError(f"Falta configurar SAP {env}: {', '.join(missing)} (menú Configuración).")
    return conn


settings = Settings()
