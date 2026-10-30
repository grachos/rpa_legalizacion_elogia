"""RPA: available budget per cost center from GR55 (report group ZFM1, Report Writer -> Excel)."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re
import time

import pandas as pd

from .config import settings
from .sap_gui import SapError, SapSession
from .sap_security import SecurityPromptWatcher

CC_ROW = re.compile(r"^\s*\*\s+(\d{8})\b")
SHEETS = {"travel": "GASTOS VARIOS", "tolls": "PEAJES"}


def parse_report_rows(rows) -> dict[str, float]:
    """Rows of the GR55 sheet -> {cost_center: available}. Uses '*  CC  name' rows, column F (Disponible)."""
    out: dict[str, float] = {}
    for row in rows:
        if not row or len(row) < 5:  # UsedRange may start at column A or B
            continue
        label = next((c for c in row if isinstance(c, str) and c.strip()), "")
        m = CC_ROW.match(str(label or ""))
        if not m:
            continue
        value = row[-1]
        if isinstance(value, str):  # e.g. "1.234.567-"
            neg = value.strip().endswith("-")
            value = float(value.replace(".", "").replace(",", ".").rstrip("- ") or 0) * (-1 if neg else 1)
        out[m.group(1)] = out.get(m.group(1), 0.0) + float(value or 0)
    return out


def _excel():
    import win32com.client

    try:
        return win32com.client.GetActiveObject("Excel.Application")
    except Exception:  # noqa: BLE001 - Excel not open yet
        return None


def _sap_workbooks(xl) -> set[str]:
    if xl is None:
        return set()
    return {wb.FullName for wb in xl.Workbooks if "~SAP{" in wb.FullName}


def _wait_report_workbook(before: set[str], timeout: int = 90):
    """Waits for the workbook SAP opens through Office integration and returns it."""
    end = time.time() + timeout
    while time.time() < end:
        xl = _excel()
        if xl is not None:
            for wb in xl.Workbooks:
                if "~SAP{" in wb.FullName and wb.FullName not in before:
                    return wb
        time.sleep(1)
    raise SapError("SAP no abrió el Excel del informe en el tiempo esperado. Revisa si quedó "
                   "alguna ventana de SAP pendiente y vuelve a intentar.")


def _run_one(sap: SapSession, pospre: str) -> dict[str, float]:
    s = sap.session
    sap.transaction("GR55")
    s.findById("wnd[0]/usr/ctxtRGRWJ-JOB").text = settings.gr55_group
    s.findById("wnd[0]").sendVKey(8)
    sap.raise_on_error("Grupo de informes")
    s.findById("wnd[0]/usr/ctxt_4FFICTR-LOW").text = settings.gr55_cc_low
    s.findById("wnd[0]/usr/ctxt_4FFICTR-HIGH").text = settings.gr55_cc_high
    s.findById("wnd[0]/usr/ctxt_4FFIPEX-LOW").text = pospre
    s.findById("wnd[0]/usr/ctxt_4FFIPEX-HIGH").text = pospre
    before = _sap_workbooks(_excel())
    s.findById("wnd[0]").sendVKey(8)
    wb = _wait_report_workbook(before)
    try:
        ws = wb.Worksheets(1)
        if str(ws.Name).upper() == "NO_DATA":
            return {}
        rows = ws.UsedRange.Value or ()
        return parse_report_rows(rows)
    finally:
        wb.Close(SaveChanges=False)
        s.findById("wnd[0]").sendVKey(3)  # back to selection


def fetch_balances(env: str | None = None) -> dict[str, dict[str, float]]:
    """Runs GR55 for travel expenses and cash tolls. Returns {'travel': {...}, 'tolls': {...}}."""
    import pythoncom

    pythoncom.CoInitialize()  # needed when called from a UI thread
    try:
        with SapSession(env) as sap, SecurityPromptWatcher():
            return {
                "travel": _run_one(sap, settings.pospre_travel),
                "tolls": _run_one(sap, settings.pospre_tolls),
            }
    finally:
        pythoncom.CoUninitialize()


def save_balances(balances: dict[str, dict[str, float]], env: str) -> Path:
    """Saves balances in the same layout as the manual file (sheets GASTOS VARIOS / PEAJES)."""
    path = settings.output_dir / f"saldos_sap_{env}_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for key, sheet in SHEETS.items():
            df = pd.DataFrame(sorted(balances[key].items()), columns=["Centro de costo", "Disponible"])
            df.to_excel(xw, sheet_name=sheet, index=False)
    return path
