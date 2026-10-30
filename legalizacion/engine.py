"""Builds the legalization tables (BD, SAP, differences, CADI, registration)."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import unicodedata

import pandas as pd

from .config import settings

TOLL_CONCEPT = "PEAJES EN EFECTIVOS"
ELECTRONIC_TOLL = "PEAJE ELECTRONICO"
SAP_COLUMNS = ["cuenta_mayor", "texto", "dh", "valor", "saldo_sap", "diferencia",
               "col_h", "indicador", "col_j", "centro_costo", "cuenta_fija", "tercero"]
REGISTRO_COLUMNS = [c for c in SAP_COLUMNS if c not in ("saldo_sap", "diferencia")]
FV50_BATCH = 4


def _norm(text: str) -> str:
    text = unicodedata.normalize("NFKD", str(text)).encode("ascii", "ignore").decode()
    return " ".join(text.upper().split())


def _plate(text: str) -> str:
    return str(text).replace(" ", "").replace("-", "").upper().strip()


def _find_sheet(path: Path, wanted: str) -> str:
    sheets = pd.ExcelFile(path).sheet_names
    for s in sheets:
        if _norm(s) == _norm(wanted):
            return s
    raise ValueError(f"Sheet '{wanted}' not found in {path.name}. Sheets: {sheets}")


# ---------- Inputs ----------

def load_cost_centers(path: Path | None = None, sheet: str | None = None) -> dict[str, str]:
    """Plate -> cost center from the drivers assignment workbook."""
    path = Path(path or settings.drivers_file)
    df = pd.read_excel(path, sheet_name=_find_sheet(path, sheet or settings.drivers_sheet), dtype=str)
    df = df.iloc[:, :2].dropna()
    df.columns = ["placa", "cc"]
    return {_plate(p): c.strip().split(".")[0] for p, c in zip(df.placa, df.cc) if c.strip()}


def _parse_balance_table(df: pd.DataFrame) -> dict[str, float]:
    df = df.iloc[:, :2].copy()
    df.columns = ["cc", "disponible"]
    df["cc"] = df["cc"].astype(str).str.strip().str.split(".").str[0]
    df = df[df["cc"].str.fullmatch(r"\d{8}")]  # drops header, totals and blanks
    df["disponible"] = pd.to_numeric(df["disponible"], errors="coerce").fillna(0)
    return df.groupby("cc")["disponible"].sum().to_dict()


def load_balances_excel(path: Path) -> dict[str, dict[str, float]]:
    """Cleaned GR55 balances (sheets GASTOS VARIOS / PEAJES): {'travel': {cc: disp}, 'tolls': {...}}."""
    path = Path(path)
    out = {}
    for key, sheet in (("travel", "GASTOS VARIOS"), ("tolls", "PEAJES")):
        out[key] = _parse_balance_table(pd.read_excel(path, sheet_name=_find_sheet(path, sheet), header=None))
    return out


# ---------- Engine ----------

@dataclass
class AccountResult:
    name: str
    bd: pd.DataFrame
    sap: pd.DataFrame
    diferencias: pd.DataFrame
    cadi: pd.DataFrame
    registro: pd.DataFrame
    alerts: list[str] = field(default_factory=list)


@dataclass
class LegalizationResult:
    travel: AccountResult
    tolls: AccountResult
    source_rows: int

    def tables(self) -> dict[str, pd.DataFrame]:
        return {
            "BD gastos": self.travel.bd, "SAP1": self.travel.sap,
            "BD peajes": self.tolls.bd, "SAP2": self.tolls.sap,
            "DIFERENCIAS": pd.concat([self.travel.diferencias, self.tolls.diferencias], ignore_index=True),
            "CADI": pd.concat([self.travel.cadi, self.tolls.cadi], ignore_index=True),
            "REGISTRO": pd.concat([self.travel.registro, self.tolls.registro], ignore_index=True),
            "ALERTAS": pd.DataFrame({"alerta": self.travel.alerts + self.tolls.alerts}),
        }


def _prepare(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw.copy()
    df["concepto_n"] = df["Concepto gasto"].map(_norm)
    df = df[df["concepto_n"] != ELECTRONIC_TOLL]  # API already excludes it; kept as a safeguard
    df["gasto"] = pd.to_numeric(df["Gasto"], errors="coerce").fillna(0)
    df["placa"] = df["placa_sap"].map(_plate)
    df["tercero"] = df["Tercero"].map(lambda s: " ".join(str(s).split()))
    return df


def _build_bd(df: pd.DataFrame, cc_map: dict[str, str]) -> tuple[pd.DataFrame, list[str]]:
    bd = (df.groupby(["placa", "tercero"], as_index=False)["gasto"].sum()
            .rename(columns={"placa": "vehiculo", "gasto": "total_gasto"}))
    bd = bd[bd["total_gasto"] > 0]
    bd["centro_costo"] = bd["vehiculo"].map(cc_map)
    alerts = [f"Placa {p} sin centro de costo ({t}: {v:,.0f})"
              for p, t, v in bd.loc[bd["centro_costo"].isna(), ["vehiculo", "tercero", "total_gasto"]].values]
    return bd.sort_values(["centro_costo", "tercero"], na_position="last").reset_index(drop=True), alerts


def _build_sap(bd: pd.DataFrame, gl: str, balances: dict[str, float]) -> tuple[pd.DataFrame, list[str]]:
    ok = bd.dropna(subset=["centro_costo"])
    sap = pd.DataFrame({
        "cuenta_mayor": gl, "texto": "", "dh": "", "valor": ok["total_gasto"].round(0).astype("int64"),
        "saldo_sap": ok["centro_costo"].map(lambda c: balances.get(c, 0.0)),
        "col_h": "", "indicador": settings.sap_indicator, "col_j": "",
        "centro_costo": ok["centro_costo"], "cuenta_fija": settings.sap_fixed_account, "tercero": ok["tercero"],
    })
    sap["diferencia"] = sap["saldo_sap"] - sap["valor"]
    missing = sorted(set(ok["centro_costo"]) - set(balances))
    alerts = [f"CC {c} sin saldo en SAP (cuenta {gl})" for c in missing]
    return sap[SAP_COLUMNS].reset_index(drop=True), alerts


def _build_cadi(sap: pd.DataFrame, balances: dict[str, float], pospre: str, observacion: str,
                donor_pool: set[str]) -> tuple[pd.DataFrame, pd.DataFrame, list[str]]:
    """Deficit per CC (all its rows) is covered by donors' surplus (balance - own expense)."""
    need = sap.groupby("centro_costo")["valor"].sum()
    deficits = {cc: need[cc] - balances.get(cc, 0) for cc in need.index if need[cc] > balances.get(cc, 0)}
    surplus = {cc: balances[cc] - need.get(cc, 0) for cc in donor_pool & set(balances)
               if balances[cc] - need.get(cc, 0) > 0}

    diff = sap[sap["centro_costo"].isin(deficits)].copy()
    diff["faltante_cc"] = diff["centro_costo"].map(deficits)

    rows, alerts = [], []
    for receiver, amount in sorted(deficits.items(), key=lambda x: -x[1]):
        remaining = amount
        # Prefer a single donor that covers the whole deficit (smallest sufficient surplus).
        single = sorted((s, cc) for cc, s in surplus.items() if s >= remaining and cc != receiver)
        donors = [single[0][1]] if single else [cc for cc, _ in sorted(surplus.items(), key=lambda x: -x[1])
                                               if cc != receiver]
        for donor in donors:
            if remaining <= 0:
                break
            take = min(remaining, surplus[donor])
            if take <= 0:
                continue
            rows.append({"emisor": donor, "receptor": receiver, "pospre": pospre,
                         "importe": int(round(take)), "observacion": observacion})
            surplus[donor] -= take
            remaining -= take
        if remaining > 0:
            alerts.append(f"CC {receiver}: no hay saldo para cubrir {remaining:,.0f} ({observacion})")
    cadi = pd.DataFrame(rows, columns=["emisor", "receptor", "pospre", "importe", "observacion"])
    return diff.reset_index(drop=True), cadi, alerts


def _registro(sap: pd.DataFrame) -> pd.DataFrame:
    reg = sap[REGISTRO_COLUMNS].copy()
    reg.insert(0, "lote", (reg.index // FV50_BATCH) + 1)
    return reg


def _account(name, df, cc_map, gl, pospre, observacion, balances) -> AccountResult:
    bd, a1 = _build_bd(df, cc_map)
    sap, a2 = _build_sap(bd, gl, balances)
    # Donors must be cost centers being legalized in this same account.
    diff, cadi, a3 = _build_cadi(sap, balances, pospre, observacion, set(sap["centro_costo"]))
    return AccountResult(name, bd, sap, diff, cadi, _registro(sap), a1 + a2 + a3)


def build(raw: pd.DataFrame, cc_map: dict[str, str], balances: dict[str, dict[str, float]]) -> LegalizationResult:
    df = _prepare(raw)
    tolls = df[df["concepto_n"] == TOLL_CONCEPT]
    travel = df[df["concepto_n"] != TOLL_CONCEPT]
    s = settings
    return LegalizationResult(
        travel=_account("Gastos de viaje", travel, cc_map, s.gl_travel, s.pospre_travel,
                        "GASTOS DE VIAJES", balances["travel"]),
        tolls=_account("Peajes en efectivo", tolls, cc_map, s.gl_tolls, s.pospre_tolls,
                       "PEAJES EN EFECTIVO", balances["tolls"]),
        source_rows=len(df),
    )


def export_excel(result: LegalizationResult, path: Path) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for name, table in result.tables().items():
            table.to_excel(xw, sheet_name=name, index=False)
    return path


def export_run(result: LegalizationResult, base_dir: Path, meta: dict) -> Path:
    """Writes output/<timestamp>/ with legalizacion.xlsx + one JSON per table (RPA input)."""
    from datetime import datetime
    import json

    run_dir = Path(base_dir) / datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir.mkdir(parents=True, exist_ok=True)
    export_excel(result, run_dir / "legalizacion.xlsx")
    for name, table in result.tables().items():
        slug = _norm(name).lower().replace(" ", "_")
        records = json.loads(table.to_json(orient="records", force_ascii=False))
        (run_dir / f"{slug}.json").write_text(json.dumps(records, ensure_ascii=False, indent=1), encoding="utf-8")
    (run_dir / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    return run_dir
