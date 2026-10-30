"""Increment 2 check: build all tables from a saved API CSV + balances Excel.

Usage: python -m scripts.test_engine output/viaje_sap_2026-10-01_2026-10-30.csv "Gastos varios y peajes.xlsx"
"""
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

from legalizacion import engine
from legalizacion.config import settings


def main() -> None:
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    raw = pd.read_csv(sys.argv[1], dtype=str, keep_default_na=False)
    result = engine.build(raw, engine.load_cost_centers(), engine.load_balances_excel(Path(sys.argv[2])))
    for name, t in result.tables().items():
        total = f" | valor: {t['valor'].sum():,.0f}" if "valor" in t else ""
        total += f" | importe: {t['importe'].sum():,.0f}" if "importe" in t else ""
        print(f"{name:12} {len(t):5} filas{total}")
    out = engine.export_excel(result, settings.output_dir / f"legalizacion_{datetime.now():%Y%m%d_%H%M}.xlsx")
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
