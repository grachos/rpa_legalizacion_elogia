"""Increment 1 check: download the crossed report and save it to output/.

Usage: python -m scripts.test_api 2026-10-01 2026-10-15
"""
import sys
from datetime import date

from legalizacion.api_client import ApiError, LegalizacionApi
from legalizacion.config import settings


def ask_date(label: str, default: date) -> date:
    while True:
        raw = input(f"{label} [{default}]: ").strip()
        if not raw:
            return default
        try:
            value = date.fromisoformat(raw)
        except ValueError:
            print("  Formato invalido, usa YYYY-MM-DD.")
            continue
        if not 2020 <= value.year <= date.today().year:
            print(f"  Año fuera de rango (2020-{date.today().year}).")
            continue
        return value


def main() -> None:
    if len(sys.argv) == 3:
        start, end = (date.fromisoformat(a) for a in sys.argv[1:])
    else:
        start = ask_date("Fecha inicial (YYYY-MM-DD)", date.today().replace(day=1))
        end = ask_date("Fecha final   (YYYY-MM-DD)", date.today())
    try:
        run(start, end)
    except ApiError as e:
        sys.exit(f"[ERROR] {e}")


def run(start: date, end: date) -> None:
    api = LegalizacionApi()
    print("Waking up API...", api.wake_up())
    df = api.legalizacion_viaje_sap(start, end)
    print(f"Rows: {len(df)}")
    print("Columns:", list(df.columns))
    print(df.head(5).to_string())
    settings.output_dir.mkdir(exist_ok=True)
    out = settings.output_dir / f"viaje_sap_{start}_{end}.csv"
    df.to_csv(out, index=False, encoding="utf-8-sig")
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
