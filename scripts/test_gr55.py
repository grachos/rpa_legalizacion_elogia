"""Increment 4 check: fetch GR55 balances from SAP and save them to output/."""
import sys

from legalizacion.gr55 import fetch_balances, save_balances
from legalizacion.config import sap_connection
from legalizacion.sap_gui import SapError


def main() -> None:
    env = (sys.argv[1] if len(sys.argv) > 1 else sap_connection().env).upper()
    try:
        b = fetch_balances(env)
    except (SapError, ValueError) as e:
        sys.exit(f"[ERROR] {e}")
    for k, v in b.items():
        print(f"{k:7} {len(v):4} CC | disponible total: {sum(v.values()):,.0f}")
    print("Guardado:", save_balances(b, env))


if __name__ == "__main__":
    main()
