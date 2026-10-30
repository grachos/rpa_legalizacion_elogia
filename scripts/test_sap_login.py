"""Increment 4 check: log in to SAP (SAP_ENV) and open GR55."""
import sys

from legalizacion.sap_gui import SapError, SapSession


def main() -> None:
    env = sys.argv[1] if len(sys.argv) > 1 else None
    try:
        with SapSession(env) as sap:
            info = sap.session.Info
            print(f"Conectado: sistema {info.SystemName}, mandante {info.Client}, usuario {info.User}")
            sap.transaction("GR55")
            print("Ventana:", sap.session.findById("wnd[0]").Text)
            input("GR55 abierta. Revisa SAP y presiona Enter para cerrar...")
    except (SapError, ValueError) as e:
        sys.exit(f"[ERROR] {e}")


if __name__ == "__main__":
    main()
