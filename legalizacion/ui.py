"""Desktop UI (CustomTkinter) to generate and review the legalization tables."""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import json
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import pandas as pd

from . import engine
from .api_client import LegalizacionApi
from .config import ROOT, settings

STATE_FILE = ROOT / ".ui_state.json"
MONEY_COLUMNS = {"valor", "saldo_sap", "diferencia", "total_gasto", "importe", "faltante_cc"}
TOTAL_COLUMNS = ("valor", "importe", "total_gasto")
NEGATIVE_COLUMNS = ("diferencia",)


def _load_state() -> dict:
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _save_state(state: dict) -> None:
    try:
        STATE_FILE.write_text(json.dumps(state, indent=1), encoding="utf-8")
    except OSError:
        pass


def _money(v) -> str:
    try:
        return f"{float(v):,.0f}".replace(",", ".")
    except (TypeError, ValueError):
        return "" if pd.isna(v) else str(v)


def _cell(col: str, v) -> str:
    if col in MONEY_COLUMNS:
        return _money(v)
    return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)


class TableView(ctk.CTkFrame):
    """Treeview with scrollbars, totals footer and copy-to-clipboard."""

    def __init__(self, master):
        super().__init__(master, fg_color="transparent")
        self.df = pd.DataFrame()
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", pady=(0, 4))
        self.summary = ctk.CTkLabel(bar, text="Sin datos", anchor="w")
        self.summary.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(bar, text="Copiar (sin encabezado)", width=170, command=self.copy).pack(side="right")

        holder = tk.Frame(self)
        holder.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(holder, show="headings", selectmode="extended")
        ys = ttk.Scrollbar(holder, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(holder, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        holder.rowconfigure(0, weight=1)
        holder.columnconfigure(0, weight=1)
        self.tree.tag_configure("neg", foreground="#C62828")
        self.tree.tag_configure("odd", background="#F5F7FA")

    def show(self, df: pd.DataFrame) -> None:
        self.df = df.reset_index(drop=True)
        cols = list(self.df.columns)
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
            anchor = "e" if c in MONEY_COLUMNS else "w"
            width = 260 if c in ("tercero", "alerta") else 110
            self.tree.column(c, width=width, anchor=anchor, stretch=c in ("tercero", "alerta"))
        neg_cols = [c for c in NEGATIVE_COLUMNS if c in cols]
        for i, row in enumerate(self.df.itertuples(index=False)):
            rec = row._asdict() if hasattr(row, "_asdict") else dict(zip(cols, row))
            values = [_cell(c, v) for c, v in zip(cols, row)]
            tags = ["odd"] if i % 2 else []
            if any((pd.to_numeric(rec.get(c), errors="coerce") or 0) < 0 for c in neg_cols):
                tags.append("neg")
            self.tree.insert("", "end", values=values, tags=tags)
        parts = [f"{len(self.df)} filas"]
        for c in TOTAL_COLUMNS:
            if c in cols:
                parts.append(f"Total {c}: $ {_money(self.df[c].sum())}")
        if neg_cols:
            n = int((pd.to_numeric(self.df[neg_cols[0]], errors="coerce") < 0).sum())
            parts.append(f"En negativo: {n}")
        self.summary.configure(text="   |   ".join(parts))

    def copy(self) -> None:
        """Copies selected rows (or all) as tab-separated text, ready to paste in SAP."""
        sel = self.tree.selection() or self.tree.get_children()
        if not sel:
            return
        idx = [self.tree.index(i) for i in sel]
        data = self.df.iloc[idx].copy()
        for c in data.columns:
            data[c] = data[c].map(lambda v, c=c: "" if pd.isna(v) else (f"{v:.0f}" if c in MONEY_COLUMNS else str(v)))
        text = "\n".join("\t".join(r) for r in data.astype(str).values.tolist())
        self.clipboard_clear()
        self.clipboard_append(text)
        self.summary.configure(text=self.summary.cget("text").split("   ·")[0] + f"   · {len(idx)} filas copiadas")


class App(ctk.CTk):
    TABS = ["BD gastos", "SAP1", "BD peajes", "SAP2", "DIFERENCIAS", "CADI", "REGISTRO", "ALERTAS"]

    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("light")
        self.title("Auto legalización SAP")
        self.geometry("1280x760")
        self.minsize(980, 600)
        self.state_data = _load_state()
        self.result: engine.LegalizationResult | None = None
        self.meta: dict = {}

        self._build_header()
        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        self.views = {}
        for name in self.TABS:
            view = TableView(self.tabs.add(name))
            view.pack(fill="both", expand=True)
            self.views[name] = view
        self.status = ctk.CTkLabel(self, text="Listo.", anchor="w")
        self.status.pack(fill="x", padx=14, pady=(0, 8))

    # ---------- layout ----------
    def _build_header(self) -> None:
        top = ctk.CTkFrame(self)
        top.pack(fill="x", padx=12, pady=12)
        today = date.today()

        ctk.CTkLabel(top, text="Fecha inicial").grid(row=0, column=0, padx=(10, 4), pady=8)
        self.start_var = tk.StringVar(value=today.replace(day=1).isoformat())
        ctk.CTkEntry(top, textvariable=self.start_var, width=110).grid(row=0, column=1)
        ctk.CTkLabel(top, text="Fecha final").grid(row=0, column=2, padx=(14, 4))
        self.end_var = tk.StringVar(value=today.isoformat())
        ctk.CTkEntry(top, textvariable=self.end_var, width=110).grid(row=0, column=3)

        ctk.CTkLabel(top, text="Saldos SAP").grid(row=0, column=4, padx=(14, 4))
        default_bal = self.state_data.get("balances_file") or str(ROOT / "Gastos varios y peajes.xlsx")
        self.bal_var = tk.StringVar(value=default_bal)
        ctk.CTkEntry(top, textvariable=self.bal_var, width=300).grid(row=0, column=5, sticky="ew")
        ctk.CTkButton(top, text="…", width=32, command=self._pick_balances).grid(row=0, column=6, padx=4)

        ctk.CTkLabel(top, text="Ambiente SAP").grid(row=0, column=7, padx=(14, 4))
        self.env_var = tk.StringVar(value=os.getenv("SAP_ENV", "QAS").upper())
        ctk.CTkSegmentedButton(top, values=["QAS", "PRD"], variable=self.env_var).grid(row=0, column=8)

        self.run_btn = ctk.CTkButton(top, text="Generar", width=110, command=self._run)
        self.run_btn.grid(row=0, column=9, padx=(14, 4))
        self.export_btn = ctk.CTkButton(top, text="Exportar", width=110, state="disabled", command=self._export)
        self.export_btn.grid(row=0, column=10, padx=(4, 10))
        self.sap_btn = ctk.CTkButton(top, text="Traer saldos de SAP (GR55)", width=200,
                                     fg_color="#2E7D32", hover_color="#1B5E20", command=self._fetch_sap)
        self.sap_btn.grid(row=1, column=5, sticky="w", pady=(0, 8))
        self.sap_info = ctk.CTkLabel(top, text="o selecciona un Excel de saldos con …", anchor="w")
        self.sap_info.grid(row=1, column=6, columnspan=5, sticky="w", padx=8, pady=(0, 8))
        top.columnconfigure(5, weight=1)

        self.progress = ctk.CTkProgressBar(self, mode="indeterminate", height=6)

    def _pick_balances(self) -> None:
        path = filedialog.askopenfilename(title="Archivo de saldos SAP (GR55)",
                                          initialdir=str(Path(self.bal_var.get()).parent),
                                          filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            self.bal_var.set(path)

    # ---------- actions ----------
    def _busy(self, on: bool, text: str = "") -> None:
        self.run_btn.configure(state="disabled" if on else "normal")
        if on:
            self.progress.pack(fill="x", padx=12, before=self.tabs)
            self.progress.start()
        else:
            self.progress.stop()
            self.progress.pack_forget()
        if text:
            self.status.configure(text=text)

    def _fetch_sap(self) -> None:
        env = self.env_var.get()
        if env == "PRD" and not messagebox.askyesno("SAP", "Se consultará la GR55 en PRODUCTIVO (solo lectura). ¿Continuar?"):
            return
        self.sap_btn.configure(state="disabled")
        self._busy(True, f"Consultando GR55 en SAP {env}…")
        threading.Thread(target=self._sap_worker, args=(env,), daemon=True).start()

    def _sap_worker(self, env: str) -> None:
        from .gr55 import fetch_balances, save_balances

        try:
            balances = fetch_balances(env)
            if not balances["travel"] and not balances["tolls"]:
                raise ValueError("La GR55 no devolvió datos.")
            path = save_balances(balances, env)
            info = (f"SAP {env}: {len(balances['travel'])} CC gastos · {len(balances['tolls'])} CC peajes")
            self.after(0, lambda: self._sap_done(path, info))
        except Exception as e:  # noqa: BLE001 - shown to the user
            msg = str(e)
            self.after(0, lambda: (self.sap_btn.configure(state="normal"),
                                   self._busy(False, f"Error SAP: {msg}"), messagebox.showerror("SAP", msg)))

    def _sap_done(self, path: Path, info: str) -> None:
        self.bal_var.set(str(path))
        self.sap_btn.configure(state="normal")
        self.sap_info.configure(text=f"{info} · {datetime.now():%H:%M}")
        self._busy(False, f"Saldos guardados en {path.name}. Ya puedes Generar.")

    def _run(self) -> None:
        try:
            start, end = date.fromisoformat(self.start_var.get().strip()), date.fromisoformat(self.end_var.get().strip())
        except ValueError:
            messagebox.showerror("Fechas", "Usa el formato YYYY-MM-DD.")
            return
        if end < start:
            messagebox.showerror("Fechas", "La fecha final es anterior a la inicial.")
            return
        bal = Path(self.bal_var.get().strip())
        if not bal.exists():
            messagebox.showerror("Saldos", f"No existe el archivo de saldos:\n{bal}")
            return
        self.state_data["balances_file"] = str(bal)
        _save_state(self.state_data)
        self._busy(True, "Consultando Silogtran (la API puede tardar si estaba dormida)…")
        threading.Thread(target=self._worker, args=(start, end, bal), daemon=True).start()

    def _worker(self, start: date, end: date, bal: Path) -> None:
        try:
            api = LegalizacionApi()
            api.wake_up()
            raw = api.legalizacion_viaje_sap(start, end)
            self.after(0, lambda: self.status.configure(text=f"{len(raw)} registros recibidos. Calculando…"))
            if raw.empty:
                raise ValueError("La API no devolvió registros para ese rango.")
            result = engine.build(raw, engine.load_cost_centers(), engine.load_balances_excel(bal))
            meta = {"fecha_inicial": start, "fecha_final": end, "saldos": str(bal),
                    "ambiente_sap": self.env_var.get(), "registros_api": len(raw)}
            self.after(0, lambda: self._show(result, meta))
        except Exception as e:  # noqa: BLE001 - shown to the user
            msg = str(e)
            self.after(0, lambda: (self._busy(False, f"Error: {msg}"), messagebox.showerror("Error", msg)))

    def _show(self, result: engine.LegalizationResult, meta: dict) -> None:
        self.result, self.meta = result, meta
        for name, table in result.tables().items():
            self.views[name].show(table)
        n_alerts = len(result.travel.alerts) + len(result.tolls.alerts)
        self.export_btn.configure(state="normal")
        self._busy(False, f"Listo: {meta['registros_api']} registros · {meta['fecha_inicial']} a {meta['fecha_final']} · "
                          f"{n_alerts} alertas · Ambiente {meta['ambiente_sap']}")
        self.tabs.set("ALERTAS" if n_alerts else "SAP1")

    def _export(self) -> None:
        if not self.result:
            return
        try:
            folder = engine.export_run(self.result, settings.output_dir, self.meta)
        except PermissionError:
            messagebox.showerror("Exportar", "No se pudo escribir. ¿Hay un archivo abierto en Excel?")
            return
        self.status.configure(text=f"Exportado en {folder}")
        if messagebox.askyesno("Exportar", f"Archivos guardados en:\n{folder}\n\n¿Abrir la carpeta?"):
            os.startfile(folder)  # Windows only


def main() -> None:
    App().mainloop()
