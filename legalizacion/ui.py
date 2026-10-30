"""Desktop UI (CustomTkinter) to generate and review the legalization tables."""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import json
import os
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import pandas as pd

from . import engine
from .api_client import LegalizacionApi
from .config import FIELDS, ROOT, is_configured, read_env, settings, write_env

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

    def __init__(self, master, name: str):
        super().__init__(master, fg_color="transparent")
        self.name = name
        self.df = pd.DataFrame()
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", pady=(0, 4))
        self.summary = ctk.CTkLabel(bar, text="Sin datos", anchor="w")
        self.summary.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(bar, text="Exportar a Excel", width=170, command=self.export).pack(side="right")

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

    def export(self) -> None:
        """Saves this table alone to an Excel file and opens it."""
        if self.df.empty:
            messagebox.showinfo("Exportar", "No hay datos en esta tabla.")
            return
        settings.output_dir.mkdir(parents=True, exist_ok=True)
        slug = self.name.replace(" ", "_")
        path = filedialog.asksaveasfilename(
            title=f"Exportar {self.name}", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialdir=str(settings.output_dir), initialfile=f"{slug}_{datetime.now():%Y%m%d_%H%M}.xlsx")
        if not path:
            return
        try:
            with pd.ExcelWriter(path, engine="openpyxl") as xw:
                self.df.to_excel(xw, sheet_name=self.name[:31], index=False)
        except PermissionError:
            messagebox.showerror("Exportar", "No se pudo escribir. ¿El archivo está abierto en Excel?")
            return
        os.startfile(path)  # Windows only

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
        self.after(50, self._maximize)
        self.minsize(980, 600)
        self.state_data = _load_state()
        self.result: engine.LegalizationResult | None = None
        self.meta: dict = {}

        self._build_header()
        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(fill="both", expand=True, padx=12, pady=(0, 6))
        self.views = {}
        for name in self.TABS:
            view = TableView(self.tabs.add(name), name)
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

        ctk.CTkLabel(top, text="Ambiente SAP").grid(row=0, column=4, padx=(14, 4))
        self.env_var = tk.StringVar(value="PRD")
        ctk.CTkSegmentedButton(top, values=["PRD", "QAS"], variable=self.env_var,
                               command=lambda _v: self._clear_balances()).grid(row=0, column=5, sticky="w")
        ctk.CTkButton(top, text="⚙ Configuración", width=130, fg_color="#607D8B", hover_color="#455A64",
                      command=self._open_settings).grid(row=0, column=8, padx=(14, 4))

        self.run_btn = ctk.CTkButton(top, text="Generar", width=110, state="disabled", command=self._run)
        self.run_btn.grid(row=0, column=9, padx=(14, 4))
        self.export_btn = ctk.CTkButton(top, text="Exportar", width=110, state="disabled", command=self._export)
        self.export_btn.grid(row=0, column=10, padx=(4, 10))
        self.sap_btn = ctk.CTkButton(top, text="Traer saldos de SAP (GR55)", width=200,
                                     fg_color="#2E7D32", hover_color="#1B5E20", command=self._fetch_sap)
        self.sap_btn.grid(row=1, column=0, columnspan=2, sticky="w", padx=(10, 4), pady=(0, 8))
        self.sap_info = ctk.CTkLabel(top, text="Paso 1: trae los saldos de SAP. Paso 2: Generar.", anchor="w")
        self.sap_info.grid(row=1, column=2, columnspan=9, sticky="w", padx=8, pady=(0, 8))
        top.columnconfigure(7, weight=1)
        self.balances: dict | None = None
        self.balances_meta: dict = {}

        self.progress_box = ctk.CTkFrame(self, fg_color="transparent")
        self.progress = ctk.CTkProgressBar(self.progress_box, mode="determinate", height=14)
        self.progress.pack(side="left", fill="x", expand=True)
        self.progress_lbl = ctk.CTkLabel(self.progress_box, text="", width=420, anchor="w")
        self.progress_lbl.pack(side="left", padx=(10, 0))
        self._pct = self._ceil = 0.0
        self._step_text, self._step_t0, self._ticking = "", 0.0, False

    def _clear_balances(self) -> None:
        if self.balances is not None:
            self.balances, self.balances_meta = None, {}
            self.run_btn.configure(state="disabled")
            self.sap_info.configure(text="Cambiaste de ambiente: vuelve a traer los saldos de SAP.")

    def _maximize(self) -> None:
        try:
            self.state("zoomed")  # Windows
        except tk.TclError:
            self.attributes("-zoomed", True)  # Linux (dev)

    def _open_settings(self) -> None:
        SettingsDialog(self)

    # ---------- actions ----------
    def _busy(self, on: bool, text: str = "") -> None:
        self.run_btn.configure(state="disabled" if on or self.balances is None else "normal")
        if on:
            self._pct = self._ceil = 0.0
            self.progress.set(0)
            self.progress_box.pack(fill="x", padx=14, pady=(0, 6), before=self.tabs)
            self._ticking = True
            self._tick()
        else:
            self._ticking = False
            self.progress_box.pack_forget()
        if text:
            self.status.configure(text=text)

    def _step(self, start: float, ceiling: float, text: str) -> None:
        """Thread-safe: starts a stage at `start`% that creeps toward `ceiling`% while it runs."""
        def apply():
            self._pct, self._ceil = max(self._pct, start), ceiling
            self._step_text, self._step_t0 = text, time.time()
            self._render()
        self.after(0, apply)

    def _tick(self) -> None:
        if not self._ticking:
            return
        if self._pct < self._ceil:  # long stages: approach the ceiling without reaching it
            self._pct += max((self._ceil - self._pct) * 0.04, 0.05)
            self._pct = min(self._pct, self._ceil - 0.5 if self._ceil < 100 else 100)
        self._render()
        self.after(400, self._tick)

    def _render(self) -> None:
        secs = int(time.time() - self._step_t0) if self._step_t0 else 0
        self.progress.set(self._pct / 100)
        self.progress_lbl.configure(text=f"{self._pct:5.0f} %  ·  {self._step_text}  ({secs} s)")

    def _fetch_sap(self) -> None:
        env = self.env_var.get()
        self.sap_btn.configure(state="disabled")
        self._busy(True, f"Consultando GR55 en SAP {env}…")
        threading.Thread(target=self._sap_worker, args=(env,), daemon=True).start()

    def _sap_worker(self, env: str) -> None:
        from .gr55 import fetch_balances, save_balances

        try:
            balances = fetch_balances(env, progress=self._step)
            if not balances["travel"] and not balances["tolls"]:
                raise ValueError("La GR55 no devolvió datos.")
            path = save_balances(balances, env)  # audit copy
            info = (f"Saldos SAP {env}: {len(balances['travel'])} CC gastos · {len(balances['tolls'])} CC peajes")
            self.after(0, lambda: self._sap_done(balances, path, env, info))
        except Exception as e:  # noqa: BLE001 - shown to the user
            msg = str(e)
            self.after(0, lambda: (self.sap_btn.configure(state="normal"),
                                   self._busy(False, f"Error SAP: {msg}"), messagebox.showerror("SAP", msg)))

    def _sap_done(self, balances: dict, path: Path, env: str, info: str) -> None:
        self.balances = balances
        self.balances_meta = {"saldos": str(path), "ambiente_sap": env, "saldos_hora": datetime.now()}
        self.sap_btn.configure(state="normal")
        self.sap_info.configure(text=f"{info} · {datetime.now():%H:%M}")
        self._busy(False, "Saldos listos. Ya puedes Generar.")

    def _run(self) -> None:
        try:
            start, end = date.fromisoformat(self.start_var.get().strip()), date.fromisoformat(self.end_var.get().strip())
        except ValueError:
            messagebox.showerror("Fechas", "Usa el formato YYYY-MM-DD.")
            return
        if end < start:
            messagebox.showerror("Fechas", "La fecha final es anterior a la inicial.")
            return
        if self.balances is None:
            messagebox.showwarning("Saldos", "Primero trae los saldos de SAP (GR55).")
            return
        self._busy(True, "Consultando Silogtran (la API puede tardar si estaba dormida)…")
        threading.Thread(target=self._worker, args=(start, end, self.balances), daemon=True).start()

    def _worker(self, start: date, end: date, balances: dict) -> None:
        try:
            api = LegalizacionApi()
            self._step(0, 15, "Despertando la API de Silogtran")
            api.wake_up()
            self._step(15, 70, f"Descargando legalizaciones {start} a {end}")
            raw = api.legalizacion_viaje_sap(start, end)
            if raw.empty:
                raise ValueError("La API no devolvió registros para ese rango.")
            self._step(70, 80, f"{len(raw)} registros recibidos · leyendo centros de costo")
            cc_map = engine.load_cost_centers()
            self._step(80, 95, "Cruzando legalizaciones con saldos SAP")
            result = engine.build(raw, cc_map, balances)
            self._step(95, 100, "Mostrando tablas")
            meta = {"fecha_inicial": start, "fecha_final": end, "registros_api": len(raw), **self.balances_meta}
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


class SettingsDialog(ctk.CTkToplevel):
    """Edits the .env values (API, SAP users, files, accounts)."""

    def __init__(self, app: "App"):
        super().__init__(app)
        self.title("Configuración")
        self.geometry("620x640")
        self.transient(app)
        self.grab_set()
        values = read_env()
        self.vars: dict[str, tk.StringVar] = {}
        body = ctk.CTkScrollableFrame(self)
        body.pack(fill="both", expand=True, padx=12, pady=(12, 6))
        body.columnconfigure(1, weight=1)
        row = 0
        for section, fields in FIELDS:
            ctk.CTkLabel(body, text=section, font=ctk.CTkFont(weight="bold")).grid(
                row=row, column=0, columnspan=3, sticky="w", pady=(12 if row else 0, 4))
            row += 1
            for f in fields:
                var = tk.StringVar(value=values.get(f.key, f.default))
                self.vars[f.key] = var
                ctk.CTkLabel(body, text=f.label, anchor="w").grid(row=row, column=0, sticky="w", padx=(8, 8), pady=2)
                ctk.CTkEntry(body, textvariable=var, show="•" if f.secret else "").grid(
                    row=row, column=1, sticky="ew", pady=2)
                if f.kind == "file":
                    ctk.CTkButton(body, text="…", width=32, command=lambda v=var: self._pick(v)).grid(
                        row=row, column=2, padx=4)
                row += 1
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(bar, text="Guardar", width=120, command=self._save).pack(side="right")
        ctk.CTkButton(bar, text="Cancelar", width=120, fg_color="#9E9E9E", hover_color="#757575",
                      command=self.destroy).pack(side="right", padx=8)

    def _pick(self, var: tk.StringVar) -> None:
        path = filedialog.askopenfilename(parent=self, filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            var.set(path)

    def _save(self) -> None:
        try:
            write_env({k: v.get().strip() if k != "DRIVERS_SHEET" else v.get() for k, v in self.vars.items()})
        except OSError as e:
            messagebox.showerror("Configuración", f"No se pudo guardar: {e}", parent=self)
            return
        messagebox.showinfo("Configuración", "Configuración guardada.", parent=self)
        self.destroy()


def main() -> None:
    app = App()
    if not is_configured():
        app.after(300, lambda: (messagebox.showinfo(
            "Configuración", "Primera vez: completa la API key y tu usuario SAP.", parent=app),
            app._open_settings()))
    app.mainloop()
