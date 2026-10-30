# RPA Legalización Elogia

App de escritorio en Python (Windows) que construye las tablas de **legalización de gastos de viaje y peajes en efectivo** para contabilizar en SAP (FV50) y los **traslados presupuestales de CADI**. Reemplaza el proceso manual del instructivo *Legalización en SAP*.

## Flujo

```
Silogtran (API Render)  ──►  Motor de cálculo  ──►  UI (tablas)  ──►  Exportar (Excel + JSON para RPA FV50)
SAP GR55 (RPA saldos)   ──►        ▲
Asignación Conductores  ──────────┘  (placa → centro de costo)
```

1. **Legalizaciones**: `GET /api/reports/legalizacion-viaje-sap` de [apis_legalizacion_silogtran](https://github.com/grachos/apis_legalizacion_silogtran) (sin PEAJE ELECTRÓNICO ni anulados).
2. **Centros de costo**: hoja `CENTROS DE COSTOS FLOTA` de `Asignación Conductores.xlsx` (no se versiona).
3. **Saldos SAP**: RPA sobre **GR55** (grupo `ZFM1`, centro gestor 25000000–29000000) o carga manual del Excel de saldos.
4. **Cálculo**:
   - *Gastos de viaje* = todos los conceptos excepto `PEAJES EN EFECTIVOS`; *Peajes* = solo `PEAJES EN EFECTIVOS`.
   - BD gastos / BD peajes: suma por placa + tercero, con su centro de costo.
   - SAP1 / SAP2: estructura FV50 (cuenta mayor, texto y dh vacíos, valor, saldo, diferencia, V7, CC, 222222222, tercero).
   - DIFERENCIAS: CC cuyo gasto total supera el disponible.
   - CADI: emisor que también se legaliza en la misma cuenta, con sobrante suficiente (saldo − su propio gasto), sin dejarlo en negativo.
   - REGISTRO: SAP1 + SAP2 sin columnas auxiliares, en lotes de 4 para FV50.

## Cuentas

| Concepto | Cuenta FV50 | Pospre (GR55 y CADI) |
|---|---|---|
| Gastos de viaje | 6130105595 | 61301055 |
| Peajes en efectivo | 6130105530 | 6130105530 |

## Instalación

Requisitos: Windows, Python 3.11+, SAP GUI con scripting habilitado, Excel.

```bat
probar_api.bat      :: crea .venv, instala dependencias, crea .env y prueba la API
```

Completar `.env` (ver `.env.example`): `API_KEY`, ambiente SAP (`SAP_ENV=QAS|PRD`) y conexión/mandante/usuario/contraseña por ambiente. Copiar `Asignación Conductores.xlsx` en la raíz del proyecto.

## Uso

```bat
iniciar_app.bat
```

1. Elegir ambiente SAP (QAS/PRD) y pulsar **Traer saldos de SAP (GR55)** (o seleccionar un Excel de saldos).
2. Elegir fechas y pulsar **Generar**.
3. Revisar las pestañas (filas en rojo = diferencia negativa) y **Exportar** → `output/<fecha_hora>/` con `legalizacion.xlsx`, un JSON por tabla y `meta.json`.

El botón **Exportar a Excel** de cada pestaña guarda esa tabla sola en un .xlsx y lo abre.

## Scripts de prueba

| Script | Qué prueba |
|---|---|
| `probar_api.bat` | Descarga del informe cruzado de la API |
| `probar_motor.bat` | Motor de cálculo con el último CSV + Excel de saldos |
| `probar_sap.bat` | Login SAP y apertura de GR55 |
| `probar_gr55.bat` | RPA de saldos GR55 |

## Estructura

```
legalizacion/
  config.py         # settings desde .env
  api_client.py     # cliente API Silogtran (Render)
  engine.py         # tablas BD, SAP, diferencias, CADI, registro, exportación
  sap_gui.py        # conexión/login SAP GUI Scripting
  gr55.py           # RPA saldos GR55 (lee el Excel que abre SAP por integración Office)
  sap_security.py   # acepta los avisos "Seguridad SAP GUI" de los temporales ~SAP{...}
  ui.py             # interfaz CustomTkinter
scripts/            # pruebas por incremento
docs/gr55_record_qas.vbs  # grabación SAP de la pantalla de selección GR55
```

## Pendiente (fase 2)

- RPA de contabilización en **FV50** a partir de `registro.json`.
- Registro de traslados en **CADI** a partir de `cadi.json`.
