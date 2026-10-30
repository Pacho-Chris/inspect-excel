# /// script
# requires-python = ">=3.10"
# dependencies = ["pandas>=2.2", "openpyxl>=3.1"]
# ///
"""Plantilla de consulta. Copia a `queries/<slug>.py` y ajusta.

Convenciones (mirar el excel_schema.json):
- header_row  -> usar en skiprows
- name        -> usar en names=[...]
- letter      -> usar en usecols="A:C" (solo las columnas pedidas)
- type/semantic -> dtype={} y parse_dates=[]

El encabezado PEP 723 (líneas # ///) hace el script autocontenido: se ejecuta con
`uv run queries/<slug>.py` sin proyecto ni `--with`. No lo borres.

Salida: la query imprime el resultado como JSON (`records`) a stdout y nada más.
Cualquier nota de progreso va a stderr. El JSON se puede:
  - encadenar:  uv run queries/a.py | uv run queries/b.py
  - guardar:    uv run queries/a.py > resultado.json
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

# Rutas robustas (no dependen del directorio desde donde se ejecute).
HERE = Path(__file__).parent            # queries/
ROOT = HERE.parent                      # directorio de trabajo (donde está el Excel)
XLSX = ROOT / "ventas.xlsx"             # nombre del archivo Excel (ajustar)

SHEET = "Ventas"
HEADER_ROW = 2                          # header_row del schema (fila Excel, 1-based)
COLUMNS = ["Fecha", "Cliente", "Monto"]  # names del schema
LETTERS = "A:C"                         # letras de esas columnas

df = pd.read_excel(
    XLSX,
    sheet_name=SHEET,
    header=None,
    skiprows=HEADER_ROW,                 # salta título + cabecera
    usecols=LETTERS,
    names=COLUMNS,
    dtype={"Monto": "float64"},          # opcional, según schema
    # parse_dates=["Fecha"],             # si hay fechas
)

# Consulta de ejemplo: filtrar
out = df[df["Monto"] > 1000]

# Emisión del resultado como JSON a stdout (nada más va a stdout).
print(out.to_json(orient="records", date_format="iso", force_ascii=False))

# Para consumir el JSON de otra query (sin re-leer el Excel):
#
#   import sys
#   df = pd.read_json(sys.stdin, orient="records")
#   res = df.groupby("Cliente")["Monto"].sum().reset_index()
#   print(res.to_json(orient="records", date_format="iso", force_ascii=False))
