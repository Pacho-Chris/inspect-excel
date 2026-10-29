# /// script
# requires-python = ">=3.10"
# dependencies = ["pandas>=2.2", "openpyxl>=3.1"]
# ///
"""Plantilla de consulta. Copia a `queries/<slug>/<slug>.py` y ajusta.

Convenciones (mirar el excel_schema.json):
- header_row  -> usar en skiprows
- name        -> usar en names=[...]
- letter      -> usar en usecols="A:C" (solo las columnas pedidas)
- type/semantic -> dtype={} y parse_dates=[]

El encabezado PEP 723 (líneas # ///) hace el script autocontenido: se ejecuta con
`uv run queries/<slug>/<slug>.py` sin proyecto ni `--with`. No lo borres.

El CSV se escribe en la carpeta de la query y se sobrescribe en cada ejecución.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)
pd.set_option("display.max_rows", 60)

# Rutas robustas (no dependen del directorio desde donde se ejecute).
HERE = Path(__file__).parent            # queries/<slug>/
ROOT = HERE.parent.parent               # directorio de trabajo (donde está el Excel)
XLSX = ROOT / "ventas.xlsx"             # nombre del archivo Excel (ajustar)
SLUG = Path(__file__).stem              # nombre de la query = nombre del .py
OUT_CSV = HERE / f"{SLUG}.csv"          # resultado (se sobrescribe en cada ejecución)

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

out.to_csv(OUT_CSV, index=False)         # recrea el CSV en cada ejecución

print(out.head(50).to_string(index=False))
print(f"\n{len(out)} filas de {len(df)}")
print(f"CSV → {OUT_CSV}")
