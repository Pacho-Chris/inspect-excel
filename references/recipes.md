# Recetas de consulta

Todas asumen que ya leíste `excel_schema.json` y conoces `header_row`,
`data_start_row`, `name`, `letter` y `type` de las columnas involucradas.

Cada consulta va en su propia carpeta:

```text
queries/<slug>/
  <slug>.py    # script autocontenido (encabezado PEP 723)
  <slug>.csv   # resultado, se sobrescribe en cada ejecución
```

Ejecutar (desde la raíz del proyecto, sin `--with`):

```bash
uv run queries/<slug>/<slug>.py
```

Base común para cargar una hoja con las columnas pedidas. El encabezado
PEP 723 y las rutas por `__file__` son obligatorios:

```python
# /// script
# requires-python = ">=3.10"
# dependencies = ["pandas>=2.2", "openpyxl>=3.1"]
# ///
from pathlib import Path
import pandas as pd

HERE = Path(__file__).parent
ROOT = HERE.parent.parent
XLSX = ROOT / "ventas.xlsx"
OUT_CSV = HERE / (Path(__file__).stem + ".csv")   # <slug>.csv, se sobrescribe

df = pd.read_excel(
    XLSX,
    sheet_name="Ventas",
    header=None,
    skiprows=2,               # header_row
    usecols="A:C",            # letras de las columnas
    names=["Fecha", "Cliente", "Monto"],  # names del schema
)
```

## Seleccionar columnas

Busca en el schema sus `letter` y
`name` y úsalos arriba. No leas todas las columnas y luego filtres.

## Filtrar

```python
out = df[df["Monto"] > 1000]                 # numérico
out = df[df["Cliente"].str.contains("A", na=False)]  # texto
out = df[df["Cliente"].isin(["A", "B"])]     # lista
out = df[df["Fecha"] >= "2024-01-01"]        # fechas (tras parse_dates)
out = df[df["Nota"].notna()]                 # no nulos
```

Múltiples condiciones:

```python
m = (df["Monto"] > 1000) & (df["Pagado"] == True)  # noqa: E712
out = df[m]
```

## Ordenar y top N

```python
out = df.sort_values("Monto", ascending=False).head(10)
```

## Agrupar y agregar

```python
res = df.groupby("Cliente", as_index=False)["Monto"].agg(total=("sum"), n=("count"))
res = res.sort_values("total", ascending=False)
```

## Fechas

```python
df = pd.read_excel(
    XLSX, sheet_name="Ventas", header=None, skiprows=2,
    usecols="A:C", names=["Fecha", "Cliente", "Monto"],
    parse_dates=["Fecha"],
)
res = df.groupby(df["Fecha"].dt.month)["Monto"].sum()   # total por mes
```

Si las fechas son ambiguas (`dd/mm` vs `mm/dd`), pregúntale al usuario y usa
`dayfirst=True` o `format="%d/%m/%Y"` en `read_excel`/`to_datetime`.

## Preservar ceros a la izquierda (IDs)

```python
df = pd.read_excel(..., dtype={"ID": "string"})
```

La columna `ID` con `warnings: ["numeric_stored_as_text"]` debe leerse como
texto, no como número.

## Múltiples hojas

```python
a = pd.read_excel(XLSX, sheet_name="Ventas", header=None, skiprows=2, usecols="A:C", names=["Fecha", "Cliente", "Monto"])
b = pd.read_excel(XLSX, sheet_name="Clientes", header=None, skiprows=2, usecols="A:F", names=["ID", "Nombre", "Activo", "Edad", "Categoria", "Saldo"])
```

## Unir hojas (join)

```python
j = pd.merge(a, b, left_on="Cliente", right_on="Nombre", how="left")
```

## Valores distintos / enum

```python
df["Categoria"].value_counts(dropna=False)
df["Categoria"].unique()
```

## Exportar resultado (obligatorio: sobrescribe el CSV de la query)

```python
out.to_csv(OUT_CSV, index=False)          # CSV, en la carpeta de la query
# out.to_parquet(HERE / "<slug>.parquet", index=False)  # si es muy grande
```

Siempre escribe el resultado a `OUT_CSV` (o Parquet) dentro de la carpeta de la
query. Devuelve la ruta y un resumen (`shape`, primeras filas), no el contenido
completo.

## Leer solo una parte (archivos muy grandes)

```python
df = pd.read_excel(..., nrows=5000)   # primeras filas, para explorar
```
