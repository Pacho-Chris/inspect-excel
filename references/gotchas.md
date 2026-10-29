# Trampas frecuentes al leer Excel

## Dependencias de los scripts de consulta

Cada `queries/<slug>/<slug>.py` lleva un encabezado PEP 723 con
`pandas` y `openpyxl`:

```python
# /// script
# requires-python = ">=3.10"
# dependencies = ["pandas>=2.2", "openpyxl>=3.1"]
# ///
```

Eso permite `uv run queries/<slug>/<slug>.py` sin proyecto ni `--with`. Si
borras el encabezado, `uv run` falla con `ModuleNotFoundError: No module named
'pandas'`. No hace falta inicializar un proyecto uv en el directorio del
usuario.

## Fórmulas

`inspect_excel.py` abre con `data_only=True`: lee el **valor cacheado** que dejó
Excel. Si el archivo lo generó otra herramienta (pandas/openpyxl) y la fórmula
no tiene caché, la celda aparece como `null`. Solución: abrir y guardar el
archivo en Excel/LibreOffice para que recalcule y guarde el caché, y regenerar
el schema.

## Fechas ambiguas

Excel guarda las fechas como números con formato. `openpyxl` las convierte a
`datetime`. El schema las marca `date` o `datetime`. Si una columna es
`text` pero contiene fechas, o el formato es `mm/dd` y el usuario espera
`dd/mm`, confirma con el usuario y usa `dayfirst` / `format`.

## Ceros a la izquierda

Un ID `0001` guardado como texto aparece `text` con
`warnings: ["numeric_stored_as_text"]`. Léelo con `dtype={"ID": "string"}`.
Si está guardado como número, Excel ya perdió los ceros (`semantic:
"leading_zeros"` + `warnings: ["leading_zeros_lossy"]`); no se pueden recuperar
desde el Excel.

## Celdas combinadas

`openpyxl` en modo lectura solo expone el valor de la celda superior-izquierda
de una celda combinada; el resto aparece como `null`. Si necesitas repetir el
valor, usa `df.ffill()` sobre la columna.

## Filas de totales / pies

Si el Excel tiene una fila de "Total" al final, exclúyela con
`--footer-rows 1` al inspeccionar, o filtra en el script
(`df = df[~df["Cliente"].str.contains("total", case=False, na=False)]`).
En tablas de Excel con fila de totales, el inspector la excluye
automáticamente (`totalsRowCount`).

## Cabeceras desplazadas

Si la cabecera no está en la fila 1 (hay títulos arriba), el inspector intenta
detectarla entre las primeras `--header-scan` (20) filas. Si falla, fuerza con
`--header-row N`.

## Hojas ocultas

El schema marca `hidden: true`. No asumas que una hoja oculta es irrelevante;
pregunta antes de ignorarla.

## Columnas mixtas

Una columna con números y texto se marca `mixed` con el desglose en `types`.
Al leerla, pandas la deja `object`; decide con el usuario si limpiarla o
filtrarla.

## `NaN` vs vacío

`null` en el schema = celda vacía (o sin caché). En pandas aparece como `NaN`.
Para contar vacíos reales usa `df["col"].isna()`.

## `.xls` y `.xlsb`

No soportados. Pide al usuario que guarde como `.xlsx` y vuelva a inspeccionar.

## Archivos muy grandes

- Genera el schema con `--max-rows 20000` para no escanear todo.
- En los scripts usa `usecols` (solo columnas necesarias) y `nrows` para
  explorar.
- Para consultas repetidas sobre el mismo archivo grande, exporta una vez a
  Parquet y consulta el Parquet.
