---
name: inspect-excel
description: >
  Lee archivos Excel (.xlsx/.xlsm), detecta hojas y tablas de Excel, infiere
  tipos de datos y genera un excel_schema.json para consultar sin abrir el
  Excel en el contexto. Úsala cuando el usuario comparta un Excel o pida
  seleccionar columnas, filtrar, agrupar, unir hojas o resumir datos de un
  spreadsheet. No usar para editar estilos/formato, crear dashboards ni para
  CSV.
license: MIT
compatibility: Requiere uv y Python 3.10+. Entrada .xlsx/.xlsm. Windows, macOS y Linux.
metadata:
  version: "1.0.0"
---

# inspect-excel

Consulta archivos Excel masivos **sin abrirlos en el contexto de la IA**. El CLI
`inspect_excel.py` genera un `excel_schema.json` (metadatos: hojas, tablas,
columnas, tipos, nulos, muestras). Después la IA lee solo ese schema y genera
scripts de pandas mínimos para responder consultas del usuario.

```text
inspect_excel.py   → inspecciona el Excel y escribe excel_schema.json (se conserva)
excel_schema.json  → contrato de metadatos (lo único que lee la IA)
queries/<slug>/    → carpeta por consulta, con <slug>.py y <slug>.csv
```

## Cuándo usarla

- El usuario comparte un `.xlsx` / `.xlsm` o menciona un archivo Excel.
- Pide: "en la hoja X las columnas 1, 2 y 3", "filtrar", "agrupar", "unir
  hojas", "resumir", "total por cliente", "cuántos registros cumplen...".

No usarla para:

- Editar formato, estilos, colores o celdas combinadas del Excel.

## Flujo obligatorio

### 1. Inspeccionar (nunca leer el Excel a ciegas)

No abras el `.xlsx` con pandas ni vuelques celdas al chat. Primero genera el
schema:

```bash
uv run ~/.agents/skills/inspect-excel/inspect_excel.py /ruta/a/datos.xlsx
```

- Un solo Excel → escribe `excel_schema.json` junto al archivo.
- Varios Excel en un directorio → escribe `<nombre>.schema.json` por cada uno.
- Archivos grandes: añade `--max-rows 20000` para limitar el escaneo de
  estadísticas (el total de filas sigue saliendo en `rows`).

### 2. Leer el schema

Lee solo `excel_schema.json`. Identifica libro, hoja, tabla y columnas por
`pos`, `letter` y `name`. Fíjate en:

- `type` / `pandas`: tipo inferido y dtype sugerido.
- `semantic`: `currency`, `percent`, `leading_zeros`.
- `nulls`, `unique`, `enum`, `min`, `max`, `samples`.
- `warnings`: `numeric_stored_as_text`, `leading_zeros_lossy`, etc.
- `header_row` y `data_start_row`: dónde empiezan los datos.

Para ver datos rápido sin escribir un script:

```bash
uv run ~/.agents/skills/inspect-excel/inspect_excel.py preview excel_schema.json \
  --sheet Ventas --columns 1,2,3 --limit 10
```

### 3. Aclarar ambigüedades (solo lo necesario)

- ¿Qué hoja/tabla usar si hay varias?
- Fechas `dd/mm` vs `mm/dd` si el formato es ambiguo.
- Columnas de dinero: confirmar que son `currency` y moneda.
- Si hay cabecera desplazada o filas de totales: usar `--header-row` /
  `--footer-rows` y regenerar.

### 4. Generar el script de consulta

Cada consulta va en **su propia carpeta**:

```text
queries/<slug>/<slug>.py    # script
queries/<slug>/<slug>.csv   # resultado, se sobrescribe en cada ejecución
```

El `<slug>` es el nombre de la query en minúsculas con guiones bajos
(ej. "top10 vendedores por marca" → `top10_vendedores_por_marca`).

Copia `references/query_template.py` a esa carpeta y ajusta. Reglas obligatorias:

- **Encabezado PEP 723** al inicio del script (lo hace autocontenido y evita
  `ModuleNotFoundError: pandas`):

  ```python
  # /// script
  # requires-python = ">=3.10"
  # dependencies = ["pandas>=2.2", "openpyxl>=3.1"]
  # ///
  ```

- Rutas robustas con `__file__`: `HERE = Path(__file__).parent`,
  `ROOT = HERE.parent.parent`, `OUT_CSV = HERE / "<slug>.csv"`.
- Escribir el resultado con `out.to_csv(OUT_CSV, index=False)` (sobrescribe en
  cada ejecución).
- Usar `header=None` + `skiprows=header_row` + `names=[...]` para que manden las
  columnas del schema (funciona con cabeceras duplicadas o vacías).
- Usar `usecols="A:C"` (o las letras exactas) para leer solo lo pedido.
- Usar `dtype={"ID": "string"}` cuando el schema indique ceros a la izquierda.
- Usar `parse_dates=[...]` para fechas.

Ejecutarlo (sin `--with`, sin proyecto uv):

```bash
uv run queries/<slug>/<slug>.py
```

### 5. Devolver un resultado compacto

Nunca imprimas miles de filas. El script ya escribe el CSV en la carpeta de la
query. Devuelve:

- `head(n)`, `shape`, agregados (`groupby`), y
- la ruta del CSV/Parquet generado + un resumen.

### 6. Detectar cambios

Si el Excel cambió, regenera el schema. Para comprobar si está al día:

```bash
uv run ~/.agents/skills/inspect-excel/inspect_excel.py validate excel_schema.json
```

Salida `0` = al día, `2` = desactualizado.

## Comandos

```bash
inspect_excel.py datos.xlsx                      # inspecciona y escribe el schema
inspect_excel.py data/ -o excel_schema.json      # directorio, schema combinado
inspect_excel.py datos.xlsx --header-row 3 --footer-rows 1
inspect_excel.py datos.xlsx --max-rows 20000 --sample 8
inspect_excel.py validate excel_schema.json
inspect_excel.py preview excel_schema.json --sheet Ventas --columns 1,2,3 --limit 10
```

## Contrato del schema

Ver `references/schema.md` para el detalle campo a campo.

Resumen de tipos: `text`, `integer`, `number`, `boolean`, `date`, `datetime`,
`time`, `mixed`, `empty`.

## Anti-patrones

- Leer el `.xlsx` completo en el contexto (pandas, raw, base64) en vez del schema.
- Volcar miles de filas al chat.
- Ignorar `header_row` / `data_start_row` y leer con `header=0` por defecto.
- Tratar un ID con ceros a la izquierda como número.
- Modificar el Excel original.
- Re-generar el schema a cada consulta sin necesidad (usar `validate`).
- Usar `read_excel` sin `usecols` en archivos grandes.
- Generar el `.py` y el `.csv` sueltos en la raíz (siempre en
  `queries/<slug>/`).
- Quitar el encabezado PEP 723 del script (rompe `uv run` con
  `ModuleNotFoundError`).
- Inicializar un proyecto uv en el directorio del usuario: no hace falta, cada
  script es autocontenido.

## Criterio de terminado

- `excel_schema.json` generado y leído (no se abrió el Excel en el contexto).
- La consulta quedó en `queries/<slug>/<slug>.py` con su encabezado PEP 723.
- El script corre con `uv run queries/<slug>/<slug>.py` y genera
  `queries/<slug>/<slug>.csv`.
- El resultado se devolvió compacto (head/agregados + ruta del CSV).
- No se modificó el Excel original.

## Referencias

- `references/schema.md`: formato completo del schema.
- `references/recipes.md`: recetas de consulta (filtro, groupby, join, export).
- `references/query_template.py`: plantilla de script.
- `references/gotchas.md`: fechas, fórmulas, celdas combinadas, ceros a la izquierda.
