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
queries/<slug>.py  → script por consulta; imprime el resultado como JSON a stdout
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

- `type`: tipo inferido (el dtype de `read_excel` se deriva de la tabla de tipos).
- `semantic`: `currency`, `percent`, `leading_zeros`.
- `nulls`, `unique`, `enum`, `min`, `max`, `samples`.
- `warnings`: `numeric_stored_as_text`, `leading_zeros_lossy`, etc.
- `header_row`: dónde empieza la cabecera (los datos empiezan en `header_row + 1`).

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

Cada consulta es un único script plano:

```text
queries/<slug>.py   # imprime el resultado como JSON a stdout
```

El `<slug>` es el nombre de la query en minúsculas con guiones bajos
(ej. "top10 vendedores por marca" → `top10_vendedores_por_marca`).

Copia `references/query_template.py` a `queries/<slug>.py` y ajusta. Reglas
obligatorias:

- **Encabezado PEP 723** al inicio del script (lo hace autocontenido y evita
  `ModuleNotFoundError: pandas`):

  ```python
  # /// script
  # requires-python = ">=3.10"
  # dependencies = ["pandas>=2.2", "openpyxl>=3.1"]
  # ///
  ```

- Rutas robustas con `__file__`: `HERE = Path(__file__).parent` (queries/),
  `ROOT = HERE.parent`.
- **Emitir el resultado como JSON a stdout**, y nada más:
  `print(out.to_json(orient="records", date_format="iso", force_ascii=False))`.
  Cualquier nota va a `stderr`.
- Usar `header=None` + `skiprows=header_row` + `names=[...]` para que manden las
  columnas del schema (funciona con cabeceras duplicadas o vacías).
- Usar `usecols="A:C"` (o las letras exactas) para leer solo lo pedido.
- Usar `dtype={"ID": "string"}` cuando el schema indique ceros a la izquierda.
- Usar `parse_dates=[...]` para fechas.

Ejecutarlo (sin `--with`, sin proyecto uv):

```bash
uv run queries/<slug>.py
```

Para encadenar o guardar (lo decide el usuario):

```bash
uv run queries/a.py | uv run queries/b.py   # una query consume el JSON de otra
uv run queries/<slug>.py > resultado.json   # guardar a archivo
```

### 5. Devolver un resultado compacto

Ejecuta el script, captura el JSON de stdout y devuelve un resumen compacto:
`head(n)`, `shape`, agregados (`groupby`) y el JSON (o un recorte si es largo).
Nunca imprimas miles de filas ni JSON gigante en el chat.

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
- Mezclar notas o logs en stdout (rompe el JSON; deben ir a `stderr`).
- Ignorar `header_row` y leer con `header=0` por defecto.
- Tratar un ID con ceros a la izquierda como número.
- Modificar el Excel original.
- Re-generar el schema a cada consulta sin necesidad (usar `validate`).
- Usar `read_excel` sin `usecols` en archivos grandes.
- Generar el `.py` suelto en la raíz (siempre en `queries/<slug>.py`).
- Quitar el encabezado PEP 723 del script (rompe `uv run` con
  `ModuleNotFoundError`).
- Inicializar un proyecto uv en el directorio del usuario: no hace falta, cada
  script es autocontenido.

## Criterio de terminado

- `excel_schema.json` generado y leído (no se abrió el Excel en el contexto).
- La consulta quedó en `queries/<slug>.py` con su encabezado PEP 723.
- El script corre con `uv run queries/<slug>.py` e imprime JSON válido a stdout
  (nada más en stdout).
- El resultado se devolvió compacto (head/agregados + el JSON).
- No se modificó el Excel original.

## Referencias

- `references/schema.md`: formato completo del schema.
- `references/recipes.md`: recetas de consulta (filtro, groupby, join, emitir JSON, encadenar).
- `references/query_template.py`: plantilla de script.
- `references/gotchas.md`: fechas, fórmulas, celdas combinadas, ceros a la izquierda.
