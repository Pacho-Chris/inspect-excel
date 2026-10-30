# Contrato del `excel_schema.json`

El schema es la única fuente que la IA lee sobre el Excel. No contiene los datos
completos: solo estructura, tipos y estadísticas. Está pensado para ser mínimo:
solo los campos necesarios para escribir consultas.

## Nivel raíz

```json
{
  "version": 2,
  "generated_at": "2026-10-29T00:00:00+00:00",
  "workbooks": {
    "datos.xlsx": { "...": "workbook" }
  }
}
```

- `version`: versión del formato del schema.
- `generated_at`: instante de generación (ISO 8601 UTC).
- `workbooks`: mapa `ruta_relativa -> workbook`. La ruta es relativa a la
  ubicación del archivo del schema (normalmente junto al Excel).

## Workbook

```json
{
  "size": 12345,
  "mtime": "2026-10-28T19:20:00+00:00",
  "sheet_count": 3,
  "sheets": { "Ventas": { "...": "sheet" } }
}
```

- `size` / `mtime`: firma del archivo. `validate` usa estos campos para detectar
  cambios.
- `sheet_count`: número de hojas (worksheets) del libro.
- `sheets`: mapa `nombre_hoja -> sheet`. Una hoja con varias tablas aparece
  como `"Hoja::NombreTabla"` por cada tabla.

## Sheet

```json
{
  "kind": "table",
  "table": "TablaVentas",
  "range": "A2:F12",
  "header_row": 2,
  "rows": 10,
  "sampled": false,
  "hidden": false,
  "columns": [ "...": "column" ]
}
```

- `kind`: `table` (tabla de Excel / ListObject) o `range` (hoja sin tabla).
- `table`: nombre de la tabla si `kind == "table"`.
- `range`: referencia de celdas (informativa).
- `header_row`: fila (1-based) de la cabecera. `0` = no hay cabecera. La primera
  fila de datos es `header_row + 1` (usar con `skiprows`).
- `rows`: filas de datos no vacías.
- `sampled`: `true` si `--max-rows` limitó el escaneo de estadísticas.
- `hidden`: `true` si la hoja está oculta o muy oculta.

## Column

```json
{
  "pos": 1,
  "letter": "A",
  "name": "Fecha",
  "type": "date",
  "semantic": "currency",
  "format": "dd/mm/yyyy",
  "nulls": 0,
  "errors": 2,
  "unique": 10,
  "min": "2024-01-01",
  "max": "2024-12-31",
  "decimal_places": 2,
  "samples": ["2024-01-01", "2024-01-02"],
  "enum": [{ "value": "A", "count": 40 }],
  "types": { "integer": 4, "text": 2 },
  "warnings": []
}
```

- `pos`: posición de columna en Excel (1-based). Es lo que el usuario llama
  "columna 1, 2, 3".
- `letter`: letra de columna (`A`, `B`, ...). Se usa en `usecols`.
- `name`: nombre único (cabecera o `column_A` si no hay). Se deduplica con `_2`.
  Se usa en `names=[...]`.
- `type`: ver tabla de tipos abajo.
- `semantic`: `currency`, `percent`, `leading_zeros` (solo si aplica).
- `format`: formato numérico de Excel más frecuente. Útil para desambiguar
  fechas `dd/mm` vs `mm/dd`.
- `nulls`: celdas nulas.
- `errors`: celdas de error (solo aparece si `> 0`).
- `unique`: valores distintos no nulos. `null` si se superó `--max-distinct`.
- `min` / `max`: mínimo y máximo (numéricos o fechas ISO; solo numérico/fecha).
- `decimal_places`: decimales observados en columnas `number`.
- `samples`: hasta `--sample` valores de ejemplo (JSON-safe).
- `enum`: top valores con frecuencia, solo en columnas `text` de baja
  cardinalidad (2..20 valores distintos con repeticiones).
- `types`: desglose por tipo cuando `type == "mixed"`.
- `warnings`: `numeric_stored_as_text`, `numeric_text_among_values`,
  `leading_zeros_lossy`.

Los campos opcionales se omiten del JSON si no aplican (no van como `null`).

## Tipos y dtype sugerido

El dtype de `read_excel` se deriva del `type` (ya no se guarda en el JSON):

| `type`     | Significado                      | dtype `read_excel` |
| ---------- | -------------------------------- | ------------------ |
| `text`     | texto                            | `string`          |
| `integer`  | entero                           | `Int64`           |
| `number`   | decimal                          | `float64`         |
| `boolean`  | booleano                         | `boolean`         |
| `date`     | fecha (sin hora)                 | `datetime64[ns]`  |
| `datetime` | fecha y hora                     | `datetime64[ns]`  |
| `time`     | hora                             | `object`          |
| `mixed`    | varios tipos en la misma columna | `object`          |
| `empty`    | columna sin datos                | `object`          |

## Mapping posiciones ↔ letras

El `pos` es la posición dentro de la hoja. La `letter` es la letra de Excel. Para
el script se usa la letra en `usecols="A:C"` y el `name` en `names=[...]`.
