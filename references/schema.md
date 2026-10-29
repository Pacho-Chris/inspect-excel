# Contrato del `excel_schema.json`

El schema es la única fuente que la IA lee sobre el Excel. No contiene los datos
completos: solo estructura, tipos y estadísticas.

## Nivel raíz

```json
{
  "version": 1,
  "generated_at": "2026-10-29T00:00:00+00:00",
  "workbooks": {
    "datos.xlsx": { "...": "workbook" }
  }
}
```

- `version`: versión del formato del schema.
- `generated_at`: instante de generación (ISO 8601 UTC).
- `workbooks`: mapa `ruta_relativa -> workbook`. La ruta es relativa a la
  ubicación del archivo del schema (normalmente `datos.xlsx`).

## Workbook

```json
{
  "size": 12345,
  "mtime": "2026-10-28T19:20:00+00:00",
  "sha256": "abc...",
  "sheets": { "Ventas": { "...": "sheet" } }
}
```

- `size` / `mtime` / `sha256`: firma del archivo. `sha256` se omite si se usó
  `--no-hash`. `validate` usa estos campos para detectar cambios.
- `sheets`: mapa `nombre_hoja -> sheet`. Una hoja con varias tablas aparece
  como `"Hoja::NombreTabla"` por cada tabla.

## Sheet

```json
{
  "kind": "table",
  "table": "TablaVentas",
  "range": "A2:F12",
  "header_row": 2,
  "data_start_row": 3,
  "rows": 10,
  "nonempty_rows": 10,
  "sampled": false,
  "scanned_rows": 10,
  "hidden": false,
  "columns": [ "...": "column" ]
}
```

- `kind`: `table` (tabla de Excel / ListObject) o `range` (hoja sin tabla).
- `table`: nombre de la tabla si `kind == "table"`.
- `range`: referencia de celdas (informativa).
- `header_row`: fila (1-based) de la cabecera. `0` = no hay cabecera.
- `data_start_row`: primera fila de datos (1-based). Usar con `skiprows`.
- `rows`: filas de datos no vacías.
- `nonempty_rows`: ídem (equivale a `rows`; se mantiene por claridad).
- `sampled`: `true` si `--max-rows` limitó el escaneo de estadísticas.
- `scanned_rows`: filas realmente usadas para las estadísticas.
- `hidden`: `true` si la hoja está oculta o muy oculta.

## Column

```json
{
  "pos": 1,
  "letter": "A",
  "name": "Fecha",
  "type": "date",
  "pandas": "datetime64[ns]",
  "semantic": null,
  "format": "dd/mm/yyyy",
  "nulls": 0,
  "errors": 0,
  "count": 10,
  "unique": 10,
  "unique_capped": false,
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
- `letter`: letra de columna (`A`, `B`, ...).
- `name`: nombre único (cabecera o `column_A` si no hay). Se deduplica con `_2`.
- `type`: ver tabla de tipos abajo.
- `pandas`: dtype sugerido para `read_excel`.
- `semantic`: `currency`, `percent`, `leading_zeros` (opcional).
- `format`: formato numérico de Excel más frecuente.
- `nulls` / `errors` / `count`: celdas nulas / de error / no nulas.
- `unique`: valores distintos no nulos. `null` + `unique_capped: true` si se
  superó `--max-distinct`.
- `min` / `max`: mínimo y máximo (numéricos o fechas ISO).
- `decimal_places`: decimales observados en columnas `number`.
- `samples`: hasta `--sample` valores de ejemplo (JSON-safe).
- `enum`: top valores con frecuencia, solo en columnas `text` de baja
  cardinalidad (2..20 valores distintos con repeticiones).
- `types`: desglose por tipo cuando `type == "mixed"`.
- `warnings`: `numeric_stored_as_text`, `numeric_text_among_values`,
  `leading_zeros_lossy`.

## Tipos

| `type`     | Significado                      | `pandas`         |
| ---------- | -------------------------------- | ---------------- |
| `text`     | texto                            | `string`         |
| `integer`  | entero                           | `Int64`          |
| `number`   | decimal                          | `float64`        |
| `boolean`  | booleano                         | `boolean`        |
| `date`     | fecha (sin hora)                 | `datetime64[ns]` |
| `datetime` | fecha y hora                     | `datetime64[ns]` |
| `time`     | hora                             | `object`         |
| `mixed`    | varios tipos en la misma columna | `object`         |
| `empty`    | columna sin datos                | `object`         |

## Mapping posiciones ↔ letras

El `pos` es la posición dentro de la
hoja. La `letter` es la letra de Excel. Para el script se usa la letra en
`usecols="A:C"` y el `name` en `names=[...]`.
