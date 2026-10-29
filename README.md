# inspect-excel

Skill para consultar archivos Excel (`.xlsx` / `.xlsm`) **sin abrirlos en el
contexto de la IA**. Un CLI inspecciona el archivo, detecta hojas y tablas de
Excel, infiere tipos de datos y genera un `excel_schema.json`. Después la IA
lee solo ese schema y escribe scripts mínimos de `pandas` para responder
consultas.

```text
inspect_excel.py   → inspecciona el Excel y escribe excel_schema.json (se conserva)
excel_schema.json  → contrato de metadatos (lo único que lee la IA)
queries/<slug>/    → carpeta por consulta, con <slug>.py y <slug>.csv
```

## Requisitos

- [uv](https://docs.astral.sh/uv/) (instala Python y las dependencias solo).
- Python 3.10+ (lo gestiona `uv` automáticamente).

No necesitas instalar `pandas` ni `openpyxl` manualmente: cada script declara
sus dependencias en un encabezado PEP 723 y `uv` se encarga.

## Instalación

Copia (o crea un enlace a) la carpeta `inspect-excel/` dentro de tu directorio
de skills del agente:

```bash
# OpenCode / Claude Code / otros que usen ~/.agents/skills
ln -s "$(pwd)/inspect-excel" ~/.agents/skills/inspect-excel
```

La skill queda disponible como `inspect-excel`.

## Uso

Con un Excel en el directorio de trabajo, pídele a la IA, por ejemplo:

> "Mira ventas.xlsx y en la hoja Ventas necesito las columnas 1, 2 y 3 con Monto > 1000"

La IA:

1. Ejecuta el inspector:
   ```bash
   uv run ~/.agents/skills/inspect-excel/inspect_excel.py ventas.xlsx
   ```
2. Lee solo el `excel_schema.json` generado.
3. Escribe `queries/<slug>/<slug>.py` (autocontenido) y lo ejecuta:
   ```bash
   uv run queries/<slug>/<slug>.py
   ```
4. Genera `queries/<slug>/<slug>.csv` (se sobrescribe en cada ejecución) y
   devuelve un resumen compacto.

## CLI

```bash
uv run inspect_excel.py datos.xlsx                      # inspecciona y escribe el schema
uv run inspect_excel.py data/ -o excel_schema.json      # directorio, schema combinado
uv run inspect_excel.py datos.xlsx --header-row 3 --footer-rows 1
uv run inspect_excel.py datos.xlsx --max-rows 20000 --sample 8
uv run inspect_excel.py validate excel_schema.json      # ¿sigue al día? (0 = sí, 2 = no)
uv run inspect_excel.py preview excel_schema.json --sheet Ventas --columns 1,2,3 --limit 10
```

## Tests

```bash
uv run --with pytest --with openpyxl pytest -q
```

## Documentación

- `SKILL.md` — flujo obligatorio para el agente.
- `references/schema.md` — formato completo del `excel_schema.json`.
- `references/recipes.md` — recetas de consulta (filtro, groupby, join, export).
- `references/query_template.py` — plantilla de script.
- `references/gotchas.md` — fechas, fórmulas, celdas combinadas, ceros a la izquierda.

## Licencia

MIT. Ver `LICENSE`.
