#!/usr/bin/env python3
# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1"]
# ///
"""inspect_excel.py — Inspecciona archivos Excel (.xlsx/.xlsm), detecta hojas,
tablas de Excel (ListObjects), filas de cabecera e infiere tipos de datos, y
genera un esquema JSON compacto para que una IA consulte los datos sin abrir el
Excel ni cargarlo en el contexto.

Uso:
  uv run inspect_excel.py datos.xlsx
  uv run inspect_excel.py data/ --output excel_schema.json
  uv run inspect_excel.py validate datos.schema.json
  uv run inspect_excel.py preview datos.schema.json --sheet Ventas --columns 1,2,3 --limit 10
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import date, datetime, time, timezone
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries

SCHEMA_VERSION = 1
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
EXCEL_EXTS = (".xlsx", ".xlsm")
UNSUPPORTED_EXTS = (".xls", ".xlsb", ".xltx", ".xltm", ".ods", ".csv")


# --------------------------------------------------------------------------- #
# Utilidades
# --------------------------------------------------------------------------- #

def local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def iso(value) -> str:
    if isinstance(value, datetime):
        if value.tzinfo is not None:
            value = value.astimezone(timezone.utc)
        if value.time() == time(0, 0, 0):
            return value.date().isoformat()
        return value.isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, time):
        return value.isoformat()
    return value


def iso_mtime(st) -> str:
    return datetime.fromtimestamp(st.st_mtime, timezone.utc).isoformat()


def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def any_nonempty(vals) -> bool:
    return any(v is not None and (not isinstance(v, str) or v.strip() != "") for v in vals)


def to_sample(value):
    if isinstance(value, (datetime, date, time)):
        return iso(value)
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, str):
        return value[:200]
    return str(value)


_NUM_LIKE = re.compile(r"^[+-]?\s*[0-9][0-9 ,.]*$")


def looks_numeric(s: str) -> bool:
    return bool(_NUM_LIKE.match(s.strip()))


def decimal_places(value) -> int:
    s = repr(float(value))
    if "e" in s or "E" in s:
        return 0
    if "." in s:
        return min(10, len(s.split(".")[1].rstrip("0")))
    return 0


def semantic_of(fmt: str):
    if not fmt:
        return None
    if re.search(r"[$€£¥]", fmt) or "[$" in fmt or "S/" in fmt:
        return "currency"
    if re.search(r"(^|\s)Q(\s|$)", fmt):
        return "currency"
    if "%" in fmt:
        return "percent"
    if re.fullmatch(r"[0#?]+", fmt) and "0" in fmt:
        return "leading_zeros"
    return None


# --------------------------------------------------------------------------- #
# Lectura de metadatos (tablas, estado de hoja) desde el XML del paquete
# --------------------------------------------------------------------------- #

def _resolve_part(target: str, base_dir: str = "xl") -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base_dir, target))


def read_workbook_xml(path: Path) -> dict:
    """Devuelve {sheet_name: {"state": str, "tables": [table_info, ...]}}."""
    info: dict = {}
    try:
        zf = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError):
        return info
    with zf:
        names = set(zf.namelist())
        try:
            wb = ET.fromstring(zf.read("xl/workbook.xml"))
        except (KeyError, ET.ParseError):
            return info

        rid2target: dict = {}
        try:
            rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            for rel in rels:
                if local(rel.tag) == "Relationship":
                    rid2target[rel.get("Id")] = rel.get("Target")
        except (KeyError, ET.ParseError):
            pass

        for sheet in wb.findall(f".//{NS}sheet"):
            name = sheet.get("name")
            entry = {"state": sheet.get("state", "visible"), "tables": []}
            rid = sheet.get(R_NS + "id")
            if rid is None:
                rid = sheet.get("id")
            target = rid2target.get(rid)
            if target:
                sheet_path = _resolve_part(target, "xl")
                rels_path = posixpath.join(
                    posixpath.dirname(sheet_path),
                    "_rels",
                    posixpath.basename(sheet_path) + ".rels",
                )
                if rels_path in names:
                    try:
                        srels = ET.fromstring(zf.read(rels_path))
                        for rel in srels:
                            if local(rel.tag) != "Relationship":
                                continue
                            if not rel.get("Type", "").endswith("/table"):
                                continue
                            ttarget = rel.get("Target")
                            tpath = _resolve_part(ttarget, posixpath.dirname(sheet_path))
                            if tpath not in names:
                                continue
                            try:
                                t = ET.fromstring(zf.read(tpath))
                            except ET.ParseError:
                                continue
                            cols = []
                            for tc in t.findall(f".//{NS}tableColumn"):
                                cols.append({
                                    "name": tc.get("name"),
                                    "totalsRowLabel": tc.get("totalsRowLabel"),
                                })
                            entry["tables"].append({
                                "name": t.get("name") or t.get("displayName"),
                                "displayName": t.get("displayName") or t.get("name"),
                                "ref": t.get("ref"),
                                "headerRowCount": int(t.get("headerRowCount") or 1),
                                "totalsRowCount": int(t.get("totalsRowCount") or 0),
                                "columns": cols,
                            })
                    except (ET.ParseError, KeyError):
                        pass
            info[name] = entry
    return info


# --------------------------------------------------------------------------- #
# Clasificación de celdas e inferencia
# --------------------------------------------------------------------------- #

def classify_cell(v, cell) -> str:
    if v is None:
        return "null"
    dt = getattr(cell, "data_type", None)
    nf = getattr(cell, "number_format", None) or ""
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return "null"
        if dt == "e" or (s.startswith("#") and re.match(r"^#[A-Z0-9_/]+[!?]$", s)):
            return "error"
        return "text"
    if dt == "e":
        return "error"
    if isinstance(v, bool):
        return "boolean"
    if isinstance(v, datetime):
        if v.time() == time(0, 0, 0) and not re.search(r"[hHsS]", nf):
            return "date"
        return "datetime"
    if isinstance(v, date):
        return "date"
    if isinstance(v, time):
        return "time"
    if isinstance(v, int):
        return "integer"
    if isinstance(v, float):
        return "number"
    return "other"


def new_acc() -> dict:
    return {
        "count": 0, "nulls": 0, "errors": 0,
        "types": Counter(),
        "uniq": set(), "uniq_capped": False,
        "enum": Counter(), "enum_capped": False,
        "samples": [],
        "nmin": None, "nmax": None,
        "dmin": None, "dmax": None,
        "formats": Counter(),
        "decimals": 0,
        "numeric_text": 0,
        "warnings": set(),
    }


def add_value(acc, kind, v, cell, args):
    if kind == "null":
        acc["nulls"] += 1
        return
    if kind == "error":
        acc["errors"] += 1
        return
    acc["count"] += 1
    acc["types"][kind] += 1

    if len(acc["samples"]) < args.sample:
        acc["samples"].append(to_sample(v))

    if not acc["uniq_capped"]:
        if len(acc["uniq"]) >= args.max_distinct:
            acc["uniq_capped"] = True
        else:
            acc["uniq"].add(v)

    if not acc["enum_capped"]:
        if len(acc["enum"]) >= args.max_distinct:
            acc["enum_capped"] = True
        else:
            acc["enum"][v] += 1

    if kind in ("integer", "number"):
        if acc["nmin"] is None or v < acc["nmin"]:
            acc["nmin"] = v
        if acc["nmax"] is None or v > acc["nmax"]:
            acc["nmax"] = v
        if kind == "number":
            acc["decimals"] = max(acc["decimals"], decimal_places(v))
    elif kind in ("date", "datetime"):
        if acc["dmin"] is None or v < acc["dmin"]:
            acc["dmin"] = v
        if acc["dmax"] is None or v > acc["dmax"]:
            acc["dmax"] = v
    elif kind == "text":
        if looks_numeric(v):
            acc["numeric_text"] += 1

    nf = getattr(cell, "number_format", None)
    if nf and nf != "General":
        acc["formats"][nf] += 1


def consolidate(acc, args) -> dict:
    col = {
        "type": None,
        "pandas": None,
        "nulls": acc["nulls"],
        "errors": acc["errors"],
        "count": acc["count"],
        "unique": None,
        "unique_capped": acc["uniq_capped"],
        "samples": acc["samples"],
        "warnings": sorted(acc["warnings"]),
    }
    non_null = acc["count"]
    kinds = set(acc["types"])

    if non_null == 0:
        col["type"] = "empty"
        col["pandas"] = "object"
    elif kinds <= {"integer", "number"}:
        if kinds == {"integer"}:
            col["type"] = "integer"
            col["pandas"] = "Int64"
        else:
            col["type"] = "number"
            col["pandas"] = "float64"
            col["decimal_places"] = acc["decimals"]
        col["min"] = acc["nmin"]
        col["max"] = acc["nmax"]
    elif kinds <= {"date", "datetime"}:
        col["type"] = "datetime" if "datetime" in kinds else "date"
        col["pandas"] = "datetime64[ns]"
        col["min"] = iso(acc["dmin"])
        col["max"] = iso(acc["dmax"])
    elif kinds == {"boolean"}:
        col["type"] = "boolean"
        col["pandas"] = "boolean"
    elif kinds == {"time"}:
        col["type"] = "time"
        col["pandas"] = "object"
    elif kinds == {"text"}:
        col["type"] = "text"
        col["pandas"] = "string"
        if acc["numeric_text"] >= non_null:
            col["warnings"].append("numeric_stored_as_text")
    else:
        col["type"] = "mixed"
        col["pandas"] = "object"
        col["types"] = dict(acc["types"])
        if acc["numeric_text"] > 0:
            col["warnings"].append("numeric_text_among_values")

    if acc["formats"]:
        fmt = acc["formats"].most_common(1)[0][0]
        col["format"] = fmt
        sem = semantic_of(fmt)
        if sem:
            col["semantic"] = sem
            if sem == "leading_zeros" and col["type"] in ("integer", "number"):
                col["warnings"].append("leading_zeros_lossy")

    if not acc["uniq_capped"]:
        col["unique"] = len(acc["uniq"])

    if (
        col["type"] == "text"
        and not acc["uniq_capped"]
        and 2 <= len(acc["uniq"]) <= 20
        and len(acc["uniq"]) < non_null
    ):
        col["enum"] = [
            {"value": to_sample(v), "count": acc["enum"].get(v, 0)}
            for v, _ in acc["enum"].most_common(20)
        ]

    # Compactar: quitar claves con valor None (excepto warnings, siempre presente)
    return {k: v for k, v in col.items() if v is not None or k == "warnings"}


def make_names(header_vals, ncols, xml_names) -> list:
    names = []
    used = {}
    for i in range(ncols):
        base = None
        if xml_names and i < len(xml_names) and xml_names[i]:
            base = str(xml_names[i]).strip()
        if not base and header_vals and i < len(header_vals):
            hv = header_vals[i]
            if hv is not None and (not isinstance(hv, str) or hv.strip() != ""):
                base = str(hv).strip()
        if not base:
            base = "column_" + get_column_letter(i + 1)
        key = base
        if key in used:
            used[key] += 1
            base = f"{key}_{used[key]}"
        else:
            used[key] = 0
        names.append(base)
    return names


def build_column(acc, name, excel_col, args) -> dict:
    base = consolidate(acc, args)
    return {
        "pos": excel_col,
        "letter": get_column_letter(excel_col),
        "name": name,
        **base,
    }


def trim_trailing(columns: list) -> list:
    while columns and columns[-1]["type"] == "empty" and columns[-1]["name"].startswith("column_"):
        columns.pop()
    return columns


def detect_header(buffer, max_col) -> int:
    """Devuelve la fila (1-based) que parece cabecera, o 0 si no hay."""
    best = 0
    best_score = float("-inf")
    for ri, cells in buffer:
        vals = [c.value for c in cells]
        nonempty = [v for v in vals if v is not None and (not isinstance(v, str) or v.strip() != "")]
        if len(nonempty) < 2:
            continue
        n_text = sum(1 for v in nonempty if isinstance(v, str))
        n_num = sum(1 for v in nonempty if isinstance(v, (int, float, datetime, date)))
        if n_text == 0:
            # una fila sin ninguna etiqueta de texto no puede ser cabecera
            continue
        uniq = len({to_sample(v) for v in nonempty})
        ratio = uniq / len(nonempty)
        score = len(nonempty) * 2 + n_text * 3 + ratio * 5
        if n_num > n_text:
            score -= 10
        if score > best_score:
            best_score = score
            best = ri
    return best


# --------------------------------------------------------------------------- #
# Inspección de hojas
# --------------------------------------------------------------------------- #

def inspect_table(ws, table, args) -> dict:
    min_col, min_row, max_col, max_row = range_boundaries(table["ref"])
    hrc = int(table["headerRowCount"] or 1)
    totals = int(table["totalsRowCount"] or 0)
    header_row = min_row + hrc - 1 if hrc > 0 else 0
    data_start = min_row + hrc
    data_end = max_row - totals
    ncols = max_col - min_col + 1

    header_vals = [None] * ncols
    xml_names = [c["name"] for c in table.get("columns", [])]
    while len(xml_names) < ncols:
        xml_names.append(None)

    accs = [new_acc() for _ in range(ncols)]
    nonempty_rows = 0
    scanned = 0
    sampled = False

    for ri, cells in enumerate(
        ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col),
        start=min_row,
    ):
        vals = [c.value for c in cells]
        if ri == header_row:
            header_vals = vals
            continue
        if ri < data_start or ri > data_end:
            continue
        if not any_nonempty(vals):
            continue
        nonempty_rows += 1
        if args.max_rows and scanned >= args.max_rows:
            sampled = True
            continue
        scanned += 1
        for i in range(ncols):
            v = vals[i] if i < len(vals) else None
            c = cells[i] if i < len(cells) else None
            add_value(accs[i], classify_cell(v, c), v, c, args)

    names = make_names(header_vals, ncols, xml_names)
    columns = trim_trailing([build_column(accs[i], names[i], min_col + i, args) for i in range(ncols)])

    return {
        "kind": "table",
        "table": table["displayName"] or table["name"],
        "range": table["ref"],
        "header_row": header_row,
        "data_start_row": data_start,
        "rows": nonempty_rows,
        "nonempty_rows": nonempty_rows,
        "sampled": sampled,
        "scanned_rows": scanned,
        "columns": columns,
    }


def inspect_range(ws, args) -> dict:
    max_col = ws.max_column
    max_row = ws.max_row
    scan = args.header_scan
    header_row = args.header_row  # None => detección automática
    header_vals = []
    ncols = max_col if max_col else 0
    accs = []
    buffer = []
    nonempty_rows = 0
    scanned = 0
    sampled = False
    data_end = (max_row - args.footer_rows) if max_row else None

    def finalize_header():
        nonlocal ncols, accs, header_vals
        if not ncols:
            if max_col:
                ncols = max_col
            elif header_row and header_vals:
                ncols = len(header_vals)
            else:
                ncols = max((len(cells) for _, cells in buffer), default=0)
        if not header_vals and header_row:
            for ri, cells in buffer:
                if ri == header_row:
                    header_vals = [c.value for c in cells]
                    break
        while len(header_vals) < ncols:
            header_vals.append(None)
        accs = [new_acc() for _ in range(ncols)]

    def process(cells):
        nonlocal nonempty_rows, scanned, sampled
        vals = [c.value for c in cells]
        if not any_nonempty(vals):
            return
        nonempty_rows += 1
        if args.max_rows and scanned >= args.max_rows:
            sampled = True
            return
        scanned += 1
        for i in range(ncols):
            v = vals[i] if i < len(vals) else None
            c = cells[i] if i < len(cells) else None
            add_value(accs[i], classify_cell(v, c), v, c, args)

    it = ws.iter_rows(min_row=1, min_col=1, max_col=max_col or None)
    for ri, cells in enumerate(it, start=1):
        if header_row is None:
            buffer.append((ri, cells))
            if ri >= scan:
                header_row = detect_header(buffer, max_col)
                finalize_header()
                for bri, bcells in buffer:
                    if header_row == 0 or bri > header_row:
                        if data_end is None or bri <= data_end:
                            process(bcells)
            continue
        if ri == header_row:
            header_vals = [c.value for c in cells]
            continue
        if ri < header_row:
            continue
        if not accs:
            finalize_header()
        if data_end is not None and ri > data_end:
            continue
        process(cells)

    if header_row is None:
        header_row = detect_header(buffer, max_col)
        finalize_header()
        for bri, bcells in buffer:
            if header_row == 0 or bri > header_row:
                if data_end is None or bri <= data_end:
                    process(bcells)

    if not accs:
        finalize_header()

    data_start = header_row + 1 if header_row else 1

    names = make_names(header_vals, ncols, [])
    columns = trim_trailing([build_column(accs[i], names[i], i + 1, args) for i in range(ncols)])
    ncols = len(columns)

    rng = f"A1:{get_column_letter(ncols)}{max_row or nonempty_rows}" if ncols else "A1:A1"
    return {
        "kind": "range",
        "table": None,
        "range": rng,
        "header_row": header_row,
        "data_start_row": data_start,
        "rows": nonempty_rows,
        "nonempty_rows": nonempty_rows,
        "sampled": sampled,
        "scanned_rows": scanned,
        "columns": columns,
    }


def inspect_workbook(path: Path, args) -> dict:
    xml = read_workbook_xml(path)
    wb = load_workbook(path, read_only=True, data_only=True)
    sheets = {}
    try:
        for ws in wb.worksheets:
            name = ws.title
            state = xml.get(name, {}).get("state", getattr(ws, "sheet_state", "visible"))
            hidden = state != "visible"
            tables = [] if args.no_table else xml.get(name, {}).get("tables", [])
            if args.sheet and name not in args.sheet:
                continue
            if tables:
                valid = [t for t in tables if t.get("ref")]
                if len(valid) == 1:
                    meta = inspect_table(ws, valid[0], args)
                    meta["hidden"] = hidden
                    sheets[name] = meta
                else:
                    for t in valid:
                        key = f"{name}::{t.get('displayName') or t.get('name') or 'Table'}"
                        meta = inspect_table(ws, t, args)
                        meta["hidden"] = hidden
                        meta["sheet"] = name
                        sheets[key] = meta
            else:
                meta = inspect_range(ws, args)
                meta["hidden"] = hidden
                sheets[name] = meta
    finally:
        wb.close()
    return sheets


def inspect_file(path: Path, args) -> dict:
    st = path.stat()
    entry = {
        "size": st.st_size,
        "mtime": iso_mtime(st),
        "sheets": inspect_workbook(path, args),
    }
    if not args.no_hash:
        entry["sha256"] = file_sha256(path)
    return entry


# --------------------------------------------------------------------------- #
# Salida de schema y resumen
# --------------------------------------------------------------------------- #

def write_schema(out: Path, workbooks: dict):
    doc = {
        "version": SCHEMA_VERSION,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "workbooks": workbooks,
    }
    with open(out, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, indent=2)


def print_summary(entry: dict):
    for sname, meta in entry["sheets"].items():
        kind = meta.get("kind")
        label = f"[tabla {meta.get('table')}]" if kind == "table" else "[rango]"
        hid = " (oculta)" if meta.get("hidden") else ""
        print(f"   {sname} {label}{hid} · {meta.get('rows', 0)} filas · {len(meta.get('columns', []))} cols")
        for c in meta.get("columns", []):
            warns = f"  ⚠ {', '.join(c['warnings'])}" if c.get("warnings") else ""
            sem = f" [{c['semantic']}]" if c.get("semantic") else ""
            nulls = c.get("nulls", 0)
            print(f"      {c['pos']:>3} {c['letter']:>2}  {c['name']:<20} {c['type']:<10}{sem} nulls={nulls}{warns}")


def collect_files(paths) -> list:
    files = []
    for p in paths:
        pp = Path(p)
        if pp.is_dir():
            found = []
            for g in ("*.xlsx", "*.xlsm"):
                found.extend(pp.glob(g))
            found = sorted(found, key=lambda f: f.name.lower())
            if not found:
                print(f"[Aviso] Sin archivos .xlsx/.xlsm en '{pp}'.")
            files.extend(found)
        elif pp.is_file():
            ext = pp.suffix.lower()
            if ext in EXCEL_EXTS:
                files.append(pp)
            elif ext in UNSUPPORTED_EXTS:
                print(f"[Aviso] Formato no soportado: {pp.name}. Convierte a .xlsx.")
            else:
                print(f"[Aviso] Ignorado (no es Excel): {pp.name}.")
        else:
            print(f"[Error] No existe: {pp}")
    # dedupe conservando orden
    seen = set()
    out = []
    for f in files:
        rp = f.resolve()
        if rp not in seen:
            seen.add(rp)
            out.append(f)
    return out


# --------------------------------------------------------------------------- #
# Comandos
# --------------------------------------------------------------------------- #

def run_inspect(args) -> int:
    files = collect_files(args.paths)
    if not files:
        print("[Error] No se encontraron archivos Excel (.xlsx/.xlsm).")
        return 1

    if args.output:
        out = Path(args.output)
        out.parent.mkdir(parents=True, exist_ok=True)
        workbooks = {}
        for f in files:
            print(f"-> {f.name}")
            entry = inspect_file(f, args)
            key = os.path.relpath(f, out.parent)
            workbooks[key] = entry
            print_summary(entry)
        write_schema(out, workbooks)
        print(f"\nEsquema: {out.resolve()}")
    else:
        for f in files:
            print(f"-> {f.name}")
            entry = inspect_file(f, args)
            if len(files) == 1:
                out = f.with_name("excel_schema.json")
            else:
                out = f.with_name(f.stem + ".schema.json")
            write_schema(out, {f.name: entry})
            print_summary(entry)
            print(f"   Esquema: {out}")
    return 0


def run_validate(args) -> int:
    schemas = []
    for p in args.paths:
        pp = Path(p)
        if pp.is_dir():
            schemas.extend(sorted(pp.glob("excel_schema.json")))
            schemas.extend(sorted(pp.glob("*.schema.json")))
        elif pp.exists():
            schemas.append(pp)
        else:
            print(f"[Error] No existe: {p}")
            return 1
    if not schemas:
        print("[Error] No se encontraron schemas JSON.")
        return 1

    problems = []
    ok = 0
    for sp in schemas:
        try:
            data = json.loads(sp.read_text(encoding="utf-8"))
        except Exception as e:
            problems.append(f"{sp.name}: JSON inválido ({e})")
            continue
        for key, wb in data.get("workbooks", {}).items():
            fp = (sp.parent / key).resolve()
            if not fp.exists():
                problems.append(f"{key}: archivo no encontrado")
                continue
            st = fp.stat()
            if wb.get("size") is not None and st.st_size != wb["size"]:
                problems.append(f"{key}: tamaño cambió")
                continue
            if wb.get("mtime") and iso_mtime(st) != wb["mtime"]:
                problems.append(f"{key}: modificado")
                continue
            if wb.get("sha256") and file_sha256(fp) != wb["sha256"]:
                problems.append(f"{key}: contenido cambió")
                continue
            ok += 1
            print(f"✓ {key}")
    if problems:
        print("\nDesactualizado:")
        for p in problems:
            print(f"  ✗ {p}")
        return 2
    print(f"\n{ok} archivo(s) al día.")
    return 0


def resolve_columns(meta, spec) -> list:
    cols = meta.get("columns", [])
    if not spec:
        return cols
    by_letter = {c["letter"]: c for c in cols}
    by_pos = {str(c["pos"]): c for c in cols}
    by_name = {c["name"]: c for c in cols}
    result = []
    seen = set()
    for token in spec.split(","):
        t = token.strip()
        if not t:
            continue
        if re.fullmatch(r"[A-Za-z]+:[A-Za-z]+", t):
            a, b = t.split(":")
            ai, bi = column_index_from_string(a.upper()), column_index_from_string(b.upper())
            lo, hi = sorted((ai, bi))
            for i in range(lo, hi + 1):
                c = by_letter.get(get_column_letter(i))
                if c and id(c) not in seen:
                    seen.add(id(c))
                    result.append(c)
            continue
        tup = t.upper()
        if tup in by_letter:
            c = by_letter[tup]
        elif t in by_pos:
            c = by_pos[t]
        elif t in by_name:
            c = by_name[t]
        else:
            print(f"[Aviso] Columna '{t}' no encontrada; se omite.")
            continue
        if id(c) not in seen:
            seen.add(id(c))
            result.append(c)
    return result


def format_value(v) -> str:
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.isoformat(sep=" ") if v.time() != time(0, 0, 0) else v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, time):
        return v.isoformat()
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, float):
        if v == int(v) and abs(v) < 1e15:
            return str(int(v))
        return str(v)
    return str(v)


def run_preview(args) -> int:
    sp = Path(args.schema)
    try:
        data = json.loads(sp.read_text(encoding="utf-8"))
    except Exception as e:
        print(f"[Error] No se pudo leer el schema: {e}")
        return 1

    target = None
    workbook_path = None
    actual_sheet = None
    for key, wb in data.get("workbooks", {}).items():
        for sname, meta in wb.get("sheets", {}).items():
            base = sname.split("::")[-1]
            if base == args.sheet:
                target = meta
                workbook_path = (sp.parent / key).resolve()
                actual_sheet = sname.split("::")[0] if "::" in sname else sname
                break
        if target:
            break
    if not target:
        print(f"[Error] Hoja '{args.sheet}' no encontrada en el schema.")
        return 1

    columns = resolve_columns(target, args.columns)
    if not columns:
        print("[Error] Sin columnas para mostrar.")
        return 1

    wb = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        ws = wb[actual_sheet]
    except KeyError:
        print(f"[Error] Hoja '{actual_sheet}' no existe en {workbook_path.name}.")
        return 1

    col_idxs = [column_index_from_string(c["letter"]) for c in columns]
    min_c, max_c = min(col_idxs), max(col_idxs)
    start = target.get("data_start_row", 1)

    rows = [[c["name"] for c in columns]]
    printed = 0
    for cells in ws.iter_rows(min_row=start, min_col=min_c, max_col=max_c):
        row = [None] * (max_c - min_c + 1)
        for cell in cells:
            row[cell.column - min_c] = cell.value
        selected = [format_value(row[idx - min_c]) for idx in col_idxs]
        if not any(s for s in selected):
            continue
        rows.append(selected)
        printed += 1
        if printed >= args.limit:
            break
    wb.close()

    widths = [0] * len(rows[0])
    for r in rows:
        for i, cell in enumerate(r):
            widths[i] = max(widths[i], len(cell))
    for r in rows:
        print("  ".join(cell.ljust(widths[i]) for i, cell in enumerate(r)).rstrip())
    print(f"\n{printed} fila(s) mostradas.")
    return 0


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="inspect_excel.py",
        description="Inspecciona Excel (.xlsx/.xlsm) y genera un schema JSON para consultas por IA.",
    )
    sub = p.add_subparsers(dest="command")

    ip = sub.add_parser("inspect", help="Genera el schema de uno o más Excel")
    ip.add_argument("paths", nargs="+", help="Archivos .xlsx/.xlsm o directorios")
    ip.add_argument("--output", "-o", help="Escribe un único schema combinado aquí")
    ip.add_argument("--header-row", type=int, help="Forzar fila de cabecera (1-based)")
    ip.add_argument("--footer-rows", type=int, default=0, help="Filas de totales al final a excluir")
    ip.add_argument("--max-rows", type=int, default=0, help="Máx. filas de datos a escanear (0 = todas)")
    ip.add_argument("--sample", type=int, default=5, help="Nº de valores de muestra por columna")
    ip.add_argument("--max-distinct", type=int, default=10000, help="Tope de valores distintos por columna")
    ip.add_argument("--header-scan", type=int, default=20, help="Filas a revisar para detectar cabecera")
    ip.add_argument("--sheet", action="append", help="Solo esta hoja (repetible)")
    ip.add_argument("--no-table", action="store_true", help="Ignorar tablas de Excel")
    ip.add_argument("--no-hash", action="store_true", help="No calcular SHA-256")

    vp = sub.add_parser("validate", help="Verifica que el schema esté al día")
    vp.add_argument("paths", nargs="+", help="Schema(s) JSON o directorios")

    pp = sub.add_parser("preview", help="Muestra filas usando el schema")
    pp.add_argument("schema", help="Ruta del schema JSON")
    pp.add_argument("--sheet", required=True, help="Hoja (o hoja::tabla)")
    pp.add_argument("--columns", help="'1,2,3', 'A:C' o nombres, separados por coma")
    pp.add_argument("--limit", type=int, default=10, help="Máx. filas a mostrar")

    return p


def main(argv=None) -> int:
    raw = list(sys.argv[1:] if argv is None else argv)
    subcommands = {"inspect", "validate", "preview", "-h", "--help"}
    if raw and raw[0] not in subcommands:
        raw = ["inspect"] + raw

    parser = build_parser()
    args = parser.parse_args(raw)

    if args.command == "validate":
        return run_validate(args)
    if args.command == "preview":
        return run_preview(args)
    return run_inspect(args)


if __name__ == "__main__":
    sys.exit(main())
