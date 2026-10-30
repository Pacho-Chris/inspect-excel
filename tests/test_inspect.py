import json
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import inspect_excel as ie
from fixture_factory import build_ventas


class Args:
    sample = 5
    max_rows = 0
    max_distinct = 10000
    header_scan = 20
    footer_rows = 0
    header_row = None
    no_table = False
    sheet = None


def inspect(path: Path, **kw) -> dict:
    a = Args()
    for k, v in kw.items():
        setattr(a, k, v)
    return ie.inspect_file(path, a)["sheets"]


def col(meta: dict, name: str) -> dict:
    for c in meta["columns"]:
        if c["name"] == name:
            return c
    raise AssertionError(f"columna {name!r} no encontrada en {[c['name'] for c in meta['columns']]}")


@pytest.fixture()
def xlsx(tmp_path):
    p = tmp_path / "ventas.xlsx"
    build_ventas(p)
    return p


def test_detecta_tabla(xlsx):
    sheets = inspect(xlsx)
    ventas = sheets["Ventas"]
    assert ventas["kind"] == "table"
    assert ventas["table"] == "TablaVentas"
    assert ventas["range"] == "A2:F12"
    assert ventas["header_row"] == 2
    assert ventas["rows"] == 10


def test_tipos_ventas(xlsx):
    ventas = inspect(xlsx)["Ventas"]
    fecha = col(ventas, "Fecha")
    assert fecha["type"] == "date"
    assert fecha["letter"] == "A" and fecha["pos"] == 1

    monto = col(ventas, "Monto")
    assert monto["type"] == "number"
    assert monto["semantic"] == "currency"
    assert monto["decimal_places"] == 2

    iva = col(ventas, "IVA")
    assert iva["semantic"] == "percent"

    pagado = col(ventas, "Pagado")
    assert pagado["type"] == "boolean"

    nota = col(ventas, "Nota")
    assert nota["type"] == "text"
    assert nota["nulls"] > 0
    assert nota["unique"] == 3  # ok, pendiente, revisar


def test_clientes_cabecera_con_titulo(xlsx):
    sheets = inspect(xlsx)
    clientes = sheets["Clientes"]
    assert clientes["kind"] == "range"
    assert clientes["header_row"] == 2
    assert clientes["rows"] == 6


def test_clientes_tipos(xlsx):
    clientes = inspect(xlsx)["Clientes"]

    idc = col(clientes, "ID")
    assert idc["type"] == "text"
    assert "numeric_stored_as_text" in idc["warnings"]

    activo = col(clientes, "Activo")
    assert activo["type"] == "boolean"

    edad = col(clientes, "Edad")
    assert edad["type"] == "integer"
    assert edad["min"] == 25 and edad["max"] == 41

    categoria = col(clientes, "Categoria")
    assert categoria["type"] == "text"
    assert categoria["enum"] is not None
    assert {e["value"] for e in categoria["enum"]} == {"A", "B", "C"}

    saldo = col(clientes, "Saldo")
    assert saldo["type"] == "mixed"
    assert saldo["types"] == {"integer": 4, "text": 2}


def test_hoja_oculta(xlsx):
    sheets = inspect(xlsx)
    assert sheets["Backup"]["hidden"] is True
    assert sheets["Ventas"]["hidden"] is False


def test_sin_cabecera(xlsx):
    sheets = inspect(xlsx)
    s = sheets["SinCabecera"]
    assert s["header_row"] == 0
    assert s["rows"] == 3
    assert [c["name"] for c in s["columns"]] == ["column_A", "column_B", "column_C"]
    assert all(c["type"] == "integer" for c in s["columns"])


def test_hoja_vacia(xlsx):
    sheets = inspect(xlsx)
    v = sheets["Vacia"]
    assert v["rows"] == 0
    assert v["columns"] == []


def test_schema_serializable(xlsx):
    a = Args()
    entry = ie.inspect_file(xlsx, a)
    json.dumps(entry, ensure_ascii=False)  # no debe lanzar


def test_flujo_cli_y_validate(xlsx, tmp_path):
    schema = tmp_path / "excel_schema.json"
    rc = ie.main(["inspect", str(xlsx), "-o", str(schema)])
    assert rc == 0
    assert schema.exists()

    data = json.loads(schema.read_text(encoding="utf-8"))
    assert data["version"] == 2
    assert "ventas.xlsx" in data["workbooks"]
    assert data["workbooks"]["ventas.xlsx"]["sheet_count"] == 5

    # al día
    assert ie.main(["validate", str(schema)]) == 0

    # modificar el Excel -> stale
    build_ventas(xlsx)  # reescribe (cambia mtime/tamaño/hash)
    assert ie.main(["validate", str(schema)]) == 2


def test_preview(xlsx, tmp_path, capsys):
    schema = tmp_path / "excel_schema.json"
    assert ie.main(["inspect", str(xlsx), "-o", str(schema)]) == 0

    rc = ie.main(["preview", str(schema), "--sheet", "Ventas", "--columns", "1,2,3", "--limit", "4"])
    assert rc == 0
    out = capsys.readouterr().out
    assert "Fecha" in out and "Cliente" in out and "Monto" in out
    # filas de datos limitadas a 4
    data_lines = [l for l in out.splitlines() if l.strip().startswith("2024")]
    assert len(data_lines) == 4


def test_preview_columnas_por_letra_y_nombre(xlsx, tmp_path, capsys):
    schema = tmp_path / "excel_schema.json"
    assert ie.main(["inspect", str(xlsx), "-o", str(schema)]) == 0
    assert ie.main(["preview", str(schema), "--sheet", "Clientes", "--columns", "B,C", "--limit", "3"]) == 0
    out = capsys.readouterr().out
    assert "Nombre" in out and "Activo" in out
