"""Construye un libro Excel de ejemplo con casos representativos.

Contenido:
- "Ventas": tabla de Excel real (TablaVentas), fila de título, cabecera,
  fechas, moneda, porcentaje, booleanos y nulos.
- "Clientes": rango con fila de título y cabecera detectada, IDs con ceros a
  la izquierda, booleanos, enteros, enum y columna mixta.
- "Backup": hoja oculta.
- "SinCabecera": solo números, sin cabecera.
- "Vacia": hoja vacía.
"""

from __future__ import annotations

import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


def build_ventas(path: Path) -> Path:
    wb = Workbook()

    # --- Ventas (tabla real) ---
    ws = wb.active
    ws.title = "Ventas"
    ws.append(["Reporte de Ventas 2024"])  # fila 1: título
    ws.append(["Fecha", "Cliente", "Monto", "IVA", "Pagado", "Nota"])  # fila 2
    fechas = [datetime.datetime(2024, 1, i + 1) for i in range(10)]
    clientes = ["Cliente A", "Cliente B", "Cliente C", "Cliente D", "Cliente E"] * 2
    montos = [100.0, 250.5, 1234.56, 89.99, 500.0, 75.25, 999.99, 44.44, 320.0, 12.34]
    ivas = [0.12] * 10
    pagados = [True, False, True, True, False, True, False, True, True, False]
    notas = ["ok", None, "pendiente", None, "ok", "revisar", None, "ok", None, "ok"]
    for i in range(10):
        ws.append([fechas[i], clientes[i], montos[i], ivas[i], pagados[i], notas[i]])
    for r in range(3, 13):
        ws.cell(row=r, column=1).number_format = "dd/mm/yyyy"
        ws.cell(row=r, column=3).number_format = '$#,##0.00'
        ws.cell(row=r, column=4).number_format = "0.00%"
    tab = Table(displayName="TablaVentas", ref="A2:F12")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(tab)

    # --- Clientes (rango con título y cabecera detectada) ---
    ws2 = wb.create_sheet("Clientes")
    ws2.append(["Listado de Clientes"])  # fila 1: título
    ws2.append(["ID", "Nombre", "Activo", "Edad", "Categoria", "Saldo"])  # fila 2
    ids = ["0001", "0002", "0003", "0004", "0005", "0006"]
    nombres = ["Ana", "Luis", "Mia", "Pablo", "Sara", "Leo"]
    activos = [True, True, False, True, False, True]
    edades = [34, 28, 41, 25, 37, 30]
    cats = ["A", "B", "A", "C", "B", "A"]
    saldos = [10, 20, "pendiente", 30, "n/a", 40]
    for i in range(6):
        ws2.append([ids[i], nombres[i], activos[i], edades[i], cats[i], saldos[i]])

    # --- Backup (oculta) ---
    ws3 = wb.create_sheet("Backup")
    ws3.append(["X", "Y"])
    ws3.append([1, 2])
    ws3.sheet_state = "hidden"

    # --- SinCabecera (solo números) ---
    ws4 = wb.create_sheet("SinCabecera")
    ws4.append([1, 2, 3])
    ws4.append([4, 5, 6])
    ws4.append([7, 8, 9])

    # --- Vacia ---
    wb.create_sheet("Vacia")

    wb.save(path)
    return path
