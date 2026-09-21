"""Descarga registros de venta_mdh_raw del año 2026, incluidos archivados."""
from __future__ import annotations

from pathlib import Path
import sys
import xlsxwriter

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from odoo_jsonrpc_client import client_from_environment

MODEL = "venta_mdh_raw"
YEAR = "2026"
PAGE_SIZE = 5000
OUTPUT = Path(__file__).resolve().parent / "ventas_mdh_raw_2026.xlsx"
FIELDS = ["id", "ean", "articulo", "ano", "mes", "marca", "nombre_marca", "categoria_valorizacion"]


def main() -> None:
    odoo = client_from_environment()
    odoo.authenticate()
    domain = [("ano", "=", YEAR)]
    context = {"active_test": False}
    total = odoo.execute_kw(MODEL, "search_count", [domain], {"context": context})
    workbook = xlsxwriter.Workbook(OUTPUT, {"constant_memory": True})
    header = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E78", "align": "center"})
    text = workbook.add_format({"font_name": "Arial", "font_size": 10})
    sheet = workbook.add_worksheet("Ventas MDH Raw 2026")
    sheet.hide_gridlines(2); sheet.freeze_panes(1, 0)
    sheet.write_row(0, 0, FIELDS, header)
    sheet.set_column("A:A", 14, text); sheet.set_column("B:C", 20, text)
    sheet.set_column("D:E", 12, text); sheet.set_column("F:H", 24, text)
    offset = 0; row = 1
    while True:
        records = odoo.execute_kw(MODEL, "search_read", [domain], {
            "fields": FIELDS, "limit": PAGE_SIZE, "offset": offset, "order": "id", "context": context,
        })
        if not records:
            break
        for record in records:
            sheet.write_row(row, 0, [record.get(field) or "" for field in FIELDS], text)
            row += 1
        offset += len(records)
        print(f"Registros descargados: {offset}/{total}", flush=True)
    workbook.close()
    print(f"Archivo creado: {OUTPUT.resolve()} | Registros: {total}", flush=True)


if __name__ == "__main__":
    main()
