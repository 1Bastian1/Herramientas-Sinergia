"""Exporta a Excel los productos de la unidad TXD asociados al período 2026.

Sólo realiza lecturas por JSON-RPC. Requiere ODOO_URL, ODOO_DB,
ODOO_USERNAME y ODOO_PASSWORD como variables de entorno.
"""

from __future__ import annotations

from pathlib import Path

import xlsxwriter
try:
    from tqdm import tqdm
except ImportError:
    class tqdm:  # Fallback compatible con entornos sin tqdm.
        def __init__(self, total=None, **_kwargs):
            self.total = total
            self.count = 0
        def __enter__(self): return self
        def __exit__(self, *_args): return False
        def update(self, amount):
            self.count += amount
            print(f"Productos descargados: {self.count}/{self.total}")

from odoo_jsonrpc_client import client_from_environment


MODEL = "levantamiento_rep.producto"
TXD_UNIT_ID = 3
YEAR_PERIOD_ID = 6
UNIT_LABEL = "TXD"
PAGE_SIZE = 2_000
MAX_DATA_ROWS = 999_999
OUTPUT = Path("productos_txd_2026.xlsx")

HEADERS = [
    "id_producto", "sku_producto", "sku_unidad_negocio", "descripcion",
    "marca", "origen", "unidad_de_negocio", "periodos", "es_txd",
]


def m2o_name(value):
    return value[1] if isinstance(value, list) and len(value) > 1 else ""


def main() -> None:
    odoo = client_from_environment()
    odoo.authenticate()
    domain = [("unidad_de_negocio", "in", [TXD_UNIT_ID]), ("periodo", "in", [YEAR_PERIOD_ID])]
    total = odoo.execute_kw(MODEL, "search_count", [domain])
    fields = ["name", "sku_unidad_negocio", "descripcion", "marca", "origen", "unidad_de_negocio", "periodo", "es_txd"]

    workbook = xlsxwriter.Workbook(OUTPUT, {"constant_memory": True})
    header_format = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E78", "align": "center", "valign": "vcenter", "text_wrap": True})
    text_format = workbook.add_format({"font_name": "Arial", "font_size": 10})
    part = 0
    row_number = MAX_DATA_ROWS
    offset = 0

    with tqdm(total=total, unit="producto", desc=f"Descargando productos {UNIT_LABEL} 2026") as progress:
        while True:
            records = odoo.execute_kw(MODEL, "search_read", [domain], {
                "fields": fields, "limit": PAGE_SIZE, "offset": offset, "order": "id",
            })
            if not records:
                break
            offset += len(records)
            progress.update(len(records))
            for record in records:
                if row_number >= MAX_DATA_ROWS:
                    part += 1
                    row_number = 0
                    sheet = workbook.add_worksheet(f"Productos {UNIT_LABEL} 2026 {part}")
                    sheet.hide_gridlines(2)
                    sheet.freeze_panes(1, 0)
                    sheet.write_row(0, 0, HEADERS, header_format)
                    sheet.set_column("A:A", 14, text_format)
                    sheet.set_column("B:C", 18, text_format)
                    sheet.set_column("D:D", 45, text_format)
                    sheet.set_column("E:I", 22, text_format)
                units = ", ".join(map(str, record["unidad_de_negocio"] or []))
                periods = ", ".join(map(str, record["periodo"] or []))
                sheet.write_row(row_number + 1, 0, [
                    record["id"], record["name"] or "", record["sku_unidad_negocio"] or "",
                    record["descripcion"] or "", m2o_name(record["marca"]), record["origen"] or "",
                    units, periods, record["es_txd"],
                ], text_format)
                row_number += 1
    workbook.close()
    print(f"Archivo creado: {OUTPUT.resolve()} | Productos: {total} | Hojas: {part}")


if __name__ == "__main__":
    main()
