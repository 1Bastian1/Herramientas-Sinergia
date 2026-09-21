"""Compara productos TXD con la venta más reciente de cada SKU en 2026.

Si un SKU tiene varias ventas en su mes máximo, se conserva la de mayor ID.
"""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import xlsxwriter
from openpyxl import load_workbook

PRODUCTS = Path("productos_txd_2026.xlsx")
SALES = Path("ventas_txd_2026.xlsx")
OUTPUT = Path("comparacion_ultimo_mes_por_sku_txd_2026.xlsx")


def key(value) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    return re.sub(r"[^A-Z0-9]", "", "".join(c for c in text if not unicodedata.combining(c)).upper())


def origin_key(value) -> str:
    return {
        "NAC": "NACIONAL", "NACIONAL": "NACIONAL",
        "IMP": "IMPORTADO", "IMPORT": "IMPORTADO", "IMPORTADO": "IMPORTADO",
    }.get(key(value), key(value))


def rows(path):
    book = load_workbook(path, read_only=True, data_only=True)
    for sheet in book.worksheets:
        headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
        for values in sheet.iter_rows(min_row=2, values_only=True):
            yield dict(zip(headers, values))


def main() -> None:
    products = {}
    for product in rows(PRODUCTS):
        sku = str(product.get("sku_unidad_negocio") or "").strip()
        if sku:
            products.setdefault(sku, []).append(product)

    # Selección local de una venta por SKU: mes máximo y luego ID máximo.
    latest_sales = {}
    for sale in rows(SALES):
        sku = str(sale.get("sku_unidad_negocio") or "").strip()
        if not sku:
            continue
        candidate = (str(sale.get("periodo_venta") or ""), int(sale.get("id_venta_txd") or 0))
        current = latest_sales.get(sku)
        if current is None or candidate > current[0]:
            latest_sales[sku] = (candidate, sale)

    workbook = xlsxwriter.Workbook(OUTPUT, {"constant_memory": True})
    header_format = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E78", "align": "center", "text_wrap": True})
    text_format = workbook.add_format({"font_name": "Arial", "font_size": 10})
    headers = ["sku_unidad_negocio", "id_producto", "nombre_producto", "valor_producto", "valor_producto_normalizado", "id_venta", "periodo_venta", "nombre_producto_venta", "valor_venta", "valor_venta_normalizado"]

    sheets = {}
    counts = {"Marca vs marca venta": 0, "Marca vs codigo_tipo_marca": 0, "Origen distinto": 0}
    for title in counts:
        sheet = workbook.add_worksheet(title)
        sheet.hide_gridlines(2); sheet.freeze_panes(1, 0)
        sheet.write_row(0, 0, headers, header_format)
        sheet.set_column("A:B", 18, text_format); sheet.set_column("C:C", 34, text_format)
        sheet.set_column("D:E", 23, text_format); sheet.set_column("F:G", 18, text_format)
        sheet.set_column("H:H", 34, text_format); sheet.set_column("I:J", 23, text_format)
        sheets[title] = sheet

    for sku, sale_data in latest_sales.items():
        sale = sale_data[1]
        for product in products.get(sku, []):
            checks = [
                ("Marca vs marca venta", product.get("marca") or "", sale.get("marca_venta") or "", key),
                ("Marca vs codigo_tipo_marca", product.get("marca") or "", sale.get("codigo_tipo_marca") or "", key),
                ("Origen distinto", product.get("origen") or "", sale.get("origen_venta") or "", origin_key),
            ]
            for title, product_value, sale_value, normalize in checks:
                if normalize(product_value) != normalize(sale_value):
                    row = counts[title] + 1
                    sheets[title].write_row(row, 0, [
                        sku, product.get("id_producto"), product.get("sku_producto") or "", product_value, normalize(product_value),
                        sale.get("id_venta_txd"), sale.get("periodo_venta") or "", sale.get("nombre_producto_venta") or "", sale_value, normalize(sale_value),
                    ], text_format)
                    counts[title] += 1
    workbook.close()
    print(f"Archivo: {OUTPUT.resolve()} | Marca vs marca: {counts['Marca vs marca venta']} | Marca vs código: {counts['Marca vs codigo_tipo_marca']} | Origen: {counts['Origen distinto']} | SKU con venta: {len(latest_sales)}")


if __name__ == "__main__":
    main()
