"""Crea un libro revisable: origen normalizado y marca con similitud fuzzy."""
from __future__ import annotations

import re
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path

import xlsxwriter
from openpyxl import load_workbook

PRODUCTS = Path("productos_txd_2026.xlsx")
SALES = Path("ventas_txd_2026.xlsx")
OUTPUT = Path("comparacion_txd_ultimo_mes_revisable.xlsx")


def rows(path):
    workbook = load_workbook(path, read_only=True, data_only=True)
    for sheet in workbook.worksheets:
        headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
        for values in sheet.iter_rows(min_row=2, values_only=True):
            yield dict(zip(headers, values))


def origin_key(value) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = re.sub(r"[^A-Z0-9]", "", text.upper())
    return {
        "NAC": "NACIONAL", "NACIONAL": "NACIONAL",
        "IMP": "IMPORTADO", "IMPORT": "IMPORTADO", "IMPORTADO": "IMPORTADO",
    }.get(text, text)


def similarity(left, right) -> float:
    # Sin alias ni reemplazos de marca: sólo una métrica de similitud textual.
    return round(SequenceMatcher(None, str(left or "").casefold(), str(right or "").casefold()).ratio() * 100, 1)


def main() -> None:
    products_by_sku = {}
    for product in rows(PRODUCTS):
        sku = str(product.get("sku_unidad_negocio") or "").strip()
        if sku:
            products_by_sku.setdefault(sku, []).append(product)

    latest_by_sku = {}
    for sale in rows(SALES):
        sku = str(sale.get("sku_unidad_negocio") or "").strip()
        candidate = (str(sale.get("periodo_venta") or ""), int(sale.get("id_venta_txd") or 0))
        if sku and (sku not in latest_by_sku or candidate > latest_by_sku[sku][0]):
            latest_by_sku[sku] = (candidate, sale)

    workbook = xlsxwriter.Workbook(OUTPUT, {"constant_memory": True})
    header = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E78", "align": "center", "valign": "vcenter", "text_wrap": True})
    product_compare = workbook.add_format({"font_name": "Arial", "font_size": 10, "bg_color": "#DDEBF7"})
    sale_compare = workbook.add_format({"font_name": "Arial", "font_size": 10, "bg_color": "#FFF2CC"})
    normal = workbook.add_format({"font_name": "Arial", "font_size": 10})
    pct = workbook.add_format({"font_name": "Arial", "font_size": 10, "num_format": "0.0\"%\""})

    origin_headers = ["sku", "id producto", "nombre producto", "origen producto", "origen producto comparable", "id venta", "código ítem venta", "período venta", "nombre venta", "origen venta", "origen venta comparable"]
    brand_headers = ["sku", "id producto", "nombre producto", "marca producto", "id venta", "código ítem venta", "período venta", "nombre venta", "marca venta", "similitud de marca", "origen producto", "origen venta"]
    origin_sheet = workbook.add_worksheet("Origen distinto")
    brand_sheet = workbook.add_worksheet("Marca distinta")

    for sheet, headers in ((origin_sheet, origin_headers), (brand_sheet, brand_headers)):
        sheet.hide_gridlines(2); sheet.freeze_panes(1, 0)
        sheet.write_row(0, 0, headers, header)
        sheet.set_column(0, 1, 18, normal); sheet.set_column(2, 2, 34, normal)
        sheet.set_column(3, len(headers) - 1, 22, normal)

    # Columnas que se comparan: producto en azul y venta en amarillo.
    origin_sheet.set_column(3, 3, 22, product_compare)
    origin_sheet.set_column(9, 9, 22, sale_compare)
    brand_sheet.set_column(3, 3, 22, product_compare)
    brand_sheet.set_column(8, 8, 22, sale_compare)
    brand_sheet.set_column(9, 9, 18, pct)
    brand_sheet.conditional_format("J2:J1048576", {"type": "3_color_scale", "min_color": "#F8696B", "mid_color": "#FFEB84", "max_color": "#63BE7B"})

    origin_count = brand_count = 0
    for sku, (_, sale) in latest_by_sku.items():
        for product in products_by_sku.get(sku, []):
            product_origin, sale_origin = product.get("origen") or "", sale.get("origen_venta") or ""
            product_brand, sale_brand = product.get("marca") or "", sale.get("codigo_tipo_marca") or ""
            if origin_key(product_origin) != origin_key(sale_origin):
                values = [sku, product.get("id_producto"), product.get("sku_producto") or "", product_origin, origin_key(product_origin), sale.get("id_venta_txd"), sale.get("sku_unidad_negocio") or "", sale.get("periodo_venta") or "", sale.get("nombre_producto_venta") or "", sale_origin, origin_key(sale_origin)]
                for col, value in enumerate(values):
                    origin_sheet.write(origin_count + 1, col, value, product_compare if col == 3 else sale_compare if col == 9 else normal)
                origin_count += 1
            if str(product_brand) != str(sale_brand):
                values = [sku, product.get("id_producto"), product.get("sku_producto") or "", product_brand, sale.get("id_venta_txd"), sale.get("sku_unidad_negocio") or "", sale.get("periodo_venta") or "", sale.get("nombre_producto_venta") or "", sale_brand, similarity(product_brand, sale_brand) / 100, product_origin, sale_origin]
                for col, value in enumerate(values):
                    brand_sheet.write(brand_count + 1, col, value, product_compare if col == 3 else sale_compare if col == 8 else pct if col == 9 else normal)
                brand_count += 1
    workbook.close()
    print(f"Archivo: {OUTPUT.resolve()} | Origen: {origin_count} | Marca: {brand_count}")


if __name__ == "__main__":
    main()
