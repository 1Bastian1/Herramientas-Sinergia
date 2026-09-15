# TXD Odoo — transferencia de trabajo

Esta carpeta contiene los scripts para descargar y comparar productos TXD y ventas TXD desde Odoo mediante JSON-RPC. No contiene credenciales ni archivos descargados.

## Objetivo logrado

1. Descargar productos de `levantamiento_rep.producto` asociados a la unidad de negocio TXD y al período 2026.
2. Descargar todas las ventas 2026 de `venta_txd_raw` cuyos `codigo_item` aparecen entre los SKU de los productos descargados.
3. Para comparar, seleccionar por cada SKU la venta de su mes más actual. Si hay varias ventas en ese mismo mes, se utiliza la de ID mayor.
4. Entregar dos hojas: diferencias de marca y diferencias de origen.

## Modelos y llave de cruce

| Modelo | Campo usado |
|---|---|
| `levantamiento_rep.producto` | `sku_unidad_negocio` |
| `venta_txd_raw` | `codigo_item` |

La llave es: `producto.sku_unidad_negocio = venta.codigo_item`.

Filtros de productos ya verificados en esta instancia:

- Unidad de negocio TXD: ID `3`.
- Período 2026: ID `6`.
- Ventas de 2026: campo `mes` con el patrón `2026-%`.

## Preparación en el nuevo equipo

```powershell
pip install -r requirements_txd.txt
```

Configura las credenciales sólo en variables de entorno:

```powershell
$env:ODOO_URL = "https://tu-instancia.odoo.com"
$env:ODOO_DB = "tu_base"
$env:ODOO_USERNAME = "tu_usuario"
$env:ODOO_PASSWORD = "tu_contrasena"
```

Nunca escribir las credenciales dentro de los scripts ni subirlas al repositorio.

## Flujo de ejecución

```powershell
python descargar_productos_txd_2026.py
python descargar_ventas_txd_2026.py
python comparar_ultimo_mes_por_sku.py
python generar_comparacion_txd_revisable.py
```

Los nombres esperados de archivos son:

- Entrada/resultado de productos: `productos_txd_2026.xlsx`.
- Resultado de ventas: `ventas_txd_2026.xlsx`.
- Comparación base por último mes: `comparacion_ultimo_mes_por_sku_txd_2026.xlsx`.
- Comparación visual: `comparacion_txd_ultimo_mes_revisable.xlsx`.

## Scripts

- `odoo_jsonrpc_client.py`: cliente JSON-RPC; autentica por `/jsonrpc` y sólo se usa para lectura en este flujo.
- `descargar_productos_txd_2026.py`: descarga productos TXD 2026. Incluye ID, SKU, descripción, marca y origen.
- `descargar_ventas_txd_2026.py`: descarga todas las ventas 2026 de los SKU del Excel de productos. Procesa lotes de 10.000 SKU y crea nuevas hojas si supera el límite de Excel.
- `comparar_ultimo_mes_por_sku.py`: selecciona localmente una venta por SKU mediante `max(mes, id)` y genera diferencias básicas.
- `generar_comparacion_txd_revisable.py`: genera la versión visual para revisión. Resalta la columna del producto en azul y la de venta en amarillo. En marcas calcula similitud fuzzy y aplica escala de rojo a verde.

## Reglas de comparación

### Origen

Se comparan valores normalizados para evitar falsos positivos:

- `NAC` y `NACIONAL` se consideran equivalentes.
- `IMP`, `IMPORT` e `IMPORTADO` se consideran equivalentes.

El archivo final conserva tanto el valor original como el comparable.

### Marca

No se aplican alias ni equivalencias de marca. Se compara el texto original y se calcula una similitud fuzzy de 0% a 100% para facilitar la revisión:

- Rojo: poco parecido.
- Amarillo: similitud intermedia.
- Verde: muy parecido.

## Nota de rendimiento

Las descargas se hacen por JSON-RPC y son estrictamente de sólo lectura. La descarga de ventas puede tardar por el volumen. `tqdm` muestra porcentaje, velocidad y tiempo estimado cuando está instalado; los scripts dejan salida de progreso utilizable también en un log.
