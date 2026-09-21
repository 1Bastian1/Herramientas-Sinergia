# Herramientas Sinergia

Repositorio de herramientas operativas usadas por Sinergia para apoyar procesos de revisión, comparación y actualización de información en Odoo.

## Herramientas disponibles

### [CENCO] revision de campos producto vs ventas

Herramienta para revisar consistencia entre los datos maestros de productos CENCO/TXD y la información registrada en ventas.

Su objetivo principal es detectar diferencias entre lo que existe en el modelo de productos y lo que aparece en ventas para los mismos SKU. Actualmente el flujo trabaja con productos TXD del período 2026 y ventas TXD 2026 descargadas desde Odoo por JSON-RPC.

Qué hace:

- Descarga productos desde `levantamiento_rep.producto` usando como llave el SKU de unidad de negocio.
- Descarga ventas desde `venta_txd_raw` para los SKU encontrados en productos.
- Cruza producto contra venta usando `producto.sku_unidad_negocio = venta.codigo_item`.
- Selecciona, por cada SKU, la venta del mes más reciente y, si hay más de una en ese mes, usa la de mayor ID.
- Genera archivos Excel para revisar diferencias de marca y origen.
- Resalta visualmente diferencias para facilitar revisión manual.

Archivos principales:

- `odoo_jsonrpc_client.py`: cliente de conexión JSON-RPC a Odoo.
- `descargar_productos_txd_2026.py`: descarga productos TXD 2026.
- `descargar_ventas_txd_2026.py`: descarga ventas TXD 2026 relacionadas a esos SKU.
- `comparar_ultimo_mes_por_sku.py`: cruza productos y ventas tomando la venta más reciente por SKU.
- `generar_comparacion_txd_revisable.py`: genera un Excel revisable con formato visual.

Notas importantes:

- No guardar credenciales dentro de los scripts.
- Usar variables de entorno para `ODOO_URL`, `ODOO_DB`, `ODOO_USERNAME` y `ODOO_PASSWORD`.
- La herramienta es de lectura; no modifica datos en Odoo.

### [CENCO] actualizacion de maestros en modelo productos

Carpeta destinada a la herramienta de actualización de maestros en el modelo de productos.

La idea de esta herramienta es centralizar scripts, plantillas y documentación para preparar o ejecutar actualizaciones masivas sobre datos maestros de productos CENCO en Odoo.

Uso esperado:

- Preparar archivos de entrada para actualización de productos.
- Validar columnas obligatorias antes de cargar cambios.
- Documentar reglas de negocio para modificación de maestros.
- Mantener separados los procesos de actualización respecto a los procesos de sólo revisión.

Estado actual:

- Carpeta creada como base de trabajo.
- Pendiente agregar scripts o plantillas específicas cuando se defina el flujo exacto de actualización.

## Convenciones

- Las herramientas asociadas a CENCO deben comenzar con el prefijo `[CENCO]`.
- Cada herramienta debe tener su propio `README.md` cuando tenga scripts o pasos de ejecución propios.
- Los archivos con credenciales, descargas pesadas o resultados temporales no deben subirse al repositorio.
