# Generador de Archivo Contable (SIRE LyJ)

Aplicación Streamlit para generar el Excel mensual de **Compras** o **Ventas**
con la plantilla habitual (hoja `Hoja1`, 80 columnas, fórmulas y fila de totales).

## Archivos

| Archivo | Función |
|---|---|
| `app.py` | Interfaz Streamlit (3 pasos: tipo y periodo, comprobantes, vista previa y descarga). |
| `excel_builder.py` | Construye el `.xlsx` respetando formatos y fórmulas de la plantilla. |
| `catalogo.py` | Valores por defecto: RUC propio, proveedores/clientes con su RUC y series, detalles habituales. |
| `.streamlit/config.toml` | Tema claro con acentos verdes. |
| `requirements.txt` | Dependencias. |

## Ejecutar en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicar en Streamlit Community Cloud

1. Sube esta carpeta a un repositorio de GitHub (puede ser la raíz del repo o una subcarpeta).
2. En https://share.streamlit.io crea una app nueva apuntando a `app.py`
   (si está en subcarpeta, indica `generador_sire/app.py`).
3. Streamlit instala `requirements.txt` automáticamente.

## Mantener los valores por defecto

Para agregar un proveedor o cliente nuevo de forma permanente, edita `catalogo.py`
(listas `PROVEEDORES` y `CLIENTES`). Desde la app siempre se puede escribir un tercero
nuevo sin tocar el código.

## Fórmulas que se escriben en el Excel

- `CONCATENADO` = `=CONCATENATE(G,"-",H,"-",J)` en cada fila.
- `IGV / IPM DG` = `=O*0.18` cuando el toggle "IGV automático" está activo (por defecto en Ventas).
- `Total CP` = `=SUM(O:X)` cuando "Total CP automático" está activo (por defecto).
- Fila TOTAL = `=SUM(col2:col{última+1})` en las columnas O…Y, con borde superior fino e inferior doble.
- Sección de mes anterior: se escribe debajo del total con su propia fila de totales.
