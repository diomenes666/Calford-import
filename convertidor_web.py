import os
import sys
import math
import re
from datetime import datetime
import pandas as pd
import streamlit as st
import io

# Configuración de página con diseño ancho y título moderno
st.set_page_config(
    page_title="E-Commerce Matrix Connector",
    page_icon="🚀",
    layout="centered",
    initial_sidebar_state="expanded"
)

# Inyección de estilos CSS personalizados para lograr una estética premium
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        
        /* Modificar fuente general */
        html, body, [class*="css"] {
            font-family: 'Inter', sans-serif;
        }
        
        /* Estilos de las cajas del Uploader */
        .stFileUploader {
            border: 2px dashed #cbd5e1;
            border-radius: 12px;
            padding: 10px;
            background-color: #f8fafc;
            transition: all 0.3s ease;
        }
        .stFileUploader:hover {
            border-color: #0284c7;
            background-color: #f0f9ff;
        }
        
        /* Título del panel lateral */
        .sidebar .sidebar-content {
            background-color: #0f172a;
        }
        
        /* Botón de descarga principal */
        div.stDownloadButton > button {
            background: linear-gradient(135deg, #0284c7, #0369a1) !important;
            color: white !important;
            border: none !important;
            padding: 12px 24px !important;
            border-radius: 8px !important;
            font-weight: 600 !important;
            font-size: 16px !important;
            box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1), 0 2px 4px -2px rgb(0 0 0 / 0.1) !important;
            transition: all 0.2s ease !important;
            width: 100% !important;
        }
        div.stDownloadButton > button:hover {
            background: linear-gradient(135deg, #0369a1, #075985) !important;
            transform: translateY(-1px) !important;
            box-shadow: 0 10px 15px -3px rgb(0 0 0 / 0.1), 0 4px 6px -4px rgb(0 0 0 / 0.1) !important;
        }
    </style>
""", unsafe_allow_html=True)

# =====================================================================
# 1. UTILIDADES COMPARTIDAS (LIMPIEZA Y HELPERS)
# =====================================================================

_EMOJI_PATTERN = re.compile(
    "["
    "\U0001F000-\U0001FFFF"
    "\U00002000-\U000027FF"
    "\U00002B00-\U00002BFF"
    "\U0000FE00-\U0000FE0F"
    "\U0000FFF0-\U0000FFFF"
    "\u200b-\u200f"
    "\u2028-\u202f"
    "]+",
    flags=re.UNICODE
)

def eliminar_emojis(texto):
    return _EMOJI_PATTERN.sub('', str(texto))

def limpiar_nombre(texto):
    if pd.isna(texto):
        return ""
    texto_str = eliminar_emojis(str(texto))
    texto_limpio = re.sub(r'-[^-]+-', '', texto_str)
    texto_limpio = texto_limpio.replace("'", "").replace("-", "")
    return " ".join(texto_limpio.split())

def limpiar_descripcion(html_texto):
    if pd.isna(html_texto):
        return ""
    texto = str(html_texto)
    texto = texto.replace('\\n', '<br>')
    texto = texto.replace('_x000D_', '\n')
    texto = eliminar_emojis(texto)
    texto = re.sub(r'[ \t]+', ' ', texto).strip()
    return texto

def calcular_stock_shopstar(val):
    try:
        stock = float(val)
        if pd.isna(stock) or stock <= 0:
            return 0
        if stock in [1, 2]:
            return 1
        return math.ceil(stock * 0.25)
    except (ValueError, TypeError):
        return 0

def calcular_precio_especial_shopstar(val):
    try:
        precio = float(val)
        if pd.isna(precio) or precio <= 0:
            return 0
        precio_ajustado = precio * 1.26
        return math.ceil((precio_ajustado - 9) / 10) * 10 + 9
    except (ValueError, TypeError):
        return 0

def limpiar_imagenes(val):
    if pd.isna(val):
        return ""
    enlaces = [url.strip() for url in str(val).split(',')]
    enlaces = [url for url in enlaces if url]
    return ",".join(enlaces)

# =====================================================================
# 2. MOTOR SHOPSTAR (INTACTO — 97 columnas)
# =====================================================================

COLUMNAS_PLANTILLA = [
    'Link Imagenes', 'Categoria', 'Nombre Producto', 'Nombre SKU', 'SKU',
    'Descripcion', 'Marca', 'Keywords', 'Peso', 'Alto', 'Ancho', 'Largo',
    'Stock', 'Precio Base', 'Precio Especial', 'Precio Especial Inicio',
    'Precio Especial Hasta', 'Producto_incluyeluces', 'Producto_edad',
    'Producto_tipodefijacion', 'Producto_edadsugeridadeuso',
    'Producto_profundidad(cm)', 'Producto_tamaño/largo(cm)', 'Producto_incluye',
    'Producto__excerptdescription', 'Producto_pesodelproducto(kg)',
    'Producto_incluyemovimiento', 'Producto_ancho', 'Producto_ancho(cm)',
    'Producto_edadminimarecomendada', 'Producto_material', 'Producto_observaciones',
    'Producto_sonidos', 'Producto_peso', 'Producto_FuenteDeEnergía',
    'Producto_alto(cm)', 'Producto_recomendacionesdeuso',
    'Producto_cantidaddejugadores', 'Producto_color', 'Producto_alto',
    'Producto_capacidad', 'Producto_impermeable', 'Producto_composicion',
    'Producto_certificaciones', 'Producto_comousarlo', 'Producto_nombrecomercial',
    'Producto_largo(cm)', 'Producto_vibraciones', 'Producto_descripciondelproducto',
    'Producto_advertenciadeuso', 'Producto_genero', 'Producto_cantidaddeposiciones',
    'Producto_accesorios', 'Producto_modelo', 'Producto_garantia',
    'Producto_nomenclatura', 'Producto_Medidas', 'Producto_IncluyeBaterías',
    'Producto_marca', 'Producto_unidadesporpaquete', 'Producto_descripcion',
    'Producto_lavable', 'Producto_resistenciamaxima', 'Producto_pesodelproducto',
    'Producto_largo', 'Producto_pesodelproducto(g)', 'Producto_JugueteDidáctico',
    'Producto_medidasempaque', 'Producto_edadsugerida', 'Producto_DetalleDeSurtido',
    'Producto_Luces', 'Producto_cantidaddepiezas', 'Producto_sku', 'Producto_piezas',
    'Producto_diseño', 'Producto_Sonido', 'Producto_rapidez',
    'Producto_pesomaximoporusuario', 'Producto_color/diseño',
    'Producto_advertenciasdeuso', 'Producto_tipodeproducto', 'Producto_peso(kg)',
    'Producto_caracteristicas', 'Producto_textura', 'Producto_antialergico',
    'Producto_masinformacion', 'Link', 'Meta Title', 'Meta Descripcion', 'Same Day'
]

def procesar_logica_shopstar(df_wp, df_marcas_maestro):
    logs = []

    columnas_requeridas_wp = [
        'SKU', 'Nombre', 'Descripción', 'Marcas',
        'Inventario', 'Precio normal', 'Imágenes', 'URL externa',
        'Peso (kg)', 'Longitud (cm)', 'Anchura (cm)', 'Altura (cm)'
    ]
    columnas_faltantes_wp = [col for col in columnas_requeridas_wp if col not in df_wp.columns]
    if columnas_faltantes_wp:
        logs.append((f"❌ ERROR: El CSV de WordPress no tiene la estructura correcta.\nFaltan:\n" +
                     "\n".join([f"  ⚠️ '{c}'" for c in columnas_faltantes_wp]), "error"))
        return None, logs, True

    if 'MARCA WP' not in df_marcas_maestro.columns or 'MARCA SS' not in df_marcas_maestro.columns:
        logs.append(("❌ ERROR: La tabla de marcas debe tener los encabezados 'MARCA WP' y 'MARCA SS'.", "error"))
        return None, logs, True

    dict_marcas = dict(zip(
        df_marcas_maestro['MARCA WP'].astype(str).str.strip().str.lower(),
        df_marcas_maestro['MARCA SS'].astype(str).str.strip()
    ))

    marcas_en_wp = df_wp['Marcas'].dropna().unique()
    marcas_faltantes = [str(m).strip() for m in marcas_en_wp if str(m).strip().lower() not in dict_marcas]
    if marcas_faltantes:
        logs.append(("❌ PROCESO DETENIDO: Marcas nuevas sin registrar en equivalencias:\n" +
                     "\n".join([f"  • {m}" for m in marcas_faltantes]), "error"))
        return None, logs, True

    logs.append(("✨ Aplicando transformaciones y limpiando textos...", "info"))
    n = len(df_wp)

    precio_para_especial = df_wp['Precio rebajado'].fillna(0).astype(float)
    precio_normal = df_wp['Precio normal'].fillna(0).astype(float)
    precio_base_especial = precio_para_especial.where(precio_para_especial > 0, precio_normal)

    df_out = pd.DataFrame('', index=range(n), columns=COLUMNAS_PLANTILLA)

    df_out['Link Imagenes']          = df_wp['Imágenes'].apply(limpiar_imagenes)
    df_out['Categoria']              = '1038-Infantil/Juguetes/Coleccionables'
    df_out['Nombre Producto']        = df_wp['Nombre'].apply(limpiar_nombre)
    df_out['Nombre SKU']             = df_out['Nombre Producto']
    df_out['SKU']                    = df_wp['SKU'].fillna('')
    df_out['Descripcion']            = df_wp['Descripción'].apply(limpiar_descripcion)
    df_out['Marca']                  = df_wp['Marcas'].astype(str).str.strip().str.lower().map(dict_marcas)
    df_out['Keywords']               = ''
    df_out['Peso']                   = df_wp['Peso (kg)'].fillna(0)
    df_out['Alto']                   = df_wp['Altura (cm)'].fillna(0)
    df_out['Ancho']                  = df_wp['Anchura (cm)'].fillna(0)
    df_out['Largo']                  = df_wp['Longitud (cm)'].fillna(0)
    df_out['Stock']                  = df_wp['Inventario'].apply(calcular_stock_shopstar)
    df_out['Precio Especial']        = precio_base_especial.apply(calcular_precio_especial_shopstar)
    df_out['Precio Base']            = (df_out['Precio Especial'] * 1.5).round(0).astype(int)
    df_out['Precio Especial Inicio'] = datetime.now().strftime('%d/%m/%Y')
    df_out['Precio Especial Hasta']  = '05/19/2050 23:24:07'

    df_out['Link'] = (
        df_out['Nombre Producto'].str.lower().str.replace(' ', '-', regex=False)
        + '-'
        + df_out['SKU'].str.lower()
    )

    logs.append((f"✅ Conversión finalizada con éxito. Se procesaron {len(df_out)} productos (97 columnas).", "success"))
    return df_out, logs, False

# =====================================================================
# 3. MOTOR FALABELLA (NUEVO)
# =====================================================================

# Las 4 filas de encabezados técnicos fijos que exige Falabella
FALABELLA_FILA1 = [
    'V1 Ejemplo de conjunto de atributos "Juguetes y juegos"',
    "Principales", "Principales", "Principales", "Principales", "Principales",
    "Variaciones", "Variaciones", "Variaciones",
    "Precio", "Precio", "Precio", "Precio", "Precio",
    "Especificaciones", "Especificaciones", "Especificaciones", "Especificaciones",
    "Especificaciones", "Especificaciones", "Especificaciones", "Especificaciones",
    "Garantía y Envío", "Garantía y Envío", "Garantía y Envío", "Garantía y Envío",
    "Garantía y Envío", "Garantía y Envío", "Garantía y Envío", "Garantía y Envío",
    "Garantía y Envío", "Garantía y Envío", "Garantía y Envío",
    "Imágenes", "Imágenes", "Imágenes", "Imágenes",
    "Imágenes", "Imágenes", "Imágenes", "Imágenes"
]

FALABELLA_FILA2 = [
    "Indica el nombre...", "Select the brand...", "Ingresa el modelo...",
    "Ingresa la descripción...", "Selecciona la categoría...", "País donde fue fabricado...",
    "Código único...", "Código de barras...", "Atributo de variación...",
    "Cantidad disponible...", "Precio normal...", "Precio rebajado...",
    "Fecha de inicio...", "Fecha de término...",
    "Selecciona el tipo...", "Selecciona el grupo...", "Indica si contiene...",
    "Características de salud...", "Color del producto...", "Material principal...",
    "Personaje...", "Peso del producto...",
    "Condición...", "Nombre en Chino...", "Nombre en Inglés...", "Detalles...",
    "Garantía...", "Garantía...", "Contenido...",
    "Ancho del paquete...", "Largo del paquete...", "Alto del paquete...", "Peso del paquete...",
    "URL imagen...", "URL imagen 2...", "URL imagen 3...", "URL imagen 4...",
    "URL imagen 5...", "URL imagen 6...", "URL imagen 7...", "URL imagen 8..."
]

FALABELLA_FILA3 = [
    "Nombre #39", "Marca #26", "Modelo #32", "Descripción #53",
    "Categoría primaria #1", "País de producción #59",
    "SKU del vendedor #29", "Código de barras #56", "Variación #1312",
    "QuantityFalabella #25", "PriceFalabella #52", "SalePriceFalabella #18",
    "SaleStartDateFalabella #45", "SaleEndDateFalabella #31",
    "TipoDeFigura #334136", "GrupoDeEdad #1301", "PiezasPequenas #1302",
    "CaracteristicasDeSalud #1303", "Color #1317", "Material #1322",
    "Personajes #1313", "WeightOfTheProduct #1306",
    "Condición del Producto #22", "NameCn #133815", "NameEn #133816",
    "Detalles de la condición del Producto #49",
    "Garantía del producto #35", "Garantía del vendedor #9",
    "Contenido del paquete #19",
    "Ancho del paquete #60", "Largo del paquete #33",
    "Alto del paquete #47", "Peso del paquete #8",
    "Imagen principal #IM1", "Imagen2 #IM2", "Imagen3 #IM3", "Imagen4 #IM4",
    "Imagen5 #IM5", "Imagen6 #IM6", "Imagen7 #IM7", "Imagen8 #IM8"
]

FALABELLA_FILA4 = [
    "", "", "", "Esto es un párrafo", "", "", "", "1234567890", "",
    "10", "999.99", "99.99", "2023-10-01", "2030-12-31",
    "", "", "", "", "", "", "", "",
    "Nuevo", "", "", "", "", "", "",
    "10", "10", "10", "1",
    "", "", "", "", "", "", "", ""
]

def limpiar_sku_falabella(val):
    """Elimina cualquier carácter no alfanumérico del SKU."""
    if pd.isna(val):
        return ""
    return re.sub(r'[^A-Za-z0-9]', '', str(val))

def limpiar_nombre_falabella(texto):
    """Elimina -PREVENTA-, -OFERTA-, -LIQUIDACION- y comillas."""
    if pd.isna(texto):
        return ""
    texto_str = str(texto)
    for tag in ['-PREVENTA-', '-OFERTA-', '-LIQUIDACION-']:
        texto_str = texto_str.replace(tag, '')
    texto_str = texto_str.replace('"', '').replace("'", '')
    return " ".join(texto_str.split())

def calcular_sale_price_falabella(val):
    """
    Precio psicológico:
    1. precio * 1.26
    2. Redondear a la decena más cercana
    3. Restar 1
    Siempre entero.
    """
    try:
        precio = float(val)
        if pd.isna(precio) or precio <= 0:
            return 0
        precio_ajustado = precio * 1.26
        redondeado = round(precio_ajustado / 10) * 10
        return int(redondeado - 1)
    except (ValueError, TypeError):
        return 0

def calcular_stock_falabella(val):
    """Inventario × 0.5, redondeado hacia arriba (math.ceil)."""
    try:
        stock = float(val)
        if pd.isna(stock) or stock <= 0:
            return 0
        return math.ceil(stock * 0.5)
    except (ValueError, TypeError):
        return 0

def procesar_logica_falabella(df_wp, df_marcas_maestro):
    logs = []

    # --- Validar columnas WP requeridas ---
    columnas_requeridas_wp = [
        'SKU', 'Nombre', 'Descripción', 'Marcas',
        'Inventario', 'Imágenes',
        'Peso (kg)', 'Longitud (cm)', 'Anchura (cm)', 'Altura (cm)'
    ]
    columnas_faltantes_wp = [col for col in columnas_requeridas_wp if col not in df_wp.columns]
    if columnas_faltantes_wp:
        logs.append((
            "❌ ERROR: El CSV de WordPress no tiene la estructura correcta.\nFaltan:\n" +
            "\n".join([f"  ⚠️ '{c}'" for c in columnas_faltantes_wp]), "error"
        ))
        return None, logs, True

    # --- Validar columnas de marcas ---
    if 'MARCA WP' not in df_marcas_maestro.columns or 'MARCA FALABELLA' not in df_marcas_maestro.columns:
        logs.append((
            "❌ ERROR: La tabla de marcas debe tener las columnas 'MARCA WP' y 'MARCA FALABELLA'.", "error"
        ))
        return None, logs, True

    # --- Diccionario de marcas Falabella ---
    dict_marcas_fal = dict(zip(
        df_marcas_maestro['MARCA WP'].astype(str).str.strip().str.lower(),
        df_marcas_maestro['MARCA FALABELLA'].astype(str).str.strip()
    ))

    # --- Validar que TODAS las marcas del CSV tengan equivalencia Falabella ---
    marcas_en_wp = df_wp['Marcas'].dropna().unique()
    marcas_sin_equivalencia = [
        str(m).strip() for m in marcas_en_wp
        if str(m).strip().lower() not in dict_marcas_fal
    ]
    if marcas_sin_equivalencia:
        logs.append((
            "❌ PROCESO DETENIDO: Las siguientes marcas no tienen equivalencia en la columna 'MARCA FALABELLA' de tu tabla:\n" +
            "\n".join([f"  • {m}" for m in marcas_sin_equivalencia]) +
            "\n\nAgrega estas marcas a tu tabla de equivalencias y vuelve a intentarlo.",
            "error"
        ))
        return None, logs, True

    logs.append(("✨ Aplicando transformaciones Falabella...", "info"))

    hoy = datetime.now().strftime('%Y-%m-%d')
    n = len(df_wp)

    # Precio origen: rebajado si existe, si no normal
    precio_rebajado = pd.to_numeric(df_wp.get('Precio rebajado', pd.Series([0]*n)), errors='coerce').fillna(0)
    precio_normal   = pd.to_numeric(df_wp.get('Precio normal',   pd.Series([0]*n)), errors='coerce').fillna(0)
    precio_origen   = precio_rebajado.where(precio_rebajado > 0, precio_normal)

    # Calcular precios — siempre enteros sin decimales
    sale_prices = precio_origen.apply(calcular_sale_price_falabella).astype(int)
    list_prices = (sale_prices * 1.5).apply(math.ceil).astype(int)

    # Calcular stock — siempre entero sin decimales
    stocks = df_wp['Inventario'].apply(calcular_stock_falabella).astype(int)

    # Mapeo de marcas (ya validado arriba, no puede llegar una sin equivalencia)
    def mapear_marca_fal(marca_wp):
        key = str(marca_wp).strip().lower()
        return dict_marcas_fal.get(key, "")

    marcas_fal = df_wp['Marcas'].apply(mapear_marca_fal)

    # Columna GTIN (puede no existir en todos los WP exports)
    col_gtin = 'GTIN, UPC, EAN o ISBN'
    gtin_col = df_wp[col_gtin] if col_gtin in df_wp.columns else pd.Series([''] * n)

    # Medidas con fallback
    def safe_num(series, default):
        return pd.to_numeric(series, errors='coerce').fillna(default)

    anchos  = safe_num(df_wp.get('Anchura (cm)',  pd.Series([10]*n)), 10)
    largos  = safe_num(df_wp.get('Longitud (cm)', pd.Series([10]*n)), 10)
    altos   = safe_num(df_wp.get('Altura (cm)',   pd.Series([10]*n)), 10)
    pesos   = safe_num(df_wp.get('Peso (kg)',      pd.Series([0.5]*n)), 0.5)

    # Reemplazar 0 con defaults
    anchos = anchos.where(anchos > 0, 10)
    largos = largos.where(largos > 0, 10)
    altos  = altos.where(altos  > 0, 10)
    pesos  = pesos.where(pesos  > 0, 0.5)

    # Construir filas de datos (una por producto)
    rows = []
    for i in range(n):
        # Distribuir imágenes en 8 columnas
        imgs_raw = df_wp['Imágenes'].iloc[i]
        imgs = []
        if not pd.isna(imgs_raw):
            imgs = [u.strip() for u in str(imgs_raw).split(',') if u.strip()]
        imgs_padded = (imgs + [''] * 8)[:8]

        nombre_limpio = limpiar_nombre_falabella(df_wp['Nombre'].iloc[i])
        sku_limpio    = limpiar_sku_falabella(df_wp['SKU'].iloc[i])
        desc_limpia   = limpiar_descripcion(df_wp['Descripción'].iloc[i])

        fila = [
            nombre_limpio,                  # Nombre #39
            marcas_fal.iloc[i],             # Marca #26
            "",                             # Modelo #32 — vacío
            desc_limpia,                    # Descripción #53
            "2316 - Juguetes y juegos / Muñecas|marionetas|figuras de acción|peluches / Muñecos de acción no eléctricos",  # Categoría primaria #1
            "",                             # País de producción #59 — se deja vacío
            sku_limpio,                     # SKU del vendedor #29
            gtin_col.iloc[i] if not pd.isna(gtin_col.iloc[i]) else "",  # Código de barras #56
            "",                             # Variación #1312
            int(stocks.iloc[i]),            # QuantityFalabella #25
            int(list_prices.iloc[i]),       # PriceFalabella #52
            int(sale_prices.iloc[i]),       # SalePriceFalabella #18
            hoy,                            # SaleStartDateFalabella #45
            "2050-01-01",                   # SaleEndDateFalabella #31
            "Figura coleccionable",         # TipoDeFigura #334136
            "Todas las edades",             # GrupoDeEdad #1301
            "Sí",                           # PiezasPequenas #1302
            "",                             # CaracteristicasDeSalud #1303
            "",                             # Color #1317
            "",                             # Material #1322
            "",                             # Personajes #1313
            "",                             # WeightOfTheProduct #1306
            "Nuevo",                        # Condición del Producto #22
            "",                             # NameCn #133815
            "",                             # NameEn #133816
            "",                             # Detalles de la condición #49
            "",                             # Garantía del producto #35
            "",                             # Garantía del vendedor #9
            "",                             # Contenido del paquete #19
            anchos.iloc[i],                 # Ancho del paquete #60
            largos.iloc[i],                 # Largo del paquete #33
            altos.iloc[i],                  # Alto del paquete #47
            pesos.iloc[i],                  # Peso del paquete #8
        ] + imgs_padded                     # Imagen principal #IM1 … Imagen8 #IM8

        rows.append(fila)

    # Armar DataFrame de productos (sin cabeceras propias)
    df_productos = pd.DataFrame(rows)

    logs.append((f"✅ Falabella: {n} productos procesados con estructura de 4 filas técnicas.", "success"))
    return df_productos, logs, False


def generar_excel_falabella(df_productos, plantilla_bytes=None):
    """
    Si se recibe la plantilla original (bytes), la carga como base y escribe
    los productos desde la fila 5, preservando las 4 filas de cabecera con
    todo su formato, colores y anchos de columna originales.
    Si no se recibe plantilla, construye el archivo desde cero (fallback).
    """
    import openpyxl as _openpyxl

    output = io.BytesIO()

    if plantilla_bytes is not None:
        wb = _openpyxl.load_workbook(io.BytesIO(plantilla_bytes))
        ws = wb['Subir plantilla']
        # Limpiar cualquier dato previo desde fila 5
        for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
            for cell in row:
                cell.value = None
        # Inyectar productos fila por fila desde la fila 5
        for row_idx, row_data in enumerate(df_productos.itertuples(index=False), start=5):
            for col_idx, value in enumerate(row_data, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value if value != '' else None)
        wb.save(output)
    else:
        # Fallback: construir desde cero
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_header = pd.DataFrame([
                FALABELLA_FILA1,
                FALABELLA_FILA2,
                FALABELLA_FILA3,
                FALABELLA_FILA4,
            ])
            df_header.to_excel(writer, index=False, header=False, sheet_name='Subir plantilla')
            df_productos.to_excel(
                writer, index=False, header=False,
                sheet_name='Subir plantilla', startrow=4
            )

    return output.getvalue()

# =====================================================================
# 4. INTERFAZ GRÁFICA DE USUARIO (GUI DE STREAMLIT)
# =====================================================================

st.markdown("""
    <div style="background: linear-gradient(135deg, #1e293b, #0f172a); padding: 25px; border-radius: 12px; margin-bottom: 25px; text-align: center; color: white; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);">
        <h1 style="margin: 0; font-size: 28px; font-weight: 700; letter-spacing: -0.025em; font-family: 'Inter', sans-serif;">E-Commerce Matrix Connector v3.0</h1>
        <p style="margin: 8px 0 0 0; color: #94a3b8; font-size: 15px;">Panel de Conversión de Catálogos WordPress → Shopstar (97 col.) & Falabella</p>
    </div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.image("https://img.icons8.com/clouds/100/null/data-configuration.png", width=100)
    st.markdown("### Guía de Uso")
    st.markdown("""
    1. **Equivalencias:** Sube tu Excel (`.xlsx`) con una hoja `MARCAS` que contenga las columnas:
       - `MARCA WP` → nombre en WordPress
       - `MARCA SS` → nombre oficial en Shopstar
       - `MARCA FALABELLA` → nombre oficial en Falabella
    2. **Catálogo WP:** Sube el CSV (`.csv`) exportado de WordPress.
    3. **Canal:** Pulsa el botón del canal que deseas generar:
       - 📦 **Shopstar** → plantilla de 97 columnas
       - 🔥 **Falabella** → plantilla con 4 filas técnicas + productos
    """)
    st.divider()
    st.info("💡 **Calford Import — v3.0**\nShopstar intacto · Motor Falabella añadido.")

col1, col2 = st.columns(2)

with col1:
    st.markdown("##### 📁 1. Tabla de Equivalencias de Marcas")
    file_maestro = st.file_uploader(
        "Sube el Excel de Marcas (.xlsx / .csv)",
        type=["xlsx", "csv"], key="maestro"
    )

with col2:
    st.markdown("##### 📝 2. Catálogo WordPress")
    file_wp = st.file_uploader(
        "Sube el archivo de WordPress (.csv / .xlsx)",
        type=["csv", "xlsx"], key="wordpress"
    )

st.markdown("##### 🗂️ 3. Plantilla oficial Falabella *(para Falabella — opcional pero recomendado)*")
file_plantilla_fal = st.file_uploader(
    "Sube la plantilla limpia de Falabella (.xlsx) para preservar formato exacto",
    type=["xlsx"], key="plantilla_falabella"
)

st.divider()

# Botones de canal — se muestran siempre; la lógica valida que haya archivos
col_btn1, col_btn2 = st.columns(2)
btn_shopstar  = col_btn1.button("📦 Generar Formato Shopstar",  use_container_width=True)
btn_falabella = col_btn2.button("🔥 Generar Formato Falabella", use_container_width=True)

# ── Helper para cargar archivos ──────────────────────────────────────
def cargar_df(file_obj):
    """Carga un archivo subido (csv o xlsx) como DataFrame."""
    if file_obj is None:
        return None
    name = file_obj.name.lower()
    if name.endswith('.xlsx'):
        try:
            return pd.read_excel(file_obj, sheet_name='MARCAS')
        except Exception:
            file_obj.seek(0)
            return pd.read_excel(file_obj, sheet_name=0)
    else:
        return pd.read_csv(file_obj, sep=',', dtype=str)

def cargar_wp(file_obj):
    """Carga el reporte de WordPress (csv o xlsx)."""
    if file_obj is None:
        return None
    name = file_obj.name.lower()
    if name.endswith('.xlsx'):
        df = pd.read_excel(file_obj, dtype=str)
    else:
        df = pd.read_csv(file_obj, sep=',', dtype=str)
    # Convertir columnas numéricas
    for col in ['Inventario', 'Precio normal', 'Precio rebajado',
                'Peso (kg)', 'Longitud (cm)', 'Anchura (cm)', 'Altura (cm)']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df

# ── CANAL SHOPSTAR ───────────────────────────────────────────────────
if btn_shopstar:
    if not file_maestro or not file_wp:
        st.warning("⚠️ Sube ambos archivos antes de generar.")
    else:
        st.markdown("##### ⚙️ Registro de Actividad — Shopstar:")
        with st.spinner("Procesando catálogo Shopstar..."):
            try:
                df_marcas  = cargar_df(file_maestro)
                df_wp_data = cargar_wp(file_wp)

                df_resultado, logs, has_error = procesar_logica_shopstar(df_wp_data, df_marcas)

                for msg, tipo in logs:
                    if tipo == "error":
                        st.error(msg)
                    elif tipo == "success":
                        st.success(msg)
                    else:
                        st.info(msg)

                if not has_error and df_resultado is not None:
                    st.markdown(f"##### 🔍 Vista Previa ({len(df_resultado.columns)} columnas):")
                    st.dataframe(df_resultado.head(5), use_container_width=True)

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_resultado.to_excel(writer, index=False)
                    excel_data = output.getvalue()

                    nombre_archivo = f"Plantilla_Shopstar_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    st.markdown("##### 📥 Listo para descargar:")
                    st.download_button(
                        label="🚀 Descargar Plantilla Shopstar (.xlsx)",
                        data=excel_data,
                        file_name=nombre_archivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"❌ ERROR CRÍTICO AL PROCESAR SHOPSTAR: {str(e)}")

# ── CANAL FALABELLA ──────────────────────────────────────────────────
if btn_falabella:
    if not file_maestro or not file_wp:
        st.warning("⚠️ Sube ambos archivos antes de generar.")
    else:
        st.markdown("##### ⚙️ Registro de Actividad — Falabella:")
        with st.spinner("Procesando catálogo Falabella..."):
            try:
                # Resetear punteros por si Shopstar ya los leyó
                file_maestro.seek(0)
                file_wp.seek(0)

                df_marcas  = cargar_df(file_maestro)
                df_wp_data = cargar_wp(file_wp)

                df_productos, logs, has_error = procesar_logica_falabella(df_wp_data, df_marcas)

                for msg, tipo in logs:
                    if tipo == "error":
                        st.error(msg)
                    elif tipo == "success":
                        st.success(msg)
                    else:
                        st.info(msg)

                if not has_error and df_productos is not None:
                    # Vista previa (sin los encabezados técnicos, solo productos)
                    st.markdown(f"##### 🔍 Vista Previa de productos ({len(df_productos.columns)} columnas):")
                    preview = df_productos.copy()
                    preview.columns = FALABELLA_FILA3
                    st.dataframe(preview.head(5), use_container_width=True)

                    # Leer bytes de la plantilla si se subió
                    plantilla_bytes = None
                    if file_plantilla_fal is not None:
                        file_plantilla_fal.seek(0)
                        plantilla_bytes = file_plantilla_fal.read()
                        st.info("✅ Usando plantilla oficial de Falabella como base — formato preservado.")
                    else:
                        st.warning("⚠️ No se subió plantilla de Falabella. Se generará el archivo sin formato de colores.")

                    excel_data = generar_excel_falabella(df_productos, plantilla_bytes)
                    nombre_archivo = f"Plantilla_Falabella_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    st.markdown("##### 📥 Listo para descargar:")
                    st.download_button(
                        label="🔥 Descargar Plantilla Falabella (.xlsx)",
                        data=excel_data,
                        file_name=nombre_archivo,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"❌ ERROR CRÍTICO AL PROCESAR FALABELLA: {str(e)}")

# ── Estado inicial (sin archivos) ────────────────────────────────────
if not file_maestro and not file_wp:
    st.info("⬆️ Sube los dos archivos en el panel superior para habilitar los botones.")
