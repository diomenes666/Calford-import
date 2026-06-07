import os
import math
import re
import io
from datetime import datetime
import pandas as pd
import streamlit as st
import openpyxl

# =====================================================================
# CONFIGURACIÓN DE PÁGINA
# =====================================================================
st.set_page_config(
    page_title="E-Commerce Matrix Connector",
    page_icon="🚀",
    layout="centered",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
        .stFileUploader {
            border: 2px dashed #cbd5e1; border-radius: 12px; padding: 10px;
            background-color: #f8fafc; transition: all 0.3s ease;
        }
        .stFileUploader:hover { border-color: #0284c7; background-color: #f0f9ff; }
        div.stDownloadButton > button {
            background: linear-gradient(135deg, #0284c7, #0369a1) !important;
            color: white !important; border: none !important;
            padding: 12px 24px !important; border-radius: 8px !important;
            font-weight: 600 !important; font-size: 16px !important;
            width: 100% !important;
        }
        div.stDownloadButton > button:hover {
            background: linear-gradient(135deg, #0369a1, #075985) !important;
            transform: translateY(-1px) !important;
        }
    </style>
""", unsafe_allow_html=True)

# =====================================================================
# DICCIONARIO DE CATEGORÍAS FALABELLA
# Clave  = nombre que verá el usuario en el selectbox
# Valor  = archivo dentro de plantillas_falabella/
# =====================================================================
CATEGORIAS_FALABELLA = {
    "2316 - Juguetes y juegos / Muñecos de acción no eléctricos":           "juguetes_y_juegos.xlsx",
    "1510 - Juguetes y juegos / Peluches y otras muñecas":                  "peluches.xlsx",
    "956  - Juguetes y juegos / Bloques de construcción (Lego)":            "lego.xlsx",
    "2065 - Juguetes y juegos / Juegos de cartas":                          "cartas.xlsx",
    "449  - Juguetes y juegos / Rompecabezas":                              "rompecabezas.xlsx",
    "1259 - Juguetes y juegos / Juegos de tablero":                         "juego_de_mesa.xlsx",
    "3303 - Ropa y accesorios / Pijamas":                                   "pijamas.xlsx",
    "2898 - Ropa y accesorios / Polos y camisetas":                         "polo.xlsx",
    "463  - Hogar / Ropa de cama":                                          "ropa_de_cama.xlsx",
}

# Categoría cuyo país de producción es Dinamarca (excepción a China)
CATEGORIA_LEGO = "956  - Juguetes y juegos / Bloques de construcción (Lego)"

# =====================================================================
# 1. UTILIDADES COMPARTIDAS
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

def limpiar_nombre_shopstar(texto):
    if pd.isna(texto):
        return ""
    texto_str = eliminar_emojis(str(texto))
    texto_limpio = re.sub(r'-[^-]+-', '', texto_str)
    texto_limpio = texto_limpio.replace("'", "").replace("-", "")
    return " ".join(texto_limpio.split())

def limpiar_nombre_falabella(texto):
    if pd.isna(texto):
        return ""
    s = eliminar_emojis(str(texto))
    for tag in ['-PREVENTA-', '-OFERTA-', '-LIQUIDACION-']:
        s = s.replace(tag, '')
    return " ".join(s.replace('"', '').replace("'", '').split())

def limpiar_descripcion(html_texto):
    if pd.isna(html_texto):
        return ""
    texto = str(html_texto)
    texto = texto.replace('\\n', '<br>')
    texto = texto.replace('_x000D_', '\n')
    texto = eliminar_emojis(texto)
    texto = re.sub(r'[ \t]+', ' ', texto).strip()
    return texto

def limpiar_imagenes(val):
    if pd.isna(val):
        return ""
    enlaces = [url.strip() for url in str(val).split(',') if url.strip()]
    return ",".join(enlaces)

def limpiar_sku_falabella(val):
    if pd.isna(val):
        return ""
    return re.sub(r'[^A-Za-z0-9]', '', str(val))

def safe_num(val, default):
    try:
        v = float(val)
        return v if not pd.isna(v) and v > 0 else default
    except (ValueError, TypeError):
        return default

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

def procesar_logica_shopstar(df_wp, df_marcas_maestro):
    logs = []

    columnas_requeridas_wp = [
        'SKU', 'Nombre', 'Descripción', 'Marcas',
        'Inventario', 'Precio normal', 'Imágenes', 'URL externa',
        'Peso (kg)', 'Longitud (cm)', 'Anchura (cm)', 'Altura (cm)'
    ]
    columnas_faltantes_wp = [col for col in columnas_requeridas_wp if col not in df_wp.columns]
    if columnas_faltantes_wp:
        logs.append((
            "❌ ERROR: El CSV de WordPress no tiene la estructura correcta.\nFaltan:\n" +
            "\n".join([f"  ⚠️ '{c}'" for c in columnas_faltantes_wp]), "error"
        ))
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
        logs.append((
            "❌ PROCESO DETENIDO: Marcas nuevas sin registrar en equivalencias:\n" +
            "\n".join([f"  • {m}" for m in marcas_faltantes]), "error"
        ))
        return None, logs, True

    logs.append(("✨ Aplicando transformaciones y limpiando textos...", "info"))
    n = len(df_wp)

    precio_para_especial = df_wp['Precio rebajado'].fillna(0).astype(float)
    precio_normal        = df_wp['Precio normal'].fillna(0).astype(float)
    precio_base_especial = precio_para_especial.where(precio_para_especial > 0, precio_normal)

    df_out = pd.DataFrame('', index=range(n), columns=COLUMNAS_PLANTILLA)

    df_out['Link Imagenes']          = df_wp['Imágenes'].apply(limpiar_imagenes)
    df_out['Categoria']              = '1038-Infantil/Juguetes/Coleccionables'
    df_out['Nombre Producto']        = df_wp['Nombre'].apply(limpiar_nombre_shopstar)
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
    df_out['Link']                   = (
        df_out['Nombre Producto'].str.lower().str.replace(' ', '-', regex=False)
        + '-' + df_out['SKU'].str.lower()
    )

    logs.append((f"✅ Conversión finalizada con éxito. Se procesaron {len(df_out)} productos (97 columnas).", "success"))
    return df_out, logs, False

# =====================================================================
# 3. MOTOR FALABELLA — MULTICATEGORÍA CON MAPEO POR NOMBRE
# =====================================================================

def calcular_sale_price_falabella(val):
    try:
        p = float(val)
        if pd.isna(p) or p <= 0:
            return 0
        return int(round(p * 1.26 / 10) * 10 - 1)
    except (ValueError, TypeError):
        return 0

def calcular_stock_falabella(val):
    try:
        s = float(val)
        if pd.isna(s) or s <= 0:
            return 0
        return math.ceil(s * 0.5)
    except (ValueError, TypeError):
        return 0

def procesar_logica_falabella(df_wp, df_marcas_maestro, categoria_sel):
    logs = []

    # --- Ruta dinámica a la plantilla ---
    nombre_archivo = CATEGORIAS_FALABELLA[categoria_sel]
    ruta_plantilla = os.path.join("plantillas_falabella", nombre_archivo)

    try:
        wb_base = openpyxl.load_workbook(ruta_plantilla)
    except FileNotFoundError:
        logs.append((
            f"❌ No se encontró la plantilla '{ruta_plantilla}'.\n"
            "Asegúrate de subir la carpeta 'plantillas_falabella/' junto al script en GitHub.",
            "error"
        ))
        return None, None, logs, True

    ws_base = wb_base['Subir plantilla']

    # --- Extraer nombres técnicos de la fila 4 (índice openpyxl = fila 4) ---
    n_cols      = ws_base.max_column
    col_nombres = [ws_base.cell(row=4, column=c).value or "" for c in range(1, n_cols + 1)]

    # --- Leer categoría primaria desde la hoja Categorías ---
    ws_cat         = wb_base['Categorías']
    categoria_prim = ""
    for row in ws_cat.iter_rows(max_row=10, max_col=3, values_only=True):
        for cell in row:
            if cell and str(cell).strip() and str(cell).strip() != "PrimaryCategory":
                categoria_prim = str(cell).strip()
                break
        if categoria_prim:
            break

    # --- Validaciones previas ---
    columnas_requeridas_wp = ['SKU', 'Nombre', 'Descripción', 'Marcas', 'Inventario', 'Imágenes']
    faltantes = [c for c in columnas_requeridas_wp if c not in df_wp.columns]
    if faltantes:
        logs.append((
            "❌ ERROR: Faltan columnas en el CSV de WordPress:\n" +
            "\n".join([f"  ⚠️ '{c}'" for c in faltantes]), "error"
        ))
        return None, None, logs, True

    if 'MARCA WP' not in df_marcas_maestro.columns or 'MARCA FALABELLA' not in df_marcas_maestro.columns:
        logs.append(("❌ ERROR: La tabla de marcas debe tener columnas 'MARCA WP' y 'MARCA FALABELLA'.", "error"))
        return None, None, logs, True

    dict_marcas = dict(zip(
        df_marcas_maestro['MARCA WP'].astype(str).str.strip().str.lower(),
        df_marcas_maestro['MARCA FALABELLA'].astype(str).str.strip()
    ))

    marcas_en_wp         = df_wp['Marcas'].dropna().unique()
    marcas_sin_equiv     = [str(m).strip() for m in marcas_en_wp if str(m).strip().lower() not in dict_marcas]
    if marcas_sin_equiv:
        logs.append((
            "❌ PROCESO DETENIDO: Marcas sin equivalencia en 'MARCA FALABELLA':\n" +
            "\n".join([f"  • {m}" for m in marcas_sin_equiv]) +
            "\n\nAgrega estas marcas a tu tabla de equivalencias y vuelve a intentarlo.",
            "error"
        ))
        return None, None, logs, True

    logs.append((f"✨ Procesando con plantilla: {nombre_archivo} ({n_cols} columnas)...", "info"))

    hoy = datetime.now().strftime('%Y-%m-%d')
    n   = len(df_wp)

    # Precio origen: rebajado si existe, si no normal
    precio_rebajado = pd.to_numeric(df_wp.get('Precio rebajado', pd.Series([0]*n)), errors='coerce').fillna(0)
    precio_normal   = pd.to_numeric(df_wp.get('Precio normal',   pd.Series([0]*n)), errors='coerce').fillna(0)
    precio_origen   = precio_rebajado.where(precio_rebajado > 0, precio_normal)

    sale_prices = precio_origen.apply(calcular_sale_price_falabella).astype(int)
    list_prices = (sale_prices * 1.5).apply(math.ceil).astype(int)
    stocks      = df_wp['Inventario'].apply(calcular_stock_falabella).astype(int)

    # País según categoría
    pais_produccion = "Dinamarca" if categoria_sel == CATEGORIA_LEGO else "China"

    # Columna GTIN
    col_gtin = 'GTIN, UPC, EAN o ISBN'
    gtin_series = df_wp[col_gtin] if col_gtin in df_wp.columns else pd.Series([''] * n)

    # ── Construir filas producto ──────────────────────────────────────
    filas_productos = []

    for i in range(n):
        # Inicializar todas las celdas vacías indexadas por nombre técnico
        fila = {col: "" for col in col_nombres}

        # Distribuir imágenes en columnas IM1…IM8
        imgs_raw    = df_wp['Imágenes'].iloc[i]
        imgs        = [u.strip() for u in str(imgs_raw).split(',') if u.strip()] if not pd.isna(imgs_raw) else []
        imgs_padded = (imgs + [''] * 8)[:8]

        # Medidas con fallback
        ancho_pkg = safe_num(df_wp.get('Anchura (cm)',  pd.Series([10]*n)).iloc[i], 10)
        largo_pkg = safe_num(df_wp.get('Longitud (cm)', pd.Series([10]*n)).iloc[i], 10)
        alto_pkg  = safe_num(df_wp.get('Altura (cm)',   pd.Series([10]*n)).iloc[i], 10)
        peso_pkg  = safe_num(df_wp.get('Peso (kg)',     pd.Series([0.5]*n)).iloc[i], 0.5)

        # ── MAPEO UNIVERSAL POR NOMBRE (coincidencia parcial) ─────────
        img_idx = 0
        for col in col_nombres:
            if not col:
                continue

            # Nombre
            if col.startswith("Nombre #"):
                fila[col] = limpiar_nombre_falabella(df_wp['Nombre'].iloc[i])

            # Marca
            elif col.startswith("Marca #"):
                key = str(df_wp['Marcas'].iloc[i]).strip().lower()
                fila[col] = dict_marcas.get(key, "")

            # Modelo — siempre vacío
            elif col.startswith("Modelo #"):
                fila[col] = ""

            # Descripción
            elif col.startswith("Descripción #"):
                fila[col] = limpiar_descripcion(df_wp['Descripción'].iloc[i])

            # Categoría primaria
            elif "Categoría primaria" in col:
                fila[col] = categoria_prim

            # País de producción
            elif "País de producción" in col:
                fila[col] = pais_produccion

            # SKU del vendedor
            elif "SKU del vendedor" in col:
                fila[col] = limpiar_sku_falabella(df_wp['SKU'].iloc[i])

            # Código de barras
            elif "Código de barras" in col:
                gtin_val = gtin_series.iloc[i]
                fila[col] = "" if pd.isna(gtin_val) else str(gtin_val)

            # Variación — siempre '...'
            elif col.startswith("Variación #") or col.startswith("Variacion #"):
                fila[col] = "..."

            # Stock
            elif "QuantityFalabella" in col:
                fila[col] = int(stocks.iloc[i])

            # Precio de lista
            elif "PriceFalabella" in col:
                fila[col] = int(list_prices.iloc[i])

            # Precio rebajado (sale)
            elif "SalePriceFalabella" in col:
                fila[col] = int(sale_prices.iloc[i])

            # Fecha inicio
            elif "SaleStartDateFalabella" in col:
                fila[col] = hoy

            # Fecha fin
            elif "SaleEndDateFalabella" in col:
                fila[col] = "2050-01-01"

            # Condición del Producto
            elif "Condición del Producto" in col:
                fila[col] = "Nuevo"

            # GrupoDeEdad
            elif "GrupoDeEdad" in col:
                fila[col] = "Todas las edades"

            # PiezasPequenas
            elif "PiezasPequenas" in col:
                fila[col] = "Sí"

            # CaracteristicasDeSalud — siempre "Sin BPA" en todas las categorías
            elif "CaracteristicasDeSalud" in col:
                fila[col] = "Sin BPA"

            # WeightOfTheProduct = Peso del paquete
            elif "WeightOfTheProduct" in col:
                fila[col] = peso_pkg

            # Medidas del paquete
            elif "Ancho del paquete" in col:
                fila[col] = ancho_pkg
            elif "Largo del paquete" in col:
                fila[col] = largo_pkg
            elif "Alto del paquete" in col:
                fila[col] = alto_pkg
            elif "Peso del paquete" in col:
                fila[col] = peso_pkg

            # Imágenes secuenciales
            elif "Imagen principal" in col or col.startswith("Imagen") and "#IM" in col:
                fila[col] = imgs_padded[img_idx] if img_idx < 8 else ""
                img_idx += 1

        # ── EXCEPCIONES POR CATEGORÍA (valores adicionales específicos) ──
        # Bloque escalable: agregar aquí nuevas reglas sin romper otras categorías

        # (reservado para futuras columnas como Color, Material, etc.)

        filas_productos.append(fila)

    # Convertir lista de dicts a DataFrame con las columnas en el orden exacto de la plantilla
    df_productos = pd.DataFrame(filas_productos, columns=col_nombres)

    logs.append((
        f"✅ Falabella: {n} productos procesados · {n_cols} columnas · categoría: {categoria_prim}",
        "success"
    ))
    return df_productos, wb_base, logs, False


def generar_excel_falabella(df_productos, wb_base):
    """
    Toma el workbook original como base (preservando formato, colores, filas 1-4)
    e inyecta los productos desde la fila 5 en adelante.
    """
    ws = wb_base['Subir plantilla']

    # Limpiar filas previas desde fila 5
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
        for cell in row:
            cell.value = None

    # Inyectar productos
    for row_idx, row_data in enumerate(df_productos.itertuples(index=False), start=5):
        for col_idx, value in enumerate(row_data, start=1):
            ws.cell(row=row_idx, column=col_idx, value=value if value != '' else None)

    output = io.BytesIO()
    wb_base.save(output)
    return output.getvalue()


# =====================================================================
# 4. CARGADORES DE ARCHIVOS
# =====================================================================

def cargar_maestro(file_obj):
    if file_obj is None:
        return None
    name = file_obj.name.lower()
    if name.endswith('.xlsx'):
        try:
            return pd.read_excel(file_obj, sheet_name='MARCAS')
        except Exception:
            file_obj.seek(0)
            return pd.read_excel(file_obj, sheet_name=0)
    return pd.read_csv(file_obj, sep=',', dtype=str)

def cargar_wp(file_obj):
    if file_obj is None:
        return None
    name = file_obj.name.lower()
    df = pd.read_excel(file_obj, dtype=str) if name.endswith('.xlsx') else pd.read_csv(file_obj, sep=',', dtype=str)
    for col in ['Inventario', 'Precio normal', 'Precio rebajado',
                'Peso (kg)', 'Longitud (cm)', 'Anchura (cm)', 'Altura (cm)']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df


# =====================================================================
# 5. INTERFAZ DE USUARIO
# =====================================================================

st.markdown("""
    <div style="background: linear-gradient(135deg, #1e293b, #0f172a); padding: 25px;
                border-radius: 12px; margin-bottom: 25px; text-align: center; color: white;">
        <h1 style="margin: 0; font-size: 28px; font-weight: 700; font-family: 'Inter', sans-serif;">
            E-Commerce Matrix Connector v4.0
        </h1>
        <p style="margin: 8px 0 0 0; color: #94a3b8; font-size: 15px;">
            WordPress → Shopstar (97 col.) &nbsp;|&nbsp; Falabella Multicategoría
        </p>
    </div>
""", unsafe_allow_html=True)

# ── SIDEBAR ──────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://img.icons8.com/clouds/100/null/data-configuration.png", width=100)
    st.markdown("### Guía de Uso")
    st.markdown("""
    1. **Marcas:** Sube el Excel con hoja `MARCAS` que contenga `MARCA WP`, `MARCA SS` y `MARCA FALABELLA`.
    2. **Catálogo WP:** Sube el CSV/XLSX exportado de WordPress.
    3. **Categoría Falabella:** Selecciona la categoría del lote que vas a procesar.
    4. **Canal:** Pulsa el botón del canal deseado.
    """)
    st.divider()

    st.markdown("##### 🗂️ Categoría Falabella")
    categoria_seleccionada = st.selectbox(
        "Selecciona la categoría del lote actual:",
        options=list(CATEGORIAS_FALABELLA.keys()),
        key="categoria_fal"
    )
    st.divider()
    st.info("💡 **Calford Import — v4.0**\nShopstar intacto · Motor Falabella multicategoría.")

# ── UPLOADERS ─────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    st.markdown("##### 📁 1. Tabla de Equivalencias de Marcas")
    file_maestro = st.file_uploader(
        "Excel de marcas (.xlsx / .csv)",
        type=["xlsx", "csv"], key="maestro"
    )

with col2:
    st.markdown("##### 📝 2. Catálogo WordPress")
    file_wp = st.file_uploader(
        "Archivo WordPress (.csv / .xlsx)",
        type=["csv", "xlsx"], key="wordpress"
    )

st.divider()

# ── BOTONES DE CANAL ──────────────────────────────────────────────────
col_btn1, col_btn2 = st.columns(2)
btn_shopstar  = col_btn1.button("📦 Generar Formato Shopstar",  use_container_width=True)
btn_falabella = col_btn2.button("🔥 Generar Formato Falabella", use_container_width=True)

# ── CANAL SHOPSTAR ────────────────────────────────────────────────────
if btn_shopstar:
    if not file_maestro or not file_wp:
        st.warning("⚠️ Sube ambos archivos antes de generar.")
    else:
        st.markdown("##### ⚙️ Registro de Actividad — Shopstar:")
        with st.spinner("Procesando catálogo Shopstar..."):
            try:
                df_marcas  = cargar_maestro(file_maestro)
                df_wp_data = cargar_wp(file_wp)
                df_resultado, logs, has_error = procesar_logica_shopstar(df_wp_data, df_marcas)

                for msg, tipo in logs:
                    if tipo == "error":   st.error(msg)
                    elif tipo == "success": st.success(msg)
                    else: st.info(msg)

                if not has_error and df_resultado is not None:
                    st.markdown(f"##### 🔍 Vista Previa ({len(df_resultado.columns)} columnas):")
                    st.dataframe(df_resultado.head(5), use_container_width=True)

                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_resultado.to_excel(writer, index=False)

                    st.markdown("##### 📥 Listo para descargar:")
                    st.download_button(
                        label="🚀 Descargar Plantilla Shopstar (.xlsx)",
                        data=output.getvalue(),
                        file_name=f"Plantilla_Shopstar_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"❌ ERROR CRÍTICO AL PROCESAR SHOPSTAR: {str(e)}")

# ── CANAL FALABELLA ───────────────────────────────────────────────────
if btn_falabella:
    if not file_maestro or not file_wp:
        st.warning("⚠️ Sube ambos archivos antes de generar.")
    else:
        st.markdown(f"##### ⚙️ Registro de Actividad — Falabella ({categoria_seleccionada}):")
        with st.spinner("Procesando catálogo Falabella..."):
            try:
                file_maestro.seek(0)
                file_wp.seek(0)

                df_marcas  = cargar_maestro(file_maestro)
                df_wp_data = cargar_wp(file_wp)

                df_productos, wb_base, logs, has_error = procesar_logica_falabella(
                    df_wp_data, df_marcas, categoria_seleccionada
                )

                for msg, tipo in logs:
                    if tipo == "error":   st.error(msg)
                    elif tipo == "success": st.success(msg)
                    else: st.info(msg)

                if not has_error and df_productos is not None:
                    st.markdown(f"##### 🔍 Vista Previa ({len(df_productos.columns)} columnas):")
                    st.dataframe(df_productos.head(5), use_container_width=True)

                    excel_data = generar_excel_falabella(df_productos, wb_base)
                    st.markdown("##### 📥 Listo para descargar:")
                    st.download_button(
                        label="🔥 Descargar Plantilla Falabella (.xlsx)",
                        data=excel_data,
                        file_name=f"Plantilla_Falabella_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"❌ ERROR CRÍTICO AL PROCESAR FALABELLA: {str(e)}")

# ── Estado inicial ────────────────────────────────────────────────────
if not file_maestro and not file_wp:
    st.info("⬆️ Sube los dos archivos y selecciona la categoría para habilitar los botones.")
