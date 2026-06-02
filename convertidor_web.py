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
# 1. LÓGICA DE LIMPIEZA Y NEGOCIO (EL MOTOR)
# =====================================================================

# Patrón amplio que cubre todos los bloques de emojis y símbolos Unicode
_EMOJI_PATTERN = re.compile(
    "["
    "\U0001F000-\U0001FFFF"   # Todos los emojis (cartas, caras, objetos, símbolos, etc.)
    "\U00002000-\U000027FF"   # Símbolos generales, flechas, operadores, Dingbats
    "\U00002B00-\U00002BFF"   # Símbolos misceláneos y flechas
    "\U0000FE00-\U0000FE0F"   # Selectores de variantes
    "\U0000FFF0-\U0000FFFF"   # Especiales
    "\u200b-\u200f"           # Espacios de ancho cero y control de dirección
    "\u2028-\u202f"           # Separadores y espacios especiales
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
    # 1) Reemplazar \\n (literal backslash-n del CSV) por <br>
    texto = texto.replace('\\n', '<br>')
    # 2) _x000D_ es el \r (carriage return) que Excel conserva como salto real
    texto = texto.replace('_x000D_', '\n')
    # 3) Eliminar emojis
    texto = eliminar_emojis(texto)
    # 4) Limpiar espacios y tabulaciones sobrantes sin romper el HTML
    texto = re.sub(r'[ \t]+', ' ', texto).strip()
    return texto

def calcular_stock(val):
    try:
        stock = float(val)
        if pd.isna(stock) or stock <= 0:
            return 0
        if stock in [1, 2]:
            return 1
        return math.ceil(stock * 0.25)
    except (ValueError, TypeError):
        return 0

def calcular_precio_especial(val):
    try:
        precio = float(val)
        if pd.isna(precio) or precio <= 0:
            return 0
        precio_ajustado = precio * 1.26
        return math.ceil((precio_ajustado - 9) / 10) * 10 + 9
    except (ValueError, TypeError):
        return 0

# Las 97 columnas exactas de la plantilla Shopstar (en orden)
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

# =====================================================================
# 2. PROCESADOR PRINCIPAL
# =====================================================================

def procesar_logica_shopstar(df_wp, df_marcas_maestro):
    logs = []
    
    # Auditoría columnas WordPress
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

    # Auditoría tabla de marcas
    if 'MARCA WP' not in df_marcas_maestro.columns or 'MARCA SS' not in df_marcas_maestro.columns:
        logs.append(("❌ ERROR: La tabla de marcas debe tener los encabezados 'MARCA WP' y 'MARCA SS'.", "error"))
        return None, logs, True

    # Diccionario de marcas
    dict_marcas = dict(zip(
        df_marcas_maestro['MARCA WP'].astype(str).str.strip().str.lower(),
        df_marcas_maestro['MARCA SS'].astype(str).str.strip()
    ))

    # Validar marcas
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

    # Construir df con las columnas exactas de la plantilla, vacías por defecto
    df_out = pd.DataFrame('', index=range(n), columns=COLUMNAS_PLANTILLA)

    # Rellenar columnas calculadas
    df_out['Link Imagenes']         = df_wp['Imágenes'].fillna('')
    df_out['Categoria']             = '1038-Infantil/Juguetes/Coleccionables'
    df_out['Nombre Producto']       = df_wp['Nombre'].apply(limpiar_nombre)
    df_out['Nombre SKU']            = df_out['Nombre Producto']
    df_out['SKU']                   = df_wp['SKU'].fillna('')
    df_out['Descripcion']           = df_wp['Descripción'].apply(limpiar_descripcion)
    df_out['Marca']                 = df_wp['Marcas'].astype(str).str.strip().str.lower().map(dict_marcas)
    df_out['Keywords']              = ''
    df_out['Peso']                  = df_wp['Peso (kg)'].fillna(0)
    df_out['Alto']                  = df_wp['Altura (cm)'].fillna(0)
    df_out['Ancho']                 = df_wp['Anchura (cm)'].fillna(0)
    df_out['Largo']                 = df_wp['Longitud (cm)'].fillna(0)
    df_out['Stock']                 = df_wp['Inventario'].apply(calcular_stock)
    df_out['Precio Base']           = precio_normal
    df_out['Precio Especial']       = precio_base_especial.apply(calcular_precio_especial)
    df_out['Precio Especial Inicio'] = datetime.now().strftime('%d/%m/%Y')
    df_out['Precio Especial Hasta']  = '05/19/2050 23:24:07'

    # =CONCATENAR(MINUSC(SUSTITUIR(NombreProducto," ","-")),"-",MINUSC(SKU))
    df_out['Link'] = (
        df_out['Nombre Producto'].str.lower().str.replace(' ', '-', regex=False)
        + '-'
        + df_out['SKU'].str.lower()
    )

    logs.append((f"✅ Conversión finalizada con éxito. Se procesaron {len(df_out)} productos (97 columnas).", "success"))
    return df_out, logs, False

# =====================================================================
# 3. INTERFAZ GRÁFICA DE USUARIO (GUI DE STREAMLIT)
# =====================================================================

st.markdown("""
    <div style="background: linear-gradient(135deg, #1e293b, #0f172a); padding: 25px; border-radius: 12px; margin-bottom: 25px; text-align: center; color: white; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);">
        <h1 style="margin: 0; font-size: 28px; font-weight: 700; letter-spacing: -0.025em; font-family: 'Inter', sans-serif;">E-Commerce Matrix Connector v2.0</h1>
        <p style="margin: 8px 0 0 0; color: #94a3b8; font-size: 15px;">Panel de Conversión de Catálogos de WordPress a Shopstar (97 Columnas)</p>
    </div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.image("https://img.icons8.com/clouds/100/null/data-configuration.png", width=100)
    st.markdown("### Guía de Uso")
    st.markdown("""
    1. **Equivalencias:** Sube tu archivo Excel (`.xlsx`) que contenga una pestaña llamada `MARCAS` (o la primera hoja) con las columnas `MARCA WP` y `MARCA SS`.
    2. **Catálogo WP:** Sube el archivo CSV (`.csv`) exportado de WordPress.
    3. **Procesamiento:** La web validará las marcas e inyectará dimensiones y peso del propio catálogo de WordPress.
    4. **Descargar:** Haz clic en **Descargar Plantilla Shopstar** para obtener el archivo Excel estructurado con las 97 columnas oficiales.
    """)
    st.divider()
    st.info("💡 **Calford Import - v2.0**.\nSoporte completo para 97 columnas, eliminación de emojis y mapeo avanzado de saltos de línea.")

col1, col2 = st.columns(2)

with col1:
    st.markdown("##### 📁 1. Tabla de Equivalencias")
    file_maestro = st.file_uploader("Sube el archivo Excel de Marcas", type=["xlsx", "csv"], key="maestro")

with col2:
    st.markdown("##### 📝 2. Catálogo WordPress")
    file_wp = st.file_uploader("Sube el archivo CSV de WordPress", type=["csv"], key="wordpress")

st.divider()

if file_maestro and file_wp:
    st.markdown("##### ⚙️ Registro de Actividad:")
    
    with st.spinner("Procesando catálogos..."):
        try:
            # 1. Cargar archivo de equivalencias (marcas)
            if file_maestro.name.endswith('.xlsx'):
                try:
                    df_marcas = pd.read_excel(file_maestro, sheet_name='MARCAS')
                except Exception:
                    df_marcas = pd.read_excel(file_maestro, sheet_name=0)
            else:
                df_marcas = pd.read_csv(file_maestro, sep=',')
            
            # 2. Cargar CSV de WordPress
            df_wp_data = pd.read_csv(file_wp, sep=',', dtype=str)
            
            # Convertir columnas numéricas necesarias como en convertidor.py
            for col in ['Inventario', 'Precio normal', 'Precio rebajado',
                        'Peso (kg)', 'Longitud (cm)', 'Anchura (cm)', 'Altura (cm)']:
                if col in df_wp_data.columns:
                    df_wp_data[col] = pd.to_numeric(df_wp_data[col], errors='coerce').fillna(0)
            
            # 3. Procesar
            df_resultado, logs, has_error = procesar_logica_shopstar(df_wp_data, df_marcas)
            
            for log_msg, log_type in logs:
                if log_type == "error":
                    st.error(log_msg)
                elif log_type == "success":
                    st.success(log_msg)
                else:
                    st.info(log_msg)
                    
            if not has_error and df_resultado is not None:
                st.markdown(f"##### 🔍 Vista Previa ({len(df_resultado.columns)} columnas):")
                st.dataframe(df_resultado.head(5), use_container_width=True)
                
                # Convertir a excel en memoria
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_resultado.to_excel(writer, index=False)
                excel_data = output.getvalue()
                
                st.markdown("##### 📥 Listo para descargar:")
                nombre_archivo = f"Plantilla_Shopstar_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                
                st.download_button(
                    label="🚀 Descargar Plantilla Shopstar (.xlsx)",
                    data=excel_data,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"❌ ERROR CRÍTICO AL PROCESAR: {str(e)}")
else:
    st.warning("⚠️ Sube ambos archivos para iniciar el procesamiento.")
