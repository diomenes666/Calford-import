import os
import sys
import math
import re
import unicodedata
from datetime import datetime

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import scrolledtext

try:
    import openpyxl
    import pandas as pd
except ImportError:
    root_err = tk.Tk()
    root_err.withdraw()
    messagebox.showerror("Falta Librería", "Por favor, abre la terminal de Windows (cmd) y ejecuta:\npip install pandas openpyxl")
    sys.exit()

# =====================================================================
# 1. LÓGICA DE LIMPIEZA Y NEGOCIO
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
    """
    Replica: =SUSTITUIR(BUSCARX(...), "\\n", "<br>")
    Orden correcto: primero \\n -> <br>, luego _x000D_ -> \n (salto real).
    Así _x000D_\\n_x000D_\\n produce \n<br>\n<br> igual que Excel.
    """
    if pd.isna(html_texto):
        return ""
    texto = str(html_texto)

    # 1) Reemplazar \\n (literal backslash-n del CSV) por <br>  ← igual que Excel SUSTITUIR
    texto = texto.replace('\\n', '<br>')

    # 2) _x000D_ es el \r (carriage return) que Excel conserva como salto real
    texto = texto.replace('_x000D_', '\n')

    # 3) Eliminar emojis / iconos no permitidos por Shopstar
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
    """
    Replica: =CEILING.MATH(precio_rebajado * 1.26 - 9, 10) + 9
    Usa precio rebajado si existe, si no usa precio normal.
    """
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
# 2. PROCESADOR PRINCIPAL
# =====================================================================

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

def procesar_logica_shopstar(df_wp, df_marcas_maestro, log_callback):
    # Auditoría columnas WordPress
    columnas_requeridas_wp = [
        'SKU', 'Nombre', 'Descripción', 'Marcas',
        'Inventario', 'Precio normal', 'Imágenes', 'URL externa',
        'Peso (kg)', 'Longitud (cm)', 'Anchura (cm)', 'Altura (cm)'
    ]
    columnas_faltantes_wp = [col for col in columnas_requeridas_wp if col not in df_wp.columns]
    if columnas_faltantes_wp:
        log_callback("❌ ERROR: El CSV de WordPress no tiene la estructura correcta.\nFaltan:\n" +
                     "\n".join([f"  ⚠️ '{c}'" for c in columnas_faltantes_wp]), "error")
        return None

    # Auditoría tabla de marcas
    if 'MARCA WP' not in df_marcas_maestro.columns or 'MARCA SS' not in df_marcas_maestro.columns:
        log_callback("❌ ERROR: La tabla de marcas debe tener los encabezados 'MARCA WP' y 'MARCA SS'.", "error")
        return None

    # Diccionario de marcas (case-insensitive)
    dict_marcas = dict(zip(
        df_marcas_maestro['MARCA WP'].astype(str).str.strip().str.lower(),
        df_marcas_maestro['MARCA SS'].astype(str).str.strip()
    ))

    # Validar que todas las marcas del WP estén en el maestro
    marcas_en_wp = df_wp['Marcas'].dropna().unique()
    marcas_faltantes = [str(m).strip() for m in marcas_en_wp if str(m).strip().lower() not in dict_marcas]
    if marcas_faltantes:
        log_callback("❌ PROCESO DETENIDO: Marcas nuevas sin registrar en equivalencias:\n" +
                     "\n".join([f"  • {m}" for m in marcas_faltantes]), "error")
        return None

    log_callback("✨ Aplicando transformaciones y limpiando textos...", "info")

    n = len(df_wp)

    # Precio a usar para Precio Especial: precio rebajado si existe, sino precio normal
    precio_para_especial = df_wp['Precio rebajado'].fillna(0).astype(float)
    precio_normal = df_wp['Precio normal'].fillna(0).astype(float)
    precio_base_especial = precio_para_especial.where(precio_para_especial > 0, precio_normal)

    # Construir df con las columnas exactas de la plantilla, vacías por defecto
    df_out = pd.DataFrame('', index=range(n), columns=COLUMNAS_PLANTILLA)

    # Rellenar las columnas que se calculan
    df_out['Link Imagenes']         = df_wp['Imágenes'].apply(limpiar_imagenes)
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
    df_out['Precio Especial']       = precio_base_especial.apply(calcular_precio_especial)
    df_out['Precio Base']           = (df_out['Precio Especial'] * 1.5).round(0).astype(int)
    df_out['Precio Especial Inicio'] = datetime.now().strftime('%d/%m/%Y')
    df_out['Precio Especial Hasta']  = '05/19/2050 23:24:07'

    # =CONCATENAR(MINUSC(SUSTITUIR(NombreProducto," ","-")),"-",MINUSC(SKU))
    df_out['Link'] = (
        df_out['Nombre Producto'].str.lower().str.replace(' ', '-', regex=False)
        + '-'
        + df_out['SKU'].str.lower()
    )

    return df_out

# =====================================================================
# 3. INTERFAZ GRÁFICA
# =====================================================================

class AplicacionConvertidor:
    def __init__(self, root):
        self.root = root
        self.root.title("E-Commerce Matrix Connector v2.0")
        self.root.geometry("680x520")
        self.root.configure(bg="#f4f6f9")

        lbl_titulo = tk.Label(root, text="Panel de Conversión de Catálogos", font=("Arial", 14, "bold"), bg="#1e293b", fg="white", pady=12)
        lbl_titulo.pack(fill=tk.X)

        frame_archivos = tk.LabelFrame(root, text=" Seleccionar Archivos del Proceso ", font=("Arial", 10, "bold"), bg="#f4f6f9", bd=2, padx=10, pady=10)
        frame_archivos.pack(fill=tk.X, padx=15, pady=15)

        lbl_m = tk.Label(frame_archivos, text="Tabla Equivalencias (Excel/CSV):", bg="#f4f6f9", font=("Arial", 9))
        lbl_m.grid(row=0, column=0, sticky="w", pady=6)
        self.ent_maestro = tk.Entry(frame_archivos, width=48, bd=1, font=("Arial", 9), fg="#0f172a")
        self.ent_maestro.grid(row=0, column=1, padx=5)
        btn_buscar_m = tk.Button(frame_archivos, text="Examinar...", command=self.seleccionar_equivalencias, font=("Arial", 9), bg="#cbd5e1")
        btn_buscar_m.grid(row=0, column=2)

        lbl_wp = tk.Label(frame_archivos, text="Export WordPress (.csv):", bg="#f4f6f9", font=("Arial", 9))
        lbl_wp.grid(row=1, column=0, sticky="w", pady=6)
        self.ent_wordpress = tk.Entry(frame_archivos, width=48, bd=1, font=("Arial", 9), fg="#0f172a")
        self.ent_wordpress.grid(row=1, column=1, padx=5)
        btn_buscar_wp = tk.Button(frame_archivos, text="Examinar...", command=self.seleccionar_wordpress, font=("Arial", 9), bg="#cbd5e1")
        btn_buscar_wp.grid(row=1, column=2)

        frame_canales = tk.LabelFrame(root, text=" Plataforma de Destino (Generar Formato) ", font=("Arial", 10, "bold"), bg="#f4f6f9", bd=2, padx=10, pady=10)
        frame_canales.pack(fill=tk.X, padx=15, pady=5)

        self.btn_shopstar = tk.Button(frame_canales, text="🚀 Generar Shopstar (.xlsx)", font=("Arial", 10, "bold"), bg="#0284c7", fg="white", width=24, height=2, command=self.convertir_shopstar)
        self.btn_shopstar.grid(row=0, column=0, padx=8, pady=5)

        self.btn_ripley = tk.Button(frame_canales, text="🔒 Generar Ripley", font=("Arial", 10), bg="#e2e8f0", fg="#94a3b8", width=18, state=tk.DISABLED)
        self.btn_ripley.grid(row=0, column=1, padx=8, pady=5)

        self.btn_saga = tk.Button(frame_canales, text="🔒 Generar Saga", font=("Arial", 10), bg="#e2e8f0", fg="#94a3b8", width=18, state=tk.DISABLED)
        self.btn_saga.grid(row=0, column=2, padx=8, pady=5)

        lbl_log = tk.Label(root, text="Registro de Actividad / Errores de Consistencia:", bg="#f4f6f9", font=("Arial", 9, "bold"))
        lbl_log.pack(anchor="w", padx=15, pady=(15, 0))

        self.txt_log = scrolledtext.ScrolledText(root, height=10, width=78, font=("Consolas", 9), bg="#1e1e1e", fg="#e2e8f0")
        self.txt_log.pack(padx=15, pady=5, fill=tk.BOTH, expand=True)

    def seleccionar_equivalencias(self):
        archivo = filedialog.askopenfilename(filetypes=[("Archivos de Datos", "*.xlsx *.csv")])
        if archivo:
            self.ent_maestro.delete(0, tk.END)
            self.ent_maestro.insert(0, archivo)

    def seleccionar_wordpress(self):
        archivo = filedialog.askopenfilename(filetypes=[("Archivos CSV", "*.csv")])
        if archivo:
            self.ent_wordpress.delete(0, tk.END)
            self.ent_wordpress.insert(0, archivo)

    def escribir_log(self, mensaje, tipo="info"):
        self.txt_log.insert(tk.END, mensaje + "\n")
        self.txt_log.see(tk.END)
        if tipo == "error":
            messagebox.showerror("Validación Fallida", "El proceso se detuvo debido a inconsistencias en los datos.")

    def convertir_shopstar(self):
        self.txt_log.delete('1.0', tk.END)
        file_wp = self.ent_wordpress.get()
        file_maestro = self.ent_maestro.get()

        if not file_wp or not file_maestro:
            messagebox.showwarning("Faltan Datos", "Por favor, selecciona tanto el archivo de WordPress como el de Equivalencias.")
            return

        try:
            self.escribir_log("⏳ Leyendo archivos seleccionados...")
            df_wp = pd.read_csv(file_wp, sep=',', dtype=str)

            # Convertir columnas numéricas necesarias
            for col in ['Inventario', 'Precio normal', 'Precio rebajado',
                        'Peso (kg)', 'Longitud (cm)', 'Anchura (cm)', 'Altura (cm)']:
                if col in df_wp.columns:
                    df_wp[col] = pd.to_numeric(df_wp[col], errors='coerce').fillna(0)

            # Cargar tabla de marcas (hoja MARCAS del Excel o CSV)
            if file_maestro.endswith('.xlsx'):
                try:
                    df_marcas = pd.read_excel(file_maestro, sheet_name='MARCAS')
                except Exception:
                    df_marcas = pd.read_excel(file_maestro, sheet_name=0)
            else:
                df_marcas = pd.read_csv(file_maestro, sep=',')

            resultado = procesar_logica_shopstar(df_wp, df_marcas, self.escribir_log)

            if resultado is not None:
                ruta_salida = os.path.splitext(file_wp)[0] + "_Plantilla_Shopstar.xlsx"
                resultado.to_excel(ruta_salida, index=False, engine='openpyxl')

                self.escribir_log(f"\n🎉 ¡PLANTILLA EXCEL GENERADA CON ÉXITO!\n   {len(resultado)} productos procesados.\n👉 {ruta_salida}")
                messagebox.showinfo("Completado", "El archivo Excel (.xlsx) para Shopstar se creó exitosamente.")

        except Exception as e:
            self.escribir_log(f"❌ ERROR CRÍTICO AL PROCESAR: {str(e)}", "error")


if __name__ == "__main__":
    ventana = tk.Tk()
    app = AplicacionConvertidor(ventana)
    ventana.mainloop()
