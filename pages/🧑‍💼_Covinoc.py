# ======================================================================================
# ARCHIVO: Pagina_Covinoc.py (v9 - Filtros Avanzados y Selección Total en Tab 1)
# MODIFICADO: Se añaden KPIs a todas las pestañas.
#             Se añade filtro de exclusión de clientes en Tab 1.
#             Se añade selección por checkbox (data_editor) en Tab 1.
#             Se optimiza Tab 3 para agrupar facturas por cliente en el mensaje
#             y se usa link 'wa.me' para abrir app de escritorio.
# 
# MODIFICACIÓN ACTUAL:
#           Se añade filtro de rango por "Días Vencido" en Tab 1.
#           Se añaden botones "Seleccionar Todos" / "Deseleccionar Todos" en Tab 1.
#           Se asegura que los KPIs y la tabla en Tab 1 respondan a TODOS los filtros.
# ======================================================================================
import streamlit as st
import pandas as pd
import toml
import os
from io import BytesIO, StringIO
import plotly.express as px
import plotly.graph_objects as go
import unicodedata
import re
from datetime import datetime
import dropbox
import glob
import urllib.parse # <-- IMPORTADO PARA WHATSAPP

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Gestión Covinoc",
    page_icon="🛡️",
    layout="wide"
)

# --- PALETA DE COLORES Y CSS (Copiada de Tablero_Principal.py para consistencia) ---
PALETA_COLORES = {
    "primario": "#003865",
    "secundario": "#0058A7",
    "acento": "#FFC300",
    "fondo_claro": "#F0F2F6",
    "texto_claro": "#FFFFFF",
    "texto_oscuro": "#31333F",
    "alerta_rojo": "#D32F2F",
    "alerta_naranja": "#F57C00",
    "alerta_amarillo": "#FBC02D",
    "exito_verde": "#388E3C"
}

# ================== INICIO DE LA MODIFICACIÓN (Datos Vendedores) ==================
# Diccionario de Vendedores y Teléfonos (normalizados)
# Las claves DEBEN coincidir con la salida de normalizar_nombre()
VENDEDORES_WHATSAPP = {
    "HUGO NELSON ZAPATA RAYO": "+573117658075",
    "TANIA RESTREPO BENJUMEA": "+573207425966",
    "DIEGO MAURICIO GARCIA RENGIFO": "+573205046277",
    "PABLO CESAR MAFLA BAÑOL": "+573103738523",
    "GUSTAVO ADOLFO PEREZ SANTA": "+573103663945",
    "ELISABETH CAROLINA IBARRA MANSO": "+573156224689",
    "CARLOS ALBERTO CASTRILLON LOPEZ": "+573147577658",
    "LEIVYN GRABIEL GARCIA MUNOZ": "+573127574279",
    "LEDUYN MELGAREJO ARIAS": "+573006620143",
    "JERSON ATEHORTUA OLARTE": "+573104952606"
}
# =================== FIN DE LA MODIFICACIÓN (Datos Vendedores) ===================

st.markdown(f"""
<style>
    .stApp {{ background-color: {PALETA_COLORES['fondo_claro']}; }}
    /* Modificación para métricas: añadir sombra y padding */
    .stMetric {{ 
        background-color: #FFFFFF; 
        border-radius: 10px; 
        padding: 20px; 
        border: 1px solid #CCCCCC;
        box-shadow: 0 4px 8px rgba(0,0,0,0.05);
    }}
    .stTabs [data-baseweb="tab-list"] {{ gap: 24px; }}
    .stTabs [data-baseweb="tab"] {{ height: 50px; white-space: pre-wrap; background-color: transparent; border-radius: 4px 4px 0px 0px; border-bottom: 2px solid #C0C0C0; }}
    .stTabs [aria-selected="true"] {{ border-bottom: 2px solid {PALETA_COLORES['primario']}; color: {PALETA_COLORES['primario']}; font-weight: bold; }}
    div[data-baseweb="input"], div[data-baseweb="select"], div[data-baseweb="text-area"] {{ background-color: #FFFFFF; border: 1.5px solid {PALETA_COLORES['secundario']}; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); padding-left: 5px; }}
    /* CSS para el st.data_editor */
    .stDataFrame {{ padding-top: 10px; }}
</style>
""", unsafe_allow_html=True)


# ======================================================================================
# --- LÓGICA DE CARGA DE DATOS (REUTILIZADA Y ADAPTADA) ---
# ======================================================================================

# --- Funciones Auxiliares Reutilizadas ---
def normalizar_nombre(nombre: str) -> str:
    """Normaliza nombres de columnas y datos para comparación."""
    if not isinstance(nombre, str): return ""
    nombre = nombre.upper().strip().replace('.', '')
    nombre = ''.join(c for c in unicodedata.normalize('NFD', nombre) if unicodedata.category(c) != 'Mn')
    return ' '.join(nombre.split())

ZONAS_SERIE = { "PEREIRA": [155, 189, 158, 439], "MANIZALES": [157, 238], "ARMENIA": [156] }

def procesar_cartera(df: pd.DataFrame) -> pd.DataFrame:
    """Procesa el dataframe de cartera principal (copiado de Tablero_Principal.py)."""
    df_proc = df.copy()
    if 'importe' not in df_proc.columns: df_proc['importe'] = 0
    if 'numero' not in df_proc.columns: df_proc['numero'] = '0'
    if 'dias_vencido' not in df_proc.columns: df_proc['dias_vencido'] = 0
    if 'nomvendedor' not in df_proc.columns: df_proc['nomvendedor'] = None
    if 'serie' not in df_proc.columns: df_proc['serie'] = ''

    df_proc['importe'] = pd.to_numeric(df_proc['importe'], errors='coerce').fillna(0)
    df_proc['numero'] = df_proc['numero'].astype(str) 
    df_proc['serie'] = df_proc['serie'].astype(str) 
    df_proc['dias_vencido'] = pd.to_numeric(df_proc['dias_vencido'], errors='coerce').fillna(0)
    df_proc['nomvendedor_norm'] = df_proc['nomvendedor'].apply(normalizar_nombre)
    ZONAS_SERIE_STR = {zona: [str(s) for s in series] for zona, series in ZONAS_SERIE.items()}
    
    def asignar_zona_robusta(valor_serie):
        if pd.isna(valor_serie): return "OTRAS ZONAS"
        numeros_en_celda = re.findall(r'\d+', str(valor_serie))
        if not numeros_en_celda: return "OTRAS ZONAS"
        for zona, series_clave_str in ZONAS_SERIE_STR.items():
            if set(numeros_en_celda) & set(series_clave_str): return zona
        return "OTRAS ZONAS"
    
    df_proc['zona'] = df_proc['serie'].apply(asignar_zona_robusta)
    bins = [-float('inf'), 0, 15, 30, 60, float('inf')]; labels = ['Al día', '1-15 días', '16-30 días', '31-60 días', 'Más de 60 días']
    df_proc['edad_cartera'] = pd.cut(df_proc['dias_vencido'], bins=bins, labels=labels, right=True)
    return df_proc

# --- Funciones de Carga de Dropbox ---

@st.cache_data(ttl=600)
def cargar_datos_cartera_dropbox():
    """Carga los datos de CARTERA más recientes desde el archivo CSV en Dropbox."""
    try:
        APP_KEY = st.secrets["dropbox"]["app_key"]
        APP_SECRET = st.secrets["dropbox"]["app_secret"]
        REFRESH_TOKEN = st.secrets["dropbox"]["refresh_token"]

        with dropbox.Dropbox(app_key=APP_KEY, app_secret=APP_SECRET, oauth2_refresh_token=REFRESH_TOKEN) as dbx:
            path_archivo_dropbox = '/data/cartera_detalle.csv'
            metadata, res = dbx.files_download(path=path_archivo_dropbox)
            contenido_csv = res.content.decode('latin-1')

            nombres_columnas_originales = [
                'Serie', 'Numero', 'Fecha Documento', 'Fecha Vencimiento', 'Cod Cliente',
                'NombreCliente', 'Nit', 'Poblacion', 'Provincia', 'Telefono1', 'Telefono2',
                'NomVendedor', 'Entidad Autoriza', 'E-Mail', 'Importe', 'Descuento',
                'Cupo Aprobado', 'Dias Vencido'
            ]

            df = pd.read_csv(
                StringIO(contenido_csv), 
                header=None, 
                names=nombres_columnas_originales, 
                sep='|', 
                engine='python',
                dtype={'Serie': str, 'Numero': str, 'Nit': str}
            )
            
            df_renamed = df.rename(columns=lambda x: normalizar_nombre(x).lower().replace(' ', '_'))
            df_renamed = df_renamed.loc[:, ~df_renamed.columns.duplicated()]
            df_renamed['fecha_documento'] = pd.to_datetime(df_renamed['fecha_documento'], errors='coerce')
            df_renamed['fecha_vencimiento'] = pd.to_datetime(df_renamed['fecha_vencimiento'], errors='coerce')
            
            return df_renamed
    except Exception as e:
        st.error(f"Error al cargar 'cartera_detalle.csv' desde Dropbox: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=600)
def cargar_reporte_transacciones_dropbox():
    """Carga el REPORTE TRANSACCIONES (Covinoc) desde un archivo Excel en Dropbox."""
    try:
        APP_KEY = st.secrets["dropbox"]["app_key"]
        APP_SECRET = st.secrets["dropbox"]["app_secret"]
        REFRESH_TOKEN = st.secrets["dropbox"]["refresh_token"]

        with dropbox.Dropbox(app_key=APP_KEY, app_secret=APP_SECRET, oauth2_refresh_token=REFRESH_TOKEN) as dbx:
            path_archivo_dropbox = '/data/reporteTransacciones.xlsx'
            
            metadata, res = dbx.files_download(path=path_archivo_dropbox)
            
            df = pd.read_excel(
                BytesIO(res.content),
                dtype={'DOCUMENTO': str, 'TITULO_VALOR': str, 'ESTADO': str} # Forzamos columnas clave a string
            )
            
            df.columns = [normalizar_nombre(c).lower().replace(' ', '_') for c in df.columns]

            return df
    except Exception as e:
        st.error(f"Error al cargar 'reporteTransacciones.xlsx' desde Dropbox: {e}")
        st.warning("Asegúrate de que el archivo 'reporteTransacciones.xlsx' exista en la carpeta '/data/' de Dropbox.")
        return pd.DataFrame()

# --- Funciones de Normalización de Claves ---

def normalizar_nit_simple(nit_str: str) -> str:
    """Limpia un NIT, quitando todo lo que no sea un número."""
    if not isinstance(nit_str, str):
        return ""
    return re.sub(r'\D', '', nit_str)

def normalizar_factura_simple(fact_str: str) -> str:
    """Limpia un número de factura (para Covinoc) quitando espacios, puntos, guiones."""
    if not isinstance(fact_str, str):
        return ""
    return fact_str.split('.')[0].strip().upper().replace(' ', '').replace('-', '')

def normalizar_factura_cartera(row):
    """Concatena Serie y Numero para Cartera, limpiándolos."""
    serie = str(row['serie']).strip().upper()
    numero = str(row['numero']).split('.')[0].strip()
    return (serie + numero).replace(' ', '').replace('-', '')


# --- Función Principal de Procesamiento y Cruce ---

@st.cache_data
def cargar_y_comparar_datos():
    """
    Orquesta la carga y cruce con la lógica v6:
    1. Cruce inteligente de NIT/Documento y Factura/Titulo_Valor.
    2. Filtra series 'W', 'X' y las terminadas en 'U'.
    3. Lógica de Aviso No Pago >= 25 días.
    4. Lógica de Ajustes Parciales (Covinoc > Cartera).
    """
    
    # 1. Cargar Cartera Ferreinox
    df_cartera_raw = cargar_datos_cartera_dropbox()
    if df_cartera_raw.empty:
        st.error("No se pudo cargar 'cartera_detalle.csv'. El cruce no puede continuar.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    df_cartera = procesar_cartera(df_cartera_raw)
    
    # ===================== INICIO DE LA MODIFICACIÓN (Filtro Series) =====================
    # Filtrar series W, X (en cualquier parte) y series que terminan en U (Anticipos, etc.)
    if 'serie' in df_cartera.columns:
        df_cartera = df_cartera[~df_cartera['serie'].astype(str).str.contains('W|X', case=False, na=False)]
        df_cartera = df_cartera[~df_cartera['serie'].astype(str).str.upper().str.endswith('U', na=False)]
    # ====================== FIN DE LA MODIFICACIÓN (Filtro Series) =======================

    # 2. Cargar Reporte Transacciones (Covinoc)
    df_covinoc = cargar_reporte_transacciones_dropbox()
    if df_covinoc.empty:
        st.error("No se pudo cargar 'reporteTransacciones.xlsx'. El cruce no puede continuar.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # 3. Preparar Claves de Cruce (Lógica Avanzada)

    # 3.1. Normalizar NIT de Cartera y crear un Set para búsqueda rápida
    df_cartera['nit_norm_cartera'] = df_cartera['nit'].apply(normalizar_nit_simple)
    set_nits_cartera = set(df_cartera['nit_norm_cartera'].unique())

    # 3.2. Función de Normalización Inteligente para Covinoc
    def encontrar_nit_en_cartera(doc_str_covinoc):
        if not isinstance(doc_str_covinoc, str): return None
        doc_norm = normalizar_nit_simple(doc_str_covinoc)
        if doc_norm in set_nits_cartera:
            return doc_norm
        doc_norm_base = doc_norm[:-1] 
        if doc_norm_base in set_nits_cartera:
            return doc_norm_base 
        return None 

    # 3.3. Aplicar la normalización inteligente a Covinoc
    df_covinoc['nit_norm_cartera'] = df_covinoc['documento'].apply(encontrar_nit_en_cartera)

    # 3.4. Normalizar Facturas en ambos DFs
    df_cartera['factura_norm'] = df_cartera.apply(normalizar_factura_cartera, axis=1)
    df_covinoc['factura_norm'] = df_covinoc['titulo_valor'].apply(normalizar_factura_simple)

    # 3.5. Crear Clave Única
    df_cartera['clave_unica'] = df_cartera['nit_norm_cartera'] + '_' + df_cartera['factura_norm']
    df_covinoc['clave_unica'] = df_covinoc['nit_norm_cartera'] + '_' + df_covinoc['factura_norm']
    
    # 3.6. Normalizar columna 'estado' de Covinoc para filtros
    df_covinoc['estado_norm'] = df_covinoc['estado'].astype(str).str.upper().str.strip()
    
    # 4. Lógica de Cruces y Pestañas
    
    # --- Tab 4: Reclamadas (Informativo) ---
    df_reclamadas = df_covinoc[df_covinoc['estado_norm'] == 'RECLAMADA'].copy()
    
    # --- Tab 1: Facturas a Subir ---
    # 1. Obtener lista de clientes protegidos (todos los NITs que coincidieron en Covinoc)
    nits_protegidos = df_covinoc['nit_norm_cartera'].dropna().unique()
    # 2. Filtrar cartera a solo esos clientes
    df_cartera_protegida = df_cartera[df_cartera['nit_norm_cartera'].isin(nits_protegidos)].copy()
    # 3. Obtener *todas* las claves únicas que ya existen en Covinoc
    set_claves_covinoc_total = set(df_covinoc['clave_unica'].dropna().unique())
    # 4. Las facturas a subir son las de clientes protegidos que NO están en Covinoc
    df_a_subir = df_cartera_protegida[
        ~df_cartera_protegida['clave_unica'].isin(set_claves_covinoc_total)
    ].copy()

    # --- Tab 2: Exoneraciones ---
    # 1. Filtrar Covinoc a solo facturas "comparables" (excluir cerradas)
    
    # ================== INICIO DE LA MODIFICACIÓN (Excluir 'EXONERADA') ==================
    # Se añade 'EXONERADA' a la lista para que no aparezcan en la pestaña 2.
    estados_cerrados = ['EFECTIVA', 'NEGADA', 'EXONERADA']
    # =================== FIN DE LA MODIFICACIÓN (Excluir 'EXONERADA') ===================
    
    df_covinoc_comparable = df_covinoc[~df_covinoc['estado_norm'].isin(estados_cerrados)].copy()
    # 2. Obtener *todas* las claves únicas que existen en Cartera
    set_claves_cartera_total = set(df_cartera['clave_unica'].dropna().unique())
    # 3. Las facturas a exonerar son las "comparables" de Covinoc que NO están en Cartera
    df_a_exonerar = df_covinoc_comparable[
        (~df_covinoc_comparable['clave_unica'].isin(set_claves_cartera_total)) &
        (df_covinoc_comparable['nit_norm_cartera'].notna()) # Solo las que tienen un NIT coincidente
    ].copy()

    # --- Intersección para Tabs 3 y 5 ---
    df_interseccion = pd.merge(
        df_cartera,
        df_covinoc, 
        on='clave_unica',
        how='inner', 
        suffixes=('_cartera', '_covinoc') 
    )
    
    # ===================== INICIO DE LA CORRECCIÓN (KeyError) =====================
    # Renombramos manually las columnas que no colisionaron pero que el 
    # código posterior espera que tengan sufijos.
    
    columnas_a_renombrar = {
        # De df_cartera
        'importe': 'importe_cartera',
        'nombrecliente': 'nombrecliente_cartera',
        'nit': 'nit_cartera',
        'nomvendedor': 'nomvendedor_cartera',
        'fecha_vencimiento': 'fecha_vencimiento_cartera',
        'dias_vencido': 'dias_vencido_cartera',

        # De df_covinoc
        'saldo': 'saldo_covinoc',
        'estado': 'estado_covinoc',
        'estado_norm': 'estado_norm_covinoc',
        'vencimiento': 'vencimiento_covinoc'
    }

    # Renombramos solo las que existen en el DF fusionado
    cols_existentes = df_interseccion.columns
    renombres_aplicables = {k: v for k, v in columnas_a_renombrar.items() if k in cols_existentes}
    df_interseccion.rename(columns=renombres_aplicables, inplace=True)
    
    # ====================== FIN DE LA CORRECCIÓN (KeyError) =======================


    # --- Tab 3: Aviso de No Pago ---
    # ===================== INICIO DE LA MODIFICACIÓN (Lógica Aviso No Pago) =====================
    # Facturas en intersección CON VENCIMIENTO MAYOR O IGUAL A 25 DÍAS
    df_aviso_no_pago = df_interseccion[
        df_interseccion['dias_vencido_cartera'] >= 25
    ].copy()
    # ====================== FIN DE LA MODIFICACIÓN (Lógica Aviso No Pago) =======================

    # --- Tab 5: Ajustes por Abonos ---
    # 1. Convertir 'importe_cartera' y 'saldo_covinoc' a numérico para comparación
    df_interseccion['importe_cartera'] = pd.to_numeric(df_interseccion['importe_cartera'], errors='coerce').fillna(0)
    df_interseccion['saldo_covinoc'] = pd.to_numeric(df_interseccion['saldo_covinoc'], errors='coerce').fillna(0)
    
    # ===================== INICIO DE LA MODIFICACIÓN (Lógica Ajustes) =====================
    # 2. Facturas en intersección donde el Saldo en Covinoc es MAYOR al Importe en Cartera
    #    (Significa que Ferreinox recibió un abono que Covinoc no tiene)
    df_ajustes = df_interseccion[
        (df_interseccion['saldo_covinoc'] > df_interseccion['importe_cartera'])
    ].copy()
    
    # 3. Calcular la diferencia (El monto a "exonerar" parcialmente en Covinoc)
    df_ajustes['diferencia'] = df_ajustes['saldo_covinoc'] - df_ajustes['importe_cartera']
    # ====================== FIN DE LA MODIFICACIÓN (Lógica Ajustes) =======================

    return df_a_subir, df_a_exonerar, df_aviso_no_pago, df_reclamadas, df_ajustes


# ======================================================================================
# --- FUNCIONES AUXILIARES PARA EXCEL ---
# ======================================================================================

# ================== INICIO DE LA MODIFICACIÓN (Lógica Tipo Documento) ==================
def get_tipo_doc_from_nit_col(nit_str_raw: str) -> str:
    """
    Determina si un documento es NIT ('N') o Cédula ('C') [MODIFICADO].
    - Es 'N' si contiene guión ('-') o si los números empiezan por 8 o 9.
    - En CUALQUIER otro caso, se asume 'C'.
    """
    if not isinstance(nit_str_raw, str) or pd.isna(nit_str_raw):
        return 'C' # Default a Cédula de Ciudadanía si es nulo o no string
    
    nit_str_raw_clean = nit_str_raw.strip().upper()
    
    # --- Regla 1: Prioridad NIT (N) ---
    # Si contiene guión, es NIT
    if '-' in nit_str_raw_clean:
        return 'N'
    
    # Limpiamos para análisis numérico
    nit_norm = re.sub(r'\D', '', nit_str_raw_clean)
    length = len(nit_norm)
    
    if length == 0:
        return 'C' # Default si está vacío después de limpiar
        
    # Si empieza con 8xx, 9xx (prefijos comunes de NIT)
    if (nit_norm.startswith('8') or nit_norm.startswith('9')):
        return 'N'
        
    # --- Regla 2: Todo lo demás es Cédula (C) ---
    # Ya que no fue 'N' por guión ni por prefijo 8/9,
    # cualquier otra cosa (longitud 7, 8, 10, 11, con letras, etc.)
    # se forzará a 'C' según la solicitud.
    return 'C'
# =================== FIN DE LA MODIFICACIÓN (Lógica Tipo Documento) ===================

# ================== INICIO DE LA MODIFICACIÓN (Formato Fecha YYYY/MM/DD) ==================
def format_date(date_obj) -> str:
    """Formatea un objeto de fecha a 'YYYY/mm/dd' o devuelve None."""
    if pd.isna(date_obj):
        return None
    try:
        # Cambiado de '%d/%m/%Y' a '%Y/%m/%d'
        return pd.to_datetime(date_obj).strftime('%Y/%m/%d')
    except Exception:
        return None
# =================== FIN DE LA MODIFICACIÓN (Formato Fecha YYYY/MM/DD) ===================

def to_excel(df: pd.DataFrame) -> bytes:
    """Convierte un DataFrame a un archivo Excel en memoria (bytes)."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Facturas')
    processed_data = output.getvalue()
    return processed_data


# ======================================================================================
# --- BLOQUE PRINCIPAL DE LA APP ---
# ======================================================================================
def main():
    # --- Lógica de Autenticación (Copiada 1:1 de Tablero_Principal.py) ---
    if 'authentication_status' not in st.session_state:
        st.session_state['authentication_status'] = False
        st.session_state['acceso_general'] = False
        st.session_state['vendedor_autenticado'] = None

    if not st.session_state['authentication_status']:
        st.title("Acceso al Módulo de Cartera Protegida")
        try:
            general_password = st.secrets["general"]["password"]
            vendedores_secrets = st.secrets["vendedores"]
  M       except Exception as e:
            st.error(f"Error al cargar las contraseñas desde los secretos: {e}")
            st.stop()
        
        password = st.text_input("Introduce la contraseña:", type="password", key="password_input_covinoc")
        
        if st.button("Ingresar"):
            if password == str(general_password):
                st.session_state['authentication_status'] = True
                st.session_state['acceso_general'] = True
                st.session_state['vendedor_autenticado'] = "General"
                st.rerun()
            else:
                for vendedor_key, pass_vendedor in vendedores_secrets.items():
                    if password == str(pass_vendedor):
                        st.session_state['authentication_status'] = True
                        st.session_state['acceso_general'] = False
                        st.session_state['vendedor_autenticado'] = vendedor_key
                        st.rerun()
                        break
                if not st.session_state['authentication_status']:
                    st.error("Contraseña incorrecta.")
    else:
        # --- Aplicación Principal (Usuario Autenticado) ---
        st.title("🛡️ Gestión de Cartera Protegida (Covinoc)")

        if st.button("🔄 Recargar Datos (Dropbox)"):
            st.cache_data.clear()
            st.success("Caché limpiado. Recargando datos de Cartera y Covinoc...")
            st.rerun()

        # --- Barra Lateral (Sidebar) ---
        with st.sidebar:
            try:
                st.image("LOGO FERREINOX SAS BIC 2024.png", use_container_width=True)
            except Exception:
                st.warning("Logo no encontrado.")
            
            st.success(f"Usuario: {st.session_state['vendedor_autenticado']}")
            
            if st.button("Cerrar Sesión"):
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()
            
            st.markdown("---")
            st.info("Esta página compara la cartera de Ferreinox con el reporte de transacciones de Covinoc.")
            
            # ================== INICIO DE LA CORRECCIÓN DEL ERROR ==================
            # La siguiente línea causaba el error 'MediaFileStorageError' porque
            # el archivo 'image_5019c6.png' no se encontraba.
            # Lo he comentado. Si tienes el archivo en la misma carpeta que este
            # script, puedes quitar el '#' para mostrar la imagen.
            
            # st.image(
            #      "image_5019c6.png", 
            #      caption="Instructivo Carga Masiva (Referencia)"
            # )
            # =================== FIN DE LA CORRECCIÓN DEL ERROR ===================

        # --- Carga y Procesamiento de Datos ---
        with st.spinner("Cargando y comparando archivos de Dropbox..."):
            df_a_subir, df_a_exonerar, df_aviso_no_pago, df_reclamadas, df_ajustes = cargar_y_comparar_datos()

        if df_a_subir.empty and df_a_exonerar.empty and df_aviso_no_pago.empty and df_reclamadas.empty and df_ajustes.empty:
            try:
                with dropbox.Dropbox(app_key=st.secrets["dropbox"]["app_key"], app_secret=st.secrets["dropbox"]["app_secret"], oauth2_refresh_token=st.secrets["dropbox"]["refresh_token"]) as dbx:
                    dbx.files_get_metadata('/data/cartera_detalle.csv')
                    dbx.files_get_metadata('/data/reporteTransacciones.xlsx')
                st.warning("Se cargaron los archivos, pero no se encontraron diferencias para las 5 categorías.")
            except Exception as e:
                st.error(f"No se pudieron cargar los archivos base. Verifica la conexión o los nombres de archivo en Dropbox: {e}")
                st.stop()


        # --- Contenedor Principal con Pestañas ---
        st.markdown("---")
        
        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            f"1. Facturas a Subir ({len(df_a_subir)})", 
            f"2. Exoneraciones ({len(df_a_exonerar)})", 
            f"3. Avisos de No Pago ({len(df_aviso_no_pago)})",
            f"4. Reclamadas ({len(df_reclamadas)})",
            f"5. Ajustes Parciales ({len(df_ajustes)})"
        ])

        with tab1:
            st.subheader("Facturas a Subir a Covinoc")
            st.markdown("Facturas de **clientes protegidos** que están en **Cartera Ferreinox** pero **NO** en Covinoc. (Excluye series W, X y terminadas en U).")
            
            if df_a_subir.empty:
                st.info("No hay facturas pendientes por subir.")
            else:
                # ================== INICIO MODIFICACIÓN: Filtros Adicionales ==================
                st.markdown("---")
                st.subheader("Filtros Adicionales")
                
                # --- Filtro 1: Excluir Clientes ---
                clientes_unicos = sorted(df_a_subir['nombrecliente'].dropna().unique())
                clientes_excluidos = st.multiselect(
                    "Excluir Clientes del Listado:",
                    options=clientes_unicos,
                    default=[],
                    help="Seleccione uno o más clientes para ocultar sus facturas de la lista de selección."
                )
                
                # --- Filtro 2: Días Vencido (NUEVO) ---
                if not df_a_subir['dias_vencido'].empty:
                    min_dias = int(df_a_subir['dias_vencido'].min())
                    max_dias = int(df_a_subir['dias_vencido'].max())
                    # Aseguramos que el slider tenga un rango válido
                    if min_dias == max_dias:
                        max_dias += 1
                else:
                    min_dias, max_dias = 0, 1
                    
                dias_range = st.slider(
                    "Filtrar por Días Vencido:", 
                    min_value=min_dias, 
                    max_value=max_dias, 
                    value=(min_dias, max_dias),
                    help="Seleccione el rango de días de vencimiento a incluir."
                )
                
                # Aplicar AMBOS filtros
                df_a_subir_filtrado = df_a_subir[
                    (~df_a_subir['nombrecliente'].isin(clientes_excluidos)) &
                    (df_a_subir['dias_vencido'] >= dias_range[0]) &
                    (df_a_subir['dias_vencido'] <= dias_range[1])
                ].copy()
                # =================== FIN MODIFICACIÓN: Filtros Adicionales ====================

                # ================== INICIO MODIFICACIÓN: Indicadores (Reactivos a Filtros) ==================
                st.markdown("---")
                st.subheader("Indicadores de Gestión (Facturas Filtradas)")
                
                kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
                try:
                    # Los KPIs AHORA usan df_a_subir_filtrado
                    monto_total_filtrado = pd.to_numeric(df_a_subir_filtrado['importe'], errors='coerce').sum()
                    clientes_unicos_filtrados = df_a_subir_filtrado['nombrecliente'].nunique()
                except Exception:
                    monto_total_filtrado = 0
                    clientes_unicos_filtrados = 0

                kpi_col1.metric(
                    label="Nº Facturas (Filtradas)",
                    value=f"{len(df_a_subir_filtrado)}"
                )
                kpi_col2.metric(
                    label="Monto Total (Filtrado)",
                    value=f"${monto_total_filtrado:,.0f}"
                )
                kpi_col3.metric(
                    label="Nº Clientes (Filtrados)",
                    value=f"{clientes_unicos_filtrados}"
                )
                st.markdown("---")
                # =================== FIN MODIFICACIÓN: Indicadores (Reactivos a Filtros) ===================
            
                # ================== INICIO MODIFICACIÓN: Selección de Facturas (con "Seleccionar Todos") ==================
                st.subheader("Selección de Facturas para Descarga")
                st.info("Utilice las casillas de la columna 'Seleccionar' para elegir qué facturas desea incluir en el archivo Excel.")

                # --- Lógica de Botones "Seleccionar Todos" (NUEVO) ---
                # Usamos session_state para manejar el estado de selección y la clave del editor
                if 'data_editor_key_tab1' not in st.session_state:
                    st.session_state.data_editor_key_tab1 = "data_editor_subir_0"
                if 'default_select_val_tab1' not in st.session_state:
                    st.session_state.default_select_val_tab1 = False

                sel_col1, sel_col2 = st.columns(2)
                with sel_col1:
                    if st.button("✅ Seleccionar Todos (Visible en Filtro)", use_container_width=True):
                        st.session_state.default_select_val_tab1 = True
                        # Cambiamos la clave para forzar la recreación del data_editor
                        st.session_state.data_editor_key_tab1 = f"data_editor_subir_{int(st.session_state.data_editor_key_tab1.split('_')[-1]) + 1}"
                        st.rerun()
                with sel_col2:
                    if st.button("◻️ Deseleccionar Todos (Visible en Filtro)", use_container_width=True):
                        st.session_state.default_select_val_tab1 = False
                        st.session_state.data_editor_key_tab1 = f"data_editor_subir_{int(st.session_state.data_editor_key_tab1.split('_')[-1]) + 1}"
                        st.rerun()
                # --- Fin Lógica Botones ---

                columnas_mostrar_subir = ['nombrecliente', 'nit', 'serie', 'numero', 'factura_norm', 'fecha_vencimiento', 'dias_vencido', 'importe', 'nomvendedor', 'clave_unica']
                columnas_existentes_subir = [col for col in columnas_mostrar_subir if col in df_a_subir_filtrado.columns]
                
                # Preparamos el DF para el editor, AHORA basado en df_a_subir_filtrado
                df_para_seleccionar = df_a_subir_filtrado[columnas_existentes_subir].copy()
                # Insertamos la columna 'Seleccionar' con el valor por defecto de session_state
                df_para_seleccionar.insert(0, "Seleccionar", st.session_state.default_select_val_tab1) 
                
                # Columnas que no deben ser editables (todas excepto 'Seleccionar')
                columnas_deshabilitadas = [col for col in df_para_seleccionar.columns if col != 'Seleccionar']

                # Usamos st.data_editor para la selección
                df_editado = st.data_editor(
                    df_para_seleccionar,
                    use_container_width=True,
                    hide_index=True,
                    # Configuración de la columna de selección
                    column_config={
                        "Seleccionar": st.column_config.CheckboxColumn(
                            "Seleccionar",
                            default=st.session_state.default_select_val_tab1,
                        ),
                        "importe": st.column_config.NumberColumn(
                            "Importe",
                            format="$ %d"
                        ),
                        "dias_vencido": st.column_config.NumberColumn(
                            "Días Vencido",
                            format="%d días"
                        )
                    },
                    # Deshabilitamos la edición de las columnas de datos
                    disabled=columnas_deshabilitadas, 
                    # Usamos la clave dinámica de session_state
                    key=st.session_state.data_editor_key_tab1 
                )
                
                # Filtramos las filas que fueron seleccionadas del editor
                df_seleccionado = df_editado[df_editado["Seleccionar"] == True].copy()
                
                st.markdown(f"**Facturas seleccionadas para descarga: {len(df_seleccionado)}**")
                # =================== FIN MODIFICACIÓN: Selección de Facturas ===================

                # --- Lógica de Descarga Excel (Tab 1) - MODIFICADA ---
                # Ahora usa df_seleccionado en lugar de df_a_subir
                if not df_seleccionado.empty:
                    df_subir_excel = pd.DataFrame()
                    df_subir_excel['TIPO_DOCUMENTO'] = df_seleccionado['nit'].apply(get_tipo_doc_from_nit_col)
                    df_subir_excel['DOCUMENTO'] = df_seleccionado['nit']
                    df_subir_excel['TITULO_VALOR'] = df_seleccionado['factura_norm']
                    df_subir_excel['VALOR'] = pd.to_numeric(df_seleccionado['importe'], errors='coerce').fillna(0).astype(int)
                    df_subir_excel['FECHA'] = pd.to_datetime(df_seleccionado['fecha_vencimiento'], errors='coerce').apply(format_date)
                    df_subir_excel['CODIGO_CONSULTA'] = 986638
                    excel_data_subir = to_excel(df_subir_excel)
                else:
                    excel_data_subir = b""

                st.download_button(
                    label="📥 Descargar Excel para Subida (SÓLO SELECCIONADAS)", 
                    data=excel_data_subir, 
                    file_name="1_facturas_a_subir_SELECCIONADAS.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                    # Se deshabilita si no hay nada seleccionado
                    disabled=df_seleccionado.empty 
                )

        with tab2:
            st.subheader("Facturas a Exonerar de Covinoc")
            st.markdown("Facturas en **Covinoc** (que no están 'Efectiva', 'Negada' o 'Exonerada') pero **NO** en la Cartera Ferreinox.")
            
            # ================== INICIO MODIFICACIÓN: Indicadores (Goal 1) ==================
            st.markdown("---")
            st.subheader("Indicadores de Gestión")
            
            kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
            try:
                monto_total_exonerar = pd.to_numeric(df_a_exonerar['saldo'], errors='coerce').sum()
                clientes_unicos_exonerar = df_a_exonerar['cliente'].nunique()
            except Exception:
                monto_total_exonerar = 0
                clientes_unicos_exonerar = 0

            kpi_col1.metric(
                label="Nº Facturas a Exonerar",
                value=f"{len(df_a_exonerar)}"
            )
            kpi_col2.metric(
                label="Monto Total a Exonerar",
                value=f"${monto_total_exonerar:,.0f}"
            )
            kpi_col3.metric(
                label="Nº Clientes Afectados",
                value=f"{clientes_unicos_exonerar}"
            )
            st.markdown("---")
            # =================== FIN MODIFICACIÓN: Indicadores (Goal 1) ===================

            columnas_mostrar_exonerar = ['cliente', 'documento', 'titulo_valor', 'factura_norm', 'saldo', 'fecha', 'vencimiento', 'estado', 'clave_unica']
            columnas_existentes_exonerar = [col for col in columnas_mostrar_exonerar if col in df_a_exonerar.columns]
            
            st.dataframe(df_a_exonerar[columnas_existentes_exonerar], use_container_width=True, hide_index=True)

            # --- Lógica de Descarga Excel (Tab 2) ---
            if not df_a_exonerar.empty:
                df_exonerar_excel = pd.DataFrame()
                df_exonerar_excel['TIPO_DOCUMENTO'] = df_a_exonerar['documento'].apply(get_tipo_doc_from_nit_col)
                # ================== INICIO DE LA MODIFICACIÓN SOLICITADA ==================
                # Se usa el 'documento' original de Covinoc
                df_exonerar_excel['DOCUMENTO'] = df_a_exonerar['documento']
                # =================== FIN DE LA MODIFICACIÓN SOLICITADA ===================
                df_exonerar_excel['TITULO_VALOR'] = df_a_exonerar['factura_norm']
                df_exonerar_excel['VALOR'] = pd.to_numeric(df_a_exonerar['saldo'], errors='coerce').fillna(0).astype(int)
                df_exonerar_excel['FECHA'] = pd.to_datetime(df_a_exonerar['vencimiento'], errors='coerce').apply(format_date)
                excel_data_exonerar = to_excel(df_exonerar_excel)
            else:
                excel_data_exonerar = b""

            st.download_button(
                label="📥 Descargar Excel para Exoneración (Formato Covinoc)", 
                data=excel_data_exonerar, 
                file_name="2_exoneraciones_totales.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                disabled=df_a_exonerar.empty 
            )

        with tab3:
            st.subheader("Facturas para Aviso de No Pago")
            st.markdown("Facturas que están **en ambos reportes** Y tienen un vencimiento **>= 25 días**.")
            
            # ================== INICIO MODIFICACIÓN: Indicadores (Goal 1) ==================
            st.markdown("---")
            st.subheader("Indicadores de Gestión")
            
            kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
            try:
                monto_total_aviso = pd.to_numeric(df_aviso_no_pago['importe_cartera'], errors='coerce').sum()
                clientes_unicos_aviso = df_aviso_no_pago['nombrecliente_cartera'].nunique()
            except Exception:
                monto_total_aviso = 0
                clientes_unicos_aviso = 0

            kpi_col1.metric(
                label="Nº Facturas en Aviso",
                value=f"{len(df_aviso_no_pago)}"
            )
            kpi_col2.metric(
                label="Monto Total en Aviso",
                value=f"${monto_total_aviso:,.0f}"
            )
            kpi_col3.metric(
                label="Nº Clientes en Aviso",
                value=f"{clientes_unicos_aviso}"
            )
            st.markdown("---")
            # =================== FIN MODIFICACIÓN: Indicadores (Goal 1) ===================

            columnas_mostrar_aviso = [
                'nombrecliente_cartera', 'nit_cartera', 'factura_norm_cartera', 'fecha_vencimiento_cartera', 'dias_vencido_cartera', 
                'importe_cartera', 'nomvendedor_cartera', 'saldo_covinoc', 'estado_covinoc', 'clave_unica'
            ]
            
            columnas_existentes_aviso = [col for col in columnas_mostrar_aviso if col in df_aviso_no_pago.columns]
            
            # Dataframe original
            st.dataframe(df_aviso_no_pago[columnas_existentes_aviso], use_container_width=True, hide_index=True)

            # --- Lógica de Descarga Excel (Tab 3) ---
            if not df_aviso_no_pago.empty:
                df_aviso_excel = pd.DataFrame()
                # ================== INICIO DE LA MODIFICACIÓN SOLICITADA ==================
                # Se usa el 'documento' original de Covinoc para TIPO y DOCUMENTO
                df_aviso_excel['TIPO_DOCUMENTO'] = df_aviso_no_pago['documento'].apply(get_tipo_doc_from_nit_col)
AN             df_aviso_excel['DOCUMENTO'] = df_aviso_no_pago['documento']
                # =================== FIN DE LA MODIFICACIÓN SOLICITADA ===================
                df_aviso_excel['TITULO_VALOR'] = df_aviso_no_pago['factura_norm_cartera']
                df_aviso_excel['VALOR'] = pd.to_numeric(df_aviso_no_pago['importe_cartera'], errors='coerce').fillna(0).astype(int)
                df_aviso_excel['FECHA'] = pd.to_datetime(df_aviso_no_pago['fecha_vencimiento_cartera'], errors='coerce').apply(format_date)
                excel_data_aviso = to_excel(df_aviso_excel)
            else:
                excel_data_aviso = b""

            # Botón de descarga original
            st.download_button(
                label="📥 Descargar Excel para Aviso de No Pago (Formato Covinoc)", 
                data=excel_data_aviso, 
                file_name="3_aviso_no_pago.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                disabled=df_aviso_no_pago.empty
            )
            
            # ================== INICIO DE LA MODIFICACIÓN (Gestión WhatsApp v2 - Agrupada) ==================
            st.markdown("---")
            st.subheader("🚀 Gestión de Avisos por Vendedor (WhatsApp)")
            
            if df_aviso_no_pago.empty:
                st.info("No hay facturas en 'Aviso de No Pago' para gestionar.")
            else:
                st.info("Seleccione los vendedores para preparar los mensajes de gestión.")
                
                vendedores_unicos = sorted(df_aviso_no_pago['nomvendedor_cartera'].dropna().unique())
                vendedores_seleccionados = st.multiselect(
                    "Vendedores a gestionar:", 
                    options=vendedores_unicos, 
                    default=[]
                )

                if not vendedores_seleccionados:
                    st.write("Seleccione uno o más vendedores para continuar.")
                else:
                    df_aviso_filtrado = df_aviso_no_pago[
                        df_aviso_no_pago['nomvendedor_cartera'].isin(vendedores_seleccionados)
                    ].copy()
                    
                    grouped = df_aviso_filtrado.groupby('nomvendedor_cartera')
                    
                    for vendor_name, group_df in grouped:
                        st.markdown(f"---")
                        st.markdown(f"#### Vendedor: **{vendor_name}** ({len(group_df)} facturas)")
                        
                        # Buscar teléfono
                        vendor_name_norm = normalizar_nombre(vendor_name)
                        phone_encontrado = VENDEDORES_WHATSAPP.get(vendor_name_norm, "")
                        
                        col1, col2 = st.columns([0.4, 0.6])
                        
                        with col1:
                            # El teléfono se carga aquí y es editable por el usuario
                            phone_manual = st.text_input(
                                "Teléfono (Ej: +57311...):", 
                                value=phone_encontrado, 
                                key=f"phone_{vendor_name_norm}"
                            )
                        
                        # Construir el mensaje
                        try:
                            nombre_corto = vendor_name.split(' ')[0].capitalize()
                        except Exception:
                            nombre_corto = vendor_name

                        # Mensaje de encabezado actualizado
                        mensaje_header = f"¡Hola {nombre_corto}! 👋\n\nPor favor, te pido gestionar la siguiente cartera que está próxima a **Aviso de No Pago en Covinoc** (>= 25 días vencidos):\n"
                        
                        # Agrupar facturas por cliente
                        mensaje_clientes_facturas = []
                        grouped_by_client = group_df.groupby('nombrecliente_cartera')
                        
                        for client_name, client_df in grouped_by_client:
                            cliente_str = str(client_name).strip()
                            mensaje_clientes_facturas.append(f"\n• *Cliente:* {cliente_str}")
                            
                            # Iterar sobre las facturas de ESE cliente
                            for _, row in client_df.iterrows():
                                factura = str(row['factura_norm_cartera']).strip()
                                try:
                                    valor = float(row['importe_cartera'])
                                    valor_str = f"${valor:,.0f}"
                                except (ValueError, TypeError):
                                    valor_str = str(row['importe_cartera'])
                                dias = row['dias_vencido_cartera']
                                
                                # Añadir detalles de la factura
                                mensaje_clientes_facturas.append(f"    - *Factura:* {factura} | *Valor:* {valor_str} | *Días Vencidos:* {dias}")

                        # Unir todo el mensaje
                        mensaje_completo = mensaje_header + "\n".join(mensaje_clientes_facturas) + "\n\nQuedo atento a cualquier novedad. ¡Gracias!"
                        
                        # Limpiar teléfono y codificar mensaje
                        phone_limpio = phone_manual.replace(' ', '').replace('+', '').strip()
                        if phone_limpio and not phone_limpio.startswith("57"):
                                phone_limpio = f"57{phone_limpio}" # Asegurar código de país

                        mensaje_url_encoded = urllib.parse.quote_plus(mensaje_completo)
                        
                        # URL actualizada para usar wa.me (permite app de escritorio)
AN                     url_whatsapp = f"https://wa.me/{phone_limpio}?text={mensaje_url_encoded}"
                        
                        with col2:
                            st.write(" ") # Spacer para alinear el botón verticalmente
                            st.link_button(
                                "📲 Enviar a WhatsApp (Web/App)", # Texto de botón actualizado
                                url_whatsapp, 
                                use_container_width=True, 
                                disabled=(not phone_manual)
                            )
                        
                        with st.expander("Ver detalle de facturas y mensaje completo"):
                            st.dataframe(group_df[columnas_existentes_aviso], use_container_width=True, hide_index=True)
                            st.text_area(
                                "Mensaje a Enviar:",AN                           value=mensaje_completo, 
                                height=300, # Altura aumentada
                                key=f"msg_{vendor_name_norm}",
                                disabled=True
                            )
            # =================== FIN DE LA MODIFICACIÓN (Gestión WhatsApp v2 - Agrupada) ===================

        with tab4:
            st.subheader("Facturas en Reclamación (Informativo)")
            st.markdown("Facturas que figuran en Covinoc con estado **'Reclamada'**.")

            # ================== INICIO MODIFICACIÓN: Indicadores (Goal 1) ==================
            st.markdown("---")
            st.subheader("Indicadores de Gestión")
            
            kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
            try:
                monto_total_reclamadas = pd.to_numeric(df_reclamadas['saldo'], errors='coerce').sum()
                clientes_unicos_reclamadas = df_reclamadas['cliente'].nunique()
            except Exception:
                monto_total_reclamadas = 0
                clientes_unicos_reclamadas = 0

            kpi_col1.metric(
                label="Nº Facturas Reclamadas",
                value=f"{len(df_reclamadas)}"
AN         )
            kpi_col2.metric(
                label="Monto Total Reclamado",
                value=f"${monto_total_reclamadas:,.0f}"
            )
            kpi_col3.metric(
                label="Nº Clientes",
                value=f"{clientes_unicos_reclamadas}"
            )
            st.markdown("---")
            # =================== FIN MODIFICACIÓN: Indicadores (Goal 1) ===================
            
            columnas_mostrar_reclamadas = ['cliente', 'documento', 'titulo_valor', 'factura_norm', 'saldo', 'fecha', 'vencimiento', 'estado', 'clave_unica']
            columnas_existentes_reclamadas = [col for col in columnas_mostrar_reclamadas if col in df_reclamadas.columns]
            
            st.dataframe(df_reclamadas[columnas_existentes_reclamadas], use_container_width=True, hide_index=True)

        with tab5:
            st.subheader("Ajustes por Abonos Parciales")
            st.markdown("Facturas en **ambos reportes** donde el **Saldo Covinoc es MAYOR** al **Importe Cartera** (implica un abono no reportado).")
            
            # ================== INICIO MODIFICACIÓN: Indicadores (Goal 1) ==================
            st.markdown("---")
            st.subheader("Indicadores de Gestión")
            
            kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
            try:
                # 'diferencia' ya se calcula como numérica en la carga de datos
                monto_total_ajuste = pd.to_numeric(df_ajustes['diferencia'], errors='coerce').sum()
                clientes_unicos_ajuste = df_ajustes['nombrecliente_cartera'].nunique()
            except Exception:
                monto_total_ajuste = 0
                clientes_unicos_ajuste = 0

            kpi_col1.metric(
                label="Nº Facturas para Ajuste",
                value=f"{len(df_ajustes)}"
            )
            kpi_col2.metric(
                label="Monto Total a Ajustar",
                value=f"${monto_total_ajuste:,.0f}"
            )
            kpi_col3.metric(
                label="Nº Clientes Afectados",
                value=f"{clientes_unicos_ajuste}"
            )
            st.markdown("---")
AN         # =================== FIN MODIFICACIÓN: Indicadores (Goal 1) ===================

            columnas_mostrar_ajustes = [
                'nombrecliente_cartera', 'nit_cartera', 'factura_norm_cartera', 'importe_cartera', 
                'saldo_covinoc', 'diferencia', 'dias_vencido_cartera', 'estado_covinoc', 'clave_unica'
            ]
            columnas_existentes_ajustes = [col for col in columnas_mostrar_ajustes if col in df_ajustes.columns]
            
            # Formatear columnas para mejor visualización
            df_ajustes_display = df_ajustes[columnas_existentes_ajustes].copy()
            for col_moneda in ['importe_cartera', 'saldo_covinoc', 'diferencia']:
                if col_moneda in df_ajustes_display.columns:
                    df_ajustes_display[col_moneda] = df_ajustes_display[col_moneda].map('${:,.0f}'.format)
AN             
            st.dataframe(df_ajustes_display, use_container_width=True, hide_index=True)
            
            # --- Lógica de Descarga Excel (Tab 5) ---
            if not df_ajustes.empty:
                df_ajustes_excel = pd.DataFrame()
                # ================== INICIO DE LA MODIFICACIÓN SOLICITADA ==================
                # Se usa el 'documento' original de Covinoc para TIPO y DOCUMENTO
                df_ajustes_excel['TIPO_DOCUMENTO'] = df_ajustes['documento'].apply(get_tipo_doc_from_nit_col)
                df_ajustes_excel['DOCUMENTO'] = df_ajustes['documento']
                # =================== FIN DE LA MODIFICACIÓN SOLICITLADA ===================
an             df_ajustes_excel['TITULO_VALOR'] = df_ajustes['factura_norm_cartera']
                # El VALOR a exonerar es la DIFERENCIA
    dias_range         df_ajustes_excel['VALOR'] = pd.to_numeric(df_ajustes['diferencia'], errors='coerce').fillna(0).astype(int)
                df_ajustes_excel['FECHA'] = pd.to_datetime(df_ajustes['fecha_vencimiento_cartera'], errors='coerce').apply(format_date)
                excel_data_ajustes = to_excel(df_ajustes_excel)
            else:
                excel_data_ajustes = b""

            st.download_button(
                label="📥 Descargar Excel de Ajuste (Exoneración Parcial)", 
                data=excel_data_ajustes, 
                file_name="5_ajustes_exoneracion_parcial.xlsx",
AN               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                disabled=df_ajustes.empty
            )


if __name__ == '__main__':
    main()
