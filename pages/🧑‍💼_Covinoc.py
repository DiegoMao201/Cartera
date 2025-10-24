# ======================================================================================
# ARCHIVO: Pagina_Covinoc.py (v5 - Lógica de Estados y Pestañas Múltiples)
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
st.markdown(f"""
<style>
    .stApp {{ background-color: {PALETA_COLORES['fondo_claro']}; }}
    .stMetric {{ background-color: #FFFFFF; border-radius: 10px; padding: 15px; border: 1px solid #CCCCCC; }}
    .stTabs [data-baseweb="tab-list"] {{ gap: 24px; }}
    .stTabs [data-baseweb="tab"] {{ height: 50px; white-space: pre-wrap; background-color: transparent; border-radius: 4px 4px 0px 0px; border-bottom: 2px solid #C0C0C0; }}
    .stTabs [aria-selected="true"] {{ border-bottom: 2px solid {PALETA_COLORES['primario']}; color: {PALETA_COLORES['primario']}; font-weight: bold; }}
    div[data-baseweb="input"], div[data-baseweb="select"], div[data-baseweb="text-area"] {{ background-color: #FFFFFF; border: 1.5px solid {PALETA_COLORES['secundario']}; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); padding-left: 5px; }}
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
    Orquesta la carga y cruce con la lógica v5:
    1. Cruce inteligente de NIT/Documento y Factura/Titulo_Valor.
    2. Filtra por Estados ('Efectiva', 'Negada', 'Reclamada').
    3. Separa la lógica en 5 pestañas, incluyendo Ajustes Parciales.
    """
    
    # 1. Cargar Cartera Ferreinox
    df_cartera_raw = cargar_datos_cartera_dropbox()
    if df_cartera_raw.empty:
        st.error("No se pudo cargar 'cartera_detalle.csv'. El cruce no puede continuar.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    df_cartera = procesar_cartera(df_cartera_raw)
    
    if 'serie' in df_cartera.columns:
        df_cartera = df_cartera[~df_cartera['serie'].astype(str).str.contains('W|X', case=False, na=False)]

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
    estados_cerrados = ['EFECTIVA', 'NEGADA']
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
    
    # ===================== INICIO DE LA CORRECCIÓN =====================
    # El bloque 'rename' original (líneas 276-281) ha sido reemplazado.
    
    # --- CORRECCIÓN: Renombrar columnas no colisionantes ---
    # pd.merge(suffixes=...) solo añade sufijos a columnas que están en *ambos* DFs
    # (p.ej. 'factura_norm' se vuelve 'factura_norm_cartera' y 'factura_norm_covinoc').
    # Las columnas únicas (p.ej. 'importe' de cartera, 'saldo' de covinoc) no reciben sufijo.
    # El código más adelante espera que algunas de estas SÍ tengan sufijo,
    # mientras que otras (como 'dias_vencido') las espera sin sufijo.
    # Renombramos manualmente solo las columnas que el código espera con sufijo.

    columnas_a_renombrar = {
        # De df_cartera
        'importe': 'importe_cartera',         # Usado en Tab 5
        'nombrecliente': 'nombrecliente_cartera', # Usado en Tab 3 y 5
        'nit': 'nit_cartera',                 # Usado en Tab 3 y 5
        'nomvendedor': 'nomvendedor_cartera', # Usado en Tab 3

        # De df_covinoc
        'saldo': 'saldo_covinoc',             # Usado en Tab 3 y 5
        'estado': 'estado_covinoc',           # Usado en Tab 3 y 5
        'estado_norm': 'estado_norm_covinoc'  # Usado en Tab 5
    }

    # Renombramos solo las que existen en el DF fusionado
    cols_existentes = df_interseccion.columns
    renombres_aplicables = {k: v for k, v in columnas_a_renombrar.items() if k in cols_existentes}
    df_interseccion.rename(columns=renombres_aplicables, inplace=True)
    
    # Las columnas 'dias_vencido' y 'fecha_vencimiento' (de cartera)
    # se usan sin sufijo más adelante (Tab 3), lo cual ahora es correcto.
    
    # ====================== FIN DE LA CORRECCIÓN =======================


    # --- Tab 3: Aviso de No Pago ---
    # Facturas en intersección CON VENCIMIENTO ENTRE 55 y 58 DÍAS
    df_aviso_no_pago = df_interseccion[
        df_interseccion['dias_vencido'].between(55, 58)
    ].copy()

    # --- Tab 5: Ajustes por Abonos ---
    # 1. Convertir 'importe_cartera' y 'saldo_covinoc' a numérico para comparación
    # (Esta línea ahora funciona gracias a la corrección anterior)
    df_interseccion['importe_cartera'] = pd.to_numeric(df_interseccion['importe_cartera'], errors='coerce').fillna(0)
    df_interseccion['saldo_covinoc'] = pd.to_numeric(df_interseccion['saldo_covinoc'], errors='coerce').fillna(0)
    
    # 2. Facturas en intersección, 'Exonerada Parcial' Y saldos diferentes
    # (Esta línea ahora funciona gracias a la corrección anterior)
    df_ajustes = df_interseccion[
        (df_interseccion['estado_norm_covinoc'] == 'EXONERADA PARCIAL') & 
        (df_interseccion['importe_cartera'] != df_interseccion['saldo_covinoc'])
    ].copy()
    
    # 3. Calcular la diferencia
    # (Esta línea ahora funciona gracias a la corrección anterior)
    df_ajustes['diferencia'] = df_ajustes['importe_cartera'] - df_ajustes['saldo_covinoc']

    return df_a_subir, df_a_exonerar, df_aviso_no_pago, df_reclamadas, df_ajustes


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
        except Exception as e:
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
            st.markdown("Facturas de **clientes protegidos** que están en **Cartera Ferreinox** pero **NO** en Covinoc.")
            
            columnas_mostrar_subir = ['nombrecliente', 'nit', 'serie', 'numero', 'factura_norm', 'fecha_vencimiento', 'dias_vencido', 'importe', 'nomvendedor', 'clave_unica']
            columnas_existentes_subir = [col for col in columnas_mostrar_subir if col in df_a_subir.columns]
            
            st.dataframe(df_a_subir[columnas_existentes_subir], use_container_width=True, hide_index=True)
            st.download_button(
                label="📥 Descargar Excel para Subida (Próximamente)", data="", file_name="subir_covinoc.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", disabled=True 
            )

        with tab2:
            st.subheader("Facturas a Exonerar de Covinoc")
            st.markdown("Facturas en **Covinoc** (que no están 'Efectiva' o 'Negada') pero **NO** en la Cartera Ferreinox.")
            
            columnas_mostrar_exonerar = ['cliente', 'documento', 'titulo_valor', 'factura_norm', 'saldo', 'fecha', 'vencimiento', 'estado', 'clave_unica']
            columnas_existentes_exonerar = [col for col in columnas_mostrar_exonerar if col in df_a_exonerar.columns]
            
            st.dataframe(df_a_exonerar[columnas_existentes_exonerar], use_container_width=True, hide_index=True)
            st.download_button(
                label="📥 Descargar Excel para Exoneración (Próximamente)", data="", file_name="exonerar_covinoc.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", disabled=True 
            )

        with tab3:
            st.subheader("Facturas para Aviso de No Pago")
            st.markdown("Facturas que están **en ambos reportes** Y tienen un vencimiento **entre 55 y 58 días**.")
            
            columnas_mostrar_aviso = [
                'nombrecliente_cartera', 'nit_cartera', 'factura_norm_cartera', 'fecha_vencimiento', 'dias_vencido', 
                'importe_cartera', 'nomvendedor_cartera', 'saldo_covinoc', 'estado_covinoc', 'clave_unica'
            ]
            # Usamos df_interseccion como base para las columnas, ya que df_aviso_no_pago es un subconjunto
            columnas_existentes_aviso = [col for col in columnas_mostrar_aviso if col in df_interseccion.columns] 
            
            st.dataframe(df_aviso_no_pago[columnas_existentes_aviso], use_container_width=True, hide_index=True)
            st.download_button(
                label="📥 Descargar Excel para Aviso de No Pago (Próximamente)", data="", file_name="aviso_no_pago_covinoc.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", disabled=True
            )

        with tab4:
            st.subheader("Facturas en Reclamación (Informativo)")
            st.markdown("Facturas que figuran en Covinoc con estado **'Reclamada'**.")
            
            columnas_mostrar_reclamadas = ['cliente', 'documento', 'titulo_valor', 'factura_norm', 'saldo', 'fecha', 'vencimiento', 'estado', 'clave_unica']
            columnas_existentes_reclamadas = [col for col in columnas_mostrar_reclamadas if col in df_reclamadas.columns]
            
            st.dataframe(df_reclamadas[columnas_existentes_reclamadas], use_container_width=True, hide_index=True)

        with tab5:
            st.subheader("Ajustes por Abonos Parciales")
            st.markdown("Facturas en **ambos reportes** con estado **'Exonerada Parcial'** y **saldos diferentes**.")
            
            columnas_mostrar_ajustes = [
                'nombrecliente_cartera', 'nit_cartera', 'factura_norm_cartera', 'importe_cartera', 
                'saldo_covinoc', 'diferencia', 'dias_vencido', 'estado_covinoc', 'clave_unica'
            ]
            columnas_existentes_ajustes = [col for col in columnas_mostrar_ajustes if col in df_ajustes.columns]
            
            # Formatear la columna 'diferencia' para mejor visualización
            df_ajustes_display = df_ajustes[columnas_existentes_ajustes].copy()
            if 'diferencia' in df_ajustes_display.columns:
                df_ajustes_display['diferencia'] = df_ajustes_display['diferencia'].map('${:,.0f}'.format)
            
            st.dataframe(df_ajustes_display, use_container_width=True, hide_index=True)


if __name__ == '__main__':
    main()
