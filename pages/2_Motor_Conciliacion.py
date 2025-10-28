# ======================================================================================
# ARCHIVO: pages/2_Motor_Conciliacion.py
# (Versión CORREGIDA Y MEJORADA - 27 de Octubre, 2025)
#
# MODIFICACIÓN (Gemini):
# 1.  Se actualiza `cargar_planilla_bancos` para filtrar solo los pagos
#     pendientes de identificar, basándose en la nueva lógica de negocio:
#     - VALOR > 0 (Es un ingreso)
#     - RECIBO está VACÍO (No es un pago de cartera/crédito)
#     - DESTINO está VACÍO (No se ha identificado)
#
# 2.  Se actualiza `cargar_ventas_diarias` para usar la estructura de 18 columnas
#     del script 'Resumen Mensual.py' (ventas_detalle.csv).
#     - Se asume que una "Venta de Contado" se identifica por
#       TipoDocumento == 'CONTADO'.
#     - Esto corrige el 'KeyError: fecha' al usar 'fecha_venta' y
#       'valor_venta' como columnas fuente.
#
# 3.  Se mantiene `cargar_y_procesar_cartera` sin cambios, ya que procesa
#     el archivo de cartera de crédito (cartera_detalle.csv), que es
#     distinto al archivo de ventas (ventas_detalle.csv).
# ======================================================================================
import streamlit as st
import pandas as pd
import dropbox
from io import StringIO, BytesIO
import re
import unicodedata
from datetime import datetime
from fuzzywuzzy import fuzz
import gspread
from gspread_dataframe import set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials
import logging

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Motor de Conciliación",
    page_icon="🤖",
    layout="wide"
)

# --- PALETA DE COLORES Y CSS (Copiado de tu app principal para consistencia) ---
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
</style>
""", unsafe_allow_html=True)

# ======================================================================================
# --- 1. FUNCIONES DE CONEXIÓN A SERVICIOS (DROPBOX Y GOOGLE SHEETS) ---
# ======================================================================================

@st.cache_resource(ttl=3600)
def get_dbx_client(secrets_key):
    """
    Crea un cliente de Dropbox usando una clave de secrets.toml específica.
    Esto nos permite conectarnos a 'dropbox' o 'dropbox_ventas'.
    """
    try:
        if secrets_key not in st.secrets:
            st.error(f"Error: No se encontró la configuración '[{secrets_key}]' en secrets.toml.")
            st.stop()
            
        creds = st.secrets[secrets_key]
        return dropbox.Dropbox(
            app_key=creds["app_key"],
            app_secret=creds["app_secret"],
            oauth2_refresh_token=creds["refresh_token"]
        )
    except Exception as e:
        st.error(f"Error conectando a Dropbox (App: {secrets_key}): {e}")
        st.stop()

@st.cache_resource(ttl=3600)
def connect_to_google_sheets():
    """Conecta con Google Sheets usando las credenciales del Service Account."""
    try:
        scope = [
            'https://www.googleapis.com/auth/spreadsheets',
            'https://www.googleapis.com/auth/drive'
        ]
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        st.error(f"Error conectando a Google Sheets: {e}")
        st.info("Asegúrate de haber añadido la sección [gcp_service_account] a tu secrets.toml.")
        st.stop()

def get_gsheet_worksheet(g_client, sheet_url, worksheet_name):
    """Accede a una pestaña específica de un Google Sheet por URL."""
    try:
        sheet = g_client.open_by_url(sheet_url)
        return sheet.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"Error: No se encontró la pestaña '{worksheet_name}' en tu Google Sheet.")
        st.info(f"Asegúrate de que la pestaña exista y que el nombre en secrets.toml ('{worksheet_name}') sea correcto.")
        st.stop()
    except Exception as e:
        st.error(f"Error abriendo Google Sheet: {e}")
        st.info("Asegúrate de haber compartido tu Google Sheet con el 'client_email' del robot.")
        st.stop()

def download_file_from_dropbox(dbx_client, file_path):
    """Descarga el contenido de un archivo desde Dropbox."""
    try:
        # Aseguramos que el path sea absoluto (inicie con /)
        if not file_path.startswith('/'):
            file_path = '/' + file_path

        metadata, res = dbx_client.files_download(path=file_path)
        return res.content
    except dropbox.exceptions.ApiError as e:
        st.error(f"Error en API de Dropbox al descargar {file_path}: {e}")
        st.info(f"Verifica que la ruta '{file_path}' sea correcta.")
        return None
    except Exception as e:
        st.error(f"Error inesperado al descargar {file_path}: {e}")
        return None

# ======================================================================================
# --- 2. FUNCIONES DE CARGA Y TRANSFORMACIÓN DE DATOS (ETL) ---
# ======================================================================================

def normalizar_texto(texto: str) -> str:
    """Limpia y estandariza texto para facilitar los cruces."""
    if not isinstance(texto, str):
        return ""
    texto = texto.upper().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'[^A-Z0-9\s-]', '', texto) # Permitimos guiones (para Nits y facturas)
    palabras_irrelevantes = ['PAGO', 'TRANSF', 'TRANSFERENCIA', 'CONSIGNACION', 'FACTURA', 'REF', 'BALOTO', 'EFECTY', 'PSE', 'NRO', 'CTA']
    for palabra in palabras_irrelevantes:
        texto = texto.replace(palabra, '')
    return ' '.join(texto.split()) # Normalizar espacios

@st.cache_data(ttl=600)
def cargar_y_procesar_cartera():
    """
    Carga y procesa la cartera (crédito) desde Dropbox (App 'dropbox').
    Esta función usa el path hardcodeado '/data/cartera_detalle.csv'.
    """
    dbx_client = get_dbx_client("dropbox")
    path_archivo_dropbox = '/data/cartera_detalle.csv'
    
    content = download_file_from_dropbox(dbx_client, path_archivo_dropbox)
    
    if content:
        df = pd.read_csv(StringIO(content.decode('latin-1')), sep='|', header=None, names=[
            'Serie', 'Numero', 'Fecha Documento', 'Fecha Vencimiento', 'Cod Cliente',
            'NombreCliente', 'Nit', 'Poblacion', 'Provincia', 'Telefono1', 'Telefono2',
            'NomVendedor', 'Entidad Autoriza', 'E-Mail', 'Importe', 'Descuento',
            'Cupo Aprobado', 'Dias Vencido'
        ])
        
        # Renombramos columnas a un formato limpio
        df.columns = [col.upper().strip().replace(' ', '_') for col in df.columns]
        
        df['IMPORTE'] = pd.to_numeric(df['IMPORTE'], errors='coerce').fillna(0)
        df['NUMERO'] = pd.to_numeric(df['NUMERO'], errors='coerce').fillna(0)
        df.loc[df['NUMERO'] < 0, 'IMPORTE'] *= -1
        df['DIAS_VENCIDO'] = pd.to_numeric(df['DIAS_VENCIDO'], errors='coerce').fillna(0)
        
        # Columnas clave para el motor de conciliación
        df['id_factura_unica'] = df['SERIE'].astype(str) + '-' + df['NUMERO'].astype(str)
        df['nit_norm'] = df['NIT'].astype(str).str.replace(r'[^0-9]', '', regex=True) # Limpia solo a números
        df['nombre_norm'] = df['NOMBRECLIENTE'].apply(normalizar_texto)
        
        # Filtramos solo lo que está pendiente de pago
        df_pendiente = df[df['IMPORTE'] > 0].copy()
        return df_pendiente
        
    return pd.DataFrame()

@st.cache_data(ttl=600)
def cargar_planilla_bancos(path_planilla_bancos):
    """
    Carga y limpia la planilla de bancos desde Dropbox (App 'dropbox').
    NUEVA LÓGICA: Filtra solo los ingresos (valor > 0) donde 'RECIBO' y
    'DESTINO' están vacíos, ya que esos son los pendientes de identificar.
    """
    dbx_client = get_dbx_client("dropbox")
    content = download_file_from_dropbox(dbx_client, path_planilla_bancos)
    
    if content:
        try:
            df = pd.read_excel(BytesIO(content))
        except Exception as e:
            st.warning(f"No se pudo leer como Excel ({path_planilla_bancos}), intentando como CSV... ({e})")
            try:
                df = pd.read_csv(BytesIO(content), sep=';') 
            except Exception as e2:
                st.error(f"No se pudo leer el archivo de bancos: {e2}")
                return pd.DataFrame()
        
        columnas_esperadas = ['FECHA', 'SUCURSAL BANCO', 'TIPO DE TRANSACCION', 'CUENTA', 
                              'EMPRESA', 'VALOR', 'BANCO REFRENCIA INTERNA', 'DESTINO', 
                              'RECIBO', 'FECHA RECIBO']
        
        if len(df.columns) > len(columnas_esperadas):
            st.warning(f"Advertencia en 'planilla_bancos': Se encontraron {len(df.columns)} columnas, pero se esperaban {len(columnas_esperadas)}. Se usarán solo las primeras {len(columnas_esperadas)}.")
            df = df.iloc[:, :len(columnas_esperadas)]
        elif len(df.columns) < len(columnas_esperadas):
                 st.error(f"Error en 'planilla_bancos': Se esperaban {len(columnas_esperadas)} columnas pero se encontraron {len(df.columns)}.")
                 st.info("El archivo parece estar incompleto o corrupto.")
                 return pd.DataFrame()
            
        df.columns = columnas_esperadas
        
        df_limpio = df.copy()
        df_limpio['fecha'] = pd.to_datetime(df_limpio['FECHA'], errors='coerce')
        df_limpio['valor'] = pd.to_numeric(df_limpio['VALOR'], errors='coerce').fillna(0)
        
        # --- INICIO DE LA NUEVA LÓGICA DE FILTRADO ---
        
        # 1. Filtramos solo los que son INGRESOS
        df_ingresos = df_limpio[df_limpio['valor'] > 0].copy()
        
        if df_ingresos.empty:
            st.info("El archivo de bancos no contiene movimientos de ingreso (valor > 0).")
            return pd.DataFrame()

        # 2. Normalizamos columnas de lógica para el filtro
        # Usamos .fillna('') para que los nulos se traten como strings vacíos
        df_ingresos['RECIBO_norm'] = df_ingresos['RECIBO'].fillna('').astype(str).str.strip()
        df_ingresos['DESTINO_norm'] = df_ingresos['DESTINO'].fillna('').astype(str).str.strip()

        # 3. APLICAMOS LÓGICA DE NEGOCIO:
        #    Queremos solo los pagos que NO son de cartera (RECIBO está vacío)
        #    Y que NO están identificados (DESTINO está vacío)
        df_pendientes = df_ingresos[
            (df_ingresos['RECIBO_norm'] == '') &
            (df_ingresos['DESTINO_norm'] == '')
        ].copy()
        
        # 4. Si no hay pendientes, retornamos un DF vacío
        if df_pendientes.empty:
            st.info("Se leyeron los movimientos bancarios, pero no se encontraron nuevos pagos pendientes de identificar (Recibo y Destino vacíos).")
            # Devolvemos un DF con la misma estructura para evitar errores
            columnas_finales = list(df_ingresos.columns) + ['descripcion_banco', 'texto_match', 'id_banco_unico']
            return pd.DataFrame(columns=columnas_finales)
        
        st.success(f"Se encontraron {len(df_pendientes)} nuevos pagos pendientes de identificar.")

        # 5. Continuamos el procesamiento SOLO con los pendientes
        df_pendientes['descripcion_banco'] = (
            df_pendientes['TIPO DE TRANSACCION'].fillna('').astype(str) + ' ' +
            df_pendientes['BANCO REFRENCIA INTERNA'].fillna('').astype(str) + ' ' +
            df_pendientes['DESTINO'].fillna('').astype(str) # Aunque esté vacío, lo mantenemos
        )
        df_pendientes['texto_match'] = df_pendientes['descripcion_banco'].apply(normalizar_texto)
        df_pendientes = df_pendientes.reset_index(drop=True)
        
        df_pendientes['id_banco_unico'] = df_pendientes.apply(
            lambda row: f"B-{row['fecha'].strftime('%Y%m%d')}-{int(row['valor'])}-{row.name}", axis=1
        )
        
        # --- FIN DE LA NUEVA LÓGICA ---
        
        return df_pendientes

@st.cache_data(ttl=600)
def cargar_ventas_diarias(path_ventas_diarias):
    """
    Carga las ventas diarias ('ventas_detalle.csv') desde Dropbox (App 'dropbox_ventas').
    Usa la estructura de 18 columnas de 'Resumen Mensual.py'.
    Asume que 'CONTADO' es un 'TipoDocumento'.
    """
    dbx_client = get_dbx_client("dropbox_ventas") 
    content = download_file_from_dropbox(dbx_client, path_ventas_diarias)
    
    if not content:
        st.error(f"No se pudo descargar el archivo de ventas '{path_ventas_diarias}' desde la App 'dropbox_ventas'.")
        return pd.DataFrame()

    # Columnas correctas (de Resumen Mensual.py)
    nombres_columnas = [
        'anio', 'mes', 'fecha_venta', 'Serie', 'TipoDocumento', 
        'codigo_vendedor', 'nomvendedor', 'cliente_id', 'nombre_cliente', 
        'codigo_articulo', 'nombre_articulo', 'categoria_producto', 
        'linea_producto', 'marca_producto', 'valor_venta', 
        'unidades_vendidas', 'costo_unitario', 'super_categoria'
    ]

    try:
        # Leemos el CSV usando la nueva lista de 18 columnas
        df = pd.read_csv(BytesIO(content), sep=';', encoding='latin-1', header=None, names=nombres_columnas)
    except Exception as e:
        st.error(f"No se pudo leer el archivo de ventas ({path_ventas_diarias}): {e}")
        return pd.DataFrame()

    try:
        # Limpieza básica
        df['fecha_venta'] = pd.to_datetime(df['fecha_venta'], errors='coerce')
        df['valor_venta'] = pd.to_numeric(df['valor_venta'], errors='coerce').fillna(0)
        df['TipoDocumento'] = df['TipoDocumento'].fillna('').astype(str).apply(normalizar_texto)
        
        # --- NUEVA LÓGICA DE FILTRO ---
        # Asumimos que 'CONTADO' es un TipoDocumento
        VALOR_FORMA_PAGO_CONTADO = 'CONTADO'
        df_contado = df[df['TipoDocumento'] == VALOR_FORMA_PAGO_CONTADO].copy()

        if df_contado.empty:
            st.warning(f"No se encontraron ventas de contado (TipoDocumento = '{VALOR_FORMA_PAGO_CONTADO}') en el archivo de ventas.")
            # Devolvemos un DF vacío con la estructura esperada
            cols_esperadas = ['fecha', 'valor_contado', 'forma_pago', 'cliente']
            return pd.DataFrame(columns=cols_esperadas)

        # Renombramos para que coincida con el motor de conciliación
        df_std = df_contado.rename(columns={
            'fecha_venta': 'fecha',
            'valor_venta': 'valor_contado',
            'TipoDocumento': 'forma_pago',
            'nombre_cliente': 'cliente'
        })
        
        # Seleccionamos solo las columnas que usa el Nivel 3
        return df_std[['fecha', 'valor_contado', 'forma_pago', 'cliente']]

    except KeyError as e:
        st.error(f"Error en 'cargar_ventas_diarias': No se encontró la columna {e}.")
        st.info(f"Esto sugiere que el archivo en '{path_ventas_diarias}' no coincide con la estructura de 18 columnas de 'ventas_detalle.csv'.")
        return pd.DataFrame()

# ======================================================================================
# --- 3. MOTOR DE CONCILIACIÓN AUTOMÁTICA ---
# ======================================================================================

def run_auto_reconciliation(df_bancos, df_cartera, df_ventas):
    """
    Ejecuta el motor de conciliación en cascada.
    """
    st.write("Iniciando motor de conciliación automática...")
    
    conciliados = []
    ids_banco_conciliados = set()
    ids_factura_conciliadas = set()
    
    df_bancos_pendientes = df_bancos.copy()
    
    # --- NIVEL 1: MATCH PERFECTO (ID de Factura en Descripción) ---
    mapa_facturas = {row['id_factura_unica']: row for _, row in df_cartera.iterrows()}
    
    for idx, pago in df_bancos_pendientes.iterrows():
        # Busca patrones como "123-456"
        matches = re.findall(r'(\d+-\d+)', pago['texto_match'])
        for id_factura_potencial in matches:
            if id_factura_potencial in mapa_facturas:
                factura = mapa_facturas[id_factura_potencial]
                
                if abs(pago['valor'] - factura['IMPORTE']) < 1000: # Tolerancia de $1000
                    pago['status'] = 'Conciliado (Auto N1 - ID Factura)'
                    pago['id_factura_asignada'] = id_factura_potencial
                    pago['cliente_asignado'] = factura['NOMBRECLIENTE']
                    conciliados.append(pago)
                    
                    ids_banco_conciliados.add(pago['id_banco_unico'])
                    ids_factura_conciliadas.add(factura['id_factura_unica'])
                    break 
        if pago['id_banco_unico'] in ids_banco_conciliados:
            continue 
            
    # --- NIVEL 2: MATCH POR NIT + VALOR EXACTO ---
    cartera_restante = df_cartera[~df_cartera['id_factura_unica'].isin(ids_factura_conciliadas)]
    mapa_nits = cartera_restante.groupby('nit_norm')['IMPORTE'].apply(list).to_dict()

    for _, pago in df_bancos_pendientes[~df_bancos_pendientes['id_banco_unico'].isin(ids_banco_conciliados)].iterrows():
        # Busca NITS (números de 8 a 10 dígitos)
        nits_potenciales = re.findall(r'(\d{8,10})', pago['texto_match'])
        for nit in nits_potenciales:
            if nit in mapa_nits:
                facturas_del_nit = mapa_nits[nit]
                for valor_factura in facturas_del_nit:
                    if abs(pago['valor'] - valor_factura) < 1000:
                        factura_match = cartera_restante[
                            (cartera_restante['nit_norm'] == nit) & 
                            (cartera_restante['IMPORTE'] == valor_factura)
                        ].iloc[0]
                        
                        pago['status'] = 'Conciliado (Auto N2 - NIT+Valor)'
                        pago['id_factura_asignada'] = factura_match['id_factura_unica']
                        pago['cliente_asignado'] = factura_match['NOMBRECLIENTE']
                        conciliados.append(pago)
                        
                        ids_banco_conciliados.add(pago['id_banco_unico'])
                        ids_factura_conciliadas.add(factura_match['id_factura_unica'])
                        mapa_nits[nit].remove(valor_factura) 
                        break 
            if pago['id_banco_unico'] in ids_banco_conciliados:
                break 

    # --- NIVEL 3: MATCH VENTAS DE CONTADO ---
    if not df_ventas.empty:
        df_ventas['fecha_str'] = df_ventas['fecha'].dt.strftime('%Y-%m-%d')
        mapa_ventas = df_ventas.groupby('fecha_str')['valor_contado'].apply(list).to_dict()
        
        for _, pago in df_bancos_pendientes[~df_bancos_pendientes['id_banco_unico'].isin(ids_banco_conciliados)].iterrows():
            fecha_pago_str = pago['fecha'].strftime('%Y-%m-%d')
            if fecha_pago_str in mapa_ventas:
                for valor_venta in mapa_ventas[fecha_pago_str]:
                    if abs(pago['valor'] - valor_venta) < 100: # Tolerancia muy baja para contado
                        pago['status'] = 'Conciliado (Auto N3 - Venta Contado)'
                        pago['id_factura_asignada'] = f"CONTADO-{fecha_pago_str}"
                        pago['cliente_asignado'] = "Venta Contado" 
                        conciliados.append(pago)
                        
                        ids_banco_conciliados.add(pago['id_banco_unico'])
                        mapa_ventas[fecha_pago_str].remove(valor_venta) 
                        break
    
    # --- Finalizar ---
    df_conciliados = pd.DataFrame(conciliados)
    df_no_conciliados = df_bancos_pendientes[~df_bancos_pendientes['id_banco_unico'].isin(ids_banco_conciliados)].copy()
    df_no_conciliados['status'] = 'Pendiente (Revisión Manual)'
    
    st.success(f"Motor finalizado: {len(df_conciliados)} pagos conciliados automáticamente.")
    return df_conciliados, df_no_conciliados

# ======================================================================================
# --- 4. APLICACIÓN PRINCIPAL DE STREAMLIT ---
# ======================================================================================

def main_app():
    
    st.title("🤖 Motor de Conciliación Bancaria")
    st.markdown("Carga, procesa y concilia los extractos bancarios contra la cartera y las ventas de contado.")

    # --- Validar Autenticación ---
    if not st.session_state.get('authentication_status', False):
        st.warning("Por favor, inicia sesión desde la página principal para acceder a esta herramienta.")
        st.stop()
    
    # --- Cargar configuración desde secrets.toml ---
    try:
        # Configuración de Google Sheets
        G_SHEET_URL = st.secrets["google_sheets"]["sheet_url"]
        G_SHEET_TAB_CONCILIADOS = st.secrets["google_sheets"]["tab_conciliados"]
        
        # Path de Bancos (de la App 'dropbox')
        PATH_PLANILLA_BANCOS = st.secrets["dropbox"]["path_bancos"]
        # Path de Ventas (de la App 'dropbox_ventas')
        PATH_VENTAS_DIARIAS = st.secrets["dropbox_ventas"]["path_ventas"]
        
    except KeyError as e:
        st.error(f"Error: Falta una clave en tu archivo secrets.toml: {e}")
        st.info("Revisa la estructura de ejemplo y asegúrate de que todas las claves [dropbox], [dropbox_ventas] y [google_sheets] existan.")
        st.stop()

    # --- Inicializar session_state para guardar los datos ---
    if 'data_loaded' not in st.session_state:
        st.session_state.data_loaded = False
        st.session_state.df_bancos = pd.DataFrame()
        st.session_state.df_cartera = pd.DataFrame()
        st.session_state.df_ventas = pd.DataFrame()
        st.session_state.df_conciliados_auto = pd.DataFrame()
        st.session_state.df_pendientes = pd.DataFrame()

    # --- BOTÓN DE CARGA Y PROCESAMIENTO ---
    st.header("Ejecutar Proceso ETL y Conciliación")
    
    if st.button("🔄 Cargar y Conciliar Datos", type="primary"):
        with st.spinner("Conectando y cargando datos..."):
            
            st.session_state.df_cartera = cargar_y_procesar_cartera()
            # Carga solo los bancos PENDIENTES (Recibo y Destino vacíos)
            st.session_state.df_bancos = cargar_planilla_bancos(PATH_PLANILLA_BANCOS) 
            # Carga solo las ventas de CONTADO (TipoDocumento == 'CONTADO')
            st.session_state.df_ventas = cargar_ventas_diarias(PATH_VENTAS_DIARIAS) 
            st.session_state.data_loaded = True
            
            if st.session_state.df_cartera.empty:
                st.error("Error: La cartera no pudo ser cargada. Revisa el path '/data/cartera_detalle.csv' y el archivo.")
                st.session_state.data_loaded = False
            
            # Ya no fallamos si bancos o ventas están vacíos, 
            # porque la función de carga puede devolver vacío si no hay pendientes.
            elif st.session_state.df_bancos.empty:
                st.warning("No se encontraron NUEVOS movimientos bancarios pendientes de identificar.")
                # Continuamos, puede que solo queramos ver el histórico
            
            else:
                st.success(f"Datos cargados: {len(st.session_state.df_cartera)} facturas (cartera), {len(st.session_state.df_bancos)} mov. bancarios (pendientes), {len(st.session_state.df_ventas)} ventas (contado).")
        
        if st.session_state.data_loaded: 
            with st.spinner("Ejecutando motor de conciliación automática..."):
                g_client = connect_to_google_sheets()
                ws_conciliados = get_gsheet_worksheet(g_client, G_SHEET_URL, G_SHEET_TAB_CONCILIADOS)
                
                try:
                    df_historico_gsheet = pd.DataFrame(ws_conciliados.get_all_records())
                except Exception as e:
                    st.warning(f"No se pudo leer historial de Google Sheets (puede estar vacía): {e}")
                    df_historico_gsheet = pd.DataFrame()

                ids_ya_conciliados = set()
                if not df_historico_gsheet.empty and 'id_banco_unico' in df_historico_gsheet.columns:
                    ids_ya_conciliados = set(df_historico_gsheet['id_banco_unico'])
                    st.write(f"Se encontraron {len(ids_ya_conciliados)} registros ya conciliados en Google Sheets.")
                
                bancos_a_procesar = st.session_state.df_bancos[
                    ~st.session_state.df_bancos['id_banco_unico'].isin(ids_ya_conciliados)
                ]
                
                if bancos_a_procesar.empty:
                    st.info("No se encontraron nuevos movimientos bancarios para procesar (ya estaban en G-Sheets o no había pendientes).")
                    st.session_state.df_pendientes = pd.DataFrame()
                    st.session_state.df_conciliados_auto = pd.DataFrame()
                else:
                    df_auto, df_pend = run_auto_reconciliation(
                        bancos_a_procesar, 
                        st.session_state.df_cartera, 
                        st.session_state.df_ventas
                    )
                    
                    st.session_state.df_conciliados_auto = df_auto
                    st.session_state.df_pendientes = df_pend
                    
                    if not df_auto.empty:
                        with st.spinner("Guardando conciliados automáticos en Google Sheets..."):
                            df_auto_save = df_auto.copy()
                            for col in df_auto_save.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]']):
                                df_auto_save[col] = df_auto_save[col].astype(str)
                            
                            existing_data = ws_conciliados.get_all_values()
                            list_of_lists_to_add = df_auto_save.values.tolist()
                            
                            if len(existing_data) > 0:
                                ws_conciliados.append_rows(list_of_lists_to_add, value_input_option='USER_ENTERED')
                            else: 
                                set_with_dataframe(ws_conciliados, df_auto_save)
                                
                            st.success(f"{len(df_auto)} pagos automáticos guardados en Google Sheets.")

    # --- RESULTADOS DE LA CONCILIACIÓN ---
    if st.session_state.data_loaded:
        st.header("Resultados de la Conciliación")
        
        # ==================================================================
        # --- INICIO DE LA CORRECCIÓN ---
        # Si el DataFrame está vacío (porque no se procesó nada),
        # no podemos acceder a ['valor'].sum() o dará KeyError.
        # Verificamos si está vacío primero.
        # ==================================================================
        
        total_auto = st.session_state.df_conciliados_auto['valor'].sum() if not st.session_state.df_conciliados_auto.empty else 0
        total_pendiente = st.session_state.df_pendientes['valor'].sum() if not st.session_state.df_pendientes.empty else 0
        total_recibido_nuevos = total_auto + total_pendiente

        # ==================================================================
        # --- FIN DE LA CORRECCIÓN ---
        # ==================================================================

        kpi_cols = st.columns(3)
        kpi_cols[0].metric("🏦 Nuevos Pagos (Pendientes de ID)", f"${total_recibido_nuevos:,.0f}")
        kpi_cols[1].metric("✅ Conciliado (Automático)", f"${total_auto:,.0f}")
        kpi_cols[2].metric("❓ Pendiente (Manual)", f"${total_pendiente:,.0f}", delta=f"{len(st.session_state.df_pendientes)} transacciones")

        tab_manual, tab_auto, tab_fuentes = st.tabs(
            ["📝 **PENDIENTE DE ASIGNACIÓN MANUAL**", "🤖 Conciliados Automáticamente", "🗂️ Datos Fuente Cargados"]
        )

        with tab_manual:
            if st.session_state.df_pendientes.empty:
                st.success("¡Excelente! No hay pagos pendientes de revisión manual.")
            else:
                st.info(f"Se encontraron {len(st.session_state.df_pendientes)} pagos que requieren tu atención.")
                
                # Preparamos las opciones para el selectbox
                clientes_cartera = st.session_state.df_cartera.drop_duplicates(subset=['nit_norm']) \
                                                             .set_index('nit_norm')['NOMBRECLIENTE'].to_dict()
                opciones_clientes = {nit: f"{nombre} (NIT: {nit})" for nit, nombre in clientes_cartera.items()}
                # Añadimos opciones para "Venta Contado" y "Otro"
                opciones_clientes["CONTADO"] = "Venta Contado (No ligar a factura)"
                opciones_clientes["OTRO"] = "Otro (Gasto, Préstamo, etc.)"


                for idx, pago in st.session_state.df_pendientes.iterrows():
                    container_key = f"pago_{pago['id_banco_unico']}"
                    with st.expander(f"**{pago['fecha'].strftime('%Y-%m-%d')} - ${pago['valor']:,.0f}** - {pago['descripcion_banco']}", expanded=True):
                        
                        col_pago, col_asignacion = st.columns([1, 2])
                        
                        with col_pago:
                            st.markdown("**Detalle del Pago:**")
                            columnas_a_mostrar = [
                                'fecha', 'valor', 'descripcion_banco', 'SUCURSAL BANCO',
                                'TIPO DE TRANSACCION', 'BANCO REFRENCIA INTERNA'
                            ]
                            # Filtramos solo las columnas que existen en el DF 'pago'
                            columnas_existentes = [col for col in columnas_a_mostrar if col in pago.index]
                            st.dataframe(pago[columnas_existentes])
                        
                        with col_asignacion:
                            st.markdown("**Asignar Pago:**")
                            
                            nit_seleccionado = st.selectbox(
                                "Buscar Cliente por NIT o Nombre (o asignar a Contado/Otro):",
                                options=[""] + list(opciones_clientes.keys()),
                                format_func=lambda nit: "Selecciona un cliente..." if nit == "" else opciones_clientes[nit],
                                key=f"cliente_sel_{container_key}"
                            )
                            
                            if nit_seleccionado:
                                # --- LÓGICA PARA ASIGNACIÓN MANUAL ---
                                
                                # CASO 1: Es un cliente de cartera
                                if nit_seleccionado not in ["CONTADO", "OTRO"]:
                                    facturas_cliente = st.session_state.df_cartera[
                                        st.session_state.df_cartera['nit_norm'] == nit_seleccionado
                                    ].sort_values(by='DIAS_VENCIDO', ascending=False)
                                    
                                    if facturas_cliente.empty:
                                        st.warning("Este cliente no tiene facturas pendientes en cartera.")
                                    else:
                                        opciones_facturas = {
                                            row['id_factura_unica']: f"Fact: {row['id_factura_unica']} | Valor: ${row['IMPORTE']:,.0f} | Venc: {row['DIAS_VENCIDO']} días"
                                            for _, row in facturas_cliente.iterrows()
                                        }
                                        
                                        facturas_seleccionadas = st.multiselect(
                                            "Selecciona la(s) factura(s) que cubre este pago:",
                                            options=opciones_facturas.keys(),
                                            format_func=lambda id_fact: opciones_facturas[id_fact],
                                            key=f"fact_sel_{container_key}"
                                        )
                                        
                                        if st.button("💾 Guardar Conciliación (Cartera)", key=f"btn_save_cartera_{container_key}"):
                                            if not facturas_seleccionadas:
                                                st.error("Debes seleccionar al menos una factura.")
                                            else:
                                                pago_conciliado = pago.copy()
                                                pago_conciliado['status'] = 'Conciliado (Manual - Cartera)'
                                                pago_conciliado['id_factura_asignada'] = ", ".join(facturas_seleccionadas)
                                                pago_conciliado['cliente_asignado'] = clientes_cartera[nit_seleccionado]
                                                
                                                # Guardar en Google Sheets
                                                with st.spinner("Guardando..."):
                                                    g_client = connect_to_google_sheets()
                                                    ws = get_gsheet_worksheet(g_client, G_SHEET_URL, G_SHEET_TAB_CONCILIADOS)
                                                    
                                                    df_to_save = pd.DataFrame([pago_conciliado])
                                                    for col in df_to_save.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]']):
                                                        df_to_save[col] = df_to_save[col].astype(str)
                                                    
                                                    ws.append_rows(df_to_save.values.tolist(), value_input_option='USER_ENTERED')
                                                    
                                                    st.session_state.df_pendientes = st.session_state.df_pendientes.drop(idx)
                                                    st.success(f"¡Pago de {clientes_cartera[nit_seleccionado]} guardado!")
                                                    st.rerun()

                                # CASO 2: Es Venta de Contado u Otro
                                else:
                                    comentario_manual = st.text_input("Añadir comentario (ej. 'Abono cliente XXX', 'Pago datáfono')", key=f"comentario_{container_key}")
                                    
                                    if st.button(f"💾 Guardar como '{opciones_clientes[nit_seleccionado]}'", key=f"btn_save_otro_{container_key}"):
                                        if nit_seleccionado == "OTRO" and not comentario_manual:
                                            st.error("Debes añadir un comentario si seleccionas 'Otro'.")
                                        else:
                                            pago_conciliado = pago.copy()
                                            pago_conciliado['status'] = f'Conciliado (Manual - {nit_seleccionado})'
                                            pago_conciliado['id_factura_asignada'] = comentario_manual if comentario_manual else nit_seleccionado
                                            pago_conciliado['cliente_asignado'] = opciones_clientes[nit_seleccionado]
                                            
                                            # Guardar en Google Sheets
                                            with st.spinner("Guardando..."):
                                                g_client = connect_to_google_sheets()
                                                ws = get_gsheet_worksheet(g_client, G_SHEET_URL, G_SHEET_TAB_CONCILIADOS)
                                                
                                                df_to_save = pd.DataFrame([pago_conciliado])
                                                for col in df_to_save.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]']):
                                                    df_to_save[col] = df_to_save[col].astype(str)
                                                
                                                ws.append_rows(df_to_save.values.tolist(), value_input_option='USER_ENTERED')
                                                
                                                st.session_state.df_pendientes = st.session_state.df_pendientes.drop(idx)
                                                st.success(f"¡Pago guardado como {opciones_clientes[nit_seleccionado]}!")
                                                st.rerun()

        with tab_auto:
            st.info("Estos son los pagos que el motor identificó y guardó en Google Sheets automáticamente.")
            st.dataframe(st.session_state.df_conciliados_auto, use_container_width=True)

        with tab_fuentes:
            st.subheader("Cartera Pendiente (Fuente)")
            st.dataframe(st.session_state.df_cartera, use_container_width=True)
            
            st.subheader("Planilla Bancos (Fuente - Solo Pendientes de ID)")
            st.dataframe(st.session_state.df_bancos, use_container_width=True)
            
            st.subheader("Ventas de Contado (Fuente)")
            st.dataframe(st.session_state.df_ventas, use_container_width=True)

# --- Punto de entrada ---
if __name__ == '__main__':
    main_app()
