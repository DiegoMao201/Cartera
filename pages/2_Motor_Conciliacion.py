# ======================================================================================
# ARCHIVO: pages/2_Motor_Conciliacion.py
# (Versión MODIFICADA POR GEMINI - 27 de Octubre, 2025)
#
# MODIFICACIÓN (Gemini):
# 1.  Se re-arquitectura el flujo para cumplir la solicitud del usuario:
#     - Se crea una "Base Maestra" de Bancos en un Google Sheet.
#     - El script ahora tiene dos flujos:
#       a) Un BATCH (Admin) que lee el crudo de Dropbox, lo enriquece
#          inteligentemente contra Cartera y Ventas, y SOBRE-ESCRIBE
#          la hoja "Bancos_Master" en Google Sheets.
#       b) Un flujo INTERACTIVO (Usuario) que LEE desde "Bancos_Master"
#          y filtra solo los pagos pendientes para la asignación manual.
#
# 2.  Se crea `cargar_planilla_bancos_RAW` para leer el archivo de bancos
#     de Dropbox sin aplicar NINGÚN filtro.
#
# 3.  Se crea `correr_batch_conciliacion_inteligente` (reemplaza al antiguo
#     `run_auto_reconciliation`). Este motor corre sobre TODO el
#     dataframe de bancos y añade columnas ('match_status', 'match_cliente', etc.).
#     - Se añade Lógica de Match por Nombre (Fuzzywuzzy) como Nivel 3.
#
# 4.  Se modifica `cargar_planilla_bancos` (ahora `cargar_pendientes_desde_master`)
#     para LEER desde el G-Sheet "Bancos_Master" y filtrar los pendientes.
#
# 5.  La UI de `main_app` se divide en "Paso 1: Actualizar Base Maestra" y
#     "Paso 2: Asignación Manual".
# ======================================================================================

import streamlit as st
import pandas as pd
import dropbox
from io import StringIO, BytesIO
import re
import unicodedata
from datetime import datetime
from fuzzywuzzy import fuzz, process
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
        
        # Optimizamos para el match de nombres
        df_pendiente = df_pendiente[df_pendiente['nombre_norm'] != '']
        return df_pendiente
        
    return pd.DataFrame()

@st.cache_data(ttl=600)
def cargar_planilla_bancos_RAW(path_planilla_bancos):
    """
    (NUEVA FUNCIÓN)
    Carga la planilla de bancos CRUDA desde Dropbox, sin filtros,
    solo para estandarizarla antes del batch.
    """
    dbx_client = get_dbx_client("dropbox")
    content = download_file_from_dropbox(dbx_client, path_planilla_bancos)
    
    if not content:
        st.error(f"No se pudo descargar la planilla de bancos de: {path_planilla_bancos}")
        return pd.DataFrame()

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
    
    # --- NO APLICAMOS FILTROS ---
    # Procesamos todo el archivo
    
    df_limpio['descripcion_banco'] = (
        df_limpio['TIPO DE TRANSACCION'].fillna('').astype(str) + ' ' +
        df_limpio['BANCO REFRENCIA INTERNA'].fillna('').astype(str) + ' ' +
        df_limpio['DESTINO'].fillna('').astype(str)
    )
    df_limpio['texto_match'] = df_limpio['descripcion_banco'].apply(normalizar_texto)
    df_limpio = df_limpio.reset_index(drop=True)
    
    df_limpio['id_banco_unico'] = df_limpio.apply(
        lambda row: f"B-{row['fecha'].strftime('%Y%m%d') if pd.notna(row['fecha']) else 'SINFCHA'}-{int(row['valor'])}-{row.name}", axis=1
    )
    
    return df_limpio

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
        
@st.cache_data(ttl=600)
def cargar_pendientes_desde_master(g_client, sheet_url, master_tab_name):
    """
    (FUNCIÓN MODIFICADA)
    Lee la "Base Maestra" desde Google Sheets y filtra los
    movimientos que están PENDIENTES de asignación manual.
    """
    st.write(f"Cargando datos desde G-Sheet '{master_tab_name}'...")
    try:
        ws_master = get_gsheet_worksheet(g_client, sheet_url, master_tab_name)
        df_master = pd.DataFrame(ws_master.get_all_records())
        
        if df_master.empty:
            st.warning("La Base Maestra de Bancos está vacía. Ejecuta el 'Paso 1' primero.")
            return pd.DataFrame(), pd.DataFrame()
            
        # Convertir columnas numéricas y de fecha
        df_master['valor'] = pd.to_numeric(df_master['valor'], errors='coerce').fillna(0)
        df_master['fecha'] = pd.to_datetime(df_master['fecha'], errors='coerce')

        # Filtramos los que SÍ están conciliados
        df_conciliados = df_master[
            df_master['match_status'] != 'Pendiente (Revisión Manual)'
        ].copy()
        
        # Filtramos los que NO están conciliados (pendientes)
        df_pendientes = df_master[
            df_master['match_status'] == 'Pendiente (Revisión Manual)'
        ].copy()

        st.success(f"Carga desde G-Sheet completa: {len(df_conciliados)} conciliados (auto), {len(df_pendientes)} pendientes (manual).")
        return df_conciliados, df_pendientes

    except Exception as e:
        st.error(f"Error al leer la Base Maestra de Google Sheets: {e}")
        st.info("Asegúrate de que la pestaña 'Bancos_Master' exista y tenga datos.")
        return pd.DataFrame(), pd.DataFrame()


# ======================================================================================
# --- 3. MOTOR DE CONCILIACIÓN (BATCH INTELIGENTE) ---
# ======================================================================================

def correr_batch_conciliacion_inteligente(df_bancos_raw, df_cartera, df_ventas, df_historico_manual):
    """
    (NUEVA FUNCIÓN - REEMPLAZA A run_auto_reconciliation)
    
    Ejecuta el motor de conciliación en MODO BATCH sobre TODOS los
    movimientos de bancos. Añade columnas de 'match' en lugar de
    filtrar.
    """
    st.write("Iniciando batch de conciliación inteligente...")
    
    df_bancos = df_bancos_raw.copy()
    
    # Preparamos el DF de salida
    df_bancos['match_status'] = 'Pendiente (Revisión Manual)'
    df_bancos['match_cliente'] = ''
    df_bancos['match_factura_id'] = ''
    
    # Pre-filtramos movimientos que no son ingresos
    egresos_idx = df_bancos[df_bancos['valor'] <= 0].index
    df_bancos.loc[egresos_idx, 'match_status'] = 'Egreso (No Aplica)'
    
    # Pre-filtramos movimientos que YA están identificados en la planilla
    identificados_idx = df_bancos[
        (df_bancos['RECIBO'].fillna('').astype(str).str.strip() != '') |
        (df_bancos['DESTINO'].fillna('').astype(str).str.strip() != '')
    ].index
    df_bancos.loc[identificados_idx, 'match_status'] = 'Identificado (Origen)'
    df_bancos.loc[identificados_idx, 'match_cliente'] = 'Asignado en Planilla'
    
    # Pre-filtramos movimientos que YA fueron conciliados MANUALMENTE (histórico)
    ids_manuales = set()
    if not df_historico_manual.empty and 'id_banco_unico' in df_historico_manual.columns:
        ids_manuales = set(df_historico_manual['id_banco_unico'])
        
    manuales_idx = df_bancos[df_bancos['id_banco_unico'].isin(ids_manuales)].index
    
    if not manuales_idx.empty:
        st.write(f"Omitiendo {len(manuales_idx)} registros ya conciliados manualmente.")
        mapa_manual = df_historico_manual.set_index('id_banco_unico')
        
        def map_status_manual(row):
            if row['id_banco_unico'] in mapa_manual.index:
                data = mapa_manual.loc[row['id_banco_unico']]
                row['match_status'] = data.get('status', 'Conciliado (Manual - Histórico)')
                row['match_cliente'] = data.get('cliente_asignado', 'Manual')
                row['match_factura_id'] = data.get('id_factura_asignada', 'Manual')
            return row
            
        df_bancos.loc[manuales_idx] = df_bancos.loc[manuales_idx].apply(map_status_manual, axis=1)

    # --- INICIO DEL MOTOR DE CRUCE INTELIGENTE ---
    
    # IDs de bancos que aún se pueden procesar
    ids_pendientes_proceso = df_bancos[
        (df_bancos['valor'] > 0) &
        (df_bancos['match_status'] == 'Pendiente (Revisión Manual)')
    ]['id_banco_unico']
    
    if ids_pendientes_proceso.empty:
        st.info("No se encontraron nuevos movimientos para el batch de conciliación.")
        return df_bancos

    st.write(f"Ejecutando motor sobre {len(ids_pendientes_proceso)} nuevos movimientos...")
    
    # --- NIVEL 1: MATCH PERFECTO (ID de Factura en Descripción) ---
    mapa_facturas = {row['id_factura_unica']: row for _, row in df_cartera.iterrows()}
    ids_factura_conciliadas = set()
    
    for idx, pago in df_bancos[df_bancos['id_banco_unico'].isin(ids_pendientes_proceso)].iterrows():
        matches = re.findall(r'(\d+-\d+)', pago['texto_match'])
        for id_factura_potencial in matches:
            if id_factura_potencial in mapa_facturas:
                factura = mapa_facturas[id_factura_potencial]
                if abs(pago['valor'] - factura['IMPORTE']) < 1000: # Tolerancia $1000
                    df_bancos.loc[idx, 'match_status'] = 'Conciliado (Auto N1 - ID Factura)'
                    df_bancos.loc[idx, 'match_cliente'] = factura['NOMBRECLIENTE']
                    df_bancos.loc[idx, 'match_factura_id'] = id_factura_potencial
                    ids_pendientes_proceso = ids_pendientes_proceso.drop(pago['id_banco_unico'])
                    ids_factura_conciliadas.add(id_factura_potencial)
                    break 

    # --- NIVEL 2: MATCH POR NIT + VALOR EXACTO ---
    cartera_restante = df_cartera[~df_cartera['id_factura_unica'].isin(ids_factura_conciliadas)]
    mapa_nits = cartera_restante.groupby('nit_norm')['IMPORTE'].apply(list).to_dict()

    for idx, pago in df_bancos[df_bancos['id_banco_unico'].isin(ids_pendientes_proceso)].iterrows():
        nits_potenciales = re.findall(r'(\d{8,10})', pago['texto_match'])
        for nit in nits_potenciales:
            if nit in mapa_nits:
                for valor_factura in mapa_nits[nit]:
                    if abs(pago['valor'] - valor_factura) < 1000:
                        factura_match = cartera_restante[
                            (cartera_restante['nit_norm'] == nit) & 
                            (cartera_restante['IMPORTE'] == valor_factura)
                        ].iloc[0]
                        
                        df_bancos.loc[idx, 'match_status'] = 'Conciliado (Auto N2 - NIT+Valor)'
                        df_bancos.loc[idx, 'match_cliente'] = factura_match['NOMBRECLIENTE']
                        df_bancos.loc[idx, 'match_factura_id'] = factura_match['id_factura_unica']
                        
                        ids_pendientes_proceso = ids_pendientes_proceso.drop(pago['id_banco_unico'])
                        ids_factura_conciliadas.add(factura_match['id_factura_unica'])
                        mapa_nits[nit].remove(valor_factura) 
                        break 
            if pago['id_banco_unico'] not in ids_pendientes_proceso:
                break 

    # --- NIVEL 3: MATCH POR NOMBRE (FUZZY) + VALOR ---
    # (Este es el cruce "inteligente" por nombre que solicitaste)
    cartera_restante = df_cartera[~df_cartera['id_factura_unica'].isin(ids_factura_conciliadas)]
    
    # Creamos un mapa de nombres únicos y sus facturas
    mapa_nombres = cartera_restante.groupby('nombre_norm').agg(
        n_facturas=('id_factura_unica', 'count'),
        importe_total=('IMPORTE', 'sum'),
        nit=('nit_norm', 'first'),
        nombre_real=('NOMBRECLIENTE', 'first')
    ).reset_index()
    
    lista_nombres_cartera = mapa_nombres['nombre_norm'].tolist()

    if lista_nombres_cartera: # Solo si hay nombres en cartera
        for idx, pago in df_bancos[df_bancos['id_banco_unico'].isin(ids_pendientes_proceso)].iterrows():
            if not pago['texto_match']:
                continue

            # Buscamos el mejor match de nombre
            mejor_match = process.extractOne(pago['texto_match'], lista_nombres_cartera, scorer=fuzz.partial_ratio, score_cutoff=90)
            
            if mejor_match:
                nombre_encontrado = mejor_match[0]
                cliente_data = mapa_nombres[mapa_nombres['nombre_norm'] == nombre_encontrado].iloc[0]
                
                # Verificamos si el valor pagado se acerca al total de la deuda de ese cliente
                if abs(pago['valor'] - cliente_data['importe_total']) < 5000: # Tolerancia alta
                    
                    df_bancos.loc[idx, 'match_status'] = 'Conciliado (Auto N3 - Nombre+ValorTotal)'
                    df_bancos.loc[idx, 'match_cliente'] = cliente_data['nombre_real']
                    df_bancos.loc[idx, 'match_factura_id'] = f"Abono Total Deuda (NIT {cliente_data['nit']})"
                    ids_pendientes_proceso = ids_pendientes_proceso.drop(pago['id_banco_unico'])
                    # No removemos facturas de cartera, ya que es un abono general
                    break


    # --- NIVEL 4: MATCH VENTAS DE CONTADO ---
    if not df_ventas.empty:
        df_ventas['fecha_str'] = df_ventas['fecha'].dt.strftime('%Y-%m-%d')
        mapa_ventas = df_ventas.groupby('fecha_str')['valor_contado'].apply(list).to_dict()
        
        for idx, pago in df_bancos[df_bancos['id_banco_unico'].isin(ids_pendientes_proceso)].iterrows():
            fecha_pago_str = pago['fecha'].strftime('%Y-%m-%d')
            if fecha_pago_str in mapa_ventas:
                for valor_venta in mapa_ventas[fecha_pago_str]:
                    if abs(pago['valor'] - valor_venta) < 100: # Tolerancia baja
                        df_bancos.loc[idx, 'match_status'] = 'Conciliado (Auto N4 - Venta Contado)'
                        df_bancos.loc[idx, 'match_cliente'] = 'Venta Contado'
                        df_bancos.loc[idx, 'match_factura_id'] = f"CONTADO-{fecha_pago_str}"
                        ids_pendientes_proceso = ids_pendientes_proceso.drop(pago['id_banco_unico'])
                        mapa_ventas[fecha_pago_str].remove(valor_venta) 
                        break

    # --- Finalizar ---
    total_auto = len(df_bancos_raw) - len(ids_pendientes_proceso) - len(egresos_idx) - len(identificados_idx)
    st.success(f"Batch finalizado: {total_auto} pagos conciliados automáticamente.")
    
    # Devolvemos el DF completo y enriquecido
    return df_bancos


# ======================================================================================
# --- 4. APLICACIÓN PRINCIPAL DE STREAMLIT ---
# ======================================================================================

def main_app():
    
    st.title("🤖 Motor de Conciliación Bancaria (v2 - Base Maestra)")
    st.markdown("Proceso de 2 pasos: **1.** Actualizar la Base Maestra (Admin) y **2.** Asignar Pendientes (Usuario).")

    # --- Validar Autenticación ---
    if not st.session_state.get('authentication_status', False):
        st.warning("Por favor, inicia sesión desde la página principal para acceder a esta herramienta.")
        st.stop()
    
    # --- Cargar configuración desde secrets.toml ---
    try:
        # Configuración de Google Sheets
        G_SHEET_URL = st.secrets["google_sheets"]["sheet_url"]
        G_SHEET_TAB_CONCILIADOS_MANUAL = st.secrets["google_sheets"]["tab_conciliados"]
        G_SHEET_TAB_BANCOS_MASTER = st.secrets["google_sheets"]["tab_bancos_master"] # <-- NUEVA CLAVE
        
        # Path de Bancos (de la App 'dropbox')
        PATH_PLANILLA_BANCOS = st.secrets["dropbox"]["path_bancos"]
        # Path de Ventas (de la App 'dropbox_ventas')
        PATH_VENTAS_DIARIAS = st.secrets["dropbox_ventas"]["path_ventas"]
        
    except KeyError as e:
        st.error(f"Error: Falta una clave en tu archivo secrets.toml: {e}")
        st.info("Revisa la estructura de ejemplo y asegúrate de que [google_sheets] tenga 'tab_bancos_master'.")
        st.stop()

    # --- Inicializar session_state para guardar los datos ---
    if 'data_loaded' not in st.session_state:
        st.session_state.data_loaded = False
        st.session_state.df_cartera = pd.DataFrame()
        st.session_state.df_conciliados_auto = pd.DataFrame()
        st.session_state.df_pendientes = pd.DataFrame()

    # ==================================================================
    # --- PASO 1: BATCH DE ACTUALIZACIÓN DE BASE MAESTRA ---
    # ==================================================================
    st.markdown("---")
    st.header("PASO 1: [ADMIN] Actualizar Base Maestra")
    st.info("Este proceso lee el archivo de bancos de Dropbox, lo cruza con Cartera y Ventas, y sobre-escribe la 'Base Maestra' en Google Sheets. **Ejecutar 1 vez al día.**")
    
    if st.button("🚀 Ejecutar Batch y Actualizar 'Bancos_Master' en G-Sheets"):
        
        g_client = connect_to_google_sheets()
        
        with st.spinner("Cargando fuentes de datos (Cartera, Bancos Crudo, Ventas)..."):
            df_cartera = cargar_y_procesar_cartera()
            df_bancos_raw = cargar_planilla_bancos_RAW(PATH_PLANILLA_BANCOS)
            df_ventas = cargar_ventas_diarias(PATH_VENTAS_DIARIAS)

            if df_bancos_raw.empty or df_cartera.empty:
                st.error("No se pudo cargar la Cartera o la Planilla de Bancos. El batch no puede continuar.")
                st.stop()
        
        with st.spinner("Cargando historial de conciliaciones manuales..."):
            try:
                ws_manuales = get_gsheet_worksheet(g_client, G_SHEET_URL, G_SHEET_TAB_CONCILIADOS_MANUAL)
                df_historico_manual = pd.DataFrame(ws_manuales.get_all_records())
            except Exception as e:
                st.warning(f"No se pudo leer historial manual (puede estar vacía): {e}")
                df_historico_manual = pd.DataFrame()

        with st.spinner("Ejecutando motor de conciliación inteligente (Batch)..."):
            df_bancos_enriquecido = correr_batch_conciliacion_inteligente(
                df_bancos_raw, df_cartera, df_ventas, df_historico_manual
            )
        
        with st.spinner(f"Guardando {len(df_bancos_enriquecido)} registros en G-Sheet '{G_SHEET_TAB_BANCOS_MASTER}'..."):
            ws_master = get_gsheet_worksheet(g_client, G_SHEET_URL, G_SHEET_TAB_BANCOS_MASTER)
            
            # Limpiamos el DF para guardarlo (G-Sheets no acepta NaT/NaN)
            df_to_save = df_bancos_enriquecido.fillna('')
            for col in df_to_save.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]']):
                df_to_save[col] = df_to_save[col].astype(str)
            
            ws_master.clear() # Borramos la hoja
            set_with_dataframe(ws_master, df_to_save) # La sobre-escribimos
            st.success(f"¡Éxito! Base Maestra '{G_SHEET_TAB_BANCOS_MASTER}' actualizada con {len(df_to_save)} registros.")
            st.session_state.data_loaded = False # Forzamos recarga de pendientes

    # ==================================================================
    # --- PASO 2: CARGA Y ASIGNACIÓN MANUAL ---
    # ==================================================================
    st.markdown("---")
    st.header("PASO 2: [USUARIO] Cargar Pendientes y Asignar")
    st.info("Este proceso lee la 'Base Maestra' de Google Sheets y carga SÓLO los pagos que el robot no pudo identificar, para que los asignes manualmente.")
    
    if st.button("🔄 Cargar Pendientes de Asignación Manual", type="primary"):
        with st.spinner("Cargando datos desde la Base Maestra de Google Sheets..."):
            
            g_client = connect_to_google_sheets()
            st.session_state.df_cartera = cargar_y_procesar_cartera()
            
            df_auto, df_pend = cargar_pendientes_desde_master(
                g_client, G_SHEET_URL, G_SHEET_TAB_BANCOS_MASTER
            )
            
            st.session_state.df_conciliados_auto = df_auto
            st.session_state.df_pendientes = df_pend
            
            if st.session_state.df_cartera.empty:
                st.error("Error: La cartera no pudo ser cargada. Revisa el path '/data/cartera_detalle.csv'.")
            else:
                st.session_state.data_loaded = True

    # --- RESULTADOS DE LA CONCILIACIÓN (Leídos desde G-Sheets) ---
    if st.session_state.data_loaded:
        st.header("Resultados de la Conciliación (desde Base Maestra)")
        
        total_auto = st.session_state.df_conciliados_auto['valor'].sum() if not st.session_state.df_conciliados_auto.empty else 0
        total_pendiente = st.session_state.df_pendientes['valor'].sum() if not st.session_state.df_pendientes.empty else 0
        total_recibido = total_auto + total_pendiente # Solo de los > 0

        kpi_cols = st.columns(3)
        kpi_cols[0].metric("🏦 Total Identificado (Base Maestra)", f"${total_recibido:,.0f}")
        kpi_cols[1].metric("✅ Conciliado (Automático)", f"${total_auto:,.0f}")
        kpi_cols[2].metric("❓ Pendiente (Manual)", f"${total_pendiente:,.0f}", delta=f"{len(st.session_state.df_pendientes)} transacciones")

        tab_manual, tab_auto, tab_fuentes = st.tabs(
            ["📝 **PENDIENTE DE ASIGNACIÓN MANUAL**", "🤖 Conciliados (Automáticos)", "🗂️ Cartera (Fuente)"]
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
                opciones_clientes["CONTADO"] = "Venta Contado (No ligar a factura)"
                opciones_clientes["OTRO"] = "Otro (Gasto, Préstamo, etc.)"


                for idx, pago in st.session_state.df_pendientes.iterrows():
                    container_key = f"pago_{pago['id_banco_unico']}"
                    with st.expander(f"**{pago['fecha']} - ${pago['valor']:,.0f}** - {pago['descripcion_banco']}", expanded=True):
                        
                        col_pago, col_asignacion = st.columns([1, 2])
                        
                        with col_pago:
                            st.markdown("**Detalle del Pago:**")
                            columnas_a_mostrar = [
                                'fecha', 'valor', 'descripcion_banco', 'SUCURSAL BANCO',
                                'TIPO DE TRANSACCION', 'BANCO REFRENCIA INTERNA',
                                'id_banco_unico' # Mostramos el ID único
                            ]
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
                                                
                                                # Guardar en Google Sheets (Histórico Manual)
                                                with st.spinner("Guardando en G-Sheet 'Conciliados_Historico' (y removiendo de 'Pendientes')..."):
                                                    g_client = connect_to_google_sheets()
                                                    ws = get_gsheet_worksheet(g_client, G_SHEET_URL, G_SHEET_TAB_CONCILIADOS_MANUAL)
                                                    
                                                    df_to_save = pd.DataFrame([pago_conciliado])
                                                    for col in df_to_save.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]']):
                                                        df_to_save[col] = df_to_save[col].astype(str)
                                                    
                                                    # Preparamos las columnas que existen en la hoja de conciliados
                                                    # (Usamos las de la hoja original para no dañarla)
                                                    cols_finales_en_df = [
                                                        'FECHA', 'SUCURSAL BANCO', 'TIPO DE TRANSACCION', 'CUENTA',
                                                        'EMPRESA', 'VALOR', 'BANCO REFRENCIA INTERNA', 'DESTINO', 'RECIBO',
                                                        'FECHA RECIBO', 'fecha', 'valor', 'RECIBO_norm', 'DESTINO_norm',
                                                        'descripcion_banco', 'texto_match', 'id_banco_unico',
                                                        'status', 'id_factura_asignada', 'cliente_asignado'
                                                    ]
                                                    
                                                    # Creamos un DF vacío con todas las columnas posibles y lo llenamos
                                                    df_final_save = pd.DataFrame(columns=cols_finales_en_df)
                                                    df_final_save = pd.concat([df_final_save, df_to_save], ignore_index=True)
                                                    df_final_save = df_final_save[cols_finales_en_df].fillna('')

                                                    # Leemos los headers actuales de la hoja de manuales
                                                    headers = ws.row_values(1)
                                                    if not headers: # Si la hoja está vacía, escribimos headers
                                                        set_with_dataframe(ws, df_final_save[cols_finales_en_df])
                                                    else:
                                                        # Si ya tiene headers, solo apilamos los valores
                                                        ws.append_rows(df_final_save[cols_finales_en_df].values.tolist(), value_input_option='USER_ENTERED')
                                                    
                                                    st.session_state.df_pendientes = st.session_state.df_pendientes.drop(idx)
                                                    st.success(f"¡Pago de {clientes_cartera[nit_seleccionado]} guardado en G-Sheets!")
                                                    st.info("El registro se marcará como 'Manual' la próxima vez que se ejecute el 'Paso 1' (Batch).")
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
                                            
                                            # Guardar en Google Sheets (Histórico Manual)
                                            with st.spinner("Guardando en G-Sheet 'Conciliados_Historico' (y removiendo de 'Pendientes')..."):
                                                g_client = connect_to_google_sheets()
                                                ws = get_gsheet_worksheet(g_client, G_SHEET_URL, G_SHEET_TAB_CONCILIADOS_MANUAL)
                                                
                                                df_to_save = pd.DataFrame([pago_conciliado])
                                                for col in df_to_save.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]']):
                                                    df_to_save[col] = df_to_save[col].astype(str)

                                                # (Misma lógica de guardado que el caso 1)
                                                cols_finales_en_df = [
                                                    'FECHA', 'SUCURSAL BANCO', 'TIPO DE TRANSACCION', 'CUENTA',
                                                    'EMPRESA', 'VALOR', 'BANCO REFRENCIA INTERNA', 'DESTINO', 'RECIBO',
                                                    'FECHA RECIBO', 'fecha', 'valor', 'RECIBO_norm', 'DESTINO_norm',
                                                    'descripcion_banco', 'texto_match', 'id_banco_unico',
                                                    'status', 'id_factura_asignada', 'cliente_asignado'
                                                ]
                                                df_final_save = pd.DataFrame(columns=cols_finales_en_df)
                                                df_final_save = pd.concat([df_final_save, df_to_save], ignore_index=True)
                                                df_final_save = df_final_save[cols_finales_en_df].fillna('')

                                                headers = ws.row_values(1)
                                                if not headers:
                                                    set_with_dataframe(ws, df_final_save[cols_finales_en_df])
                                                else:
                                                    ws.append_rows(df_final_save[cols_finales_en_df].values.tolist(), value_input_option='USER_ENTERED')
                                                
                                                st.session_state.df_pendientes = st.session_state.df_pendientes.drop(idx)
                                                st.success(f"¡Pago guardado como {opciones_clientes[nit_seleccionado]} en G-Sheets!")
                                                st.info("El registro se marcará como 'Manual' la próxima vez que se ejecute el 'Paso 1' (Batch).")
                                                st.rerun()

        with tab_auto:
            st.info("Estos son los pagos que el motor identificó automáticamente (leídos desde la Base Maestra).")
            columnas_auto = [
                'fecha', 'valor', 'descripcion_banco', 'match_status', 'match_cliente', 'match_factura_id', 'id_banco_unico'
            ]
            columnas_existentes_auto = [col for col in columnas_auto if col in st.session_state.df_conciliados_auto.columns]
            st.dataframe(st.session_state.df_conciliados_auto[columnas_existentes_auto], use_container_width=True)

        with tab_fuentes:
            st.subheader("Cartera Pendiente (Fuente)")
            st.dataframe(st.session_state.df_cartera, use_container_width=True)


# --- Punto de entrada ---
if __name__ == '__main__':
    main_app()
