# ======================================================================================
# ARCHIVO: pages/2_Motor_Conciliacion.py
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
def get_dbx_client(secrets_key="dropbox"):
    """Crea un cliente de Dropbox usando las credenciales de secrets.toml."""
    try:
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
        st.stop()

def get_gsheet_worksheet(g_client, sheet_url, worksheet_name):
    """Accede a una pestaña específica de un Google Sheet por URL."""
    try:
        sheet = g_client.open_by_url(sheet_url)
        return sheet.worksheet(worksheet_name)
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"Error: No se encontró la pestaña '{worksheet_name}' en tu Google Sheet.")
        st.stop()
    except Exception as e:
        st.error(f"Error abriendo Google Sheet: {e}")
        st.stop()

def download_file_from_dropbox(dbx_client, file_path):
    """Descarga el contenido de un archivo desde Dropbox."""
    try:
        metadata, res = dbx_client.files_download(path=file_path)
        return res.content
    except dropbox.exceptions.ApiError as e:
        st.error(f"Error en API de Dropbox al descargar {file_path}: {e}")
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
    # Quitar tildes
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    # Quitar caracteres especiales y dejar solo letras, números y espacios
    texto = re.sub(r'[^A-Z0-9\s]', '', texto)
    # Quitar palabras comunes de bancos
    palabras_irrelevantes = ['PAGO', 'TRANSF', 'TRANSFERENCIA', 'CONSIGNACION', 'FACTURA', 'REF', 'BALOTO', 'EFECTY', 'PSE']
    for palabra in palabras_irrelevantes:
        texto = texto.replace(palabra, '')
    return ' '.join(texto.split()) # Normalizar espacios

@st.cache_data(ttl=600)
def cargar_cartera_actualizada():
    """Carga y procesa la cartera desde Dropbox (App 1)."""
    dbx_client = get_dbx_client("dropbox")
    content = download_file_from_dropbox(dbx_client, '/data/cartera_detalle.csv')
    if content:
        df = pd.read_csv(StringIO(content.decode('latin-1')), sep='|', header=None, names=[
            'Serie', 'Numero', 'Fecha Documento', 'Fecha Vencimiento', 'Cod Cliente',
            'NombreCliente', 'Nit', 'Poblacion', 'Provincia', 'Telefono1', 'Telefono2',
            'NomVendedor', 'Entidad Autoriza', 'E-Mail', 'Importe', 'Descuento',
            'Cupo Aprobado', 'Dias Vencido'
        ])
        
        # Procesamiento básico (similar al de tu app principal)
        df.columns = [normalizar_texto(col).lower().replace(' ', '_') for col in df.columns]
        df['importe'] = pd.to_numeric(df['importe'], errors='coerce').fillna(0)
        df['numero'] = pd.to_numeric(df['numero'], errors='coerce').fillna(0)
        df.loc[df['numero'] < 0, 'importe'] *= -1
        df['dias_vencido'] = pd.to_numeric(df['dias_vencido'], errors='coerce').fillna(0)
        
        # Columnas clave para conciliación
        df['id_factura_unica'] = df['serie'].astype(str) + '-' + df['numero'].astype(str)
        df['nit_norm'] = df['nit'].astype(str).str.replace(r'\D', '', regex=True)
        df['nombre_norm'] = df['nombrecliente'].apply(normalizar_texto)
        
        # Filtrar solo cartera pendiente de pago
        df_pendiente = df[df['importe'] > 0].copy()
        return df_pendiente
    return pd.DataFrame()

@st.cache_data(ttl=600)
def cargar_planilla_bancos(path_planilla_bancos):
    """Carga y limpia la planilla de bancos desde Dropbox (App 1)."""
    dbx_client = get_dbx_client("dropbox")
    content = download_file_from_dropbox(dbx_client, path_planilla_bancos)
    if content:
        try:
            # Asumimos que es un Excel, como mencionaste "planilla"
            df = pd.read_excel(BytesIO(content))
        except Exception as e:
            st.warning(f"No se pudo leer como Excel, intentando como CSV... ({e})")
            try:
                df = pd.read_csv(BytesIO(content), sep=';') # Ajustar separador si es necesario
            except Exception as e2:
                st.error(f"No se pudo leer el archivo de bancos: {e2}")
                return pd.DataFrame()
        
        # Nombres de columnas como los diste
        columnas_esperadas = ['FECHA', 'SUCURSAL BANCO', 'TIPO DE TRANSACCION', 'CUENTA', 
                              'EMPRESA', 'VALOR', 'BANCO REFRENCIA INTERNA', 'DESTINO', 
                              'RECIBO', 'FECHA RECIBO']
        df.columns = columnas_esperadas
        
        # --- Limpieza y Transformación (ETL) ---
        df_limpio = df.copy()
        df_limpio['fecha'] = pd.to_datetime(df_limpio['FECHA'], errors='coerce')
        df_limpio['valor'] = pd.to_numeric(df_limpio['VALOR'], errors='coerce').fillna(0)
        
        # Crear la columna de descripción unificada para el match
        df_limpio['descripcion_banco'] = (
            df_limpio['TIPO DE TRANSACCION'].fillna('') + ' ' +
            df_limpio['BANCO REFRENCIA INTERNA'].fillna('').astype(str) + ' ' +
            df_limpio['DESTINO'].fillna('')
        )
        
        # Crear columna de texto normalizado para el motor
        df_limpio['texto_match'] = df_limpio['descripcion_banco'].apply(normalizar_texto)
        
        # Filtrar solo ingresos
        df_ingresos = df_limpio[df_limpio['valor'] > 0].reset_index(drop=True)
        
        # Añadir un ID único a cada movimiento bancario
        df_ingresos['id_banco_unico'] = [f"B-{i+1}-{int(row['valor'])}" for i, row in df_ingresos.iterrows()]
        return df_ingresos
    return pd.DataFrame()

@st.cache_data(ttl=600)
def cargar_ventas_diarias(path_ventas_diarias):
    """Carga las ventas diarias (contado) desde Dropbox (App 2)."""
    dbx_client = get_dbx_client("dropbox_ventas") # ¡Usamos la otra credencial!
    content = download_file_from_dropbox(dbx_client, path_ventas_diarias)
    if content:
        try:
            # Asumimos CSV, ajusta si es Excel
            df = pd.read_csv(BytesIO(content), sep=';') 
        except Exception as e:
            st.error(f"No se pudo leer el archivo de ventas: {e}")
            return pd.DataFrame()
        
        # --- Limpieza y Transformación (ETL) ---
        # !! ESTA PARTE DEPENDE DE TU ARCHIVO DE VENTAS !!
        # Asumiré columnas 'Fecha', 'Cliente', 'Total_Factura', 'Forma_Pago'
        df.columns = [normalizar_texto(col).lower().replace(' ', '_') for col in df.columns]
        df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce')
        df['valor_contado'] = pd.to_numeric(df['total_factura'], errors='coerce').fillna(0)
        
        # Filtrar solo ventas que se pagan de contado
        df_contado = df[df['forma_pago'] == 'CONTADO'].copy() # Ajusta esta lógica
        return df_contado
    return pd.DataFrame()

# ======================================================================================
# --- 3. MOTOR DE CONCILIACIÓN AUTOMÁTICA ---
# ======================================================================================

def run_auto_reconciliation(df_bancos, df_cartera, df_ventas):
    """
    Ejecuta el motor de conciliación en cascada.
    Compara Bancos vs Cartera y Bancos vs Ventas de Contado.
    """
    st.write("Iniciando motor de conciliación automática...")
    
    conciliados = []
    ids_banco_conciliados = set()
    ids_factura_conciliadas = set()
    
    df_bancos_pendientes = df_bancos.copy()
    
    # --- NIVEL 1: MATCH PERFECTO (ID de Factura en Descripción) ---
    # Asume que el ID es 'SERIE-NUMERO' (ej. '155-12345')
    
    # Pre-calcular un set de facturas para búsqueda rápida
    mapa_facturas = {row['id_factura_unica']: row for _, row in df_cartera.iterrows()}
    
    for _, pago in df_bancos_pendientes.iterrows():
        # Buscar patrones como '155-12345' o '155 12345'
        matches = re.findall(r'(\d+[\s-]?\d+)', pago['texto_match'])
        for match in matches:
            id_factura_potencial = re.sub(r'\s', '-', match) # Normalizar '155 12345' a '155-12345'
            
            if id_factura_potencial in mapa_facturas:
                factura = mapa_facturas[id_factura_potencial]
                
                # Comprobar si el valor coincide (con tolerancia)
                if abs(pago['valor'] - factura['importe']) < 1000: # Tolerancia de $1000
                    pago['status'] = 'Conciliado (Auto N1 - ID Factura)'
                    pago['id_factura_asignada'] = id_factura_potencial
                    pago['cliente_asignado'] = factura['nombrecliente']
                    conciliados.append(pago)
                    
                    ids_banco_conciliados.add(pago['id_banco_unico'])
                    ids_factura_conciliadas.add(factura['id_factura_unica'])
                    break # Salir del bucle de matches
        if pago['id_banco_unico'] in ids_banco_conciliados:
            continue # Ir al siguiente pago
            
    # --- NIVEL 2: MATCH POR NIT + VALOR EXACTO ---
    # Pre-calcular un mapa de NITs a facturas
    mapa_nits = df_cartera[~df_cartera['id_factura_unica'].isin(ids_factura_conciliadas)] \
                .groupby('nit_norm')['importe'].apply(list).to_dict()

    for _, pago in df_bancos_pendientes[~df_bancos_pendientes['id_banco_unico'].isin(ids_banco_conciliados)].iterrows():
        # Extraer números largos (potenciales NITs)
        nits_potenciales = re.findall(r'(\d{8,10})', pago['texto_match'])
        for nit in nits_potenciales:
            if nit in mapa_nits:
                facturas_del_nit = mapa_nits[nit]
                # Buscar si el valor del pago coincide con alguna factura de ese NIT
                for valor_factura in facturas_del_nit:
                    if abs(pago['valor'] - valor_factura) < 1000:
                        # Encontramos un match
                        factura_match = df_cartera[
                            (df_cartera['nit_norm'] == nit) & 
                            (df_cartera['importe'] == valor_factura) &
                            (~df_cartera['id_factura_unica'].isin(ids_factura_conciliadas))
                        ].iloc[0]
                        
                        pago['status'] = 'Conciliado (Auto N2 - NIT+Valor)'
                        pago['id_factura_asignada'] = factura_match['id_factura_unica']
                        pago['cliente_asignado'] = factura_match['nombrecliente']
                        conciliados.append(pago)
                        
                        ids_banco_conciliados.add(pago['id_banco_unico'])
                        ids_factura_conciliadas.add(factura_match['id_factura_unica'])
                        mapa_nits[nit].remove(valor_factura) # Evitar doble asignación
                        break # Salir del bucle de facturas
            if pago['id_banco_unico'] in ids_banco_conciliados:
                break # Salir del bucle de NITs

    # --- NIVEL 3: MATCH VENTAS DE CONTADO ---
    # Compara pagos con ventas de contado del mismo día y valor
    if not df_ventas.empty:
        df_ventas['fecha_str'] = df_ventas['fecha'].dt.strftime('%Y-%m-%d')
        mapa_ventas = df_ventas.groupby('fecha_str')['valor_contado'].apply(list).to_dict()
        
        for _, pago in df_bancos_pendientes[~df_bancos_pendientes['id_banco_unico'].isin(ids_banco_conciliados)].iterrows():
            fecha_pago_str = pago['fecha'].strftime('%Y-%m-%d')
            if fecha_pago_str in mapa_ventas:
                for valor_venta in mapa_ventas[fecha_pago_str]:
                    if abs(pago['valor'] - valor_venta) < 100: # Tolerancia pequeña para contado
                        pago['status'] = 'Conciliado (Auto N3 - Venta Contado)'
                        pago['id_factura_asignada'] = f"CONTADO-{fecha_pago_str}"
                        pago['cliente_asignado'] = "Venta Contado" # Se puede mejorar si ventas tiene cliente
                        conciliados.append(pago)
                        
                        ids_banco_conciliados.add(pago['id_banco_unico'])
                        mapa_ventas[fecha_pago_str].remove(valor_venta) # Evitar doble asignación
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

    # --- Validar Autenticación (copiado de tu app principal) ---
    if not st.session_state.get('authentication_status', False):
        st.warning("Por favor, inicia sesión desde la página principal para acceder a esta herramienta.")
        st.stop()
    
    # --- PASO 1: DEFINIR PARÁMETROS ---
    st.header("1. Configuración de Archivos")
    
    # --- !! INFORMACIÓN QUE NECESITO DE TI !! ---
    with st.expander("Paths de Archivos (Ajustar si es necesario)", expanded=True):
        col1, col2 = st.columns(2)
        
        with col1:
            st.info("Google Sheets (Base de Datos)")
            default_gsheet_url = "https://docs.google.com/spreadsheets/d/TU_URL_UNICA_VA_AQUI/edit" # !! REEMPLAZAR !!
            G_SHEET_URL = st.text_input("URL del Google Sheet:", default_gsheet_url)
            G_SHEET_TAB_CONCILIADOS = st.text_input("Pestaña de Transacciones:", "Transacciones_Conciliadas")

        with col2:
            st.info("Dropbox Paths")
            PATH_PLANILLA_BANCOS = st.text_input("Path Planilla Bancos (App 1):", "/data/planilla_bancos.xlsx") # !! CONFIRMAR !!
            PATH_VENTAS_DIARIAS = st.text_input("Path Ventas Diarias (App 2):", "/Ventas/ventas_diarias.csv") # !! CONFIRMAR !!
    
    if G_SHEET_URL == default_gsheet_url:
        st.error("¡Acción Requerida! Por favor, reemplaza la URL de Google Sheets con tu propia URL.")
        st.stop()

    # --- Inicializar session_state para guardar los datos ---
    if 'data_loaded' not in st.session_state:
        st.session_state.data_loaded = False
        st.session_state.df_bancos = pd.DataFrame()
        st.session_state.df_cartera = pd.DataFrame()
        st.session_state.df_ventas = pd.DataFrame()
        st.session_state.df_conciliados_auto = pd.DataFrame()
        st.session_state.df_pendientes = pd.DataFrame()

    # --- PASO 2: BOTÓN DE CARGA Y PROCESAMIENTO ---
    st.header("2. Ejecutar Proceso ETL y Conciliación")
    
    if st.button("🔄 Cargar y Conciliar Datos", type="primary"):
        with st.spinner("Conectando y cargando datos..."):
            # Cargar todos los datos
            st.session_state.df_cartera = cargar_cartera_actualizada()
            st.session_state.df_bancos = cargar_planilla_bancos(PATH_PLANILLA_BANCOS)
            st.session_state.df_ventas = cargar_ventas_diarias(PATH_VENTAS_DIARIAS)
            st.session_state.data_loaded = True
            
            if st.session_state.df_bancos.empty or st.session_state.df_cartera.empty:
                st.error("Error: La planilla de bancos o la cartera no pudieron ser cargadas. Revisa los paths.")
                st.session_state.data_loaded = False
            else:
                st.success("Datos de Cartera, Bancos y Ventas cargados.")
        
        with st.spinner("Ejecutando motor de conciliación automática..."):
            # Cargar historial de Google Sheets para no duplicar
            g_client = connect_to_google_sheets()
            ws_conciliados = get_gsheet_worksheet(g_client, G_SHEET_URL, G_SHEET_TAB_CONCILIADOS)
            df_historico_gsheet = pd.DataFrame(ws_conciliados.get_all_records())
            
            ids_ya_conciliados = set()
            if not df_historico_gsheet.empty and 'id_banco_unico' in df_historico_gsheet.columns:
                ids_ya_conciliados = set(df_historico_gsheet['id_banco_unico'])
                st.write(f"Se encontraron {len(ids_ya_conciliados)} registros ya conciliados en Google Sheets.")
            
            # Filtrar bancos que ya están en la BD
            bancos_a_procesar = st.session_state.df_bancos[
                ~st.session_state.df_bancos['id_banco_unico'].isin(ids_ya_conciliados)
            ]
            
            if bancos_a_procesar.empty:
                st.info("No se encontraron nuevos movimientos bancarios para procesar.")
                st.session_state.df_pendientes = pd.DataFrame()
                st.session_state.df_conciliados_auto = pd.DataFrame()
            else:
                # Ejecutar el motor
                df_auto, df_pend = run_auto_reconciliation(
                    bancos_a_procesar, 
                    st.session_state.df_cartera, 
                    st.session_state.df_ventas
                )
                
                # Guardar resultados en session_state
                st.session_state.df_conciliados_auto = df_auto
                st.session_state.df_pendientes = df_pend
                
                # Guardar automáticos en Google Sheets
                if not df_auto.empty:
                    with st.spinner("Guardando conciliados automáticos en Google Sheets..."):
                        # Convertir columnas de fecha a string para GSheets
                        df_auto_save = df_auto.copy()
                        for col in df_auto_save.select_dtypes(include=['datetime64[ns]']).columns:
                            df_auto_save[col] = df_auto_save[col].astype(str)
                        
                        # Obtener las filas existentes para añadir las nuevas
                        existing_data = ws_conciliados.get_all_values()
                        list_of_lists_to_add = df_auto_save.values.tolist()
                        
                        # Si hay cabeceras, añadir solo los datos
                        if len(existing_data) > 0:
                            ws_conciliados.append_rows(list_of_lists_to_add, value_input_option='USER_ENTERED')
                        else: # Si está vacía, añadir con cabeceras
                            set_with_dataframe(ws_conciliados, df_auto_save)
                            
                        st.success(f"{len(df_auto)} pagos automáticos guardados en Google Sheets.")

    # --- PASO 3: KPIs Y TABS DE RESULTADOS ---
    if st.session_state.data_loaded:
        st.header("3. Resultados de la Conciliación")
        
        total_recibido = st.session_state.df_bancos['valor'].sum()
        total_auto = st.session_state.df_conciliados_auto['valor'].sum()
        total_pendiente = st.session_state.df_pendientes['valor'].sum()

        kpi_cols = st.columns(3)
        kpi_cols[0].metric("🏦 Total Recibido (Bancos)", f"${total_recibido:,.0f}")
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
                
                # Crear la lista de clientes para el buscador
                clientes_cartera = st.session_state.df_cartera.drop_duplicates(subset=['nit_norm']) \
                                       .set_index('nit_norm')['nombrecliente'].to_dict()
                opciones_clientes = {nit: f"{nombre} (NIT: {nit})" for nit, nombre in clientes_cartera.items()}

                # --- El Centro de Comando Manual ---
                for idx, pago in st.session_state.df_pendientes.iterrows():
                    container_key = f"pago_{pago['id_banco_unico']}"
                    with st.expander(f"**{pago['fecha'].strftime('%Y-%m-%d')} - ${pago['valor']:,.0f}** - {pago['descripcion_banco']}", expanded=True):
                        
                        col_pago, col_asignacion = st.columns([1, 2])
                        
                        with col_pago:
                            st.markdown("**Detalle del Pago:**")
                            st.dataframe(pago.drop(['texto_match', 'status']))
                        
                        with col_asignacion:
                            st.markdown("**Asignar Pago a Cliente:**")
                            
                            # 1. Buscar Cliente
                            nit_seleccionado = st.selectbox(
                                "Buscar Cliente por NIT o Nombre:",
                                options=[""] + list(opciones_clientes.keys()),
                                format_func=lambda nit: "Selecciona un cliente..." if nit == "" else opciones_clientes[nit],
                                key=f"cliente_sel_{container_key}"
                            )
                            
                            if nit_seleccionado:
                                # 2. Mostrar Facturas de ese cliente
                                facturas_cliente = st.session_state.df_cartera[
                                    st.session_state.df_cartera['nit_norm'] == nit_seleccionado
                                ].sort_values(by='dias_vencido', ascending=False)
                                
                                if facturas_cliente.empty:
                                    st.warning("Este cliente no tiene facturas pendientes en cartera.")
                                else:
                                    # Convertir facturas a un formato legible para multiselect
                                    opciones_facturas = {
                                        row['id_factura_unica']: f"Fact: {row['id_factura_unica']} | Valor: ${row['importe']:,.0f} | Venc: {row['dias_vencido']} días"
                                        for _, row in facturas_cliente.iterrows()
                                    }
                                    
                                    facturas_seleccionadas = st.multiselect(
                                        "Selecciona la(s) factura(s) que cubre este pago:",
                                        options=opciones_facturas.keys(),
                                        format_func=lambda id_fact: opciones_facturas[id_fact],
                                        key=f"fact_sel_{container_key}"
                                    )
                                    
                                    # 3. Botón de Guardar
                                    if st.button("💾 Guardar Conciliación Manual", key=f"btn_save_{container_key}"):
                                        if not facturas_seleccionadas:
                                            st.error("Debes seleccionar al menos una factura.")
                                        else:
                                            with st.spinner("Guardando..."):
                                                # Preparar la fila para Google Sheets
                                                pago_conciliado = pago.copy()
                                                pago_conciliado['status'] = 'Conciliado (Manual)'
                                                pago_conciliado['id_factura_asignada'] = ", ".join(facturas_seleccionadas)
                                                pago_conciliado['cliente_asignado'] = clientes_cartera[nit_seleccionado]
                                                
                                                # Convertir a DataFrame y limpiar para GSheets
                                                df_to_save = pd.DataFrame([pago_conciliado])
                                                for col in df_to_save.select_dtypes(include=['datetime64[ns]']).columns:
                                                    df_to_save[col] = df_to_save[col].astype(str)
                                                
                                                # Guardar en Google Sheets
                                                g_client = connect_to_google_sheets()
                                                ws = get_gsheet_worksheet(g_client, G_SHEET_URL, G_SHEET_TAB_CONCILIADOS)
                                                ws.append_rows(df_to_save.values.tolist(), value_input_option='USER_ENTERED')
                                                
                                                # Actualizar el session_state (quitar el pago de pendientes)
                                                st.session_state.df_pendientes = st.session_state.df_pendientes.drop(idx)
                                                st.success(f"¡Pago de {clientes_cartera[nit_seleccionado]} guardado!")
                                                st.rerun() # Recargar la página para que desaparezca el expander

        with tab_auto:
            st.info("Estos son los pagos que el motor identificó y guardó en Google Sheets automáticamente.")
            st.dataframe(st.session_state.df_conciliados_auto, use_container_width=True)

        with tab_fuentes:
            st.subheader("Cartera Pendiente (Fuente)")
            st.dataframe(st.session_state.df_cartera, use_container_width=True)
            
            st.subheader("Planilla Bancos (Fuente)")
            st.dataframe(st.session_state.df_bancos, use_container_width=True)
            
            st.subheader("Ventas de Contado (Fuente)")
            st.dataframe(st.session_state.df_ventas, use_container_width=True)

# --- Punto de entrada ---
if __name__ == '__main__':
    main_app()
