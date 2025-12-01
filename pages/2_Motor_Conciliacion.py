# ======================================================================================
# ARCHIVO: pages/2_Motor_Conciliacion.py
# (Versión ENTERPRISE - v4.0 - El Sabueso de Abonos)
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
import openpyxl 
import time

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Motor de Conciliación Avanzado",
    page_icon="🕵️‍♂️",
    layout="wide"
)

# --- ESTILOS CSS ---
PALETA = {
    "azul_banco": "#003865",
    "gris_fondo": "#F0F2F6",
    "amarillo_alerta": "#FFC300"
}
st.markdown(f"""
<style>
    .stApp {{ background-color: {PALETA['gris_fondo']}; }}
    .stMetric {{ background-color: white; border-left: 5px solid {PALETA['azul_banco']}; }}
    div[data-testid="stExpander"] details summary {{ font-weight: bold; font-size: 1.1em; }}
</style>
""", unsafe_allow_html=True)

# ======================================================================================
# --- 1. UTILIDADES Y CONEXIONES (INFRAESTRUCTURA) ---
# ======================================================================================

@st.cache_data
def df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Conciliacion_Master')
    return output.getvalue()

def normalizar_texto_avanzado(texto):
    """Limpieza profunda para el motor de inteligencia (Sabueso)."""
    if not isinstance(texto, str): return ""
    texto = texto.upper().strip()
    # Normalización unicode
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    # Dejamos solo letras y números (quitamos puntuación que estorba al regex)
    texto = re.sub(r'[^A-Z0-9\s]', ' ', texto) 
    # Palabras irrelevantes para el match de nombre (Stop Words)
    palabras_basura = [
        'PAGO', 'TRANSF', 'TRANSFERENCIA', 'CONSIGNACION', 'ABONO', 'CTA', 'AHORROS', 
        'CORRIENTE', 'SUC', 'VIRTUAL', 'APP', 'NEQUI', 'DAVIPLATA', 'ACH', 'PSE', 'NIT', 
        'REF', 'FACTURA', 'VALOR', 'SALDO'
    ]
    for p in palabras_basura:
        texto = re.sub(r'\b' + p + r'\b', '', texto)
    return ' '.join(texto.split())

def extraer_posibles_nits(texto):
    """Busca secuencias numéricas que parecen NITs (7 a 11 dígitos)."""
    if not isinstance(texto, str): return []
    # Regex: \b indica límite de palabra, \d{7,11} busca entre 7 y 11 dígitos seguidos
    return re.findall(r'\b\d{7,11}\b', texto)

# --- CONEXIONES ---

@st.cache_resource
def get_dbx_client(secrets_key):
    try:
        creds = st.secrets[secrets_key]
        return dropbox.Dropbox(
            app_key=creds["app_key"],
            app_secret=creds["app_secret"],
            oauth2_refresh_token=creds["refresh_token"]
        )
    except Exception as e:
        st.error(f"Error Dropbox ({secrets_key}): {e}")
        return None

@st.cache_resource
def connect_to_google_sheets():
    try:
        scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_service_account"]), scope)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Error G-Sheets: {e}")
        return None

def download_from_dropbox(dbx, path):
    try:
        _, res = dbx.files_download(path)
        return res.content
    except Exception as e:
        st.error(f"Error descargando {path}: {e}")
        return None

def get_worksheet(client, url, tab_name):
    try:
        return client.open_by_url(url).worksheet(tab_name)
    except Exception as e:
        st.error(f"No se encontró la pestaña '{tab_name}' en el Sheet. {e}")
        return None

# ======================================================================================
# --- 2. CARGA DE DATOS (ETL - LECTURA INTELIGENTE) ---
# ======================================================================================

@st.cache_data(ttl=600)
def cargar_cartera():
    """Carga cartera y calcula deuda total por cliente."""
    dbx = get_dbx_client("dropbox")
    content = download_from_dropbox(dbx, '/data/cartera_detalle.csv')
    if not content: return pd.DataFrame()

    try:
        df = pd.read_csv(StringIO(content.decode('latin-1')), sep='|', header=None)
        # Asignar nombres según estructura conocida
        if len(df.columns) >= 17:
            df = df.iloc[:, :18] # Asegurar columnas
            df.columns = [
                'Serie', 'Numero', 'FechaDoc', 'FechaVenc', 'CodCliente', 'NombreCliente', 
                'Nit', 'Poblacion', 'Provincia', 'Tel1', 'Tel2', 'Vendedor', 
                'Autoriza', 'Email', 'Importe', 'Descuento', 'Cupo', 'DiasVenc'
            ]
        
        # Limpieza
        df['Importe'] = pd.to_numeric(df['Importe'], errors='coerce').fillna(0)
        df['Numero'] = pd.to_numeric(df['Numero'], errors='coerce').fillna(0)
        # Ajuste signos negativos
        df.loc[df['Numero'] < 0, 'Importe'] *= -1
        
        # NIT Normalizado (Clave para cruce)
        df['nit_norm'] = df['Nit'].astype(str).str.replace(r'[^0-9]', '', regex=True)
        # Nombre Normalizado
        df['nombre_norm'] = df['NombreCliente'].apply(normalizar_texto_avanzado)
        
        # Filtrar solo deuda positiva
        df = df[df['Importe'] > 0].copy()
        
        return df
    except Exception as e:
        st.error(f"Error procesando cartera: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=600)
def cargar_bancos_blue_headers(path_archivo):
    """
    Carga el archivo con estructura de IMAGEN 3/4 (Encabezados Azules).
    Busca columnas clave: FECHA, CUENTA, EMPRESA, VALOR.
    """
    dbx = get_dbx_client("dropbox")
    content = download_from_dropbox(dbx, path_archivo)
    if not content: return pd.DataFrame()

    try:
        # Intentamos leer. A veces los headers no están en la fila 0.
        # Leemos primeras filas para detectar donde empiezan los headers azules
        df_preview = pd.read_excel(BytesIO(content), nrows=10, header=None)
        
        header_row_idx = 0
        for idx, row in df_preview.iterrows():
            row_str = row.astype(str).str.upper().values
            if 'FECHA' in row_str and 'VALOR' in row_str:
                header_row_idx = idx
                break
        
        # Leemos de nuevo con el header correcto
        df = pd.read_excel(BytesIO(content), header=header_row_idx)
        
        # Normalizar nombres de columnas (strip y upper)
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # Validar columnas críticas
        cols_requeridas = ['FECHA', 'VALOR']
        if not all(col in df.columns for col in cols_requeridas):
            st.error(f"El archivo no tiene columnas FECHA y VALOR. Columnas detectadas: {list(df.columns)}")
            return pd.DataFrame()

        # Limpieza de datos
        df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
        df['VALOR'] = pd.to_numeric(df['VALOR'], errors='coerce').fillna(0)
        
        # --- CREACIÓN DE LA SÚPER COLUMNA DE RASTREO (SABUESO) ---
        # Concatenamos todas las columnas de texto que puedan tener pistas (Imagen 4 muestra CUENTA y EMPRESA)
        cols_pistas = ['SUCURSAL BANCO', 'TIPO DE TRANSACCION', 'CUENTA', 'EMPRESA', 'DESTINO', 'BANCO REFRENCIA INTERNA']
        
        df['texto_completo_pistas'] = ''
        for col in cols_pistas:
            if col in df.columns:
                df['texto_completo_pistas'] += df[col].fillna('').astype(str) + ' '
        
        # Normalizamos para búsqueda
        df['texto_analisis'] = df['texto_completo_pistas'].apply(normalizar_texto_avanzado)
        
        # ID Único para trazabilidad
        df['id_unico'] = df.apply(lambda x: f"MOV-{x['FECHA'].strftime('%Y%m%d')}-{int(x['VALOR'])}-{x.name}", axis=1)
        
        # Inicializar estados
        df['Estado'] = 'Pendiente'
        df['Cliente_Identificado'] = ''
        df['NIT_Encontrado'] = ''
        df['Tipo_Hallazgo'] = ''
        
        return df

    except Exception as e:
        st.error(f"Error crítico leyendo archivo de bancos: {e}")
        return pd.DataFrame()

# ======================================================================================
# --- 3. EL CEREBRO DEL SABUESO (LOGICA DE CONCILIACIÓN) ---
# ======================================================================================

def motor_sabueso_conciliacion(df_bancos, df_cartera, df_kb):
    """
    Núcleo de inteligencia artificial para encontrar pagadores.
    """
    st.write("🕵️‍♂️ **Iniciando el Sabueso... Escaneando descripciones, cuentas y valores...**")
    
    # 1. PREPARACIÓN DE MAPAS (Para velocidad extrema)
    # Mapa: NIT -> Nombre Cliente
    mapa_nit_cliente = df_cartera.groupby('nit_norm')['NombreCliente'].first().to_dict()
    # Mapa: NIT -> Deuda Total (Suma de Importes)
    mapa_deuda_total = df_cartera.groupby('nit_norm')['Importe'].sum().to_dict()
    # Lista de nombres para búsqueda difusa
    lista_nombres_unicos = df_cartera['nombre_norm'].unique().tolist()
    
    # Mapa: Memoria del Robot (Base de Conocimiento G-Sheets)
    memoria_robot = {}
    if not df_kb.empty:
        # Asumimos KB tiene: 'texto_clave', 'nit_asociado'
        for _, row in df_kb.iterrows():
            k = str(row.get('texto_banco_norm', '')).strip()
            v = str(row.get('nit_cliente', '')).strip()
            if k and v: memoria_robot[k] = v

    resultados = []
    
    # Barra de progreso visual
    progreso = st.progress(0)
    total_filas = len(df_bancos)

    # 2. ITERACIÓN (Fila por fila del banco)
    for idx, row in df_bancos.iterrows():
        progreso.progress((idx + 1) / total_filas)
        
        # Datos actuales
        texto_analisis = row['texto_analisis']
        valor_pago = row['VALOR']
        posibles_nits = extraer_posibles_nits(row['texto_completo_pistas'])
        
        info_match = {
            'Estado': 'Pendiente',
            'Cliente_Identificado': '',
            'NIT_Encontrado': '',
            'Tipo_Hallazgo': '',
            'Diferencia_Valor': 0
        }
        
        match_found = False

        # --- NIVEL 0: MEMORIA DEL ROBOT (APRENDIZAJE PREVIO) ---
        # Buscamos si alguna palabra clave o cuenta ya está en memoria
        # Esto cubre casos como "CUENTA 91200..." -> "VICTORIA PLAZA"
        for clave_memoria in memoria_robot:
            if clave_memoria in texto_analisis:
                nit_mem = memoria_robot[clave_memoria]
                if nit_mem in mapa_nit_cliente:
                    info_match['Cliente_Identificado'] = mapa_nit_cliente[nit_mem]
                    info_match['NIT_Encontrado'] = nit_mem
                    info_match['Tipo_Hallazgo'] = "0. Memoria (Aprendizaje)"
                    match_found = True
                    break
        
        # --- NIVEL 1: BÚSQUEDA DE NITS OCULTOS (REGEX) ---
        if not match_found:
            for nit_candidato in posibles_nits:
                if nit_candidato in mapa_nit_cliente:
                    info_match['Cliente_Identificado'] = mapa_nit_cliente[nit_candidato]
                    info_match['NIT_Encontrado'] = nit_candidato
                    info_match['Tipo_Hallazgo'] = f"1. NIT Detectado ({nit_candidato})"
                    match_found = True
                    break
        
        # --- NIVEL 2: BÚSQUEDA POR NOMBRE (FUZZY) ---
        # Solo si el texto es lo suficientemente largo
        if not match_found and len(texto_analisis) > 5:
            match_nombre, score = process.extractOne(texto_analisis, lista_nombres_unicos, scorer=fuzz.partial_ratio)
            if score >= 85: # Umbral alto para seguridad
                # Recuperar NIT del nombre
                nit_fuzzy = df_cartera[df_cartera['nombre_norm'] == match_nombre]['nit_norm'].iloc[0]
                nombre_real = mapa_nit_cliente.get(nit_fuzzy, match_nombre)
                
                info_match['Cliente_Identificado'] = nombre_real
                info_match['NIT_Encontrado'] = nit_fuzzy
                info_match['Tipo_Hallazgo'] = f"2. Nombre Similar ({score}%)"
                match_found = True

        # --- NIVEL 3: ANÁLISIS FINANCIERO (ABONO VS PAGO TOTAL) ---
        if match_found:
            nit_final = info_match['NIT_Encontrado']
            deuda_cliente = mapa_deuda_total.get(nit_final, 0)
            diff = deuda_cliente - valor_pago
            info_match['Diferencia_Valor'] = diff
            
            # Lógica de Negocio:
            if abs(diff) < 2000: # Margen de error pequeño ($2000 pesos)
                info_match['Estado'] = "CONCILIADO - PAGO TOTAL"
            elif valor_pago < deuda_cliente:
                info_match['Estado'] = "CONCILIADO - ABONO A CARTERA"
            elif valor_pago > deuda_cliente:
                info_match['Estado'] = "REVISAR - PAGO MAYOR A DEUDA"
            else:
                info_match['Estado'] = "IDENTIFICADO"
        
        # Guardamos resultado
        row_res = row.to_dict()
        row_res.update(info_match)
        resultados.append(row_res)

    return pd.DataFrame(resultados)

# ======================================================================================
# --- 4. INTERFAZ GRÁFICA (FRONTEND STREAMLIT) ---
# ======================================================================================

def main_app():
    st.title("🏦 Sistema Maestro de Conciliación Bancaria")
    st.markdown("""
    Este sistema utiliza **Inteligencia de Datos** para leer extractos bancarios crudos, 
    buscar NITs escondidos en columnas de texto, y cruzar con cartera para identificar 
    Pagos Totales y Abonos.
    """)
    
    # Tabs principales
    tab_admin, tab_usuario, tab_resultados = st.tabs(["⚙️ Panel Admin (Batch)", "👤 Asignación Manual", "📊 Reportes"])
    
    # --- TAB 1: ADMIN (BATCH RUN) ---
    with tab_admin:
        st.header("Ejecución del Motor (Batch)")
        st.info("Este proceso lee 'planilla_bancos.xlsx' (Encabezados Azules) y 'cartera_detalle.csv' de Dropbox.")
        
        if st.button("🚀 INICIAR CONCILIACIÓN AUTOMÁTICA", type="primary"):
            
            # 1. Cargar Datos
            with st.spinner("Leyendo Cartera y Extractos Bancarios..."):
                df_cartera = cargar_cartera()
                path_bancos = st.secrets["dropbox"]["path_bancos"] # Asegúrate que esto apunte al archivo azul
                df_bancos = cargar_bancos_blue_headers(path_bancos)
            
            if df_cartera.empty or df_bancos.empty:
                st.error("Fallo en la carga de archivos. Revisa logs.")
                st.stop()
                
            st.success(f"Datos cargados: {len(df_bancos)} movimientos bancarios, {len(df_cartera)} facturas en cartera.")

            # 2. Cargar Conocimiento (G-Sheets)
            g_client = connect_to_google_sheets()
            ws_kb = get_worksheet(g_client, st.secrets["google_sheets"]["sheet_url"], st.secrets["google_sheets"]["tab_knowledge_base"])
            df_kb = pd.DataFrame(ws_kb.get_all_records()) if ws_kb else pd.DataFrame()

            # 3. Ejecutar Sabueso
            df_resultado = motor_sabueso_conciliacion(df_bancos, df_cartera, df_kb)
            
            # 4. Guardar en G-Sheet Master
            with st.spinner("Guardando resultados en la Nube (Bancos_Master)..."):
                ws_master = get_worksheet(g_client, st.secrets["google_sheets"]["sheet_url"], st.secrets["google_sheets"]["tab_bancos_master"])
                if ws_master:
                    ws_master.clear()
                    # Convertir fechas a string para GSheets
                    df_save = df_resultado.copy()
                    for c in df_save.select_dtypes(include=['datetime']):
                        df_save[c] = df_save[c].astype(str)
                    df_save = df_save.fillna('')
                    set_with_dataframe(ws_master, df_save)
                    st.balloons()
                    st.success("¡Conciliación Finalizada y Guardada!")

    # --- TAB 2: USUARIO (MANUAL) ---
    with tab_usuario:
        st.header("Asignación de Pendientes")
        
        if st.button("🔄 Cargar Pendientes"):
            g_client = connect_to_google_sheets()
            ws_master = get_worksheet(g_client, st.secrets["google_sheets"]["sheet_url"], st.secrets["google_sheets"]["tab_bancos_master"])
            if ws_master:
                df_master = pd.DataFrame(ws_master.get_all_records())
                # Filtrar pendientes
                df_pendientes = df_master[df_master['Estado'] == 'Pendiente']
                st.session_state['pendientes'] = df_pendientes
                st.session_state['data_loaded'] = True
        
        if st.session_state.get('data_loaded'):
            df_p = st.session_state['pendientes']
            st.metric("Movimientos Pendientes", len(df_p))
            
            if not df_p.empty:
                pago_sel = st.selectbox("Seleccionar Movimiento", df_p['descripcion_banco'] + " - $" + df_p['VALOR'].astype(str))
                # (Aquí iría la lógica de asignación manual con escritura en KB - igual que en v3)
                st.info("Funcionalidad de asignación manual lista para conectar.")

    # --- TAB 3: REPORTES (TU IMAGEN 2) ---
    with tab_resultados:
        st.header("Planilla Bancos Consolidada")
        st.write("Descarga el archivo idéntico a tu formato deseado (Imagen 2).")
        
        # Botón para descargar lo que hay en G-Sheet Master formateado
        if st.button("📥 Generar Excel Consolidado"):
            g_client = connect_to_google_sheets()
            ws_master = get_worksheet(g_client, st.secrets["google_sheets"]["sheet_url"], st.secrets["google_sheets"]["tab_bancos_master"])
            if ws_master:
                df_final = pd.DataFrame(ws_master.get_all_records())
                
                # Filtrar y ordenar columnas para que se vea como Imagen 2
                cols_deseadas = ['FECHA', 'Cliente_Identificado', 'Tipo_Hallazgo', 'Estado', 'VALOR', 'texto_completo_pistas']
                # Asegurar que existan
                cols_existentes = [c for c in cols_deseadas if c in df_final.columns]
                
                excel_bytes = df_to_excel(df_final[cols_existentes])
                
                st.download_button(
                    label="Descargar Planilla Pereira.xlsx",
                    data=excel_bytes,
                    file_name="Planilla_Bancos_Pereira_Consolidada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == '__main__':
    main_app()
