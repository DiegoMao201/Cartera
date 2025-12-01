# ======================================================================================
# ARCHIVO: pages/2_Motor_Conciliacion.py
# (Versión v5.0 - Carga Manual de Planilla Pereira + Corrección NaT)
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
import openpyxl 

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Motor Conciliación Pereira", page_icon="🏦", layout="wide")

# ======================================================================================
# --- 1. CONEXIONES Y UTILIDADES ---
# ======================================================================================

@st.cache_resource
def get_dbx_client(secrets_key):
    """Conexión a Dropbox solo para CARTERA (que dijiste que sí está ahí)."""
    try:
        if secrets_key not in st.secrets: return None
        creds = st.secrets[secrets_key]
        return dropbox.Dropbox(
            app_key=creds["app_key"],
            app_secret=creds["app_secret"],
            oauth2_refresh_token=creds["refresh_token"]
        )
    except: return None

@st.cache_resource
def connect_to_google_sheets():
    """Conexión a G-Sheets para guardar el resultado maestro."""
    try:
        scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_service_account"]), scope)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Error conectando a Google Sheets: {e}")
        return None

def download_from_dropbox(dbx, path):
    try:
        _, res = dbx.files_download(path)
        return res.content
    except Exception as e:
        st.error(f"Error descargando {path} de Dropbox: {e}")
        return None

def df_to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Conciliacion')
    return output.getvalue()

def normalizar_texto_avanzado(texto):
    if not isinstance(texto, str): return ""
    texto = texto.upper().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'[^A-Z0-9\s]', ' ', texto) 
    palabras_basura = ['PAGO', 'TRANSF', 'TRANSFERENCIA', 'CONSIGNACION', 'ABONO', 'CTA', 'NIT', 'REF', 'FACTURA']
    for p in palabras_basura:
        texto = re.sub(r'\b' + p + r'\b', '', texto)
    return ' '.join(texto.split())

def extraer_posibles_nits(texto):
    if not isinstance(texto, str): return []
    return re.findall(r'\b\d{7,11}\b', texto)

# ======================================================================================
# --- 2. CARGA DE ARCHIVOS (AQUÍ ESTÁ LA MAGIA) ---
# ======================================================================================

@st.cache_data(ttl=600)
def cargar_cartera():
    """Carga la cartera desde Dropbox (Asumiendo que esta sí está configurada)."""
    dbx = get_dbx_client("dropbox")
    if not dbx: 
        st.warning("⚠️ No hay conexión a Dropbox configurada para Cartera.")
        return pd.DataFrame()
        
    content = download_from_dropbox(dbx, '/data/cartera_detalle.csv')
    if not content: return pd.DataFrame()

    try:
        df = pd.read_csv(StringIO(content.decode('latin-1')), sep='|', header=None)
        # Ajustamos a las primeras 18 columnas estándar
        df = df.iloc[:, :18]
        df.columns = [
            'Serie', 'Numero', 'FechaDoc', 'FechaVenc', 'CodCliente', 'NombreCliente', 
            'Nit', 'Poblacion', 'Provincia', 'Tel1', 'Tel2', 'Vendedor', 
            'Autoriza', 'Email', 'Importe', 'Descuento', 'Cupo', 'DiasVenc'
        ]
        
        df['Importe'] = pd.to_numeric(df['Importe'], errors='coerce').fillna(0)
        df['Numero'] = pd.to_numeric(df['Numero'], errors='coerce').fillna(0)
        df.loc[df['Numero'] < 0, 'Importe'] *= -1
        
        df['nit_norm'] = df['Nit'].astype(str).str.replace(r'[^0-9]', '', regex=True)
        df['nombre_norm'] = df['NombreCliente'].apply(normalizar_texto_avanzado)
        
        return df[df['Importe'] > 0].copy()
    except Exception as e:
        st.error(f"Error leyendo cartera: {e}")
        return pd.DataFrame()

def cargar_planilla_pereira_desde_upload(uploaded_file):
    """
    Lee el archivo 'PLANILLA BANCOS PEREIRA' que subes manualmente.
    Maneja el error NaT y busca encabezados azules.
    """
    try:
        # 1. Detectar encabezados automáticamente
        # Leemos las primeras 10 filas para buscar dónde dice "FECHA" y "VALOR"
        df_preview = pd.read_excel(uploaded_file, nrows=10, header=None)
        header_idx = 0
        found = False
        
        for idx, row in df_preview.iterrows():
            row_vals = row.astype(str).str.upper().values
            if 'FECHA' in row_vals and 'VALOR' in row_vals:
                header_idx = idx
                found = True
                break
        
        if not found:
            st.error("❌ No encontré las columnas 'FECHA' y 'VALOR' en las primeras 10 filas.")
            return pd.DataFrame()

        # 2. Leer archivo real desde la fila detectada
        # Importante: Volvemos al inicio del archivo
        uploaded_file.seek(0) 
        df = pd.read_excel(uploaded_file, header=header_idx)
        
        # Normalizar columnas
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # 3. CORRECCIÓN ERROR NaT (CRÍTICO)
        # Convertimos fecha, forzando errores a NaT (Not a Time)
        df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
        # Convertimos valor
        df['VALOR'] = pd.to_numeric(df['VALOR'], errors='coerce').fillna(0)
        
        # Eliminamos filas que sean basura (donde no hay fecha válida)
        df = df.dropna(subset=['FECHA'])
        
        # 4. Crear columna de análisis (SABUESO)
        cols_clave = ['SUCURSAL BANCO', 'TIPO DE TRANSACCION', 'CUENTA', 'EMPRESA', 'DESTINO']
        df['texto_analisis'] = ''
        for col in cols_clave:
            if col in df.columns:
                df['texto_analisis'] += df[col].fillna('').astype(str) + ' '
                
        df['texto_norm'] = df['texto_analisis'].apply(normalizar_texto_avanzado)

        # 5. Crear ID Seguro (Corrección del error strftime)
        def safe_id(row):
            try:
                # Si la fecha es NaT, ponemos un placeholder
                f_str = row['FECHA'].strftime('%Y%m%d') if pd.notnull(row['FECHA']) else "00000000"
                return f"MOV-{f_str}-{int(row['VALOR'])}-{row.name}"
            except:
                return f"MOV-ERR-{row.name}"

        df['id_unico'] = df.apply(safe_id, axis=1)
        
        # Inicializar columnas de resultado
        df['Estado'] = 'Pendiente'
        df['Cliente_Identificado'] = ''
        df['NIT_Encontrado'] = ''
        df['Tipo_Hallazgo'] = ''
        
        return df

    except Exception as e:
        st.error(f"Error leyendo el archivo Excel: {e}")
        return pd.DataFrame()

# ======================================================================================
# --- 3. MOTOR INTELIGENTE (SABUESO) ---
# ======================================================================================

def ejecutar_motor(df_bancos, df_cartera, df_kb):
    st.info("🏃‍♂️ Ejecutando motor de identificación...")
    
    # Preparar mapas de búsqueda rápida
    mapa_nit_nombre = df_cartera.groupby('nit_norm')['NombreCliente'].first().to_dict()
    mapa_nit_deuda = df_cartera.groupby('nit_norm')['Importe'].sum().to_dict()
    lista_nombres = df_cartera['nombre_norm'].unique().tolist()
    
    # Memoria KB
    memoria = {}
    if not df_kb.empty:
        for _, row in df_kb.iterrows():
            k = str(row.get('texto_banco_norm','')).strip()
            v = str(row.get('nit_cliente','')).strip()
            if k and v: memoria[k] = v

    resultados = []
    
    bar = st.progress(0)
    total = len(df_bancos)
    
    for i, row in df_bancos.iterrows():
        bar.progress((i+1)/total)
        
        res = {
            'Estado': 'Pendiente', 
            'Cliente_Identificado': '', 
            'NIT_Encontrado': '',
            'Tipo_Hallazgo': '',
            'Diferencia_Deuda': 0
        }
        
        txt = row['texto_norm']
        val = row['VALOR']
        match = False
        
        # 1. Memoria
        for k_mem in memoria:
            if k_mem in txt:
                nit = memoria[k_mem]
                if nit in mapa_nit_nombre:
                    res['NIT_Encontrado'] = nit
                    res['Cliente_Identificado'] = mapa_nit_nombre[nit]
                    res['Tipo_Hallazgo'] = '0. Memoria'
                    match = True
                    break
        
        # 2. NIT en Texto
        if not match:
            nits_txt = extraer_posibles_nits(row['texto_analisis'])
            for n in nits_txt:
                if n in mapa_nit_nombre:
                    res['NIT_Encontrado'] = n
                    res['Cliente_Identificado'] = mapa_nit_nombre[n]
                    res['Tipo_Hallazgo'] = f'1. NIT Detectado ({n})'
                    match = True
                    break
        
        # 3. Nombre Fuzzy
        if not match and len(txt) > 4:
            found_name, score = process.extractOne(txt, lista_nombres, scorer=fuzz.partial_ratio)
            if score >= 85:
                nit = df_cartera[df_cartera['nombre_norm'] == found_name]['nit_norm'].iloc[0]
                res['NIT_Encontrado'] = nit
                res['Cliente_Identificado'] = mapa_nit_nombre.get(nit, found_name)
                res['Tipo_Hallazgo'] = f'2. Nombre Similar ({score}%)'
                match = True
        
        # 4. Clasificación Financiera
        if match:
            nit = res['NIT_Encontrado']
            deuda = mapa_nit_deuda.get(nit, 0)
            diff = deuda - val
            res['Diferencia_Deuda'] = diff
            
            if abs(diff) < 2000:
                res['Estado'] = 'CONCILIADO - PAGO TOTAL'
            elif val < deuda:
                res['Estado'] = 'CONCILIADO - ABONO'
            else:
                res['Estado'] = 'REVISAR - PAGO MAYOR A DEUDA'
        
        # Unir datos
        full_row = row.to_dict()
        full_row.update(res)
        resultados.append(full_row)
        
    return pd.DataFrame(resultados)

# ======================================================================================
# --- 4. INTERFAZ PRINCIPAL ---
# ======================================================================================

def main():
    st.title("🏦 Conciliación Bancaria - Planilla Pereira")
    
    # PESTAÑA ÚNICA Y CLARA
    st.markdown("### Paso 1: Carga de Archivos")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("1. Archivo del Banco (Crudo)")
        uploaded_file = st.file_uploader("Arrastra aquí 'PLANILLA BANCOS PEREIRA.xlsx'", type=["xlsx"])
    
    with col2:
        st.subheader("2. Cartera (Desde Dropbox)")
        if st.button("🔄 Cargar Cartera"):
            with st.spinner("Descargando cartera..."):
                df_c = cargar_cartera()
                if not df_c.empty:
                    st.session_state['df_cartera'] = df_c
                    st.success(f"Cartera cargada: {len(df_c)} facturas.")
                else:
                    st.error("No se pudo cargar la cartera.")
    
    # VALIDACIÓN DE ESTADO
    if uploaded_file and 'df_cartera' in st.session_state:
        st.divider()
        st.markdown("### Paso 2: Ejecución")
        
        if st.button("🚀 EJECUTAR MOTOR (SABUESO)", type="primary"):
            # 1. Leer Banco
            df_bancos = cargar_planilla_pereira_desde_upload(uploaded_file)
            
            if not df_bancos.empty:
                st.info(f"Leídos {len(df_bancos)} movimientos del banco.")
                
                # 2. Leer KB (Opcional)
                g_client = connect_to_google_sheets()
                df_kb = pd.DataFrame()
                if g_client:
                    try:
                        ws = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"]).worksheet(st.secrets["google_sheets"]["tab_knowledge_base"])
                        df_kb = pd.DataFrame(ws.get_all_records())
                    except: pass
                
                # 3. Motor
                df_resultado = ejecutar_motor(df_bancos, st.session_state['df_cartera'], df_kb)
                
                # 4. Mostrar Resultados
                st.success("¡Conciliación Terminada!")
                st.dataframe(df_resultado[['FECHA', 'VALOR', 'texto_analisis', 'Cliente_Identificado', 'Estado', 'Tipo_Hallazgo']])
                
                # 5. Guardar Master
                if g_client:
                    try:
                        ws_master = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"]).worksheet(st.secrets["google_sheets"]["tab_bancos_master"])
                        ws_master.clear()
                        # Limpieza para GSheets
                        df_save = df_resultado.copy()
                        for c in df_save.select_dtypes(['datetime']): df_save[c] = df_save[c].astype(str)
                        df_save = df_save.fillna('')
                        set_with_dataframe(ws_master, df_save)
                        st.toast("Guardado en Google Sheets 'Bancos_Master'")
                    except Exception as e:
                        st.warning(f"No se pudo guardar en GSheets: {e}")
                
                # 6. Descargar
                excel = df_to_excel(df_resultado)
                st.download_button("📥 Descargar Planilla Consolidada", excel, "Planilla_Pereira_Identificada.xlsx")

if __name__ == "__main__":
    main()
