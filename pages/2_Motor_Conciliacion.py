# ======================================================================================
# ARCHIVO: pages/2_Motor_Conciliacion.py
# (Versión v6.0 - Super Motor de Inteligencia + Excel de Lujo para Tesorería)
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
import xlsxwriter # Necesario para el Excel bonito

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Motor Conciliación Pereira", page_icon="🧠", layout="wide")

# ======================================================================================
# --- 1. CONEXIONES Y UTILIDADES ---
# ======================================================================================

@st.cache_resource
def get_dbx_client(secrets_key):
    """Conexión a Dropbox para descargar la Cartera actualizada."""
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
    """Conexión a G-Sheets (Cerebro del sistema)."""
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

def normalizar_texto_avanzado(texto):
    """Limpieza profunda para que el robot entienda mejor."""
    if not isinstance(texto, str): return ""
    texto = texto.upper().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'[^A-Z0-9\s]', ' ', texto) 
    
    # Palabras que ensucian el análisis (Stopwords bancarios)
    palabras_basura = [
        'PAGO', 'TRANSF', 'TRANSFERENCIA', 'CONSIGNACION', 'ABONO', 'CTA', 'NIT', 
        'REF', 'FACTURA', 'OFI', 'SUC', 'ACH', 'PSE', 'NOMINA', 'PROVEEDOR', 
        'COMPRA', 'VENTA', 'VALOR', 'NETO'
    ]
    for p in palabras_basura:
        texto = re.sub(r'\b' + p + r'\b', '', texto)
    
    return ' '.join(texto.split())

def extraer_posibles_nits(texto):
    """Busca cualquier secuencia numérica que parezca un NIT (7 a 11 dígitos)."""
    if not isinstance(texto, str): return []
    return re.findall(r'\b\d{7,11}\b', texto)

# ======================================================================================
# --- 2. EXCEL DE LUJO (PARA LA LÍDER DE TESORERÍA) ---
# ======================================================================================

def generar_excel_bonito(df):
    """
    Genera un Excel con formato profesional, colores y filtros.
    Distingue visualmente lo conciliado de lo pendiente.
    """
    output = BytesIO()
    # Usamos XlsxWriter como motor
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Conciliacion_Pereira')
        workbook = writer.book
        worksheet = writer.sheets['Conciliacion_Pereira']
        
        # --- DEFINICIÓN DE FORMATOS ---
        
        # Formato Encabezados
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#1F497D', # Azul oscuro corporativo
            'font_color': 'white',
            'border': 1
        })
        
        # Formato Moneda
        money_format = workbook.add_format({'num_format': '$ #,##0', 'border': 1})
        
        # Formato Fecha
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1})
        
        # Formato Celdas Normales
        cell_format = workbook.add_format({'border': 1})
        
        # --- FORMATOS CONDICIONALES (SEMÁFORO) ---
        
        # Verde: Conciliado Total
        green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1})
        # Amarillo: Abono o Parcial
        yellow_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'border': 1})
        # Rojo: Revisar / Error
        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1})

        # --- APLICAR FORMATOS ---
        
        # 1. Ajustar ancho de columnas y aplicar encabezado
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
            # Anchos estimados
            if "TEXTO" in str(value).upper(): width = 50
            elif "CLIENTE" in str(value).upper(): width = 40
            elif "FECHA" in str(value).upper(): width = 15
            elif "VALOR" in str(value).upper() or "DEUDA" in str(value).upper(): width = 18
            else: width = 20
            worksheet.set_column(col_num, col_num, width)

        # 2. Aplicar formatos a los datos
        # Detectar columnas clave
        col_estado_idx = df.columns.get_loc('Estado') if 'Estado' in df.columns else -1
        col_valor_idx = df.columns.get_loc('VALOR') if 'VALOR' in df.columns else -1
        col_fecha_idx = df.columns.get_loc('FECHA') if 'FECHA' in df.columns else -1
        
        nrow = len(df) + 1
        
        # Aplicar formato condicional a toda la fila basado en el Estado
        if col_estado_idx != -1:
            # Regla Verde (Conciliado)
            worksheet.conditional_format(1, 0, nrow, len(df.columns)-1, {
                'type': 'formula',
                'criteria': f'=$P{col_estado_idx+1}="CONCILIADO - PAGO TOTAL"', # La P es un ejemplo, se ajusta dinámicamente en excel
                # Usamos text string matching para simplificar
                'criteria': f'=SEARCH("TOTAL", ${chr(65+col_estado_idx)}2)',
                'format': green_format
            })
            # Regla Amarilla (Abono)
            worksheet.conditional_format(1, 0, nrow, len(df.columns)-1, {
                'type': 'formula',
                'criteria': f'=SEARCH("ABONO", ${chr(65+col_estado_idx)}2)',
                'format': yellow_format
            })
            # Regla Roja (Revisar)
            worksheet.conditional_format(1, 0, nrow, len(df.columns)-1, {
                'type': 'formula',
                'criteria': f'=SEARCH("REVISAR", ${chr(65+col_estado_idx)}2)',
                'format': red_format
            })

        # Aplicar formato de moneda a la columna VALOR
        if col_valor_idx != -1:
            worksheet.set_column(col_valor_idx, col_valor_idx, 18, money_format)

        # Aplicar formato de fecha
        if col_fecha_idx != -1:
            worksheet.set_column(col_fecha_idx, col_fecha_idx, 15, date_format)

        # Activar Autofiltros
        worksheet.autofilter(0, 0, nrow, len(df.columns)-1)

    return output.getvalue()

# ======================================================================================
# --- 3. CARGA DE ARCHIVOS ---
# ======================================================================================

@st.cache_data(ttl=600)
def cargar_cartera():
    """Carga la cartera, que es la 'Verdad' contra la que conciliamos."""
    dbx = get_dbx_client("dropbox")
    if not dbx: 
        st.warning("⚠️ Sin conexión a Dropbox.")
        return pd.DataFrame()
        
    content = download_from_dropbox(dbx, '/data/cartera_detalle.csv')
    if not content: return pd.DataFrame()

    try:
        df = pd.read_csv(StringIO(content.decode('latin-1')), sep='|', header=None)
        # Ajustamos a las columnas estándar
        df = df.iloc[:, :18]
        df.columns = [
            'Serie', 'Numero', 'FechaDoc', 'FechaVenc', 'CodCliente', 'NombreCliente', 
            'Nit', 'Poblacion', 'Provincia', 'Tel1', 'Tel2', 'Vendedor', 
            'Autoriza', 'Email', 'Importe', 'Descuento', 'Cupo', 'DiasVenc'
        ]
        
        df['Importe'] = pd.to_numeric(df['Importe'], errors='coerce').fillna(0)
        df['nit_norm'] = df['Nit'].astype(str).str.replace(r'[^0-9]', '', regex=True)
        df['nombre_norm'] = df['NombreCliente'].apply(normalizar_texto_avanzado)
        
        # Filtramos solo lo que tiene deuda positiva
        return df[df['Importe'] > 0].copy()
    except Exception as e:
        st.error(f"Error leyendo cartera: {e}")
        return pd.DataFrame()

def cargar_planilla_pereira_desde_upload(uploaded_file):
    """
    Lee el archivo 'PLANILLA BANCOS PEREIRA'.
    Inteligente: Busca dónde empieza la tabla real.
    """
    try:
        # 1. Detectar encabezados
        df_preview = pd.read_excel(uploaded_file, nrows=15, header=None)
        header_idx = 0
        found = False
        
        for idx, row in df_preview.iterrows():
            row_str = row.astype(str).str.upper().values
            if 'FECHA' in row_str and 'VALOR' in row_str:
                header_idx = idx
                found = True
                break
        
        if not found:
            st.error("❌ No encontré las columnas 'FECHA' y 'VALOR'. Revisa el archivo.")
            return pd.DataFrame()

        # 2. Leer datos reales
        uploaded_file.seek(0) 
        df = pd.read_excel(uploaded_file, header=header_idx)
        
        # Normalizar columnas
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # 3. Limpieza de Fechas y Valores
        df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
        df['VALOR'] = pd.to_numeric(df['VALOR'], errors='coerce').fillna(0)
        df = df.dropna(subset=['FECHA']) # Borrar filas sin fecha
        
        # 4. Crear columna 'SABUESO' (Texto completo para analizar)
        cols_clave = ['SUCURSAL BANCO', 'TIPO DE TRANSACCION', 'CUENTA', 'EMPRESA', 'DESTINO', 'DETALLE', 'NOTAS']
        df['texto_analisis'] = ''
        for col in cols_clave:
            if col in df.columns:
                df['texto_analisis'] += df[col].fillna('').astype(str) + ' '
                
        df['texto_norm'] = df['texto_analisis'].apply(normalizar_texto_avanzado)

        # 5. ID Único para rastreo
        def safe_id(row):
            f_str = row['FECHA'].strftime('%Y%m%d') if pd.notnull(row['FECHA']) else "0000"
            return f"MOV-{f_str}-{int(row['VALOR'])}-{row.name}"

        df['id_unico'] = df.apply(safe_id, axis=1)
        
        # Columnas de resultado vacías
        df['Estado'] = 'Pendiente'
        df['Cliente_Identificado'] = ''
        df['NIT_Encontrado'] = ''
        df['Tipo_Hallazgo'] = ''
        
        return df

    except Exception as e:
        st.error(f"Error procesando Excel: {e}")
        return pd.DataFrame()

# ======================================================================================
# --- 4. MOTOR INTELIGENTE v6.0 (CEREBRO) ---
# ======================================================================================

def ejecutar_motor_inteligente(df_bancos, df_cartera, df_kb):
    st.info("🧠 Cerebro activado: Cruzando Cartera + Base de Conocimiento + Algoritmos Fuzzy...")
    
    # 1. Preparar índices de velocidad (Hash Maps)
    # Mapa: NIT -> Nombre Cliente
    mapa_nit_nombre = df_cartera.groupby('nit_norm')['NombreCliente'].first().to_dict()
    # Mapa: NIT -> Deuda Total
    mapa_nit_deuda = df_cartera.groupby('nit_norm')['Importe'].sum().to_dict()
    # Lista de nombres para Fuzzy
    lista_nombres = df_cartera['nombre_norm'].unique().tolist()
    
    # 2. Cargar 'Memoria' de Google Sheets (Knowledge Base)
    # Formato esperado en KB: Col A (Texto Banco), Col B (NIT Real)
    memoria_inteligente = {}
    if not df_kb.empty:
        # Aseguramos nombres de columnas si la KB no tiene cabeceras perfectas
        cols_kb = df_kb.columns
        if len(cols_kb) >= 2:
            for _, row in df_kb.iterrows():
                try:
                    # Normalizamos la clave (texto banco) y guardamos el valor (NIT)
                    key = normalizar_texto_avanzado(str(row.iloc[0]))
                    val = str(row.iloc[1]).strip()
                    if key and val:
                        memoria_inteligente[key] = val
                except: pass

    resultados = []
    
    bar = st.progress(0)
    total_filas = len(df_bancos)
    
    for i, row in df_bancos.iterrows():
        bar.progress((i+1)/total_filas)
        
        # Datos del movimiento
        txt_banco = row['texto_norm']
        txt_crudo = row['texto_analisis']
        valor_banco = row['VALOR']
        
        res = {
            'Estado': 'PENDIENTE', 
            'Cliente_Identificado': 'NO IDENTIFICADO', 
            'NIT_Encontrado': '',
            'Tipo_Hallazgo': '',
            'Diferencia_Deuda': 0,
            'Deuda_Total_Cartera': 0
        }
        
        match_found = False
        nit_candidato = None

        # --- NIVEL 0: MEMORIA (Inteligencia Histórica) ---
        # Si este texto exacto ya lo clasificamos antes en la KB
        if not match_found:
            for k_mem in memoria_inteligente:
                if k_mem in txt_banco and len(k_mem) > 5: # Evitar matches cortos
                    nit_candidato = memoria_inteligente[k_mem]
                    res['Tipo_Hallazgo'] = '🧠 0. Memoria Histórica (KB)'
                    match_found = True
                    break

        # --- NIVEL 1: NIT EXACTO EN TEXTO ---
        if not match_found:
            posibles_nits = extraer_posibles_nits(txt_crudo)
            for pn in posibles_nits:
                if pn in mapa_nit_nombre:
                    nit_candidato = pn
                    res['Tipo_Hallazgo'] = f'🔍 1. NIT Detectado en Ref ({pn})'
                    match_found = True
                    break
        
        # --- NIVEL 2: FUZZY MATCH (Nombre parecido) ---
        if not match_found and len(txt_banco) > 5:
            # Usamos token_set_ratio que es mejor cuando las palabras están desordenadas
            match_name, score = process.extractOne(txt_banco, lista_nombres, scorer=fuzz.token_set_ratio)
            
            if score >= 88: # Umbral alto para seguridad
                # Buscamos el NIT de ese nombre
                nit_candidato = df_cartera[df_cartera['nombre_norm'] == match_name]['nit_norm'].iloc[0]
                res['Tipo_Hallazgo'] = f'🤖 2. IA Nombre Similar ({score}%)'
                match_found = True

        # --- EVALUACIÓN FINANCIERA ---
        if match_found and nit_candidato:
            nombre_real = mapa_nit_nombre.get(nit_candidato, "Nombre Desconocido")
            deuda_actual = mapa_nit_deuda.get(nit_candidato, 0)
            
            res['NIT_Encontrado'] = nit_candidato
            res['Cliente_Identificado'] = nombre_real
            res['Deuda_Total_Cartera'] = deuda_actual
            
            diferencia = deuda_actual - valor_banco
            res['Diferencia_Deuda'] = diferencia
            
            # Lógica de conciliación
            if abs(diferencia) < 2000: # Diferencia menor a 2000 pesos
                res['Estado'] = 'CONCILIADO - PAGO TOTAL'
            elif valor_banco < deuda_actual:
                res['Estado'] = 'CONCILIADO - ABONO PARCIAL'
            elif valor_banco > deuda_actual:
                res['Estado'] = 'REVISAR - PAGO MAYOR A DEUDA'
            else:
                res['Estado'] = 'REVISAR - CASO ATÍPICO'

        # Unimos datos originales con resultados
        row_dict = row.to_dict()
        row_dict.update(res)
        resultados.append(row_dict)
        
    return pd.DataFrame(resultados)

# ======================================================================================
# --- 5. INTERFAZ PRINCIPAL ---
# ======================================================================================

def main():
    st.title("🏦 Super Motor de Conciliación - Pereira")
    st.markdown("""
    Esta herramienta cruza **Planilla Bancaria** vs **Cartera (Dropbox)**.
    Aprende de la hoja `Knowledge_Base` en Google Sheets para mejorar cada mes.
    """)
    
    # --- PANTALLA DE CARGA ---
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("1. Archivo Banco (Excel)")
        uploaded_file = st.file_uploader("Sube la Planilla Pereira", type=["xlsx", "xls"])
    
    with col2:
        st.subheader("2. Cartera (Nube)")
        if st.button("☁️ Actualizar Cartera desde Dropbox", type="secondary"):
            with st.spinner("Conectando a Dropbox..."):
                df_c = cargar_cartera()
                if not df_c.empty:
                    st.session_state['df_cartera'] = df_c
                    st.success(f"✅ Cartera cargada: {len(df_c)} facturas pendientes.")
                else:
                    st.error("❌ Falló la carga de cartera.")

    # Verificar si tenemos cartera en memoria
    if 'df_cartera' in st.session_state:
        st.caption(f"Cartera activa: {len(st.session_state['df_cartera'])} registros.")
    else:
        st.warning("⚠️ Primero carga la cartera (Botón en columna derecha).")

    st.divider()

    # --- EJECUCIÓN DEL MOTOR ---
    if uploaded_file and 'df_cartera' in st.session_state:
        if st.button("🚀 EJECUTAR SUPER MOTOR", type="primary", use_container_width=True):
            
            # 1. Cargar Banco
            with st.status("Procesando datos...", expanded=True) as status:
                st.write("📂 Leyendo archivo del banco...")
                df_bancos = cargar_planilla_pereira_desde_upload(uploaded_file)
                
                if df_bancos.empty:
                    st.error("El archivo del banco parece vacío o inválido.")
                    status.update(label="Error", state="error")
                    return

                # 2. Conectar a Google Sheets para LEER INTELIGENCIA (KB)
                st.write("🧠 Consultando Base de Conocimiento (GSheets)...")
                g_client = connect_to_google_sheets()
                df_kb = pd.DataFrame()
                
                if g_client:
                    try:
                        # Intentamos leer la hoja 'Knowledge_Base'
                        # Si no existe, no pasa nada, el motor funciona sin ella pero no aprende
                        sh = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"])
                        try:
                            ws_kb = sh.worksheet("Knowledge_Base") # Nombre exacto de la pestaña
                            df_kb = pd.DataFrame(ws_kb.get_all_records())
                            st.write(f"✅ Conocimiento cargado: {len(df_kb)} reglas aprendidas.")
                        except:
                            st.warning("⚠️ No encontré la hoja 'Knowledge_Base'. El motor usará inteligencia estándar.")
                    except Exception as e:
                        st.error(f"Error GSheets: {e}")

                # 3. Correr Algoritmo
                st.write("🤖 Ejecutando comparaciones Fuzzy + NIT + Memoria...")
                df_resultado = ejecutar_motor_inteligente(df_bancos, st.session_state['df_cartera'], df_kb)
                
                status.update(label="¡Proceso Completado!", state="complete", expanded=False)

            # --- RESULTADOS ---
            st.success(f"Procesados {len(df_resultado)} movimientos.")
            
            # Métricas rápidas
            conciliados = len(df_resultado[df_resultado['Estado'].str.contains("CONCILIADO")])
            pendientes = len(df_resultado) - conciliados
            col_m1, col_m2 = st.columns(2)
            col_m1.metric("✅ Conciliados Automáticamente", conciliados)
            col_m2.metric("⚠️ Pendientes de Revisión", pendientes)

            # Vista Previa
            st.dataframe(
                df_resultado[['FECHA', 'VALOR', 'texto_analisis', 'Cliente_Identificado', 'Estado', 'Tipo_Hallazgo']],
                use_container_width=True
            )

            # --- EXPORTACIÓN Y GUARDADO ---
            col_d1, col_d2 = st.columns(2)
            
            with col_d1:
                # Generar Excel Bonito
                excel_data = generar_excel_bonito(df_resultado)
                st.download_button(
                    label="📥 Descargar Excel Formateado (Para Tesorería)",
                    data=excel_data,
                    file_name=f"Conciliacion_Pereira_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            with col_d2:
                # Guardar en GSheets Master (Sobrescribir vista actual)
                if g_client:
                    try:
                        sh = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"])
                        ws_master = sh.worksheet(st.secrets["google_sheets"]["tab_bancos_master"])
                        ws_master.clear()
                        
                        # Preparar datos para GSheets (fechas a string para evitar errores JSON)
                        df_save = df_resultado.copy()
                        for c in df_save.select_dtypes(['datetime']): df_save[c] = df_save[c].astype(str)
                        df_save = df_save.fillna('')
                        
                        set_with_dataframe(ws_master, df_save)
                        st.info("☁️ Resultados sincronizados con GSheets 'Bancos_Master'.")
                        st.caption("Nota: Para mejorar la inteligencia futura, agrega los casos difíciles a la hoja 'Knowledge_Base'.")
                    except Exception as e:
                        st.error(f"No se pudo guardar en la nube: {e}")

if __name__ == "__main__":
    main()
