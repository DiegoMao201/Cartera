# ======================================================================================
# ARCHIVO: pages/2_Motor_Conciliacion.py
# (Versión v8.0 - Limpieza Robusta de Moneda Colombiana + Inteligencia Financiera)
# ======================================================================================

import streamlit as st
import pandas as pd
import dropbox
from io import StringIO, BytesIO
import re
import unicodedata
from datetime import datetime, timedelta
from fuzzywuzzy import fuzz, process
import gspread
from gspread_dataframe import set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials
import xlsxwriter

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Motor Conciliación Pereira", page_icon="🧠", layout="wide")

# ======================================================================================
# --- 1. CONEXIONES Y UTILIDADES ---
# ======================================================================================

@st.cache_resource
def get_dbx_client(secrets_key):
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
    try:
        scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_service_account"]), scope)
        return gspread.authorize(creds)
    except Exception as e:
        return None

def download_from_dropbox(dbx, path):
    try:
        _, res = dbx.files_download(path)
        return res.content
    except Exception as e:
        st.error(f"Error descargando {path} de Dropbox: {e}")
        return None

def normalizar_texto_avanzado(texto):
    if not isinstance(texto, str): return ""
    texto = texto.upper().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'[^A-Z0-9\s]', ' ', texto) 
    palabras_basura = [
        'PAGO', 'TRANSF', 'TRANSFERENCIA', 'CONSIGNACION', 'ABONO', 'CTA', 'NIT', 
        'REF', 'FACTURA', 'OFI', 'SUC', 'ACH', 'PSE', 'NOMINA', 'PROVEEDOR', 
        'COMPRA', 'VENTA', 'VALOR', 'NETO', 'PLANILLA'
    ]
    for p in palabras_basura:
        texto = re.sub(r'\b' + p + r'\b', '', texto)
    return ' '.join(texto.split())

def extraer_posibles_nits(texto):
    if not isinstance(texto, str): return []
    return re.findall(r'\b\d{7,11}\b', texto)

def extraer_dinero_de_texto(texto):
    """
    Busca números con formato de moneda en el texto si el valor es 0.
    """
    if not isinstance(texto, str): return 0.0
    matches = re.findall(r'(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?)', texto)
    
    valores_candidatos = []
    for m in matches:
        # Limpieza agresiva para convertir texto a float
        clean_m = m.replace(',', '').replace('.', '')
        try:
            val = float(clean_m)
            # Regla: Si > 1000, probablemente es dinero y no un código
            if val > 1000: 
                valores_candidatos.append(val)
        except: pass
    
    if valores_candidatos:
        return max(valores_candidatos)
    return 0.0

def limpiar_moneda_colombiana(val):
    """
    NUEVA FUNCIÓN DE LIMPIEZA ROBUSTA
    Convierte formatos como '1.000.000,00' o '$ 500.000' a float puro.
    Asume formato colombiano: Punto (.) para miles, Coma (,) para decimales.
    """
    # 1. Si ya es número, devolverlo como float
    if isinstance(val, (int, float)):
        return float(val) if pd.notnull(val) else 0.0
    
    # 2. Si es string, limpiar
    s = str(val).strip()
    if not s or s.lower() == 'nan': return 0.0

    # Quitar símbolos de moneda y espacios
    s = s.replace('$', '').replace('USD', '').replace('COP', '').strip()
    
    # Lógica específica para formato "1.234.567,00"
    # Paso A: Eliminar los puntos de miles (1.234 -> 1234)
    s = s.replace('.', '')
    
    # Paso B: Reemplazar la coma decimal por punto (1234,56 -> 1234.56)
    s = s.replace(',', '.')
    
    try:
        return float(s)
    except ValueError:
        return 0.0

# ======================================================================================
# --- 2. EXCEL DE LUJO ---
# ======================================================================================

def generar_excel_bonito(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Conciliacion_Pereira')
        workbook = writer.book
        worksheet = writer.sheets['Conciliacion_Pereira']
        
        # Formatos
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'top',
            'fg_color': '#1F497D', 'font_color': 'white', 'border': 1
        })
        money_format = workbook.add_format({'num_format': '$ #,##0', 'border': 1})
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1})
        
        # Colores Semáforo
        green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1}) # Conciliado
        blue_format = workbook.add_format({'bg_color': '#BDD7EE', 'font_color': '#1F497D', 'border': 1}) # Descuento
        orange_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700', 'border': 1}) # Retencion/Abono
        red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1}) # Revisar

        # Anchos
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            width = 20
            val_str = str(value).upper()
            if "TEXTO" in val_str: width = 50
            elif "CLIENTE" in val_str: width = 35
            elif "FECHA" in val_str: width = 15
            elif "VALOR" in val_str or "DEUDA" in val_str: width = 18
            worksheet.set_column(col_num, col_num, width)

        # Formatos Condicionales
        col_estado_idx = df.columns.get_loc('Estado') if 'Estado' in df.columns else -1
        nrow = len(df) + 1
        
        if col_estado_idx != -1:
            col_letter = chr(65 + col_estado_idx) # Asumiendo columnas < 26 (A-Z)
            # Verde: Total
            worksheet.conditional_format(1, 0, nrow, len(df.columns)-1, {
                'type': 'formula', 'criteria': f'=ISNUMBER(SEARCH("TOTAL", ${col_letter}2))', 'format': green_format
            })
            # Azul: Descuento
            worksheet.conditional_format(1, 0, nrow, len(df.columns)-1, {
                'type': 'formula', 'criteria': f'=ISNUMBER(SEARCH("DESCUENTO", ${col_letter}2))', 'format': blue_format
            })
            # Naranja: Retenciones/Abono
            worksheet.conditional_format(1, 0, nrow, len(df.columns)-1, {
                'type': 'formula', 'criteria': f'=OR(ISNUMBER(SEARCH("RETENCION", ${col_letter}2)), ISNUMBER(SEARCH("ABONO", ${col_letter}2)))',
                'format': orange_format
            })
            # Rojo: Revisar
            worksheet.conditional_format(1, 0, nrow, len(df.columns)-1, {
                'type': 'formula', 'criteria': f'=ISNUMBER(SEARCH("REVISAR", ${col_letter}2))', 'format': red_format
            })

        # Aplicar formato moneda
        cols_moneda = ['Valor_Banco_Calc', 'Deuda_Total_Cartera', 'Diferencia', 'Ahorro_Descuento_3%', 'Impuesto_Estimado']
        for col_name in cols_moneda:
            if col_name in df.columns:
                idx = df.columns.get_loc(col_name)
                worksheet.set_column(idx, idx, 18, money_format)

        worksheet.autofilter(0, 0, nrow, len(df.columns)-1)

    return output.getvalue()

# ======================================================================================
# --- 3. CARGA DE ARCHIVOS ---
# ======================================================================================

@st.cache_data(ttl=600)
def cargar_cartera():
    dbx = get_dbx_client("dropbox")
    if not dbx: 
        st.warning("⚠️ Sin conexión a Dropbox.")
        return pd.DataFrame()
        
    content = download_from_dropbox(dbx, '/data/cartera_detalle.csv')
    if not content: return pd.DataFrame()

    try:
        df = pd.read_csv(StringIO(content.decode('latin-1')), sep='|', header=None)
        df = df.iloc[:, :18]
        df.columns = [
            'Serie', 'Numero', 'FechaDoc', 'FechaVenc', 'CodCliente', 'NombreCliente', 
            'Nit', 'Poblacion', 'Provincia', 'Tel1', 'Tel2', 'Vendedor', 
            'Autoriza', 'Email', 'Importe', 'Descuento', 'Cupo', 'DiasVenc'
        ]
        
        df['Importe'] = pd.to_numeric(df['Importe'], errors='coerce').fillna(0)
        df['nit_norm'] = df['Nit'].astype(str).str.replace(r'[^0-9]', '', regex=True)
        df['nombre_norm'] = df['NombreCliente'].apply(normalizar_texto_avanzado)
        df['FechaDoc'] = pd.to_datetime(df['FechaDoc'], errors='coerce')
        
        return df[df['Importe'] > 0].copy()
    except Exception as e:
        st.error(f"Error leyendo cartera: {e}")
        return pd.DataFrame()

def cargar_planilla_pereira_desde_upload(uploaded_file):
    try:
        # Detectar encabezado
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
            st.error("❌ No encontré columnas FECHA/VALOR en las primeras 15 filas.")
            return pd.DataFrame()

        uploaded_file.seek(0) 
        # Leemos todo como string primero para evitar que Pandas "adivine" mal los números europeos
        df = pd.read_excel(uploaded_file, header=header_idx)
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # 1. Limpieza de FECHA
        df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
        df = df.dropna(subset=['FECHA'])
        
        # 2. Limpieza de VALOR (Aplicando la nueva función robusta)
        if 'VALOR' in df.columns:
            # Primero forzamos a string para que nuestra funcion limpie bien los puntos y comas
            df['VALOR_TEMP'] = df['VALOR'].astype(str).apply(limpiar_moneda_colombiana)
        else:
            df['VALOR_TEMP'] = 0.0

        # 3. Preparar Texto para Análisis
        cols_clave = ['SUCURSAL BANCO', 'TIPO DE TRANSACCION', 'CUENTA', 'EMPRESA', 'DESTINO', 'DETALLE', 'NOTAS']
        df['texto_analisis'] = ''
        for col in cols_clave:
            if col in df.columns:
                df['texto_analisis'] += df[col].fillna('').astype(str) + ' '
        df['texto_norm'] = df['texto_analisis'].apply(normalizar_texto_avanzado)

        # 4. Rescate de valores (Si la limpieza dio 0, buscar en texto)
        def definir_valor_final(row):
            val_limpio = row['VALOR_TEMP']
            if val_limpio == 0:
                # Intentar rescatar del texto
                return extraer_dinero_de_texto(row['texto_analisis'])
            return val_limpio

        df['Valor_Banco_Calc'] = df.apply(definir_valor_final, axis=1)
        df['Fue_Rescatado_Texto'] = (df['VALOR_TEMP'] == 0) & (df['Valor_Banco_Calc'] > 0)

        # ID Único
        def safe_id(row):
            f_str = row['FECHA'].strftime('%Y%m%d') if pd.notnull(row['FECHA']) else "0000"
            return f"MOV-{f_str}-{int(row['Valor_Banco_Calc'])}-{row.name}"

        df['id_unico'] = df.apply(safe_id, axis=1)
        
        return df

    except Exception as e:
        st.error(f"Error procesando Excel: {e}")
        return pd.DataFrame()

# ======================================================================================
# --- 4. MOTOR INTELIGENTE ---
# ======================================================================================

def ejecutar_motor_inteligente(df_bancos, df_cartera, df_kb):
    st.info("🧠 Cerebro Activado: Cruzando datos con limpieza monetaria avanzada...")
    
    mapa_nit_nombre = df_cartera.groupby('nit_norm')['NombreCliente'].first().to_dict()
    lista_nombres = df_cartera['nombre_norm'].unique().tolist()
    
    memoria_inteligente = {}
    if not df_kb.empty:
        if len(df_kb.columns) >= 2:
            for _, row in df_kb.iterrows():
                try:
                    k = normalizar_texto_avanzado(str(row.iloc[0]))
                    v = str(row.iloc[1]).strip()
                    if k and v: memoria_inteligente[k] = v
                except: pass

    resultados = []
    hoy = datetime.now()
    bar = st.progress(0)
    total_filas = len(df_bancos)
    
    for i, row in df_bancos.iterrows():
        bar.progress((i+1)/total_filas)
        
        txt_banco = row['texto_norm']
        txt_crudo = row['texto_analisis']
        valor_banco = row['Valor_Banco_Calc']
        
        res = {
            'Estado': 'PENDIENTE', 
            'Cliente_Identificado': 'NO IDENTIFICADO', 
            'NIT_Encontrado': '',
            'Tipo_Hallazgo': '',
            'Deuda_Total_Cartera': 0,
            'Diferencia': 0,
            'Ahorro_Descuento_3%': 0,
            'Impuesto_Estimado': 0,
            'Notas_Robot': ''
        }
        
        if row['Fue_Rescatado_Texto']:
            res['Notas_Robot'] += "⚠️ Dinero leído del texto (celda valor vacía o 0). "

        match_found = False
        nit_candidato = None

        # 1. Memoria
        if not match_found:
            for k_mem in memoria_inteligente:
                if k_mem in txt_banco and len(k_mem) > 5:
                    nit_candidato = memoria_inteligente[k_mem]
                    res['Tipo_Hallazgo'] = '🧠 Memoria Histórica'
                    match_found = True
                    break
        
        # 2. NIT
        if not match_found:
            posibles_nits = extraer_posibles_nits(txt_crudo)
            for pn in posibles_nits:
                if pn in mapa_nit_nombre:
                    nit_candidato = pn
                    res['Tipo_Hallazgo'] = f'🔍 NIT Detectado ({pn})'
                    match_found = True
                    break
        
        # 3. Fuzzy
        if not match_found and len(txt_banco) > 5:
            match_name, score = process.extractOne(txt_banco, lista_nombres, scorer=fuzz.token_set_ratio)
            if score >= 88:
                nit_candidato = df_cartera[df_cartera['nombre_norm'] == match_name]['nit_norm'].iloc[0]
                res['Tipo_Hallazgo'] = f'🤖 Nombre Similar ({score}%)'
                match_found = True

        # Análisis Financiero
        if match_found and nit_candidato:
            nombre_real = mapa_nit_nombre.get(nit_candidato, "Desconocido")
            res['NIT_Encontrado'] = nit_candidato
            res['Cliente_Identificado'] = nombre_real
            
            facturas_cliente = df_cartera[df_cartera['nit_norm'] == nit_candidato]
            deuda_total = facturas_cliente['Importe'].sum()
            res['Deuda_Total_Cartera'] = deuda_total
            
            diferencia = deuda_total - valor_banco
            res['Diferencia'] = diferencia
            
            # Tolerancia para pagos exactos
            if abs(diferencia) < 2000:
                res['Estado'] = '✅ CONCILIADO - PAGO TOTAL'
            
            elif valor_banco < deuda_total:
                # Chequeo de Descuento (3%)
                deuda_con_dcto = deuda_total * 0.97
                diff_dcto = abs(deuda_con_dcto - valor_banco)
                
                if diff_dcto < 5000:
                    res['Estado'] = '💎 CONCILIADO - CON DESCUENTO 3%'
                    res['Ahorro_Descuento_3%'] = deuda_total * 0.03
                    res['Notas_Robot'] += "Descuento pronto pago aplicado. "
                else:
                     # Chequeo Impuestos
                     base_estimada = deuda_total / 1.19
                     rete_iva_est = base_estimada * 0.19 * 0.15
                     rete_fuente_est = base_estimada * 0.025
                     
                     pago_esperado_full = deuda_total - rete_fuente_est - rete_iva_est
                     pago_esperado_rf = deuda_total - rete_fuente_est
                     
                     tolerance = 5000
                     
                     if abs(pago_esperado_full - valor_banco) < tolerance:
                         res['Estado'] = '🏢 CONCILIADO - CON RETENCIONES (G.C.)'
                         res['Impuesto_Estimado'] = rete_fuente_est + rete_iva_est
                         res['Notas_Robot'] += "ReteIVA + ReteFuente detectados. "
                     elif abs(pago_esperado_rf - valor_banco) < tolerance:
                         res['Estado'] = '🏢 CONCILIADO - CON RETEFUENTE'
                         res['Impuesto_Estimado'] = rete_fuente_est
                         res['Notas_Robot'] += "ReteFuente (2.5%) detectada. "
                     else:
                         res['Estado'] = '⚠️ ABONO PARCIAL O DIFERENCIA'

            elif valor_banco > deuda_total:
                res['Estado'] = '❌ REVISAR - PAGO MAYOR A DEUDA'

        row_dict = row.to_dict()
        row_dict.update(res)
        resultados.append(row_dict)
        
    return pd.DataFrame(resultados)

# ======================================================================================
# --- 5. INTERFAZ ---
# ======================================================================================

def main():
    st.title("🏦 Super Motor - Limpieza y Conciliación")
    st.markdown("Versión v8.0: Limpieza de moneda colombiana (. para miles, , para decimales).")
    
    col1, col2 = st.columns([1, 1])
    with col1:
        uploaded_file = st.file_uploader("Planilla Banco (Cualquier formato de número)", type=["xlsx", "xls"])
    with col2:
        if st.button("☁️ Sincronizar Cartera", type="secondary"):
            with st.spinner("Conectando..."):
                df_c = cargar_cartera()
                if not df_c.empty:
                    st.session_state['df_cartera'] = df_c
                    st.success(f"Cartera: {len(df_c)} registros.")
                else: st.error("Error Dropbox.")

    if 'df_cartera' in st.session_state:
        st.caption(f"Cartera activa: {len(st.session_state['df_cartera'])} facturas.")

    st.divider()

    if uploaded_file and 'df_cartera' in st.session_state:
        if st.button("🚀 INICIAR PROCESO", type="primary", use_container_width=True):
            
            with st.status("Procesando...", expanded=True) as status:
                st.write("🧹 Limpiando formatos de moneda (colombianos/europeos)...")
                df_bancos = cargar_planilla_pereira_desde_upload(uploaded_file)
                
                if df_bancos.empty:
                    st.error("Archivo no válido.")
                    status.update(label="Error", state="error")
                    return

                g_client = connect_to_google_sheets()
                df_kb = pd.DataFrame()
                if g_client:
                    try:
                        sh = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"])
                        df_kb = pd.DataFrame(sh.worksheet("Knowledge_Base").get_all_records())
                    except: pass

                st.write("🤖 Calculando coincidencias...")
                df_resultado = ejecutar_motor_inteligente(df_bancos, st.session_state['df_cartera'], df_kb)
                status.update(label="¡Listo!", state="complete", expanded=False)

            # Métricas y Tabla
            c_tot = len(df_resultado[df_resultado['Estado'].str.contains("TOTAL")])
            c_imp = len(df_resultado[df_resultado['Estado'].str.contains("RETENCIONES") | df_resultado['Estado'].str.contains("RETEFUENTE")])
            
            c1, c2 = st.columns(2)
            c1.metric("Pagos Totales", c_tot)
            c2.metric("Con Impuestos", c_imp)
            
            st.dataframe(df_resultado[['FECHA', 'Valor_Banco_Calc', 'Cliente_Identificado', 'Estado', 'Impuesto_Estimado']], use_container_width=True)

            # Descarga y Guardado
            c_d, c_s = st.columns(2)
            with c_d:
                excel_data = generar_excel_bonito(df_resultado)
                st.download_button("📥 Bajar Excel Resultados", excel_data, "Conciliacion_Pereira_v8.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
            
            with c_s:
                if g_client:
                    try:
                        sh = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"])
                        ws = sh.worksheet(st.secrets["google_sheets"]["tab_bancos_master"])
                        ws.clear()
                        df_save = df_resultado.copy()
                        for c in df_save.select_dtypes(['datetime']): df_save[c] = df_save[c].astype(str)
                        df_save = df_save.fillna('')
                        set_with_dataframe(ws, df_save)
                        st.success("Guardado en Nube.")
                    except: pass

if __name__ == "__main__":
    main()
