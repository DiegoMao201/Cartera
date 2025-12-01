# ======================================================================================
# ARCHIVO: pages/2_Motor_Conciliacion.py
# (Versión v7.0 - Super Inteligencia: Descuentos, Impuestos y Rescate de Valores en 0)
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
        # st.error(f"Error conectando a Google Sheets: {e}") # Silenciar error en UI si no es crítico
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
    SUPER PODER: Si el valor es 0, busca números con formato de moneda en el texto.
    Ej: 'Pago (313,885)' -> Retorna 313885.0
    """
    if not isinstance(texto, str): return 0.0
    # Busca patrones como 313,885 o 313.885 o 1.000.000
    # Regex explica: digitos seguidos de coma/punto y mas digitos
    matches = re.findall(r'(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?)', texto)
    
    valores_candidatos = []
    for m in matches:
        # Limpiar puntuación para convertir a float
        clean_m = m.replace(',', '').replace('.', '')
        try:
            val = float(clean_m)
            # Regla heurística: Si tiene decimales implícitos (ej 313885 en texto suele ser 313885 pesos)
            # A veces Excel lee 313.885 como 313 mil
            if val > 1000: # Ignoramos números pequeños que parezcan códigos
                valores_candidatos.append(val)
        except: pass
    
    if valores_candidatos:
        return max(valores_candidatos) # Asumimos que el valor más grande es el pago
    return 0.0

# ======================================================================================
# --- 2. EXCEL DE LUJO (VERSION FINANCIERA) ---
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

        # 1. Anchos y Encabezados
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            width = 20
            if "TEXTO" in str(value).upper(): width = 50
            elif "CLIENTE" in str(value).upper(): width = 35
            elif "FECHA" in str(value).upper(): width = 15
            elif "VALOR" in str(value).upper() or "DEUDA" in str(value).upper(): width = 18
            worksheet.set_column(col_num, col_num, width)

        # 2. Formatos Condicionales Inteligentes
        col_estado_idx = df.columns.get_loc('Estado') if 'Estado' in df.columns else -1
        col_val_banco = df.columns.get_loc('Valor_Banco_Calc') if 'Valor_Banco_Calc' in df.columns else -1
        
        nrow = len(df) + 1
        
        if col_estado_idx != -1:
            # Verde: Pago Total Exacto
            worksheet.conditional_format(1, 0, nrow, len(df.columns)-1, {
                'type': 'formula', 'criteria': f'=SEARCH("TOTAL", ${chr(65+col_estado_idx)}2)',
                'format': green_format
            })
            # Azul: Descuento Pronto Pago
            worksheet.conditional_format(1, 0, nrow, len(df.columns)-1, {
                'type': 'formula', 'criteria': f'=SEARCH("DESCUENTO", ${chr(65+col_estado_idx)}2)',
                'format': blue_format
            })
            # Naranja: Retenciones o Abonos
            worksheet.conditional_format(1, 0, nrow, len(df.columns)-1, {
                'type': 'formula', 'criteria': f'=OR(SEARCH("RETENCION", ${chr(65+col_estado_idx)}2), SEARCH("ABONO", ${chr(65+col_estado_idx)}2))',
                'format': orange_format
            })
            # Rojo: Revisar
            worksheet.conditional_format(1, 0, nrow, len(df.columns)-1, {
                'type': 'formula', 'criteria': f'=SEARCH("REVISAR", ${chr(65+col_estado_idx)}2)',
                'format': red_format
            })

        # Formato Moneda a columnas numéricas
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
        
        # Convertir Fechas para cálculo de antigüedad
        df['FechaDoc'] = pd.to_datetime(df['FechaDoc'], errors='coerce')
        
        return df[df['Importe'] > 0].copy()
    except Exception as e:
        st.error(f"Error leyendo cartera: {e}")
        return pd.DataFrame()

def cargar_planilla_pereira_desde_upload(uploaded_file):
    try:
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
            st.error("❌ No encontré columnas FECHA/VALOR.")
            return pd.DataFrame()

        uploaded_file.seek(0) 
        df = pd.read_excel(uploaded_file, header=header_idx)
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
        df['VALOR'] = pd.to_numeric(df['VALOR'], errors='coerce').fillna(0)
        df = df.dropna(subset=['FECHA'])
        
        # --- NUEVA LÓGICA DE EXTRACCIÓN DE TEXTO ---
        cols_clave = ['SUCURSAL BANCO', 'TIPO DE TRANSACCION', 'CUENTA', 'EMPRESA', 'DESTINO', 'DETALLE', 'NOTAS']
        df['texto_analisis'] = ''
        for col in cols_clave:
            if col in df.columns:
                df['texto_analisis'] += df[col].fillna('').astype(str) + ' '
                
        df['texto_norm'] = df['texto_analisis'].apply(normalizar_texto_avanzado)

        # Corrección de Valor 0: Si es 0, intentar leer del texto
        def corregir_valor_cero(row):
            if row['VALOR'] == 0 or pd.isna(row['VALOR']):
                # Intentar extraer del texto
                val_rescatado = extraer_dinero_de_texto(row['texto_analisis'])
                return val_rescatado
            return row['VALOR']

        df['Valor_Banco_Calc'] = df.apply(corregir_valor_cero, axis=1)
        df['Fue_Rescatado_Texto'] = (df['VALOR'] == 0) & (df['Valor_Banco_Calc'] > 0)

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
# --- 4. SUPER MOTOR INTELIGENTE (SCENARIO MATCHER) ---
# ======================================================================================

def ejecutar_motor_inteligente(df_bancos, df_cartera, df_kb):
    st.info("🧠 Cerebro Financiero Activado: Buscando Descuentos, Retenciones y Errores...")
    
    # Índices
    mapa_nit_nombre = df_cartera.groupby('nit_norm')['NombreCliente'].first().to_dict()
    # Para calculo avanzado, necesitamos agrupar por NIT pero mantener detalle de fechas
    # mapa_nit_deuda = df_cartera.groupby('nit_norm')['Importe'].sum().to_dict() 
    lista_nombres = df_cartera['nombre_norm'].unique().tolist()
    
    # KB
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
        valor_banco = row['Valor_Banco_Calc'] # Usamos el valor corregido/rescatado
        
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
        
        # --- PASO 1: IDENTIFICACIÓN DEL CLIENTE ---
        match_found = False
        nit_candidato = None

        # A. Memoria
        if not match_found:
            for k_mem in memoria_inteligente:
                if k_mem in txt_banco and len(k_mem) > 5:
                    nit_candidato = memoria_inteligente[k_mem]
                    res['Tipo_Hallazgo'] = '🧠 Memoria (KB)'
                    match_found = True
                    break
        
        # B. NIT Exacto
        if not match_found:
            posibles_nits = extraer_posibles_nits(txt_crudo)
            for pn in posibles_nits:
                if pn in mapa_nit_nombre:
                    nit_candidato = pn
                    res['Tipo_Hallazgo'] = f'🔍 NIT Detectado ({pn})'
                    match_found = True
                    break
        
        # C. Fuzzy Name
        if not match_found and len(txt_banco) > 5:
            match_name, score = process.extractOne(txt_banco, lista_nombres, scorer=fuzz.token_set_ratio)
            if score >= 88:
                nit_candidato = df_cartera[df_cartera['nombre_norm'] == match_name]['nit_norm'].iloc[0]
                res['Tipo_Hallazgo'] = f'🤖 IA Nombre Similar ({score}%)'
                match_found = True

        # --- PASO 2: ANÁLISIS FINANCIERO AVANZADO (Descuentos e Impuestos) ---
        if match_found and nit_candidato:
            nombre_real = mapa_nit_nombre.get(nit_candidato, "Desconocido")
            res['NIT_Encontrado'] = nit_candidato
            res['Cliente_Identificado'] = nombre_real
            
            # Traer TODAS las facturas de este cliente
            facturas_cliente = df_cartera[df_cartera['nit_norm'] == nit_candidato]
            deuda_total = facturas_cliente['Importe'].sum()
            res['Deuda_Total_Cartera'] = deuda_total
            
            diferencia = deuda_total - valor_banco
            res['Diferencia'] = diferencia
            
            if row['Fue_Rescatado_Texto']:
                res['Notas_Robot'] += "⚠️ VALOR 0 EN EXCEL -> DINERO LEÍDO DEL TEXTO. "

            # --- ESCENARIO 1: PAGO TOTAL EXACTO ---
            if abs(diferencia) < 2000:
                res['Estado'] = '✅ CONCILIADO - PAGO TOTAL'
            
            # --- ESCENARIO 2: DESCUENTO PRONTO PAGO (3%) ---
            elif valor_banco < deuda_total:
                # Filtrar facturas "jóvenes" (<= 30 días de antigüedad desde FechaDoc)
                # Asumimos que hoy es la fecha de análisis, o usamos la fecha del pago
                fecha_pago = row['FECHA'] if pd.notnull(row['FECHA']) else hoy
                
                # Calculamos si aplicando 3% a las facturas recientes, el valor cuadra
                # Base de facturas que podrian tener descuento
                monto_candidato_descuento = 0
                monto_sin_descuento = 0
                
                # Lógica simplificada: Si aplicamos 3% a TODA la deuda, ¿cuadra?
                deuda_con_dcto = deuda_total * 0.97
                diff_dcto = abs(deuda_con_dcto - valor_banco)
                
                # O quizas solo a las facturas jovenes?
                # Vamos a probar la regla general del cliente primero
                if diff_dcto < 5000: # Tolerancia de 5 mil pesos
                    res['Estado'] = '💎 CONCILIADO - CON DESCUENTO 3%'
                    res['Ahorro_Descuento_3%'] = deuda_total * 0.03
                    res['Notas_Robot'] += "Cliente aplicó 3% de descuento por pronto pago. "
                
                else:
                     # --- ESCENARIO 3: RETENCIONES (GRAN CONTRIBUYENTE) ---
                     # ReteIVA usual es 15% del IVA. El IVA es el 19% de la Base.
                     # Base = Deuda / 1.19
                     # IVA = Base * 0.19
                     # ReteIVA = IVA * 0.15 => (Deuda / 1.19) * 0.19 * 0.15
                     # ReteFuente usual = 2.5% de la Base => (Deuda / 1.19) * 0.025
                     
                     base_estimada = deuda_total / 1.19
                     rete_iva_est = base_estimada * 0.19 * 0.15
                     rete_fuente_est = base_estimada * 0.025
                     rete_ica_est = base_estimada * 0.005 # 5 por mil promedio
                     
                     # Probamos combinaciones
                     # 1. Solo ReteFuente
                     pago_esperado_rf = deuda_total - rete_fuente_est
                     # 2. ReteFuente + ReteIVA
                     pago_esperado_full = deuda_total - rete_fuente_est - rete_iva_est
                     # 3. Solo ReteIVA
                     pago_esperado_ri = deuda_total - rete_iva_est
                     
                     tolerance = 5000
                     
                     if abs(pago_esperado_full - valor_banco) < tolerance:
                         res['Estado'] = '🏢 CONCILIADO - CON RETENCIONES (G.C.)'
                         res['Impuesto_Estimado'] = rete_fuente_est + rete_iva_est
                         res['Notas_Robot'] += "Detectado Gran Contribuyente (ReteIVA + ReteFuente aplicados). "
                     elif abs(pago_esperado_rf - valor_banco) < tolerance:
                         res['Estado'] = '🏢 CONCILIADO - CON RETEFUENTE'
                         res['Impuesto_Estimado'] = rete_fuente_est
                         res['Notas_Robot'] += "Detectada ReteFuente (2.5%). "
                     elif valor_banco < deuda_total:
                         res['Estado'] = '⚠️ ABONO PARCIAL O DIFERENCIA'
                     else:
                         res['Estado'] = '❌ REVISAR - PAGO MAYOR'

            elif valor_banco > deuda_total:
                res['Estado'] = '❌ REVISAR - PAGO MAYOR A DEUDA'

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
    **Novedades v7.0:**
    1. 🕵️‍♂️ **Detector de ceros:** Si el Excel dice $0, el robot lee el texto y busca el dinero.
    2. 📉 **Descuentos:** Detecta pagos con 3% de descuento por pronto pago (<= 30 días).
    3. 🏢 **Impuestos:** Calcula automáticamente si faltó dinero por ReteIVA/ReteFuente.
    """)
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.subheader("1. Planilla Banco")
        uploaded_file = st.file_uploader("Sube el Excel (incluso con valores en 0)", type=["xlsx", "xls"])
    
    with col2:
        st.subheader("2. Cartera Dropbox")
        if st.button("☁️ Sincronizar Cartera", type="secondary"):
            with st.spinner("Bajando datos..."):
                df_c = cargar_cartera()
                if not df_c.empty:
                    st.session_state['df_cartera'] = df_c
                    st.success(f"✅ Cartera Lista: {len(df_c)} facturas.")
                else:
                    st.error("❌ Error en Dropbox.")

    if 'df_cartera' in st.session_state:
        st.caption(f"Cartera activa: {len(st.session_state['df_cartera'])} registros.")
    else:
        st.warning("⚠️ Carga la cartera primero.")

    st.divider()

    if uploaded_file and 'df_cartera' in st.session_state:
        if st.button("🚀 ANALIZAR FINANZAS", type="primary", use_container_width=True):
            
            with st.status("El Robot está trabajando...", expanded=True) as status:
                st.write("📂 Leyendo Excel y rescatando valores en 0...")
                df_bancos = cargar_planilla_pereira_desde_upload(uploaded_file)
                
                if df_bancos.empty:
                    st.error("Archivo vacío.")
                    status.update(label="Error", state="error")
                    return

                # Mostrar cuántos valores se rescataron del texto
                rescatados = df_bancos['Fue_Rescatado_Texto'].sum()
                if rescatados > 0:
                    st.warning(f"👁️ OJO: Se detectaron {rescatados} filas con valor 0. El robot extrajo el dinero leyendo la descripción.")

                st.write("🧠 Consultando Inteligencia Histórica...")
                g_client = connect_to_google_sheets()
                df_kb = pd.DataFrame()
                if g_client:
                    try:
                        sh = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"])
                        try:
                            df_kb = pd.DataFrame(sh.worksheet("Knowledge_Base").get_all_records())
                        except: pass
                    except: pass

                st.write("🤖 Calculando Impuestos, Descuentos y Cruces...")
                df_resultado = ejecutar_motor_inteligente(df_bancos, st.session_state['df_cartera'], df_kb)
                
                status.update(label="¡Análisis Financiero Completado!", state="complete", expanded=False)

            # Métricas
            c_total = len(df_resultado[df_resultado['Estado'].str.contains("PAGO TOTAL")])
            c_dcto = len(df_resultado[df_resultado['Estado'].str.contains("DESCUENTO")])
            c_imp = len(df_resultado[df_resultado['Estado'].str.contains("RETENCIONES") | df_resultado['Estado'].str.contains("RETEFUENTE")])
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Pago Exacto", c_total)
            m2.metric("Con Descuento (3%)", c_dcto)
            m3.metric("Con Impuestos (G.C.)", c_imp)

            st.dataframe(df_resultado[['FECHA', 'Valor_Banco_Calc', 'Cliente_Identificado', 'Estado', 'Ahorro_Descuento_3%', 'Impuesto_Estimado', 'Notas_Robot']], use_container_width=True)

            # Exportar
            c_down, c_save = st.columns(2)
            with c_down:
                excel_data = generar_excel_bonito(df_resultado)
                st.download_button(
                    "📥 Descargar Reporte Financiero Inteligente",
                    excel_data,
                    f"Conciliacion_Pereira_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            
            with c_save:
                if g_client:
                    try:
                        sh = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"])
                        ws_master = sh.worksheet(st.secrets["google_sheets"]["tab_bancos_master"])
                        ws_master.clear()
                        df_save = df_resultado.copy()
                        for c in df_save.select_dtypes(['datetime']): df_save[c] = df_save[c].astype(str)
                        df_save = df_save.fillna('')
                        set_with_dataframe(ws_master, df_save)
                        st.success("☁️ Base de Datos actualizada en Google Sheets.")
                    except Exception as e:
                        st.error(f"Error nube: {e}")

if __name__ == "__main__":
    main()
