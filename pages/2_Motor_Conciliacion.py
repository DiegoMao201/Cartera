# ======================================================================================
# ARCHIVO: pages/2_Motor_Conciliacion.py
# (Versión v9.1 - "El Auditor": Corrección de Indentación y Limpieza Total)
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
import xlsxwriter
import itertools

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Motor Conciliación Pro v9", page_icon="🕵️‍♂️", layout="wide")

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
    except Exception:
        return None

def download_from_dropbox(dbx, path):
    try:
        _, res = dbx.files_download(path)
        return res.content
    except Exception as e:
        st.error(f"Error descargando {path}: {e}")
        return None

def normalizar_texto_avanzado(texto):
    if not isinstance(texto, str): return ""
    texto = texto.upper().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'[^A-Z0-9\s]', ' ', texto) 
    palabras_basura = [
        'PAGO', 'TRANSF', 'TRANSFERENCIA', 'CONSIGNACION', 'ABONO', 'CTA', 'NIT', 
        'REF', 'FACTURA', 'OFI', 'SUC', 'ACH', 'PSE', 'NOMINA', 'PROVEEDOR', 
        'COMPRA', 'VENTA', 'VALOR', 'NETO', 'PLANILLA', 'S A', 'SAS', 'LTDA'
    ]
    for p in palabras_basura:
        texto = re.sub(r'\b' + p + r'\b', '', texto)
    return ' '.join(texto.split())

def extraer_posibles_nits(texto):
    if not isinstance(texto, str): return []
    # Busca secuencias de 7 a 11 dígitos que suelen ser NITs
    return re.findall(r'\b\d{7,11}\b', texto)

def limpiar_moneda_colombiana(val):
    """
    Convierte formatos como '1.000.000,00' o '$ 500.000' a float puro.
    Asume formato colombiano: Punto (.) miles, Coma (,) decimales.
    """
    if isinstance(val, (int, float)):
        return float(val) if pd.notnull(val) else 0.0
    
    s = str(val).strip()
    if not s or s.lower() == 'nan': return 0.0

    s = s.replace('$', '').replace('USD', '').replace('COP', '').strip()
    # 1.234.567,00 -> 1234567.00
    s = s.replace('.', '') # Quitar miles
    s = s.replace(',', '.') # Convertir decimal
    
    try:
        return float(s)
    except ValueError:
        return 0.0

def extraer_dinero_de_texto(texto):
    if not isinstance(texto, str): return 0.0
    matches = re.findall(r'(\d{1,3}(?:[.,]\d{3})*(?:[.,]\d{2})?)', texto)
    valores = []
    for m in matches:
        clean_m = m.replace(',', '').replace('.', '')
        try:
            val = float(clean_m)
            if val > 1000: valores.append(val)
        except: pass
    return max(valores) if valores else 0.0

# ======================================================================================
# --- 2. EXCEL DE ALTO IMPACTO (FORMATO V9) ---
# ======================================================================================

def generar_excel_profesional(df, resumen_kpis):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        
        # --- HOJA 1: RESUMEN GERENCIAL ---
        workbook = writer.book
        sheet_resumen = workbook.add_worksheet("Dashboard")
        sheet_resumen.hide_gridlines(2)
        
        # Estilos Dashboard
        style_title = workbook.add_format({'bold': True, 'font_size': 16, 'font_color': '#1F497D'})
        style_kpi_label = workbook.add_format({'bold': True, 'bg_color': '#E7E6E6', 'border': 1})
        style_kpi_value = workbook.add_format({'bold': True, 'num_format': '#,##0', 'border': 1, 'align': 'center'})
        
        sheet_resumen.write('B2', "RESUMEN DE CONCILIACIÓN - PEREIRA", style_title)
        
        # Tabla KPIs
        kpis = [
            ("Total Movimientos", resumen_kpis['total_tx']),
            ("Conciliados (Match Exacto)", resumen_kpis['exactos']),
            ("Conciliados (Con Descuento)", resumen_kpis['descuentos']),
            ("Conciliados (Impuestos)", resumen_kpis['impuestos']),
            ("Parciales / Abonos", resumen_kpis['parciales']),
            ("Sin Identificar", resumen_kpis['sin_id'])
        ]
        
        row = 4
        for label, val in kpis:
            sheet_resumen.write(row, 1, label, style_kpi_label)
            sheet_resumen.write(row, 2, val, style_kpi_value)
            row += 1

        # --- HOJA 2: DETALLE ---
        df.to_excel(writer, index=False, sheet_name='Detalle_Conciliacion')
        worksheet = writer.sheets['Detalle_Conciliacion']
        
        # Estilos Detalle
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'fg_color': '#203764', 'font_color': 'white', 'border': 1
        })
        currency_fmt = workbook.add_format({'num_format': '$ #,##0.00', 'border': 1})
        text_fmt = workbook.add_format({'text_wrap': False, 'border': 1})
        
        # Colores de Estado
        fmt_green = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
        fmt_blue = workbook.add_format({'bg_color': '#BDD7EE', 'font_color': '#1F497D'})
        fmt_red = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        fmt_yellow = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700'})

        # Configurar columnas
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
            # Anchos automáticos inteligentes
            col_name = str(value).upper()
            width = 15
            if "CLIENTE" in col_name or "RAZON" in col_name: width = 35
            elif "DETALLE" in col_name or "NOTAS" in col_name: width = 50
            elif "FACTURAS" in col_name: width = 30
            elif "VALOR" in col_name or "SALDO" in col_name: width = 18
            
            worksheet.set_column(col_num, col_num, width, text_fmt)
            
            # Aplicar formato moneda a columnas de dinero
            if any(x in col_name for x in ['VALOR', 'DEUDA', 'DIFERENCIA', 'AHORRO', 'IMPUESTO', 'SALDO']):
                worksheet.set_column(col_num, col_num, width, currency_fmt)

        # Formato Condicional en columna ESTADO
        try:
            col_idx = df.columns.get_loc('Estado')
            col_letter = chr(65 + col_idx)
            last_row = len(df) + 1
            
            rng = f"{col_letter}2:{col_letter}{last_row}"
            
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'EXACTO', 'format': fmt_green})
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'DESCUENTO', 'format': fmt_blue})
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'REVISAR', 'format': fmt_red})
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'ABONO', 'format': fmt_yellow})
        except: pass

        worksheet.freeze_panes(1, 0)
        
    return output.getvalue()

# ======================================================================================
# --- 3. CARGA DE DATOS ---
# ======================================================================================

@st.cache_data(ttl=600)
def cargar_cartera_detalle():
    """Carga y procesa el detalle de facturas pendientes"""
    dbx = get_dbx_client("dropbox")
    if not dbx: return pd.DataFrame()
    
    content = download_from_dropbox(dbx, '/data/cartera_detalle.csv')
    if not content: return pd.DataFrame()

    try:
        df = pd.read_csv(StringIO(content.decode('latin-1')), sep='|', header=None)
        # Ajustar a columnas reales de tu archivo
        df = df.iloc[:, :18] 
        df.columns = [
            'Serie', 'Numero', 'FechaDoc', 'FechaVenc', 'CodCliente', 'NombreCliente', 
            'Nit', 'Poblacion', 'Provincia', 'Tel1', 'Tel2', 'Vendedor', 
            'Autoriza', 'Email', 'Importe', 'Descuento', 'Cupo', 'DiasVenc'
        ]
        
        # Limpieza
        df['Importe'] = pd.to_numeric(df['Importe'], errors='coerce').fillna(0)
        df['nit_norm'] = df['Nit'].astype(str).str.replace(r'[^0-9]', '', regex=True)
        df['nombre_norm'] = df['NombreCliente'].apply(normalizar_texto_avanzado)
        df['FechaDoc'] = pd.to_datetime(df['FechaDoc'], errors='coerce')
        
        # Solo facturas con saldo positivo
        return df[df['Importe'] > 100].copy()
    except Exception as e:
        st.error(f"Error estructura cartera: {e}")
        return pd.DataFrame()

def procesar_planilla_bancos(uploaded_file):
    try:
        # Previsualizar para encontrar encabezado
        df_temp = pd.read_excel(uploaded_file, nrows=15, header=None)
        header_idx = 0
        for idx, row in df_temp.iterrows():
            if 'FECHA' in row.astype(str).str.upper().values and 'VALOR' in row.astype(str).str.upper().values:
                header_idx = idx
                break
        
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, header=header_idx)
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # Limpieza Fechas
        df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
        df = df.dropna(subset=['FECHA'])
        
        # Limpieza Valor
        if 'VALOR' in df.columns:
            df['Valor_Banco'] = df['VALOR'].apply(limpiar_moneda_colombiana)
        else:
            df['Valor_Banco'] = 0.0

        # Crear texto de búsqueda
        cols_txt = [c for c in df.columns if c not in ['FECHA', 'VALOR', 'Valor_Banco']]
        df['Texto_Completo'] = df[cols_txt].fillna('').astype(str).agg(' '.join, axis=1)
        df['Texto_Norm'] = df['Texto_Completo'].apply(normalizar_texto_avanzado)
        
        # Rescate de dinero si columna valor es 0
        mask_zero = df['Valor_Banco'] == 0
        df.loc[mask_zero, 'Valor_Banco'] = df.loc[mask_zero, 'Texto_Completo'].apply(extraer_dinero_de_texto)
        df['Rescatado'] = mask_zero & (df['Valor_Banco'] > 0)
        
        return df
    except Exception as e:
        st.error(f"Error leyendo Excel: {e}")
        return pd.DataFrame()

# ======================================================================================
# --- 4. ALGORITMO "EL AUDITOR" (MATCHING LOGIC) ---
# ======================================================================================

def analizar_cliente(nombre_banco, valor_pago, facturas_cliente):
    """
    Analiza un pago contra las facturas abiertas de un cliente específico.
    Retorna un diccionario con el resultado del análisis.
    """
    res = {
        'Estado': '⚠️ SIN COINCIDENCIA CLARA',
        'Facturas_Conciliadas': '',
        'Detalle_Operacion': '',
        'Diferencia': 0,
        'Tipo_Ajuste': 'Ninguno',
        'Ahorro_Dcto': 0,
        'Impuesto_Est': 0
    }
    
    if facturas_cliente.empty:
        res['Detalle_Operacion'] = "Cliente identificado pero sin cartera pendiente."
        res['Diferencia'] = valor_pago * -1 # Saldo a favor
        return res

    facturas = facturas_cliente[['Numero', 'Importe', 'FechaDoc']].sort_values('FechaDoc').to_dict('records')
    total_deuda = sum(f['Importe'] for f in facturas)
    
    # --- ESCENARIO 1: PAGO TOTAL DEUDA ---
    if abs(valor_pago - total_deuda) < 1000:
        res['Estado'] = '✅ MATCH EXACTO (TOTAL)'
        res['Facturas_Conciliadas'] = 'TODAS'
        res['Detalle_Operacion'] = f"Pago cubre las {len(facturas)} facturas pendientes."
        return res
        
    # --- ESCENARIO 2: BUSQUEDA COMBINATORIA (Pago de facturas especificas) ---
    # Intentamos encontrar si el pago suma exactamente a 1, 2 o 3 facturas.
    # Limitamos a combinaciones de 3 para rendimiento.
    found_combo = False
    
    for r in range(1, 4): 
        if r > len(facturas): break
        for combo in itertools.combinations(facturas, r):
            suma_combo = sum(c['Importe'] for c in combo)
            numeros_combo = ", ".join([str(c['Numero']) for c in combo])
            
            # A. Match Exacto
            if abs(valor_pago - suma_combo) < 500:
                res['Estado'] = '✅ MATCH FACTURAS ESPECÍFICAS'
                res['Facturas_Conciliadas'] = numeros_combo
                res['Detalle_Operacion'] = f"Suma exacta de {r} factura(s)."
                found_combo = True
                break
                
            # B. Match con Descuento (3% aprox)
            suma_dcto = suma_combo * 0.97
            if abs(valor_pago - suma_dcto) < 2000:
                res['Estado'] = '💎 CONCILIADO CON DESCUENTO'
                res['Facturas_Conciliadas'] = numeros_combo
                res['Tipo_Ajuste'] = 'Descuento Pronto Pago'
                res['Ahorro_Dcto'] = suma_combo - valor_pago
                res['Detalle_Operacion'] = f"Pago con 3% Dcto sobre facturas: {numeros_combo}"
                found_combo = True
                break
        if found_combo: break
    
    if found_combo: return res

    # --- ESCENARIO 3: LOGICA FIFO (Abono a la deuda mas vieja) ---
    # Si no cuadró ninguna combinación, asumimos que paga desde la más vieja
    acumulado = 0
    facturas_cubiertas = []
    
    for f in facturas:
        if acumulado + f['Importe'] <= valor_pago + 1000: # Tolerancia pequeña
            acumulado += f['Importe']
            facturas_cubiertas.append(str(f['Numero']))
        else:
            break
            
    saldo_restante = total_deuda - valor_pago
    
    if facturas_cubiertas:
        res['Estado'] = '⚠️ ABONO PARCIAL (FIFO)'
        res['Facturas_Conciliadas'] = ", ".join(facturas_cubiertas)
        res['Diferencia'] = saldo_restante
        res['Detalle_Operacion'] = f"Cubre facturas antiguas. Queda debiendo ${saldo_restante:,.0f}"
    else:
        # Chequeo impuestos sobre el total
        base_est = total_deuda / 1.19
        rete_fuente = base_est * 0.025
        rete_iva = (base_est * 0.19) * 0.15
        pago_con_imptos = total_deuda - rete_fuente - rete_iva
        
        if abs(valor_pago - pago_con_imptos) < 5000:
            res['Estado'] = '🏢 CONCILIADO (IMPUESTOS)'
            res['Impuesto_Est'] = rete_fuente + rete_iva
            res['Detalle_Operacion'] = "Coincide con total menos ReteFuente y ReteIVA."
        else:
            res['Estado'] = '❌ REVISAR MANUALMENTE'
            res['Diferencia'] = saldo_restante
            res['Detalle_Operacion'] = f"Monto no cuadra con facturas. Deuda Total: ${total_deuda:,.0f}"

    return res

def correr_motor_inteligente(df_bancos, df_cartera, df_kb):
    st.info("🔎 Iniciando auditoría detallada...")
    
    # Preparar mapas de búsqueda
    mapa_nit = df_cartera.groupby('nit_norm')['NombreCliente'].first().to_dict()
    lista_nombres = df_cartera['nombre_norm'].unique().tolist()
    
    # Cargar base de conocimiento
    memoria = {}
    if not df_kb.empty:
        for _, row in df_kb.iterrows():
            try: memoria[str(row[0]).strip().upper()] = str(row[1]).strip()
            except: pass

    resultados = []
    bar = st.progress(0)
    
    for i, row in df_bancos.iterrows():
        bar.progress((i+1)/len(df_bancos))
        
        item = row.to_dict()
        txt = row['Texto_Norm']
        val = row['Valor_Banco']
        
        nit_found = None
        nombre_cliente = "NO IDENTIFICADO"
        
        # 1. Identificación
        # A. Memoria
        for k, v in memoria.items():
            if k in txt:
                nit_found = v
                break
        
        # B. NIT en texto
        if not nit_found:
            posibles = extraer_posibles_nits(row['Texto_Completo'])
            for p in posibles:
                if p in mapa_nit:
                    nit_found = p
                    break
        
        # C. Fuzzy Name
        if not nit_found and len(txt) > 5:
            match, score = process.extractOne(txt, lista_nombres, scorer=fuzz.token_set_ratio)
            if score >= 88:
                nit_found = df_cartera[df_cartera['nombre_norm'] == match]['nit_norm'].iloc[0]

        # 2. Análisis Financiero
        if nit_found:
            nombre_cliente = mapa_nit.get(nit_found, "Cliente")
            facturas_open = df_cartera[df_cartera['nit_norm'] == nit_found]
            
            analisis = analizar_cliente(nombre_cliente, val, facturas_open)
            
            item.update(analisis)
            item['Cliente_Identificado'] = nombre_cliente
            item['NIT'] = nit_found
        else:
            item['Estado'] = '❓ NO IDENTIFICADO'
            item['Cliente_Identificado'] = ''
            item['Detalle_Operacion'] = 'Falta información para cruzar.'
            
        resultados.append(item)
        
    return pd.DataFrame(resultados)

# ======================================================================================
# --- 5. INTERFAZ GRÁFICA ---
# ======================================================================================

def main():
    st.title("🏦 Conciliador Financiero v9.1")
    st.markdown("**El Auditor Digital:** Identificación de facturas específicas, descuentos e impuestos.")
    
    # --- BARRA LATERAL ---
    with st.sidebar:
        st.header("Configuración")
        uploaded_file = st.file_uploader("📂 Cargar Planilla Banco (.xlsx)", type=["xlsx"])
        
        if st.button("🔄 Sincronizar Cartera Dropbox"):
            with st.spinner("Descargando..."):
                df_c = cargar_cartera_detalle()
                if not df_c.empty:
                    st.session_state['cartera'] = df_c
                    st.success(f"Cartera: {len(df_c)} facturas activas.")
                else:
                    st.error("Falló conexión Dropbox")

        st.divider()
        if 'cartera' in st.session_state:
            st.info(f"Facturas en memoria: {len(st.session_state['cartera'])}")
            st.dataframe(st.session_state['cartera'].head(3), use_container_width=True, hide_index=True)
        else:
            st.warning("⚠️ Carga la cartera primero")

    # --- PANEL PRINCIPAL ---
    if uploaded_file and 'cartera' in st.session_state:
        if st.button("🚀 EJECUTAR CONCILIACIÓN", type="primary", use_container_width=True):
            
            # 1. Leer Banco
            df_bancos = procesar_planilla_bancos(uploaded_file)
            if df_bancos.empty:
                st.error("El archivo de banco no parece válido o está vacío.")
                return

            # 2. Leer KB (Google Sheets)
            df_kb = pd.DataFrame()
            g_client = connect_to_google_sheets()
            if g_client:
                try:
                    sh = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"])
                    df_kb = pd.DataFrame(sh.worksheet("Knowledge_Base").get_all_records())
                except: pass

            # 3. Correr Motor
            df_res = correr_motor_inteligente(df_bancos, st.session_state['cartera'], df_kb)
            
            # 4. Estadísticas
            kpis = {
                'total_tx': len(df_res),
                'exactos': len(df_res[df_res['Estado'].str.contains('EXACTO')]),
                'descuentos': len(df_res[df_res['Estado'].str.contains('DESCUENTO')]),
                'impuestos': len(df_res[df_res['Estado'].str.contains('IMPUESTOS')]),
                'parciales': len(df_res[df_res['Estado'].str.contains('PARCIAL') | df_res['Estado'].str.contains('ABONO')]),
                'sin_id': len(df_res[df_res['Estado'].str.contains('NO IDENTIFICADO')])
            }
            
            # Mostrar Métricas
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Conciliación Perfecta", kpis['exactos'] + kpis['descuentos'])
            col2.metric("Con Impuestos", kpis['impuestos'])
            col3.metric("Abonos/Parciales", kpis['parciales'])
            col4.metric("Por Revisar", kpis['sin_id'])
            
            st.divider()
            
            # Tabla Interactiva
            st.subheader("📋 Vista Previa de Resultados")
            cols_show = ['FECHA', 'Valor_Banco', 'Cliente_Identificado', 'Estado', 'Facturas_Conciliadas', 'Detalle_Operacion']
            
            def color_estado(val):
                color = 'black'
                if '✅' in val: color = 'green'
                elif '💎' in val: color = 'blue'
                elif '❌' in val: color = 'red'
                elif '⚠️' in val: color = 'orange'
                return f'color: {color}'

            st.dataframe(
                df_res[cols_show].style.map(color_estado, subset=['Estado']),
                use_container_width=True,
                height=400
            )

            # 5. Descargar Excel
            excel_data = generar_excel_profesional(df_res, kpis)
            
            c_down, c_save = st.columns(2)
            with c_down:
                st.download_button(
                    label="💾 Descargar Reporte Conciliación (.xlsx)",
                    data=excel_data,
                    file_name=f"Conciliacion_Auditor_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            
            with c_save:
                if g_client and st.button("☁️ Guardar en Nube (Google Drive)"):
                    try:
                        sh = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"])
                        ws = sh.worksheet(st.secrets["google_sheets"]["tab_bancos_master"])
                        ws.clear()
                        # Formato string para fechas
                        df_save = df_res.copy().fillna('')
                        for c in df_save.select_dtypes(['datetime']): df_save[c] = df_save[c].astype(str)
                        set_with_dataframe(ws, df_save)
                        st.success("¡Datos sincronizados con éxito!")
                    except Exception as e:
                        st.error(f"Error subiendo a nube: {e}")

if __name__ == "__main__":
    main()
