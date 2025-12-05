# ======================================================================================
# ARCHIVO: pages/2_Motor_Conciliacion.py
# (Versión v19 - "Omnisciente Detallada": Facturas Visibles + Info Cliente Completa)
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
from gspread_dataframe import set_with_dataframe, get_as_dataframe
from oauth2client.service_account import ServiceAccountCredentials
import xlsxwriter
import itertools
import hashlib
from collections import defaultdict

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Motor Conciliación V19", page_icon="🕵️‍♂️", layout="wide")

# ======================================================================================
# --- 1. CONEXIONES Y UTILIDADES ---
# ======================================================================================

@st.cache_resource
def get_dbx_client(secrets_key):
    """Conexión persistente a Dropbox"""
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
    """Conexión persistente a Google Sheets"""
    try:
        scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        if "gcp_service_account" in st.secrets:
            creds = ServiceAccountCredentials.from_json_keyfile_dict(dict(st.secrets["gcp_service_account"]), scope)
            return gspread.authorize(creds)
        return None
    except Exception:
        return None

def download_from_dropbox(dbx, path):
    """Descarga binaria de un archivo en Dropbox"""
    try:
        _, res = dbx.files_download(path)
        return res.content
    except Exception as e:
        st.error(f"Error descargando {path}: {e}")
        return None

def generar_id_unico(row, index):
    """Huella digital única para evitar duplicados y rastrear filas"""
    try:
        fecha_str = str(row.get('FECHA', ''))
        val_str = str(row.get('Valor_Banco', 0))
        txt_str = str(row.get('Texto_Completo', '')).strip()
        # Incluimos index para asegurar unicidad absoluta en la sesión
        raw_str = f"{index}_{fecha_str}{val_str}{txt_str}"
        return hashlib.md5(raw_str.encode('utf-8')).hexdigest()
    except:
        return f"ID_ERROR_{index}"

def normalizar_texto_avanzado(texto):
    """Limpieza profunda para IA y Fuzzy Matching"""
    if not isinstance(texto, str): return ""
    texto = texto.upper().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'[^A-Z0-9\s]', ' ', texto) 
    
    palabras_basura = [
        'PAGO', 'TRANSF', 'TRANSFERENCIA', 'CONSIGNACION', 'ABONO', 'CTA', 'NIT', 
        'REF', 'FACTURA', 'OFI', 'SUC', 'ACH', 'PSE', 'NOMINA', 'PROVEEDOR', 
        'COMPRA', 'VENTA', 'VALOR', 'NETO', 'PLANILLA', 'S A', 'SAS', 'LTDA',
        'COLOMBIA', 'BANCOLOMBIA', 'DAVIVIENDA', 'BBVA', 'BOGOTA', 'OCCIDENTE',
        'NEQUI', 'DAVIPLATA', 'TRANSACCION', 'ELECTRONICA', 'RECIBIDO', 'DESDE', 'TERCERO',
        'CONSORCIO', 'UNION', 'TEMPORAL', 'GRP', 'GROUP'
    ]
    for p in palabras_basura:
        texto = re.sub(r'\b' + p + r'\b', ' ', texto)
    return ' '.join(texto.split())

def extraer_posibles_nits(texto):
    if not isinstance(texto, str): return []
    clean_txt = texto.replace('.', '').replace('-', '')
    return re.findall(r'\b\d{7,11}\b', clean_txt)

def limpiar_moneda_colombiana(val):
    if isinstance(val, (int, float)):
        return float(val) if pd.notnull(val) else 0.0
    s = str(val).strip()
    if not s or s.lower() == 'nan': return 0.0
    s = s.replace('$', '').replace('USD', '').replace('COP', '').strip()
    s = s.replace('.', '') 
    s = s.replace(',', '.') 
    try: return float(s)
    except ValueError: return 0.0

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
# --- 2. GENERADORES DE EXCEL (OPERATIVO Y GERENCIAL) ---
# ======================================================================================

def generar_excel_operativo(df):
    """Genera el excel de trabajo diario"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        
        # --- PREPARAR DATOS ---
        cols_export = [
            'FECHA', 'Valor_Banco', 'Texto_Completo', 
            'Cliente_Identificado', 'NIT', 
            'Facturas_Conciliadas', 'Detalle_Operacion', 'Estado', 
            'Diferencia', 'Tipo_Ajuste', 'Status_Gestion', 'Sugerencia_IA', 'ID_Unico'
        ]
        # Asegurar columnas
        for c in cols_export:
            if c not in df.columns: df[c] = ''
            
        df_export = df[cols_export].copy()
        
        # --- FORMATO ---
        df_export.to_excel(writer, index=False, sheet_name='Detalle_Conciliacion')
        worksheet = writer.sheets['Detalle_Conciliacion']
        workbook = writer.book
        
        header_fmt = workbook.add_format({'bold': True, 'fg_color': '#203764', 'font_color': 'white', 'border': 1})
        currency_fmt = workbook.add_format({'num_format': '$ #,##0.00'})
        
        for i, col in enumerate(df_export.columns):
            worksheet.write(0, i, col, header_fmt)
            width = 18
            if col in ['Texto_Completo', 'Cliente_Identificado', 'Detalle_Operacion']: width = 45
            worksheet.set_column(i, i, width)

        # Aplicar formato moneda
        col_val = df_export.columns.get_loc('Valor_Banco')
        worksheet.set_column(col_val, col_val, 20, currency_fmt)
        
    return output.getvalue()

def generar_reporte_gerencial(df):
    """Genera el reporte consolidado MES a MES"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Crear columna Mes para agrupación
        if not pd.api.types.is_datetime64_any_dtype(df['FECHA']):
            df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
        
        df['Mes_Año'] = df['FECHA'].dt.strftime('%Y-%m')
        
        # 1. HOJA RESUMEN (Pivot Table)
        pivot = df.pivot_table(
            index='Mes_Año', 
            columns='Estado', 
            values='Valor_Banco', 
            aggfunc='count', 
            fill_value=0
        )
        pivot.to_excel(writer, sheet_name='Resumen_Mensual')
        
        # Formato Resumen
        ws_res = writer.sheets['Resumen_Mensual']
        style_header = workbook.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white'})
        ws_res.write(0, 0, "Periodo", style_header)
        
        # 2. HOJA PENDIENTES
        df_pend = df[
            (df['Cliente_Identificado'] == "") | 
            (df['Status_Gestion'] == "PENDIENTE")
        ].copy()
        df_pend.to_excel(writer, sheet_name='Pendientes_Gestion', index=False)
        
        # 3. HOJA TOTAL HISTÓRICO
        df.to_excel(writer, sheet_name='Data_Completa', index=False)
        
    return output.getvalue()

# ======================================================================================
# --- 3. CARGA DE DATOS ---
# ======================================================================================

@st.cache_data(ttl=600)
def cargar_cartera_dropbox():
    """Carga Facturas Abiertas desde Dropbox"""
    dbx = get_dbx_client("dropbox")
    if not dbx: return pd.DataFrame()
    
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
        
        return df[df['Importe'] > 100].copy()
    except Exception as e:
        st.error(f"Error estructura cartera: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=600)
def cargar_historico_dropbox():
    """Carga Historial Consolidado"""
    dbx = get_dbx_client("dropbox")
    if not dbx: return pd.DataFrame()
    
    content = download_from_dropbox(dbx, '/data/planilla_bancos.xlsx')
    if not content: return pd.DataFrame()
    
    try:
        df = pd.read_excel(BytesIO(content))
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        col_cliente = 'EMPRESA' if 'EMPRESA' in df.columns else None
        cols_texto = []
        if 'TIPO DE TRANSACCION' in df.columns: cols_texto.append('TIPO DE TRANSACCION')
        if 'BANCO REFRENCIA INTERNA' in df.columns: cols_texto.append('BANCO REFRENCIA INTERNA')
        if 'DESTINO' in df.columns: cols_texto.append('DESTINO')
        
        if not col_cliente:
             col_cliente = next((c for c in df.columns if 'CLIENTE' in c or 'IDENTIFICADO' in c), None)
        
        if col_cliente:
            df['HISTORIA_CLIENTE'] = df[col_cliente].astype(str).str.strip()
            if cols_texto:
                df['HISTORIA_TEXTO_RAW'] = df[cols_texto].fillna('').astype(str).agg(' '.join, axis=1)
            else:
                col_gen = next((c for c in df.columns if 'TEXTO' in c or 'DESCRIPCION' in c), 'HISTORIA_CLIENTE')
                df['HISTORIA_TEXTO_RAW'] = df[col_gen].astype(str)
                
            df['HISTORIA_TEXTO'] = df['HISTORIA_TEXTO_RAW'].apply(normalizar_texto_avanzado)
            return df
        else:
            return pd.DataFrame()
            
    except Exception as e:
        st.error(f"Error leyendo Histórico Dropbox: {e}")
        return pd.DataFrame()

def procesar_archivo_manual(uploaded_file):
    """Procesa el archivo del día a día"""
    try:
        df_temp = pd.read_excel(uploaded_file, nrows=15, header=None)
        header_idx = 0
        for idx, row in df_temp.iterrows():
            row_str = row.astype(str).str.upper().values
            if 'FECHA' in row_str and ('VALOR' in row_str or 'IMPORTE' in row_str or 'CREDITO' in row_str):
                header_idx = idx
                break
        
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, header=header_idx)
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        if 'FECHA' not in df.columns: 
             cols_fecha = [c for c in df.columns if 'FECHA' in c]
             if cols_fecha: df.rename(columns={cols_fecha[0]: 'FECHA'}, inplace=True)

        if 'VALOR' not in df.columns:
            cols_valor = [c for c in df.columns if 'VALOR' in c or 'CREDITO' in c or 'IMPORTE' in c]
            if cols_valor: df.rename(columns={cols_valor[0]: 'VALOR'}, inplace=True)
        
        df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
        df = df.dropna(subset=['FECHA'])
        
        if 'VALOR' in df.columns:
            df['Valor_Banco'] = df['VALOR'].apply(limpiar_moneda_colombiana)
        else:
            df['Valor_Banco'] = 0.0

        cols_exclude = ['FECHA', 'VALOR', 'Valor_Banco', 'SALDO', 'DEBITO']
        cols_txt = [c for c in df.columns if c not in cols_exclude]
        df['Texto_Completo'] = df[cols_txt].fillna('').astype(str).agg(' '.join, axis=1)
        df['Texto_Norm'] = df['Texto_Completo'].apply(normalizar_texto_avanzado)
        
        mask_zero = df['Valor_Banco'] == 0
        df.loc[mask_zero, 'Valor_Banco'] = df.loc[mask_zero, 'Texto_Completo'].apply(extraer_dinero_de_texto)
        
        # Index artificial para generar ID único
        df = df.reset_index(drop=True)
        df['ID_Unico'] = [generar_id_unico(row, idx) for idx, row in df.iterrows()]
        
        return df
    except Exception as e:
        st.error(f"Error procesando archivo manual: {e}")
        return pd.DataFrame()

# ======================================================================================
# --- 4. LÓGICA OMNISCIENTE ---
# ======================================================================================

def analizar_deuda_cliente(nombre_cliente, nit_cliente, valor_pago, df_cartera):
    res = {
        'Estado': '⚠️ SIN COINCIDENCIA VALOR',
        'Facturas_Conciliadas': '',
        'Detalle_Operacion': '',
        'Diferencia': 0,
        'Tipo_Ajuste': 'Ninguno',
        'Impuesto_Est': 0
    }
    
    facturas = df_cartera[df_cartera['nit_norm'] == nit_cliente]
    if facturas.empty:
        facturas = df_cartera[df_cartera['NombreCliente'].str.contains(nombre_cliente, case=False, na=False)]
    
    if facturas.empty:
        res['Detalle_Operacion'] = "Cliente identificado por Nombre/NIT, pero no tiene facturas abiertas en cartera."
        res['Diferencia'] = valor_pago * -1
        return res

    facturas_list = facturas[['Numero', 'Importe']].to_dict('records')
    total_deuda = sum(f['Importe'] for f in facturas_list)
    
    # 1. MATCH EXACTO TOTAL
    if abs(valor_pago - total_deuda) < 1000:
        res['Estado'] = '✅ MATCH EXACTO (TOTAL)'
        res['Facturas_Conciliadas'] = 'TODAS'
        res['Detalle_Operacion'] = f"Pago total de {len(facturas_list)} facturas pendientes."
        return res
        
    # 2. MATCH FACTURAS ESPECÍFICAS (Combinatoria)
    found = False
    # Probamos combinaciones de 1 a 4 facturas
    for r in range(1, 5): 
        if found: break
        for combo in itertools.combinations(facturas_list, r):
            suma_combo = sum(c['Importe'] for c in combo)
            numeros = ", ".join([str(c['Numero']) for c in combo])
            
            # Match Exacto de Subconjunto
            if abs(valor_pago - suma_combo) < 500:
                res['Estado'] = '✅ FACTURAS ESPECÍFICAS'
                res['Facturas_Conciliadas'] = numeros
                res['Detalle_Operacion'] = f"Suma exacta de: {numeros}"
                found = True
                break
            
            # Match con Descuento Pronto Pago (~3%)
            if abs(valor_pago - (suma_combo * 0.97)) < 2000:
                res['Estado'] = '💎 CONCILIADO C/DCTO'
                res['Facturas_Conciliadas'] = numeros
                res['Detalle_Operacion'] = f"Posible Dcto Pronto Pago sobre: {numeros}"
                res['Tipo_Ajuste'] = "Descuento Pronto Pago"
                found = True
                break

    if found: return res

    # 3. IMPUESTOS
    base_est = total_deuda / 1.19
    rete_fuente = base_est * 0.025
    rete_iva = (base_est * 0.19) * 0.15
    pago_imptos = total_deuda - rete_fuente - rete_iva
    
    if abs(valor_pago - pago_imptos) < 5000:
        res['Estado'] = '🏢 CONCILIADO (IMPUESTOS)'
        res['Impuesto_Est'] = rete_fuente + rete_iva
        res['Detalle_Operacion'] = "Coincide monto total menos retenciones estimadas."
        res['Facturas_Conciliadas'] = 'TODAS (Probable)'
        return res

    # 4. ABONO PARCIAL
    res['Estado'] = '⚠️ ABONO / PARCIAL'
    res['Diferencia'] = total_deuda - valor_pago
    res['Detalle_Operacion'] = f"No cruza exacto. Deuda Total: ${total_deuda:,.0f}. Diferencia: ${res['Diferencia']:,.0f}"
    
    return res

def motor_omnisciente(df_manual, df_cartera, df_historico, df_kb):
    st.info("🧠 Procesando: Memoria Histórica + Knowledge Base + Cartera...")
    
    # 1. MEMORIA UNIFICADA (Historial + KB)
    memoria_unificada = {}
    
    if not df_kb.empty:
        for _, row in df_kb.iterrows():
            try: memoria_unificada[str(row[0]).strip()] = str(row[1]).strip()
            except: pass
            
    if not df_historico.empty:
        for _, row in df_historico.iterrows():
            txt = str(row.get('HISTORIA_TEXTO', ''))
            cli = str(row.get('HISTORIA_CLIENTE', ''))
            if len(txt) > 5 and cli != '' and cli.lower() != 'nan':
                memoria_unificada[txt] = cli

    # 2. INDICE PALABRAS CLAVE (Cartera)
    word_to_nits = defaultdict(set)
    mapa_nit_nombre = df_cartera.groupby('nit_norm')['NombreCliente'].first().to_dict()
    
    for _, row in df_cartera.iterrows():
        nombre_norm = str(row['nombre_norm'])
        nit = row['nit_norm']
        for w in nombre_norm.split():
            if len(w) > 3: word_to_nits[w].add(nit)
            
    unique_keywords = {w: list(ns)[0] for w, ns in word_to_nits.items() if len(ns) == 1}

    # 3. ITERACIÓN
    resultados = []
    progress_bar = st.progress(0)
    
    for idx, row in df_manual.iterrows():
        progress_bar.progress((idx + 1) / len(df_manual))
        
        item = row.to_dict()
        txt_norm = row['Texto_Norm']
        val_pago = row['Valor_Banco']
        
        cliente_detectado = None
        nit_detectado = None
        metodo_deteccion = ""
        
        # A. MEMORIA (Prioridad Absoluta)
        if txt_norm in memoria_unificada:
            cliente_detectado = memoria_unificada[txt_norm]
            metodo_deteccion = "🧠 Memoria / KB"
            match_nit = df_cartera[df_cartera['NombreCliente'] == cliente_detectado]
            if not match_nit.empty:
                nit_detectado = match_nit.iloc[0]['nit_norm']
        
        # B. CARTERA (Si no hay memoria)
        if not cliente_detectado:
            # B1. NIT en Texto
            nits_found = extraer_posibles_nits(row['Texto_Completo'])
            for n in nits_found:
                if n in mapa_nit_nombre:
                    nit_detectado = n
                    cliente_detectado = mapa_nit_nombre[n]
                    metodo_deteccion = "🆔 NIT encontrado en Texto"
                    break
            
            # B2. Palabra Clave
            if not cliente_detectado:
                for palabra in txt_norm.split():
                    if palabra in unique_keywords:
                        nit_detectado = unique_keywords[palabra]
                        cliente_detectado = mapa_nit_nombre[nit_detectado]
                        metodo_deteccion = f"🔑 Palabra Clave '{palabra}'"
                        break
            
            # B3. Fuzzy
            if not cliente_detectado:
                posibles_nombres = df_cartera['nombre_norm'].unique()
                match, score = process.extractOne(txt_norm, posibles_nombres, scorer=fuzz.token_set_ratio)
                if score >= 85:
                    cliente_detectado = df_cartera[df_cartera['nombre_norm'] == match]['NombreCliente'].iloc[0]
                    nit_detectado = df_cartera[df_cartera['nombre_norm'] == match]['nit_norm'].iloc[0]
                    metodo_deteccion = f"≈ Similitud Nombre ({score}%)"

        # C. RESULTADO
        item['Cliente_Identificado'] = cliente_detectado if cliente_detectado else ""
        item['NIT'] = nit_detectado if nit_detectado else ""
        item['Sugerencia_IA'] = metodo_deteccion
        item['Status_Gestion'] = 'PENDIENTE'
        
        if cliente_detectado and nit_detectado:
            analisis = analizar_deuda_cliente(cliente_detectado, nit_detectado, val_pago, df_cartera)
            item.update(analisis)
        else:
            # Radar Monto (Último recurso)
            match_monto = df_cartera[
                (df_cartera['Importe'] >= val_pago - 100) & 
                (df_cartera['Importe'] <= val_pago + 100)
            ]
            if not match_monto.empty:
                cand = match_monto.iloc[0]
                item['Estado'] = '💡 SUGERENCIA MONTO'
                item['Sugerencia_IA'] = "Coincidencia solo por Valor"
                item['Cliente_Identificado'] = cand['NombreCliente'] # Sugerencia visual
                item['NIT'] = cand['nit_norm']
                item['Detalle_Operacion'] = f"Monto coincide con Factura {cand['Numero']} de {cand['NombreCliente']}"
                item['Facturas_Conciliadas'] = str(cand['Numero'])
            else:
                item['Estado'] = '❓ NO IDENTIFICADO'
                item['Detalle_Operacion'] = "Sin coincidencias claras."

        resultados.append(item)
        
    return pd.DataFrame(resultados)

# ======================================================================================
# --- 5. INTERFAZ PRINCIPAL ---
# ======================================================================================

def main():
    st.title("📊 Conciliación Omnisciente V19")
    st.markdown("Plataforma Integral: Visión Completa de Facturas, NITs y Detalles de Conciliación")

    # --- BARRA LATERAL: CARGA DE DATOS ---
    with st.sidebar:
        st.header("1. Fuentes de Datos (Dropbox)")
        
        if st.button("🔄 Cargar Cartera (Dropbox)"):
            with st.spinner("Descargando..."):
                df_c = cargar_cartera_dropbox()
                if not df_c.empty:
                    st.session_state['cartera'] = df_c
                    st.success(f"Cartera: {len(df_c)} regs")
                else: st.error("Error Cartera")
        
        if st.button("📚 Cargar Historial (Dropbox)"):
            with st.spinner("Descargando Historial..."):
                df_h = cargar_historico_dropbox()
                if not df_h.empty:
                    st.session_state['historico'] = df_h
                    st.success(f"Historial Cargado: {len(df_h)} regs")
                else: st.error("Error Historial")
                
        if 'cartera' in st.session_state: st.info(f"📂 Cartera Activa: {len(st.session_state['cartera'])}")
        if 'historico' in st.session_state: st.info(f"🧠 Memoria Activa: {len(st.session_state['historico'])}")
        st.divider()

    # --- PANEL SUPERIOR: OPERACIÓN ---
    st.subheader("2. Operación Diaria")
    uploaded_file = st.file_uploader("Sube el Archivo Manual Diario (.xlsx)", type=["xlsx"])

    if uploaded_file and 'cartera' in st.session_state:
        if st.button("🚀 EJECUTAR MOTOR IA (ANÁLISIS COMPLETO)", type="primary", use_container_width=True):
            
            # 1. Leer Manual
            df_manual = procesar_archivo_manual(uploaded_file)
            if df_manual.empty:
                st.error("Error leyendo archivo manual.")
                return
            
            # 2. Leer KB Google Sheets
            df_kb = pd.DataFrame()
            g_client = connect_to_google_sheets()
            if g_client:
                try:
                    sh = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"])
                    try:
                        df_kb = get_as_dataframe(sh.worksheet("Knowledge_Base"), header=None)
                    except:
                        st.warning("KB vacía, se creará al guardar.")
                except: pass
            
            # 3. Correr Motor
            df_res = motor_omnisciente(
                df_manual, 
                st.session_state['cartera'], 
                st.session_state.get('historico', pd.DataFrame()), 
                df_kb
            )
            st.session_state['resultado_final'] = df_res

    # --- SECCIÓN DE RESULTADOS Y FILTROS ---
    if 'resultado_final' in st.session_state:
        df_master = st.session_state['resultado_final']
        
        # --- FILTROS (SIDEBAR ADICIONAL O EXPANDER) ---
        with st.expander("🔎 FILTROS AVANZADOS (Mes, Fechas, Estado)", expanded=True):
            col_f1, col_f2, col_f3, col_f4 = st.columns(4)
            
            # Filtro Fechas
            min_date = df_master['FECHA'].min()
            max_date = df_master['FECHA'].max()
            date_range = col_f1.date_input("Rango de Fechas", [min_date, max_date])
            
            # Filtro Estado Conciliación
            estados_disponibles = sorted(df_master['Estado'].unique())
            sel_estado = col_f2.multiselect("Estado Conciliación", estados_disponibles, default=estados_disponibles)
            
            # Filtro Gestión
            gestiones_disponibles = df_master['Status_Gestion'].unique()
            sel_gestion = col_f3.multiselect("Estado Gestión", gestiones_disponibles, default=gestiones_disponibles)
            
            # Filtro Sugerencia IA
            sugerencias = sorted(df_master['Sugerencia_IA'].astype(str).unique())
            sel_ia = col_f4.multiselect("Tipo Hallazgo IA", sugerencias, default=sugerencias)

        # APLICAR FILTROS
        mask_fecha = (df_master['FECHA'].dt.date >= date_range[0]) & (df_master['FECHA'].dt.date <= date_range[1]) if len(date_range) == 2 else True
        mask_estado = df_master['Estado'].isin(sel_estado)
        mask_gestion = df_master['Status_Gestion'].isin(sel_gestion)
        mask_ia = df_master['Sugerencia_IA'].astype(str).isin(sel_ia)
        
        df_view = df_master[mask_fecha & mask_estado & mask_gestion & mask_ia].copy()
        
        st.divider()
        
        # --- KPIS VISUALES ---
        kpis = {
            'total': len(df_view),
            'pendientes': len(df_view[df_view['Status_Gestion'] == 'PENDIENTE']),
            'monto': df_view['Valor_Banco'].sum()
        }
        c1, c2, c3 = st.columns(3)
        c1.metric("Registros Filtrados", kpis['total'])
        c2.metric("Pendientes de Gestión", kpis['pendientes'], delta_color="inverse")
        c3.metric("Monto Total Vista", f"${kpis['monto']:,.0f}")

        # --- EDITOR DE DATOS (AQUÍ ESTÁ LA MEJORA VISUAL) ---
        st.write("### 📝 Detalle de Conciliación (Edita aquí)")
        lista_clientes = sorted(st.session_state['cartera']['NombreCliente'].unique().tolist())
        
        # Configuración de Columnas para máxima visibilidad
        col_config = {
            "Status_Gestion": st.column_config.SelectboxColumn("Gestión", options=['PENDIENTE', 'REGISTRADA'], required=True, width="small"),
            "Cliente_Identificado": st.column_config.SelectboxColumn("Cliente", options=lista_clientes, width="large"),
            "Valor_Banco": st.column_config.NumberColumn("Valor Pago", format="$ %d", width="small"),
            "FECHA": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY", width="small"),
            "NIT": st.column_config.TextColumn("NIT Detectado", width="medium"),
            "Facturas_Conciliadas": st.column_config.TextColumn("Facturas Cruzadas", width="medium", help="Facturas que suman el valor del pago"),
            "Detalle_Operacion": st.column_config.TextColumn("Explicación IA", width="large"),
            "Estado": st.column_config.TextColumn("Estado Match", width="medium"),
            "Sugerencia_IA": st.column_config.TextColumn("Método Detección", width="medium")
        }
        
        # Seleccionamos y ordenamos las columnas que quieres ver
        cols_view = [
            'Status_Gestion', 
            'FECHA', 
            'Valor_Banco', 
            'Cliente_Identificado', 
            'NIT', 
            'Facturas_Conciliadas', 
            'Detalle_Operacion', 
            'Estado', 
            'Sugerencia_IA', 
            'ID_Unico'
        ]
        
        edited_df = st.data_editor(
            df_view[cols_view], 
            use_container_width=True, 
            column_config=col_config, 
            key="editor_filtrado",
            num_rows="dynamic",
            height=600
        )
        
        # --- SINCRONIZACIÓN DE CAMBIOS ---
        # Si el usuario edita la vista filtrada, actualizamos el DF Master usando ID_Unico
        if not edited_df.equals(df_view[cols_view]):
            for idx, row in edited_df.iterrows():
                id_unico = row['ID_Unico']
                # Actualizar campos clave en el master
                idx_master = df_master[df_master['ID_Unico'] == id_unico].index
                if not idx_master.empty:
                    st.session_state['resultado_final'].loc[idx_master, 'Status_Gestion'] = row['Status_Gestion']
                    st.session_state['resultado_final'].loc[idx_master, 'Cliente_Identificado'] = row['Cliente_Identificado']

        # --- BOTONES DE ACCIÓN ---
        st.divider()
        c_excel, c_informe, c_save = st.columns(3)
        
        with c_excel:
            # Descarga Operativa
            excel_op = generar_excel_operativo(edited_df)
            st.download_button("💾 Descargar Vista Actual (Operativo)", data=excel_op, file_name="Conciliacion_Operativa.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            
        with c_informe:
            # Descarga Gerencial
            excel_mg = generar_reporte_gerencial(st.session_state['resultado_final'])
            st.download_button("📊 Descargar Informe Gerencial (Mes a Mes)", data=excel_mg, file_name="Reporte_Consolidado_Mensual.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            
        with c_save:
            if st.button("☁️ GUARDAR Y ENTRENAR IA", type="primary", use_container_width=True):
                g_client = connect_to_google_sheets()
                if not g_client:
                    st.error("Error conexión Google Sheets.")
                else:
                    try:
                        sh = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"])
                        
                        # 1. Guardar Maestro
                        ws_master = sh.worksheet(st.secrets["google_sheets"]["tab_bancos_master"])
                        df_final_save = st.session_state['resultado_final'].copy()
                        df_final_save = df_final_save.fillna('')
                        df_final_save['FECHA'] = df_final_save['FECHA'].astype(str)
                        ws_master.clear()
                        set_with_dataframe(ws_master, df_final_save)
                        
                        # 2. Entrenar IA (KB)
                        try: ws_kb = sh.worksheet("Knowledge_Base")
                        except: ws_kb = sh.add_worksheet(title="Knowledge_Base", rows=1000, cols=2)

                        nuevos_registros = df_final_save[
                            (df_final_save['Status_Gestion'] == 'REGISTRADA') & 
                            (df_final_save['Cliente_Identificado'] != '')
                        ]
                        
                        if not nuevos_registros.empty:
                            data_kb = []
                            for _, r in nuevos_registros.iterrows():
                                txt_norm = normalizar_texto_avanzado(str(r['Texto_Completo']))
                                cli = str(r['Cliente_Identificado']).strip()
                                if len(txt_norm) > 5 and cli:
                                    data_kb.append([txt_norm, cli])
                            
                            if data_kb:
                                ws_kb.append_rows(data_kb)
                                st.toast(f"🧠 IA aprendió {len(data_kb)} nuevos patrones.")
                        
                        st.success("✅ Guardado Exitoso y Aprendizaje Completado")
                    except Exception as e:
                        st.error(f"Error guardando: {e}")

if __name__ == "__main__":
    main()
