# ======================================================================================
# ARCHIVO: pages/2_Motor_Conciliacion.py
# (Versión v17 - "Omnisciente Adaptativo": Fusión Dropbox + GSheets + IA Entrenable)
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
st.set_page_config(page_title="Motor Conciliación Pro v17", page_icon="👁️", layout="wide")

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
    """Huella digital única para evitar duplicados"""
    try:
        fecha_str = str(row.get('FECHA', ''))
        val_str = str(row.get('Valor_Banco', 0))
        txt_str = str(row.get('Texto_Completo', '')).strip()
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
# --- 2. GENERADOR DE EXCEL PROFESIONAL ---
# ======================================================================================

def generar_excel_profesional(df, resumen_kpis):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        
        # --- HOJA DASHBOARD ---
        workbook = writer.book
        sheet_resumen = workbook.add_worksheet("Dashboard")
        sheet_resumen.hide_gridlines(2)
        
        style_title = workbook.add_format({'bold': True, 'font_size': 16, 'font_color': '#1F497D'})
        style_kpi_label = workbook.add_format({'bold': True, 'bg_color': '#E7E6E6', 'border': 1})
        style_kpi_value = workbook.add_format({'bold': True, 'num_format': '#,##0', 'border': 1, 'align': 'center'})
        
        sheet_resumen.write('B2', "RESUMEN DE CONCILIACIÓN - OMNISCIENTE", style_title)
        
        kpis = [
            ("Total Movimientos", resumen_kpis.get('total_tx', 0)),
            ("Conciliados (Match Exacto)", resumen_kpis.get('exactos', 0)),
            ("Conciliados (Con Descuento)", resumen_kpis.get('descuentos', 0)),
            ("Conciliados (Impuestos)", resumen_kpis.get('impuestos', 0)),
            ("Parciales / Abonos", resumen_kpis.get('parciales', 0)),
            ("Histórico / IA", resumen_kpis.get('historico', 0)),
            ("Sin Identificar", resumen_kpis.get('sin_id', 0))
        ]
        
        row = 4
        for label, val in kpis:
            sheet_resumen.write(row, 1, label, style_kpi_label)
            sheet_resumen.write(row, 2, val, style_kpi_value)
            row += 1

        # --- HOJA DETALLE ---
        cols_export = [
            'FECHA', 'Valor_Banco', 'Texto_Completo', 'Cliente_Identificado', 
            'NIT', 'Estado', 'Facturas_Conciliadas', 'Detalle_Operacion', 
            'Diferencia', 'Tipo_Ajuste', 'Status_Gestion', 'Sugerencia_IA', 'ID_Unico'
        ]
        for c in cols_export:
            if c not in df.columns: df[c] = ''
            
        df_export = df[cols_export].copy()
        df_export.to_excel(writer, index=False, sheet_name='Detalle_Conciliacion')
        worksheet = writer.sheets['Detalle_Conciliacion']
        
        header_format = workbook.add_format({
            'bold': True, 'text_wrap': True, 'valign': 'vcenter', 'align': 'center',
            'fg_color': '#203764', 'font_color': 'white', 'border': 1
        })
        currency_fmt = workbook.add_format({'num_format': '$ #,##0.00', 'border': 1})
        text_fmt = workbook.add_format({'text_wrap': False, 'border': 1})
        
        # Colores Condicionales
        fmt_green = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
        fmt_blue = workbook.add_format({'bg_color': '#BDD7EE', 'font_color': '#1F497D'})
        fmt_red = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
        fmt_yellow = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C5700'})
        fmt_purple = workbook.add_format({'bg_color': '#E1D5E7', 'font_color': '#604A7B'})

        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(0, col_num, value, header_format)
            col_name = str(value).upper()
            width = 20
            if "CLIENTE" in col_name or "TEXTO" in col_name: width = 45
            elif "DETALLE" in col_name: width = 50
            
            worksheet.set_column(col_num, col_num, width, text_fmt)
            if any(x in col_name for x in ['VALOR', 'DIFERENCIA']):
                worksheet.set_column(col_num, col_num, width, currency_fmt)

        try:
            col_idx = df_export.columns.get_loc('Estado')
            col_letter = chr(65 + col_idx)
            last_row = len(df_export) + 1
            rng = f"{col_letter}2:{col_letter}{last_row}"
            
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'EXACTO', 'format': fmt_green})
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'DESCUENTO', 'format': fmt_blue})
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'REVISAR', 'format': fmt_red})
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'ABONO', 'format': fmt_yellow})
            worksheet.conditional_format(rng, {'type': 'text', 'criteria': 'containing', 'value': 'IA', 'format': fmt_purple})
        except: pass

        worksheet.freeze_panes(1, 0)
        
    return output.getvalue()

# ======================================================================================
# --- 3. CARGA DE DATOS (AQUÍ ESTÁ LA SOLUCIÓN AL ERROR DE COLUMNAS) ---
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
    """
    Carga Historial Consolidado adaptándose a tus columnas específicas:
    FECHA, SUCURSAL BANCO, TIPO DE TRANSACCION, CUENTA, EMPRESA, VALOR...
    """
    dbx = get_dbx_client("dropbox")
    if not dbx: return pd.DataFrame()
    
    content = download_from_dropbox(dbx, '/data/planilla_bancos.xlsx')
    if not content: return pd.DataFrame()
    
    try:
        df = pd.read_excel(BytesIO(content))
        # Normalizar nombres de columnas a mayúsculas y sin espacios extra
        df.columns = [str(c).strip().upper() for c in df.columns]
        
        # --- LÓGICA DE MAPEO DE COLUMNAS PERSONALIZADA ---
        
        # 1. Identificar CLIENTE (Tu columna es 'EMPRESA')
        col_cliente = 'EMPRESA' if 'EMPRESA' in df.columns else None
        
        # 2. Identificar DESCRIPCIÓN (Combinamos TIPO y REFERENCIA)
        cols_texto = []
        if 'TIPO DE TRANSACCION' in df.columns: cols_texto.append('TIPO DE TRANSACCION')
        if 'BANCO REFRENCIA INTERNA' in df.columns: cols_texto.append('BANCO REFRENCIA INTERNA')
        if 'DESTINO' in df.columns: cols_texto.append('DESTINO')
        
        # Si no encuentro tus columnas específicas, busco las genéricas
        if not col_cliente:
             col_cliente = next((c for c in df.columns if 'CLIENTE' in c or 'IDENTIFICADO' in c), None)
        
        # --- PROCESAMIENTO ---
        
        if col_cliente:
            # Normalizamos el cliente histórico
            df['HISTORIA_CLIENTE'] = df[col_cliente].astype(str).str.strip()
            
            # Construimos el Texto Histórico concatenando las columnas detectadas
            if cols_texto:
                df['HISTORIA_TEXTO_RAW'] = df[cols_texto].fillna('').astype(str).agg(' '.join, axis=1)
            else:
                # Fallback si no hay columnas específicas
                col_gen = next((c for c in df.columns if 'TEXTO' in c or 'DESCRIPCION' in c), 'HISTORIA_CLIENTE')
                df['HISTORIA_TEXTO_RAW'] = df[col_gen].astype(str)
                
            # Aplicamos normalización avanzada para que la IA entienda
            df['HISTORIA_TEXTO'] = df['HISTORIA_TEXTO_RAW'].apply(normalizar_texto_avanzado)
            
            return df
        else:
            st.warning("No se encontró la columna 'EMPRESA' ni similar en el Histórico.")
            st.write("Columnas detectadas:", df.columns.tolist())
            return pd.DataFrame()
            
    except Exception as e:
        st.error(f"Error leyendo Histórico Dropbox: {e}")
        return pd.DataFrame()

def procesar_archivo_manual(uploaded_file):
    """Procesa el archivo del día a día"""
    try:
        # Detectar cabecera dinámicamente
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

        # Crear Texto Completo
        cols_exclude = ['FECHA', 'VALOR', 'Valor_Banco', 'SALDO', 'DEBITO']
        cols_txt = [c for c in df.columns if c not in cols_exclude]
        df['Texto_Completo'] = df[cols_txt].fillna('').astype(str).agg(' '.join, axis=1)
        df['Texto_Norm'] = df['Texto_Completo'].apply(normalizar_texto_avanzado)
        
        # Rescatar valores cero
        mask_zero = df['Valor_Banco'] == 0
        df.loc[mask_zero, 'Valor_Banco'] = df.loc[mask_zero, 'Texto_Completo'].apply(extraer_dinero_de_texto)
        
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
        'Estado': '⚠️ SIN COINCIDENCIA DE VALOR',
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
        res['Detalle_Operacion'] = "Cliente identificado, pero sin cartera pendiente."
        res['Diferencia'] = valor_pago * -1
        return res

    facturas_list = facturas[['Numero', 'Importe']].to_dict('records')
    total_deuda = sum(f['Importe'] for f in facturas_list)
    
    # 1. MATCH EXACTO TOTAL
    if abs(valor_pago - total_deuda) < 1000:
        res['Estado'] = '✅ MATCH EXACTO (TOTAL)'
        res['Facturas_Conciliadas'] = 'TODAS'
        res['Detalle_Operacion'] = f"Pago total deuda."
        return res
        
    # 2. MATCH FACTURAS ESPECÍFICAS
    found = False
    for r in range(1, 4):
        if found: break
        for combo in itertools.combinations(facturas_list, r):
            suma_combo = sum(c['Importe'] for c in combo)
            numeros = ", ".join([str(c['Numero']) for c in combo])
            
            if abs(valor_pago - suma_combo) < 500:
                res['Estado'] = '✅ FACTURAS ESPECÍFICAS'
                res['Facturas_Conciliadas'] = numeros
                res['Detalle_Operacion'] = "Suma exacta de facturas."
                found = True
                break
            
            if abs(valor_pago - (suma_combo * 0.97)) < 2000:
                res['Estado'] = '💎 CONCILIADO C/DCTO'
                res['Facturas_Conciliadas'] = numeros
                res['Detalle_Operacion'] = "Pago con aprox 3% descuento."
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
        res['Detalle_Operacion'] = "Coincide menos retenciones."
        return res

    # 4. ABONO PARCIAL
    res['Estado'] = '⚠️ ABONO / PARCIAL'
    res['Diferencia'] = total_deuda - valor_pago
    res['Detalle_Operacion'] = f"Abono a deuda. Resta ${total_deuda - valor_pago:,.0f}"
    
    return res

def motor_omnisciente(df_manual, df_cartera, df_historico, df_kb):
    """
    EL CEREBRO. Cruza:
    1. Knowledge Base (Google Sheets)
    2. Historial (Dropbox - Ahora corregido)
    3. Cartera (NIT, Palabras Clave, Fuzzy)
    """
    st.info("🧠 Procesando: Memoria Histórica + Knowledge Base + Cartera...")
    
    # 1. MEMORIA UNIFICADA (Historial + KB)
    memoria_unificada = {}
    
    # Cargar KB (Entrenamientos previos)
    if not df_kb.empty:
        for _, row in df_kb.iterrows():
            try: memoria_unificada[str(row[0]).strip()] = str(row[1]).strip()
            except: pass
            
    # Cargar Historial (Dropbox)
    if not df_historico.empty:
        for _, row in df_historico.iterrows():
            # Usamos las columnas normalizadas que creamos en cargar_historico_dropbox
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
            # Buscar NIT en cartera si existe
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
                    metodo_deteccion = "🆔 NIT en Texto"
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
                    metodo_deteccion = f"≈ Similitud ({score}%)"

        # C. RESULTADO
        item['Cliente_Identificado'] = cliente_detectado if cliente_detectado else ""
        item['NIT'] = nit_detectado if nit_detectado else ""
        item['Sugerencia_IA'] = metodo_deteccion
        item['Status_Gestion'] = 'PENDIENTE'
        
        if cliente_detectado and nit_detectado:
            analisis = analizar_deuda_cliente(cliente_detectado, nit_detectado, val_pago, df_cartera)
            item.update(analisis)
        else:
            # Radar Monto
            match_monto = df_cartera[
                (df_cartera['Importe'] >= val_pago - 100) & 
                (df_cartera['Importe'] <= val_pago + 100)
            ]
            if not match_monto.empty:
                cand = match_monto.iloc[0]
                item['Estado'] = '💡 SUGERENCIA MONTO'
                item['Sugerencia_IA'] = f"Monto coincide con {cand['NombreCliente']}"
                item['Detalle_Operacion'] = f"Posible Factura {cand['Numero']}"
            else:
                item['Estado'] = '❓ NO IDENTIFICADO'
                item['Detalle_Operacion'] = "Sin coincidencias claras."

        resultados.append(item)
        
    return pd.DataFrame(resultados)

# ======================================================================================
# --- 5. INTERFAZ PRINCIPAL ---
# ======================================================================================

def main():
    st.title("👁️ Motor Omnisciente Total (v17)")
    st.markdown("Fusión: Dropbox + Google Sheets + IA + Columnas Personalizadas")

    # --- BARRA LATERAL ---
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
            with st.spinner("Descargando Historial con tus columnas..."):
                df_h = cargar_historico_dropbox()
                if not df_h.empty:
                    st.session_state['historico'] = df_h
                    st.success(f"Historial Cargado: {len(df_h)} regs")
                else: st.error("Error Historial")
                
        if 'cartera' in st.session_state: st.info(f"📂 Cartera Activa: {len(st.session_state['cartera'])}")
        if 'historico' in st.session_state: st.info(f"🧠 Memoria Activa: {len(st.session_state['historico'])}")

    # --- PANEL CENTRAL ---
    st.subheader("2. Operación Diaria")
    uploaded_file = st.file_uploader("Sube el Archivo Manual Diario (.xlsx)", type=["xlsx"])

    if uploaded_file and 'cartera' in st.session_state:
        if st.button("🚀 EJECUTAR ESCANEO PROFUNDO", type="primary", use_container_width=True):
            
            # 1. Leer Manual
            df_manual = procesar_archivo_manual(uploaded_file)
            if df_manual.empty:
                st.error("Error leyendo archivo manual.")
                return
            
            # 2. Leer KB de Google Sheets
            df_kb = pd.DataFrame()
            g_client = connect_to_google_sheets()
            if g_client:
                try:
                    sh = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"])
                    # Intentar leer KB, si no existe creamos vacío
                    try:
                        df_kb = get_as_dataframe(sh.worksheet("Knowledge_Base"), header=None)
                    except:
                        st.warning("Hoja 'Knowledge_Base' no encontrada, se creará al guardar.")
                    st.toast("Conectado a Google Sheets (KB)")
                except Exception as e:
                    st.warning(f"No se pudo cargar KB de Google: {e}")
            
            # 3. Correr Motor
            df_res = motor_omnisciente(
                df_manual, 
                st.session_state['cartera'], 
                st.session_state.get('historico', pd.DataFrame()), 
                df_kb
            )
            st.session_state['resultado_final'] = df_res

    # --- RESULTADOS ---
    if 'resultado_final' in st.session_state:
        df_res = st.session_state['resultado_final']
        st.divider()
        
        # KPIs
        kpis = {
            'total_tx': len(df_res),
            'exactos': len(df_res[df_res['Estado'].str.contains('EXACTO', na=False)]),
            'descuentos': len(df_res[df_res['Estado'].str.contains('DCTO', na=False)]),
            'impuestos': len(df_res[df_res['Estado'].str.contains('IMPUESTOS', na=False)]),
            'parciales': len(df_res[df_res['Estado'].str.contains('PARCIAL', na=False)]),
            'historico': len(df_res[df_res['Sugerencia_IA'].str.contains('Memoria', na=False)]),
            'sin_id': len(df_res[df_res['Estado'].str.contains('NO IDENTIFICADO', na=False)])
        }
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Conciliados Auto", kpis['exactos'] + kpis['descuentos'] + kpis['impuestos'])
        c2.metric("Parciales/Abonos", kpis['parciales'])
        c3.metric("Por Memoria", kpis['historico'])
        c4.metric("Pendientes", kpis['sin_id'])
        
        # Editor
        lista_clientes = sorted(st.session_state['cartera']['NombreCliente'].unique().tolist())
        col_config = {
            "Status_Gestion": st.column_config.SelectboxColumn("Gestión", options=['PENDIENTE', 'REGISTRADA'], required=True),
            "Cliente_Identificado": st.column_config.SelectboxColumn("Cliente", options=lista_clientes, width="large"),
            "Valor_Banco": st.column_config.NumberColumn("Valor", format="$ %d")
        }
        cols_view = ['Status_Gestion', 'FECHA', 'Valor_Banco', 'Cliente_Identificado', 'Estado', 'Sugerencia_IA', 'Detalle_Operacion', 'ID_Unico']
        
        edited_df = st.data_editor(
            df_res[cols_view], 
            use_container_width=True, 
            column_config=col_config, 
            key="editor_final"
        )
        
        # ACCIONES
        c_down, c_save = st.columns(2)
        
        with c_down:
            st.session_state['resultado_final'].update(edited_df)
            excel_data = generar_excel_profesional(st.session_state['resultado_final'], kpis)
            st.download_button("💾 Descargar Excel Profesional", data=excel_data, file_name="Conciliacion_Pro.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            
        with c_save:
            if st.button("☁️ GUARDAR Y ENTRENAR IA (Google Sheets)", type="primary", use_container_width=True):
                g_client = connect_to_google_sheets()
                if not g_client:
                    st.error("No se pudo conectar a Google Sheets.")
                else:
                    try:
                        sh = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"])
                        
                        # 1. Guardar Maestro (Estado Actual)
                        ws_master = sh.worksheet(st.secrets["google_sheets"]["tab_bancos_master"])
                        df_final = st.session_state['resultado_final'].copy()
                        df_final = df_final.fillna('')
                        df_final['FECHA'] = df_final['FECHA'].astype(str)
                        ws_master.clear()
                        set_with_dataframe(ws_master, df_final)
                        
                        # 2. Entrenar IA (Crear/Actualizar Knowledge_Base)
                        try:
                            ws_kb = sh.worksheet("Knowledge_Base")
                        except:
                            ws_kb = sh.add_worksheet(title="Knowledge_Base", rows=1000, cols=2)

                        # Filtramos lo que ya fue gestionado manualmente para "enseñarle" al robot
                        nuevos_registros = df_final[
                            (df_final['Status_Gestion'] == 'REGISTRADA') & 
                            (df_final['Cliente_Identificado'] != '')
                        ]
                        
                        if not nuevos_registros.empty:
                            data_kb = []
                            for _, r in nuevos_registros.iterrows():
                                # Texto normalizado del extracto -> Nombre Cliente
                                txt_raw = str(r['Texto_Completo'])
                                txt_norm = normalizar_texto_avanzado(txt_raw)
                                cli = str(r['Cliente_Identificado']).strip()
                                
                                if len(txt_norm) > 5 and cli:
                                    data_kb.append([txt_norm, cli])
                            
                            if data_kb:
                                # Usamos append para que la BD crezca con el tiempo
                                ws_kb.append_rows(data_kb)
                                st.toast(f"🧠 IA aprendió {len(data_kb)} nuevos patrones.")
                        
                        st.success("✅ Guardado en la Nube y Entrenamiento Completado")
                    except Exception as e:
                        st.error(f"Error guardando en Google Sheets: {e}")

if __name__ == "__main__":
    main()
