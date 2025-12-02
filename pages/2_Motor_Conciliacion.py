# ======================================================================================
# ARCHIVO: pages/2_Motor_Conciliacion.py
# (Versión v14 - "El Omnisciente": Deep Token Search + Keyword Mapping)
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
st.set_page_config(page_title="Motor Conciliación Pro v14", page_icon="👁️", layout="wide")

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

def generar_id_unico(row, index):
    """
    Crea una huella digital única (MD5).
    Incluye el 'index' para garantizar unicidad absoluta.
    """
    fecha_str = str(row['FECHA'])
    val_str = str(row['Valor_Banco'])
    txt_str = str(row['Texto_Completo']).strip()
    raw_str = f"{index}_{fecha_str}{val_str}{txt_str}"
    return hashlib.md5(raw_str.encode('utf-8')).hexdigest()

def normalizar_texto_avanzado(texto):
    """Limpieza profunda de texto para mejorar el Fuzzy Matching"""
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
        texto = re.sub(r'\b' + p + r'\b', ' ', texto) # Reemplazar por espacio para no pegar palabras
    return ' '.join(texto.split())

def extraer_posibles_nits(texto):
    """Extrae números que parezcan NITs, incluyendo formatos con puntos"""
    if not isinstance(texto, str): return []
    # Busca nits limpios (890900123) o con puntos (890.900.123)
    clean_txt = texto.replace('.', '').replace('-', '')
    return re.findall(r'\b\d{7,11}\b', clean_txt)

def limpiar_moneda_colombiana(val):
    """Convierte texto financiero colombiano a float"""
    if isinstance(val, (int, float)):
        return float(val) if pd.notnull(val) else 0.0
    
    s = str(val).strip()
    if not s or s.lower() == 'nan': return 0.0

    s = s.replace('$', '').replace('USD', '').replace('COP', '').strip()
    s = s.replace('.', '') # Quitar miles
    s = s.replace(',', '.') # Convertir decimal
    
    try:
        return float(s)
    except ValueError:
        return 0.0

def extraer_dinero_de_texto(texto):
    """Intenta rescatar montos numéricos de la descripción"""
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
        
        # --- HOJA 1: RESUMEN GERENCIAL ---
        workbook = writer.book
        sheet_resumen = workbook.add_worksheet("Dashboard")
        sheet_resumen.hide_gridlines(2)
        
        style_title = workbook.add_format({'bold': True, 'font_size': 16, 'font_color': '#1F497D'})
        style_kpi_label = workbook.add_format({'bold': True, 'bg_color': '#E7E6E6', 'border': 1})
        style_kpi_value = workbook.add_format({'bold': True, 'num_format': '#,##0', 'border': 1, 'align': 'center'})
        
        sheet_resumen.write('B2', "RESUMEN DE CONCILIACIÓN - OMNISCIENTE", style_title)
        
        kpis = [
            ("Total Movimientos", resumen_kpis['total_tx']),
            ("Conciliados (Match Exacto)", resumen_kpis['exactos']),
            ("Conciliados (Con Descuento)", resumen_kpis['descuentos']),
            ("Conciliados (Impuestos)", resumen_kpis['impuestos']),
            ("Parciales / Abonos", resumen_kpis['parciales']),
            ("Histórico (Ya procesado)", resumen_kpis['historico']),
            ("Sin Identificar", resumen_kpis['sin_id'])
        ]
        
        row = 4
        for label, val in kpis:
            sheet_resumen.write(row, 1, label, style_kpi_label)
            sheet_resumen.write(row, 2, val, style_kpi_value)
            row += 1

        # --- HOJA 2: DETALLE ---
        cols_export = [
            'FECHA', 'Valor_Banco', 'Texto_Completo', 'Cliente_Identificado', 
            'NIT', 'Estado', 'Facturas_Conciliadas', 'Detalle_Operacion', 
            'Diferencia', 'Tipo_Ajuste', 'Status_Gestion', 'Sugerencia_IA'
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
        fmt_purple = workbook.add_format({'bg_color': '#E1D5E7', 'font_color': '#604A7B'}) # Nuevo para IA

        for col_num, value in enumerate(df_export.columns.values):
            worksheet.write(0, col_num, value, header_format)
            col_name = str(value).upper()
            width = 20
            if "CLIENTE" in col_name or "TEXTO" in col_name: width = 45
            elif "DETALLE" in col_name: width = 50
            
            worksheet.set_column(col_num, col_num, width, text_fmt)
            if any(x in col_name for x in ['VALOR', 'DIFERENCIA']):
                worksheet.set_column(col_num, col_num, width, currency_fmt)

        # Formato Condicional
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
# --- 3. CARGA DE DATOS ---
# ======================================================================================

@st.cache_data(ttl=600)
def cargar_cartera_detalle():
    """Carga Cartera desde Dropbox"""
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
        df['FechaDoc'] = pd.to_datetime(df['FechaDoc'], errors='coerce')
        
        # Generar lista de palabras clave únicas por cliente para la búsqueda profunda
        return df[df['Importe'] > 100].copy()
    except Exception as e:
        st.error(f"Error estructura cartera: {e}")
        return pd.DataFrame()

def procesar_planilla_bancos(uploaded_file):
    """Procesamiento inteligente del Excel Bancario"""
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
        
        df['ID_Transaccion'] = [generar_id_unico(row, idx) for idx, row in df.iterrows()]
        
        return df
    except Exception as e:
        st.error(f"Error leyendo Excel: {e}")
        return pd.DataFrame()

# ======================================================================================
# --- 4. ALGORITMO OMNISCIENTE ---
# ======================================================================================

def construir_indice_palabras_unicas(df_cartera):
    """
    Crea un mapa de palabras clave únicas.
    Si la palabra "ASUL" aparece solo en un NIT, mapeamos ASUL -> NIT.
    Esto permite encontrar al cliente incluso si el texto está sucio.
    """
    word_to_nits = defaultdict(set)
    for idx, row in df_cartera.iterrows():
        nit = row['nit_norm']
        nombre = str(row['nombre_norm'])
        # Tokenizar
        words = nombre.split()
        for w in words:
            if len(w) > 3: # Ignorar palabras muy cortas
                word_to_nits[w].add(nit)
    
    # Filtrar solo las palabras que apuntan a un único cliente (Determinísticas)
    unique_keyword_map = {}
    for w, nits in word_to_nits.items():
        if len(nits) == 1:
            unique_keyword_map[w] = list(nits)[0]
            
    return unique_keyword_map

def analizar_cliente(nombre_banco, valor_pago, facturas_cliente):
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
        res['Diferencia'] = valor_pago * -1
        return res

    facturas = facturas_cliente[['Numero', 'Importe', 'FechaDoc']].sort_values('FechaDoc').to_dict('records')
    total_deuda = sum(f['Importe'] for f in facturas)
    
    # 1. MATCH EXACTO
    if abs(valor_pago - total_deuda) < 1000:
        res['Estado'] = '✅ MATCH EXACTO (TOTAL)'
        res['Facturas_Conciliadas'] = 'TODAS'
        res['Detalle_Operacion'] = f"Pago total deuda."
        return res
        
    # 2. MATCH COMBINATORIO
    found_combo = False
    for r in range(1, 5): # Aumentado a 5 facturas
        if r > len(facturas): break
        for combo in itertools.combinations(facturas, r):
            suma_combo = sum(c['Importe'] for c in combo)
            numeros_combo = ", ".join([str(c['Numero']) for c in combo])
            
            # A. Exacto
            if abs(valor_pago - suma_combo) < 500:
                res['Estado'] = '✅ MATCH FACTURAS ESPECÍFICAS'
                res['Facturas_Conciliadas'] = numeros_combo
                res['Detalle_Operacion'] = f"Suma exacta de {r} factura(s)."
                found_combo = True
                break
                
            # B. Descuento (~3%)
            suma_dcto = suma_combo * 0.97
            if abs(valor_pago - suma_dcto) < 2000:
                res['Estado'] = '💎 CONCILIADO CON DESCUENTO'
                res['Facturas_Conciliadas'] = numeros_combo
                res['Tipo_Ajuste'] = 'Descuento Pronto Pago'
                res['Ahorro_Dcto'] = suma_combo - valor_pago
                res['Detalle_Operacion'] = f"Pago con 3% Dcto."
                found_combo = True
                break
        if found_combo: break
    
    if found_combo: return res

    # 3. FIFO / IMPUESTOS
    acumulado = 0
    facturas_cubiertas = []
    
    for f in facturas:
        if acumulado + f['Importe'] <= valor_pago + 1000:
            acumulado += f['Importe']
            facturas_cubiertas.append(str(f['Numero']))
        else:
            break
            
    saldo_restante = total_deuda - valor_pago
    
    if facturas_cubiertas:
        res['Estado'] = '⚠️ ABONO PARCIAL (FIFO)'
        res['Facturas_Conciliadas'] = ", ".join(facturas_cubiertas)
        res['Diferencia'] = saldo_restante
        res['Detalle_Operacion'] = f"Abona a deuda. Resta ${saldo_restante:,.0f}"
    else:
        # Chequeo impuestos
        base_est = total_deuda / 1.19
        rete_fuente = base_est * 0.025
        rete_iva = (base_est * 0.19) * 0.15
        pago_con_imptos = total_deuda - rete_fuente - rete_iva
        
        if abs(valor_pago - pago_con_imptos) < 5000:
            res['Estado'] = '🏢 CONCILIADO (IMPUESTOS)'
            res['Impuesto_Est'] = rete_fuente + rete_iva
            res['Detalle_Operacion'] = "Coincide menos retenciones."
        else:
            res['Estado'] = '❌ REVISAR MANUALMENTE'
            res['Diferencia'] = saldo_restante
            res['Detalle_Operacion'] = f"No cruza. Deuda Total: ${total_deuda:,.0f}"

    return res

def buscar_match_global_por_monto(valor_pago, df_cartera_completa):
    if valor_pago < 10000: return "" 
    match_val = df_cartera_completa[
        (df_cartera_completa['Importe'] >= valor_pago - 100) & 
        (df_cartera_completa['Importe'] <= valor_pago + 100)
    ]
    if not match_val.empty:
        mejor = match_val.iloc[0]
        return f"💡 Monto coincide con '{mejor['NombreCliente']}' (Fac: {mejor['Numero']})"
    return ""

def correr_motor_supremo(df_bancos, df_cartera, df_kb, df_historial):
    st.info("👁️ Iniciando Motor Omnisciente: Buscando 'ASUL' y patrones ocultos...")
    
    mapa_nit = df_cartera.groupby('nit_norm')['NombreCliente'].first().to_dict()
    lista_nombres = df_cartera['nombre_norm'].unique().tolist()
    
    # INDICE DE PALABRAS CLAVE (La Magia para "ASUL")
    indice_palabras = construir_indice_palabras_unicas(df_cartera)

    # 1. Preparar Historial
    mapa_historia = {}
    if not df_historial.empty:
        df_historial = df_historial.loc[:, ~df_historial.columns.str.contains('^Unnamed')]
        if 'ID_Transaccion' in df_historial.columns:
            df_historial['ID_Transaccion'] = df_historial['ID_Transaccion'].astype(str)
            df_historial = df_historial.drop_duplicates(subset=['ID_Transaccion'])
            mapa_historia = df_historial.set_index('ID_Transaccion').to_dict('index')

    # 2. Preparar Knowledge Base
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
        
        # --- A. CHECK HISTORIAL ---
        current_id = str(item.get('ID_Transaccion', ''))
        if current_id in mapa_historia:
            hist_data = mapa_historia[current_id]
            item['Status_Gestion'] = hist_data.get('Status_Gestion', 'REGISTRADA')
            item['Cliente_Identificado'] = hist_data.get('Cliente_Identificado', '')
            item['Estado'] = '🔒 HISTORIAL (YA REGISTRADO)'
            item['Detalle_Operacion'] = 'Transacción procesada anteriormente.'
            item['Facturas_Conciliadas'] = hist_data.get('Facturas_Conciliadas', '')
            resultados.append(item)
            continue
        
        # --- B. ANÁLISIS NUEVO ---
        item['Status_Gestion'] = 'PENDIENTE'
        item['Sugerencia_IA'] = ''
        
        txt_full = str(row['Texto_Completo']).upper()
        txt_norm = row['Texto_Norm']
        val = row['Valor_Banco']
        
        nit_found = None
        origen_match = ""
        
        # B1. Memoria (Knowledge Base)
        for k, v in memoria.items():
            if k in txt_norm:
                nit_found = v
                origen_match = "IA (Memoria)"
                break
        
        # B2. NIT Explicito (Regex)
        if not nit_found:
            posibles = extraer_posibles_nits(txt_full)
            for p in posibles:
                # Limpiar puntos para cruzar
                p_clean = p.replace('.', '').replace('-', '')
                if p_clean in mapa_nit:
                    nit_found = p_clean
                    origen_match = "NIT Encontrado"
                    break
        
        # B3. KEYWORD MATCHING (Solución "ASUL")
        # Rompe la frase en palabras y busca si alguna palabra es única de un cliente
        if not nit_found:
            palabras_banco = txt_norm.split()
            for palabra in palabras_banco:
                if len(palabra) > 3 and palabra in indice_palabras:
                    nit_found = indice_palabras[palabra]
                    origen_match = f"Palabra Clave '{palabra}'"
                    break

        # B4. Búsqueda de Nombre Contenido (String Containment)
        # Si el nombre normalizado del cliente (ej: "ASUL") está dentro de la frase
        if not nit_found:
            for nombre_cli in lista_nombres:
                # Solo si el nombre es lo suficientemente largo para evitar falsos positivos
                if len(nombre_cli) > 4 and nombre_cli in txt_norm:
                     nit_found = df_cartera[df_cartera['nombre_norm'] == nombre_cli]['nit_norm'].iloc[0]
                     origen_match = "Nombre Contenido"
                     break

        # B5. Fuzzy Logic (Último recurso)
        if not nit_found and len(txt_norm) > 5:
            # Token Set Ratio es potente para desorden
            match, score = process.extractOne(txt_norm, lista_nombres, scorer=fuzz.token_set_ratio)
            if score >= 85: # Bajamos un poco umbral para ser más agresivos
                nit_found = df_cartera[df_cartera['nombre_norm'] == match]['nit_norm'].iloc[0]
                origen_match = f"Fuzzy Logic ({score}%)"

        # --- C. MOTOR FINANCIERO ---
        if nit_found:
            nombre_cliente = mapa_nit.get(nit_found, "Cliente")
            facturas_open = df_cartera[df_cartera['nit_norm'] == nit_found]
            analisis = analizar_cliente(nombre_cliente, val, facturas_open)
            
            item.update(analisis)
            item['Cliente_Identificado'] = nombre_cliente
            item['NIT'] = nit_found
            item['Sugerencia_IA'] = f"Detectado por: {origen_match}"
            
        else:
            item['Estado'] = '❓ NO IDENTIFICADO'
            item['Cliente_Identificado'] = ''
            item['Detalle_Operacion'] = 'Falta información para cruzar.'
            
            # B6. Radar Global (Monto)
            sugerencia = buscar_match_global_por_monto(val, df_cartera)
            if sugerencia:
                item['Sugerencia_IA'] = sugerencia
                item['Estado'] = '💡 SUGERENCIA IA'

        resultados.append(item)
        
    return pd.DataFrame(resultados)

# ======================================================================================
# --- 5. INTERFAZ GRÁFICA MAESTRA ---
# ======================================================================================

def main():
    st.title("👁️ Motor Conciliación Omnisciente v14")
    st.markdown("**Novedad:** Detecta palabras clave ocultas (ej. 'ASUL') dentro de descripciones complejas.")
    
    # --- BARRA LATERAL ---
    with st.sidebar:
        st.header("1. Carga de Datos")
        uploaded_file = st.file_uploader("📂 Planilla Banco (.xlsx)", type=["xlsx"])
        
        if st.button("🔄 Sincronizar Cartera (Dropbox)"):
            with st.spinner("Descargando cartera actualizada..."):
                df_c = cargar_cartera_detalle()
                if not df_c.empty:
                    st.session_state['cartera'] = df_c
                    st.success(f"Cartera OK: {len(df_c)} facturas.")
                else: st.error("Error conectando a Dropbox")
        
        st.divider()
        st.header("2. Filtros de Vista")
        filtro_fecha = st.empty()
        filtro_mes = st.empty()
        filtro_estado = st.empty()
        filtro_gestion = st.empty()

    # --- PANEL PRINCIPAL ---
    if uploaded_file and 'cartera' in st.session_state:
        
        if st.button("🚀 EJECUTAR ESCANEO PROFUNDO", type="primary", use_container_width=True):
            # 1. Leer Banco
            df_bancos = procesar_planilla_bancos(uploaded_file)
            if df_bancos.empty:
                st.error("Archivo de banco vacío o ilegible.")
                return

            # 2. Leer Nube
            g_client = connect_to_google_sheets()
            df_kb = pd.DataFrame()
            df_hist = pd.DataFrame()
            
            if g_client:
                try:
                    sh = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"])
                    try: df_kb = get_as_dataframe(sh.worksheet("Knowledge_Base"))
                    except: pass
                    try: 
                        ws_hist = sh.worksheet(st.secrets["google_sheets"]["tab_bancos_master"])
                        df_hist = get_as_dataframe(ws_hist)
                        df_hist = df_hist.dropna(how='all')
                    except: pass
                except: pass

            # 3. Correr Motor
            df_res = correr_motor_supremo(df_bancos, st.session_state['cartera'], df_kb, df_hist)
            st.session_state['resultados_supremo'] = df_res
            st.rerun()

        # --- PANTALLA DE RESULTADOS ---
        if 'resultados_supremo' in st.session_state:
            df = st.session_state['resultados_supremo'].copy()
            
            # KPIs
            kpis = {
                'total_tx': len(df),
                'exactos': len(df[df['Estado'].str.contains('EXACTO', na=False)]),
                'descuentos': len(df[df['Estado'].str.contains('DESCUENTO', na=False)]),
                'impuestos': len(df[df['Estado'].str.contains('IMPUESTOS', na=False)]),
                'parciales': len(df[df['Estado'].str.contains('PARCIAL|ABONO', regex=True, na=False)]),
                'historico': len(df[df['Estado'].str.contains('HISTORIAL', na=False)]),
                'sin_id': len(df[df['Estado'].str.contains('NO IDENTIFICADO|SUGERENCIA', regex=True, na=False)])
            }
            
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Conciliados Auto", kpis['exactos'] + kpis['descuentos'] + kpis['impuestos'])
            c2.metric("Parciales/Abonos", kpis['parciales'])
            c3.metric("Ya en Historial", kpis['historico'])
            c4.metric("Por Gestionar", kpis['sin_id'])
            
            st.divider()
            
            # --- LÓGICA DE FILTROS ---
            df['Mes'] = df['FECHA'].dt.strftime('%Y-%m')
            
            with filtro_fecha:
                min_d = df['FECHA'].min().date()
                max_d = df['FECHA'].max().date()
                rango = st.date_input("📅 Rango", value=(min_d, max_d), min_value=min_d, max_value=max_d)
            
            with filtro_mes:
                meses = sorted(df['Mes'].unique())
                sel_mes = st.multiselect("🗓️ Mes", meses, default=meses)
            
            with filtro_estado:
                estados = sorted(df['Estado'].unique())
                sel_estado = st.multiselect("📊 Estado", estados, default=estados)
                
            with filtro_gestion:
                gestiones = sorted(df['Status_Gestion'].unique())
                sel_gestion = st.multiselect("📝 Gestión", gestiones, default=gestiones)
                
            # Aplicar Filtros
            mask = (df['Mes'].isin(sel_mes)) & (df['Estado'].isin(sel_estado)) & (df['Status_Gestion'].isin(sel_gestion))
            if isinstance(rango, tuple) and len(rango) == 2:
                mask = mask & (df['FECHA'].dt.date >= rango[0]) & (df['FECHA'].dt.date <= rango[1])
            
            df_filtered = df[mask].copy()

            # --- EDITOR DE DATOS ---
            st.subheader(f"📋 Panel de Gestión ({len(df_filtered)} registros)")
            
            lista_clientes = sorted(st.session_state['cartera']['NombreCliente'].unique().tolist())
            
            col_config = {
                "Status_Gestion": st.column_config.SelectboxColumn("Estado Gestión", options=['PENDIENTE', 'REGISTRADA'], required=True),
                "Cliente_Identificado": st.column_config.SelectboxColumn("Cliente (Editar)", options=lista_clientes, width="large"),
                "Sugerencia_IA": st.column_config.TextColumn("IA Detectó", disabled=True, width="medium"),
                "Valor_Banco": st.column_config.NumberColumn("Valor", format="$ %d"),
                "FECHA": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY", disabled=True)
            }
            
            cols_show = [
                'Status_Gestion', 'FECHA', 'Valor_Banco', 'Cliente_Identificado', 
                'Estado', 'Sugerencia_IA', 'Detalle_Operacion', 'Facturas_Conciliadas', 
                'Texto_Completo', 'ID_Transaccion'
            ]
            
            edited_df = st.data_editor(
                df_filtered[cols_show],
                column_config=col_config,
                use_container_width=True,
                height=500,
                num_rows="fixed",
                key="editor_supremo"
            )
            
            # --- ACCIONES FINALES ---
            col_btn1, col_btn2 = st.columns([1, 1])
            
            # A. Descargar Excel
            with col_btn1:
                st.session_state['resultados_supremo'].update(edited_df)
                df_final_export = st.session_state['resultados_supremo'].copy()
                excel_data = generar_excel_profesional(df_final_export, kpis)
                
                st.download_button(
                    label="💾 Descargar Reporte Profesional (.xlsx)",
                    data=excel_data,
                    file_name=f"Conciliacion_Omnisciente_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="secondary",
                    use_container_width=True
                )

            # B. Guardar en Nube
            with col_btn2:
                if st.button("☁️ GUARDAR CAMBIOS Y ENTRENAR IA", type="primary", use_container_width=True):
                    try:
                        st.session_state['resultados_supremo'].update(edited_df)
                        df_final = st.session_state['resultados_supremo'].copy()
                        
                        g_client = connect_to_google_sheets()
                        if g_client:
                            sh = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"])
                            ws = sh.worksheet(st.secrets["google_sheets"]["tab_bancos_master"])
                            
                            df_save = df_final.copy()
                            df_save['FECHA'] = df_save['FECHA'].astype(str)
                            df_save = df_save.fillna('')
                            
                            ws.clear()
                            set_with_dataframe(ws, df_save)
                            
                            # Entrenar IA
                            nuevos_manuales = df_final[
                                (df_final['Status_Gestion'] == 'REGISTRADA') & 
                                (df_final['Estado'].str.contains('NO IDENTIFICADO', na=False)) &
                                (df_final['Cliente_Identificado'] != '')
                            ]
                            
                            if not nuevos_manuales.empty:
                                ws_kb = sh.worksheet("Knowledge_Base")
                                data_kb = []
                                for _, r in nuevos_manuales.iterrows():
                                    # Entrenamos con la frase normalizada para mejorar futuros matches
                                    key = str(r['Texto_Norm'])[:40].strip().upper()
                                    val = str(r['Cliente_Identificado']).strip()
                                    if len(key) > 5 and len(val) > 3:
                                        data_kb.append([key, val])
                                
                                if data_kb:
                                    ws_kb.append_rows(data_kb)
                                    st.toast(f"🧠 La IA aprendió {len(data_kb)} nuevos patrones.")
                            
                            st.success("✅ ¡Datos sincronizados exitosamente!")
                            
                    except Exception as e:
                        st.error(f"Error al guardar: {e}")

if __name__ == "__main__":
    main()
