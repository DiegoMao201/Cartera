# ======================================================================================
# ARCHIVO: pages/2_Motor_Conciliacion.py
# (Versión v11 - "El Híbrido": Lógica v9.1 + Superpoderes de Edición y Memoria)
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

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Motor Conciliación v11", page_icon="🕵️‍♂️", layout="wide")

# ======================================================================================
# --- 1. CONEXIONES Y UTILIDADES (Igual a v9.1 + Hashlib) ---
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

def generar_id_unico(row):
    """Crea una huella digital única para cada movimiento bancario"""
    # Usamos Fecha + Valor + Texto para identificar la transacción siempre
    raw_str = f"{row['FECHA']}{row['Valor_Banco']}{row['Texto_Completo']}".strip()
    return hashlib.md5(raw_str.encode('utf-8')).hexdigest()

def normalizar_texto_avanzado(texto):
    if not isinstance(texto, str): return ""
    texto = texto.upper().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    texto = re.sub(r'[^A-Z0-9\s]', ' ', texto) 
    palabras_basura = [
        'PAGO', 'TRANSF', 'TRANSFERENCIA', 'CONSIGNACION', 'ABONO', 'CTA', 'NIT', 
        'REF', 'FACTURA', 'OFI', 'SUC', 'ACH', 'PSE', 'NOMINA', 'PROVEEDOR', 
        'COMPRA', 'VENTA', 'VALOR', 'NETO', 'PLANILLA', 'S A', 'SAS', 'LTDA', 
        'COLOMBIA', 'BANCOLOMBIA', 'DAVIVIENDA'
    ]
    for p in palabras_basura:
        texto = re.sub(r'\b' + p + r'\b', '', texto)
    return ' '.join(texto.split())

def extraer_posibles_nits(texto):
    if not isinstance(texto, str): return []
    return re.findall(r'\b\d{7,11}\b', texto)

def limpiar_moneda_colombiana(val):
    if isinstance(val, (int, float)):
        return float(val) if pd.notnull(val) else 0.0
    s = str(val).strip()
    if not s or s.lower() == 'nan': return 0.0
    s = s.replace('$', '').replace('USD', '').replace('COP', '').strip()
    s = s.replace('.', '')
    s = s.replace(',', '.')
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
# --- 2. CARGA DE DATOS (Igual a v9.1) ---
# ======================================================================================

@st.cache_data(ttl=600)
def cargar_cartera_detalle():
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
        return df[df['Importe'] > 100].copy()
    except Exception as e:
        st.error(f"Error estructura cartera: {e}")
        return pd.DataFrame()

def procesar_planilla_bancos(uploaded_file):
    try:
        df_temp = pd.read_excel(uploaded_file, nrows=15, header=None)
        header_idx = 0
        for idx, row in df_temp.iterrows():
            if 'FECHA' in row.astype(str).str.upper().values and 'VALOR' in row.astype(str).str.upper().values:
                header_idx = idx
                break
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, header=header_idx)
        df.columns = [str(c).strip().upper() for c in df.columns]
        df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')
        df = df.dropna(subset=['FECHA'])
        
        if 'VALOR' in df.columns:
            df['Valor_Banco'] = df['VALOR'].apply(limpiar_moneda_colombiana)
        else:
            df['Valor_Banco'] = 0.0

        cols_txt = [c for c in df.columns if c not in ['FECHA', 'VALOR', 'Valor_Banco']]
        df['Texto_Completo'] = df[cols_txt].fillna('').astype(str).agg(' '.join, axis=1)
        df['Texto_Norm'] = df['Texto_Completo'].apply(normalizar_texto_avanzado)
        
        # Rescate de dinero
        mask_zero = df['Valor_Banco'] == 0
        df.loc[mask_zero, 'Valor_Banco'] = df.loc[mask_zero, 'Texto_Completo'].apply(extraer_dinero_de_texto)
        
        # --- NUEVO: Generar ID para Historial ---
        df['ID_Transaccion'] = df.apply(generar_id_unico, axis=1)
        
        return df
    except Exception as e:
        st.error(f"Error leyendo Excel: {e}")
        return pd.DataFrame()

# ======================================================================================
# --- 3. ALGORITMO (NUCLEO v9.1 + EXTENSIONES INTELIGENTES) ---
# ======================================================================================

def analizar_cliente(nombre_banco, valor_pago, facturas_cliente):
    """
    TU LÓGICA FAVORITA DE LA V9.1 (INTACTA)
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
        res['Diferencia'] = valor_pago * -1
        return res

    facturas = facturas_cliente[['Numero', 'Importe', 'FechaDoc']].sort_values('FechaDoc').to_dict('records')
    total_deuda = sum(f['Importe'] for f in facturas)
    
    # 1. MATCH EXACTO TOTAL
    if abs(valor_pago - total_deuda) < 1000:
        res['Estado'] = '✅ MATCH EXACTO (TOTAL)'
        res['Facturas_Conciliadas'] = 'TODAS'
        res['Detalle_Operacion'] = f"Pago cubre las {len(facturas)} facturas pendientes."
        return res
        
    # 2. MATCH COMBINATORIO
    found_combo = False
    for r in range(1, 4): 
        if r > len(facturas): break
        for combo in itertools.combinations(facturas, r):
            suma_combo = sum(c['Importe'] for c in combo)
            numeros_combo = ", ".join([str(c['Numero']) for c in combo])
            
            if abs(valor_pago - suma_combo) < 500:
                res['Estado'] = '✅ MATCH FACTURAS ESPECÍFICAS'
                res['Facturas_Conciliadas'] = numeros_combo
                res['Detalle_Operacion'] = f"Suma exacta de {r} factura(s)."
                found_combo = True
                break
                
            suma_dcto = suma_combo * 0.97
            if abs(valor_pago - suma_dcto) < 2000:
                res['Estado'] = '💎 CONCILIADO CON DESCUENTO'
                res['Facturas_Conciliadas'] = numeros_combo
                res['Tipo_Ajuste'] = 'Descuento Pronto Pago'
                res['Detalle_Operacion'] = f"Pago con 3% Dcto sobre facturas: {numeros_combo}"
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
        else: break
            
    saldo_restante = total_deuda - valor_pago
    
    if facturas_cubiertas:
        res['Estado'] = '⚠️ ABONO PARCIAL (FIFO)'
        res['Facturas_Conciliadas'] = ", ".join(facturas_cubiertas)
        res['Diferencia'] = saldo_restante
        res['Detalle_Operacion'] = f"Cubre antiguas. Queda debiendo ${saldo_restante:,.0f}"
    else:
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
            res['Detalle_Operacion'] = f"Monto no cuadra. Deuda Total: ${total_deuda:,.0f}"

    return res

def buscar_match_global(valor_pago, df_cartera_completa):
    """
    NUEVA FUNCIÓN: Si no encontramos al cliente, buscamos el DINERO en toda la base.
    Retorna una sugerencia de texto.
    """
    # 1. Búsqueda exacta unitaria
    match_val = df_cartera_completa[
        (df_cartera_completa['Importe'] >= valor_pago - 50) & 
        (df_cartera_completa['Importe'] <= valor_pago + 50)
    ]
    
    if not match_val.empty:
        mejor_candidato = match_val.iloc[0]
        return f"💡 SUGERENCIA IA: Valor exacto en cliente '{mejor_candidato['NombreCliente']}' (Fac: {mejor_candidato['Numero']})"
    
    # 2. Si es un valor grande, intentamos buscar combinaciones rápidas en clientes grandes
    # (Para no hacer lento el proceso, solo miramos clientes con deuda total > pago)
    return ""

def correr_motor_con_memoria(df_bancos, df_cartera, df_kb, df_historial):
    st.info("🔎 Iniciando auditoría con Memoria y Radar Global...")
    
    # Mapas de búsqueda
    mapa_nit = df_cartera.groupby('nit_norm')['NombreCliente'].first().to_dict()
    lista_nombres = df_cartera['nombre_norm'].unique().tolist()
    
    # Mapa de historial (Para persistencia)
    mapa_historia = df_historial.set_index('ID_Transaccion').to_dict('index') if not df_historial.empty else {}

    # Memoria KB
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
        
        # --- PASO 0: VERIFICAR SI YA EXISTE EN HISTORIAL (PERSISTENCIA) ---
        if item['ID_Transaccion'] in mapa_historia:
            hist_data = mapa_historia[item['ID_Transaccion']]
            # Si ya fue gestionada, respetamos la data guardada
            item.update(hist_data) # Sobrescribe con lo que viene de Google Sheets
            item['Estado_Analisis'] = '🔒 HISTORIAL (YA REGISTRADO)'
            resultados.append(item)
            continue
        
        # --- SI ES NUEVO, CORREMOS TU LÓGICA v9.1 ---
        # Inicializar valores por defecto
        item['Status_Gestion'] = 'PENDIENTE' 
        item['Sugerencia_IA'] = ''
        
        txt = row['Texto_Norm']
        val = row['Valor_Banco']
        
        nit_found = None
        nombre_cliente = ""
        
        # 1. Identificación
        for k, v in memoria.items():
            if k in txt:
                nit_found = v
                break
        
        if not nit_found:
            posibles = extraer_posibles_nits(row['Texto_Completo'])
            for p in posibles:
                if p in mapa_nit:
                    nit_found = p
                    break
                    
        if not nit_found and len(txt) > 5:
            match, score = process.extractOne(txt, lista_nombres, scorer=fuzz.token_set_ratio)
            if score >= 88:
                nit_found = df_cartera[df_cartera['nombre_norm'] == match]['nit_norm'].iloc[0]

        # 2. Análisis Financiero (Llamamos a tu función original)
        if nit_found:
            nombre_cliente = mapa_nit.get(nit_found, "Cliente")
            facturas_open = df_cartera[df_cartera['nit_norm'] == nit_found]
            analisis = analizar_cliente(nombre_cliente, val, facturas_open)
            
            item.update(analisis) # Mezclamos resultado v9.1
            item['Cliente_Identificado'] = nombre_cliente
            item['NIT'] = nit_found
            item['Estado_Analisis'] = analisis['Estado'] # Guardamos el estado original
            
        else:
            # --- NUEVO: SI NO ENCUENTRA CLIENTE, ACTIVA EL RADAR GLOBAL ---
            item['Estado'] = '❓ NO IDENTIFICADO'
            item['Estado_Analisis'] = '❓ NO IDENTIFICADO'
            item['Cliente_Identificado'] = ''
            item['Detalle_Operacion'] = 'Falta información para cruzar.'
            
            # Aquí la IA busca en TODA la cartera a ver si ese dinero existe en otro lado
            sugerencia = buscar_match_global(val, df_cartera)
            if sugerencia:
                item['Sugerencia_IA'] = sugerencia
                item['Estado_Analisis'] = '💡 SUGERENCIA IA'

        resultados.append(item)
        
    return pd.DataFrame(resultados)

# ======================================================================================
# --- 4. INTERFAZ GRÁFICA (FILTROS Y EDICIÓN) ---
# ======================================================================================

def main():
    st.title("🏦 Conciliador Financiero v11")
    st.markdown("**El Híbrido:** Lógica probada v9 + Edición, Filtros y Memoria.")
    
    # --- BARRA LATERAL ---
    with st.sidebar:
        st.header("1. Carga de Datos")
        uploaded_file = st.file_uploader("📂 Planilla Banco (.xlsx)", type=["xlsx"])
        
        if st.button("🔄 Sincronizar Cartera"):
            with st.spinner("Descargando..."):
                df_c = cargar_cartera_detalle()
                if not df_c.empty:
                    st.session_state['cartera'] = df_c
                    st.success(f"Cartera: {len(df_c)} facturas.")
                else: st.error("Error Dropbox")
        
        st.divider()
        st.header("2. Filtros de Vista")
        # Los filtros se renderizan vacíos hasta que haya datos
        filtro_mes = st.empty()
        filtro_estado = st.empty()
        filtro_gestion = st.empty()

    # --- PANEL PRINCIPAL ---
    if uploaded_file and 'cartera' in st.session_state:
        
        if st.button("🚀 EJECUTAR CONCILIACIÓN", type="primary", use_container_width=True):
            # 1. Leer Banco
            df_bancos = procesar_planilla_bancos(uploaded_file)
            
            # 2. Leer Google Sheets (Base Conocimiento + Historial)
            g_client = connect_to_google_sheets()
            df_kb = pd.DataFrame()
            df_hist = pd.DataFrame()
            
            if g_client:
                try:
                    sh = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"])
                    # Cargar KB
                    try: df_kb = get_as_dataframe(sh.worksheet("Knowledge_Base"))
                    except: pass
                    # Cargar Historial (Bancos Master)
                    try: 
                        ws_hist = sh.worksheet(st.secrets["google_sheets"]["tab_bancos_master"])
                        df_hist = get_as_dataframe(ws_hist)
                        # Limpiar filas vacías si existen
                        df_hist = df_hist.dropna(how='all')
                    except: pass
                except: pass

            # 3. Correr Motor
            df_res = correr_motor_con_memoria(df_bancos, st.session_state['cartera'], df_kb, df_hist)
            st.session_state['resultados_v11'] = df_res
            st.rerun() # Recargar para aplicar filtros

        # --- PANTALLA DE RESULTADOS Y EDICIÓN ---
        if 'resultados_v11' in st.session_state:
            df = st.session_state['resultados_v11'].copy()
            
            # Asegurar columna Mes
            df['Mes'] = df['FECHA'].dt.strftime('%Y-%m')
            
            # --- RENDERIZAR FILTROS EN SIDEBAR ---
            with filtro_mes:
                meses = sorted(df['Mes'].unique())
                sel_mes = st.multiselect("📅 Mes", meses, default=meses)
            
            with filtro_estado:
                estados = sorted(df['Estado_Analisis'].unique())
                sel_estado = st.multiselect("📊 Estado Conciliación", estados, default=estados)
                
            with filtro_gestion:
                # Asegurar que Status_Gestion existe y no tiene nulos feos
                df['Status_Gestion'] = df['Status_Gestion'].fillna('PENDIENTE')
                gestiones = sorted(df['Status_Gestion'].unique())
                sel_gestion = st.multiselect("📝 Estado Gestión", gestiones, default=gestiones)
            
            # --- APLICAR FILTROS ---
            mask = (df['Mes'].isin(sel_mes)) & (df['Estado_Analisis'].isin(sel_estado)) & (df['Status_Gestion'].isin(sel_gestion))
            df_filtered = df[mask].copy()
            
            st.divider()
            st.subheader(f"📋 Panel de Gestión ({len(df_filtered)} registros)")
            
            # Obtener lista de clientes para dropdown (Opcional, si son muchos mejor texto)
            lista_clientes_cartera = sorted(st.session_state['cartera']['NombreCliente'].unique().tolist())
            
            # --- CONFIGURAR EDITOR ---
            column_config = {
                "Status_Gestion": st.column_config.SelectboxColumn(
                    "Estado (Editable)",
                    options=['PENDIENTE', 'REGISTRADA'],
                    required=True,
                    width="medium"
                ),
                "Cliente_Identificado": st.column_config.SelectboxColumn(
                    "Cliente (Seleccionar)",
                    options=lista_clientes_cartera,
                    width="large",
                    required=False
                ),
                "Sugerencia_IA": st.column_config.TextColumn(
                    "Sugerencias IA",
                    disabled=True,
                    width="medium"
                ),
                "Valor_Banco": st.column_config.NumberColumn(
                    "Valor",
                    format="$ %d"
                ),
                "Detalle_Operacion": st.column_config.TextColumn("Detalle Sistema", disabled=True),
                "FECHA": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY", disabled=True)
            }
            
            cols_show = [
                'Status_Gestion', 'FECHA', 'Valor_Banco', 'Cliente_Identificado', 
                'Estado_Analisis', 'Sugerencia_IA', 'Detalle_Operacion', 'Facturas_Conciliadas', 
                'Texto_Completo', 'ID_Transaccion'
            ]
            
            edited_df = st.data_editor(
                df_filtered[cols_show],
                column_config=column_config,
                use_container_width=True,
                height=500,
                num_rows="fixed",
                key="editor_datos"
            )
            
            # --- BOTÓN DE GUARDADO ---
            col1, col2 = st.columns([1,3])
            with col1:
                if st.button("💾 GUARDAR CAMBIOS EN LA NUBE", type="primary"):
                    try:
                        # 1. Actualizar Dataframe Maestro en Memoria
                        # Usamos ID_Transaccion como índice para actualizar solo lo cambiado
                        df_master = st.session_state['resultados_v11'].set_index('ID_Transaccion')
                        df_changes = edited_df.set_index('ID_Transaccion')
                        
                        # Actualizamos solo las columnas que permite editar el usuario
                        df_master.update(df_changes[['Status_Gestion', 'Cliente_Identificado']])
                        
                        # Recuperamos el DF completo actualizado
                        df_final = df_master.reset_index()
                        st.session_state['resultados_v11'] = df_final # Guardar en sesión
                        
                        # 2. Enviar a Google Sheets
                        g_client = connect_to_google_sheets()
                        if g_client:
                            sh = g_client.open_by_url(st.secrets["google_sheets"]["sheet_url"])
                            ws = sh.worksheet(st.secrets["google_sheets"]["tab_bancos_master"])
                            
                            # Preparamos datos para guardar (Convertir fechas a string)
                            df_save = df_final.copy()
                            df_save['FECHA'] = df_save['FECHA'].astype(str)
                            df_save = df_save.fillna('')
                            
                            # Sobrescribimos la hoja Master (Es lo más seguro para mantener integridad)
                            ws.clear()
                            set_with_dataframe(ws, df_save)
                            
                            # 3. Aprendizaje Automático (Opcional)
                            # Si marcaste REGISTRADA manual, agregamos a KB
                            nuevos_manuales = df_final[
                                (df_final['Status_Gestion'] == 'REGISTRADA') & 
                                (df_final['Estado_Analisis'] == '❓ NO IDENTIFICADO') &
                                (df_final['Cliente_Identificado'] != '')
                            ]
                            
                            if not nuevos_manuales.empty:
                                ws_kb = sh.worksheet("Knowledge_Base")
                                data_kb = []
                                for _, r in nuevos_manuales.iterrows():
                                    # Guardamos una palabra clave del texto
                                    key = r['Texto_Completo'][:25].strip()
                                    val = r['Cliente_Identificado']
                                    data_kb.append([key, val])
                                ws_kb.append_rows(data_kb)
                                st.toast(f"🧠 Aprendí {len(data_kb)} nuevos patrones.")

                            st.success("✅ ¡Datos sincronizados y guardados exitosamente!")
                        
                    except Exception as e:
                        st.error(f"Error al guardar: {e}")

if __name__ == "__main__":
    main()
