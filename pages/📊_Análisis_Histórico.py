# ======================================================================================
# ARCHIVO: pages/📊_Análisis_Histórico.py
# ======================================================================================
import streamlit as st
import pandas as pd
import glob
import re
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Análisis Histórico", page_icon="📊", layout="wide")

# --- GUARDIA DE SEGURIDAD ---
if 'authentication_status' not in st.session_state or not st.session_state['authentication_status']:
    st.warning("Por favor, inicie sesión en el 📈 Tablero Principal para acceder a esta página.")
    st.stop()

# --- CÓDIGO DE LA PÁGINA ---
st.title("📊 Análisis Histórico de Cartera")

@st.cache_data
def cargar_datos_historicos():
    mapa_columnas = {
        'Serie': 'serie', 'Número': 'numero', 'Fecha Documento': 'fecha_documento',
        'Fecha Vencimiento': 'fecha_vencimiento', 'Fecha Saldado': 'fecha_saldado',
        'NOMBRECLIENTE': 'nombrecliente', 'Población': 'poblacion', 'Provincia': 'provincia',
        'IMPORTE': 'importe', 'RIESGOCONCEDIDO': 'riesgoconcedido', 'NOMVENDEDOR': 'nomvendedor',
        'DIAS_VENCIDO': 'dias_vencido', 'Estado': 'estado'
    }
    lista_archivos = sorted(glob.glob("Cartera_*.xlsx"))
    if not lista_archivos: return pd.DataFrame()
    lista_df = []
    for archivo in lista_archivos:
        try:
            df = pd.read_excel(archivo)
            if not df.empty: df = df.iloc[:-1]
            df['Serie'] = df['Serie'].astype(str)
            df = df[~df['Serie'].str.contains('W|X', case=False, na=False)]
            df.rename(columns=mapa_columnas, inplace=True)
            lista_df.append(df)
        except Exception as e:
            st.warning(f"No se pudo procesar el archivo {archivo}: {e}")
    if not lista_df: return pd.DataFrame()
    df_completo = pd.concat(lista_df, ignore_index=True)
    df_completo.dropna(subset=['numero', 'nombrecliente'], inplace=True)
    df_completo.sort_values(by=['fecha_documento', 'fecha_saldado'], ascending=[True, True], na_position='first', inplace=True)
    df_historico_unico = df_completo.drop_duplicates(subset=['numero'], keep='last')
    for col in ['fecha_documento', 'fecha_vencimiento', 'fecha_saldado']:
        df_historico_unico[col] = pd.to_datetime(df_historico_unico[col], errors='coerce')
    df_historico_unico['importe'] = pd.to_numeric(df_historico_unico['importe'], errors='coerce').fillna(0)
    df_pagadas = df_historico_unico.dropna(subset=['fecha_saldado', 'fecha_documento']).copy()
    if not df_pagadas.empty:
        df_pagadas['dias_de_pago'] = (df_pagadas['fecha_saldado'] - df_pagadas['fecha_documento']).dt.days
        df_historico_unico = pd.merge(df_historico_unico, df_pagadas[['numero', 'dias_de_pago']], on='numero', how='left')
    return df_historico_unico

df_historico_base = cargar_datos_historicos()

if df_historico_base.empty:
    st.warning("No se encontraron archivos de datos históricos."); st.stop()

# --- Filtros de la Página ---
st.sidebar.header("Filtros de Análisis")
acceso_general = st.session_state.get('acceso_general', False)
vendedor_autenticado = st.session_state.get('vendedor_autenticado', None)

# Filtro por Vendedor
if acceso_general:
    vendedores = ["Todos"] + sorted(df_historico_base['nomvendedor'].dropna().unique())
    vendedor_sel_hist = st.sidebar.selectbox("Vendedor:", vendedores)
else:
    vendedor_sel_hist = vendedor_autenticado

if vendedor_sel_hist == "Todos":
    df_historico = df_historico_base.copy()
else:
    df_historico = df_historico_base[df_historico_base['nomvendedor'] == vendedor_sel_hist].copy()

min_date = df_historico['fecha_documento'].min().date()
max_date_saldado = df_historico['fecha_saldado'].max()
max_date_doc = df_historico['fecha_documento'].max()
max_date = max(max_date_saldado, max_date_doc).date() if pd.notna(max_date_saldado) else max_date_doc.date()
default_start_date = max(min_date, max_date - relativedelta(months=12))
fecha_inicio, fecha_fin = st.sidebar.date_input("Rango de Fechas:", (default_start_date, max_date), min_value=min_date, max_value=max_date)

if not fecha_inicio or not fecha_fin or fecha_inicio > fecha_fin:
    st.error("Por favor, selecciona un rango de fechas válido."); st.stop()

fecha_inicio, fecha_fin = pd.to_datetime(fecha_inicio), pd.to_datetime(fecha_fin)
df_periodo = df_historico[(df_historico['fecha_documento'].between(fecha_inicio, fecha_fin)) | (df_historico['fecha_saldado'].between(fecha_inicio, fecha_fin))]

# --- El resto de la página (KPIs y Gráficos) ---
# ... (El código de esta parte no cambia, se pega tal cual estaba)
if df_periodo.empty:
    st.warning("No hay datos en el período seleccionado."); st.stop()

ventas_periodo = df_periodo[df_periodo['fecha_documento'].between(fecha_inicio, fecha_fin)]
total_ventas = ventas_periodo['importe'].sum()
cobros_periodo = df_periodo[df_periodo['fecha_saldado'].between(fecha_inicio, fecha_fin)]
total_cobrado = cobros_periodo['importe'].sum()
dso_periodo = cobros_periodo['dias_de_pago'].mean() if not cobros_periodo.empty else 0
snapshot_final = df_historico[df_historico['fecha_documento'] <= fecha_fin]
facturas_abiertas_al_final = snapshot_final[(snapshot_final['fecha_saldado'].isnull()) | (snapshot_final['fecha_saldado'] > fecha_fin)]
facturas_vencidas_al_final = facturas_abiertas_al_final[facturas_abiertas_al_final['fecha_vencimiento'] < fecha_fin]
saldo_vencido_final = facturas_vencidas_al_final['importe'].sum()
st.markdown("### Resumen del Período Seleccionado")
col1, col2, col3, col4 = st.columns(4)
# ... (y el resto del código hasta el final)
