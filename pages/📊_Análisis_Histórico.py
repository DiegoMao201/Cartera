# ======================================================================================
# ARCHIVO: pages/📊_Análisis_Histórico.py (Versión Corregida)
# ======================================================================================
import streamlit as st
import pandas as pd
import glob
import re
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Análisis Histórico", page_icon="📊", layout="wide")

st.title("📊 Análisis Histórico de Cartera")

@st.cache_data
def cargar_y_procesar_historicos():
    mapa_columnas = {
        'Serie': 'serie', 'Número': 'numero', 'Fecha Documento': 'fecha_documento',
        'Fecha Vencimiento': 'fecha_vencimiento', 'Fecha Saldado': 'fecha_saldado',
        'NOMBRECLIENTE': 'nombrecliente', 'Población': 'poblacion', 'Provincia': 'provincia',
        'IMPORTE': 'importe', 'RIESGOCONCEDIDO': 'riesgoconcedido', 'NOMVENDEDOR': 'nomvendedor',
        'DIAS_VENCIDO': 'dias_vencido', 'Estado': 'estado'
    }
    
    lista_archivos = glob.glob("Cartera_*.xlsx")
    if not lista_archivos:
        return pd.DataFrame()

    lista_df = []
    for archivo in lista_archivos:
        try:
            df = pd.read_excel(archivo)
            df['Serie'] = df['Serie'].astype(str)
            df = df[~df['Serie'].str.contains('W|X', case=False, na=False)]
            df.rename(columns=mapa_columnas, inplace=True)
            lista_df.append(df)
        except Exception as e:
            st.warning(f"No se pudo procesar el archivo {archivo}: {e}")
            
    if not lista_df:
        return pd.DataFrame()

    df_historico = pd.concat(lista_df, ignore_index=True).drop_duplicates()
    
    for col in ['fecha_documento', 'fecha_vencimiento', 'fecha_saldado']:
        df_historico[col] = pd.to_datetime(df_historico[col], errors='coerce')
    
    df_historico['importe'] = pd.to_numeric(df_historico['importe'], errors='coerce').fillna(0)
    
    df_historico_pagadas = df_historico.dropna(subset=['fecha_saldado', 'fecha_documento']).copy()
    df_historico_pagadas['dias_de_pago'] = (df_historico_pagadas['fecha_saldado'] - df_historico_pagadas['fecha_documento']).dt.days
    
    return df_historico_pagadas

df_historico = cargar_y_procesar_historicos()

if df_historico.empty:
    st.warning("No se encontraron archivos de datos históricos con el formato 'Cartera_AAAA-MM.xlsx'.")
    st.info("Asegúrate de tener al menos dos archivos históricos en la carpeta principal para poder ver las tendencias.")
    st.stop()

# --- Filtros de Fecha ---
st.sidebar.header("Filtros de Análisis")
min_date = df_historico['fecha_documento'].min().date()
max_date = df_historico['fecha_documento'].max().date()

default_start_date = max(min_date, max_date - relativedelta(months=12))

fecha_inicio, fecha_fin = st.sidebar.date_input(
    "Selecciona el Rango de Fechas",
    (default_start_date, max_date),
    min_value=min_date,
    max_value=max_date
)

if not fecha_inicio or not fecha_fin or fecha_inicio > fecha_fin:
    st.error("Por favor, selecciona un rango de fechas válido.")
    st.stop()

fecha_inicio = pd.to_datetime(fecha_inicio)
fecha_fin = pd.to_datetime(fecha_fin)

# --- Cálculos para KPIs y Gráficos ---
ventas_periodo = df_historico[df_historico['fecha_documento'].between(fecha_inicio, fecha_fin)]
total_ventas = ventas_periodo['importe'].sum()

cobros_periodo = df_historico[df_historico['fecha_saldado'].between(fecha_inicio, fecha_fin)]
total_cobrado = cobros_periodo['importe'].sum()

dso_periodo = cobros_periodo['dias_de_pago'].mean() if not cobros_periodo.empty else 0

snapshot_final = df_historico[df_historico['fecha_documento'] <= fecha_fin]
saldo_vencido_final = snapshot_final[snapshot_final['fecha_saldado'].isnull() & (snapshot_final['fecha_vencimiento'] < fecha_fin)]['importe'].sum()

# --- Renderizado de KPIs ---
st.markdown("### Resumen del Período Seleccionado")
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("📈 Ventas Emitidas", f"${total_ventas:,.0f}")
with col2:
    st.metric("✅ Total Cobrado", f"${total_cobrado:,.0f}")
with col3:
    st.metric("🔄 Rotación de Cartera (DSO)", f"{dso_periodo:.0f} días", help="Días promedio que se tardó en cobrar las facturas saldadas en este período.")
with col4:
    st.metric("🔥 Saldo Vencido al Final", f"${saldo_vencido_final:,.0f}", help=f"Cartera que quedó vencida al {fecha_fin.strftime('%Y-%m-%d')}")

st.markdown("---")

# --- Gráficos de Evolución Mensual ---
st.subheader("Análisis de Evolución Mensual")

df_graficos = df_historico.copy()

# --- MODIFICACIÓN: Forma más robusta de agrupar por mes ---
# La línea anterior (.dt.to_period('M').to_timestamp()) causaba un error de tipo.
# Esta nueva forma es más estable y logra el mismo resultado.
df_graficos['mes_documento'] = df_graficos['fecha_documento'].dt.to_period('M').start_time
df_graficos['mes_saldado'] = df_graficos['fecha_saldado'].dt.to_period('M').start_time

ventas_mes = df_graficos.groupby('mes_documento')['importe'].sum().reset_index().rename(columns={'mes_documento': 'mes', 'importe': 'Ventas'})
cobros_mes = df_graficos.groupby('mes_saldado')['importe'].sum().reset_index().rename(columns={'mes_saldado': 'mes', 'importe': 'Cobros'})
dso_mes = df_graficos.groupby('mes_saldado')['dias_de_pago'].mean().reset_index().rename(columns={'mes_saldado': 'mes', 'dias_de_pago': 'DSO'})

df_final_graficos = pd.merge(ventas_mes, cobros_mes, on='mes', how='outer').fillna(0)
df_final_graficos = pd.merge(df_final_graficos, dso_mes, on='mes', how='outer')
df_final_graficos = df_final_graficos.sort_values('mes')
df_final_graficos_filtrado = df_final_graficos[df_final_graficos['mes'].between(fecha_inicio, fecha_fin)]

st.markdown("#### Flujo de Caja Mensual")
st.bar_chart(df_final_graficos_filtrado, x='mes', y=['Ventas', 'Cobros'], color=["#1f77b4", "#2ca02c"])

st.markdown("#### Eficiencia de Cobro (Rotación/DSO en días)")
st.line_chart(df_final_graficos_filtrado.set_index('mes')['DSO'])
