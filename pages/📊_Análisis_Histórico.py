# ======================================================================================
# ARCHIVO: pages/📊_Análisis_Histórico.py (Versión Profesional)
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
    
    # Limpieza y conversión de tipos
    for col in ['fecha_documento', 'fecha_vencimiento', 'fecha_saldado']:
        df_historico[col] = pd.to_datetime(df_historico[col], errors='coerce')
    
    df_historico['importe'] = pd.to_numeric(df_historico['importe'], errors='coerce').fillna(0)
    
    # Cálculo de Días de Pago (DSO por factura)
    df_historico['dias_de_pago'] = (df_historico['fecha_saldado'] - df_historico['fecha_documento']).dt.days
    
    return df_historico

df_historico = cargar_y_procesar_historicos()

if df_historico.empty:
    st.warning("No se encontraron archivos de datos históricos con el formato 'Cartera_AAAA-MM.xlsx'.")
    st.info("Asegúrate de tener al menos dos archivos históricos en la carpeta principal para poder ver las tendencias.")
    st.stop()

# --- Filtros de Fecha ---
st.sidebar.header("Filtros de Análisis")
min_date = df_historico['fecha_documento'].min().date()
max_date = df_historico['fecha_documento'].max().date()

# Por defecto, analizamos los últimos 12 meses, si hay datos
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

# Convertir fechas de vuelta a datetime para filtrar
fecha_inicio = pd.to_datetime(fecha_inicio)
fecha_fin = pd.to_datetime(fecha_fin)

# --- Cálculos para KPIs y Gráficos ---

# 1. Ventas: Facturas emitidas en el período
ventas_periodo = df_historico[df_historico['fecha_documento'].between(fecha_inicio, fecha_fin)]
total_ventas = ventas_periodo['importe'].sum()

# 2. Cobros: Facturas saldadas en el período
cobros_periodo = df_historico[df_historico['fecha_saldado'].between(fecha_inicio, fecha_fin)]
total_cobrado = cobros_periodo['importe'].sum()

# 3. Rotación de Cartera (DSO): Días promedio de pago de lo que se cobró en el período
dso_periodo = cobros_periodo['dias_de_pago'].mean() if not cobros_periodo.empty else 0

# 4. Saldo Vencido al final del período
# Tomamos el snapshot más reciente dentro del rango para el saldo
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

# Preparar datos para gráficos mensuales
df_graficos = df_historico.copy()
df_graficos['mes_documento'] = df_graficos['fecha_documento'].dt.to_period('M').to_timestamp()
df_graficos['mes_saldado'] = df_graficos['fecha_saldado'].dt.to_period('M').to_timestamp()

# Ventas y cobros por mes
ventas_mes = df_graficos.groupby('mes_documento')['importe'].sum().reset_index().rename(columns={'mes_documento': 'mes', 'importe': 'Ventas'})
cobros_mes = df_graficos.groupby('mes_saldado')['importe'].sum().reset_index().rename(columns={'mes_saldado': 'mes', 'importe': 'Cobros'})

# DSO por mes
dso_mes = df_graficos.groupby('mes_saldado')['dias_de_pago'].mean().reset_index().rename(columns={'mes_saldado': 'mes', 'dias_de_pago': 'DSO'})

# Unir dataframes para graficar
df_final_graficos = pd.merge(ventas_mes, cobros_mes, on='mes', how='outer').fillna(0)
df_final_graficos = pd.merge(df_final_graficos, dso_mes, on='mes', how='outer')
df_final_graficos = df_final_graficos.sort_values('mes')
df_final_graficos_filtrado = df_final_graficos[df_final_graficos['mes'].between(fecha_inicio.to_period('M').to_timestamp(), fecha_fin.to_period('M').to_timestamp())]

# Gráfico 1: Flujo de Caja (Ventas vs. Cobros)
st.markdown("#### Flujo de Caja Mensual")
st.bar_chart(df_final_graficos_filtrado, x='mes', y=['Ventas', 'Cobros'], color=["#1f77b4", "#2ca02c"])

# Gráfico 2: Eficiencia de Cobro (DSO)
st.markdown("#### Eficiencia de Cobro (Rotación/DSO en días)")
st.line_chart(df_final_graficos_filtrado.set_index('mes')['DSO'])
