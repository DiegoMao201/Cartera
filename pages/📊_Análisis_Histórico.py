# ======================================================================================
# ARCHIVO: pages/📊_Análisis_Histórico.py (Versión Mejorada)
# ======================================================================================
import streamlit as st
import pandas as pd
import glob
import re
from datetime import datetime
from dateutil.relativedelta import relativedelta
import unicodedata
import plotly.graph_objects as go # --- MEJORA ---

st.set_page_config(page_title="Análisis Histórico", page_icon="📊", layout="wide")

if 'authentication_status' not in st.session_state or not st.session_state['authentication_status']:
    st.warning("Por favor, inicie sesión en el 📈 Tablero Principal para acceder a esta página.")
    st.stop()

def normalizar_nombre(nombre: str) -> str:
    if not isinstance(nombre, str): return ""
    nombre = nombre.upper().strip().replace('.', '')
    nombre = ''.join(c for c in unicodedata.normalize('NFD', nombre) if unicodedata.category(c) != 'Mn')
    return ' '.join(nombre.split())

@st.cache_data
def cargar_datos_historicos():
    # Tu función de carga de datos original (está bien construida)
    mapa_columnas = {
        'Serie': 'serie', 'Número': 'numero', 'Fecha Documento': 'fecha_documento',
        'Fecha Vencimiento': 'fecha_vencimiento', 'Fecha Saldado': 'fecha_saldado',
        'NOMBRECLIENTE': 'nombrecliente', 'Población': 'poblacion', 'Provincia': 'provincia',
        'IMPORTE': 'importe', 'RIESGOCONCEDIDO': 'riesgoconcedido', 'NOMVENDEDOR': 'nomvendedor',
        'DIAS_VENCIDO': 'dias_vencido', 'Estado': 'estado', 'Cod. Cliente': 'cod_cliente',
        'e-mail': 'e_mail'
    }
    lista_archivos = sorted(glob.glob("Cartera_*.xlsx"))
    if not lista_archivos: return pd.DataFrame()
    lista_df = []
    for archivo in lista_archivos:
        try:
            df = pd.read_excel(archivo)
            if not df.empty: df = df.iloc[:-1]
            if 'e-mail' not in df.columns: df['e-mail'] = None
            if 'Cod. Cliente' not in df.columns: df['Cod. Cliente'] = None
            df['Serie'] = df['Serie'].astype(str)
            df = df[~df['Serie'].str.contains('W|X', case=False, na=False)]
            df.rename(columns=mapa_columnas, inplace=True)
            lista_df.append(df)
        except Exception as e:
            st.warning(f"No se pudo procesar el archivo {archivo}: {e}")
    if not lista_df: return pd.DataFrame()
    
    df_completo = pd.concat(lista_df, ignore_index=True)
    df_completo.dropna(subset=['numero', 'nombrecliente'], inplace=True)
    df_completo['nomvendedor_norm'] = df_completo['nomvendedor'].apply(normalizar_nombre)
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

st.title("📊 Análisis Histórico de Cartera")
df_historico_base = cargar_datos_historicos()

if df_historico_base.empty:
    st.warning("No se encontraron archivos de datos históricos."); st.stop()

# --- FILTROS (Sin cambios) ---
st.sidebar.header("Filtros de Análisis")
acceso_general = st.session_state.get('acceso_general', False)
vendedor_autenticado = st.session_state.get('vendedor_autenticado', None)
if acceso_general:
    vendedores = ["Todos"] + sorted(df_historico_base['nomvendedor'].dropna().unique())
    vendedor_sel_hist = st.sidebar.selectbox("Vendedor:", vendedores)
else:
    vendedor_sel_hist = vendedor_autenticado
if vendedor_sel_hist == "Todos":
    df_historico = df_historico_base.copy()
else:
    df_historico = df_historico_base[df_historico_base['nomvendedor_norm'] == normalizar_nombre(vendedor_sel_hist)].copy()

if df_historico.empty or df_historico['fecha_documento'].isnull().all():
    st.warning("No hay datos para el vendedor seleccionado en el historial."); st.stop()

min_date = df_historico['fecha_documento'].min().date()
max_date_saldado = df_historico['fecha_saldado'].max()
max_date_doc = df_historico['fecha_documento'].max()
max_date = max(max_date_saldado, max_date_doc).date() if pd.notna(max_date_saldado) else max_date_doc.date()
default_start_date = max(min_date, max_date - relativedelta(months=12))
fecha_inicio, fecha_fin = st.sidebar.date_input("Rango de Fechas:", (default_start_date, max_date), min_value=min_date, max_value=max_date)

if not fecha_inicio or not fecha_fin or fecha_inicio > fecha_fin:
    st.error("Por favor, selecciona un rango de fechas válido."); st.stop()

fecha_inicio, fecha_fin = pd.to_datetime(fecha_inicio), pd.to_datetime(fecha_fin)
df_periodo = df_historico[
    (df_historico['fecha_documento'] >= fecha_inicio) & (df_historico['fecha_documento'] <= fecha_fin) |
    (df_historico['fecha_saldado'] >= fecha_inicio) & (df_historico['fecha_saldado'] <= fecha_fin)
].copy()

if df_periodo.empty:
    st.warning("No hay datos de facturas emitidas o saldadas en el período de fechas seleccionado."); st.stop()

# --- MEJORA: Cálculo de KPIs Financieros ---
# Cartera abierta al inicio del período
snapshot_inicial = df_historico[df_historico['fecha_documento'] < fecha_inicio]
saldo_inicial = snapshot_inicial[(snapshot_inicial['fecha_saldado'].isnull()) | (snapshot_inicial['fecha_saldado'] >= fecha_inicio)]['importe'].sum()

# Ventas y cobros dentro del período
ventas_periodo = df_historico[df_historico['fecha_documento'].between(fecha_inicio, fecha_fin)]
total_ventas = ventas_periodo['importe'].sum()
cobros_periodo = df_historico[df_historico['fecha_saldado'].between(fecha_inicio, fecha_fin)]
total_cobrado = cobros_periodo['importe'].sum()

# Cartera abierta al final del período
snapshot_final = df_historico[df_historico['fecha_documento'] <= fecha_fin]
facturas_abiertas_al_final = snapshot_final[(snapshot_final['fecha_saldado'].isnull()) | (snapshot_final['fecha_saldado'] > fecha_fin)]
saldo_final_total = facturas_abiertas_al_final['importe'].sum()
facturas_vencidas_al_final = facturas_abiertas_al_final[facturas_abiertas_al_final['fecha_vencimiento'] < fecha_fin]
saldo_vencido_final = facturas_vencidas_al_final['importe'].sum()

# KPIs
dso_periodo = cobros_periodo['dias_de_pago'].mean() if not cobros_periodo.empty else 0
flujo_neto = total_cobrado - total_ventas
universo_cobrable = saldo_inicial + total_ventas
cer = (total_cobrado / universo_cobrable) * 100 if universo_cobrable > 0 else 0
indice_morosidad = (saldo_vencido_final / saldo_final_total) * 100 if saldo_final_total > 0 else 0

st.markdown("### Diagnóstico Financiero del Período")
col1, col2, col3, col4 = st.columns(4)
with col1: st.metric("💰 Eficiencia de Cobro (CER)", f"{cer:.1f}%", help="Porcentaje cobrado del total que se debía cobrar (Saldo Inicial + Ventas). Más alto es mejor.")
with col2: st.metric("🔥 Índice de Morosidad", f"{indice_morosidad:.1f}%", help="Porcentaje de la cartera pendiente que está vencida al final del período. Más bajo es mejor.")
with col3: st.metric("🔄 Rotación (DSO)", f"{dso_periodo:.0f} días", help="Días promedio que se tardó en cobrar las facturas saldadas en este período.")
with col4: st.metric("🌊 Flujo Neto de Cartera", f"${flujo_neto:,.0f}", help="Cobros (-) Ventas. Positivo significa que entró más dinero del que salió en nuevas facturas.")

# --- MEJORA: Diagnóstico con IA más profundo ---
st.markdown("#### Asistente de Diagnóstico IA")
st.markdown('<hr style="border:1px solid #e0e0e0">', unsafe_allow_html=True)
if cer > 85: st.success(f"**✅ Excelente Eficiencia de Cobro ({cer:.1f}%):** La gestión ha sido muy efectiva, recuperando una alta proporción de la cartera cobrable.")
elif cer > 70: st.info(f"**👍 Buena Eficiencia de Cobro ({cer:.1f}%):** Se ha recuperado una parte importante de la cartera. Hay margen para optimizar.")
else: st.warning(f"**⚠️ Baja Eficiencia de Cobro ({cer:.1f}%):** La recuperación está por debajo de lo óptimo. Es crucial revisar estrategias de cobro.")

if indice_morosidad < 15: st.success(f"**✅ Cartera Saludable ({indice_morosidad:.1f}%):** El nivel de morosidad es bajo, indicando una cartera de clientes de buena calidad.")
elif indice_morosidad < 30: st.info(f"**👍 Cartera Controlada ({indice_morosidad:.1f}%):** La morosidad es manejable, pero requiere monitoreo constante.")
else: st.error(f"**🚨 Cartera de Riesgo ({indice_morosidad:.1f}%):** Un alto porcentaje de la cartera está en mora. Requiere acción inmediata para mitigar pérdidas.")

if flujo_neto < 0 and cer < 70:
    st.error("**🔥 ALERTA CRÍTICA DE FLUJO:** La cartera está creciendo sin un respaldo de cobros eficientes. El riesgo de liquidez es alto.")

st.markdown('<hr style="border:1px solid #e0e0e0">', unsafe_allow_html=True)

# --- MEJORA: Nueva sección de Análisis de Tendencias Rodantes ---
st.subheader("Análisis de Tendencias Rodantes (Medias Móviles de 3 Meses)")

# Preparación de datos mensuales
df_graficos = df_periodo.copy()
df_graficos['mes_documento'] = pd.to_datetime(df_graficos['fecha_documento'].dt.strftime('%Y-%m-01'), errors='coerce')
df_graficos['mes_saldado'] = pd.to_datetime(df_graficos['fecha_saldado'].dt.strftime('%Y-%m-01'), errors='coerce')

ventas_mes = df_graficos.groupby('mes_documento')['importe'].sum()
cobros_mes = df_graficos.groupby('mes_saldado')['importe'].sum()
dso_mes = df_graficos.groupby('mes_saldado')['dias_de_pago'].mean()

df_final_graficos = pd.concat([ventas_mes, cobros_mes, dso_mes], axis=1).fillna(0)
df_final_graficos.index.name = 'mes'
df_final_graficos.columns = ['Ventas', 'Cobros', 'DSO']
df_final_graficos = df_final_graficos.sort_index().reset_index()
df_final_graficos = df_final_graficos[df_final_graficos['mes'].between(fecha_inicio, fecha_fin)]

# Calcular tendencias rodantes
df_final_graficos['DSO_tendencia'] = df_final_graficos['DSO'].replace(0, pd.NA).rolling(window=3, min_periods=1, center=True).mean()
df_final_graficos['CER_mes'] = (df_final_graficos['Cobros'] / df_final_graficos['Ventas'].replace(0, 1)) * 100 # CER simplificado mensual
df_final_graficos['CER_tendencia'] = df_final_graficos['CER_mes'].rolling(window=3, min_periods=1, center=True).mean()

if not df_final_graficos.empty:
    chart1, chart2 = st.columns(2)
    with chart1:
        st.markdown("#### Tendencia de Rotación (DSO)")
        fig = go.Figure()
        fig.add_trace(go.Bar(x=df_final_graficos['mes'], y=df_final_graficos['DSO'], name='DSO Mensual', marker_color='lightblue'))
        fig.add_trace(go.Scatter(x=df_final_graficos['mes'], y=df_final_graficos['DSO_tendencia'], name='Tendencia (3 Meses)', mode='lines', line=dict(color='darkblue', width=3)))
        fig.update_layout(title_text='Evolución del DSO vs. Tendencia Rodante', yaxis_title='Días')
        st.plotly_chart(fig, use_container_width=True)

    with chart2:
        st.markdown("#### Tendencia de Eficiencia de Cobro")
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(x=df_final_graficos['mes'], y=df_final_graficos['CER_mes'], name='Eficiencia Mensual', marker_color='lightgreen'))
        fig2.add_trace(go.Scatter(x=df_final_graficos['mes'], y=df_final_graficos['CER_tendencia'], name='Tendencia (3 Meses)', mode='lines', line=dict(color='darkgreen', width=3)))
        fig2.update_layout(title_text='Eficiencia de Cobro vs. Tendencia Rodante', yaxis_title='Eficiencia (%)')
        st.plotly_chart(fig2, use_container_width=True)

    # Diagnóstico de la tendencia
    st.markdown("##### Diagnóstico de las Tendencias")
    dso_tendencia = df_final_graficos['DSO_tendencia'].dropna()
    if len(dso_tendencia) > 1:
        cambio_dso = dso_tendencia.iloc[-1] - dso_tendencia.iloc[0]
        if cambio_dso < -2:
            st.success(f"**📈 Tendencia de DSO positiva:** La velocidad de cobro está mejorando consistentemente (reducción de {abs(cambio_dso):.0f} días).")
        elif cambio_dso > 2:
            st.warning(f"**📉 Tendencia de DSO a revisar:** El tiempo para cobrar está aumentando de forma sostenida (aumento de {cambio_dso:.0f} días).")
        else:
            st.info("**⏸️ Tendencia de DSO estable:** La rotación de cartera se mantiene sin cambios significativos.")
else:
    st.info("No hay suficientes datos mensuales en el período para generar gráficos de tendencia.")
