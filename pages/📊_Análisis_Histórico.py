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

# --- CORRECCIÓN: El nombre de la función ahora es consistente ---
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
    if not lista_archivos:
        return pd.DataFrame()

    lista_df = []
    for archivo in lista_archivos:
        try:
            match = re.search(r'(\d{4})-(\d{2})', archivo)
            if not match: continue

            df = pd.read_excel(archivo)
            
            if not df.empty:
                df = df.iloc[:-1]

            df['Serie'] = df['Serie'].astype(str)
            df = df[~df['Serie'].str.contains('W|X', case=False, na=False)]
            
            df.rename(columns=mapa_columnas, inplace=True)
            lista_df.append(df)
        except Exception as e:
            st.warning(f"No se pudo procesar el archivo {archivo}: {e}")
            
    if not lista_df:
        return pd.DataFrame()

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

# --- CORRECCIÓN: La llamada a la función ahora usa el nombre correcto ---
df_historico = cargar_datos_historicos()

if df_historico.empty:
    st.warning("No se encontraron archivos de datos históricos con el formato 'Cartera_AAAA-MM.xlsx'.")
    st.stop()

# --- El resto del script sigue igual ---
st.sidebar.header("Filtros de Análisis")
min_date = df_historico['fecha_documento'].min().date()
max_date_saldado = df_historico['fecha_saldado'].max()
max_date_doc = df_historico['fecha_documento'].max()
max_date = max(max_date_saldado, max_date_doc).date() if pd.notna(max_date_saldado) else max_date_doc.date()
default_start_date = max(min_date, max_date - relativedelta(months=12))
fecha_inicio, fecha_fin = st.sidebar.date_input("Selecciona el Rango de Fechas", (default_start_date, max_date), min_value=min_date, max_value=max_date)

if not fecha_inicio or not fecha_fin or fecha_inicio > fecha_fin:
    st.error("Por favor, selecciona un rango de fechas válido."); st.stop()

fecha_inicio, fecha_fin = pd.to_datetime(fecha_inicio), pd.to_datetime(fecha_fin)

df_periodo = df_historico[(df_historico['fecha_documento'] >= fecha_inicio) & (df_historico['fecha_documento'] <= fecha_fin) | (df_historico['fecha_saldado'] >= fecha_inicio) & (df_historico['fecha_saldado'] <= fecha_fin)]
if df_periodo.empty:
    st.warning("No hay datos de facturas emitidas o saldadas en el período de fechas seleccionado."); st.stop()

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
with col1: st.metric("📈 Ventas Emitidas", f"${total_ventas:,.0f}")
with col2: st.metric("✅ Total Cobrado", f"${total_cobrado:,.0f}")
with col3: st.metric("🔄 Rotación de Cartera (DSO)", f"{dso_periodo:.0f} días", help="Días promedio que se tardó en cobrar las facturas saldadas en este período.")
with col4: st.metric("🔥 Saldo Vencido al Final", f"${saldo_vencido_final:,.0f}", help=f"Cartera que quedó vencida al {fecha_fin.strftime('%Y-%m-%d')}")

st.markdown("#### Análisis y Conclusiones del Período")
st.markdown('<hr style="border:1px solid #e0e0e0">', unsafe_allow_html=True)
diferencia_flujo = total_cobrado - total_ventas
if diferencia_flujo >= 0:
    st.success(f"**✅ Gestión de Flujo Positiva:** En este período se ha cobrado **${diferencia_flujo:,.0f} más** de lo que se ha vendido.")
else:
    st.warning(f"**⚠️ Crecimiento de Cartera:** Las ventas han superado a los cobros por **${abs(diferencia_flujo):,.0f}**.")
if dso_periodo <= 30:
    st.success(f"**✅ Eficiencia Óptima:** La rotación de cartera de **{dso_periodo:.0f} días** es excelente.")
elif dso_periodo <= 60:
    st.info(f"**👍 Eficiencia Aceptable:** La rotación de **{dso_periodo:.0f} días** es buena.")
else:
    st.error(f"**🚨 Alerta de Eficiencia:** La rotación de **{dso_periodo:.0f} días** es elevada.")

st.markdown('<hr style="border:1px solid #e0e0e0">', unsafe_allow_html=True)
st.subheader("Análisis de Evolución Mensual")

df_graficos = df_periodo.copy()
df_graficos['mes_documento'] = pd.to_datetime(df_graficos['fecha_documento'].dt.strftime('%Y-%m-01'), errors='coerce')
df_graficos['mes_saldado'] = pd.to_datetime(df_graficos['fecha_saldado'].dt.strftime('%Y-%m-01'), errors='coerce')
ventas_mes = df_graficos.groupby('mes_documento')['importe'].sum().reset_index().rename(columns={'mes_documento': 'mes', 'importe': 'Ventas'})
cobros_mes = df_graficos.groupby('mes_saldado')['importe'].sum().reset_index().rename(columns={'mes_saldado': 'mes', 'importe': 'Cobros'})
dso_mes = df_graficos.groupby('mes_saldado')['dias_de_pago'].mean().reset_index().rename(columns={'mes_saldado': 'mes', 'dias_de_pago': 'DSO'})
df_final_graficos = pd.merge(ventas_mes, cobros_mes, on='mes', how='outer').fillna(0)
df_final_graficos = pd.merge(df_final_graficos, dso_mes, on='mes', how='outer')
df_final_graficos = df_final_graficos.sort_values('mes').reset_index(drop=True)

if not df_final_graficos.empty:
    st.markdown("#### Flujo de Caja Mensual (Ventas vs. Cobros)")
    st.bar_chart(df_final_graficos, x='mes', y=['Ventas', 'Cobros'], color=["#1f77b4", "#2ca02c"])
    st.markdown("#### Eficiencia de Cobro Mensual (Evolución del DSO en días)")
    st.line_chart(df_final_graficos.set_index('mes')['DSO'])
    if len(df_final_graficos) > 1:
        dso_filtrado = df_final_graficos['DSO'].dropna()
        if len(dso_filtrado) > 1:
            dso_inicial = dso_filtrado.iloc[0]
            dso_final = dso_filtrado.iloc[-1]
            cambio_dso = dso_final - dso_inicial
            st.markdown("##### Diagnóstico de la Tendencia de Eficiencia")
            if cambio_dso < -1:
                st.success(f"**Tendencia Positiva:** La eficiencia de cobro ha **mejorado**. El tiempo promedio se ha **reducido en {abs(cambio_dso):.0f} días**.")
            elif cambio_dso > 1:
                st.warning(f"**Tendencia a Revisar:** La eficiencia de cobro ha **disminuido**. Se tarda en promedio **{cambio_dso:.0f} días más** en cobrar.")
            else:
                st.info("**Tendencia Estable:** La eficiencia de cobro se ha mantenido estable.")
