# ======================================================================================
# ARCHIVO: pages/🧑‍💼_Perfil_de_Cliente.py
# ======================================================================================
import streamlit as st
import pandas as pd
import glob
import re

st.set_page_config(page_title="Perfil de Cliente", page_icon="🧑‍💼", layout="wide")

st.title("🧑‍💼 Perfil de Pagador por Cliente")

@st.cache_data
def cargar_datos_historicos():
    mapa_columnas = {
        'Serie': 'serie', 'Número': 'numero', 'Fecha Documento': 'fecha_documento',
        'Fecha Vencimiento': 'fecha_vencimiento', 'Fecha Saldado': 'fecha_saldado',
        'NOMBRECLIENTE': 'nombrecliente', 'Población': 'poblacion', 'Provincia': 'provincia',
        'IMPORTE': 'importe', 'RIESGOCONCEDIDO': 'riesgoconcedido', 'NOMVENDEDOR': 'nomvendedor',
        'DIAS_VENCIDO': 'dias_vencido', 'Estado': 'estado'
    }
    lista_archivos = glob.glob("Cartera_*.xlsx")
    if not lista_archivos: return pd.DataFrame()
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
    if not lista_df: return pd.DataFrame()
    return pd.concat(lista_df, ignore_index=True)

df_historico_completo = cargar_datos_historicos()

if df_historico_completo.empty:
    st.warning("No se encontraron archivos de datos históricos con el formato 'Cartera_AAAA-MM.xlsx'.")
    st.info("Esta sección requiere datos históricos para funcionar.")
    st.stop()

# --- Buscador de Clientes ---
lista_clientes = sorted(df_historico_completo['nombrecliente'].dropna().unique())
cliente_sel = st.selectbox("Selecciona un cliente para analizar su comportamiento de pago:", [""] + lista_clientes)

if cliente_sel:
    df_cliente = df_historico_completo[df_historico_completo['nombrecliente'] == cliente_sel].copy()
    
    df_cliente['fecha_documento'] = pd.to_datetime(df_cliente['fecha_documento'], errors='coerce')
    df_cliente['fecha_saldado'] = pd.to_datetime(df_cliente['fecha_saldado'], errors='coerce')
    
    df_pagadas = df_cliente.dropna(subset=['fecha_saldado'])
    if not df_pagadas.empty:
        df_pagadas['dias_de_pago'] = (df_pagadas['fecha_saldado'] - df_pagadas['fecha_documento']).dt.days
    
    # --- Mostrar KPIs del Cliente ---
    st.markdown("---")
    st.subheader(f"Análisis de {cliente_sel}")

    if not df_pagadas.empty and df_pagadas['dias_de_pago'].notna().any():
        avg_dias_pago = df_pagadas['dias_de_pago'].mean()
        
        if avg_dias_pago <= 30: calificacion = "✅ Pagador Excelente"
        elif avg_dias_pago <= 60: calificacion = "👍 Pagador Bueno"
        elif avg_dias_pago <= 90: calificacion = "⚠️ Pagador Lento"
        else: calificacion = "🚨 Pagador de Riesgo"

        col1, col2 = st.columns(2)
        with col1:
            st.metric("Días Promedio de Pago", f"{avg_dias_pago:.0f} días")
        with col2:
            st.metric("Calificación", calificacion)
    else:
        st.info("Este cliente no tiene un historial de facturas pagadas para calcular su comportamiento.")

    # Mostrar historial de facturas
    st.markdown("---")
    st.subheader("Historial Completo de Facturas")
    st.dataframe(df_cliente)
