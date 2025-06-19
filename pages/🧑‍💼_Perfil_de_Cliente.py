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
    
    lista_archivos = sorted(glob.glob("Cartera_*.xlsx"))
    if not lista_archivos: return pd.DataFrame()
    
    lista_df = []
    for archivo in lista_archivos:
        try:
            df = pd.read_excel(archivo)
            if not df.empty:
                df = df.iloc[:-1]
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

df_historico_completo = cargar_datos_historicos()

if df_historico_completo.empty:
    st.warning("No se encontraron archivos de datos históricos con el formato 'Cartera_AAAA-MM.xlsx'.")
    st.stop()

# --- Buscador de Clientes ---
lista_clientes = sorted(df_historico_completo['nombrecliente'].dropna().unique())
cliente_sel = st.selectbox("Selecciona un cliente para analizar su comportamiento de pago:", [""] + lista_clientes)

if cliente_sel:
    df_cliente = df_historico_completo[df_historico_completo['nombrecliente'] == cliente_sel].copy()
    
    df_pagadas = df_cliente.dropna(subset=['dias_de_pago'])
    
    st.markdown("---")
    st.subheader(f"Análisis de {cliente_sel}")

    if not df_pagadas.empty:
        avg_dias_pago = df_pagadas['dias_de_pago'].mean()
        
        if avg_dias_pago <= 30: calificacion = "✅ Pagador Excelente"
        elif avg_dias_pago <= 60: calificacion = "👍 Pagador Bueno"
        elif avg_dias_pago <= 90: calificacion = "⚠️ Pagador Lento"
        else: calificacion = "🚨 Pagador de Riesgo"

        col1, col2 = st.columns(2)
        with col1:
            st.metric("Días Promedio de Pago", f"{avg_dias_pago:.0f} días", help="Promedio de días que tarda el cliente en pagar una factura desde su emisión.")
        with col2:
            st.metric("Calificación", calificacion)
    else:
        st.info("Este cliente no tiene un historial de facturas pagadas para calcular su comportamiento.")

    st.markdown("---")
    st.subheader("Historial Completo de Facturas del Cliente")
    st.dataframe(df_cliente[['numero', 'fecha_documento', 'fecha_vencimiento', 'fecha_saldado', 'dias_de_pago', 'importe']].sort_values(by="fecha_documento", ascending=False))
