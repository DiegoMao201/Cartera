# ======================================================================================
# ARCHIVO: pages/📊_Análisis_Histórico.py
# ======================================================================================
import streamlit as st
import pandas as pd
import glob
import re

st.set_page_config(page_title="Análisis Histórico", page_icon="📊", layout="wide")

st.title("📊 Análisis Histórico de Cartera")

# --- Función para normalizar nombres (debe ser consistente) ---
def normalizar_columna(nombre: str) -> str:
    # Una versión simplificada solo para estandarizar
    nombre = ''.join(c for c in re.sub(r'[^A-Z0-9_]', '', nombre.upper()) if not c.isdigit())
    return nombre.replace(' ', '_')

@st.cache_data
def cargar_datos_historicos():
    # Mapa de nombres de columnas de los archivos históricos al formato estándar
    mapa_columnas = {
        'Serie': 'serie', 'Número': 'numero', 'Fecha Documento': 'fecha_documento',
        'Fecha Vencimiento': 'fecha_vencimiento', 'Fecha Saldado': 'fecha_saldado',
        'NOMBRECLIENTE': 'nombrecliente', 'Población': 'poblacion', 'Provincia': 'provincia',
        'IMPORTE': 'importe', 'RIESGOCONCEDIDO': 'riesgoconcedido', 'NOMVENDEDOR': 'nomvendedor',
        'DIAS_VENCIDO': 'dias_vencido', 'Estado': 'estado'
    }
    
    lista_archivos = glob.glob("Cartera_*.xlsx")
    if not lista_archivos:
        return pd.DataFrame() # Devuelve un DataFrame vacío si no hay archivos históricos

    lista_df = []
    for archivo in lista_archivos:
        try:
            # Extraer la fecha del nombre del archivo
            match = re.search(r'(\d{4})-(\d{2})', archivo)
            if match:
                año, mes = map(int, match.groups())
                fecha_reporte = pd.to_datetime(f'{año}-{mes}-01')
                
                df = pd.read_excel(archivo)
                df.rename(columns=mapa_columnas, inplace=True)
                df['fecha_reporte'] = fecha_reporte
                lista_df.append(df)
        except Exception as e:
            st.warning(f"No se pudo procesar el archivo {archivo}: {e}")
            
    if not lista_df:
        return pd.DataFrame()

    df_historico = pd.concat(lista_df, ignore_index=True)
    df_historico['importe'] = pd.to_numeric(df_historico['importe'], errors='coerce').fillna(0)
    df_historico['dias_vencido'] = pd.to_numeric(df_historico['dias_vencido'], errors='coerce').fillna(0)
    
    return df_historico

df_historico = cargar_datos_historicos()

if df_historico.empty:
    st.warning("No se encontraron archivos de datos históricos con el formato 'Cartera_AAAA-MM.xlsx'.")
    st.info("Asegúrate de tener al menos dos archivos históricos en la carpeta principal para poder ver las tendencias.")
    st.stop()

# --- Cálculos para los gráficos ---
# Agrupamos por fecha de reporte para ver la evolución mensual
historico_mensual = df_historico.groupby('fecha_reporte').agg(
    cartera_total=('importe', 'sum'),
    cartera_vencida=('importe', lambda x: x[df_historico.loc[x.index, 'dias_vencido'] > 0].sum())
).reset_index()

historico_mensual = historico_mensual.sort_values('fecha_reporte')

st.subheader("Evolución de la Cartera Total vs. Vencida")
st.line_chart(historico_mensual, x='fecha_reporte', y=['cartera_total', 'cartera_vencida'])

st.subheader("Evolución del Porcentaje de Cartera Vencida")
historico_mensual['porc_vencido'] = (historico_mensual['cartera_vencida'] / historico_mensual['cartera_total'] * 100).fillna(0)
st.line_chart(historico_mensual, x='fecha_reporte', y='porc_vencido')

st.subheader("Datos Históricos Consolidados")
st.dataframe(df_historico)
