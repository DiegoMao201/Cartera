# ======================================================================================
# SISTEMA INTEGRAL DE GESTIÓN DE CARTERA Y COBRANZA - FERREINOX SAS BIC (V. ULTRA)
# ======================================================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import os
import glob
import re
import unicodedata
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.worksheet.table import Table, TableStyleInfo
from urllib.parse import quote

# --- 1. CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(
    page_title="Centro de Mando: Cobranza Estratégica",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS Avanzados para separar visualmente las secciones
st.markdown("""
<style>
    .stApp { background-color: #f0f2f6; }
    .metric-card {
        background-color: white; padding: 20px; border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1); text-align: center;
        border-top: 5px solid #003366;
    }
    .big-font { font-size: 24px !important; font-weight: bold; color: #003366; }
    .status-badge { padding: 4px 8px; border-radius: 4px; font-weight: bold; color: white; }
    
    /* Pestañas personalizadas */
    div[data-testid="stTabs"] button { font-weight: bold; font-size: 16px; }
</style>
""", unsafe_allow_html=True)

# ======================================================================================
# --- 2. MOTOR DE DATOS (Ingestión y Limpieza) ---
# ======================================================================================

def normalizar_texto(texto):
    if not isinstance(texto, str): return str(texto)
    return unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode("utf-8").upper().strip()

def limpiar_moneda(valor):
    if pd.isna(valor): return 0.0
    s_val = str(valor).strip()
    s_val = re.sub(r'[^\d.,-]', '', s_val) # Quitar símbolos
    if not s_val: return 0.0
    try:
        # Lógica para detectar si es 1.000,00 (Latino) o 1,000.00 (USA)
        if ',' in s_val and '.' in s_val:
            if s_val.rfind(',') > s_val.rfind('.'): # Caso Latino
                s_val = s_val.replace('.', '').replace(',', '.')
            else: # Caso USA
                s_val = s_val.replace(',', '')
        elif ',' in s_val:
            parts = s_val.split(',')
            if len(parts[-1]) != 3: s_val = s_val.replace(',', '.') # Es decimal
            else: s_val = s_val.replace(',', '') # Son miles
        return float(s_val)
    except: return 0.0

@st.cache_data(ttl=300)
def cargar_datos():
    """Carga archivos locales (Excel o CSV) automáticamente."""
    df = pd.DataFrame()
    archivos = glob.glob("Cartera*.xlsx") + glob.glob("Cartera*.csv")
    
    if not archivos:
        return pd.DataFrame(), "No se encontró archivo 'Cartera...'"
    
    archivo = max(archivos, key=os.path.getctime) # El más reciente
    try:
        if archivo.endswith('.csv'):
            df = pd.read_csv(archivo, dtype=str, encoding='latin-1')
        else:
            df = pd.read_excel(archivo, dtype=str)
    except Exception as e:
        return pd.DataFrame(), f"Error leyendo archivo: {e}"

    # Mapeo Inteligente de Columnas
    cols_map = {
        'cliente': ['nombre', 'cliente', 'razon social', 'tercero'],
        'nit': ['nit', 'identificacion', 'cedula'],
        'saldo': ['saldo', 'importe', 'total', 'valor'],
        'dias_mora': ['dias', 'mora', 'vencido', 'antiguedad'],
        'telefono': ['tel', 'celular', 'movil'],
        'vendedor': ['vendedor', 'asesor', 'comercial'],
        'email': ['mail', 'correo'],
        'fecha_venc': ['vencimiento', 'fecha venc']
    }
    
    df.columns = [normalizar_texto(c) for c in df.columns]
    renombres = {}
    
    for key, patterns in cols_map.items():
        for col in df.columns:
            if any(p.upper() in col for p in patterns):
                renombres[col] = key
                break
    
    df.rename(columns=renombres, inplace=True)
    
    # Validar columnas mínimas
    req = ['cliente', 'saldo', 'dias_mora']
    if not all(c in df.columns for c in req):
        return pd.DataFrame(), f"Faltan columnas clave. Detectadas: {list(df.columns)}"

    # Limpieza de tipos
    df['saldo'] = df['saldo'].apply(limpiar_moneda)
    df['dias_mora'] = pd.to_numeric(df['dias_mora'], errors='coerce').fillna(0)
    df['cliente'] = df['cliente'].fillna("Desconocido").astype(str)
    
    # Asegurar campos opcionales
    for c in ['telefono', 'email', 'vendedor', 'nit']:
        if c not in df.columns: df[c] = 'N/A'
            
    return df[df['saldo'] != 0], f"Cargado: {archivo}"

# ======================================================================================
# --- 3. CEREBRO DE ESTRATEGIA Y MENSAJES ---
# ======================================================================================

def segmentar_cartera(df):
    """Clasifica al cliente y genera el mensaje de WhatsApp perfecto."""
    
    def generar_mensaje(row):
        cliente = str(row['cliente']).split()[0].title() # Primer nombre bonito
        saldo = f"${row['saldo']:,.0f}"
        dias = row['dias_mora']
        
        if dias <= 0:
            return f"Hola {cliente}, de Ferreinox. Esperamos que estés muy bien. Te confirmamos que tu estado de cuenta está al día. ¡Gracias por tu puntualidad!"
        elif dias <= 15:
            return f"Hola {cliente}, un saludo cordial de Ferreinox. Te recordamos amablemente un saldo pendiente de {saldo} vencido hace {int(dias)} días. Agradecemos tu gestión."
        elif dias <= 30:
            return f"Hola {cliente}. En Ferreinox valoramos tu crédito. Notamos una factura de {saldo} con {int(dias)} días de vencimiento. ¿Nos ayudas con la fecha de pago para actualizar el sistema?"
        elif dias <= 60:
            return f"IMPORTANTE {cliente}: Su cuenta presenta {int(dias)} días de mora por {saldo}. Por favor contáctenos hoy para evitar suspensión de despachos."
        else:
            return f"URGENTE {cliente}: Cartera en estado PRE-JURÍDICO. Saldo: {saldo} ({int(dias)} días). Evite reporte negativo y costos de abogados gestionando su pago hoy."

    def clasificar(dias):
        if dias <= 0: return "✅ Al Día"
        if dias <= 30: return "🟡 Preventivo"
        if dias <= 60: return "🟠 Administrativo"
        if dias <= 90: return "🔴 Pre-Jurídico"
        return "⚫ Castigo/Abogado"

    df['Estado'] = df['dias_mora'].apply(clasificar)
    df['Mensaje_WhatsApp'] = df.apply(generar_mensaje, axis=1)
    
    # Generar Link de WhatsApp
    def crear_link(row):
        tel = str(row['telefono'])
        tel = re.sub(r'\D', '', tel) # Solo números
        if len(tel) < 10: return None
        if not tel.startswith('57'): tel = '57' + tel # Asumir Colombia
        msg = quote(row['Mensaje_WhatsApp'])
        return f"https://wa.me/{tel}?text={msg}"

    df['Link_WA'] = df.apply(crear_link, axis=1)
    return df

# ======================================================================================
# --- 4. INTERFAZ PRINCIPAL (DASHBOARD) ---
# ======================================================================================

def main():
    st.markdown("<h1 style='text-align: center; color: #003366;'>🛡️ Centro de Gestión de Cartera Ferreinox</h1>", unsafe_allow_html=True)
    
    # 1. Carga de Datos
    df_raw, status_msg = cargar_datos()
    
    if df_raw.empty:
        st.error(f"❌ {status_msg}")
        st.info("Sube un archivo Excel llamado 'Cartera.xlsx' en la misma carpeta.")
        with st.expander("Ver formato de archivo requerido"):
            st.write("El Excel debe tener columnas como: Cliente, Nit, Saldo, Dias Mora, Telefono.")
        st.stop()
    
    df = segmentar_cartera(df_raw)

    # 2. Sidebar de Filtros
    with st.sidebar:
        st.image("https://cdn-icons-png.flaticon.com/512/2503/2503657.png", width=80)
        st.markdown("### 🔍 Filtros Globales")
        
        vendedores = ["TODOS"] + sorted(list(df['vendedor'].astype(str).unique()))
        filtro_vendedor = st.selectbox("Vendedor / Zona", vendedores)
        
        if filtro_vendedor != "TODOS":
            df = df[df['vendedor'] == filtro_vendedor]

        st.markdown("---")
        st.markdown("### 📊 Descargas")
        # Generar Excel Simple
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Cartera_Gestionada', index=False)
        
        st.download_button(
            label="📥 Bajar Base Completa",
            data=buffer,
            file_name="Cartera_Procesada.xlsx",
            mime="application/vnd.ms-excel"
        )

    # 3. KPIs Generales
    total = df['saldo'].sum()
    vencido = df[df['dias_mora'] > 0]['saldo'].sum()
    aldia = total - vencido
    pct_mora = (vencido/total)*100 if total > 0 else 0
    
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown(f"<div class='metric-card'><h3>💰 Total Cartera</h3><p class='big-font'>${total:,.0f}</p></div>", unsafe_allow_html=True)
    k2.markdown(f"<div class='metric-card'><h3>🔥 Vencido (Mora)</h3><p class='big-font' style='color:#b71c1c'>${vencido:,.0f}</p></div>", unsafe_allow_html=True)
    k3.markdown(f"<div class='metric-card'><h3>✅ Al Día (Corriente)</h3><p class='big-font' style='color:#2e7d32'>${aldia:,.0f}</p></div>", unsafe_allow_html=True)
    k4.markdown(f"<div class='metric-card'><h3>📉 Índice de Mora</h3><p class='big-font'>{pct_mora:.1f}%</p></div>", unsafe_allow_html=True)

    st.write("---")

    # 4. Pestañas de Gestión
    tab_cobro, tab_prev, tab_analisis = st.tabs(["🚨 GESTIÓN DE COBRANZA", "✅ PREVENTIVO / AL DÍA", "📈 INTELEIGENCIA"])

    # --- TAB A: COBRANZA (Mora > 0) ---
    with tab_cobro:
        st.subheader("⚔️ Sala de Guerra: Clientes en Mora")
        
        df_mora = df[df['dias_mora'] > 0].copy()
        df_mora = df_mora.sort_values(by=['dias_mora', 'saldo'], ascending=[False, False])
        
        # Filtro rápido por rango
        rango_filtro = st.radio("Filtrar por gravedad:", ["Todos", "1-30 Días", "31-60 Días", "> 60 Días (Crítico)"], horizontal=True)
        
        if rango_filtro == "1-30 Días": df_mora = df_mora[df_mora['dias_mora'] <= 30]
        elif rango_filtro == "31-60 Días": df_mora = df_mora[(df_mora['dias_mora'] > 30) & (df_mora['dias_mora'] <= 60)]
        elif rango_filtro == "> 60 Días (Crítico)": df_mora = df_mora[df_mora['dias_mora'] > 60]

        # Configuración de columnas para mostrar el enlace de WhatsApp bonito
        st.data_editor(
            df_mora[['cliente', 'saldo', 'dias_mora', 'Estado', 'Link_WA', 'telefono', 'vendedor']],
            column_config={
                "Link_WA": st.column_config.LinkColumn(
                    "📱 Acción WhatsApp",
                    help="Clic para abrir WhatsApp Web",
                    validate="^https://wa\.me/.*",
                    display_text="💬 ENVIAR COBRO"
                ),
                "saldo": st.column_config.NumberColumn("Deuda Total", format="$ %d"),
                "dias_mora": st.column_config.ProgressColumn(
                    "Días Mora", min_value=0, max_value=120, format="%f días"
                ),
            },
            hide_index=True,
            use_container_width=True,
            height=600
        )

    # --- TAB B: PREVENTIVO (Mora <= 0) ---
    with tab_prev:
        st.subheader("🤝 Fidelización y Recordatorios (Clientes al día)")
        st.info("Estos clientes no deben nada vencido. Úsalos para: 1. Agradecer pago 2. Ofrecer nuevos productos 3. Recordar factura próxima a vencer.")
        
        df_aldia = df[df['dias_mora'] <= 0].copy()
        df_aldia = df_aldia.sort_values(by='fecha_venc', ascending=True) # Mostrar próximos a vencer
        
        st.data_editor(
            df_aldia[['cliente', 'saldo', 'fecha_venc', 'Link_WA', 'telefono', 'vendedor']],
            column_config={
                "Link_WA": st.column_config.LinkColumn(
                    "📱 Contactar",
                    display_text="👋 SALUDAR"
                ),
                "saldo": st.column_config.NumberColumn("Saldo Corriente", format="$ %d"),
            },
            hide_index=True,
            use_container_width=True
        )

    # --- TAB C: ANALYTICS ---
    with tab_analisis:
        col1, col2 = st.columns(2)
        
        with col1:
            # Gráfico de Pastel
            fig_pie = px.pie(df, values='saldo', names='Estado', title='Distribución de Cartera por Estado', hole=0.4, color_discrete_sequence=px.colors.sequential.RdBu)
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with col2:
            # Gráfico de Barras Top Deudores
            top_10 = df.sort_values('saldo', ascending=False).head(10)
            fig_bar = px.bar(top_10, x='saldo', y='cliente', orientation='h', title='Top 10 Clientes con Mayor Deuda', text_auto='.2s')
            fig_bar.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_bar, use_container_width=True)

if __name__ == '__main__':
    main()
