# ======================================================================================
# SISTEMA INTEGRAL DE GESTIÓN DE CARTERA Y COBRANZA (V. CORREGIDA)
# Autor: Gemini AI para Ferreinox SAS BIC
# ======================================================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import dropbox
import io
import os
import glob
import re
import unicodedata
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from urllib.parse import quote

# Intentar importar yagmail de forma segura
try:
    import yagmail
except ImportError:
    yagmail = None

# --- 1. CONFIGURACIÓN DE LA PÁGINA Y ESTILOS ---
st.set_page_config(
    page_title="Centro de Mando: Cobranza Estratégica",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS
st.markdown("""
<style>
    .stApp { background-color: #f8f9fa; font-family: 'Segoe UI', sans-serif; }
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border-left: 5px solid #0d6efd;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }
    h1, h2, h3 { color: #003366; }
    .action-btn-wa {
        background-color: #25D366; color: white !important;
        padding: 8px 16px; border-radius: 50px; text-decoration: none;
        font-weight: 600; display: inline-block; border: 1px solid #1da851;
    }
    .action-btn-email {
        background-color: #EA4335; color: white !important;
        padding: 8px 16px; border-radius: 50px; text-decoration: none;
        font-weight: 600; display: inline-block; border: 1px solid #c53929;
    }
</style>
""", unsafe_allow_html=True)

# ======================================================================================
# --- 2. MOTOR DE INGESTIÓN Y LIMPIEZA DE DATOS ---
# ======================================================================================

def normalizar_texto(texto):
    """Elimina acentos y caracteres especiales."""
    if not isinstance(texto, str): return str(texto)
    texto = unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode("utf-8")
    return texto.upper().strip()

def limpiar_moneda(valor):
    """Convierte moneda a float de forma robusta."""
    if pd.isna(valor): return 0.0
    s_val = str(valor).strip()
    s_val = re.sub(r'[^\d.,-]', '', s_val) # Dejar solo números, puntos y comas
    if not s_val: return 0.0

    try:
        # Lógica para detectar si es 1.000,00 (Europa/Latam) o 1,000.00 (USA)
        if ',' in s_val and '.' in s_val:
            if s_val.rfind(',') > s_val.rfind('.'): # Coma al final: decimal
                s_val = s_val.replace('.', '').replace(',', '.')
            else: # Punto al final: decimal
                s_val = s_val.replace(',', '')
        elif ',' in s_val:
            parts = s_val.split(',')
            # Si el último bloque tiene 3 dígitos exactos, asumimos miles (ej: 100,000)
            # A menos que sea muy corto (ej: 50,5)
            if len(parts[-1]) != 3: 
                s_val = s_val.replace(',', '.') 
            else:
                s_val = s_val.replace(',', '') 
        return float(s_val)
    except:
        return 0.0

@st.cache_data(ttl=300)
def cargar_datos_maestros():
    """Carga y normaliza datos."""
    df_raw = pd.DataFrame()
    origen = "Desconocido"

    # 1. Intento Dropbox
    try:
        if "dropbox" in st.secrets:
            dbx = dropbox.Dropbox(
                app_key=st.secrets["dropbox"]["app_key"],
                app_secret=st.secrets["dropbox"]["app_secret"],
                oauth2_refresh_token=st.secrets["dropbox"]["refresh_token"]
            )
            _, res = dbx.files_download(path='/data/cartera_detalle.csv')
            try:
                df_raw = pd.read_csv(io.StringIO(res.content.decode('latin-1')), sep='|', header=None, dtype=str)
            except:
                df_raw = pd.read_csv(io.StringIO(res.content.decode('latin-1')), sep=',', header=0, dtype=str)
            origen = "Nube (Dropbox)"
    except Exception:
        pass 

    # 2. Intento Local
    if df_raw.empty:
        archivos = glob.glob("Cartera_*.xlsx") + glob.glob("Cartera_*.csv")
        if archivos:
            archivo_reciente = max(archivos, key=os.path.getctime)
            try:
                if archivo_reciente.endswith('.csv'):
                    df_raw = pd.read_csv(archivo_reciente, dtype=str, encoding='latin-1')
                else:
                    df_raw = pd.read_excel(archivo_reciente, dtype=str)
                origen = f"Local ({archivo_reciente})"
            except: pass

    if df_raw.empty:
        return pd.DataFrame(), "Sin Datos"

    # --- LIMPIEZA ---
    # Normalizar columnas
    cols_map = {
        'NombreCliente': 'cliente', 'Nit': 'nit', 'NomVendedor': 'vendedor',
        'Importe': 'saldo', 'Dias Vencido': 'dias_mora', 'E-Mail': 'email',
        'Telefono1': 'telefono', 'Numero': 'factura', 'Fecha Vencimiento': 'fecha_venc'
    }
    
    # Estandarizar nombres de columnas del archivo
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    mapping_final = {}
    for k, v in cols_map.items():
        for col_real in df_raw.columns:
            if normalizar_texto(k) in normalizar_texto(col_real):
                mapping_final[col_real] = v
                break
    df_raw.rename(columns=mapping_final, inplace=True)
    
    # Rellenar columnas faltantes
    required = ['cliente', 'saldo', 'dias_mora', 'vendedor', 'nit', 'factura', 'email', 'telefono']
    for req in required:
        if req not in df_raw.columns:
            df_raw[req] = 'N/A'

    # --- CORRECCIÓN DEL ERROR DE TIPOS (AQUÍ ESTÁ EL ARREGLO) ---
    # Convertir columnas de texto explícitamente a String y rellenar nulos
    text_cols = ['cliente', 'vendedor', 'nit', 'email', 'telefono', 'factura']
    for col in text_cols:
        df_raw[col] = df_raw[col].fillna("SIN DEFINIR").astype(str).str.strip().str.upper()

    # Conversión Numérica
    df_raw['saldo'] = df_raw['saldo'].apply(limpiar_moneda)
    df_raw['dias_mora'] = pd.to_numeric(df_raw['dias_mora'], errors='coerce').fillna(0)
    
    # Filtro final
    df_raw = df_raw[df_raw['saldo'] != 0]

    return df_raw, origen

# ======================================================================================
# --- 3. LÓGICA DE NEGOCIO ---
# ======================================================================================

def analizar_cartera(df):
    if df.empty: return df

    bins = [-9999, 0, 30, 60, 90, 9999]
    labels = ['Corriente (Al día)', '1 a 30 Días', '31 a 60 Días', '61 a 90 Días', '> 90 Días (Jurídico)']
    df['rango_mora'] = pd.cut(df['dias_mora'], bins=bins, labels=labels)

    def get_estado(dias):
        if dias <= 0: return "✅ Al Día"
        if dias <= 30: return "🟡 Preventivo"
        if dias <= 60: return "🟠 Administrativo"
        if dias <= 90: return "🔴 Pre-Jurídico"
        return "⚫ Castigo/Jurídico"
    
    df['estado_gestion'] = df['dias_mora'].apply(get_estado)

    def get_accion(row):
        dias = row['dias_mora']
        monto = row['saldo']
        if dias <= 0: return "Agradecer pago / Venta nueva"
        if dias <= 15: return "Recordatorio Whatsapp suave"
        if dias <= 30: return "Llamada de servicio + Email Estado Cuenta"
        if dias <= 60: return "LLAMADA DE COBRO (Compromiso fecha)"
        if dias <= 90: return "BLOQUEO DE CUPO + Carta Prejurídica"
        if monto < 100000: return "Castigar cartera (Bajo monto)"
        return "TRASLADO A ABOGADO"

    df['accion_sugerida'] = df.apply(get_accion, axis=1)
    return df

# ======================================================================================
# --- 4. GENERADOR DE REPORTES ---
# ======================================================================================

def generar_excel_gerencial(df_detalle, df_resumen_cliente):
    output = io.BytesIO()
    wb = Workbook()
    
    ws_kpi = wb.active
    ws_kpi.title = "Resumen Gerencial"
    ws_kpi['A1'] = "INFORME DE ESTADO DE CARTERA"
    
    ws_clientes = wb.create_sheet("Top Gestión")
    headers = ["Cliente", "NIT", "Vendedor", "Saldo Total", "Días Mora Max", "Estado", "Acción Sugerida", "Teléfono", "Email"]
    ws_clientes.append(headers)
    
    for _, row in df_resumen_cliente.iterrows():
        ws_clientes.append([
            row['cliente'], row['nit'], row['vendedor'], 
            row['saldo'], row['dias_mora'], 
            row['estado_gestion'], row['accion_sugerida'],
            row['telefono'], row['email']
        ])
    
    ws_data = wb.create_sheet("Data Cruda")
    rows = [df_detalle.columns.tolist()] + df_detalle.astype(str).values.tolist()
    for r in rows: ws_data.append(r)

    wb.save(output)
    return output.getvalue()

# ======================================================================================
# --- 5. INTERFAZ DE USUARIO ---
# ======================================================================================

def main():
    # --- SIDEBAR ---
    with st.sidebar:
        st.title("Panel de Control")
        user_type = st.selectbox("Perfil", ["Gerencia General", "Líder Cartera", "Vendedor"])
        
        df, source_msg = cargar_datos_maestros()
        st.caption(f"🔌 Fuente: {source_msg}")
        
        if df.empty:
            st.error("⚠️ No hay datos. Carga 'Cartera_Data.xlsx'.")
            st.stop()
            
        df_processed = analizar_cartera(df)
        
        # --- FILTROS SEGUROS ---
        st.markdown("### 🔍 Filtros")
        
        # 1. Asegurar que son strings únicos
        lista_vendedores = sorted(df_processed['vendedor'].unique().astype(str))
        vendedores = ["TODOS"] + lista_vendedores
        sel_vendedor = st.selectbox("Vendedor:", vendedores)
        
        df_view = df_processed.copy()
        if sel_vendedor != "TODOS":
            df_view = df_view[df_view['vendedor'] == sel_vendedor]
            
    # --- DASHBOARD ---
    st.markdown(f"## 📊 Estado de Cartera - Vista {user_type}")
    
    total_cobrar = df_view['saldo'].sum()
    vencido = df_view[df_view['dias_mora'] > 0]['saldo'].sum()
    critico = df_view[df_view['dias_mora'] > 60]['saldo'].sum()
    
    k1, k2, k3 = st.columns(3)
    k1.metric("💰 Cartera Total", f"${total_cobrar:,.0f}")
    k2.metric("🔥 Total Vencido", f"${vencido:,.0f}", delta_color="inverse")
    k3.metric("🚨 Riesgo Crítico", f"${critico:,.0f}", delta_color="inverse")
    
    st.markdown("---")
    
    tab_war_room, tab_analytics, tab_data = st.tabs(["⚔️ GESTIÓN", "📈 ANÁLISIS", "📂 REPORTES"])
    
    with tab_war_room:
        df_clients = df_view.groupby(['cliente', 'nit', 'vendedor', 'telefono', 'email']).agg({
            'saldo': 'sum', 'dias_mora': 'max', 'factura': 'count'
        }).reset_index()
        
        df_clients = analizar_cartera(df_clients)
        df_clients = df_clients.sort_values(by=['rango_mora', 'saldo'], ascending=[False, False])
        
        col_list, col_action = st.columns([1, 2])
        
        with col_list:
            st.markdown("### Objetivos")
            df_clients['label'] = df_clients.apply(lambda x: f"{x['cliente']} | ${x['saldo']:,.0f}", axis=1)
            # Manejo seguro si la lista está vacía
            opciones_clientes = df_clients['label'].tolist() if not df_clients.empty else ["Sin clientes"]
            target_client = st.selectbox("Seleccione Cliente:", opciones_clientes)
            
        with col_action:
            if target_client != "Sin clientes":
                client_name = target_client.split(' | ')[0]
                client_row = df_clients[df_clients['cliente'] == client_name]
                
                if not client_row.empty:
                    client_data = client_row.iloc[0]
                    st.markdown(f"### Gestión: {client_name}")
                    st.info(f"**Estrategia:** {client_data['accion_sugerida']}")
                    
                    facturas_cli = df_view[df_view['cliente'] == client_name][['factura', 'fecha_venc', 'dias_mora', 'saldo']]
                    st.dataframe(facturas_cli.style.format({'saldo': '${:,.0f}'}), use_container_width=True)
                    
                    # Generador de Links
                    phone = re.sub(r'\D', '', str(client_data['telefono']))
                    msg = quote(f"Hola {client_name}, su saldo pendiente es ${client_data['saldo']:,.0f}.")
                    
                    c1, c2 = st.columns(2)
                    if len(phone) >= 10:
                        c1.markdown(f'<a href="https://wa.me/57{phone}?text={msg}" target="_blank" class="action-btn-wa">📱 WhatsApp</a>', unsafe_allow_html=True)
                    else:
                        c1.warning("Sin teléfono válido")
    
    with tab_analytics:
        st.subheader("Radiografía de la Cartera")

        

[Image of Data Dashboard]

        
        g1, g2 = st.columns(2)
        with g1:
            df_pie = df_view.groupby('rango_mora', observed=True)['saldo'].sum().reset_index()
            fig_pie = px.pie(df_pie, values='saldo', names='rango_mora', hole=0.4, title='Por Edades')
            st.plotly_chart(fig_pie, use_container_width=True)
        with g2:
            df_top = df_clients.sort_values('saldo', ascending=False).head(10) if 'df_clients' in locals() else pd.DataFrame()
            if not df_top.empty:
                fig_bar = px.bar(df_top, x='saldo', y='cliente', orientation='h', title='Top 10 Deudores')
                st.plotly_chart(fig_bar, use_container_width=True)

    with tab_data:
        st.success("Descargar Reporte Excel")
        if 'df_clients' in locals():
            excel_file = generar_excel_gerencial(df_view, df_clients)
            st.download_button("📥 DESCARGAR EXCEL", excel_file, f"Cartera_{datetime.now().strftime('%Y%m%d')}.xlsx")

if __name__ == '__main__':
    main()
