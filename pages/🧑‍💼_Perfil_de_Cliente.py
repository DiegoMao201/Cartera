# ======================================================================================
# SISTEMA INTEGRAL DE GESTIÓN DE CARTERA Y COBRANZA (V. DEFINITIVA - GERENCIAL)
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

# Intentar importar yagmail, si no está, no romper la app
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

# Estilos CSS Profesionales
st.markdown("""
<style>
    /* Fondo y Tipografía */
    .stApp { background-color: #f8f9fa; font-family: 'Segoe UI', sans-serif; }
    
    /* Métricas Superiores */
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border-left: 5px solid #0d6efd;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }
    
    /* Títulos */
    h1, h2, h3 { color: #003366; }
    
    /* Tablas */
    .dataframe { font-size: 14px; }
    
    /* Botones de Acción */
    .action-btn-wa {
        background-color: #25D366; color: white !important;
        padding: 8px 16px; border-radius: 50px; text-decoration: none;
        font-weight: 600; display: inline-block; border: 1px solid #1da851;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .action-btn-email {
        background-color: #EA4335; color: white !important;
        padding: 8px 16px; border-radius: 50px; text-decoration: none;
        font-weight: 600; display: inline-block; border: 1px solid #c53929;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
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
    """
    Convierte moneda ($ 1.000,00 o 1,000.00) a float puro.
    Versión robusta para formatos latinos y americanos.
    """
    if pd.isna(valor): return 0.0
    s_val = str(valor).strip()
    
    # 1. Eliminar símbolos de moneda y espacios
    s_val = re.sub(r'[^\d.,-]', '', s_val)
    
    if not s_val: return 0.0

    try:
        # Detectar formato:
        # Si tiene coma y punto, decidir cuál es decimal
        if ',' in s_val and '.' in s_val:
            if s_val.rfind(',') > s_val.rfind('.'): # Ejemplo: 1.500,50 (Latino)
                s_val = s_val.replace('.', '').replace(',', '.')
            else: # Ejemplo: 1,500.50 (Inglés)
                s_val = s_val.replace(',', '')
        elif ',' in s_val:
            # Si solo tiene comas:
            # Si la coma separa 1 o 2 dígitos al final, es decimal (ej: 50,5)
            # Si separa 3, asumimos miles, a menos que sea un número pequeño con decimales largos
            parts = s_val.split(',')
            if len(parts[-1]) != 3: 
                s_val = s_val.replace(',', '.') # Es decimal
            else:
                s_val = s_val.replace(',', '') # Es miles
        
        return float(s_val)
    except:
        return 0.0

@st.cache_data(ttl=300)
def cargar_datos_maestros():
    """Carga unificada: Intenta Dropbox primero, luego Local."""
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
            # Ajustar ruta según tu Dropbox real
            _, res = dbx.files_download(path='/data/cartera_detalle.csv')
            # Intentar leer csv con diferentes separadores si falla
            try:
                df_raw = pd.read_csv(io.StringIO(res.content.decode('latin-1')), sep='|', header=None, dtype=str)
            except:
                df_raw = pd.read_csv(io.StringIO(res.content.decode('latin-1')), sep=',', header=0, dtype=str)
            
            origen = "Nube (Dropbox)"
            
            # Asignar nombres si viene sin header (ajustar según tu CSV real)
            nombres_cols_manual = [
                'Serie','Numero','Fecha Documento','Fecha Vencimiento','Cod Cliente',
                'NombreCliente','Nit','Poblacion','Provincia','Telefono1','Telefono2',
                'NomVendedor','Entidad Autoriza','E-Mail','Importe','Descuento',
                'Cupo Aprobado','Dias Vencido'
            ]
            if len(df_raw.columns) == len(nombres_cols_manual):
                df_raw.columns = nombres_cols_manual
    except Exception as e:
        pass # Falló Dropbox, seguimos a local

    # 2. Intento Local (Fallback)
    if df_raw.empty:
        # Busca archivos Excel o CSV que empiecen por Cartera_
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

    # --- LIMPIEZA Y ESTANDARIZACIÓN ---
    cols_map = {
        'NombreCliente': 'cliente', 'Nit': 'nit', 'NomVendedor': 'vendedor',
        'Importe': 'saldo', 'Dias Vencido': 'dias_mora', 'E-Mail': 'email',
        'Telefono1': 'telefono', 'Numero': 'factura', 'Fecha Vencimiento': 'fecha_venc'
    }
    
    # Normalizar columnas actuales
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    
    # Mapear a nombres estándar
    mapping_final = {}
    for k, v in cols_map.items():
        for col_real in df_raw.columns:
            # Búsqueda flexible (insensible a mayúsculas/acentos)
            if normalizar_texto(k) in normalizar_texto(col_real):
                mapping_final[col_real] = v
                break
    
    df_raw.rename(columns=mapping_final, inplace=True)
    
    # Asegurar columnas críticas
    required = ['cliente', 'saldo', 'dias_mora', 'vendedor', 'nit', 'factura', 'email', 'telefono']
    for req in required:
        if req not in df_raw.columns:
            df_raw[req] = 0 if req in ['saldo', 'dias_mora'] else 'N/A'

    # Conversión de Tipos
    df_raw['saldo'] = df_raw['saldo'].apply(limpiar_moneda)
    df_raw['dias_mora'] = pd.to_numeric(df_raw['dias_mora'], errors='coerce').fillna(0)
    
    # Filtros de limpieza
    df_raw = df_raw[df_raw['cliente'].astype(str).str.len() > 2]
    df_raw = df_raw[df_raw['saldo'] != 0]

    return df_raw, origen

# ======================================================================================
# --- 3. LÓGICA DE NEGOCIO ---
# ======================================================================================

def analizar_cartera(df):
    """Segmentación y Estrategia."""
    if df.empty: return df

    # 1. Rangos de Edad
    bins = [-9999, 0, 30, 60, 90, 9999]
    labels = ['Corriente (Al día)', '1 a 30 Días', '31 a 60 Días', '61 a 90 Días', '> 90 Días (Jurídico)']
    df['rango_mora'] = pd.cut(df['dias_mora'], bins=bins, labels=labels)

    # 2. Estado Visual
    def get_estado(dias):
        if dias <= 0: return "✅ Al Día"
        if dias <= 30: return "🟡 Preventivo"
        if dias <= 60: return "🟠 Administrativo"
        if dias <= 90: return "🔴 Pre-Jurídico"
        return "⚫ Castigo/Jurídico"
    
    df['estado_gestion'] = df['dias_mora'].apply(get_estado)

    # 3. Acción Recomendada
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
# --- 4. GENERADOR DE REPORTES GERENCIALES ---
# ======================================================================================

def generar_excel_gerencial(df_detalle, df_resumen_cliente):
    output = io.BytesIO()
    wb = Workbook()
    
    # --- HOJA 1: RESUMEN EJECUTIVO (KPIs) ---
    ws_kpi = wb.active
    ws_kpi.title = "Resumen Gerencial"
    
    ws_kpi['A1'] = "INFORME DE ESTADO DE CARTERA - FERREINOX"
    ws_kpi['A1'].font = Font(size=16, bold=True, color="003366")
    ws_kpi['A2'] = f"Generado el: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    
    total_cartera = df_detalle['saldo'].sum()
    total_vencido = df_detalle[df_detalle['dias_mora'] > 0]['saldo'].sum()
    pct_vencido = (total_vencido / total_cartera) if total_cartera else 0
    
    ws_kpi['A4'] = "Cartera Total"; ws_kpi['B4'] = total_cartera
    ws_kpi['A5'] = "Cartera Vencida"; ws_kpi['B5'] = total_vencido
    ws_kpi['A6'] = "% Deterioro"; ws_kpi['B6'] = pct_vencido
    
    for cell in ['B4', 'B5']: ws_kpi[cell].number_format = '$ #,##0'
    ws_kpi['B6'].number_format = '0.0%'
    
    # --- HOJA 2: TOP CLIENTES ---
    ws_clientes = wb.create_sheet("Top Gestión (Pareto)")
    headers = ["Cliente", "NIT", "Vendedor", "Saldo Total", "Días Mora Max", "Estado", "Acción Sugerida", "Teléfono", "Email"]
    ws_clientes.append(headers)
    
    header_fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    
    for col_num, header in enumerate(headers, 1):
        cell = ws_clientes.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        ws_clientes.column_dimensions[get_column_letter(col_num)].width = 20
        
    df_sorted = df_resumen_cliente.sort_values(by='saldo', ascending=False)
    
    for _, row in df_sorted.iterrows():
        ws_clientes.append([
            row['cliente'], row['nit'], row['vendedor'], 
            row['saldo'], row['dias_mora'], 
            row['estado_gestion'], row['accion_sugerida'],
            row['telefono'], row['email']
        ])
    
    tab = Table(displayName="TablaClientes", ref=f"A1:I{len(df_sorted)+1}")
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    tab.tableStyleInfo = style
    ws_clientes.add_table(tab)
    
    for row in range(2, len(df_sorted)+2):
        ws_clientes[f'D{row}'].number_format = '$ #,##0'
        
    # --- HOJA 3: DETALLE FACTURAS ---
    ws_data = wb.create_sheet("Data Cruda")
    # Convertir a lista de listas para openpyxl
    rows = [df_detalle.columns.tolist()] + df_detalle.astype(str).values.tolist()
    for r in rows: ws_data.append(r)

    wb.save(output)
    return output.getvalue()

# ======================================================================================
# --- 5. INTERFAZ DE USUARIO (DASHBOARD) ---
# ======================================================================================

def main():
    # --- SIDEBAR ---
    with st.sidebar:
        st.title("Panel de Control")
        user_type = st.selectbox("Perfil", ["Gerencia General", "Líder Cartera", "Vendedor"])
        
        df, source_msg = cargar_datos_maestros()
        st.caption(f"🔌 Fuente: {source_msg}")
        
        if df.empty:
            st.error("⚠️ No se encontraron datos.")
            st.info("Por favor, sube un archivo llamado 'Cartera_Data.xlsx' a la carpeta del proyecto o configura Dropbox.")
            st.stop()
            
        df_processed = analizar_cartera(df)
        
        # Filtros
        st.markdown("### 🔍 Filtros")
        vendedores = ["TODOS"] + sorted(list(df_processed['vendedor'].unique()))
        sel_vendedor = st.selectbox("Vendedor:", vendedores)
        
        df_view = df_processed.copy()
        if sel_vendedor != "TODOS":
            df_view = df_view[df_view['vendedor'] == sel_vendedor]
            
    # --- ÁREA PRINCIPAL ---
    st.markdown(f"## 📊 Estado de Cartera - Vista {user_type}")
    
    total_cobrar = df_view['saldo'].sum()
    vencido = df_view[df_view['dias_mora'] > 0]['saldo'].sum()
    critico = df_view[df_view['dias_mora'] > 60]['saldo'].sum()
    potencial = df_view[(df_view['dias_mora'] > 0) & (df_view['dias_mora'] <= 30)]['saldo'].sum()

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("💰 Cartera Total", f"${total_cobrar:,.0f}")
    k2.metric("🔥 Total Vencido", f"${vencido:,.0f}", f"{(vencido/total_cobrar)*100:.1f}%", delta_color="inverse")
    k3.metric("🚨 Riesgo Crítico", f"${critico:,.0f}", "> 60 días", delta_color="inverse")
    k4.metric("🎯 Recaudo Inmediato", f"${potencial:,.0f}", "< 30 días", delta_color="normal")
    
    st.markdown("---")
    
    tab_war_room, tab_analytics, tab_data = st.tabs(["⚔️ GESTIÓN", "📈 ANÁLISIS", "📂 REPORTES"])
    
    # --- TAB 1: GESTIÓN ---
    with tab_war_room:
        df_clients = df_view.groupby(['cliente', 'nit', 'vendedor', 'telefono', 'email']).agg({
            'saldo': 'sum',
            'dias_mora': 'max',
            'factura': 'count'
        }).reset_index()
        
        df_clients = analizar_cartera(df_clients)
        df_clients = df_clients.sort_values(by=['rango_mora', 'saldo'], ascending=[False, False])
        
        col_list, col_action = st.columns([1, 2])
        
        with col_list:
            st.markdown("### Objetivos")
            df_clients['label'] = df_clients.apply(lambda x: f"{x['cliente']} | ${x['saldo']:,.0f}", axis=1)
            target_client = st.selectbox("Seleccione Cliente:", df_clients['label'])
            
            client_name = target_client.split(' | ')[0]
            client_data = df_clients[df_clients['cliente'] == client_name].iloc[0]
            
            st.info(f"**Estrategia:** {client_data['accion_sugerida']}")
            
        with col_action:
            st.markdown(f"### Gestión: {client_name}")
            facturas_cli = df_view[df_view['cliente'] == client_name][['factura', 'fecha_venc', 'dias_mora', 'saldo']]
            st.dataframe(facturas_cli.style.format({'saldo': '${:,.0f}'}), use_container_width=True)
            
            # Botones de contacto
            phone = str(client_data['telefono']).replace('.0', '')
            phone_clean = re.sub(r'\D', '', phone)
            msg_wa = quote(f"Hola {client_name}, recordamos su saldo pendiente de ${client_data['saldo']:,.0f}.")
            
            c1, c2 = st.columns(2)
            if len(phone_clean) >= 10:
                c1.markdown(f'<a href="https://wa.me/57{phone_clean}?text={msg_wa}" target="_blank" class="action-btn-wa">📱 WhatsApp</a>', unsafe_allow_html=True)
            else:
                c1.warning("Sin teléfono válido")
                
            c2.markdown(f'<a href="mailto:{client_data["email"]}?subject=Estado de Cuenta&body={msg_wa}" class="action-btn-email">📧 Correo</a>', unsafe_allow_html=True)

    # --- TAB 2: ANÁLISIS ---
    with tab_analytics:
        st.subheader("Radiografía de la Cartera")
        
        # Se eliminaron los tags de imagen que causaban error de sintaxis
        
        g1, g2 = st.columns(2)
        with g1:
            df_pie = df_view.groupby('rango_mora', observed=True)['saldo'].sum().reset_index()
            fig_pie = px.pie(df_pie, values='saldo', names='rango_mora', hole=0.4, title='Por Edades')
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with g2:
            df_top = df_clients.sort_values('saldo', ascending=False).head(10)
            fig_bar = px.bar(df_top, x='saldo', y='cliente', orientation='h', title='Top 10 Deudores')
            fig_bar.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_bar, use_container_width=True)

    # --- TAB 3: REPORTES ---
    with tab_data:
        st.success("Descargue el reporte procesado para Excel.")
        excel_file = generar_excel_gerencial(df_view, df_clients)
        st.download_button(
            label="📥 DESCARGAR EXCEL GERENCIAL",
            data=excel_file,
            file_name=f"Reporte_Cartera_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.dataframe(df_view)

if __name__ == '__main__':
    main()
