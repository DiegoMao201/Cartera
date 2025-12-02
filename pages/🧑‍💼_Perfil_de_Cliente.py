import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import os
import glob
import re
import unicodedata
from datetime import datetime
from urllib.parse import quote
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from fpdf import FPDF
import yagmail # Necesario para enviar correos: pip install yagmail

# --- CONFIGURACIÓN VISUAL PROFESIONAL ---
st.set_page_config(
    page_title="Centro de Mando: Cobranza Ferreinox",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Paleta de Colores y CSS Corporativo
COLOR_PRIMARIO = "#003366"  # Azul oscuro corporativo
COLOR_ACCION = "#FFC300"    # Amarillo para acciones
COLOR_FONDO = "#f4f6f9"
st.markdown(f"""
<style>
    .main {{ background-color: {COLOR_FONDO}; }}
    /* Métricas */
    .stMetric {{ background-color: white; padding: 15px; border-radius: 8px; border-left: 5px solid {COLOR_PRIMARIO}; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
    /* Títulos */
    h1, h2, h3 {{ color: {COLOR_PRIMARIO}; }}
    /* Botones de acción */
    .stButton>button {{ width: 100%; border-radius: 5px; font-weight: bold; }}
    /* Estilo para Link WhatsApp */
    a.wa-link {{
        text-decoration: none; display: inline-block; padding: 6px 12px;
        background-color: #25D366; color: white; border-radius: 4px; font-weight: bold;
    }}
    a.wa-link:hover {{ background-color: #128C7E; }}
</style>
""", unsafe_allow_html=True)

# ======================================================================================
# 1. MOTOR DE CONEXIÓN Y LIMPIEZA DE DATOS
# ======================================================================================

def normalizar_texto(texto):
    if not isinstance(texto, str): return str(texto)
    texto = unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode("utf-8").upper().strip()
    return re.sub(r'[^\w\s\.]', '', texto).strip()

def limpiar_moneda(valor):
    if pd.isna(valor): return 0.0
    s_val = str(valor).strip()
    s_val = re.sub(r'[^\d.,-]', '', s_val)
    if not s_val: return 0.0
    try:
        # Lógica para detectar miles y decimales
        if ',' in s_val and '.' in s_val:
            if s_val.rfind(',') > s_val.rfind('.'): # Formato europeo/latino 1.000,00
                s_val = s_val.replace('.', '').replace(',', '.')
            else: # Formato USA 1,000.00
                s_val = s_val.replace(',', '')
        elif ',' in s_val: # Solo comas
            if len(s_val.split(',')[-1]) == 2: s_val = s_val.replace(',', '.') # Asume decimal
            else: s_val = s_val.replace(',', '') # Asume miles
        return float(s_val)
    except:
        return 0.0

def mapear_y_limpiar_df(df):
    df.columns = [normalizar_texto(c) for c in df.columns]
    
    # Diccionario inteligente de columnas
    mapa = {
        'cliente': ['NOMBRE', 'RAZON SOCIAL', 'TERCERO', 'CLIENTE', 'NOMBRECLIENTE'],
        'nit': ['NIT', 'IDENTIFICACION', 'CEDULA', 'RUT'],
        'saldo': ['IMPORTE', 'SALDO', 'TOTAL', 'DEUDA', 'VALOR'],
        'dias': ['DIAS', 'VENCIDO', 'MORA', 'ANTIGUEDAD', 'DIAS VENCIDO'],
        'telefono': ['TEL', 'MOVIL', 'CELULAR', 'TELEFONO', 'TELEFONO1'],
        'vendedor': ['VENDEDOR', 'ASESOR', 'COMERCIAL', 'NOMVENDEDOR'],
        'factura': ['NUMERO', 'FACTURA', 'DOC', 'SERIE'],
        'email': ['CORREO', 'EMAIL', 'E-MAIL', 'MAIL']
    }
    
    renombres = {}
    for standard, variantes in mapa.items():
        for col in df.columns:
            col_norm = normalizar_texto(col)
            if standard not in renombres.values() and any(v in col_norm for v in variantes):
                renombres[col] = standard
                break
    
    df.rename(columns=renombres, inplace=True)
    
    # Validación mínima
    req = ['cliente', 'saldo', 'dias']
    if not all(c in df.columns for c in req):
        return None, f"Faltan columnas críticas (Cliente, Saldo, Días). Detectadas: {list(df.columns)}"

    # Conversión de tipos
    df['saldo'] = df['saldo'].apply(limpiar_moneda)
    df['dias'] = pd.to_numeric(df['dias'], errors='coerce').fillna(0).astype(int)
    
    for c in ['telefono', 'vendedor', 'nit', 'factura', 'email']:
        if c not in df.columns: df[c] = 'N/A'
        else: df[c] = df[c].fillna('N/A').astype(str)

    # Limpieza final
    df = df[df['saldo'] != 0] # Mantiene saldos negativos (anticipos) y positivos
    return df, "OK"

@st.cache_data(ttl=300)
def cargar_datos_automaticos():
    """Busca archivos automáticamente en la carpeta local."""
    # Buscar archivos Excel o CSV que parezcan de cartera
    archivos = glob.glob("Cartera*.xlsx") + glob.glob("*.csv") + glob.glob("*.xlsx")
    
    if not archivos:
        return None, "No se encontraron archivos locales."
    
    # Priorizar archivos que empiecen por 'Cartera'
    archivo_prioritario = next((f for f in archivos if "Cartera" in f), archivos[0])
    
    try:
        if archivo_prioritario.endswith('.csv'):
            df = pd.read_csv(archivo_prioritario, sep=None, engine='python', encoding='latin-1', dtype=str)
        else:
            df = pd.read_excel(archivo_prioritario, dtype=str)
            
        df_proc, status = mapear_y_limpiar_df(df)
        if df_proc is None: return None, status
        return df_proc, f"Conectado automáticamente a: {os.path.basename(archivo_prioritario)}"
    except Exception as e:
        return None, f"Error leyendo {archivo_prioritario}: {str(e)}"

# ======================================================================================
# 2. INTELIGENCIA DE NEGOCIO (ESTRATEGIA)
# ======================================================================================

def segmentar_cartera(df):
    bins = [-float('inf'), 0, 30, 60, 90, float('inf')]
    labels = ["🟢 Al Día", "🟡 Preventivo (1-30)", "🟠 Riesgo (31-60)", "🔴 Crítico (61-90)", "⚫ Legal (+90)"]
    df['Rango'] = pd.cut(df['dias'], bins=bins, labels=labels)
    return df

def calcular_kpis(df):
    total = df['saldo'].sum()
    vencido = df[df['dias'] > 0]['saldo'].sum()
    pct_vencido = (vencido / total * 100) if total else 0
    clientes_mora = df[df['dias'] > 0]['cliente'].nunique()
    return total, vencido, pct_vencido, clientes_mora

def generar_link_wa(telefono, cliente, saldo, dias, facturas):
    tel = re.sub(r'\D', '', str(telefono))
    if len(tel) == 10: tel = '57' + tel # Asume Colombia
    if len(tel) < 10: return None
    
    cliente_corto = str(cliente).split()[0].title()
    
    if dias <= 0:
        msg = f"Hola {cliente_corto}, de Ferreinox. ¡Gracias por mantener tu cuenta al día! Adjunto tu estado de cuenta."
    elif dias < 30:
        msg = f"Hola {cliente_corto}, recordatorio amable de Ferreinox. Tienes un saldo de ${saldo:,.0f} vencido hace {dias} días. Agradecemos tu pago hoy."
    else:
        msg = f"URGENTE {cliente_corto}: Su cuenta en Ferreinox presenta ${saldo:,.0f} con {dias} días de mora. Requerimos pago inmediato para evitar bloqueos."
        
    return f"https://wa.me/{tel}?text={quote(msg)}"

# ======================================================================================
# 3. GENERADORES (PDF Y EXCEL)
# ======================================================================================

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 15)
        self.set_text_color(0, 51, 102)
        self.cell(0, 10, 'ESTADO DE CUENTA - FERREINOX SAS BIC', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

def crear_pdf(df_cliente):
    pdf = PDF()
    pdf.add_page()
    
    # Datos Cliente
    row = df_cliente.iloc[0]
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(0, 6, f"Cliente: {row['cliente']}", 0, 1)
    pdf.cell(0, 6, f"NIT/ID: {row['nit']}", 0, 1)
    pdf.cell(0, 6, f"Fecha Corte: {datetime.now().strftime('%Y-%m-%d')}", 0, 1)
    pdf.ln(5)
    
    # Tabla
    pdf.set_fill_color(200, 220, 255)
    pdf.cell(30, 8, "Factura", 1, 0, 'C', 1)
    pdf.cell(30, 8, "Días Mora", 1, 0, 'C', 1)
    pdf.cell(40, 8, "Vendedor", 1, 0, 'C', 1)
    pdf.cell(40, 8, "Saldo", 1, 1, 'C', 1)
    
    pdf.set_font("Arial", '', 10)
    total = 0
    for _, item in df_cliente.iterrows():
        total += item['saldo']
        pdf.cell(30, 7, str(item['factura']), 1)
        pdf.cell(30, 7, str(item['dias']), 1, 0, 'C')
        pdf.cell(40, 7, str(item['vendedor'])[:18], 1) # Truncar nombre largo
        pdf.cell(40, 7, f"${item['saldo']:,.0f}", 1, 1, 'R')
        
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(100, 8, "TOTAL A PAGAR", 1, 0, 'R')
    pdf.cell(40, 8, f"${total:,.0f}", 1, 1, 'R')
    
    return bytes(pdf.output())

def crear_excel_gerencial(df, kpis):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen Gerencial"
    
    # Estilos
    header_style = Font(bold=True, color="FFFFFF")
    fill_blue = PatternFill("solid", fgColor="003366")
    
    # KPIs en Excel
    ws['A1'] = "REPORTE GERENCIAL DE CARTERA"
    ws['A1'].font = Font(size=14, bold=True)
    
    kpi_labels = ["Total Cartera", "Total Vencido", "% Mora", "Clientes en Mora"]
    kpi_values = [kpis[0], kpis[1], kpis[2]/100, kpis[3]]
    formats = ['$#,##0', '$#,##0', '0.0%', '0']
    
    for i, (lab, val, fmt) in enumerate(zip(kpi_labels, kpi_values, formats)):
        ws.cell(row=3, column=i+1, value=lab).font = Font(bold=True)
        c = ws.cell(row=4, column=i+1, value=val)
        c.number_format = fmt
        
    # Tabla Detalle
    ws['A6'] = "DETALLE DE CLIENTES (Priorizado por Deuda)"
    cols = ['cliente', 'nit', 'factura', 'vendedor', 'dias', 'Rango', 'saldo', 'telefono', 'email']
    
    # Headers
    for col_num, col_name in enumerate(cols, 1):
        c = ws.cell(row=7, column=col_num, value=col_name.upper())
        c.fill = fill_blue
        c.font = header_style
        
    # Data
    for row_num, row_data in enumerate(df[cols].values, 8):
        for col_num, val in enumerate(row_data, 1):
            c = ws.cell(row=row_num, column=col_num, value=val)
            if col_num == 7: c.number_format = '$#,##0' # Saldo
            
    # Filtros
    ws.auto_filter.ref = f"A7:I{len(df)+7}"
    
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ======================================================================================
# 4. FUNCIÓN DE ENVÍO DE CORREO
# ======================================================================================
def enviar_correo(destinatario, asunto, cuerpo, pdf_bytes, nombre_pdf):
    try:
        # Credenciales: Intenta leer de secrets, si no, usa inputs temporales
        email_user = st.session_state.get('email_user', '')
        email_pass = st.session_state.get('email_pass', '')
        
        if not email_user or not email_pass:
            st.error("⚠️ Configura las credenciales de correo en la barra lateral.")
            return False

        yag = yagmail.SMTP(email_user, email_pass)
        
        # Guardar PDF temporalmente
        path_pdf = f"temp_{nombre_pdf}"
        with open(path_pdf, "wb") as f:
            f.write(pdf_bytes)
            
        yag.send(
            to=destinatario,
            subject=asunto,
            contents=[cuerpo, path_pdf] # Adjunta el PDF
        )
        
        os.remove(path_pdf) # Limpiar
        return True
    except Exception as e:
        st.error(f"Error enviando correo: {e}")
        return False

# ======================================================================================
# 5. DASHBOARD PRINCIPAL
# ======================================================================================

def main():
    # --- BARRA LATERAL: CONFIGURACIÓN Y FILTROS ---
    with st.sidebar:
        st.image("https://cdn-icons-png.flaticon.com/512/9322/9322127.png", width=50)
        st.header("⚙️ Configuración")
        
        # Credenciales de Correo (Sesión)
        with st.expander("📧 Configurar Correo (Gmail/Outlook)"):
            st.session_state['email_user'] = st.text_input("Tu Correo", value=st.session_state.get('email_user', ''))
            st.session_state['email_pass'] = st.text_input("Tu Contraseña de Aplicación", type="password", value=st.session_state.get('email_pass', ''))
            st.caption("Nota: Para Gmail usa 'Contraseña de Aplicación'.")

        st.divider()
        st.header("🔍 Filtros Operativos")
        
    # --- CARGA DE DATOS ---
    df, status = cargar_datos_automaticos()
    
    # Opción manual si falla la automática
    if df is None:
        st.warning(f"{status} - Por favor sube el archivo manualmente:")
        uploaded = st.file_uploader("Subir Excel/CSV", type=['xlsx', 'csv'])
        if uploaded:
            if uploaded.name.endswith('.csv'):
                df_raw = pd.read_csv(uploaded, sep=None, engine='python', encoding='latin-1', dtype=str)
            else:
                df_raw = pd.read_excel(uploaded, dtype=str)
            df, status = mapear_y_limpiar_df(df_raw)
    
    if df is None:
        st.stop()

    # --- PROCESAMIENTO ---
    df = segmentar_cartera(df)
    
    # Filtros Dinámicos
    vendedores = ["TODOS"] + sorted(df['vendedor'].unique().tolist())
    filtro_vendedor = st.sidebar.selectbox("Filtrar por Vendedor", vendedores)
    
    if filtro_vendedor != "TODOS":
        df_view = df[df['vendedor'] == filtro_vendedor].copy()
    else:
        df_view = df.copy()

    total, vencido, pct_mora, clientes_mora = calcular_kpis(df_view)

    # --- ENCABEZADO Y KPIS ---
    st.title(f"Centro de Cobranza: {status}")
    
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("💰 Cartera Total", f"${total:,.0f}")
    k2.metric("⚠️ Cartera Vencida", f"${vencido:,.0f}", f"{pct_mora:.1f}% del total")
    k3.metric("👥 Clientes en Mora", clientes_mora)
    k4.metric("📅 Fecha Corte", datetime.now().strftime("%d-%m-%Y"))
    
    st.divider()

    # --- TABS DE GESTIÓN ---
    tab_lider, tab_gerente, tab_datos = st.tabs(["👩‍💼 LÍDER (Gestión)", "👨‍💼 GERENTE (Visión)", "📥 EXPORTAR"])

    # ==============================================================================
    # TAB LÍDER: GESTIÓN DE COBRO 1 A 1
    # ==============================================================================
    with tab_lider:
        st.subheader("🎯 Gestión Directa de Clientes")
        
        # Agrupar por Cliente para gestión
        df_agrupado = df_view[df_view['dias'] > 0].groupby('cliente').agg({
            'saldo': 'sum',
            'dias': 'max',
            'factura': 'count', # Número de facturas
            'telefono': 'first',
            'email': 'first',
            'vendedor': 'first',
            'nit': 'first'
        }).reset_index().sort_values('saldo', ascending=False)
        
        # Selector de cliente
        cliente_sel = st.selectbox("🔍 Selecciona Cliente a Gestionar (Ordenado por Deuda)", df_agrupado['cliente'])
        
        if cliente_sel:
            data_cli = df_agrupado[df_agrupado['cliente'] == cliente_sel].iloc[0]
            detalle_facturas = df_view[df_view['cliente'] == cliente_sel]
            
            c1, c2 = st.columns([1, 2])
            
            with c1:
                st.info(f"**Deuda Total:** ${data_cli['saldo']:,.0f}")
                st.warning(f"**Días Máx Mora:** {data_cli['dias']} días")
                st.text(f"📞 {data_cli['telefono']}")
                st.text(f"📧 {data_cli['email']}")
                
                # Generar PDF en memoria
                pdf_bytes = crear_pdf(detalle_facturas)
                
                # --- BOTÓN WHATSAPP ---
                link_wa = generar_link_wa(data_cli['telefono'], cliente_sel, data_cli['saldo'], data_cli['dias'], data_cli['factura'])
                if link_wa:
                    st.markdown(f"""<a href="{link_wa}" target="_blank" class="wa-link">📱 ABRIR WHATSAPP CON GUION</a>""", unsafe_allow_html=True)
                else:
                    st.error("Número de teléfono inválido para WhatsApp")

            with c2:
                st.write("#### 📄 Estado de Cuenta y Correo")
                # Vista previa rápida de facturas
                st.dataframe(detalle_facturas[['factura', 'fecha_vencimiento', 'dias', 'saldo'] if 'fecha_vencimiento' in detalle_facturas else ['factura', 'dias', 'saldo']], height=150)
                
                # --- ENVÍO DE CORREO ---
                with st.form("form_email"):
                    email_dest = st.text_input("Destinatario", value=data_cli['email'])
                    asunto_msg = f"Estado de Cuenta - {cliente_sel} - {datetime.now().strftime('%d/%m')}"
                    submit_email = st.form_submit_button("📧 ENVIAR CORREO AHORA")
                    
                    if submit_email:
                        if enviar_correo(email_dest, asunto_msg, "Adjunto encontrará su estado de cuenta detallado.", pdf_bytes, "EstadoCuenta.pdf"):
                            st.success(f"✅ Correo enviado a {email_dest}")
                        else:
                            st.error("❌ Falló el envío. Revisa credenciales.")

    # ==============================================================================
    # TAB GERENTE: VISIÓN ESTRATÉGICA
    # ==============================================================================
    with tab_gerente:
        st.subheader("📊 Análisis de Cartera")
        
        g1, g2 = st.columns(2)
        
        with g1:
            st.markdown("**1. Distribución por Riesgo**")
            df_pie = df_view.groupby('Rango')['saldo'].sum().reset_index()
            fig_pie = px.pie(df_pie, names='Rango', values='saldo', color='Rango', 
                             color_discrete_map={"🟢 Al Día": "green", "🟡 Preventivo (1-30)": "yellow", "🟠 Riesgo (31-60)": "orange", "🔴 Crítico (61-90)": "red", "⚫ Legal (+90)": "black"})
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with g2:
            st.markdown("**2. Top 10 Clientes Morosos (Pareto)**")
            top_cli = df_view[df_view['dias']>0].groupby('cliente')['saldo'].sum().nlargest(10).reset_index()
            fig_bar = px.bar(top_cli, x='saldo', y='cliente', orientation='h', text_auto='.2s', title="Deuda Vencida")
            fig_bar.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_bar, use_container_width=True)
            
        st.markdown("**3. Desempeño por Vendedor**")
        resumen_vendedor = df_view.groupby('vendedor').agg(
            Cartera_Total=('saldo', 'sum'),
            Vencido=('saldo', lambda x: x[df_view.loc[x.index, 'dias'] > 0].sum())
        ).reset_index()
        resumen_vendedor['% Vencido'] = (resumen_vendedor['Vencido'] / resumen_vendedor['Cartera_Total'] * 100)
        st.dataframe(resumen_vendedor.style.format({'Cartera_Total': '${:,.0f}', 'Vencido': '${:,.0f}', '% Vencido': '{:.1f}%'}).background_gradient(subset=['% Vencido'], cmap='RdYlGn_r'), use_container_width=True)

    # ==============================================================================
    # TAB DATOS: EXPORTAR EXCEL
    # ==============================================================================
    with tab_datos:
        st.subheader("📥 Descargas")
        
        excel_data = crear_excel_gerencial(df_view, [total, vencido, pct_mora, clientes_mora])
        
        st.download_button(
            label="💾 DESCARGAR REPORTE GERENCIAL (EXCEL)",
            data=excel_data,
            file_name=f"Reporte_Cartera_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.subheader("🔎 Datos Crudos")
        st.dataframe(df_view)

if __name__ == "__main__":
    main()
