# ======================================================================================
# ARCHIVO: pages/1_🚀_Estrategia_Cobranza.py
# VERSIÓN: WAR ROOM "SUPER GESTIÓN" (Correos Masivos, WhatsApp, Sin Errores)
# ======================================================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO, StringIO
import dropbox
import glob
import unicodedata
import re
from datetime import datetime, timedelta
from urllib.parse import quote
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import yagmail # Necesario para enviar correos reales

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="War Room Cobranza", page_icon="🚀", layout="wide")

# --- ESTILOS CSS PROFESIONALES ---
st.markdown("""
<style>
    .stApp { background-color: #F4F6F9; }
    .stMetric { background-color: #FFFFFF; border-radius: 8px; padding: 15px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-left: 5px solid #0058A7; }
    .big-font { font-size: 20px !important; font-weight: bold; color: #1F2937; }
    .management-card { background-color: #FFFFFF; padding: 25px; border-radius: 12px; border: 1px solid #E5E7EB; box-shadow: 0 4px 10px rgba(0,0,0,0.05); margin-bottom: 20px; }
    .whatsapp-btn { 
        background-color: #25D366; color: white !important; padding: 10px 20px; 
        border-radius: 8px; text-decoration: none; font-weight: bold; 
        display: block; text-align: center; margin-top: 10px;
    }
    .email-btn { 
        background-color: #EA4335; color: white !important; padding: 10px 20px; 
        border-radius: 8px; text-decoration: none; font-weight: bold; 
        display: block; text-align: center; margin-top: 10px;
    }
    .status-badge { padding: 5px 10px; border-radius: 15px; font-weight: bold; font-size: 12px; color: white;}
</style>
""", unsafe_allow_html=True)

# ======================================================================================
# --- 1. CARGA DE DATOS BLINDADA (Cero Errores de Columnas) ---
# ======================================================================================

def normalizar_texto(texto: str) -> str:
    if not isinstance(texto, str): return ""
    texto = texto.upper().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')

def limpiar_columnas(df):
    """Normaliza los nombres de las columnas para evitar KeyError."""
    # 1. Quitar espacios, guiones, puntos y pasar a minúsculas
    df.columns = [c.strip().lower().replace('-', '_').replace('.', '').replace(' ', '_') for c in df.columns]
    
    # 2. Mapeo de sinónimos a nombres estándar
    mapa = {
        'email': 'e_mail',
        'correo': 'e_mail',
        'correo_electronico': 'e_mail',
        'mail': 'e_mail',
        'telefono': 'telefono1',
        'celular': 'telefono1',
        'nombre_cliente': 'nombrecliente',
        'cliente': 'nombrecliente',
        'vendedor': 'nomvendedor',
        'nombre_vendedor': 'nomvendedor',
        'dias_mora': 'dias_vencido',
        'dias': 'dias_vencido',
        'valor': 'importe',
        'saldo': 'importe',
        'total': 'importe'
    }
    df = df.rename(columns=mapa)
    
    # 3. Garantizar columnas críticas (si no existen, crearlas vacías)
    cols_criticas = ['e_mail', 'telefono1', 'nomvendedor', 'nombrecliente', 'nit', 'importe', 'dias_vencido']
    for col in cols_criticas:
        if col not in df.columns:
            df[col] = "No registrado" if col not in ['importe', 'dias_vencido'] else 0
            
    return df

@st.cache_data(ttl=600)
def cargar_datos_maestros():
    df_final = pd.DataFrame()
    
    # --- INTENTO 1: DROPBOX ---
    try:
        if "dropbox" in st.secrets:
            APP_KEY = st.secrets["dropbox"]["app_key"]
            APP_SECRET = st.secrets["dropbox"]["app_secret"]
            REFRESH_TOKEN = st.secrets["dropbox"]["refresh_token"]
            with dropbox.Dropbox(app_key=APP_KEY, app_secret=APP_SECRET, oauth2_refresh_token=REFRESH_TOKEN) as dbx:
                _, res = dbx.files_download(path='/data/cartera_detalle.csv')
                # Cargar sin nombres primero para ver qué llega, o usar nombres forzados
                nombres_cols = ['Serie', 'Numero', 'Fecha Documento', 'Fecha Vencimiento', 'Cod Cliente',
                                'NombreCliente', 'Nit', 'Poblacion', 'Provincia', 'Telefono1', 'Telefono2',
                                'NomVendedor', 'Entidad Autoriza', 'E-Mail', 'Importe', 'Descuento',
                                'Cupo Aprobado', 'Dias Vencido']
                df_dropbox = pd.read_csv(StringIO(res.content.decode('latin-1')), header=None, names=nombres_cols, sep='|', engine='python')
                df_final = pd.concat([df_final, df_dropbox])
    except Exception: pass 

    # --- INTENTO 2: ARCHIVOS LOCALES ---
    archivos = glob.glob("Cartera_*.xlsx")
    for archivo in archivos:
        try:
            df_hist = pd.read_excel(archivo)
            if not df_hist.empty:
                if "Total" in str(df_hist.iloc[-1, 0]): df_hist = df_hist.iloc[:-1]
                df_final = pd.concat([df_final, df_hist])
        except Exception: pass

    if df_final.empty: return pd.DataFrame()

    # --- LIMPIEZA PROFUNDA ---
    df_final = limpiar_columnas(df_final)
    
    # Convertir numéricos
    df_final['importe'] = pd.to_numeric(df_final['importe'], errors='coerce').fillna(0)
    df_final['dias_vencido'] = pd.to_numeric(df_final['dias_vencido'], errors='coerce').fillna(0)
    
    # Filtrar notas crédito/basura
    if 'serie' in df_final.columns:
        df_final = df_final[~df_final['serie'].astype(str).str.contains('W|X', case=False, na=False)]
    
    # Eliminar duplicados de columnas
    df_final = df_final.loc[:, ~df_final.columns.duplicated()]
    
    return df_final

# ======================================================================================
# --- 2. CEREBRO DE PRIORIZACIÓN ---
# ======================================================================================

def procesar_cartera(df):
    if df.empty: return pd.DataFrame()
    df = df[df['importe'] > 0].copy() # Solo lo que deben
    
    # Agrupar por Cliente
    cols_group = ['nombrecliente', 'nit', 'nomvendedor', 'telefono1', 'e_mail']
    # Asegurar que existen en el DF actual
    cols_validas = [c for c in cols_group if c in df.columns]
    
    kpis = df.groupby(cols_validas).agg({
        'importe': 'sum',
        'dias_vencido': 'max',
        'numero': 'count'
    }).reset_index()
    
    # Definir Estrategia y Prioridad
    def estrategia(dias):
        if dias > 90: return "🔴 JURÍDICO"
        if dias > 60: return "⛔ PRE-JURÍDICO"
        if dias > 30: return "🟠 COBRO ACTIVO"
        if dias > 0: return "🟡 PREVENTIVO"
        return "🟢 AL DÍA"

    kpis['Estado'] = kpis['dias_vencido'].apply(estrategia)
    
    # Score: 60% Días Mora + 40% Monto Deuda
    kpis['Prioridad'] = (
        (kpis['dias_vencido'].clip(upper=120) / 1.2) * 0.6 + 
        (kpis['importe'].clip(upper=20000000) / 20000000 * 100) * 0.4
    )
    
    return kpis.sort_values('Prioridad', ascending=False).reset_index(drop=True)

# ======================================================================================
# --- 3. EXCEL ---
# ======================================================================================
def generar_excel(df):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Matriz de Gestión"
    
    headers = ["Prioridad", "Cliente", "NIT", "Teléfono", "Email", "Deuda Total", "Días Mora", "Estado", "Vendedor"]
    ws.append(headers)
    
    # Estilo Header
    fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
    font = Font(color="FFFFFF", bold=True)
    for c in range(1, len(headers)+1):
        cell = ws.cell(row=1, column=c)
        cell.fill = fill
        cell.font = font
        
    for _, row in df.iterrows():
        ws.append([
            f"{row['Prioridad']:.1f}",
            row['nombrecliente'],
            row['nit'],
            row['telefono1'],
            row['e_mail'],
            row['importe'],
            row['dias_vencido'],
            row['Estado'],
            row['nomvendedor']
        ])
    
    # Tabla
    tab = Table(displayName="DatosCartera", ref=f"A1:I{len(df)+1}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    
    wb.save(output)
    return output.getvalue()

# ======================================================================================
# --- 4. INTERFAZ PRINCIPAL ---
# ======================================================================================

def main():
    st.title("🚀 Centro de Comando de Cobranza (War Room)")
    
    # 1. Cargar
    df_raw = cargar_datos_maestros()
    if df_raw.empty:
        st.error("No hay datos. Cargue archivos en Dropbox o en la carpeta local.")
        st.stop()
        
    # 2. Procesar
    df_gestion = procesar_cartera(df_raw)
    
    # 3. Sidebar Filtros
    st.sidebar.header("🔍 Filtros de Gestión")
    vendedores = ["TODOS"] + sorted(df_gestion['nomvendedor'].unique().tolist())
    filtro_vend = st.sidebar.selectbox("Vendedor:", vendedores)
    
    estados = ["TODOS"] + sorted(df_gestion['Estado'].unique().tolist())
    filtro_estado = st.sidebar.selectbox("Estado:", estados)
    
    # Aplicar Filtros
    df_view = df_gestion.copy()
    if filtro_vend != "TODOS":
        df_view = df_view[df_view['nomvendedor'] == filtro_vend]
    if filtro_estado != "TODOS":
        df_view = df_view[df_view['Estado'] == filtro_estado]
        
    # 4. KPIs Globales
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Clientes a Gestionar", f"{len(df_view)}")
    k2.metric("Total Cartera Filtro", f"${df_view['importe'].sum():,.0f}")
    k3.metric("Ticket Promedio", f"${df_view['importe'].mean():,.0f}")
    criticos = len(df_view[df_view['dias_vencido'] > 60])
    k4.metric("🚨 Casos Críticos (>60 días)", f"{criticos}")
    
    st.markdown("---")

    # ==================================================================================
    # ZONA DE SUPER GESTIÓN (Pestañas)
    # ==================================================================================
    tab_gestion, tab_analisis, tab_admin = st.tabs(["⚡ GESTIÓN UNO A UNO", "📊 VISIÓN GENERAL", "📥 EXPORTAR"])

    # --- PESTAÑA 1: GESTIÓN UNO A UNO (Poderosa) ---
    with tab_gestion:
        st.info("💡 Seleccione un cliente de la lista para activar las herramientas de cobro (WhatsApp y Email).")
        
        # Selector de Cliente con Búsqueda
        lista_clientes = df_view['nombrecliente'].unique()
        cliente_sel = st.selectbox("👤 Buscar Cliente (Escriba nombre o NIT):", lista_clientes)
        
        if cliente_sel:
            # Datos del Cliente Seleccionado
            data_cli = df_view[df_view['nombrecliente'] == cliente_sel].iloc[0]
            
            # --- TARJETA DE GESTIÓN ---
            col_info, col_accion = st.columns([1, 1.5])
            
            with col_info:
                st.markdown(f"""
                <div class="management-card">
                    <h3 style="color:#003366; margin:0;">{data_cli['nombrecliente']}</h3>
                    <p style="color:#666; font-size:14px;">NIT: {data_cli['nit']}</p>
                    <hr>
                    <div style="display:flex; justify-content:space-between;">
                        <div>
                            <small>DEUDA TOTAL</small><br>
                            <b style="font-size:22px; color:#D32F2F;">${data_cli['importe']:,.0f}</b>
                        </div>
                        <div style="text-align:right;">
                            <small>DÍAS MORA</small><br>
                            <b style="font-size:22px; color:#F57C00;">{data_cli['dias_vencido']}</b>
                        </div>
                    </div>
                    <div style="margin-top:15px;">
                        <span style="background:#EEE; padding:5px 10px; border-radius:5px;">{data_cli['Estado']}</span>
                        <span style="float:right;">🧑‍💼 {data_cli['nomvendedor']}</span>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            with col_accion:
                st.markdown("### 🛠️ Herramientas de Contacto")
                
                # --- HERRAMIENTA 1: WHATSAPP ---
                with st.expander("📱 Enviar WhatsApp", expanded=True):
                    # Limpieza Teléfono
                    tel_raw = str(data_cli['telefono1']).replace('.0', '')
                    tel_clean = re.sub(r'\D', '', tel_raw)
                    if len(tel_clean) < 10: tel_clean = "" 
                    
                    c_wa1, c_wa2 = st.columns([1, 2])
                    tel_final = c_wa1.text_input("Celular (+57):", value=tel_clean)
                    
                    # Plantilla Dinámica
                    if data_cli['dias_vencido'] > 60:
                        msg_wa = f"Hola {cliente_sel}, le escribe el área de Cartera de Ferreinox. Su saldo de ${data_cli['importe']:,.0f} presenta {data_cli['dias_vencido']} días de mora. Requerimos pago inmediato para evitar reporte."
                    else:
                        msg_wa = f"Hola {cliente_sel}, un saludo de Ferreinox. Recordamos amablemente su saldo pendiente de ${data_cli['importe']:,.0f}. Agradecemos su gestión."
                    
                    msg_final = c_wa2.text_area("Mensaje:", value=msg_wa, height=100)
                    
                    if tel_final:
                        link = f"https://wa.me/57{tel_final}?text={quote(msg_final)}"
                        st.markdown(f'<a href="{link}" target="_blank" class="whatsapp-btn">🚀 Abrir WhatsApp</a>', unsafe_allow_html=True)
                    else:
                        st.warning("Sin teléfono válido.")

                # --- HERRAMIENTA 2: EMAIL ---
                with st.expander("📧 Enviar Correo Electrónico", expanded=False):
                    email_reg = str(data_cli['e_mail']) if pd.notna(data_cli['e_mail']) and '@' in str(data_cli['e_mail']) else ""
                    
                    email_dest = st.text_input("Para:", value=email_reg)
                    asunto_mail = st.text_input("Asunto:", value=f"Estado de Cuenta - {cliente_sel}")
                    
                    cuerpo_mail = st.text_area("Cuerpo del Correo:", value=f"""
Estimados {cliente_sel},

Adjuntamos relación de su estado de cuenta a la fecha.
Saldo Total: ${data_cli['importe']:,.0f}
Días de Vencimiento: {data_cli['dias_vencido']} días.

Agradecemos su pronta gestión y soporte de pago.

Cordialmente,
Equipo de Cartera Ferreinox
                    """, height=150)
                    
                    if st.button("📨 Enviar Correo Ahora"):
                        if not email_dest:
                            st.error("Falta el correo destinatario.")
                        else:
                            try:
                                # Intentar enviar con Yagmail
                                user = st.secrets["email_credentials"]["sender_email"]
                                password = st.secrets["email_credentials"]["sender_password"]
                                yag = yagmail.SMTP(user, password)
                                yag.send(to=email_dest, subject=asunto_mail, contents=cuerpo_mail)
                                st.success(f"¡Correo enviado a {email_dest}!")
                            except KeyError:
                                st.error("⚠️ Faltan credenciales en secrets.toml. Usando método manual.")
                                # Fallback a Mailto
                                body_encoded = quote(cuerpo_mail)
                                subject_encoded = quote(asunto_mail)
                                link_mail = f"mailto:{email_dest}?subject={subject_encoded}&body={body_encoded}"
                                st.markdown(f'<a href="{link_mail}" target="_blank" class="email-btn">Abrir Gestor de Correo (Outlook/Gmail)</a>', unsafe_allow_html=True)
                            except Exception as e:
                                st.error(f"Error al enviar: {e}")

        st.markdown("---")
        st.subheader("📋 Lista Táctica de Clientes (Filtro Actual)")
        st.dataframe(
            df_view[['Prioridad', 'nombrecliente', 'nit', 'telefono1', 'e_mail', 'dias_vencido', 'importe', 'Estado']],
            column_config={
                "Prioridad": st.column_config.ProgressColumn("Urgencia", min_value=0, max_value=100),
                "importe": st.column_config.NumberColumn("Deuda", format="$%d"),
                "e_mail": "Email",
            },
            hide_index=True,
            use_container_width=True
        )

    # --- PESTAÑA 2: VISIÓN GENERAL ---
    with tab_analisis:
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Distribución por Estado")
            fig_pie = px.pie(df_view, names='Estado', values='importe', hole=0.4, title="Cartera por Estado de Gestión")
            st.plotly_chart(fig_pie, use_container_width=True)
        with c2:
            st.subheader("Mapa de Riesgo")
            fig_scat = px.scatter(df_view, x='dias_vencido', y='importe', color='Estado', size='importe', hover_name='nombrecliente', title="Mora vs Importe")
            st.plotly_chart(fig_scat, use_container_width=True)

    # --- PESTAÑA 3: EXPORTAR ---
    with tab_admin:
        st.header("📥 Descargar Reportes")
        st.write("Descargue la matriz completa para trabajar en Excel si lo prefiere.")
        
        excel_bytes = generar_excel(df_gestion)
        st.download_button(
            label="💾 Descargar Matriz de Gestión (.xlsx)",
            data=excel_bytes,
            file_name="Gestion_Cartera_Total.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

if __name__ == "__main__":
    main()
