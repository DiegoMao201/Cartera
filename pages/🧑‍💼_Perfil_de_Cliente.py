import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO, StringIO
import dropbox
import glob
import unicodedata
import re
from urllib.parse import quote
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.worksheet.table import Table, TableStyleInfo
import yagmail

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="War Room Cobranza", page_icon="🚀", layout="wide")

# --- ESTILOS CSS MEJORADOS ---
st.markdown("""
<style>
    .stApp { background-color: #F4F6F9; }
    .stMetric { background-color: #FFFFFF; border-radius: 8px; padding: 15px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-left: 5px solid #0058A7; }
    .management-card { background-color: #FFFFFF; padding: 25px; border-radius: 12px; border: 1px solid #E5E7EB; box-shadow: 0 4px 10px rgba(0,0,0,0.05); margin-bottom: 20px; }
    .whatsapp-btn { 
        background-color: #25D366; color: white !important; padding: 10px 20px; 
        border-radius: 8px; text-decoration: none; font-weight: bold; 
        display: block; text-align: center; margin-top: 10px; border: 1px solid #20b85c;
    }
    .email-btn { 
        background-color: #EA4335; color: white !important; padding: 10px 20px; 
        border-radius: 8px; text-decoration: none; font-weight: bold; 
        display: block; text-align: center; margin-top: 10px; border: 1px solid #d62516;
    }
    .whatsapp-btn:hover, .email-btn:hover { opacity: 0.9; }
</style>
""", unsafe_allow_html=True)

# ======================================================================================
# --- 1. CARGA DE DATOS BLINDADA (CORRECCIÓN CRÍTICA AQUI) ---
# ======================================================================================

def limpiar_columnas(df):
    """Normaliza nombres y elimina duplicados para evitar el TypeError."""
    if df.empty: return df

    # 1. Normalización básica
    df.columns = [str(c).strip().lower().replace('-', '_').replace('.', '').replace(' ', '_') for c in df.columns]
    
    # 2. Mapeo inteligente
    mapa = {
        'email': 'e_mail', 'correo': 'e_mail', 'mail': 'e_mail', 'correo_electronico': 'e_mail',
        'telefono': 'telefono1', 'celular': 'telefono1', 'movil': 'telefono1',
        'nombre_cliente': 'nombrecliente', 'cliente': 'nombrecliente', 'razon_social': 'nombrecliente',
        'vendedor': 'nomvendedor', 'nombre_vendedor': 'nomvendedor',
        'dias_mora': 'dias_vencido', 'dias': 'dias_vencido', 'vencimiento': 'dias_vencido',
        'valor': 'importe', 'saldo': 'importe', 'total': 'importe', 'deuda': 'importe',
        'nit': 'nit', 'cedula': 'nit'
    }
    df = df.rename(columns=mapa)
    
    # 3. CRÍTICO: Eliminar columnas duplicadas INMEDIATAMENTE después de renombrar
    # Esto previene que existan dos columnas 'importe'
    df = df.loc[:, ~df.columns.duplicated()]
    
    # 4. Garantizar columnas base
    cols_criticas = ['e_mail', 'telefono1', 'nomvendedor', 'nombrecliente', 'nit', 'importe', 'dias_vencido']
    for col in cols_criticas:
        if col not in df.columns:
            df[col] = 0 if col in ['importe', 'dias_vencido'] else "No registrado"
            
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
                # Forzar lectura como string para evitar conversiones prematuras
                df_dropbox = pd.read_csv(StringIO(res.content.decode('latin-1')), sep='|', dtype=str)
                df_final = pd.concat([df_final, df_dropbox])
    except Exception as e:
        print(f"Nota: No se cargó de Dropbox: {e}")

    # --- INTENTO 2: LOCAL ---
    archivos = glob.glob("Cartera_*.xlsx")
    for archivo in archivos:
        try:
            df_hist = pd.read_excel(archivo, dtype=str) # Leer todo como texto para limpiar despues
            if not df_hist.empty:
                # Limpiar filas de totales basura
                df_hist = df_hist[~df_hist.iloc[:, 0].astype(str).str.contains("Total", na=False, case=False)]
                df_final = pd.concat([df_final, df_hist])
        except Exception: pass

    if df_final.empty: return pd.DataFrame()

    # --- LIMPIEZA PROFUNDA ---
    df_final = limpiar_columnas(df_final)
    
    # Conversión Numérica Robusta (Elimina $, comas, puntos raros)
    for col in ['importe', 'dias_vencido']:
        # Limpiar caracteres no numéricos excepto el punto decimal y el menos
        df_final[col] = df_final[col].astype(str).str.replace(r'[$,]', '', regex=True)
        df_final[col] = pd.to_numeric(df_final[col], errors='coerce').fillna(0)

    # Filtros adicionales de basura
    if 'serie' in df_final.columns:
        df_final = df_final[~df_final['serie'].astype(str).str.contains('W|X', case=False, na=False)]
    
    return df_final

# ======================================================================================
# --- 2. CEREBRO DE PRIORIZACIÓN ---
# ======================================================================================

def procesar_cartera(df):
    if df.empty: return pd.DataFrame()
    df = df[df['importe'] > 1000].copy() # Filtrar saldos insignificantes
    
    # Agrupar por Cliente (Usando un fillna previo para agrupar bien)
    cols_group = ['nombrecliente', 'nit', 'nomvendedor', 'telefono1', 'e_mail']
    df[cols_group] = df[cols_group].fillna("Sin Datos")
    
    kpis = df.groupby(cols_group).agg({
        'importe': 'sum',
        'dias_vencido': 'max', # Tomamos la factura más vieja
        'nit': 'count' # Usamos nit para contar facturas
    }).rename(columns={'nit': 'num_facturas'}).reset_index()
    
    # Estrategia
    def estrategia(dias):
        if dias > 90: return "🔴 JURÍDICO"
        if dias > 60: return "⛔ PRE-JURÍDICO"
        if dias > 30: return "🟠 COBRO ACTIVO"
        if dias > 0: return "🟡 PREVENTIVO"
        return "🟢 AL DÍA"

    kpis['Estado'] = kpis['dias_vencido'].apply(estrategia)
    
    # Score de Prioridad (Normalizado 0-100)
    kpis['Prioridad'] = (
        (kpis['dias_vencido'].clip(upper=120) / 120 * 60) + 
        (kpis['importe'].clip(upper=50000000) / 50000000 * 40)
    )
    
    return kpis.sort_values('Prioridad', ascending=False).reset_index(drop=True)

# ======================================================================================
# --- 3. EXPORTACIÓN EXCEL ---
# ======================================================================================
def generar_excel(df):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Matriz de Cobro"
    
    headers = ["Score", "Cliente", "NIT", "Teléfono", "Email", "Deuda Total", "Días Mora", "Estado", "Vendedor", "# Facturas"]
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
            row['nomvendedor'],
            row['num_facturas']
        ])
    
    # Convertir a Tabla Excel real
    tab = Table(displayName="TablaCartera", ref=f"A1:J{len(df)+1}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    
    wb.save(output)
    return output.getvalue()

# ======================================================================================
# --- 4. INTERFAZ ---
# ======================================================================================

def main():
    st.title("🚀 Centro de Comando de Cobranza")
    
    # 1. Carga
    with st.spinner('Conectando bases de datos y consolidando cartera...'):
        df_raw = cargar_datos_maestros()
    
    if df_raw.empty:
        st.warning("⚠️ No se encontraron datos de cartera. Verifique la conexión a Dropbox o cargue archivos 'Cartera_*.xlsx' en la carpeta.")
        return # Detener ejecución limpiamente

    # 2. Procesamiento
    df_gestion = procesar_cartera(df_raw)
    
    # 3. Sidebar Filtros
    st.sidebar.markdown("### 🔍 Filtros Inteligentes")
    
    lista_vendedores = ["TODOS"] + sorted(list(df_gestion['nomvendedor'].unique()))
    filtro_vend = st.sidebar.selectbox("Vendedor:", lista_vendedores)
    
    lista_estados = ["TODOS"] + sorted(list(df_gestion['Estado'].unique()))
    filtro_estado = st.sidebar.selectbox("Estado de Cartera:", lista_estados)
    
    # Aplicar filtros
    df_view = df_gestion.copy()
    if filtro_vend != "TODOS":
        df_view = df_view[df_view['nomvendedor'] == filtro_vend]
    if filtro_estado != "TODOS":
        df_view = df_view[df_view['Estado'] == filtro_estado]
        
    # 4. KPIs
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Clientes", f"{len(df_view)}")
    col2.metric("Cartera Total", f"${df_view['importe'].sum():,.0f}")
    criticos = df_view[df_view['dias_vencido'] > 60].shape[0]
    col3.metric("🚨 Riesgo Alto", f"{criticos} Clientes")
    ticket = df_view['importe'].mean() if not df_view.empty else 0
    col4.metric("Ticket Promedio", f"${ticket:,.0f}")
    
    st.markdown("---")
    
    # 5. TABS PRINCIPALES
    tab1, tab2, tab3 = st.tabs(["⚡ GESTIÓN RÁPIDA", "📊 ANÁLISIS VISUAL", "📥 DESCARGAS"])
    
    # --- PESTAÑA 1: GESTIÓN ---
    with tab1:
        c1, c2 = st.columns([1, 2])
        with c1:
            st.markdown("### 👤 Seleccionar Cliente")
            opciones_cliente = df_view['nombrecliente'].unique()
            cliente_sel = st.selectbox("Busque por nombre:", opciones_cliente)
            
        if cliente_sel:
            data_cli = df_view[df_view['nombrecliente'] == cliente_sel].iloc[0]
            
            # Tarjeta de Datos
            st.markdown(f"""
            <div class="management-card">
                <h3 style="color:#003366; margin:0;">{data_cli['nombrecliente']}</h3>
                <p><b>NIT:</b> {data_cli['nit']} | <b>Vendedor:</b> {data_cli['nomvendedor']}</p>
                <hr>
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <div>
                        <span style="font-size:12px; color:#666;">DEUDA TOTAL</span><br>
                        <span style="font-size:24px; font-weight:bold; color:#B71C1C;">${data_cli['importe']:,.0f}</span>
                    </div>
                    <div>
                        <span style="font-size:12px; color:#666;">DÍAS DE MORA</span><br>
                        <span style="font-size:24px; font-weight:bold; color:#FF6F00;">{int(data_cli['dias_vencido'])} días</span>
                    </div>
                    <div style="text-align:right;">
                        <span class="status-badge" style="background-color:#0058A7;">{data_cli['Estado']}</span>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            # Acciones
            col_wa, col_email = st.columns(2)
            
            # --- WHATSAPP ---
            with col_wa:
                st.markdown("#### 💬 WhatsApp")
                tel_limpio = re.sub(r'\D', '', str(data_cli['telefono1']))
                tel_input = st.text_input("Número (+57):", value=tel_limpio if len(tel_limpio) >= 10 else "")
                
                msg_base = f"Hola {data_cli['nombrecliente']}, le saludamos de Cartera. "
                if data_cli['dias_vencido'] > 60:
                    msg_base += f"Su saldo vencido es de ${data_cli['importe']:,.0f} con {int(data_cli['dias_vencido'])} días de mora. Requerimos pago inmediato."
                else:
                    msg_base += f"Recordamos amablemente su saldo pendiente de ${data_cli['importe']:,.0f}. Agradecemos su gestión."
                
                msg_wa = st.text_area("Mensaje:", value=msg_base, height=100)
                
                if tel_input:
                    link_wa = f"https://wa.me/57{tel_input}?text={quote(msg_wa)}"
                    st.markdown(f'<a href="{link_wa}" target="_blank" class="whatsapp-btn">Enviar WhatsApp 🚀</a>', unsafe_allow_html=True)
            
            # --- EMAIL ---
            with col_email:
                st.markdown("#### 📧 Email")
                email_val = str(data_cli['e_mail']) if "@" in str(data_cli['e_mail']) else ""
                destinatario = st.text_input("Correo Destino:", value=email_val)
                asunto = st.text_input("Asunto:", value=f"Estado de Cuenta - {data_cli['nombrecliente']}")
                
                if st.button("Enviar Correo Electrónico"):
                    if not destinatario:
                        st.error("Falta el correo.")
                    else:
                        try:
                            # Intento usar credenciales si existen
                            user = st.secrets["email_credentials"]["sender_email"]
                            pwd = st.secrets["email_credentials"]["sender_password"]
                            yag = yagmail.SMTP(user, pwd)
                            yag.send(to=destinatario, subject=asunto, contents=msg_base)
                            st.success(f"Correo enviado a {destinatario}")
                        except Exception:
                            # Fallback a mailto client side
                            link_mail = f"mailto:{destinatario}?subject={quote(asunto)}&body={quote(msg_base)}"
                            st.markdown(f'<a href="{link_mail}" target="_blank" class="email-btn">Abrir Outlook/Gmail Local</a>', unsafe_allow_html=True)
                            st.info("Nota: Se abrio su gestor de correo local porque no hay credenciales automáticas configuradas.")

    # --- PESTAÑA 2: GRÁFICOS ---
    with tab2:
        if not df_view.empty:
            c1, c2 = st.columns(2)
            with c1:
                fig = px.sunburst(df_view, path=['Estado', 'nomvendedor'], values='importe', title="Distribución de Cartera")
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig2 = px.scatter(df_view, x="dias_vencido", y="importe", size="importe", color="Estado", hover_name="nombrecliente", title="Mapa de Riesgo (Mora vs Deuda)")
                st.plotly_chart(fig2, use_container_width=True)
                
    # --- PESTAÑA 3: DESCARGAS ---
    with tab3:
        st.subheader("Generar Reportes")
        st.write("Descargue la matriz de gestión actual con filtros aplicados.")
        excel_data = generar_excel(df_view)
        st.download_button(
            label="💾 Descargar Excel (.xlsx)",
            data=excel_data,
            file_name="Gestion_Cartera_Corte.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

if __name__ == "__main__":
    main()
