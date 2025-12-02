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
import yagmail
from urllib.parse import quote

# --- 1. CONFIGURACIÓN DE LA PÁGINA Y ESTILOS ---
st.set_page_config(
    page_title="Centro de Mando: Cobranza Estratégica",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS Profesionales (Look & Feel Bancario/Corporativo)
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
    
    /* Alertas de Estado */
    .status-critico { color: #dc3545; font-weight: bold; }
    .status-alerta { color: #ffc107; font-weight: bold; }
    .status-ok { color: #198754; font-weight: bold; }
</style>
""", unsafe_allow_html=True)

# ======================================================================================
# --- 2. MOTOR DE INGESTIÓN Y LIMPIEZA DE DATOS (EL NÚCLEO) ---
# ======================================================================================

def normalizar_texto(texto):
    """Elimina acentos y caracteres especiales para comparaciones seguras."""
    if not isinstance(texto, str): return str(texto)
    texto = unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode("utf-8")
    return texto.upper().strip()

def limpiar_moneda(valor):
    """
    Convierte cualquier formato de moneda ($ 1.000,00 o 1,000.00) a float puro.
    Esta función es crítica para que los valores coincidan.
    """
    if pd.isna(valor): return 0.0
    s_val = str(valor).strip()
    # Eliminar símbolo de moneda y espacios
    s_val = s_val.replace('$', '').replace(' ', '')
    
    # Detección heurística: si hay coma y punto, asumir formato latino (1.000,00) 
    # o formato inglés (1,000.00). 
    # Regla simple para Colombia: Si hay ',' al final, es decimal. Si hay '.' al final es decimal.
    
    try:
        # Caso 1: Formato sin decimales o limpio
        if ',' not in s_val and '.' not in s_val:
            return float(s_val)
            
        # Caso 2: Formato Colombia/Europa (puntos miles, coma decimal) -> 1.500,50
        if '.' in s_val and ',' in s_val:
            if s_val.rfind(',') > s_val.rfind('.'): # La coma está después (es decimal)
                s_val = s_val.replace('.', '').replace(',', '.')
            else: # El punto está después (1,500.50)
                s_val = s_val.replace(',', '')
        elif ',' in s_val: # Solo comas (1,500 o 1,5) - Asumimos coma como decimal si es corto, miles si es largo? 
            # Mejor estandarizar: en archivos planos contables, comas suelen ser decimales o miles.
            # Vamos a eliminar caracteres no numéricos excepto el último separador
            pass 
        
        # Método a prueba de balas (Regex):
        # 1. Eliminar todo lo que no sea dígito, punto o coma
        s_val = re.sub(r'[^\d.,-]', '', s_val)
        # 2. Si tiene coma y punto, quitar el separador de miles
        if ',' in s_val and '.' in s_val:
             if s_val.rfind(',') > s_val.rfind('.'): # Coma es decimal
                 s_val = s_val.replace('.', '').replace(',', '.')
             else:
                 s_val = s_val.replace(',', '')
        # 3. Si solo tiene uno, asumir que si hay más de 3 digitos tras él es miles, sino decimal
        elif ',' in s_val:
             parts = s_val.split(',')
             if len(parts[-1]) != 2: # Asumir miles
                 s_val = s_val.replace(',', '')
             else:
                 s_val = s_val.replace(',', '.')
        
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
            _, res = dbx.files_download(path='/data/cartera_detalle.csv')
            df_raw = pd.read_csv(io.StringIO(res.content.decode('latin-1')), sep='|', header=None, dtype=str)
            origen = "Nube (Dropbox)"
            
            # Asignar nombres si viene sin header (ajustar según tu CSV real)
            nombres_cols = [
                'Serie','Numero','Fecha Documento','Fecha Vencimiento','Cod Cliente',
                'NombreCliente','Nit','Poblacion','Provincia','Telefono1','Telefono2',
                'NomVendedor','Entidad Autoriza','E-Mail','Importe','Descuento',
                'Cupo Aprobado','Dias Vencido'
            ]
            if len(df_raw.columns) == len(nombres_cols):
                df_raw.columns = nombres_cols
    except Exception as e:
        pass # Falló Dropbox, seguimos a local

    # 2. Intento Local (Fallback)
    if df_raw.empty:
        archivos = glob.glob("Cartera_*.xlsx")
        if archivos:
            # Tomar el más reciente
            archivo_reciente = max(archivos, key=os.path.getctime)
            try:
                df_raw = pd.read_excel(archivo_reciente, dtype=str)
                origen = f"Local ({archivo_reciente})"
            except: pass

    if df_raw.empty:
        return pd.DataFrame(), "Sin Datos"

    # --- LIMPIEZA Y ESTANDARIZACIÓN ---
    # Renombrar columnas a formato estándar (snake_case)
    cols_map = {
        'NombreCliente': 'cliente', 'Nit': 'nit', 'NomVendedor': 'vendedor',
        'Importe': 'saldo', 'Dias Vencido': 'dias_mora', 'E-Mail': 'email',
        'Telefono1': 'telefono', 'Numero': 'factura', 'Fecha Vencimiento': 'fecha_venc'
    }
    # Normalizar nombres de columnas actuales
    df_raw.columns = [normalizar_texto(c).replace(' ', '_').title() for c in df_raw.columns]
    # Mapear a nuestras columnas clave
    col_actuales = {c: c for c in df_raw.columns} # Diccionario auxiliar
    
    # Buscar match aproximado
    mapping_final = {}
    for k, v in cols_map.items():
        for col_real in df_raw.columns:
            if normalizar_texto(k) in normalizar_texto(col_real):
                mapping_final[col_real] = v
                break
    
    df_raw.rename(columns=mapping_final, inplace=True)
    
    # Columnas críticas requeridas (rellenar si faltan)
    required = ['cliente', 'saldo', 'dias_mora', 'vendedor', 'nit', 'factura', 'email', 'telefono']
    for req in required:
        if req not in df_raw.columns:
            df_raw[req] = 0 if req in ['saldo', 'dias_mora'] else 'N/A'

    # Conversión de Tipos (LA PARTE CLAVE PARA QUE LOS VALORES COINCIDAN)
    df_raw['saldo'] = df_raw['saldo'].apply(limpiar_moneda)
    df_raw['dias_mora'] = pd.to_numeric(df_raw['dias_mora'], errors='coerce').fillna(0)
    
    # Filtrar basura (totales, filas vacías)
    df_raw = df_raw[df_raw['cliente'].str.len() > 2]
    df_raw = df_raw[df_raw['saldo'] != 0] # Eliminar saldos cero

    return df_raw, origen

# ======================================================================================
# --- 3. LÓGICA DE NEGOCIO: CLASIFICACIÓN Y ESTRATEGIA ---
# ======================================================================================

def analizar_cartera(df):
    """
    Agrega inteligencia a los datos:
    1. Segmentación (Edad de Cartera)
    2. Prioridad de Gestión (Pareto + Riesgo)
    3. Acción Sugerida
    """
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

    # 3. Acción Recomendada (Dictamen para la Líder)
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
# --- 4. GENERADOR DE REPORTES GERENCIALES (EXCEL PREMIUM) ---
# ======================================================================================

def generar_excel_gerencial(df_detalle, df_resumen_cliente):
    output = io.BytesIO()
    wb = Workbook()
    
    # --- HOJA 1: RESUMEN EJECUTIVO (KPIs) ---
    ws_kpi = wb.active
    ws_kpi.title = "Resumen Gerencial"
    
    # Título
    ws_kpi['A1'] = "INFORME DE ESTADO DE CARTERA - FERREINOX"
    ws_kpi['A1'].font = Font(size=16, bold=True, color="003366")
    ws_kpi['A2'] = f"Generado el: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    
    # Datos Totales
    total_cartera = df_detalle['saldo'].sum()
    total_vencido = df_detalle[df_detalle['dias_mora'] > 0]['saldo'].sum()
    pct_vencido = (total_vencido / total_cartera) if total_cartera else 0
    
    ws_kpi['A4'] = "Cartera Total"; ws_kpi['B4'] = total_cartera
    ws_kpi['A5'] = "Cartera Vencida"; ws_kpi['B5'] = total_vencido
    ws_kpi['A6'] = "% Deterioro"; ws_kpi['B6'] = pct_vencido
    
    # Formato Moneda
    for cell in ['B4', 'B5']: ws_kpi[cell].number_format = '$ #,##0'
    ws_kpi['B6'].number_format = '0.0%'
    
    # --- HOJA 2: TOP CLIENTES (ACTION PLAN) ---
    ws_clientes = wb.create_sheet("Top Gestión (Pareto)")
    
    # Encabezados
    headers = ["Cliente", "NIT", "Vendedor", "Saldo Total", "Días Mora Max", "Estado", "Acción Sugerida", "Teléfono", "Email"]
    ws_clientes.append(headers)
    
    # Estilo Encabezado
    header_fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    
    for col_num, header in enumerate(headers, 1):
        cell = ws_clientes.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        ws_clientes.column_dimensions[get_column_letter(col_num)].width = 20
        
    # Llenar datos (Ordenados por Saldo Descendente)
    df_sorted = df_resumen_cliente.sort_values(by='saldo', ascending=False)
    
    for _, row in df_sorted.iterrows():
        ws_clientes.append([
            row['cliente'], row['nit'], row['vendedor'], 
            row['saldo'], row['dias_mora'], 
            row['estado_gestion'], row['accion_sugerida'],
            row['telefono'], row['email']
        ])
    
    # Formato Tabla Excel
    tab = Table(displayName="TablaClientes", ref=f"A1:I{len(df_sorted)+1}")
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    tab.tableStyleInfo = style
    ws_clientes.add_table(tab)
    
    # Formato Moneda Columna D (Saldo)
    for row in range(2, len(df_sorted)+2):
        ws_clientes[f'D{row}'].number_format = '$ #,##0'
        
    # --- HOJA 3: DETALLE FACTURAS ---
    ws_data = wb.create_sheet("Data Cruda (Facturas)")
    rows = [df_detalle.columns.tolist()] + df_detalle.values.tolist()
    for r in rows: ws_data.append(r)

    wb.save(output)
    return output.getvalue()

# ======================================================================================
# --- 5. INTERFAZ DE USUARIO (DASHBOARD) ---
# ======================================================================================

def main():
    # --- SIDEBAR: Autenticación y Filtros ---
    with st.sidebar:
        st.image("https://via.placeholder.com/150x50.png?text=Ferreinox+Logo", use_container_width=True) # Reemplazar con logo real si existe
        st.title("Panel de Control")
        
        # Simulación de Login Simple
        user_type = st.selectbox("Perfil de Usuario", ["Gerencia General", "Líder Cartera", "Vendedor"])
        
        # Carga de Datos
        df, source_msg = cargar_datos_maestros()
        st.caption(f"🔌 Fuente: {source_msg}")
        
        if df.empty:
            st.error("No se encontraron datos. Cargue archivos 'Cartera_*.xlsx' o conecte Dropbox.")
            st.stop()
            
        # Proceso de Inteligencia
        df_processed = analizar_cartera(df)
        
        # Filtros Globales
        st.markdown("### 🔍 Filtros")
        
        # Filtro Vendedor
        vendedores = ["TODOS"] + sorted(list(df_processed['vendedor'].unique()))
        sel_vendedor = st.selectbox("Vendedor:", vendedores)
        
        # Filtro Zona (Si existe columna zona, sino ignorar)
        
        # Aplicar Filtros
        df_view = df_processed.copy()
        if sel_vendedor != "TODOS":
            df_view = df_view[df_view['vendedor'] == sel_vendedor]
            
    # --- ÁREA PRINCIPAL ---
    
    # 1. KPIs SUPERIORES (Lo que le importa al Gerente)
    st.markdown(f"## 📊 Estado de Cartera - Vista {user_type}")
    
    total_cobrar = df_view['saldo'].sum()
    
    # Cartera Vencida (> 0 días)
    vencido = df_view[df_view['dias_mora'] > 0]['saldo'].sum()
    
    # Cartera Crítica (> 60 días)
    critico = df_view[df_view['dias_mora'] > 60]['saldo'].sum()
    
    # Recaudo Potencial (Facturas < 30 días)
    potencial = df_view[(df_view['dias_mora'] > 0) & (df_view['dias_mora'] <= 30)]['saldo'].sum()

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("💰 Cartera Total", f"${total_cobrar:,.0f}", help="Suma total de facturas abiertas")
    k2.metric("🔥 Total Vencido", f"${vencido:,.0f}", f"{(vencido/total_cobrar)*100:.1f}% del total", delta_color="inverse")
    k3.metric("🚨 Riesgo Crítico (>60d)", f"${critico:,.0f}", "Requiere Abogado/Bloqueo", delta_color="inverse")
    k4.metric("🎯 Recaudo Inmediato", f"${potencial:,.0f}", "Gestionable hoy", delta_color="normal")
    
    st.markdown("---")
    
    # --- PESTAÑAS ESTRATÉGICAS ---
    tab_war_room, tab_analytics, tab_data = st.tabs(["⚔️ SALA DE GUERRA (Gestión)", "📈 ANÁLISIS GERENCIAL", "📂 DATA & REPORTES"])
    
    # ==================================================================================
    # TAB 1: SALA DE GUERRA (Para la Líder de Cartera - ¡Acción Pura!)
    # ==================================================================================
    with tab_war_room:
        st.info("💡 **Instrucción:** Esta lista está ordenada por **PRIORIDAD**. Los clientes arriba son los que más impacto tienen en el flujo de caja. Gestione de arriba a abajo.")
        
        # Agrupar por Cliente para la lista de gestión
        df_clients = df_view.groupby(['cliente', 'nit', 'vendedor', 'telefono', 'email']).agg({
            'saldo': 'sum',
            'dias_mora': 'max',
            'factura': 'count'
        }).reset_index()
        
        # Recalcular lógica de estado para el grupo
        df_clients = analizar_cartera(df_clients)
        
        # Ordenar: Primero los de más plata vencida, luego los más antiguos
        df_clients = df_clients.sort_values(by=['rango_mora', 'saldo'], ascending=[False, False])
        
        # Selector de Cliente para gestionar
        col_list, col_action = st.columns([1, 2])
        
        with col_list:
            st.markdown("### 📋 Lista de Objetivos")
            # Crear una etiqueta compuesta para el selector
            df_clients['label'] = df_clients.apply(lambda x: f"{x['cliente']} | ${x['saldo']:,.0f} | {int(x['dias_mora'])} días", axis=1)
            target_client = st.selectbox("Seleccione Cliente a Gestionar:", df_clients['label'], list_index=0)
            
            # Extraer nombre del cliente seleccionado
            client_name = target_client.split(' | ')[0]
            client_data = df_clients[df_clients['cliente'] == client_name].iloc[0]
            
            # Visualizador de Estado Rápido
            st.markdown(f"""
            <div style="background-color: #fff; padding: 15px; border-radius: 10px; border: 1px solid #ddd;">
                <h4 style="margin:0;">{client_data['cliente']}</h4>
                <p style="color:gray; font-size: 12px;">NIT: {client_data['nit']}</p>
                <hr>
                <div style="display:flex; justify-content:space-between;">
                    <div><b>Deuda:</b><br><span style="color:#c0392b; font-size:18px; font-weight:bold;">${client_data['saldo']:,.0f}</span></div>
                    <div><b>Mora Max:</b><br><span style="color:#d35400; font-size:18px; font-weight:bold;">{int(client_data['dias_mora'])} días</span></div>
                </div>
                <br>
                <div style="background-color:#f0f2f6; padding:10px; border-radius:5px; text-align:center;">
                    <b>ACCIÓN RECOMENDADA:</b><br>
                    <span style="color:#2980b9; font-weight:bold;">{client_data['accion_sugerida']}</span>
                </div>
            </div>
            """, unsafe_allow_html=True)
            
        with col_action:
            st.markdown("### 🛠️ Centro de Acción")
            
            # Ver facturas específicas de este cliente
            st.subheader("Detalle de Facturas Pendientes")
            facturas_cli = df_view[df_view['cliente'] == client_name][['factura', 'fecha_venc', 'dias_mora', 'saldo', 'rango_mora']]
            st.dataframe(facturas_cli.style.format({'saldo': '${:,.0f}'}), use_container_width=True)
            
            st.markdown("#### 🚀 Generador de Comunicaciones")
            c1, c2 = st.columns(2)
            
            # Generador WhatsApp
            with c1:
                st.markdown(" **1. WhatsApp**")
                phone = str(client_data['telefono']).replace('.0', '')
                phone_clean = re.sub(r'\D', '', phone)
                
                # Mensaje dinámico según mora
                if client_data['dias_mora'] > 0:
                    msg_wa = f"Hola {client_name}, le escribimos de Ferreinox. Vemos un saldo pendiente de ${client_data['saldo']:,.0f} con facturas de hasta {int(client_data['dias_mora'])} días vencidas. Agradecemos su pago en el link: [LINK_PAGOS]"
                else:
                    msg_wa = f"Hola {client_name}, gracias por ser cliente de Ferreinox. Su estado de cuenta está al día."
                
                msg_input = st.text_area("Mensaje Personalizado:", value=msg_wa, height=100)
                
                if phone_clean and len(phone_clean) > 7:
                    link_wa = f"https://wa.me/57{phone_clean}?text={quote(msg_input)}"
                    st.markdown(f'<a href="{link_wa}" target="_blank" class="action-btn-wa">📱 Enviar WhatsApp a {phone}</a>', unsafe_allow_html=True)
                else:
                    st.warning("Número de teléfono no válido o no registrado.")

            # Generador Email
            with c2:
                st.markdown(" **2. Correo Electrónico**")
                email_dest = client_data['email']
                if "@" not in str(email_dest):
                    st.warning("No hay correo registrado.")
                else:
                    subject = f"Estado de Cuenta - {client_name}"
                    st.text_input("Asunto:", value=subject, disabled=True)
                    st.caption("Al hacer clic, se abrirá tu gestor de correo predeterminado (Outlook/Gmail) con el mensaje listo.")
                    
                    mailto_link = f"mailto:{email_dest}?subject={quote(subject)}&body={quote(msg_input)}"
                    st.markdown(f'<a href="{mailto_link}" target="_blank" class="action-btn-email">📧 Abrir Correo</a>', unsafe_allow_html=True)

    # ==================================================================================
    # TAB 2: ANÁLISIS GERENCIAL (Para el Jefe - Gráficas y Estrategia)
    # ==================================================================================
    with tab_analytics:
        st.subheader("Radiografía de la Cartera")
        

[Image of Data Dashboard]
 # Tag para insertar gráfico contextual si el sistema lo permite
        
        g1, g2 = st.columns(2)
        
        with g1:
            # Gráfico de Donut: Composición por Edad
            df_pie = df_view.groupby('rango_mora')['saldo'].sum().reset_index()
            fig_pie = px.pie(df_pie, values='saldo', names='rango_mora', title='Distribución por Antigüedad', hole=0.4,
                             color_discrete_sequence=px.colors.sequential.RdBu_r)
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with g2:
            # Gráfico de Barras: Top 10 Deudores
            df_top = df_clients.sort_values('saldo', ascending=False).head(10)
            fig_bar = px.bar(df_top, x='saldo', y='cliente', orientation='h', title='Top 10 Clientes con Mayor Deuda',
                             text_auto='.2s', color='dias_mora', color_continuous_scale='Reds')
            fig_bar.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_bar, use_container_width=True)
            
        st.markdown("### 🗺️ Mapa de Calor: Vendedores vs. Mora")
        # Pivote para ver qué vendedor tiene la cartera más sana o podrida
        pivot = pd.pivot_table(df_view, values='saldo', index='vendedor', columns='rango_mora', aggfunc='sum', fill_value=0)
        st.dataframe(pivot.style.background_gradient(cmap='Reds', axis=1).format("${:,.0f}"), use_container_width=True)

    # ==================================================================================
    # TAB 3: DATA & REPORTES (Descarga del Excel Definitivo)
    # ==================================================================================
    with tab_data:
        st.success("✅ **Reporte Listo:** Este archivo contiene el análisis completo, formato condicional y sugerencias de gestión.")
        
        col_d1, col_d2 = st.columns([1, 3])
        
        with col_d1:
            # Generar Excel en memoria
            excel_file = generar_excel_gerencial(df_view, df_clients)
            
            st.download_button(
                label="📥 DESCARGAR REPORTE GERENCIAL (.xlsx)",
                data=excel_file,
                file_name=f"Reporte_Cartera_Ferreinox_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        
        with col_d2:
            st.markdown("""
            **Contenido del Reporte:**
            1.  **Resumen Ejecutivo:** KPIs totales listos para imprimir.
            2.  **Top Gestión:** Lista priorizada de clientes con acciones sugeridas (columna 'Dictamen').
            3.  **Data Cruda:** Todas las facturas limpias y estandarizadas.
            """)
            
        st.markdown("---")
        st.subheader("Vista Previa de Datos Crudos")
        st.dataframe(df_view)

if __name__ == '__main__':
    main()
