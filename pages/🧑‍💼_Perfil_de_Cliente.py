# ======================================================================================
# ARCHIVO: pages/🧑‍💼_Perfil_de_Cliente.py (Versión Potenciada)
# ======================================================================================
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import glob
import re
import unicodedata
from datetime import datetime
from fpdf import FPDF
from io import BytesIO

st.set_page_config(page_title="Perfil de Cliente", page_icon="🧑‍💼", layout="wide")

# --- GUARDIA DE SEGURIDAD ---
if 'authentication_status' not in st.session_state or not st.session_state['authentication_status']:
    st.warning("Por favor, inicie sesión en el 📈 Tablero Principal para acceder a esta página.")
    st.stop()

# --- FUNCIONES AUXILIARES (Sin cambios en las funciones PDF) ---
class PDF(FPDF):
    def header(self):
        try:
            self.image("LOGO FERREINOX SAS BIC 2024.png", 10, 8, 80)
        except FileNotFoundError:
            self.set_font('Arial', 'B', 12); self.cell(80, 10, 'Logo no encontrado', 0, 0, 'L')
        self.set_font('Arial', 'B', 18); self.cell(0, 10, 'Estado de Cuenta', 0, 1, 'R')
        self.set_font('Arial', 'I', 9); self.cell(0, 10, f'Generado el: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', 0, 1, 'R')
        self.ln(5); self.set_line_width(0.5); self.set_draw_color(220, 220, 220); self.line(10, 35, 200, 35); self.ln(10)

    def footer(self):
        self.set_y(-40)
        self.set_font('Arial', 'I', 9); self.set_text_color(100, 100, 100)
        self.cell(0, 6, "Para ingresar al portal de pagos, utiliza el NIT como 'usuario' y el Codigo de Cliente como 'codigo unico interno'.", 0, 1, 'C')
        self.set_font('Arial', 'B', 11); self.set_text_color(0, 0, 0)
        self.cell(0, 8, 'Realiza tu pago de forma facil y segura aqui:', 0, 1, 'C')
        self.set_font('Arial', 'BU', 12); self.set_text_color(4, 88, 167)
        link = "https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/"
        self.cell(0, 10, "Portal de Pagos Ferreinox SAS BIC", 0, 1, 'C', link=link)

def generar_pdf_estado_cuenta(datos_cliente: pd.DataFrame):
    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=45)
    pdf.add_page()
    if datos_cliente.empty:
        pdf.set_font('Arial', 'B', 12); pdf.cell(0, 10, 'No se encontraron facturas para este cliente.', 0, 1, 'C')
        return bytes(pdf.output())
    datos_cliente_ordenados = datos_cliente.sort_values(by='fecha_vencimiento', ascending=True)
    info_cliente = datos_cliente_ordenados.iloc[0]
    
    hoy = pd.to_datetime(datetime.now())
    datos_cliente_ordenados['dias_vencido_hoy'] = (hoy - datos_cliente_ordenados['fecha_vencimiento']).dt.days

    pdf.set_font('Arial', 'B', 11); pdf.cell(40, 10, 'Cliente:', 0, 0); pdf.set_font('Arial', '', 11); pdf.cell(0, 10, info_cliente.get('nombrecliente', ''), 0, 1)
    cod_cliente_val = info_cliente.get('cod_cliente')
    cod_cliente_str = str(int(cod_cliente_val)) if pd.notna(cod_cliente_val) else "N/A"
    pdf.set_font('Arial', 'B', 11); pdf.cell(40, 10, 'Codigo de Cliente:', 0, 0); pdf.set_font('Arial', '', 11)
    pdf.cell(0, 10, cod_cliente_str, 0, 1); pdf.ln(5)
    
    pdf.set_font('Arial', '', 10); mensaje = "Apreciado cliente, a continuacion encontrara el detalle de su estado de cuenta a la fecha. Le agradecemos por su continua confianza en Ferreinox SAS BIC y le invitamos a revisar los vencimientos para mantener su cartera al dia."
    pdf.set_text_color(128, 128, 128); pdf.multi_cell(0, 5, mensaje, 0, 'J'); pdf.set_text_color(0, 0, 0); pdf.ln(10)
    
    pdf.set_font('Arial', 'B', 10); pdf.set_fill_color(0, 56, 101); pdf.set_text_color(255, 255, 255)
    pdf.cell(30, 10, 'Factura', 1, 0, 'C', 1); pdf.cell(40, 10, 'Fecha Factura', 1, 0, 'C', 1)
    pdf.cell(40, 10, 'Fecha Vencimiento', 1, 0, 'C', 1); pdf.cell(40, 10, 'Importe', 1, 1, 'C', 1)
    
    pdf.set_font('Arial', '', 10)
    total_importe = 0
    for _, row in datos_cliente_ordenados.iterrows():
        pdf.set_text_color(0, 0, 0)
        if row.get('dias_vencido_hoy', 0) > 0 and pd.isnull(row.get('fecha_saldado')):
            pdf.set_fill_color(248, 241, 241)
        else:
            pdf.set_fill_color(255, 255, 255)
        total_importe += row.get('importe', 0)
        numero_factura_str = str(int(row.get('numero'))) if pd.notna(row.get('numero')) else "N/A"
        fecha_doc_str = row['fecha_documento'].strftime('%d/%m/%Y') if pd.notna(row.get('fecha_documento')) else "N/A"
        fecha_ven_str = row['fecha_vencimiento'].strftime('%d/%m/%Y') if pd.notna(row.get('fecha_vencimiento')) else "N/A"
        pdf.cell(30, 10, numero_factura_str, 1, 0, 'C', 1)
        pdf.cell(40, 10, fecha_doc_str, 1, 0, 'C', 1)
        pdf.cell(40, 10, fecha_ven_str, 1, 0, 'C', 1)
        pdf.cell(40, 10, f"${row.get('importe', 0):,.0f}", 1, 1, 'R', 1)
        
    pdf.set_text_color(0, 0, 0)
    pdf.set_font('Arial', 'B', 10); pdf.set_fill_color(0, 56, 101); pdf.set_text_color(255, 255, 255)
    pdf.cell(110, 10, 'TOTAL ADEUDADO', 1, 0, 'R', 1)
    pdf.cell(40, 10, f"${total_importe:,.0f}", 1, 1, 'R', 1)
    
    return bytes(pdf.output())

def normalizar_nombre(nombre: str) -> str:
    if not isinstance(nombre, str): return ""
    nombre = nombre.upper().strip().replace('.', '')
    nombre = ''.join(c for c in unicodedata.normalize('NFD', nombre) if unicodedata.category(c) != 'Mn')
    return ' '.join(nombre.split())

st.title("🧑‍💼 Dossier de Inteligencia de Cliente")

# --- Carga de Datos (Tu código original, bien hecho) ---
@st.cache_data
def cargar_datos_historicos():
    mapa_columnas = {
        'Serie': 'serie', 'Número': 'numero', 'Fecha Documento': 'fecha_documento',
        'Fecha Vencimiento': 'fecha_vencimiento', 'Fecha Saldado': 'fecha_saldado',
        'NOMBRECLIENTE': 'nombrecliente', 'Población': 'poblacion', 'Provincia': 'provincia',
        'IMPORTE': 'importe', 'RIESGOCONCEDIDO': 'riesgoconcedido', 'NOMVENDEDOR': 'nomvendedor',
        'DIAS_VENCIDO': 'dias_vencido', 'Estado': 'estado', 'Cod. Cliente': 'cod_cliente', 'e-mail': 'e_mail'
    }
    lista_archivos = sorted(glob.glob("Cartera_*.xlsx"))
    if not lista_archivos: return pd.DataFrame()
    lista_df = []
    for archivo in lista_archivos:
        try:
            df = pd.read_excel(archivo)
            if not df.empty: df = df.iloc[:-1]
            for col in ['e-mail', 'Cod. Cliente']:
                if col not in df.columns: df[col] = None
            df['Serie'] = df['Serie'].astype(str)
            df = df[~df['Serie'].str.contains('W|X', case=False, na=False)]
            df.rename(columns=mapa_columnas, inplace=True)
            lista_df.append(df)
        except Exception as e:
            st.warning(f"No se pudo procesar el archivo {archivo}: {e}")
    if not lista_df: return pd.DataFrame()
    df_completo = pd.concat(lista_df, ignore_index=True)
    df_completo.dropna(subset=['numero', 'nombrecliente'], inplace=True)
    df_completo['nomvendedor_norm'] = df_completo['nomvendedor'].apply(normalizar_nombre)
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

@st.cache_data
def cargar_datos_actuales():
    try:
        df = pd.read_excel("Cartera.xlsx")
        if not df.empty: df = df.iloc[:-1]
        df_renamed = df.rename(columns=lambda x: normalizar_nombre(x).lower().replace(' ', '_'))
        if 'serie' in df_renamed.columns:
            df_renamed['serie'] = df_renamed['serie'].astype(str)
            df_filtrado = df_renamed[~df_renamed['serie'].str.contains('W|X', case=False, na=False)]
        else:
            df_filtrado = df_renamed
        for col in ['fecha_documento', 'fecha_vencimiento']:
            if col in df_filtrado.columns:
                df_filtrado[col] = pd.to_datetime(df_filtrado[col], errors='coerce')
        hoy = pd.to_datetime(datetime.now())
        if 'fecha_vencimiento' in df_filtrado.columns:
            df_filtrado['dias_vencido'] = (hoy - df_filtrado['fecha_vencimiento']).dt.days
        return df_filtrado
    except FileNotFoundError:
        return pd.DataFrame()

# --- CÓDIGO PRINCIPAL DE LA PÁGINA ---
df_historico_completo = cargar_datos_historicos()
df_cartera_actual = cargar_datos_actuales()

if df_historico_completo.empty and df_cartera_actual.empty:
    st.error("No se encontraron archivos de datos (ni históricos `Cartera_*.xlsx` ni actuales `Cartera.xlsx`)."); st.stop()

acceso_general = st.session_state.get('acceso_general', False)
vendedor_autenticado = st.session_state.get('vendedor_autenticado', None)
if not acceso_general:
    df_historico_filtrado = df_historico_completo[df_historico_completo['nomvendedor_norm'] == normalizar_nombre(vendedor_autenticado)].copy()
else:
    df_historico_filtrado = df_historico_completo.copy()

lista_clientes = sorted(df_historico_filtrado['nombrecliente'].dropna().unique())
if not lista_clientes:
    st.info("No tienes clientes asignados en el historial de datos."); st.stop()
    
cliente_sel = st.selectbox("Selecciona un cliente para analizar y gestionar su cuenta:", [""] + lista_clientes, format_func=lambda x: "Seleccione un cliente..." if x == "" else x)

if cliente_sel:
    df_cliente_historico = df_historico_filtrado[df_historico_filtrado['nombrecliente'] == cliente_sel].copy()
    
    # --- CÁLCULO DE KPIS Y DATOS PARA GRÁFICOS ---
    # Datos de cartera actual para el cliente
    df_cliente_actual = pd.DataFrame()
    if not df_cartera_actual.empty:
        df_cliente_actual_temp = df_cartera_actual.dropna(subset=['nombrecliente'])
        df_cliente_actual = df_cliente_actual_temp[df_cliente_actual_temp['nombrecliente'] == cliente_sel].copy()

    # Deuda vencida actual
    total_vencido_cliente = 0
    df_vencidas_actual_cliente = pd.DataFrame()
    if not df_cliente_actual.empty:
        df_vencidas_actual_cliente = df_cliente_actual[df_cliente_actual['dias_vencido'] > 0]
        total_vencido_cliente = df_vencidas_actual_cliente['importe'].sum()

    # KPIs de comportamiento histórico
    df_pagadas_reales = df_cliente_historico[(df_cliente_historico['dias_de_pago'].notna()) & (df_cliente_historico['importe'] > 0)]
    avg_dias_pago = df_pagadas_reales['dias_de_pago'].mean() if not df_pagadas_reales.empty else 0
    std_dias_pago = df_pagadas_reales['dias_de_pago'].std() if not df_pagadas_reales.empty else 0
    valor_historico = df_cliente_historico['importe'].sum()
    ultima_compra = df_cliente_historico['fecha_documento'].max()
    
    # Calificación de consistencia
    if std_dias_pago < 7: consistencia = "✅ Muy Consistente"
    elif std_dias_pago < 15: consistencia = "👍 Consistente"
    else: consistencia = "⚠️ Inconsistente"

    # --- MEJORA: Pestañas reorganizadas para un mejor flujo de análisis ---
    tab1, tab2, tab3 = st.tabs(["📊 Resumen y Tendencias", "💳 Cartera Actual", "✉️ Gestión y Comunicación"])

    with tab1:
        st.subheader(f"Diagnóstico General: {cliente_sel}")

        # --- MEJORA: Fila de KPIs extendida ---
        kpi_cols = st.columns(5)
        with kpi_cols[0]:
            st.metric("Días Promedio de Pago", f"{avg_dias_pago:.0f} días" if avg_dias_pago else "N/A")
        with kpi_cols[1]:
            st.metric("Consistencia de Pago", consistencia if avg_dias_pago else "N/A", help="Mide la variabilidad de sus días de pago. Menor es mejor.")
        with kpi_cols[2]:
            st.metric("🔥 Deuda Vencida Actual", f"${total_vencido_cliente:,.0f}")
        with kpi_cols[3]:
            st.metric("Última Compra", ultima_compra.strftime('%d/%m/%Y') if pd.notna(ultima_compra) else "N/A")
        with kpi_cols[4]:
            st.metric("Valor Histórico Total", f"${valor_historico:,.0f}")
        
        # --- MEJORA: Resumen con IA ---
        with st.expander("🤖 **Análisis del Asistente IA**", expanded=True):
            resumen_ia = []
            if avg_dias_pago > 45:
                resumen_ia.append(f"<li>🔴 **Pagador Lento:** El promedio de pago de <b>{avg_dias_pago:.0f} días</b> supera el plazo de 30 días, indicando un hábito de pago demorado.</li>")
            elif avg_dias_pago > 30:
                 resumen_ia.append(f"<li>🟡 **Pagador Oportuno:** Con <b>{avg_dias_pago:.0f} días</b> promedio, el cliente paga ligeramente por encima del plazo, pero de forma aceptable.</li>")
            else:
                 resumen_ia.append(f"<li>🟢 **Pagador Ejemplar:** El cliente paga en un promedio de <b>{avg_dias_pago:.0f} días</b>, demostrando excelente liquidez y compromiso.</li>")

            if consistencia == "⚠️ Inconsistente":
                resumen_ia.append(f"<li>🟡 **Comportamiento Errático:** La alta variabilidad en sus fechas de pago lo convierte en un cliente poco predecible.</li>")
            
            if total_vencido_cliente > 0:
                 resumen_ia.append(f"<li>🔴 **Requiere Acción:** Actualmente posee una deuda vencida de <b>${total_vencido_cliente:,.0f}</b> que debe ser gestionada.</li>")
            
            st.markdown("<ul>" + "".join(resumen_ia) + "</ul>", unsafe_allow_html=True)

        st.markdown("---")
        
        # --- MEJORA: Gráficos de Tendencias ---
        st.subheader("Evolución del Comportamiento del Cliente")
        chart_cols = st.columns(2)
        with chart_cols[0]:
            if not df_pagadas_reales.empty:
                fig_tendencia_pago = px.line(df_pagadas_reales.sort_values('fecha_documento'), 
                                             x='fecha_documento', y='dias_de_pago', 
                                             title="Tendencia de Días de Pago", markers=True,
                                             labels={'fecha_documento': 'Fecha de Factura', 'dias_de_pago': 'Días para Pagar'})
                fig_tendencia_pago.add_trace(go.Scatter(x=df_pagadas_reales['fecha_documento'], y=df_pagadas_reales['dias_de_pago'], mode='lines', line=dict(color='rgba(0,0,0,0)'), name='Tendencia (OLS)'))
                # Línea de tendencia (regresión)
                trendline = px.scatter(df_pagadas_reales, x='fecha_documento', y='dias_de_pago', trendline="ols", trendline_color_override="red").data[1]
                fig_tendencia_pago.add_trace(trendline)
                st.plotly_chart(fig_tendencia_pago, use_container_width=True)
            else:
                st.info("No hay suficientes datos de facturas pagadas para mostrar la tendencia de pago.")

        with chart_cols[1]:
            df_compras = df_cliente_historico[df_cliente_historico['importe']>0].set_index('fecha_documento').resample('QE')['importe'].sum().reset_index()
            if not df_compras.empty:
                fig_volumen_compra = px.bar(df_compras, x='fecha_documento', y='importe',
                                            title="Volumen de Compra por Trimestre",
                                            labels={'fecha_documento': 'Trimestre', 'importe': 'Monto Total Comprado'})
                st.plotly_chart(fig_volumen_compra, use_container_width=True)
            else:
                st.info("No hay datos de compras para mostrar el volumen.")

    with tab2:
        st.subheader(f"Análisis de la Cartera Actual para {cliente_sel}")
        if not df_cliente_actual.empty:
            col1, col2 = st.columns([1,2])
            with col1:
                # --- MEJORA: Gráfico de dona para la deuda actual ---
                st.write("#### Composición de la Deuda")
                if not df_vencidas_actual_cliente.empty:
                    bins = [-float('inf'), 0, 15, 30, 60, float('inf')]
                    labels = ['Al día', '1-15 días', '16-30 días', '31-60 días', 'Más de 60 días']
                    
                    # Usar datos de cartera actual para la composición
                    df_cliente_actual['edad_cartera'] = pd.cut(df_cliente_actual['dias_vencido'], bins=bins, labels=labels, right=True)
                    df_edades = df_cliente_actual.groupby('edad_cartera', observed=True)['importe'].sum().reset_index()

                    fig_dona = px.pie(df_edades, values='importe', names='edad_cartera', 
                                      title='Deuda por Antigüedad', hole=.4,
                                      color_discrete_map={'Al día': 'green', '1-15 días': '#FFD700', '16-30 días': 'orange', '31-60 días': 'darkorange', 'Más de 60 días': 'red'})
                    st.plotly_chart(fig_dona, use_container_width=True)
                else:
                    st.success("¡Excelente! El cliente no tiene deuda vencida.")

            with col2:
                st.write("#### Detalle de Facturas Pendientes")
                st.dataframe(df_cliente_actual[['numero', 'fecha_documento', 'fecha_vencimiento', 'dias_vencido', 'importe']], use_container_width=True)
        else:
            st.info("Este cliente no tiene facturas en la cartera actual (archivo Cartera.xlsx).")

    with tab3:
        # Tu código original para esta pestaña, que ya es excelente.
        st.subheader(f"Herramientas de Comunicación para: {cliente_sel}")
        col1, col2 = st.columns(2)
        with col1:
            st.write("#### 1. Generar Estado de Cuenta en PDF")
            st.download_button(
                label="📄 Descargar Estado de Cuenta Histórico (PDF)",
                data=generar_pdf_estado_cuenta(df_cliente_historico), # PDF con todo el histórico
                file_name=f"Estado_Cuenta_Historico_{normalizar_nombre(cliente_sel).replace(' ', '_')}.pdf",
                mime="application/pdf"
            )
        with col2:
            st.write("#### 2. Preparar Email de Cobro")
            asunto_sugerido = f"Recordatorio de pago y estado de cuenta - Ferreinox SAS BIC"
            cuerpo_mensaje = f"""Estimados Sres. de {cliente_sel},

Le saludamos cordialmente desde Ferreinox SAS BIC.

Nos ponemos en contacto con usted para recordarle amablemente sobre su saldo pendiente. Actualmente, sus facturas vencidas suman un total de **${total_vencido_cliente:,.0f}**.

Adjuntamos a este correo su estado de cuenta detallado para su revisión.

Puede realizar su pago de forma fácil y segura a través de nuestro portal en línea:
https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/

Para ingresar, por favor utilice su NIT como 'usuario' y su Código de Cliente como 'código único interno'.

Agradecemos de antemano su pronta atención a este asunto. Si ya ha realizado el pago, por favor haga caso omiso de este mensaje.

Atentamente,
Equipo de Cartera
Ferreinox SAS BIC"""
            st.text_input("Asunto del Correo:", value=asunto_sugerido)
            st.text_area("Cuerpo del Correo (listo para copiar):", value=cuerpo_mensaje, height=350)
