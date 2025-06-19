# ======================================================================================
# ARCHIVO: pages/🧑‍💼_Perfil_de_Cliente.py (Versión Final Definitiva)
# ======================================================================================
import streamlit as st
import pandas as pd
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

# --- FUNCIONES AUXILIARES ---
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
    
    # Recalcular días vencido para el PDF basado en la fecha actual
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
        # Usar el valor recalculado para el color
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

st.title("🧑‍💼 Perfil de Pagador por Cliente")

@st.cache_data
def cargar_datos_historicos():
    mapa_columnas = {
        'Serie': 'serie', 'Número': 'numero', 'Fecha Documento': 'fecha_documento',
        'Fecha Vencimiento': 'fecha_vencimiento', 'Fecha Saldado': 'fecha_saldado',
        'NOMBRECLIENTE': 'nombrecliente', 'Población': 'poblacion', 'Provincia': 'provincia',
        'IMPORTE': 'importe', 'RIESGOCONCEDIDO': 'riesgoconcedido', 'NOMVENDEDOR': 'nomvendedor',
        'DIAS_VENCIDO': 'dias_vencido', 'Estado': 'estado', 'Cod. Cliente': 'cod_cliente',
        'e-mail': 'e_mail'
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

# --- CÓDIGO PRINCIPAL DE LA PÁGINA ---
df_historico_completo = cargar_datos_historicos()
if df_historico_completo.empty:
    st.warning("No se encontraron archivos de datos históricos."); st.stop()

acceso_general = st.session_state.get('acceso_general', False)
vendedor_autenticado = st.session_state.get('vendedor_autenticado', None)
if not acceso_general:
    df_historico_filtrado = df_historico_completo[df_historico_completo['nomvendedor_norm'] == normalizar_nombre(vendedor_autenticado)].copy()
else:
    df_historico_filtrado = df_historico_completo.copy()
lista_clientes = sorted(df_historico_filtrado['nombrecliente'].dropna().unique())
if not lista_clientes:
    st.info("No tienes clientes asignados en el historial de datos."); st.stop()
    
cliente_sel = st.selectbox("Selecciona un cliente para analizar y gestionar su cuenta:", [""] + lista_clientes)
if cliente_sel:
    df_cliente = df_historico_filtrado[df_historico_filtrado['nombrecliente'] == cliente_sel].copy()
    tab1, tab2 = st.tabs(["📊 Análisis del Cliente", "✉️ Gestión y Comunicación"])
    with tab1:
        st.subheader(f"Análisis de Comportamiento: {cliente_sel}")
        df_pagadas_reales = df_cliente[(df_cliente['dias_de_pago'].notna()) & (df_cliente['importe'] > 0)]
        if not df_pagadas_reales.empty:
            avg_dias_pago = df_pagadas_reales['dias_de_pago'].mean()
            if avg_dias_pago <= 30: calificacion = "✅ Pagador Excelente"
            elif avg_dias_pago <= 60: calificacion = "👍 Pagador Bueno"
            elif avg_dias_pago <= 90: calificacion = "⚠️ Pagador Lento"
            else: calificacion = "🚨 Pagador de Riesgo"
            col1, col2 = st.columns(2)
            with col1: st.metric("Días Promedio de Pago (Ventas)", f"{avg_dias_pago:.0f} días", help="Promedio de días que tarda el cliente en pagar las facturas de VENTA.")
            with col2: st.metric("Calificación", calificacion)
        else:
            st.info("Este cliente no tiene un historial de facturas de VENTA pagadas para calcular su comportamiento.")
        st.subheader("Historial Completo de Transacciones")
        st.dataframe(df_cliente[['numero', 'fecha_documento', 'fecha_vencimiento', 'fecha_saldado', 'dias_de_pago', 'importe']].sort_values(by="fecha_documento", ascending=False))
    with tab2:
        st.subheader(f"Herramientas de Comunicación para: {cliente_sel}")
        
        # --- LÓGICA CORREGIDA PARA CALCULAR DEUDA VENCIDA ACTUAL ---
        hoy = pd.to_datetime(datetime.now())
        df_abiertas_cliente = df_cliente[df_cliente['fecha_saldado'].isnull()]
        if not df_abiertas_cliente.empty:
            df_vencidas_hoy = df_abiertas_cliente[df_abiertas_cliente['fecha_vencimiento'] < hoy].copy()
            total_vencido_cliente = df_vencidas_hoy['importe'].sum()
        else:
            total_vencido_cliente = 0
            
        col1, col2 = st.columns(2)
        with col1:
            st.write("#### 1. Generar Estado de Cuenta en PDF")
            st.download_button(
                label="📄 Descargar Estado de Cuenta (PDF)",
                data=generar_pdf_estado_cuenta(df_cliente),
                file_name=f"Estado_Cuenta_{normalizar_nombre(cliente_sel).replace(' ', '_')}.pdf",
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
Ferreinox SAS BIC
"""
            st.text_input("Asunto del Correo:", value=asunto_sugerido)
            st.text_area("Cuerpo del Correo (listo para copiar):", value=cuerpo_mensaje, height=350)
