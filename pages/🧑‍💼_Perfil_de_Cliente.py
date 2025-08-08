# ======================================================================================
# ARCHIVO: pages/🧑‍💼_Perfil_de_Cliente.py (Versión Definitiva 2.0)
# ======================================================================================
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import glob
import re
import unicodedata
from datetime import datetime, timedelta
from fpdf import FPDF
from io import BytesIO
from urllib.parse import quote
import yagmail
import os
import tempfile

# --- CONFIGURACIÓN DE PÁGINA Y ESTILOS ---
st.set_page_config(page_title="Perfil de Cliente", page_icon="🧑‍💼", layout="wide")

st.markdown("""
<style>
    .stMetric {
        background-color: #FFFFFF;
        border-radius: 10px;
        padding: 15px;
        border: 1px solid #CCCCCC;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: transparent;
        border-radius: 4px 4px 0px 0px;
        border-bottom: 2px solid #C0C0C0;
    }
    .stTabs [aria-selected="true"] {
        border-bottom: 2px solid #003865;
        color: #003865;
        font-weight: bold;
    }
    .button {
        display: inline-block;
        padding: 10px 20px;
        color: white;
        background-color: #25D366; /* Verde WhatsApp */
        border-radius: 5px;
        text-align: center;
        text-decoration: none;
        font-weight: bold;
        width: 95%; /* Ocupa casi todo el ancho de la columna */
        box-sizing: border-box; /* Asegura que el padding no desborde */
    }
    .stContainer {
        border: 1px solid #DDDDDD;
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)


# --- GUARDIA DE SEGURIDAD ---
if 'authentication_status' not in st.session_state or not st.session_state['authentication_status']:
    st.warning("Por favor, inicie sesión en el 📈 Tablero Principal para acceder a esta página.")
    st.stop()

# ======================================================================================
# --- CLASE PDF Y FUNCIONES DE GENERACIÓN (POTENCIADAS) ---
# ======================================================================================

# Clase PDF tomada del Tablero Principal para consistencia en el branding
class PDF(FPDF):
    def __init__(self, orientation='P', unit='mm', format='A4', title='Estado de Cuenta'):
        super().__init__(orientation, unit, format)
        self.title_doc = title

    def header(self):
        try:
            self.image("LOGO FERREINOX SAS BIC 2024.png", 10, 8, 80)
        except FileNotFoundError:
            self.set_font('Arial', 'B', 12)
            self.cell(80, 10, 'Logo no encontrado', 0, 0, 'L')
        self.set_font('Arial', 'B', 18)
        self.cell(0, 10, self.title_doc, 0, 1, 'R')
        self.set_font('Arial', 'I', 9)
        self.cell(0, 10, f'Generado el: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', 0, 1, 'R')
        self.ln(5)
        self.set_line_width(0.5)
        self.set_draw_color(220, 220, 220)
        self.line(10, 35, 200, 35)
        self.ln(10)

    def footer(self):
        self.set_y(-40)
        self.set_font('Arial', 'I', 9)
        self.set_text_color(100, 100, 100)
        self.cell(0, 6, "Para ingresar al portal de pagos, utiliza el NIT como 'usuario' y el Codigo de Cliente como 'codigo unico interno'.", 0, 1, 'C')
        self.set_font('Arial', 'B', 11)
        self.set_text_color(0, 0, 0)
        self.cell(0, 8, 'Realiza tu pago de forma facil y segura aqui:', 0, 1, 'C')
        self.set_font('Arial', 'BU', 12)
        self.set_text_color(4, 88, 167)
        link = "https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/"
        self.cell(0, 10, "Portal de Pagos Ferreinox SAS BIC", 0, 1, 'C', link=link)

# --- NUEVA FUNCIÓN: PDF para Cartera Actual (Tab 2) ---
def generar_pdf_cartera_actual(datos_cliente: pd.DataFrame):
    if datos_cliente.empty:
        return None
    pdf = PDF(title='Cartera Pendiente')
    pdf.set_auto_page_break(auto=True, margin=45)
    pdf.add_page()
    info_cliente = datos_cliente.iloc[0]

    pdf.set_font('Arial', 'B', 11); pdf.cell(40, 10, 'Cliente:', 0, 0); pdf.set_font('Arial', '', 11); pdf.cell(0, 10, info_cliente.get('nombrecliente', ''), 0, 1)
    cod_cliente_val = info_cliente.get('cod_cliente')
    cod_cliente_str = str(int(cod_cliente_val)) if pd.notna(cod_cliente_val) else "N/A"
    pdf.set_font('Arial', 'B', 11); pdf.cell(40, 10, 'Codigo de Cliente:', 0, 0); pdf.set_font('Arial', '', 11)
    pdf.cell(0, 10, cod_cliente_str, 0, 1); pdf.ln(5)

    pdf.set_font('Arial', '', 10)
    mensaje = "A continuacion se presenta el detalle de todas las facturas pendientes de pago a la fecha actual."
    pdf.set_text_color(128, 128, 128); pdf.multi_cell(0, 5, mensaje, 0, 'J'); pdf.set_text_color(0, 0, 0); pdf.ln(10)

    pdf.set_font('Arial', 'B', 10); pdf.set_fill_color(0, 56, 101); pdf.set_text_color(255, 255, 255)
    pdf.cell(25, 10, 'Factura', 1, 0, 'C', 1); pdf.cell(35, 10, 'Fecha Factura', 1, 0, 'C', 1)
    pdf.cell(35, 10, 'Fecha Venc.', 1, 0, 'C', 1); pdf.cell(30, 10, 'Dias Vencido', 1, 0, 'C', 1); pdf.cell(35, 10, 'Importe', 1, 1, 'C', 1)

    pdf.set_font('Arial', '', 10)
    total_importe = 0
    total_vencido = 0
    df_ordenado = datos_cliente.sort_values(by='fecha_vencimiento', ascending=True)

    for _, row in df_ordenado.iterrows():
        pdf.set_text_color(0, 0, 0)
        dias_vencido = row.get('dias_vencido', 0)
        if dias_vencido > 0:
            pdf.set_fill_color(248, 241, 241) # Rojo claro para vencidas
            total_vencido += row.get('importe', 0)
        else:
            pdf.set_fill_color(230, 245, 230) # Verde claro para al día
        total_importe += row.get('importe', 0)
        
        pdf.cell(25, 10, str(int(row.get('numero', 0))), 1, 0, 'C', 1)
        pdf.cell(35, 10, pd.to_datetime(row.get('fecha_documento')).strftime('%d/%m/%Y'), 1, 0, 'C', 1)
        pdf.cell(35, 10, pd.to_datetime(row.get('fecha_vencimiento')).strftime('%d/%m/%Y'), 1, 0, 'C', 1)
        pdf.cell(30, 10, str(int(dias_vencido)), 1, 0, 'C', 1)
        pdf.cell(35, 10, f"${row.get('importe', 0):,.0f}", 1, 1, 'R', 1)

    pdf.ln(10)
    pdf.set_font('Arial', 'B', 10); pdf.set_fill_color(0, 56, 101); pdf.set_text_color(255, 255, 255)
    pdf.cell(125, 10, 'TOTAL VENCIDO', 1, 0, 'R', 1); pdf.cell(35, 10, f"${total_vencido:,.0f}", 1, 1, 'R', 1)
    pdf.cell(125, 10, 'TOTAL CARTERA PENDIENTE', 1, 0, 'R', 1); pdf.cell(35, 10, f"${total_importe:,.0f}", 1, 1, 'R', 1)

    return bytes(pdf.output())

# --- MEJORADA FUNCIÓN: PDF Histórico con filtro y movimientos ---
def generar_pdf_historico_filtrado(df_cliente_historico, fecha_inicio, fecha_fin):
    pdf = PDF(title='Extracto de Movimientos')
    pdf.set_auto_page_break(auto=True, margin=45)
    pdf.add_page()
    info_cliente = df_cliente_historico.iloc[0]

    # Info Cliente
    pdf.set_font('Arial', 'B', 11); pdf.cell(40, 8, 'Cliente:', 0, 0); pdf.set_font('Arial', '', 11); pdf.cell(0, 8, info_cliente.get('nombrecliente', ''), 0, 1)
    cod_cliente_val = info_cliente.get('cod_cliente')
    cod_cliente_str = str(int(cod_cliente_val)) if pd.notna(cod_cliente_val) else "N/A"
    pdf.set_font('Arial', 'B', 11); pdf.cell(40, 8, 'Codigo Cliente:', 0, 0); pdf.set_font('Arial', '', 11); pdf.cell(0, 8, cod_cliente_str, 0, 1)
    pdf.set_font('Arial', 'B', 11); pdf.cell(40, 8, 'Periodo:', 0, 0); pdf.set_font('Arial', '', 11); pdf.cell(0, 8, f"{fecha_inicio.strftime('%d/%m/%Y')} al {fecha_fin.strftime('%d/%m/%Y')}", 0, 1)
    pdf.ln(5)

    # Preparar datos de movimientos
    df_facturas = df_cliente_historico[['fecha_documento', 'numero', 'importe']].copy()
    df_facturas.rename(columns={'fecha_documento': 'fecha', 'numero': 'documento', 'importe': 'debito'}, inplace=True)
    df_facturas['credito'] = 0
    df_facturas['tipo'] = 'Factura'

    df_pagos = df_cliente_historico[df_cliente_historico['fecha_saldado'].notna()][['fecha_saldado', 'numero', 'importe']].copy()
    df_pagos.rename(columns={'fecha_saldado': 'fecha', 'numero': 'documento', 'importe': 'credito'}, inplace=True)
    df_pagos['debito'] = 0
    df_pagos['tipo'] = 'Pago'

    df_movimientos = pd.concat([df_facturas, df_pagos]).sort_values(by='fecha').reset_index(drop=True)
    
    # Filtrar por rango de fechas
    saldo_anterior = df_movimientos[df_movimientos['fecha'] < fecha_inicio]['debito'].sum() - df_movimientos[df_movimientos['fecha'] < fecha_inicio]['credito'].sum()
    df_movimientos_periodo = df_movimientos[(df_movimientos['fecha'] >= fecha_inicio) & (df_movimientos['fecha'] <= fecha_fin)].copy()

    # Calcular saldo corrido
    df_movimientos_periodo['saldo'] = saldo_anterior + df_movimientos_periodo['debito'].cumsum() - df_movimientos_periodo['credito'].cumsum()

    # Tabla de Movimientos
    pdf.set_font('Arial', 'B', 9); pdf.set_fill_color(0, 56, 101); pdf.set_text_color(255, 255, 255)
    pdf.cell(25, 10, 'Fecha', 1, 0, 'C', 1); pdf.cell(20, 10, 'Tipo', 1, 0, 'C', 1)
    pdf.cell(30, 10, 'Documento N°', 1, 0, 'C', 1); pdf.cell(35, 10, 'Debito (+)', 1, 0, 'C', 1)
    pdf.cell(35, 10, 'Credito (-)', 1, 0, 'C', 1); pdf.cell(40, 10, 'Saldo', 1, 1, 'C', 1)
    
    pdf.set_font('Arial', '', 9); pdf.set_fill_color(230, 230, 230); pdf.set_text_color(0,0,0)
    pdf.cell(105, 8, 'Saldo Anterior al Periodo', 1, 0, 'R', 1)
    pdf.cell(35, 8, '', 1, 0, 'R', 1)
    pdf.cell(40, 8, f"${saldo_anterior:,.0f}", 1, 1, 'R', 1)

    pdf.set_font('Arial', '', 9)
    for _, row in df_movimientos_periodo.iterrows():
        pdf.set_text_color(0,0,0)
        pdf.set_fill_color(255, 255, 255)
        pdf.cell(25, 8, row['fecha'].strftime('%d/%m/%Y'), 1, 0, 'C', 1)
        pdf.cell(20, 8, row['tipo'], 1, 0, 'C', 1)
        pdf.cell(30, 8, str(int(row['documento'])), 1, 0, 'C', 1)
        pdf.cell(35, 8, f"${row['debito']:,.0f}" if row['debito'] > 0 else '-', 1, 0, 'R', 1)
        pdf.cell(35, 8, f"${row['credito']:,.0f}" if row['credito'] > 0 else '-', 1, 0, 'R', 1)
        pdf.cell(40, 8, f"${row['saldo']:,.0f}", 1, 1, 'R', 1)

    # Resumen Final
    total_debitos_periodo = df_movimientos_periodo['debito'].sum()
    total_creditos_periodo = df_movimientos_periodo['credito'].sum()
    saldo_final = saldo_anterior + total_debitos_periodo - total_creditos_periodo

    pdf.ln(5)
    pdf.set_font('Arial', 'B', 10); pdf.set_fill_color(0, 56, 101); pdf.set_text_color(255, 255, 255)
    pdf.cell(145, 10, 'TOTAL FACTURADO (PERIODO)', 1, 0, 'R', 1); pdf.cell(40, 10, f"${total_debitos_periodo:,.0f}", 1, 1, 'R', 1)
    pdf.cell(145, 10, 'TOTAL PAGADO (PERIODO)', 1, 0, 'R', 1); pdf.cell(40, 10, f"${total_creditos_periodo:,.0f}", 1, 1, 'R', 1)
    pdf.set_font('Arial', 'B', 11)
    pdf.cell(145, 10, 'SALDO FINAL A LA FECHA', 1, 0, 'R', 1); pdf.cell(40, 10, f"${saldo_final:,.0f}", 1, 1, 'R', 1)

    return bytes(pdf.output())


def normalizar_nombre(nombre: str) -> str:
    if not isinstance(nombre, str): return ""
    nombre = nombre.upper().strip().replace('.', '')
    nombre = ''.join(c for c in unicodedata.normalize('NFD', nombre) if unicodedata.category(c) != 'Mn')
    return ' '.join(nombre.split())

# --- FUNCIONES DE CARGA DE DATOS (CONSOLIDADO) ---
@st.cache_data(ttl=600)
def cargar_datos_consolidados():
    # Carga de Datos Históricos (Excel)
    mapa_columnas = {
        'Serie': 'serie', 'Número': 'numero', 'Fecha Documento': 'fecha_documento',
        'Fecha Vencimiento': 'fecha_vencimiento', 'Fecha Saldado': 'fecha_saldado',
        'NOMBRECLIENTE': 'nombrecliente', 'Población': 'poblacion', 'Provincia': 'provincia',
        'IMPORTE': 'importe', 'RIESGOCONCEDIDO': 'riesgoconcedido', 'NOMVENDEDOR': 'nomvendedor',
        'DIAS_VENCIDO': 'dias_vencido', 'Estado': 'estado', 'Cod. Cliente': 'cod_cliente',
        'e-mail': 'e_mail', 'Nit': 'nit', 'Telefono1': 'telefono1'
    }
    lista_archivos = sorted(glob.glob("Cartera_*.xlsx"))
    df_historico = pd.DataFrame()
    if lista_archivos:
        lista_df = []
        for archivo in lista_archivos:
            try:
                df = pd.read_excel(archivo)
                if not df.empty and "Total" in str(df.iloc[-1, 0]):
                    df = df.iloc[:-1]
                lista_df.append(df)
            except Exception as e:
                st.warning(f"No se pudo procesar {archivo}: {e}")
        if lista_df:
            df_historico = pd.concat(lista_df, ignore_index=True)
            df_historico.rename(columns=mapa_columnas, inplace=True)

    # Carga de Datos Actuales (Excel 'Cartera.xlsx')
    df_cartera_actual = pd.DataFrame()
    try:
        df_actual = pd.read_excel("Cartera.xlsx")
        if not df_actual.empty and "Total" in str(df_actual.iloc[-1, 0]):
            df_actual = df_actual.iloc[:-1]
        
        # Renombramos columnas de forma segura
        df_actual.columns = [mapa_columnas.get(col, col.lower().replace(' ', '_')) for col in df_actual.columns]
        df_cartera_actual = df_actual

    except FileNotFoundError:
        pass # No es un error si no existe, los históricos son la base

    # Combinar y procesar
    if df_historico.empty and df_cartera_actual.empty:
        return pd.DataFrame(), pd.DataFrame()

    df_completo = pd.concat([df_historico, df_cartera_actual], ignore_index=True)
    df_completo = df_completo.loc[:,~df_completo.columns.duplicated()] # Prevenir errores de columnas duplicadas
    df_completo.dropna(subset=['numero', 'nombrecliente'], inplace=True)
    
    # Normalización y limpieza
    for col in ['fecha_documento', 'fecha_vencimiento', 'fecha_saldado']:
        if col in df_completo.columns:
            df_completo[col] = pd.to_datetime(df_completo[col], errors='coerce')

    df_completo['importe'] = pd.to_numeric(df_completo['importe'], errors='coerce').fillna(0)
    df_completo['nomvendedor_norm'] = df_completo['nomvendedor'].apply(normalizar_nombre)
    df_completo.sort_values(by=['fecha_documento', 'fecha_saldado'], ascending=[True, True], na_position='first', inplace=True)
    df_historico_unico = df_completo.drop_duplicates(subset=['serie', 'numero'], keep='last').copy()
    
    # Calcular dias de pago
    df_pagadas = df_historico_unico.dropna(subset=['fecha_saldado', 'fecha_documento']).copy()
    if not df_pagadas.empty:
        df_pagadas['dias_de_pago'] = (df_pagadas['fecha_saldado'] - df_pagadas['fecha_documento']).dt.days
        df_historico_unico = pd.merge(df_historico_unico, df_pagadas[['serie', 'numero', 'dias_de_pago']], on=['serie', 'numero'], how='left')

    # Calcular dias vencido para cartera actual
    hoy = pd.to_datetime(datetime.now())
    if 'fecha_vencimiento' in df_cartera_actual.columns:
        df_cartera_actual['dias_vencido'] = (hoy - df_cartera_actual['fecha_vencimiento']).dt.days.fillna(0)
    
    return df_historico_unico, df_cartera_actual


# ======================================================================================
# --- CÓDIGO PRINCIPAL DE LA PÁGINA ---
# ======================================================================================

st.title("🧑‍💼 Dossier de Inteligencia de Cliente")

df_historico_completo, df_cartera_actual = cargar_datos_consolidados()

if df_historico_completo.empty:
    st.error("No se encontraron archivos de datos históricos (`Cartera_*.xlsx`). La aplicación no puede continuar."); st.stop()

# Filtrar por vendedor autenticado
acceso_general = st.session_state.get('acceso_general', False)
vendedor_autenticado = st.session_state.get('vendedor_autenticado', None)
if not acceso_general:
    df_historico_filtrado_vendedor = df_historico_completo[df_historico_completo['nomvendedor_norm'] == normalizar_nombre(vendedor_autenticado)].copy()
else:
    df_historico_filtrado_vendedor = df_historico_completo.copy()

lista_clientes = sorted(df_historico_filtrado_vendedor['nombrecliente'].dropna().unique())
if not lista_clientes:
    st.info("No tienes clientes asignados en el historial de datos."); st.stop()
    
cliente_sel = st.selectbox("Selecciona un cliente para analizar y gestionar su cuenta:", [""] + lista_clientes, format_func=lambda x: "Seleccione un cliente..." if x == "" else x)

if cliente_sel:
    # --- FILTRADO DE DATOS PARA CLIENTE SELECCIONADO ---
    df_cliente_historico = df_historico_filtrado_vendedor[df_historico_filtrado_vendedor['nombrecliente'] == cliente_sel].copy()
    df_cliente_actual = pd.DataFrame()
    if not df_cartera_actual.empty:
        df_cliente_actual_temp = df_cartera_actual.dropna(subset=['nombrecliente'])
        # Asegurar que el nombre del cliente coincida para la cartera actual
        df_cliente_actual = df_cliente_actual_temp[df_cliente_actual_temp['nombrecliente'] == cliente_sel].copy()

    # --- CÁLCULO DE KPIS EXTENDIDO ---
    # Históricos
    df_pagadas_reales = df_cliente_historico[(df_cliente_historico['dias_de_pago'].notna()) & (df_cliente_historico['importe'] > 0)]
    avg_dias_pago = df_pagadas_reales['dias_de_pago'].mean() if not df_pagadas_reales.empty else 0
    std_dias_pago = df_pagadas_reales['dias_de_pago'].std() if not df_pagadas_reales.empty else 0
    valor_historico = df_cliente_historico['importe'].sum()
    ultima_compra = df_cliente_historico['fecha_documento'].max()
    ultimo_pago = df_cliente_historico['fecha_saldado'].max()
    
    compras_ordenadas = df_cliente_historico.sort_values('fecha_documento')
    frecuencia_compra = compras_ordenadas['fecha_documento'].diff().mean().days if len(compras_ordenadas) > 1 else 0

    # Actuales
    deuda_total_actual = df_cliente_actual['importe'].sum() if not df_cliente_actual.empty else 0
    df_vencidas_actual_cliente = df_cliente_actual[df_cliente_actual['dias_vencido'] > 0] if not df_cliente_actual.empty else pd.DataFrame()
    deuda_vencida_actual = df_vencidas_actual_cliente['importe'].sum() if not df_vencidas_actual_cliente.empty else 0
    max_dias_retraso = df_vencidas_actual_cliente['dias_vencido'].max() if not df_vencidas_actual_cliente.empty else 0
    
    # Consistencia
    if std_dias_pago < 7: consistencia = "✅ Muy Consistente"
    elif std_dias_pago < 15: consistencia = "👍 Consistente"
    else: consistencia = "⚠️ Inconsistente"

    # Datos para comunicación
    info_reciente = df_cliente_historico.iloc[-1]
    nit_cliente = str(info_reciente.get('nit', 'N/A')).split('.')[0]
    cod_cliente = str(int(info_reciente.get('cod_cliente'))) if pd.notna(info_reciente.get('cod_cliente')) else "N/A"
    email_cliente = info_reciente.get('e_mail', 'Correo no disponible')
    telefono_raw = str(info_reciente.get('telefono1', ''))
    telefono_cliente = telefono_raw.split('.')[0] if '.' in telefono_raw else telefono_raw


    # ======================================================================================
    # --- INTERFAZ DE PESTAÑAS ---
    # ======================================================================================
    tab1, tab2, tab3 = st.tabs(["📊 **Resumen y Tendencias**", "💳 **Cartera Actual Detallada**", "✉️ **Gestión, Comunicación y Extractos**"])

    with tab1:
        st.subheader(f"Diagnóstico General: {cliente_sel}")

        # --- MEJORA: KPIs en dos filas ---
        st.markdown("##### Comportamiento Histórico de Pago")
        kpi_cols1 = st.columns(4)
        kpi_cols1[0].metric("Días Promedio de Pago", f"{avg_dias_pago:.0f} días" if avg_dias_pago else "N/A")
        kpi_cols1[1].metric("Consistencia de Pago", consistencia, help="Mide la variabilidad (desviación estándar) de sus días de pago. Menor es mejor.")
        kpi_cols1[2].metric("Frecuencia de Compra", f"Cada {frecuencia_compra:.0f} días" if frecuencia_compra > 0 else "N/A", help="Tiempo promedio entre compras.")
        kpi_cols1[3].metric("Último Pago Registrado", ultimo_pago.strftime('%d/%m/%Y') if pd.notna(ultimo_pago) else "N/A")

        st.markdown("##### Situación Actual y Valor del Cliente")
        kpi_cols2 = st.columns(4)
        kpi_cols2[0].metric("🔥 Deuda Vencida Actual", f"${deuda_vencida_actual:,.0f}")
        kpi_cols2[1].metric("💰 Deuda Total Actual", f"${deuda_total_actual:,.0f}")
        kpi_cols2[2].metric("📅 Última Compra", ultima_compra.strftime('%d/%m/%Y') if pd.notna(ultima_compra) else "N/A")
        kpi_cols2[3].metric("🏆 Valor Histórico Total", f"${valor_historico:,.0f}")
        
        # --- MEJORA: Resumen con IA ---
        with st.expander("🤖 **Análisis del Asistente IA**", expanded=True):
            resumen_ia = []
            if avg_dias_pago == 0:
                resumen_ia.append(f"<li>🔵 **Cliente Nuevo:** No hay historial de pagos registrado para analizar un comportamiento.</li>")
            elif avg_dias_pago > 45:
                resumen_ia.append(f"<li>🔴 **Pagador Lento:** Paga en promedio a <b>{avg_dias_pago:.0f} días</b>, un indicador de riesgo.</li>")
            elif avg_dias_pago > 30:
                resumen_ia.append(f"<li>🟡 **Pagador Oportuno:** Paga a <b>{avg_dias_pago:.0f} días</b>, ligeramente por encima del plazo pero aceptable.</li>")
            else:
                resumen_ia.append(f"<li>🟢 **Pagador Ejemplar:** Paga en un promedio de <b>{avg_dias_pago:.0f} días</b>, demostrando excelente compromiso.</li>")

            if consistencia == "⚠️ Inconsistente":
                resumen_ia.append(f"<li>🟡 **Comportamiento Errático:** La alta variabilidad en sus fechas de pago (desv. est. de {std_dias_pago:.0f} días) lo hace poco predecible.</li>")
            
            if deuda_vencida_actual > 0:
                resumen_ia.append(f"<li>🔴 **Requiere Acción Inmediata:** Actualmente tiene una deuda vencida de <b>${deuda_vencida_actual:,.0f}</b> con facturas de hasta <b>{max_dias_retraso:.0f} días</b> de retraso.</li>")
            else:
                 resumen_ia.append(f"<li>🟢 **Cartera Sana:** El cliente se encuentra al día con sus pagos. ¡Excelente!</li>")

            if pd.notna(ultima_compra) and (datetime.now() - ultima_compra).days > 90:
                 resumen_ia.append(f"<li>🔵 **Cliente Inactivo:** Han pasado más de 90 días desde su última compra. Posible oportunidad de reactivación.</li>")

            st.markdown("<ul>" + "".join(resumen_ia) + "</ul>", unsafe_allow_html=True)

        st.markdown("---")
        
        # --- Gráficos de Tendencias ---
        st.subheader("Evolución del Comportamiento del Cliente")
        chart_cols = st.columns(2)
        with chart_cols[0]:
            if not df_pagadas_reales.empty:
                fig_tendencia_pago = px.line(df_pagadas_reales.sort_values('fecha_documento'), 
                                             x='fecha_documento', y='dias_de_pago', 
                                             title="Tendencia de Días de Pago", markers=True,
                                             labels={'fecha_documento': 'Fecha de Factura', 'dias_de_pago': 'Días para Pagar'})
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
            col1, col2 = st.columns([1, 2])
            with col1:
                st.write("#### Composición de la Deuda")
                bins = [-float('inf'), 0, 15, 30, 60, float('inf')]
                labels = ['Al día', '1-15 días', '16-30 días', '31-60 días', 'Más de 60 días']
                df_cliente_actual['edad_cartera'] = pd.cut(df_cliente_actual['dias_vencido'], bins=bins, labels=labels, right=True)
                df_edades = df_cliente_actual.groupby('edad_cartera', observed=True)['importe'].sum().reset_index()

                if not df_edades.empty:
                    fig_dona = px.pie(df_edades, values='importe', names='edad_cartera', 
                                      title='Deuda por Antigüedad', hole=.4,
                                      color_discrete_map={'Al día': '#388E3C', '1-15 días': '#FBC02D', '16-30 días': '#F57C00', '31-60 días': 'darkorange', 'Más de 60 días': '#D32F2F'})
                    st.plotly_chart(fig_dona, use_container_width=True)
                else:
                    st.success("¡Excelente! El cliente no tiene deuda pendiente.")
            
            with col2:
                st.write("#### Detalle de Facturas Pendientes")
                st.dataframe(df_cliente_actual[['numero', 'fecha_documento', 'fecha_vencimiento', 'dias_vencido', 'importe']], height=300, use_container_width=True)
                
                pdf_cartera_bytes = generar_pdf_cartera_actual(df_cliente_actual)
                st.download_button(
                    label="📄 Descargar Detalle de Cartera Pendiente (PDF)",
                    data=pdf_cartera_bytes,
                    file_name=f"Cartera_Pendiente_{normalizar_nombre(cliente_sel).replace(' ', '_')}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
        else:
            st.success("✅ ¡Felicidades! Este cliente no tiene facturas pendientes en la cartera actual.")

    with tab3:
        st.subheader(f"Herramientas de Gestión para: {cliente_sel}")
        
        with st.container(border=True):
            st.markdown("#### 📜 Generar Extracto Histórico de Movimientos")
            
            col_fecha1, col_fecha2, col_btn = st.columns([1,1,1])
            with col_fecha1:
                fecha_inicio = st.date_input("Fecha de Inicio", value=datetime.now() - timedelta(days=365))
            with col_fecha2:
                fecha_fin = st.date_input("Fecha de Fin", value=datetime.now())
            
            if fecha_inicio > fecha_fin:
                st.error("Error: La fecha de inicio no puede ser posterior a la fecha de fin.")
            else:
                pdf_historico_bytes = generar_pdf_historico_filtrado(df_cliente_historico, pd.to_datetime(fecha_inicio), pd.to_datetime(fecha_fin))
                with col_btn:
                    st.download_button(
                        label="📄 Descargar Extracto (PDF)",
                        data=pdf_historico_bytes,
                        file_name=f"Extracto_Historico_{normalizar_nombre(cliente_sel).replace(' ', '_')}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        key="download_historico"
                    )

        st.markdown("---")
        
        with st.container(border=True):
            st.markdown("#### 📞 Herramientas de Comunicación")
            col_email, col_whatsapp = st.columns(2)

            with col_email:
                st.subheader("✉️ Enviar por Correo Electrónico")
                email_destino = st.text_input("Verificar o modificar correo:", value=email_cliente)
                
                if st.button("📧 Enviar Correo con Estado de Cuenta", use_container_width=True):
                    if not email_destino or email_destino == 'Correo no disponible' or '@' not in email_destino:
                        st.error("Dirección de correo no válida o no disponible.")
                    else:
                        try:
                            sender_email = st.secrets["email_credentials"]["sender_email"]
                            sender_password = st.secrets["email_credentials"]["sender_password"]
                            
                            portal_link = "https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/"
                            instrucciones = f"<b>Instrucciones de acceso:</b><br> &nbsp; • <b>Usuario:</b> {nit_cliente} (Tu NIT)<br> &nbsp; • <b>Código Único:</b> {cod_cliente}"
                            
                            pdf_para_adjuntar = generar_pdf_cartera_actual(df_cliente_actual) if not df_cliente_actual.empty else generar_pdf_historico_filtrado(df_cliente_historico, pd.to_datetime(ultima_compra) - timedelta(days=30), datetime.now()) if pd.notna(ultima_compra) else b''

                            if deuda_vencida_actual > 0:
                                asunto = f"Recordatorio de Saldo Pendiente - {cliente_sel}"
                                cuerpo_html = f"""
                                <html><body>
                                <p>Estimado(a) {cliente_sel},</p>
                                <p>Recibe un cordial saludo de Ferreinox SAS BIC.</p>
                                <p>Nos ponemos en contacto para recordarle su saldo pendiente. Actualmente, sus facturas vencidas suman un total de <b>${deuda_vencida_actual:,.0f}</b>, y su factura más antigua tiene <b>{max_dias_retraso:.0f} días</b> de vencida.</p>
                                <p>Adjunto a este correo, encontrará su estado de cuenta detallado para su revisión.</p>
                                <p>Para su comodidad, puede realizar el pago a través de nuestro <a href='{portal_link}'><b>Portal de Pagos en Línea</b></a>.</p>
                                <p>{instrucciones}</p>
                                <p>Si ya ha realizado el pago, por favor, haz caso omiso de este recordatorio. Atentamente,<br><b>Área de Cartera Ferreinox SAS BIC</b></p>
                                </body></html>"""
                            else:
                                asunto = f"Tu Estado de Cuenta al día - {cliente_sel}"
                                cuerpo_html = f"""
                                <html><body>
                                <p>Estimado(a) {cliente_sel},</p>
                                <p>Recibe un cordial saludo de Ferreinox SAS BIC.</p>
                                <p>Nos complace informarte que tu cuenta se encuentra al día. ¡Agradecemos tu excelente gestión y puntualidad!</p>
                                <p>Para tu control y referencia, adjuntamos a este correo tu estado de cuenta completo.</p>
                                <p>Recuerda que para futuras consultas o pagos, nuestro <a href='{portal_link}'><b>Portal de Pagos</b></a> está a tu disposición.</p>
                                <p>{instrucciones}</p>
                                <p>Gracias por tu confianza. Atentamente,<br><b>Área de Cartera Ferreinox SAS BIC</b></p>
                                </body></html>"""
                            
                            with st.spinner(f"Enviando correo a {email_destino}..."):
                                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                                    tmp.write(pdf_para_adjuntar)
                                    tmp_path = tmp.name
                                yag = yagmail.SMTP(sender_email, sender_password)
                                yag.send(to=email_destino, subject=asunto, contents=cuerpo_html, attachments=tmp_path)
                                os.remove(tmp_path)
                                st.success(f"¡Correo enviado exitosamente a {email_destino}!")
                        except Exception as e:
                            st.error(f"Error al enviar el correo: {e}")
                            st.error("Verifica las credenciales de correo en los 'secrets' de Streamlit.")

            with col_whatsapp:
                st.subheader("📲 Enviar por WhatsApp")
                numero_completo_para_mostrar = f"+57{telefono_cliente}" if telefono_cliente else "+57"
                numero_destino_wa = st.text_input("Verificar o modificar número de WhatsApp:", value=numero_completo_para_mostrar)
                
                if deuda_vencida_actual > 0:
                    mensaje_whatsapp = (
                        f"👋 ¡Hola {cliente_sel}! Te saludamos desde Ferreinox SAS BIC.\n\n"
                        f"Te recordamos que tu estado de cuenta presenta un valor total vencido de *${deuda_vencida_actual:,.0f}*. Tu factura más antigua tiene *{max_dias_retraso:.0f} días* de vencida.\n\n"
                        f"Para ponerte al día, puedes usar nuestro Portal de Pagos:\n"
                        f"🔗 https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/\n\n"
                        f"Tus datos de acceso son:\n"
                        f"👤 *Usuario:* {nit_cliente} (Tu NIT)\n"
                        f"🔑 *Código Único:* {cod_cliente}\n\n"
                        f"¡Agradecemos tu pronta gestión!"
                    )
                else:
                    mensaje_whatsapp = (
                        f"👋 ¡Hola {cliente_sel}! Te saludamos desde Ferreinox SAS BIC.\n\n"
                        f"¡Excelentes noticias! Tu cartera se encuentra al día. Agradecemos tu puntualidad y confianza.\n\n"
                        f"Recuerda que para futuras gestiones puedes usar nuestro Portal de Pagos:\n"
                        f"🔗 https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/\n\n"
                        f"Tus datos de acceso son:\n"
                        f"👤 *Usuario:* {nit_cliente} (Tu NIT)\n"
                        f"🔑 *Código Único:* {cod_cliente}\n\n"
                        f"¡Que tengas un excelente día!"
                    )

                mensaje_codificado = quote(mensaje_whatsapp)
                numero_limpio = re.sub(r'\D', '', numero_destino_wa)
                if numero_limpio:
                    url_whatsapp = f"https://wa.me/{numero_limpio}?text={mensaje_codificado}"
                    st.markdown(f'<a href="{url_whatsapp}" target="_blank" class="button">📱 Enviar Recordatorio</a>', unsafe_allow_html=True)
                else:
                    st.warning("Ingresa un número de teléfono válido.")
