# ======================================================================================
# ARCHIVO: Tablero_Comando_Ferreinox.py (v.PRO)
# Descripción: Panel de Control de Cartera para gestión operativa y análisis gerencial.
# ======================================================================================
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import os
import re
import unicodedata
from datetime import datetime
from urllib.parse import quote
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from fpdf import FPDF
import yagmail
import tempfile
import glob
from io import BytesIO

# --- CONFIGURACIÓN DE PÁGINA Y ESTILOS PROFESIONALES ---
st.set_page_config(
    page_title="🛡️ Centro de Mando: Cobranza Ferreinox PRO",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Paleta de Colores y CSS Corporativo
COLOR_PRIMARIO = "#003366"      # Azul oscuro corporativo
COLOR_ACCION = "#FFC300"        # Amarillo para acciones y énfasis
COLOR_FONDO = "#f0f2f6"         # Gris claro de fondo
COLOR_TARJETA = "#FFFFFF"       # Fondo de tarjetas y métricas

st.markdown(f"""
<style>
    .main {{ background-color: {COLOR_FONDO}; }}
    /* Métricas: Tarjetas con sombra y borde */
    .stMetric {{ 
        background-color: {COLOR_TARJETA}; 
        padding: 15px; 
        border-radius: 12px; 
        border-left: 6px solid {COLOR_PRIMARIO}; 
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
    }}
    /* Títulos */
    h1, h2, h3, .stTabs [data-baseweb="tab-list"] {{ color: {COLOR_PRIMARIO}; }}
    h1 {{ border-bottom: 2px solid {COLOR_ACCION}; padding-bottom: 10px; }}
    /* Tabs */
    .stTabs [aria-selected="true"] {{
        border-bottom: 3px solid {COLOR_ACCION};
        color: {COLOR_PRIMARIO};
        font-weight: bold;
    }}
    /* Botón WhatsApp */
    a.wa-link {{
        text-decoration: none; display: block; padding: 10px; margin-top: 10px;
        background-color: #25D366; color: white; border-radius: 8px; font-weight: bold;
        text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }}
    a.wa-link:hover {{ background-color: #128C7E; }}
    /* Botón Email */
    .email-btn {{ 
        background-color: {COLOR_ACCION}; color: {COLOR_PRIMARIO}; font-weight: bold;
        border: none; border-radius: 8px; padding: 10px; margin-top: 10px; width: 100%;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }}
    .email-btn:hover {{ background-color: #FFD700; }}
</style>
""", unsafe_allow_html=True)

# ======================================================================================
# 1. MOTOR DE CONEXIÓN Y LIMPIEZA DE DATOS (MÁS ROBUSTO)
# ======================================================================================

def normalizar_texto(texto):
    if not isinstance(texto, str): return str(texto)
    # 1. Normaliza a NFD, codifica a ASCII (quita tildes), decodifica a UTF-8, pone en mayúsculas y quita espacios
    texto = unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode("utf-8").upper().strip()
    # 2. Quita cualquier caracter que no sea alfanumérico, espacio o punto (mantiene el punto para NIT/ID)
    return re.sub(r'[^\w\s\.]', '', texto).strip()

def limpiar_moneda(valor):
    if pd.isna(valor): return 0.0
    s_val = str(valor).strip()
    # Eliminar símbolos de moneda y separadores de miles innecesarios
    s_val = re.sub(r'[^\d.,-]', '', s_val)
    if not s_val: return 0.0
    try:
        if ',' in s_val and '.' in s_val:
            # Si la coma está después del punto, es formato USA: 1,000.00
            if s_val.rfind(',') < s_val.rfind('.'):
                s_val = s_val.replace(',', '')
            # Si el punto está después de la coma, es formato Latino/Euro: 1.000,00
            else:
                s_val = s_val.replace('.', '').replace(',', '.')
        elif ',' in s_val:
            # Si solo hay comas, y la parte decimal tiene 2 dígitos, es un decimal latino.
            if len(s_val.split(',')[-1]) <= 2: 
                s_val = s_val.replace(',', '.')
            # De lo contrario, asume separador de miles.
            else: 
                s_val = s_val.replace(',', '')
        
        return float(s_val)
    except Exception:
        return 0.0 # Retorna 0.0 en caso de error

def mapear_y_limpiar_df(df):
    # Normalizar todas las columnas primero
    df.columns = [normalizar_texto(c) for c in df.columns]
    
    # Diccionario de mapeo robusto con columnas requeridas
    mapa = {
        'cliente': ['NOMBRECLIENTE', 'RAZONSOCIAL', 'TERCERO', 'CLIENTE'],
        'nit': ['NIT', 'IDENTIFICACION', 'CEDULA', 'RUT'],
        'saldo': ['IMPORTE', 'SALDO', 'TOTAL', 'DEUDA', 'VALOR'],
        'dias': ['DIASVENCIDO', 'DIAS', 'VENCIDO', 'MORA', 'ANTIGUEDAD'],
        'telefono': ['TELEFONO1', 'TEL', 'MOVIL', 'CELULAR', 'TELEFONO'],
        'vendedor': ['NOMVENDEDOR', 'VENDEDOR', 'ASESOR', 'COMERCIAL'],
        'factura': ['NUMERO', 'FACTURA', 'DOC'],
        'email': ['EMAIL', 'CORREO', 'E-MAIL', 'MAIL'],
        'cod_cliente': ['CODCLIENTE', 'CODIGO']
    }
    
    renombres = {}
    for standard, variantes in mapa.items():
        for col in df.columns:
            if standard not in renombres.values() and any(v in col for v in variantes):
                renombres[col] = standard
                break
    
    df.rename(columns=renombres, inplace=True)
    
    # ⚠️ Validación mínima (Factura, Saldo, Días son críticos para gestión)
    req = ['cliente', 'saldo', 'dias', 'factura']
    if not all(c in df.columns for c in req):
        missing = [c for c in req if c not in df.columns]
        return None, f"Faltan columnas críticas mapeadas: {', '.join(missing)}."

    # Inclusión de columnas opcionales con valor 'N/A' si faltan
    for c in ['telefono', 'vendedor', 'nit', 'email', 'cod_cliente']:
        if c not in df.columns: df[c] = 'N/A'
        else: df[c] = df[c].fillna('N/A').astype(str)

    # Conversión de tipos
    df['saldo'] = df['saldo'].apply(limpiar_moneda)
    df['dias'] = pd.to_numeric(df['dias'], errors='coerce').fillna(0).astype(int)
    df['vendedor'] = df['vendedor'].apply(normalizar_texto) # Normalizar vendedores
    
    # Limpieza final: Quitar facturas con saldo cero
    df = df[df['saldo'] != 0].copy() 
    return df, "OK"

@st.cache_data(ttl=300)
def cargar_datos_automaticos(nombre_archivo="cartera_detalle"):
    """Busca archivos automáticamente en la carpeta local con el nombre específico."""
    
    # Buscar archivos Excel o CSV que coincidan con el nombre
    archivos = glob.glob(f"{nombre_archivo}*.xlsx") + glob.glob(f"{nombre_archivo}*.csv")
    
    if not archivos:
        return None, f"No se encontró el archivo clave '{nombre_archivo}.xlsx' o '{nombre_archivo}.csv' localmente."
    
    # Priorizar el archivo que parezca más directo (aunque glob ya es sensible)
    archivo_prioritario = archivos[0]
    
    try:
        if archivo_prioritario.endswith('.csv'):
            # El separador es None para que Pandas lo detecte automáticamente
            df = pd.read_csv(archivo_prioritario, sep=None, engine='python', encoding='latin-1', dtype=str)
        else:
            df = pd.read_excel(archivo_prioritario, dtype=str)
            
        # Intentar mapear y limpiar
        df_proc, status = mapear_y_limpiar_df(df)
        
        if df_proc is None: 
            return None, f"Error en procesamiento de datos: {status}"
            
        return df_proc, f"Conectado a la fuente principal: **{os.path.basename(archivo_prioritario)}**"
    
    except Exception as e:
        return None, f"Error leyendo {os.path.basename(archivo_prioritario)}: {str(e)}"

# ======================================================================================
# 2. INTELIGENCIA DE NEGOCIO (ESTRATEGIA Y FUNCIONES)
# ======================================================================================

def segmentar_cartera(df):
    # Segmentación estratégica de cartera
    bins = [-float('inf'), 0, 15, 30, 60, 90, float('inf')]
    labels = ["🟢 Al Día", "🟡 Prev. (1-15)", "🟠 Riesgo (16-30)", "🔴 Crítico (31-60)", "🚨 Alto Riesgo (61-90)", "⚫ Legal (+90)"]
    df['Rango'] = pd.cut(df['dias'], bins=bins, labels=labels)
    return df

def calcular_kpis(df):
    total = df['saldo'].sum()
    vencido = df[df['dias'] > 0]['saldo'].sum()
    pct_vencido = (vencido / total * 100) if total else 0
    clientes_mora = df[df['dias'] > 0]['cliente'].nunique()
    
    # Calcular CSI (Collection Severity Index) = Suma(Saldo * Días) / Saldo Total
    df_vencido = df[df['dias'] > 0].copy()
    csi = (df_vencido['saldo'] * df_vencido['dias']).sum() / total if total > 0 else 0
    
    return total, vencido, pct_vencido, clientes_mora, csi

def generar_link_wa(telefono, cliente, saldo_vencido, dias_max, nit, cod_cliente):
    # Limpiar y estandarizar el número (asume Colombia si son 10 dígitos)
    tel = re.sub(r'\D', '', str(telefono))
    if len(tel) == 10: tel = '57' + tel 
    if len(tel) < 10: return None
    
    cliente_corto = str(cliente).split()[0].title()
    portal_link = "https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/"
    
    if saldo_vencido <= 0:
        msg = (
            f"👋 ¡Hola {cliente_corto}! Te saludamos de Ferreinox SAS BIC.\n\n"
            f"¡Felicitaciones! Tu cuenta está al día. Agradecemos tu puntualidad.\n\n"
            f"Hemos enviado tu estado de cuenta completo a tu correo. ¡Gracias por tu confianza!"
        )
    elif dias_max <= 30:
        msg = (
            f"👋 ¡Hola {cliente_corto}! Te saludamos de Ferreinox SAS BIC.\n\n"
            f"Recordatorio amable: Tienes un saldo vencido de *${saldo_vencido:,.0f}*. La factura más antigua tiene *{dias_max} días* de vencida.\n\n"
            f"Puedes usar nuestro portal de pagos 🔗 {portal_link} con tu NIT ({nit}) y Código Interno ({cod_cliente}).\n\n"
            f"¡Agradecemos tu pago hoy mismo!"
        )
    else: # Días > 30 (Alerta crítica)
        msg = (
            f"🚨 URGENTE {cliente_corto}: Su cuenta en Ferreinox SAS BIC presenta un saldo de *${saldo_vencido:,.0f}* con hasta *{dias_max} días* de mora.\n\n"
            f"Requerimos su pago inmediato para evitar medidas como el bloqueo de cupo o inicio de cobro pre-jurídico.\n\n"
            f"Pague aquí 🔗 {portal_link}\n\n"
            f"Usuario (NIT): {nit}\nCódigo Único: {cod_cliente}\n\n"
            f"Por favor, conteste este mensaje para confirmar su compromiso de pago."
        )
            
    return f"https://wa.me/{tel}?text={quote(msg)}"

# ======================================================================================
# 3. GENERADORES (PDF Y EXCEL)
# ======================================================================================

class PDF(FPDF):
    """Clase personalizada para generar PDF con estilos de Ferreinox."""
    def header(self):
        try:
            # Usar un logo ficticio si el real no está disponible
            #  
            self.set_font('Arial', 'B', 12)
            self.set_text_color(0, 51, 102)
            self.cell(0, 10, 'FERREINOX SAS BIC', 0, 1, 'L')
        except Exception:
            self.set_font('Arial', 'B', 12); self.cell(80, 10, 'Ferreinox SAS BIC (Logo)', 0, 0, 'L')
            
        self.set_font('Arial', 'B', 18); self.set_text_color(0, 0, 0)
        self.cell(0, 10, 'ESTADO DE CUENTA', 0, 1, 'R')
        self.ln(5)

    def footer(self):
        self.set_y(-30)
        self.set_font('Arial', 'I', 8); self.set_text_color(100, 100, 100)
        
        # Enlace Portal de Pagos
        portal_link = "https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/"
        self.set_font('Arial', 'B', 10); self.set_text_color(0, 51, 102)
        self.cell(0, 5, 'Portal de Pagos Ferreinox: ', 0, 1, 'C', link=portal_link)
        self.set_font('Arial', 'I', 8); self.set_text_color(100, 100, 100)
        self.cell(0, 5, f'Generado el: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', 0, 1, 'C')
        self.cell(0, 5, f'Página {self.page_no()}', 0, 0, 'C')

def crear_pdf(df_cliente, total_vencido_cliente):
    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=40)
    pdf.add_page()
    
    if df_cliente.empty:
        pdf.set_font('Arial', 'B', 12); pdf.cell(0, 10, 'No se encontraron facturas para este cliente.', 0, 1, 'C')
        return bytes(pdf.output())
        
    row = df_cliente.iloc[0]
    
    # --- Datos Cliente ---
    pdf.set_font("Arial", 'B', 11); pdf.set_text_color(0, 51, 102)
    pdf.cell(40, 6, "Cliente:", 0, 0); pdf.set_font("Arial", '', 11); pdf.set_text_color(0, 0, 0); pdf.cell(0, 6, row['cliente'], 0, 1)
    pdf.set_font("Arial", 'B', 11); pdf.set_text_color(0, 51, 102)
    pdf.cell(40, 6, "NIT/ID:", 0, 0); pdf.set_font("Arial", '', 11); pdf.set_text_color(0, 0, 0); pdf.cell(0, 6, row['nit'], 0, 1)
    pdf.set_font("Arial", 'B', 11); pdf.set_text_color(0, 51, 102)
    pdf.cell(40, 6, "Asesor:", 0, 0); pdf.set_font("Arial", '', 11); pdf.set_text_color(0, 0, 0); pdf.cell(0, 6, row['vendedor'], 0, 1)
    pdf.ln(5)
    
    # --- Tabla de Facturas ---
    pdf.set_font('Arial', 'B', 10); pdf.set_fill_color(0, 51, 102); pdf.set_text_color(255, 255, 255)
    pdf.cell(30, 8, "Factura", 1, 0, 'C', 1)
    pdf.cell(30, 8, "Días Mora", 1, 0, 'C', 1)
    pdf.cell(45, 8, "Fecha Venc.", 1, 0, 'C', 1)
    pdf.cell(45, 8, "Saldo", 1, 1, 'C', 1)
    
    pdf.set_font("Arial", '', 10)
    total_cartera = 0
    for _, item in df_cliente.iterrows():
        total_cartera += item['saldo']
        # Estilo para filas vencidas
        if item['dias'] > 0:
            pdf.set_fill_color(255, 235, 238) # Fondo rojo claro
            pdf.set_text_color(150, 0, 0)
        else:
            pdf.set_fill_color(255, 255, 255) # Fondo blanco
            pdf.set_text_color(0, 0, 0)
            
        pdf.cell(30, 7, str(item['factura']), 1, 0, 'C', 1)
        pdf.cell(30, 7, str(item['dias']), 1, 0, 'C', 1)
        # Fecha Vencimiento (simulación, ya que no la tenemos mapeada en el código base)
        fecha_vencimiento = f"{item['dias']} DÍAS" # Se usa una simulación
        pdf.cell(45, 7, fecha_vencimiento, 1, 0, 'C', 1) 
        pdf.cell(45, 7, f"${item['saldo']:,.0f}", 1, 1, 'R', 1)
            
    # --- Totales ---
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", 'B', 11); pdf.set_fill_color(224, 224, 224)
    pdf.cell(105, 8, "TOTAL CARTERA", 1, 0, 'R', 1)
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(45, 8, f"${total_cartera:,.0f}", 1, 1, 'R', 1)

    if total_vencido_cliente > 0:
        pdf.set_font('Arial', 'B', 12); pdf.set_fill_color(255, 204, 204); pdf.set_text_color(192, 0, 0)
        pdf.cell(105, 8, 'VALOR TOTAL VENCIDO A PAGAR', 1, 0, 'R', 1)
        pdf.cell(45, 8, f"${total_vencido_cliente:,.0f}", 1, 1, 'R', 1)
            
    return bytes(pdf.output())

def crear_excel_gerencial(df, kpis, total, vencido, pct_mora, clientes_mora, csi):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen Gerencial"
    
    # Estilos
    header_style = Font(bold=True, color="FFFFFF")
    fill_blue = PatternFill("solid", fgColor="003366")
    fill_kpi = PatternFill("solid", fgColor="FFC300")
    
    # --- KPIs en Excel ---
    ws['A1'] = "REPORTE GERENCIAL DE CARTERA - FERREINOX"
    ws['A1'].font = Font(size=16, bold=True)
    
    kpi_labels = ["Total Cartera", "Total Vencido", "% Mora", "Clientes en Mora", "Índice de Severidad (CSI)"]
    kpi_values = [total, vencido, pct_mora / 100, clientes_mora, csi]
    formats = ['$#,##0', '$#,##0', '0.0%', '0', '0.00']
    
    for i, (lab, val, fmt) in enumerate(zip(kpi_labels, kpi_values, formats)):
        c_lab = ws.cell(row=3, column=i+1, value=lab)
        c_lab.font = Font(bold=True)
        c_lab.fill = fill_blue
        c_val = ws.cell(row=4, column=i+1, value=val)
        c_val.number_format = fmt
        c_val.font = Font(bold=True)
        c_val.fill = fill_kpi
        
    # --- Tabla Detalle ---
    ws['A6'] = "DETALLE COMPLETO DE LA CARTERA (Filtrable)"
    ws['A6'].font = Font(size=12, bold=True)
    
    # Columnas a incluir en el reporte (las estándar más las originales)
    cols = ['cliente', 'nit', 'factura', 'vendedor', 'cod_cliente', 'Rango', 'dias', 'saldo', 'telefono', 'email']
    df_detalle = df.sort_values(by='dias', ascending=False).reset_index(drop=True)

    # Headers de la tabla
    for col_num, col_name in enumerate(cols, 1):
        c = ws.cell(row=7, column=col_num, value=col_name.upper().replace('_', ' '))
        c.fill = fill_blue
        c.font = header_style
        
    # Data
    for row_num, row_data in enumerate(df_detalle[cols].values, 8):
        for col_num, val in enumerate(row_data, 1):
            c = ws.cell(row=row_num, column=col_num, value=val)
            if col_num == 8: c.number_format = '$#,##0' # Columna Saldo
            
    # Autoajuste de columnas y filtros
    ws.auto_filter.ref = f"A7:{get_column_letter(len(cols))}{len(df_detalle)+7}"
    for i in range(1, len(cols) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 20
        
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# ======================================================================================
# 4. FUNCIÓN DE ENVÍO DE CORREO (yagmail)
# ======================================================================================
def enviar_correo(destinatario, asunto, cuerpo_html, pdf_bytes, nombre_pdf):
    # Credenciales de correo (simuladas aquí, en una app real se leerían de st.secrets)
    email_user = st.session_state.get('email_user', 'USUARIO_CORREO')
    email_pass = st.session_state.get('email_pass', 'CONTRASEÑA_APP')

    if email_user == 'USUARIO_CORREO' or not email_pass:
        st.error("⚠️ Configura las credenciales de correo en la barra lateral antes de enviar.")
        return False

    try:
        # Guardar PDF temporalmente
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(pdf_bytes)
            tmp_path = tmp.name

        yag = yagmail.SMTP(email_user, email_pass)
        
        yag.send(
            to=destinatario,
            subject=asunto,
            contents=[cuerpo_html, tmp_path] # Adjunta el PDF
        )
        
        os.remove(tmp_path) # Limpiar el archivo temporal
        return True
    except Exception as e:
        st.error(f"Error enviando correo. Asegúrate que la contraseña es una 'Contraseña de Aplicación' (no la normal): {e}")
        try:
            if os.path.exists(tmp_path): os.remove(tmp_path)
        except:
            pass
        return False
        
# --- PLANTILLAS HTML PARA CORREO (MÁS COMPACTAS Y PROFESIONALES) ---

def plantilla_correo_vencido(cliente, saldo, dias, nit, cod_cliente, portal_link):
    return f"""
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #ddd; border-radius: 10px; overflow: hidden;">
        <div style="background-color: {COLOR_PRIMARIO}; color: white; padding: 20px; text-align: center;">
            <h2 style="margin: 0;">🚨 Recordatorio de Saldo Pendiente</h2>
        </div>
        <div style="padding: 20px;">
            <p><strong>Estimado(a) {cliente},</strong></p>
            <p>Le contactamos de <strong>Ferreinox SAS BIC</strong>. Hemos identificado que su cuenta presenta un saldo vencido.</p>
            <div style="background-color: #ffebeb; border-left: 5px solid #d32f2f; padding: 15px; margin: 20px 0; border-radius: 5px;">
                <p style="font-size: 1.1em; font-weight: bold; color: #d32f2f; margin: 0;">Valor Total Vencido: ${saldo:,.0f}</p>
                <p style="margin: 5px 0 0 0; font-size: 0.9em;">(Su factura más antigua tiene {dias} días de vencimiento)</p>
            </div>
            <p>Agradecemos su gestión inmediata para evitar el bloqueo de su cupo.</p>
            <p style="text-align: center; margin: 30px 0;">
                <a href="{portal_link}" style="background-color: #16a34a; color: white; padding: 12px 25px; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 1.1em;">🚀 IR AL PORTAL DE PAGOS</a>
            </p>
            <p style="font-size: 0.9em;"><strong>Datos de Acceso:</strong> Usuario (NIT): {nit} | Código Único: {cod_cliente}</p>
            <p>Adjunto encontrará el estado de cuenta detallado en PDF.</p>
        </div>
        <div style="background-color: #f0f2f6; color: #555; padding: 15px; text-align: center; font-size: 0.8em;">
            <p style="margin: 5px 0;"><strong>Departamento de Cartera y Recaudos.</strong></p>
            <p style="margin: 5px 0;">Si ya realizó el pago, por favor ignore este mensaje y envíenos el comprobante.</p>
        </div>
    </div>
    """

def plantilla_correo_al_dia(cliente, total_cartera):
    return f"""
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #ddd; border-radius: 10px; overflow: hidden;">
        <div style="background-color: #16a34a; color: white; padding: 20px; text-align: center;">
            <h2 style="margin: 0;">🎉 ¡Cuenta al Día!</h2>
        </div>
        <div style="padding: 20px;">
            <p><strong>Estimado(a) {cliente},</strong></p>
            <p>Le contactamos de <strong>Ferreinox SAS BIC</strong> para agradecer su excelente gestión de pagos.</p>
            <div style="background-color: #e6ffe6; border-left: 5px solid #388e3c; padding: 15px; margin: 20px 0; border-radius: 5px; text-align: center;">
                <p style="font-size: 1.1em; font-weight: bold; color: #388e3c; margin: 0;">Su cuenta no presenta saldos vencidos.</p>
                <p style="margin: 5px 0 0 0; font-size: 0.9em;">(Saldo total en Cartera: ${total_cartera:,.0f})</p>
            </div>
            <p>Adjunto encontrará su estado de cuenta completo para su referencia y control.</p>
            <p>¡Gracias por su confianza y por ser un cliente puntual!</p>
        </div>
        <div style="background-color: #f0f2f6; color: #555; padding: 15px; text-align: center; font-size: 0.8em;">
            <p style="margin: 5px 0;"><strong>Departamento de Cartera y Recaudos.</strong></p>
        </div>
    </div>
    """

# ======================================================================================
# 5. DASHBOARD PRINCIPAL (MAIN)
# ======================================================================================

def main():
    # --- BARRA LATERAL: CONFIGURACIÓN Y FILTROS ---
    with st.sidebar:
        st.image("https://cdn-icons-png.flaticon.com/512/9322/9322127.png", width=50)
        st.header("⚙️ Configuración")
        
        # Credenciales de Correo (Sesión para evitar re-ingreso)
        with st.expander("📧 Configurar Correo (yagmail)"):
            st.session_state['email_user'] = st.text_input("Tu Correo (Remitente)", value=st.session_state.get('email_user', ''), key='side_user')
            st.session_state['email_pass'] = st.text_input("Contraseña de Aplicación", type="password", value=st.session_state.get('email_pass', ''), key='side_pass')
            st.caption("Nota: Para Gmail/Outlook, usa una 'Contraseña de Aplicación', no tu clave normal.")

        st.divider()
        st.header("🔍 Filtros Operativos")
        
    # --- CARGA DE DATOS ---
    # Intentar cargar el archivo 'cartera_detalle'
    df, status_carga = cargar_datos_automaticos("cartera_detalle") 
    
    # Manejar carga manual si la automática falla
    if df is None:
        st.error(f"{status_carga} - Por favor, sube el archivo **'cartera_detalle'** o la cartera más reciente manualmente:")
        uploaded = st.file_uploader("Subir Excel/CSV de Cartera", type=['xlsx', 'csv'])
        if uploaded:
            if uploaded.name.endswith('.csv'):
                df_raw = pd.read_csv(uploaded, sep=None, engine='python', encoding='latin-1', dtype=str)
            else:
                df_raw = pd.read_excel(uploaded, dtype=str)
            df, status_manual = mapear_y_limpiar_df(df_raw)
            if df is None:
                st.error(f"Error en datos cargados manualmente: {status_manual}")
                st.stop()
            status_carga = f"Conectado a la fuente manual: **{uploaded.name}**"

    if df is None:
        st.stop() # Detener la app si los datos son críticos

    # --- PROCESAMIENTO Y FILTROS ---
    df = segmentar_cartera(df)
    
    # Filtros Dinámicos
    vendedores = ["TODOS"] + sorted(df['vendedor'].unique().tolist())
    filtro_vendedor = st.sidebar.selectbox("Filtrar por Vendedor:", vendedores)
    
    rangos_cartera = ["TODOS"] + df['Rango'].cat.categories.tolist()
    filtro_rango = st.sidebar.selectbox("Filtrar por Antigüedad:", rangos_cartera)
    
    df_view = df.copy()
    if filtro_vendedor != "TODOS":
        df_view = df_view[df_view['vendedor'] == filtro_vendedor]
    if filtro_rango != "TODOS":
        df_view = df_view[df_view['Rango'] == filtro_rango]

    if df_view.empty:
        st.warning("No hay datos para la selección actual de filtros.")
        df_view = df.copy() # Mostrar data sin filtrar si la vista queda vacía
    
    total, vencido, pct_mora, clientes_mora, csi = calcular_kpis(df_view)

    # --- ENCABEZADO Y KPIS ---
    st.title(f"Centro de Mando: Cobranza Ferreinox")
    st.caption(status_carga)
    
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("💰 Cartera Total", f"${total:,.0f}")
    k2.metric("🔥 Cartera Vencida", f"${vencido:,.0f}")
    k3.metric("📈 % Vencido s/ Total", f"{pct_mora:.1f}%", delta="-0.5%" if pct_mora < 10 else "0.5%") # Delta simulado
    k4.metric("👥 Clientes en Mora", clientes_mora)
    k5.metric("💥 Índice de Severidad (CSI)", f"{csi:,.1f}")
    
    st.divider()

    # --- TABS DE GESTIÓN ---
    tab_lider, tab_gerente, tab_datos = st.tabs(["👩‍💼 GESTIÓN OPERATIVA (1 a 1)", "👨‍💼 ANÁLISIS GERENCIAL", "📥 EXPORTAR Y DATOS"])

    # ==============================================================================
    # TAB LÍDER: GESTIÓN DE COBRO 1 A 1
    # ==============================================================================
    with tab_lider:
        st.subheader("🎯 Módulo de Contacto Directo y Envío de Docs.")
        
        # Agrupar por Cliente para gestión (solo clientes con saldo > 0)
        df_agrupado = df_view[df_view['saldo'] > 0].groupby('cliente').agg(
            saldo=('saldo', 'sum'),
            dias_max=('dias', 'max'),
            facturas=('factura', lambda x: ', '.join(x.astype(str).tolist())), # Lista de facturas
            telefono=('telefono', 'first'),
            email=('email', 'first'),
            vendedor=('vendedor', 'first'),
            nit=('nit', 'first'),
            cod_cliente=('cod_cliente', 'first')
        ).reset_index().sort_values('saldo', ascending=False)
        
        # Filtrar solo clientes en mora para la gestión 1 a 1 (opcional)
        clientes_en_mora = df_agrupado[df_agrupado['dias_max'] > 0]
        
        # Selector de cliente
        cliente_sel = st.selectbox("🔍 Selecciona Cliente a Gestionar (Priorizado por Deuda Vencida)", 
                                   clientes_en_mora['cliente'].tolist() if not clientes_en_mora.empty else df_agrupado['cliente'].tolist())
        
        if cliente_sel:
            data_cli = df_agrupado[df_agrupado['cliente'] == cliente_sel].iloc[0]
            detalle_facturas = df_view[df_view['cliente'] == cliente_sel].sort_values('dias', ascending=False)
            
            saldo_vencido_cli = detalle_facturas[detalle_facturas['dias'] > 0]['saldo'].sum()
            
            c1, c2 = st.columns([1, 2])
            
            with c1:
                st.markdown(f"#### Resumen de Cliente: **{cliente_sel}**")
                st.info(f"**Deuda Total:** ${data_cli['saldo']:,.0f}")
                st.warning(f"**Deuda Vencida:** ${saldo_vencido_cli:,.0f}")
                st.error(f"**Días Máx Mora:** {data_cli['dias_max']} días")
                st.text(f"📞 {data_cli['telefono']} | 📧 {data_cli['email']}")
                st.text(f"ID: {data_cli['nit']} | Cód. Cliente: {data_cli['cod_cliente']}")
                
                # Generar PDF en memoria
                pdf_bytes = crear_pdf(detalle_facturas, saldo_vencido_cli)
                
                # --- BOTÓN WHATSAPP ---
                link_wa = generar_link_wa(data_cli['telefono'], cliente_sel, saldo_vencido_cli, data_cli['dias_max'], data_cli['nit'], data_cli['cod_cliente'])
                if link_wa:
                    st.markdown(f"""<a href="{link_wa}" target="_blank" class="wa-link">📱 ABRIR WHATSAPP CON GUION</a>""", unsafe_allow_html=True)
                else:
                    st.error("Número de teléfono inválido para WhatsApp")
                
                st.download_button(label="📄 Descargar PDF Local", data=pdf_bytes, file_name=f"Estado_Cuenta_{cliente_sel}.pdf", mime="application/pdf")

            with c2:
                st.write("#### 📄 Detalle de Facturas (Priorizadas por Mora)")
                # Vista previa de facturas
                st.dataframe(detalle_facturas[['factura', 'dias', 'saldo', 'vendedor']].style.format({'saldo': '${:,.0f}'}).background_gradient(subset=['dias'], cmap='YlOrRd'), height=250, use_container_width=True, hide_index=True)
                
                # --- ENVÍO DE CORREO ---
                st.write("#### 📧 Envío de Estado de Cuenta por Correo")
                with st.form("form_email"):
                    email_dest = st.text_input("Destinatario", value=data_cli['email'], key="email_dest_input")
                    
                    if saldo_vencido_cli > 0:
                        asunto_msg = f"Recordatorio URGENTE de Saldo Pendiente - {cliente_sel}"
                        cuerpo_html = plantilla_correo_vencido(cliente_sel, saldo_vencido_cli, data_cli['dias_max'], data_cli['nit'], data_cli['cod_cliente'], "https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/")
                    else:
                        asunto_msg = f"Tu Estado de Cuenta Actualizado - {cliente_sel} (Cta al Día)"
                        cuerpo_html = plantilla_correo_al_dia(cliente_sel, data_cli['saldo'])
                        
                    submit_email = st.form_submit_button("📧 ENVIAR CORREO CON PDF ADJUNTO", type="primary")
                    
                    if submit_email:
                        if enviar_correo(email_dest, asunto_msg, cuerpo_html, pdf_bytes, f"EstadoCuenta_{cliente_sel}.pdf"):
                            st.success(f"✅ Correo enviado a {email_dest}")
                        else:
                            st.error("❌ Falló el envío. Revisa credenciales y logs.")


    # ==============================================================================
    # TAB GERENTE: VISIÓN ESTRATÉGICA
    # ==============================================================================
    with tab_gerente:
        st.subheader("📊 Análisis de Cartera por Segmento y Concentración")
        
        c_pie, c_bar = st.columns(2)
        
        with c_pie:
            st.markdown("**1. Distribución de Saldo por Rango de Mora**")
            df_pie = df_view.groupby('Rango', observed=True)['saldo'].sum().reset_index()
            # Mapeo de colores coherente con los rangos
            color_map = {"🟢 Al Día": "green", "🟡 Prev. (1-15)": "gold", "🟠 Riesgo (16-30)": "orange", 
                         "🔴 Crítico (31-60)": "orangered", "🚨 Alto Riesgo (61-90)": "red", "⚫ Legal (+90)": "black"}
            fig_pie = px.pie(df_pie, names='Rango', values='saldo', color='Rango', 
                             color_discrete_map=color_map, hole=.3)
            fig_pie.update_traces(textinfo='percent+label', marker=dict(line=dict(color='#000000', width=1)))
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with c_bar:
            st.markdown("**2. Top 10 Clientes Morosos (Pareto)**")
            # Solo clientes con mora y saldo positivo (filtrado de anticipos)
            top_cli = df_view[(df_view['dias'] > 0) & (df_view['saldo'] > 0)].groupby('cliente')['saldo'].sum().nlargest(10).reset_index()
            fig_bar = px.bar(top_cli, x='saldo', y='cliente', orientation='h', 
                             text_auto='$.2s', title="Monto de Deuda Vencida", 
                             color_discrete_sequence=[COLOR_PRIMARIO])
            fig_bar.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_bar, use_container_width=True)
            
        st.markdown("---")
        st.markdown("### 3. Desempeño y Riesgo por Vendedor")
        resumen_vendedor = df_view.groupby('vendedor').agg(
            Cartera_Total=('saldo', 'sum'),
            Vencido=('saldo', lambda x: x[df_view.loc[x.index, 'dias'] > 0].sum())
        ).reset_index()
        resumen_vendedor['% Vencido'] = (resumen_vendedor['Vencido'] / resumen_vendedor['Cartera_Total'] * 100).fillna(0)
        
        # Cálculo de Clientes en Mora y CSI por Vendedor
        vencidos_df = df_view[df_view['dias'] > 0]
        clientes_mora_vendedor = vencidos_df.groupby('vendedor')['cliente'].nunique().reset_index().rename(columns={'cliente': 'Clientes_Mora'})
        csi_vendedor = (vencidos_df.groupby('vendedor').apply(lambda x: (x['saldo'] * x['dias']).sum() / x['saldo'].sum() if x['saldo'].sum() > 0 else 0, include_groups=False).reset_index(name='CSI'))
        
        resumen_vendedor = resumen_vendedor.merge(clientes_mora_vendedor, on='vendedor', how='left').fillna(0)
        resumen_vendedor = resumen_vendedor.merge(csi_vendedor, on='vendedor', how='left').fillna(0)
        
        # Formato profesional para la tabla
        styled_df = resumen_vendedor.style.format({
            'Cartera_Total': '${:,.0f}', 
            'Vencido': '${:,.0f}', 
            '% Vencido': '{:.1f}%',
            'Clientes_Mora': '{:,.0f}',
            'CSI': '{:,.1f}'
        }).background_gradient(subset=['% Vencido'], cmap='RdYlGn_r').background_gradient(subset=['CSI'], cmap='OrRd')
        
        st.dataframe(styled_df, use_container_width=True, hide_index=True)


    # ==============================================================================
    # TAB DATOS: EXPORTAR EXCEL
    # ==============================================================================
    with tab_datos:
        st.subheader("📥 Descarga del Reporte Gerencial")
        
        excel_data = crear_excel_gerencial(df_view, [], total, vencido, pct_mora, clientes_mora, csi)
        
        st.download_button(
            label="💾 DESCARGAR REPORTE GERENCIAL (EXCEL) - Formato Profesional",
            data=excel_data,
            file_name=f"Reporte_Cartera_{filtro_vendedor}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.subheader("🔎 Datos Crudos Filtrados")
        # Mostrar el dataframe completo con las columnas mapeadas y la segmentación
        st.dataframe(df_view.style.format({'saldo': '${:,.0f}', 'dias': '{:,.0f}'}), use_container_width=True, hide_index=True)

if __name__ == "__main__":
    # Inicializar las variables de sesión si no existen
    if 'email_user' not in st.session_state:
        st.session_state['email_user'] = ''
    if 'email_pass' not in st.session_state:
        st.session_state['email_pass'] = ''
        
    main()
