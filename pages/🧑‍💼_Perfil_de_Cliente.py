# ======================================================================================
# ARCHIVO: Tablero_Comando_Ferreinox_PRO.py (v.FINAL UNIFICADA)
# Descripción: Panel de Control de Cartera para gestión operativa, análisis gerencial y
#              conexión automática a Dropbox con autenticación.
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
from io import BytesIO, StringIO
import dropbox # Conexión a Dropbox
import toml # Para manejo de secretos

# --- CONFIGURACIÓN DE PÁGINA Y ESTILOS PROFESIONALES ---
st.set_page_config(
    page_title="🛡️ Centro de Mando: Cobranza Ferreinox PRO",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Paleta de Colores y CSS Corporativo (Unificada)
COLOR_PRIMARIO = "#003865"       # Azul oscuro corporativo (Similar al unificado)
COLOR_ACCION = "#FFC300"         # Amarillo para acciones y énfasis
COLOR_FONDO = "#f0f2f6"          # Gris claro de fondo
COLOR_TARJETA = "#FFFFFF"        # Fondo de tarjetas y métricas
COLOR_ALERTA_CRITICA = "#D32F2F" # Rojo para alertas

st.markdown(f"""
<style>
    .stApp {{ background-color: {COLOR_FONDO}; }}
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
    /* Botón Email (Para uniformidad con el código 1) */
    .email-btn {{ 
        background-color: {COLOR_ACCION}; color: {COLOR_PRIMARIO}; font-weight: bold;
        border: none; border-radius: 8px; padding: 10px; margin-top: 10px; width: 100%;
        box-shadow: 0 2px 4px rgba(0,0,0,0.2);
    }}
    .email-btn:hover {{ background-color: #FFD700; }}
    
    /* Input/Select estilo profesional (tomado del código 2) */
    div[data-baseweb="input"], div[data-baseweb="select"], div.st-multiselect, div.st-text-area {{ background-color: #FFFFFF; border: 1.5px solid {COLOR_PRIMARIO}; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); padding-left: 5px; }}

</style>
""", unsafe_allow_html=True)


# ======================================================================================
# 1. MOTOR DE CONEXIÓN, LIMPIEZA Y PROCESAMIENTO
# ======================================================================================

def normalizar_texto(texto):
    if not isinstance(texto, str): return str(texto)
    # Normaliza, quita tildes, pone en mayúsculas y quita caracteres especiales (mantiene espacios y puntos)
    texto = unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode("utf-8").upper().strip()
    return re.sub(r'[^\w\s\.]', '', texto).strip()

def normalizar_nombre(nombre: str) -> str:
    """Función para normalizar nombres de vendedores/clientes para la autenticación."""
    if not isinstance(nombre, str): return ""
    nombre = nombre.upper().strip().replace('.', '')
    nombre = ''.join(c for c in unicodedata.normalize('NFD', nombre) if unicodedata.category(c) != 'Mn')
    return ' '.join(nombre.split())

def mapear_y_limpiar_df(df_raw):
    """Mapea y limpia el DataFrame según las columnas esperadas."""
    
    # 1. Mapeo de columnas con el robusto sistema de variantes
    df = df_raw.copy()
    
    # Normalizar todas las columnas primero
    df.columns = [normalizar_texto(c) for c in df.columns]
    
    # Diccionario de mapeo robusto con columnas requeridas
    mapa = {
        'nombrecliente': ['NOMBRECLIENTE', 'RAZONSOCIAL', 'TERCERO', 'CLIENTE'],
        'nit': ['NIT', 'IDENTIFICACION', 'CEDULA', 'RUT'],
        'importe': ['IMPORTE', 'SALDO', 'TOTAL', 'DEUDA', 'VALOR'],
        'dias_vencido': ['DIASVENCIDO', 'DIAS', 'VENCIDO', 'MORA', 'ANTIGUEDAD'],
        'telefono1': ['TELEFONO1', 'TEL', 'MOVIL', 'CELULAR', 'TELEFONO'],
        'nomvendedor': ['NOMVENDEDOR', 'VENDEDOR', 'ASESOR', 'COMERCIAL'],
        'numero': ['NUMERO', 'FACTURA', 'DOC'],
        'e_mail': ['EMAIL', 'CORREO', 'E-MAIL', 'MAIL'],
        'cod_cliente': ['CODCLIENTE', 'CODIGO'],
        # Columnas críticas del segundo script para complementar (si existen)
        'serie': ['SERIE', 'ZONA_SERIE'], 
        'fecha_documento': ['FECHADOCUMENTO', 'FECHAFACTURA', 'FECHA_DOC'],
        'fecha_vencimiento': ['FECHAVENCIMIENTO', 'FECHA_VENC'],
        'poblacion': ['POBLACION', 'CIUDAD']
    }
    
    renombres = {}
    for standard, variantes in mapa.items():
        for col in df.columns:
            # Buscar coincidencia parcial, pero priorizar si ya está mapeada
            if standard not in renombres.values() and any(v in col for v in variantes):
                renombres[col] = standard
                break
            
    df.rename(columns=renombres, inplace=True)
    
    # ⚠️ Validación mínima (Nombre, Importe, Días, Número son críticos)
    req = ['nombrecliente', 'importe', 'dias_vencido', 'numero']
    if not all(c in df.columns for c in req):
        missing = [c for c in req if c not in df.columns]
        return None, f"Faltan columnas críticas mapeadas: {', '.join(missing)}."

    # Inclusión de columnas opcionales con valor 'N/A' si faltan (para evitar KeyError)
    for c in mapa.keys():
        if c not in df.columns: df[c] = 'N/A'
        else: df[c] = df[c].fillna('N/A')

    # 2. Conversión de tipos y limpieza de datos
    df['importe'] = pd.to_numeric(df['importe'], errors='coerce').fillna(0) # Limpieza de moneda no es necesaria si el CSV de Dropbox trae numéricos limpios
    df['dias_vencido'] = pd.to_numeric(df['dias_vencido'], errors='coerce').fillna(0).astype(int)
    
    # Asegurar que el importe sea negativo si el número de factura es negativo (Notas Crédito)
    df['numero_float'] = pd.to_numeric(df['numero'], errors='coerce').fillna(0)
    df.loc[df['numero_float'] < 0, 'importe'] *= -1
    df.drop(columns=['numero_float'], inplace=True)
    
    # Normalizar vendedores para filtrado de seguridad
    df['nomvendedor_norm'] = df['nomvendedor'].apply(normalizar_nombre) 
    
    # Convertir fechas (si existen)
    if 'fecha_documento' in df.columns:
        df['fecha_documento'] = pd.to_datetime(df['fecha_documento'], errors='coerce')
    if 'fecha_vencimiento' in df.columns:
        df['fecha_vencimiento'] = pd.to_datetime(df['fecha_vencimiento'], errors='coerce')
    
    # 3. Limpieza final: Quitar facturas con saldo cero
    df = df[df['importe'] != 0].copy()  
    
    # 4. Asignación de Segmentos (del código 2)
    ZONAS_SERIE = { "PEREIRA": [155, 189, 158, 439], "MANIZALES": [157, 238], "ARMENIA": [156] }
    ZONAS_SERIE_STR = {zona: [str(s) for s in series] for zona, series in ZONAS_SERIE.items()}
    def asignar_zona_robusta(valor_serie):
        if pd.isna(valor_serie) or valor_serie == 'N/A': return "OTRAS ZONAS"
        numeros_en_celda = re.findall(r'\d+', str(valor_serie))
        if not numeros_en_celda: return "OTRAS ZONAS"
        for zona, series_clave_str in ZONAS_SERIE_STR.items():
            if set(numeros_en_celda) & set(series_clave_str): return zona
        return "OTRAS ZONAS"
    df['zona'] = df['serie'].apply(asignar_zona_robusta)
    
    # Segmentación estratégica de cartera (del código 1)
    bins = [-float('inf'), 0, 15, 30, 60, 90, float('inf')]
    labels = ["🟢 Al Día", "🟡 Prev. (1-15)", "🟠 Riesgo (16-30)", "🔴 Crítico (31-60)", "🚨 Alto Riesgo (61-90)", "⚫ Legal (+90)"]
    df['Rango'] = pd.cut(df['dias_vencido'], bins=bins, labels=labels, right=True)

    return df, "OK"

@st.cache_data(ttl=600) # Carga automática desde Dropbox (USANDO EL DECORADOR SOLICITADO)
def cargar_datos_automaticos_dropbox():
    """Carga los datos más recientes desde el archivo CSV en Dropbox y los mapea."""
    try:
        # Se asume que las credenciales están en st.secrets
        APP_KEY = st.secrets["dropbox"]["app_key"]
        APP_SECRET = st.secrets["dropbox"]["app_secret"]
        REFRESH_TOKEN = st.secrets["dropbox"]["refresh_token"]

        with dropbox.Dropbox(app_key=APP_KEY, app_secret=APP_SECRET, oauth2_refresh_token=REFRESH_TOKEN) as dbx:
            path_archivo_dropbox = st.secrets["dropbox"].get("file_path", '/data/cartera_detalle.csv') # Usar valor por defecto si no está en secrets
            metadata, res = dbx.files_download(path=path_archivo_dropbox)
            
            # Decodificación y lectura del CSV (asumiendo separador | y encoding latin-1)
            contenido_csv = res.content.decode('latin-1')
            
            # Se usa el primer encabezado para mapear automáticamente, no una lista fija.
            df = pd.read_csv(StringIO(contenido_csv), sep='|', engine='python', dtype=str)

        # Mapear y limpiar el DataFrame
        df_proc, status = mapear_y_limpiar_df(df)
            
        if df_proc is None:  
            return None, f"Error en procesamiento de datos (Dropbox): {status}"
            
        return df_proc, f"Conectado a la fuente principal: **Dropbox ({metadata.name})**"
            
    except toml.TomlDecodeError:
        return None, "Error: Las credenciales de Dropbox no están configuradas correctamente en `secrets.toml`."
    except KeyError as ke:
        return None, f"Error: La clave de Dropbox no se encontró en `secrets.toml`: {ke}"
    except Exception as e:
        return None, f"Error al cargar datos desde Dropbox: {e}"

# ======================================================================================
# 2. INTELIGENCIA DE NEGOCIO (ESTRATEGIA Y FUNCIONES)
# ======================================================================================

def calcular_kpis(df):
    """Calcula los KPIs principales de cobranza."""
    total = df['importe'].sum()
    vencido_df = df[df['dias_vencido'] > 0]
    vencido = vencido_df['importe'].sum()
    pct_vencido = (vencido / total * 100) if total else 0
    clientes_mora = vencido_df['nombrecliente'].nunique()
    
    # Calcular CSI (Collection Severity Index) = Suma(Importe * Días) / Importe Total
    csi = (vencido_df['importe'] * vencido_df['dias_vencido']).sum() / total if total > 0 else 0
    
    # Antigüedad Promedio Vencida
    antiguedad_prom_vencida = (vencido_df['importe'] * vencido_df['dias_vencido']).sum() / vencido if vencido > 0 else 0
    
    return total, vencido, pct_vencido, clientes_mora, csi, antiguedad_prom_vencida

def generar_analisis_cartera(kpis: dict):
    """Genera comentarios de análisis IA basados en KPIs (tomado y mejorado del código 2)."""
    comentarios = []
    
    # 1. Análisis de % Vencido
    if kpis['porcentaje_vencido'] > 30: 
        comentarios.append(f"<li>🔴 **Alerta Crítica (%):** El <b>{kpis['porcentaje_vencido']:.1f}%</b> de la cartera está vencida. Requiere acciones urgentes en todos los frentes.</li>")
    elif kpis['porcentaje_vencido'] > 15: 
        comentarios.append(f"<li>🟡 **Advertencia (%):** Con un <b>{kpis['porcentaje_vencido']:.1f}%</b> de cartera vencida, es prioritario intensificar gestiones en el corto plazo.</li>")
    else: 
        comentarios.append(f"<li>🟢 **Saludable (%):** El porcentaje de cartera vencida (<b>{kpis['porcentaje_vencido']:.1f}%</b>) está en un nivel manejable y eficiente.</li>")
        
    # 2. Análisis de Antigüedad Promedio
    if kpis['antiguedad_prom_vencida'] > 60: 
        comentarios.append(f"<li>🔴 **Riesgo Alto (Días):** Antigüedad promedio de <b>{kpis['antiguedad_prom_vencida']:.0f} días</b>. La deuda está muy envejecida; priorizar clientes con +90 días.</li>")
    elif kpis['antiguedad_prom_vencida'] > 30: 
        comentarios.append(f"<li>🟠 **Atención Requerida (Días):** Antigüedad promedio de <b>{kpis['antiguedad_prom_vencida']:.0f} días</b>. Concentrar esfuerzos en el rango 31-60 para evitar paso a legal.</li>")
    else:
        comentarios.append(f"<li>🟡 **Gestión Preventiva (Días):** La antigüedad es baja (<b>{kpis['antiguedad_prom_vencida']:.0f} días</b>), enfóquese en la gestión *pre-vencimiento* (1-15 días).</li>")
        
    # 3. Análisis de CSI (Severidad)
    if kpis['csi'] > 15: 
        comentarios.append(f"<li>🚨 **Severidad Crítica (CSI: {kpis['csi']:.1f}):** Indica un impacto muy alto. Probablemente hay *clientes muy grandes* con deuda antigua. Focalización extrema.</li>")
    elif kpis['csi'] > 5: 
        comentarios.append(f"<li>🟠 **Severidad Moderada (CSI: {kpis['csi']:.1f}):** Existe riesgo. Hay focos de deuda que, por valor o antigüedad, afectan el indicador.</li>")
    else: 
        comentarios.append(f"<li>🟢 **Severidad Baja (CSI: {kpis['csi']:.1f}):** Impacto bajo, lo que sugiere que la cartera vencida no es excesivamente antigua ni concentrada en grandes montos.</li>")
        
    return "<ul>" + "".join(comentarios) + "</ul>"

def generar_link_wa(telefono, cliente, saldo_vencido, dias_max, nit, cod_cliente):
    """Genera el link de WhatsApp con mensaje pre-cargado."""
    # Limpiar y estandarizar el número (asume Colombia si son 10 dígitos)
    tel = re.sub(r'\D', '', str(telefono))
    # Intentar corregir formato para Colombia (si el número es 10 dígitos, añadir 57)
    if len(tel) == 10 and tel.startswith('3'): tel = '57' + tel 
    if len(tel) < 10: return None
    
    # Tomar solo el primer nombre para un trato más corto
    cliente_corto = str(cliente).split()[0].title()
    portal_link = "https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/"
    
    if saldo_vencido <= 0:
        msg = (
            f"👋 ¡Hola {cliente_corto}! Te saludamos de Ferreinox SAS BIC.\n\n"
            f"¡Felicitaciones! Tu cuenta está al día. Agradecemos tu puntualidad.\n\n"
            f"Te hemos enviado tu estado de cuenta completo a tu correo. ¡Gracias por tu confianza!"
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
    """Clase personalizada para generar PDF con estilos de Ferreinox (Unificado y mejorado)."""
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.set_text_color(0, 51, 102) # Color Primario
        # Intenta usar logo si existe, si no, usa texto
        try:
            # Reemplazar con la ruta real de su logo si lo aloja
            # self.image("LOGO_FERREINOX_SAS_BIC_2024.png", 10, 8, 80) 
            self.cell(80, 10, 'FERREINOX SAS BIC', 0, 0, 'L')
        except RuntimeError: 
            self.cell(80, 10, 'FERREINOX SAS BIC', 0, 0, 'L')
            
        self.set_font('Arial', 'B', 18); self.set_text_color(0, 0, 0)
        self.cell(0, 10, 'ESTADO DE CUENTA', 0, 1, 'R')
        self.set_font('Arial', 'I', 9); self.cell(0, 10, f'Generado el: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}', 0, 1, 'R')
        self.ln(5)

    def footer(self):
        self.set_y(-30)
        self.set_font('Arial', 'I', 8); self.set_text_color(100, 100, 100)
        
        # Enlace Portal de Pagos
        portal_link = "https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/"
        self.set_font('Arial', 'B', 10); self.set_text_color(0, 51, 102)
        self.cell(0, 5, 'Portal de Pagos Ferreinox: ', 0, 1, 'C', link=portal_link)
        self.set_font('Arial', 'I', 8); self.set_text_color(100, 100, 100)
        self.cell(0, 5, f'Página {self.page_no()}', 0, 0, 'C')

def crear_pdf(df_cliente, total_vencido_cliente):
    """Genera el PDF de estado de cuenta detallado por cliente."""
    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=40)
    pdf.add_page()
    
    if df_cliente.empty:
        pdf.set_font('Arial', 'B', 12); pdf.cell(0, 10, 'No se encontraron facturas para este cliente.', 0, 1, 'C')
        return bytes(pdf.output())
        
    row = df_cliente.iloc[0]
    
    # --- Datos Cliente ---
    pdf.set_font("Arial", 'B', 11); pdf.set_text_color(0, 51, 102)
    pdf.cell(40, 6, "Cliente:", 0, 0); pdf.set_font("Arial", '', 11); pdf.set_text_color(0, 0, 0); pdf.cell(0, 6, row['nombrecliente'], 0, 1)
    pdf.set_font("Arial", 'B', 11); pdf.set_text_color(0, 51, 102)
    pdf.cell(40, 6, "NIT/ID:", 0, 0); pdf.set_font("Arial", '', 11); pdf.set_text_color(0, 0, 0); pdf.cell(0, 6, row['nit'], 0, 1)
    pdf.set_font("Arial", 'B', 11); pdf.set_text_color(0, 51, 102)
    pdf.cell(40, 6, "Asesor:", 0, 0); pdf.set_font("Arial", '', 11); pdf.set_text_color(0, 0, 0); pdf.cell(0, 6, row['nomvendedor'], 0, 1)
    pdf.ln(5)
    
    # --- Tabla de Facturas ---
    pdf.set_font('Arial', 'B', 10); pdf.set_fill_color(0, 51, 102); pdf.set_text_color(255, 255, 255)
    pdf.cell(25, 8, "Factura", 1, 0, 'C', 1)
    pdf.cell(25, 8, "Días Mora", 1, 0, 'C', 1)
    pdf.cell(35, 8, "Fecha Doc.", 1, 0, 'C', 1)
    pdf.cell(35, 8, "Fecha Venc.", 1, 0, 'C', 1)
    pdf.cell(40, 8, "Saldo", 1, 1, 'C', 1)
    
    pdf.set_font("Arial", '', 10)
    total_cartera = 0
    for _, item in df_cliente.sort_values(by='dias_vencido', ascending=False).iterrows():
        total_cartera += item['importe']
        
        # Estilo para filas vencidas
        if item['dias_vencido'] > 0:
            pdf.set_fill_color(255, 235, 238) # Fondo rojo claro
            pdf.set_text_color(150, 0, 0)
        else:
            pdf.set_fill_color(255, 255, 255) # Fondo blanco
            pdf.set_text_color(0, 0, 0)
            
        fecha_doc_str = item['fecha_documento'].strftime('%d/%m/%Y') if pd.notna(item['fecha_documento']) else 'N/A'
        fecha_venc_str = item['fecha_vencimiento'].strftime('%d/%m/%Y') if pd.notna(item['fecha_vencimiento']) else 'N/A'
        
        pdf.cell(25, 7, str(item['numero']), 1, 0, 'C', 1)
        pdf.cell(25, 7, str(item['dias_vencido']), 1, 0, 'C', 1)
        pdf.cell(35, 7, fecha_doc_str, 1, 0, 'C', 1)
        pdf.cell(35, 7, fecha_venc_str, 1, 0, 'C', 1)
        pdf.cell(40, 7, f"${item['importe']:,.0f}", 1, 1, 'R', 1)
        
    # --- Totales ---
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", 'B', 11); pdf.set_fill_color(224, 224, 224)
    pdf.cell(120, 8, "TOTAL CARTERA", 1, 0, 'R', 1)
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(40, 8, f"${total_cartera:,.0f}", 1, 1, 'R', 1)

    if total_vencido_cliente > 0:
        pdf.set_font('Arial', 'B', 12); pdf.set_fill_color(255, 204, 204); pdf.set_text_color(192, 0, 0)
        pdf.cell(120, 8, 'VALOR TOTAL VENCIDO A PAGAR', 1, 0, 'R', 1)
        pdf.cell(40, 8, f"${total_vencido_cliente:,.0f}", 1, 1, 'R', 1)
            
    return bytes(pdf.output())

def crear_excel_gerencial(df, total, vencido, pct_mora, clientes_mora, csi, antiguedad_prom_vencida):
    """Genera el reporte ejecutivo en Excel con estilos y fórmulas (tomado del código 2)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumen Gerencial"
    
    # Estilos
    header_style = Font(bold=True, color="FFFFFF")
    fill_blue = PatternFill("solid", fgColor="003865")
    fill_kpi = PatternFill("solid", fgColor="FFC300")
    
    # --- KPIs en Excel ---
    ws['A1'] = "REPORTE GERENCIAL DE CARTERA - FERREINOX"
    ws['A1'].font = Font(size=16, bold=True)
    
    kpi_labels = ["Total Cartera", "Total Vencido", "% Mora", "Clientes en Mora", "Antigüedad Prom. Vencida", "Índice de Severidad (CSI)"]
    kpi_values = [total, vencido, pct_mora / 100, clientes_mora, antiguedad_prom_vencida, csi]
    formats = ['$#,##0', '$#,##0', '0.0%', '0', '0.0', '0.0']
    
    for i, (lab, val, fmt) in enumerate(zip(kpi_labels, kpi_values, formats)):
        col_letter = get_column_letter(i+1)
        c_lab = ws.cell(row=3, column=i+1, value=lab)
        c_lab.font = Font(bold=True); c_lab.fill = fill_blue; c_lab.alignment = Alignment(horizontal='center')
        ws.column_dimensions[col_letter].width = 20
        
        c_val = ws.cell(row=4, column=i+1, value=val)
        c_val.number_format = fmt
        c_val.font = Font(bold=True, color=COLOR_PRIMARIO); c_val.fill = fill_kpi
        c_val.alignment = Alignment(horizontal='center')

    # --- Tabla Detalle ---
    ws['A6'] = "DETALLE COMPLETO DE LA CARTERA (Filtrable)"
    ws['A6'].font = Font(size=12, bold=True)
    
    # Columnas a incluir en el reporte
    cols = ['nombrecliente', 'nit', 'numero', 'nomvendedor', 'cod_cliente', 'Rango', 'zona', 'dias_vencido', 'importe', 'telefono1', 'e_mail']
    df_detalle = df[cols].sort_values(by='dias_vencido', ascending=False).reset_index(drop=True)

    # Headers de la tabla (fila 7)
    for col_num, col_name in enumerate(cols, 1):
        c = ws.cell(row=7, column=col_num, value=col_name.upper().replace('_', ' '))
        c.fill = fill_blue
        c.font = header_style
        
    # Data (a partir de fila 8)
    for row_num, row_data in enumerate(df_detalle.values, 8):
        for col_num, val in enumerate(row_data, 1):
            c = ws.cell(row=row_num, column=col_num, value=val)
            if col_num == 9: c.number_format = '$#,##0' # Columna Saldo
            
    # Autoajuste de columnas y filtros
    ws.auto_filter.ref = f"A7:{get_column_letter(len(cols))}{len(df_detalle)+7}"
    for i in range(1, len(cols) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 20 if i != 1 else 35
        
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ======================================================================================
# 4. FUNCIÓN DE ENVÍO DE CORREO (yagmail) Y PLANTILLAS HTML
# ======================================================================================
def enviar_correo(destinatario, asunto, cuerpo_html, pdf_bytes):
    """Función para enviar correo con el PDF adjunto, usando yagmail y st.secrets."""
    try:
        email_user = st.secrets["email_credentials"]["sender_email"]
        email_pass = st.secrets["email_credentials"]["sender_password"]
    except KeyError:
        st.error("⚠️ Configura las credenciales de correo (sender_email y sender_password) en `secrets.toml` antes de enviar.")
        return False

    if not email_user or not email_pass:
        st.error("⚠️ Credenciales de correo incompletas. Revisa `secrets.toml`.")
        return False

    # VALIDACIÓN BÁSICA DE DESTINATARIO
    if not destinatario or '@' not in destinatario:
        st.error("⚠️ El correo electrónico del destinatario no es válido.")
        return False

    try:
        # Guardar PDF temporalmente
        tmp_path = ''
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(pdf_bytes)
            tmp_path = tmp.name

        with st.spinner(f"Enviando correo a {destinatario}..."):
            # Conexión con yagmail
            yag = yagmail.SMTP(email_user, email_pass)
            
            yag.send(
                to=destinatario,
                subject=asunto,
                contents=[cuerpo_html, tmp_path] # Adjunta el PDF
            )
        
        os.remove(tmp_path) # Limpiar el archivo temporal
        return True
    except Exception as e:
        st.error(f"Error enviando correo. Asegúrate que la contraseña es una 'Contraseña de Aplicación' (no la normal) y que el remitente está configurado: {e}")
        try:
            if os.path.exists(tmp_path): os.remove(tmp_path)
        except:
            pass
        return False
        
# --- PLANTILLAS HTML PROFESIONALES (Tomadas del código 2) ---

def plantilla_correo_vencido(cliente, saldo, dias, nit, cod_cliente, portal_link):
    """Plantilla de correo para clientes con deuda vencida."""
    # Uso de HTML complejo para mayor compatibilidad con clientes de correo (MJML-style)
    # El contenido se mantiene idéntico al código 2, asegurando la profesionalidad
    dias_max_vencido = int(dias)
    return f"""
    <!doctype html><html xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><head><title>Recordatorio Amistoso de Saldo Vencido - Ferreinox</title><meta http-equiv="X-UA-Compatible" content="IE=edge"><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style type="text/css">#outlook a {{ padding:0; }}
    body {{ margin:0;padding:0;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%; }}
    table, td {{ border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt; }}
    img {{ border:0;height:auto;line-height:100%; outline:none;text-decoration:none;-ms-interpolation-mode:bicubic; }}
    p {{ display:block;margin:13px 0; }}</style><link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet" type="text/css"><style type="text/css">@import url(https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap);</style><style type="text/css">@media only screen and (min-width:480px) {{
        .mj-column-per-100 {{ width:100% !important; max-width: 100%; }}
        .mj-column-per-50 {{ width:50% !important; max-width: 50%; }}
        }}</style><style media="screen and (min-width:480px)">.moz-text-html .mj-column-per-100 {{ width:100% !important; max-width: 100%; }}
        .moz-text-html .mj-column-per-50 {{ width:50% !important; max-width: 50%; }}</style><style type="text/css"></style><style type="text/css">.greeting-strong {{
        color: #1e40af;
        font-weight: 600;
        }}
        .whatsapp-button table {{
        width: 100% !important;
        }}</style></head><body style="word-spacing:normal;background-color:#f3f4f6;"><div style="background-color:#f3f4f6;"><div class="email-container" style="background:#FFFFFF;background-color:#FFFFFF;margin:0px auto;border-radius:24px;max-width:600px;"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#FFFFFF;background-color:#FFFFFF;width:100%;border-radius:24px;"><tbody><tr><td style="direction:ltr;font-size:0px;padding:0;text-align:center;"><div style="background:#1e3a8a;background-color:#1e3a8a;margin:0px auto;max-width:600px;"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#1e3a8a;background-color:#1e3a8a;width:100%;"><tbody><tr><td style="direction:ltr;font-size:0px;padding:30px 30px;text-align:center;"><div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="vertical-align:top;" width="100%"><tbody><tr><td align="center" style="font-size:0px;padding:10px 25px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:28px;font-weight:700;line-height:1.6;text-align:center;color:#ffffff;">Recordatorio de Saldo Pendiente</div></td></tr></tbody></table></div></td></tr></tbody></table></div><div style="background:#ffffff;background-color:#ffffff;margin:0px auto;max-width:600px;"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#ffffff;background-color:#ffffff;width:100%;"><tbody><tr><td style="direction:ltr;font-size:0px;padding:40px 40px 20px 40px;text-align:center;"><div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="vertical-align:top;" width="100%"><tbody><tr><td align="left" style="font-size:0px;padding:10px 25px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:18px;font-weight:500;line-height:1.6;text-align:left;color:#374151;">Hola, <span class="greeting-strong">{cliente}</span> 👋</div></td></tr><tr><td align="left" style="font-size:0px;padding:10px 25px;padding-bottom:20px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:16px;line-height:1.6;text-align:left;color:#6b7280;">Te contactamos de parte de <strong>Ferreinox SAS BIC</strong> para recordarte amablemente sobre tu estado de cuenta. Hemos identificado un saldo vencido y te invitamos a revisarlo.</div></td></tr><tr><td align="center" style="font-size:0px;padding:10px 0;word-break:break-word;"><p style="border-top:solid 2px #3b82f6;font-size:1px;margin:0px auto;width:100%;"></p></td></tr></tbody></table></div></td></tr></tbody></table></div><div style="background:#ffffff;background-color:#ffffff;margin:0px auto;max-width:600px;"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#ffffff;background-color:#ffffff;width:100%;"><tbody><tr><td style="direction:ltr;font-size:0px;padding:10px 40px;text-align:center;"><div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="background-color:#fee2e2;border-radius:20px;vertical-align:top;" width="100%"><tbody><tr><td align="center" style="font-size:0px;padding:25px 0 10px 0;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:48px;line-height:1.6;text-align:center;color:#374151;">⚠️</div></td></tr><tr><td align="center" style="font-size:0px;padding:10px 25px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:24px;font-weight:700;line-height:1.6;text-align:center;color:#991b1b;">Valor Total Vencido</div></td></tr><tr><td align="center" style="font-size:0px;padding:5px 25px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:40px;font-weight:700;line-height:1.6;text-align:center;color:#991b1b;">${saldo:,.0f}</div></td></tr><tr><td align="center" style="font-size:0px;padding:5px 25px 30px 25px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:16px;line-height:1.6;text-align:center;color:#b91c1c;">Tu factura más antigua tiene <strong>{dias_max_vencido} días</strong> de vencimiento.</div></td></tr></tbody></table></div></td></tr></tbody></table></div><div style="background:#ffffff;background-color:#ffffff;margin:0px auto;max-width:600px;"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#ffffff;background-color:#ffffff;width:100%;"><tbody><tr><td style="direction:ltr;font-size:0px;padding:20px 40px;text-align:center;"><div class="mj-column-per-50 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:middle;width:100%;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="background-color:#f8fafc;border-radius:16px;vertical-align:middle;" width="100%"><tbody><tr><td align="left" style="font-size:0px;padding:20px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:16px;font-weight:700;line-height:1.2;text-align:left;color:#334155;">NIT/CC</div><div style="font-family:Inter, -apple-system, sans-serif;font-size:20px;font-weight:700;line-height:1.2;text-align:left;color:#1e293b;">{nit}</div></td></tr><tr><td align="left" style="font-size:0px;padding:20px;padding-top:0;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:16px;font-weight:700;line-height:1.2;text-align:left;color:#334155;">CÓDIGO INTERNO</div><div style="font-family:Inter, -apple-system, sans-serif;font-size:20px;font-weight:700;line-height:1.2;text-align:left;color:#1e293b;">{cod_cliente}</div></td></tr></tbody></table></div><div class="mj-column-per-50 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:middle;width:100%;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="vertical-align:middle;" width="100%"><tbody><tr><td align="center" style="font-size:0px;padding:10px 25px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:16px;font-weight:500;line-height:1.6;text-align:center;color:#475569;">Usa estos datos en nuestro portal de pagos.</div></td></tr><tr><td align="center" vertical-align="middle" style="font-size:0px;padding:10px 25px;word-break:break-word;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse:separate;line-height:100%;"><tr><td align="center" bgcolor="#16a34a" role="presentation" style="border:none;border-radius:12px;cursor:auto;mso-padding-alt:16px 25px;background:#16a34a;" valign="middle"><a href="{portal_link}" style="display:inline-block;background:#16a34a;color:#ffffff;font-family:Inter, -apple-system, sans-serif;font-size:16px;font-weight:600;line-height:120%;margin:0;text-decoration:none;text-transform:none;padding:16px 25px;mso-padding-alt:0px;border-radius:12px;" target="_blank">🚀 Realizar Pago</a></td></tr></table></td></tr></tbody></table></div></td></tr></tbody></table></div><div style="background:#ffffff;background-color:#ffffff;margin:0px auto;max-width:600px;"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#ffffff;background-color:#ffffff;width:100%;"><tbody><tr><td style="direction:ltr;font-size:0px;padding:20px 40px;text-align:center;"><div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%"><tbody><tr><td style="background-color:#f8fafc;border-left:5px solid #3b82f6;border-radius:16px;vertical-align:top;padding:20px;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%"><tbody><tr><td align="left" style="font-size:0px;padding:10px 25px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:16px;font-weight:500;line-height:1.6;text-align:left;color:#475569;">💡 <strong>Nota:</strong> Si ya realizaste el pago, por favor omite este mensaje. Para tu control, hemos adjuntado tu estado de cuenta en PDF.</div></td></tr></tbody></table></td></tr></tbody></table></div></td></tr></tbody></table></div><div style="background:#1f2937;background-color:#1f2937;margin:0px auto;max-width:600px;"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#1f2937;background-color:#1f2937;width:100%;"><tbody><tr><td style="direction:ltr;font-size:0px;padding:30px;text-align:center;"><div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="vertical-align:top;" width="100%"><tbody><tr><td align="center" style="font-size:0px;padding:10px 25px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:18px;font-weight:600;line-height:1.6;text-align:center;color:#ffffff;">Área de Cartera y Recaudos</div></td></tr><tr><td align="center" style="font-size:0px;padding:10px 25px;padding-bottom:20px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:16px;line-height:1.6;text-align:center;color:#e5e7eb;"><strong>Líneas de Atención WhatsApp</strong></div></td></tr><tr><td align="center" vertical-align="middle" class="whatsapp-button" style="font-size:0px;padding:10px 25px;word-break:break-word;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse:separate;line-height:100%;"><tr><td align="center" bgcolor="#25d366" role="presentation" style="border:none;border-radius:12px;cursor:auto;mso-padding-alt:10px 25px;background:#25d366;" valign="middle"><a href="https://wa.me/573165219904" style="display:inline-block;background:#25d366;color:#ffffff;font-family:Inter, -apple-system, sans-serif;font-size:13px;font-weight:500;line-height:120%;margin:0;text-decoration:none;text-transform:none;padding:10px 25px;mso-padding-alt:0px;border-radius:12px;" target="_blank">📱 Armenia: 316 5219904</a></td></tr></table></td></tr><tr><td align="center" vertical-align="middle" class="whatsapp-button" style="font-size:0px;padding:10px 25px;padding-top:12px;word-break:break-word;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse:separate;line-height:100%;"><tr><td align="center" bgcolor="#25d366" role="presentation" style="border:none;border-radius:12px;cursor:auto;mso-padding-alt:10px 25px;background:#25d366;" valign="middle"><a href="https://wa.me/573108501359" style="display:inline-block;background:#25d366;color:#ffffff;font-family:Inter, -apple-system, sans-serif;font-size:13px;font-weight:500;line-height:120%;margin:0;text-decoration:none;text-transform:none;padding:10px 25px;mso-padding-alt:0px;border-radius:12px;" target="_blank">📱 Manizales: 310 8501359</a></td></tr></table></td></tr><tr><td align="center" vertical-align="middle" class="whatsapp-button" style="font-size:0px;padding:10px 25px;padding-top:12px;word-break:break-word;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse:separate;line-height:100%;"><tr><td align="center" bgcolor="#25d366" role="presentation" style="border:none;border-radius:12px;cursor:auto;mso-padding-alt:10px 25px;background:#25d366;" valign="middle"><a href="https://wa.me/573142087169" style="display:inline-block;background:#25d366;color:#ffffff;font-family:Inter, -apple-system, sans-serif;font-size:13px;font-weight:500;line-height:120%;margin:0;text-decoration:none;text-transform:none;padding:10px 25px;mso-padding-alt:0px;border-radius:12px;" target="_blank">📱 Pereira: 314 2087169</a></td></tr></table></td></tr><tr><td align="center" style="font-size:0px;padding:30px 0 20px 0;word-break:break-word;"><p style="border-top:solid 1px #4b5563;font-size:1px;margin:0px auto;width:100%;"></p></td></tr><tr><td align="center" style="font-size:0px;padding:10px 25px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:14px;line-height:1.6;text-align:center;color:#9ca3af;">© 2025 Ferreinox SAS BIC - Todos los derechos reservados</div></td></tr></tbody></table></div></td></tr></tbody></table></div></td></tr></tbody></table></div></div></body></html>
    """

def plantilla_correo_al_dia(cliente, total_cartera):
    """Plantilla de correo para clientes con cuenta al día."""
    # Uso de HTML complejo para mayor compatibilidad con clientes de correo (MJML-style)
    # El contenido se mantiene idéntico al código 2, asegurando la profesionalidad
    return f"""
    <!doctype html><html xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><head><title>Tu Estado de Cuenta Actualizado - Ferreinox</title><meta http-equiv="X-UA-Compatible" content="IE=edge"><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style type="text/css">#outlook a {{ padding:0; }}
    body {{ margin:0;padding:0;-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%; }}
    table, td {{ border-collapse:collapse;mso-table-lspace:0pt;mso-table-rspace:0pt; }}
    img {{ border:0;height:auto;line-height:100%; outline:none;text-decoration:none;-ms-interpolation-mode:bicubic; }}
    p {{ display:block;margin:13px 0; }}</style><link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet" type="text/css"><style type="text/css">@import url(https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap);</style><style type="text/css">@media only screen and (min-width:480px) {{
        .mj-column-per-100 {{ width:100% !important; max-width: 100%; }}
        }}</style><style media="screen and (min-width:480px)">.moz-text-html .mj-column-per-100 {{ width:100% !important; max-width: 100%; }}</style><style type="text/css"></style><style type="text/css">.greeting-strong {{
        color: #1e40af;
        font-weight: 600;
        }}
        .whatsapp-button table {{
        width: 100% !important;
        }}</style></head><body style="word-spacing:normal;background-color:#f3f4f6;"><div style="background-color:#f3f4f6;"><div class="email-container" style="background:#FFFFFF;background-color:#FFFFFF;margin:0px auto;border-radius:24px;max-width:600px;"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#FFFFFF;background-color:#FFFFFF;width:100%;border-radius:24px;"><tbody><tr><td style="direction:ltr;font-size:0px;padding:0;text-align:center;"><div style="background:#1e3a8a;background-color:#1e3a8a;margin:0px auto;max-width:600px;"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#1e3a8a;background-color:#1e3a8a;width:100%;"><tbody><tr><td style="direction:ltr;font-size:0px;padding:30px 30px;text-align:center;"><div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="vertical-align:top;" width="100%"><tbody><tr><td align="center" style="font-size:0px;padding:10px 25px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:28px;font-weight:700;line-height:1.6;text-align:center;color:#ffffff;">Estado de Cuenta Actualizado</div></td></tr></tbody></table></div></td></tr></tbody></table></div><div style="background:#ffffff;background-color:#ffffff;margin:0px auto;max-width:600px;"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#ffffff;background-color:#ffffff;width:100%;"><tbody><tr><td style="direction:ltr;font-size:0px;padding:40px 40px 20px 40px;text-align:center;"><div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="vertical-align:top;" width="100%"><tbody><tr><td align="left" style="font-size:0px;padding:10px 25px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:18px;font-weight:500;line-height:1.6;text-align:left;color:#374151;">Hola, <span class="greeting-strong">{cliente}</span> ✨</div></td></tr><tr><td align="left" style="font-size:0px;padding:10px 25px;padding-bottom:20px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:16px;line-height:1.6;text-align:left;color:#6b7280;">Recibe un cordial saludo del equipo de <strong>Ferreinox SAS BIC</strong>.</div></td></tr><tr><td align="center" style="font-size:0px;padding:10px 0;word-break:break-word;"><p style="border-top:solid 2px #3b82f6;font-size:1px;margin:0px auto;width:100%;"></p></td></tr></tbody></table></div></td></tr></tbody></table></div><div style="background:#ffffff;background-color:#ffffff;margin:0px auto;max-width:600px;"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#ffffff;background-color:#ffffff;width:100%;"><tbody><tr><td style="direction:ltr;font-size:0px;padding:10px 40px;text-align:center;"><div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="background-color:#10b981;border-radius:20px;vertical-align:top;" width="100%"><tbody><tr><td align="center" style="font-size:0px;padding:25px 0 10px 0;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:48px;line-height:1.6;text-align:center;color:#374151;">🎉</div></td></tr><tr><td align="center" style="font-size:0px;padding:10px 25px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:24px;font-weight:700;line-height:1.6;text-align:center;color:#ffffff;">¡Felicitaciones!</div></td></tr><tr><td align="center" style="font-size:0px;padding:5px 25px 30px 25px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:16px;line-height:1.6;text-align:center;color:#ffffff;">Tu cuenta no presenta saldos vencidos.<br>Agradecemos enormemente tu puntualidad y excelente gestión de pagos.</div></td></tr></tbody></table></div></td></tr></tbody></table></div><div style="background:#ffffff;background-color:#ffffff;margin:0px auto;max-width:600px;"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#ffffff;background-color:#ffffff;width:100%;"><tbody><tr><td style="direction:ltr;font-size:0px;padding:20px 40px;text-align:center;"><div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%"><tbody><tr><td style="background-color:#f8fafc;border-left:5px solid #3b82f6;border-radius:16px;vertical-align:top;padding:20px;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" width="100%"><tbody><tr><td align="left" style="font-size:0px;padding:10px 25px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:16px;font-weight:500;line-height:1.6;text-align:left;color:#475569;">📄 Para tu control y referencia, hemos adjuntado tu estado de cuenta completo en formato PDF a este correo electrónico.</div></td></tr></tbody></table></td></tr></tbody></table></div></td></tr></tbody></table></div><div style="background:#1f2937;background-color:#1f2937;margin:0px auto;max-width:600px;"><table align="center" border="0" cellpadding="0" cellspacing="0" role="presentation" style="background:#1f2937;background-color:#1f2937;width:100%;"><tbody><tr><td style="direction:ltr;font-size:0px;padding:30px;text-align:center;"><div class="mj-column-per-100 mj-outlook-group-fix" style="font-size:0px;text-align:left;direction:ltr;display:inline-block;vertical-align:top;width:100%;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="vertical-align:top;" width="100%"><tbody><tr><td align="center" style="font-size:0px;padding:10px 25px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:18px;font-weight:600;line-height:1.6;text-align:center;color:#ffffff;">Área de Cartera y Recaudos</div></td></tr><tr><td align="center" style="font-size:0px;padding:10px 25px;padding-bottom:20px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:16px;line-height:1.6;text-align:center;color:#e5e7eb;"><strong>Líneas de Atención WhatsApp</strong></div></td></tr><tr><td align="center" vertical-align="middle" class="whatsapp-button" style="font-size:0px;padding:10px 25px;word-break:break-word;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse:separate;line-height:100%;"><tr><td align="center" bgcolor="#25d366" role="presentation" style="border:none;border-radius:12px;cursor:auto;mso-padding-alt:10px 25px;background:#25d366;" valign="middle"><a href="https://wa.me/573165219904" style="display:inline-block;background:#25d366;color:#ffffff;font-family:Inter, -apple-system, sans-serif;font-size:13px;font-weight:500;line-height:120%;margin:0;text-decoration:none;text-transform:none;padding:10px 25px;mso-padding-alt:0px;border-radius:12px;" target="_blank">📱 Armenia: 316 5219904</a></td></tr></table></td></tr><tr><td align="center" vertical-align="middle" class="whatsapp-button" style="font-size:0px;padding:10px 25px;padding-top:12px;word-break:break-word;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse:separate;line-height:100%;"><tr><td align="center" bgcolor="#25d366" role="presentation" style="border:none;border-radius:12px;cursor:auto;mso-padding-alt:10px 25px;background:#25d366;" valign="middle"><a href="https://wa.me/573108501359" style="display:inline-block;background:#25d366;color:#ffffff;font-family:Inter, -apple-system, sans-serif;font-size:13px;font-weight:500;line-height:120%;margin:0;text-decoration:none;text-transform:none;padding:10px 25px;mso-padding-alt:0px;border-radius:12px;" target="_blank">📱 Manizales: 310 8501359</a></td></tr></table></td></tr><tr><td align="center" vertical-align="middle" class="whatsapp-button" style="font-size:0px;padding:10px 25px;padding-top:12px;word-break:break-word;"><table border="0" cellpadding="0" cellspacing="0" role="presentation" style="border-collapse:separate;line-height:100%;"><tr><td align="center" bgcolor="#25d366" role="presentation" style="border:none;border-radius:12px;cursor:auto;mso-padding-alt:10px 25px;background:#25d366;" valign="middle"><a href="https://wa.me/573142087169" style="display:inline-block;background:#25d366;color:#ffffff;font-family:Inter, -apple-system, sans-serif;font-size:13px;font-weight:500;line-height:120%;margin:0;text-decoration:none;text-transform:none;padding:10px 25px;mso-padding-alt:0px;border-radius:12px;" target="_blank">📱 Pereira: 314 2087169</a></td></tr></table></td></tr><tr><td align="center" style="font-size:0px;padding:30px 0 20px 0;word-break:break-word;"><p style="border-top:solid 1px #4b5563;font-size:1px;margin:0px auto;width:100%;"></p></td></tr><tr><td align="center" style="font-size:0px;padding:10px 25px;word-break:break-word;"><div style="font-family:Inter, -apple-system, sans-serif;font-size:14px;line-height:1.6;text-align:center;color:#9ca3af;">© 2025 Ferreinox SAS BIC - Todos los derechos reservados</div></td></tr></tbody></table></div></td></tr></tbody></table></div></td></tr></tbody></table></div></div></body></html>
    """
    
# ======================================================================================
# 5. DASHBOARD PRINCIPAL (MAIN)
# ======================================================================================

def main():
    # --- AUTENTICACIÓN (Tomado del código 2) ---
    if 'authentication_status' not in st.session_state:
        st.session_state['authentication_status'] = False
        st.session_state['acceso_general'] = False
        st.session_state['vendedor_autenticado'] = None
        st.session_state['email_destino_temp'] = ''
        st.session_state['whatsapp_destino_temp'] = ''

    if not st.session_state['authentication_status']:
        st.title("🔐 Acceso al Centro de Mando: Cobranza Ferreinox")
        try:
            general_password = st.secrets["general"]["password"]
            vendedores_secrets = st.secrets["vendedores"]
        except Exception as e:
            st.error(f"Error al cargar las contraseñas desde `secrets.toml`: {e}. Por favor, verifique su configuración.")
            st.stop()
            
        password = st.text_input("Introduce la contraseña:", type="password", key="password_input")
        if st.button("Ingresar"):
            
            if password == str(general_password):
                st.session_state['authentication_status'] = True
                st.session_state['acceso_general'] = True
                st.session_state['vendedor_autenticado'] = "GERENTE_GENERAL"
                st.rerun()
            else:
                for vendedor_key, pass_vendedor in vendedores_secrets.items():
                    if password == str(pass_vendedor):
                        st.session_state['authentication_status'] = True
                        st.session_state['acceso_general'] = False
                        st.session_state['vendedor_autenticado'] = vendedor_key
                        st.rerun()
                        break
                if not st.session_state['authentication_status']:
                    st.error("Contraseña incorrecta. Intente de nuevo.")
        st.stop()
        
    # --- LÓGICA DE LA APP (Una vez autenticado) ---
    st.title("🛡️ Centro de Mando: Cobranza Ferreinox PRO")
    
    # --- BARRA LATERAL: CONFIGURACIÓN Y FILTROS ---
    with st.sidebar:
        st.header("👤 Sesión y Control")
        st.success(f"Usuario: **{st.session_state['vendedor_autenticado']}**")
        if st.button("Cerrar Sesión"):
            for key in list(st.session_state.keys()): del st.session_state[key]
            st.rerun()
            
        st.divider()
        if st.button("🔄 Recargar Datos (Dropbox)", type="primary"):
            st.cache_data.clear()
            st.success("Caché limpiado. Recargando datos de Dropbox...")
            st.rerun()

    # --- CARGA DE DATOS ---
    df, status_carga = cargar_datos_automaticos_dropbox()
    st.caption(status_carga)

    if df is None:
        st.error("🚨 No se pudieron cargar datos funcionales. Revise las credenciales de Dropbox y el formato del archivo.")
        st.stop()
        
    # --- FILTROS DINÁMICOS ---
    st.sidebar.header("🔍 Filtros Operativos")
    
    # 1. Filtro Vendedor (General puede ver todos)
    if st.session_state['acceso_general']:
        vendedores_disponibles = ["TODOS"] + sorted(df['nomvendedor'].unique().tolist())
        filtro_vendedor = st.sidebar.selectbox("Filtrar por Vendedor:", vendedores_disponibles)
        if filtro_vendedor != "TODOS":
            # Filtrar por el nombre normalizado para mayor seguridad
            df_view = df[df['nomvendedor_norm'] == normalizar_nombre(filtro_vendedor)].copy()
        else:
            df_view = df.copy()
    else:
        # Vendedor solo ve su cartera
        vendedor_actual_norm = normalizar_nombre(st.session_state['vendedor_autenticado'])
        df_view = df[df['nomvendedor_norm'] == vendedor_actual_norm].copy()
        st.sidebar.info(f"Vista: Solo mi Cartera ({st.session_state['vendedor_autenticado']})")
        
    # 2. Filtro Rango de Antigüedad
    rangos_cartera = ["TODOS"] + df['Rango'].cat.categories.tolist()
    filtro_rango = st.sidebar.selectbox("Filtrar por Antigüedad:", rangos_cartera)
    if filtro_rango != "TODOS":
        df_view = df_view[df_view['Rango'] == filtro_rango]

    # 3. Filtro Zona
    zonas_disponibles = ["TODAS LAS ZONAS"] + sorted(df_view['zona'].unique().tolist())
    filtro_zona = st.sidebar.selectbox("Filtrar por Zona:", zonas_disponibles)
    if filtro_zona != "TODAS LAS ZONAS":
        df_view = df_view[df_view['zona'] == filtro_zona]

    if df_view.empty:
        st.warning("No hay datos para la selección actual de filtros.")
        st.stop() 

    # --- CÁLCULO DE KPIS CON DATOS FILTRADOS ---
    total, vencido, pct_mora, clientes_mora, csi, antiguedad_prom_vencida = calcular_kpis(df_view)

    # --- ENCABEZADO Y KPIS ---
    st.header("Indicadores Clave de Rendimiento (KPIs)")
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("💰 Cartera Total", f"${total:,.0f}")
    k2.metric("🔥 Cartera Vencida", f"${vencido:,.0f}")
    k3.metric("📈 % Vencido s/ Total", f"{pct_mora:.1f}%", delta=f"{pct_mora - 10:.1f}%" if pct_mora > 10 else "N/A") # Delta simulado
    k4.metric("👥 Clientes en Mora", clientes_mora)
    k5.metric("⏳ Antigüedad Prom.", f"{antiguedad_prom_vencida:.0f} días")
    k6.metric("💥 Índice de Severidad (CSI)", f"{csi:,.1f}")
    
    # Análisis IA (mejorado del código 2)
    with st.expander("🤖 **Análisis y Recomendaciones del Asistente IA**", expanded=pct_mora > 15):
        kpis_dict = {'porcentaje_vencido': pct_mora, 'antiguedad_prom_vencida': antiguedad_prom_vencida, 'csi': csi}
        analisis = generar_analisis_cartera(kpis_dict)
        st.markdown(analisis, unsafe_allow_html=True)
    
    st.divider()

    # --- TABS DE GESTIÓN ---
    tab_lider, tab_gerente, tab_datos = st.tabs(["👩‍💼 GESTIÓN OPERATIVA (1 a 1)", "👨‍💼 ANÁLISIS GERENCIAL", "📥 EXPORTAR Y DATOS"])

    # ==============================================================================
    # TAB LÍDER: GESTIÓN DE COBRO 1 A 1
    # ==============================================================================
    with tab_lider:
        st.subheader("🎯 Módulo de Contacto Directo y Envío de Docs.")
        
        # Agrupar por Cliente para gestión (solo clientes con saldo > 0)
        df_agrupado = df_view[df_view['importe'] > 0].groupby('nombrecliente').agg(
            saldo=('importe', 'sum'),
            saldo_vencido=('importe', lambda x: x[df_view.loc[x.index, 'dias_vencido'] > 0].sum()),
            dias_max=('dias_vencido', 'max'),
            telefono=('telefono1', 'first'),
            email=('e_mail', 'first'),
            vendedor=('nomvendedor', 'first'),
            nit=('nit', 'first'),
            cod_cliente=('cod_cliente', 'first')
        ).reset_index().sort_values('saldo_vencido', ascending=False)
        
        clientes_a_mostrar = df_agrupado['nombrecliente'].tolist()
        
        # Selector de cliente
        cliente_sel = st.selectbox("🔍 Selecciona Cliente a Gestionar (Priorizado por Deuda Vencida)", 
                                 [""] + clientes_a_mostrar, 
                                 format_func=lambda x: '--- Selecciona un cliente ---' if x == "" else x)
        
        if cliente_sel:
            data_cli = df_agrupado[df_agrupado['nombrecliente'] == cliente_sel].iloc[0]
            detalle_facturas = df_view[df_view['nombrecliente'] == cliente_sel].sort_values('dias_vencido', ascending=False)
            
            saldo_vencido_cli = data_cli['saldo_vencido']
            
            # Limpieza de datos
            email_cli = data_cli['email'] if data_cli['email'] not in ['N/A', '', None] else 'Correo no disponible'
            # Asumiendo que telefono es un string con posible punto decimal
            telefono_raw = str(data_cli['telefono']).split('.')[0].strip()
            telefono_cli = f"+57{re.sub(r'\D', '', telefono_raw)}" if len(re.sub(r'\D', '', telefono_raw)) == 10 else telefono_raw
            
            c1, c2 = st.columns([1, 2])
            
            with c1:
                st.markdown(f"#### Resumen de Cliente: **{cliente_sel}**")
                st.info(f"**Deuda Total:** ${data_cli['saldo']:,.0f}")
                st.markdown(f'<div style="background-color: #fee2e2; border-left: 5px solid {COLOR_ALERTA_CRITICA}; padding: 10px; border-radius: 5px;">'
                            f'**Deuda Vencida:** ${saldo_vencido_cli:,.0f}'
                            f'</div>', unsafe_allow_html=True)
                st.error(f"**Días Máx Mora:** {int(data_cli['dias_max'])} días")
                st.text(f"📞 {telefono_cli} | 📧 {email_cli}")
                st.text(f"ID: {data_cli['nit']} | Cód. Cliente: {data_cli['cod_cliente']}")
                
                # Generar PDF en memoria
                pdf_bytes = crear_pdf(detalle_facturas, saldo_vencido_cli)
                
                # --- BOTÓN WHATSAPP ---
                link_wa = generar_link_wa(telefono_cli, cliente_sel, saldo_vencido_cli, data_cli['dias_max'], data_cli['nit'], data_cli['cod_cliente'])
                if link_wa and len(re.sub(r'\D', '', link_wa)) >= 10:
                    st.markdown(f"""<a href="{link_wa}" target="_blank" class="wa-link">📱 ABRIR WHATSAPP CON GUION</a>""", unsafe_allow_html=True)
                else:
                    st.error("Número de teléfono inválido para WhatsApp")
                
                st.download_button(label="📄 Descargar PDF Local", data=pdf_bytes, file_name=f"Estado_Cuenta_{normalizar_nombre(cliente_sel).replace(' ', '_')}.pdf", mime="application/pdf")

            with c2:
                st.write("#### 📄 Detalle de Facturas (Priorizadas por Mora)")
                # Vista previa de facturas
                st.dataframe(detalle_facturas[['numero', 'fecha_vencimiento', 'dias_vencido', 'importe', 'Rango']].style.format({'importe': '${:,.0f}'}).background_gradient(subset=['dias_vencido'], cmap='YlOrRd'), height=250, use_container_width=True, hide_index=True)
                
                # --- ENVÍO DE CORREO ---
                st.write("#### 📧 Envío de Estado de Cuenta por Correo")
                with st.form("form_email"):
                    email_dest = st.text_input("Destinatario", value=email_cli, key="email_dest_input")
                    
                    if saldo_vencido_cli > 0:
                        asunto_msg = f"Recordatorio URGENTE de Saldo Pendiente - {cliente_sel}"
                        portal_link_email = "https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/"
                        cuerpo_html = plantilla_correo_vencido(cliente_sel, saldo_vencido_cli, data_cli['dias_max'], data_cli['nit'], data_cli['cod_cliente'], portal_link_email)
                    else:
                        asunto_msg = f"Tu Estado de Cuenta Actualizado - {cliente_sel} (Cta al Día)"
                        cuerpo_html = plantilla_correo_al_dia(cliente_sel, data_cli['saldo'])
                        
                    submit_email = st.form_submit_button("📧 ENVIAR CORREO CON PDF ADJUNTO", type="primary")
                    
                    if submit_email:
                        if enviar_correo(email_dest, asunto_msg, cuerpo_html, pdf_bytes):
                            st.success(f"✅ Correo enviado a {email_dest}")
                        else:
                            st.error("❌ Falló el envío. Revisa credenciales y logs.")


    # ==============================================================================
    # TAB GERENTE: VISIÓN ESTRATÉGICA
    # ==============================================================================
    with tab_gerente:
        st.subheader("📊 Análisis de Cartera por Segmento y Concentración")
        
        c_pie, c_bar = st.columns(2)
        
        # --- Gráfico de Distribución por Rango de Mora ---
        with c_pie:
            st.markdown("**1. Distribución de Saldo por Rango de Mora** ")
            df_pie = df_view.groupby('Rango', observed=True)['importe'].sum().reset_index()
            # Mapeo de colores coherente con los rangos (del código 1)
            color_map = {"🟢 Al Día": "green", "🟡 Prev. (1-15)": "gold", "🟠 Riesgo (16-30)": "orange", 
                         "🔴 Crítico (31-60)": "orangered", "🚨 Alto Riesgo (61-90)": "red", "⚫ Legal (+90)": "black"}
            fig_pie = px.pie(df_pie, names='Rango', values='importe', color='Rango', 
                             color_discrete_map=color_map, hole=.3)
            fig_pie.update_traces(textinfo='percent+label', marker=dict(line=dict(color='#000000', width=1)))
            st.plotly_chart(fig_pie, use_container_width=True)
            
        # --- Top 10 Clientes Morosos (Pareto) ---
        with c_bar:
            st.markdown("**2. Top 10 Clientes Morosos (Pareto)** ")
            # Solo clientes con mora y saldo positivo 
            top_cli = df_view[(df_view['dias_vencido'] > 0) & (df_view['importe'] > 0)].groupby('nombrecliente')['importe'].sum().nlargest(10).reset_index()
            fig_bar = px.bar(top_cli, x='importe', y='nombrecliente', orientation='h', 
                             text_auto='$.2s', title="Monto de Deuda Vencida (Top 10)", 
                             color_discrete_sequence=[COLOR_PRIMARIO])
            fig_bar.update_layout(yaxis={'categoryorder':'total ascending'}, xaxis_title="Monto Vencido", yaxis_title="Cliente")
            st.plotly_chart(fig_bar, use_container_width=True)
            
        st.markdown("---")
        st.markdown("### 3. Desempeño y Riesgo por Vendedor")
        
        # Calcular métricas por Vendedor
        resumen_vendedor = df_view.groupby('nomvendedor_norm').agg(
            nomvendedor=('nomvendedor', 'first'),
            Cartera_Total=('importe', 'sum'),
            Vencido=('importe', lambda x: x[df_view.loc[x.index, 'dias_vencido'] > 0].sum())
        ).reset_index()
        resumen_vendedor['% Vencido'] = (resumen_vendedor['Vencido'] / resumen_vendedor['Cartera_Total'] * 100).fillna(0)
        
        vencidos_df = df_view[df_view['dias_vencido'] > 0]
        clientes_mora_vendedor = vencidos_df.groupby('nomvendedor_norm')['nombrecliente'].nunique().reset_index().rename(columns={'nombrecliente': 'Clientes_Mora'})
        
        # CSI por Vendedor
        # El groupby.apply requiere la columna 'nomvendedor_norm' para funcionar correctamente en el merge posterior.
        csi_vendedor = vencidos_df.groupby('nomvendedor_norm').apply(
            lambda x: (x['importe'] * x['dias_vencido']).sum() / df_view[df_view['nomvendedor_norm'] == x.name]['importe'].sum() if df_view[df_view['nomvendedor_norm'] == x.name]['importe'].sum() > 0 else 0, 
            include_groups=False
        ).reset_index(name='CSI')
        
        resumen_vendedor = resumen_vendedor.merge(clientes_mora_vendedor, on='nomvendedor_norm', how='left').fillna(0)
        resumen_vendedor = resumen_vendedor.merge(csi_vendedor, on='nomvendedor_norm', how='left').fillna(0)
        
        # Formato profesional para la tabla
        styled_df = resumen_vendedor.drop(columns=['nomvendedor_norm']).rename(columns={'nomvendedor': 'Vendedor'}).style.format({
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
        st.subheader("📥 Descarga del Reporte Gerencial y Datos Crudos")
        
        excel_data = crear_excel_gerencial(df_view, total, vencido, pct_mora, clientes_mora, csi, antiguedad_prom_vencida)
        
        st.download_button(
            label="💾 DESCARGAR REPORTE GERENCIAL (EXCEL) - Formato Profesional",
            data=excel_data,
            file_name=f"Reporte_Cartera_Ferreinox_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        
        st.subheader("🔎 Datos Crudos Filtrados")
        # Mostrar el dataframe completo con las columnas clave
        cols_mostrar = ['nombrecliente', 'nit', 'numero', 'fecha_documento', 'fecha_vencimiento', 'dias_vencido', 'importe', 'Rango', 'nomvendedor', 'zona', 'telefono1', 'e_mail']
        st.dataframe(df_view[cols_mostrar].style.format({'importe': '${:,.0f}', 'dias_vencido': '{:,.0f}'}), use_container_width=True, hide_index=True)

if __name__ == "__main__":
    main()
