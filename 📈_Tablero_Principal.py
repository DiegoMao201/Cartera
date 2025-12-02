import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import re
import unicodedata
from datetime import datetime
from urllib.parse import quote
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# --- CONFIGURACIÓN VISUAL PROFESIONAL ---
st.set_page_config(
    page_title="Centro de Mando: Cobranza Estratégica",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Paleta de Colores y CSS Corporativo
COLOR_PRIMARIO = "#003366"  # Azul oscuro corporativo
COLOR_RIESGO_CRITICO = "#B30000" # Rojo oscuro
COLOR_RIESGO_ALTO = "#FF9900" # Naranja
COLOR_RIESGO_MEDIO = "#FFD700" # Dorado/Amarillo
COLOR_RIESGO_BAJO = "#008000" # Verde
st.markdown(f"""
<style>
    .main {{ background-color: #f4f6f9; }}
    .stMetric {{ background-color: white; padding: 15px; border-radius: 8px; border-left: 5px solid {COLOR_PRIMARIO}; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }}
    div[data-testid="stExpander"] div[role="button"] p {{ font-size: 1.1rem; font-weight: bold; color: {COLOR_PRIMARIO}; }}
    /* Botón de WhatsApp estilizado para la gestión */
    div[data-testid="stDataEditor"] a[target="_blank"] {{
        background-color: #25D366; color: white; padding: 5px 10px; border-radius: 5px; text-decoration: none; font-weight: bold;
        display: inline-block; text-align: center; font-size: 12px;
    }}
    .kpi-title {{ font-size: 1.2rem; font-weight: 600; color: {COLOR_PRIMARIO}; }}
</style>
""", unsafe_allow_html=True)

# ======================================================================================
# 1. MOTOR DE INGESTIÓN Y LIMPIEZA DE DATOS (Inteligencia de Columnas)
# ======================================================================================

def normalizar_texto(texto):
    """Elimina tildes, símbolos y pone mayúsculas para mapeo."""
    if not isinstance(texto, str): return str(texto)
    texto = unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode("utf-8").upper().strip()
    return re.sub(r'[^\w\s]', '', texto).strip()

def limpiar_moneda(valor):
    """Limpia formatos de moneda, tolerante a comas y puntos."""
    if pd.isna(valor): return 0.0
    s_val = str(valor).strip()
    s_val = re.sub(r'[^\d.,-]', '', s_val)
    if not s_val: return 0.0
    try:
        # Intenta manejar formatos (1.000,00 vs 1,000.00)
        if s_val.count(',') > 1 and s_val.count('.') == 0: s_val = s_val.replace(',', '') # Miles con coma
        elif s_val.count('.') > 1 and s_val.count(',') == 0: s_val = s_val.replace('.', '') # Miles con punto
        elif s_val.count(',') == 1 and s_val.count('.') == 1:
            if s_val.rfind(',') > s_val.rfind('.'): s_val = s_val.replace('.', '').replace(',', '.') # Latino (ej: 1.000,00 -> 1000.00)
            else: s_val = s_val.replace(',', '') # USA (ej: 1,000.00 -> 1000.00)
        elif s_val.count('.') == 1 and s_val.count(',') == 0: # Para 1000.00
            s_val = s_val.replace(',', '')
        
        return float(s_val.replace(' ', ''))
    except Exception:
        return 0.0

def mapear_y_limpiar_df(df):
    """Mapea, limpia y valida las columnas."""
    df.columns = [normalizar_texto(c) for c in df.columns]
    
    # Mapeo de Columnas Críticas y Opcionales
    mapa = {
        'cliente': ['NOMBRE', 'RAZON SOCIAL', 'TERCERO', 'CLIENTE'],
        'nit': ['NIT', 'IDENTIFICACION', 'CEDULA', 'RUT'],
        'saldo': ['IMPORTE', 'SALDO', 'TOTAL', 'DEUDA', 'VALOR'],
        'dias': ['DIAS', 'VENCIDO', 'MORA', 'ANTIGUEDAD'],
        'telefono': ['TEL', 'MOVIL', 'CELULAR', 'TELEFONO', 'CONTACTO'],
        'vendedor': ['VENDEDOR', 'ASESOR', 'COMERCIAL', 'NOMVENDEDOR'],
        'factura': ['NUMERO', 'FACTURA', 'DOC', 'SERIE']
    }
    
    renombres = {}
    for standard, variantes in mapa.items():
        for col in df.columns:
            # Normalización ampliada para mejor coincidencia
            col_norm = normalizar_texto(col).replace(' ', '')
            if standard not in renombres.values() and any(normalizar_texto(v).replace(' ', '') in col_norm for v in variantes):
                renombres[col] = standard
                break
    
    df.rename(columns=renombres, inplace=True)
    
    # --- VALIDACIÓN CRÍTICA ---
    req = ['cliente', 'saldo', 'dias']
    if not all(c in df.columns for c in req):
        missing = [c for c in req if c not in df.columns]
        return None, f"Faltan columnas críticas: {', '.join(missing)}. Columnas detectadas: {list(df.columns)}"

    # --- LIMPIEZA Y CONVERSIÓN ---
    df['saldo'] = df['saldo'].apply(limpiar_moneda)
    df['dias'] = pd.to_numeric(df['dias'], errors='coerce').fillna(0).astype(int)
    
    # Asegurar campos opcionales
    for c in ['telefono', 'vendedor', 'nit', 'factura']:
        if c not in df.columns: 
            df[c] = 'N/A'
        else:
            df[c] = df[c].fillna('N/A').astype(str)

    # Agrupar por Cliente, NIT y Vendedor para tener un SALDO CONSOLIDADO
    df_consolidado = df.groupby(['cliente', 'nit', 'vendedor', 'telefono']).agg(
        saldo=('saldo', 'sum'),
        dias_max=('dias', 'max'),
        facturas=('factura', lambda x: ', '.join(x.unique().astype(str)))
    ).reset_index()
    
    df_consolidado.rename(columns={'dias_max': 'dias'}, inplace=True)
    
    return df_consolidado[df_consolidado['saldo'] > 0], "Datos limpios y listos. Cartera consolidada por cliente."


@st.cache_data(ttl=600)
def cargar_datos_hibrido(archivo_subido=None):
    """Procesa el archivo subido o el archivo local más reciente."""
    df = None
    status = "Esperando un archivo de cartera (*Cartera*.xlsx o .csv)."

    if archivo_subido is not None:
        try:
            if archivo_subido.name.endswith('.csv'):
                df = pd.read_csv(archivo_subido, sep=None, engine='python', encoding='latin-1', dtype=str)
            else:
                df = pd.read_excel(archivo_subido, engine='openpyxl', dtype=str)
            
            status = f"Archivo cargado: {archivo_subido.name}"
        except Exception as e:
            return None, f"Error leyendo el archivo subido: {str(e)}"
    
    if df is not None:
        return mapear_y_limpiar_df(df)
        
    return None, status

# ======================================================================================
# 2. CEREBRO DE ESTRATEGIA Y GUIONES (Guía de Acción)
# ======================================================================================

def generar_estrategia(row):
    """Segmenta la cartera, asigna prioridad y genera el guion de WhatsApp detallado."""
    dias = row['dias']
    saldo = row['saldo']
    cliente_nombre_corto = str(row['cliente']).split()[0].title()
    
    # Lógica de Semáforo y Prioridad (Prioridad: Mayor valor = Mayor urgencia)
    if dias <= 0:
        estado = "🟢 Corriente"
        prioridad = 0
        color_hex = COLOR_RIESGO_BAJO
        msg = (
            f"✅ ¡Hola {cliente_nombre_corto}! Te saludamos de Ferreinox SAS BIC. \n"
            f"Tu estado de cuenta consolidado con nosotros está **AL DÍA**. ¡Gracias por tu excelente hábito de pago! \n"
            f"Saldo Total: ${saldo:,.0f}"
        )
    elif dias <= 15:
        estado = "🟡 Preventivo (1-15 días)"
        prioridad = 1
        color_hex = COLOR_RIESGO_MEDIO
        msg = (
            f"🔔 ¡Hola {cliente_nombre_corto}! Amable recordatorio de Ferreinox. \n"
            f"Hemos notado un saldo pendiente de *${saldo:,.0f}* con vencimiento máximo hace *{dias} días*.\n"
            f"🙏 Por favor, confírmanos la fecha exacta en la que se realizará el pago para actualizar tu estado. \n"
            f"Facturas involucradas: {row['facturas'][:50]}..."
        )
    elif dias <= 30:
        estado = "🟠 Administrativo (16-30 días)"
        prioridad = 2
        color_hex = COLOR_RIESGO_ALTO
        msg = (
            f"⚠️ ATENCIÓN {cliente_nombre_corto}: Notificación de Cartera Vencida (Ferreinox).\n"
            f"Tienes un total de *${saldo:,.0f}* vencido hace *{dias} días*.\n"
            f"🚨 Esto afecta directamente tu cupo de crédito y futuros despachos. \n"
            f"📞 Necesitamos tu GESTIÓN INMEDIATA para saldar o establecer un compromiso de pago firme. Contáctanos urgente."
        )
    elif dias <= 60:
        estado = "🔴 Alto Riesgo (31-60 días)"
        prioridad = 3
        color_hex = COLOR_RIESGO_CRITICO
        msg = (
            f"🔥 **URGENTE ACCIÓN DE PAGO** {cliente_nombre_corto} (Ferreinox).\n"
            f"El saldo de *${saldo:,.0f}* tiene *{dias} días* de mora.\n"
            f"🚫 **ADVERTENCIA:** Tu cuenta está bajo REVISIÓN para BLOQUEO TOTAL de despachos. \n"
            f"⛔ Evita el reporte negativo a centrales de riesgo. ¡Sálvalo hoy mismo!"
        )
    else:
        estado = "⚫ Pre-Jurídico (+60 días)"
        prioridad = 4
        color_hex = "#000000" # Negro
        msg = (
            f"⚖️ **ACCIÓN LEGAL INMINENTE** {cliente_nombre_corto} (Ferreinox).\n"
            f"El saldo de *${saldo:,.0f}* supera los *60 DÍAS* de vencimiento. \n"
            f"❌ Tu deuda ha sido escalada a la fase PRE-JURÍDICA. \n"
            f"El costo de la deuda aumentará con los honorarios de cobro. Exige el link de pago y liquida ahora."
        )
        
    # Añadir link de Portal de Pagos para clientes con mora
    if dias > 0:
         portal_link = "https://tu-portal-de-pagos.com/recaudo" # URL de ejemplo, cámbiala a la real
         msg += f"\n\n🔗 *Link de Pago Rápido: {portal_link}*"

    return pd.Series([estado, prioridad, msg, color_hex])

def crear_link_whatsapp(row):
    """Genera el enlace de WhatsApp, limpiando y estandarizando el número (+57)."""
    tel = str(row['telefono']).strip()
    tel = re.sub(r'\D', '', tel) # Quita todo lo que no sea dígito
    if len(tel) < 10: return None # Número incompleto
    
    # Asume código de país +57 (Colombia) si solo tiene 10 dígitos (ej: 310xxxxxxx)
    if len(tel) == 10 and not tel.startswith('57'): 
        tel = '57' + tel 
    # Asegura que no haya prefijos erróneos y que el número sea funcional para WA
    elif len(tel) > 12: # Si tiene más de 12 dígitos (ej: 0057310xxxxxxx), toma los últimos 10
         tel = tel[-10:]
         if not tel.startswith('57'):
              tel = '57' + tel

    # La columna 'Mensaje_WhatsApp' debe existir en el DataFrame antes de llamar a esta función
    if 'Mensaje_WhatsApp' not in row:
        return None
        
    return f"https://wa.me/{tel}?text={quote(row['Mensaje_WhatsApp'])}"

# ======================================================================================
# 3. EXPORTACIÓN PROFESIONAL (Generador de Reporte Gerencial con Formato)
# ======================================================================================

def generar_excel_gerencial(df_input):
    """Genera un archivo Excel con formato corporativo, resumen y análisis de Pareto."""
    output = io.BytesIO()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Reporte Cobranza Estrategica"
    
    # --- ESTILOS ---
    fill_header = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
    font_header = Font(color="FFFFFF", bold=True)
    fill_summary = PatternFill(start_color="F0F8FF", end_color="F0F8FF", fill_type="solid")
    font_bold_red = Font(bold=True, color="FF0000")
    font_bold_green = Font(bold=True, color="006400")
    money_fmt = '"$"#,##0'
    
    # --- METADATOS Y TÍTULO ---
    sheet['A1'] = "REPORTE DE COBRANZA ESTRATÉGICA CONSOLIDADA"
    sheet['A1'].font = Font(size=18, bold=True, color="003366")
    sheet['A2'] = f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    sheet['A2'].font = Font(italic=True)

    # --- DATOS DE LA CARTERA (Clientes Únicos consolidados) ---
    cols_to_export = ['cliente', 'nit', 'vendedor', 'telefono', 'dias', 'saldo', 'Estado', 'facturas']
    
    start_row = 5
    # Headers
    sheet.append([c.replace('_', ' ').upper() for c in cols_to_export])
    for col_idx, cell in enumerate(sheet[start_row], 1):
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal='center')
        
    # Rows
    for row in df_input[cols_to_export].itertuples(index=False):
        sheet.append(row)
    
    ultima_fila_datos = sheet.max_row
    
    # Formato de Moneda y Colores por Días Vencidos
    saldo_col_idx = cols_to_export.index('saldo') + 1
    dias_col_idx = cols_to_export.index('dias') + 1
    
    # Rellenar filas y aplicar formato de moneda
    for row_idx, row in enumerate(sheet.iter_rows(min_row=start_row + 1, max_row=ultima_fila_datos), start=start_row + 1):
        row[saldo_col_idx - 1].number_format = money_fmt
        
        # Colores por Mora
        dias = row[dias_col_idx - 1].value
        if isinstance(dias, int):
            if dias > 60: row[dias_col_idx - 1].fill = PatternFill(start_color="FFCCCC", fill_type='solid') 
            elif dias > 30: row[dias_col_idx - 1].fill = PatternFill(start_color="FFEB99", fill_type='solid')
            elif dias > 0: row[dias_col_idx - 1].fill = PatternFill(start_color="FFFFCC", fill_type='solid')

    # Autoajuste de columnas
    for i, col_name in enumerate(cols_to_export, 1):
        ancho = 30 if col_name in ['cliente', 'facturas'] else 18
        sheet.column_dimensions[get_column_letter(i)].width = ancho

    # --- RESUMEN EJECUTIVO (KPIs) ---
    
    ultima_fila_resumen = ultima_fila_datos + 3
    
    sheet[f"A{ultima_fila_resumen}"] = "RESUMEN EJECUTIVO DE KPIs"
    sheet[f"A{ultima_fila_resumen}"].font = Font(size=14, bold=True, color="003366")
    
    # Fórmulas de Totales
    total_cartera_formula = f"=SUBTOTAL(9,{get_column_letter(saldo_col_idx)}{start_row + 1}:{get_column_letter(saldo_col_idx)}{ultima_fila_datos})"
    total_vencido_formula = f"=SUMIF({get_column_letter(dias_col_idx)}{start_row+1}:{get_column_letter(dias_col_idx)}{ultima_fila_datos}, \">0\", {get_column_letter(saldo_col_idx)}{start_row+1}:{get_column_letter(saldo_col_idx)}{ultima_fila_datos})"
    
    sheet[f"B{ultima_fila_resumen + 1}"] = "TOTAL CARTERA:"
    sheet[f"C{ultima_fila_resumen + 1}"] = total_cartera_formula
    sheet[f"C{ultima_fila_resumen + 1}"].number_format = money_fmt
    sheet[f"C{ultima_fila_resumen + 1}"].font = font_bold_green
    
    sheet[f"B{ultima_fila_resumen + 2}"] = "TOTAL VENCIDO (MORA > 0):"
    sheet[f"C{ultima_fila_resumen + 2}"] = total_vencido_formula
    sheet[f"C{ultima_fila_resumen + 2}"].number_format = money_fmt
    sheet[f"C{ultima_fila_resumen + 2}"].font = font_bold_red
    
    sheet[f"B{ultima_fila_resumen + 3}"] = "% CARTERA VENCIDA:"
    sheet[f"C{ultima_fila_resumen + 3}"] = f"={get_column_letter(3)}{ultima_fila_resumen + 2}/{get_column_letter(3)}{ultima_fila_resumen + 1}"
    sheet[f"C{ultima_fila_resumen + 3}"].number_format = "0.0%"
    sheet[f"C{ultima_fila_resumen + 3}"].font = Font(bold=True)
    
    for r in range(ultima_fila_resumen + 1, ultima_fila_resumen + 4):
        sheet[f"B{r}"].fill = fill_summary
        sheet[f"C{r}"].fill = PatternFill(start_color="FFFFFF", fill_type='solid')

    # --- ANÁLISIS DE PARETO (Top 20% Clientes) ---
    
    df_vencido = df_input[df_input['dias'] > 0].copy()
    if not df_vencido.empty:
        df_vencido_sorted = df_vencido.sort_values(by='saldo', ascending=False)
        total_vencido = df_vencido_sorted['saldo'].sum()
        df_vencido_sorted['acumulado'] = df_vencido_sorted['saldo'].cumsum()
        
        pareto_limit_idx = len(df_vencido_sorted[df_vencido_sorted['acumulado'] <= total_vencido * 0.8])
        df_pareto = df_vencido_sorted.head(pareto_limit_idx + 1)

        ultima_fila_pareto = ultima_fila_resumen + 6
        sheet[f"A{ultima_fila_pareto}"] = "CLIENTES ESTRATÉGICOS (80/20)"
        sheet[f"A{ultima_fila_pareto}"].font = Font(size=14, bold=True, color="003366")
        
        pareto_cols = ['cliente', 'dias', 'saldo']
        
        # Headers Pareto
        sheet.append([c.replace('_', ' ').upper() for c in pareto_cols])
        pareto_header_row = ultima_fila_pareto + 1
        for col_idx, cell in enumerate(sheet[pareto_header_row], 1):
            cell.fill = fill_header
            cell.font = font_header
            cell.alignment = Alignment(horizontal='center')
        
        # Rows Pareto
        for row in df_pareto[pareto_cols].itertuples(index=False):
            sheet.append(row)
        
        # Formato Moneda Pareto
        for r in range(pareto_header_row + 1, sheet.max_row + 1):
            sheet[f"C{r}"].number_format = money_fmt
            sheet[f"C{r}"].font = font_bold_red
    
    workbook.save(output)
    return output.getvalue()

# ======================================================================================
# 4. INTERFAZ GRÁFICA (DASHBOARD)
# ======================================================================================

def generar_analisis_ia(kpis: dict):
    """Genera comentarios de análisis basados en KPIs."""
    comentarios = []
    
    # Análisis de % Vencido
    if kpis['pct_vencido'] > 20: 
        comentarios.append(f"<li>🚨 **Riesgo Crítico:** Un **{kpis['pct_vencido']:.1f}%** de cartera vencida es muy alto. El foco debe estar en la **recuperación inmediata** de la mora > 60 días.</li>")
    elif kpis['pct_vencido'] > 10: 
        comentarios.append(f"<li>⚠️ **Advertencia:** El **{kpis['pct_vencido']:.1f}%** está en mora. Esto afecta el flujo de caja. Implementar la **gestión preventiva (1-30 días)** es la clave para evitar que se convierta en crítico.</li>")
    else: 
        comentarios.append(f"<li>🟢 **Saludable:** El porcentaje de mora es bajo (**{kpis['pct_vencido']:.1f}%**). Mantener la estrategia preventiva activa.</li>")
        
    # Análisis de CSI (Índice de Severidad de Cobranza)
    if kpis['csi'] > 30:
         comentarios.append(f"<li>🔥 **Severidad Extrema (CSI: {kpis['csi']:.0f}):** La deuda está envejeciendo rápidamente o el monto de la mora antigua es masivo. **Necesitas un plan de pago especial (Acuerdos, abonos) para los clientes TOP 10.**</li>")
    elif kpis['csi'] > 15:
        comentarios.append(f"<li>🟡 **Severidad Moderada (CSI: {kpis['csi']:.0f}):** Indica que la morosidad no solo es alta, sino también **ANTIGUA**. Es crucial enfocarse en la cartera de **31-60 días**.</li>")
    else:
        comentarios.append(f"<li>✨ **Severidad Baja (CSI: {kpis['csi']:.0f}):** La mora es relativamente reciente. La gestión de cobranza está siendo efectiva en fases tempranas.</li>")
        
    # Recomendación de Gestión Inmediata
    top_n = kpis['num_pareto_clientes']
    total_n = kpis['num_total_clientes_mora']
    if top_n > 0:
        comentarios.append(f"<li>🎯 **Foco Operativo:** Concentrar el **80% de tu tiempo** de gestión en los **{top_n} clientes** que componen el 80% de la deuda vencida (Regla de Pareto).</li>")
    else:
        comentarios.append(f"<li>😌 **Foco Operativo:** Como no hay clientes críticos por Pareto, la gestión debe ser **exhaustiva y equitativa** con todos los clientes con mora.</li>")

    return "<ul>" + "".join(comentarios) + "</ul>"

def main():
    col_logo, col_titulo = st.columns([1, 5])
    with col_titulo:
        st.title("🛡️ Centro de Mando: Cobranza Estratégica")
        st.markdown(f"**Ferreinox SAS BIC** | Panel Operativo y Gerencial | Última actualización: **{datetime.now().strftime('%d/%m/%Y %H:%M')}**")
    
    st.markdown("---")
    
    # --- CARGA HÍBRIDA (Opción de subir archivo) ---
    st.sidebar.header("📁 Carga de Datos")
    uploaded_file = st.sidebar.file_uploader("Sube el archivo de Cartera (Excel/CSV)", type=['xlsx', 'csv'], help="El archivo debe contener al menos las columnas: NombreCliente, Saldo/Importe, Días Vencido")
    
    if uploaded_file is None:
        st.info("Sube el archivo de Cartera para empezar el análisis.")
        return

    df_raw, status = cargar_datos_hibrido(uploaded_file)
    st.sidebar.caption(status)
    
    if df_raw is None:
        st.error(status)
        return

    # Aplicar Estrategia a la Cartera Consolidada
    df = df_raw.copy()
    df[['Estado', 'Prioridad', 'Mensaje_WhatsApp', 'Color_Estado']] = df.apply(generar_estrategia, axis=1)
    df['Link_WA'] = df.apply(crear_link_whatsapp, axis=1)
    
    # --- SIDEBAR: FILTROS ---
    with st.sidebar:
        st.header("🔍 Filtros de Gestión")
        
        # Filtro Vendedor
        vendedores = ["TODOS"] + sorted(list(df['vendedor'].unique()))
        sel_vendedor = st.selectbox("Vendedor / Zona", vendedores)
        if sel_vendedor != "TODOS":
            df = df[df['vendedor'] == sel_vendedor]

        # Filtro Estado (Prioridad) - Ordenado de más crítico a menos
        estados_ordenados = sorted(df['Estado'].unique(), key=lambda x: [c for c in ['⚫', '🔴', '🟠', '🟡', '🟢'] if c in x][0], reverse=True)
        sel_estado = st.selectbox("Estado de Mora", ["TODOS"] + estados_ordenados)
        if sel_estado != "TODOS":
            df = df[df['Estado'] == sel_estado]
            
        st.markdown("---")
        # Mostrar solo Mora por defecto, a menos que se quiera ver todo
        mostrar_vencida = st.checkbox("Mostrar SOLO Cartera Vencida (días > 0)", value=True)
        if mostrar_vencida:
            df = df[df['dias'] > 0]


    if df.empty:
        st.warning("No hay datos que coincidan con los filtros aplicados.")
        return

    # --- CÁLCULO DE KPIs ---
    total_cartera = df['saldo'].sum()
    df_vencido = df[df['dias'] > 0]
    total_vencido = df_vencido['saldo'].sum()
    total_critico = df[df['dias'] >= 60]['saldo'].sum()
    pct_mora = (total_vencido/total_cartera)*100 if total_cartera > 0 else 0
    
    # CSI (Credit Severity Index) = Suma(Saldo * Días) / Total Cartera
    if total_cartera > 0:
        csi = (df_vencido['saldo'] * df_vencido['dias']).sum() / total_cartera
    else:
        csi = 0

    # Análisis de Pareto
    client_debt = df_vencido.groupby('cliente')['saldo'].sum().sort_values(ascending=False)
    num_total_clientes_mora = len(client_debt)
    num_clientes_pareto = 0
    if total_vencido > 0:
        client_debt_cumsum = client_debt.cumsum()
        pareto_limit = total_vencido * 0.80
        num_clientes_pareto = len(client_debt_cumsum[client_debt_cumsum <= pareto_limit]) + 1
        
    kpis_dict = {
        'pct_vencido': pct_mora,
        'csi': csi,
        'num_pareto_clientes': num_clientes_pareto,
        'num_total_clientes_mora': num_total_clientes_mora
    }

    # --- KPIs SUPERIORES (KPIs) ---
    st.header("🔑 Indicadores Clave de Rendimiento (KPIs)")
    k1, k2, k3, k4 = st.columns(4)
    k1.markdown('<p class="kpi-title">💰 Cartera Total (Filtrada)</p>', unsafe_allow_html=True)
    k1.metric("", f"${total_cartera:,.0f}")
    
    k2.markdown('<p class="kpi-title">⚠️ Total Vencido</p>', unsafe_allow_html=True)
    k2.metric("", f"${total_vencido:,.0f}", f"{pct_mora:.1f}% del Total")
    
    k3.markdown('<p class="kpi-title">🔥 Total Crítico (+60 Días)</p>', unsafe_allow_html=True)
    k3.metric("", f"${total_critico:,.0f}", "Prioridad Máxima")
    
    k4.markdown('<p class="kpi-title">💥 Índice de Severidad (CSI)</p>', unsafe_allow_html=True)
    k4.metric("", f"{csi:,.1f}", "Días promedio de impacto")

    with st.expander("🤖 **Análisis y Plan de Acción del Asistente IA**", expanded=False):
        st.markdown(generar_analisis_ia(kpis_dict), unsafe_allow_html=True)
        
    st.markdown("---")

    # --- PESTAÑAS PRINCIPALES ---
    tab_accion, tab_analisis, tab_export = st.tabs(["🚀 GESTIÓN DIARIA: ACCIÓN RÁPIDA", "📊 ANÁLISIS GERENCIAL", "📥 EXPORTAR Y DATOS"])

    # --------------------------------------------------------
    # TAB 1: GESTIÓN (Prioridad para Líder de Cartera)
    # --------------------------------------------------------
    with tab_accion:
        st.subheader("🎯 Clientes a Contactar (Prioridad: Crítico > Alto Riesgo)")
        st.caption(f"Lista de **{df['cliente'].nunique()} clientes únicos** ordenados por la mayor prioridad de cobro.")

        # Preparar datos para la tabla interactiva
        # Ordenar por Prioridad (desc), luego Días (desc), luego Saldo (desc)
        df_display = df.sort_values(by=['Prioridad', 'dias', 'saldo'], ascending=[False, False, False]).copy()
        
        # Columnas clave para la acción
        columnas_accion = ['cliente', 'nit', 'dias', 'saldo', 'Estado', 'vendedor', 'telefono', 'Link_WA']
        
        # Columna para el guion (opcional para visualización rápida)
        df_display['Guion de WhatsApp'] = df_display['Mensaje_WhatsApp'].apply(lambda x: x.split('\n')[0] + "...")

        st.data_editor(
            df_display[columnas_accion + ['Guion de WhatsApp']],
            column_config={
                "Link_WA": st.column_config.LinkColumn(
                    "📱 ACCIÓN WA",
                    help="Clic para abrir WhatsApp Web con el guion listo",
                    validate="^https://wa\.me/.*",
                    display_text="💬 ENVIAR GUION"
                ),
                "saldo": st.column_config.NumberColumn("Deuda Consolidada", format="💰 $ %d"),
                "dias": st.column_config.NumberColumn("Días Mora (Máx)", format="⏳ %d días", min_value=0),
                "Estado": st.column_config.TextColumn("ESTADO (Riesgo)", width="medium"),
                "cliente": st.column_config.TextColumn("CLIENTE (Razón Social)", width="large"),
                "vendedor": st.column_config.TextColumn("Asesor Comercial"),
                "Guion de WhatsApp": st.column_config.TextColumn("Guion de WhatsApp (Vista Previa)"),
            },
            hide_index=True,
            use_container_width=True,
            height=600
        )

    # --------------------------------------------------------
    # TAB 2: ANÁLISIS (Visión Estratégica para Gerencia)
    # --------------------------------------------------------
    with tab_analisis:
        st.subheader("📈 Concentración de Deuda y Distribución de Antigüedad")
        c1, c2 = st.columns(2)
        
        # Distribución de Antigüedad por Saldo
        with c1:
            st.markdown("**1. Distribución de Cartera por Riesgo**")
            df_edades_resumen = df.groupby('Estado', observed=True)['saldo'].sum().reset_index()
            fig_pie = px.pie(
                df_edades_resumen, values='saldo', names='Estado', hole=0.4, 
                color_discrete_map={row['Estado']: row['Color_Estado'] for _, row in df[['Estado', 'Color_Estado']].drop_duplicates().iterrows()},
                title="Monto Total por Estado de Cobranza (Semáforo)"
            )
            fig_pie.update_traces(textinfo='percent+label', marker=dict(line=dict(color='#FFFFFF', width=1)))
            st.plotly_chart(fig_pie, use_container_width=True)
            
        # Top 10 Clientes (Pareto Visual)
        with c2:
            st.markdown("**2. Top 10 Clientes con Mayor Deuda Vencida**")
            df_top = df_vencido.sort_values(by='saldo', ascending=False).head(10)
            if not df_top.empty:
                fig_bar = px.bar(
                    df_top, x='saldo', y='cliente', orientation='h', 
                    text_auto='$.2s', color='dias', 
                    color_continuous_scale='Reds',
                    labels={'saldo': 'Monto Vencido', 'cliente': 'Cliente'},
                    title="Top Clientes por Monto Vencido (Color: Días de Mora)"
                )
                fig_bar.update_layout(yaxis={'categoryorder':'total ascending'}, showlegend=True, coloraxis_colorbar_title='Días Mora')
                st.plotly_chart(fig_bar, use_container_width=True)
            else:
                 st.info("No hay clientes con cartera vencida para mostrar.")
                 
        # Análisis por Vendedor
        if sel_vendedor == "TODOS" and df['vendedor'].nunique() > 1:
            st.markdown("---")
            st.subheader("Desempeño Comparativo por Vendedor/Zona")
            
            # Recalcular el resumen por vendedor
            df_vendedor_resumen = df_raw.copy() # Usamos el raw consolidado antes del filtro de sidebar
            df_vendedor_resumen[['Estado', 'Prioridad', 'Mensaje_WhatsApp', 'Color_Estado']] = df_vendedor_resumen.apply(generar_estrategia, axis=1)
            
            df_vendedor_agg = df_vendedor_resumen.groupby('vendedor').agg(
                Total_Cartera=('saldo', 'sum'),
                Total_Vencido=('saldo', lambda x: x[df_vendedor_resumen.loc[x.index, 'dias'] > 0].sum()),
                Clientes_Mora=('cliente', lambda x: x[df_vendedor_resumen.loc[x.index, 'dias'] > 0].nunique()),
                Clientes_Total=('cliente', 'nunique')
            ).reset_index()
            
            df_vendedor_agg['% Vencido s/ Total'] = (df_vendedor_agg['Total_Vencido'] / df_vendedor_agg['Total_Cartera'] * 100).fillna(0)
            df_vendedor_agg['% Clientes Mora'] = (df_vendedor_agg['Clientes_Mora'] / df_vendedor_agg['Clientes_Total'] * 100).fillna(0)
            
            df_vendedor_agg = df_vendedor_agg.sort_values('% Vencido s/ Total', ascending=False)
            
            st.dataframe(df_vendedor_agg.style.format(
                {'Total_Cartera': '${:,.0f}', 'Total_Vencido': '${:,.0f}', 
                 '% Vencido s/ Total': '{:.1f}%', '% Clientes Mora': '{:.1f}%'}
            ).background_gradient(subset=['% Vencido s/ Total'], cmap='Reds'), 
            use_container_width=True, hide_index=True)


    # --------------------------------------------------------
    # TAB 3: EXPORTACIÓN (Datos y Descargas)
    # --------------------------------------------------------
    with tab_export:
        st.subheader("📥 Descarga de Reportes y Detalle de Datos")
        
        col_dl, col_raw = st.columns([1, 2])
        
        with col_dl:
            st.markdown("**Reporte Listo para Gerencia (Excel)**")
            # Usa el generador de Excel mejorado
            excel_data = generar_excel_gerencial(df)
            st.download_button(
                label="✅ DESCARGAR REPORTE GERENCIAL (Formato Excel)",
                data=excel_data,
                file_name=f"Reporte_Cartera_Estrategica_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            st.caption("Incluye KPIs, fórmulas, formato y análisis de Pareto, listo para presentar.")
            
        with col_raw:
            st.markdown("**Vista Detallada de la Base de Datos Filtrada**")
            # Muestra el detalle de la cartera consolidada filtrada
            st.dataframe(
                df.drop(columns=['Prioridad', 'Mensaje_WhatsApp', 'Link_WA', 'Color_Estado'], errors='ignore'), 
                use_container_width=True, 
                height=300
            )

if __name__ == "__main__":
    main()
