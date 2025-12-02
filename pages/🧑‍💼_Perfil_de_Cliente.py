import streamlit as st
import pandas as pd
import plotly.express as px
import io
import os
import glob
import re
import unicodedata
from datetime import datetime
from urllib.parse import quote
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment

# --- CONFIGURACIÓN VISUAL PROFESIONAL ---
st.set_page_config(
    page_title="Centro de Mando: Cobranza Ferreinox",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Paleta de Colores y CSS Corporativo
COLOR_PRIMARIO = "#003366"  # Azul oscuro corporativo
st.markdown(f"""
<style>
    .main {{ background-color: #f4f6f9; }}
    .stMetric {{ background-color: white; padding: 15px; border-radius: 8px; border-left: 5px solid {COLOR_PRIMARIO}; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }}
    div[data-testid="stExpander"] div[role="button"] p {{ font-size: 1.1rem; font-weight: bold; color: {COLOR_PRIMARIO}; }}
    /* Asegurar que las pestañas sean visibles */
    div[data-testid="stTabs"] button {{ font-weight: bold; font-size: 16px; }}
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
            if s_val.rfind(',') > s_val.rfind('.'): s_val = s_val.replace('.', '').replace(',', '.') # Latino
            else: s_val = s_val.replace(',', '') # USA
        return float(s_val.replace(',', '').replace(' ', ''))
    except:
        return 0.0

def mapear_y_limpiar_df(df):
    """Mapea, limpia y valida las columnas."""
    df.columns = [normalizar_texto(c) for c in df.columns]
    
    # Mapeo de Columnas Críticas y Opcionales (con múltiples sinónimos)
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
            if standard not in renombres.values() and any(v in col for v in variantes):
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
    
    # Asegurar campos opcionales para evitar el error 'telefono'
    for c in ['telefono', 'vendedor', 'nit', 'factura']:
        if c not in df.columns: 
            df[c] = 'N/A'
        else:
            df[c] = df[c].fillna('N/A').astype(str)

    return df[df['saldo'] > 0], "Datos limpios y listos."


@st.cache_data(ttl=600)
def cargar_datos():
    """Busca el archivo de cartera más reciente y lo procesa."""
    archivos = glob.glob("*Cartera*.xlsx") + glob.glob("*Cartera*.csv")
    
    if not archivos:
        return None, "No se encontró ningún archivo con el nombre '*Cartera*'."
    
    archivo = max(archivos, key=os.path.getctime)
    
    try:
        if archivo.endswith('.csv'):
            df = pd.read_csv(archivo, sep=None, engine='python', encoding='latin-1', dtype=str)
        else:
            # openpyxl es más estable para Excel
            df = pd.read_excel(archivo, engine='openpyxl', dtype=str)
            
        df_procesado, status = mapear_y_limpiar_df(df)
        
        if df_procesado is None:
            return None, f"Error en la estructura del archivo {archivo}: {status}"
        
        return df_procesado, f"Datos cargados y limpios de: {archivo}"
        
    except Exception as e:
        return None, f"Error leyendo {archivo}: {str(e)}"

# ======================================================================================
# 2. CEREBRO DE ESTRATEGIA Y GUIONES (Guía de Acción)
# ======================================================================================

def generar_estrategia(row):
    """Segmenta la cartera y genera el guion de WhatsApp basado en el riesgo."""
    dias = row['dias']
    saldo = row['saldo']
    cliente = str(row['cliente']).split()[0].title()
    
    # Lógica de Semáforo y Guion profesional
    if dias <= 0:
        estado = "🟢 Corriente"
        prioridad = 3
        msg = f"Hola {cliente}, saludamos de Ferreinox. Su estado de cuenta está al día. ¡Gracias por su excelente hábito de pago!"
    elif dias <= 15:
        estado = "🟡 Preventivo (1-15)"
        prioridad = 2
        msg = f"Hola {cliente}, amable recordatorio de Ferreinox. Tienes un saldo pendiente de ${saldo:,.0f} vencido hace {dias} días. ¿Nos confirmas la fecha de pago, por favor?"
    elif dias <= 30:
        estado = "🟠 Administrativo (16-30)"
        prioridad = 1
        msg = f"ATENCIÓN {cliente}: Notamos una factura de ${saldo:,.0f} con {dias} días de vencimiento. Requerimos tu gestión inmediata para actualizar tu estado crediticio."
    elif dias <= 60:
        estado = "🔴 Alto Riesgo (31-60)"
        prioridad = 0
        msg = f"URGENTE {cliente}: Saldo de ${saldo:,.0f} ({dias} días). Por políticas internas, tu crédito está bajo revisión para bloqueo de despachos. Contáctanos de inmediato."
    else:
        estado = "⚫ Pre-Jurídico (+60)"
        prioridad = -1
        msg = f"ACCIÓN LEGAL {cliente}: Su cuenta está en estado PRE-JURÍDICO. Saldo: ${saldo:,.0f}. Evite reporte negativo a centrales de riesgo y honorarios de cobranza. Gestiónalo hoy."
        
    return pd.Series([estado, prioridad, msg])

def crear_link_whatsapp(row):
    """Genera el enlace de WhatsApp, limpiando y estandarizando el número (+57)."""
    tel = str(row['telefono']).strip()
    tel = re.sub(r'\D', '', tel) # Quita todo lo que no sea dígito
    if len(tel) < 10: return None # Número incompleto
    if len(tel) == 10: tel = '57' + tel # Asume Colombia
    if len(tel) > 12: tel = tel[-10:] # Si tiene código de área, toma los últimos 10 (ej: 0321)

    return f"https://wa.me/{tel}?text={quote(row['Mensaje_WhatsApp'])}"

# ======================================================================================
# 3. EXPORTACIÓN PROFESIONAL
# ======================================================================================

def generar_excel_gerencial(df_input):
    """Genera un archivo Excel con formato corporativo y resumen ejecutivo."""
    output = io.BytesIO()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Cartera Ferreinox - Gerencial"
    
    # Estilos
    header_fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    summary_fill = PatternFill(start_color="F0F8FF", end_color="F0F8FF", fill_type="solid")
    money_fmt = '"$"#,##0'
    
    # Columnas a Exportar
    cols = ['cliente', 'nit', 'factura', 'vendedor', 'telefono', 'dias', 'saldo', 'Estado']
    cols_to_export = [c for c in cols if c in df_input.columns]
    
    # Título
    sheet['A1'] = "REPORTE EJECUTIVO DE CARTERA"
    sheet['A1'].font = Font(size=18, bold=True, color="003366")
    sheet['A2'] = f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    sheet['A2'].font = Font(italic=True)

    # Headers
    start_row = 4
    sheet.append([c.upper() for c in cols_to_export])
    for col_idx, cell in enumerate(sheet[start_row], 1):
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center')
        
    # Rows
    for row in df_input[cols_to_export].itertuples(index=False):
        sheet.append(row)
    
    # Formato de Moneda y Colores por Días Vencidos
    saldo_col_idx = cols_to_export.index('saldo') + 1
    dias_col_idx = cols_to_export.index('dias') + 1
    
    for row_idx, row in enumerate(sheet.iter_rows(min_row=start_row + 1, max_row=sheet.max_row), start=start_row + 1):
        # Formato de saldo
        row[saldo_col_idx - 1].number_format = money_fmt
        
        # Colores por Mora
        dias = row[dias_col_idx - 1].value
        if isinstance(dias, int):
            if dias > 60: row[dias_col_idx - 1].fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type='solid') # Rojo suave
            elif dias > 30: row[dias_col_idx - 1].fill = PatternFill(start_color="FFEB99", end_color="FFEB99", fill_type='solid') # Naranja suave
            elif dias > 0: row[dias_col_idx - 1].fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type='solid') # Amarillo suave

    # Autoajuste de columnas
    for i, col_name in enumerate(cols_to_export, 1):
        ancho = 25 if col_name in ['cliente', 'vendedor'] else 15
        sheet.column_dimensions[chr(64 + i)].width = ancho

    # Resumen Ejecutivo al final
    ultima_fila = sheet.max_row
    
    sheet[f"A{ultima_fila + 2}"] = "RESUMEN EJECUTIVO:"
    sheet[f"A{ultima_fila + 2}"].font = Font(bold=True)
    
    # Total Cartera
    sheet[f"E{ultima_fila + 3}"] = "TOTAL CARTERA:"
    sheet[f"F{ultima_fila + 3}"] = f"=SUBTOTAL(9,{chr(64 + saldo_col_idx)}{start_row + 1}:{chr(64 + saldo_col_idx)}{ultima_fila})"
    sheet[f"F{ultima_fila + 3}"].number_format = money_fmt
    sheet[f"F{ultima_fila + 3}"].font = Font(bold=True, color="006400") # Verde Oscuro
    sheet[f"E{ultima_fila + 3}"].fill = summary_fill

    # Total Vencido
    sheet[f"E{ultima_fila + 4}"] = "TOTAL VENCIDO (Mora > 0):"
    # Fórmula para sumar solo si los días de mora son > 0
    dias_col_letra = chr(64 + dias_col_idx)
    saldo_col_letra = chr(64 + saldo_col_idx)
    sheet[f"F{ultima_fila + 4}"] = f"=SUMIF({dias_col_letra}{start_row+1}:{dias_col_letra}{ultima_fila}, \">0\", {saldo_col_letra}{start_row+1}:{saldo_col_letra}{ultima_fila})"
    sheet[f"F{ultima_fila + 4}"].number_format = money_fmt
    sheet[f"F{ultima_fila + 4}"].font = Font(bold=True, color="FF0000") # Rojo
    sheet[f"E{ultima_fila + 4}"].fill = summary_fill
    
    workbook.save(output)
    return output.getvalue()

# ======================================================================================
# 4. INTERFAZ GRÁFICA (DASHBOARD)
# ======================================================================================

def main():
    col_logo, col_titulo = st.columns([1, 5])
    with col_titulo:
        st.title("🛡️ Centro de Mando: Cobranza Estratégica")
        st.markdown(f"**Ferreinox SAS BIC** | Panel Operativo y Gerencial | Última actualización: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        st.button("🔄 Recargar Datos (Buscar nuevo archivo 'Cartera')", on_click=st.cache_data.clear)

    # --- CARGA Y PROCESAMIENTO ---
    df_raw, status = cargar_datos()
    if df_raw is None:
        st.error(status)
        st.info("Asegúrese de que su archivo se llame `*Cartera*.xlsx` o `*Cartera*.csv` y contenga las columnas: **Cliente**, **Saldo** y **Días Mora**.")
        return

    # Aplicar Estrategia
    df = df_raw.copy()
    df[['Estado', 'Prioridad', 'Mensaje_WhatsApp']] = df.apply(generar_estrategia, axis=1)
    df['Link_WA'] = df.apply(crear_link_whatsapp, axis=1)
    
    # --- SIDEBAR: FILTROS ---
    with st.sidebar:
        st.header("🔍 Filtros de Gestión")
        
        # Filtro Vendedor
        vendedores = ["TODOS"] + sorted(list(df['vendedor'].unique()))
        sel_vendedor = st.selectbox("Vendedor / Zona", vendedores)
        if sel_vendedor != "TODOS":
            df = df[df['vendedor'] == sel_vendedor]

        # Filtro Estado (Prioridad)
        estados = ["TODOS"] + sorted(list(df['Estado'].unique()), key=lambda x: (x.count('⚫'), x.count('🔴'), x.count('🟠'), x.count('🟡'), x.count('🟢'))) # Ordenar por gravedad
        sel_estado = st.selectbox("Estado de Mora", estados)
        if sel_estado != "TODOS":
            df = df[df['Estado'] == sel_estado]
        
        st.markdown("---")
        if st.checkbox("Mostrar solo Cartera Vencida"):
             df = df[df['dias'] > 0]


    if df.empty:
         st.warning("No hay datos que coincidan con los filtros aplicados.")
         return

    # --- KPIs SUPERIORES ---
    total = df['saldo'].sum()
    vencido = df[df['dias'] > 0]['saldo'].sum()
    critico = df[df['dias'] >= 60]['saldo'].sum()
    pct_mora = (vencido/total)*100 if total > 0 else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("💰 Cartera Total", f"${total:,.0f}")
    k2.metric("⚠️ Total Vencido", f"${vencido:,.0f}", f"{pct_mora:.1f}% del total")
    k3.metric("🔥 Crítico (+60 Días)", f"${critico:,.0f}", "Prioridad de Cobro Máxima")
    k4.metric("👥 Clientes Morosos", f"{df[df['dias'] > 0]['cliente'].nunique()}")

    # --- PESTAÑAS PRINCIPALES ---
    tab_accion, tab_analisis, tab_export = st.tabs(["🚀 GESTIÓN DIARIA: ELIMINAR MORA", "📊 ANÁLISIS GERENCIAL", "📥 EXPORTAR Y DATOS"])

    # --------------------------------------------------------
    # TAB 1: GESTIÓN (Prioridad para Líder de Cartera)
    # --------------------------------------------------------
    with tab_accion:
        st.subheader("🎯 Tareas del Día: Clientes a Contactar")
        st.caption("La tabla está ordenada automáticamente de **CRÍTICO** a **PREVENTIVO**. Gestiona desde arriba hacia abajo.")

        # Preparar datos para la tabla interactiva
        df_display = df.sort_values(by=['Prioridad', 'dias', 'saldo'], ascending=[False, False, False]).copy()
        
        # Filtrar columnas clave para la acción
        columnas_accion = ['cliente', 'factura', 'dias', 'saldo', 'Estado', 'vendedor', 'telefono', 'Link_WA']

        # Usar st.data_editor para el LinkColumn
        st.data_editor(
            df_display[columnas_accion],
            column_config={
                "Link_WA": st.column_config.LinkColumn(
                    "📱 ACCIÓN WHATSAPP",
                    help="Clic para abrir WhatsApp Web con el guion listo",
                    validate="^https://wa\.me/.*",
                    display_text="💬 ENVIAR GUION"
                ),
                "saldo": st.column_config.NumberColumn("Deuda Total", format="$ %d"),
                "dias": st.column_config.NumberColumn("Días Mora", format="%d días", min_value=0, max_value=120),
                "Estado": st.column_config.TextColumn("ESTADO (Prioridad)", width="medium"),
                "cliente": st.column_config.TextColumn("CLIENTE (Razón Social)", width="large"),
                "telefono": st.column_config.TextColumn("Teléfono")
            },
            hide_index=True,
            use_container_width=True,
            height=600
        )

    # --------------------------------------------------------
    # TAB 2: ANÁLISIS (Visión Estratégica para Gerencia)
    # --------------------------------------------------------
    with tab_analisis:
        st.subheader("📈 Concentración y Antigüedad de la Cartera")
        c1, c2 = st.columns(2)
        
        with c1:
            st.markdown("**1. Distribución de Cartera por Riesgo**")
            fig_pie = px.pie(df, values='saldo', names='Estado', hole=0.4, 
                             color_discrete_sequence=px.colors.sequential.RdBu, 
                             color='Estado', title="Monto Total por Estado de Cobranza")
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with c2:
            st.markdown("**2. Top 10 Clientes (Deuda Máxima)**")
            df_top = df.sort_values(by='saldo', ascending=False).head(10)
            fig_bar = px.bar(df_top, x='saldo', y='cliente', orientation='h', 
                             text_auto='.2s', color='dias', 
                             color_continuous_scale='Reds')
            fig_bar.update_layout(yaxis={'categoryorder':'total ascending'}, showlegend=False)
            st.plotly_chart(fig_bar, use_container_width=True)

        # Análisis por Vendedor
        if sel_vendedor == "TODOS":
            st.markdown("---")
            st.subheader("Desempeño por Vendedor/Zona")
            df_vendedor_resumen = df.groupby('vendedor').agg(
                Total_Cartera=('saldo', 'sum'),
                Vencido=('saldo', lambda x: x[df.loc[x.index, 'dias'] > 0].sum()),
                Clientes=('cliente', 'nunique')
            ).reset_index()
            
            df_vendedor_resumen['% Vencido'] = (df_vendedor_resumen['Vencido'] / df_vendedor_resumen['Total_Cartera'] * 100).fillna(0)
            df_vendedor_resumen = df_vendedor_resumen.sort_values('% Vencido', ascending=False)
            
            st.dataframe(df_vendedor_resumen.style.format(
                {'Total_Cartera': '${:,.0f}', 'Vencido': '${:,.0f}', '% Vencido': '{:.1f}%'}
            ), use_container_width=True, hide_index=True)


    # --------------------------------------------------------
    # TAB 3: EXPORTACIÓN (Datos y Descargas)
    # --------------------------------------------------------
    with tab_export:
        st.subheader("📥 Descarga de Reportes y Detalle de Datos")
        
        col_dl, col_raw = st.columns([1, 2])
        
        with col_dl:
            st.markdown("**Reporte Listo para Gerencia**")
            excel_data = generar_excel_gerencial(df)
            st.download_button(
                label="✅ DESCARGAR EXCEL GERENCIAL FORMATEADO",
                data=excel_data,
                file_name=f"Reporte_Cartera_Ferreinox_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            st.caption("Incluye totales y formato profesional, listo para presentar.")
            
        with col_raw:
            st.markdown("**Vista de la Base de Datos Filtrada**")
            st.dataframe(df.drop(columns=['Prioridad', 'Mensaje_WhatsApp', 'Link_WA'], errors='ignore'), use_container_width=True, height=300)

if __name__ == "__main__":
    main()
