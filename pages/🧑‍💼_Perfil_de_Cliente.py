# ======================================================================================
# ARCHIVO: pages/1_Cobranza_Estrategica.py
# DESCRIPCIÓN: Centro de Comando para Gestión de Cobranza Inteligente y Prioritaria
# ======================================================================================
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO, StringIO
import dropbox
import glob
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import DataBarRule, IconSetRule, ColorScaleRule
from openpyxl.worksheet.table import Table, TableStyleInfo
import unicodedata
import re

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Gestión Estratégica de Cobranza",
    page_icon="👮‍♂️",
    layout="wide"
)

# --- PALETA DE COLORES (Coherente con Principal) ---
COLORS = {
    "nav": "#003865",
    "action": "#D32F2F",    # Rojo Urgente
    "warning": "#F57C00",   # Naranja Alerta
    "safe": "#388E3C",      # Verde Seguro
    "neutral": "#F0F2F6",
    "text": "#31333F"
}

st.markdown(f"""
<style>
    .stMetric {{ background-color: #FFFFFF; border-radius: 8px; padding: 15px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); border-left: 5px solid {COLORS['nav']}; }}
    h1, h2, h3 {{ color: {COLORS['nav']}; }}
    .big-number {{ font-size: 24px; font-weight: bold; color: {COLORS['action']}; }}
    .action-card {{ background-color: #FFF3E0; padding: 20px; border-radius: 10px; border: 1px solid {COLORS['warning']}; margin-bottom: 20px; }}
    .priority-high {{ color: #D32F2F; font-weight: bold; }}
</style>
""", unsafe_allow_html=True)

# ======================================================================================
# --- LÓGICA DE CARGA DE DATOS (Reutilizada y Optimizada) ---
# ======================================================================================

@st.cache_data(ttl=600)
def cargar_datos_inteligentes():
    # 1. Intentar Cargar Dropbox
    df_dropbox = pd.DataFrame()
    try:
        APP_KEY = st.secrets["dropbox"]["app_key"]
        APP_SECRET = st.secrets["dropbox"]["app_secret"]
        REFRESH_TOKEN = st.secrets["dropbox"]["refresh_token"]
        with dropbox.Dropbox(app_key=APP_KEY, app_secret=APP_SECRET, oauth2_refresh_token=REFRESH_TOKEN) as dbx:
            _, res = dbx.files_download(path='/data/cartera_detalle.csv')
            df_dropbox = pd.read_csv(StringIO(res.content.decode('latin-1')), header=None, sep='|', engine='python')
            df_dropbox.columns = [
                'Serie', 'Numero', 'Fecha Documento', 'Fecha Vencimiento', 'Cod Cliente',
                'NombreCliente', 'Nit', 'Poblacion', 'Provincia', 'Telefono1', 'Telefono2',
                'NomVendedor', 'Entidad Autoriza', 'E-Mail', 'Importe', 'Descuento',
                'Cupo Aprobado', 'Dias Vencido'
            ]
    except Exception:
        pass # Fallo silencioso, intentamos locales

    # 2. Intentar Cargar Locales
    df_historico = pd.DataFrame()
    archivos = glob.glob("Cartera_*.xlsx")
    if archivos:
        lista = []
        for f in archivos:
            try:
                temp = pd.read_excel(f)
                if not temp.empty and "Total" in str(temp.iloc[-1, 0]): temp = temp.iloc[:-1]
                lista.append(temp)
            except: pass
        if lista: df_historico = pd.concat(lista, ignore_index=True)

    # 3. Consolidar
    df = pd.concat([df_dropbox, df_historico], ignore_index=True)
    if df.empty: return pd.DataFrame()

    # 4. Limpieza Rápida
    df = df.loc[:, ~df.columns.duplicated()]
    df.columns = [x.lower().replace(' ', '_') for x in df.columns]
    df['importe'] = pd.to_numeric(df['importe'], errors='coerce').fillna(0)
    df['dias_vencido'] = pd.to_numeric(df['dias_vencido'], errors='coerce').fillna(0)
    df = df[~df['serie'].astype(str).str.contains('W|X', case=False, na=False)]
    
    # Manejo de notas crédito
    df.loc[df['numero'] < 0, 'importe'] = df.loc[df['numero'] < 0, 'importe'].abs() * -1
    
    return df

# ======================================================================================
# --- MOTOR DE INTELIGENCIA DE COBRANZA ---
# ======================================================================================

def segmentar_cartera(df):
    """Aplica reglas de negocio para clasificar la deuda y sugerir acciones."""
    df_calc = df.copy()
    
    # Solo nos interesa lo vencido para gestión activa, aunque mostramos todo
    df_calc['es_vencido'] = df_calc['dias_vencido'] > 0
    
    def clasificar_riesgo(row):
        dias = row['dias_vencido']
        if dias <= 0: return "🟢 Al Día"
        elif dias <= 30: return "🟡 Preventiva (0-30)"
        elif dias <= 60: return "🟠 Administrativa (31-60)"
        elif dias <= 120: return "🔴 Pre-Jurídica (61-120)"
        else: return "⚫ Jurídica / Castigo (>120)"

    def sugerir_accion(row):
        dias = row['dias_vencido']
        monto = row['importe']
        
        if dias <= 0: return "Fidelización / Venta Cruzada"
        if dias <= 15: return "Recordatorio Amable (WhatsApp)"
        if dias <= 30: return "Llamada de Servicio + Confirmar Pago"
        if dias <= 60: return "🚫 BLOQUEO DE CUPO + Llamada Firme"
        if dias <= 90: return "⚠️ Carta Pre-Jurídica + Visita Comercial"
        if dias <= 120: return "⚖️ Conciliación Urgente / Acuerdo de Pago"
        
        # Casos Extremos
        if dias > 120 and monto > 1000000: return "💀 Traslado a Abogados / Cobro Jurídico"
        if dias > 360: return "🗑️ Evaluar Castigo de Cartera"
        return "Gestión Administrativa"

    def calcular_prioridad(row):
        # Fórmula de Prioridad (0 a 100)
        # Peso Días: 60%, Peso Monto: 40% (Normalizado logarítmicamente aprox)
        if row['dias_vencido'] <= 0: return 0
        
        score_dias = min(row['dias_vencido'], 180) / 180 * 60
        score_monto = min(row['importe'], 10000000) / 10000000 * 40 
        return score_dias + score_monto

    df_calc['Etapa'] = df_calc.apply(clasificar_riesgo, axis=1)
    df_calc['Accion_Sugerida'] = df_calc.apply(sugerir_accion, axis=1)
    df_calc['Score_Prioridad'] = df_calc.apply(calcular_prioridad, axis=1)
    
    return df_calc

def generar_super_excel(df_filtrado):
    """Genera un Excel visualmente rico con formato condicional y tablas."""
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Plan de Choque"

    # Preparar datos para exportar
    cols_export = [
        'nombrecliente', 'nit', 'nomvendedor', 'telefono1', 
        'Etapa', 'Accion_Sugerida', 'dias_vencido', 'Score_Prioridad', 'importe'
    ]
    
    # Encabezados personalizados
    headers = [
        "Cliente", "NIT", "Vendedor Responsable", "Teléfono", 
        "Etapa de Cobro", "ACCIÓN RECOMENDADA", "Días Vencido", "Nivel Prioridad (0-100)", "Deuda Total"
    ]

    # Escribir Encabezados
    ws.append(headers)
    
    # Escribir Datos
    # Agrupamos por factura o cliente? El usuario pidió gestión, mejor agrupar por Cliente
    df_agrupado = df_filtrado.groupby(
        ['nombrecliente', 'nit', 'nomvendedor', 'telefono1', 'Etapa', 'Accion_Sugerida']
    ).agg({
        'dias_vencido': 'max',
        'Score_Prioridad': 'max',
        'importe': 'sum'
    }).reset_index().sort_values(by='Score_Prioridad', ascending=False)

    for r in df_agrupado[cols_export].itertuples(index=False):
        ws.append(list(r))

    # --- FORMATO EXTRAORDINARIO ---
    
    # 1. Crear Tabla
    max_row = ws.max_row
    max_col = ws.max_column
    tab = Table(displayName="TablaGestion", ref=f"A1:{get_column_letter(max_col)}{max_row}")
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    tab.tableStyleInfo = style
    ws.add_table(tab)

    # 2. Anchos de Columna
    ws.column_dimensions['A'].width = 35 # Cliente
    ws.column_dimensions['B'].width = 15 # NIT
    ws.column_dimensions['C'].width = 25 # Vendedor
    ws.column_dimensions['D'].width = 15 # Tel
    ws.column_dimensions['E'].width = 20 # Etapa
    ws.column_dimensions['F'].width = 35 # Accion
    ws.column_dimensions['G'].width = 12 # Dias
    ws.column_dimensions['H'].width = 15 # Score
    ws.column_dimensions['I'].width = 18 # Importe

    # 3. Formato Condicional: Barras de Datos para "Importe" (Columna I - 9)
    data_bar_rule = DataBarRule(start_type='min', end_type='max', color="638EC6", showValue=True, minLength=None, maxLength=None)
    ws.conditional_formatting.add(f"I2:I{max_row}", data_bar_rule)
    
    # 4. Formato Condicional: Escala de Color para "Score Prioridad" (Columna H - 8)
    color_scale_rule = ColorScaleRule(start_type='min', start_color='63BE7B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='F8696B')
    ws.conditional_formatting.add(f"H2:H{max_row}", color_scale_rule)

    # 5. Formato Condicional: Semáforo para Días Vencido (Columna G - 7)
    icon_set_rule = IconSetRule(icon_style='3TrafficLights1', type='num', values=[0, 30, 60], showValue=None, percent=None, reverse=False)
    ws.conditional_formatting.add(f"G2:G{max_row}", icon_set_rule)

    # 6. Formato de Moneda
    for cell in ws['I']:
        cell.number_format = '"$"#,##0'
        
    wb.save(output)
    return output.getvalue()

# ======================================================================================
# --- INTERFAZ DE USUARIO ---
# ======================================================================================

def main():
    st.title("👮‍♂️ Centro de Comando: Cobranza Estratégica")
    st.markdown("Identifica, prioriza y ejecuta acciones de recuperación de cartera de alto impacto.")

    if 'authentication_status' not in st.session_state or not st.session_state['authentication_status']:
        st.error("⚠️ Por favor inicia sesión en la página principal primero.")
        st.stop()

    df_raw = cargar_datos_inteligentes()
    if df_raw.empty:
        st.warning("No hay datos disponibles. Ve al Tablero Principal y recarga los datos.")
        st.stop()

    # --- PROCESAMIENTO ---
    df_proc = segmentar_cartera(df_raw)
    
    # Filtros Globales (Sidebar)
    with st.sidebar:
        st.header("🎯 Filtros de Enfoque")
        vendedor_filtro = st.selectbox("Vendedor", ["Todos"] + sorted(df_proc['nomvendedor'].dropna().unique().tolist()))
        
        if vendedor_filtro != "Todos":
            df_proc = df_proc[df_proc['nomvendedor'] == vendedor_filtro]
            
        solo_vencido = st.checkbox("Ver SOLO Cartera Vencida", value=True)
        if solo_vencido:
            df_proc = df_proc[df_proc['dias_vencido'] > 0]

    # --- KPI HEADER ---
    col1, col2, col3, col4 = st.columns(4)
    
    total_gestion = df_proc['importe'].sum()
    critico_df = df_proc[df_proc['Etapa'].str.contains('Pre-Jurídica|Jurídica')]
    total_critico = critico_df['importe'].sum()
    clientes_criticos = critico_df['nombrecliente'].nunique()
    
    col1.metric("Cartera a Gestionar (Filtro)", f"${total_gestion:,.0f}")
    col2.metric("🚨 En Riesgo Alto (>60 días)", f"${total_critico:,.0f}", delta_color="inverse", delta="Prioridad Máxima")
    col3.metric("Clientes en Riesgo Alto", f"{clientes_criticos}")
    
    # Mejor vendedor (el que menos debe) vs Peor (más debe) - Informativo
    if not df_proc.empty:
        grouped_v = df_proc.groupby('nomvendedor')['importe'].sum().sort_values()
        col4.metric("Mayor Concentración", grouped_v.index[-1] if not grouped_v.empty else "N/A", f"${grouped_v.iloc[-1]:,.0f}")

    st.markdown("---")

    # --- SECCIÓN 1: EL EMBUDO DE COBRANZA (Visualización Macro) ---
    c1, c2 = st.columns([2, 1])
    
    with c1:
        st.subheader("📡 Radar de Cartera por Etapas")
        df_funnel = df_proc.groupby('Etapa')['importe'].sum().reset_index()
        # Ordenar etapas lógicamente
        orden_etapas = ["🟢 Al Día", "🟡 Preventiva (0-30)", "🟠 Administrativa (31-60)", "🔴 Pre-Jurídica (61-120)", "⚫ Jurídica / Castigo (>120)"]
        df_funnel['Etapa'] = pd.Categorical(df_funnel['Etapa'], categories=orden_etapas, ordered=True)
        df_funnel = df_funnel.sort_values('Etapa')
        
        fig = px.funnel(df_funnel, x='importe', y='Etapa', color='Etapa', 
                        title="Flujo de Dinero Estancado por Etapa",
                        color_discrete_map={
                            "🟢 Al Día": COLORS['safe'], 
                            "🟡 Preventiva (0-30)": "#FFD54F",
                            "🟠 Administrativa (31-60)": COLORS['warning'],
                            "🔴 Pre-Jurídica (61-120)": COLORS['action'],
                            "⚫ Jurídica / Castigo (>120)": "#212121"
                        })
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.subheader("💾 Exportar Plan de Choque")
        st.info("Descarga el 'Super Excel' con semáforos y prioridades listas para imprimir y trabajar.")
        
        excel_data = generar_super_excel(df_proc)
        st.download_button(
            label="📥 DESCARGAR EXCEL MAESTRO",
            data=excel_data,
            file_name="Plan_Maestro_Cobranza.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Genera un archivo .xlsx con formato condicional avanzado.",
            use_container_width=True,
            type="primary"
        )
        st.markdown("### Resumen Rápido")
        st.write(df_funnel[['Etapa', 'importe']].style.format({'importe': '${:,.0f}'}))

    # --- SECCIÓN 2: MATRIZ DE ACCIÓN "A QUIÉN COBRAR YA" ---
    st.markdown("---")
    st.header("🔥 TOP PRIORIDAD: Gestión Inmediata")
    st.markdown("Estos clientes requieren acción **HOY**. Están ordenados por un algoritmo de *Score de Prioridad* (Monto + Antigüedad).")

    # Crear tabla bonita agrupada por cliente
    top_clientes = df_proc.groupby(['nombrecliente', 'Accion_Sugerida', 'nomvendedor']).agg({
        'importe': 'sum',
        'dias_vencido': 'max',
        'Score_Prioridad': 'max',
        'telefono1': 'first' # Tomamos el primer telefono encontrado
    }).reset_index().sort_values(by='Score_Prioridad', ascending=False).head(15)
    
    # Mostramos la tabla con columnas formateadas
    st.dataframe(
        top_clientes.style.format({'importe': '${:,.0f}', 'Score_Prioridad': '{:.1f}'})
        .background_gradient(subset=['Score_Prioridad'], cmap='Reds')
        .bar(subset=['importe'], color='#d65f5f'),
        use_container_width=True,
        column_config={
            "nombrecliente": "Cliente",
            "Accion_Sugerida": "⚡ ACCIÓN REQUERIDA",
            "nomvendedor": "Vendedor",
            "importe": "Deuda Total",
            "dias_vencido": "Días Max.",
            "telefono1": "Contacto"
        }
    )

    # --- SECCIÓN 3: GESTIÓN POR VENDEDOR (¿A quién presionar?) ---
    st.markdown("---")
    st.header("🕵️‍♂️ Gestión Comercial: ¿Qué vendedor necesita apoyo?")
    
    col_v1, col_v2 = st.columns(2)
    
    with col_v1:
        st.markdown("##### Deuda Vencida por Vendedor (Treemap)")
        # Agrupar vendedores y etapas
        df_vend = df_proc[df_proc['dias_vencido'] > 0].groupby(['nomvendedor', 'Etapa'])['importe'].sum().reset_index()
        fig_tree = px.treemap(df_vend, path=['nomvendedor', 'Etapa'], values='importe',
                              color='importe', color_continuous_scale='RdBu_r',
                              title="Tamaño de Bloques de Deuda por Vendedor")
        st.plotly_chart(fig_tree, use_container_width=True)
        
    with col_v2:
        st.markdown("##### Casos de Conciliación (Deudas Antiguas)")
        # Filtro: Más de 90 días vencido
        conciliacion = df_proc[df_proc['dias_vencido'] > 90].groupby(['nombrecliente', 'nomvendedor']).agg({'importe':'sum', 'dias_vencido':'max'}).reset_index().sort_values('dias_vencido', ascending=False)
        
        if not conciliacion.empty:
            st.warning(f"Hay **{len(conciliacion)} clientes** con facturas de más de 90 días. Candidatos a cobro jurídico.")
            st.dataframe(conciliacion.head(10).style.format({'importe': '${:,.0f}'}), use_container_width=True, hide_index=True)
        else:
            st.success("¡Excelente! No hay cartera mayor a 90 días para conciliar.")

if __name__ == '__main__':
    main()
