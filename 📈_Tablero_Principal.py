# ======================================================================================
# ARCHIVO: 📈_Tablero_Principal.py (v.Final Corregida y Ampliada)
# ======================================================================================
import streamlit as st
import pandas as pd
import toml
import os
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.table import Table, TableStyleInfo
import unicodedata
import re
from datetime import datetime
from fpdf import FPDF
import yagmail # Necesario para enviar correos
from urllib.parse import quote # Necesario para codificar URL de WhatsApp

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Tablero Principal",
    page_icon="📈",
    layout="wide"
)

# --- MEJORA: Paleta de colores centralizada y CSS para un look más profesional ---
PALETA_COLORES = {
    "primario": "#003865",
    "secundario": "#0058A7",
    "acento": "#FFC300",
    "fondo_claro": "#F0F2F6",
    "texto_claro": "#FFFFFF",
    "texto_oscuro": "#31333F",
    "alerta_rojo": "#D32F2F",
    "alerta_naranja": "#F57C00",
    "alerta_amarillo": "#FBC02D",
    "exito_verde": "#388E3C"
}

st.markdown(f"""
<style>
    .stApp {{
        background-color: {PALETA_COLORES['fondo_claro']};
    }}
    .stMetric {{
        background-color: #FFFFFF;
        border-radius: 10px;
        padding: 15px;
        border: 1px solid #CCCCCC;
    }}
    .stTabs [data-baseweb="tab-list"] {{
        gap: 24px;
    }}
    .stTabs [data-baseweb="tab"] {{
        height: 50px;
        white-space: pre-wrap;
        background-color: transparent;
        border-radius: 4px 4px 0px 0px;
        border-bottom: 2px solid #C0C0C0;
    }}
    .stTabs [aria-selected="true"] {{
        border-bottom: 2px solid {PALETA_COLORES['primario']};
        color: {PALETA_COLORES['primario']};
        font-weight: bold;
    }}
</style>
""", unsafe_allow_html=True)


# ======================================================================================
# --- CLASE PDF Y FUNCIONES AUXILIARES ---
# ======================================================================================
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

def normalizar_nombre(nombre: str) -> str:
    if not isinstance(nombre, str): return ""
    nombre = nombre.upper().strip().replace('.', '')
    nombre = ''.join(c for c in unicodedata.normalize('NFD', nombre) if unicodedata.category(c) != 'Mn')
    return ' '.join(nombre.split())

ZONAS_SERIE = { "PEREIRA": [155, 189, 158, 439], "MANIZALES": [157, 238], "ARMENIA": [156] }

def procesar_cartera(df: pd.DataFrame) -> pd.DataFrame:
    df_proc = df.copy()
    df_proc['importe'] = pd.to_numeric(df_proc['importe'], errors='coerce').fillna(0)
    df_proc['numero'] = pd.to_numeric(df_proc['numero'], errors='coerce').fillna(0)
    df_proc.loc[df_proc['numero'] < 0, 'importe'] *= -1
    df_proc['dias_vencido'] = pd.to_numeric(df_proc['dias_vencido'], errors='coerce').fillna(0)
    df_proc['nomvendedor_norm'] = df_proc['nomvendedor'].apply(normalizar_nombre)
    ZONAS_SERIE_STR = {zona: [str(s) for s in series] for zona, series in ZONAS_SERIE.items()}
    def asignar_zona_robusta(valor_serie):
        if pd.isna(valor_serie): return "OTRAS ZONAS"
        numeros_en_celda = re.findall(r'\d+', str(valor_serie))
        if not numeros_en_celda: return "OTRAS ZONAS"
        for zona, series_clave_str in ZONAS_SERIE_STR.items():
            if set(numeros_en_celda) & set(series_clave_str): return zona
        return "OTRAS ZONAS"
    df_proc['zona'] = df_proc['serie'].apply(asignar_zona_robusta)
    bins = [-float('inf'), 0, 15, 30, 60, float('inf')]; labels = ['Al día', '1-15 días', '16-30 días', '31-60 días', 'Más de 60 días']
    df_proc['edad_cartera'] = pd.cut(df_proc['dias_vencido'], bins=bins, labels=labels, right=True)
    return df_proc

def generar_excel_formateado(df: pd.DataFrame):
    output = BytesIO()
    df_export = df[['nombrecliente', 'serie', 'numero', 'fecha_documento', 'fecha_vencimiento', 'importe', 'dias_vencido']].copy()
    for col in ['fecha_documento', 'fecha_vencimiento']: df_export[col] = pd.to_datetime(df_export[col], errors='coerce').dt.strftime('%d/%m/%Y')
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Cartera', startrow=9)
        wb, ws = writer.book, writer.sheets['Cartera']
        try:
            img = XLImage("LOGO FERREINOX SAS BIC 2024.png"); img.anchor = 'A1'; img.width = 390; img.height = 130
            ws.add_image(img)
        except FileNotFoundError: ws['A1'] = "Logo no encontrado."
        fill_red, fill_orange, fill_yellow = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid'), PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid'), PatternFill(start_color='FFF9C4', end_color='FFF9C4', fill_type='solid')
        font_bold, font_green_bold = Font(bold=True), Font(bold=True, color="006400")
        first_data_row, last_data_row = 10, ws.max_row
        tab = Table(displayName="CarteraVendedor", ref=f"A{first_data_row}:G{last_data_row}"); tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        ws.add_table(tab)
        for i, ancho in enumerate([40, 10, 12, 18, 18, 18, 15], 1): ws.column_dimensions[get_column_letter(i)].width = ancho
        importe_col_idx, dias_col_idx, formato_moneda = 6, 7, '"$"#,##0'
        for row_idx, row in enumerate(ws.iter_rows(min_row=first_data_row, max_row=last_data_row), start=first_data_row):
            if row_idx == first_data_row:
                for cell in row: cell.font = font_bold; cell.alignment = Alignment(horizontal='center', vertical='center')
                continue
            row[importe_col_idx - 1].number_format = formato_moneda
            dias_cell = row[dias_col_idx - 1]
            dias = int(dias_cell.value) if str(dias_cell.value).isdigit() else 0
            if dias > 60: dias_cell.fill = fill_red
            elif dias > 30: dias_cell.fill = fill_orange
            elif dias > 0: dias_cell.fill = fill_yellow
            dias_cell.alignment = Alignment(horizontal='center')
        ws[f"E{last_data_row + 2}"] = "Tu cartera total es de:"; ws[f"E{last_data_row + 2}"].font = font_green_bold
        ws[f"F{last_data_row + 2}"] = f"=SUBTOTAL(9,F{first_data_row + 1}:F{last_data_row})"; ws[f"F{last_data_row + 2}"].number_format = formato_moneda; ws[f"F{last_data_row + 2}"].font = font_green_bold
        ws[f"E{last_data_row + 3}"] = "Facturas vencidas por valor de:"; ws[f"E{last_data_row + 3}"].font = font_green_bold
        ws[f"F{last_data_row + 3}"] = f"=SUMPRODUCT((SUBTOTAL(103,OFFSET(F{first_data_row+1},ROW(F{first_data_row+1}:F{last_data_row})-ROW(F{first_data_row+1}),0,1,1)))*(G{first_data_row+1}:G{last_data_row}>0),F{first_data_row+1}:F{last_data_row})"; ws[f"F{last_data_row + 3}"].number_format = formato_moneda; ws[f"F{last_data_row + 3}"].font = font_green_bold
    return output.getvalue()

def generar_pdf_estado_cuenta(datos_cliente: pd.DataFrame):
    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=45)
    pdf.add_page()
    if datos_cliente.empty:
        pdf.set_font('Arial', 'B', 12); pdf.cell(0, 10, 'No se encontraron facturas para este cliente.', 0, 1, 'C')
        return bytes(pdf.output())
    datos_cliente_ordenados = datos_cliente.sort_values(by='fecha_vencimiento', ascending=True)
    info_cliente = datos_cliente_ordenados.iloc[0]
    pdf.set_font('Arial', 'B', 11); pdf.cell(40, 10, 'Cliente:', 0, 0); pdf.set_font('Arial', '', 11); pdf.cell(0, 10, info_cliente['nombrecliente'], 0, 1)
    pdf.set_font('Arial', 'B', 11); pdf.cell(40, 10, 'Codigo de Cliente:', 0, 0); pdf.set_font('Arial', '', 11)
    cod_cliente_str = str(int(info_cliente['cod_cliente'])) if pd.notna(info_cliente['cod_cliente']) else "N/A"
    pdf.cell(0, 10, cod_cliente_str, 0, 1); pdf.ln(5)
    pdf.set_font('Arial', '', 10)
    mensaje = ("Apreciado cliente, a continuación encontrará el detalle de su estado de cuenta a la fecha. "
               "Le invitamos a realizar su revisión y proceder con el pago de los valores vencidos. "
               "Puede realizar su pago de forma fácil y segura a través de nuestro PORTAL DE PAGOS en línea, "
               "cuyo enlace encontrará al final de este documento.")
    pdf.set_text_color(128, 128, 128); pdf.multi_cell(0, 5, mensaje, 0, 'J'); pdf.set_text_color(0, 0, 0); pdf.ln(10)
    pdf.set_font('Arial', 'B', 10); pdf.set_fill_color(0, 56, 101); pdf.set_text_color(255, 255, 255)
    pdf.cell(30, 10, 'Factura', 1, 0, 'C', 1); pdf.cell(40, 10, 'Fecha Factura', 1, 0, 'C', 1)
    pdf.cell(40, 10, 'Fecha Vencimiento', 1, 0, 'C', 1); pdf.cell(40, 10, 'Importe', 1, 1, 'C', 1)
    pdf.set_font('Arial', '', 10)
    total_importe = 0
    for _, row in datos_cliente_ordenados.iterrows():
        pdf.set_text_color(0, 0, 0)
        if row['dias_vencido'] > 0: pdf.set_fill_color(248, 241, 241)
        else: pdf.set_fill_color(255, 255, 255)
        total_importe += row['importe']
        numero_factura_str = str(int(row['numero'])) if pd.notna(row['numero']) else "N/A"
        pdf.cell(30, 10, numero_factura_str, 1, 0, 'C', 1)
        pdf.cell(40, 10, row['fecha_documento'].strftime('%d/%m/%Y'), 1, 0, 'C', 1)
        pdf.cell(40, 10, row['fecha_vencimiento'].strftime('%d/%m/%Y'), 1, 0, 'C', 1)
        pdf.cell(40, 10, f"${row['importe']:,.0f}", 1, 1, 'R', 1)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font('Arial', 'B', 10); pdf.set_fill_color(0, 56, 101); pdf.set_text_color(255, 255, 255)
    pdf.cell(110, 10, 'TOTAL ADEUDADO', 1, 0, 'R', 1)
    pdf.cell(40, 10, f"${total_importe:,.0f}", 1, 1, 'R', 1)
    return bytes(pdf.output())

@st.cache_data
def cargar_y_procesar_datos():
    df = pd.read_excel("Cartera.xlsx")
    if not df.empty: df = df.iloc[:-1]
    df_renamed = df.rename(columns=lambda x: normalizar_nombre(x).lower().replace(' ', '_'))
    df_renamed['serie'] = df_renamed['serie'].astype(str)
    df_filtrado = df_renamed[~df_renamed['serie'].str.contains('W|X', case=False, na=False)]
    df_filtrado['fecha_documento'] = pd.to_datetime(df_filtrado['fecha_documento'], errors='coerce')
    df_filtrado['fecha_vencimiento'] = pd.to_datetime(df_filtrado['fecha_vencimiento'], errors='coerce')
    return procesar_cartera(df_filtrado)
    
def generar_analisis_cartera(kpis: dict):
    """Genera un análisis en texto basado en los KPIs calculados."""
    comentarios = []
    
    if kpis['porcentaje_vencido'] > 30:
        comentarios.append(f"<li>🔴 **Alerta Crítica:** El <b>{kpis['porcentaje_vencido']:.1f}%</b> de la cartera está vencida. Este nivel es preocupante y requiere acciones inmediatas y contundentes.</li>")
    elif kpis['porcentaje_vencido'] > 15:
        comentarios.append(f"<li>🟡 **Advertencia:** Con un <b>{kpis['porcentaje_vencido']:.1f}%</b> de cartera vencida, es un buen momento para intensificar las gestiones de cobro antes de que la situación se deteriore.</li>")
    else:
        comentarios.append(f"<li>🟢 **Saludable:** El porcentaje de cartera vencida (<b>{kpis['porcentaje_vencido']:.1f}%</b>) está en un nivel manejable y saludable. ¡Buen trabajo!</li>")

    if kpis['antiguedad_prom_vencida'] > 60:
        comentarios.append(f"<li>🔴 **Riesgo Alto:** Las deudas vencidas tienen una antigüedad promedio de <b>{kpis['antiguedad_prom_vencida']:.0f} días</b>. El riesgo de incobrabilidad es alto. Se debe priorizar la recuperación de estas cuentas antiguas.</li>")
    elif kpis['antiguedad_prom_vencida'] > 30:
        comentarios.append(f"<li>🟡 **Atención Requerida:** La antigüedad promedio de la cartera vencida es de <b>{kpis['antiguedad_prom_vencida']:.0f} días</b>. Es vital evitar que estas deudas envejezcan más.</li>")

    if kpis['csi'] > 15:
        comentarios.append(f"<li>🔴 **Severidad Crítica (CSI: {kpis['csi']:.1f}):** El impacto combinado del monto y la antigüedad de la deuda vencida es muy alto. Esto indica un problema estructural que afecta el flujo de caja.</li>")
    elif kpis['csi'] > 5:
        comentarios.append(f"<li>🟡 **Severidad Moderada (CSI: {kpis['csi']:.1f}):** El índice de severidad sugiere que, aunque la situación es manejable, hay focos de deuda antigua o de alto valor que pesan sobre la cartera total.</li>")
    else:
        comentarios.append(f"<li>🟢 **Severidad Baja (CSI: {kpis['csi']:.1f}):** El impacto ponderado de la deuda vencida es bajo, indicando una buena gestión de cobro y plazos.</li>")

    return "<ul>" + "".join(comentarios) + "</ul>"

# ======================================================================================
# --- BLOQUE PRINCIPAL DE LA APP ---
# ======================================================================================
def main():
    if 'authentication_status' not in st.session_state:
        st.session_state['authentication_status'] = False
        st.session_state['acceso_general'] = False
        st.session_state['vendedor_autenticado'] = None

    if not st.session_state['authentication_status']:
        st.title("Acceso al Tablero de Cartera")
        try:
            general_password = st.secrets["general"]["password"]
            vendedores_secrets = st.secrets["vendedores"]
        except Exception:
            st.error("Error al cargar las contraseñas desde los secretos.")
            st.stop()

        password = st.text_input("Introduce la contraseña:", type="password")
        if st.button("Ingresar"):
            if password == str(general_password):
                st.session_state['authentication_status'] = True
                st.session_state['acceso_general'] = True
                st.session_state['vendedor_autenticado'] = "General"
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
                    st.error("Contraseña incorrecta.")
    else:
        st.title("📊 Tablero de Cartera Ferreinox SAS BIC")
        with st.sidebar:
            try:
                st.image("LOGO FERREINOX SAS BIC 2024.png", use_container_width=True)
            except FileNotFoundError:
                st.warning("Logo no encontrado.")
            st.success(f"Usuario: {st.session_state['vendedor_autenticado']}")
            if st.button("Cerrar Sesión"):
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()
        
        try:
            cartera_procesada = cargar_y_procesar_datos()
        except FileNotFoundError: st.error("No se encontró el archivo 'Cartera.xlsx'."); st.stop()
        except Exception as e: st.error(f"Error al cargar o procesar 'Cartera.xlsx': {e}."); st.stop()

        st.sidebar.title("Filtros")
        if st.session_state['acceso_general']:
            vendedores_en_excel_display = ["Todos"] + sorted(cartera_procesada['nomvendedor'].dropna().unique())
            vendedor_sel = st.sidebar.selectbox("Filtrar por Vendedor:", vendedores_en_excel_display)
        else:
            vendedor_sel = st.session_state['vendedor_autenticado']
        
        lista_zonas = ["Todas las Zonas"] + list(ZONAS_SERIE.keys())
        zona_sel = st.sidebar.selectbox("Filtrar por Zona:", lista_zonas)
        lista_poblaciones = ["Todas"] + sorted(cartera_procesada['poblacion'].dropna().unique())
        poblacion_sel = st.sidebar.selectbox("Filtrar por Población:", lista_poblaciones)

        if vendedor_sel == "Todos": cartera_filtrada = cartera_procesada.copy()
        else: cartera_filtrada = cartera_procesada[cartera_procesada['nomvendedor_norm'] == normalizar_nombre(vendedor_sel)].copy()
        if zona_sel != "Todas las Zonas": cartera_filtrada = cartera_filtrada[cartera_filtrada['zona'] == zona_sel]
        if poblacion_sel != "Todas": cartera_filtrada = cartera_filtrada[cartera_filtrada['poblacion'] == poblacion_sel]

        if cartera_filtrada.empty:
            st.warning(f"No se encontraron datos para los filtros seleccionados."); st.stop()
            
        total_cartera = cartera_filtrada['importe'].sum()
        cartera_vencida_df = cartera_filtrada[cartera_filtrada['dias_vencido'] > 0]
        total_vencido = cartera_vencida_df['importe'].sum()
        porcentaje_vencido = (total_vencido / total_cartera) * 100 if total_cartera > 0 else 0
        if total_cartera > 0:
            csi = (cartera_vencida_df['importe'] * cartera_vencida_df['dias_vencido']).sum() / total_cartera
        else:
            csi = 0
        
        antiguedad_prom_vencida = (cartera_vencida_df['importe'] * cartera_vencida_df['dias_vencido']).sum() / total_vencido if total_vencido > 0 else 0
        
        st.header("Indicadores Clave de Rendimiento (KPIs)")
        kpi_cols = st.columns(5)
        kpi_cols[0].metric("💰 Cartera Total", f"${total_cartera:,.0f}")
        kpi_cols[1].metric("🔥 Cartera Vencida", f"${total_vencido:,.0f}", help="Suma del importe de facturas con días de vencimiento > 0.")
        kpi_cols[2].metric("📈 % Vencido s/ Total", f"{porcentaje_vencido:.1f}%")
        kpi_cols[3].metric("⏳ Antigüedad Prom. Vencida", f"{antiguedad_prom_vencida:.0f} días", help="Edad promedio ponderada, solo de facturas YA VENCIDAS.")
        kpi_cols[4].metric(label="💥 Índice de Severidad (CSI)", value=f"{csi:.1f}", help="Impacto ponderado de la deuda vencida sobre la cartera total. Un número más alto indica mayor riesgo.")

        with st.expander("🤖 **Análisis y Recomendaciones del Asistente IA**", expanded=True):
            kpis_dict = {
                'porcentaje_vencido': porcentaje_vencido,
                'antiguedad_prom_vencida': antiguedad_prom_vencida,
                'csi': csi
            }
            analisis = generar_analisis_cartera(kpis_dict)
            st.markdown(analisis, unsafe_allow_html=True)
        st.markdown("---")

        tab1, tab2, tab3 = st.tabs(["📊 Visión General de la Cartera", "👥 Análisis por Cliente", "📑 Detalle Completo"])

        with tab1:
            st.subheader("Distribución de Cartera por Antigüedad")
            col_grafico, col_tabla_resumen = st.columns([2, 1])
            with col_grafico:
                df_edades = cartera_filtrada.groupby('edad_cartera', observed=True)['importe'].sum().reset_index()
                color_map_edades = {'Al día': PALETA_COLORES['exito_verde'], '1-15 días': PALETA_COLORES['alerta_amarillo'], '16-30 días': PALETA_COLORES['alerta_naranja'], '31-60 días': 'darkorange', 'Más de 60 días': PALETA_COLORES['alerta_rojo']}
                fig = px.bar(df_edades, x='edad_cartera', y='importe', text_auto='.2s', title='Monto de Cartera por Rango de Días', labels={'edad_cartera': 'Antigüedad', 'importe': 'Monto Total'}, color='edad_cartera', color_discrete_map=color_map_edades)
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
            with col_tabla_resumen:
                st.subheader("Resumen por Antigüedad")
                df_edades['Porcentaje'] = (df_edades['importe'] / total_cartera * 100).map('{:.1f}%'.format) if total_cartera > 0 else '0.0%'
                df_edades['importe'] = df_edades['importe'].map('${:,.0f}'.format)
                st.dataframe(df_edades.rename(columns={'edad_cartera': 'Rango', 'importe': 'Monto'}), use_container_width=True, hide_index=True)
        
        with tab2:
            st.subheader("Análisis de Concentración de Deuda por Cliente")
            col_pareto, col_treemap = st.columns(2)
            
            with col_treemap:
                st.markdown("**Visualización de Cartera Vencida por Cliente (Treemap)**")
                df_clientes_vencidos = cartera_vencida_df.groupby('nombrecliente')['importe'].sum().reset_index()
                df_clientes_vencidos = df_clientes_vencidos[df_clientes_vencidos['importe'] > 0]
                
                fig_treemap = px.treemap(df_clientes_vencidos, path=[px.Constant("Clientes con Deuda Vencida"), 'nombrecliente'], values='importe',
                                         title='Haga clic en un recuadro para explorar',
                                         color_continuous_scale='Reds',
                                         color='importe')
                fig_treemap.update_layout(margin = dict(t=50, l=25, r=25, b=25))
                st.plotly_chart(fig_treemap, use_container_width=True)

            with col_pareto:
                st.markdown("**Clientes Clave (Principio de Pareto)**")
                client_debt = cartera_vencida_df.groupby('nombrecliente')['importe'].sum().sort_values(ascending=False)
                if not client_debt.empty:
                    client_debt_cumsum = client_debt.cumsum()
                    total_debt_vencida = client_debt.sum()
                    pareto_limit = total_debt_vencida * 0.80
                    pareto_clients = client_debt[client_debt_cumsum <= pareto_limit]
                    
                    num_total_clientes_deuda = len(client_debt)
                    num_clientes_pareto = len(pareto_clients)
                    porcentaje_clientes_pareto = (num_clientes_pareto / num_total_clientes_deuda) * 100 if num_total_clientes_deuda > 0 else 0

                    st.info(f"El **{porcentaje_clientes_pareto:.0f}%** de los clientes ({num_clientes_pareto} de {num_total_clientes_deuda}) representan aprox. el **80%** del total de la cartera vencida. Estos son:")
                    
                    df_pareto_display = pareto_clients.reset_index()
                    df_pareto_display.columns = ['Cliente', 'Monto Vencido']
                    df_pareto_display['Monto Vencido'] = df_pareto_display['Monto Vencido'].map('${:,.0f}'.format)
                    st.dataframe(df_pareto_display, height=250, hide_index=True, use_container_width=True)
                else:
                    st.info("No hay cartera vencida para analizar.")


        with tab3:
            st.subheader(f"Detalle Completo: {vendedor_sel} / {zona_sel} / {poblacion_sel}")
            st.download_button(label="📥 Descargar Reporte en Excel con Formato", data=generar_excel_formateado(cartera_filtrada), file_name=f'Cartera_{normalizar_nombre(vendedor_sel)}_{zona_sel}_{poblacion_sel}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            columnas_a_ocultar = ['provincia', 'telefono1', 'telefono2', 'entidad_autoriza', 'e_mail', 'descuento', 'cupo_aprobado', 'nomvendedor_norm', 'zona']
            cartera_para_mostrar = cartera_filtrada.drop(columns=columnas_a_ocultar, errors='ignore')
            st.dataframe(cartera_para_mostrar, use_container_width=True, hide_index=True)
            
        st.markdown("---")
        st.header("⚙️ Herramientas de Gestión")
        st.subheader("Generar y Enviar Estado de Cuenta por Cliente")
        lista_clientes = sorted(cartera_filtrada['nombrecliente'].dropna().unique())
        if not lista_clientes:
            st.warning("No hay clientes para mostrar con los filtros actuales.")
        else:
            cliente_seleccionado = st.selectbox("Busca y selecciona un cliente para gestionar su cuenta:", [""] + lista_clientes, format_func=lambda x: 'Selecciona un cliente...' if x == "" else x, key="cliente_selector")
            
            if cliente_seleccionado:
                datos_cliente_seleccionado = cartera_filtrada[cartera_filtrada['nombrecliente'] == cliente_seleccionado].copy()
                info_cliente_raw = datos_cliente_seleccionado.iloc[0]
                
                # Extraer datos clave del cliente
                correo_cliente = info_cliente_raw.get('e_mail', 'Correo no disponible')
                telefono_cliente = str(info_cliente_raw.get('telefono1', '')).split('.')[0]
                nit_cliente = str(info_cliente_raw.get('nit', 'N/A'))
                cod_cliente = str(int(info_cliente_raw['cod_cliente'])) if pd.notna(info_cliente_raw['cod_cliente']) else "N/A"

                st.write(f"**Facturas para {cliente_seleccionado}:**")
                st.dataframe(datos_cliente_seleccionado[['numero', 'fecha_documento', 'fecha_vencimiento', 'dias_vencido', 'importe']], use_container_width=True, hide_index=True)
                
                pdf_bytes = generar_pdf_estado_cuenta(datos_cliente_seleccionado)
                st.download_button(label="📄 Descargar Estado de Cuenta (PDF)", data=pdf_bytes, file_name=f"Estado_Cuenta_{normalizar_nombre(cliente_seleccionado).replace(' ', '_')}.pdf", mime="application/pdf")

                st.markdown("---")
                st.subheader("✉️ Enviar por Correo Electrónico")

                email_destino = st.text_input("Verificar o modificar correo del cliente:", value=correo_cliente)
                
                if st.button("📧 Enviar Correo con Estado de Cuenta"):
                    if not email_destino or email_destino == 'Correo no disponible':
                        st.error("Dirección de correo no válida o no disponible.")
                    else:
                        try:
                            # Cargar credenciales desde secrets
                            sender_email = st.secrets["email_credentials"]["sender_email"]
                            sender_password = st.secrets["email_credentials"]["sender_password"]

                            # Contenido del correo
                            asunto = f"Estado de Cuenta - Ferreinox SAS BIC - {cliente_seleccionado}"
                            cuerpo_html = f"""
                            <html>
                            <head>
                                <style>
                                    body {{ font-family: Arial, sans-serif; color: #333; }}
                                    .container {{ padding: 20px; border: 1px solid #ddd; border-radius: 8px; max-width: 600px; margin: auto; }}
                                    .header {{ color: #003865; font-size: 24px; font-weight: bold; }}
                                    .content {{ margin-top: 20px; }}
                                    .footer {{ margin-top: 30px; font-size: 12px; color: #777; }}
                                    .payment-info {{ background-color: #f0f2f6; padding: 15px; border-radius: 5px; margin-top: 15px; }}
                                    a {{ color: #0058A7; text-decoration: none; font-weight: bold; }}
                                </style>
                            </head>
                            <body>
                                <div class="container">
                                    <p class="header">Hola, {cliente_seleccionado}</p>
                                    <div class="content">
                                        <p>Recibe un cordial saludo de parte del equipo de Ferreinox SAS BIC.</p>
                                        <p>Adjunto a este correo, encontrarás tu estado de cuenta detallado a la fecha.</p>
                                        <p>Te invitamos a revisarlo y gestionar el pago de los valores pendientes a la brevedad posible.</p>
                                        <div class="payment-info">
                                            <p><b>Realiza tu pago de forma fácil y segura en nuestro Portal de Pagos en línea:</b></p>
                                            <p><a href="https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/">Acceder al Portal de Pagos</a></p>
                                            <p><b>Instrucciones de acceso:</b></p>
                                            <ul>
                                                <li><b>Usuario:</b> {nit_cliente} (Tu NIT)</li>
                                                <li><b>Código Único Interno:</b> {cod_cliente} (Tu Código de Cliente)</li>
                                            </ul>
                                        </div>
                                    </div>
                                    <div class="footer">
                                        <p>Si tienes alguna duda o ya realizaste el pago, por favor haz caso omiso a este mensaje o contáctanos.</p>
                                        <p>Agradecemos tu confianza y preferencia.</p>
                                        <p><b>Ferreinox SAS BIC</b></p>
                                    </div>
                                </div>
                            </body>
                            </html>
                            """
                            
                            with st.spinner(f"Enviando correo a {email_destino}..."):
                                yag = yagmail.SMTP(sender_email, sender_password)
                                yag.send(
                                    to=email_destino,
                                    subject=asunto,
                                    contents=cuerpo_html,
                                    attachments=[(f"Estado_Cuenta_{normalizar_nombre(cliente_seleccionado).replace(' ', '_')}.pdf", BytesIO(pdf_bytes))]
                                )
                            st.success(f"¡Correo enviado exitosamente a {email_destino}!")

                        except Exception as e:
                            st.error(f"Error al enviar el correo: {e}")
                            st.error("Asegúrate de haber configurado correctamente tus credenciales en el archivo secrets.toml y de tener una 'contraseña de aplicación' de Gmail.")
                
                st.markdown("---")
                st.subheader("📲 Enviar Recordatorio por WhatsApp")

                facturas_vencidas_cliente = datos_cliente_seleccionado[datos_cliente_seleccionado['dias_vencido'] > 0]
                if not facturas_vencidas_cliente.empty:
                    total_vencido_cliente = facturas_vencidas_cliente['importe'].sum()
                    dias_max_vencido = int(facturas_vencidas_cliente['dias_vencido'].max())
                    
                    mensaje_whatsapp = (
                        f"👋 ¡Hola {cliente_seleccionado}! Te saludamos desde Ferreinox SAS BIC para recordarte sobre tus facturas vencidas.\n\n"
                        f"Suma total vencida: *${total_vencido_cliente:,.0f}*\n"
                        f"Tu factura más antigua tiene: *{dias_max_vencido} días* de vencida.\n\n"
                        f"Puedes ponerte al día de forma fácil y segura en nuestro Portal de Pagos en línea:\n"
                        f"🔗 https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/\n\n"
                        f"Para ingresar, usa estos datos:\n"
                        f"👤 *Usuario:* {nit_cliente}\n"
                        f"🔑 *Código Único Interno:* {cod_cliente}\n\n"
                        f"¡Agradecemos tu pronta gestión!"
                    )
                    
                    mensaje_codificado = quote(mensaje_whatsapp)
                    
                    # Se asume que el número en la base de datos no tiene el '+' y necesita el código de país (57 para Colombia)
                    if telefono_cliente and telefono_cliente.isdigit():
                        numero_completo = f"57{telefono_cliente}"
                        url_whatsapp = f"https://wa.me/{numero_completo}?text={mensaje_codificado}"
                        st.markdown(f'<a href="{url_whatsapp}" target="_blank" class="button">📱 Enviar a WhatsApp ({telefono_cliente})</a>', unsafe_allow_html=True)
                    else:
                        st.warning(f"No se encontró un número de teléfono válido para este cliente. (Encontrado: {telefono_cliente})")

                else:
                    st.info("Este cliente no tiene facturas vencidas. ¡Excelente!")


if __name__ == '__main__':
    main()
