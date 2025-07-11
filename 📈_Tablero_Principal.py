# ======================================================================================
# ARCHIVO: 📈_Tablero_Principal.py (v.Definitiva con Corrección de Duplicados)
# ======================================================================================
import streamlit as st
import pandas as pd
import toml
import os
from io import BytesIO, StringIO
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
import yagmail
from urllib.parse import quote
import tempfile
import dropbox
import glob

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Tablero Principal",
    page_icon="📈",
    layout="wide"
)

# --- PALETA DE COLORES Y CSS ---
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
    .stApp {{ background-color: {PALETA_COLORES['fondo_claro']}; }}
    .stMetric {{ background-color: #FFFFFF; border-radius: 10px; padding: 15px; border: 1px solid #CCCCCC; }}
    .stTabs [data-baseweb="tab-list"] {{ gap: 24px; }}
    .stTabs [data-baseweb="tab"] {{ height: 50px; white-space: pre-wrap; background-color: transparent; border-radius: 4px 4px 0px 0px; border-bottom: 2px solid #C0C0C0; }}
    .stTabs [aria-selected="true"] {{ border-bottom: 2px solid {PALETA_COLORES['primario']}; color: {PALETA_COLORES['primario']}; font-weight: bold; }}
    div[data-baseweb="input"], div[data-baseweb="select"], div.st-multiselect, div.st-text-area {{ background-color: #FFFFFF; border: 1.5px solid {PALETA_COLORES['secundario']}; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); padding-left: 5px; }}
    .button {{ display: inline-block; padding: 10px 20px; color: white; background-color: #25D366; border-radius: 5px; text-align: center; text-decoration: none; font-weight: bold; }}
</style>
""", unsafe_allow_html=True)


# ======================================================================================
# --- LÓGICA DE CARGA DE DATOS HÍBRIDA ---
# ======================================================================================

@st.cache_data(ttl=600)
def cargar_datos_desde_dropbox():
    """Carga los datos más recientes desde el archivo CSV en Dropbox."""
    try:
        APP_KEY = st.secrets["dropbox"]["app_key"]
        APP_SECRET = st.secrets["dropbox"]["app_secret"]
        REFRESH_TOKEN = st.secrets["dropbox"]["refresh_token"]
        
        with dropbox.Dropbox(app_key=APP_KEY, app_secret=APP_SECRET, oauth2_refresh_token=REFRESH_TOKEN) as dbx:
            path_archivo_dropbox = '/data/cartera_detalle.csv'
            metadata, res = dbx.files_download(path=path_archivo_dropbox)
            contenido_csv = res.content.decode('latin-1')
            
            nombres_columnas_originales = [
                'Serie', 'Numero', 'Fecha Documento', 'Fecha Vencimiento', 'Cod Cliente',
                'NombreCliente', 'Nit', 'Poblacion', 'Provincia', 'Telefono1', 'Telefono2',
                'NomVendedor', 'Entidad Autoriza', 'E-Mail', 'Importe', 'Descuento',
                'Cupo Aprobado', 'Dias Vencido'
            ]
            
            df = pd.read_csv(StringIO(contenido_csv), header=None, names=nombres_columnas_originales, sep='|', engine='python')
            return df
    except Exception as e:
        st.error(f"Error al cargar datos desde Dropbox: {e}")
        return pd.DataFrame()

@st.cache_data
def cargar_datos_historicos():
    """Busca y carga todos los archivos Excel históricos locales."""
    archivos_historicos = glob.glob("Cartera_*.xlsx")
    if not archivos_historicos:
        return pd.DataFrame()

    lista_de_dataframes = []
    for archivo in archivos_historicos:
        try:
            df_hist = pd.read_excel(archivo)
            if not df_hist.empty:
                # La lógica original eliminaba la última fila, la preservamos
                if "Total" in str(df_hist.iloc[-1, 0]):
                    df_hist = df_hist.iloc[:-1]
                lista_de_dataframes.append(df_hist)
        except Exception as e:
            st.warning(f"No se pudo leer el archivo histórico {archivo}: {e}")
            
    if lista_de_dataframes:
        return pd.concat(lista_de_dataframes, ignore_index=True)
    return pd.DataFrame()

@st.cache_data
def cargar_y_procesar_datos():
    """
    Orquesta la carga de datos, los combina, limpia duplicados y procesa.
    """
    df_dropbox = cargar_datos_desde_dropbox()
    df_historico = cargar_datos_historicos()
    
    df_combinado = pd.concat([df_dropbox, df_historico], ignore_index=True)

    if df_combinado.empty:
        st.error("No se pudieron cargar datos de ninguna fuente. La aplicación no puede continuar.")
        st.stop()
        
    # --- SOLUCIÓN AL ERROR TypeError ---
    # Elimina columnas duplicadas que pueden venir de archivos Excel malformados,
    # conservando la primera aparición de cada columna.
    df_combinado = df_combinado.loc[:, ~df_combinado.columns.duplicated()]
    
    # Aplica la misma lógica de procesamiento del código original
    df_renamed = df_combinado.rename(columns=lambda x: normalizar_nombre(x).lower().replace(' ', '_'))
    
    # ***** INICIO DE LA CORRECCIÓN *****
    # Se eliminan las columnas duplicadas NUEVAMENTE después de renombrar.
    # Esto soluciona el TypeError, ya que el paso anterior puede crear nuevos duplicados
    # (ej. 'Importe' y 'importe' se convierten ambos en 'importe').
    df_renamed = df_renamed.loc[:, ~df_renamed.columns.duplicated()]
    # ***** FIN DE LA CORRECCIÓN *****
    
    df_renamed['serie'] = df_renamed['serie'].astype(str)
    df_renamed['fecha_documento'] = pd.to_datetime(df_renamed['fecha_documento'], errors='coerce')
    df_renamed['fecha_vencimiento'] = pd.to_datetime(df_renamed['fecha_vencimiento'], errors='coerce')
    
    df_filtrado = df_renamed[~df_renamed['serie'].str.contains('W|X', case=False, na=False)]
    
    return procesar_cartera(df_filtrado)


# ======================================================================================
# --- CLASE PDF Y FUNCIONES AUXILIARES (LÓGICA ORIGINAL INTACTA) ---
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
    # Estas conversiones ahora se hacen sobre un DataFrame limpio y sin duplicados
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
        tab = Table(displayName="CarteraVendedor", ref=f"A{first_data_row-1}:G{last_data_row}")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        ws.add_table(tab)
        for i, ancho in enumerate([40, 10, 12, 18, 18, 18, 15], 1): ws.column_dimensions[get_column_letter(i)].width = ancho
        importe_col_idx, dias_col_idx, formato_moneda = 6, 7, '"$"#,##0'
        for row_idx, row in enumerate(ws.iter_rows(min_row=first_data_row, max_row=last_data_row), start=first_data_row):
            if row_idx == first_data_row:
                    for cell in row:
                        cell.font = font_bold
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    continue
            row[importe_col_idx - 1].number_format = formato_moneda
            dias_cell = row[dias_col_idx - 1]
            try:
                dias = int(dias_cell.value)
                if dias > 60: dias_cell.fill = fill_red
                elif dias > 30: dias_cell.fill = fill_orange
                elif dias > 0: dias_cell.fill = fill_yellow
            except (ValueError, TypeError):
                pass
            dias_cell.alignment = Alignment(horizontal='center')

        ws[f"E{last_data_row + 2}"] = "Tu cartera total es de:"; ws[f"E{last_data_row + 2}"].font = font_green_bold
        ws[f"F{last_data_row + 2}"] = f"=SUBTOTAL(9,F{first_data_row}:F{last_data_row})"; ws[f"F{last_data_row + 2}"].number_format = formato_moneda; ws[f"F{last_data_row + 2}"].font = font_green_bold
        ws[f"E{last_data_row + 3}"] = "Facturas vencidas por valor de:"; ws[f"E{last_data_row + 3}"].font = font_green_bold
        ws[f"F{last_data_row + 3}"] = f"=SUMIF(G{first_data_row}:G{last_data_row},\">0\",F{first_data_row}:F{last_data_row})"; ws[f"F{last_data_row + 3}"].number_format = formato_moneda; ws[f"F{last_data_row + 3}"].font = font_green_bold
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
        fecha_doc_str = row['fecha_documento'].strftime('%d/%m/%Y') if pd.notna(row['fecha_documento']) else ''
        fecha_ven_str = row['fecha_vencimiento'].strftime('%d/%m/%Y') if pd.notna(row['fecha_vencimiento']) else ''
        pdf.cell(30, 10, numero_factura_str, 1, 0, 'C', 1)
        pdf.cell(40, 10, fecha_doc_str, 1, 0, 'C', 1)
        pdf.cell(40, 10, fecha_ven_str, 1, 0, 'C', 1)
        pdf.cell(40, 10, f"${row['importe']:,.0f}", 1, 1, 'R', 1)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font('Arial', 'B', 10); pdf.set_fill_color(0, 56, 101); pdf.set_text_color(255, 255, 255)
    pdf.cell(110, 10, 'TOTAL ADEUDADO', 1, 0, 'R', 1)
    pdf.cell(40, 10, f"${total_importe:,.0f}", 1, 1, 'R', 1)
    return bytes(pdf.output())

def generar_analisis_cartera(kpis: dict):
    comentarios = []
    if kpis['porcentaje_vencido'] > 30: comentarios.append(f"<li>🔴 **Alerta Crítica:** El <b>{kpis['porcentaje_vencido']:.1f}%</b> de la cartera está vencida. Requiere acciones inmediatas.</li>")
    elif kpis['porcentaje_vencido'] > 15: comentarios.append(f"<li>🟡 **Advertencia:** Con un <b>{kpis['porcentaje_vencido']:.1f}%</b> de cartera vencida, es momento de intensificar gestiones.</li>")
    else: comentarios.append(f"<li>🟢 **Saludable:** El porcentaje de cartera vencida (<b>{kpis['porcentaje_vencido']:.1f}%</b>) está en un nivel manejable.</li>")
    if kpis['antiguedad_prom_vencida'] > 60: comentarios.append(f"<li>🔴 **Riesgo Alto:** Antigüedad promedio de <b>{kpis['antiguedad_prom_vencida']:.0f} días</b>. Priorizar recuperación.</li>")
    elif kpis['antiguedad_prom_vencida'] > 30: comentarios.append(f"<li>🟡 **Atención Requerida:** Antigüedad promedio de <b>{kpis['antiguedad_prom_vencida']:.0f} días</b>. Evitar que envejezcan más.</li>")
    if kpis['csi'] > 15: comentarios.append(f"<li>🔴 **Severidad Crítica (CSI: {kpis['csi']:.1f}):** Impacto muy alto que afecta el flujo de caja.</li>")
    elif kpis['csi'] > 5: comentarios.append(f"<li>🟡 **Severidad Moderada (CSI: {kpis['csi']:.1f}):** Hay focos de deuda antigua o de alto valor que pesan.</li>")
    else: comentarios.append(f"<li>🟢 **Severidad Baja (CSI: {kpis['csi']:.1f}):** Impacto bajo, indicando buena gestión.</li>")
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
        except Exception as e:
            st.error(f"Error al cargar las contraseñas desde los secretos: {e}")
            st.stop()
        password = st.text_input("Introduce la contraseña:", type="password", key="password_input")
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
        
        if st.button("🔄 Recargar Datos (Dropbox + Locales)"):
            st.cache_data.clear()
            st.success("Caché limpiado. Recargando todos los datos...")
            st.rerun()

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

        cartera_procesada = cargar_y_procesar_datos()
        
        st.sidebar.title("Filtros")
        if st.session_state['acceso_general']:
            vendedores_en_excel_display = ["Todos"] + sorted(cartera_procesada['nomvendedor'].dropna().unique())
            vendedor_sel = st.sidebar.selectbox("Filtrar por Vendedor:", vendedores_en_excel_display)
        else:
            vendedor_sel = st.session_state['vendedor_autenticado']
        
        zonas_disponibles = ["Todas las Zonas"] + sorted(cartera_procesada['zona'].dropna().unique())
        zona_sel = st.sidebar.selectbox("Filtrar por Zona:", zonas_disponibles)
        
        poblaciones_disponibles = ["Todas"] + sorted(cartera_procesada['poblacion'].dropna().unique())
        poblacion_sel = st.sidebar.selectbox("Filtrar por Población:", poblaciones_disponibles)

        cartera_filtrada = cartera_procesada.copy()
        if vendedor_sel != "Todos":
            cartera_filtrada = cartera_filtrada[cartera_filtrada['nomvendedor_norm'] == normalizar_nombre(vendedor_sel)]
        if zona_sel != "Todas las Zonas":
            cartera_filtrada = cartera_filtrada[cartera_filtrada['zona'] == zona_sel]
        if poblacion_sel != "Todas":
            cartera_filtrada = cartera_filtrada[cartera_filtrada['poblacion'] == poblacion_sel]

        if cartera_filtrada.empty:
            st.warning(f"No se encontraron datos para los filtros seleccionados."); st.stop()
            
        total_cartera = cartera_filtrada['importe'].sum()
        cartera_vencida_df = cartera_filtrada[cartera_filtrada['dias_vencido'] > 0]
        total_vencido = cartera_vencida_df['importe'].sum()
        porcentaje_vencido = (total_vencido / total_cartera) * 100 if total_cartera > 0 else 0
        csi = (cartera_vencida_df['importe'] * cartera_vencida_df['dias_vencido']).sum() / total_cartera if total_cartera > 0 else 0
        antiguedad_prom_vencida = (cartera_vencida_df['importe'] * cartera_vencida_df['dias_vencido']).sum() / total_vencido if total_vencido > 0 else 0
        
        st.header("Indicadores Clave de Rendimiento (KPIs)")
        kpi_cols = st.columns(5)
        kpi_cols[0].metric("💰 Cartera Total", f"${total_cartera:,.0f}")
        kpi_cols[1].metric("🔥 Cartera Vencida", f"${total_vencido:,.0f}")
        kpi_cols[2].metric("📈 % Vencido s/ Total", f"{porcentaje_vencido:.1f}%")
        kpi_cols[3].metric("⏳ Antigüedad Prom. Vencida", f"{antiguedad_prom_vencida:.0f} días")
        kpi_cols[4].metric(label="💥 Índice de Severidad (CSI)", value=f"{csi:.1f}")

        with st.expander("🤖 **Análisis y Recomendaciones del Asistente IA**", expanded=True):
            kpis_dict = {'porcentaje_vencido': porcentaje_vencido, 'antiguedad_prom_vencida': antiguedad_prom_vencida, 'csi': csi}
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
                fig_treemap = px.treemap(df_clientes_vencidos, path=[px.Constant("Clientes con Deuda Vencida"), 'nombrecliente'], values='importe', title='Haga clic en un recuadro para explorar', color_continuous_scale='Reds', color='importe')
                fig_treemap.update_layout(margin = dict(t=50, l=25, r=25, b=25))
                st.plotly_chart(fig_treemap, use_container_width=True)
            with col_pareto:
                st.markdown("**Clientes Clave (Principio de Pareto)**")
                client_debt = cartera_vencida_df.groupby('nombrecliente')['importe'].sum().sort_values(ascending=False)
                if not client_debt.empty:
                    client_debt_cumsum = client_debt.cumsum()
                    total_debt_vencida = client_debt.sum()
                    pareto_limit = total_debt_vencida * 0.80
                    pareto_clients_df = client_debt.to_frame().iloc[0:len(client_debt_cumsum[client_debt_cumsum <= pareto_limit]) + 1]
                    num_total_clientes_deuda = len(client_debt)
                    num_clientes_pareto = len(pareto_clients_df)
                    porcentaje_clientes_pareto = (num_clientes_pareto / num_total_clientes_deuda) * 100 if num_total_clientes_deuda > 0 else 0
                    st.info(f"El **{porcentaje_clientes_pareto:.0f}%** de los clientes ({num_clientes_pareto} de {num_total_clientes_deuda}) representan aprox. el **80%** de la cartera vencida.")
                    df_pareto_display = pareto_clients_df.reset_index()
                    df_pareto_display.columns = ['Cliente', 'Monto Vencido']
                    df_pareto_display['Monto Vencido'] = df_pareto_display['Monto Vencido'].map('${:,.0f}'.format)
                    st.dataframe(df_pareto_display, height=250, hide_index=True, use_container_width=True)
                else:
                    st.info("No hay cartera vencida para analizar.")
        with tab3:
            st.subheader(f"Detalle Completo: {vendedor_sel} / {zona_sel} / {poblacion_sel}")
            st.download_button(label="📥 Descargar Reporte en Excel", data=generar_excel_formateado(cartera_filtrada), file_name=f'Cartera_{normalizar_nombre(vendedor_sel)}_{zona_sel}_{poblacion_sel}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            columnas_disponibles = cartera_filtrada.columns
            columnas_a_ocultar_existentes = [col for col in ['provincia', 'telefono1', 'telefono2', 'entidad_autoriza', 'e_mail', 'descuento', 'cupo_aprobado', 'nomvendedor_norm', 'zona'] if col in columnas_disponibles]
            cartera_para_mostrar = cartera_filtrada.drop(columns=columnas_a_ocultar_existentes, errors='ignore')
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
                correo_cliente = info_cliente_raw.get('e_mail', 'Correo no disponible')
                telefono_raw = str(info_cliente_raw.get('telefono1', ''))
                telefono_cliente = telefono_raw.split('.')[0] if '.' in telefono_raw else telefono_raw
                nit_cliente = str(info_cliente_raw.get('nit', 'N/A'))
                cod_cliente = str(int(info_cliente_raw['cod_cliente'])) if pd.notna(info_cliente_raw['cod_cliente']) else "N/A"
                st.write(f"**Facturas para {cliente_seleccionado}:**")
                st.dataframe(datos_cliente_seleccionado[['numero', 'fecha_documento', 'fecha_vencimiento', 'dias_vencido', 'importe']], use_container_width=True, hide_index=True)
                pdf_bytes = generar_pdf_estado_cuenta(datos_cliente_seleccionado)
                st.download_button(label="📄 Descargar Estado de Cuenta (PDF)", data=pdf_bytes, file_name=f"Estado_Cuenta_{normalizar_nombre(cliente_seleccionado).replace(' ', '_')}.pdf", mime="application/pdf")
                st.markdown("---")
                col_email, col_whatsapp = st.columns(2)

                with col_email:
                    st.subheader("✉️ Enviar por Correo Electrónico")
                    email_destino = st.text_input("Verificar o modificar correo:", value=correo_cliente)
                    
                    if st.button("📧 Enviar Correo con Estado de Cuenta"):
                        if not email_destino or email_destino == 'Correo no disponible' or '@' not in email_destino:
                            st.error("Dirección de correo no válida o no disponible.")
                        else:
                            try:
                                sender_email = st.secrets["email_credentials"]["sender_email"]
                                sender_password = st.secrets["email_credentials"]["sender_password"]
                                facturas_vencidas_cliente = datos_cliente_seleccionado[datos_cliente_seleccionado['dias_vencido'] > 0]
                                total_vencido_cliente = facturas_vencidas_cliente['importe'].sum()
                                portal_link = "https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/"
                                instrucciones = f"<b>Instrucciones de acceso:</b><br> &nbsp; • <b>Usuario:</b> {nit_cliente} (Tu NIT)<br> &nbsp; • <b>Código Único:</b> {cod_cliente}"
                                
                                if total_vencido_cliente > 0:
                                    dias_max_vencido = int(facturas_vencidas_cliente['dias_vencido'].max())
                                    asunto = f"Recordatorio de Saldo Pendiente - {cliente_seleccionado}"
                                    cuerpo_html = f"""
                                    <html><body style='font-family: Arial, sans-serif; color: #333;'>
                                        <p>Estimado(a) {cliente_seleccionado},</p>
                                        <p>Recibe un cordial saludo del Área de Cartera de Ferreinox SAS BIC.</p>
                                        <p>Nos ponemos en contacto para recordarle su saldo pendiente. Actualmente, tus facturas vencidas suman un total de <b>${total_vencido_cliente:,.0f}</b>, y tu factura más antigua tiene <b>{dias_max_vencido} días</b> de vencida.</p>
                                        <p>Adjunto a este correo, encontrara su estado de cuenta completo para su revisión.</p>
                                        <p>Para su comodidad, puede realizar el pago de forma fácil y segura a través de nuestro <a href='{portal_link}'><b>Portal de Pagos en Línea</b></a>.</p>
                                        <p>{instrucciones}</p>
                                        <p>Si ya has realizado el pago, por favor, haz caso omiso de este recordatorio. Si tienes alguna consulta, no dudes en contactarnos.</p>
                                        <p>Atentamente,<br><b>Area Cartera Ferreinox SAS BIC</b></p>
                                    </body></html>
                                    """
                                else:
                                    asunto = f"Tu Estado de Cuenta actualizado - {cliente_seleccionado}"
                                    cuerpo_html = f"""
                                    <html><body style='font-family: Arial, sans-serif; color: #333;'>
                                        <p>Estimado(a) {cliente_seleccionado},</p>
                                        <p>Recibe un cordial saludo del Área de Cartera de Ferreinox SAS BIC.</p>
                                        <p>Nos complace informarte que tu cuenta se encuentra al día. ¡Agradecemos tu excelente gestión y puntualidad en los pagos!</p>
                                        <p>Para tu control y referencia, adjuntamos a este correo tu estado de cuenta completo.</p>
                                        <p>Recuerda que para futuras consultas o pagos, nuestro <a href='{portal_link}'><b>Portal de Pagos en Línea</b></a> está siempre a tu disposición.</p>
                                        <p>{instrucciones}</p>
                                        <p>Gracias por tu confianza en nosotros.</p>
                                        <p>Atentamente,<br><b>Area Cartera Ferreinox SAS BIC</b></p>
                                    </body></html>
                                    """
                                
                                with st.spinner(f"Enviando correo a {email_destino}..."):
                                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
                                        tmp.write(pdf_bytes)
                                        tmp_path = tmp.name
                                    yag = yagmail.SMTP(sender_email, sender_password)
                                    yag.send(to=email_destino, subject=asunto, contents=cuerpo_html, attachments=tmp_path)
                                    os.remove(tmp_path)
                                st.success(f"¡Correo enviado exitosamente a {email_destino}!")
                            except Exception as e:
                                st.error(f"Error al enviar el correo: {e}")

                with col_whatsapp:
                    st.subheader("📲 Enviar por WhatsApp")
                    numero_completo_para_mostrar = f"+57{telefono_cliente}" if telefono_cliente else "+57"
                    numero_destino_wa = st.text_input("Verificar o modificar número de WhatsApp:", value=numero_completo_para_mostrar)
                    facturas_vencidas_cliente = datos_cliente_seleccionado[datos_cliente_seleccionado['dias_vencido'] > 0]
                    if not facturas_vencidas_cliente.empty:
                        total_vencido_cliente = facturas_vencidas_cliente['importe'].sum()
                        dias_max_vencido = int(facturas_vencidas_cliente['dias_vencido'].max())
                        mensaje_whatsapp = (
                            f"👋 ¡Hola {cliente_seleccionado}! Te saludamos desde Ferreinox SAS BIC.\n\n"
                            f"Tu estado de cuenta con un valor total vencido de *${total_vencido_cliente:,.0f}* ha sido enviado a tu correo. Tu factura más antigua tiene *{dias_max_vencido} días* de vencida.\n\n"
                            f"Para ponerte al día, puedes usar nuestro Portal de Pagos:\n"
                            f"🔗 https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/\n\n"
                            f"Tus datos de acceso son:\n"
                            f"👤 *Usuario:* {nit_cliente} (Tu NIT)\n"
                            f"🔑 *Código Único:* {cod_cliente}\n\n"
                            f"¡Agradecemos tu pronta gestión!"
                        )
                        mensaje_codificado = quote(mensaje_whatsapp)
                        numero_limpio = re.sub(r'\D', '', numero_destino_wa)
                        if numero_limpio:
                            url_whatsapp = f"https://wa.me/{numero_limpio}?text={mensaje_codificado}"
                            st.markdown(f'<a href="{url_whatsapp}" target="_blank" class="button">📱 Enviar a WhatsApp ({numero_destino_wa})</a>', unsafe_allow_html=True)
                        else:
                            st.warning("Ingresa un número de teléfono válido.")
                    else:
                        st.info("Este cliente no tiene facturas vencidas.")

if __name__ == '__main__':
    main()
