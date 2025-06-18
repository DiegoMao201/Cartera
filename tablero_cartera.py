# --------------------------------------------------------------------------
# --- LIBRERÍAS ---
# --------------------------------------------------------------------------
import streamlit as st
import pandas as pd
import toml
import os
from io import BytesIO
import plotly.express as px
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.table import Table, TableStyleInfo
import unicodedata

# --------------------------------------------------------------------------
# --- CONFIGURACIÓN DE LA PÁGINA ---
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Tablero de Cartera Ferreinox",
    page_icon="📊",
    layout="wide" # Usar el ancho completo de la página
)

# --------------------------------------------------------------------------
# --- FUNCIONES AUXILIARES ---
# --------------------------------------------------------------------------

def normalizar_nombre(nombre):
    """Limpia y estandariza un nombre para consistencia."""
    if not isinstance(nombre, str):
        return ""
    nombre = nombre.upper().strip().replace('.', '')
    nombre = ''.join(
        c for c in unicodedata.normalize('NFD', nombre)
        if unicodedata.category(c) != 'Mn'
    )
    nombre = ' '.join(nombre.split())
    return nombre

def procesar_cartera(df: pd.DataFrame) -> pd.DataFrame:
    """Añade cálculos y clasificaciones necesarias para el tablero."""
    df_proc = df.copy()
    
    # Asegurar tipos de datos correctos
    df_proc['importe'] = pd.to_numeric(df_proc['importe'], errors='coerce').fillna(0)
    df_proc['dias_vencido'] = pd.to_numeric(df_proc['dias_vencido'], errors='coerce').fillna(0)
    
    # Clasificación por edades de cartera
    bins = [-float('inf'), 0, 15, 30, 60, float('inf')]
    labels = ['Al día', '1-15 días', '16-30 días', '31-60 días', 'Más de 60 días']
    df_proc['edad_cartera'] = pd.cut(df_proc['dias_vencido'], bins=bins, labels=labels, right=True)
    
    return df_proc

def generar_excel_formateado(df: pd.DataFrame):
    """
    Crea un archivo Excel en memoria con formato avanzado, replicando la lógica del script original.
    """
    output = BytesIO()
    df_export = df[['nombrecliente', 'serie', 'numero', 'fecha_documento', 'fecha_vencimiento', 'importe', 'dias_vencido']].copy()
    
    # Formatear fechas como texto para evitar reconversión de Excel
    for col in ['fecha_documento', 'fecha_vencimiento']:
        df_export[col] = pd.to_datetime(df_export[col], errors='coerce').dt.strftime('%d/%m/%Y')

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Cartera', startrow=9)
        
        # --- Aplicar formato avanzado al Excel ---
        wb = writer.book
        ws = writer.sheets['Cartera']
        
        # Insertar logo (asegúrate que 'LOGO FERREINOX SAS BIC 2024.png' esté en la carpeta)
        try:
            img = XLImage("LOGO FERREINOX SAS BIC 2024.png")
            img.anchor = 'A1'
            img.width = 390
            img.height = 130
            ws.add_image(img)
        except FileNotFoundError:
            ws['A1'] = "Logo no encontrado. Asegúrate que el archivo 'LOGO FERREINOX SAS BIC 2024.png' esté en el directorio."

        # Definir estilos de celda
        fill_red = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid') # Rojo > 60
        fill_orange = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid') # Naranja > 30
        fill_yellow = PatternFill(start_color='FFF9C4', end_color='FFF9C4', fill_type='solid') # Amarillo > 0
        font_bold = Font(bold=True)
        font_green_bold = Font(bold=True, color="006400")
        
        # Formato de la tabla
        first_data_row = 10
        last_data_row = ws.max_row
        tab = Table(displayName="CarteraVendedor", ref=f"A{first_data_row}:G{last_data_row}")
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        tab.tableStyleInfo = style
        ws.add_table(tab)
        
        # Ajustar anchos de columna
        anchos = [40, 10, 12, 18, 18, 18, 15]
        for i, ancho in enumerate(anchos, 1):
            ws.column_dimensions[get_column_letter(i)].width = ancho
            
        # Formatear datos y encabezados
        importe_col_idx = 6
        dias_col_idx = 7
        formato_moneda = '"$"#,##0'
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=first_data_row, max_row=last_data_row), start=first_data_row):
            dias_cell = row[dias_col_idx - 1]
            importe_cell = row[importe_col_idx - 1]
            
            # Formato de encabezados
            if row_idx == first_data_row:
                for cell in row:
                    cell.font = font_bold
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                continue
            
            # Formato de celdas de datos
            importe_cell.number_format = formato_moneda
            dias = int(dias_cell.value) if str(dias_cell.value).isdigit() else 0
            
            if dias > 60:
                dias_cell.fill = fill_red
            elif dias > 30:
                dias_cell.fill = fill_orange
            elif dias > 0:
                dias_cell.fill = fill_yellow
            
            dias_cell.alignment = Alignment(horizontal='center')

        # Fórmulas de resumen al final
        ws[f"E{last_data_row + 2}"] = "Tu cartera total es de:"
        ws[f"E{last_data_row + 2}"].font = font_green_bold
        ws[f"F{last_data_row + 2}"] = f"=SUBTOTAL(9,F{first_data_row + 1}:F{last_data_row})"
        ws[f"F{last_data_row + 2}"].number_format = formato_moneda
        ws[f"F{last_data_row + 2}"].font = font_green_bold

        ws[f"E{last_data_row + 3}"] = "Facturas vencidas por valor de:"
        ws[f"E{last_data_row + 3}"].font = font_green_bold
        ws[f"F{last_data_row + 3}"] = f"=SUMPRODUCT((SUBTOTAL(103,OFFSET(F{first_data_row+1},ROW(F{first_data_row+1}:F{last_data_row})-ROW(F{first_data_row+1}),0,1,1)))*(G{first_data_row+1}:G{last_data_row}>0),F{first_data_row+1}:F{last_data_row})"
        ws[f"F{last_data_row + 3}"].number_format = formato_moneda
        ws[f"F{last_data_row + 3}"].font = font_green_bold

    processed_data = output.getvalue()
    return processed_data

# --------------------------------------------------------------------------
# --- AUTENTICACIÓN Y CARGA DE DATOS (YA CORREGIDO) ---
# --------------------------------------------------------------------------
# ... (Este bloque es el mismo que ya teníamos y funciona)
# ... (Pega aquí el código de autenticación completo)
# El código debería empezar con "st.sidebar.subheader("Estado de Archivos")"
# y terminar con "st.stop()" si la contraseña es incorrecta.
# --- INICIO DEL BLOQUE DE AUTENTICACIÓN ---
st.sidebar.subheader("Estado de Archivos")
if os.path.exists("carterasecrets.toml"):
    st.sidebar.success("carterasecrets.toml encontrado.")
else:
    st.sidebar.error("carterasecrets.toml NO ENCONTRADO.")
if os.path.exists("Cartera.xlsx"):
    st.sidebar.success("Cartera.xlsx encontrado.")
else:
    st.sidebar.error("Cartera.xlsx NO ENCONTRADO.")

try:
    secrets = toml.load("carterasecrets.toml")
    general_password = secrets.get("general", {}).get("password")
    vendedores_secrets = secrets.get("vendedores", {})
except FileNotFoundError:
    st.error("Archivo 'carterasecrets.toml' no encontrado.")
    st.stop()
except Exception as e:
    st.error(f"Error al cargar 'carterasecrets.toml': {e}.")
    st.stop()

password = st.text_input("Introduce la contraseña para acceder a la cartera:", type="password")

if not password:
    st.warning("Debes ingresar una contraseña para continuar.")
    st.stop()

acceso_general = False
vendedor_autenticado = None

if password == str(general_password):
    acceso_general = True
else:
    for vendedor_key, pass_vendedor in vendedores_secrets.items():
        if password == str(pass_vendedor):
            vendedor_autenticado = vendedor_key
            break

if not acceso_general and vendedor_autenticado is None:
    st.warning("Contraseña incorrecta. No tienes acceso al tablero.")
    st.stop()
# --- FIN DEL BLOQUE DE AUTENTICACIÓN ---

st.title("📊 Tablero de Cartera Ferreinox SAS BIC")

# --- Carga y procesamiento inicial de datos ---
@st.cache_data # Usamos cache para no recargar el Excel cada vez
def cargar_y_procesar_datos():
    cartera_df = pd.read_excel("Cartera.xlsx")
    cartera_df = cartera_df.rename(columns=lambda x: normalizar_nombre(x).lower().replace(' ', '_'))
    cartera_proc = procesar_cartera(cartera_df)
    return cartera_proc

try:
    cartera_procesada = cargar_y_procesar_datos()
except FileNotFoundError:
    st.error("No se encontró el archivo 'Cartera.xlsx'. Asegúrate de que está en el mismo directorio.")
    st.stop()
except Exception as e:
    st.error(f"Error al cargar o procesar 'Cartera.xlsx': {e}.")
    st.stop()

# --- FILTROS EN LA BARRA LATERAL ---
st.sidebar.title("Filtros")
vendedores_en_excel = sorted(cartera_procesada['nomvendedor'].dropna().unique())

if acceso_general:
    vendedor_sel = st.sidebar.selectbox("Selecciona el Vendedor:", ["Todos"] + vendedores_en_excel)
else:
    vendedor_sel = vendedor_autenticado
    if vendedor_sel not in vendedores_en_excel:
        st.error(f"El vendedor '{vendedor_sel}' no se encuentra en 'Cartera.xlsx'.")
        st.info("Verifica que el nombre de usuario en 'carterasecrets.toml' coincida con el de la columna 'nomvendedor' en Excel.")
        st.stop()
    st.sidebar.success(f"Mostrando cartera de:")
    st.sidebar.write(f"**{vendedor_sel}**")

# --- Filtrado final de datos ---
if vendedor_sel == "Todos":
    cartera_filtrada = cartera_procesada.copy()
else:
    cartera_filtrada = cartera_procesada[cartera_procesada['nomvendedor'] == vendedor_sel].copy()

# --- CUERPO PRINCIPAL DEL TABLERO ---
if cartera_filtrada.empty:
    st.warning(f"No se encontraron datos de cartera para la selección actual ('{vendedor_sel}').")
    st.stop()

st.markdown("---")

# --- SECCIÓN DE MÉTRICAS PRINCIPALES (KPIs) ---
total_cartera = cartera_filtrada['importe'].sum()
cartera_vencida = cartera_filtrada[cartera_filtrada['dias_vencido'] > 0]
total_vencido = cartera_vencida['importe'].sum()
porcentaje_vencido = (total_vencido / total_cartera) * 100 if total_cartera > 0 else 0

col1, col2, col3 = st.columns(3)
col1.metric("💰 Cartera Total", f"${total_cartera:,.0f}")
col2.metric("🔥 Cartera Vencida", f"${total_vencido:,.0f}")
col3.metric("📈 % Vencido s/ Total", f"{porcentaje_vencido:.1f}%")

st.markdown("---")

# --- SECCIÓN DE GRÁFICOS ---
col_grafico, col_tabla_resumen = st.columns([2, 1])

with col_grafico:
    st.subheader("Distribución de Cartera por Antigüedad")
    # Agrupar por edad y sumar importes
    df_edades = cartera_filtrada.groupby('edad_cartera')['importe'].sum().reset_index()
    
    # Crear el gráfico de barras con Plotly
    fig = px.bar(
        df_edades,
        x='edad_cartera',
        y='importe',
        text_auto='.2s',
        title='Monto de Cartera por Rango de Días',
        labels={'edad_cartera': 'Antigüedad', 'importe': 'Monto Total'},
        color='edad_cartera',
        color_discrete_map={
             'Al día': 'green',
             '1-15 días': '#FFD700', # Gold
             '16-30 días': 'orange',
             '31-60 días': 'darkorange',
             'Más de 60 días': 'red'
        }
    )
    fig.update_layout(xaxis_title=None, yaxis_title="Monto ($)", showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

with col_tabla_resumen:
    st.subheader("Resumen por Antigüedad")
    # Formatear la tabla de resumen para mejor visualización
    df_edades['Porcentaje'] = (df_edades['importe'] / total_cartera * 100).map('{:.1f}%'.format)
    df_edades['importe'] = df_edades['importe'].map('${:,.0f}'.format)
    st.dataframe(
        df_edades.rename(columns={'edad_cartera': 'Rango', 'importe': 'Monto'}),
        use_container_width=True,
        hide_index=True
    )

st.markdown("---")

# --- SECCIÓN DE DATOS DETALLADOS Y DESCARGA ---
st.subheader(f"Detalle de la Cartera - {vendedor_sel}")

# Botón de descarga primero, para mayor visibilidad
st.download_button(
    label="📥 Descargar Reporte en Excel con Formato",
    data=generar_excel_formateado(cartera_filtrada),
    file_name=f'Cartera_{normalizar_nombre(vendedor_sel).replace(" ", "_")}.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

# Mostrar la tabla detallada
st.dataframe(cartera_filtrada, use_container_width=True, hide_index=True)
