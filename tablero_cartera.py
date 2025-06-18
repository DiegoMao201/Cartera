# ======================================================================================
# --- LIBRERÍAS Y CONFIGURACIÓN INICIAL ---
# ======================================================================================
import streamlit as st
import pandas as pd
import toml # Solo para la opción de prueba local si no se usa st.secrets
import os
from io import BytesIO
import plotly.express as px
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.worksheet.table import Table, TableStyleInfo
import unicodedata

# Configuración inicial de la página de Streamlit
st.set_page_config(
    page_title="Tablero de Cartera Ferreinox",
    page_icon="📊",
    layout="wide"
)

# ======================================================================================
# --- FUNCIONES AUXILIARES ---
# ======================================================================================

def normalizar_nombre(nombre: str) -> str:
    """Limpia y estandariza un nombre para consistencia (mayúsculas, sin tildes, etc.)."""
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
    """Añade columnas calculadas y de clasificación al DataFrame principal."""
    df_proc = df.copy()
    
    # Asegurar tipos de datos numéricos correctos
    df_proc['importe'] = pd.to_numeric(df_proc['importe'], errors='coerce').fillna(0)
    df_proc['dias_vencido'] = pd.to_numeric(df_proc['dias_vencido'], errors='coerce').fillna(0)
    
    # Crear columna normalizada para el nombre del vendedor para comparaciones robustas
    df_proc['nomvendedor_norm'] = df_proc['nomvendedor'].apply(normalizar_nombre)
    
    # Clasificar la cartera por rangos de antigüedad
    bins = [-float('inf'), 0, 15, 30, 60, float('inf')]
    labels = ['Al día', '1-15 días', '16-30 días', '31-60 días', 'Más de 60 días']
    df_proc['edad_cartera'] = pd.cut(df_proc['dias_vencido'], bins=bins, labels=labels, right=True)
    
    return df_proc

def generar_excel_formateado(df: pd.DataFrame):
    """Crea un archivo Excel en memoria con formato avanzado para descargar."""
    output = BytesIO()
    df_export = df[['nombrecliente', 'serie', 'numero', 'fecha_documento', 'fecha_vencimiento', 'importe', 'dias_vencido']].copy()
    
    for col in ['fecha_documento', 'fecha_vencimiento']:
        df_export[col] = pd.to_datetime(df_export[col], errors='coerce').dt.strftime('%d/%m/%Y')

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Cartera', startrow=9)
        
        wb = writer.book
        ws = writer.sheets['Cartera']
        
        try:
            img = XLImage("LOGO FERREINOX SAS BIC 2024.png")
            img.anchor = 'A1'
            img.width = 390
            img.height = 130
            ws.add_image(img)
        except FileNotFoundError:
            ws['A1'] = "Logo no encontrado. Asegúrate que 'LOGO FERREINOX SAS BIC 2024.png' esté en el directorio."

        fill_red = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
        fill_orange = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
        fill_yellow = PatternFill(start_color='FFF9C4', end_color='FFF9C4', fill_type='solid')
        font_bold = Font(bold=True)
        font_green_bold = Font(bold=True, color="006400")
        
        first_data_row = 10
        last_data_row = ws.max_row
        tab = Table(displayName="CarteraVendedor", ref=f"A{first_data_row}:G{last_data_row}")
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        tab.tableStyleInfo = style
        ws.add_table(tab)
        
        anchos = [40, 10, 12, 18, 18, 18, 15]
        for i, ancho in enumerate(anchos, 1):
            ws.column_dimensions[get_column_letter(i)].width = ancho
            
        importe_col_idx, dias_col_idx = 6, 7
        formato_moneda = '"$"#,##0'
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=first_data_row, max_row=last_data_row), start=first_data_row):
            if row_idx == first_data_row:
                for cell in row:
                    cell.font = font_bold
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                continue
            
            row[importe_col_idx - 1].number_format = formato_moneda
            dias_cell = row[dias_col_idx - 1]
            dias = int(dias_cell.value) if str(dias_cell.value).isdigit() else 0
            
            if dias > 60: dias_cell.fill = fill_red
            elif dias > 30: dias_cell.fill = fill_orange
            elif dias > 0: dias_cell.fill = fill_yellow
            dias_cell.alignment = Alignment(horizontal='center')

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

    return output.getvalue()

@st.cache_data
def cargar_y_procesar_datos():
    """Lee el archivo Excel, elimina la fila de total y lo procesa."""
    cartera_df = pd.read_excel("Cartera.xlsx")
    
    # IMPORTANTE: Eliminamos la última fila del DataFrame, que corresponde al total.
    cartera_df = cartera_df.iloc[:-1]
    
    # El resto del procesamiento continúa con los datos ya limpios
    cartera_df = cartera_df.rename(columns=lambda x: normalizar_nombre(x).lower().replace(' ', '_'))
    cartera_proc = procesar_cartera(cartera_df)
    return cartera_proc

# ======================================================================================
# --- AUTENTICACIÓN Y CARGA DE DATOS PRINCIPAL ---
# ======================================================================================

# Carga de contraseñas desde st.secrets (funciona en local y en la nube)
try:
    general_password = st.secrets["general"]["password"]
    vendedores_secrets = st.secrets["vendedores"]
except Exception as e:
    st.error("Error al cargar las contraseñas desde los secretos.")
    st.info("Asegúrate de tener el archivo .streamlit/secrets.toml configurado correctamente si pruebas en local.")
    st.stop()

# Formulario de contraseña
password = st.text_input("Introduce la contraseña para acceder a la cartera:", type="password")

if not password:
    st.warning("Debes ingresar una contraseña para continuar.")
    st.stop()

# Verificación de contraseña
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

# ======================================================================================
# --- CARGA Y FILTRADO DE DATOS ---
# ======================================================================================

st.title("📊 Tablero de Cartera Ferreinox SAS BIC")

try:
    cartera_procesada = cargar_y_procesar_datos()
except FileNotFoundError:
    st.error("No se encontró el archivo 'Cartera.xlsx'. Asegúrate de que está en el mismo directorio.")
    st.stop()
except Exception as e:
    st.error(f"Error al cargar o procesar 'Cartera.xlsx': {e}.")
    st.stop()

# --- Filtros en la barra lateral ---
st.sidebar.title("Filtros")
vendedores_en_excel_display = sorted(cartera_procesada['nomvendedor'].dropna().unique())

if acceso_general:
    vendedor_sel = st.sidebar.selectbox("Selecciona el Vendedor:", ["Todos"] + vendedores_en_excel_display)
else:
    vendedor_autenticado_norm = normalizar_nombre(vendedor_autenticado)
    vendedores_en_excel_norm = cartera_procesada['nomvendedor_norm'].dropna().unique()

    if vendedor_autenticado_norm not in vendedores_en_excel_norm:
        st.error(f"¡Error de coincidencia! El vendedor '{vendedor_autenticado}' no se encontró en 'Cartera.xlsx'.")
        st.info("Verifica que el nombre de usuario en la configuración de 'Secrets' coincida con un vendedor en el Excel.")
        st.stop()
    
    vendedor_sel = vendedor_autenticado
    st.sidebar.success("Mostrando cartera de:")
    st.sidebar.write(f"**{vendedor_sel}**")

# --- Filtrado final del DataFrame ---
if vendedor_sel == "Todos":
    cartera_filtrada = cartera_procesada.copy()
else:
    vendedor_sel_norm = normalizar_nombre(vendedor_sel)
    cartera_filtrada = cartera_procesada[cartera_procesada['nomvendedor_norm'] == vendedor_sel_norm].copy()

# ======================================================================================
# --- RENDERIZADO DEL TABLERO ---
# ======================================================================================

if cartera_filtrada.empty:
    st.warning(f"No se encontraron datos de cartera para la selección actual ('{vendedor_sel}').")
    st.stop()

st.markdown("---")

# --- KPIs o Métricas Principales ---
total_cartera = cartera_filtrada['importe'].sum()
cartera_vencida = cartera_filtrada[cartera_filtrada['dias_vencido'] > 0]
total_vencido = cartera_vencida['importe'].sum()
porcentaje_vencido = (total_vencido / total_cartera) * 100 if total_cartera > 0 else 0

col1, col2, col3 = st.columns(3)
col1.metric("💰 Cartera Total", f"${total_cartera:,.0f}")
col2.metric("🔥 Cartera Vencida", f"${total_vencido:,.0f}")
col3.metric("📈 % Vencido s/ Total", f"{porcentaje_vencido:.1f}%")

st.markdown("---")

# --- Gráficos y Resumen por Antigüedad ---
col_grafico, col_tabla_resumen = st.columns([2, 1])
with col_grafico:
    st.subheader("Distribución de Cartera por Antigüedad")
    df_edades = cartera_filtrada.groupby('edad_cartera')['importe'].sum().reset_index()
    fig = px.bar(
        df_edades, x='edad_cartera', y='importe', text_auto='.2s', title='Monto de Cartera por Rango de Días',
        labels={'edad_cartera': 'Antigüedad', 'importe': 'Monto Total'}, color='edad_cartera',
        color_discrete_map={
             'Al día': 'green', '1-15 días': '#FFD700', '16-30 días': 'orange',
             '31-60 días': 'darkorange', 'Más de 60 días': 'red'
        }
    )
    fig.update_layout(xaxis_title=None, yaxis_title="Monto ($)", showlegend=False)
    st.plotly_chart(fig, use_container_width=True)

with col_tabla_resumen:
    st.subheader("Resumen por Antigüedad")
    df_edades['Porcentaje'] = (df_edades['importe'] / total_cartera * 100).map('{:.1f}%'.format)
    df_edades['importe'] = df_edades['importe'].map('${:,.0f}'.format)
    st.dataframe(
        df_edades.rename(columns={'edad_cartera': 'Rango', 'importe': 'Monto'}),
        use_container_width=True, hide_index=True
    )

st.markdown("---")

# --- Tabla de Datos Detallados y Descarga ---
st.subheader(f"Detalle de la Cartera - {vendedor_sel}")

st.download_button(
    label="📥 Descargar Reporte en Excel con Formato",
    data=generar_excel_formateado(cartera_filtrada),
    file_name=f'Cartera_{normalizar_nombre(vendedor_sel).replace(" ", "_")}.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

columnas_a_ocultar = [
    'provincia', 'telefono1', 'telefono2', 'entidad_autoriza', 'e-mail',
    'descuento', 'cupo_aprobado', 'castigada', 'edad_cartera', 'nomvendedor', 'nomvendedor_norm'
]
cartera_para_mostrar = cartera_filtrada.drop(columns=columnas_a_ocultar, errors='ignore')

st.dataframe(cartera_para_mostrar, use_container_width=True, hide_index=True)
cartera_para_mostrar = cartera_filtrada.drop(columns=columnas_a_ocultar, errors='ignore')
st.dataframe(cartera_para_mostrar, use_container_width=True, hide_index=True)
