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
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# --- CONFIGURACIÓN VISUAL PROFESIONAL ---
st.set_page_config(
    page_title="Centro de Mando: Cobranza Ferreinox",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS para limpiar la interfaz y darle toque corporativo
st.markdown("""
<style>
    .main { background-color: #f4f6f9; }
    .stMetric { background-color: white; padding: 15px; border-radius: 8px; border-left: 5px solid #003366; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    div[data-testid="stExpander"] div[role="button"] p { font-size: 1.1rem; font-weight: bold; color: #003366; }
    .css-1d391kg { padding-top: 1rem; }
</style>
""", unsafe_allow_html=True)

# ======================================================================================
# 1. MOTOR DE INGESTIÓN DE DATOS (Inteligencia de Columnas)
# ======================================================================================

def normalizar_texto(texto):
    """Elimina tildes y pone mayúsculas para comparar columnas."""
    if not isinstance(texto, str): return str(texto)
    return unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode("utf-8").upper().strip()

def detectar_columnas(df):
    """Mapea las columnas del ERP a nombres estándar automáticamente."""
    df.columns = [normalizar_texto(c) for c in df.columns]
    
    mapa = {
        'cliente': ['NOMBRE', 'RAZON SOCIAL', 'TERCERO', 'CLIENTE', 'NOMVENDEDOR'], # Ajustar prioridad
        'nit': ['NIT', 'IDENTIFICACION', 'CEDULA', 'RUT'],
        'saldo': ['IMPORTE', 'SALDO', 'TOTAL', 'DEUDA', 'VALOR'],
        'dias': ['DIAS', 'VENCIDO', 'MORA', 'ANTIGUEDAD'],
        'telefono': ['TEL', 'MOVIL', 'CELULAR', 'TELEFONO'],
        'vendedor': ['VENDEDOR', 'ASESOR', 'COMERCIAL', 'NOMVENDEDOR'],
        'factura': ['NUMERO', 'FACTURA', 'DOC', 'SERIE']
    }
    
    renombres = {}
    for standard, variantes in mapa.items():
        for col in df.columns:
            if any(v in col for v in variantes):
                if standard not in renombres.values(): # Evitar duplicados
                    renombres[col] = standard
                break
    
    df.rename(columns=renombres, inplace=True)
    return df

@st.cache_data(ttl=600)
def cargar_datos():
    # 1. Intentar cargar archivo local más reciente
    archivos = glob.glob("*.xlsx") + glob.glob("*.csv")
    if not archivos:
        return None, "No se encontraron archivos Excel/CSV en la carpeta."
    
    archivo = max(archivos, key=os.path.getctime)
    
    try:
        if archivo.endswith('.csv'):
            df = pd.read_csv(archivo, sep=None, engine='python', encoding='latin-1', dtype=str)
        else:
            df = pd.read_excel(archivo, dtype=str)
            
        df = detectar_columnas(df)
        
        # Limpieza de datos duros
        if 'saldo' in df.columns:
            df['saldo'] = df['saldo'].astype(str).str.replace(r'[^\d.-]', '', regex=True)
            df['saldo'] = pd.to_numeric(df['saldo'], errors='coerce').fillna(0)
        
        if 'dias' in df.columns:
            df['dias'] = pd.to_numeric(df['dias'], errors='coerce').fillna(0)
            
        if 'cliente' not in df.columns:
            return None, f"El archivo {archivo} no tiene columna de Cliente identificable."

        df['cliente'] = df['cliente'].fillna("Desconocido")
        df['telefono'] = df['telefono'].fillna("0")
        if 'vendedor' not in df.columns: df['vendedor'] = "General"
        
        # Filtrar saldos irrelevantes
        df = df[df['saldo'] > 1000] 
        
        return df, f"Datos actualizados: {archivo}"
        
    except Exception as e:
        return None, f"Error leyendo {archivo}: {str(e)}"

# ======================================================================================
# 2. CEREBRO DE ESTRATEGIA (Segmentación y Guiones)
# ======================================================================================

def generar_estrategia(row):
    dias = row['dias']
    saldo = row['saldo']
    cliente = str(row['cliente']).split()[0].title()
    
    # Lógica de Semáforo y Guion
    if dias <= 0:
        estado = "🟢 Preventivo"
        accion = "Recordatorio Amable"
        prioridad = 3
        msg = f"Hola {cliente}, saludamos de Ferreinox. Su estado de cuenta está al día. ¡Gracias por su excelente hábito de pago!"
    elif dias <= 30:
        estado = "🟡 Mora Temprana"
        accion = "Gestionar Pago"
        prioridad = 2
        msg = f"Hola {cliente}. En Ferreinox notamos una factura vencida por ${saldo:,.0f} ({int(dias)} días). ¿Nos ayudas con el soporte de pago hoy?"
    elif dias <= 60:
        estado = "🟠 Mora Media"
        accion = "Llamada Administrativa"
        prioridad = 1
        msg = f"IMPORTANTE {cliente}: Saldo pendiente de ${saldo:,.0f} con {int(dias)} días. Agradecemos contactarnos para evitar bloqueo de despachos."
    else:
        estado = "🔴 Crítico/Jurídico"
        accion = "Cobro Imperativo"
        prioridad = 0
        msg = f"URGENTE {cliente}: Cartera en etapa PRE-JURÍDICA. Saldo: ${saldo:,.0f}. Evite costos de abogados y reporte negativo gestionando su pago inmediato."
        
    return pd.Series([estado, accion, prioridad, msg])

# ======================================================================================
# 3. INTERFAZ GRÁFICA (DASHBOARD)
# ======================================================================================

def main():
    col_logo, col_titulo = st.columns([1, 5])
    with col_titulo:
        st.title("🛡️ Centro de Gestión de Cartera")
        st.markdown("**Ferreinox SAS BIC** | Panel de Control Gerencial y Operativo")

    # --- CARGA ---
    df, status = cargar_datos()
    if df is None:
        st.error(status)
        st.info("Por favor sube el archivo 'Cartera.xlsx' o 'Cartera.csv' al directorio.")
        return

    # Aplicar Estrategia
    df[['Estado', 'Accion_Sugerida', 'Prioridad', 'Mensaje_WhatsApp']] = df.apply(generar_estrategia, axis=1)
    
    # --- SIDEBAR: FILTROS ---
    st.sidebar.header("🔍 Filtros de Gestión")
    
    # Filtro Vendedor
    vendedores = ["TODOS"] + sorted(list(df['vendedor'].unique()))
    sel_vendedor = st.sidebar.selectbox("Vendedor / Zona", vendedores)
    if sel_vendedor != "TODOS":
        df = df[df['vendedor'] == sel_vendedor]

    # Filtro Estado
    estados = ["TODOS"] + sorted(list(df['Estado'].unique()))
    sel_estado = st.sidebar.selectbox("Estado de Mora", estados)
    if sel_estado != "TODOS":
        df = df[df['Estado'] == sel_estado]

    st.sidebar.markdown("---")
    st.sidebar.info(f"📁 {status}")

    # --- KPIs SUPERIORES ---
    total = df['saldo'].sum()
    vencido = df[df['dias'] > 0]['saldo'].sum()
    critico = df[df['dias'] > 60]['saldo'].sum()
    clientes_mora = df[df['dias'] > 0]['cliente'].nunique()

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("💰 Cartera Total", f"${total:,.0f}", help="Suma total de facturas")
    k2.metric("⚠️ Total Vencido", f"${vencido:,.0f}", delta="-Cartera en Riesgo", delta_color="inverse")
    k3.metric("🔥 Mora Crítica (>60)", f"${critico:,.0f}", delta="Acción Inmediata", delta_color="inverse")
    k4.metric("👥 Clientes a Gestionar", f"{clientes_mora}", "Clientes con mora > 1 día")

    # --- PESTAÑAS PRINCIPALES ---
    tab_accion, tab_analisis, tab_export = st.tabs(["🚀 GESTIÓN DIARIA (WhatsApp)", "📊 ANÁLISIS GERENCIAL", "📥 DESCARGAR REPORTES"])

    # --------------------------------------------------------
    # TAB 1: GESTIÓN (La herramienta del día a día)
    # --------------------------------------------------------
    with tab_accion:
        st.markdown("### 📋 Lista de Trabajo Priorizada")
        st.caption("Ordenada por urgencia. Usa el botón de WhatsApp para gestionar cobros en un clic.")

        # Preparar datos para la tabla interactiva
        df_display = df.sort_values(by=['Prioridad', 'dias', 'saldo'], ascending=[True, False, False]).copy()
        
        # Generar Enlace WA
        def crear_link(row):
            tel = str(row['telefono']).strip()
            tel = re.sub(r'\D', '', tel)
            if len(tel) < 10: return None
            if not tel.startswith('57'): tel = '57' + tel
            return f"https://wa.me/{tel}?text={quote(row['Mensaje_WhatsApp'])}"
            
        df_display['Link_WA'] = df_display.apply(crear_link, axis=1)

        # Tabla interactiva
        st.data_editor(
            df_display[['cliente', 'dias', 'saldo', 'Estado', 'Link_WA', 'vendedor']],
            column_config={
                "Link_WA": st.column_config.LinkColumn(
                    "📱 Acción",
                    help="Clic para abrir WhatsApp Web con el mensaje precargado",
                    validate="^https://wa\.me/.*",
                    display_text="💬 COBRAR AHORA"
                ),
                "saldo": st.column_config.NumberColumn("Deuda Total", format="$ %d"),
                "dias": st.column_config.NumberColumn("Días Mora", format="%d días"),
                "Estado": st.column_config.TextColumn("Estado", width="medium"),
                "cliente": st.column_config.TextColumn("Cliente", width="large"),
            },
            hide_index=True,
            use_container_width=True,
            height=600
        )

    # --------------------------------------------------------
    # TAB 2: ANÁLISIS (Para el Gerente / Líder)
    # --------------------------------------------------------
    with tab_analisis:
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Distribución por Estado de Mora")
            fig_pie = px.pie(df, values='saldo', names='Estado', hole=0.4, color='Estado',
                             color_discrete_map={
                                 "🟢 Preventivo": "#2ecc71",
                                 "🟡 Mora Temprana": "#f1c40f",
                                 "🟠 Mora Media": "#e67e22",
                                 "🔴 Crítico/Jurídico": "#e74c3c"
                             })
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with c2:
            st.subheader("Top 10 Clientes Morosos")
            df_top = df.sort_values(by='saldo', ascending=False).head(10)
            fig_bar = px.bar(df_top, x='saldo', y='cliente', orientation='h', 
                             text_auto='.2s', color='dias', title="Ranking por Deuda",
                             color_continuous_scale='Reds')
            fig_bar.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig_bar, use_container_width=True)

        

    # --------------------------------------------------------
    # TAB 3: EXPORTACIÓN (Excel Profesional)
    # --------------------------------------------------------
    with tab_export:
        st.subheader("Descarga de Informes")
        col_dl, col_info = st.columns([1, 2])
        
        with col_dl:
            # Generador de Excel Bonito
            def to_excel(df_input):
                output = io.BytesIO()
                workbook = Workbook()
                sheet = workbook.active
                sheet.title = "Cartera Ferreinox"
                
                # Estilos
                header_fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
                header_font = Font(color="FFFFFF", bold=True)
                money_fmt = '"$"#,##0'
                
                # Datos
                cols = ['cliente', 'nit', 'factura', 'fecha_venc', 'dias', 'saldo', 'Estado', 'vendedor', 'telefono']
                # Filtrar solo columnas que existen
                cols = [c for c in cols if c in df_input.columns]
                
                # Headers
                sheet.append([c.upper() for c in cols])
                for cell in sheet[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center')
                
                # Rows
                for row in df_input[cols].itertuples(index=False):
                    sheet.append(row)
                
                # Autoajuste básico y formato moneda
                for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                    # Asumiendo que saldo está en una columna específica, buscamos el índice
                    # Aquí simplificado: buscamos la celda que tenga valor numérico grande
                    pass 

                workbook.save(output)
                return output.getvalue()

            excel_data = to_excel(df)
            st.download_button(
                label="📥 DESCARGAR EXCEL GERENCIAL",
                data=excel_data,
                file_name=f"Cartera_Ferreinox_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
        with col_info:
            st.info("Este reporte descarga la base filtrada actual con formato profesional, lista para enviar a gerencia o imprimir.")

if __name__ == "__main__":
    main()
