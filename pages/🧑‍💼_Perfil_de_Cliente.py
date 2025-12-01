# ======================================================================================
# ARCHIVO: pages/1_🚀_Estrategia_Cobranza.py
# VERSIÓN: FINAL CORREGIDA (Sin errores de duplicados ni imports)
# ======================================================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO, StringIO
import dropbox
import glob
import unicodedata
import re
from datetime import datetime
from urllib.parse import quote  # <--- IMPORTANTE: Necesario para los links de WhatsApp
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="War Room Cobranza", page_icon="🚀", layout="wide")

# --- ESTILOS CSS ---
st.markdown("""
<style>
    .stApp { background-color: #F8F9FA; }
    div.stMetric { background-color: #FFFFFF; border: 1px solid #E0E0E0; border-radius: 8px; padding: 15px; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
    .big-font { font-size: 18px !important; font-weight: bold; color: #333; }
    .action-call { background-color: #FFEBEE; color: #C62828; padding: 4px 8px; border-radius: 4px; font-weight: bold; }
    .action-email { background-color: #E3F2FD; color: #1565C0; padding: 4px 8px; border-radius: 4px; font-weight: bold; }
    .button-wa { text-decoration: none; background-color: #25D366; color: white; padding: 8px 15px; border-radius: 5px; font-weight: bold; display: inline-block; }
</style>
""", unsafe_allow_html=True)

# ======================================================================================
# --- 1. LÓGICA DE CARGA DE DATOS (ROBUSTA Y SIN ERRORES) ---
# ======================================================================================

def normalizar_nombre(nombre: str) -> str:
    if not isinstance(nombre, str): return ""
    nombre = nombre.upper().strip().replace('.', '')
    nombre = ''.join(c for c in unicodedata.normalize('NFD', nombre) if unicodedata.category(c) != 'Mn')
    return ' '.join(nombre.split())

@st.cache_data(ttl=600)
def cargar_datos_maestros():
    """Carga datos de Dropbox y Locales, limpiando duplicados para evitar errores."""
    df_final = pd.DataFrame()
    
    # 1. Intentar Dropbox
    try:
        if "dropbox" in st.secrets:
            APP_KEY = st.secrets["dropbox"]["app_key"]
            APP_SECRET = st.secrets["dropbox"]["app_secret"]
            REFRESH_TOKEN = st.secrets["dropbox"]["refresh_token"]
            
            with dropbox.Dropbox(app_key=APP_KEY, app_secret=APP_SECRET, oauth2_refresh_token=REFRESH_TOKEN) as dbx:
                # Descarga el archivo
                _, res = dbx.files_download(path='/data/cartera_detalle.csv')
                csv_content = res.content.decode('latin-1')
                
                # Nombres de columnas fijos para evitar confusiones
                nombres_cols = ['Serie', 'Numero', 'Fecha Documento', 'Fecha Vencimiento', 'Cod Cliente',
                                'NombreCliente', 'Nit', 'Poblacion', 'Provincia', 'Telefono1', 'Telefono2',
                                'NomVendedor', 'Entidad Autoriza', 'E-Mail', 'Importe', 'Descuento',
                                'Cupo Aprobado', 'Dias Vencido']
                
                df_dropbox = pd.read_csv(StringIO(csv_content), header=None, names=nombres_cols, sep='|', engine='python')
                df_final = pd.concat([df_final, df_dropbox])
    except Exception as e:
        # Si falla Dropbox, no detenemos el código, seguimos con locales
        pass 

    # 2. Intentar Archivos Locales (Excel)
    archivos = glob.glob("Cartera_*.xlsx")
    for archivo in archivos:
        try:
            df_hist = pd.read_excel(archivo)
            if not df_hist.empty:
                # Eliminar fila de totales si existe
                if "Total" in str(df_hist.iloc[-1, 0]): 
                    df_hist = df_hist.iloc[:-1]
                df_final = pd.concat([df_final, df_hist])
        except Exception: 
            pass

    if df_final.empty:
        return pd.DataFrame()

    # --- FASE DE LIMPIEZA CRÍTICA (Evita el TypeError) ---
    
    # 1. Normalizar nombres de columnas a minúsculas y sin espacios
    df_final = df_final.rename(columns=lambda x: normalizar_nombre(x).lower().replace(' ', '_'))
    
    # 2. ELIMINAR COLUMNAS DUPLICADAS (Esta es la corrección clave)
    # Si existen dos columnas llamadas 'importe', esto deja solo la primera.
    df_final = df_final.loc[:, ~df_final.columns.duplicated()]
    
    # 3. Conversión de tipos segura
    if 'importe' in df_final.columns:
        df_final['importe'] = pd.to_numeric(df_final['importe'], errors='coerce').fillna(0)
    else:
        # Si no hay columna importe, creamos una vacía para que no falle
        df_final['importe'] = 0

    if 'dias_vencido' in df_final.columns:
        df_final['dias_vencido'] = pd.to_numeric(df_final['dias_vencido'], errors='coerce').fillna(0)
        
    if 'fecha_vencimiento' in df_final.columns:
        df_final['fecha_vencimiento'] = pd.to_datetime(df_final['fecha_vencimiento'], errors='coerce')
    
    # Filtro de notas crédito si existe la columna serie
    if 'serie' in df_final.columns:
        df_final = df_final[~df_final['serie'].astype(str).str.contains('W|X', case=False, na=False)]
    
    return df_final

# ======================================================================================
# --- 2. CEREBRO: MATRIZ DE RIESGO ---
# ======================================================================================

def aplicar_inteligencia_cobranza(df):
    """Clasifica la cartera y asigna prioridades."""
    if df.empty: return pd.DataFrame()

    df = df.copy()
    # Solo deuda positiva
    df = df[df['importe'] > 0]
    
    # Verificar columnas necesarias
    cols_necesarias = ['nombrecliente', 'nit', 'nomvendedor', 'telefono1', 'e_mail']
    for col in cols_necesarias:
        if col not in df.columns:
            df[col] = "Desconocido" # Rellenar si falta alguna columna

    # Agrupar por Cliente
    cliente_kpis = df.groupby(['nombrecliente', 'nit', 'nomvendedor', 'telefono1', 'e_mail']).agg({
        'importe': 'sum',
        'dias_vencido': 'max',  # Peor vencimiento
        'numero': 'count'       # Cantidad facturas
    }).reset_index()
    
    # Reglas de Negocio
    def determinar_accion(row):
        dias = row['dias_vencido']
        monto = row['importe']
        
        if dias > 90:
            return "🔴 JURÍDICO / PRE-JURÍDICO"
        elif dias > 60:
            return "🟠 CONCILIACIÓN URGENTE"
        elif dias > 30:
            if monto > 5000000: return "🟡 GESTIÓN TELEFÓNICA (Prioridad)"
            else: return "🟡 GESTIÓN ADMINISTRATIVA"
        elif dias > 0:
            return "🟢 RECORDATORIO AMABLE"
        else:
            return "🔵 AL DÍA / PREVENTIVO"

    def calcular_prioridad(row):
        # Score 0-100
        if row['dias_vencido'] <= 0: return 0
        score_dias = min(row['dias_vencido'], 120) / 1.2
        score_monto = min(row['importe'] / 10000000, 1) * 100
        return round((score_dias * 0.6) + (score_monto * 0.4), 1)

    cliente_kpis['Accion_Sugerida'] = cliente_kpis.apply(determinar_accion, axis=1)
    cliente_kpis['Score_Riesgo'] = cliente_kpis.apply(calcular_prioridad, axis=1)
    
    return cliente_kpis.sort_values(by='Score_Riesgo', ascending=False)

# ======================================================================================
# --- 3. EXCEL AVANZADO ---
# ======================================================================================

def generar_excel_estrategico(df_estrategico):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Matriz de Cobro"
    
    headers = ["Prioridad", "Cliente", "NIT", "Vendedor", "Teléfono", "Deuda Total", "Días Vencido (Max)", "# Facturas", "Acción Sugerida", "Score Riesgo"]
    ws.append(headers)
    
    header_fill = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(col_num)].width = 20
    ws.column_dimensions['B'].width = 40

    for index, row in df_estrategico.iterrows():
        ws.append([
            index + 1,
            str(row['nombrecliente']),
            str(row['nit']),
            str(row['nomvendedor']),
            str(row['telefono1']),
            row['importe'],
            row['dias_vencido'],
            row['numero'],
            row['Accion_Sugerida'],
            row['Score_Riesgo']
        ])
    
    # Crear Tabla
    filas = len(df_estrategico) + 1
    if filas > 1:
        tab = Table(displayName="TablaCobranza", ref=f"A1:J{filas}")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tab)
        
        # Formatos
        for r in range(2, filas + 1):
            ws[f'F{r}'].number_format = '"$"#,##0' # Moneda
        
        # Escala de Color (Verde a Rojo en Días Vencido)
        ws.conditional_formatting.add(f"G2:G{filas}", 
            ColorScaleRule(start_type="num", start_value=0, start_color="63BE7B",
                           mid_type="num", mid_value=45, mid_color="FFEB84",
                           end_type="num", end_value=90, end_color="F8696B"))

    wb.save(output)
    return output.getvalue()

# ======================================================================================
# --- 4. INTERFAZ GRÁFICA (MAIN) ---
# ======================================================================================

def main():
    st.title("🚀 Central de Estrategia de Cobranza")
    st.markdown("Analítica avanzada para la recuperación de cartera.")
    
    if st.button("🔄 Recargar Datos"):
        st.cache_data.clear()
        st.rerun()

    # Cargar datos
    df_raw = cargar_datos_maestros()
    
    if df_raw.empty:
        st.warning("No se encontraron datos. Verifique que los archivos estén cargados o la conexión a Dropbox.")
        st.stop()
        
    # Procesar inteligencia
    df_inteligente = aplicar_inteligencia_cobranza(df_raw)
    
    if df_inteligente.empty:
        st.success("¡Excelente! No hay cartera vencida o datos pendientes por gestionar.")
        st.stop()

    # --- KPIs ---
    col1, col2, col3, col4 = st.columns(4)
    total_riesgo = df_inteligente[df_inteligente['dias_vencido'] > 0]['importe'].sum()
    critico_juridico = df_inteligente[df_inteligente['dias_vencido'] > 90]['importe'].sum()
    clientes_gestion = df_inteligente[df_inteligente['dias_vencido'] > 0].shape[0]
    
    # Obtener el top cliente de forma segura
    top_deudor = "N/A"
    if not df_inteligente.empty:
        top_deudor = df_inteligente.iloc[0]['nombrecliente']

    col1.metric("🔥 Total en Riesgo", f"${total_riesgo:,.0f}")
    col2.metric("⚖️ Crítico (>90 días)", f"${critico_juridico:,.0f}")
    col3.metric("👥 Clientes a Gestionar", f"{clientes_gestion}")
    col4.metric("🚨 Prioridad #1", f"{str(top_deudor)[:15]}...")

    st.markdown("---")

    # --- GRÁFICOS ---
    col_matrix, col_vendor = st.columns([2, 1])
    
    with col_matrix:
        st.subheader("🎯 Matriz de Priorización")
        st.info("Arriba a la derecha: **Mayor Deuda + Más Antigüedad** (Atacar Primero)")
        
        df_chart = df_inteligente[df_inteligente['dias_vencido'] > 0]
        if not df_chart.empty:
            fig = px.scatter(
                df_chart,
                x="dias_vencido",
                y="importe",
                size="importe",
                color="Accion_Sugerida",
                hover_name="nombrecliente",
                title="Mapa de Calor de Riesgo",
                labels={"dias_vencido": "Días de Mora", "importe": "Valor Deuda"},
                color_discrete_map={
                    "🔴 JURÍDICO / PRE-JURÍDICO": "#D32F2F",
                    "🟠 CONCILIACIÓN URGENTE": "#F57C00",
                    "🟡 GESTIÓN TELEFÓNICA (Prioridad)": "#FBC02D",
                    "🟡 GESTIÓN ADMINISTRATIVA": "#FFEB3B",
                    "🟢 RECORDATORIO AMABLE": "#388E3C"
                }
            )
            fig.add_vline(x=90, line_dash="dash", line_color="red")
            fig.update_layout(height=400)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Sin datos para graficar.")

    with col_vendor:
        st.subheader("🕵️‍♂️ Foco por Vendedor")
        df_vend = df_inteligente[df_inteligente['dias_vencido'] > 30].groupby('nomvendedor')['importe'].sum().reset_index().sort_values('importe', ascending=False)
        if not df_vend.empty:
            fig_bar = px.bar(df_vend, x='importe', y='nomvendedor', orientation='h', title="Cartera >30 Días", text_auto='.2s')
            fig_bar.update_layout(height=400)
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("No hay cartera mayor a 30 días.")

    # --- DESCARGA EXCEL ---
    st.markdown("### 📥 Herramientas")
    excel_data = generar_excel_estrategico(df_inteligente)
    st.download_button(
        label="📥 DESCARGAR MATRIZ DE GESTIÓN (.xlsx)",
        data=excel_data,
        file_name=f"Estrategia_Cobranza_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )

    # --- LISTADO TÁCTICO ---
    st.markdown("---")
    st.subheader("📝 Gestión Táctica")
    
    filtro = st.multiselect("Filtrar Acción:", df_inteligente['Accion_Sugerida'].unique(), default=df_inteligente['Accion_Sugerida'].unique())
    df_show = df_inteligente[df_inteligente['Accion_Sugerida'].isin(filtro)].copy()
    
    st.dataframe(
        df_show[['Score_Riesgo', 'nombrecliente', 'nit', 'Accion_Sugerida', 'dias_vencido', 'importe', 'nomvendedor', 'telefono1']],
        column_config={
            "Score_Riesgo": st.column_config.ProgressColumn("Riesgo", min_value=0, max_value=100, format="%d"),
            "importe": st.column_config.NumberColumn("Deuda Total", format="$%d"),
            "dias_vencido": st.column_config.NumberColumn("Días Mora", format="%d días"),
            "telefono1": st.column_config.TextColumn("Teléfono")
        },
        use_container_width=True,
        hide_index=True,
        height=500
    )

    # --- ACCIÓN RÁPIDA (WHATSAPP) ---
    if not df_show.empty:
        st.markdown("### ⚡ Gestión Rápida (Top 1 de la lista filtrada)")
        cliente = df_show.iloc[0]
        
        c1, c2 = st.columns(2)
        with c1:
            tel_raw = str(cliente['telefono1']).replace('.0', '')
            tel_clean = re.sub(r'\D', '', tel_raw) # Solo números
            
            msg = f"Hola {cliente['nombrecliente']}, le escribimos de Ferreinox. Notamos un saldo pendiente de ${cliente['importe']:,.0f}. Agradecemos gestionar su pago."
            
            if len(tel_clean) >= 10:
                link_wa = f"https://wa.me/57{tel_clean}?text={quote(msg)}"
                st.markdown(f"""
                <a href="{link_wa}" target="_blank" class="button-wa">
                📱 Enviar WhatsApp a {cliente['nombrecliente']}
                </a>
                """, unsafe_allow_html=True)
            else:
                st.warning(f"Teléfono no válido para WhatsApp: {tel_raw}")
        
        with c2:
            st.info(f"📧 **Correo:** {cliente['e_mail']}")

if __name__ == "__main__":
    main()
