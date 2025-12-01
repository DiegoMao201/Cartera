# ======================================================================================
# ARCHIVO: pages/1_🚀_Estrategia_Cobranza.py
# VERSIÓN: FINAL CORREGIDA (Fix KeyError e_mail y normalización robusta)
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
from datetime import datetime, timedelta
from urllib.parse import quote
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="War Room Cobranza", page_icon="🚀", layout="wide")

# --- ESTILOS CSS PERSONALIZADOS ---
st.markdown("""
<style>
    .stApp { background-color: #F4F6F9; }
    .stMetric { background-color: #FFFFFF; border-radius: 8px; padding: 15px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-left: 5px solid #0058A7; }
    .big-font { font-size: 20px !important; font-weight: bold; color: #1F2937; }
    .action-card { background-color: #FFFFFF; padding: 20px; border-radius: 10px; border: 1px solid #E5E7EB; box-shadow: 0 4px 6px rgba(0,0,0,0.05); margin-bottom: 20px; }
    .whatsapp-btn { 
        background-color: #25D366; color: white !important; padding: 10px 20px; 
        border-radius: 50px; text-decoration: none; font-weight: bold; 
        display: inline-flex; align-items: center; gap: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.2);
    }
    .whatsapp-btn:hover { background-color: #128C7E; color: white !important; }
    .priority-badge { background-color: #FEF2F2; color: #991B1B; padding: 4px 10px; border-radius: 15px; font-size: 12px; font-weight: bold; border: 1px solid #FCA5A5; }
</style>
""", unsafe_allow_html=True)

# ======================================================================================
# --- 1. CARGA DE DATOS (Mismo motor robusto con FIX de columnas) ---
# ======================================================================================

def normalizar_nombre(nombre: str) -> str:
    if not isinstance(nombre, str): return ""
    nombre = nombre.upper().strip().replace('.', '')
    nombre = ''.join(c for c in unicodedata.normalize('NFD', nombre) if unicodedata.category(c) != 'Mn')
    return ' '.join(nombre.split())

@st.cache_data(ttl=600)
def cargar_datos_maestros():
    df_final = pd.DataFrame()
    
    # 1. Dropbox
    try:
        if "dropbox" in st.secrets:
            APP_KEY = st.secrets["dropbox"]["app_key"]
            APP_SECRET = st.secrets["dropbox"]["app_secret"]
            REFRESH_TOKEN = st.secrets["dropbox"]["refresh_token"]
            with dropbox.Dropbox(app_key=APP_KEY, app_secret=APP_SECRET, oauth2_refresh_token=REFRESH_TOKEN) as dbx:
                _, res = dbx.files_download(path='/data/cartera_detalle.csv')
                csv_content = res.content.decode('latin-1')
                # Definimos nombres explícitos para asegurar consistencia
                nombres_cols = ['Serie', 'Numero', 'Fecha Documento', 'Fecha Vencimiento', 'Cod Cliente',
                                'NombreCliente', 'Nit', 'Poblacion', 'Provincia', 'Telefono1', 'Telefono2',
                                'NomVendedor', 'Entidad Autoriza', 'E-Mail', 'Importe', 'Descuento',
                                'Cupo Aprobado', 'Dias Vencido']
                df_dropbox = pd.read_csv(StringIO(csv_content), header=None, names=nombres_cols, sep='|', engine='python')
                df_final = pd.concat([df_final, df_dropbox])
    except Exception: pass 

    # 2. Locales
    archivos = glob.glob("Cartera_*.xlsx")
    for archivo in archivos:
        try:
            df_hist = pd.read_excel(archivo)
            if not df_hist.empty:
                if "Total" in str(df_hist.iloc[-1, 0]): df_hist = df_hist.iloc[:-1]
                df_final = pd.concat([df_final, df_hist])
        except Exception: pass

    if df_final.empty: return pd.DataFrame()

    # --- LIMPIEZA Y NORMALIZACIÓN DE COLUMNAS ---
    # Convertimos todo a minúsculas y reemplazamos espacios por guiones bajos
    df_final = df_final.rename(columns=lambda x: normalizar_nombre(x).lower().replace(' ', '_'))
    
    # FIX: Renombrar variaciones comunes de email para evitar KeyError
    mapa_renombre = {
        'e-mail': 'e_mail',
        'email': 'e_mail',
        'correo': 'e_mail',
        'telefono': 'telefono1',
        'nombre_cliente': 'nombrecliente'
    }
    df_final = df_final.rename(columns=mapa_renombre)
    
    # Asegurar que existan las columnas críticas aunque estén vacías
    columnas_criticas = ['e_mail', 'telefono1', 'nomvendedor', 'nombrecliente', 'nit']
    for col in columnas_criticas:
        if col not in df_final.columns:
            df_final[col] = "No registrado"

    # Eliminar duplicados de columnas si se generaron
    df_final = df_final.loc[:, ~df_final.columns.duplicated()]
    
    # Conversión de tipos
    if 'importe' in df_final.columns: df_final['importe'] = pd.to_numeric(df_final['importe'], errors='coerce').fillna(0)
    else: df_final['importe'] = 0
    
    if 'dias_vencido' in df_final.columns: df_final['dias_vencido'] = pd.to_numeric(df_final['dias_vencido'], errors='coerce').fillna(0)
    
    if 'fecha_vencimiento' in df_final.columns: df_final['fecha_vencimiento'] = pd.to_datetime(df_final['fecha_vencimiento'], errors='coerce')
    
    if 'serie' in df_final.columns: df_final = df_final[~df_final['serie'].astype(str).str.contains('W|X', case=False, na=False)]
    
    return df_final

# ======================================================================================
# --- 2. MOTOR DE ESTRATEGIA (CEREBRO) ---
# ======================================================================================

def procesar_estrategia(df):
    if df.empty: return pd.DataFrame()
    df = df.copy()
    df = df[df['importe'] > 0] # Solo deuda real
    
    # Agrupar por Cliente
    # Usamos las columnas que aseguramos en la carga
    cols_group = ['nombrecliente', 'nit', 'nomvendedor', 'telefono1', 'e_mail']
    # Doble chequeo por si alguna se perdió, aunque el fix de arriba lo previene
    cols_existentes = [c for c in cols_group if c in df.columns]
    
    cliente_kpis = df.groupby(cols_existentes).agg({
        'importe': 'sum',
        'dias_vencido': 'max',
        'numero': 'count',
        'fecha_vencimiento': 'min'
    }).reset_index()

    # Definir Estrategia
    def definir_accion(row):
        dias = row['dias_vencido']
        
        if dias > 120: return "🔴 JURÍDICO INMEDIATO"
        if dias > 90: return "⛔ BLOQUEO Y CONCILIACIÓN"
        if dias > 60: return "🟠 COBRO ADMINISTRATIVO FUERTE"
        if dias > 30: return "🟡 GESTIÓN TELEFÓNICA"
        if dias > 0: return "🟢 RECORDATORIO DE PAGO"
        return "🔵 PREVENTIVO / AL DÍA"

    def definir_prioridad(row):
        # Score de 0 a 100 donde 100 es "Llamar YA"
        p_dias = min(row['dias_vencido'], 120) / 1.2 # Max 100 pts por días
        p_monto = min(row['importe'] / 10000000, 1) * 100 # Max 100 pts por monto (Base 10M)
        return (p_dias * 0.6) + (p_monto * 0.4)

    cliente_kpis['Estrategia'] = cliente_kpis.apply(definir_accion, axis=1)
    cliente_kpis['Prioridad_Score'] = cliente_kpis.apply(definir_prioridad, axis=1)
    
    # Ordenar: Lo más grave y con más plata primero
    return cliente_kpis.sort_values(by='Prioridad_Score', ascending=False)

# ======================================================================================
# --- 3. GENERADOR DE EXCEL SUPER PODEROSO ---
# ======================================================================================

def generar_excel_master(df_estrategia):
    output = BytesIO()
    wb = Workbook()
    
    # --- HOJA 1: MATRIZ DE ACCIÓN (Para el Líder de Cobros) ---
    ws_accion = wb.active
    ws_accion.title = "1. Matriz de Acción"
    
    headers_accion = ["Prioridad", "Cliente", "NIT", "Teléfono", "Deuda Total", "Días Mora (Max)", "Acción Requerida", "Vendedor"]
    ws_accion.append(headers_accion)
    
    # Estilos
    header_style = PatternFill(start_color="1F2937", end_color="1F2937", fill_type="solid")
    font_white = Font(color="FFFFFF", bold=True)
    
    for col, header in enumerate(headers_accion, 1):
        cell = ws_accion.cell(row=1, column=col)
        cell.fill = header_style
        cell.font = font_white
        ws_accion.column_dimensions[get_column_letter(col)].width = 20
    ws_accion.column_dimensions['B'].width = 40

    for idx, row in df_estrategia.iterrows():
        # Uso seguro de .get por si acaso
        ws_accion.append([
            f"{row.get('Prioridad_Score', 0):.1f}",
            row.get('nombrecliente', ''),
            row.get('nit', ''),
            row.get('telefono1', ''),
            row.get('importe', 0),
            row.get('dias_vencido', 0),
            row.get('Estrategia', ''),
            row.get('nomvendedor', '')
        ])
    
    # Formato Tabla y Condicional
    filas = len(df_estrategia) + 1
    if filas > 1:
        tab = Table(displayName="TablaAccion", ref=f"A1:H{filas}")
        tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
        ws_accion.add_table(tab)
    
    # Formato Moneda
    for r in range(2, filas + 1):
        ws_accion[f'E{r}'].number_format = '"$"#,##0'
    
    # --- HOJA 2: VISIÓN GERENCIAL (Pareto y Bloqueos) ---
    ws_gerencia = wb.create_sheet("2. Estrategia Gerencial")
    ws_gerencia.append(["REPORTE DE ALTO NIVEL PARA GERENCIA", "", "", ""])
    ws_gerencia.merge_cells('A1:D1')
    ws_gerencia['A1'].font = Font(size=14, bold=True, color="003366")
    ws_gerencia.append([])
    
    # Top 10 Deudores (Pareto)
    ws_gerencia.append(["TOP 10 CLIENTES CRÍTICOS (Foco Jurídico/Gerencial)"])
    ws_gerencia['A3'].font = Font(bold=True)
    ws_gerencia.append(["Cliente", "Deuda Total", "Días Mora", "Vendedor"])
    
    top_10 = df_estrategia.head(10)
    for _, row in top_10.iterrows():
        ws_gerencia.append([
            row.get('nombrecliente', ''), 
            row.get('importe', 0), 
            row.get('dias_vencido', 0), 
            row.get('nomvendedor', '')
        ])
        
    # Formatos hoja 2
    ws_gerencia.column_dimensions['A'].width = 40
    for r in range(5, 16):
        ws_gerencia[f'B{r}'].number_format = '"$"#,##0'

    wb.save(output)
    return output.getvalue()

# ======================================================================================
# --- 4. INTERFAZ PRINCIPAL ---
# ======================================================================================

def main():
    st.title("🚀 Centro de Comando de Cobranza (War Room)")
    
    # Carga de datos
    df_raw = cargar_datos_maestros()
    if df_raw.empty:
        st.error("⚠️ No hay datos cargados. Verifique Dropbox o archivos locales.")
        st.stop()
        
    df_kpi = procesar_estrategia(df_raw)
    
    # --- PESTAÑAS ESTRATÉGICAS ---
    tab_accion, tab_evolucion, tab_gerencia = st.tabs([
        "⚡ ZONA DE ACCIÓN (Líder)", 
        "📈 EVOLUCIÓN (Analista)", 
        "🧠 ESTRATEGIA (Gerencia)"
    ])

    # ==================================================================================
    # PESTAÑA 1: ZONA DE ACCIÓN (Foco: Ejecución Rápida)
    # ==================================================================================
    with tab_accion:
        st.markdown("### 🎯 Foco del Día: Gestión Táctica")
        
        col_filtro, col_kpi_dia = st.columns([1, 3])
        with col_filtro:
            filtro_accion = st.selectbox(
                "Filtrar por Tipo de Gestión:",
                ["TODOS", "🔴 URGENTE (Jurídico/Conciliación)", "🟡 GESTIÓN (Cobro Telefónico)", "🟢 PREVENTIVO"]
            )
        
        # Filtrado inteligente
        df_show = df_kpi.copy()
        if filtro_accion == "🔴 URGENTE (Jurídico/Conciliación)":
            df_show = df_show[df_show['dias_vencido'] > 90]
        elif filtro_accion == "🟡 GESTIÓN (Cobro Telefónico)":
            df_show = df_show[(df_show['dias_vencido'] > 30) & (df_show['dias_vencido'] <= 90)]
        elif filtro_accion == "🟢 PREVENTIVO":
            df_show = df_show[(df_show['dias_vencido'] > 0) & (df_show['dias_vencido'] <= 30)]

        with col_kpi_dia:
            kpi1, kpi2, kpi3 = st.columns(3)
            kpi1.metric("Clientes en Lista", f"{len(df_show)}")
            kpi2.metric("Monto a Gestionar", f"${df_show['importe'].sum():,.0f}")
            if not df_show.empty:
                kpi3.metric("Ticket Promedio Deuda", f"${df_show['importe'].mean():,.0f}")

        st.markdown("---")
        
        # --- TARJETA DE ACCIÓN RÁPIDA (TOP 1) ---
        if not df_show.empty:
            cliente_top = df_show.iloc[0] # El más crítico
            
            st.markdown(f"""
            <div class="action-card">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <span class="big-font">🔥 PRIORIDAD #1: {cliente_top['nombrecliente']}</span>
                    <span class="priority-badge">{cliente_top['Estrategia']}</span>
                </div>
                <div style="margin-top: 10px; color: #4B5563;">
                    <strong>NIT:</strong> {cliente_top['nit']} | <strong>Vendedor:</strong> {cliente_top['nomvendedor']}
                </div>
                <hr style="margin: 10px 0;">
                <div style="display:flex; justify-content:space-around; text-align:center;">
                    <div>
                        <div style="font-size:12px; color:#6B7280;">DEUDA TOTAL</div>
                        <div style="font-size:24px; font-weight:bold; color:#DC2626;">${cliente_top['importe']:,.0f}</div>
                    </div>
                    <div>
                        <div style="font-size:12px; color:#6B7280;">DÍAS MORA</div>
                        <div style="font-size:24px; font-weight:bold; color:#D97706;">{cliente_top['dias_vencido']}</div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # --- GENERADOR DE WHATSAPP ---
            c_wa_1, c_wa_2 = st.columns([2, 1])
            with c_wa_1:
                st.write("**📝 Redactar Mensaje:**")
                # Limpieza de teléfono segura
                tel_raw = str(cliente_top.get('telefono1', ''))
                tel_limpio = re.sub(r'\D', '', tel_raw)
                
                telefono_editable = st.text_input("Confirmar Teléfono (+57):", value=tel_limpio, key="tel_focus")
                
                # Plantilla dinámica según estado
                if cliente_top['dias_vencido'] > 90:
                    msg_default = f"Hola {cliente_top['nombrecliente']}, le escribe el Área Jurídica de Ferreinox. Requerimos conciliación urgente de su saldo pendiente de ${cliente_top['importe']:,.0f} para evitar acciones adicionales."
                else:
                    msg_default = f"Hola {cliente_top['nombrecliente']}, saludamos de Ferreinox. Recordamos amablemente su saldo pendiente de ${cliente_top['importe']:,.0f}. Agradecemos su gestión."
                
                mensaje_final = st.text_area("Mensaje a enviar:", value=msg_default, height=100)
                
                if telefono_editable:
                    link_wa = f"https://wa.me/57{telefono_editable}?text={quote(mensaje_final)}"
                    st.markdown(f'<a href="{link_wa}" target="_blank" class="whatsapp-btn">📱 Enviar WhatsApp Ahora</a>', unsafe_allow_html=True)
                else:
                    st.warning("⚠️ Ingrese un número válido para habilitar el botón.")
            
            with c_wa_2:
                st.info("💡 **Tip:** Si el cliente solicita soporte o factura, recuerda actualizar el correo en el sistema ERP.")
                # FIX KEYERROR: Uso seguro de .get()
                email_safe = cliente_top.get('e_mail', 'No registrado')
                st.write(f"📧 **Email registrado:** {email_safe}")

        # --- LISTADO TÁCTICO COMPLETO ---
        st.markdown("### 📋 Listado de Gestión (Siguientes en la fila)")
        
        # Selección segura de columnas para mostrar
        cols_display = ['Prioridad_Score', 'nombrecliente', 'telefono1', 'dias_vencido', 'importe', 'Estrategia', 'nomvendedor']
        # Intersección con las columnas que realmente existen
        cols_final_display = [c for c in cols_display if c in df_show.columns]
        
        st.dataframe(
            df_show.iloc[1:][cols_final_display],
            column_config={
                "Prioridad_Score": st.column_config.ProgressColumn("Urgencia", min_value=0, max_value=100),
                "importe": st.column_config.NumberColumn("Deuda", format="$%d"),
                "telefono1": st.column_config.TextColumn("Teléfono (Editable)", width="medium")
            },
            hide_index=True,
            use_container_width=True
        )

    # ==================================================================================
    # PESTAÑA 2: EVOLUCIÓN (Foco: Tendencias y Análisis)
    # ==================================================================================
    with tab_evolucion:
        st.header("📉 Análisis de Evolución de Cartera")
        st.info("Este módulo analiza cómo se distribuye la deuda hoy en comparación con los rangos de tiempo.")

        # Gráfico de Aging (Sustituto de evolución temporal si no hay histórico diario)
        bins = [0, 30, 60, 90, 120, 1000]
        labels = ['0-30 Días (Corriente)', '31-60 Días (Vencido)', '61-90 Días (Crítico)', '91-120 Días (Pre-Jurídico)', '>120 Días (Jurídico)']
        df_kpi['Rango_Mora'] = pd.cut(df_kpi['dias_vencido'], bins=bins, labels=labels)
        
        col_chart1, col_chart2 = st.columns(2)
        
        with col_chart1:
            st.subheader("💰 ¿Dónde está atrapado el dinero?")
            fig_pie = px.pie(df_kpi, values='importe', names='Rango_Mora', title='Distribución de Deuda por Edades', hole=0.4,
                             color_discrete_sequence=px.colors.sequential.RdBu_r)
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with col_chart2:
            st.subheader("⚠️ Concentración de Riesgo")
            # Scatter plot: Días vencido vs Monto
            if not df_kpi.empty:
                fig_scatter = px.scatter(
                    df_kpi[df_kpi['dias_vencido']>0], 
                    x="dias_vencido", y="importe", 
                    size="importe", color="Estrategia",
                    hover_name="nombrecliente",
                    title="Mapa de Calor: Mora vs Valor",
                    color_discrete_map={
                        "🔴 JURÍDICO INMEDIATO": "red",
                        "⛔ BLOQUEO Y CONCILIACIÓN": "darkred",
                        "🟠 COBRO ADMINISTRATIVO FUERTE": "orange",
                        "🟡 GESTIÓN TELEFÓNICA": "gold",
                        "🟢 RECORDATORIO DE PAGO": "green"
                    }
                )
                # Línea de peligro
                fig_scatter.add_vline(x=90, line_dash="dash", line_color="red", annotation_text="Límite Jurídico")
                st.plotly_chart(fig_scatter, use_container_width=True)

    # ==================================================================================
    # PESTAÑA 3: ESTRATEGIA GERENCIAL (Foco: Decisiones de Alto Nivel)
    # ==================================================================================
    with tab_gerencia:
        st.header("🧠 Tablero de Control Gerencial")
        st.markdown("Herramientas para la toma de decisiones: **Bloqueos, Cupos y Foco Comercial**.")
        
        # KPIs Gerenciales
        total_risk = df_kpi[df_kpi['dias_vencido'] > 60]['importe'].sum()
        pct_risk = (total_risk / df_kpi['importe'].sum()) * 100 if df_kpi['importe'].sum() > 0 else 0
        
        g1, g2, g3 = st.columns(3)
        g1.metric("🚨 Cartera en Riesgo Alto (>60 días)", f"${total_risk:,.0f}")
        g2.metric("% Cartera Contaminada", f"{pct_risk:.1f}%")
        g3.metric("Clientes a Bloquear (>90 días)", f"{len(df_kpi[df_kpi['dias_vencido'] > 90])}")

        st.markdown("### ⛔ Listado Sugerido para Bloqueo de Crédito")
        st.markdown("Estos clientes tienen **más de 90 días de mora**. Se sugiere suspender despachos inmediatamente.")
        
        df_block = df_kpi[df_kpi['dias_vencido'] > 90].sort_values('importe', ascending=False)
        
        cols_block = ['nombrecliente', 'nit', 'importe', 'dias_vencido', 'nomvendedor']
        cols_final_block = [c for c in cols_block if c in df_block.columns]
        
        st.dataframe(
            df_block[cols_final_block],
            column_config={
                "importe": st.column_config.NumberColumn("Deuda Riesgosa", format="$%d"),
                "dias_vencido": st.column_config.NumberColumn("Días Mora", format="%d ⚠️"),
            },
            hide_index=True,
            use_container_width=True
        )

        st.markdown("### 📥 Descarga de Inteligencia de Negocios")
        st.write("Descargue el reporte maestro con múltiples hojas: Matriz de Acción, Resumen Gerencial y Listas de Bloqueo.")
        
        excel_data = generar_excel_master(df_kpi)
        st.download_button(
            label="📥 DESCARGAR REPORTE ESTRATÉGICO COMPLETO (.xlsx)",
            data=excel_data,
            file_name=f"Estrategia_Cobranza_Maestra_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

if __name__ == "__main__":
    main()
