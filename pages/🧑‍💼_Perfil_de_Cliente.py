# ======================================================================================
# SISTEMA INTEGRAL DE GESTIÓN DE CARTERA Y COBRANZA (V. FINAL - SIN ERRORES)
# ======================================================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import dropbox
import io
import os
import glob
import re
import unicodedata
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from urllib.parse import quote
import yagmail

# --- 1. CONFIGURACIÓN DE LA PÁGINA Y ESTILOS ---
st.set_page_config(
    page_title="Centro de Mando: Cobranza Estratégica",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos CSS
st.markdown("""
<style>
    .stApp { background-color: #f8f9fa; }
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border-left: 5px solid #0d6efd;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }
    .action-btn-wa {
        background-color: #25D366; color: white !important;
        padding: 8px 16px; border-radius: 50px; text-decoration: none;
        font-weight: 600; display: inline-block; border: 1px solid #1da851;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .action-btn-email {
        background-color: #EA4335; color: white !important;
        padding: 8px 16px; border-radius: 50px; text-decoration: none;
        font-weight: 600; display: inline-block; border: 1px solid #c53929;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
</style>
""", unsafe_allow_html=True)

# ======================================================================================
# --- 2. FUNCIONES DE LIMPIEZA Y CARGA ---
# ======================================================================================

def normalizar_texto(texto):
    """Elimina acentos y caracteres especiales."""
    if not isinstance(texto, str): return str(texto)
    texto = unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode("utf-8")
    return texto.upper().strip()

def limpiar_moneda(valor):
    """Convierte texto de moneda a número float de forma robusta."""
    if pd.isna(valor): return 0.0
    s_val = str(valor).strip().replace('$', '').replace(' ', '')
    try:
        # Eliminar todo excepto dígitos, puntos y comas
        s_val = re.sub(r'[^\d.,-]', '', s_val)
        if not s_val: return 0.0
        
        # Lógica para determinar separador de miles vs decimales
        if ',' in s_val and '.' in s_val:
            if s_val.rfind(',') > s_val.rfind('.'): # Coma es decimal (Europeo/Col)
                s_val = s_val.replace('.', '').replace(',', '.')
            else: # Punto es decimal (USA)
                s_val = s_val.replace(',', '')
        elif ',' in s_val:
            # Si solo hay comas, asumimos que es decimal si hay 1 o 2 digitos al final
            parts = s_val.split(',')
            if len(parts[-1]) != 2: 
                s_val = s_val.replace(',', '') # Miles
            else:
                s_val = s_val.replace(',', '.') # Decimal
        return float(s_val)
    except:
        return 0.0

@st.cache_data(ttl=300)
def cargar_datos_maestros():
    """Carga datos desde Dropbox o Archivos Locales."""
    df_raw = pd.DataFrame()
    origen = "Desconocido"

    # Intentar Dropbox
    try:
        if "dropbox" in st.secrets:
            dbx = dropbox.Dropbox(
                app_key=st.secrets["dropbox"]["app_key"],
                app_secret=st.secrets["dropbox"]["app_secret"],
                oauth2_refresh_token=st.secrets["dropbox"]["refresh_token"]
            )
            _, res = dbx.files_download(path='/data/cartera_detalle.csv')
            df_raw = pd.read_csv(io.StringIO(res.content.decode('latin-1')), sep='|', header=None, dtype=str)
            origen = "Nube (Dropbox)"
            
            # Asignar nombres estándar si viene sin header
            nombres_cols = [
                'Serie','Numero','Fecha Documento','Fecha Vencimiento','Cod Cliente',
                'NombreCliente','Nit','Poblacion','Provincia','Telefono1','Telefono2',
                'NomVendedor','Entidad Autoriza','E-Mail','Importe','Descuento',
                'Cupo Aprobado','Dias Vencido'
            ]
            if len(df_raw.columns) == len(nombres_cols):
                df_raw.columns = nombres_cols
    except Exception:
        pass 

    # Intentar Local si Dropbox falla
    if df_raw.empty:
        archivos = glob.glob("Cartera_*.xlsx")
        if archivos:
            archivo_reciente = max(archivos, key=os.path.getctime)
            try:
                df_raw = pd.read_excel(archivo_reciente, dtype=str)
                origen = f"Local ({archivo_reciente})"
            except: pass

    if df_raw.empty:
        return pd.DataFrame(), "Sin Datos"

    # Limpieza de Columnas
    cols_map = {
        'NombreCliente': 'cliente', 'Nit': 'nit', 'NomVendedor': 'vendedor',
        'Importe': 'saldo', 'Dias Vencido': 'dias_mora', 'E-Mail': 'email',
        'Telefono1': 'telefono', 'Numero': 'factura', 'Fecha Vencimiento': 'fecha_venc'
    }
    
    # Renombrar columnas buscando coincidencia parcial
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    mapping_final = {}
    for col_real in df_raw.columns:
        col_norm = normalizar_texto(col_real)
        for key, val in cols_map.items():
            if normalizar_texto(key) in col_norm:
                mapping_final[col_real] = val
                break
    
    df_raw.rename(columns=mapping_final, inplace=True)
    
    # Rellenar columnas faltantes
    required = ['cliente', 'saldo', 'dias_mora', 'vendedor', 'nit', 'factura', 'email', 'telefono']
    for req in required:
        if req not in df_raw.columns:
            df_raw[req] = 0 if req in ['saldo', 'dias_mora'] else 'N/A'

    # Convertir Tipos
    df_raw['saldo'] = df_raw['saldo'].apply(limpiar_moneda)
    df_raw['dias_mora'] = pd.to_numeric(df_raw['dias_mora'], errors='coerce').fillna(0)
    
    # Eliminar filas basura
    df_raw = df_raw[df_raw['saldo'] != 0]

    return df_raw, origen

def analizar_cartera(df):
    """Calcula rangos, estados y acciones sugeridas."""
    if df.empty: return df

    # Rangos de Edad
    bins = [-9999, 0, 30, 60, 90, 9999]
    labels = ['Corriente (Al día)', '1 a 30 Días', '31 a 60 Días', '61 a 90 Días', '> 90 Días (Jurídico)']
    df['rango_mora'] = pd.cut(df['dias_mora'], bins=bins, labels=labels)

    # Acción Sugerida
    def get_accion(row):
        dias = row['dias_mora']
        if dias <= 0: return "Agradecer pago"
        if dias <= 15: return "Recordatorio Whatsapp"
        if dias <= 30: return "Llamada de servicio"
        if dias <= 60: return "LLAMADA DE COBRO"
        if dias <= 90: return "BLOQUEO + Prejurídico"
        return "TRASLADO ABOGADO"

    df['accion_sugerida'] = df.apply(get_accion, axis=1)
    return df

# ======================================================================================
# --- 3. GENERACIÓN DE EXCEL ---
# ======================================================================================

def generar_excel_gerencial(df_detalle, df_resumen_cliente):
    output = io.BytesIO()
    wb = Workbook()
    
    # Hoja 1: Resumen
    ws_kpi = wb.active
    ws_kpi.title = "Resumen Gerencial"
    ws_kpi['A1'] = "INFORME DE ESTADO DE CARTERA"
    ws_kpi['A1'].font = Font(size=14, bold=True)
    
    total = df_detalle['saldo'].sum()
    vencido = df_detalle[df_detalle['dias_mora'] > 0]['saldo'].sum()
    
    ws_kpi['A3'] = "Total Cartera"; ws_kpi['B3'] = total
    ws_kpi['A4'] = "Total Vencido"; ws_kpi['B4'] = vencido
    ws_kpi['B3'].number_format = '$ #,##0'
    ws_kpi['B4'].number_format = '$ #,##0'
    
    # Hoja 2: Detalle Clientes
    ws_cli = wb.create_sheet("Top Clientes")
    headers = ["Cliente", "NIT", "Vendedor", "Saldo Total", "Días Mora Max", "Acción Sugerida", "Teléfono"]
    ws_cli.append(headers)
    
    # Estilo Header
    fill = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
    font = Font(color="FFFFFF", bold=True)
    for c in range(1, len(headers)+1):
        cell = ws_cli.cell(row=1, column=c)
        cell.fill = fill
        cell.font = font
        
    for _, row in df_resumen_cliente.sort_values('saldo', ascending=False).iterrows():
        ws_cli.append([
            row['cliente'], row['nit'], row['vendedor'], row['saldo'], 
            row['dias_mora'], row['accion_sugerida'], row['telefono']
        ])
        
    # Formato Tabla
    tab = Table(displayName="TablaClientes", ref=f"A1:G{len(df_resumen_cliente)+1}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws_cli.add_table(tab)
    
    # Hoja 3: Data Cruda
    ws_data = wb.create_sheet("Data Cruda")
    rows = [df_detalle.columns.tolist()] + df_detalle.values.tolist()
    for r in rows: ws_data.append(r)

    wb.save(output)
    return output.getvalue()

# ======================================================================================
# --- 4. APP PRINCIPAL ---
# ======================================================================================

def main():
    with st.sidebar:
        st.title("Panel de Control")
        df, source_msg = cargar_datos_maestros()
        st.caption(f"Fuente: {source_msg}")
        
        if df.empty:
            st.error("No hay datos. Cargue archivos o conecte Dropbox.")
            st.stop()
            
        df_processed = analizar_cartera(df)
        
        # Filtros
        vendedores = ["TODOS"] + sorted(list(df_processed['vendedor'].unique()))
        sel_vendedor = st.selectbox("Filtrar Vendedor:", vendedores)
        
        df_view = df_processed.copy()
        if sel_vendedor != "TODOS":
            df_view = df_view[df_view['vendedor'] == sel_vendedor]
            
    # KPIs Principales
    st.markdown("## 📊 Estado de Cartera")
    total = df_view['saldo'].sum()
    vencido = df_view[df_view['dias_mora'] > 0]['saldo'].sum()
    critico = df_view[df_view['dias_mora'] > 60]['saldo'].sum()
    
    k1, k2, k3 = st.columns(3)
    k1.metric("💰 Cartera Total", f"${total:,.0f}")
    k2.metric("🔥 Vencido", f"${vencido:,.0f}")
    k3.metric("🚨 Crítico (>60d)", f"${critico:,.0f}")
    
    st.markdown("---")
    
    tab1, tab2, tab3 = st.tabs(["⚔️ GESTIÓN (Sala de Guerra)", "📈 ANÁLISIS", "📥 REPORTES"])
    
    # --- TAB 1: GESTIÓN ---
    with tab1:
        st.info("Lista priorizada por impacto en flujo de caja y antigüedad.")
        
        # Agrupar por cliente
        df_clients = df_view.groupby(['cliente', 'nit', 'vendedor', 'telefono', 'email']).agg({
            'saldo': 'sum', 'dias_mora': 'max'
        }).reset_index()
        
        df_clients = analizar_cartera(df_clients)
        df_clients = df_clients.sort_values(by=['saldo', 'dias_mora'], ascending=[False, False])
        
        # Selector
        df_clients['label'] = df_clients.apply(lambda x: f"{x['cliente']} | ${x['saldo']:,.0f} | {int(x['dias_mora'])} días", axis=1)
        target = st.selectbox("Seleccione Cliente a Gestionar:", df_clients['label'])
        
        cli_name = target.split(' | ')[0]
        cli_data = df_clients[df_clients['cliente'] == cli_name].iloc[0]
        
        c1, c2 = st.columns([1, 2])
        with c1:
            st.markdown(f"""
            <div style="background:#fff; padding:15px; border:1px solid #ddd; border-radius:10px;">
                <h3>{cli_data['cliente']}</h3>
                <p><b>NIT:</b> {cli_data['nit']}</p>
                <hr>
                <p style="font-size:18px; color:#c0392b;"><b>Deuda: ${cli_data['saldo']:,.0f}</b></p>
                <p style="font-size:16px; color:#d35400;">Mora Max: {int(cli_data['dias_mora'])} días</p>
                <div style="background:#eee; padding:5px; text-align:center;"><b>{cli_data['accion_sugerida']}</b></div>
            </div>
            """, unsafe_allow_html=True)
            
        with c2:
            st.subheader("Generar Acción")
            phone = re.sub(r'\D', '', str(cli_data['telefono']))
            msg = f"Hola {cli_name}, le escribimos de Ferreinox. Su saldo pendiente es ${cli_data['saldo']:,.0f}. Agradecemos su pago."
            
            st.text_area("Mensaje:", value=msg, height=100)
            if len(phone) > 7:
                link = f"https://wa.me/57{phone}?text={quote(msg)}"
                st.markdown(f'<a href="{link}" target="_blank" class="action-btn-wa">📱 Enviar WhatsApp</a>', unsafe_allow_html=True)
            else:
                st.warning("Teléfono no válido para WhatsApp.")

    # --- TAB 2: ANÁLISIS ---
    with tab2:
        

[Image of Data Dashboard]

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            df_pie = df_view.groupby('rango_mora', observed=False)['saldo'].sum().reset_index()
            fig = px.pie(df_pie, values='saldo', names='rango_mora', title='Antigüedad de Deuda')
            st.plotly_chart(fig, use_container_width=True)
        with col_g2:
            df_top = df_clients.head(10)
            fig2 = px.bar(df_top, x='saldo', y='cliente', orientation='h', title="Top 10 Deudores", color='dias_mora')
            fig2.update_layout(yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig2, use_container_width=True)

    # --- TAB 3: REPORTES ---
    with tab3:
        st.success("Reporte listo para descargar con formato gerencial.")
        excel_data = generar_excel_gerencial(df_view, df_clients)
        st.download_button(
            "📥 Descargar Excel Gerencial",
            data=excel_data,
            file_name=f"Reporte_Cartera_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
        st.dataframe(df_view)

if __name__ == '__main__':
    main()
