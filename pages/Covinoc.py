# ======================================================================================
# ARCHIVO: Pagina_Covinoc.py (Página secundaria para cruces de cartera protegida)
# ======================================================================================
import streamlit as st
import pandas as pd
import toml
import os
from io import BytesIO, StringIO
import plotly.express as px
import plotly.graph_objects as go
import unicodedata
import re
from datetime import datetime
import dropbox
import glob

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Gestión Covinoc",
    page_icon="🛡️",
    layout="wide"
)

# --- PALETA DE COLORES Y CSS (Copiada de Tablero_Principal.py para consistencia) ---
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
</style>
""", unsafe_allow_html=True)


# ======================================================================================
# --- LÓGICA DE CARGA DE DATOS (REUTILIZADA Y ADAPTADA) ---
# ======================================================================================

# --- Funciones Auxiliares Reutilizadas ---
def normalizar_nombre(nombre: str) -> str:
    """Normaliza nombres para comparación."""
    if not isinstance(nombre, str): return ""
    nombre = nombre.upper().strip().replace('.', '')
    nombre = ''.join(c for c in unicodedata.normalize('NFD', nombre) if unicodedata.category(c) != 'Mn')
    return ' '.join(nombre.split())

ZONAS_SERIE = { "PEREIRA": [155, 189, 158, 439], "MANIZALES": [157, 238], "ARMENIA": [156] }

def procesar_cartera(df: pd.DataFrame) -> pd.DataFrame:
    """Procesa el dataframe de cartera principal (copiado de Tablero_Principal.py)."""
    df_proc = df.copy()
    # Aseguramos que las columnas clave existan antes de procesar
    if 'importe' not in df_proc.columns: df_proc['importe'] = 0
    if 'numero' not in df_proc.columns: df_proc['numero'] = 0
    if 'dias_vencido' not in df_proc.columns: df_proc['dias_vencido'] = 0
    if 'nomvendedor' not in df_proc.columns: df_proc['nomvendedor'] = None
    if 'serie' not in df_proc.columns: df_proc['serie'] = None

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

# --- Funciones de Carga de Dropbox ---

@st.cache_data(ttl=600)
def cargar_datos_cartera_dropbox():
    """Carga los datos de CARTERA más recientes desde el archivo CSV en Dropbox."""
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
            
            # Renombrar columnas a un formato estándar (minúsculas, guion bajo)
            df_renamed = df.rename(columns=lambda x: normalizar_nombre(x).lower().replace(' ', '_'))
            df_renamed = df_renamed.loc[:, ~df_renamed.columns.duplicated()]
            
            # Procesamiento básico de fechas y tipos
            df_renamed['fecha_documento'] = pd.to_datetime(df_renamed['fecha_documento'], errors='coerce')
            df_renamed['fecha_vencimiento'] = pd.to_datetime(df_renamed['fecha_vencimiento'], errors='coerce')
            
            return df_renamed
    except Exception as e:
        st.error(f"Error al cargar 'cartera_detalle.csv' desde Dropbox: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=600)
def cargar_reporte_transacciones_dropbox():
    """Carga el REPORTE TRANSACCIONES (Covinoc) desde un archivo Excel en Dropbox."""
    try:
        APP_KEY = st.secrets["dropbox"]["app_key"]
        APP_SECRET = st.secrets["dropbox"]["app_secret"]
        REFRESH_TOKEN = st.secrets["dropbox"]["refresh_token"]

        with dropbox.Dropbox(app_key=APP_KEY, app_secret=APP_SECRET, oauth2_refresh_token=REFRESH_TOKEN) as dbx:
            # *** ¡IMPORTANTE! Asunción: El archivo se llama 'reporteTransacciones.xlsx' ***
            # *** Si se llama diferente, cambia el nombre aquí ***
            path_archivo_dropbox = '/data/reporteTransacciones.xlsx'
            
            metadata, res = dbx.files_download(path=path_archivo_dropbox)
            
            # Leemos el Excel, forzando las columnas clave a ser texto (string)
            # Esto evita que Excel convierta números grandes en notación científica o elimine ceros
            df = pd.read_excel(
                BytesIO(res.content),
                dtype={
                    'CLIENTE': str,
                    'DOCUMENTO': str
                }
            )
            return df
    except Exception as e:
        st.error(f"Error al cargar 'reporteTransacciones.xlsx' desde Dropbox: {e}")
        st.warning("Asegúrate de que el archivo 'reporteTransacciones.xlsx' exista en la carpeta '/data/' de Dropbox.")
        return pd.DataFrame()

# --- Función Principal de Procesamiento y Cruce ---

@st.cache_data
def cargar_y_comparar_datos():
    """
    Orquesta la carga de ambos archivos, los procesa y realiza los cruces lógicos.
    """
    # 1. Cargar y procesar Cartera Ferreinox
    df_cartera_raw = cargar_datos_cartera_dropbox()
    if df_cartera_raw.empty:
        st.error("No se pudo cargar 'cartera_detalle.csv'. El cruce no puede continuar.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
    df_cartera = procesar_cartera(df_cartera_raw)
    
    # Filtrar series W y X como en el tablero principal
    if 'serie' in df_cartera.columns:
        df_cartera = df_cartera[~df_cartera['serie'].astype(str).str.contains('W|X', case=False, na=False)]

    # 2. Cargar Reporte Transacciones (Covinoc)
    df_covinoc = cargar_reporte_transacciones_dropbox()
    if df_covinoc.empty:
        st.error("No se pudo cargar 'reporteTransacciones.xlsx'. El cruce no puede continuar.")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    # 3. Preparar Claves Únicas para el Cruce
    # Usamos .astype(str).str.strip() para asegurar consistencia
    
    # Clave de Cartera: cod_cliente + numero
    df_cartera['cod_cliente_str'] = df_cartera['cod_cliente'].astype(str).str.strip().str.split('.').str[0]
    df_cartera['numero_str'] = df_cartera['numero'].astype(str).str.strip().str.split('.').str[0]
    df_cartera['clave_unica'] = df_cartera['cod_cliente_str'] + '_' + df_cartera['numero_str']
    
    # Clave de Covinoc: CLIENTE + DOCUMENTO
    df_covinoc['CLIENTE_str'] = df_covinoc['CLIENTE'].astype(str).str.strip().str.split('.').str[0]
    df_covinoc['DOCUMENTO_str'] = df_covinoc['DOCUMENTO'].astype(str).str.strip().str.split('.').str[0]
    df_covinoc['clave_unica'] = df_covinoc['CLIENTE_str'] + '_' + df_covinoc['DOCUMENTO_str']

    # 4. Lógica de Cruces
    
    # ---- Lógica 1: Facturas a Subir (Están en Cartera, NO en Covinoc) ----
    # Usamos un merge 'left' con 'indicator=True' para encontrar las que solo están en la izquierda (Cartera)
    df_merge_subir = pd.merge(
        df_cartera,
        df_covinoc[['clave_unica']], # Solo necesitamos la clave de covinoc para comparar
        on='clave_unica',
        how='left',
        indicator=True
    )
    # Filtramos por '_merge' == 'left_only'
    df_a_subir = df_merge_subir[df_merge_subir['_merge'] == 'left_only'].copy()
    # Limpiamos columnas innecesarias del merge
    df_a_subir = df_a_subir.drop(columns=['_merge', 'cod_cliente_str', 'numero_str'])

    # ---- Lógica 2: Exoneraciones (Están en Covinoc, NO en Cartera) ----
    # Hacemos el merge al revés: Covinoc (left) vs Cartera (right)
    df_merge_exonerar = pd.merge(
        df_covinoc,
        df_cartera[['clave_unica']], # Solo necesitamos la clave de cartera para comparar
        on='clave_unica',
        how='left',
        indicator=True
    )
    # Filtramos por '_merge' == 'left_only' (solo están en Covinoc)
    df_a_exonerar = df_merge_exonerar[df_merge_exonerar['_merge'] == 'left_only'].copy()
    df_a_exonerar = df_a_exonerar.drop(columns=['_merge', 'CLIENTE_str', 'DOCUMENTO_str'])

    # ---- Lógica 3: Avisos de No Pago (Intersección + Vencidas < 58 días) ----
    # Usamos un merge 'inner' para encontrar las que están en AMBOS archivos
    df_interseccion = pd.merge(
        df_cartera,
        df_covinoc[['clave_unica']], # Solo necesitamos la clave para confirmar la intersección
        on='clave_unica',
        how='inner' # 'inner' significa que la 'clave_unica' debe estar en ambos
    )
    
    # De la intersección, filtramos por las condiciones de días
    # "menor a 58 dias" -> asumimos VENCIDAS (dias_vencido > 0) y (dias_vencido < 58)
    # Esto es: 1 a 57 días.
    df_aviso_no_pago = df_interseccion[
        (df_interseccion['dias_vencido'] > 0) & 
        (df_interseccion['dias_vencido'] < 58)
    ].copy()
    df_aviso_no_pago = df_aviso_no_pago.drop(columns=['cod_cliente_str', 'numero_str'])

    return df_a_subir, df_a_exonerar, df_aviso_no_pago


# ======================================================================================
# --- BLOQUE PRINCIPAL DE LA APP ---
# ======================================================================================
def main():
    # --- Lógica de Autenticación (Copiada 1:1 de Tablero_Principal.py) ---
    if 'authentication_status' not in st.session_state:
        st.session_state['authentication_status'] = False
        st.session_state['acceso_general'] = False
        st.session_state['vendedor_autenticado'] = None

    if not st.session_state['authentication_status']:
        st.title("Acceso al Módulo de Cartera Protegida")
        try:
            general_password = st.secrets["general"]["password"]
            vendedores_secrets = st.secrets["vendedores"]
        except Exception as e:
            st.error(f"Error al cargar las contraseñas desde los secretos: {e}")
            st.stop()
        
        password = st.text_input("Introduce la contraseña:", type="password", key="password_input_covinoc")
        
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
        # --- Aplicación Principal (Usuario Autenticado) ---
        st.title("🛡️ Gestión de Cartera Protegida (Covinoc)")

        if st.button("🔄 Recargar Datos (Dropbox)"):
            st.cache_data.clear()
            st.success("Caché limpiado. Recargando datos de Cartera y Covinoc...")
            st.rerun()

        # --- Barra Lateral (Sidebar) ---
        with st.sidebar:
            try:
                st.image("LOGO FERREINOX SAS BIC 2024.png", use_container_width=True)
            except Exception: # Usamos Excepción genérica por si falta el archivo
                st.warning("Logo no encontrado.")
            
            st.success(f"Usuario: {st.session_state['vendedor_autenticado']}")
            
            if st.button("Cerrar Sesión"):
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()
            
            st.markdown("---")
            st.info("Esta página compara la cartera de Ferreinox con el reporte de transacciones de Covinoc.")

        # --- Carga y Procesamiento de Datos ---
        with st.spinner("Cargando y comparando archivos de Dropbox..."):
            df_a_subir, df_a_exonerar, df_aviso_no_pago = cargar_y_comparar_datos()

        if df_a_subir.empty and df_a_exonerar.empty and df_aviso_no_pago.empty:
            st.warning("No se pudieron cargar los datos o no se encontraron diferencias. Verifica los archivos en Dropbox.")
            st.stop()

        # --- Contenedor Principal con Pestañas ---
        st.markdown("---")
        
        tab1, tab2, tab3 = st.tabs([
            f"1. Facturas a Subir ({len(df_a_subir)})", 
            f"2. Exoneraciones ({len(df_a_exonerar)})", 
            f"3. Avisos de No Pago ({len(df_aviso_no_pago)})"
        ])

        with tab1:
            st.subheader("Facturas a Subir a Covinoc")
            st.markdown("Listado de facturas que están en la **Cartera Ferreinox** pero **NO** en el reporte de transacciones de Covinoc.")
            
            columnas_mostrar_subir = ['nombrecliente', 'cod_cliente', 'numero', 'fecha_vencimiento', 'dias_vencido', 'importe', 'nit', 'nomvendedor']
            columnas_existentes_subir = [col for col in columnas_mostrar_subir if col in df_a_subir.columns]
            
            st.dataframe(df_a_subir[columnas_existentes_subir], use_container_width=True, hide_index=True)
            
            # Placeholder para el botón de descarga
            st.download_button(
                label="📥 Descargar Excel para Subida (Próximamente)",
                data="", # Aquí iría la función que genera el Excel
                file_name="subir_covinoc.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                disabled=True # Deshabilitado por ahora
            )

        with tab2:
            st.subheader("Facturas a Exonerar de Covinoc")
            st.markdown("Listado de facturas que están en el reporte de **Covinoc** pero **NO** en la Cartera Ferreinox (ej. ya fueron pagadas).")
            
            st.dataframe(df_a_exonerar, use_container_width=True, hide_index=True)
            
            # Placeholder para el botón de descarga
            st.download_button(
                label="📥 Descargar Excel para Exoneración (Próximamente)",
                data="", # Aquí iría la función que genera el Excel
                file_name="exonerar_covinoc.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                disabled=True # Deshabilitado por ahora
            )

        with tab3:
            st.subheader("Facturas para Aviso de No Pago")
            st.markdown("Listado de facturas que están **en ambos reportes** (Cartera y Covinoc) Y tienen un vencimiento **entre 1 y 57 días**.")
            
            columnas_mostrar_aviso = ['nombrecliente', 'cod_cliente', 'numero', 'fecha_vencimiento', 'dias_vencido', 'importe', 'nit', 'nomvendedor']
            columnas_existentes_aviso = [col for col in columnas_mostrar_aviso if col in df_aviso_no_pago.columns]
            
            st.dataframe(df_aviso_no_pago[columnas_existentes_aviso], use_container_width=True, hide_index=True)
            
            # Placeholder para el botón de descarga
            st.download_button(
                label="📥 Descargar Excel para Aviso de No Pago (Próximamente)",
                data="", # Aquí iría la función que genera el Excel
                file_name="aviso_no_pago_covinoc.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                disabled=True # Deshabilitado por ahora
            )

if __name__ == '__main__':
    main()
