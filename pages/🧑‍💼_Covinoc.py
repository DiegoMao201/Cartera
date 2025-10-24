# ======================================================================================
# ARCHIVO: Pagina_Covinoc.py (v2 - Cruce por NIT/Documento y Numero/Titulo_Valor)
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
    """Normaliza nombres de columnas y datos para comparación."""
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

            # Columnas originales de 'cartera_detalle.csv'
            nombres_columnas_originales = [
                'Serie', 'Numero', 'Fecha Documento', 'Fecha Vencimiento', 'Cod Cliente',
                'NombreCliente', 'Nit', 'Poblacion', 'Provincia', 'Telefono1', 'Telefono2',
                'NomVendedor', 'Entidad Autoriza', 'E-Mail', 'Importe', 'Descuento',
                'Cupo Aprobado', 'Dias Vencido'
            ]

            df = pd.read_csv(
                StringIO(contenido_csv), 
                header=None, 
                names=nombres_columnas_originales, 
                sep='|', 
                engine='python',
                # Forzamos las columnas clave de cruce a ser texto
                dtype={'Nit': str, 'Numero': str}
            )
            
            # Renombrar columnas a un formato estándar (minúsculas, guion bajo)
            df_renamed = df.rename(columns=lambda x: normalizar_nombre(x).lower().replace(' ', '_'))
            df_renamed = df_renamed.loc[:, ~df_renamed.columns.duplicated()]
            
            # Procesamiento básico de fechas
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
            path_archivo_dropbox = '/data/reporteTransacciones.xlsx'
            
            metadata, res = dbx.files_download(path=path_archivo_dropbox)
            
            # Leemos el Excel, forzando las columnas clave de cruce a ser texto (string)
            df = pd.read_excel(
                BytesIO(res.content),
                dtype={
                    'DOCUMENTO': str,
                    'TITULO_VALOR': str
                }
            )
            
            # Normalizamos los nombres de las columnas leídas del Excel
            # Ej: 'TITULO_VALOR' -> 'titulo_valor', 'AVISO_NO_PAGO' -> 'aviso_no_pago'
            df.columns = [normalizar_nombre(c).lower().replace(' ', '_') for c in df.columns]

            return df
    except Exception as e:
        st.error(f"Error al cargar 'reporteTransacciones.xlsx' desde Dropbox: {e}")
        st.warning("Asegúrate de que el archivo 'reporteTransacciones.xlsx' exista en la carpeta '/data/' de Dropbox.")
        return pd.DataFrame()

# --- Función Principal de Procesamiento y Cruce ---

def normalizar_nit(nit_str: str) -> str:
    """Limpia un NIT, quitando dígito de verificación, puntos, y letras."""
    if not isinstance(nit_str, str):
        return ""
    # 1. Quita el dígito de verificación (ej. '901571997-1' -> '901571997')
    nit_base = nit_str.split('-')[0]
    # 2. Quita cualquier cosa que no sea un número (puntos, espacios, 'NIT ')
    nit_limpio = re.sub(r'\D', '', nit_base)
    return nit_limpio.strip()

def normalizar_factura(factura_str: str) -> str:
    """Limpia un número de factura, quitando decimales .0"""
    if not isinstance(factura_str, str):
        return ""
    # Quita '.0' en caso de que venga de un float (ej. '982751997.0' -> '982751997')
    return factura_str.split('.')[0].strip()


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
    
    # Claves de Cartera:
    df_cartera['nit_norm'] = df_cartera['nit'].apply(normalizar_nit)
    df_cartera['factura_norm'] = df_cartera['numero'].apply(normalizar_factura)
    df_cartera['clave_unica'] = df_cartera['nit_norm'] + '_' + df_cartera['factura_norm']

    # Claves de Covinoc:
    df_covinoc['nit_norm'] = df_covinoc['documento'].apply(normalizar_nit)
    df_covinoc['factura_norm'] = df_covinoc['titulo_valor'].apply(normalizar_factura)
    df_covinoc['clave_unica'] = df_covinoc['nit_norm'] + '_' + df_covinoc['factura_norm']

    # 4. Lógica de Cruces
    
    # ---- Lógica 1: Facturas a Subir ----
    # "solo las facturas de los clientes que estan en el archivo reporte"
    # 1. Obtener la lista de NITs normalizados que SÍ están en Covinoc
    nits_en_covinoc = df_covinoc['nit_norm'].unique()
    
    # 2. Filtrar la cartera de Ferreinox a solo esos clientes
    df_cartera_clientes_protegidos = df_cartera[df_cartera['nit_norm'].isin(nits_en_covinoc)].copy()
    
    # 3. De esos clientes, encontrar las facturas (clave_unica) que NO están en Covinoc
    df_merge_subir = pd.merge(
        df_cartera_clientes_protegidos,
        df_covinoc[['clave_unica']], # Solo necesitamos la clave de covinoc para comparar
        on='clave_unica',
        how='left',
        indicator=True
    )
    # Filtramos por '_merge' == 'left_only' (solo están en Cartera)
    df_a_subir = df_merge_subir[df_merge_subir['_merge'] == 'left_only'].copy()
    
    # Limpiamos columnas innecesarias
    df_a_subir = df_a_subir.drop(columns=['_merge'])

    # ---- Lógica 2: Exoneraciones (Están en Covinoc, NO en Cartera) ----
    df_merge_exonerar = pd.merge(
        df_covinoc,
        df_cartera[['clave_unica']], # Solo necesitamos la clave de cartera para comparar
        on='clave_unica',
        how='left',
        indicator=True
    )
    # Filtramos por '_merge' == 'left_only' (solo están en Covinoc)
    df_a_exonerar = df_merge_exonerar[df_merge_exonerar['_merge'] == 'left_only'].copy()
    df_a_exonerar = df_a_exonerar.drop(columns=['_merge'])

    # ---- Lógica 3: Avisos de No Pago (Intersección + Vencidas < 58 días) ----
    # Hacemos el merge 'inner' para encontrar las que están en AMBOS archivos
    # Empezamos con Cartera (left) para mantener 'dias_vencido'
    df_interseccion = pd.merge(
        df_cartera,
        df_covinoc, 
        on='clave_unica',
        how='inner', # 'inner' = solo las que están en ambos
        suffixes=('_cartera', '_covinoc') # Renombra columnas duplicadas ej. 'cliente_cartera', 'cliente_covinoc'
    )
    
    # De la intersección, filtramos por las condiciones de días (1 a 57 días)
    df_aviso_no_pago = df_interseccion[
        (df_interseccion['dias_vencido'] > 0) & 
        (df_interseccion['dias_vencido'] < 58)
    ].copy()

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
            except Exception:
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
            if 'df_cartera' not in locals() or 'df_covinoc' not in locals():
                st.error("No se pudieron cargar los archivos base. Verifica la conexión o los nombres de archivo en Dropbox.")
                st.stop()
            else:
                 st.warning("Se cargaron los archivos, pero no se encontraron diferencias para las 3 categorías.")


        # --- Contenedor Principal con Pestañas ---
        st.markdown("---")
        
        tab1, tab2, tab3 = st.tabs([
            f"1. Facturas a Subir ({len(df_a_subir)})", 
            f"2. Exoneraciones ({len(df_a_exonerar)})", 
            f"3. Avisos de No Pago ({len(df_aviso_no_pago)})"
        ])

        with tab1:
            st.subheader("Facturas a Subir a Covinoc")
            st.markdown("Facturas de **clientes protegidos** que están en **Cartera Ferreinox** pero **NO** en Covinoc.")
            
            columnas_mostrar_subir = ['nombrecliente', 'nit', 'numero', 'fecha_vencimiento', 'dias_vencido', 'importe', 'nomvendedor', 'clave_unica']
            columnas_existentes_subir = [col for col in columnas_mostrar_subir if col in df_a_subir.columns]
            
            st.dataframe(df_a_subir[columnas_existentes_subir], use_container_width=True, hide_index=True)
            
            # Placeholder para el botón de descarga
            st.download_button(
                label="📥 Descargar Excel para Subida (Próximamente)",
                data="", 
                file_name="subir_covinoc.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                disabled=True 
            )

        with tab2:
            st.subheader("Facturas a Exonerar de Covinoc")
            st.markdown("Facturas que están en **Covinoc** pero **NO** en la Cartera Ferreinox (ej. ya fueron pagadas).")
            
            # Mostramos las columnas originales de Covinoc
            columnas_mostrar_exonerar = ['cliente', 'documento', 'titulo_valor', 'saldo', 'fecha', 'vencimiento', 'estado', 'clave_unica']
            columnas_existentes_exonerar = [col for col in columnas_mostrar_exonerar if col in df_a_exonerar.columns]
            
            st.dataframe(df_a_exonerar[columnas_existentes_exonerar], use_container_width=True, hide_index=True)
            
            # Placeholder para el botón de descarga
            st.download_button(
                label="📥 Descargar Excel para Exoneración (Próximamente)",
                data="", 
                file_name="exonerar_covinoc.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                disabled=True 
            )

        with tab3:
            st.subheader("Facturas para Aviso de No Pago")
            st.markdown("Facturas que están **en ambos reportes** Y tienen un vencimiento **entre 1 y 57 días**.")
            
            # Mostramos una combinación de columnas de Cartera (la fuente principal) y Covinoc
            columnas_mostrar_aviso = [
                'nombrecliente', 'nit', 'numero', 'fecha_vencimiento', 'dias_vencido', 
                'importe', 'nomvendedor', 'saldo_covinoc', 'estado_covinoc', 'clave_unica'
            ]
            # Renombramos las columnas post-merge para que sean legibles
            df_aviso_no_pago_display = df_aviso_no_pago.rename(columns={
                'nombrecliente_cartera': 'nombrecliente',
                'nit_cartera': 'nit',
                'numero_cartera': 'numero',
                'importe_cartera': 'importe',
                'nomvendedor_cartera': 'nomvendedor',
                'saldo_covinoc': 'saldo_covinoc',
                'estado_covinoc': 'estado_covinoc'
            })
            
            columnas_existentes_aviso = [col for col in columnas_mostrar_aviso if col in df_aviso_no_pago_display.columns]
            
            st.dataframe(df_aviso_no_pago_display[columnas_existentes_aviso], use_container_width=True, hide_index=True)
            
            # Placeholder para el botón de descarga
            st.download_button(
                label="📥 Descargar Excel para Aviso de No Pago (Próximamente)",
                data="", 
                file_name="aviso_no_pago_covinoc.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                disabled=True
            )

if __name__ == '__main__':
    main()
