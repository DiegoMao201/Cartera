import streamlit as st
import pandas as pd
import numpy as np
import toml
import os # Importar la librería os para verificar la existencia del archivo

# --- Depuración - Comprobación de archivos ---
st.sidebar.subheader("Estado de Archivos")
if os.path.exists("carterasecrets.toml"):
    st.sidebar.success("carterasecrets.toml encontrado.")
else:
    st.sidebar.error("carterasecrets.toml NO ENCONTRADO. Asegúrate de que esté en el mismo directorio.")
if os.path.exists("Cartera.xlsx"):
    st.sidebar.success("Cartera.xlsx encontrado.")
else:
    st.sidebar.error("Cartera.xlsx NO ENCONTRADO. Asegúrate de que esté en el mismo directorio.")

# --- Autenticación por contraseña usando carterasecrets.toml ---
try:
    secrets = toml.load("carterasecrets.toml")
    general_password = secrets["general"]["password"]
    vendedores_secrets = secrets["vendedores"]
except FileNotFoundError:
    st.error("Archivo 'carterasecrets.toml' no encontrado. Contacta al administrador.")
    st.stop()
except Exception as e:
    st.error(f"Error al cargar 'carterasecrets.toml': {e}. Contacta al administrador.")
    st.stop()

password = st.text_input("Introduce la contraseña para acceder a la cartera:", type="password")

# Determinar tipo de acceso y vendedor asociado
acceso_general = False
vendedor_autenticado = None
if password == str(general_password):
    acceso_general = True
else:
    for vendedor_key, pass_vendedor in vendedores_secrets.items():
        if password == str(pass_vendedor):
            vendedor_autenticado = vendedor_key # Guarda la clave exacta del TOML
            break

if not acceso_general and vendedor_autenticado is None:
    st.warning("Debes ingresar una contraseña válida para acceder al tablero.")
    st.stop()

st.title("📊 Tablero de Cartera Ferreinox SAS BIC")

# --- Cargar y limpiar datos ---
try:
    import janitor
    cartera = pd.read_excel("Cartera.xlsx")
    cartera = cartera.clean_names()
except ImportError:
    st.error("Falta la librería pyjanitor. Instálala con: pip install pyjanitor")
    st.stop()
except FileNotFoundError:
    st.error("No se encontró el archivo 'Cartera.xlsx'. Asegúrate de que está en el mismo directorio que tu script de Streamlit.")
    st.stop()
except Exception as e:
    st.error(f"Error al cargar o limpiar 'Cartera.xlsx': {e}. Asegúrate de que el archivo no esté corrupto y tenga el formato correcto.")
    st.stop()

# --- Depuración - Carga y limpieza de datos ---
st.sidebar.subheader("Depuración de Datos de Cartera")
st.sidebar.write(f"Total de filas cargadas en Cartera.xlsx: {len(cartera)}")
st.sidebar.write(f"Columnas después de clean_names(): {cartera.columns.tolist()}")

if 'nomvendedor' in cartera.columns:
    st.sidebar.write("Columna 'nomvendedor' encontrada.")
    unique_vendedores_excel = cartera['nomvendedor'].dropna().unique().tolist()
    st.sidebar.write(f"Valores únicos en 'nomvendedor' (de Cartera.xlsx):")
    for v in unique_vendedores_excel:
        st.sidebar.text(f"- '{v}' (longitud: {len(str(v))})")
    st.sidebar.write(f"Tipo de datos de 'nomvendedor': {cartera['nomvendedor'].dtype}")
else:
    st.sidebar.error("¡ERROR! La columna 'nomvendedor' no se encontró en el DataFrame después de clean_names().")
    st.stop()
# --- Fin Depuración - Carga y limpieza de datos ---


# --- Filtros de vendedor y cliente ---
# Los vendedores de la lista de selección provienen de los datos reales del Excel
vendedores_en_excel = sorted(cartera['nomvendedor'].dropna().unique())

if acceso_general:
    vendedor_sel = st.selectbox("Selecciona el vendedor:", ["Todos"] + vendedores_en_excel)
else:
    vendedor_sel = vendedor_autenticado # Este ya es el nombre exacto de la clave del TOML
    # Verificamos si el vendedor autenticado realmente existe en los datos de Excel
    if vendedor_sel not in vendedores_en_excel:
        st.error(f"El vendedor '{vendedor_sel}' (autenticado) no se encontró en los datos de la columna 'nomvendedor' de 'Cartera.xlsx'.")
        st.info("Por favor, verifica que el nombre de usuario en 'carterasecrets.toml' sea EXACTAMENTE igual a algún nombre en la columna 'nomvendedor' de tu 'Cartera.xlsx', incluyendo mayúsculas, m[...]")
        st.stop() # Detenemos la ejecución si el vendedor no se encuentra
    st.info(f"Acceso restringido: solo puedes ver tu propia cartera ({vendedor_sel})")


# --- Depuración - Vendedor Seleccionado/Autenticado ---
st.sidebar.subheader("Depuración de Selección de Vendedor")
st.sidebar.write(f"Vendedor seleccionado/autenticado: '{vendedor_sel}'")
st.sidebar.write(f"Tipo de dato de vendedor_sel: {type(vendedor_sel)}")
# --- Fin Depuración - Vendedor Seleccionado/Autenticado ---

if vendedor_sel == "Todos":
    cartera_filtrada = cartera.copy()
else:
    # Filtramos por el nombre real del vendedor que viene del TOML/Excel
    cartera_filtrada = cartera[cartera['nomvendedor'] == vendedor_sel].copy()

# --- Depuración - Filtrado por Vendedor ---
st.sidebar.subheader("Depuración de Filtrado Final")
st.sidebar.write(f"Número de filas en cartera_filtrada después del filtro de vendedor: {len(cartera_filtrada)}")

# Puedes agregar aquí una comprobación adicional, por ejemplo:
if cartera_filtrada.empty:
    st.warning("No hay datos en la cartera filtrada para el vendedor seleccionado.")
else:
    st.success("Datos cargados y filtrados correctamente.")
