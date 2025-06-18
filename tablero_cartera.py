import streamlit as st
import pandas as pd
import numpy as np
import toml
import os # Importar la librería os para verificar la existencia del archivo

# --- Autenticación y Carga de Secretos ---
# Nota para el usuario:
# 1. Asegúrate de que tu archivo se llame EXACTAMENTE 'carterasecrets.toml' (con una sola extensión .toml).
# 2. Este archivo debe estar en EL MISMO DIRECTORIO que tu script de Python.

# Estructura esperada del archivo 'carterasecrets.toml':
#
# [general]
# password = "tu_password_general"
#
# [vendedores]
# "NOMBRE VENDEDOR 1" = "password_vendedor_1"
# "NOMBRE VENDEDOR 2" = "password_vendedor_2"
# # Importante: "NOMBRE VENDEDOR 1" debe ser idéntico al valor en la columna 'nomvendedor' de tu Excel.

st.sidebar.subheader("Estado de Archivos")
if os.path.exists("carterasecrets.toml"):
    st.sidebar.success("carterasecrets.toml encontrado.")
else:
    st.sidebar.error("carterasecrets.toml NO ENCONTRADO. Asegúrate de que esté en el mismo directorio y que el nombre sea correcto.")
if os.path.exists("Cartera.xlsx"):
    st.sidebar.success("Cartera.xlsx encontrado.")
else:
    st.sidebar.error("Cartera.xlsx NO ENCONTRADO. Asegúrate de que esté en el mismo directorio.")


try:
    secrets = toml.load("carterasecrets.toml")
    general_password = secrets.get("general", {}).get("password")
    # CORRECCIÓN: Acceder a los vendedores desde el diccionario 'secrets' cargado.
    # Usamos .get('vendedores', {}) para evitar un error si la sección [vendedores] no existe.
    vendedores_secrets = secrets.get("vendedores", {})

except FileNotFoundError:
    st.error("Archivo 'carterasecrets.toml' no encontrado. Por favor, revisa las instrucciones en la barra lateral.")
    st.stop()
except Exception as e:
    st.error(f"Error al cargar 'carterasecrets.toml': {e}. Contacta al administrador.")
    st.stop()

# --- Interfaz de Autenticación ---
password = st.text_input("Introduce la contraseña para acceder a la cartera:", type="password")

if not password:
    st.warning("Debes ingresar una contraseña para continuar.")
    st.stop()

# Determinar tipo de acceso y vendedor asociado
acceso_general = False
vendedor_autenticado = None

if password == str(general_password):
    acceso_general = True
else:
    # CORRECCIÓN: Iterar sobre 'vendedores_secrets' que ahora sí está definido.
    for vendedor_key, pass_vendedor in vendedores_secrets.items():
        if password == str(pass_vendedor):
            vendedor_autenticado = vendedor_key # Guarda la clave exacta del TOML
            break

if not acceso_general and vendedor_autenticado is None:
    st.warning("Contraseña incorrecta. No tienes acceso al tablero.")
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

# --- Depuración - Carga y limpieza de datos (en la barra lateral) ---
st.sidebar.subheader("Depuración de Datos de Cartera")
st.sidebar.write(f"Total de filas cargadas en Cartera.xlsx: {len(cartera)}")
st.sidebar.write(f"Columnas después de clean_names(): {cartera.columns.tolist()}")

if 'nomvendedor' in cartera.columns:
    st.sidebar.write("Columna 'nomvendedor' encontrada.")
    unique_vendedores_excel = cartera['nomvendedor'].dropna().unique().tolist()
