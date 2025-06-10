import streamlit as st
import pandas as pd
import numpy as np
import pywhatkit
from janitor import clean_names
import toml

# --- Autenticación por contraseña usando carterasecrets.toml ---
try:
    secrets = toml.load("carterasecrets.toml")
    general_password = secrets["general"]["password"]
    vendedores_secrets = secrets["vendedores"]
except Exception:
    st.error("No se encontraron las contraseñas en carterasecrets.toml. Contacta al administrador.")
    st.stop()

password = st.text_input("Introduce la contraseña para acceder a la cartera:", type="password")

# Determinar tipo de acceso y vendedor asociado
acceso_general = False
vendedor_autenticado = None
for vendedor, pass_vendedor in vendedores_secrets.items():
    if password == str(pass_vendedor):
        vendedor_autenticado = vendedor
        break
if password == str(general_password):
    acceso_general = True
elif vendedor_autenticado is None:
    st.warning("Debes ingresar una contraseña válida para acceder al tablero.")
    st.stop()

st.title("📊 Tablero de Cartera Ferreinox SAS BIC")

# --- Cargar y limpiar datos ---
cartera = pd.read_excel("Cartera.xlsx")
cartera = cartera.clean_names()

# --- Filtros de vendedor y cliente ---
vendedores = sorted(cartera['nomvendedor'].dropna().unique())
if acceso_general:
    vendedor_sel = st.selectbox("Selecciona el vendedor:", ["Todos"] + vendedores)
else:
    vendedor_sel = vendedor_autenticado
    st.info(f"Acceso restringido: solo puedes ver tu propia cartera ({vendedor_sel})")

if vendedor_sel == "Todos":
    cartera_filtrada = cartera.copy()
else:
    cartera_filtrada = cartera[cartera['nomvendedor'] == vendedor_sel].copy()

clientes = sorted(cartera_filtrada['nombrecliente'].dropna().unique())
cliente_sel = st.selectbox("Selecciona el cliente:", ["Todos"] + clientes)

if cliente_sel == "Todos":
    cartera_cliente = cartera_filtrada.copy()
else:
    cartera_cliente = cartera_filtrada[cartera_filtrada['nombrecliente'] == cliente_sel].copy()

# --- Indicadores ---
cartera_cliente['importe'] = pd.to_numeric(cartera_cliente['importe'].replace({'[$,]': ''}, regex=True), errors='coerce').fillna(0)
total_cartera = cartera_cliente['importe'].sum()
cartera_vencida = cartera_cliente[cartera_cliente['dias_vencido'] > 0]['importe'].sum()

col1, col2 = st.columns(2)
col1.metric("Cartera Total", f"${total_cartera:,.0f}")
col2.metric("Cartera Vencida", f"${cartera_vencida:,.0f}")

# --- Tabla de cartera filtrada ---
st.dataframe(cartera_cliente[[
    'nombrecliente', 'serie', 'numero', 'fecha_documento', 'fecha_vencimiento', 'importe', 'dias_vencido', 'telefono', 'cod_cliente', 'nomvendedor'
]].sort_values(['nombrecliente', 'dias_vencido'], ascending=[True, False]), use_container_width=True)

# --- Botón de WhatsApp ---
link_pago_base = "https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/"

if cliente_sel != "Todos" and not cartera_cliente.empty:
    cliente_row = cartera_cliente.iloc[0]
    telefono = str(cliente_row['telefono']) if 'telefono' in cliente_row and pd.notna(cliente_row['telefono']) else ''
    nombre_real = cliente_row['nombrecliente']
    vendedor_real = cliente_row['nomvendedor']
    facturas_vencidas = ', '.join([
        f"*{str(int(x)) if isinstance(x, float) and x.is_integer() else str(x)}*"
        for x in cartera_cliente[cartera_cliente['dias_vencido'] > 0]['numero']
    ])
    total_vencido = cartera_cliente[cartera_cliente['dias_vencido'] > 0]['importe'].sum()
    cod_cliente_val = cliente_row['cod_cliente'] if 'cod_cliente' in cliente_row else ''
    if isinstance(cod_cliente_val, float) and pd.notna(cod_cliente_val) and cod_cliente_val.is_integer():
        cod_cliente = str(int(cod_cliente_val))
    else:
        cod_cliente = str(cod_cliente_val)
    mensaje = (
        f"Señores (a) {nombre_real},\n"
        f"Ferreinox SAS BIC te recuerda que las facturas {facturas_vencidas} se vencieron.\n"
        f"Tu total vencido es de ${total_vencido:,.0f}.\n"
        f"Puedes pagarlas a través del siguiente link: {link_pago_base}\n"
        f"Tu código de cliente es: {cod_cliente}\n"
        f"Si tienes alguna duda aquí estoy yo, {vendedor_real}, para ayudarte."
    )
    if st.button("Enviar mensaje de cobro por WhatsApp Web"):
        if telefono and telefono != 'nan' and telefono != '':
            try:
                pywhatkit.sendwhatmsg_instantly(f"+57{telefono[-10:]}", mensaje, wait_time=10, tab_close=True)
                st.success(f"Mensaje de cobro enviado a {telefono}")
            except Exception as e:
                st.error(f"No se pudo enviar el mensaje por WhatsApp: {e}")
        else:
            st.error("No encontramos teléfono para enviar el mensaje. Por favor ingresa al fichero del cliente y anexa un teléfono válido.")
else:
    st.info("Selecciona un cliente para habilitar el envío de mensaje de cobro.")
