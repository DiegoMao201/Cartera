import streamlit as st
import pandas as pd
import numpy as np
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
try:
    import janitor
    cartera = pd.read_excel("Cartera.xlsx")
    cartera = cartera.clean_names()
except ImportError:
    st.error("Falta la librería pyjanitor. Instálala con: pip install pyjanitor")
    st.stop()
except FileNotFoundError:
    st.error("No se encontró el archivo 'Cartera.xlsx'. Asegúrate de que esté en la misma carpeta que tu script.")
    st.stop()
except Exception as e:
    st.error(f"Error al cargar o limpiar los datos de Cartera.xlsx: {e}")
    st.stop()


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

# Mensaje si no hay datos para el vendedor seleccionado
if cartera_filtrada.empty:
    st.warning(f"No se encontraron datos de cartera para el vendedor: **{vendedor_sel}**.")
    st.stop() # Detiene la ejecución del resto del script si no hay datos

clientes = sorted(cartera_filtrada['nombrecliente'].dropna().unique())
cliente_sel = st.selectbox("Selecciona el cliente:", ["Todos"] + clientes)

if cliente_sel == "Todos":
    cartera_cliente = cartera_filtrada.copy()
else:
    cartera_cliente = cartera_filtrada[cartera_filtrada['nombrecliente'] == cliente_sel].copy()

# --- Indicadores ---
# Convertir 'importe' a numérico, manejando posibles formatos de moneda
cartera_cliente['importe'] = pd.to_numeric(cartera_cliente['importe'].astype(str).str.replace(r'[$,.]', '', regex=True), errors='coerce').fillna(0)


total_cartera = cartera_cliente['importe'].sum()
cartera_vencida = cartera_cliente[cartera_cliente['dias_vencido'] > 0]['importe'].sum()

col1, col2 = st.columns(2)
col1.metric("Cartera Total", f"${total_cartera:,.0f}")
col2.metric("Cartera Vencida", f"${cartera_vencida:,.0f}")

# --- Tabla de cartera filtrada ---
columnas_deseadas = [
    'nombrecliente', 'serie', 'numero', 'fecha_documento', 'fecha_vencimiento', 'importe', 'dias_vencido', 'telefono', 'cod_cliente', 'nomvendedor'
]
columnas_existentes = [col for col in columnas_deseadas if col in cartera_cliente.columns]

if not cartera_cliente.empty:
    st.dataframe(
        cartera_cliente[columnas_existentes].sort_values(['nombrecliente', 'dias_vencido'], ascending=[True, False]),
        use_container_width=True
    )
else:
    st.info("No hay datos de cartera para mostrar con los filtros seleccionados.")

# --- Botón de WhatsApp ---
link_pago_base = "https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/"

if cliente_sel != "Todos" and not cartera_cliente.empty:
    cliente_row = cartera_cliente.iloc[0] # Se toma la primera fila del cliente seleccionado
    telefono = str(cliente_row['telefono']) if 'telefono' in cliente_row and pd.notna(cliente_row['telefono']) else ''
    nombre_real = cliente_row['nombrecliente']
    vendedor_real = cliente_row['nomvendedor']
    
    # Asegúrate de que 'numero' es un tipo de dato que se puede convertir a entero si es numérico
    facturas_vencidas = ', '.join([
        f"*{str(int(x)) if isinstance(x, (float, np.number)) and not pd.isna(x) and x.is_integer() else str(x)}*"
        for x in cartera_cliente[cartera_cliente['dias_vencido'] > 0]['numero']
    ])
    
    total_vencido = cartera_cliente[cartera_cliente['dias_vencido'] > 0]['importe'].sum()
    
    cod_cliente_val = cliente_row['cod_cliente'] if 'cod_cliente' in cliente_row else ''
    if isinstance(cod_cliente_val, (float, np.number)) and pd.notna(cod_cliente_val) and cod_cliente_val.is_integer():
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
        try:
            import pywhatkit
            if telefono and telefono != 'nan' and telefono != '':
                try:
                    # pywhatkit.sendwhatmsg_instantly requires a valid phone number format.
                    # Assuming phone numbers in Colombia are 10 digits after +57.
                    # Ensure you handle cases where 'telefono' might not be a clean 10-digit number.
                    clean_telefono = telefono.replace(" ", "").replace("-", "") # Eliminar espacios y guiones
                    if len(clean_telefono) >= 10:
                        pywhatkit.sendwhatmsg_instantly(f"+57{clean_telefono[-10:]}", mensaje, wait_time=10, tab_close=True)
                        st.success(f"Mensaje de cobro enviado a {telefono}")
                    else:
                        st.error("El número de teléfono no parece tener el formato correcto para Colombia (mínimo 10 dígitos).")
                except Exception as e:
                    st.error(f"No se pudo enviar el mensaje por WhatsApp: {e}")
            else:
                st.error("No encontramos teléfono para enviar el mensaje. Por favor ingresa al fichero del cliente y anexa un teléfono válido.")
        except ImportError:
            st.error("La librería pywhatkit no está instalada. Instálala con: pip install pywhatkit")
        except Exception as e:
            st.error(f"La función de WhatsApp solo está disponible en equipos con entorno gráfico y navegador. Intenta desde tu computador personal. Error: {e}")
else:
    st.info("Selecciona un cliente para habilitar el envío de mensaje de cobro.")
