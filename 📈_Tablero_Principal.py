import os
import json
import requests
import logging
import threading
import gspread
import tempfile
import glob
import re
import unicodedata
import dropbox
from io import StringIO
import pandas as pd # Necesario para procesar la cartera
from google.oauth2.service_account import Credentials
from datetime import datetime
from flask import Flask, request, make_response
import google.generativeai as genai

# --- CONFIGURACIÓN DE LOGGING Y FLASK ---
app = Flask(__name__)
# Configuración del Logger para ver los logs en el terminal
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
app.logger.setLevel(logging.INFO)

# --- CARGAR VARIABLES DE ENTORNO ---

# WhatsApp (Webhooks)
WHATSAPP_VERIFY_TOKEN = os.environ.get('WHATSAPP_VERIFY_TOKEN')
WHATSAPP_ACCESS_TOKEN = os.environ.get('WHATSAPP_ACCESS_TOKEN')
WHATSAPP_PHONE_NUMBER_ID = os.environ.get('WHATSAPP_PHONE_NUMBER_ID')

# Gemini
GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY')

# Dropbox (Datos de Cartera)
DBX_APP_KEY = os.environ.get('DBX_APP_KEY')
DBX_APP_SECRET = os.environ.get('DBX_APP_SECRET')
DBX_REFRESH_TOKEN = os.environ.get('DBX_REFRESH_TOKEN')
DBX_FILE_PATH = os.environ.get('DBX_FILE_PATH', '/data/cartera_detalle.csv') # Ruta por defecto

# Google Sheets (Log)
GCP_JSON_STR = os.environ.get('GCP_SERVICE_ACCOUNT_JSON')
GOOGLE_SHEET_NAME = os.environ.get('GOOGLE_SHEET_NAME')
GOOGLE_WORKSHEET_NAME = os.environ.get('GOOGLE_WORKSHEET_NAME')

# --- ESTADO EN MEMORIA ---
user_chats = {}
processed_message_ids = set() 
user_security_context = {} # Almacena el último intento de consulta sensible (temporal)
CARTERA_PROCESADA_DF = pd.DataFrame() # Cache para los datos de cartera

# ----------------------------------------------------------------------
## 📊 INICIALIZACIÓN DE GOOGLE SHEETS (LOG DE CHAT)
# ----------------------------------------------------------------------
worksheet = None 
temp_creds_file = None 

def init_google_sheets():
    global worksheet, temp_creds_file
    try:
        if not GCP_JSON_STR or not GOOGLE_SHEET_NAME or not GOOGLE_WORKSHEET_NAME:
            app.logger.warning("Variables de Google Sheets no configuradas. El log de chats está desactivado.")
            return

        with tempfile.NamedTemporaryFile(mode='w', delete=False) as temp_file:
            temp_file.write(GCP_JSON_STR)
            temp_creds_file = temp_file.name
        
        client_gspread = gspread.service_account(filename=temp_creds_file)
        sheet = client_gspread.open(GOOGLE_SHEET_NAME)
        worksheet = sheet.worksheet(GOOGLE_WORKSHEET_NAME)
        
        if not worksheet.get_all_values():
            worksheet.append_row(["Timestamp", "Numero_Usuario", "Mensaje_Usuario", "Respuesta_Bot", "Herramienta_Usada"])
        
        app.logger.info(f"Conectado a Google Sheets para Logging: {GOOGLE_SHEET_NAME} -> {GOOGLE_WORKSHEET_NAME}")

    except Exception as e:
        app.logger.error(f"Error al inicializar Google Sheets: {e}")
        worksheet = None 
    finally:
        if temp_creds_file and os.path.exists(temp_creds_file):
            os.remove(temp_creds_file)

init_google_sheets()

# ----------------------------------------------------------------------
## 🗃️ LÓGICA DE DATOS DE CARTERA (ADAPTACIÓN DE STREAMLIT)
# ----------------------------------------------------------------------

# --- Adaptación de Funciones Auxiliares de Streamlit ---
def normalizar_nombre(nombre: str) -> str:
    if not isinstance(nombre, str): return ""
    nombre = nombre.upper().strip().replace('.', '')
    nombre = ''.join(c for c in unicodedata.normalize('NFD', nombre) if unicodedata.category(c) != 'Mn')
    return ' '.join(nombre.split())

ZONAS_SERIE = { "PEREIRA": [155, 189, 158, 439], "MANIZALES": [157, 238], "ARMENIA": [156] }

def procesar_cartera(df: pd.DataFrame) -> pd.DataFrame:
    df_proc = df.copy()
    # Limpieza y conversión de tipos
    df_proc['importe'] = pd.to_numeric(df_proc['importe'], errors='coerce').fillna(0)
    df_proc['numero'] = pd.to_numeric(df_proc['numero'], errors='coerce').fillna(0)
    df_proc.loc[df_proc['numero'] < 0, 'importe'] *= -1
    df_proc['dias_vencido'] = pd.to_numeric(df_proc['dias_vencido'], errors='coerce').fillna(0)
    df_proc['nomvendedor_norm'] = df_proc['NomVendedor'].apply(normalizar_nombre)
    
    # Asignación de Zonas
    ZONAS_SERIE_STR = {zona: [str(s) for s in series] for zona, series in ZONAS_SERIE.items()}
    def asignar_zona_robusta(valor_serie):
        if pd.isna(valor_serie): return "OTRAS ZONAS"
        numeros_en_celda = re.findall(r'\d+', str(valor_serie))
        if not numeros_en_celda: return "OTRAS ZONAS"
        for zona, series_clave_str in ZONAS_SERIE_STR.items():
            if set(numeros_en_celda) & set(series_clave_str): return zona
        return "OTRAS ZONAS"
    df_proc['zona'] = df_proc['Serie'].apply(asignar_zona_robusta)
    
    # Clasificación de Edad de Cartera
    bins = [-float('inf'), 0, 15, 30, 60, float('inf')]
    labels = ['Al día', '1-15 días', '16-30 días', '31-60 días', 'Más de 60 días']
    df_proc['edad_cartera'] = pd.cut(df_proc['dias_vencido'], bins=bins, labels=labels, right=True)
    
    # Renombrar para consistencia con el código de Streamlit (se usa en las Tools)
    df_proc.rename(columns=lambda x: normalizar_nombre(x).lower().replace(' ', '_'), inplace=True)
    
    return df_proc

def cargar_datos_desde_dropbox():
    """ADAPTACIÓN: Carga datos desde Dropbox usando variables de entorno."""
    if not all([DBX_APP_KEY, DBX_APP_SECRET, DBX_REFRESH_TOKEN]):
        app.logger.error("Credenciales de Dropbox no configuradas.")
        return pd.DataFrame()

    try:
        with dropbox.Dropbox(app_key=DBX_APP_KEY, app_secret=DBX_APP_SECRET, oauth2_refresh_token=DBX_REFRESH_TOKEN) as dbx:
            metadata, res = dbx.files_download(path=DBX_FILE_PATH)
            contenido_csv = res.content.decode('latin-1')

            nombres_columnas_originales = [
                'Serie', 'Numero', 'Fecha Documento', 'Fecha Vencimiento', 'Cod Cliente',
                'NombreCliente', 'Nit', 'Poblacion', 'Provincia', 'Telefono1', 'Telefono2',
                'NomVendedor', 'Entidad Autoriza', 'E-Mail', 'Importe', 'Descuento',
                'Cupo Aprobado', 'Dias Vencido'
            ]

            df = pd.read_csv(StringIO(contenido_csv), header=None, names=nombres_columnas_originales, sep='|', engine='python')
            app.logger.info("Datos de Dropbox cargados exitosamente.")
            return df
    except Exception as e:
        app.logger.error(f"Error al cargar datos desde Dropbox: {e}")
        return pd.DataFrame()

def cargar_datos_historicos():
    """Carga los archivos Excel históricos locales (simplificado)."""
    # En un entorno de cloud deployment (como Heroku/Cloud Run), esta función
    # probablemente no cargaría archivos locales y debería ser reemplazada por 
    # una carga desde S3/Google Cloud Storage. Aquí se mantiene la estructura.
    # Por ahora, retorna un DF vacío para evitar errores si no hay archivos locales.
    return pd.DataFrame()

def cargar_y_procesar_datos():
    """Orquesta la carga de datos, los combina, limpia y procesa, con caching en memoria."""
    global CARTERA_PROCESADA_DF
    
    # Recargar solo si el DF está vacío (o se podría añadir un TTL si es necesario)
    if CARTERA_PROCESADA_DF.empty:
        app.logger.info("Recargando datos de cartera desde cero...")
        df_dropbox = cargar_datos_desde_dropbox()
        df_historico = cargar_datos_historicos() # Esto probablemente será vacío en la nube
        
        df_combinado = pd.concat([df_dropbox, df_historico], ignore_index=True)

        if df_combinado.empty:
            app.logger.error("No se pudieron cargar datos de ninguna fuente. La app no funcionará.")
            return pd.DataFrame()

        df_combinado = df_combinado.loc[:, ~df_combinado.columns.duplicated()]
        
        # Eliminar filas donde el 'Importe' sea NaN o 0 después de la limpieza inicial
        df_combinado.dropna(subset=['Importe'], inplace=True)
        
        CARTERA_PROCESADA_DF = procesar_cartera(df_combinado)
        app.logger.info(f"Procesamiento de cartera finalizado. {len(CARTERA_PROCESADA_DF)} registros cargados.")
    
    return CARTERA_PROCESADA_DF.copy()

# Carga inicial de datos al iniciar el servidor (esto puede tomar tiempo)
cargar_y_procesar_datos() 

# ----------------------------------------------------------------------
## 🛡️ FUNCIONES DE HERRAMIENTA (TOOLS) PARA GEMINI - SEGURAS
# ----------------------------------------------------------------------

# Se definen las funciones de análisis (como en Streamlit) para que Gemini las use internamente
def generar_analisis_cartera_texto(kpis: dict):
    """Genera un resumen de texto de los KPIs para el bot."""
    comentarios = []
    
    comentarios.append(f"El porcentaje de cartera vencida es del {kpis['porcentaje_vencido']:.1f}%.")
    
    if kpis['antiguedad_prom_vencida'] > 0:
        comentarios.append(f"La antigüedad promedio de la cartera vencida es de {kpis['antiguedad_prom_vencida']:.0f} días.")
    else:
        comentarios.append("No hay cartera vencida para analizar su antigüedad.")

    if kpis['porcentaje_vencido'] > 30: 
        comentarios.append("Recomendación: ¡ALERTA CRÍTICA! Urge contactar a los clientes con más de 60 días vencidos.")
    elif kpis['porcentaje_vencido'] > 15: 
        comentarios.append("Recomendación: Es importante intensificar las gestiones de cobro para evitar el envejecimiento.")
    else:
        comentarios.append("Recomendación: La cartera está saludable, mantén el seguimiento proactivo.")
    
    return " ".join(comentarios)

def obtener_analisis_cartera(vendedor: str = "Total") -> str:
    """
    [TOOL] Calcula los KPIs clave (CSI, % Vencido) para un vendedor o la cartera total.
    Retorna un resumen ejecutivo. (No requiere datos sensibles del cliente).
    """
    cartera_procesada = cargar_y_procesar_datos()
    
    # Filtrar por vendedor si se especificó
    if vendedor and vendedor.lower() != "total":
        cartera_filtrada = cartera_procesada[
            cartera_procesada['nomvendedor_norm'] == normalizar_nombre(vendedor)
        ]
        if cartera_filtrada.empty:
            return f"No hay datos de cartera para el vendedor: {vendedor}."
    else:
        cartera_filtrada = cartera_procesada.copy()

    total_cartera = cartera_filtrada['importe'].sum()
    cartera_vencida_df = cartera_filtrada[cartera_filtrada['dias_vencido'] > 0]
    total_vencido = cartera_vencida_df['importe'].sum()
    
    porcentaje_vencido = (total_vencido / total_cartera) * 100 if total_cartera > 0 else 0
    # Cálculo de CSI y Antigüedad (como en Streamlit)
    csi = (cartera_vencida_df['importe'] * cartera_vencida_df['dias_vencido']).sum() / total_cartera if total_cartera > 0 else 0
    antiguedad_prom_vencida = (cartera_vencida_df['importe'] * cartera_vencida_df['dias_vencido']).sum() / total_vencido if total_vencido > 0 else 0
    
    kpis = {
        'total_cartera': total_cartera,
        'total_vencido': total_vencido,
        'porcentaje_vencido': porcentaje_vencido,
        'csi': csi,
        'antiguedad_prom_vencida': antiguedad_prom_vencida,
    }
    
    resumen_analisis = generar_analisis_cartera_texto(kpis)
    
    return f"Métricas clave: Cartera Total: ${total_cartera:,.0f}. Cartera Vencida: ${total_vencido:,.0f}. {resumen_analisis}"

def consultar_estado_cliente_seguro(nit: str, codigo_cliente: str) -> str:
    """
    [TOOL] Consulta el estado de cuenta. Requiere credenciales validadas.
    Retorna un resumen de la deuda total y vencida.
    """
    # 1. Validación de entradas
    if not nit or not codigo_cliente:
        return "Error: Faltan el NIT o el Código de Cliente para realizar la consulta."

    try:
        cartera_procesada = cargar_y_procesar_datos()
        
        # 2. Búsqueda por NIT y Código (Ambos deben coincidir por seguridad)
        # Se asume que 'cod_cliente' en el DataFrame es numérico o string
        datos_cliente_seleccionado = cartera_procesada[
            (cartera_procesada['nit'].astype(str) == str(nit).strip()) &
            (cartera_procesada['cod_cliente'].astype(str) == str(codigo_cliente).strip())
        ].copy()

        if datos_cliente_seleccionado.empty:
            return "Las credenciales no coinciden o no hay un estado de cuenta activo con esos datos. Por favor, verifica el NIT y el Código de Cliente."

        # 3. Cálculo de métricas
        total_cartera_cliente = datos_cliente_seleccionado['importe'].sum()
        facturas_vencidas_cliente = datos_cliente_seleccionado[
            datos_cliente_seleccionado['dias_vencido'] > 0
        ]
        total_vencido_cliente = facturas_vencidas_cliente['importe'].sum()

        nombre_cliente = datos_cliente_seleccionado.iloc[0]['nombrecliente']

        # 4. Generar Respuesta (incluyendo la URL de pagos)
        portal_link = "https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/"
        
        if total_vencido_cliente > 0:
            dias_max_vencido = int(facturas_vencidas_cliente['dias_vencido'].max())
            
            respuesta = (
                f"Hola {nombre_cliente}. Tu *deuda total es de ${total_cartera_cliente:,.0f}*. "
                f"De este monto, *${total_vencido_cliente:,.0f} está vencido*. "
                f"La factura con más antigüedad tiene {dias_max_vencido} días vencida. "
                f"Puedes pagar en nuestro portal: {portal_link}"
            )
        else:
            respuesta = (
                f"¡Hola {nombre_cliente}! ¡Felicitaciones! *No tienes facturas vencidas*. "
                f"Tu cartera total activa es de ${total_cartera_cliente:,.0f}. "
                f"Puedes consultar tus pagos futuros aquí: {portal_link}"
            )

        return respuesta

    except Exception as e:
        app.logger.error(f"Error en consultar_estado_cliente_seguro: {e}")
        return "Lo siento, hubo un error interno al consultar tu estado de cuenta. Intenta de nuevo más tarde."

# ----------------------------------------------------------------------
## 🧠 INICIALIZACIÓN DE GEMINI
# ----------------------------------------------------------------------
model = None
try:
    if not GEMINI_API_KEY:
        raise ValueError("Error: La variable 'GEMINI_API_KEY' no está configurada.")

    genai.configure(api_key=GEMINI_API_KEY)

    # --- Instrucción de Sistema para el Tono y Rol ---
    system_instruction = (
        "Eres **Ferreinox CRM AI**, un asistente de servicio al cliente, CRM y analista de cartera experto para FERREINOX SAS BIC. "
        "Tu función principal es consultar estados de cuenta (usando el NIT y Código de Cliente), analizar la gestión de cartera y ofrecer información general de la empresa. "
        "Habla de forma amable, cercana y natural, resumiendo la información de manera concisa. "
        "Tutea al cliente y nunca digas que eres un bot o una inteligencia artificial. "
        "**PROTOCOLO DE SEGURIDAD (MÁXIMA PRIORIDAD):** "
        "1. Si el cliente pide su estado de cuenta, deuda, o información financiera, *DEBES* pedirle su **NIT** y su **Código de Cliente** para poder usar la herramienta `consultar_estado_cliente_seguro`. NUNCA uses la herramienta sin tener AMBOS datos del cliente. "
        "2. Si la consulta es de análisis de negocio (ej: '¿cómo va la cartera?', '¿quién vende más?'), usa la herramienta `obtener_analisis_cartera` sin pedir credenciales. "
        "3. La página web de la empresa es www.ferreinox.co."
    )
    
    # --- Definición de las Herramientas (Tools) ---
    tools_list = [
        consultar_estado_cliente_seguro, # CRM - REQUIERE SEGURIDAD
        obtener_analisis_cartera       # BI - NO REQUIERE SEGURIDAD
    ]
    
    model = genai.GenerativeModel(
        model_name="gemini-pro-latest",
        system_instruction=system_instruction,
        tools=tools_list
    )

    # Prueba de inicialización
    model.generate_content("Hola") 
    app.logger.info("Modelo Gemini (Ferreinox CRM AI) inicializado exitosamente con Tools.")

except Exception as e:
    app.logger.error(f"Error fatal al configurar Google AI Studio o Tools: {e}")
    model = None

# ----------------------------------------------------------------------
## 💬 FUNCIONES AUXILIARES DE CHAT
# ----------------------------------------------------------------------

def send_whatsapp_message(to_number, message_text):
    """Envía un mensaje de texto de WhatsApp."""
    # (El código de esta función permanece igual al original)
    if not WHATSAPP_ACCESS_TOKEN or not WHATSAPP_PHONE_NUMBER_ID:
        app.logger.error("Error: Tokens de WhatsApp no configurados.")
        return

    url = f"https://graph.facebook.com/v19.0/{WHATSAPP_PHONE_NUMBER_ID}/messages"
    headers = {
        "Authorization": f"Bearer {WHATSAPP_ACCESS_TOKEN}",
        "Content-Type": "application/json",
    }
    data = {
        "messaging_product": "whatsapp",
        "to": to_number,
        "type": "text",
        "text": {"body": message_text},
    }

    try:
        response = requests.post(url, headers=headers, json=data)
        response.raise_for_status()
        app.logger.info(f"Respuesta enviada a {to_number}: {response.json()}")
    except requests.exceptions.RequestException as e:
        app.logger.error(f"Error al enviar mensaje de WhatsApp: {e}")
        if e.response is not None:
            app.logger.error(f"Respuesta del error de WhatsApp: {e.response.text}")

def log_to_google_sheet(timestamp, phone, user_msg, bot_msg, tool_used="N/A"):
    """Registra la conversación en la hoja de Google Sheets."""
    global worksheet
    if worksheet is None: return

    try:
        worksheet.append_row([timestamp, phone, user_msg, bot_msg, tool_used])
        app.logger.info(f"Chat loggeado en Google Sheets para {phone}")
    except Exception as e:
        app.logger.error(f"Error al escribir en Google Sheets: {e}")

def process_message_in_thread(user_phone_number, user_message, message_id):
    """
    Función que se ejecuta en un hilo separado para procesar el mensaje,
    incluyendo la lógica de Tool Calling de Gemini.
    """
    global model, user_chats, processed_message_ids

    if message_id in processed_message_ids:
        app.logger.warning(f"Mensaje duplicado (ID: {message_id}). Ignorando.")
        return
    processed_message_ids.add(message_id)
    if len(processed_message_ids) > 1000:
        processed_message_ids.clear() 

    if model is None:
        send_whatsapp_message(user_phone_number, "Lo siento, el servicio de IA de Ferreinox no está disponible.")
        return

    if user_phone_number not in user_chats:
        app.logger.info(f"Creando nueva sesión de chat para {user_phone_number}")
        user_chats[user_phone_number] = model.start_chat(history=[])
    
    chat_session = user_chats[user_phone_number]
    
    # Manejo del comando /reset
    if user_message.strip().lower() == "/reset":
        user_chats[user_phone_number] = model.start_chat(history=[])
        gemini_reply = "¡Listo! Empecemos de nuevo. ¿En qué te puedo ayudar?"
        send_whatsapp_message(user_phone_number, gemini_reply)
        log_to_google_sheet(datetime.now().isoformat(), user_phone_number, user_message, gemini_reply, "Reset")
        return

    try:
        app.logger.info(f"Enviando a Gemini...")
        
        # --- PRIMERA LLAMADA: Enviar el mensaje del usuario ---
        response = chat_session.send_message(user_message)
        
        # --- LÓGICA DE TOOL CALLING ---
        tool_function_name = "N/A"
        tool_response_text = ""
        
        while response.function_calls:
            function_calls = response.function_calls
            tool_calls = []

            for fc in function_calls:
                tool_function_name = fc.name
                app.logger.info(f"Gemini quiere llamar a la función: {tool_function_name}")
                
                # Obtener la función de Python por nombre
                func_to_call = globals().get(tool_function_name)
                
                if not func_to_call:
                    app.logger.error(f"Función no definida: {tool_function_name}")
                    tool_output = {"result": f"Error: Herramienta {tool_function_name} no encontrada."}
                else:
                    try:
                        # Ejecutar la función con los argumentos de Gemini
                        args = dict(fc.args)
                        app.logger.info(f"Argumentos para {tool_function_name}: {args}")
                        tool_output = func_to_call(**args)
                        
                        tool_response_text = tool_output # Guardar el resultado para el log
                        
                        # El resultado de la función se debe enviar de vuelta a Gemini
                        tool_calls.append(genai.types.ToolResponse(
                            name=tool_function_name,
                            response={'result': tool_output} # Formato requerido por la API
                        ))
                    except Exception as e:
                        app.logger.error(f"Error al ejecutar la herramienta {tool_function_name}: {e}")
                        tool_calls.append(genai.types.ToolResponse(
                            name=tool_function_name,
                            response={'result': f"Error en la ejecución de la función: {e}"}
                        ))
                
            # --- SEGUNDA LLAMADA: Enviar los resultados de la herramienta de vuelta a Gemini ---
            if tool_calls:
                response = chat_session.send_message(tool_calls)
            else:
                break # Salir si no hay llamadas de herramientas o hay errores
        
        # --- RESPUESTA FINAL ---
        gemini_reply = response.text
        app.logger.info(f"Respuesta final de Gemini: {gemini_reply[:50]}...")

    except Exception as e:
        app.logger.error(f"Error fatal en el proceso de chat o Tool Calling: {e}", exc_info=True)
        gemini_reply = "Perdona, hubo un error grave en la comunicación. ¿Puedes repetirme tu pregunta?"
        # Eliminar chat para un inicio limpio
        if user_phone_number in user_chats:
            del user_chats[user_phone_number]

    # 1. Enviar respuesta a WhatsApp
    send_whatsapp_message(user_phone_number, gemini_reply)

    # 2. Registrar en Google Sheets
    timestamp = datetime.now().isoformat()
    log_to_google_sheet(timestamp, user_phone_number, user_message, gemini_reply, tool_function_name)


# ----------------------------------------------------------------------
## 🌐 RUTAS DEL WEBHOOK
# ----------------------------------------------------------------------
@app.route('/webhook', methods=['GET', 'POST'])
def webhook():
    # --- Verificación del Webhook (GET) ---
    if request.method == 'GET':
        if request.args.get('hub.mode') == 'subscribe' and request.args.get('hub.verify_token') == WHATSAPP_VERIFY_TOKEN:
            app.logger.info("¡Webhook verificado!")
            return make_response(request.args.get('hub.challenge'), 200)
        else:
            app.logger.warning("Error de verificación. Tokens no coinciden.")
            return make_response('Error de verificación', 403)

    # --- Procesamiento de Mensajes (POST) ---
    if request.method == 'POST':
        data = request.get_json()
        app.logger.info(f"Payload recibido: {json.dumps(data, indent=2)[:500]}...")

        try:
            # Navegación segura en el JSON de WhatsApp
            if (data.get('entry') and 
                data['entry'][0].get('changes') and 
                data['entry'][0]['changes'][0].get('value') and 
                data['entry'][0]['changes'][0]['value'].get('messages') and
                data['entry'][0]['changes'][0]['value']['messages'][0]):

                message_info = data['entry'][0]['changes'][0]['value']['messages'][0]

                if message_info['type'] == 'text':
                    user_message = message_info['text']['body']
                    user_phone_number = message_info['from']
                    message_id = message_info['id'] 

                    app.logger.info(f"Mensaje de {user_phone_number} (ID: {message_id}): {user_message}")

                    # Despachar el trabajo a un hilo para evitar timeout de WhatsApp
                    processing_thread = threading.Thread(
                        target=process_message_in_thread,
                        args=(user_phone_number, user_message, message_id)
                    )
                    processing_thread.start()
                    
                    # Respuesta inmediata a WhatsApp (200 OK)
                    return make_response('EVENT_RECEIVED', 200)

            # Si no es un mensaje de texto procesable, igual respondemos OK
            app.logger.info("Payload recibido, pero ignorado (no es un mensaje de texto, es un estado, etc.).")
            return make_response('EVENT_RECEIVED', 200)

        except KeyError as e:
            app.logger.error(f"KeyError: Payload con estructura inesperada. Clave: {e}")
            return make_response('EVENT_RECEIVED', 200)
        except Exception as e:
            app.logger.error(f"Error general procesando el webhook POST: {e}", exc_info=True)
            return make_response('EVENT_RECEIVED', 200)

# ----------------------------------------------------------------------
## ▶️ INICIO DE LA APLICACIÓN
# ----------------------------------------------------------------------
if __name__ == '__main__':
    # Esto es principalmente para desarrollo local. En producción, el host lo maneja el servicio.
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port, debug=False)
