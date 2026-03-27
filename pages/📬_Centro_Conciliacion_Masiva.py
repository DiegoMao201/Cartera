import base64
import html
import json
import os
import re
import time
import unicodedata
from datetime import datetime
from io import BytesIO, StringIO
from urllib import error as urllib_error
from urllib import request as urllib_request

import dropbox
import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components
from fpdf import FPDF


st.set_page_config(
    page_title="Centro de Conciliacion Masiva",
    page_icon="📬",
    layout="wide",
    initial_sidebar_state="expanded",
)


COLOR_PRIMARIO = "#B21917"
COLOR_SECUNDARIO = "#E73537"
COLOR_TERCIARIO = "#F0833A"
COLOR_ACCION = "#F9B016"
COLOR_FONDO_CLARO = "#FEF4C0"
COLOR_FONDO_APP = "#F7F8FA"
COLOR_TEXTO = "#1F2937"
COLOR_OK = "#127A43"
COLOR_WARN = "#B4690E"
COLOR_DANGER = "#A61B1B"
PORTAL_PAGOS = "https://ferreinoxtiendapintuco.epayco.me/recaudo/ferreinoxrecaudoenlinea/"
WHATSAPP_LINEAS = [
    ("Armenia", "316 5219904", "https://wa.me/573165219904"),
    ("Manizales", "310 8501359", "https://wa.me/573108501359"),
    ("Pereira", "314 2087169", "https://wa.me/573142087169"),
]
HISTORY_FILE = os.path.join(os.path.dirname(__file__), "_historial_conciliacion_envios.csv")


st.markdown(
    f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Quicksand:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {{
        font-family: 'Quicksand', sans-serif;
    }}

    .stApp {{
        background:
            radial-gradient(circle at top right, rgba(249,176,22,0.15), transparent 24%),
            radial-gradient(circle at left top, rgba(178,25,23,0.08), transparent 30%),
            {COLOR_FONDO_APP};
    }}

    .block-container {{
        padding-top: 1.2rem;
        padding-bottom: 2rem;
    }}

    .hero-card {{
        background: linear-gradient(135deg, {COLOR_PRIMARIO} 0%, #7F1715 100%);
        color: white;
        border-radius: 22px;
        padding: 28px 30px;
        box-shadow: 0 18px 40px rgba(127, 23, 21, 0.18);
        margin-bottom: 1rem;
    }}

    .hero-card h1 {{
        margin: 0;
        font-size: 2rem;
        font-weight: 700;
    }}

    .hero-card p {{
        margin: 0.45rem 0 0 0;
        font-size: 1.02rem;
        color: rgba(255,255,255,0.92);
    }}

    .mini-note {{
        display: inline-block;
        margin-top: 0.9rem;
        background: rgba(255,255,255,0.12);
        border: 1px solid rgba(255,255,255,0.14);
        padding: 9px 12px;
        border-radius: 999px;
        font-size: 0.9rem;
    }}

    .panel-card {{
        background: white;
        border-radius: 18px;
        padding: 18px 20px;
        border: 1px solid rgba(148, 163, 184, 0.18);
        box-shadow: 0 8px 28px rgba(15, 23, 42, 0.05);
        margin-bottom: 1rem;
    }}

    .section-title {{
        font-size: 1.05rem;
        font-weight: 700;
        color: {COLOR_PRIMARIO};
        margin-bottom: 0.5rem;
    }}

    .metric-hint {{
        color: #6B7280;
        font-size: 0.92rem;
        margin-top: 0.25rem;
    }}

    .status-pill {{
        display: inline-block;
        border-radius: 999px;
        padding: 0.25rem 0.7rem;
        font-size: 0.82rem;
        font-weight: 700;
        margin-right: 0.35rem;
        margin-bottom: 0.35rem;
    }}

    .pill-ok {{ background: rgba(18, 122, 67, 0.12); color: {COLOR_OK}; }}
    .pill-warn {{ background: rgba(180, 105, 14, 0.12); color: {COLOR_WARN}; }}
    .pill-danger {{ background: rgba(166, 27, 27, 0.12); color: {COLOR_DANGER}; }}
    .pill-neutral {{ background: rgba(55, 65, 81, 0.08); color: #374151; }}

    .stMetric {{
        background: white;
        border-radius: 16px;
        padding: 14px;
        border-left: 6px solid {COLOR_PRIMARIO};
        box-shadow: 0 10px 26px rgba(15, 23, 42, 0.05);
    }}

    .stTabs [data-baseweb="tab"] {{
        border-radius: 10px 10px 0 0;
        font-weight: 700;
    }}

    .stTabs [aria-selected="true"] {{
        color: {COLOR_PRIMARIO};
        border-bottom: 3px solid {COLOR_PRIMARIO};
        background: rgba(249, 176, 22, 0.10);
    }}

    div.stButton > button:first-child {{
        background: linear-gradient(135deg, {COLOR_ACCION} 0%, #F5C443 100%);
        color: #111827;
        border: none;
        border-radius: 10px;
        font-weight: 700;
        box-shadow: 0 8px 18px rgba(249, 176, 22, 0.16);
    }}

    div.stButton > button:hover {{
        background: linear-gradient(135deg, {COLOR_TERCIARIO} 0%, {COLOR_ACCION} 100%);
        color: white;
    }}
</style>
""",
    unsafe_allow_html=True,
)


EMAIL_REGEX = re.compile(r"^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$", re.IGNORECASE)


def normalizar_nombre(texto: str) -> str:
    if not isinstance(texto, str):
        return ""
    texto = texto.upper().strip().replace('.', '')
    texto = ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn')
    return ' '.join(texto.split())


def normalizar_texto(texto: str) -> str:
    if not isinstance(texto, str):
        return str(texto)
    texto = unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode('utf-8').upper().strip()
    return re.sub(r'[^\w\s\.]', '', texto).strip()


def limpiar_nit(valor) -> str:
    if pd.isna(valor):
        return ""
    return re.sub(r'\D', '', str(valor))


def normalizar_email(valor) -> str:
    if pd.isna(valor):
        return ""
    email = str(valor).strip().lower().replace(' ', '')
    return email


def email_es_valido(valor: str) -> bool:
    if not valor:
        return False
    return EMAIL_REGEX.match(valor) is not None


def extraer_correos_prueba(texto: str) -> list[str]:
    if not texto:
        return []
    partes = re.split(r'[;,\n\r\t ]+', texto)
    correos = []
    for parte in partes:
        correo = normalizar_email(parte)
        if correo and correo not in correos:
            correos.append(correo)
    return correos


def hex_to_rgb(hex_color: str):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))


def obtener_columna_email(df: pd.DataFrame) -> str:
    for candidate in ["e_mail", "email", "correo", "mail"]:
        if candidate in df.columns:
            return candidate
    return ""


def obtener_columna_telefono(df: pd.DataFrame) -> str:
    for candidate in ["telefono1", "telefono", "tel1", "celular"]:
        if candidate in df.columns:
            return candidate
    return ""


def procesar_dataframe_robusto(df_raw: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()
    df.columns = [normalizar_texto(c).lower().replace(' ', '_') for c in df.columns]

    for columna in ["nomvendedor", "serie", "nombrecliente", "poblacion", "provincia"]:
        if columna not in df.columns:
            df[columna] = ""

    for columna in ["importe", "numero", "dias_vencido", "cod_cliente"]:
        if columna not in df.columns:
            df[columna] = 0

    df["nomvendedor"] = df["nomvendedor"].astype(str).str.strip()
    df["serie"] = df["serie"].astype(str)
    df["importe"] = pd.to_numeric(df["importe"], errors="coerce").fillna(0)
    df["numero"] = pd.to_numeric(df["numero"], errors="coerce").fillna(0)
    df.loc[df["numero"] < 0, "importe"] *= -1
    df["dias_vencido"] = pd.to_numeric(df["dias_vencido"], errors="coerce").fillna(0).astype(int)
    df["cod_cliente"] = pd.to_numeric(df["cod_cliente"], errors="coerce")
    df["nomvendedor_norm"] = df["nomvendedor"].apply(normalizar_nombre)
    df["nit_clean"] = df["nit"].apply(limpiar_nit) if "nit" in df.columns else ""

    if "fecha_documento" in df.columns:
        df["fecha_documento"] = pd.to_datetime(df["fecha_documento"], errors="coerce")
    else:
        df["fecha_documento"] = pd.NaT

    if "fecha_vencimiento" in df.columns:
        df["fecha_vencimiento"] = pd.to_datetime(df["fecha_vencimiento"], errors="coerce")
    else:
        df["fecha_vencimiento"] = pd.NaT

    zonas_serie = {"PEREIRA": [155, 189, 158, 439], "MANIZALES": [157, 238], "ARMENIA": [156]}
    zonas_serie_str = {zona: [str(s) for s in series] for zona, series in zonas_serie.items()}

    def asignar_zona_robusta(valor_serie):
        if pd.isna(valor_serie):
            return "OTRAS ZONAS"
        numeros_en_celda = re.findall(r'\d+', str(valor_serie))
        if not numeros_en_celda:
            return "OTRAS ZONAS"
        for zona, series_clave_str in zonas_serie_str.items():
            if set(numeros_en_celda) & set(series_clave_str):
                return zona
        return "OTRAS ZONAS"

    df["zona"] = df["serie"].apply(asignar_zona_robusta)
    df = df[~df["serie"].str.contains('W|X', case=False, na=False)].copy()

    bins = [-float('inf'), 0, 15, 30, 60, 90, float('inf')]
    labels = ["Al Dia", "1-15 dias", "16-30 dias", "31-60 dias", "61-90 dias", "+90 dias"]
    df["rango_mora"] = pd.cut(df["dias_vencido"], bins=bins, labels=labels, right=True)
    df = df[df["importe"] != 0].copy()
    return df


@st.cache_data(ttl=600)
def cargar_cartera_dropbox() -> tuple[pd.DataFrame | None, str]:
    try:
        app_key = st.secrets["dropbox"]["app_key"]
        app_secret = st.secrets["dropbox"]["app_secret"]
        refresh_token = st.secrets["dropbox"]["refresh_token"]
    except Exception as exc:
        return None, f"Error leyendo secretos de Dropbox: {exc}"

    try:
        with dropbox.Dropbox(
            app_key=app_key,
            app_secret=app_secret,
            oauth2_refresh_token=refresh_token,
        ) as dbx:
            path_archivo = "/data/cartera_detalle.csv"
            metadata, res = dbx.files_download(path=path_archivo)
            contenido_csv = res.content.decode("latin-1")
            columnas = [
                "Serie", "Numero", "Fecha Documento", "Fecha Vencimiento", "Cod Cliente",
                "NombreCliente", "Nit", "Poblacion", "Provincia", "Telefono1", "Telefono2",
                "NomVendedor", "Entidad Autoriza", "E-Mail", "Importe", "Descuento",
                "Cupo Aprobado", "Dias Vencido",
            ]
            df = pd.read_csv(
                StringIO(contenido_csv),
                header=None,
                names=columnas,
                sep="|",
                engine="python",
            )
            return procesar_dataframe_robusto(df), f"Conectado a Dropbox: {metadata.name}"
    except Exception as exc:
        return None, f"Error cargando cartera desde Dropbox: {exc}"


def construir_resumen_clientes(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    email_col = obtener_columna_email(df)
    tel_col = obtener_columna_telefono(df)
    if not email_col:
        df = df.copy()
        df["e_mail"] = ""
        email_col = "e_mail"

    if not tel_col:
        df = df.copy()
        df["telefono1"] = ""
        tel_col = "telefono1"

    base = df.copy()
    base["email_normalizado"] = base[email_col].apply(normalizar_email)
    base["telefono_normalizado"] = base[tel_col].astype(str).str.replace(r'\D', '', regex=True)
    base["importe_vencido"] = base["importe"].where(base["dias_vencido"] > 0, 0)
    base["documento_max_vencido"] = base["dias_vencido"].where(base["dias_vencido"] > 0, 0)
    base["cliente_key"] = (
        base["nit_clean"].astype(str)
        + "::"
        + base["cod_cliente"].fillna(0).astype(int).astype(str)
        + "::"
        + base["nombrecliente"].astype(str)
    )

    resumen = base.groupby("cliente_key", dropna=False).agg(
        nombrecliente=("nombrecliente", "first"),
        nit=("nit", "first"),
        nit_clean=("nit_clean", "first"),
        cod_cliente=("cod_cliente", "first"),
        nomvendedor=("nomvendedor", "first"),
        nomvendedor_norm=("nomvendedor_norm", "first"),
        zona=("zona", "first"),
        poblacion=("poblacion", "first"),
        provincia=("provincia", "first"),
        correo=("email_normalizado", "first"),
        correo_original=(email_col, "first"),
        telefono=("telefono_normalizado", "first"),
        telefono_original=(tel_col, "first"),
        saldo_total=("importe", "sum"),
        saldo_vencido=("importe_vencido", "sum"),
        dias_max_mora=("documento_max_vencido", "max"),
        facturas=("numero", "count"),
        fecha_ultima_factura=("fecha_documento", "max"),
        fecha_proximo_vencimiento=("fecha_vencimiento", "min"),
    ).reset_index()

    correos_duplicados = resumen[resumen["correo"] != ""]["correo"].value_counts()
    set_correos_compartidos = set(correos_duplicados[correos_duplicados > 1].index)

    def clasificar_estado_correo(row):
        correo = row["correo"]
        if not correo:
            return "Sin correo"
        if not email_es_valido(correo):
            return "Correo invalido"
        if correo in set_correos_compartidos:
            return "Correo compartido"
        return "Listo"

    resumen["estado_correo"] = resumen.apply(clasificar_estado_correo, axis=1)
    resumen["segmento_envio"] = "Conciliacion"
    resumen.loc[resumen["saldo_vencido"] <= 0, "segmento_envio"] = "Informativo"
    resumen.loc[resumen["dias_max_mora"] >= 31, "segmento_envio"] = "Prioritario"
    resumen["cliente_label"] = resumen.apply(
        lambda row: f"{row['nombrecliente']} | {row['saldo_vencido']:,.0f} vencido | {row['correo'] or 'sin correo'}",
        axis=1,
    )
    return resumen.sort_values(["saldo_vencido", "saldo_total"], ascending=[False, False]).reset_index(drop=True)


def preparar_reporte_calidad(resumen: pd.DataFrame) -> pd.DataFrame:
    if resumen.empty:
        return pd.DataFrame()
    reporte = resumen[[
        "nombrecliente",
        "nit",
        "cod_cliente",
        "nomvendedor",
        "zona",
        "poblacion",
        "correo_original",
        "correo",
        "estado_correo",
        "telefono_original",
        "saldo_vencido",
        "saldo_total",
        "facturas",
        "dias_max_mora",
    ]].copy()
    reporte.rename(columns={
        "nombrecliente": "Cliente",
        "nit": "NIT",
        "cod_cliente": "Codigo Cliente",
        "nomvendedor": "Vendedor",
        "zona": "Zona",
        "poblacion": "Poblacion",
        "correo_original": "Correo Base",
        "correo": "Correo Normalizado",
        "estado_correo": "Estado Correo",
        "telefono_original": "Telefono Base",
        "saldo_vencido": "Saldo Vencido",
        "saldo_total": "Saldo Total",
        "facturas": "Facturas",
        "dias_max_mora": "Max Dias Mora",
    }, inplace=True)
    return reporte


def dataframe_a_excel(dataframes: dict[str, pd.DataFrame]) -> bytes:
    salida = BytesIO()
    with pd.ExcelWriter(salida, engine="xlsxwriter") as writer:
        for hoja, df in dataframes.items():
            safe_name = hoja[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
            worksheet = writer.sheets[safe_name]
            workbook = writer.book
            encabezado = workbook.add_format({
                "bold": True,
                "fg_color": COLOR_PRIMARIO,
                "font_color": "#FFFFFF",
                "border": 1,
            })
            for idx, col in enumerate(df.columns):
                worksheet.write(0, idx, col, encabezado)
                ancho = min(max(len(str(col)) + 4, 14), 42)
                worksheet.set_column(idx, idx, ancho)
    return salida.getvalue()


def cargar_historial_envios() -> pd.DataFrame:
    columnas = [
        "Fecha", "Campana", "Modo", "Cliente", "Destino", "Correo Cliente", "Estado Correo",
        "Saldo Vencido", "Resultado", "Detalle", "Vendedor", "Zona", "Estrategia",
    ]
    if not os.path.exists(HISTORY_FILE):
        return pd.DataFrame(columns=columnas)
    try:
        historial = pd.read_csv(HISTORY_FILE)
        for columna in columnas:
            if columna not in historial.columns:
                historial[columna] = ""
        return historial[columnas].copy()
    except Exception:
        return pd.DataFrame(columns=columnas)


def guardar_historial_envios(df_resultados: pd.DataFrame):
    if df_resultados.empty:
        return
    historial_actual = cargar_historial_envios()
    historial_nuevo = pd.concat([historial_actual, df_resultados], ignore_index=True)
    historial_nuevo.to_csv(HISTORY_FILE, index=False)


class PDFEstadoCuenta(FPDF):
    def header(self):
        try:
            self.image("LOGO FERREINOX SAS BIC 2024.png", 10, 8, 72)
        except Exception:
            self.set_font("Helvetica", "B", 12)
            self.cell(72, 10, "FERREINOX S.A.S. BIC", 0, 0, "L")

        self.set_font("Helvetica", "B", 18)
        self.set_text_color(*hex_to_rgb(COLOR_PRIMARIO))
        self.cell(0, 10, "Estado de Cuenta", 0, 1, "R")
        self.set_font("Helvetica", "", 9)
        self.set_text_color(110, 110, 110)
        self.cell(0, 7, f"Fecha de generacion: {datetime.now().strftime('%Y-%m-%d %H:%M')}", 0, 1, "R")
        self.ln(6)

    def footer(self):
        self.set_y(-28)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(110, 110, 110)
        self.multi_cell(
            0,
            4,
            "Este documento soporta la conciliacion de cartera y el control administrativo del cliente.",
            0,
            "C",
        )
        self.set_font("Helvetica", "B", 9)
        self.set_text_color(*hex_to_rgb(COLOR_SECUNDARIO))
        self.cell(0, 5, "Portal de Pagos Ferreinox", 0, 1, "C", link=PORTAL_PAGOS)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(120, 120, 120)
        self.cell(0, 4, f"Pagina {self.page_no()}", 0, 0, "C")


def crear_pdf_cliente(df_cliente: pd.DataFrame, saldo_vencido: float) -> bytes:
    pdf = PDFEstadoCuenta()
    pdf.set_auto_page_break(auto=True, margin=30)
    pdf.add_page()

    if df_cliente.empty:
        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 10, "Sin facturas para este cliente.", 0, 1, "C")
        return bytes(pdf.output())

    fila = df_cliente.iloc[0]
    rgb_primario = hex_to_rgb(COLOR_PRIMARIO)
    rgb_terciario = hex_to_rgb(COLOR_TERCIARIO)

    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(*rgb_primario)
    pdf.cell(36, 7, "Cliente:", 0, 0)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 7, str(fila.get("nombrecliente", "")), 0, 1)

    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(*rgb_primario)
    pdf.cell(36, 7, "NIT:", 0, 0)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 7, str(fila.get("nit", "")), 0, 1)

    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(*rgb_primario)
    pdf.cell(36, 7, "Codigo Cliente:", 0, 0)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", "", 11)
    cod_cliente = fila.get("cod_cliente")
    cod_texto = str(int(cod_cliente)) if pd.notna(cod_cliente) else "N/A"
    pdf.cell(0, 7, cod_texto, 0, 1)
    pdf.ln(4)

    mensaje = (
        "A continuacion se presenta el detalle consolidado de su estado de cuenta. "
        "Este envio tiene enfoque de conciliacion administrativa y validacion de saldos, "
        "facilitando la revision oportuna de la informacion financiera del cliente."
    )
    pdf.set_text_color(95, 95, 95)
    pdf.set_font("Helvetica", "", 10)
    pdf.multi_cell(0, 5, mensaje, 0, "J")
    pdf.ln(5)

    pdf.set_font("Helvetica", "B", 10)
    pdf.set_fill_color(*rgb_primario)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(26, 8, "Factura", 1, 0, "C", 1)
    pdf.cell(28, 8, "Dias Mora", 1, 0, "C", 1)
    pdf.cell(38, 8, "Fecha Doc.", 1, 0, "C", 1)
    pdf.cell(38, 8, "Fecha Venc.", 1, 0, "C", 1)
    pdf.cell(40, 8, "Saldo", 1, 1, "C", 1)

    total = 0
    pdf.set_font("Helvetica", "", 10)
    for _, item in df_cliente.sort_values(by=["dias_vencido", "fecha_vencimiento"], ascending=[False, True]).iterrows():
        total += float(item.get("importe", 0) or 0)
        if int(item.get("dias_vencido", 0) or 0) > 0:
            pdf.set_fill_color(255, 245, 238)
            pdf.set_text_color(*rgb_terciario)
        else:
            pdf.set_fill_color(255, 255, 255)
            pdf.set_text_color(0, 0, 0)

        fecha_doc = item.get("fecha_documento")
        fecha_venc = item.get("fecha_vencimiento")
        fecha_doc_txt = fecha_doc.strftime("%d/%m/%Y") if pd.notna(fecha_doc) else "-"
        fecha_venc_txt = fecha_venc.strftime("%d/%m/%Y") if pd.notna(fecha_venc) else "-"
        numero = item.get("numero")
        numero_txt = str(int(numero)) if pd.notna(numero) else "-"
        dias_txt = str(int(item.get("dias_vencido", 0) or 0))

        pdf.cell(26, 8, numero_txt, 1, 0, "C", 1)
        pdf.cell(28, 8, dias_txt, 1, 0, "C", 1)
        pdf.cell(38, 8, fecha_doc_txt, 1, 0, "C", 1)
        pdf.cell(38, 8, fecha_venc_txt, 1, 0, "C", 1)
        pdf.cell(40, 8, f"${float(item.get('importe', 0) or 0):,.0f}", 1, 1, "R", 1)

    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(130, 9, "TOTAL CARTERA", 1, 0, "R")
    pdf.cell(40, 9, f"${total:,.0f}", 1, 1, "R")

    if saldo_vencido > 0:
        pdf.set_fill_color(*rgb_primario)
        pdf.set_text_color(255, 255, 255)
        pdf.cell(130, 9, "TOTAL VENCIDO", 1, 0, "R", 1)
        pdf.cell(40, 9, f"${saldo_vencido:,.0f}", 1, 1, "R", 1)

    return bytes(pdf.output())


def construir_texto_asunto(cliente: str, estrategia: str, saldo_vencido: float) -> str:
    if estrategia == "Seguimiento prioritario" and saldo_vencido > 0:
        return f"Revision prioritaria de cartera - {cliente}"
    if estrategia == "Cierre operativo":
        return f"Actualizacion de estado de cuenta - {cliente}"
    return f"Conciliacion de cartera Ferreinox - {cliente}"


def construir_resumen_bloques(cliente_row: pd.Series) -> str:
    bloques = [
        ("Saldo total", f"${cliente_row['saldo_total']:,.0f}"),
        ("Saldo vencido", f"${cliente_row['saldo_vencido']:,.0f}"),
        ("Facturas", f"{int(cliente_row['facturas'])}"),
        ("Max mora", f"{int(cliente_row['dias_max_mora'])} dias"),
    ]
    html_bloques = []
    for titulo, valor in bloques:
        html_bloques.append(
            f"""
            <td style=\"padding:8px;\">
                <table role=\"presentation\" width=\"100%\" style=\"background:#fff7e8;border:1px solid #f6d6a1;border-radius:16px;\">
                    <tr><td style=\"padding:14px 16px 6px 16px;font-family:Quicksand,Arial,sans-serif;font-size:13px;color:#7c2d12;font-weight:700;\">{titulo}</td></tr>
                    <tr><td style=\"padding:0 16px 14px 16px;font-family:Quicksand,Arial,sans-serif;font-size:24px;color:#7f1715;font-weight:700;\">{valor}</td></tr>
                </table>
            </td>
            """
        )
    return "".join(html_bloques)


def plantilla_correo_conciliacion(cliente_row: pd.Series, estrategia: str) -> str:
    cliente = html.escape(str(cliente_row["nombrecliente"]))
    nit = html.escape(str(cliente_row.get("nit", "")))
    codigo = str(int(cliente_row["cod_cliente"])) if pd.notna(cliente_row["cod_cliente"]) else "N/A"
    saldo_vencido = float(cliente_row["saldo_vencido"])
    saldo_total = float(cliente_row["saldo_total"])
    dias_max = int(cliente_row["dias_max_mora"])

    if estrategia == "Seguimiento prioritario" and saldo_vencido > 0:
        titulo = "Revision prioritaria de cartera"
        subtitulo = (
            "Identificamos saldos vencidos que ameritan una validacion prioritaria para alinear cartera, despacho y continuidad operativa."
        )
        etiqueta = "Accion prioritaria"
        fondo_aviso = "#fff1f2"
        color_aviso = "#9f1239"
    elif estrategia == "Cierre operativo":
        titulo = "Estado de cuenta para cierre operativo"
        subtitulo = (
            "Compartimos el consolidado actualizado de su cartera para facilitar la conciliacion administrativa y el control interno."
        )
        etiqueta = "Control administrativo"
        fondo_aviso = "#eff6ff"
        color_aviso = "#1d4ed8"
    else:
        titulo = "Conciliacion de cartera Ferreinox"
        subtitulo = (
            "Enviamos el estado de cuenta actualizado con enfoque de conciliacion, revision de saldos y soporte administrativo."
        )
        etiqueta = "Conciliacion cordial"
        fondo_aviso = "#fff7ed"
        color_aviso = "#b45309"

    titulo_secundario = "Mesa central de cartera Ferreinox"
    bloque_principal = (
        f"Saldo vencido por validar: <strong>${saldo_vencido:,.0f}</strong>" if saldo_vencido > 0 else
        f"Saldo total para control: <strong>${saldo_total:,.0f}</strong>"
    )

    tono_accion = (
        "Sugerimos validar este corte con su equipo contable y confirmar cualquier novedad de pago, compensacion o cruce en proceso."
        if saldo_vencido > 0 else
        "El documento adjunto queda como soporte de control para su cierre administrativo y seguimiento interno."
    )

    lineas_whatsapp = []
    for ciudad, numero, link in WHATSAPP_LINEAS:
        lineas_whatsapp.append(
            f"<a href=\"{link}\" style=\"display:inline-block;margin:6px 4px;padding:10px 14px;border-radius:999px;background:#25d366;color:#ffffff;font-family:Quicksand,Arial,sans-serif;font-size:12px;font-weight:700;text-decoration:none;\">{ciudad}: {numero}</a>"
        )

    return f"""
<!doctype html>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Quicksand:wght@400;500;600;700&display=swap');
    </style>
</head>
<body style="margin:0;padding:24px;background:#eef2f7;">
    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#f3f4f6;">
        <tr>
            <td align="center">
                <table role="presentation" width="680" cellspacing="0" cellpadding="0" style="max-width:680px;background:#ffffff;border-radius:30px;overflow:hidden;box-shadow:0 24px 55px rgba(15,23,42,0.14);">
                    <tr>
                        <td style="background:linear-gradient(135deg, {COLOR_PRIMARIO} 0%, #7f1715 62%, #c98d10 100%);padding:36px 36px 32px 36px;position:relative;">
                            <div style="display:inline-block;background:rgba(255,255,255,0.12);border:1px solid rgba(255,255,255,0.18);padding:7px 12px;border-radius:999px;font-family:Quicksand,Arial,sans-serif;font-size:12px;color:#ffffff;font-weight:700;letter-spacing:0.3px;">{etiqueta}</div>
                            <div style="font-family:Quicksand,Arial,sans-serif;font-size:32px;line-height:1.12;color:#ffffff;font-weight:700;margin-top:14px;max-width:430px;">{titulo}</div>
                            <div style="font-family:Quicksand,Arial,sans-serif;font-size:16px;line-height:1.7;color:rgba(255,255,255,0.94);margin-top:12px;max-width:470px;">{subtitulo}</div>
                            <div style="margin-top:18px;font-family:Quicksand,Arial,sans-serif;font-size:13px;color:rgba(255,255,255,0.88);font-weight:600;">{titulo_secundario}</div>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding:32px 36px 10px 36px;">
                            <div style="font-family:Quicksand,Arial,sans-serif;font-size:19px;color:{COLOR_TEXTO};font-weight:700;">Apreciado cliente, {cliente}</div>
                            <div style="font-family:Quicksand,Arial,sans-serif;font-size:15px;line-height:1.8;color:#4b5563;margin-top:14px;">
                                Compartimos su estado de cuenta adjunto como soporte para la conciliacion de cartera. El objetivo es facilitar la verificacion de la informacion, validar diferencias si existen y mantener su expediente financiero al dia.
                            </div>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding:8px 28px 0 28px;">
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
                                <tr>
                                    {construir_resumen_bloques(cliente_row)}
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding:18px 36px 0 36px;">
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:{fondo_aviso};border:1px solid rgba(127,23,21,0.10);border-radius:18px;">
                                <tr>
                                    <td style="padding:18px 20px;font-family:Quicksand,Arial,sans-serif;font-size:15px;line-height:1.7;color:{color_aviso};">
                                        {bloque_principal}. Si ya existe pago aplicado o cruce pendiente, este correo sirve como punto de control para su equipo administrativo y el nuestro.
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding:16px 36px 0 36px;">
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#111827;border-radius:20px;overflow:hidden;">
                                <tr>
                                    <td style="padding:18px 20px 12px 20px;font-family:Quicksand,Arial,sans-serif;font-size:13px;color:#fbbf24;font-weight:700;letter-spacing:0.3px;">Lectura recomendada del envio</td>
                                </tr>
                                <tr>
                                    <td style="padding:0 20px 20px 20px;font-family:Quicksand,Arial,sans-serif;font-size:15px;line-height:1.8;color:#e5e7eb;">{tono_accion}</td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding:22px 36px 0 36px;">
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0">
                                <tr>
                                    <td width="48%" valign="top" style="padding-right:12px;">
                                        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#faf5ef;border:1px solid #f5dac0;border-radius:18px;">
                                            <tr><td style="padding:18px 20px 8px 20px;font-family:Quicksand,Arial,sans-serif;font-size:13px;color:#7c2d12;font-weight:700;">NIT / Documento</td></tr>
                                            <tr><td style="padding:0 20px 18px 20px;font-family:Quicksand,Arial,sans-serif;font-size:22px;color:#111827;font-weight:700;">{nit}</td></tr>
                                            <tr><td style="padding:0 20px 8px 20px;font-family:Quicksand,Arial,sans-serif;font-size:13px;color:#7c2d12;font-weight:700;">Codigo Cliente</td></tr>
                                            <tr><td style="padding:0 20px 20px 20px;font-family:Quicksand,Arial,sans-serif;font-size:22px;color:#111827;font-weight:700;">{codigo}</td></tr>
                                        </table>
                                    </td>
                                    <td width="52%" valign="top" style="padding-left:12px;">
                                        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#fffdf8;border:1px solid #fde6b1;border-radius:18px;">
                                            <tr><td style="padding:18px 20px 10px 20px;font-family:Quicksand,Arial,sans-serif;font-size:13px;color:#92400e;font-weight:700;">Acceso rapido</td></tr>
                                            <tr><td style="padding:0 20px 10px 20px;font-family:Quicksand,Arial,sans-serif;font-size:15px;line-height:1.7;color:#4b5563;">Puede usar el portal de pagos o responder este correo para validar novedades sobre su conciliacion.</td></tr>
                                            <tr><td style="padding:4px 20px 20px 20px;"><a href="{PORTAL_PAGOS}" style="display:inline-block;background:{COLOR_ACCION};color:#111827;padding:14px 22px;border-radius:12px;font-family:Quicksand,Arial,sans-serif;font-size:14px;font-weight:700;text-decoration:none;">Ir al portal de pagos</a></td></tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding:22px 36px 0 36px;">
                            <div style="font-family:Quicksand,Arial,sans-serif;font-size:15px;line-height:1.8;color:#4b5563;">
                                Adjuntamos el estado de cuenta en PDF para su revision. Si desea actualizar correos de facturacion, tesoreria o cartera, puede responder este mensaje y así corregimos la base para futuros envios.
                            </div>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding:18px 36px 0 36px;">
                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="background:#fff9ec;border:1px solid #f6d38a;border-radius:18px;">
                                <tr>
                                    <td style="padding:18px 20px;font-family:Quicksand,Arial,sans-serif;font-size:14px;line-height:1.8;color:#6b4f0c;">
                                        <strong>Nota de claridad:</strong> este correo tiene enfoque de conciliacion y organizacion de informacion financiera. Si su equipo ya cuenta con soporte de pago o novedad pendiente de aplicacion, bastara con responder este mensaje para actualizar el caso.
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td style="padding:26px 36px 32px 36px;background:#1f2937;">
                            <div style="font-family:Quicksand,Arial,sans-serif;font-size:18px;color:#ffffff;font-weight:700;text-align:center;">Area de Cartera y Recaudos</div>
                            <div style="font-family:Quicksand,Arial,sans-serif;font-size:14px;line-height:1.7;color:#d1d5db;text-align:center;margin-top:10px;">Canales de atencion Ferreinox</div>
                            <div style="text-align:center;margin-top:10px;">{''.join(lineas_whatsapp)}</div>
                            <div style="font-family:Quicksand,Arial,sans-serif;font-size:12px;line-height:1.7;color:#9ca3af;text-align:center;margin-top:16px;">Envio operativo de conciliacion de cartera. Si ya aplico el pago o existe acuerdo en proceso, por favor ignore este recordatorio.</div>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>
"""


def cuerpo_texto_plano(cliente_row: pd.Series, estrategia: str) -> str:
    cliente = str(cliente_row["nombrecliente"])
    saldo_total = float(cliente_row["saldo_total"])
    saldo_vencido = float(cliente_row["saldo_vencido"])
    dias_max = int(cliente_row["dias_max_mora"])
    return (
        f"Ferreinox S.A.S. BIC\n\n"
        f"{estrategia}\n\n"
        f"Cliente: {cliente}\n"
        f"Saldo total: ${saldo_total:,.0f}\n"
        f"Saldo vencido: ${saldo_vencido:,.0f}\n"
        f"Maximo de mora: {dias_max} dias\n"
        f"Portal de pagos: {PORTAL_PAGOS}\n\n"
        "Adjuntamos el estado de cuenta en PDF para control y conciliacion administrativa."
    )


def enviar_con_sendgrid(
    api_key: str,
    from_email: str,
    from_name: str,
    to_email: str,
    subject: str,
    html_content: str,
    plain_content: str,
    attachment_bytes: bytes,
    attachment_name: str,
    cliente_row: pd.Series,
) -> tuple[bool, str]:
    payload = {
        "personalizations": [{
            "to": [{"email": to_email, "name": str(cliente_row["nombrecliente"])}],
            "custom_args": {
                "cliente": str(cliente_row["nombrecliente"]),
                "nit": str(cliente_row.get("nit_clean", "")),
                "zona": str(cliente_row.get("zona", "")),
            },
        }],
        "from": {"email": from_email, "name": from_name},
        "subject": subject,
        "content": [
            {"type": "text/plain", "value": plain_content},
            {"type": "text/html", "value": html_content},
        ],
        "attachments": [{
            "content": base64.b64encode(attachment_bytes).decode("utf-8"),
            "type": "application/pdf",
            "filename": attachment_name,
            "disposition": "attachment",
        }],
        "tracking_settings": {
            "open_tracking": {"enable": True},
            "click_tracking": {"enable": True, "enable_text": False},
        },
        "categories": ["conciliacion-cartera", "ferreinox"],
    }

    request = urllib_request.Request(
        url="https://api.sendgrid.com/v3/mail/send",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )

    try:
        with urllib_request.urlopen(request, timeout=45) as response:
            status = response.getcode()
            if 200 <= status < 300:
                return True, f"HTTP {status}"
            return False, f"HTTP {status}"
    except urllib_error.HTTPError as exc:
        detalle = exc.read().decode("utf-8", errors="ignore")
        return False, f"HTTP {exc.code}: {detalle[:220]}"
    except Exception as exc:
        return False, str(exc)


def render_login():
    st.markdown(
        f"""
        <div class="hero-card">
            <h1>Centro de Conciliacion Masiva</h1>
            <p>Acceso protegido para la gestion estructurada de envios de cartera Ferreinox.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    try:
        general_password = st.secrets["general"]["password"]
        vendedores_secrets = st.secrets["vendedores"]
    except Exception as exc:
        st.error(f"Error leyendo secretos de autenticacion: {exc}")
        st.stop()

    password = st.text_input("Contraseña", type="password", key="password_conciliacion")
    if st.button("Ingresar al centro de control"):
        if password == str(general_password):
            st.session_state["authentication_status"] = True
            st.session_state["acceso_general"] = True
            st.session_state["vendedor_autenticado"] = "GERENCIA"
            st.rerun()
        for key, value in vendedores_secrets.items():
            if password == str(value):
                st.session_state["authentication_status"] = True
                st.session_state["acceso_general"] = False
                st.session_state["vendedor_autenticado"] = key
                st.rerun()
        st.error("Acceso denegado")
    st.stop()


def inicializar_estado():
    if "authentication_status" not in st.session_state:
        st.session_state["authentication_status"] = False
        st.session_state["acceso_general"] = False
        st.session_state["vendedor_autenticado"] = None
    if "seleccion_clientes_conciliacion" not in st.session_state:
        st.session_state["seleccion_clientes_conciliacion"] = []
    if "reporte_envio_conciliacion" not in st.session_state:
        st.session_state["reporte_envio_conciliacion"] = []
    if "historial_envio_conciliacion" not in st.session_state:
        st.session_state["historial_envio_conciliacion"] = cargar_historial_envios().to_dict("records")


def aplicar_filtros(df_resumen: pd.DataFrame) -> pd.DataFrame:
    df_view = df_resumen.copy()
    st.sidebar.header("Filtros de control")

    if st.session_state.get("acceso_general", False):
        vendedores = ["TODOS"] + sorted(df_view["nomvendedor"].dropna().astype(str).unique().tolist())
        vendedor_sel = st.sidebar.selectbox("Vendedor", vendedores)
        if vendedor_sel != "TODOS":
            df_view = df_view[df_view["nomvendedor_norm"] == normalizar_nombre(vendedor_sel)]
    else:
        vendedor_activo = normalizar_nombre(st.session_state.get("vendedor_autenticado", ""))
        df_view = df_view[df_view["nomvendedor_norm"] == vendedor_activo]
        st.sidebar.info(f"Vista del vendedor: {st.session_state.get('vendedor_autenticado', 'N/A')}")

    zonas = ["TODAS"] + sorted(df_view["zona"].dropna().astype(str).unique().tolist())
    zona_sel = st.sidebar.selectbox("Zona", zonas)
    if zona_sel != "TODAS":
        df_view = df_view[df_view["zona"] == zona_sel]

    poblaciones = ["TODAS"] + sorted(df_view["poblacion"].dropna().astype(str).unique().tolist())
    poblacion_sel = st.sidebar.selectbox("Poblacion", poblaciones)
    if poblacion_sel != "TODAS":
        df_view = df_view[df_view["poblacion"] == poblacion_sel]

    rangos = ["TODOS"] + [str(valor) for valor in df_view["segmento_envio"].dropna().unique().tolist()]
    segmento_sel = st.sidebar.selectbox("Segmento envio", rangos)
    if segmento_sel != "TODOS":
        df_view = df_view[df_view["segmento_envio"] == segmento_sel]

    estados = ["TODOS", "Listo", "Correo compartido", "Correo invalido", "Sin correo"]
    estado_sel = st.sidebar.selectbox("Estado de correo", estados)
    if estado_sel != "TODOS":
        df_view = df_view[df_view["estado_correo"] == estado_sel]

    solo_con_vencido = st.sidebar.toggle("Solo con saldo vencido", value=True)
    if solo_con_vencido:
        df_view = df_view[df_view["saldo_vencido"] > 0]

    minimo_vencido = st.sidebar.number_input("Saldo vencido minimo", min_value=0, value=0, step=50000)
    if minimo_vencido > 0:
        df_view = df_view[df_view["saldo_vencido"] >= minimo_vencido]

    minimo_dias = st.sidebar.slider("Dias de mora minimo", min_value=0, max_value=180, value=0, step=5)
    if minimo_dias > 0:
        df_view = df_view[df_view["dias_max_mora"] >= minimo_dias]

    return df_view.reset_index(drop=True)


def render_sidebar(status_carga: str):
    with st.sidebar:
        try:
            st.image("LOGO FERREINOX SAS BIC 2024.png", use_container_width=True)
        except Exception:
            pass
        st.success(f"Usuario activo: {st.session_state.get('vendedor_autenticado', 'N/A')}")
        if st.button("Cerrar sesion"):
            st.session_state.clear()
            st.rerun()
        if st.button("Recargar Dropbox", type="primary"):
            st.cache_data.clear()
            st.rerun()
        st.caption(status_carga)


def seleccionar_clientes(df_view: pd.DataFrame):
    listos = df_view[df_view["estado_correo"].isin(["Listo", "Correo compartido"])]["cliente_label"].tolist()
    disponibles = df_view["cliente_label"].tolist()

    col_a, col_b, col_c = st.columns([1, 1, 2])
    with col_a:
        if st.button("Seleccionar listos del filtro"):
            st.session_state["seleccion_clientes_conciliacion"] = listos
    with col_b:
        if st.button("Limpiar seleccion"):
            st.session_state["seleccion_clientes_conciliacion"] = []
    with col_c:
        st.caption("La seleccion masiva solo toma clientes visibles con el filtro actual.")

    seleccion = st.multiselect(
        "Clientes seleccionados para el envio o la vista previa",
        options=disponibles,
        default=[valor for valor in st.session_state.get("seleccion_clientes_conciliacion", []) if valor in disponibles],
        key="multiselect_conciliacion_clientes",
    )
    st.session_state["seleccion_clientes_conciliacion"] = seleccion
    return df_view[df_view["cliente_label"].isin(seleccion)].copy()


def render_metricas(df_view: pd.DataFrame):
    total_clientes = len(df_view)
    total_listos = int(df_view["estado_correo"].eq("Listo").sum())
    total_compartidos = int(df_view["estado_correo"].eq("Correo compartido").sum())
    total_invalidos = int(df_view["estado_correo"].eq("Correo invalido").sum())
    total_sin_correo = int(df_view["estado_correo"].eq("Sin correo").sum())
    total_vencido = float(df_view["saldo_vencido"].sum())

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Clientes visibles", total_clientes)
    c2.metric("Listos", total_listos)
    c3.metric("Compartidos", total_compartidos)
    c4.metric("Invalidos", total_invalidos)
    c5.metric("Sin correo", total_sin_correo)
    c6.metric("Vencido visible", f"${total_vencido:,.0f}")


def render_control_tab(df_view: pd.DataFrame, seleccionados: pd.DataFrame):
    st.markdown('<div class="section-title">Base filtrada para conciliacion y envio</div>', unsafe_allow_html=True)
    columnas = [
        "nombrecliente", "nomvendedor", "zona", "poblacion", "correo", "estado_correo",
        "saldo_vencido", "saldo_total", "facturas", "dias_max_mora",
    ]
    vista = df_view[columnas].copy().rename(columns={
        "nombrecliente": "Cliente",
        "nomvendedor": "Vendedor",
        "zona": "Zona",
        "poblacion": "Poblacion",
        "correo": "Correo",
        "estado_correo": "Estado Correo",
        "saldo_vencido": "Saldo Vencido",
        "saldo_total": "Saldo Total",
        "facturas": "Facturas",
        "dias_max_mora": "Max Mora",
    })
    st.dataframe(
        vista.style.format({"Saldo Vencido": "${:,.0f}", "Saldo Total": "${:,.0f}"}),
        use_container_width=True,
        hide_index=True,
    )

    col_1, col_2 = st.columns([2, 1])
    with col_1:
        st.markdown('<div class="panel-card"><div class="section-title">Estado de seleccion</div>', unsafe_allow_html=True)
        st.write(f"Clientes seleccionados: {len(seleccionados)}")
        if not seleccionados.empty:
            st.write(f"Saldo vencido en seleccion: ${seleccionados['saldo_vencido'].sum():,.0f}")
            st.write(f"Clientes con correo utilizable: {int(seleccionados['estado_correo'].isin(['Listo', 'Correo compartido']).sum())}")
        st.markdown('</div>', unsafe_allow_html=True)
    with col_2:
        conteo = df_view["estado_correo"].value_counts().reset_index()
        if not conteo.empty:
            fig = px.pie(conteo, names="estado_correo", values="count", hole=0.45, color="estado_correo",
                         color_discrete_map={
                             "Listo": COLOR_OK,
                             "Correo compartido": COLOR_ACCION,
                             "Correo invalido": COLOR_TERCIARIO,
                             "Sin correo": COLOR_PRIMARIO,
                         })
            fig.update_layout(margin=dict(t=10, l=0, r=0, b=0), showlegend=True)
            st.plotly_chart(fig, use_container_width=True)


def render_preview_tab(df_base: pd.DataFrame, resumen: pd.DataFrame, seleccionados: pd.DataFrame):
    universo = seleccionados if not seleccionados.empty else resumen
    if universo.empty:
        st.info("No hay clientes disponibles para vista previa.")
        return

    opciones = universo["cliente_label"].tolist()
    cliente_label = st.selectbox("Cliente para vista previa", opciones, key="preview_cliente_conciliacion")
    estrategia = st.selectbox(
        "Estilo del correo",
        ["Conciliacion cordial", "Cierre operativo", "Seguimiento prioritario"],
        key="preview_estrategia_conciliacion",
    )
    fila = universo[universo["cliente_label"] == cliente_label].iloc[0]
    df_cliente = df_base[df_base["cliente_key"] == fila["cliente_key"]].copy()
    pdf_bytes = crear_pdf_cliente(df_cliente, float(fila["saldo_vencido"]))
    html_body = plantilla_correo_conciliacion(fila, estrategia)

    info_1, info_2, info_3, info_4 = st.columns(4)
    info_1.metric("Correo", fila["correo"] or "Sin correo")
    info_2.metric("Estado", fila["estado_correo"])
    info_3.metric("Saldo vencido", f"${fila['saldo_vencido']:,.0f}")
    info_4.metric("Facturas", int(fila["facturas"]))

    st.download_button(
        "Descargar PDF del cliente",
        data=pdf_bytes,
        file_name=f"Estado_Cuenta_{normalizar_nombre(str(fila['nombrecliente'])).replace(' ', '_')}.pdf",
        mime="application/pdf",
    )

    st.markdown('<div class="panel-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Laboratorio de prueba a 3 correos</div>', unsafe_allow_html=True)
    correos_prueba_texto = st.text_area(
        "Correos de prueba",
        value="",
        placeholder="correo1@dominio.com\ncorreo2@dominio.com\ncorreo3@dominio.com",
        help="Escribe hasta 3 correos separados por coma o salto de linea.",
        key="preview_correos_prueba_texto",
    )
    estrategia_prueba = st.selectbox(
        "Estilo para la prueba",
        ["Conciliacion cordial", "Cierre operativo", "Seguimiento prioritario"],
        key="preview_estrategia_prueba_conciliacion",
    )
    correos_prueba = extraer_correos_prueba(correos_prueba_texto)
    invalidos = [correo for correo in correos_prueba if not email_es_valido(correo)]
    if len(correos_prueba) > 3:
        st.warning("Solo se permiten hasta 3 correos de prueba. Se usaran los primeros 3.")
        correos_prueba = correos_prueba[:3]
    if invalidos:
        st.error(f"Correos invalidos detectados: {', '.join(invalidos)}")
    elif correos_prueba:
        st.caption(f"Destino de prueba: {', '.join(correos_prueba)}")

    if st.button("Enviar prueba del cliente seleccionado", disabled=not correos_prueba or bool(invalidos)):
        if "sendgrid" not in st.secrets:
            st.error("No existe configuracion de SendGrid en secrets.")
        else:
            api_key = st.secrets["sendgrid"].get("api_key", "")
            from_email = st.secrets["sendgrid"].get("from_email", "")
            from_name = st.secrets["sendgrid"].get("from_name", "Ferreinox S.A.S. BIC")
            if not api_key or not from_email:
                st.error("La configuracion de SendGrid esta incompleta.")
            else:
                resultados_prueba = []
                progress = st.progress(0)
                for idx, correo_destino in enumerate(correos_prueba, start=1):
                    asunto = f"[PRUEBA] {construir_texto_asunto(str(fila['nombrecliente']), estrategia_prueba, float(fila['saldo_vencido']))}"
                    html_prueba = plantilla_correo_conciliacion(fila, estrategia_prueba)
                    texto_prueba = cuerpo_texto_plano(fila, estrategia_prueba)
                    nombre_pdf = f"PRUEBA_{normalizar_nombre(str(fila['nombrecliente'])).replace(' ', '_')}.pdf"
                    ok, detalle = enviar_con_sendgrid(
                        api_key=api_key,
                        from_email=from_email,
                        from_name=from_name,
                        to_email=correo_destino,
                        subject=asunto,
                        html_content=html_prueba,
                        plain_content=texto_prueba,
                        attachment_bytes=pdf_bytes,
                        attachment_name=nombre_pdf,
                        cliente_row=fila,
                    )
                    resultados_prueba.append({
                        "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Campana": "PRUEBA VISUAL",
                        "Modo": "Prueba",
                        "Cliente": fila["nombrecliente"],
                        "Destino": correo_destino,
                        "Correo Cliente": fila["correo"],
                        "Estado Correo": fila["estado_correo"],
                        "Saldo Vencido": float(fila["saldo_vencido"]),
                        "Resultado": "Enviado" if ok else "Error",
                        "Detalle": detalle,
                        "Vendedor": fila["nomvendedor"],
                        "Zona": fila["zona"],
                        "Estrategia": estrategia_prueba,
                    })
                    progress.progress(idx / len(correos_prueba))

                df_prueba = pd.DataFrame(resultados_prueba)
                guardar_historial_envios(df_prueba)
                st.session_state["historial_envio_conciliacion"] = cargar_historial_envios().to_dict("records")
                total_ok = int(df_prueba["Resultado"].eq("Enviado").sum())
                if total_ok == len(df_prueba):
                    st.success(f"Prueba completada. {total_ok} correos enviados.")
                else:
                    st.warning("La prueba termino con algunas novedades. Revisa el historial.")
                st.dataframe(df_prueba, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    components.html(html_body, height=980, scrolling=True)


def render_envio_tab(df_base: pd.DataFrame, seleccionados: pd.DataFrame):
    st.markdown('<div class="section-title">Despacho masivo con SendGrid</div>', unsafe_allow_html=True)

    if "sendgrid" not in st.secrets:
        st.error(
            "Falta la seccion sendgrid en secrets.toml. Debe existir con api_key, from_email y from_name."
        )
        st.code(
            "[sendgrid]\napi_key = \"SG.xxxxx\"\nfrom_email = \"tiendapintucopereira@ferreinox.co\"\nfrom_name = \"Ferreinox S.A.S. BIC\""
        )
        return

    if seleccionados.empty:
        st.warning("Selecciona clientes en la pestaña de control antes de enviar.")
        return

    elegibles = seleccionados[seleccionados["estado_correo"].isin(["Listo", "Correo compartido"])].copy()
    bloqueados = seleccionados[~seleccionados["estado_correo"].isin(["Listo", "Correo compartido"])].copy()

    c1, c2, c3 = st.columns(3)
    c1.metric("Seleccionados", len(seleccionados))
    c2.metric("Elegibles para envio", len(elegibles))
    c3.metric("Bloqueados por calidad", len(bloqueados))

    nombre_campana = st.text_input(
        "Nombre de campana o corte",
        value=f"Corte cartera {datetime.now().strftime('%Y-%m-%d')}",
        key="nombre_campana_conciliacion",
        help="Este nombre quedara guardado en el historial del centro de control.",
    )

    estrategia = st.selectbox(
        "Estrategia de correo para el lote",
        ["Conciliacion cordial", "Cierre operativo", "Seguimiento prioritario"],
        key="estrategia_envio_conciliacion",
    )
    incluir_compartidos = st.toggle("Permitir correos compartidos", value=False)
    pausa_ms = st.slider("Pausa entre envios (milisegundos)", min_value=0, max_value=1500, value=150, step=50)

    st.markdown('<div class="panel-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Modo prueba de lote</div>', unsafe_allow_html=True)
    st.caption("Usa este bloque para hacer pruebas controladas sin tocar los correos reales de los clientes.")
    clientes_prueba_lote = st.multiselect(
        "Clientes muestra para la prueba",
        options=elegibles["cliente_label"].tolist(),
        default=elegibles["cliente_label"].tolist()[:1],
        key="clientes_prueba_lote_conciliacion",
    )
    correos_prueba_lote_texto = st.text_area(
        "Correos de laboratorio",
        value="",
        placeholder="gerencia@ferreinox.co\npruebas@ferreinox.co\notro@dominio.com",
        key="correos_prueba_lote_texto",
    )
    correos_prueba_lote = extraer_correos_prueba(correos_prueba_lote_texto)
    if len(correos_prueba_lote) > 3:
        st.warning("En modo prueba solo se permiten 3 correos destino. Se usaran los primeros 3.")
        correos_prueba_lote = correos_prueba_lote[:3]
    invalidos_prueba = [correo for correo in correos_prueba_lote if not email_es_valido(correo)]
    if invalidos_prueba:
        st.error(f"Correos de laboratorio invalidos: {', '.join(invalidos_prueba)}")

    if st.button("Ejecutar prueba de lote", disabled=not clientes_prueba_lote or not correos_prueba_lote or bool(invalidos_prueba)):
        api_key = st.secrets["sendgrid"].get("api_key", "")
        from_email = st.secrets["sendgrid"].get("from_email", "")
        from_name = st.secrets["sendgrid"].get("from_name", "Ferreinox S.A.S. BIC")
        if not api_key or not from_email:
            st.error("La configuracion de SendGrid esta incompleta.")
        else:
            muestra = elegibles[elegibles["cliente_label"].isin(clientes_prueba_lote)].copy().head(3)
            resultados_prueba_lote = []
            progress = st.progress(0)
            total_iteraciones = max(len(muestra) * len(correos_prueba_lote), 1)
            avance = 0
            for _, fila in muestra.iterrows():
                df_cliente = df_base[df_base["cliente_key"] == fila["cliente_key"]].copy()
                pdf_bytes = crear_pdf_cliente(df_cliente, float(fila["saldo_vencido"]))
                nombre_pdf = f"PRUEBA_{normalizar_nombre(str(fila['nombrecliente'])).replace(' ', '_')}.pdf"
                asunto = f"[PRUEBA] {construir_texto_asunto(str(fila['nombrecliente']), estrategia, float(fila['saldo_vencido']))}"
                cuerpo_html = plantilla_correo_conciliacion(fila, estrategia)
                cuerpo_txt = cuerpo_texto_plano(fila, estrategia)
                for correo_destino in correos_prueba_lote:
                    ok, detalle = enviar_con_sendgrid(
                        api_key=api_key,
                        from_email=from_email,
                        from_name=from_name,
                        to_email=correo_destino,
                        subject=asunto,
                        html_content=cuerpo_html,
                        plain_content=cuerpo_txt,
                        attachment_bytes=pdf_bytes,
                        attachment_name=nombre_pdf,
                        cliente_row=fila,
                    )
                    resultados_prueba_lote.append({
                        "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Campana": f"PRUEBA - {nombre_campana}",
                        "Modo": "Prueba",
                        "Cliente": fila["nombrecliente"],
                        "Destino": correo_destino,
                        "Correo Cliente": fila["correo"],
                        "Estado Correo": fila["estado_correo"],
                        "Saldo Vencido": float(fila["saldo_vencido"]),
                        "Resultado": "Enviado" if ok else "Error",
                        "Detalle": detalle,
                        "Vendedor": fila["nomvendedor"],
                        "Zona": fila["zona"],
                        "Estrategia": estrategia,
                    })
                    avance += 1
                    progress.progress(avance / total_iteraciones)
            df_prueba_lote = pd.DataFrame(resultados_prueba_lote)
            guardar_historial_envios(df_prueba_lote)
            st.session_state["historial_envio_conciliacion"] = cargar_historial_envios().to_dict("records")
            st.success("Prueba de lote finalizada. Revisa abajo el resultado y luego procede con el envio real.")
            st.dataframe(df_prueba_lote, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    if not incluir_compartidos:
        elegibles = elegibles[elegibles["estado_correo"] == "Listo"].copy()

    st.info(
        f"El lote actual enviaria {len(elegibles)} correos. Los clientes sin correo o con correo invalido quedan fuera y se reportan para depuracion."
    )

    if not bloqueados.empty:
        with st.expander("Ver clientes bloqueados antes del envio"):
            st.dataframe(
                bloqueados[["nombrecliente", "correo_original", "estado_correo", "nomvendedor", "zona", "saldo_vencido"]]
                .rename(columns={
                    "nombrecliente": "Cliente",
                    "correo_original": "Correo Base",
                    "estado_correo": "Estado",
                    "nomvendedor": "Vendedor",
                    "zona": "Zona",
                    "saldo_vencido": "Saldo Vencido",
                })
                .style.format({"Saldo Vencido": "${:,.0f}"}),
                use_container_width=True,
                hide_index=True,
            )

    if st.button("Ejecutar envio masivo", type="primary", disabled=elegibles.empty):
        api_key = st.secrets["sendgrid"].get("api_key", "")
        from_email = st.secrets["sendgrid"].get("from_email", "")
        from_name = st.secrets["sendgrid"].get("from_name", "Ferreinox S.A.S. BIC")
        if not api_key or not from_email:
            st.error("La configuracion de SendGrid esta incompleta.")
            return

        progress = st.progress(0)
        estado = st.empty()
        resultados = []

        for idx, (_, fila) in enumerate(elegibles.iterrows(), start=1):
            estado.info(f"Enviando {idx} de {len(elegibles)}: {fila['nombrecliente']}")
            df_cliente = df_base[df_base["cliente_key"] == fila["cliente_key"]].copy()
            pdf_bytes = crear_pdf_cliente(df_cliente, float(fila["saldo_vencido"]))
            asunto = construir_texto_asunto(str(fila["nombrecliente"]), estrategia, float(fila["saldo_vencido"]))
            cuerpo_html = plantilla_correo_conciliacion(fila, estrategia)
            cuerpo_txt = cuerpo_texto_plano(fila, estrategia)
            nombre_pdf = f"Estado_Cuenta_{normalizar_nombre(str(fila['nombrecliente'])).replace(' ', '_')}.pdf"

            ok, detalle = enviar_con_sendgrid(
                api_key=api_key,
                from_email=from_email,
                from_name=from_name,
                to_email=str(fila["correo"]),
                subject=asunto,
                html_content=cuerpo_html,
                plain_content=cuerpo_txt,
                attachment_bytes=pdf_bytes,
                attachment_name=nombre_pdf,
                cliente_row=fila,
            )

            resultados.append({
                "Fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Campana": nombre_campana,
                "Modo": "Produccion",
                "Cliente": fila["nombrecliente"],
                "Destino": fila["correo"],
                "Correo Cliente": fila["correo"],
                "Estado Correo": fila["estado_correo"],
                "Saldo Vencido": float(fila["saldo_vencido"]),
                "Resultado": "Enviado" if ok else "Error",
                "Detalle": detalle,
                "Vendedor": fila["nomvendedor"],
                "Zona": fila["zona"],
                "Estrategia": estrategia,
            })
            progress.progress(idx / len(elegibles))
            if pausa_ms:
                time.sleep(pausa_ms / 1000)

        st.session_state["reporte_envio_conciliacion"] = resultados
        guardar_historial_envios(pd.DataFrame(resultados))
        st.session_state["historial_envio_conciliacion"] = cargar_historial_envios().to_dict("records")
        total_ok = sum(1 for item in resultados if item["Resultado"] == "Enviado")
        total_error = len(resultados) - total_ok
        estado.empty()
        if total_error == 0:
            st.success(f"Envio finalizado. {total_ok} correos enviados correctamente.")
        else:
            st.warning(f"Envio finalizado. Exitosos: {total_ok}. Con error: {total_error}.")

    resultados = pd.DataFrame(st.session_state.get("reporte_envio_conciliacion", []))
    if not resultados.empty:
        st.markdown('<div class="section-title">Reporte del ultimo despacho</div>', unsafe_allow_html=True)
        st.dataframe(resultados.style.format({"Saldo Vencido": "${:,.0f}"}), use_container_width=True, hide_index=True)
        st.download_button(
            "Descargar reporte del envio",
            data=dataframe_a_excel({"Reporte Envio": resultados}),
            file_name=f"Reporte_Envio_Conciliacion_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def render_historial_tab():
    historial = pd.DataFrame(st.session_state.get("historial_envio_conciliacion", []))
    st.markdown('<div class="section-title">Historial de campanas y pruebas</div>', unsafe_allow_html=True)
    if historial.empty:
        st.info("Aun no hay historial guardado en este centro de control.")
        return

    campanas = ["TODAS"] + sorted(historial["Campana"].dropna().astype(str).unique().tolist())
    modos = ["TODOS"] + sorted(historial["Modo"].dropna().astype(str).unique().tolist())
    resultados = ["TODOS"] + sorted(historial["Resultado"].dropna().astype(str).unique().tolist())

    col1, col2, col3 = st.columns(3)
    with col1:
        campana_sel = st.selectbox("Campana", campanas, key="historial_campana_sel")
    with col2:
        modo_sel = st.selectbox("Modo", modos, key="historial_modo_sel")
    with col3:
        resultado_sel = st.selectbox("Resultado", resultados, key="historial_resultado_sel")

    vista = historial.copy()
    if campana_sel != "TODAS":
        vista = vista[vista["Campana"] == campana_sel]
    if modo_sel != "TODOS":
        vista = vista[vista["Modo"] == modo_sel]
    if resultado_sel != "TODOS":
        vista = vista[vista["Resultado"] == resultado_sel]

    k1, k2, k3 = st.columns(3)
    k1.metric("Registros", len(vista))
    k2.metric("Enviados", int(vista["Resultado"].eq("Enviado").sum()))
    k3.metric("Errores", int(vista["Resultado"].eq("Error").sum()))

    st.dataframe(vista.style.format({"Saldo Vencido": "${:,.0f}"}), use_container_width=True, hide_index=True)
    st.download_button(
        "Descargar historial completo",
        data=dataframe_a_excel({"Historial": historial, "Vista filtrada": vista}),
        file_name=f"Historial_Conciliacion_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def render_calidad_tab(reporte_calidad: pd.DataFrame, df_view: pd.DataFrame):
    st.markdown('<div class="section-title">Alertas de calidad de datos</div>', unsafe_allow_html=True)

    col1, col2 = st.columns([1, 2])
    with col1:
        conteo = df_view["estado_correo"].value_counts().reset_index()
        if not conteo.empty:
            fig = px.bar(
                conteo,
                x="estado_correo",
                y="count",
                color="estado_correo",
                color_discrete_map={
                    "Listo": COLOR_OK,
                    "Correo compartido": COLOR_ACCION,
                    "Correo invalido": COLOR_TERCIARIO,
                    "Sin correo": COLOR_PRIMARIO,
                },
            )
            fig.update_layout(margin=dict(t=20, l=0, r=0, b=0), showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
    with col2:
        st.markdown(
            """
            <div class="panel-card">
                <div class="section-title">Como usar esta pestaña</div>
                <div class="metric-hint">
                    Aqui ves exactamente quienes no tienen correo, quienes lo tienen mal escrito y quienes comparten un mismo correo.
                    Esta pagina no modifica tu base original, pero si te entrega la lista precisa para depurar Dropbox o tu origen maestro y mantener los siguientes cortes al dia.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.dataframe(
        reporte_calidad.style.format({"Saldo Vencido": "${:,.0f}", "Saldo Total": "${:,.0f}"}),
        use_container_width=True,
        hide_index=True,
    )
    st.download_button(
        "Descargar reporte de calidad de correos",
        data=dataframe_a_excel({"Calidad Correos": reporte_calidad}),
        file_name=f"Calidad_Correos_Ferreinox_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def main():
    inicializar_estado()
    if not st.session_state.get("authentication_status", False):
        render_login()

    df_base, status = cargar_cartera_dropbox()
    render_sidebar(status)
    if df_base is None or df_base.empty:
        st.error("No fue posible cargar la cartera. Revisa Dropbox y vuelve a intentar.")
        st.stop()

    resumen = construir_resumen_clientes(df_base)
    if resumen.empty:
        st.warning("No hay clientes disponibles en la cartera actual.")
        st.stop()

    df_view = aplicar_filtros(resumen)

    st.markdown(
        f"""
        <div class="hero-card">
            <h1>Centro de Control de Conciliacion de Cartera</h1>
            <p>Pagina nueva, aislada del resto de la app, conectada a tu cartera actual para validar correos, preparar lotes y despachar estados de cuenta PDF con SendGrid.</p>
            <div class="mini-note">Clientes filtrados: {len(df_view)} | Fecha de corte: {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    render_metricas(df_view)
    seleccionados = seleccionar_clientes(df_view)
    reporte_calidad = preparar_reporte_calidad(df_view)

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Centro de control",
        "Vista previa",
        "Envio masivo",
        "Calidad de datos",
        "Historial",
    ])

    with tab1:
        render_control_tab(df_view, seleccionados)

    with tab2:
        render_preview_tab(df_base, df_view, seleccionados)

    with tab3:
        render_envio_tab(df_base, seleccionados)

    with tab4:
        render_calidad_tab(reporte_calidad, df_view)

    with tab5:
        render_historial_tab()


if __name__ == "__main__":
    main()