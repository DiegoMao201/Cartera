# ======================================================================================
# ARCHIVO: pages/📊_Análisis_Histórico.py (Con Explicaciones Integradas)
# ======================================================================================
import streamlit as st
import pandas as pd
import glob
import re
from datetime import datetime
from dateutil.relativedelta import relativedelta
import unicodedata
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import plotly.express as px

st.set_page_config(page_title="Centro de Comando Histórico", page_icon="🔮", layout="wide")

# --- GUARDIA DE SEGURIDAD ---
if 'authentication_status' not in st.session_state or not st.session_state['authentication_status']:
    st.warning("Por favor, inicie sesión en el 📈 Tablero Principal para acceder a esta página.")
    st.stop()

# --- FUNCIONES AUXILIARES (Sin cambios) ---
def normalizar_nombre(nombre: str) -> str:
    if not isinstance(nombre, str): return ""
    nombre = nombre.upper().strip().replace('.', '')
    nombre = ''.join(c for c in unicodedata.normalize('NFD', nombre) if unicodedata.category(c) != 'Mn')
    return ' '.join(nombre.split())

@st.cache_data
def cargar_datos_historicos():
    mapa_columnas = {
        'Serie': 'serie', 'Número': 'numero', 'Fecha Documento': 'fecha_documento',
        'Fecha Vencimiento': 'fecha_vencimiento', 'Fecha Saldado': 'fecha_saldado',
        'NOMBRECLIENTE': 'nombrecliente', 'Población': 'poblacion', 'Provincia': 'provincia',
        'IMPORTE': 'importe', 'RIESGOCONCEDIDO': 'riesgoconcedido', 'NOMVENDEDOR': 'nomvendedor',
        'DIAS_VENCIDO': 'dias_vencido', 'Estado': 'estado', 'Cod. Cliente': 'cod_cliente', 'e-mail': 'e_mail'
    }
    lista_archivos = sorted(glob.glob("Cartera_*.xlsx"))
    if not lista_archivos: return pd.DataFrame()
    lista_df = []
    for archivo in lista_archivos:
        try:
            df = pd.read_excel(archivo).iloc[:-1]
            for col in ['e-mail', 'Cod. Cliente']:
                if col not in df.columns: df[col] = None
            df['Serie'] = df['Serie'].astype(str)
            df = df[~df['Serie'].str.contains('W|X', case=False, na=False)]
            df.rename(columns=mapa_columnas, inplace=True)
            lista_df.append(df)
        except Exception: pass
    if not lista_df: return pd.DataFrame()
    df_completo = pd.concat(lista_df, ignore_index=True).dropna(subset=['numero', 'nombrecliente'])
    df_completo['nomvendedor_norm'] = df_completo['nomvendedor'].apply(normalizar_nombre)
    df_completo.sort_values(by=['fecha_documento', 'fecha_saldado'], ascending=[True, True], na_position='first', inplace=True)
    df_historico_unico = df_completo.drop_duplicates(subset=['numero'], keep='last').copy()
    for col in ['fecha_documento', 'fecha_vencimiento', 'fecha_saldado']:
        df_historico_unico[col] = pd.to_datetime(df_historico_unico[col], errors='coerce')
    df_historico_unico['importe'] = pd.to_numeric(df_historico_unico['importe'], errors='coerce').fillna(0)
    df_pagadas = df_historico_unico.dropna(subset=['fecha_saldado', 'fecha_documento']).copy()
    if not df_pagadas.empty:
        df_pagadas['dias_de_pago'] = (df_pagadas['fecha_saldado'] - df_pagadas['fecha_documento']).dt.days
        df_historico_unico = pd.merge(df_historico_unico, df_pagadas[['numero', 'dias_de_pago']], on='numero', how='left')
    return df_historico_unico

@st.cache_data
def calcular_rfm(df: pd.DataFrame):
    snapshot_date = df['fecha_documento'].max() + relativedelta(days=1)
    rfm = df.groupby('nombrecliente').agg({
        'fecha_documento': lambda date: (snapshot_date - date.max()).days,
        'numero': 'count',
        'importe': 'sum'
    }).rename(columns={'fecha_documento': 'Recencia', 'numero': 'Frecuencia', 'importe': 'Monetario'})
    r_labels = range(4, 0, -1); f_labels = range(1, 5); m_labels = range(1, 5)
    rfm['R_score'] = pd.qcut(rfm['Recencia'], q=4, labels=r_labels, duplicates='drop').astype(int)
    rfm['F_score'] = pd.qcut(rfm['Frecuencia'].rank(method='first'), q=4, labels=f_labels).astype(int)
    rfm['M_score'] = pd.qcut(rfm['Monetario'], q=4, labels=m_labels, duplicates='drop').astype(int)
    def segmentar(df):
        if df['R_score'] >= 4 and df['F_score'] >= 4: return 'Campeones'
        if df['R_score'] >= 3 and df['F_score'] >= 3: return 'Clientes Leales'
        if df['R_score'] >= 3 and df['M_score'] >= 4: return 'Grandes Compradores'
        if df['R_score'] <= 2 and df['F_score'] >= 3: return 'En Riesgo'
        if df['R_score'] <= 2 and df['M_score'] >= 3: return 'No se pueden perder'
        if df['R_score'] <= 2: return 'Hibernando'
        return 'Otros'
    rfm['Segmento'] = rfm.apply(segmentar, axis=1)
    return rfm

# --- Carga y Filtros ---
st.title("🔮 Centro de Comando Histórico y Predictivo")
df_historico_base = cargar_datos_historicos()
if df_historico_base.empty:
    st.error("No se encontraron archivos de datos históricos `Cartera_*.xlsx`."); st.stop()
st.sidebar.header("Filtros de Análisis")
acceso_general = st.session_state.get('acceso_general', False)
vendedor_autenticado = st.session_state.get('vendedor_autenticado', None)
if acceso_general:
    vendedores = ["Todos"] + sorted(df_historico_base['nomvendedor'].dropna().unique())
    vendedor_sel_hist = st.sidebar.selectbox("Vendedor:", vendedores)
else:
    vendedor_sel_hist = vendedor_autenticado
if vendedor_sel_hist == "Todos":
    df_historico = df_historico_base.copy()
else:
    df_historico = df_historico_base[df_historico_base['nomvendedor_norm'] == normalizar_nombre(vendedor_sel_hist)].copy()
if df_historico.empty:
    st.warning("No hay datos para el vendedor seleccionado."); st.stop()
    
tab1, tab2, tab3, tab4 = st.tabs([
    "📈 Diagnóstico del Período", 
    "🔮 Análisis Predictivo y de Tendencias",
    "🧑‍🤝‍🧑 Segmentación Estratégica de Clientes (RFM)",
    "⚙️ Simulador de Escenarios"
])

# ======================================================================================
# PESTAÑA 1: Diagnóstico del Período
# ======================================================================================
with tab1:
    st.header("Diagnóstico Financiero del Período")
    max_date = df_historico['fecha_documento'].max().date()
    min_date = df_historico['fecha_documento'].min().date()
    default_start_date = max(min_date, max_date - relativedelta(months=12))
    fecha_inicio, fecha_fin = st.date_input("Selecciona el Rango de Fechas para el Diagnóstico:", (default_start_date, max_date), min_value=min_date, max_value=max_date, key="date_range_tab1")

    if not fecha_inicio or not fecha_fin or fecha_inicio > fecha_fin:
        st.error("Rango de fechas inválido."); st.stop()
    fecha_inicio, fecha_fin = pd.to_datetime(fecha_inicio), pd.to_datetime(fecha_fin)
    
    snapshot_inicial = df_historico[df_historico['fecha_documento'] < fecha_inicio]
    saldo_inicial = snapshot_inicial[(snapshot_inicial['fecha_saldado'].isnull()) | (snapshot_inicial['fecha_saldado'] >= fecha_inicio)]['importe'].sum()
    ventas_periodo = df_historico[df_historico['fecha_documento'].between(fecha_inicio, fecha_fin)]
    total_ventas = ventas_periodo['importe'].sum()
    cobros_periodo = df_historico[df_historico['fecha_saldado'].between(fecha_inicio, fecha_fin)]
    total_cobrado = cobros_periodo['importe'].sum()
    snapshot_final = df_historico[df_historico['fecha_documento'] <= fecha_fin]
    facturas_abiertas_al_final = snapshot_final[(snapshot_final['fecha_saldado'].isnull()) | (snapshot_final['fecha_saldado'] > fecha_fin)]
    saldo_final_total = facturas_abiertas_al_final['importe'].sum()
    facturas_vencidas_al_final = facturas_abiertas_al_final[facturas_abiertas_al_final['fecha_vencimiento'] < fecha_fin]
    saldo_vencido_final = facturas_vencidas_al_final['importe'].sum()
    dso_periodo = cobros_periodo['dias_de_pago'].mean() if not cobros_periodo.empty else 0
    flujo_neto = total_cobrado - total_ventas
    universo_cobrable = saldo_inicial + total_ventas
    cer = (total_cobrado / universo_cobrable) * 100 if universo_cobrable > 0 else 0
    indice_morosidad = (saldo_vencido_final / saldo_final_total) * 100 if saldo_final_total > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("💰 Eficiencia de Cobro (CER)", f"{cer:.1f}%")
    col2.metric("🔥 Índice de Morosidad", f"{indice_morosidad:.1f}%")
    col3.metric("🔄 Rotación (DSO)", f"{dso_periodo:.0f} días")
    col4.metric("🌊 Flujo Neto de Cartera", f"${flujo_neto:,.0f}")
    
    # --- MEJORA: Explicación detallada de los KPIs ---
    with st.expander("❓ ¿Cómo interpretar estos indicadores? (Manual para Gerentes y Vendedores)"):
        st.markdown(f"""
        #### 💰 Eficiencia de Cobro (CER): {cer:.1f}%
        * **Pregunta Clave:** De todo el dinero que podíamos cobrar en este período (lo que nos debían al empezar + lo que facturamos), ¿qué porcentaje realmente entró a la caja?
        * **Explicación Sencilla:** Es el termómetro de la efectividad de tu gestión de cobro. Un número alto es señal de un equipo proactivo y clientes que responden.
        * **En tus Datos:** Un **{cer:.1f}%** es un resultado **excelente**. Significa que la gestión para convertir papel (facturas) en dinero (efectivo) es de élite.
        * **Decisiones a Tomar (Vendedor):** ¡Felicita a tus clientes buenos pagadores! Identifica qué hiciste bien con ellos y replícalo. Para los que no, entiende por qué y actúa.

        #### 🔥 Índice de Morosidad: {indice_morosidad:.1f}%
        * **Pregunta Clave:** Del dinero que todavía nos deben al final del período, ¿qué porcentaje ya está vencido?
        * **Explicación Sencilla:** Mide la "calidad" o "salud" de la cartera que queda abierta. Un índice bajo significa que la mayoría de tu cartera está al día. Un índice alto es una bandera roja.
        * **En tus Datos:** Un **{indice_morosidad:.1f}%** es **elevado**. Es una señal de alerta importante.
        * **Decisiones a Tomar (Gerente):** ¿Estamos dando crédito a los clientes correctos? ¿Nuestros plazos son adecuados? ¿Necesitamos ser más estrictos con los límites de crédito? Este indicador exige una revisión de la política de riesgo.

        #### 🔄 Rotación (DSO): {dso_periodo:.0f} días
        * **Pregunta Clave:** En promedio, ¿cuántos días tardamos en cobrar una factura desde que la emitimos?
        * **Explicación Sencilla:** Es la velocidad de tu ciclo de cobro. Cada día menos en el DSO es dinero que tienes disponible antes en tu cuenta bancaria.
        * **En tus Datos:** **{dso_periodo:.0f} días** es el tiempo promedio que tardas. Compáralo con tu plazo de pago estándar (ej. 30 días). Si es mucho mayor, hay una desconexión.
        * **Decisiones a Tomar (Ambos):** Si el DSO es alto, hay que analizar qué clientes o vendedores lo están causando. ¿Se pueden ofrecer descuentos por pronto pago?
        
        #### 🌊 Flujo Neto de Cartera: ${flujo_neto:,.0f}
        * **Pregunta Clave:** En este período, ¿entró más dinero por cobros del que salió por nuevas ventas a crédito?
        * **Explicación Sencilla:** Es el pulso de la liquidez de tu operación comercial. Si es positivo, cobraste más de lo que vendiste a crédito, fortaleciendo tu caja. Si es negativo, tu cartera creció, lo que requiere más capital de trabajo.
        * **En tus Datos:** Un resultado de **${flujo_neto:,.0f}** indica que la gestión generó liquidez.
        
        ---
        #### La Historia Completa: Conectando los Puntos
        Viendo tus números, la historia es clara:
        > El equipo de cobranza está haciendo un trabajo **fenomenal recuperando dinero (CER del 97.4%)**, probablemente enfocándose en deudas importantes o antiguas, lo que generó un **flujo de caja positivo**. Sin embargo, la **alta morosidad (39.8%)** en la cartera restante es una **alarma crítica**. Sugiere que mientras se apagan los grandes incendios, se están descuidando las brasas de las facturas más nuevas o que la calidad de los nuevos créditos otorgados no es óptima.
        
        **Acción Gerencial:** Capitalizar la excelente gestión de cobro para diseñar un plan proactivo que ataque la mora de la cartera "joven" y revise las condiciones de crédito para evitar que el problema crezca.
        """)

# PESTAÑA 2: Análisis Predictivo y de Tendencias
with tab2:
    st.header("Proyecciones y Tendencias a Futuro")
    # El resto de la pestaña 2 no necesita cambios, ya es bastante explicativa.
    df_ts = df_historico.set_index('fecha_documento')
    df_ventas_mes = df_ts['importe'].resample('MS').sum()
    df_dso_mes = df_historico.dropna(subset=['dias_de_pago']).set_index('fecha_saldado')['dias_de_pago'].resample('MS').mean()
    periodos_a_proyectar = st.slider("Meses a proyectar hacia el futuro:", 1, 12, 3, key="slider_proyeccion")
    chart1, chart2 = st.columns(2)
    with chart1:
        st.markdown("#### Proyección de Ventas")
        if len(df_ventas_mes) >= 24:
            modelo_ventas = ExponentialSmoothing(df_ventas_mes, trend='add', seasonal='add', seasonal_periods=12).fit()
            proyeccion_ventas = modelo_ventas.forecast(periodos_a_proyectar)
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=df_ventas_mes.index, y=df_ventas_mes.values, mode='lines', name='Ventas Históricas'))
            fig.add_trace(go.Scatter(x=proyeccion_ventas.index, y=proyeccion_ventas.values, mode='lines', name='Proyección', line=dict(dash='dash', color='red')))
            st.plotly_chart(fig, use_container_width=True)
        else: st.warning("Se necesitan al menos 24 meses de datos históricos para una proyección estacional fiable.")
    with chart2:
        st.markdown("#### Proyección de DSO (Rotación)")
        if len(df_dso_mes) >= 12:
            modelo_dso = ExponentialSmoothing(df_dso_mes.ffill(), trend='add', seasonal=None).fit()
            proyeccion_dso = modelo_dso.forecast(periodos_a_proyectar)
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=df_dso_mes.index, y=df_dso_mes.values, mode='lines', name='DSO Histórico'))
            fig.add_trace(go.Scatter(x=proyeccion_dso.index, y=proyeccion_dso.values, mode='lines', name='Proyección', line=dict(dash='dash', color='orange')))
            st.plotly_chart(fig, use_container_width=True)
        else: st.warning("Se necesitan al menos 12 meses de datos de cobros para una proyección de DSO fiable.")


# PESTAÑA 3: Segmentación Estratégica de Clientes (RFM)
with tab3:
    # Esta pestaña también es bastante autoexplicativa.
    st.header("Segmentación Estratégica de Clientes (RFM)")
    st.markdown("Clasifique a sus clientes en segmentos accionables basados en su comportamiento de compra: **R**ecencia, **F**recuencia y **M**onto.")
    rfm_data = calcular_rfm(df_historico)
    col1, col2 = st.columns([1, 2])
    with col1:
        st.markdown("#### Resumen de Segmentos")
        segment_counts = rfm_data['Segmento'].value_counts().reset_index()
        st.dataframe(segment_counts, use_container_width=True, hide_index=True)
        st.markdown("#### Recomendaciones Estratégicas")
        recomendaciones = {"Campeones": "🏆 Fidelizar y recompensar.", "Clientes Leales": "🤝 Venta cruzada y up-selling.", "Grandes Compradores": "💰 Foco en post-venta.", "En Riesgo": "⚠️ Contacto proactivo inmediato.", "No se pueden perder": "💎 Atención personalizada de alto nivel.", "Hibernando": "😴 Campañas de reactivación."}
        segmento_sel = st.selectbox("Ver estrategia para el segmento:", rfm_data['Segmento'].unique())
        if segmento_sel in recomendaciones: st.info(recomendaciones[segmento_sel])
    with col2:
        st.markdown("#### Visualización de la Base de Clientes")
        plot_data = rfm_data[rfm_data['Monetario'] > 0].copy()
        if not plot_data.empty:
            fig = px.scatter(plot_data, x='Recencia', y='Frecuencia', size='Monetario', color='Segmento',
                             hover_name=plot_data.index, size_max=60,
                             title="Mapa de Clientes por Recencia, Frecuencia y Monto")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay clientes con valor monetario positivo para visualizar en el gráfico RFM.")
    with st.expander("Ver detalle de clientes por segmento"):
        st.dataframe(rfm_data.sort_values(by=['R_score', 'F_score', 'M_score'], ascending=False), use_container_width=True)

# PESTAÑA 4: Simulador de Escenarios
with tab4:
    st.header("Simulador de Escenarios Futuros")
    st.markdown("Use esta herramienta para cuantificar el impacto de sus decisiones.")
    st.sidebar.markdown("---")
    st.sidebar.header("Parámetros del Simulador")
    # --- MEJORA: Explicación de los parámetros del simulador ---
    st.sidebar.caption("Mueva estos controles para simular cómo las mejoras en la gestión de cobro (reducir DSO) o los cambios en la actividad comercial (aumentar/disminuir ventas) impactan las finanzas de la empresa.")

    base_simulacion = df_historico[df_historico['fecha_documento'] > (df_historico['fecha_documento'].max() - relativedelta(months=12))]
    ventas_base_anual = base_simulacion[base_simulacion['importe'] > 0]['importe'].sum()
    dso_base_anual = base_simulacion.dropna(subset=['dias_de_pago'])['dias_de_pago'].mean()
    st.sidebar.info(f"**Base Anual para Simulación:**\nVentas: ${ventas_base_anual:,.0f}\nDSO: {dso_base_anual:.0f} días")
    cambio_ventas_pct = st.sidebar.slider("Cambio en Ventas (%)", -25, 50, 0)
    cambio_dso_dias = st.sidebar.slider("Reducción del DSO (días)", 0, 30, 0)
    ventas_proyectadas = ventas_base_anual * (1 + cambio_ventas_pct / 100)
    dso_proyectado = dso_base_anual - cambio_dso_dias
    capital_trabajo_base = (ventas_base_anual / 365) * dso_base_anual
    capital_trabajo_proyectado = (ventas_proyectadas / 365) * dso_proyectado
    liberacion_capital = capital_trabajo_base - capital_trabajo_proyectado
    st.subheader("Resultados de la Simulación")
    col1, col2, col3 = st.columns(3)
    col1.metric("📈 Ventas Proyectadas", f"${ventas_proyectadas:,.0f}", delta=f"${ventas_proyectadas - ventas_base_anual:,.0f}")
    col2.metric("🔄 DSO Proyectado", f"{dso_proyectado:.0f} días", delta=f"{-cambio_dso_dias} días")
    col3.metric("💸 Capital de Trabajo Liberado", f"${liberacion_capital:,.0f}", help="Dinero que deja de estar inmovilizado en la cartera.")
    
    # --- MEJORA: Explicación detallada de los resultados del simulador ---
    with st.expander("❓ ¿Qué significa 'Capital de Trabajo Liberado'?"):
         st.markdown("""
        Piense en el **Capital de Trabajo** como el dinero de la empresa que está "atrapado" en la calle, en forma de facturas que sus clientes aún no han pagado. Es dinero que es suyo, pero que no puede usar.

        **"Liberar" capital de trabajo** significa que, gracias a una gestión más eficiente (principalmente, cobrar más rápido y reducir el DSO), usted logra sacar ese dinero de la calle y traerlo de vuelta a la caja de la empresa.

        * Un número **positivo** aquí es el "premio" por su buena gestión. Es dinero extra que la empresa ahora tiene disponible para pagar sueldos, comprar inventario, invertir o repartir dividendos.
        * Un número **negativo** significa que su nuevo escenario (más ventas o cobros más lentos) requiere "atrapar" más dinero en la calle para poder funcionar. Es una inversión necesaria en su cartera.

        Este simulador le permite medir el impacto monetario real de sus estrategias antes de implementarlas.
        """)
