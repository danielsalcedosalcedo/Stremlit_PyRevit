# app.py – Aplicación Streamlit para Revisión de Curvas S de Proyecto
# Vinculación ITEM (Excel) ↔ Assembly Code (Revit/PyRevit)
# Desarrollado por Daniel Salcedo – Coordinador BIM

import streamlit as st
import pandas as pd
import io as _io
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))

import matplotlib.pyplot as plt

from logic.loader import (
    load_excel_sheet_names,
    load_datos_sheet,
    load_avance_semanal,
    load_avance_diario,
    load_pyrevit_json,
    load_pyrevit_csv,
    validar_columnas_pyrevit,
    resumen_pyrevit,
)
from logic.processor import (
    calcular_curva_s_semanal,
    calcular_curva_s_diaria,
    calcular_curva_s_real_pyrevit,
    vincular_item_pyrevit,
    tabla_avance_por_item,
    avance_por_nivel_pyrevit,
    obtener_lista_actividades,
    kpis_semanal,
)
from logic.charts import (
    grafico_curva_s_semanal_plotly,
    grafico_curva_s_diaria_plotly,
    grafico_barras_actividades_plotly,
    grafico_avance_por_item_plotly,
    grafico_avance_por_nivel_plotly,
    grafico_curva_real_pyrevit_plotly,
)

# ═══════════════════════════════════════════════════════════════════════
# FUNCIONES CACHEADAS PARA OPTIMIZAR STREAMLIT
# ═══════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def get_curva_s_semanal(datos_sem, actividad=None):
    return calcular_curva_s_semanal(datos_sem, actividad=actividad)

@st.cache_data(show_spinner=False)
def get_curva_s_diaria(datos_dia, actividad=None):
    return calcular_curva_s_diaria(datos_dia, actividad=actividad)

@st.cache_data(show_spinner=False)
def get_curva_s_real_pyrevit(df_pyrevit, frecuencia="W", fecha_corte=None):
    return calcular_curva_s_real_pyrevit(df_pyrevit, frecuencia=frecuencia, fecha_corte=fecha_corte)

@st.cache_data(show_spinner=False)
def get_vinculado(df_datos, df_pyrevit):
    return vincular_item_pyrevit(df_datos, df_pyrevit)


def crear_registros_avance_manual(registros):
    """Convierte avances manuales acumulados en registros compatibles con PyRevit."""
    filas = []
    for registro in registros:
        total = float(registro["cantidad_total"])
        ejecutado = float(registro["cantidad_ejecutada"])
        base = {
            "assembly_code": str(registro["ITEM"]).strip(),
            "assembly_description": registro.get("actividad", ""),
            "unidad": registro.get("unidad", "unid"),
        }
        if ejecutado > 0:
            filas.append({
                **base,
                "cantidad_total": ejecutado,
                "ejecutado": True,
                "fecha": pd.to_datetime(registro["fecha"]),
                "origen": "Manual",
            })
        pendiente = max(total - ejecutado, 0.0)
        if pendiente > 0:
            filas.append({
                **base,
                "cantidad_total": pendiente,
                "ejecutado": False,
                "fecha": pd.NaT,
                "origen": "Manual",
            })
    return pd.DataFrame(filas)


# ═══════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE PÁGINA
# ═══════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Echeverría Izquierdo – Control de Avance BIM",
    page_icon="🔷",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap');

:root {
    --ei-blue:    #00A0E3;
    --ei-blue-dark:#0077B6;
    --ei-navy:    #0A2A62;
    --ei-navy-2:  #123A7A;
    --bg:         #FFFFFF;
    --panel:      #F7F9FC;
    --card:       #FFFFFF;
    --border:     #E6E9EE;
    --green:      #2E9E6B;
    --amber:      #D4A017;
    --red:        #C0392B;
    --text:       #2C3338;
    --sub:        #6B7280;
    --muted:      #9AA3AE;
}

html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"], .stApp {
    background-color: #FFFFFF !important;
    color: var(--text) !important;
    font-family: 'Montserrat', system-ui, sans-serif !important;
}

[data-testid="stHeader"],
#MainMenu, header[data-testid="stHeader"],
.stDeployButton, div[data-testid="stToolbar"],
div[data-testid="stAppToolbar"], .stAppToolbar,
footer, [data-testid="stStatusWidget"],
[data-testid="stDecoration"] {
    display: none !important;
    visibility: hidden !important;
    height: 0 !important;
}

.block-container {
    padding-top: 0.4rem !important;
    padding-bottom: 2rem !important;
    max-width: 1280px;
}

[data-testid="stSidebar"] {
    background-color: #FFFFFF !important;
    border-right: 1px solid var(--border);
}
[data-testid="stSidebar"] * {
    font-family: 'Montserrat', system-ui, sans-serif !important;
}

[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid var(--border);
    border-radius: 0;
    padding: 14px 16px;
    box-shadow: none;
}
[data-testid="stMetricLabel"]  { color: var(--sub) !important; font-size:.72rem; font-weight:600; letter-spacing:.08em; text-transform:uppercase; }
[data-testid="stMetricValue"]  { color: var(--ei-navy) !important; font-size:1.45rem; font-weight:700; }

[data-testid="stExpander"] {
    background: #FFFFFF;
    border: 1px solid var(--border) !important;
    border-radius: 0;
}

[data-testid="stFileUploader"] {
    background: #F7F9FC;
    border: 1px dashed #C5CDD6;
    border-radius: 0;
}
[data-testid="stFileUploader"]:hover { border-color: var(--ei-blue); }

button[kind="primary"] {
    background: var(--ei-blue) !important;
    color: #FFFFFF !important;
    border: none;
    border-radius: 0;
    font-weight: 600;
    letter-spacing: .04em;
    text-transform: uppercase;
    font-size: .78rem;
}
button[kind="primary"]:hover { background: var(--ei-blue-dark) !important; }

.ei-topbar {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 16px;
    padding: 6px 4px 12px 4px;
    border-bottom: 1px solid #EEF1F4;
    background: #FFFFFF;
    flex-wrap: wrap;
}
.ei-brand {
    display: flex;
    align-items: center;
    gap: 12px;
    min-width: 240px;
    text-decoration: none;
}
.ei-mark { width: 42px; height: 42px; flex-shrink: 0; }
.ei-wordmark {
    display: flex;
    flex-direction: column;
    line-height: 1.05;
}
.ei-wordmark span {
    font-family: 'Montserrat', sans-serif;
    font-weight: 800;
    color: var(--ei-blue);
    letter-spacing: .04em;
    font-size: 1.05rem;
}
.ei-nav {
    display: flex;
    flex-wrap: wrap;
    align-items: center;
    justify-content: flex-end;
    gap: 4px 16px;
    flex: 1;
}
.ei-nav a {
    font-size: .72rem;
    color: #4B5563;
    text-decoration: none;
    font-weight: 500;
    white-space: nowrap;
}
.ei-nav a:hover, .ei-nav a.active {
    color: var(--ei-blue);
}
.ei-nav a.active { font-weight: 700; }

.ei-banner {
    position: relative;
    margin: 0 -1rem 1.5rem -1rem;
    min-height: 128px;
    background:
        linear-gradient(180deg, rgba(10,42,98,.58) 0%, rgba(10,42,98,.78) 100%),
        url("https://images.unsplash.com/photo-1477959858617-67f85cf4f1df?auto=format&fit=crop&w=1800&q=70") center 40%/cover no-repeat,
        linear-gradient(90deg, #0A2A62 0%, #1B4F9A 45%, #0E3578 100%);
    overflow: hidden;
    display: flex;
    align-items: center;
    justify-content: center;
    text-align: center;
    color: #FFFFFF;
}
.ei-banner::after {
    content: "";
    position: absolute;
    left: 0; right: 0; bottom: 0;
    height: 38px;
    background: repeating-linear-gradient(
        90deg,
        rgba(255,255,255,.06) 0 18px,
        rgba(255,255,255,0) 18px 36px
    );
    mask-image: linear-gradient(to top, #000 0%, transparent 100%);
}
.ei-banner-inner { position: relative; z-index: 1; padding: 22px 16px; }
.ei-banner h1 {
    margin: 0;
    font-size: 1.55rem;
    font-weight: 700;
    letter-spacing: .04em;
    color: #FFFFFF;
}
.ei-banner p {
    margin: 8px 0 0 0;
    font-size: .82rem;
    color: #FFFFFF !important;
    font-weight: 500;
}

.section-title {
    font-size: .95rem;
    font-weight: 700;
    color: var(--ei-navy);
    letter-spacing: .08em;
    text-transform: uppercase;
    border-bottom: 2px solid var(--ei-blue);
    padding: 0 0 8px 0;
    margin: 1.8rem 0 1rem 0;
}

.info-box {
    background: #F7F9FC;
    border: 1px solid var(--border);
    border-left: 3px solid var(--ei-blue);
    border-radius: 0;
    padding: 12px 16px;
    margin: 8px 0 14px 0;
    font-size: .82rem;
    color: var(--sub);
}

.ei-footer {
    margin-top: 2.4rem;
    padding: 28px 8px 18px 8px;
    border-top: 1px solid #EEF1F4;
    background: #FFFFFF;
}
.ei-units {
    display: flex;
    flex-wrap: wrap;
    justify-content: space-between;
    gap: 18px 24px;
    padding: 8px 0 22px 0;
    border-bottom: 1px solid #EEF1F4;
}
.ei-unit {
    color: #9AA3AE;
    font-size: .62rem;
    font-weight: 700;
    letter-spacing: .08em;
    text-transform: uppercase;
    text-align: center;
    line-height: 1.35;
    min-width: 110px;
}
.ei-unit strong {
    display: block;
    font-size: .72rem;
    color: #8B939C;
}
.ei-footer-bottom {
    display: flex;
    flex-wrap: wrap;
    justify-content: space-between;
    align-items: center;
    gap: 12px;
    padding-top: 16px;
    font-size: .72rem;
    color: var(--muted);
}
.ei-footer-links {
    display: flex;
    flex-wrap: wrap;
    gap: 8px 18px;
}
.ei-footer-links span {
    color: #9AA3AE;
    font-weight: 600;
    letter-spacing: .08em;
    text-transform: uppercase;
    font-size: .68rem;
}
.ei-copy { color: var(--ei-blue); font-weight: 500; }

div[data-testid="stMarkdownContainer"] p { color: var(--sub); }
hr { border-color: var(--border) !important; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════════════════════════════════
def _init_state():
    defaults = {
        "df_datos":        None,
        "avance_semanal":  None,
        "avance_diario":   None,
        "df_pyrevit":      None,
        "df_vinculado":    None,
        "avances_manuales": [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # Auto-cargar datos por defecto si están disponibles en disco y la sesión está vacía
    if st.session_state.df_datos is None:
        excel_path = os.path.join(os.path.dirname(__file__), "Demoliciones_Avance.xlsx")
        if os.path.exists(excel_path):
            try:
                with open(excel_path, "rb") as f:
                    raw = f.read()
                st.session_state.df_datos = load_datos_sheet(_io.BytesIO(raw))
                st.session_state.avance_semanal = load_avance_semanal(_io.BytesIO(raw))
                st.session_state.avance_diario = load_avance_diario(_io.BytesIO(raw))
            except Exception:
                pass

    if st.session_state.df_pyrevit is None:
        for json_path in [
            os.path.join(os.path.dirname(__file__), "data", "pyrevit_elementos_detallados.json"),
            os.path.join(os.path.dirname(__file__), "Aplicacion_pyRevit", "data", "pyrevit_elementos_detallados.json"),
        ]:
            if os.path.exists(json_path) and os.path.getsize(json_path) > 5000:
                try:
                    st.session_state.df_pyrevit = load_pyrevit_json(json_path)
                    break
                except Exception:
                    pass

    if st.session_state.df_vinculado is None and st.session_state.df_datos is not None and st.session_state.df_pyrevit is not None:
        try:
            st.session_state.df_vinculado = vincular_item_pyrevit(
                st.session_state.df_datos, st.session_state.df_pyrevit
            )
        except Exception:
            pass

_init_state()


# ═══════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("""
    <div style='padding:8px 0 16px 0;'>
      <div style='font-size:.68rem;font-weight:700;letter-spacing:.14em;color:#00A0E3;text-transform:uppercase;'>Ei Holding</div>
      <div style='font-size:1.05rem;font-weight:800;color:#0A2A62;margin-top:4px;letter-spacing:.03em;'>Control BIM</div>
      <div style='font-size:.75rem;color:#6B7280;margin-top:4px;'>Curvas S de proyecto</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("**Información del proyecto**")
    proyecto_nombre = st.text_input("Nombre del Proyecto", value="Demoliciones OC")
    proyecto_etapa  = st.text_input("Etapa / Frente", value="Estructura")
    proyecto_fecha  = st.date_input("Fecha de corte", format="DD-MM-YYYY")

    st.markdown("---")
    st.markdown("""
    <div style='font-size:.72rem;color:#6B7280;line-height:1.7;'>
    <b style='color:#00A0E3;'>Vinculación ITEM ↔ Revit</b><br>
    El campo <code>ITEM</code> del Excel corresponde al <b>Assembly Code</b> del modelo Revit.<br><br>
    PyRevit exporta elementos con:<br>
    • <code>assembly_code</code> (= ITEM)<br>
    • <code>cantidad_total</code> (m², ml, m³)<br>
    • <code>ejecutado</code> (True/False)<br>
    • <code>fecha</code> (YYYY-MM-DD)<br>
    • <code>unidad</code>, <code>nivel</code> (opcional)<br><br>
    <b style='color:#00A0E3;'>Formatos aceptados</b><br>
    Excel: <code>.xlsx</code> / <code>.xls</code><br>
    PyRevit: <code>.json</code> / <code>.csv</code>
    </div>
    """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════
# CABECERA
# ═══════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="ei-topbar">
  <nav class="ei-nav">
    <a href="#">Demoliciones</a>
  </nav>
</div>
""", unsafe_allow_html=True)

st.markdown(
    f"""
    <div class="ei-banner">
      <div class="ei-banner-inner">
        <h1>Control de Avance BIM</h1>
        <p>Proyecto: <b>{proyecto_nombre}</b> &nbsp;|&nbsp; Etapa: <b>{proyecto_etapa}</b>
        &nbsp;|&nbsp; Corte: <b>{proyecto_fecha.strftime("%d-%m-%Y")}</b></p>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ═══════════════════════════════════════════════════════════════════════
# SECCIÓN 1 – CARGA EXCEL
# ═══════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">1. Datos programados (Excel)</div>',
            unsafe_allow_html=True)

col_u1, col_u2 = st.columns([2, 1])
with col_u1:
    excel_file = st.file_uploader(
        "Cargar Excel de partidas (.xlsx / .xls)",
        type=["xlsx", "xls"], key="excel_up",
        help="Debe tener las hojas: Datos, Avance Semanal y Avance Diario")

if excel_file is not None:
    with st.spinner("Leyendo Excel..."):
        try:
            raw = excel_file.read()
            hojas = load_excel_sheet_names(_io.BytesIO(raw))
            with col_u2:
                st.success(f"✅ {excel_file.name}")
                st.caption(f"Hojas: {', '.join(hojas)}")

            st.session_state.df_datos = load_datos_sheet(_io.BytesIO(raw))

            if "Avance Semanal" in hojas:
                st.session_state.avance_semanal = load_avance_semanal(_io.BytesIO(raw))
            if "Avance Diario" in hojas:
                st.session_state.avance_diario = load_avance_diario(_io.BytesIO(raw))

            # Re-vincular si ya hay PyRevit cargado
            if st.session_state.df_pyrevit is not None:
                st.session_state.df_vinculado = get_vinculado(
                    st.session_state.df_datos, st.session_state.df_pyrevit)

        except Exception as e:
            st.error(f"❌ Error: {e}")

if st.session_state.df_datos is not None:
    with st.expander("📋 Ver partidas", expanded=False):
        df_show = st.session_state.df_datos.copy()
        for col in df_show.columns:
            if pd.api.types.is_datetime64_any_dtype(df_show[col]):
                df_show[col] = df_show[col].dt.strftime("%d-%m-%Y")
        # Resaltar columna ITEM
        st.info("💡 La columna **ITEM** es la clave de vinculación con el Codigo de Montaje en Revit.")
        st.dataframe(df_show, use_container_width=True, hide_index=True)
        st.caption(f"Total partidas: **{len(st.session_state.df_datos)}** | "
                   f"ITEMs únicos: **{st.session_state.df_datos['ITEM'].nunique()}**")


# ═══════════════════════════════════════════════════════════════════════
# SECCIÓN 2 – CARGA PYREVIT
# ═══════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">2. Avances desde PyRevit (Codigo de Montaje)</div>',
            unsafe_allow_html=True)

st.markdown("""
<div class="info-box">
  <b>Formato esperado del archivo PyRevit:</b><br>
  Columnas: <code>assembly_code</code> · <code>cantidad_total</code> · <code>ejecutado</code> (True/False)
  · <code>fecha</code> (YYYY-MM-DD) · <code>unidad</code> (m2/ml/m3) · <code>nivel</code><br>
  El campo <code>assembly_code</code> debe coincidir con el <b>ITEM</b> del Excel (ej: "3.10.2.1").
</div>
""", unsafe_allow_html=True)

col_py1, col_py2, col_py3 = st.columns([1, 1, 1])
with col_py1:
    json_file = st.file_uploader("Avances JSON (PyRevit)", type=["json"], key="json_up")
with col_py2:
    csv_file  = st.file_uploader("Avances CSV (PyRevit)",  type=["csv"],  key="csv_up")

df_pyr_nuevo = None
with col_py3:
    if json_file is not None:
        try:
            df_pyr_nuevo = load_pyrevit_json(json_file)
            st.success(f"✅ JSON: {len(df_pyr_nuevo)} elementos")
        except Exception as e:
            st.error(f"❌ JSON: {e}")
    elif csv_file is not None:
        try:
            df_pyr_nuevo = load_pyrevit_csv(csv_file)
            st.success(f"✅ CSV: {len(df_pyr_nuevo)} elementos")
        except Exception as e:
            st.error(f"❌ CSV: {e}")

if df_pyr_nuevo is not None:
    st.session_state.df_pyrevit = df_pyr_nuevo
    # Vincular automáticamente si hay Excel cargado
    if st.session_state.df_datos is not None:
        try:
            st.session_state.df_vinculado = get_vinculado(
                st.session_state.df_datos, df_pyr_nuevo)
        except Exception as e:
            st.warning(f"Aviso al vincular: {e}")

# Validación y resumen PyRevit
if st.session_state.df_pyrevit is not None:
    df_pyr = st.session_state.df_pyrevit
    val    = validar_columnas_pyrevit(df_pyr)
    res    = resumen_pyrevit(df_pyr)

    # Estado de columnas
    if val["ok"]:
        st.success("✅ Columnas requeridas presentes en el archivo PyRevit")
    else:
        st.warning(f"⚠️ Columnas faltantes: `{'`, `'.join(val['columnas_faltantes'])}`")

    for adv in val["advertencias"]:
        st.caption(f"ℹ️ {adv}")

    # KPIs PyRevit
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Elementos totales",    res.get("total_elementos", 0))
    c2.metric("Ejecutados (True)",    res.get("elementos_ejecutados", 0))
    c3.metric("Pendientes",           res.get("elementos_pendientes", 0))
    c4.metric("% Ejec. (conteo)",     f"{res.get('pct_ejecutado_count', 0):.1f}%")
    c5.metric("Assembly Codes",       res.get("assembly_codes_unicos", "-"))

    if "cantidad_ejecutada" in res:
        ca, cb, cc = st.columns(3)
        ca.metric("Cantidad total (Revit)",    f"{res.get('cantidad_total_sum',0):.2f}")
        cb.metric("Cantidad ejecutada",        f"{res.get('cantidad_ejecutada',0):.2f}")
        cc.metric("% Ejec. (por cantidad)",    f"{res.get('pct_cantidad_ejec',0):.1f}%")

    with st.expander("📋 Ver datos PyRevit cargados", expanded=False):
        st.dataframe(df_pyr, use_container_width=True, hide_index=True)


# ═══════════════════════════════════════════════════════════════════════
# AVANCE MANUAL PARA ITEMS SIN DATOS DE REVIT
# ═══════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">Avance manual para ITEMs sin datos Revit</div>',
            unsafe_allow_html=True)
st.markdown(
    '<div class="info-box">Registra la cantidad total de la partida y su cantidad ejecutada acumulada a la fecha indicada. '
    'El avance se incorpora a la tabla de control y a la curva real. Puedes volver a seleccionar un ITEM para actualizarlo. '
    'Los datos se mantienen durante la sesión actual; descarga el CSV como respaldo.</div>',
    unsafe_allow_html=True,
)

if st.session_state.df_datos is None or "ITEM" not in st.session_state.df_datos.columns:
    st.info("Carga primero el Excel con la hoja Datos para elegir un ITEM.")
else:
    items_excel = st.session_state.df_datos["ITEM"].dropna().astype(str).str.strip()
    items_excel = [item for item in items_excel.unique() if item and item.lower() != "nan"]
    codigos_revit = set()
    if st.session_state.df_pyrevit is not None and "assembly_code" in st.session_state.df_pyrevit.columns:
        codigos_revit = set(
            st.session_state.df_pyrevit["assembly_code"].dropna().astype(str).str.strip()
        )
    items_manuales = {str(r["ITEM"]).strip() for r in st.session_state.avances_manuales}
    items_capturables = [item for item in items_excel if item not in codigos_revit or item in items_manuales]

    if not items_capturables:
        st.success("Todos los ITEMs del Excel tienen datos Revit. No hay partidas para captura manual.")
    else:
        item_manual = st.selectbox(
            "ITEM sin datos Revit (o ya registrado manualmente)",
            items_capturables,
            key="select_item_avance_manual",
        )
        fila_excel = st.session_state.df_datos[
            st.session_state.df_datos["ITEM"].astype(str).str.strip() == item_manual
        ]
        actividad_manual = ""
        if not fila_excel.empty and "Nombre de tarea" in fila_excel.columns:
            actividad_manual = str(fila_excel.iloc[0]["Nombre de tarea"])
        registro_existente = next(
            (r for r in st.session_state.avances_manuales if str(r["ITEM"]).strip() == item_manual),
            None,
        )

        with st.form("form_avance_manual", clear_on_submit=False):
            st.caption(f"Partida: **{actividad_manual or item_manual}**")
            c_manual1, c_manual2, c_manual3 = st.columns(3)
            with c_manual1:
                cantidad_total_manual = st.number_input(
                    "Cantidad total de la partida", min_value=0.0,
                    value=float(registro_existente["cantidad_total"]) if registro_existente else 1.0,
                    step=1.0, format="%.3f", key=f"manual_total_{item_manual}",
                )
            with c_manual2:
                cantidad_ejecutada_manual = st.number_input(
                    "Cantidad ejecutada acumulada", min_value=0.0,
                    value=float(registro_existente["cantidad_ejecutada"]) if registro_existente else 0.0,
                    step=1.0, format="%.3f", key=f"manual_ejecutado_{item_manual}",
                )
            with c_manual3:
                unidades_disponibles = ["m2", "m3", "ml", "unid"]
                unidad_manual = st.selectbox(
                    "Unidad", unidades_disponibles,
                    index=unidades_disponibles.index(registro_existente["unidad"])
                    if registro_existente and registro_existente.get("unidad") in unidades_disponibles else 0,
                    key=f"manual_unidad_{item_manual}",
                )
            fecha_manual = st.date_input(
                "Fecha del avance acumulado",
                value=pd.to_datetime(registro_existente["fecha"]).date() if registro_existente else proyecto_fecha,
                format="DD-MM-YYYY", key=f"manual_fecha_{item_manual}",
            )
            guardar_manual = st.form_submit_button("Guardar / actualizar avance manual", type="primary")

        if guardar_manual:
            if cantidad_total_manual <= 0:
                st.error("La cantidad total debe ser mayor que cero.")
            elif cantidad_ejecutada_manual > cantidad_total_manual:
                st.error("La cantidad ejecutada no puede superar la cantidad total.")
            else:
                nueva_fila_manual = {
                    "ITEM": item_manual, "actividad": actividad_manual,
                    "cantidad_total": float(cantidad_total_manual),
                    "cantidad_ejecutada": float(cantidad_ejecutada_manual),
                    "unidad": unidad_manual, "fecha": pd.Timestamp(fecha_manual),
                }
                st.session_state.avances_manuales = [
                    r for r in st.session_state.avances_manuales
                    if str(r["ITEM"]).strip() != item_manual
                ] + [nueva_fila_manual]
                st.success(f"Avance manual guardado para el ITEM {item_manual}.")

        if registro_existente and st.button("Eliminar avance manual de este ITEM", key=f"eliminar_manual_{item_manual}"):
            st.session_state.avances_manuales = [
                r for r in st.session_state.avances_manuales
                if str(r["ITEM"]).strip() != item_manual
            ]
            st.rerun()

        if st.session_state.avances_manuales:
            df_manuales_vista = pd.DataFrame(st.session_state.avances_manuales).copy()
            df_manuales_vista["fecha"] = pd.to_datetime(df_manuales_vista["fecha"]).dt.strftime("%d-%m-%Y")
            st.dataframe(df_manuales_vista, use_container_width=True, hide_index=True)
            st.download_button(
                "Descargar respaldo de avances manuales (CSV)",
                data=pd.DataFrame(st.session_state.avances_manuales).to_csv(index=False).encode("utf-8-sig"),
                file_name="avances_manuales.csv", mime="text/csv", key="download_avances_manuales",
            )

df_manual_pyrevit = crear_registros_avance_manual(st.session_state.avances_manuales)
if st.session_state.df_pyrevit is not None and not st.session_state.df_pyrevit.empty:
    codigos_manual = set(df_manual_pyrevit["assembly_code"].astype(str)) if not df_manual_pyrevit.empty else set()
    df_pyrevit_base = st.session_state.df_pyrevit
    if codigos_manual and "assembly_code" in df_pyrevit_base.columns:
        df_pyrevit_base = df_pyrevit_base[
            ~df_pyrevit_base["assembly_code"].astype(str).str.strip().isin(codigos_manual)
        ]
    df_pyrevit_combinado = pd.concat(
        [df_pyrevit_base, df_manual_pyrevit], ignore_index=True, sort=False
    )
elif not df_manual_pyrevit.empty:
    df_pyrevit_combinado = df_manual_pyrevit
else:
    df_pyrevit_combinado = st.session_state.df_pyrevit


# ═══════════════════════════════════════════════════════════════════════
# SECCIÓN 3 – VINCULACIÓN ITEM ↔ ASSEMBLY CODE
# ═══════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">3. Vinculación ITEM ↔ Assembly Code</div>',
            unsafe_allow_html=True)

if st.session_state.df_datos is None:
    st.info("👆 Carga primero el Excel para ver la vinculación con Revit y los avances manuales.")
elif df_pyrevit_combinado is None or df_pyrevit_combinado.empty:
    st.info("👆 Carga PyRevit o ingresa avances manuales para vincular con los ITEMs del Excel.")
else:
    df_vinc = get_vinculado(st.session_state.df_datos, df_pyrevit_combinado)
    df_tabla_ctrl = tabla_avance_por_item(df_vinc)

    # KPIs de la vinculación
    if "pct_ejecutado" in df_vinc.columns:
        items_vinculados  = (df_vinc["pct_ejecutado"] > 0).sum()
        items_sin_vincular = (df_vinc["pct_ejecutado"] == 0).sum()
        pct_global = df_vinc["pct_ejecutado"].mean()

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("ITEMs en Excel",         len(df_vinc))
        k2.metric("ITEMs con avance",       int(items_vinculados))
        k3.metric("ITEMs sin avance",       int(items_sin_vincular))
        k4.metric("% Ejec. Promedio",       f"{pct_global:.1f}%")

    # Tabla de control
    st.markdown("**📊 Tabla de control de avance por ITEM (ordenado por Descripción de Montaje)**")
    if not df_tabla_ctrl.empty:
        # Mapa de colores para % Avance
        def color_estado(val):
            if isinstance(val, str):
                if "Crítico"     in val: return "color: #EF4444"
                if "En progreso" in val: return "color: #F59E0B"
                if "Avanzado"    in val: return "color: #10B981"
                if "Completado"  in val: return "color: #7C3AED"
            return ""

        st.dataframe(
            df_tabla_ctrl,
            use_container_width=True,
            hide_index=True,
            column_config={
                "ITEM":       st.column_config.TextColumn("ITEM", width="small"),
                "% Avance":   st.column_config.ProgressColumn(
                                  "% Avance", min_value=0, max_value=100, format="%.1f%%"),
                "Estado":     st.column_config.TextColumn("Estado", width="medium"),
            }
        )

    # Gráfico de avance por ITEM
    fig_item = grafico_avance_por_item_plotly(df_vinc)
    st.plotly_chart(fig_item, use_container_width=True)

    # Items sin vincular (advertencia)
    if "pct_ejecutado" in df_vinc.columns and "assembly_code" in df_vinc.columns:
        sin_match = df_vinc[df_vinc["assembly_code"].isna()]["ITEM"].tolist()
        if sin_match:
            with st.expander(f"⚠️ {len(sin_match)} ITEMs sin datos de avance", expanded=False):
                st.caption("Estos ITEMs del Excel no tienen Assembly Code en PyRevit ni avance manual:")
                st.write(sin_match)


# ═══════════════════════════════════════════════════════════════════════
# SECCIÓN 4 – CURVA S (PROGRAMADO VS REAL)
# ═══════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-title">4. Curva S (Programado vs Real)</div>', unsafe_allow_html=True)

if st.session_state.avance_semanal is None and st.session_state.avance_diario is None:
    st.info("👆 Carga el Excel para visualizar la Curva S.")
else:
    # Selector simple de vista: Semanal o Diaria
    # Controles: Frecuencia (Semanal/Diaria) y Selector de Partida
    col_sel1, col_sel2 = st.columns([1, 2])
    with col_sel1:
        frecuencia_vista = st.radio(
            "Frecuencia de avance:",
            ["📅 Semanal", "📆 Diaria"],
            horizontal=True,
            key="radio_frecuencia_curva"
        )

    # Obtener lista de actividades/partidas disponibles
    lista_act = []
    if "Semanal" in frecuencia_vista and st.session_state.avance_semanal is not None:
        lista_act = obtener_lista_actividades(st.session_state.avance_semanal)
    elif "Diaria" in frecuencia_vista and st.session_state.avance_diario is not None:
        lista_act = obtener_lista_actividades(st.session_state.avance_diario)

    opciones_partidas = ["— Todas las partidas (Global) —"] + lista_act

    with col_sel2:
        partida_sel = st.selectbox(
            "Filtrar por Partida / Actividad:",
            opciones_partidas,
            key="select_partida_curva_s"
        )

    actividad_filtro = None if partida_sel == "— Todas las partidas (Global) —" else partida_sel

    # Filtrar df_pyrevit por partida si corresponde
    df_pyr_filtrado = df_pyrevit_combinado
    if df_pyr_filtrado is not None and actividad_filtro is not None:
        act_clean = str(actividad_filtro).strip()
        masks = []

        # ── Estrategia 1: Buscar ITEM en la hoja 'Datos' del Excel que
        #    corresponda al nombre de tarea seleccionado, luego filtrar
        #    PyRevit por ese assembly_code (la relación más confiable).
        import re as _re
        _prefix_pat = _re.compile(r'^\[[\w\s]+\]\s*', _re.IGNORECASE)
        items_encontrados = set()
        if st.session_state.df_datos is not None and "ITEM" in st.session_state.df_datos.columns:
            if "Nombre de tarea" in st.session_state.df_datos.columns:
                act_norm_cmp = _prefix_pat.sub('', act_clean).strip().lower()
                for _, r in st.session_state.df_datos.iterrows():
                    nt   = str(r.get("Nombre de tarea", "")).strip()
                    item = str(r["ITEM"]).strip()
                    # Normalizar quitando prefijo [M2]/[M3]/[UN]/etc. de ambos lados
                    nt_norm = _prefix_pat.sub('', nt).strip().lower()
                    if nt and (nt_norm == act_norm_cmp or nt_norm in act_norm_cmp or act_norm_cmp in nt_norm):
                        items_encontrados.add(item)

        if items_encontrados and "assembly_code" in df_pyr_filtrado.columns:
            code_series = df_pyr_filtrado["assembly_code"].astype(str).str.strip()
            masks.append(code_series.isin(items_encontrados))

        # ── Estrategia 2: Comparar directamente contra assembly_description,
        #    normalizando prefijos como [M2], [m3], [ml], etc.
        if "assembly_description" in df_pyr_filtrado.columns:
            desc_norm = df_pyr_filtrado["assembly_description"].astype(str).str.strip().apply(
                lambda d: _prefix_pat.sub('', d).strip()
            )
            act_norm_str = _prefix_pat.sub('', act_clean).strip()
            masks.append(desc_norm.str.lower() == act_norm_str.lower())
            masks.append(desc_norm.apply(
                lambda d: bool(d and (d.lower() in act_norm_str.lower() or act_norm_str.lower() in d.lower()))
            ))

        if masks:
            final_mask = pd.Series(False, index=df_pyr_filtrado.index)
            for m in masks:
                final_mask |= m
            res_filt = df_pyr_filtrado[final_mask]
            if not res_filt.empty:
                df_pyr_filtrado = res_filt

    # Calcular curva real desde PyRevit si está disponible
    df_real_chart = None
    if df_pyr_filtrado is not None and not df_pyr_filtrado.empty:
        try:
            freq_code = "W" if "Semanal" in frecuencia_vista else "D"
            corte_ts = pd.to_datetime(proyecto_fecha) if proyecto_fecha is not None else None
            df_real_chart = get_curva_s_real_pyrevit(df_pyr_filtrado, frecuencia=freq_code, fecha_corte=corte_ts)
        except Exception:
            df_real_chart = None

    if "Semanal" in frecuencia_vista:
        if st.session_state.avance_semanal is not None:
            datos_sem = st.session_state.avance_semanal
            df_s = get_curva_s_semanal(datos_sem, actividad=actividad_filtro)

            titulo_chart = f"Curva S Semanal – {partida_sel}" if actividad_filtro else "Curva S Semanal – Programado (Excel) vs Real (BIM)"

            fig_s = grafico_curva_s_semanal_plotly(
                df_s,
                titulo=titulo_chart,
                df_real_pyrevit=df_real_chart
            )
            st.plotly_chart(fig_s, use_container_width=True)

            # KPIs comparativos (sin variación)
            pct_prog = df_s["acumulado"].iloc[-1] if not df_s.empty else 0.0
            pct_real = df_real_chart["pct_ejecutado_acum"].max() if (df_real_chart is not None and not df_real_chart.empty) else 0.0

            k1, k2, k3 = st.columns(3)
            k1.metric("% Programado Acumulado", f"{pct_prog:.1f}%")
            k2.metric("% Real Ejecutado (BIM)", f"{pct_real:.1f}%" if df_real_chart is not None else "Sin datos")
            k3.metric("Semanas Evaluadas", len(df_s))

            with st.expander("📊 Ver datos de la Curva S Semanal", expanded=False):
                df_show_s = df_s[["semana_label", "semana_rango", "acumulado"]].copy()
                df_show_s.columns = ["Semana", "Rango", "Programado %"]
                st.dataframe(df_show_s, use_container_width=True, hide_index=True)
        else:
            st.warning("El archivo Excel no contiene la hoja 'Avance Semanal'.")

    else:
        if st.session_state.avance_diario is not None:
            datos_dia = st.session_state.avance_diario
            df_d = get_curva_s_diaria(datos_dia, actividad=actividad_filtro)

            titulo_chart_d = f"Curva S Diaria – {partida_sel}" if actividad_filtro else "Curva S Diaria – Programado (Excel) vs Real (BIM)"

            fig_d = grafico_curva_s_diaria_plotly(
                df_d,
                titulo=titulo_chart_d,
                df_real_pyrevit=df_real_chart
            )
            st.plotly_chart(fig_d, use_container_width=True)

            pct_prog_d = df_d["acumulado"].iloc[-1] if not df_d.empty else 0.0
            pct_real_d = df_real_chart["pct_ejecutado_acum"].max() if (df_real_chart is not None and not df_real_chart.empty) else 0.0

            k1, k2, k3 = st.columns(3)
            k1.metric("% Programado Acumulado", f"{pct_prog_d:.1f}%")
            k2.metric("% Real Ejecutado (BIM)", f"{pct_real_d:.1f}%" if df_real_chart is not None else "Sin datos")
            k3.metric("Días Registrados", len(df_d))

            with st.expander("📊 Ver datos de la Curva S Diaria", expanded=False):
                df_show_d = df_d.copy()
                df_show_d["fecha"] = pd.to_datetime(df_show_d["fecha"]).dt.strftime("%d-%m-%Y")
                df_show_d.columns = ["Fecha", "Parcial %", "Programado %"]
                st.dataframe(df_show_d, use_container_width=True, hide_index=True)


# ═══════════════════════════════════════════════════════════════════════
# FOOTER
# ═══════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="ei-footer">
  <div class="ei-units">
  </div>
  <div class="ei-footer-bottom">
    <div class="ei-copy">© Desarrollado por Daniel Salcedo Salcedo 2026.</div>
    <div class="ei-footer-links">
    </div>
    <div>Coodrinador BIM · Daniel Salcedo</div>
  </div>
</div>
""", unsafe_allow_html=True)
