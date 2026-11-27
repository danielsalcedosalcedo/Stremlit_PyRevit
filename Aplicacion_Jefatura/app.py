"""Seguimiento y control del avance BIM · Boulevard Gastronómico Mall Plaza El Trébol.

PROPUESTA DE MEJORA DE app.py
=============================
Esta versión rediseña la interfaz y conserva la lógica de datos del tablero original:

* Se mantienen sin cambios las funciones de lectura, filtrado y cálculo
  (`_load_local_data`, `_filtered_as_of`, `_build_item_summary`,
  `_build_weighted_global_curve`, `_activity_choices`, las exportaciones a Excel y
  `_calcular_kpis_curva_s`) y las llamadas a `logic.loader`, `logic.processor`,
  `logic.charts` y `logic.curva_lb`.
* Cambia la presentación: barra lateral con tres vistas por rol (Gerencia,
  Planificación y BIM), un solo frente de obra a la vez, resumen ejecutivo, estados por
  partida (atrasada, en riesgo, en plazo, ejecutada, programada), cronograma,
  ficha de partida y control de calidad de los datos del modelo.
* Los colores de la curva S se remapean desde aquí (`_theme_fig`), sin editar
  `logic/curva_lb.py`.

Requisitos sugeridos: Streamlit 1.37 o superior (usa `st.dialog` y la selección de filas
de `st.dataframe`; con versiones anteriores la ficha se abre con un selector).
"""

from __future__ import annotations

import html
import io
import re
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
from openpyxl.styles import Alignment, Font, PatternFill
import streamlit as st

ROOT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT_DIR))

from logic.charts import (  # noqa: E402
    grafico_avance_por_item_plotly,
    grafico_comparativo_cantidades_plotly,
    grafico_curva_real_pyrevit_plotly,
    grafico_curva_s_diaria_plotly,
    grafico_curva_s_semanal_plotly,
)
from logic.curva_lb import (  # noqa: E402
    construir_base_curva,
    grafico_curva_lb_plotly,
    tabla_curva_lb,
)
from logic.loader import (  # noqa: E402
    load_avance_diario,
    load_avance_semanal,
    load_datos_sheet,
    load_excel_sheet_names,
    load_pyrevit_csv,
    load_pyrevit_json,
)
from logic.processor import (  # noqa: E402
    calcular_comparativo_cantidades_periodo,
    calcular_curva_s_diaria,
    calcular_curva_s_real_pyrevit,
    calcular_curva_s_semanal,
    detectar_unidad_pyrevit,
    extraer_cantidad_pyrevit,
    generar_reporte_diario_partidas,
)

st.set_page_config(
    page_title="Seguimiento BIM | El Trébol",
    page_icon="🔷",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Estilos ────────────────────────────────────────────────────────────────
# Sala de control de obra: neutros fríos, bordes de 1 px, números tabulares y estados
# que combinan color y texto. La barra lateral usa el azul marino corporativo.
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&display=swap');
    :root {
        --bg:#F1F3F6; --surface:#FFFFFF; --ink:#17212D; --ink2:#536172;
        --line:#D3DAE2; --line2:#B7C1CC; --grid:#E7EBF0;
        --plan:#6C7B8E; --real:#0B6BB5;
        --ok:#2A7D4F; --risk:#A5620A; --late:#B42318; --sched:#6B7A8C;
        --nav:#0A2A62; --nav2:#123A80; --navink:#E8EEF8; --navmute:#9DB1D3; --navline:#26467F;
    }
    html, body, [data-testid="stAppViewContainer"], .stApp {
        background:var(--bg) !important; color:var(--ink) !important;
        font-family:'IBM Plex Sans','Segoe UI',system-ui,sans-serif !important;
        font-variant-numeric:tabular-nums;
    }
    [data-testid="stHeader"] {background:transparent;}
    [data-testid="stToolbar"], [data-testid="stDecoration"], #MainMenu, footer {display:none !important;}
    .block-container {max-width:1280px !important; padding:1rem 1.6rem 3rem !important;}

    /* Barra lateral */
    [data-testid="stSidebar"] {background:var(--nav) !important; border-right:0;}
    [data-testid="stSidebar"] *, [data-testid="stSidebar"] label {color:var(--navink) !important;}
    [data-testid="stSidebar"] [data-testid="stRadio"] > label {display:none;}
    [data-testid="stSidebar"] [role="radiogroup"] {gap:2px;}
    [data-testid="stSidebar"] [role="radiogroup"] label {
        display:flex; box-sizing:border-box; padding:8px 10px; border-radius:3px; border:1px solid transparent; width:100%; cursor:pointer;}
    [data-testid="stSidebar"] [role="radiogroup"] label:hover {background:var(--nav2);}
    [data-testid="stSidebar"] [role="radiogroup"] label > div:first-child:not(:has([data-testid="stMarkdownContainer"])),
    [data-testid="stSidebar"] [role="radiogroup"] label > div > div:first-child:not([data-testid="stMarkdownContainer"]) {display:none;}
    [data-testid="stSidebar"] [role="radiogroup"] label:has(input:checked) {
        background:var(--nav2); border-color:var(--navline); box-shadow:inset 3px 0 0 var(--navink); font-weight:600;}
    .side-brand h1 {font-size:1rem; font-weight:600; line-height:1.25; margin:0; color:var(--navink);}
    .side-brand p {font-size:.8rem; color:var(--navmute); margin:.25rem 0 0;}
    .side-h {font-size:.8rem; color:var(--navmute) !important; margin:1.1rem 0 .25rem;}
    .side-foot {font-size:.8rem; color:var(--navmute) !important; border-top:1px solid var(--navline); padding-top:.7rem; margin-top:1.4rem;}
    [data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"] {display:none !important;}

    /* Navegación superior */
    .nav-brand {background:var(--nav); color:var(--navink); padding:.9rem 1.1rem; border-radius:3px; margin-bottom:.9rem;}
    .nav-brand h1 {font-size:1.05rem; font-weight:600; line-height:1.25; margin:0; color:var(--navink);}
    .nav-brand p {font-size:.8rem; color:var(--navmute); margin:.2rem 0 0;}
    .nav-h {font-size:.78rem; color:var(--ink2); margin:0 0 .25rem; font-weight:500;}
    .nav-txt {font-size:1rem; font-weight:600; color:var(--ink); padding:.35rem 0;}

    /* Cabecera de vista */
    .vhead {display:flex; flex-wrap:wrap; justify-content:space-between; align-items:flex-end; gap:.4rem 1rem; margin:.2rem 0 .9rem;}
    .vhead h2 {font-size:1.45rem; font-weight:600; letter-spacing:-.01em; margin:0; padding:0;}
    .vhead p {color:var(--ink2); margin:.15rem 0 0;}
    .vhead .cut {font-size:.85rem; color:var(--ink2); text-align:right;}
    .vhead .cut b {color:var(--ink); font-weight:600;}

    /* Superficies */
    [data-testid="stVerticalBlockBorderWrapper"] {background:var(--surface); border-color:var(--line) !important; border-radius:3px !important;}
    .panel-h {display:flex; justify-content:space-between; align-items:center; gap:.5rem 1rem; flex-wrap:wrap; margin:.1rem 0 .55rem;}
    .panel-h h3 {font-size:.95rem; font-weight:600; margin:0; padding:0;}
    .panel-h span {font-size:.8rem; color:var(--ink2);}
    .muted {color:var(--ink2);}

    /* Resumen ejecutivo */
    .kpis {display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:1px; background:var(--line);
        border:1px solid var(--line); border-radius:3px; overflow:hidden; margin-bottom:.7rem;}
    .kpis.kpis3 {grid-template-columns:repeat(3,minmax(0,1fr));}
    .kpi {background:var(--surface); padding:.75rem .9rem; min-width:0;}
    .kpi .l {font-size:.8rem; color:var(--ink2);}
    .kpi .v {font-size:1.7rem; font-weight:600; line-height:1.15; margin-top:.1rem; letter-spacing:-.01em;}
    .kpi .v small {font-size:.9rem; font-weight:500; color:var(--ink2); margin-left:3px; letter-spacing:0;}
    .kpi .s {font-size:.85rem; color:var(--ink2); margin-top:.1rem;}
    @media (max-width:1180px) {
        .kpis:not(.kpis3) {grid-template-columns:repeat(2,minmax(0,1fr));}
        .kpis:not(.kpis3) .kpi:last-child {grid-column:1 / -1;}
    }
    @media (max-width:560px) {.kpis, .kpis.kpis3 {grid-template-columns:1fr;}}
    .concl {font-size:.98rem; max-width:96ch; margin:.1rem 0 .9rem;}
    .concl strong {font-weight:600;} .concl span {color:var(--ink2);}

    /* Estados */
    .sw {display:inline-block; width:9px; height:9px; border-radius:2px; flex:none; background:var(--c);}
    .s-done, .s-ok {--c:var(--ok);} .s-risk {--c:var(--risk);} .s-late {--c:var(--late);} .s-sched, .s-nodata {--c:var(--sched);}
    .pill {display:inline-flex; align-items:center; gap:6px; padding:2px 8px 2px 7px; border-radius:3px; white-space:nowrap;
        border:1px solid color-mix(in srgb, var(--c) 45%, #fff); background:color-mix(in srgb, var(--c) 10%, #fff);
        color:var(--c); font-size:.8rem; font-weight:500;}
    .s-done.pill .sw {border-radius:50%;}
    .stackbar {display:flex; height:10px; margin-top:.6rem; border:1px solid var(--line2); border-radius:2px; overflow:hidden; gap:1px; background:var(--line2);}
    .stackbar i {display:block; background:var(--c);}
    .legend {display:flex; flex-wrap:wrap; gap:2px 12px; margin:.5rem 0 0; padding:0; list-style:none; font-size:.8rem;}
    .legend li {display:inline-flex; align-items:center; gap:5px;}
    .help {position:relative; display:inline-grid; place-items:center; width:15px; height:15px; margin-left:5px; border:1px solid var(--ink2);
        border-radius:50%; font-size:10px; font-weight:600; color:var(--ink2); cursor:help; line-height:1; vertical-align:1px;}
    .help::after {content:attr(data-tip); position:absolute; top:calc(100% + 6px); left:-8px; width:max-content; max-width:250px; background:var(--ink);
        color:#fff; padding:7px 9px; border-radius:3px; font-size:.8rem; font-weight:400; line-height:1.35; text-align:left; white-space:normal;
        opacity:0; pointer-events:none; z-index:999;}
    .help:hover::after {opacity:1;} .help.end::after {left:auto; right:-8px;}
    .crit-row {display:grid; grid-template-columns:minmax(0,1fr) auto; gap:3px 12px; padding:.55rem 0; border-top:1px solid var(--line);}
    .crit-row:first-child {border-top:0;}
    .crit-row .n {font-weight:500; min-width:0;} .crit-row .c {color:var(--ink2); font-weight:400; margin-right:6px;}
    .crit-row .d {font-weight:600; text-align:right; color:var(--late);}
    .bar {position:relative; height:6px; background:var(--grid); border-radius:1px; grid-column:1 / -1;}
    .bar i {position:absolute; inset:0 auto 0 0; background:var(--real); border-radius:1px;}
    .bar b {position:absolute; top:-4px; bottom:-4px; width:2px; background:var(--ink); margin-left:-1px;}
    .note {padding:.6rem .8rem; border:1px solid var(--line2); border-left:3px solid var(--risk); background:color-mix(in srgb, var(--risk) 7%, #fff);
        border-radius:3px; font-size:.9rem; margin-bottom:.7rem;}

    /* Controles de Streamlit */
    [data-testid="stWidgetLabel"] p, label[data-testid="stWidgetLabel"] {font-size:.8rem !important; color:var(--ink2) !important;}
    [data-baseweb="select"] > div, [data-baseweb="input"] {border-radius:3px !important; border-color:var(--line2) !important; background:#fff !important;}
    .stButton > button, .stDownloadButton > button {border-radius:3px; border:1px solid var(--line2); background:#fff; color:var(--ink); font-weight:500; padding:.3rem .8rem;}
    .stButton > button:hover, .stDownloadButton > button:hover {border-color:var(--ink); color:var(--ink); background:var(--bg);}
    .stButton > button:disabled {opacity:.5;}
    [data-testid="stMain"] [data-testid="stRadio"] [role="radiogroup"] {gap:0;}
    [data-testid="stMain"] [data-testid="stRadio"] [role="radiogroup"] label {
        border:1px solid var(--line2); padding:4px 12px; background:#fff; margin:0; margin-right:-1px; cursor:pointer;}
    [data-testid="stMain"] [data-testid="stRadio"] [role="radiogroup"] label > div:first-child:not(:has([data-testid="stMarkdownContainer"])),
    [data-testid="stMain"] [data-testid="stRadio"] [role="radiogroup"] label > div > div:first-child:not([data-testid="stMarkdownContainer"]) {display:none;}
    [data-testid="stMain"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) {background:var(--ink); color:#fff; border-color:var(--ink);}
    [data-testid="stMain"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) p {color:#fff !important;}
    [data-testid="stDataFrame"] {border:1px solid var(--line); border-radius:3px;}
    .footer {margin-top:2rem; padding:14px 2px; border-top:1px solid var(--line); color:var(--ink2); font-size:.78rem;}
    hr {border-color:var(--line) !important;}
    </style>
    """,
    unsafe_allow_html=True,
)

DEMOLICIONES_EXCEL_CANDIDATES = (
    ROOT_DIR / "data" / "programado" / "Demoliciones_Avance.xlsx",
    ROOT_DIR / "Demoliciones_Avance.xlsx",
)
HORMIGONES_EXCEL_CANDIDATES = (
    ROOT_DIR / "data" / "programado" / "Hormigones_Avance.xlsx",
    ROOT_DIR / "Hormigones_Avance.xlsx",
)
EEMM_EXCEL_CANDIDATES = (
    ROOT_DIR / "data" / "programado" / "EEMM_Avance.xlsx",
    ROOT_DIR / "EEMM_Avance.xlsx",
)
DEMOLICIONES_PYREVIT_CANDIDATES = (
    ROOT_DIR / "data" / "pyrevit_elementos_detallados.json",
    ROOT_DIR / "Aplicacion_pyRevit" / "data" / "pyrevit_elementos_detallados.json",
    ROOT_DIR / "data" / "pyrevit_elementos_detallados.csv",
)
HORMIGONES_PYREVIT_CANDIDATES = DEMOLICIONES_PYREVIT_CANDIDATES
EEMM_IFC_DIRECTORY = ROOT_DIR / "data" / "Modelos IFC"
# Electricidad, Sanitario y PCI: avance BIM desde exportaciones PyRevit de la categoría Tuberías y tubos.
ELECTRICIDAD_EXCEL_CANDIDATES = (ROOT_DIR / "data" / "programado" / "Electricidad_Avance.xlsx",)
ELECTRICIDAD_PYREVIT_CANDIDATES = (
    ROOT_DIR / "data" / "pyrevit_electricidad_tuberias.json",
    ROOT_DIR / "data" / "pyrevit_electricidad_tuberias.csv",
)
SANITARIO_EXCEL_CANDIDATES = (ROOT_DIR / "data" / "programado" / "Sanitario_Avance.xlsx",)
SANITARIO_PYREVIT_CANDIDATES = (
    ROOT_DIR / "data" / "pyrevit_tuberias_elementos_detallados.json",
    ROOT_DIR / "data" / "pyrevit_sanitario_tuberias.json",
    ROOT_DIR / "data" / "pyrevit_sanitario_tuberias.csv",
)
PCI_EXCEL_CANDIDATES = (ROOT_DIR / "data" / "programado" / "PCI_Avance.xlsx",)
PCI_PYREVIT_CANDIDATES = (
    ROOT_DIR / "data" / "pyrevit_pci_tuberias.json",
    ROOT_DIR / "data" / "pyrevit_pci_tuberias.csv",
)


def _first_existing(paths: tuple[Path, ...]) -> Path | None:
    return next((path for path in paths if path.is_file() and path.stat().st_size > 0), None)


@st.cache_data(show_spinner=False)
def _load_ifc_data(
    ifc_path: str,
    modified_at: float,
    include_geometry: bool,
    visible_ids: tuple[int, ...] = (),
) -> dict:
    """Carga propiedades IFC y cachea geometría hasta que cambia el archivo."""
    del modified_at
    from logic.ifc_viewer import load_ifc_model

    return load_ifc_model(
        ifc_path,
        include_geometry=include_geometry,
        element_ids=set(visible_ids) if include_geometry else None,
    )


# Las hojas Datos de Electricidad, Sanitario y PCI usan EDT / Actividad / Inicio en lugar de ITEM / Nombre de tarea / Comienzo.
_DATOS_RENAME = {"EDT": "ITEM", "Actividad": "Nombre de tarea", "Inicio": "Comienzo"}


def _normalize_datos(df: pd.DataFrame | None) -> pd.DataFrame | None:
    """Unifica los nombres de columna de la hoja Datos para que todos los frentes usen las mismas claves."""
    if df is None or "ITEM" in df.columns:
        return df
    return df.rename(columns={old: new for old, new in _DATOS_RENAME.items() if old in df.columns})


@st.cache_data(show_spinner=False)
def _load_local_data(
    excel_path: str | None,
    excel_mtime: float,
    pyrevit_paths: tuple[str, ...],
    pyrevit_mtimes: tuple[float, ...],
    require_pyrevit: bool = True,
) -> dict:
    """Lee las fuentes fijas y renueva la caché cuando cambia su fecha de modificación."""
    del excel_mtime, pyrevit_mtimes  # Se conservan como parte de la clave de caché.
    result = {
        "df_datos": None,
        "avance_semanal": None,
        "avance_diario": None,
        "df_pyrevit": None,
        "pyrevit_path": None,
        "errores": [],
    }

    if excel_path:
        try:
            hojas = load_excel_sheet_names(excel_path)
            result["df_datos"] = _normalize_datos(load_datos_sheet(excel_path))
            if "Avance Semanal" in hojas:
                result["avance_semanal"] = load_avance_semanal(excel_path)
            if "Avance Diario" in hojas:
                result["avance_diario"] = load_avance_diario(excel_path)
        except Exception as exc:
            result["errores"].append(f"No se pudo leer el Excel programado: {exc}")

    if not require_pyrevit:
        return result

    if pyrevit_paths:
        errores_pyrevit = []
        try:
            for pyrevit_path in pyrevit_paths:
                try:
                    if pyrevit_path.lower().endswith(".csv"):
                        df_pyrevit = load_pyrevit_csv(pyrevit_path)
                    else:
                        df_pyrevit = load_pyrevit_json(pyrevit_path)

                    if df_pyrevit.empty:
                        errores_pyrevit.append(f"{pyrevit_path}: el archivo no contiene registros.")
                        continue
                    if "assembly_code" not in df_pyrevit.columns:
                        errores_pyrevit.append(f"{pyrevit_path}: falta la columna assembly_code.")
                        continue

                    result["df_pyrevit"] = df_pyrevit
                    result["pyrevit_path"] = pyrevit_path
                    break
                except Exception as exc:
                    errores_pyrevit.append(f"{pyrevit_path}: {exc}")
        except Exception as exc:
            errores_pyrevit.append(str(exc))

        if result["df_pyrevit"] is None:
            result["errores"].append(
                "No se pudo cargar una exportación BIM válida. " + " | ".join(errores_pyrevit)
            )
    else:
        result["errores"].append(
            "No se encontró ningún archivo de exportación PyRevit en las rutas del repositorio."
        )

    return result


def _filtered_as_of(df: pd.DataFrame, cutoff: date) -> pd.DataFrame:
    """Conserva todos los pendientes y solo los ejecutados hasta la fecha de corte."""
    if df.empty or "ejecutado" not in df.columns:
        return df.copy()

    result = df.copy()
    executed = result["ejecutado"].fillna(False).astype(bool)
    dates = pd.to_datetime(result.get("fecha"), format="mixed", dayfirst=True, errors="coerce")
    if "ei_fecha_ejecucion" in result.columns:
        fallback = pd.to_datetime(
            result["ei_fecha_ejecucion"], format="mixed", dayfirst=True, errors="coerce"
        )
        dates = dates.fillna(fallback)
    earliest = dates[executed].dropna().min()
    if pd.notna(earliest):
        dates = dates.fillna(earliest)

    result["_fecha_control"] = dates
    cutoff_ts = pd.Timestamp(cutoff)
    keep = (~executed) | (result["_fecha_control"].notna() & (result["_fecha_control"] <= cutoff_ts))
    return result.loc[keep].copy()


def _build_item_summary(
    df: pd.DataFrame,
    progress_as_of: pd.DataFrame | None = None,
    unit_override: str | None = None,
) -> pd.DataFrame:
    """Resume cantidades del alcance completo y avance opcional a una fecha de corte."""
    columns = [
        "ITEM",
        "assembly_code",
        "assembly_description",
        "Nombre de tarea",
        "Unidad",
        "Cantidad total",
        "Cantidad ejecutada",
        "pct_ejecutado",
        "Elementos",
        "Elementos ejecutados",
    ]
    if df.empty or "assembly_code" not in df.columns:
        return pd.DataFrame(columns=columns)

    rows = []
    for code, group in df.groupby("assembly_code", dropna=True):
        code_text = str(code).strip()
        if not code_text or code_text.lower() == "nan":
            continue

        description = ""
        if "assembly_description" in group.columns:
            modes = group["assembly_description"].dropna().astype(str).mode()
            description = modes.iloc[0] if not modes.empty else ""

        progress_group = group
        if progress_as_of is not None:
            progress_group = progress_as_of.loc[
                progress_as_of["assembly_code"].astype(str).str.strip() == code_text
            ]
        executed = (
            progress_group["ejecutado"].fillna(False).astype(bool)
            if "ejecutado" in progress_group
            else pd.Series(False, index=progress_group.index)
        )
        done = progress_group.loc[executed]
        unit = unit_override or detectar_unidad_pyrevit(group)
        total = float(extraer_cantidad_pyrevit(group, unit).sum())
        quantity_done = float(extraer_cantidad_pyrevit(done, unit).sum())
        pct = min(100.0, quantity_done / total * 100) if total > 0 else 0.0
        rows.append({
            "ITEM": code_text,
            "assembly_code": code_text,
            "assembly_description": description,
            "Nombre de tarea": description or code_text,
            "Unidad": unit,
            "Cantidad total": round(float(total), 2),
            "Cantidad ejecutada": round(float(quantity_done), 2),
            "pct_ejecutado": round(pct, 2),
            "Elementos": int(len(group)),
            "Elementos ejecutados": int(executed.sum()),
        })

    return pd.DataFrame(rows, columns=columns).sort_values(
        "pct_ejecutado", ascending=True
    )


def _build_weighted_global_curve(
    planned_data: dict,
    activity_choices: list[dict[str, str]],
    item_summary: pd.DataFrame,
    unit: str,
    frequency: str,
) -> tuple[pd.DataFrame, float, float]:
    """Calcula el programado global como cantidad BIM programada / cantidad BIM total."""
    if frequency == "W":
        curve = calcular_curva_s_semanal(planned_data)
    else:
        curve = calcular_curva_s_diaria(planned_data)
    if curve.empty:
        return curve, 0.0, 0.0

    unit_items = item_summary[item_summary["Unidad"] == unit]
    totals_by_code = dict(zip(unit_items["ITEM"], unit_items["Cantidad total"]))
    total_bim = float(unit_items["Cantidad total"].sum())
    if total_bim <= 0:
        return pd.DataFrame(), 0.0, 0.0

    activity_codes = {
        choice["activity"]: choice["code"]
        for choice in activity_choices
        if choice["activity"] in planned_data["actividades"]
    }
    weighted_percent = pd.Series(0.0, index=range(len(curve)))
    scheduled_bim = 0.0
    weighted_codes: set[str] = set()
    for activity_name, code_value in activity_codes.items():
        code = str(code_value).strip()
        weight = float(totals_by_code.get(code, 0.0))
        if weight <= 0 or code in weighted_codes:
            continue
        percentages = planned_data["actividades"][activity_name].get("acumulado", [])
        if len(percentages) != len(curve):
            continue
        weighted_percent = weighted_percent.add(
            pd.Series(percentages, dtype="float64") * weight,
            fill_value=0.0,
        )
        scheduled_bim += weight
        weighted_codes.add(code)

    curve["acumulado"] = (weighted_percent / total_bim).round(2)
    curve["parcial"] = curve["acumulado"].diff().fillna(curve["acumulado"])
    return curve, total_bim, scheduled_bim


def _normalize_task_name(value: object) -> str:
    """Normaliza nombres para enlazar las actividades programadas con Excel."""
    text = "" if value is None or pd.isna(value) else str(value).strip()
    text = re.sub(r"^\[[^\]]+\]\s*", "", text)
    return re.sub(r"\s+", " ", text).casefold()


def _activity_choices(planned_data: dict, df_datos: pd.DataFrame | None) -> list[dict[str, str]]:
    """Arma opciones de actividad mostrando ITEM y nombre de tarea."""
    activities = planned_data.get("actividades", {})
    rows = df_datos.to_dict("records") if df_datos is not None and not df_datos.empty else []
    valid_items = {
        str(row.get("ITEM", "")).strip()
        for row in rows
        if str(row.get("ITEM", "")).strip()
    }
    choices = []

    for activity_name, details in activities.items():
        scheduled_item = str(details.get("item", "")).strip()
        match = next(
            (
                row for row in rows
                if scheduled_item
                and str(row.get("ITEM", "")).strip() == scheduled_item
            ),
            None,
        )
        if match is None:
            activity_normalized = _normalize_task_name(activity_name)
            matches = [
                row for row in rows
                if _normalize_task_name(row.get("Nombre de tarea", "")) == activity_normalized
            ]
            if len(matches) == 1:
                match = matches[0]

        item_value = (
            match.get("ITEM", "")
            if match
            else scheduled_item
            if scheduled_item in valid_items
            or re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]*", scheduled_item)
            else ""
        )
        item_code = "" if pd.isna(item_value) else str(item_value).strip()
        task_value = match.get("Nombre de tarea", "") if match else activity_name
        task_name = "" if pd.isna(task_value) else str(task_value).strip()
        if not task_name:
            task_name = str(activity_name).strip()
        choices.append({
            "item": item_code or "Sin ITEM",
            "task": task_name,
            "activity": str(activity_name),
            "code": item_code,
        })

    return choices


def _build_excel_export(
    reports: dict[str, pd.DataFrame],
    codes_by_activity: dict[str, str] | None = None,
) -> bytes:
    """Construye un libro Excel con una hoja por partida."""
    output = io.BytesIO()
    used_names: set[str] = set()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for report_key, report in reports.items():
            code = str(
                report.attrs.get("codigo_partida")
                or (codes_by_activity or {}).get(report_key)
                or report_key
            ).strip()
            sheet_name = re.sub(r"[\[\]:*?/\\]", "_", code).strip("' ")[:31]
            if not sheet_name:
                sheet_name = "Partida"
            base_name = sheet_name
            suffix = 2
            while sheet_name.casefold() in used_names:
                suffix_text = f"_{suffix}"
                sheet_name = f"{base_name[:31 - len(suffix_text)]}{suffix_text}"
                suffix += 1
            used_names.add(sheet_name.casefold())
            worksheet = writer.book.create_sheet(sheet_name)
            task_name = report.attrs.get("nombre_partida", report_key)
            worksheet.merge_cells("A1:E1")
            title = worksheet["A1"]
            title.value = f"{code} - {task_name}"
            title.font = Font(
                name="Calibri", size=14, bold=True, color="FFFFFF"
            )
            title.fill = PatternFill(
                fill_type="solid", fgColor="0A2A62"
            )
            title.alignment = Alignment(
                horizontal="left", vertical="center"
            )
            worksheet.row_dimensions[1].height = 26
            report.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
                startrow=2,
            )
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes = "A4"
            worksheet.auto_filter.ref = f"A3:E{worksheet.max_row}"
            worksheet.column_dimensions["A"].width = 15
            for column in ("B", "D"):
                worksheet.column_dimensions[column].width = 22
                for cell in worksheet[column][3:]:
                    cell.number_format = '0.00"%"'
            for column in ("C", "E"):
                worksheet.column_dimensions[column].width = 28
                for cell in worksheet[column][3:]:
                    cell.number_format = "#,##0.000"
            for cell in worksheet["A"][3:]:
                cell.number_format = "DD-MM-YYYY"
    return output.getvalue()


def _build_global_daily_reports(
    df_pyrevit: pd.DataFrame,
    item_summary: pd.DataFrame,
    activities: dict[str, dict],
    codes_by_activity: dict[str, str],
    fecha_inicio: pd.Timestamp,
    fecha_fin: pd.Timestamp,
    unit_labels: dict[str, str] | None = None,
) -> dict[str, pd.DataFrame]:
    """Genera avances diarios globales separados por unidad BIM."""
    start_date = pd.Timestamp(fecha_inicio).normalize()
    end_date = pd.Timestamp(fecha_fin).normalize()
    if start_date > end_date:
        raise ValueError("La fecha inicial del reporte no puede ser posterior a la fecha final.")

    dates = pd.date_range(start_date, end_date, freq="D")
    unit_labels = unit_labels or {
        "m2": "m²", "m3": "m³", "ml": "ml", "unid": "unitario"
    }
    if item_summary.empty:
        return {}

    totals_by_code = dict(zip(
        item_summary["assembly_code"].astype(str).str.strip(),
        item_summary["Cantidad total"],
    ))
    units_by_code = dict(zip(
        item_summary["assembly_code"].astype(str).str.strip(),
        item_summary["Unidad"],
    ))
    planned_by_unit = {
        unit: pd.Series(0.0, index=dates) for unit in unit_labels
    }
    scheduled_codes: set[str] = set()
    for activity, details in activities.items():
        code = str(codes_by_activity.get(activity) or details.get("item") or "").strip()
        if (
            not code
            or code in scheduled_codes
            or code not in totals_by_code
            or units_by_code[code] not in unit_labels
        ):
            continue
        scheduled_codes.add(code)

        activity_start = pd.to_datetime(details.get("inicio"), errors="coerce")
        duration_value = pd.to_numeric(details.get("dias"), errors="coerce")
        if pd.isna(activity_start) or pd.isna(duration_value) or duration_value <= 0:
            continue

        duration = int(round(float(duration_value)))
        if duration <= 0:
            continue
        workdays = pd.bdate_range(start=activity_start.normalize(), periods=duration)
        quantity_per_day = float(totals_by_code[code]) / duration
        unit = units_by_code[code]
        for workday in workdays:
            day = pd.Timestamp(workday).normalize()
            if start_date <= day <= end_date:
                planned_by_unit[unit].loc[day] += quantity_per_day

    actual_by_unit = {}
    for unit in unit_labels:
        actual_curve = calcular_curva_s_real_pyrevit(
            df_pyrevit,
            frecuencia="D",
            fecha_corte=end_date,
            unidad=unit,
        )
        actual_by_date = pd.Series(0.0, index=dates)
        for _, row in actual_curve.iterrows():
            day = pd.to_datetime(row["fecha"], errors="coerce")
            if pd.notna(day):
                day = day.normalize()
                if day in actual_by_date.index:
                    actual_by_date.loc[day] += float(
                        row["cantidad_ejecutada_periodo"]
                    )
        actual_by_unit[unit] = actual_by_date

    reports = {}
    for unit, unit_label in unit_labels.items():
        total_quantity = float(
            item_summary.loc[item_summary["Unidad"] == unit, "Cantidad total"].sum()
        )
        scheduled = planned_by_unit[unit]
        actual = actual_by_unit[unit]
        reports[unit] = pd.DataFrame({
            "Fecha": dates,
            "Avance programado (%)": (
                scheduled / total_quantity * 100 if total_quantity > 0 else 0.0
            ),
            "Avance programado (cantidad)": scheduled.to_numpy(),
            "Avance real (%)": (
                actual / total_quantity * 100 if total_quantity > 0 else 0.0
            ),
            "Avance real (cantidad)": actual.to_numpy(),
        })
        reports[unit].attrs["titulo"] = f"Global · Todas las partidas · {unit_label}"
    return reports


def _build_global_excel_export(reports: dict[str, pd.DataFrame]) -> bytes:
    """Construye un libro Excel global con una hoja por unidad."""
    output = io.BytesIO()
    unit_sheets = {
        "m2": "m2", "m3": "m3", "ml": "ml", "kg": "kg", "unid": "unitario"
    }
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for unit, report in reports.items():
            sheet_name = unit_sheets[unit]
            worksheet = writer.book.create_sheet(sheet_name)
            worksheet.merge_cells("A1:E1")
            title = worksheet["A1"]
            title.value = report.attrs["titulo"]
            title.font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
            title.fill = PatternFill(fill_type="solid", fgColor="0A2A62")
            title.alignment = Alignment(horizontal="left", vertical="center")
            worksheet.row_dimensions[1].height = 26
            report.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
                startrow=2,
            )
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes = "A4"
            worksheet.auto_filter.ref = f"A3:E{worksheet.max_row}"
            worksheet.column_dimensions["A"].width = 15
            for column in ("B", "D"):
                worksheet.column_dimensions[column].width = 22
                for cell in worksheet[column][3:]:
                    cell.number_format = '0.00"%"'
            for column in ("C", "E"):
                worksheet.column_dimensions[column].width = 28
                for cell in worksheet[column][3:]:
                    cell.number_format = "#,##0.000"
            for cell in worksheet["A"][3:]:
                cell.number_format = "DD-MM-YYYY"
    return output.getvalue()


def _calcular_kpis_curva_s(
    planned_curve: pd.DataFrame,
    real_curve: pd.DataFrame,
    frequency: str,
    cutoff: date,
) -> dict[str, float | None]:
    """Calcula el porcentaje de avance programado, real y desviación/atraso a la fecha de corte."""
    pct_prog = None
    pct_real = None

    if not planned_curve.empty and "acumulado" in planned_curve.columns:
        if frequency == "W":
            from logic.charts import _current_week_label
            curr_lbl = _current_week_label(planned_curve, cutoff)
            if curr_lbl is not None and "semana_label" in planned_curve.columns:
                match = planned_curve[planned_curve["semana_label"].astype(str) == str(curr_lbl)]
                if not match.empty:
                    pct_prog = float(match.iloc[-1]["acumulado"])
            if pct_prog is None:
                pct_prog = float(planned_curve.iloc[-1]["acumulado"])
        else:
            fechas = pd.to_datetime(planned_curve.get("fecha", []), errors="coerce")
            cutoff_ts = pd.Timestamp(cutoff)
            valid = planned_curve[fechas <= cutoff_ts]
            if not valid.empty:
                pct_prog = float(valid.iloc[-1]["acumulado"])
            elif not planned_curve.empty:
                pct_prog = float(planned_curve.iloc[-1]["acumulado"])

    if not real_curve.empty and "pct_ejecutado_acum" in real_curve.columns:
        if "fecha" in real_curve.columns:
            fechas_r = pd.to_datetime(real_curve["fecha"], errors="coerce", dayfirst=True)
            cutoff_ts = pd.Timestamp(cutoff)
            valid_r = real_curve[fechas_r <= cutoff_ts]
            if not valid_r.empty:
                pct_real = float(valid_r.iloc[-1]["pct_ejecutado_acum"])
            else:
                pct_real = float(real_curve.iloc[-1]["pct_ejecutado_acum"])
        else:
            pct_real = float(real_curve.iloc[-1]["pct_ejecutado_acum"])

    desviacion = (pct_real - pct_prog) if (pct_prog is not None and pct_real is not None) else None
    return {
        "pct_prog": pct_prog,
        "pct_real": pct_real,
        "desviacion": desviacion,
    }


# ════════════════════════════════════════════════════════════════════════════
#  PRESENTACIÓN · helpers de formato, estados y gráficos
#  (todo lo anterior es la lógica original del tablero, sin cambios)
# ════════════════════════════════════════════════════════════════════════════
def _streamlit_at_least(major: int, minor: int) -> bool:
    try:
        parts = st.__version__.split(".")
        return (int(parts[0]), int(parts[1])) >= (major, minor)
    except Exception:
        return False


# Streamlit 1.50+ reemplaza `use_container_width=True` por `width="stretch"`.
_STRETCH = {"width": "stretch"} if _streamlit_at_least(1, 50) else {"use_container_width": True}

PLAN_COLOR = "#6C7B8E"
REAL_COLOR = "#0B6BB5"
INK = "#17212D"
INK_2 = "#536172"
GRID = "#E7EBF0"
LINE_2 = "#B7C1CC"

ESTADOS = (
    ("late", "Atrasada", "Atrasadas"),
    ("risk", "En riesgo", "En riesgo"),
    ("ok", "En plazo", "En plazo"),
    ("done", "Ejecutada", "Ejecutadas"),
    ("sched", "Programada", "Programadas"),
    ("nodata", "Sin dato real", "Sin dato real"),
)
ESTADO_NOMBRE = {key: singular for key, singular, _ in ESTADOS}
ESTADO_PLURAL = {key: plural for key, _, plural in ESTADOS}
ESTADO_ORDEN = {key: index for index, (key, _, _) in enumerate(ESTADOS)}
ESTADO_COLOR = {
    "late": "#B42318", "risk": "#A5620A", "ok": "#2A7D4F",
    "done": "#2A7D4F", "sched": "#6B7A8C", "nodata": "#6B7A8C",
}
TIP_ESTADOS = (
    "Ejecutada: 100 % real. En plazo: desvío mayor que −0,5 %. En riesgo: entre −0,5 y −5 %. "
    "Atrasada: peor que −5 %. Programada: aún sin avance previsto. Sin dato real: faltan elementos del modelo."
)

VIEWS = {
    "Gerencia": "¿Cómo vamos respecto al programa y qué requiere una decisión?",
    "Planificación": "¿Dónde y cuándo nos desviamos, partida por partida?",
    "BIM": "¿Están completos los datos del modelo para medir el avance con confianza?",
}
FRENTES = {
    "Demoliciones": {
        "excel_candidates": DEMOLICIONES_EXCEL_CANDIDATES,
        "pyrevit_candidates": DEMOLICIONES_PYREVIT_CANDIDATES,
        "dashboard_key": "demoliciones",
    },
    "Hormigones": {
        "excel_candidates": HORMIGONES_EXCEL_CANDIDATES,
        "pyrevit_candidates": HORMIGONES_PYREVIT_CANDIDATES,
        "dashboard_key": "hormigones",
    },
    "EEMM": {
        "excel_candidates": EEMM_EXCEL_CANDIDATES,
        "pyrevit_candidates": (),
        "dashboard_key": "eemm",
        "forced_unit": "kg",
        "ifc_directory": EEMM_IFC_DIRECTORY,
    },
    "Electricidad": {
        "excel_candidates": ELECTRICIDAD_EXCEL_CANDIDATES,
        "pyrevit_candidates": ELECTRICIDAD_PYREVIT_CANDIDATES,
        "dashboard_key": "electricidad",
    },
    "Sanitario": {
        "excel_candidates": SANITARIO_EXCEL_CANDIDATES,
        "pyrevit_candidates": SANITARIO_PYREVIT_CANDIDATES,
        "dashboard_key": "sanitario",
    },
    "PCI": {
        "excel_candidates": PCI_EXCEL_CANDIDATES,
        "pyrevit_candidates": PCI_PYREVIT_CANDIDATES,
        "dashboard_key": "pci",
    },
}


def _fmt(value, decimals: int = 1, signed: bool = False) -> str:
    """Formato es-CL: miles con punto, decimales con coma y signo menos tipográfico."""
    if value is None or pd.isna(value):
        return "—"
    rounded = round(abs(float(value)), decimals)
    text = f"{abs(float(value)):,.{decimals}f}".replace(",", "§").replace(".", ",").replace("§", ".")
    if float(value) < 0 and rounded > 0:
        return f"−{text}"
    if signed and float(value) > 0 and rounded > 0:
        return f"+{text}"
    return text


def _esc(value) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def _html(markup: str) -> None:
    """Renderiza HTML sin pasar por el analizador de Markdown."""
    markup = re.sub(r"\n\s*", "", markup or "")
    if not markup.strip():
        return
    if hasattr(st, "html"):
        st.html(markup)
    else:  # Streamlit antiguo
        st.markdown(markup, unsafe_allow_html=True)


def _help(text: str, end: bool = False) -> str:
    cls = "help end" if end else "help"
    return f'<span class="{cls}" tabindex="0" role="img" aria-label="{_esc(text)}" data-tip="{_esc(text)}">?</span>'


def _short_name(name: str) -> str:
    """Quita el prefijo repetitivo de obra (C / F2 / OC / …) y deja lo que distingue a la partida."""
    text = re.sub(r"^\s*\[[^\]]+\]\s*", "", str(name))   # quita la etiqueta de unidad, p. ej. "[M3]"
    parts = [p.strip() for p in text.split("/")]
    parts = [p for p in parts if p and not re.fullmatch(r"C|OC|F\d", p)]
    return " · ".join(parts) or str(name)


def _status_of(deviation: float | None) -> str:
    """Misma regla del tablero original: ±0,05 puntos."""
    if deviation is None:
        return "none"
    if deviation < -0.05:
        return "late"
    if deviation > 0.05:
        return "ahead"
    return "ok"


def _dev_color(deviation: float | None) -> str:
    status = _status_of(deviation)
    return {"late": "var(--late)", "ahead": "var(--ok)", "ok": "var(--ink)"}.get(status, "var(--ink2)")


def _partida_state(real: float | None, prog: float | None) -> str:
    """Estado de una partida a partir de su % real y su % programado al corte (0–100)."""
    if real is None:
        return "sched" if (prog is not None and prog < 0.5) else "nodata"
    if prog is None:
        return "done" if real >= 99.5 else "nodata"
    if real >= 99.5:
        return "done"
    if prog < 0.5 and real < 0.5:
        return "sched"
    deviation = real - prog
    if deviation < -5:
        return "late"
    if deviation < -0.5:
        return "risk"
    return "ok"


def _pill(state: str) -> str:
    return f'<span class="pill s-{state}"><span class="sw"></span>{_esc(ESTADO_NOMBRE.get(state, state))}</span>'


def _stack_html(rows: pd.DataFrame) -> str:
    if rows.empty:
        return ""
    counts = rows["estado"].value_counts().to_dict()
    visibles = [(key, ESTADO_PLURAL[key]) for key, _, _ in ESTADOS if counts.get(key)]
    seg = "".join(
        f'<i class="s-{key}" style="flex:{counts[key]}" title="{counts[key]} {_esc(label.lower())}"></i>'
        for key, label in visibles
    )
    legend = "".join(
        f'<li class="s-{key}"><span class="sw"></span>{counts[key]} {_esc(label.lower())}</li>'
        for key, label in visibles
    )
    return f'<div class="stackbar" role="img" aria-label="Distribución de partidas por estado">{seg}</div><ul class="legend">{legend}</ul>'


# ── Gráficos: paleta nueva sin tocar logic/curva_lb.py ────────────────────
_PALETTE_REMAP = {
    "#2f6fd0": PLAN_COLOR,          # curva_lb: programado
    "#c0392b": REAL_COLOR,          # curva_lb: real (el rojo de alerta va en mayúsculas y no se toca)
    "#0077B6": PLAN_COLOR,          # charts: programado
    "#00A0E3": REAL_COLOR,          # charts: real
    "rgba(192,57,43": "rgba(11,107,181",
    "rgba(0, 119, 182": "rgba(108, 123, 142",
    "rgba(0, 160, 227": "rgba(11, 107, 181",
}


def _theme_fig(fig: go.Figure, height: int | None = None) -> go.Figure:
    """Aplica la paleta y la tipografía del nuevo diseño a una figura ya construida."""
    try:
        raw = pio.to_json(fig)
        for old, new in _PALETTE_REMAP.items():
            raw = raw.replace(old, new)
        fig = pio.from_json(raw)
    except Exception:
        pass
    fig.update_layout(
        font=dict(family="IBM Plex Sans, system-ui, sans-serif", color=INK),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#FFFFFF",
        legend=dict(bgcolor="rgba(255,255,255,0.9)", bordercolor=LINE_2, borderwidth=1),
    )
    fig.update_xaxes(gridcolor=GRID, linecolor=LINE_2)
    fig.update_yaxes(gridcolor=GRID, linecolor=LINE_2)
    if height:
        fig.update_layout(height=height)
    return fig


def _gantt_figure(rows: pd.DataFrame, cutoff: date) -> go.Figure | None:
    """Cronograma por partida: ventana programada, avance real y marca del programado al corte."""
    data = rows.dropna(subset=["inicio", "fin"]).sort_values(["inicio", "codigo"]).reset_index(drop=True)
    if data.empty:
        return None
    labels = []
    for _, row in data.iterrows():
        text = f'{row["codigo"]}  {row["corto"]}'
        labels.append(text if len(text) <= 52 else text[:51] + "…")
    day_ms = 86_400_000
    spans = [((fin - inicio).days + 1) * day_ms for inicio, fin in zip(data["inicio"], data["fin"])]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=labels, x=spans, base=list(data["inicio"]), orientation="h", width=0.55,
        marker=dict(color=GRID, line=dict(color=LINE_2, width=1)), name="Ventana programada",
        customdata=data[["nombre", "pct_prog", "pct_real"]].to_numpy(),
        hovertemplate="<b>%{customdata[0]}</b><br>Programado al corte: %{customdata[1]:.1f}%<br>Real: %{customdata[2]:.1f}%<extra></extra>",
    ))
    real_spans = [span * (0 if pd.isna(value) else value / 100) for span, value in zip(spans, data["pct_real"])]
    fig.add_trace(go.Bar(
        y=labels, x=real_spans, base=list(data["inicio"]), orientation="h", width=0.3,
        marker=dict(color=REAL_COLOR), name="Avance real", hoverinfo="skip",
    ))
    marks = data[data["pct_prog"].notna()]
    if not marks.empty:
        mark_y = [labels[i] for i in marks.index]
        mark_x = [
            inicio + pd.Timedelta(milliseconds=((fin - inicio).days + 1) * day_ms * (prog / 100))
            for inicio, fin, prog in zip(marks["inicio"], marks["fin"], marks["pct_prog"])
        ]
        fig.add_trace(go.Scatter(
            x=mark_x, y=mark_y, mode="markers", name="Programado al corte",
            marker=dict(symbol="line-ns", size=16, line=dict(color=INK, width=2.5), color=INK), hoverinfo="skip",
        ))
    cut = pd.Timestamp(cutoff)
    fig.add_shape(type="line", x0=cut, x1=cut, y0=0, y1=1, yref="paper", line=dict(color=INK_2, width=1, dash="dash"))
    fig.add_annotation(x=cut, y=1, yref="paper", text=f"Corte {cutoff:%d-%m-%Y}", showarrow=False, yanchor="bottom", font=dict(size=11, color=INK_2))
    fig.update_layout(
        barmode="overlay", height=max(260, 70 + 26 * len(data)), margin=dict(l=10, r=20, t=40, b=30),
        legend=dict(orientation="h", y=-0.08, x=0), showlegend=True,
    )
    fig.update_yaxes(autorange="reversed", automargin=True, tickfont=dict(size=11))
    fig.update_xaxes(type="date", tickformat="%d-%b")
    return _theme_fig(fig)


# ── Partidas: avance real, programado al corte, desvío y estado ─────────────
def _build_partida_rows(item_summary: pd.DataFrame, loaded: dict, cutoff: date) -> pd.DataFrame:
    """Une el resumen por código de montaje con el programado al corte de cada actividad.

    No agrega reglas de cálculo: usa `calcular_curva_s_*` y `_calcular_kpis_curva_s` del tablero.
    """
    columns = [
        "codigo", "nombre", "corto", "unidad", "total", "ejecutada", "n", "nd",
        "pct_real", "pct_prog", "desvio", "estado", "inicio", "fin",
    ]
    if item_summary is None or item_summary.empty:
        empty = pd.DataFrame(columns=columns)
        for column in ("total", "ejecutada", "pct_real", "pct_prog", "desvio"):
            empty[column] = pd.to_numeric(empty[column])
        for column in ("inicio", "fin"):
            empty[column] = pd.to_datetime(empty[column])
        return empty

    planned = loaded.get("avance_diario") or loaded.get("avance_semanal")
    frequency = "D" if loaded.get("avance_diario") else "W"
    choices = _activity_choices(planned, loaded.get("df_datos")) if planned else []
    activity_by_code = {c["code"]: c["activity"] for c in choices if c["code"] and c["activity"]}

    schedule: dict[str, tuple] = {}
    datos = loaded.get("df_datos")
    if datos is not None and not datos.empty:
        for _, row in datos.iterrows():
            code = str(row.get("ITEM", "")).strip()
            if code:
                schedule[code] = (
                    pd.to_datetime(row.get("Comienzo"), errors="coerce"),
                    pd.to_datetime(row.get("Fin"), errors="coerce"),
                )

    rows = []
    for _, item in item_summary.iterrows():
        code = str(item["ITEM"]).strip()
        pct_prog = None
        activity = activity_by_code.get(code)
        if planned and activity:
            try:
                curve = (
                    calcular_curva_s_diaria(planned, actividad=activity)
                    if frequency == "D"
                    else calcular_curva_s_semanal(planned, actividad=activity)
                )
                pct_prog = _calcular_kpis_curva_s(curve, pd.DataFrame(), frequency, cutoff)["pct_prog"]
            except Exception:
                pct_prog = None
        pct_real = float(item["pct_ejecutado"])
        start, end = schedule.get(code, (pd.NaT, pd.NaT))
        name = str(item.get("assembly_description") or item.get("Nombre de tarea") or code)
        rows.append({
            "codigo": code,
            "nombre": name,
            "corto": _short_name(name),
            "unidad": item["Unidad"],
            "total": float(item["Cantidad total"]),
            "ejecutada": float(item["Cantidad ejecutada"]),
            "n": int(item["Elementos"]),
            "nd": int(item["Elementos ejecutados"]),
            "pct_real": pct_real,
            "pct_prog": pct_prog,
            "desvio": None if pct_prog is None else pct_real - pct_prog,
            "estado": _partida_state(pct_real, pct_prog),
            "inicio": start,
            "fin": end,
        })
    result = pd.DataFrame(rows, columns=columns)
    for column in ("total", "ejecutada", "pct_real", "pct_prog", "desvio"):
        result[column] = pd.to_numeric(result[column], errors="coerce")
    for column in ("inicio", "fin"):
        result[column] = pd.to_datetime(result[column], errors="coerce")
    return result


# ── Curva S y comparativos: la lógica original, agrupada para poder mostrarla por vista ──
def _curve_bundle(
    *,
    loaded: dict,
    scope: pd.DataFrame,
    model_item_summary: pd.DataFrame,
    item_summary: pd.DataFrame,
    cutoff: date,
    frequency: str,
    planned_data: dict | None,
    activity_choices: list[dict[str, str]],
    chosen_activity: dict | None,
    selected_unit: str | None,
    unit_labels: dict[str, str],
) -> dict:
    """Calcula curvas y cantidades con las mismas llamadas y reglas del tablero original."""
    notes: list[tuple[str, str]] = []
    bundle: dict = {
        "notes": notes, "planned_curve": pd.DataFrame(), "real_curve": pd.DataFrame(),
        "kpis": None, "lb": None, "qty": None, "quantity_comparison": None,
        "curve_scope": scope, "curve_unit": selected_unit, "has_planned": bool(planned_data),
    }
    if scope is None or scope.empty or "assembly_code" not in scope.columns:
        notes.append(("info", "No hay elementos del modelo para calcular la curva en este alcance."))
        return bundle

    unit_codes: set[str] = set()
    curve_scope = scope
    if chosen_activity and chosen_activity["code"]:
        curve_scope = scope[scope["assembly_code"].astype(str).str.strip() == chosen_activity["code"]]
    elif selected_unit:
        unit_codes = set(model_item_summary.loc[
            model_item_summary["Unidad"] == selected_unit, "assembly_code"
        ].astype(str))
        curve_scope = scope[scope["assembly_code"].astype(str).str.strip().isin(unit_codes)]

    curve_unit = selected_unit
    if chosen_activity and chosen_activity["code"]:
        matching_summary = model_item_summary.loc[
            model_item_summary["assembly_code"] == chosen_activity["code"]
        ]
        if not matching_summary.empty:
            curve_unit = matching_summary.iloc[0]["Unidad"]
    bundle["curve_scope"], bundle["curve_unit"] = curve_scope, curve_unit

    real_curve = pd.DataFrame()
    if chosen_activity is not None or selected_unit:
        try:
            real_curve = calcular_curva_s_real_pyrevit(
                curve_scope,
                frecuencia=frequency,
                fecha_corte=pd.Timestamp(cutoff),
                unidad=curve_unit,
            )
        except Exception as exc:
            notes.append(("warning", f"No fue posible calcular la curva real: {exc}"))
    bundle["real_curve"] = real_curve

    if not planned_data:
        return bundle

    activity = chosen_activity["activity"] if chosen_activity else None
    if chosen_activity is None and selected_unit:
        planned_curve, total_bim, scheduled_bim = _build_weighted_global_curve(
            planned_data, activity_choices, model_item_summary, selected_unit, frequency,
        )
        if scheduled_bim < total_bim:
            notes.append((
                "warning",
                f"El Excel no tiene actividades vinculadas para {total_bim - scheduled_bim:,.2f} "
                f"{unit_labels[selected_unit]} del total BIM; esa cantidad queda incluida en el "
                "denominador global, pero no aporta avance programado.",
            ))
    elif chosen_activity is not None and frequency == "W":
        planned_curve = calcular_curva_s_semanal(planned_data, actividad=activity)
    elif chosen_activity is not None:
        planned_curve = calcular_curva_s_diaria(planned_data, actividad=activity)
    else:
        planned_curve = pd.DataFrame()
    bundle["planned_curve"] = planned_curve

    if planned_curve.empty:
        notes.append(("info", "No hay una curva programada disponible para la unidad seleccionada."))
        return bundle

    kpis = _calcular_kpis_curva_s(planned_curve, real_curve, frequency, cutoff)
    bundle["kpis"] = kpis

    real_daily = pd.DataFrame()
    if chosen_activity is not None or selected_unit:
        try:
            real_daily = calcular_curva_s_real_pyrevit(
                curve_scope,
                frecuencia="D",
                fecha_corte=pd.Timestamp(cutoff),
                unidad=curve_unit,
            )
        except Exception as exc:
            notes.append(("warning", f"No fue posible calcular la curva real diaria: {exc}"))
    # En vista semanal se pasa también el calendario diario del Excel: solo sirve
    # para deducir el año de los rangos de semana ("31-ago al 06-sep").
    excel_daily = pd.DataFrame()
    if frequency == "W" and loaded["avance_diario"]:
        try:
            excel_daily = calcular_curva_s_diaria(loaded["avance_diario"])
        except Exception:
            excel_daily = pd.DataFrame()
    lb_base = construir_base_curva(
        planned_curve if frequency == "W" else None,
        planned_curve if frequency == "D" else excel_daily,
        real_daily,
        pd.Timestamp(cutoff),
    )
    if lb_base is None:
        notes.append(("info", "No hay datos programados suficientes para esta vista."))

    scheduled_activities = {}
    if loaded["avance_diario"]:
        scheduled_activities = {
            name: details.copy()
            for name, details in loaded["avance_diario"]["actividades"].items()
        }
    activity_codes = {}
    for choice in activity_choices:
        name = choice["activity"]
        code = choice["code"]
        activity_codes[name] = code
        details = scheduled_activities.setdefault(name, {"item": code})
        if code and loaded["df_datos"] is not None and not loaded["df_datos"].empty:
            schedule_row = loaded["df_datos"].loc[
                loaded["df_datos"]["ITEM"].astype(str).str.strip() == code
            ]
            if not schedule_row.empty:
                row = schedule_row.iloc[0]
                if pd.isna(details.get("inicio")):
                    details["inicio"] = row.get("Comienzo")
                if pd.isna(details.get("dias")):
                    details["dias"] = row.get("N° días hábiles (calc.)")

    if chosen_activity is None and selected_unit:
        selected_activities = {
            choice["activity"] for choice in activity_choices
            if choice["code"] in unit_codes
        }
        scheduled_activities = {
            name: details for name, details in scheduled_activities.items()
            if name in selected_activities
        }
        activity_codes = {
            name: code for name, code in activity_codes.items()
            if name in selected_activities
        }

    comparison_unit = selected_unit
    if chosen_activity and chosen_activity["code"]:
        matching_summary = model_item_summary.loc[
            model_item_summary["assembly_code"] == chosen_activity["code"]
        ]
        if not matching_summary.empty:
            comparison_unit = matching_summary.iloc[0]["Unidad"]
    quantity_comparison = calcular_comparativo_cantidades_periodo(
        planned_curve,
        real_curve,
        curve_scope,
        scheduled_activities,
        activity_codes,
        frequency,
        pd.Timestamp(cutoff),
        actividad=chosen_activity["activity"] if chosen_activity else None,
        unidad=comparison_unit,
    )
    unit_sym = unit_labels.get(comparison_unit, comparison_unit or "")
    if chosen_activity and chosen_activity["code"]:
        code_str = str(chosen_activity["code"]).strip()
        sub_model = model_item_summary.loc[model_item_summary["assembly_code"] == code_str]
        sub_item = item_summary.loc[item_summary["assembly_code"] == code_str]
    else:
        sub_model = model_item_summary.loc[model_item_summary["Unidad"] == comparison_unit]
        sub_item = item_summary.loc[item_summary["Unidad"] == comparison_unit]

    total_cant = float(sub_model["Cantidad total"].sum()) if not sub_model.empty else 0.0
    real_cant = float(sub_item["Cantidad ejecutada"].sum()) if not sub_item.empty else 0.0
    prog_pct = kpis.get("pct_prog") if kpis else None
    prog_cant = total_cant * (prog_pct / 100.0) if (prog_pct is not None and total_cant > 0) else 0.0

    bundle["lb"] = (
        {"base": lb_base, "mode": "sem" if frequency == "W" else "dia"} if lb_base is not None else None
    )
    bundle["quantity_comparison"] = quantity_comparison
    bundle["qty"] = {
        "unit": unit_sym, "total": total_cant, "prog": prog_cant, "real": real_cant,
        "dev": real_cant - prog_cant, "n": len(sub_model),
    }
    return bundle


# ── Resumen ejecutivo ──────────────────────────────────────────────────────
def _kpi_cell(label: str, value: str, unit: str = "", sub: str = "", color: str | None = None,
              tip: str = "", tip_end: bool = False, extra: str = "") -> str:
    style = f' style="color:{color}"' if color else ""
    small = f"<small>{_esc(unit)}</small>" if unit else ""
    help_html = _help(tip, tip_end) if tip else ""
    return (
        f'<div class="kpi"><div class="l">{_esc(label)}{help_html}</div>'
        f'<div class="v"{style}>{value}{small}</div><div class="s">{sub}</div>{extra}</div>'
    )


def _kpi_strip(bundle: dict, rows: pd.DataFrame) -> str:
    kpis, qty = bundle.get("kpis"), bundle.get("qty")
    if not kpis or not qty:
        return ""
    unit = qty["unit"]
    pct_prog, pct_real, deviation = kpis.get("pct_prog"), kpis.get("pct_real"), kpis.get("desviacion")
    status = _status_of(deviation)
    pill = {
        "late": '<span class="pill s-late"><span class="sw"></span>Atrasado</span>',
        "ahead": '<span class="pill s-ok"><span class="sw"></span>Adelantado</span>',
        "ok": '<span class="pill s-ok"><span class="sw"></span>En plazo</span>',
    }.get(status, '<span class="pill s-nodata"><span class="sw"></span>Sin dato real</span>')
    dev_qty = (
        f'<span style="color:{_dev_color(deviation)}">{_fmt(qty["dev"], 0, signed=True)} {_esc(unit)}</span>'
        if deviation is not None else ""
    )
    cells = [
        _kpi_cell("Programado a la fecha", _fmt(pct_prog), "%", f'{_fmt(qty["prog"], 0)} {_esc(unit)}',
                  tip="Avance que la obra debería llevar al corte según el Excel programado, ponderado por cantidad."),
        _kpi_cell("Real ejecutado", _fmt(pct_real), "%", f'{_fmt(qty["real"], 0)} {_esc(unit)}', color="var(--real)",
                  tip="Avance medido en el modelo BIM al corte, ponderado por cantidad."),
        _kpi_cell("Desvío", _fmt(deviation, 1, signed=True), "%", f"{pill} {dev_qty}", color=_dev_color(deviation),
                  tip="Real menos programado, en %. Negativo indica atraso, positivo indica adelanto."),
        _kpi_cell("Partidas por estado", str(len(rows)), "partidas", "", tip=TIP_ESTADOS, tip_end=True,
                  extra=_stack_html(rows)),
        _kpi_cell("Alcance", _fmt(qty["total"], 0), unit, f'{qty["n"]} {"partida" if qty["n"] == 1 else "partidas"} en la vista',
                  tip="Cantidad total BIM del alcance filtrado en la unidad elegida.", tip_end=True),
    ]
    # Segunda fila: mismas magnitudes de avance y atraso, expresadas en la unidad de medida de la partida o vista.
    if deviation is None:
        dev_value, dev_sub, dev_color = "—", "Sin dato real para comparar", "var(--ink2)"
    else:
        dev_value = _fmt(qty["dev"], 2, signed=True)
        dev_sub = {"late": "Atraso a la fecha", "ahead": "Adelanto a la fecha"}.get(status, "En plazo a la fecha")
        dev_color = _dev_color(deviation)
    qty_cells = [
        _kpi_cell("Programado a la fecha · cantidad", _fmt(qty["prog"], 2), unit,
                  "Avance programado en unidad de medida", tip="Cantidad total del alcance × % programado al corte."),
        _kpi_cell("Real ejecutado · cantidad", _fmt(qty["real"], 2), unit,
                  "Avance medido en el modelo BIM", color="var(--real)"),
        _kpi_cell("Atraso / adelanto · cantidad", dev_value, unit, dev_sub, color=dev_color,
                  tip="Cantidad real menos cantidad programada al corte, en la unidad de la partida.", tip_end=True),
    ]
    return (
        f'<div class="kpis">{"".join(cells)}</div>'
        f'<div class="kpis kpis3">{"".join(qty_cells)}</div>'
    )


def _lead_html(bundle: dict, rows: pd.DataFrame, frente: str, chosen_activity: dict | None,
               model_item_summary: pd.DataFrame) -> str:
    kpis, qty = bundle.get("kpis"), bundle.get("qty")
    if not kpis or kpis.get("desviacion") is None:
        return ""
    deviation = kpis["desviacion"]
    status = _status_of(deviation)
    row = None
    if chosen_activity and chosen_activity["code"] and not rows.empty:
        match = rows[rows["codigo"] == chosen_activity["code"]]
        row = match.iloc[0] if not match.empty else None
    subject = f'{row["codigo"]} {row["corto"]}' if row is not None else frente
    if status == "ok":
        first = f"{subject} va en plazo."
    else:
        first = f'{subject} va {_fmt(abs(deviation))} % {"bajo" if status == "late" else "sobre"} lo programado.'
    second = ""
    if row is None:
        late = rows[(rows["desvio"].notna()) & (rows["desvio"] < -0.5)].sort_values("desvio") if not rows.empty else rows
        if not late.empty:
            worst = late.iloc[0]
            second = f'La partida más atrasada es {worst["codigo"]} {worst["corto"]}, con {_fmt(abs(worst["desvio"]))} % de atraso.'
            if len(late) > 1:
                others = len(late) - 1
                second += f' Otras {others} {"partida también está" if others == 1 else "partidas también están"} bajo lo programado.'
        else:
            second = "Ninguna partida está atrasada."
    elif qty and not model_item_summary.empty:
        total_unit = float(model_item_summary.loc[model_item_summary["Unidad"] == row["unidad"], "Cantidad total"].sum())
        if total_unit > 0:
            second = f'Esta partida es el {_fmt(qty["total"] / total_unit * 100, 0)} % de las cantidades de {_esc(qty["unit"])} en {frente}.'
    return f'<div class="concl"><strong>{_esc(first)}</strong> <span>{_esc(second)}</span></div>'


def _critical_html(rows: pd.DataFrame, limit: int = 6) -> str:
    if rows.empty:
        return '<p class="muted">No hay datos reales para evaluar el atraso.</p>'
    critical = rows[rows["estado"].isin(["late", "risk"])].sort_values("desvio").head(limit)
    if critical.empty:
        return '<p class="muted">Ninguna partida está atrasada o en riesgo.</p>'
    items = []
    for _, r in critical.iterrows():
        items.append(
            f'<div class="crit-row"><div class="n"><span class="c">{_esc(r["codigo"])}</span>{_esc(r["corto"])}</div>'
            f'<div class="d">{_fmt(r["desvio"], 1, signed=True)} %</div>'
            f'<div class="bar" role="img" aria-label="Real {_fmt(r["pct_real"], 0)} por ciento, programado {_fmt(r["pct_prog"], 0)} por ciento">'
            f'<i style="width:{r["pct_real"]:.1f}%"></i><b style="left:{(r["pct_prog"] or 0):.1f}%"></b></div></div>'
        )
    return "".join(items)


# ── Ficha de partida ───────────────────────────────────────────────────────
_dialog = getattr(st, "dialog", None) or getattr(st, "experimental_dialog", None)


def _ficha_body(payload: dict) -> None:
    row = payload["row"]
    state = row["estado"]
    unit = payload["unit_label"]
    _html(
        f'<div style="margin-bottom:.6rem"><div style="font-size:1.05rem;font-weight:600">{_esc(row["codigo"])} · {_esc(row["corto"])}</div>'
        f'<div class="muted" style="margin:.2rem 0 .4rem;font-size:.85rem">{_esc(row["nombre"])}</div>{_pill(state)}</div>'
    )
    start = row["inicio"].strftime("%d-%m-%Y") if pd.notna(row["inicio"]) else "—"
    end = row["fin"].strftime("%d-%m-%Y") if pd.notna(row["fin"]) else "—"
    deviation = row["desvio"]
    _html(
        '<div class="kpis" style="grid-template-columns:repeat(3,minmax(0,1fr))">'
        + _kpi_cell("Programado al corte", _fmt(row["pct_prog"]), "%")
        + _kpi_cell("Real al corte", _fmt(row["pct_real"]), "%", color="var(--real)")
        + _kpi_cell("Desvío", _fmt(deviation, 1, signed=True), "%", color=_dev_color(deviation))
        + _kpi_cell("Cantidad ejecutada", f'{_fmt(row["ejecutada"], 0)}', unit, f'de {_fmt(row["total"], 0)} {_esc(unit)}')
        + _kpi_cell("Inicio programado", start)
        + _kpi_cell("Término programado", end)
        + "</div>"
    )
    if payload.get("fig") is not None:
        st.plotly_chart(payload["fig"], **_STRETCH, theme=None)
    breakdown = payload.get("breakdown")
    if breakdown is not None and not breakdown.empty:
        st.markdown("**Avance por zona y subcontratista**")
        st.dataframe(breakdown, **_STRETCH, hide_index=True)


if _dialog is not None:
    try:
        @_dialog("Ficha de partida", width="large")
        def _ficha_dialog(payload: dict) -> None:
            _ficha_body(payload)
    except TypeError:  # versiones sin el parámetro width
        @_dialog("Ficha de partida")
        def _ficha_dialog(payload: dict) -> None:
            _ficha_body(payload)
else:
    def _ficha_dialog(payload: dict) -> None:
        with st.expander("Ficha de partida", expanded=True):
            _ficha_body(payload)


def _ficha_payload(row: pd.Series, *, loaded: dict, scope: pd.DataFrame, cutoff: date,
                   unit_labels: dict[str, str]) -> dict:
    """Arma la curva y el desglose de una partida usando las mismas funciones del tablero."""
    payload: dict = {"row": row, "unit_label": unit_labels.get(row["unidad"], "Kg" if row["unidad"] == "kg" else str(row["unidad"]))}
    code = row["codigo"]
    code_scope = scope[scope["assembly_code"].astype(str).str.strip() == code] if "assembly_code" in scope.columns else scope.iloc[0:0]
    planned = loaded.get("avance_diario")
    if planned:
        choice = next((c for c in _activity_choices(planned, loaded.get("df_datos")) if c["code"] == code), None)
        if choice:
            try:
                planned_curve = calcular_curva_s_diaria(planned, actividad=choice["activity"])
                real_curve = calcular_curva_s_real_pyrevit(
                    code_scope, frecuencia="D", fecha_corte=pd.Timestamp(cutoff), unidad=row["unidad"],
                )
                payload["fig"] = _theme_fig(
                    grafico_curva_s_diaria_plotly(planned_curve, titulo="Curva de la partida", df_real_pyrevit=real_curve),
                    height=360,
                )
            except Exception:
                payload["fig"] = None
    if not code_scope.empty and {"ei_zona", "ei_subcontratista"} <= set(code_scope.columns):
        records = []
        for (zone, contractor), group in code_scope.groupby(["ei_zona", "ei_subcontratista"], dropna=False):
            executed = group[group["ejecutado"].fillna(False).astype(bool)] if "ejecutado" in group else group.iloc[0:0]
            total = float(extraer_cantidad_pyrevit(group, row["unidad"]).sum())
            done = float(extraer_cantidad_pyrevit(executed, row["unidad"]).sum())
            records.append({
                "Zona": zone, "Subcontratista": contractor, "Elementos": len(group),
                "Ejecutado": round(done, 2), "Total": round(total, 2),
                "% avance": round(done / total * 100, 1) if total > 0 else 0.0,
            })
        payload["breakdown"] = pd.DataFrame(records)
    return payload


# ── Secciones de cada vista ────────────────────────────────────────────────
def _render_notes(notes: list[tuple[str, str]]) -> None:
    for kind, message in notes:
        (st.warning if kind == "warning" else st.info)(message)


def _render_curve_panel(bundle: dict, dashboard_key: str, frequency: str, frequency_label: str,
                        cutoff: date, selected_unit: str | None, chosen_activity: dict | None,
                        unit_labels: dict[str, str], compact: bool = False) -> None:
    """Curva S con barras de cantidad (figura original `grafico_curva_lb_plotly`, recoloreada)."""
    with _box():
        _html('<div class="panel-h"><h3>Curva S, programado y real</h3>'
              + _help("Avance acumulado en porcentaje de la cantidad total del alcance filtrado. La caja de desfase marca la diferencia al corte.", True)
              + "</div>")
        lb = bundle.get("lb")
        if lb is not None:
            col_a, col_b = st.columns(2)
            show_bars = col_a.checkbox(
                "Mostrar avance físico (barras de cantidad)", value=True,
                key=f"{dashboard_key}_lb_bars_{frequency}_{'c' if compact else 'p'}",
            )
            show_labels = col_b.checkbox(
                "Mostrar etiquetas de datos", value=not compact,
                key=f"{dashboard_key}_lb_labels_{frequency}_{'c' if compact else 'p'}",
            )
            fig = grafico_curva_lb_plotly(
                lb["base"], modo=lb["mode"], df_cantidades=bundle["quantity_comparison"],
                mostrar_barras=show_bars, etiquetas=show_labels,
            )
            st.plotly_chart(_theme_fig(fig, height=430 if compact else 520), **_STRETCH, theme=None,
                            key=f"{dashboard_key}_lb_chart_{frequency}_{'c' if compact else 'p'}")
        elif not bundle.get("has_planned"):
            st.info("La curva programada aparecerá cuando exista el Excel con la hoja de avance. La curva real BIM está disponible a continuación.")
            title = (
                f"Curva real BIM · {unit_labels[selected_unit]} · {frequency_label.lower()} · al {cutoff:%d-%m-%Y}"
                if chosen_activity is None and selected_unit
                else f"Curva real BIM · {frequency_label.lower()} · al {cutoff:%d-%m-%Y}"
            )
            st.plotly_chart(_theme_fig(grafico_curva_real_pyrevit_plotly(bundle["real_curve"], titulo=title)),
                            **_STRETCH, theme=None,
                            key=f"{dashboard_key}_real_curve_{frequency}_{'c' if compact else 'p'}")
        st.caption("La curva programada se obtiene del Excel fijo del proyecto; la real, de la exportación PyRevit.")


def _box():
    try:
        return st.container(border=True)
    except TypeError:  # Streamlit antiguo
        return st.container()


def _render_followup_table(rows: pd.DataFrame, *, dashboard_key: str, loaded: dict, scope: pd.DataFrame,
                           cutoff: date, unit_labels: dict[str, str]) -> None:
    """Tabla de seguimiento por partida: estados filtrables, orden por columna y ficha."""
    with _box():
        _html(f'<div class="panel-h"><h3>Seguimiento por partida</h3><span>{len(rows)} partidas, de mayor atraso a mayor adelanto</span></div>')
        if rows.empty:
            st.info("No hay partidas con elementos del modelo en este alcance.")
            return
        counts = rows["estado"].value_counts().to_dict()
        options = ["all"] + [k for k, _, _ in ESTADOS if counts.get(k)]
        if st.session_state.get(f"{dashboard_key}_estado", "all") not in options:
            st.session_state.pop(f"{dashboard_key}_estado", None)
        chosen = st.radio(
            "Filtrar por estado", options, horizontal=True, key=f"{dashboard_key}_estado",
            format_func=lambda k: f"Todas ({len(rows)})" if k == "all" else f"{ESTADO_PLURAL[k]} ({counts[k]})",
        )
        view = rows if chosen == "all" else rows[rows["estado"] == chosen]
        view = view.assign(_orden=view["estado"].map(ESTADO_ORDEN)).sort_values(["desvio", "codigo"], na_position="last").reset_index(drop=True)
        table = pd.DataFrame({
            "Código": view["codigo"],
            "Partida": view["corto"],
            "Estado": view["estado"].map(ESTADO_NOMBRE),
            "Programado (%)": view["pct_prog"],
            "Real (%)": view["pct_real"],
            "Desvío (%)": view["desvio"],
            "Ejecutado": view["ejecutada"],
            "Total": view["total"],
            "Unidad": view["unidad"].replace({"kg": "Kg"}),
            "Inicio": view["inicio"].dt.strftime("%d-%m-%Y"),
            "Término": view["fin"].dt.strftime("%d-%m-%Y"),
        })

        def _color_state(value: str) -> str:
            key = next((k for k, name in ESTADO_NOMBRE.items() if name == value), "nodata")
            return f"color:{ESTADO_COLOR[key]};font-weight:600"

        def _color_dev(value) -> str:
            if value is None or pd.isna(value):
                return ""
            return "color:#B42318;font-weight:600" if value < -0.05 else ("color:#2A7D4F;font-weight:600" if value > 0.05 else "")

        styler = table.style
        paint = getattr(styler, "map", None) or styler.applymap  # `map` desde pandas 2.1
        styler = paint(_color_state, subset=["Estado"])
        paint = getattr(styler, "map", None) or styler.applymap
        styler = paint(_color_dev, subset=["Desvío (%)"])
        config = {
            "Programado (%)": st.column_config.NumberColumn(format="%.1f"),
            "Real (%)": st.column_config.ProgressColumn("Real (%)", min_value=0, max_value=100, format="%.1f"),
            "Desvío (%)": st.column_config.NumberColumn(format="%+.1f"),
            "Ejecutado": st.column_config.NumberColumn(format="%.2f"),
            "Total": st.column_config.NumberColumn(format="%.2f"),
        }
        selected_index = None
        try:
            event = st.dataframe(
                styler, **_STRETCH, hide_index=True, column_config=config,
                on_select="rerun", selection_mode="single-row", key=f"{dashboard_key}_tbl",
            )
            picked = event["selection"]["rows"] if event is not None else []
            selected_index = picked[0] if picked else None
        except TypeError:  # Streamlit sin selección de filas
            st.dataframe(styler, **_STRETCH, hide_index=True, column_config=config)

        # Selector explícito: funciona en cualquier versión y permite reabrir la ficha.
        col_a, col_b = st.columns([4, 1])
        labels = [f'{r["codigo"]} · {r["corto"]}' for _, r in view.iterrows()]
        default_index = selected_index if selected_index is not None and selected_index < len(labels) else 0
        pick = col_a.selectbox("Ficha de partida", labels, index=default_index if labels else None,
                               key=f"{dashboard_key}_ficha_pick") if labels else None
        open_clicked = col_b.button("Abrir ficha", key=f"{dashboard_key}_ficha_open", **_STRETCH, disabled=not labels)

        last_key = f"{dashboard_key}_ficha_last"
        target = None
        if open_clicked and pick is not None:
            target = labels.index(pick)
        elif selected_index is not None and st.session_state.get(last_key) != view.iloc[selected_index]["codigo"]:
            target = selected_index
        if target is not None:
            row = view.iloc[target]
            st.session_state[last_key] = row["codigo"]
            _ficha_dialog(_ficha_payload(row, loaded=loaded, scope=scope, cutoff=cutoff, unit_labels=unit_labels))


def _render_exports(*, loaded: dict, scope: pd.DataFrame, model_item_summary: pd.DataFrame,
                    dashboard_key: str, forced_unit: str | None, unit_labels: dict[str, str]) -> None:
    """Exportaciones a Excel: misma lógica del tablero original."""
    with _box():
        _html('<div class="panel-h"><h3>Exportar evolución diaria de partidas</h3></div>')
        if not loaded["avance_diario"]:
            st.info("La exportación requiere la hoja Avance Diario del Excel programado.")
            return
        export_choices = _activity_choices(loaded["avance_diario"], loaded["df_datos"])
        export_codes = {choice["activity"]: choice["code"] for choice in export_choices}
        export_schedules = {
            name: details.copy()
            for name, details in loaded["avance_diario"]["actividades"].items()
        }
        for choice in export_choices:
            code = choice["code"]
            details = export_schedules.setdefault(choice["activity"], {"item": code})
            if not code or loaded["df_datos"] is None:
                continue
            schedule_row = loaded["df_datos"].loc[
                loaded["df_datos"]["ITEM"].astype(str).str.strip() == code
            ]
            if schedule_row.empty:
                continue
            row = schedule_row.iloc[0]
            if pd.isna(details.get("inicio")):
                details["inicio"] = row.get("Comienzo")
            if pd.isna(details.get("dias")):
                details["dias"] = row.get("N° días hábiles (calc.)")

        try:
            export_start = pd.Timestamp("2026-08-24")
            export_end = pd.Timestamp(date.today())
            daily_reports = generar_reporte_diario_partidas(
                scope, export_schedules, export_codes, export_start, export_end,
            )
            global_reports = _build_global_daily_reports(
                scope, model_item_summary, export_schedules, export_codes, export_start, export_end,
                unit_labels=({forced_unit: unit_labels[forced_unit]} if forced_unit else None),
            )
            download_columns = st.columns(2)
            with download_columns[0]:
                if daily_reports:
                    st.download_button(
                        "Descargar Excel con avance diario de partidas",
                        data=_build_excel_export(daily_reports, export_codes),
                        file_name=f"avance_diario_partidas_2026-08-24_al_{date.today():%Y-%m-%d}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"{dashboard_key}_download_daily_progress",
                    )
                    st.caption(
                        f"El libro incluye {len(daily_reports)} partidas con cantidades BIM, desde el "
                        "24-08-2026 hasta hoy. Los porcentajes corresponden al avance diario respecto "
                        "de la cantidad total de cada partida."
                    )
                else:
                    st.info("No hay partidas con cantidad BIM y actividad programada para exportar.")
            with download_columns[1]:
                if global_reports:
                    st.download_button(
                        "Descargar Excel Global",
                        data=_build_global_excel_export(global_reports),
                        file_name=f"avance_global_todas_las_partidas_2026-08-24_al_{date.today():%Y-%m-%d}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"{dashboard_key}_download_global_daily_progress",
                    )
                    st.caption(
                        f"Incluye hojas separadas para {', '.join(unit_labels.values())}. Los porcentajes "
                        "diarios se calculan sobre la cantidad BIM total de cada unidad."
                    )
                else:
                    st.info("No hay cantidades BIM disponibles para generar el avance global.")
        except Exception as exc:
            st.error(f"No fue posible generar el Excel de avance diario: {exc}")


def _schedule_leaf_codes(activity_choices: list[dict[str, str]], datos: pd.DataFrame | None) -> set[str]:
    """Partidas hoja del cronograma: las que llevan etiqueta de unidad ([M2], [M3], [UN]) en el Excel.

    Si el Excel no usa etiquetas (p. ej. EEMM), se toman los códigos que no tienen subcódigos.
    """
    codes = {c["code"] for c in activity_choices if c["code"]}
    if datos is not None and not datos.empty:
        tagged = {
            str(row.get("ITEM", "")).strip()
            for _, row in datos.iterrows()
            if str(row.get("ITEM", "")).strip() in codes
            and re.match(r"^\s*\[[^\]]+\]", str(row.get("Nombre de tarea", "")))
        }
        if tagged:
            return tagged
    return {c for c in codes if not any(o != c and o.startswith(c + ".") for o in codes)}


def _bim_quality(scope: pd.DataFrame, rows_all: pd.DataFrame, schedule_codes: set[str],
                 loaded: dict, item_summary: pd.DataFrame) -> None:
    """Calidad de los datos del modelo: lo que el equipo BIM debe completar para medir bien."""
    if scope is None or scope.empty:
        st.info("No hay elementos del modelo en este alcance.")
        return
    total = len(scope)
    executed = int(scope["ejecutado"].fillna(False).astype(bool).sum()) if "ejecutado" in scope.columns else 0
    contractor = scope["ei_subcontratista"].fillna("").astype(str).str.strip() if "ei_subcontratista" in scope.columns else pd.Series([""] * total)
    no_contractor = int(contractor.isin(["", "Sin Asignar"]).sum())
    zone = scope["ei_zona"].fillna("").astype(str).str.strip() if "ei_zona" in scope.columns else pd.Series([""] * total)
    no_zone = int(zone.isin(["", "General"]).sum())
    linked_codes = set(item_summary["ITEM"].astype(str)) if not item_summary.empty else set()
    unlinked = sorted(schedule_codes - linked_codes)

    cells = [
        _kpi_cell("Elementos modelados", _fmt(total, 0), "", f"{_fmt(executed / total * 100 if total else 0)} % ejecutado"),
        _kpi_cell("Elementos ejecutados", _fmt(executed, 0), "", f"{_fmt(total - executed, 0)} pendientes", color="var(--real)"),
        _kpi_cell("Partidas con elementos BIM", str(len(linked_codes & schedule_codes) if schedule_codes else len(linked_codes)),
                  f"de {len(schedule_codes)}" if schedule_codes else "", "del cronograma",
                  tip="Partidas del Excel que tienen al menos un elemento con su código de montaje en el modelo."),
        _kpi_cell("Sin subcontratista", _fmt(no_contractor, 0), "", f"{_fmt(no_contractor / total * 100 if total else 0)} % de los elementos",
                  tip='Elementos cuyo subcontratista figura como "Sin Asignar" o vacío.', tip_end=True),
        _kpi_cell("Fuera de un frente", _fmt(no_zone, 0), "", 'zona "General" o vacía',
                  tip="Elementos sin frente de obra (FRENTE_1, 2 o 3) asignado en el modelo.", tip_end=True),
    ]
    _html(f'<div class="kpis">{"".join(cells)}</div>')
    problems = []
    if no_contractor:
        problems.append(f"{_fmt(no_contractor, 0)} elementos sin subcontratista")
    if no_zone:
        problems.append(f"{_fmt(no_zone, 0)} sin frente")
    if unlinked:
        problems.append(f"{len(unlinked)} partidas del cronograma sin elementos")
    head = "Hay datos por completar en el modelo." if problems else "Los datos del modelo están completos."
    _html(f'<div class="concl"><strong>{head}</strong> <span>{("Pendientes: " + ", ".join(problems) + ".") if problems else ""}</span></div>')

    left, right = st.columns(2)
    with left:
        with _box():
            _html('<div class="panel-h"><h3>Elementos por modelo de origen</h3></div>')
            model_col = "modelo_origen" if "modelo_origen" in scope.columns else ("archivo_ifc" if "archivo_ifc" in scope.columns else None)
            if model_col:
                by_model = scope.assign(_ej=scope["ejecutado"].fillna(False).astype(bool) if "ejecutado" in scope.columns else False) \
                    .groupby(model_col).agg(Elementos=("_ej", "size"), Ejecutados=("_ej", "sum")).reset_index()
                by_model["Avance (%)"] = (by_model["Ejecutados"] / by_model["Elementos"] * 100).round(1)
                by_model = by_model.rename(columns={model_col: "Modelo"}).sort_values("Elementos", ascending=False)
                st.dataframe(by_model, **_STRETCH, hide_index=True, column_config={
                    "Avance (%)": st.column_config.ProgressColumn("Avance (%)", min_value=0, max_value=100, format="%.1f")})
            else:
                st.caption("Los datos no traen el modelo de origen.")
    with right:
        with _box():
            _html('<div class="panel-h"><h3>Elementos por zona y subcontratista</h3></div>')
            if {"ei_zona", "ei_subcontratista"} <= set(scope.columns):
                work = scope.assign(_ej=scope["ejecutado"].fillna(False).astype(bool) if "ejecutado" in scope.columns else False)
                by_zone = work.groupby(["ei_zona", "ei_subcontratista"], dropna=False).agg(
                    Elementos=("_ej", "size"), Ejecutados=("_ej", "sum")).reset_index()
                by_zone["Avance (%)"] = (by_zone["Ejecutados"] / by_zone["Elementos"] * 100).round(1)
                by_zone = by_zone.rename(columns={"ei_zona": "Zona", "ei_subcontratista": "Subcontratista"})
                st.dataframe(by_zone, **_STRETCH, hide_index=True, column_config={
                    "Avance (%)": st.column_config.ProgressColumn("Avance (%)", min_value=0, max_value=100, format="%.1f")})
            else:
                st.caption("Los datos no traen zona o subcontratista.")

    with _box():
        _html(f'<div class="panel-h"><h3>Partidas del cronograma sin elementos en el modelo</h3><span>{len(unlinked)} de {len(schedule_codes)} partidas</span></div>')
        if unlinked:
            datos = loaded.get("df_datos")
            records = []
            for code in unlinked:
                name, start, end = "", pd.NaT, pd.NaT
                if datos is not None and not datos.empty:
                    match = datos[datos["ITEM"].astype(str).str.strip() == code]
                    if not match.empty:
                        name = str(match.iloc[0].get("Nombre de tarea", "")).strip()
                        start = pd.to_datetime(match.iloc[0].get("Comienzo"), errors="coerce")
                        end = pd.to_datetime(match.iloc[0].get("Fin"), errors="coerce")
                records.append({
                    "Código": code, "Partida": _short_name(name),
                    "Inicio": start.strftime("%d-%m-%Y") if pd.notna(start) else "—",
                    "Término": end.strftime("%d-%m-%Y") if pd.notna(end) else "—",
                })
            st.dataframe(pd.DataFrame(records), **_STRETCH, hide_index=True)
        else:
            st.caption("Todas las partidas del cronograma tienen elementos en el modelo.")


def _render_item_table(item_summary: pd.DataFrame, dashboard_key: str) -> None:
    """Tabla y gráfico "Partidas / códigos" del tablero original, ahora dentro de la vista BIM."""
    if item_summary.empty:
        st.info("No hay datos agrupables por Código de montaje.")
        return
    selected_codes = st.multiselect(
        "Filtrar códigos de montaje",
        options=item_summary["ITEM"].tolist(),
        default=[],
        help="Sin selección se muestran todos los códigos.",
        key=f"{dashboard_key}_item_codes",
    )
    visible_items = item_summary[item_summary["ITEM"].isin(selected_codes)] if selected_codes else item_summary
    table = visible_items[[
        "ITEM", "assembly_description", "Unidad", "Cantidad total", "Cantidad ejecutada",
        "pct_ejecutado", "Elementos", "Elementos ejecutados",
    ]].rename(columns={"assembly_description": "Descripción", "pct_ejecutado": "% Avance"})
    table["Unidad"] = table["Unidad"].replace({"kg": "Kg"})
    st.dataframe(
        table, **_STRETCH, hide_index=True,
        column_config={
            "% Avance": st.column_config.ProgressColumn("% Avance", min_value=0, max_value=100, format="%.1f%%"),
            "Cantidad total": st.column_config.NumberColumn("Cantidad total", format="%.2f"),
            "Cantidad ejecutada": st.column_config.NumberColumn("Cantidad ejecutada", format="%.2f"),
        },
    )
    st.plotly_chart(
        _theme_fig(grafico_avance_por_item_plotly(visible_items, titulo="Avance por código de montaje")),
        **_STRETCH, theme=None, key=f"{dashboard_key}_item_progress",
    )


def _render_ifc_viewer(*, dashboard_key: str, ifc_records: pd.DataFrame, loaded_ifc_models: list,
                       selected_zone: str, selected_contractor: str, excel_codes: set) -> None:
    """Visor IFC del tablero original (misma lógica), ubicado en la vista BIM."""
    viewer_scope = ifc_records.copy()
    if selected_zone != "Todas" and "ei_zona" in viewer_scope.columns:
        viewer_scope = viewer_scope[viewer_scope["ei_zona"].astype(str) == selected_zone]
    if selected_contractor != "Todos" and "ei_subcontratista" in viewer_scope.columns:
        viewer_scope = viewer_scope[viewer_scope["ei_subcontratista"].astype(str) == selected_contractor]
    if excel_codes:
        linked_elements = viewer_scope[viewer_scope["assembly_code"].astype(str).str.strip().isin(excel_codes)]
        if not linked_elements.empty:
            viewer_scope = linked_elements

    st.markdown("#### Visor IFC · elementos ejecutados en azul")
    st.caption("El visor respeta los filtros de zona y subcontratista seleccionados arriba.")
    if len(loaded_ifc_models) > 1:
        module_options = [path.name for path, _ in loaded_ifc_models]
        selected_3d_modules = st.multiselect(
            "Modelos activos en la vista 3D",
            options=module_options,
            default=module_options,
            help="Permite prender o apagar modelos individuales en el visor 3D.",
            key=f"{dashboard_key}_ifc_3d_modules",
        )
        if selected_3d_modules:
            viewer_scope = viewer_scope[viewer_scope["archivo_ifc"].isin(selected_3d_modules)]
        else:
            st.info("Selecciona al menos un modelo para visualizar en 3D.")

    available_codes = sorted(
        code for code in viewer_scope["assembly_code"].dropna().astype(str).str.strip().unique() if code
    )
    visible_codes = st.multiselect(
        "Filtrar códigos de montaje en el modelo",
        options=available_codes,
        default=[],
        help="Sin selección se muestran todos los elementos del alcance filtrado.",
        key=f"{dashboard_key}_ifc_codes",
    )
    if visible_codes:
        viewer_scope = viewer_scope[viewer_scope["assembly_code"].astype(str).str.strip().isin(visible_codes)]

    executed_count = int(viewer_scope["ejecutado"].fillna(False).astype(bool).sum())
    total_weight = pd.to_numeric(viewer_scope["kg"], errors="coerce").fillna(0).sum()
    metric_columns = st.columns(3)
    metric_columns[0].metric("Elementos visibles", f"{len(viewer_scope):,}")
    metric_columns[1].metric("Elementos ejecutados", f"{executed_count:,}")
    metric_columns[2].metric("Peso del alcance", f"{total_weight:,.2f} Kg")
    missing_weights = int(viewer_scope["kg"].isna().sum())
    if missing_weights:
        st.warning(f"{missing_weights} elementos del alcance no tienen un peso numérico en la propiedad Weigth.")
    if not available_codes:
        st.info(
            "Los modelos IFC seleccionados aún no tienen el parámetro Código de montaje. "
            "Los elementos se visualizarán en gris hasta agregar los parámetros de avance."
        )
    st.caption(
        "Los elementos marcados como ejecutados se resaltan en azul; "
        "los restantes se muestran en gris. Usa el mouse para rotar y acercar el modelo."
    )
    if st.checkbox("Cargar geometría 3D del IFC", value=False, key=f"{dashboard_key}_load_ifc_geometry"):
        try:
            models_with_geometry = []
            total_missing_geometry = 0
            active_files = [
                path for path, _ in loaded_ifc_models
                if not viewer_scope.empty and (viewer_scope["archivo_ifc"] == path.name).any()
            ]
            with st.spinner(f"Procesando geometría 3D de {len(active_files)} modelo(s) IFC..."):
                for ifc_file in active_files:
                    file_scope = viewer_scope[viewer_scope["archivo_ifc"] == ifc_file.name]
                    file_visible_ids = set(pd.to_numeric(file_scope["ifc_id"], errors="coerce").dropna().astype(int))
                    if not file_visible_ids:
                        continue
                    model_with_geom = _load_ifc_data(
                        str(ifc_file), ifc_file.stat().st_mtime,
                        include_geometry=True, visible_ids=tuple(sorted(file_visible_ids)),
                    )
                    models_with_geometry.append((model_with_geom, file_visible_ids))
                    if model_with_geom.get("elements_without_geometry"):
                        total_missing_geometry += model_with_geom["elements_without_geometry"]

            from logic.ifc_viewer import build_ifc_figure

            if models_with_geometry:
                figure = build_ifc_figure(models_with_geometry)
                if figure.data:
                    st.plotly_chart(figure, **_STRETCH, key=f"{dashboard_key}_ifc_viewer")
                else:
                    st.info("No hay geometría visible para la combinación actual de filtros.")
                if total_missing_geometry:
                    st.warning(f"{total_missing_geometry} elementos IFC no tienen geometría renderizable.")
            else:
                st.info("No hay elementos visibles para cargar geometría.")
        except Exception as exc:
            st.error(f"No fue posible visualizar los modelos IFC: {exc}")
    else:
        st.info("Activa la casilla para procesar y mostrar la geometría 3D de los modelos seleccionados.")


# ════════════════════════════════════════════════════════════════════════════
#  PANEL POR FRENTE
# ════════════════════════════════════════════════════════════════════════════
def render_dashboard(
    excel_candidates: tuple[Path, ...],
    pyrevit_candidates: tuple[Path, ...],
    dashboard_key: str,
    view: str,
    forced_unit: str | None = None,
    ifc_directory: Path | None = None,
) -> None:
    cutoff = date.today()
    frente_name = next((n for n, c in FRENTES.items() if c["dashboard_key"] == dashboard_key), dashboard_key)
    excel_file = _first_existing(excel_candidates)
    ifc_files = (
        tuple(sorted(ifc_directory.glob("*.ifc"), key=lambda path: path.name.casefold()))
        if ifc_directory and ifc_directory.is_dir()
        else ()
    )
    selected_ifc_files: list[Path] = []
    if ifc_directory:
        if not ifc_files:
            st.error(f"No se encontraron modelos IFC en {ifc_directory}.")
            return

        # Siempre se leen todos los modelos IFC de la carpeta; no hay selección manual.
        selected_ifc_files = list(ifc_files)

    pyrevit_files = tuple(
        path for path in pyrevit_candidates if path.is_file() and path.stat().st_size > 0
    )
    loaded = _load_local_data(
        str(excel_file) if excel_file else None,
        excel_file.stat().st_mtime if excel_file else 0.0,
        tuple(str(path) for path in pyrevit_files),
        tuple(path.stat().st_mtime for path in pyrevit_files),
        require_pyrevit=ifc_directory is None,
    )
    loaded_ifc_models: list[tuple[Path, dict]] = []
    if selected_ifc_files:
        for ifc_file in selected_ifc_files:
            try:
                model_data = _load_ifc_data(str(ifc_file), ifc_file.stat().st_mtime, include_geometry=False)
                loaded_ifc_models.append((ifc_file, model_data))
            except Exception as exc:
                st.warning(f"No fue posible leer el modelo IFC {ifc_file.name}: {exc}")

    all_records = []
    for ifc_file, model_data in loaded_ifc_models:
        for rec in model_data.get("records", []):
            if "archivo_ifc" not in rec:
                rec["archivo_ifc"] = ifc_file.name
            all_records.append(rec)

    ifc_records = (
        pd.DataFrame(
            all_records,
            columns=[
                "ifc_id", "global_id", "assembly_code", "assembly_description",
                "nombre", "kg", "cantidad_total", "unidad", "ejecutado",
                "ei_zona", "ei_subcontratista", "ei_semana",
                "ei_planificacion_obra", "ei_fecha_ejecucion", "fecha",
                "archivo_ifc",
            ],
        )
        if all_records
        else pd.DataFrame()
    )
    for error in loaded["errores"]:
        st.warning(error)

    if loaded["df_datos"] is None or "ITEM" not in loaded["df_datos"].columns:
        st.error("El Excel asignado a este frente no contiene datos BIM con una columna ITEM.")
        return

    excel_codes = set(loaded["df_datos"]["ITEM"].dropna().astype(str).str.strip())
    if not excel_codes and ifc_directory is None:
        st.info("El Excel asignado a este frente no contiene códigos ITEM para vincular con BIM.")
        return

    raw_pyrevit = ifc_records if ifc_directory is not None else loaded["df_pyrevit"]
    if (raw_pyrevit is None or raw_pyrevit.empty) and ifc_directory is None:
        st.error("No hay datos BIM para mostrar. La aplicación no tiene carga manual: verifica la exportación automática en la carpeta data.")
        _html('<div class="note">Incluye una exportación PyRevit exclusiva para este frente en una de las rutas indicadas. '
              'Cada frente consulta su propio archivo BIM.</div>')
        with st.expander("Rutas BIM revisadas"):
            st.code("\n".join(str(path.relative_to(ROOT_DIR)) for path in pyrevit_candidates))
        return

    if raw_pyrevit is None:
        raw_pyrevit = pd.DataFrame()
    if "assembly_code" not in raw_pyrevit.columns and ifc_directory is None:
        st.error("La exportación encontrada no contiene el campo de Código de montaje (assembly_code).")
        return

    # Sin filtros de zona ni subcontratista: el alcance es el total del frente.
    filters_box = _box()
    selected_zone, selected_contractor = "Todas", "Todos"

    if "assembly_code" not in raw_pyrevit.columns:
        # Sin registros IFC cargados (p. ej. un modelo que no se pudo leer): alcance vacío, sin error.
        raw_pyrevit = raw_pyrevit.assign(assembly_code=pd.Series(dtype="object"))
    if not loaded_ifc_models and ifc_directory is None:
        raw_pyrevit = raw_pyrevit[raw_pyrevit["assembly_code"].astype(str).str.strip().isin(excel_codes)].copy()
    elif excel_codes:
        raw_pyrevit = raw_pyrevit[raw_pyrevit["assembly_code"].astype(str).str.strip().isin(excel_codes)].copy()
    else:
        raw_pyrevit = raw_pyrevit.iloc[0:0].copy()
    if raw_pyrevit.empty and not loaded_ifc_models and ifc_directory is None:
        st.info("La exportación BIM de este frente no tiene códigos presentes en su Excel.")
        return
    if raw_pyrevit.empty and loaded_ifc_models:
        st.info(
            "Los IFC seleccionados aún no contienen Códigos de montaje vinculados al Excel. "
            "El visor se puede usar; agrega los parámetros BIM para habilitar los filtros y avances."
        )

    scope = raw_pyrevit.copy()

    as_of = _filtered_as_of(scope, cutoff)
    model_item_summary = _build_item_summary(scope, unit_override=forced_unit)
    item_summary = _build_item_summary(scope, progress_as_of=as_of, unit_override=forced_unit)

    # ── Controles de la curva (no se usan en la vista BIM) ──────────────────
    unit_labels = {"m2": "m²", "m3": "m³", "ml": "ml", "unid": "unidades"}
    if forced_unit:
        unit_labels = {forced_unit: ("Kg" if forced_unit == "kg" else unit_labels[forced_unit])}

    frequency_label, frequency = "Semanal", "W"
    chosen_activity, selected_unit = None, None
    notes: list[tuple[str, str]] = []
    if view != "BIM":
        with filters_box:
            col_freq, col_activity, col_unit = st.columns([2, 5, 2])
            with col_freq:
                frequency_label = st.radio("Frecuencia", ["Semanal", "Diaria"], horizontal=True, key=f"{dashboard_key}_frequency")
            frequency = "W" if frequency_label == "Semanal" else "D"
            planned_data = loaded["avance_semanal"] if frequency == "W" else loaded["avance_diario"]
            if planned_data:
                activity_choices = _activity_choices(planned_data, loaded["df_datos"])
            else:
                activity_choices = [
                    {"item": row["ITEM"], "task": row["Nombre de tarea"], "activity": "", "code": row["ITEM"]}
                    for _, row in item_summary.iterrows()
                ]
            with col_activity:
                chosen_activity = st.selectbox(
                    "Partida o vista global",
                    options=[None] + activity_choices,
                    format_func=lambda option: (
                        "Global · Todas las partidas" if option is None
                        else f"{option['item']} · {_short_name(option['task'])}"
                    ),
                    key=f"{dashboard_key}_activity_{frequency}",
                )
            if chosen_activity is None:
                with col_unit:
                    available_units = list(unit_labels)
                    default_unit = forced_unit or "m2"
                    if not forced_unit and not model_item_summary.empty:
                        totals_by_unit = model_item_summary.groupby("Unidad")["Cantidad total"].sum()
                        totals_by_unit = totals_by_unit[totals_by_unit.index.isin(available_units)]
                        if not totals_by_unit.empty and totals_by_unit.max() > 0:
                            default_unit = str(totals_by_unit.idxmax())
                    selected_unit = st.selectbox(
                        "Unidad de la curva global",
                        options=available_units,
                        index=available_units.index(default_unit),
                        format_func=lambda unit: unit_labels[unit],
                        key=f"{dashboard_key}_global_unit_{frequency}",
                        help="Cada curva global considera solo partidas de la unidad elegida.",
                    )
                unit_total = (
                    float(model_item_summary.loc[model_item_summary["Unidad"] == selected_unit, "Cantidad total"].sum())
                    if not model_item_summary.empty else 0.0
                )
                if unit_total <= 0:
                    notes.append((
                        "warning",
                        f"Los datos BIM del filtro actual no contienen cantidades mayores que cero "
                        f"clasificadas como {unit_labels[selected_unit]}; no se puede calcular esa curva. "
                        "Cambia la unidad de la curva global.",
                    ))
    else:
        planned_data = loaded["avance_diario"] or loaded["avance_semanal"]
        activity_choices = _activity_choices(planned_data, loaded["df_datos"]) if planned_data else []

    rows = _build_partida_rows(item_summary, loaded, cutoff)

    # ── Vista: BIM ──────────────────────────────────────────────────────────
    if view == "BIM":
        if ifc_directory is not None and (scope is None or scope.empty):
            _html('<div class="note"><b>Este frente aún no tiene elementos con Código de montaje vinculados al Excel.</b> '
                  "Agrega los parámetros BIM a los IFC para medir el avance real.</div>")
        _bim_quality(scope, rows, _schedule_leaf_codes(activity_choices, loaded["df_datos"]), loaded, item_summary)
        with st.expander("Avance por código de montaje", expanded=False):
            _render_item_table(item_summary, dashboard_key)
        if loaded_ifc_models:
            with st.expander("Visor IFC", expanded=False):
                _render_ifc_viewer(
                    dashboard_key=dashboard_key, ifc_records=ifc_records, loaded_ifc_models=loaded_ifc_models,
                    selected_zone=selected_zone, selected_contractor=selected_contractor, excel_codes=excel_codes,
                )
        return

    # ── Vistas: Gerencia y Planificación ────────────────────────────────────
    bundle = _curve_bundle(
        loaded=loaded, scope=scope, model_item_summary=model_item_summary, item_summary=item_summary,
        cutoff=cutoff, frequency=frequency, planned_data=planned_data, activity_choices=activity_choices,
        chosen_activity=chosen_activity, selected_unit=selected_unit, unit_labels=unit_labels,
    )
    _render_notes(notes + bundle["notes"])
    _html(_kpi_strip(bundle, rows))
    _html(_lead_html(bundle, rows, frente_name, chosen_activity, model_item_summary))

    if view == "Gerencia":
        left, right = st.columns([3, 2])
        with left:
            _render_curve_panel(bundle, dashboard_key, frequency, frequency_label, cutoff, selected_unit,
                                chosen_activity, unit_labels, compact=True)
        with right:
            with _box():
                _html('<div class="panel-h"><h3>Partidas críticas</h3><span>de mayor a menor atraso</span></div>')
                _html(_critical_html(rows))
                st.caption("Abre la ficha de cada partida desde la vista Planificación.")
        return

    # Planificación
    _render_curve_panel(bundle, dashboard_key, frequency, frequency_label, cutoff, selected_unit,
                        chosen_activity, unit_labels)
    gantt = _gantt_figure(rows, cutoff) if not rows.empty else None
    with _box():
        _html('<div class="panel-h"><h3>Cronograma por partida</h3>'
              "<span>Barra clara: ventana programada. Azul: avance real. Raya: programado al corte.</span></div>")
        if gantt is not None:
            st.plotly_chart(gantt, **_STRETCH, theme=None, key=f"{dashboard_key}_gantt")
        else:
            st.caption("No hay fechas de inicio y término en el Excel para dibujar el cronograma.")
    _render_followup_table(rows, dashboard_key=dashboard_key, loaded=loaded, scope=scope, cutoff=cutoff, unit_labels=unit_labels)
    _render_exports(loaded=loaded, scope=scope, model_item_summary=model_item_summary,
                    dashboard_key=dashboard_key, forced_unit=forced_unit, unit_labels=unit_labels)


# ════════════════════════════════════════════════════════════════════════════
#  PÁGINA
# ════════════════════════════════════════════════════════════════════════════
_html(
    '<div class="nav-brand"><h1>Boulevard Gastronómico Mall Plaza El Trébol</h1>'
    "<p>Seguimiento y control del avance BIM</p></div>"
)
nav_view_col, nav_frente_col = st.columns([3, 2])
with nav_view_col:
    # Vista fija: Planificación (texto, sin selector). Gerencia y BIM quedan en el código sin mostrarse.
    selected_view = "Planificación"
    st.markdown('<div class="nav-h">Vista</div><div class="nav-txt">Planificación</div>', unsafe_allow_html=True)
with nav_frente_col:
    st.markdown('<div class="nav-h">Frente de obra</div>', unsafe_allow_html=True)
    selected_frente = st.radio("Frente de obra", list(FRENTES), key="nav_frente", horizontal=True, label_visibility="collapsed")

_html(
    f'<div class="vhead"><div><h2>{_esc(selected_view)}</h2>'
    f"<p>{_esc(VIEWS[selected_view])} Frente: {_esc(selected_frente)}.</p></div>"
    f'<div class="cut">Corte al <b>{date.today():%d-%m-%Y}</b></div></div>'
)

render_dashboard(**FRENTES[selected_frente], view=selected_view)

st.markdown(
    '<div class="footer">© Desarrollado por Daniel Salcedo Salcedo 2026 · Coordinador BIM · '
    "Datos consultados automáticamente desde las fuentes locales del proyecto.</div>",
    unsafe_allow_html=True,
)
