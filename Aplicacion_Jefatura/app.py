"""Dashboard ejecutivo de control de avance BIM, sin carga manual de archivos."""

from __future__ import annotations

import html
import re
import sys
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

ROOT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT_DIR))

from logic.charts import (  # noqa: E402
    grafico_avance_por_item_plotly,
    grafico_curva_real_pyrevit_plotly,
    grafico_curva_s_diaria_plotly,
    grafico_curva_s_semanal_plotly,
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
    calcular_curva_s_diaria,
    calcular_curva_s_real_pyrevit,
    calcular_curva_s_semanal,
    extraer_unidad_y_cantidad_de_desc,
)

st.set_page_config(
    page_title="Jefatura | Control BIM",
    page_icon="🔷",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&display=swap');
    :root {
        --ei-blue:#00A0E3; --ei-blue-dark:#0077B6; --ei-navy:#0A2A62;
        --panel:#F7F9FC; --border:#E6E9EE; --text:#2C3338;
        --sub:#6B7280; --green:#2E9E6B; --amber:#D4A017; --red:#C0392B;
    }
    html, body, [data-testid="stAppViewContainer"], .stApp {
        background:#fff !important; color:var(--text) !important;
        font-family:'Montserrat',system-ui,sans-serif !important;
    }
    .block-container {max-width:1320px; padding-top:.6rem; padding-bottom:2rem;}
    [data-testid="stSidebar"] {background:#fff; border-right:1px solid var(--border);}
    [data-testid="stMetric"] {border:1px solid var(--border); border-radius:4px; padding:14px 16px; background:#fff;}
    [data-testid="stMetricLabel"] {color:var(--sub)!important; font-size:.72rem; font-weight:600; letter-spacing:.06em; text-transform:uppercase;}
    [data-testid="stMetricValue"] {color:var(--ei-navy)!important; font-weight:700;}
    [data-testid="stTabs"] button[aria-selected="true"] {color:var(--ei-blue)!important; border-bottom-color:var(--ei-blue)!important;}
    .jefatura-banner {padding:30px 34px; margin:8px 0 22px; color:#fff;
        background:linear-gradient(115deg,#0A2A62 0%,#1264A3 58%,#00A0E3 100%);}
    .jefatura-banner h1 {margin:0;color:#fff;font-size:1.8rem;font-weight:800;letter-spacing:.02em;}
    .jefatura-banner p {margin:8px 0 0;color:#EAF6FC;font-size:.9rem;}
    .section-title {font-size:.95rem;font-weight:700;color:var(--ei-navy);letter-spacing:.07em;
        text-transform:uppercase;border-bottom:2px solid var(--ei-blue);padding-bottom:8px;margin:1.4rem 0 .9rem;}
    .info-box {background:var(--panel);border:1px solid var(--border);border-left:3px solid var(--ei-blue);
        padding:12px 16px;margin:8px 0 14px;color:var(--sub);font-size:.85rem;}
    .source-pill {display:inline-block;padding:5px 10px;background:#EAF7FC;color:#0077B6;
        font-size:.72rem;font-weight:700;letter-spacing:.04em;text-transform:uppercase;}
    .footer {margin-top:2rem;padding:18px 4px;border-top:1px solid var(--border);color:#89929C;font-size:.72rem;}
    hr {border-color:var(--border)!important;}
    </style>
    """,
    unsafe_allow_html=True,
)

EXCEL_CANDIDATES = (
    ROOT_DIR / "data" / "programado" / "Demoliciones_Avance.xlsx",
    ROOT_DIR / "Demoliciones_Avance.xlsx",
)
PYREVIT_CANDIDATES = (
    ROOT_DIR / "data" / "pyrevit_elementos_detallados.json",
    ROOT_DIR / "Aplicacion_pyRevit" / "data" / "pyrevit_elementos_detallados.json",
    ROOT_DIR / "data" / "pyrevit_elementos_detallados.csv",
)


def _first_existing(paths: tuple[Path, ...]) -> Path | None:
    return next((path for path in paths if path.is_file() and path.stat().st_size > 0), None)


@st.cache_data(show_spinner=False)
def _load_local_data(
    excel_path: str | None,
    excel_mtime: float,
    pyrevit_paths: tuple[str, ...],
    pyrevit_mtimes: tuple[float, ...],
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
            result["df_datos"] = load_datos_sheet(excel_path)
            if "Avance Semanal" in hojas:
                result["avance_semanal"] = load_avance_semanal(excel_path)
            if "Avance Diario" in hojas:
                result["avance_diario"] = load_avance_diario(excel_path)
        except Exception as exc:
            result["errores"].append(f"No se pudo leer el Excel programado: {exc}")

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


def _build_item_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Resume el avance por Assembly Code, respetando la unidad de cada montaje."""
    if df.empty or "assembly_code" not in df.columns:
        return pd.DataFrame()

    rows = []
    for code, group in df.groupby("assembly_code", dropna=True):
        code_text = str(code).strip()
        if not code_text or code_text.lower() == "nan":
            continue

        description = ""
        if "assembly_description" in group.columns:
            modes = group["assembly_description"].dropna().astype(str).mode()
            description = modes.iloc[0] if not modes.empty else ""

        executed = group["ejecutado"].fillna(False).astype(bool) if "ejecutado" in group else pd.Series(False, index=group.index)
        done = group.loc[executed]
        m2_total = pd.to_numeric(group.get("m2", pd.Series(0, index=group.index)), errors="coerce").fillna(0).sum()
        m3_total = pd.to_numeric(group.get("m3", pd.Series(0, index=group.index)), errors="coerce").fillna(0).sum()
        ml_total = pd.to_numeric(group.get("ml", pd.Series(0, index=group.index)), errors="coerce").fillna(0).sum()
        m2_done = pd.to_numeric(done.get("m2", pd.Series(0, index=done.index)), errors="coerce").fillna(0).sum()
        m3_done = pd.to_numeric(done.get("m3", pd.Series(0, index=done.index)), errors="coerce").fillna(0).sum()
        ml_done = pd.to_numeric(done.get("ml", pd.Series(0, index=done.index)), errors="coerce").fillna(0).sum()

        unit, total = extraer_unidad_y_cantidad_de_desc(
            description, m2_total, m3_total, ml_total, default_qty=float(len(group))
        )
        _, quantity_done = extraer_unidad_y_cantidad_de_desc(
            description, m2_done, m3_done, ml_done, default_qty=float(len(done))
        )
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

    return pd.DataFrame(rows).sort_values("pct_ejecutado", ascending=True) if rows else pd.DataFrame()


def _normalize_task_name(value: object) -> str:
    """Normaliza nombres para enlazar las actividades programadas con Excel."""
    text = "" if value is None or pd.isna(value) else str(value).strip()
    text = re.sub(r"^\[[^\]]+\]\s*", "", text)
    return re.sub(r"\s+", " ", text).casefold()


def _activity_choices(planned_data: dict, df_datos: pd.DataFrame | None) -> list[dict[str, str]]:
    """Arma opciones de actividad mostrando ITEM y nombre de tarea."""
    activities = planned_data.get("actividades", {})
    rows = df_datos.to_dict("records") if df_datos is not None and not df_datos.empty else []
    choices = []

    for activity_name, details in activities.items():
        scheduled_item = str(details.get("item", "")).strip()
        match = next(
            (row for row in rows if scheduled_item and str(row.get("ITEM", "")).strip() == scheduled_item),
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

        item_value = match.get("ITEM", "") if match else scheduled_item
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


excel_file = _first_existing(EXCEL_CANDIDATES)
pyrevit_files = tuple(
    path for path in PYREVIT_CANDIDATES if path.is_file() and path.stat().st_size > 0
)
loaded = _load_local_data(
    str(excel_file) if excel_file else None,
    excel_file.stat().st_mtime if excel_file else 0.0,
    tuple(str(path) for path in pyrevit_files),
    tuple(path.stat().st_mtime for path in pyrevit_files),
)
pyrevit_file = Path(loaded["pyrevit_path"]) if loaded["pyrevit_path"] else None

with st.sidebar:
    st.markdown("<div style='font-size:.68rem;font-weight:700;letter-spacing:.14em;color:#00A0E3;text-transform:uppercase'>Datos del proyecto</div>", unsafe_allow_html=True)
    st.markdown("### Panel")
    proyecto = st.text_input("Proyecto", value="Demoliciones OC")
    etapa = st.text_input("Etapa / Frente", value="Estructura")
    cutoff = st.date_input("Fecha de corte", value=date.today(), format="DD-MM-YYYY")
    st.markdown("---")
    st.caption("Este panel consulta automáticamente los archivos locales del proyecto. No permite cargar archivos desde la pantalla.")
    if pyrevit_file:
        st.success(f"PyRevit: {pyrevit_file.name}")
    else:
        st.warning("No se encontró la exportación detallada de PyRevit.")
    if excel_file:
        st.success(f"Programación: {excel_file.name}")
    else:
        st.info("Sin Excel programado; se mostrarán los datos BIM disponibles.")

safe_project = html.escape(proyecto)
safe_stage = html.escape(etapa)
st.markdown(
    f"""<div class="jefatura-banner">
        <div class="source-pill">Dashboard ejecutivo · Control BIM</div>
        <h1>Seguimiento de avance del proyecto</h1>
        <p>Proyecto: <b>{safe_project}</b> &nbsp;|&nbsp; Frente: <b>{safe_stage}</b>
        &nbsp;|&nbsp; Corte: <b>{cutoff.strftime('%d-%m-%Y')}</b></p>
    </div>""",
    unsafe_allow_html=True,
)

for error in loaded["errores"]:
    st.warning(error)

raw_pyrevit = loaded["df_pyrevit"]
if raw_pyrevit is None or raw_pyrevit.empty:
    st.error("No hay datos BIM para mostrar. La aplicación no tiene carga manual: verifica la exportación automática en la carpeta data.")
    st.markdown(
        '<div class="info-box">Incluye y confirma en GitHub una exportación con registros en '
        '<b>data/pyrevit_elementos_detallados.json</b>, '
        '<b>Aplicacion_pyRevit/data/pyrevit_elementos_detallados.json</b> o '
        '<b>data/pyrevit_elementos_detallados.csv</b>. Luego reinicia la app para que Streamlit Cloud '
        'descargue el último commit.</div>',
        unsafe_allow_html=True,
    )
    with st.expander("Rutas BIM revisadas"):
        st.code("\n".join(str(path.relative_to(ROOT_DIR)) for path in PYREVIT_CANDIDATES))
    st.stop()

if "assembly_code" not in raw_pyrevit.columns:
    st.error("La exportación encontrada no contiene el campo de Código de montaje (assembly_code).")
    st.stop()

# Filtros de lectura del panel: zona y subcontratista.
filter_cols = st.columns(2)
with filter_cols[0]:
    zone_values = sorted(raw_pyrevit["ei_zona"].dropna().astype(str).unique()) if "ei_zona" in raw_pyrevit else []
    selected_zone = st.selectbox("Zona", ["Todas"] + zone_values)
with filter_cols[1]:
    contractor_values = sorted(raw_pyrevit["ei_subcontratista"].dropna().astype(str).unique()) if "ei_subcontratista" in raw_pyrevit else []
    selected_contractor = st.selectbox("Subcontratista", ["Todos"] + contractor_values)

scope = raw_pyrevit.copy()
if selected_zone != "Todas" and "ei_zona" in scope.columns:
    scope = scope[scope["ei_zona"].astype(str) == selected_zone]
if selected_contractor != "Todos" and "ei_subcontratista" in scope.columns:
    scope = scope[scope["ei_subcontratista"].astype(str) == selected_contractor]

as_of = _filtered_as_of(scope, cutoff)
item_summary = _build_item_summary(as_of)
st.markdown('<div class="section-title">Análisis ejecutivo</div>', unsafe_allow_html=True)
tab_curve, tab_items = st.tabs(["Curva de avance", "Partidas / códigos"])

with tab_curve:
    frequency_label = st.radio("Frecuencia", ["Semanal", "Diaria"], horizontal=True, key="jefatura_frequency")
    frequency = "W" if frequency_label == "Semanal" else "D"
    planned_data = loaded["avance_semanal"] if frequency == "W" else loaded["avance_diario"]
    if planned_data:
        activity_choices = _activity_choices(planned_data, loaded["df_datos"])
    else:
        activity_choices = [
            {
                "item": row["ITEM"],
                "task": row["Nombre de tarea"],
                "activity": "",
                "code": row["ITEM"],
            }
            for _, row in item_summary.iterrows()
        ]
    chosen_activity = st.selectbox(
        "Partida · ITEM y nombre de tarea",
        options=[None] + activity_choices,
        format_func=lambda option: (
            "Global · Todas las partidas"
            if option is None
            else f"{option['item']} · {option['task']}"
        ),
        key=f"jefatura_activity_{frequency}",
    )

    curve_scope = scope
    if chosen_activity and chosen_activity["code"]:
        curve_scope = scope[
            scope["assembly_code"].astype(str).str.strip() == chosen_activity["code"]
        ]
    try:
        real_curve = calcular_curva_s_real_pyrevit(
            curve_scope,
            frecuencia=frequency,
            fecha_corte=pd.Timestamp(cutoff),
        )
    except Exception as exc:
        real_curve = pd.DataFrame()
        st.warning(f"No fue posible calcular la curva real: {exc}")

    if planned_data:
        activity = chosen_activity["activity"] if chosen_activity else None
        if frequency == "W":
            planned_curve = calcular_curva_s_semanal(planned_data, actividad=activity)
            chart = grafico_curva_s_semanal_plotly(
                planned_curve,
                titulo="Curva S · Programado vs. real",
                df_real_pyrevit=real_curve,
            )
        else:
            planned_curve = calcular_curva_s_diaria(planned_data, actividad=activity)
            chart = grafico_curva_s_diaria_plotly(
                planned_curve,
                titulo="Curva diaria · Programado vs. real",
                df_real_pyrevit=real_curve,
            )
        st.plotly_chart(chart, use_container_width=True)
        st.caption("La curva programada se obtiene del Excel fijo del proyecto; la real, de la exportación PyRevit.")
    else:
        if not excel_file:
            st.info("La curva programada aparecerá cuando exista el Excel en data/programado/Demoliciones_Avance.xlsx. La curva real BIM está disponible a continuación.")
        else:
            st.info("El Excel encontrado no incluye una hoja de avance para esta frecuencia.")
        st.plotly_chart(
            grafico_curva_real_pyrevit_plotly(
                real_curve,
                titulo=f"Curva real BIM · {frequency_label.lower()} · al {cutoff.strftime('%d-%m-%Y')}",
            ),
            use_container_width=True,
        )

with tab_items:
    if item_summary.empty:
        st.info("No hay datos agrupables por Código de montaje.")
    else:
        selected_codes = st.multiselect(
            "Filtrar códigos de montaje",
            options=item_summary["ITEM"].tolist(),
            default=[],
            help="Sin selección se muestran todos los códigos.",
        )
        visible_items = item_summary[item_summary["ITEM"].isin(selected_codes)] if selected_codes else item_summary
        table = visible_items[[
            "ITEM", "assembly_description", "Unidad", "Cantidad total", "Cantidad ejecutada",
            "pct_ejecutado", "Elementos", "Elementos ejecutados",
        ]].rename(columns={"assembly_description": "Descripción", "pct_ejecutado": "% Avance"})
        st.dataframe(
            table,
            use_container_width=True,
            hide_index=True,
            column_config={
                "% Avance": st.column_config.ProgressColumn("% Avance", min_value=0, max_value=100, format="%.1f%%"),
                "Cantidad total": st.column_config.NumberColumn("Cantidad total", format="%.2f"),
                "Cantidad ejecutada": st.column_config.NumberColumn("Cantidad ejecutada", format="%.2f"),
            },
        )
        st.plotly_chart(
            grafico_avance_por_item_plotly(visible_items, titulo="Avance por código de montaje"),
            use_container_width=True,
        )

st.markdown(
    '<div class="footer">© Desarrollado por Daniel Salcedo Salcedo 2026-Coordinador BIM · Datos consultados automáticamente desde las fuentes locales del proyecto.</div>',
    unsafe_allow_html=True,
)
