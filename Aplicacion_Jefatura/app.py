"""Dashboard ejecutivo de control de avance BIM, sin carga manual de archivos."""

from __future__ import annotations

import html
import io
import re
import sys
from datetime import date
from pathlib import Path

import pandas as pd
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
        unit = detectar_unidad_pyrevit(group)
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

    return pd.DataFrame(rows).sort_values("pct_ejecutado", ascending=True) if rows else pd.DataFrame()


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
) -> dict[str, pd.DataFrame]:
    """Genera avances diarios globales separados por unidad BIM."""
    start_date = pd.Timestamp(fecha_inicio).normalize()
    end_date = pd.Timestamp(fecha_fin).normalize()
    if start_date > end_date:
        raise ValueError("La fecha inicial del reporte no puede ser posterior a la fecha final.")

    dates = pd.date_range(start_date, end_date, freq="D")
    unit_labels = {"m2": "m²", "m3": "m³", "ml": "ml", "unid": "unitario"}
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
    unit_sheets = {"m2": "m2", "m3": "m3", "ml": "ml", "unid": "unitario"}
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
model_item_summary = _build_item_summary(scope)
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
        "Seleccione partida o vista global",
        options=[None] + activity_choices,
        format_func=lambda option: (
            "Global · Todas las partidas"
            if option is None
            else f"{option['item']} · {option['task']}"
        ),
        key=f"jefatura_activity_{frequency}",
    )

    unit_labels = {"m2": "m²", "m3": "m³", "ml": "ml", "unid": "unidades"}
    selected_unit = None
    if chosen_activity is None:
        available_units = list(unit_labels)
        selected_unit = st.selectbox(
            "Unidad de la curva global",
            options=available_units,
            index=available_units.index("m2"),
            format_func=lambda unit: unit_labels[unit],
            key=f"jefatura_global_unit_{frequency}",
            help="Cada curva global considera solo partidas de la unidad elegida.",
        )
        unit_total = (
            float(model_item_summary.loc[
                model_item_summary["Unidad"] == selected_unit, "Cantidad total"
            ].sum())
            if not model_item_summary.empty
            else 0.0
        )
        if unit_total <= 0:
            st.warning(
                f"La exportación BIM no contiene cantidades clasificadas como "
                f"{unit_labels[selected_unit]}; no se puede calcular esa curva."
            )

    curve_scope = scope
    if chosen_activity and chosen_activity["code"]:
        curve_scope = scope[
            scope["assembly_code"].astype(str).str.strip() == chosen_activity["code"]
        ]
    elif selected_unit:
        unit_codes = set(model_item_summary.loc[
            model_item_summary["Unidad"] == selected_unit, "assembly_code"
        ].astype(str))
        curve_scope = scope[
            scope["assembly_code"].astype(str).str.strip().isin(unit_codes)
        ]

    curve_unit = selected_unit
    if chosen_activity and chosen_activity["code"]:
        matching_summary = model_item_summary.loc[
            model_item_summary["assembly_code"] == chosen_activity["code"]
        ]
        if not matching_summary.empty:
            curve_unit = matching_summary.iloc[0]["Unidad"]

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
            st.warning(f"No fue posible calcular la curva real: {exc}")

    if planned_data:
        activity = chosen_activity["activity"] if chosen_activity else None
        if chosen_activity is None and selected_unit:
            planned_curve, total_bim, scheduled_bim = _build_weighted_global_curve(
                planned_data,
                activity_choices,
                model_item_summary,
                selected_unit,
                frequency,
            )
            if scheduled_bim < total_bim:
                st.warning(
                    f"El Excel no tiene actividades vinculadas para "
                    f"{total_bim - scheduled_bim:,.2f} {unit_labels[selected_unit]} "
                    "del total BIM; esa cantidad queda incluida en el denominador "
                    "global, pero no aporta avance programado."
                )
        elif chosen_activity is not None and frequency == "W":
            planned_curve = calcular_curva_s_semanal(planned_data, actividad=activity)
        elif chosen_activity is not None:
            planned_curve = calcular_curva_s_diaria(planned_data, actividad=activity)
        else:
            planned_curve = pd.DataFrame()

        if planned_curve.empty:
            st.info("No hay una curva programada disponible para la unidad seleccionada.")
        else:
            if frequency == "W":
                chart = grafico_curva_s_semanal_plotly(
                    planned_curve,
                    titulo=(
                        f"Curva S · {unit_labels[selected_unit]} · Programado vs. real"
                        if chosen_activity is None and selected_unit
                        else "Curva S · Programado vs. real"
                    ),
                    df_real_pyrevit=real_curve,
                )
            else:
                chart = grafico_curva_s_diaria_plotly(
                    planned_curve,
                    titulo=(
                        f"Curva diaria · {unit_labels[selected_unit]} · Programado vs. real"
                        if chosen_activity is None and selected_unit
                        else "Curva diaria · Programado vs. real"
                    ),
                    df_real_pyrevit=real_curve,
                )
            st.plotly_chart(chart, use_container_width=True)
        st.caption("La curva programada se obtiene del Excel fijo del proyecto; la real, de la exportación PyRevit.")
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

        if not planned_curve.empty:
            comparison_unit = selected_unit
            if chosen_activity and chosen_activity["code"]:
                matching_summary = model_item_summary.loc[
                    model_item_summary["assembly_code"]
                    == chosen_activity["code"]
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
            st.markdown("#### Avance por período · Cantidad programada vs. real")
            if quantity_comparison.empty:
                st.info("No hay cantidades comparables para el período y la unidad seleccionados.")
            else:
                st.plotly_chart(
                    grafico_comparativo_cantidades_plotly(quantity_comparison),
                    use_container_width=True,
                )
    else:
        if not excel_file:
            st.info("La curva programada aparecerá cuando exista el Excel en data/programado/Demoliciones_Avance.xlsx. La curva real BIM está disponible a continuación.")
        else:
            st.info("El Excel encontrado no incluye una hoja de avance para esta frecuencia.")
        st.plotly_chart(
            grafico_curva_real_pyrevit_plotly(
                real_curve,
                titulo=(
                    f"Curva real BIM · {unit_labels[selected_unit]} · "
                    f"{frequency_label.lower()} · al {cutoff.strftime('%d-%m-%Y')}"
                    if chosen_activity is None and selected_unit
                    else f"Curva real BIM · {frequency_label.lower()} · al "
                    f"{cutoff.strftime('%d-%m-%Y')}"
                ),
            ),
            use_container_width=True,
        )

    st.markdown("#### Exportar evolución diaria de partidas")
    if loaded["avance_diario"]:
        export_choices = _activity_choices(
            loaded["avance_diario"], loaded["df_datos"]
        )
        export_codes = {
            choice["activity"]: choice["code"] for choice in export_choices
        }
        export_schedules = {
            name: details.copy()
            for name, details in loaded["avance_diario"]["actividades"].items()
        }
        for choice in export_choices:
            code = choice["code"]
            details = export_schedules.setdefault(
                choice["activity"], {"item": code}
            )
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
                scope,
                export_schedules,
                export_codes,
                export_start,
                export_end,
            )
            global_reports = _build_global_daily_reports(
                scope,
                model_item_summary,
                export_schedules,
                export_codes,
                export_start,
                export_end,
            )
            download_columns = st.columns(2)
            with download_columns[0]:
                if daily_reports:
                    workbook = _build_excel_export(daily_reports, export_codes)
                    st.download_button(
                        "Descargar Excel con avance diario de partidas",
                        data=workbook,
                        file_name=(
                            f"avance_diario_partidas_2026-08-24_al_"
                            f"{date.today():%Y-%m-%d}.xlsx"
                        ),
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        key="jefatura_download_daily_progress",
                    )
                    st.caption(
                        f"El libro incluye {len(daily_reports)} partidas con cantidades BIM, "
                        "desde el 24-08-2026 hasta hoy. Los porcentajes corresponden al "
                        "avance diario respecto de la cantidad total de cada partida."
                    )
                else:
                    st.info("No hay partidas con cantidad BIM y actividad programada para exportar.")
            with download_columns[1]:
                if global_reports:
                    global_workbook = _build_global_excel_export(global_reports)
                    st.download_button(
                        "Descargar Excel Global",
                        data=global_workbook,
                        file_name=(
                            f"avance_global_todas_las_partidas_2026-08-24_al_"
                            f"{date.today():%Y-%m-%d}.xlsx"
                        ),
                        mime=(
                            "application/vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        key="jefatura_download_global_daily_progress",
                    )
                    st.caption(
                        "Incluye hojas separadas para m², m³, ml y unitario. "
                        "Los porcentajes diarios se calculan sobre la cantidad BIM "
                        "total de cada unidad."
                    )
                else:
                    st.info("No hay cantidades BIM disponibles para generar el avance global.")
        except Exception as exc:
            st.error(f"No fue posible generar el Excel de avance diario: {exc}")
    else:
        st.info("La exportación requiere la hoja Avance Diario del Excel programado.")

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
