# logic/charts.py
# Gráficos Matplotlib para Curvas S – con integración PyRevit por Assembly Code
# Desarrollado por Daniel Salcedo - Coordinador BIM

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as mticker
import matplotlib.dates as mdates
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ── Paleta BIM ─────────────────────────────────────────────────────────────
C = {
    "bg":         "#FFFFFF",
    "panel":      "#F7F9FC",
    "grid":       "#E6E9EE",
    "programado": "#0077B6",   # Azul corporativo EI – curva programada Excel
    "real":       "#00A0E3",   # Cian EI – curva real PyRevit (ejecutado)
    "variacion":  "#D4A017",   # Ámbar – diferencia/variación
    "alerta":     "#C0392B",   # Rojo  – crítico / atraso
    "purple":     "#0A2A62",   # Navy EI – acento / 100%
    "text":       "#2C3338",
    "subtext":    "#6B7280",
    "border":     "#C5CDD6",
}


def _apply_style(fig, *axes):
    fig.patch.set_facecolor(C["bg"])
    for ax in axes:
        ax.set_facecolor(C["panel"])
        ax.tick_params(colors=C["subtext"], labelsize=9)
        ax.xaxis.label.set_color(C["subtext"])
        ax.yaxis.label.set_color(C["subtext"])
        ax.title.set_color(C["text"])
        for sp in ax.spines.values():
            sp.set_edgecolor(C["border"])
        ax.grid(True, color=C["grid"], linewidth=0.6, alpha=0.8, linestyle="--")
        ax.set_axisbelow(True)


def _watermark(fig):
    fig.text(0.99, 0.01,
             "Desarrollado por Daniel Salcedo – Coordinador BIM",
             ha="right", va="bottom", fontsize=7.5,
             color=C["subtext"], alpha=0.6, style="italic")


# ─────────────────────────────────────────────────────────────────────────────
# CURVA S SEMANAL – Programado (Excel) vs Real (PyRevit)
# ─────────────────────────────────────────────────────────────────────────────

def grafico_curva_s_semanal(df_prog: pd.DataFrame,
                             titulo: str = "Curva S – Avance Semanal",
                             df_real_pyrevit: pd.DataFrame = None) -> plt.Figure:
    """
    df_prog: DataFrame con semana_label, acumulado (% programado del Excel)
    df_real_pyrevit: DataFrame de calcular_curva_s_real_pyrevit()
                     con columnas: fecha, pct_ejecutado_acum
    """
    fig, ax = plt.subplots(figsize=(12, 5.5))
    _apply_style(fig, ax)

    x        = list(range(len(df_prog)))
    x_labels = df_prog["semana_label"].tolist()
    y_prog   = df_prog["acumulado"].tolist()

    # Área + línea programada
    ax.fill_between(x, y_prog, alpha=0.08, color=C["programado"])
    ax.plot(x, y_prog,
            color=C["programado"], linewidth=2.5,
            marker="o", markersize=6,
            markerfacecolor=C["bg"], markeredgecolor=C["programado"],
            markeredgewidth=2, label="Programado % (Excel)", zorder=5)

    # Etiquetas en puntos programados
    for i, v in enumerate(y_prog):
        if v > 0:
            ax.annotate(f"{v:.1f}%", (i, v),
                        textcoords="offset points", xytext=(0, 10),
                        ha="center", fontsize=7.5,
                        color=C["programado"], fontweight="bold")

    ax2 = None
    # Curva real PyRevit (ejecutado acumulado %)
    if df_real_pyrevit is not None and not df_real_pyrevit.empty \
            and "pct_ejecutado_acum" in df_real_pyrevit.columns:

        # Mapear puntos reales al eje X de semanas alineado correctamente
        x_real_lbls, y_real = _alinear_real_con_semanas(df_prog, df_real_pyrevit)
        if x_real_lbls:
            x_indices = [x_labels.index(l) for l in x_real_lbls if l in x_labels]
            y_real_sub = y_real[:len(x_indices)]

            # Curva real PyRevit
            ax.fill_between(x_indices, y_real_sub, alpha=0.09, color=C["real"])
            ax.plot(x_indices, y_real_sub,
                    color=C["real"], linewidth=2.5,
                    marker="s", markersize=5,
                    markerfacecolor=C["bg"], markeredgecolor=C["real"],
                    markeredgewidth=2, label="Real % Ejecutado (BIM)",
                    linestyle="--", zorder=5)

    ax.axhline(100, color=C["subtext"], linewidth=0.8, linestyle=":", alpha=0.6)
    ax.set_xticks(x)
    ax.set_xticklabels(x_labels, rotation=30, ha="right", fontsize=8.5)
    ax.set_ylim(-5, 115)
    ax.set_ylabel("Avance Acumulado %", fontsize=10)
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.0f%%"))
    ax.set_title(titulo, fontsize=14, fontweight="bold", color=C["text"], pad=15)
    ax.legend(loc="upper left", fontsize=9,
              facecolor=C["panel"], edgecolor=C["border"],
              labelcolor=C["text"], framealpha=0.9)
    _watermark(fig)
    plt.tight_layout(pad=1.5)
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# CURVA S DIARIA – Programado (Excel) vs Real (PyRevit)
# ─────────────────────────────────────────────────────────────────────────────

def grafico_curva_s_diaria(df_prog: pd.DataFrame,
                            titulo: str = "Curva S – Avance Diario",
                            df_real_pyrevit: pd.DataFrame = None) -> plt.Figure:
    """
    df_prog: DataFrame con fecha, acumulado (programado Excel)
    df_real_pyrevit: DataFrame con fecha, pct_ejecutado_acum (PyRevit)
    """
    fig, ax = plt.subplots(figsize=(14, 5.5))
    _apply_style(fig, ax)

    fechas = pd.to_datetime(df_prog["fecha"], errors="coerce")
    y_prog = df_prog["acumulado"].tolist()

    ax.fill_between(fechas, y_prog, alpha=0.08, color=C["programado"])
    ax.plot(fechas, y_prog,
            color=C["programado"], linewidth=2.0,
            label="Programado % (Excel)", zorder=5)

    if df_real_pyrevit is not None and not df_real_pyrevit.empty \
            and "pct_ejecutado_acum" in df_real_pyrevit.columns:
        fechas_r = pd.to_datetime(df_real_pyrevit["fecha"])
        y_real   = df_real_pyrevit["pct_ejecutado_acum"].values
        ax.fill_between(fechas_r, y_real, alpha=0.09, color=C["real"])
        ax.plot(fechas_r, y_real,
                color=C["real"], linewidth=2.0,
                marker="s", markersize=4,
                markerfacecolor=C["bg"], markeredgecolor=C["real"],
                markeredgewidth=1.5, label="Real % Ejecutado (PyRevit)",
                linestyle="--", zorder=5)

    ax.axhline(100, color=C["subtext"], linewidth=0.8, linestyle=":", alpha=0.6)
    ax.set_ylim(-5, 115)
    ax.set_ylabel("Avance Acumulado %", fontsize=10)
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.0f%%"))
    ax.set_xlabel("Fecha", fontsize=10)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d-%m-%Y"))
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=3))
    plt.xticks(rotation=40, ha="right", fontsize=8)
    ax.set_title(titulo, fontsize=14, fontweight="bold", color=C["text"], pad=15)
    ax.legend(loc="upper left", fontsize=9,
              facecolor=C["panel"], edgecolor=C["border"],
              labelcolor=C["text"], framealpha=0.9)
    _watermark(fig)
    plt.tight_layout(pad=1.5)
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# AVANCE POR ITEM / ASSEMBLY CODE  (tabla vinculada)
# ─────────────────────────────────────────────────────────────────────────────

def grafico_avance_por_item(df_vinculado: pd.DataFrame,
                             titulo: str = "Avance Ejecutado por ITEM (Assembly Code)") -> plt.Figure:
    """
    Barras horizontales mostrando % ejecutado por ITEM.
    df_vinculado: resultado de vincular_item_pyrevit()
    """
    if df_vinculado is None or df_vinculado.empty or "pct_ejecutado" not in df_vinculado.columns:
        fig, ax = plt.subplots(figsize=(10, 4))
        _apply_style(fig, ax)
        ax.text(0.5, 0.5, "Sin datos de PyRevit vinculados",
                ha="center", va="center", color=C["subtext"], fontsize=11)
        return fig

    df = df_vinculado[df_vinculado["pct_ejecutado"] > 0].copy() \
        if "pct_ejecutado" in df_vinculado.columns else df_vinculado.copy()

    if df.empty:
        df = df_vinculado.copy()

    # Etiqueta: ITEM + nombre corto
    nombre_col = "Nombre de tarea" if "Nombre de tarea" in df.columns else None
    if nombre_col:
        df["_label"] = df["ITEM"].astype(str) + " – " + df[nombre_col].astype(str).str.strip().str[:45]
    else:
        df["_label"] = df["ITEM"].astype(str)

    df = df.sort_values("pct_ejecutado")

    colores = [C["alerta"]   if v < 25  else
               C["variacion"] if v < 75  else
               C["real"]      if v < 100 else
               C["purple"]
               for v in df["pct_ejecutado"]]

    fig, ax = plt.subplots(figsize=(13, max(4, len(df) * 0.55)))
    _apply_style(fig, ax)

    bars = ax.barh(df["_label"], df["pct_ejecutado"],
                   color=colores, height=0.65,
                   edgecolor=C["border"], linewidth=0.4)

    # Etiquetas con cantidad ejecutada / total si existen
    for bar, (_, row) in zip(bars, df.iterrows()):
        pct = row["pct_ejecutado"]
        extra = ""
        if "cantidad_ejecutada" in row and "cantidad_total_revit" in row:
            u = row.get("unidad", "")
            extra = f"  {row['cantidad_ejecutada']:.1f}/{row['cantidad_total_revit']:.1f} {u}"
        ax.text(min(pct + 1, 99), bar.get_y() + bar.get_height() / 2,
                f"{pct:.1f}%{extra}",
                va="center", ha="left", fontsize=7.5,
                color=C["text"], fontweight="bold")

    ax.set_xlim(0, 120)
    ax.axvline(100, color=C["subtext"], linewidth=0.8, linestyle=":")
    ax.set_xlabel("% Ejecutado (Ejecutado=True)", fontsize=10)
    ax.set_title(titulo, fontsize=13, fontweight="bold", color=C["text"], pad=12)

    patches = [
        mpatches.Patch(color=C["alerta"],    label="< 25% Crítico"),
        mpatches.Patch(color=C["variacion"], label="25-75% En progreso"),
        mpatches.Patch(color=C["real"],      label="75-99% Avanzado"),
        mpatches.Patch(color=C["purple"],    label="100% Completado"),
    ]
    ax.legend(handles=patches, loc="lower right", fontsize=8,
              facecolor=C["panel"], edgecolor=C["border"], labelcolor=C["text"])

    _watermark(fig)
    plt.tight_layout(pad=1.5)
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# AVANCE POR NIVEL (PyRevit)
# ─────────────────────────────────────────────────────────────────────────────

def grafico_avance_por_nivel(df_nivel: pd.DataFrame,
                              titulo: str = "Avance por Nivel – PyRevit") -> plt.Figure:
    """df_nivel: resultado de avance_por_nivel_pyrevit()"""
    if df_nivel is None or df_nivel.empty:
        fig, ax = plt.subplots(figsize=(8, 4))
        _apply_style(fig, ax)
        ax.text(0.5, 0.5, "Sin columna 'nivel' en los datos PyRevit",
                ha="center", va="center", color=C["subtext"], fontsize=11)
        return fig

    df = df_nivel.sort_values("pct_ejecutado")
    colores = [C["alerta"]   if v < 25 else
               C["variacion"] if v < 75 else
               C["real"]
               for v in df["pct_ejecutado"]]

    fig, ax = plt.subplots(figsize=(10, max(4, len(df) * 0.7)))
    _apply_style(fig, ax)

    ax.barh(df["nivel"].astype(str), df["pct_ejecutado"],
            color=colores, height=0.6, edgecolor=C["border"])

    for i, (_, row) in enumerate(df.iterrows()):
        v = row["pct_ejecutado"]
        ejec = row.get("cantidad_ejecutada", row.get("elementos_ejec", 0))
        total = row.get("cantidad_total", row.get("elementos_total", 0))
        ax.text(min(v + 1, 99), i, f"{v:.1f}%  ({ejec:.1f}/{total:.1f})",
                va="center", ha="left", fontsize=8.5,
                color=C["text"], fontweight="bold")

    ax.set_xlim(0, 115)
    ax.axvline(100, color=C["subtext"], linewidth=0.8, linestyle=":")
    ax.set_xlabel("% Ejecutado", fontsize=10)
    ax.set_title(titulo, fontsize=13, fontweight="bold", color=C["text"], pad=12)
    _watermark(fig)
    plt.tight_layout(pad=1.5)
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# CURVA S REAL desde PyRevit (acumulado por fecha)
# ─────────────────────────────────────────────────────────────────────────────

def grafico_curva_real_pyrevit(df_curva_real: pd.DataFrame,
                                titulo: str = "Curva Real de Ejecución – PyRevit") -> plt.Figure:
    """
    Muestra la curva de avance real acumulado de PyRevit (ejecutado por fecha).
    df_curva_real: resultado de calcular_curva_s_real_pyrevit()
    """
    if df_curva_real is None or df_curva_real.empty:
        fig, ax = plt.subplots(figsize=(10, 4))
        _apply_style(fig, ax)
        ax.text(0.5, 0.5, "Sin datos PyRevit con Ejecutado=True y fecha",
                ha="center", va="center", color=C["subtext"], fontsize=11)
        return fig

    fechas = pd.to_datetime(df_curva_real["fecha"])
    y_acum = df_curva_real["pct_ejecutado_acum"].values
    y_perd = df_curva_real["cantidad_ejecutada_periodo"].values

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(13, 7),
                                    gridspec_kw={"height_ratios": [3, 1]},
                                    sharex=True)
    _apply_style(fig, ax1, ax2)

    # Panel superior – acumulado %
    ax1.fill_between(fechas, y_acum, alpha=0.12, color=C["real"])
    ax1.plot(fechas, y_acum,
             color=C["real"], linewidth=2.5,
             marker="o", markersize=5,
             markerfacecolor=C["bg"], markeredgecolor=C["real"],
             markeredgewidth=2, label="% Ejecutado acumulado", zorder=5)
    for f, v in zip(fechas, y_acum):
        ax1.annotate(f"{v:.1f}%", (f, v),
                     textcoords="offset points", xytext=(0, 9),
                     ha="center", fontsize=7.5, color=C["real"], fontweight="bold")

    ax1.axhline(100, color=C["subtext"], linewidth=0.8, linestyle=":", alpha=0.6)
    ax1.set_ylim(-5, 115)
    ax1.set_ylabel("% Ejecutado Acumulado", fontsize=10)
    ax1.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.0f%%"))
    ax1.set_title(titulo, fontsize=13, fontweight="bold", color=C["text"], pad=12)
    ax1.legend(loc="upper left", fontsize=9,
               facecolor=C["panel"], edgecolor=C["border"],
               labelcolor=C["text"], framealpha=0.9)

    # Panel inferior – cantidad por periodo
    ax2.bar(fechas, y_perd, color=C["programado"], alpha=0.75,
            edgecolor=C["border"], width=0.8, label="Cantidad / período")
    ax2.set_ylabel("Cantidad", fontsize=9)
    ax2.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f"))
    ax2.xaxis.set_major_formatter(mdates.DateFormatter("%d-%m-%Y"))
    ax2.xaxis.set_major_locator(mdates.DayLocator(interval=1))
    plt.setp(ax2.xaxis.get_majorticklabels(), rotation=40, ha="right", fontsize=8)
    ax2.legend(loc="upper left", fontsize=8,
               facecolor=C["panel"], edgecolor=C["border"],
               labelcolor=C["text"], framealpha=0.9)

    _watermark(fig)
    plt.tight_layout(pad=1.2)
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# AVANCE POR ACTIVIDAD – Excel (barra semanal)
# ─────────────────────────────────────────────────────────────────────────────

def grafico_barras_actividades(datos: dict, semana_idx: int = -1,
                                titulo: str = "Avance por Actividad") -> plt.Figure:
    """Barras horizontales del avance acumulado Excel por actividad en semana dada."""
    actividades = datos.get("actividades", {})
    if not actividades:
        fig, ax = plt.subplots(figsize=(10, 4))
        _apply_style(fig, ax)
        ax.text(0.5, 0.5, "Sin datos", ha="center", va="center")
        return fig

    nombres, valores = [], []
    for act, d in actividades.items():
        acum = d.get("acumulado", [])
        if acum:
            idx = semana_idx if semana_idx >= 0 else len(acum) - 1
            idx = min(idx, len(acum) - 1)
            label = act[:55] + "…" if len(act) > 55 else act
            nombres.append(label)
            valores.append(acum[idx])

    pairs = sorted(zip(valores, nombres))
    valores_ord = [p[0] for p in pairs]
    nombres_ord = [p[1] for p in pairs]

    colores = [C["alerta"]   if v < 25 else
               C["variacion"] if v < 75 else
               C["real"]
               for v in valores_ord]

    fig, ax = plt.subplots(figsize=(12, max(4, len(nombres) * 0.55)))
    _apply_style(fig, ax)

    bars = ax.barh(nombres_ord, valores_ord, color=colores,
                   height=0.6, edgecolor=C["border"], linewidth=0.5)

    for bar, val in zip(bars, valores_ord):
        ax.text(min(val + 1.5, 98), bar.get_y() + bar.get_height() / 2,
                f"{val:.1f}%",
                va="center", ha="left", fontsize=8,
                color=C["text"], fontweight="bold")

    ax.set_xlim(0, 105)
    ax.axvline(100, color=C["subtext"], linewidth=0.8, linestyle=":")
    ax.set_xlabel("Avance Acumulado %", fontsize=10)
    ax.set_title(titulo, fontsize=13, fontweight="bold", color=C["text"], pad=12)

    patches = [
        mpatches.Patch(color=C["alerta"],    label="< 25% Crítico"),
        mpatches.Patch(color=C["variacion"], label="25-75% En progreso"),
        mpatches.Patch(color=C["real"],      label="> 75% Avanzado"),
    ]
    ax.legend(handles=patches, loc="lower right", fontsize=8,
              facecolor=C["panel"], edgecolor=C["border"], labelcolor=C["text"])

    _watermark(fig)
    plt.tight_layout(pad=1.5)
    return fig


# ─────────────────────────────────────────────────────────────────────────────
# VERSIONES PLOTLY INTERACTIVAS Y ULTRA-RÁPIDAS (60 FPS)
# ─────────────────────────────────────────────────────────────────────────────

def _apply_plotly_layout(fig, title_text):
    fig.update_layout(
        title=dict(
            text=f"<b>{title_text}</b>",
            font=dict(color=C["text"], size=16),
            x=0.01,
            y=0.96,
            xanchor="left",
            yanchor="top"
        ),
        paper_bgcolor=C["bg"],
        plot_bgcolor=C["panel"],
        font=dict(color=C["text"], family="Montserrat, system-ui, sans-serif"),
        margin=dict(l=60, r=40, t=75, b=50),
        hovermode="x unified",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1.0,
            font=dict(color=C["text"], size=11),
            bgcolor="rgba(255,255,255,0.9)",
            bordercolor=C["border"],
            borderwidth=1,
        ),
    )
    fig.update_xaxes(showgrid=True, gridcolor=C["grid"], zeroline=False, tickfont=dict(color=C["text"]))
    fig.update_yaxes(showgrid=True, gridcolor=C["grid"], zeroline=False, tickfont=dict(color=C["text"]))
    return fig


def _alinear_real_con_semanas(df_prog: pd.DataFrame, df_real: pd.DataFrame):
    """
    Alinea los valores reales de PyRevit con las semanas de df_prog.
    Ubica cada avance real en la semana correspondiente según rangos de fecha
    o correspondencia directa, evitando interpolaciones linspace arbitrarias.
    """
    if df_prog is None or df_prog.empty or df_real is None or df_real.empty:
        return [], []

    x_labels = df_prog["semana_label"].tolist() if "semana_label" in df_prog.columns else list(df_prog["semana"].astype(str))

    import re
    meses = {"ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6, "jul": 7, "ago": 8, "sep": 9, "oct": 10, "nov": 11, "dic": 12}
    semana_ends = []

    for r in df_prog.get("semana_rango", []):
        m = re.search(r"al\s+(\d{1,2})-([a-z]{3})", str(r).lower())
        if m:
            d = int(m.group(1))
            mon = meses.get(m.group(2), 9)
            semana_ends.append(pd.Timestamp(year=2026, month=mon, day=d, hour=23, minute=59, second=59))
        else:
            semana_ends.append(None)

    df_r = df_real.copy()
    if "fecha" in df_r.columns:
        df_r["fecha"] = pd.to_datetime(df_r["fecha"], dayfirst=True)
        df_r = df_r.sort_values("fecha")

    last_real_date = df_r["fecha"].max() if ("fecha" in df_r.columns and not df_r.empty) else None

    x_real = []
    y_real = []

    for i, lbl in enumerate(x_labels):
        ed = semana_ends[i] if i < len(semana_ends) else None
        if ed is not None and "fecha" in df_r.columns:
            sub = df_r[df_r["fecha"] <= ed]
            val = float(sub["pct_ejecutado_acum"].iloc[-1]) if not sub.empty else 0.0
            x_real.append(lbl)
            y_real.append(round(val, 2))
            if last_real_date is not None and ed >= last_real_date:
                break
        else:
            x_real.append(lbl)
            idx = min(i, len(df_r) - 1)
            y_real.append(float(df_r["pct_ejecutado_acum"].iloc[idx]))
            if i >= len(df_r) - 1:
                break

    return x_real, y_real


def grafico_curva_s_semanal_plotly(df_prog: pd.DataFrame,
                                    titulo: str = "Curva S – Avance Semanal Programado vs Real",
                                    df_real_pyrevit: pd.DataFrame = None) -> go.Figure:
    fig = go.Figure()
    x_labels = df_prog["semana_label"].tolist() if "semana_label" in df_prog.columns else list(range(len(df_prog)))
    y_prog   = df_prog["acumulado"].tolist()

    fig.add_trace(
        go.Scatter(
            x=x_labels,
            y=y_prog,
            name="Programado % (Excel)",
            mode="lines+markers+text",
            text=[f"{v:.1f}%" if v > 0 else "" for v in y_prog],
            textposition="top center",
            line=dict(color=C["programado"], width=3),
            marker=dict(size=7, color=C["programado"]),
            fill="tozeroy",
            fillcolor="rgba(0, 119, 182, 0.10)",
        )
    )

    if df_real_pyrevit is not None and not df_real_pyrevit.empty and "pct_ejecutado_acum" in df_real_pyrevit.columns:
        x_real, y_real = _alinear_real_con_semanas(df_prog, df_real_pyrevit)

        fig.add_trace(
            go.Scatter(
                x=x_real,
                y=y_real,
                name="Real % Ejecutado (BIM)",
                mode="lines+markers+text",
                text=[f"{v:.1f}%" if v > 0 else "" for v in y_real],
                textposition="bottom center",
                line=dict(color=C["real"], width=3, dash="dash"),
                marker=dict(size=8, symbol="square", color=C["real"]),
                fill="tozeroy",
                fillcolor="rgba(0, 160, 227, 0.10)",
            )
        )

    _apply_plotly_layout(fig, titulo)
    fig.update_yaxes(title_text="Avance Acumulado %", range=[-5, 115])
    return fig


def grafico_curva_s_diaria_plotly(df_prog: pd.DataFrame,
                                   titulo: str = "Curva S – Avance Diario Programado vs Real",
                                   df_real_pyrevit: pd.DataFrame = None) -> go.Figure:
    fig = go.Figure()
    fechas = pd.to_datetime(df_prog["fecha"], errors="coerce")
    y_prog = df_prog["acumulado"].tolist()

    fig.add_trace(
        go.Scatter(
            x=fechas,
            y=y_prog,
            name="Programado % (Excel)",
            mode="lines+markers+text",
            text=[f"{v:.1f}%" if pd.notna(v) and v > 0 else "" for v in y_prog],
            textposition="top center",
            textfont=dict(size=8),
            connectgaps=True,
            line=dict(color=C["programado"], width=3),
            marker=dict(size=6, color=C["programado"]),
            fill="tozeroy",
            fillcolor="rgba(0, 119, 182, 0.10)",
        )
    )

    if df_real_pyrevit is not None and not df_real_pyrevit.empty and "pct_ejecutado_acum" in df_real_pyrevit.columns:
        fechas_r = pd.to_datetime(df_real_pyrevit["fecha"], dayfirst=True)
        y_real   = df_real_pyrevit["pct_ejecutado_acum"].values
        fig.add_trace(
            go.Scatter(
                x=fechas_r,
                y=y_real,
                name="Real % Ejecutado (BIM)",
                mode="lines+markers+text",
                text=[f"{v:.1f}%" if v > 0 else "" for v in y_real],
                textposition="bottom center",
                textfont=dict(size=8),
                line=dict(color=C["real"], width=2.5, dash="dash"),
                marker=dict(size=7, symbol="square", color=C["real"]),
                fill="tozeroy",
                fillcolor="rgba(0, 160, 227, 0.10)",
            )
        )

    _apply_plotly_layout(fig, titulo)
    fig.update_xaxes(tickformat="%d-%m-%Y")
    fig.update_yaxes(title_text="Avance Acumulado %", range=[-5, 115])
    return fig


def grafico_avance_por_item_plotly(df_vinculado: pd.DataFrame,
                                    titulo: str = "Avance Ejecutado por ITEM (Assembly Code)") -> go.Figure:
    if df_vinculado is None or df_vinculado.empty or "pct_ejecutado" not in df_vinculado.columns:
        fig = go.Figure()
        fig.add_annotation(text="Sin datos de PyRevit vinculados", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False, font=dict(size=14, color=C["subtext"]))
        _apply_plotly_layout(fig, titulo)
        return fig

    df = df_vinculado.copy()
    if "assembly_description" in df.columns and df["assembly_description"].notna().any():
        df["_label"] = df["ITEM"].astype(str) + " – " + df["assembly_description"].astype(str).str.strip()
    elif "Nombre de tarea" in df.columns:
        df["_label"] = df["ITEM"].astype(str) + " – " + df["Nombre de tarea"].astype(str).str.strip()
    else:
        df["_label"] = df["ITEM"].astype(str)

    df = df.sort_values("pct_ejecutado")
    colores = [C["alerta"] if v < 25 else (C["variacion"] if v < 75 else (C["real"] if v < 100 else C["purple"])) for v in df["pct_ejecutado"]]

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            y=df["_label"],
            x=df["pct_ejecutado"],
            orientation="h",
            marker=dict(color=colores, line=dict(color=C["border"], width=1)),
            text=[f"{v:.1f}%" for v in df["pct_ejecutado"]],
            textposition="outside",
            hoverinfo="y+x",
        )
    )
    _apply_plotly_layout(fig, titulo)
    fig.update_layout(
        margin=dict(l=320, r=60, t=75, b=50),
        height=max(420, len(df) * 36)
    )
    fig.update_xaxes(title_text="% Ejecutado", range=[0, 115])
    fig.update_yaxes(automargin=True, tickfont=dict(size=11, color=C["text"]))
    return fig


def grafico_avance_por_nivel_plotly(df_nivel: pd.DataFrame,
                                     titulo: str = "Avance por Nivel – PyRevit") -> go.Figure:
    if df_nivel is None or df_nivel.empty:
        fig = go.Figure()
        fig.add_annotation(text="Sin columna 'nivel' en datos PyRevit", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False, font=dict(size=14, color=C["subtext"]))
        _apply_plotly_layout(fig, titulo)
        return fig

    df = df_nivel.sort_values("pct_ejecutado")
    colores = [C["alerta"] if v < 25 else (C["variacion"] if v < 75 else C["real"]) for v in df["pct_ejecutado"]]

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            y=df["nivel"].astype(str),
            x=df["pct_ejecutado"],
            orientation="h",
            marker=dict(color=colores, line=dict(color=C["border"], width=1)),
            text=[f"{v:.1f}%" for v in df["pct_ejecutado"]],
            textposition="outside",
        )
    )
    _apply_plotly_layout(fig, titulo)
    fig.update_xaxes(title_text="% Ejecutado", range=[0, 115])
    fig.update_layout(height=max(350, len(df) * 40))
    return fig


def grafico_curva_real_pyrevit_plotly(df_curva_real: pd.DataFrame,
                                       titulo: str = "Curva Real de Ejecución – PyRevit") -> go.Figure:
    if df_curva_real is None or df_curva_real.empty:
        fig = go.Figure()
        fig.add_annotation(text="Sin datos PyRevit con Ejecutado=True", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False, font=dict(size=14, color=C["subtext"]))
        _apply_plotly_layout(fig, titulo)
        return fig

    fechas = pd.to_datetime(df_curva_real["fecha"])
    y_acum = df_curva_real["pct_ejecutado_acum"].values
    y_period = df_curva_real["cantidad_ejecutada_periodo"].values

    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.1, row_heights=[0.7, 0.3])

    fig.add_trace(
        go.Scatter(
            x=fechas,
            y=y_acum,
            name="% Ejecutado acumulado",
            mode="lines+markers+text",
            text=[f"{v:.1f}%" for v in y_acum],
            textposition="top center",
            line=dict(color=C["real"], width=3),
            marker=dict(size=6, color=C["real"]),
            fill="tozeroy",
            fillcolor="rgba(0, 160, 227, 0.12)",
        ),
        row=1, col=1
    )

    fig.add_trace(
        go.Bar(
            x=fechas,
            y=y_period,
            name="Cantidad / período",
            marker_color=C["programado"],
            opacity=0.8,
        ),
        row=2, col=1
    )

    _apply_plotly_layout(fig, titulo)
    fig.update_xaxes(tickformat="%d-%m-%Y")
    fig.update_yaxes(title_text="% Acumulado", range=[-5, 115], row=1, col=1)
    fig.update_yaxes(title_text="Cantidad", row=2, col=1)
    return fig


def grafico_barras_actividades_plotly(datos: dict, semana_idx: int = -1,
                                       titulo: str = "Avance por Actividad") -> go.Figure:
    actividades = datos.get("actividades", {})
    if not actividades:
        fig = go.Figure()
        fig.add_annotation(text="Sin datos", xref="paper", yref="paper", x=0.5, y=0.5, showarrow=False)
        _apply_plotly_layout(fig, titulo)
        return fig

    nombres, valores = [], []
    for act, d in actividades.items():
        acum = d.get("acumulado", [])
        if acum:
            idx = semana_idx if semana_idx >= 0 else len(acum) - 1
            idx = min(idx, len(acum) - 1)
            label = act[:55] + "…" if len(act) > 55 else act
            nombres.append(label)
            valores.append(acum[idx])

    pairs = sorted(zip(valores, nombres))
    valores_ord = [p[0] for p in pairs]
    nombres_ord = [p[1] for p in pairs]

    colores = [C["alerta"] if v < 25 else (C["variacion"] if v < 75 else C["real"]) for v in valores_ord]

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            y=nombres_ord,
            x=valores_ord,
            orientation="h",
            marker=dict(color=colores, line=dict(color=C["border"], width=1)),
            text=[f"{v:.1f}%" for v in valores_ord],
            textposition="outside",
        )
    )
    _apply_plotly_layout(fig, titulo)
    fig.update_xaxes(title_text="Avance Acumulado %", range=[0, 105])
    fig.update_layout(height=max(380, len(nombres_ord) * 32))
    return fig
