# logic/curva_lb.py
# Curva de avance físico acumulado: PROGRAMADO (Excel) vs REAL (PyRevit)
# con vista SEMANAL (S1..Sn) o DIARIA, línea de corte, desfase, barras de avance
# por semana y SPI. Replica en Plotly la lógica del gráfico "Avance físico" del
# tablero BIM 4D de referencia (LB / EV), adaptada a los datos de esta app:
#   - Programado  = hojas 'Avance Semanal' / 'Avance Diario' del Excel
#   - Real (EV)   = elementos ejecutados de PyRevit (% acumulado por fecha)

from __future__ import annotations

import re

import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from logic.charts import C, _apply_plotly_layout

C_PROG = "#2f6fd0"   # programado (Excel)
C_REAL = "#c0392b"   # real / ejecutado (PyRevit)
C_CORTE = "#0ea5e9"
_MES = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"]
_MES_IDX = {m: i + 1 for i, m in enumerate(_MES)} | {"set": 9}


# ── utilidades ────────────────────────────────────────────────────────────
def _fd(ts) -> str:
    if ts is None or pd.isna(ts):
        return "—"
    ts = pd.Timestamp(ts)
    return f"{ts.day}-{_MES[ts.month - 1]}"


def _fdl(ts) -> str:
    if ts is None or pd.isna(ts):
        return "—"
    ts = pd.Timestamp(ts)
    return f"{ts.day} {_MES[ts.month - 1]} {str(ts.year)[2:]}"


def _f2(v) -> str:
    return "—" if v is None else f"{v:.2f}"


def _rango_semana(texto: str):
    """'31-ago al 06-sep' -> ((31, 8), (6, 9)) o None."""
    m = re.search(r"(\d{1,2})-([a-zA-Z]{3})\w*\s+al\s+(\d{1,2})-([a-zA-Z]{3})", str(texto))
    if not m:
        return None
    m1, m2 = _MES_IDX.get(m.group(2).lower()), _MES_IDX.get(m.group(4).lower())
    if not m1 or not m2:
        return None
    return (int(m.group(1)), m1), (int(m.group(3)), m2)


def _fechas_semanas(rangos: list, referencia: pd.Timestamp) -> list:
    """Asigna año a cada rango 'dd-mmm al dd-mmm' recorriendo el calendario."""
    out, year, prev_m = [], referencia.year, None
    for r in rangos:
        if r is None:
            out.append((None, None))
            continue
        (d1, m1), (d2, m2) = r
        if prev_m is None:
            cand = pd.Timestamp(year=year, month=m1, day=d1)
            if cand - referencia > pd.Timedelta(days=60):
                year -= 1
        elif m1 < prev_m:
            year += 1
        ini = pd.Timestamp(year=year, month=m1, day=d1)
        y2 = year + 1 if m2 < m1 else year
        fin = pd.Timestamp(year=y2, month=m2, day=d2)
        year, prev_m = y2, m2
        out.append((ini, fin))
    return out


# ── construcción de la base de datos de la curva ──────────────────────────
def construir_base_curva(df_sem: pd.DataFrame | None, df_dia: pd.DataFrame | None,
                         df_real_d: pd.DataFrame | None, fecha_corte) -> dict | None:
    """
    df_sem    : calcular_curva_s_semanal  -> semana_label, semana_rango, acumulado (%)
    df_dia    : calcular_curva_s_diaria   -> fecha, acumulado (%)
    df_real_d : calcular_curva_s_real_pyrevit(frecuencia='D', fecha_corte=...)
                -> fecha, pct_ejecutado_acum (%) ; None/vacío si no hay real
    """
    if (df_sem is None or df_sem.empty) and (df_dia is None or df_dia.empty):
        return None
    corte = pd.Timestamp(fecha_corte).normalize()

    # Serie real diaria -> función "EV acumulado a una fecha"
    tiene_real = df_real_d is not None and not df_real_d.empty
    if tiene_real:
        real = (df_real_d.assign(fecha=pd.to_datetime(df_real_d["fecha"]))
                .groupby("fecha")["pct_ejecutado_acum"].max().sort_index())
        ultimo_ev = min(real.index.max().normalize(), corte)
    else:
        real, ultimo_ev = None, None

    def ev_a(fecha):
        if not tiene_real:
            return None
        sub = real[:fecha]
        return float(sub.iloc[-1]) if len(sub) else 0.0

    base = {"fecha_corte": corte, "tiene_real": tiene_real, "semanas": None, "diario": None}
    cabo = {}

    # ── SEMANAL ───────────────────────────────────────────────────────────
    if df_sem is not None and not df_sem.empty:
        ref = pd.to_datetime(df_dia["fecha"]).min() if (df_dia is not None and not df_dia.empty) else corte
        fechas = _fechas_semanas([_rango_semana(r) for r in df_sem["semana_rango"]], ref)
        filas, prev_p, prev_e = [], 0.0, 0.0
        for (_, w), (ini, fin) in zip(df_sem.iterrows(), fechas):
            prog = float(w["acumulado"])
            ev = None
            if tiene_real and ini is not None and ini <= ultimo_ev:
                ev = ev_a(min(fin, ultimo_ev))
            pv = prog - prev_p
            evs = None if ev is None else ev - prev_e
            n = re.sub(r"\D", "", str(w["semana_label"]))
            filas.append({
                "n": int(n) if n else len(filas) + 1, "label": str(w["semana_label"]),
                "ini": ini, "fin": fin, "prog": prog, "ev": ev, "pv": pv, "evS": evs,
                "spi": (evs / pv) if (evs is not None and pv > 0.01) else None,
                "spiA": (ev / prog) if (ev is not None and prog > 0.01) else None,
            })
            prev_p = prog
            prev_e = ev if ev is not None else prev_e
        base["semanas"] = filas

        # semana del corte (la que contiene la fecha de corte; acotada al calendario)
        tope = min(corte, filas[-1]["fin"] or corte)
        sem_idx = 0
        for i, w in enumerate(filas):
            if w["ini"] is not None and w["ini"] <= tope:
                sem_idx = i
        ev_idx = max((i for i, w in enumerate(filas) if w["ev"] is not None), default=None)
        cabo.update(sem_idx=sem_idx, sem_ev_idx=ev_idx)
        if ev_idx is not None:
            cabo["desfase_sem"] = filas[ev_idx]["ev"] - filas[ev_idx]["prog"]
            cabo["semana_ev"] = filas[ev_idx]["n"]
            cabo["prog_sem"] = filas[ev_idx]["prog"]
            cabo["ev_sem"] = filas[ev_idx]["ev"]
        cabo["semanas_evaluadas"] = sem_idx + 1

    # ── DIARIO ────────────────────────────────────────────────────────────
    if df_dia is not None and not df_dia.empty:
        d = df_dia.copy()
        d["fecha"] = pd.to_datetime(d["fecha"]).dt.normalize()
        d = d.dropna(subset=["fecha"]).sort_values("fecha").reset_index(drop=True)
        fechas_d = list(d["fecha"])
        ev_d = [ev_a(f) if (tiene_real and f <= ultimo_ev) else None for f in fechas_d]
        i_corte = max((i for i, f in enumerate(fechas_d) if f <= corte), default=0)
        i_ev = max((i for i, v in enumerate(ev_d) if v is not None), default=None)
        base["diario"] = {"fechas": fechas_d, "prog": [float(v) for v in d["acumulado"]],
                          "ev": ev_d, "i_corte": i_corte, "i_ev": i_ev}
        cabo["i_corte"], cabo["i_ev"] = i_corte, i_ev
        cabo["prog_dia"] = base["diario"]["prog"][i_corte]
        if i_ev is not None:
            cabo["desfase_dia"] = ev_d[i_ev] - base["diario"]["prog"][i_ev]
            cabo["ev_dia"] = ev_d[i_ev]
            cabo["ev_fecha"] = fechas_d[i_ev]

    # KPIs globales (prefiere lectura diaria si existe)
    prog = cabo.get("prog_dia", cabo.get("prog_sem"))
    ev = cabo.get("ev_dia", cabo.get("ev_sem"))
    cabo["prog"], cabo["ev"] = prog, ev
    cabo["spi"] = (ev / prog) if (ev is not None and prog and prog > 0.01) else None
    base["corte"] = cabo
    return base


# ── gráfico principal ─────────────────────────────────────────────────────
def grafico_curva_lb_plotly(base: dict, modo: str = "sem", barras: bool = True,
                            titulo: str | None = None) -> go.Figure:
    """modo = 'sem' (semanal) | 'dia' (diario)."""
    sem = modo == "sem"
    K = base["corte"]
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    if sem:
        S = base["semanas"]
        x = [w["label"] for w in S]
        prog, ev = [w["prog"] for w in S], [w["ev"] for w in S]
        hover = [
            f"<b>Semana {w['n']}</b> · {_fd(w['ini'])} → {_fd(w['fin'])}<br>"
            f"Programado: {_f2(w['prog'])} %<br>"
            f"Real: {_f2(w['ev']) + ' %' if w['ev'] is not None else '—'}<br>"
            + (f"Desfase: {w['ev'] - w['prog']:+.2f} pts<br>" if w["ev"] is not None else "")
            + f"Semana · programado {_f2(w['pv'])} %"
            + (f" · real {_f2(w['evS'])} %" if w["evS"] is not None else "")
            + (f"<br>SPI semana {_f2(w['spi'])} · acumulado {_f2(w['spiA'])}" if w["spi"] is not None else "")
            for w in S
        ]
        i_c, i_e = K["sem_idx"], K.get("sem_ev_idx")
        x_c = x[i_c]
        x_e = x[i_e] if i_e is not None else None
    else:
        D = base["diario"]
        x, prog, ev = D["fechas"], D["prog"], D["ev"]
        hover = [
            f"<b>{_fdl(f)}</b><br>Programado: {_f2(p)} %<br>"
            f"Real: {_f2(e) + ' %' if e is not None else '— (sin ejecución registrada)'}"
            + (f"<br>Desfase: {e - p:+.2f} pts" if e is not None else "")
            for f, p, e in zip(x, prog, ev)
        ]
        i_c, i_e = K["i_corte"], K.get("i_ev")
        x_c = x[i_c]
        x_e = x[i_e] if i_e is not None else None

    # Barras de avance de la semana (programado / real) en el tercio inferior
    if sem and barras:
        pv = [w["pv"] for w in S]
        evs = [w["evS"] for w in S]
        bmax = max(max(pv), max((v or 0) for v in evs)) * 1.15 or 1
        fig.add_trace(go.Bar(x=x, y=pv, name="Programado de la semana", marker_color=C_PROG, opacity=.28,
                             hoverinfo="skip", offsetgroup="a"), secondary_y=True)
        fig.add_trace(go.Bar(x=x, y=evs, name="Real de la semana", marker_color=C_REAL, opacity=.42,
                             hoverinfo="skip", offsetgroup="b"), secondary_y=True)
        fig.update_yaxes(range=[0, bmax / 0.32], showgrid=False, showticklabels=False, secondary_y=True)
    else:
        fig.update_yaxes(visible=False, secondary_y=True)

    # Curvas
    if base["tiene_real"]:
        fig.add_trace(go.Scatter(x=x, y=ev, name="Real (PyRevit)", mode="lines+markers" if sem else "lines",
                                 line=dict(color=C_REAL, width=2.8 if sem else 2.4), marker=dict(size=5),
                                 fill="tozeroy", fillcolor="rgba(192,57,43,.06)", connectgaps=False,
                                 text=hover, hoverinfo="text"))
    fig.add_trace(go.Scatter(x=x, y=prog, name="Programado (Excel)", mode="lines+markers" if sem else "lines",
                             line=dict(color=C_PROG, width=2.4 if sem else 2), marker=dict(size=5),
                             text=hover, hoverinfo="text"))

    # Línea de corte
    txt = (f"CORTE · {x_c} ({_fd(S[i_c]['ini'])} → {_fd(S[i_c]['fin'])})" if sem
           else f"CORTE · {_fdl(x_c)}")
    fig.add_shape(type="line", x0=x_c, x1=x_c, y0=0, y1=100, line=dict(color=C_CORTE, width=1.6, dash="dash"))
    fig.add_annotation(x=x_c, y=100, text=f"<b>{txt}</b>", showarrow=False, xanchor="right",
                       yanchor="bottom", font=dict(color=C_CORTE, size=10), xshift=-6)

    # Diario: marca "real registrado hasta"
    if not sem and x_e is not None and x_e != x_c:
        fig.add_shape(type="line", x0=x_e, x1=x_e, y0=0, y1=100, line=dict(color=C_REAL, width=1, dash="dot"))
        fig.add_annotation(x=x_e, y=4, text=f"real hasta {_fd(x_e)}", showarrow=False, xanchor="right",
                           font=dict(color=C_REAL, size=9))

    # Puntos y etiquetas en el corte
    puntos = [(x_c, prog[i_c], C_PROG, "top right")]
    if i_e is not None and ev[i_e] is not None:
        puntos.append((x_e, ev[i_e], C_REAL, "bottom right"))
    for xx, v, col, pos in puntos:
        fig.add_trace(go.Scatter(x=[xx], y=[v], mode="markers+text", text=[f"<b>{v:.1f}%</b>"],
                                 textposition=pos, textfont=dict(color=col, size=11),
                                 marker=dict(size=10, color=col, line=dict(color="white", width=1.6)),
                                 showlegend=False, hoverinfo="skip"))

    # Caja de desfase
    d = K.get("desfase_sem") if sem else K.get("desfase_dia")
    if d is not None:
        dcol = C["alerta"] if d < 0 else "#2E9E6B"
        ref = (f"S{K['semana_ev']}" if sem else _fd(K["ev_fecha"]))
        cuerpo = (f"Real {_f2(K['ev_sem'])} % − Programado {_f2(K['prog_sem'])} %" if sem
                  else f"Real {_f2(K['ev_dia'])} % − Programado {_f2(base['diario']['prog'][K['i_ev']])} %")
        fig.add_annotation(
            xref="paper", yref="paper", x=0.01, y=0.98, xanchor="left", yanchor="top", showarrow=False,
            align="center", bordercolor=dcol, borderwidth=2, borderpad=6, bgcolor="rgba(255,255,255,.95)",
            text=f"<b>Desfase {ref}: {d:+.2f} pts</b><br><span style='font-size:10px'>{cuerpo}</span>",
            font=dict(color=dcol, size=13))

    _apply_plotly_layout(fig, titulo or f"Avance físico acumulado · {'Semanal' if sem else 'Diario'}")
    fig.update_layout(height=470, barmode="group", bargap=.35, hovermode="closest",
                      margin=dict(l=55, r=25, t=80, b=55))
    fig.update_yaxes(range=[0, 102], ticksuffix="%", dtick=20, secondary_y=False)
    if sem:
        fig.update_xaxes(tickangle=0, categoryorder="array", categoryarray=x)
    else:
        fig.update_xaxes(type="date", tickformat="%d-%b")
    return fig


# ── gráfico SPI semanal ───────────────────────────────────────────────────
def grafico_spi_curva_lb_plotly(base: dict) -> go.Figure | None:
    rows = [w for w in (base.get("semanas") or []) if w["spi"] is not None]
    if not rows:
        return None
    x = [w["label"] for w in rows]
    cols = [("#2E9E6B" if w["spi"] >= 1 else "#D4A017" if w["spi"] >= .7 else C["alerta"]) for w in rows]
    fig = go.Figure()
    fig.add_trace(go.Bar(x=x, y=[min(w["spi"], 2) for w in rows], marker_color=cols, opacity=.6,
                         text=[f"{w['spi']:.2f}" if w["spi"] <= 2 else f"▲{w['spi']:.1f}" for w in rows],
                         textposition="outside", name="SPI de la semana", hoverinfo="skip"))
    fig.add_trace(go.Scatter(x=x, y=[w["spiA"] for w in rows], mode="lines+markers", name="SPI acumulado",
                             line=dict(color=C_PROG, width=2)))
    fig.add_hline(y=1, line=dict(color=C["subtext"], width=1.2, dash="dash"),
                  annotation_text="SPI = 1 (en plazo)", annotation_position="top right")
    _apply_plotly_layout(fig, "SPI semanal (real / programado)")
    fig.update_layout(height=300, hovermode="x unified", margin=dict(l=55, r=25, t=70, b=40))
    fig.update_yaxes(range=[0, 2.2], dtick=.5)
    return fig


def tabla_curva_lb(base: dict, modo: str = "sem") -> pd.DataFrame:
    if modo == "sem":
        return pd.DataFrame([{
            "Semana": w["label"], "Inicio": _fdl(w["ini"]), "Fin": _fdl(w["fin"]),
            "Programado %": round(w["prog"], 2), "Real %": None if w["ev"] is None else round(w["ev"], 2),
            "Prog. semana %": round(w["pv"], 2), "Real semana %": None if w["evS"] is None else round(w["evS"], 2),
            "SPI semana": None if w["spi"] is None else round(w["spi"], 2),
        } for w in base["semanas"]])
    D = base["diario"]
    return pd.DataFrame({"Fecha": [_fdl(f) for f in D["fechas"]],
                         "Programado %": [round(v, 2) for v in D["prog"]],
                         "Real %": [None if v is None else round(v, 2) for v in D["ev"]]})
