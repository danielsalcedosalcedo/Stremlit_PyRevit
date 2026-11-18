# logic/curva_lb.py
# Curva S de avance físico: % acumulado PROGRAMADO (Excel) vs % REAL (PyRevit),
# con vista SEMANAL (S1..Sn) o DIARIA, línea de corte, desfase y barras con las
# CANTIDADES por período (programada vs real, en m2/m3/ml...). Replica en Plotly la
# lógica del gráfico "Avance físico" del tablero BIM 4D de referencia, adaptada a:
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
                         df_real_d: pd.DataFrame | None, fecha_corte,
                         fecha_referencia=None) -> dict | None:
    """
    df_sem    : calcular_curva_s_semanal  -> semana_label, semana_rango, acumulado (%)
    df_dia    : calcular_curva_s_diaria   -> fecha, acumulado (%)
    df_real_d : calcular_curva_s_real_pyrevit(frecuencia='D', fecha_corte=...)
                -> fecha, pct_ejecutado_acum (%) ; None/vacío si no hay real
    fecha_referencia : fecha del calendario del Excel, solo para deducir el año
                de los rangos de semana ('31-ago al 06-sep').
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
        if fecha_referencia is not None:      # p. ej. primera fecha del Avance Diario del Excel
            ref = pd.Timestamp(fecha_referencia)
        elif df_dia is not None and not df_dia.empty:
            ref = pd.to_datetime(df_dia["fecha"]).min()
        else:
            ref = corte
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
            })
            prev_p = prog
            prev_e = ev if ev is not None else prev_e
        # Si hay avance real posterior al fin del programado, se prolongan las semanas
        # (programado plano en su último valor) para que la curva real no se corte.
        while tiene_real and filas and filas[-1]["fin"] is not None and filas[-1]["fin"] < ultimo_ev:
            u = filas[-1]
            ini_x = u["fin"] + pd.Timedelta(days=1)
            fin_x = ini_x + pd.Timedelta(days=6)
            ev_x = ev_a(min(fin_x, ultimo_ev))
            filas.append({
                "n": u["n"] + 1, "label": f"S{u['n'] + 1:02d}", "ini": ini_x, "fin": fin_x,
                "prog": u["prog"], "ev": ev_x, "pv": 0.0, "evS": ev_x - prev_e,
            })
            prev_e = ev_x
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
        cabo["prog_corte_sem"] = filas[sem_idx]["prog"]   # programado en la semana del corte
        cabo["semanas_evaluadas"] = sem_idx + 1

    # ── DIARIO ────────────────────────────────────────────────────────────
    if df_dia is not None and not df_dia.empty:
        d = df_dia.copy()
        d["fecha"] = pd.to_datetime(d["fecha"]).dt.normalize()
        d = d.dropna(subset=["fecha"]).sort_values("fecha").reset_index(drop=True)
        fechas_d = list(d["fecha"])
        prog_d = [float(v) for v in d["acumulado"]]
        # Avance real posterior al fin del programado: se prolongan los días (programado plano)
        if tiene_real and fechas_d and fechas_d[-1] < ultimo_ev:
            extra = pd.date_range(fechas_d[-1] + pd.Timedelta(days=1), ultimo_ev, freq="D")
            fechas_d += list(extra)
            prog_d += [prog_d[-1]] * len(extra)
        ev_d = [ev_a(f) if (tiene_real and f <= ultimo_ev) else None for f in fechas_d]
        i_corte = max((i for i, f in enumerate(fechas_d) if f <= corte), default=0)
        i_ev = max((i for i, v in enumerate(ev_d) if v is not None), default=None)
        base["diario"] = {"fechas": fechas_d, "prog": prog_d,
                          "ev": ev_d, "i_corte": i_corte, "i_ev": i_ev}
        cabo["i_corte"], cabo["i_ev"] = i_corte, i_ev
        cabo["prog_dia"] = base["diario"]["prog"][i_corte]
        if i_ev is not None:
            cabo["desfase_dia"] = ev_d[i_ev] - base["diario"]["prog"][i_ev]
            cabo["ev_dia"] = ev_d[i_ev]
            cabo["ev_fecha"] = fechas_d[i_ev]

    # KPIs globales (prefiere lectura diaria si existe)
    prog = cabo.get("prog_dia", cabo.get("prog_corte_sem"))
    ev = cabo.get("ev_dia", cabo.get("ev_sem"))
    cabo["prog"], cabo["ev"] = prog, ev
    base["corte"] = cabo
    return base


# ── gráfico principal: curva S (%) + barras de cantidad por período ───────
def _barras_cantidad(df_cant: pd.DataFrame | None) -> list[dict]:
    """Extrae del comparativo de cantidades una lista de {unidad, etiqueta, periodos, prog, real}."""
    if df_cant is None or df_cant.empty:
        return []
    etiquetas = df_cant.attrs.get("unidades", {})
    out = []
    for col in df_cant.columns:
        if not col.startswith("cantidad_programada_"):
            continue
        u = col.removeprefix("cantidad_programada_")
        if f"cantidad_real_{u}" not in df_cant.columns:
            continue
        out.append({"unidad": u, "etiqueta": etiquetas.get(u, u),
                    "periodos": list(df_cant["periodo"]),
                    "prog": list(df_cant[col]), "real": list(df_cant[f"cantidad_real_{u}"])})
    return out


def grafico_curva_lb_plotly(base: dict, modo: str = "sem", df_cantidades: pd.DataFrame | None = None,
                            titulo: str | None = None) -> go.Figure:
    """
    Curva S (% programado y % real) + barras con la cantidad por período.
    modo = 'sem' (semanal) | 'dia' (diario)
    df_cantidades = calcular_comparativo_cantidades_periodo(...): periodo, cantidad_programada_<u>, cantidad_real_<u>
    """
    sem = modo == "sem"
    K = base["corte"]
    barras = _barras_cantidad(df_cantidades)
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # ── series de la curva, indexadas por clave de período ────────────────
    if sem:
        S = base["semanas"]
        clave = lambda w: w["label"]
        keys_curva = [w["label"] for w in S]
        prog_m = {w["label"]: w["prog"] for w in S}
        ev_m = {w["label"]: w["ev"] for w in S}
        hov_m = {w["label"]: (
            f"<b>Semana {w['n']}</b> · {_fd(w['ini'])} → {_fd(w['fin'])}<br>"
            f"Programado: {_f2(w['prog'])} %<br>"
            f"Real: {_f2(w['ev']) + ' %' if w['ev'] is not None else '—'}"
            + (f"<br>Desfase: {w['ev'] - w['prog']:+.2f} pts" if w["ev"] is not None else "")
        ) for w in S}
        k_corte = S[K["sem_idx"]]["label"]
        k_ev = S[K["sem_ev_idx"]]["label"] if K.get("sem_ev_idx") is not None else None
        keys_barras = [str(p) for b in barras for p in b["periodos"]]
        x = list(dict.fromkeys([str(p) for p in (barras[0]["periodos"] if barras else [])] + keys_curva))
        if not barras:
            x = keys_curva
    else:
        D = base["diario"]
        keys_curva = [pd.Timestamp(f) for f in D["fechas"]]
        prog_m = dict(zip(keys_curva, D["prog"]))
        ev_m = dict(zip(keys_curva, D["ev"]))
        hov_m = {f: (
            f"<b>{_fdl(f)}</b><br>Programado: {_f2(prog_m[f])} %<br>"
            f"Real: {_f2(ev_m[f]) + ' %' if ev_m[f] is not None else '— (sin ejecución registrada)'}"
            + (f"<br>Desfase: {ev_m[f] - prog_m[f]:+.2f} pts" if ev_m[f] is not None else "")
        ) for f in keys_curva}
        k_corte = keys_curva[K["i_corte"]]
        k_ev = keys_curva[K["i_ev"]] if K.get("i_ev") is not None else None
        fechas_barras = []
        if barras:
            fechas_barras = list(pd.to_datetime(barras[0]["periodos"], format="%d-%m-%Y", errors="coerce").normalize())
        x = sorted({*keys_curva, *[f for f in fechas_barras if not pd.isna(f)]})

    prog = [prog_m.get(k) for k in x]
    ev = [ev_m.get(k) for k in x]
    hover = [hov_m.get(k, "") for k in x]

    # ── barras de cantidad (eje derecho, ocupan el tercio inferior) ───────
    if barras:
        bmax = 0.0
        for b in barras:
            if sem:
                pm = {str(p): v for p, v in zip(b["periodos"], b["prog"])}
                rm = {str(p): v for p, v in zip(b["periodos"], b["real"])}
            else:
                ks = pd.to_datetime(b["periodos"], format="%d-%m-%Y", errors="coerce").normalize()
                pm, rm = dict(zip(ks, b["prog"])), dict(zip(ks, b["real"]))
            yp, yr = [pm.get(k) for k in x], [rm.get(k) for k in x]
            bmax = max(bmax, max([v or 0 for v in yp] + [v or 0 for v in yr] + [0]))
            fig.add_trace(go.Bar(x=x, y=yp, name=f"Cantidad programada · {b['etiqueta']}", marker_color=C_PROG,
                                 opacity=.35, offsetgroup=f"p{b['unidad']}",
                                 hovertemplate=f"%{{x}}<br>Programado: %{{y:,.2f}} {b['etiqueta']}<extra></extra>"),
                          secondary_y=True)
            fig.add_trace(go.Bar(x=x, y=yr, name=f"Cantidad real · {b['etiqueta']}", marker_color=C_REAL,
                                 opacity=.55, offsetgroup=f"r{b['unidad']}",
                                 hovertemplate=f"%{{x}}<br>Real: %{{y:,.2f}} {b['etiqueta']}<extra></extra>"),
                          secondary_y=True)
        bmax = (bmax * 1.15) or 1.0
        ticks = [round(bmax * i / 3, 2) for i in range(4)]
        fig.update_yaxes(range=[0, bmax / 0.32], tickvals=ticks, ticktext=[f"{t:,.0f}" for t in ticks],
                         showgrid=False, tickfont=dict(color=C["subtext"], size=10),
                         title_text="Cantidad por período (" + ", ".join(b["etiqueta"] for b in barras) + ")",
                         title_font=dict(color=C["subtext"], size=11), secondary_y=True)
    else:
        fig.update_yaxes(visible=False, secondary_y=True)

    # ── curvas S ──────────────────────────────────────────────────────────
    if base["tiene_real"]:
        fig.add_trace(go.Scatter(x=x, y=ev, name="Real (PyRevit) · % acumulado",
                                 mode="lines+markers" if sem else "lines",
                                 line=dict(color=C_REAL, width=2.8 if sem else 2.4), marker=dict(size=5),
                                 fill="tozeroy", fillcolor="rgba(192,57,43,.06)", connectgaps=False,
                                 text=hover, hoverinfo="text"), secondary_y=False)
    fig.add_trace(go.Scatter(x=x, y=prog, name="Programado (Excel) · % acumulado",
                             mode="lines+markers" if sem else "lines",
                             line=dict(color=C_PROG, width=2.4 if sem else 2), marker=dict(size=5),
                             connectgaps=False, text=hover, hoverinfo="text"), secondary_y=False)

    # ── línea de corte ────────────────────────────────────────────────────
    if sem:
        w = S[K["sem_idx"]]
        txt = f"CORTE · {k_corte} ({_fd(w['ini'])} → {_fd(w['fin'])})"
    else:
        txt = f"CORTE · {_fdl(k_corte)}"
    fig.add_shape(type="line", x0=k_corte, x1=k_corte, y0=0, y1=100, line=dict(color=C_CORTE, width=1.6, dash="dash"))
    fig.add_annotation(x=k_corte, y=100, text=f"<b>{txt}</b>", showarrow=False, xanchor="right",
                       yanchor="bottom", font=dict(color=C_CORTE, size=10), xshift=-6)
    if not sem and k_ev is not None and k_ev != k_corte:
        fig.add_shape(type="line", x0=k_ev, x1=k_ev, y0=0, y1=100, line=dict(color=C_REAL, width=1, dash="dot"))
        fig.add_annotation(x=k_ev, y=4, text=f"real hasta {_fd(k_ev)}", showarrow=False, xanchor="right",
                           font=dict(color=C_REAL, size=9))

    puntos = [(k_corte, prog_m.get(k_corte), C_PROG, "top right")]
    if k_ev is not None and ev_m.get(k_ev) is not None:
        puntos.append((k_ev, ev_m[k_ev], C_REAL, "bottom right"))
    for xx, v, col, pos in puntos:
        if v is None:
            continue
        fig.add_trace(go.Scatter(x=[xx], y=[v], mode="markers+text", text=[f"<b>{v:.1f}%</b>"],
                                 textposition=pos, textfont=dict(color=col, size=11),
                                 marker=dict(size=10, color=col, line=dict(color="white", width=1.6)),
                                 showlegend=False, hoverinfo="skip"), secondary_y=False)

    # ── caja de desfase ───────────────────────────────────────────────────
    d = K.get("desfase_sem") if sem else K.get("desfase_dia")
    if d is not None:
        dcol = C["alerta"] if d < 0 else "#2E9E6B"
        ref = f"S{K['semana_ev']}" if sem else _fd(K["ev_fecha"])
        cuerpo = (f"Real {_f2(K['ev_sem'])} % − Programado {_f2(K['prog_sem'])} %" if sem
                  else f"Real {_f2(K['ev_dia'])} % − Programado {_f2(base['diario']['prog'][K['i_ev']])} %")
        fig.add_annotation(
            xref="paper", yref="paper", x=0.01, y=0.98, xanchor="left", yanchor="top", showarrow=False,
            align="center", bordercolor=dcol, borderwidth=2, borderpad=6, bgcolor="rgba(255,255,255,.95)",
            text=f"<b>Desfase {ref}: {d:+.2f} pts</b><br><span style='font-size:10px'>{cuerpo}</span>",
            font=dict(color=dcol, size=13))

    _apply_plotly_layout(fig, titulo or
                         f"Curva S (% acumulado) y cantidad por período · {'Semanal' if sem else 'Diario'}")
    fig.update_layout(height=500, barmode="group", bargap=.35, hovermode="closest",
                      margin=dict(l=55, r=60 if barras else 25, t=80, b=55))
    fig.update_yaxes(range=[0, 102], ticksuffix="%", dtick=20, secondary_y=False)
    if sem:
        fig.update_xaxes(tickangle=-35 if len(x) > 14 else 0, type="category", categoryorder="array",
                         categoryarray=x, tickfont=dict(size=9))
    else:
        fig.update_xaxes(type="date", tickformat="%d-%b")
    return fig


def tabla_curva_lb(base: dict, modo: str = "sem") -> pd.DataFrame:
    """Valores de la curva S (% acumulado) por semana o por día."""
    if modo == "sem":
        return pd.DataFrame([{
            "Semana": w["label"], "Inicio": _fdl(w["ini"]), "Fin": _fdl(w["fin"]),
            "Programado %": round(w["prog"], 2), "Real %": None if w["ev"] is None else round(w["ev"], 2),
        } for w in base["semanas"]])
    D = base["diario"]
    return pd.DataFrame({"Fecha": [_fdl(f) for f in D["fechas"]],
                         "Programado %": [round(v, 2) for v in D["prog"]],
                         "Real %": [None if v is None else round(v, 2) for v in D["ev"]]})
