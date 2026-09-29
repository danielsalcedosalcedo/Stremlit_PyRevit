# logic/processor.py
# Procesamiento de datos y Curvas S – con vinculación ITEM ↔ Assembly Code
# Desarrollado por Daniel Salcedo - Coordinador BIM

import pandas as pd
import numpy as np


# ─────────────────────────────────────────────────────────────────────────────
# CURVAS S DESDE EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def calcular_curva_s_semanal(datos_semanal: dict, actividad: str = None) -> pd.DataFrame:
    """
    Curva S semanal desde la hoja 'Avance Semanal' del Excel.
    Si actividad=None, promedia todas las actividades raíz.
    Retorna DataFrame: semana, semana_label, semana_rango, parcial, acumulado
    """
    actividades = datos_semanal["actividades"]
    semanas     = datos_semanal["semanas"]

    if actividad and actividad in actividades:
        d = actividades[actividad]
        df = pd.DataFrame({
            "semana":    semanas,
            "parcial":   d.get("parcial",   [0] * len(semanas)),
            "acumulado": d.get("acumulado", [0] * len(semanas)),
        })
    else:
        raices = _filtrar_actividades_raiz(list(actividades.keys()))
        acum_sum = np.zeros(len(semanas))
        count = 0
        for act in raices:
            if act in actividades:
                a = actividades[act].get("acumulado", [])
                if len(a) == len(semanas):
                    acum_sum += np.array(a)
                    count += 1
        promedio = acum_sum / count if count > 0 else acum_sum
        df = pd.DataFrame({
            "semana":    semanas,
            "parcial":   [0] * len(semanas),
            "acumulado": promedio.tolist(),
        })

    df["semana_label"] = df["semana"].apply(lambda s: str(s).split("\n")[0])
    df["semana_rango"] = df["semana"].apply(
        lambda s: str(s).split("\n")[1] if "\n" in str(s) else ""
    )
    return df


def calcular_curva_s_diaria(datos_diario: dict, actividad: str = None) -> pd.DataFrame:
    """
    Curva S diaria desde la hoja 'Avance Diario' del Excel.
    Retorna DataFrame: fecha, parcial, acumulado
    """
    fechas      = datos_diario["fechas"]
    actividades = datos_diario["actividades"]

    if actividad and actividad in actividades:
        d = actividades[actividad]
        df = pd.DataFrame({
            "fecha":     fechas,
            "parcial":   d.get("parcial",   [0] * len(fechas)),
            "acumulado": d.get("acumulado", [0] * len(fechas)),
        })
    else:
        raices = _filtrar_actividades_raiz(list(actividades.keys()))
        acum_sum = np.zeros(len(fechas))
        count = 0
        for act in raices:
            if act in actividades:
                a = actividades[act].get("acumulado", [])
                if len(a) == len(fechas):
                    acum_sum += np.array(a)
                    count += 1
        promedio = acum_sum / count if count > 0 else acum_sum
        df = pd.DataFrame({
            "fecha":     fechas,
            "parcial":   [0] * len(fechas),
            "acumulado": promedio.tolist(),
        })

    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    return df


# ─────────────────────────────────────────────────────────────────────────────
# VINCULACIÓN PYREVIT ↔ EXCEL (ITEM = ASSEMBLY CODE)
# ─────────────────────────────────────────────────────────────────────────────

def extraer_unidad_y_cantidad_de_desc(assembly_desc, m2=0.0, m3=0.0, ml=0.0, default_qty=1.0):
    """
    Determina la unidad objetivo y la cantidad correspondiente según la
    descripción de montaje (o texto equivalente).
    Prioriza las unidades mencionadas en la descripción (m2, m3, ml, unid/ud).
    """
    if not assembly_desc:
        d = ""
    else:
        d = str(assembly_desc).lower().strip()

    # Patrones para m2
    if any(pat in d for pat in ["m2", "m²", "sqm", "metro cuadrado", "metros cuadrados"]):
        return "m2", m2 if m2 > 0 else (m3 if m3 > 0 else (ml if ml > 0 else default_qty))

    # Patrones para m3
    if any(pat in d for pat in ["m3", "m³", "cum", "metro cubico", "metros cubicos", "metro cúbico", "metros cúbicos"]):
        return "m3", m3 if m3 > 0 else (m2 if m2 > 0 else (ml if ml > 0 else default_qty))

    # Patrones para ml (metros lineales)
    if any(pat in d for pat in ["ml", "m.l", "linear", "metro lineal", "metros lineales"]):
        return "ml", ml if ml > 0 else (m2 if m2 > 0 else (m3 if m3 > 0 else default_qty))

    # Patrones para unidad / conteo
    if any(pat in d for pat in ["ud", "unid", "unidades", "unidad", "pza", "piezas", "stk"]):
        return "unid", default_qty

    # Fallback por geometría si no hay unidad explícita en la descripción:
    if m3 > 0:
        return "m3", m3
    elif m2 > 0:
        return "m2", m2
    elif ml > 0:
        return "ml", ml
    else:
        return "unid", default_qty


def vincular_item_pyrevit(df_datos: pd.DataFrame, df_pyrevit: pd.DataFrame) -> pd.DataFrame:
    """
    Une la hoja 'Datos' del Excel con los datos de PyRevit usando
    ITEM (Excel) = assembly_code (PyRevit).
    Ordena el resultado por 'assembly_description' / 'Nombre de tarea'.

    Retorna DataFrame con columnas:
      ITEM, Nombre de tarea, Comienzo, Fin,
      cantidad_total_revit, cantidad_ejecutada, pct_ejecutado,
      unidad, elementos_total, elementos_ejecutados
    """
    if df_datos is None or df_pyrevit is None:
        return pd.DataFrame()

    if "ITEM" not in df_datos.columns:
        raise ValueError("La hoja 'Datos' no tiene columna 'ITEM'.")
    if "assembly_code" not in df_pyrevit.columns:
        raise ValueError("Los datos PyRevit no tienen columna 'assembly_code'.")

    # Agrupar PyRevit por assembly_code
    agg = _agrupar_por_assembly_code(df_pyrevit)

    # Merge con datos del Excel
    df_merge = df_datos.merge(
        agg,
        left_on="ITEM",
        right_on="assembly_code",
        how="left"
    )

    # Rellenar NaN en columnas numéricas
    for col in ["cantidad_total_revit", "cantidad_ejecutada",
                "pct_ejecutado", "elementos_total", "elementos_ejecutados"]:
        if col in df_merge.columns:
            df_merge[col] = df_merge[col].fillna(0)

    if "pct_ejecutado" in df_merge.columns:
        df_merge["pct_ejecutado"] = df_merge["pct_ejecutado"].round(2)

    # Ordenar por Descripción de Montaje o Nombre de tarea si está presente
    if "assembly_description" in df_merge.columns and df_merge["assembly_description"].notna().any():
        df_merge["sort_key"] = df_merge["assembly_description"].fillna(df_merge.get("Nombre de tarea", ""))
        df_merge = df_merge.sort_values(by=["sort_key", "ITEM"]).drop(columns=["sort_key"])
    elif "Nombre de tarea" in df_merge.columns:
        df_merge = df_merge.sort_values(by=["Nombre de tarea", "ITEM"])
    elif "ITEM" in df_merge.columns:
        df_merge = df_merge.sort_values(by="ITEM")

    return df_merge


def _agrupar_por_assembly_code(df_pyrevit: pd.DataFrame) -> pd.DataFrame:
    """
    Agrupa los elementos de PyRevit por assembly_code.
    Determina las cantidades según la unidad expresada en la descripción de montaje.
    Calcula: cantidad_total_revit, cantidad_ejecutada, pct_ejecutado,
             elementos_total, elementos_ejecutados, unidad, assembly_description
    """
    if df_pyrevit.empty:
        return pd.DataFrame()

    tiene_ejecutado = "ejecutado" in df_pyrevit.columns
    tiene_cantidad  = "cantidad_total" in df_pyrevit.columns
    tiene_unidad    = "unidad" in df_pyrevit.columns
    tiene_desc      = "assembly_description" in df_pyrevit.columns

    grupos = df_pyrevit.groupby("assembly_code")
    rows = []

    for code, grp in grupos:
        row = {"assembly_code": code}
        row["elementos_total"] = len(grp)

        # Descripción del montaje
        if tiene_desc:
            mode_desc = grp["assembly_description"].mode()
            desc = mode_desc.iloc[0] if not mode_desc.empty else ""
        else:
            desc = ""
        row["assembly_description"] = desc

        if tiene_ejecutado:
            ejec = grp[grp["ejecutado"] == True]
            row["elementos_ejecutados"] = len(ejec)
        else:
            ejec = pd.DataFrame()
            row["elementos_ejecutados"] = 0

        # Si hay columnas de métricas especificas m2/m3/ml, recalcular con respecto a la descripción
        has_specific_metrics = any(c in grp.columns for c in ["m2", "m3", "ml"])
        if has_specific_metrics:
            m2_tot = grp["m2"].sum() if "m2" in grp.columns else 0.0
            m3_tot = grp["m3"].sum() if "m3" in grp.columns else 0.0
            ml_tot = grp["ml"].sum() if "ml" in grp.columns else 0.0

            m2_ejec = ejec["m2"].sum() if ("m2" in ejec.columns and not ejec.empty) else 0.0
            m3_ejec = ejec["m3"].sum() if ("m3" in ejec.columns and not ejec.empty) else 0.0
            ml_ejec = ejec["ml"].sum() if ("ml" in ejec.columns and not ejec.empty) else 0.0

            u_det, total_qty = extraer_unidad_y_cantidad_de_desc(desc, m2_tot, m3_tot, ml_tot, default_qty=float(len(grp)))
            _, ejec_qty = extraer_unidad_y_cantidad_de_desc(desc, m2_ejec, m3_ejec, ml_ejec, default_qty=float(len(ejec)))

            row["unidad"] = u_det
            row["cantidad_total_revit"] = round(total_qty, 3)
            row["cantidad_ejecutada"]   = round(ejec_qty, 3)
            row["pct_ejecutado"] = round(ejec_qty / total_qty * 100, 2) if total_qty > 0 else 0.0
        elif tiene_cantidad:
            total_qty = grp["cantidad_total"].sum()
            ejec_qty  = ejec["cantidad_total"].sum() if not ejec.empty else 0.0
            row["cantidad_total_revit"] = round(total_qty, 3)
            row["cantidad_ejecutada"]   = round(ejec_qty, 3)
            row["pct_ejecutado"] = round(ejec_qty / total_qty * 100, 2) if total_qty > 0 else 0.0
            if tiene_unidad:
                row["unidad"] = grp["unidad"].mode().iloc[0] if not grp["unidad"].mode().empty else "-"
            else:
                row["unidad"] = "-"
        else:
            row["cantidad_total_revit"] = 0.0
            row["cantidad_ejecutada"]   = 0.0
            row["unidad"] = "-"
            if row["elementos_total"] > 0:
                row["pct_ejecutado"] = round(row["elementos_ejecutados"] / row["elementos_total"] * 100, 2)
            else:
                row["pct_ejecutado"] = 0.0

        rows.append(row)

    res_df = pd.DataFrame(rows)
    if not res_df.empty and "assembly_description" in res_df.columns:
        res_df = res_df.sort_values(by=["assembly_description", "assembly_code"])
    return res_df


def _detectar_columna_cantidad(df: pd.DataFrame) -> str:
    """
    Detecta qué columna de cantidad usar para el cálculo real de la curva S.
    Prioriza m2/m3/ml según la descripción de montaje (assembly_description).
    Fallback: cantidad_total.
    """
    desc = ""
    if "assembly_description" in df.columns:
        modos = df["assembly_description"].dropna()
        if not modos.empty:
            desc = str(modos.mode().iloc[0]).lower().strip()

    tiene_m2  = "m2"  in df.columns and df["m2"].sum()  > 0
    tiene_m3  = "m3"  in df.columns and df["m3"].sum()  > 0
    tiene_ml  = "ml"  in df.columns and df["ml"].sum()  > 0
    tiene_qty = "cantidad_total" in df.columns

    # Decidir por descripción
    if any(p in desc for p in ["[m2]", "m2", "m²", "sqm", "metro cuadrado"]):
        if tiene_m2:  return "m2"
    if any(p in desc for p in ["[m3]", "m3", "m³", "cum", "metro cubico", "metro cúbico"]):
        if tiene_m3:  return "m3"
    if any(p in desc for p in ["[ml]", "ml", "m.l", "linear", "metro lineal"]):
        if tiene_ml:  return "ml"

    # Fallback geométrico
    if tiene_m2:  return "m2"
    if tiene_m3:  return "m3"
    if tiene_ml:  return "ml"
    if tiene_qty: return "cantidad_total"
    return None


def calcular_curva_s_real_pyrevit(df_pyrevit: pd.DataFrame,
                                   df_datos: pd.DataFrame = None,
                                   frecuencia: str = "D",
                                   fecha_corte: pd.Timestamp = None) -> pd.DataFrame:
    """
    Genera la Curva S real desde PyRevit.

    Usa la columna de cantidad correcta (m2 / m3 / ml / cantidad_total) según
    la descripción de montaje, exactamente igual que _agrupar_por_assembly_code.

    Parámetros:
      frecuencia: "D" = diario, "W" = semanal (fin de semana)
      fecha_corte: filtra hasta esta fecha inclusive
    """
    if df_pyrevit is None or df_pyrevit.empty:
        return pd.DataFrame()
    if "ejecutado" not in df_pyrevit.columns:
        return pd.DataFrame()

    # Detectar columna de cantidad ANTES de filtrar (con todos los elementos)
    col_qty = _detectar_columna_cantidad(df_pyrevit)
    if col_qty is None:
        return pd.DataFrame()

    # Total de referencia = TODOS los elementos de la partida (ejecutados + no ejecutados)
    total_ref = float(df_pyrevit[col_qty].sum())

    # Filtrar solo ejecutados
    df_ejec = df_pyrevit[df_pyrevit["ejecutado"] == True].copy()
    if df_ejec.empty:
        return pd.DataFrame()

    # Manejo robusto de fechas — parseo día primero (DD-MM-YYYY) y fallback a ei_fecha_ejecucion
    if "fecha" in df_ejec.columns:
        df_ejec["fecha"] = pd.to_datetime(df_ejec["fecha"], format="mixed", dayfirst=True, errors="coerce")
        if "ei_fecha_ejecucion" in df_ejec.columns:
            df_ejec["fecha"] = df_ejec["fecha"].fillna(pd.to_datetime(df_ejec["ei_fecha_ejecucion"], format="mixed", dayfirst=True, errors="coerce"))
        min_date = df_ejec["fecha"].dropna().min()
        if pd.isna(min_date):
            min_date = pd.Timestamp.now().floor("D")
        df_ejec["fecha"] = df_ejec["fecha"].fillna(min_date)
    elif "ei_fecha_ejecucion" in df_ejec.columns:
        df_ejec["fecha"] = pd.to_datetime(df_ejec["ei_fecha_ejecucion"], format="mixed", dayfirst=True, errors="coerce")
        min_date = df_ejec["fecha"].dropna().min()
        if pd.isna(min_date):
            min_date = pd.Timestamp.now().floor("D")
        df_ejec["fecha"] = df_ejec["fecha"].fillna(min_date)
    else:
        df_ejec["fecha"] = pd.Timestamp.now().floor("D")

    # Filtrar hasta fecha de corte si aplica
    if fecha_corte is not None:
        df_ejec = df_ejec[df_ejec["fecha"] <= pd.to_datetime(fecha_corte)]
        if df_ejec.empty:
            return pd.DataFrame()

    # Agrupar por período usando la columna de cantidad correcta
    agrupado = df_ejec.groupby(
        pd.Grouper(key="fecha", freq=frecuencia)
    )[col_qty].sum().reset_index()
    agrupado.columns = ["fecha", "cantidad_ejecutada_periodo"]

    agrupado = agrupado[agrupado["cantidad_ejecutada_periodo"] > 0]
    agrupado = agrupado.sort_values("fecha").reset_index(drop=True)
    agrupado["cantidad_ejecutada_acum"] = agrupado["cantidad_ejecutada_periodo"].cumsum()

    # % acumulado real
    if total_ref > 0:
        agrupado["pct_ejecutado_acum"] = (
            agrupado["cantidad_ejecutada_acum"] / total_ref * 100
        ).clip(0, 100).round(2)
    else:
        max_val = agrupado["cantidad_ejecutada_acum"].max()
        agrupado["pct_ejecutado_acum"] = (
            agrupado["cantidad_ejecutada_acum"] / max_val * 100
        ).clip(0, 100).round(2) if max_val > 0 else 0.0

    return agrupado


def tabla_avance_por_item(df_vinculado: pd.DataFrame,
                           df_pyrevit: pd.DataFrame = None) -> pd.DataFrame:
    """
    Genera tabla de control de avance por ITEM con semáforo de estado.
    Columnas: ITEM, Nombre, Inicio, Fin, Unidad, Total Revit, Ejecutado, % Avance, Estado
    """
    if df_vinculado is None or df_vinculado.empty:
        return pd.DataFrame()

    cols_mostrar = []
    for c in ["ITEM", "Nombre de tarea", "Comienzo", "Fin",
               "unidad", "cantidad_total_revit", "cantidad_ejecutada",
               "pct_ejecutado", "elementos_total", "elementos_ejecutados"]:
        if c in df_vinculado.columns:
            cols_mostrar.append(c)

    df_tabla = df_vinculado[cols_mostrar].copy()

    # Renombrar para display
    rename = {
        "Nombre de tarea":      "Actividad",
        "Comienzo":             "Inicio",
        "unidad":               "Unidad",
        "cantidad_total_revit": "Total (Revit)",
        "cantidad_ejecutada":   "Ejecutado",
        "pct_ejecutado":        "% Avance",
        "elementos_total":      "Elem. Total",
        "elementos_ejecutados": "Elem. Ejec.",
    }
    df_tabla = df_tabla.rename(columns={k: v for k, v in rename.items() if k in df_tabla.columns})

    # Semáforo
    if "% Avance" in df_tabla.columns:
        def _estado(p):
            if p == 0:    return "⬜ Sin inicio"
            if p < 25:    return "🔴 Crítico"
            if p < 75:    return "🟡 En progreso"
            if p < 100:   return "🟢 Avanzado"
            return             "✅ Completado"
        df_tabla["Estado"] = df_tabla["% Avance"].apply(_estado)

    # Formato fechas
    for col in ["Inicio", "Fin"]:
        if col in df_tabla.columns:
            df_tabla[col] = pd.to_datetime(df_tabla[col], errors="coerce").dt.strftime("%d-%m-%Y")

    return df_tabla


def avance_por_nivel_pyrevit(df_pyrevit: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula avance ejecutado por nivel/piso desde PyRevit.
    Retorna: nivel, cantidad_total, cantidad_ejecutada, pct_ejecutado, elementos_total, elementos_ejec
    """
    if "nivel" not in df_pyrevit.columns:
        return pd.DataFrame()

    tiene_cantidad = "cantidad_total" in df_pyrevit.columns
    tiene_ejec     = "ejecutado" in df_pyrevit.columns
    rows = []

    for nivel, grp in df_pyrevit.groupby("nivel"):
        row = {"nivel": nivel, "elementos_total": len(grp)}
        ejec = grp[grp["ejecutado"] == True] if tiene_ejec else pd.DataFrame()
        row["elementos_ejec"] = len(ejec)

        if tiene_cantidad:
            total = grp["cantidad_total"].sum()
            ejec_qty = ejec["cantidad_total"].sum() if not ejec.empty else 0.0
            row["cantidad_total"]    = round(total, 3)
            row["cantidad_ejecutada"] = round(ejec_qty, 3)
            row["pct_ejecutado"] = round(ejec_qty / total * 100, 2) if total > 0 else 0.0
        else:
            row["cantidad_total"]    = 0.0
            row["cantidad_ejecutada"] = 0.0
            row["pct_ejecutado"] = round(len(ejec) / len(grp) * 100, 2) if len(grp) > 0 else 0.0

        rows.append(row)

    return pd.DataFrame(rows).sort_values("pct_ejecutado", ascending=False)


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _filtrar_actividades_raiz(nombres: list) -> list:
    if not nombres:
        return []
    profundidades = {n: n.count("/") for n in nombres}
    min_prof = min(profundidades.values())
    return [n for n, p in profundidades.items() if p == min_prof]


def obtener_lista_actividades(datos: dict) -> list:
    if "actividades" in datos:
        return list(datos["actividades"].keys())
    return []


def kpis_semanal(df: pd.DataFrame, semana_actual_idx: int = -1) -> dict:
    if df.empty:
        return {}
    idx = semana_actual_idx if semana_actual_idx >= 0 else len(df) - 1
    idx = min(idx, len(df) - 1)
    prog_actual  = df["acumulado"].iloc[idx]
    prog_total   = df["acumulado"].max()
    semana_label = df["semana_label"].iloc[idx] if "semana_label" in df.columns else f"S{idx+1}"
    return {
        "avance_actual_%":  round(prog_actual, 1),
        "avance_total_%":   round(prog_total, 1),
        "semana_actual":    semana_label,
        "semanas_totales":  len(df),
    }
