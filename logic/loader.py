# logic/loader.py
# Funciones de carga y parseo de archivos Excel, JSON y CSV para la App Curvas S
# Desarrollado por Daniel Salcedo - Coordinador BIM

import pandas as pd
import numpy as np
import json
import io
import os
from datetime import datetime


# ─────────────────────────────────────────────────────────────────────────────
# CARGA DESDE EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def load_excel_sheet_names(file) -> list:
    """Retorna los nombres de las hojas del archivo Excel."""
    try:
        xls = pd.ExcelFile(file)
        return xls.sheet_names
    except Exception as e:
        raise ValueError(f"No se pudo leer el archivo Excel: {e}")


def load_datos_sheet(file) -> pd.DataFrame:
    """
    Carga la hoja 'Datos' del Excel.
    Columnas: ITEM, Nombre de tarea, Duración, Días Corridos, Comienzo, Fin, N° días hábiles
    ITEM es la clave que mapea con Assembly Code en Revit.
    """
    try:
        df = pd.read_excel(file, sheet_name="Datos", engine="openpyxl")
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(how="all")

        # Normalizar columna ITEM a string limpio (ej: "3.10", "3.10.1")
        if "ITEM" in df.columns:
            df["ITEM"] = df["ITEM"].astype(str).str.strip()

        # Normalizar fechas
        for col in df.columns:
            col_lower = col.lower()
            if "comienzo" in col_lower or "fin" in col_lower:
                df[col] = pd.to_datetime(df[col], errors="coerce")

        return df
    except Exception as e:
        raise ValueError(f"Error al leer hoja 'Datos': {e}")


def load_avance_semanal(file) -> dict:
    """
    Carga la hoja 'Avance Semanal' del Excel.
    Estructura real: ITEM (col 0) | Actividad (col 1) | Tipo (col 2) | Semanas (col 3+)
    Retorna: {semanas: [...], actividades: {nombre: {item, parcial: [...], acumulado: [...]}}}
    """
    try:
        df_raw = pd.read_excel(file, sheet_name="Avance Semanal",
                               header=0, engine="openpyxl")
        columnas = list(df_raw.columns)
        # Las semanas están desde la columna 3 en adelante
        semanas = columnas[3:]

        semanas_clean = []
        for s in semanas:
            partes = str(s).split("\n")
            codigo = partes[0].strip()
            rango  = partes[1].strip() if len(partes) > 1 else ""
            semanas_clean.append(f"{codigo}\n{rango}")

        actividades = {}
        current_act = None
        current_item = None

        for _, row in df_raw.iterrows():
            item_val = row.iloc[0]   # ITEM (ej: "3.10", "3.10.1")
            act_val  = row.iloc[1]   # Nombre de actividad
            tipo     = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
            valores  = list(row.iloc[3:])  # Valores numéricos por semana

            # Nueva actividad cuando hay nombre en columna Actividad
            if pd.notna(act_val):
                act_str = str(act_val).strip()
                if act_str.lower() in ("nan", "", "actividad", "nombre de tarea", "tarea"):
                    continue
                current_act  = act_str
                current_item = str(item_val).strip() if pd.notna(item_val) else ""
                actividades[current_act] = {
                    "item":      current_item,
                    "parcial":   [],
                    "acumulado": []
                }

            if current_act is None:
                continue

            # Conversión segura: cualquier string en columnas de datos → 0.0
            vals = []
            for v in valores:
                num = pd.to_numeric(v, errors="coerce")
                vals.append(0.0 if pd.isna(num) else float(num))

            if "parcial" in tipo.lower():
                actividades[current_act]["parcial"] = vals
            elif "acumulado" in tipo.lower():
                actividades[current_act]["acumulado"] = vals

        return {"semanas": semanas_clean, "actividades": actividades}
    except Exception as e:
        raise ValueError(f"Error al leer hoja 'Avance Semanal': {e}")



def load_avance_diario(file) -> dict:
    """
    Carga la hoja 'Avance Diario' del Excel.
    Retorna: {fechas: [...], actividades: {nombre: {parcial: [...], acumulado: [...]}}}
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb["Avance Diario"]

        all_rows = list(ws.iter_rows(values_only=True))
        fechas_row = all_rows[1]
        fechas = [f for f in fechas_row[5:] if f is not None]

        actividades  = {}
        current_act  = None

        for row in all_rows[2:]:
            act_name = row[0]
            tipo     = str(row[4]).strip() if row[4] else ""
            valores  = list(row[5:])

            if act_name:
                current_act = str(act_name).strip()
                actividades[current_act] = {
                    "inicio":    row[1],
                    "fin":       row[2],
                    "dias":      row[3],
                    "parcial":   [],
                    "acumulado": []
                }

            if current_act is None:
                continue

            # Conversión segura: ignora strings inesperados en columnas de datos
            vals = [
                float(pd.to_numeric(v, errors="coerce") or 0.0)
                if v is not None else 0.0
                for v in valores[:len(fechas)]
            ]

            if "parcial" in tipo.lower():
                actividades[current_act]["parcial"] = vals
            elif "acumulado" in tipo.lower():
                actividades[current_act]["acumulado"] = vals

        return {"fechas": fechas, "actividades": actividades}
    except Exception as e:
        raise ValueError(f"Error al leer hoja 'Avance Diario': {e}")


# ─────────────────────────────────────────────────────────────────────────────
# CARGA DESDE PYREVIT (JSON / CSV)
# Formato esperado:
#   - assembly_code : str   → coincide con ITEM del Excel (ej: "3.10.2.1")
#   - nombre        : str   → nombre del elemento Revit (opcional)
#   - unidad        : str   → "m2" | "ml" | "m3"
#   - cantidad_total: float → cantidad total del elemento
#   - ejecutado     : bool  → True si el elemento está ejecutado
#   - fecha         : str   → fecha de ejecución (YYYY-MM-DD)
#   - nivel         : str   → nivel/piso en Revit (opcional)
# ─────────────────────────────────────────────────────────────────────────────

# Alias de columnas aceptados desde PyRevit (tolerante a variaciones de nombre)
_ALIAS_ASSEMBLY      = ["assembly_code", "assemblycode", "assembly_code", "codigo_montaje", "código_de_montaje",
                        "codigo_de_montaje", "codigo_ensamblaje", "item", "cod_item"]
_ALIAS_ASSEMBLY_DESC = ["assembly_description", "assemblydescription", "descripción_de_montaje",
                        "descripcion_de_montaje", "descrip_montaje", "descripcion_montaje"]
_ALIAS_NOMBRE        = ["nombre", "name", "elemento", "element", "descripcion", "description", "tarea"]
_ALIAS_UNIDAD        = ["unidad", "unit", "ud", "tipo_unidad"]
_ALIAS_CANTIDAD      = ["cantidad_total", "cantidad", "quantity", "total", "metrado",
                        "area", "longitud", "volumen", "valor"]
_ALIAS_EJECUTADO     = ["ejecutado", "executed", "done", "completado", "completed", "status"]
_ALIAS_FECHA         = ["fecha", "date", "fecha_ejecucion", "execution_date", "periodo", "ei_fecha_ejecucion"]
_ALIAS_NIVEL         = ["nivel", "level", "piso", "floor", "planta"]


def _normalizar_columnas_pyrevit(df: pd.DataFrame) -> pd.DataFrame:
    """
    Renombra columnas del DataFrame de PyRevit a nombres estándar internos.
    Tolerante a variaciones de mayúsculas, espacios y alias. Evita columnas duplicadas.
    """
    col_map = {}
    for col in df.columns:
        lc = str(col).lower().strip().replace(" ", "_")
        if lc in _ALIAS_ASSEMBLY and "assembly_code" not in col_map.values():
            col_map[col] = "assembly_code"
        elif lc in _ALIAS_ASSEMBLY_DESC and "assembly_description" not in col_map.values():
            col_map[col] = "assembly_description"
        elif lc in _ALIAS_NOMBRE and "nombre" not in col_map.values():
            col_map[col] = "nombre"
        elif lc in _ALIAS_UNIDAD and "unidad" not in col_map.values():
            col_map[col] = "unidad"
        elif lc in _ALIAS_CANTIDAD and "cantidad_total" not in col_map.values():
            col_map[col] = "cantidad_total"
        elif lc in _ALIAS_EJECUTADO and "ejecutado" not in col_map.values():
            col_map[col] = "ejecutado"
        elif lc in _ALIAS_FECHA and "fecha" not in col_map.values():
            col_map[col] = "fecha"
        elif lc in _ALIAS_NIVEL and "nivel" not in col_map.values():
            col_map[col] = "nivel"

    df = df.rename(columns=col_map)
    df = df.loc[:, ~df.columns.duplicated(keep="first")]

    # Convertir tipos
    if "fecha" in df.columns:
        df["fecha"] = pd.to_datetime(df["fecha"], format="mixed", dayfirst=True, errors="coerce")
        if "ei_fecha_ejecucion" in df.columns:
            df["fecha"] = df["fecha"].fillna(pd.to_datetime(df["ei_fecha_ejecucion"], format="mixed", dayfirst=True, errors="coerce"))
        # Corregir posibles años erróneos de tipeo en Revit (ej: 2926 -> 2026)
        df["fecha"] = df["fecha"].apply(lambda d: d.replace(year=2026) if (pd.notna(d) and d.year > 2050) else d)
    elif "ei_fecha_ejecucion" in df.columns:
        df["fecha"] = pd.to_datetime(df["ei_fecha_ejecucion"], format="mixed", dayfirst=True, errors="coerce")
        df["fecha"] = df["fecha"].apply(lambda d: d.replace(year=2026) if (pd.notna(d) and d.year > 2050) else d)

    if "cantidad_total" in df.columns:
        df["cantidad_total"] = pd.to_numeric(df["cantidad_total"], errors="coerce").fillna(0.0)

    if "ejecutado" in df.columns:
        # Normalizar a booleano (acepta: True, "true", "si", "yes", 1, "1")
        def _to_bool(v):
            if isinstance(v, bool):
                return v
            if isinstance(v, (int, float)):
                return bool(v)
            s = str(v).strip().lower()
            return s in ["true", "verdadero", "si", "sí", "yes", "1", "x", "ejecutado"]
        df["ejecutado"] = df["ejecutado"].apply(_to_bool)

    if "assembly_code" in df.columns:
        df["assembly_code"] = df["assembly_code"].astype(str).str.strip()

    return df


def load_pyrevit_json(file) -> pd.DataFrame:
    """
    Carga avances desde JSON exportado por PyRevit.
    Formato esperado: lista de objetos con assembly_code, cantidad_total, ejecutado, fecha, etc.
    """
    try:
        if hasattr(file, "read"):
            content = file.read()
            if isinstance(content, bytes):
                content = content.decode("utf-8")
        elif isinstance(file, bytes):
            content = file.decode("utf-8")
        elif isinstance(file, str):
            if file.strip().startswith("[") or file.strip().startswith("{"):
                content = file
            elif os.path.exists(file):
                with open(file, "r", encoding="utf-8") as f:
                    content = f.read()
            else:
                content = file
        else:
            raise ValueError("Tipo de archivo no soportado para lectura JSON.")

        data = json.loads(content)

        if isinstance(data, dict):
            # Buscar la lista en claves comunes
            for key in ["data", "elements", "elementos", "items", "registros"]:
                if key in data and isinstance(data[key], list):
                    data = data[key]
                    break
            else:
                # Tomar el primer valor que sea lista
                for v in data.values():
                    if isinstance(v, list):
                        data = v
                        break

        if not isinstance(data, list):
            raise ValueError("Formato JSON no reconocido. Se esperaba una lista de objetos.")

        df = pd.DataFrame(data)
        df = _normalizar_columnas_pyrevit(df)
        return df

    except Exception as e:
        raise ValueError(f"Error al cargar JSON de PyRevit: {e}")


def load_pyrevit_csv(file) -> pd.DataFrame:
    """
    Carga avances desde CSV exportado por PyRevit.
    Detecta automáticamente separador (coma o punto y coma).
    """
    try:
        if hasattr(file, "read"):
            content = file.read()
            if isinstance(content, bytes):
                content = content.decode("utf-8-sig")
        else:
            with open(file, "r", encoding="utf-8-sig") as f:
                content = f.read()

        sep = ";" if content.count(";") > content.count(",") else ","
        df  = pd.read_csv(io.StringIO(content), sep=sep)
        df  = _normalizar_columnas_pyrevit(df)
        return df

    except Exception as e:
        raise ValueError(f"Error al cargar CSV de PyRevit: {e}")


def validar_columnas_pyrevit(df: pd.DataFrame) -> dict:
    """
    Valida que el DataFrame de PyRevit tenga las columnas mínimas requeridas.
    Retorna dict con: ok (bool), columnas_encontradas, columnas_faltantes, advertencias.
    """
    requeridas   = {"assembly_code", "cantidad_total", "ejecutado", "fecha"}
    opcionales   = {"nombre", "unidad", "nivel"}
    encontradas  = set(df.columns)
    faltantes    = requeridas - encontradas
    advertencias = []

    if "ejecutado" in encontradas:
        n_true  = df["ejecutado"].sum() if df["ejecutado"].dtype == bool else 0
        n_total = len(df)
        advertencias.append(f"{n_true}/{n_total} elementos marcados como Ejecutado=True")

    if "assembly_code" in encontradas:
        n_unicos = df["assembly_code"].nunique()
        advertencias.append(f"{n_unicos} Assembly Codes únicos en el archivo")

    return {
        "ok":                   len(faltantes) == 0,
        "columnas_encontradas": list(encontradas),
        "columnas_faltantes":   list(faltantes),
        "advertencias":         advertencias,
    }


def resumen_pyrevit(df: pd.DataFrame) -> dict:
    """Genera resumen estadístico del DataFrame de PyRevit."""
    resumen = {"total_elementos": len(df)}

    if "ejecutado" in df.columns:
        n_ejec = int(df["ejecutado"].sum())
        resumen["elementos_ejecutados"]  = n_ejec
        resumen["elementos_pendientes"]  = len(df) - n_ejec
        resumen["pct_ejecutado_count"]   = round(n_ejec / len(df) * 100, 1) if len(df) > 0 else 0

    if "cantidad_total" in df.columns:
        resumen["cantidad_total_sum"]     = round(df["cantidad_total"].sum(), 2)
        if "ejecutado" in df.columns:
            ejec = df[df["ejecutado"] == True]["cantidad_total"].sum()
            total = df["cantidad_total"].sum()
            resumen["cantidad_ejecutada"] = round(ejec, 2)
            resumen["pct_cantidad_ejec"]  = round(ejec / total * 100, 1) if total > 0 else 0

    if "assembly_code" in df.columns:
        resumen["assembly_codes_unicos"] = df["assembly_code"].nunique()

    if "nivel" in df.columns:
        resumen["niveles"] = df["nivel"].nunique()

    if "fecha" in df.columns:
        resumen["fecha_min"] = df["fecha"].min()
        resumen["fecha_max"] = df["fecha"].max()

    return resumen
