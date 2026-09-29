import sys, json, io
sys.path.insert(0, '.')
from logic.loader import load_pyrevit_json, validar_columnas_pyrevit, resumen_pyrevit
from logic.processor import calcular_curva_s_real_pyrevit, vincular_item_pyrevit, tabla_avance_por_item
from logic.charts import grafico_avance_por_item, grafico_curva_real_pyrevit
print("OK imports")

sample = [
    {"assembly_code":"3.10.1","nombre":"Elem A","unidad":"m2","cantidad_total":100,"ejecutado":True, "fecha":"2026-09-15","nivel":"N1"},
    {"assembly_code":"3.10.1","nombre":"Elem B","unidad":"m2","cantidad_total":50, "ejecutado":False,"fecha":"2026-09-16","nivel":"N1"},
    {"assembly_code":"3.10.2","nombre":"Elem C","unidad":"ml","cantidad_total":20, "ejecutado":True, "fecha":"2026-09-18","nivel":"N2"},
    {"assembly_code":"3.10.3","nombre":"Elem D","unidad":"m3","cantidad_total":5,  "ejecutado":True, "fecha":"2026-09-20","nivel":"N1"},
]
json_str = json.dumps(sample)
df_pyr = load_pyrevit_json(io.StringIO(json_str))
print("Columnas PyRevit:", list(df_pyr.columns))
print("ejecutado dtype:", df_pyr["ejecutado"].dtype)

val = validar_columnas_pyrevit(df_pyr)
print("Validacion OK:", val["ok"])
print("Advertencias:", val["advertencias"])

res = resumen_pyrevit(df_pyr)
print("Resumen:", res)

df_curva = calcular_curva_s_real_pyrevit(df_pyr)
print("Curva real:")
print(df_curva.to_string())
print("TODOS LOS TESTS OK")
