"""Análisis de la extracción Noetix_Tescom_Automatico (MC1): existencias, WIP y P/N sin movimiento.

Lee solo las hojas de datos (OnHand, Dispatch_WO, WIP_Completions, WO_SO); NUNCA abre LOGIN
(esa hoja guarda usuario/contraseña). Solo lectura; no se conecta a Noetix ni a Oracle.
Uso:  python3 -I noetix_analisis.py ARCHIVO.xlsx [DIAS_SIN_MOV] [SALIDA_DIR]
"""
import sys, pathlib
import pandas as pd

path = pathlib.Path(sys.argv[1])
dias = int(sys.argv[2]) if len(sys.argv) > 2 else 180
out = pathlib.Path(sys.argv[3]) if len(sys.argv) > 3 else pathlib.Path(".")
out.mkdir(parents=True, exist_ok=True)

def hoja(nombre):
    # fila 1 = título, fila 2 = aviso, fila 4 = encabezados
    df = pd.read_excel(path, sheet_name=nombre, header=3, engine="openpyxl")
    return df.dropna(how="all")

oh = hoja("OnHand"); wip = hoja("WIP_Completions"); dsp = hoja("Dispatch_WO"); wso = hoja("WO_SO")
ITEM = "Item SV Xxinv Item Number"

res = {}
res["onhand_filas"] = len(oh); res["onhand_pn"] = oh[ITEM].nunique()
res["onhand_cantidad_total"] = float(oh["Quantity"].sum())
por_item = (oh.groupby([ITEM, "Item Description", "Planner_Code", "Item_Type"], dropna=False)["Quantity"]
              .sum().reset_index().rename(columns={"Quantity": "Existencia"}))
por_subinv = oh.groupby("Subinventory")["Quantity"].agg(["count", "sum"]).sort_values("sum", ascending=False)

wip["Transaction Date"] = pd.to_datetime(wip["Transaction Date"], errors="coerce")
res["wip_rango"] = (str(wip["Transaction Date"].min()), str(wip["Transaction Date"].max()))
res["wip_filas"] = len(wip)
ult = wip.groupby(ITEM)["Transaction Date"].max().rename("UltimaTerminacionWIP").reset_index()
cant = wip.groupby(ITEM)["Transaction Quantity"].sum().rename("CantTerminadaPeriodo").reset_index()
base = por_item.merge(ult, on=ITEM, how="left").merge(cant, on=ITEM, how="left")
corte = wip["Transaction Date"].max()
base["DiasDesdeUltTerminacion"] = (corte - base["UltimaTerminacionWIP"]).dt.days
# Solo tiene sentido para ítems que la planta TERMINA en WIP (no comprados: PURCHASED/BUY ITEM)
FABRICADOS = {"MAKE ITEM", "FINISHED GOOD", "MTO ITEM", "CONFIGURED ITEM", "KIT", "SERVICE KIT"}
fabricado = base["Item_Type"].isin(FABRICADOS)
sin_mov = base[(base["Existencia"] > 0) & fabricado &
               (base["UltimaTerminacionWIP"].isna() | (base["DiasDesdeUltTerminacion"] > dias))]
res["tipos_item"] = oh["Item_Type"].value_counts(dropna=False).to_dict()
res["sin_mov_fabricados"] = len(sin_mov); res["sin_mov_dias"] = dias

dsp_status = dsp["Job Status"].value_counts(dropna=False).to_dict()
res["dispatch_filas"] = len(dsp); res["dispatch_status"] = dsp_status
res["dispatch_jobs"] = int(dsp["Job Name"].nunique())
res["wo_so_filas"] = len(wso)
res["wo_so_status_linea"] = wso["Order_Line_Status"].value_counts(dropna=False).head(6).to_dict()

base.sort_values("Existencia", ascending=False).to_csv(out / "existencia_por_pn.csv", index=False)
sin_mov.sort_values("Existencia", ascending=False).to_csv(out / "pn_sin_movimiento_fabricados.csv", index=False)
por_subinv.to_csv(out / "existencia_por_subinventario.csv")
for k, v in res.items(): print(k, ":", v)
print("\nTOP subinventarios por cantidad:\n", por_subinv.head(8))
print("\nTOP 8 P/N sin movimiento (fabricados):\n", sin_mov.sort_values("Existencia", ascending=False).head(8)[[ITEM, "Planner_Code", "Item_Type", "Existencia", "UltimaTerminacionWIP"]].to_string(index=False))
