"""
generar_metas.py  ·  v2
-----------------------------------------------------------
Lee el archivo "Metas y presupuestos" (una hoja por tipo, con bloques
de Presupuesto / Leads / Ventas en columnas por dia) y lo convierte en
metas.xlsx, liviano y normalizado.

    python generar_metas.py "Metas_y_presupuestos_-_Septiembre_2026.xlsx"

MIXTO se reparte 50/50 hacia Movil y Fijo, igual que en el real. Eso
reproduce el 364.320.000 que la propia hoja Fijo calcula
(304.320.000 propio + 60.000.000 de la mitad de Mixto).
"""

import os
import sys

import pandas as pd
from openpyxl import load_workbook

SALIDA = "metas.xlsx"
HOJAS = {"Movil": "MOVIL", "Fijo": "FIJO", "Mixto": "MIXTO"}
SPLIT_MIXTO = 0.50
CLASES = ["META_COSTO", "META_LEADS", "META_VENTAS"]

# (Categoria, CAMPANAS) -> nombre corto. Se revisa primero el par exacto.
PARES = {
    ("SEARCH", "SEARCH"): "Search", ("SEARCH", "SEM"): "Search",
    ("SEM", "SEM"): "Search", ("SEM", "SEARCH"): "Search",
    ("DISPLAY", "PMAX"): "Pmax", ("DISPLAY", "DISPLAY"): "Display",
    ("DISPLAY", "DEMAND GEN"): "Demand Gen",
    ("META", "META"): "Meta", ("RRSS", "RRSS"): "Meta",
    ("BING", "BING"): "Bing",
    ("BING", "PMAX"): "Bing",      # en Mixto, Bing viene rotulado como PMAX
}
SUELTOS = {"SEARCH": "Search", "SEM": "Search", "PMAX": "Pmax",
           "DISPLAY": "Display", "META": "Meta", "BING": "Bing",
           "DEMAND GEN": "Demand Gen"}


def nombre_canal(categoria, campana, avisos):
    cat = str(categoria or "").strip().upper()
    cam = str(campana or "").strip().upper()
    if (cat, cam) in PARES:
        return PARES[(cat, cam)]
    if cam in SUELTOS:
        avisos.append(f"{cat} / {cam}: se usa CAMPANAS -> {SUELTOS[cam]}")
        return SUELTOS[cam]
    if cat in SUELTOS:
        avisos.append(f"{cat} / {cam}: se usa Categoria -> {SUELTOS[cat]}")
        return SUELTOS[cat]
    avisos.append(f"{cat} / {cam}: SIN MAPEO, queda como '{cam.title()}'")
    return cam.title() or cat.title()


def encabezados_de(ws):
    """Filas donde empieza un bloque, con el titulo que lo precede."""
    salida, titulo = [], ""
    for i, fila in enumerate(ws.iter_rows(max_col=1, values_only=True), 1):
        a = str(fila[0] or "").strip()
        if not a:
            continue
        if a.lower() == "categoria":
            salida.append((i, titulo))
        elif any(k in a.lower() for k in ("presupuesto", "leads", "ventas")):
            titulo = a
    return salida


def leer_bloque(ws, fila_hdr, titulo, mes, avisos):
    """Lee un bloque. Para saber si trae costo, leads o ventas compara la
    suma diaria contra 'Propuesta Presupuesto' y 'Estimacion Leads'.
    Asi no dependemos del rotulo, que a veces esta mal puesto."""
    hdr = [c.value for c in ws[fila_hdr]]
    cols = [(j, v) for j, v in enumerate(hdr, 1)
            if hasattr(v, "year") and v.strftime("%Y-%m") == mes]
    if not cols:
        return []

    filas = []
    for i in range(fila_hdr + 1, ws.max_row + 1):
        cat = ws.cell(row=i, column=1).value
        cam = ws.cell(row=i, column=2).value
        etq = str(cat or "").strip().lower()
        etq2 = str(cam or "").strip().lower()
        # Corta en la fila de totales (el rotulo puede venir en la 1a o la 2a
        # columna), al empezar otro bloque, o cuando la fila esta vacia.
        if (etq in ("total", "categoria") or etq2 == "total"
                or any(k in etq for k in ("presupuesto", "leads diario",
                                          "leads diarios", "ventas diario"))
                or (not etq and not etq2)):
            break
        vals = {}
        for j, f in cols:
            v = ws.cell(row=i, column=j).value
            vals[f] = float(v) if isinstance(v, (int, float)) else 0.0
        filas.append((cat, cam, ws.cell(row=i, column=4).value,
                      ws.cell(row=i, column=5).value, vals))

    if not filas:
        return []

    suma = sum(sum(f[4].values()) for f in filas)
    presup = sum(f[2] for f in filas if isinstance(f[2], (int, float)))
    estim = sum(f[3] for f in filas if isinstance(f[3], (int, float)))

    def cerca(a, b):
        return bool(b) and abs(a - b) / abs(b) < 0.02

    if cerca(suma, presup):
        clase = "META_COSTO"
    elif "venta" in titulo.lower():
        clase = "META_VENTAS"
    elif cerca(suma, estim):
        clase = "META_LEADS"
    else:
        avisos.append(f"bloque de la fila {fila_hdr} ('{titulo}'): no clasificado")
        return []

    salida = []
    for cat, cam, _, _, vals in filas:
        canal = nombre_canal(cat, cam, avisos)
        for fecha, v in vals.items():
            if v:
                salida.append({"FECHA": fecha, "CANAL2": canal, clase: v})
    return salida


def normalizar(ruta, mes=None):
    wb = load_workbook(ruta, data_only=True)
    avisos = []

    if mes is None:                       # el mes que mas se repite
        vistos = {}
        for hoja in HOJAS:
            if hoja in wb.sheetnames:
                for c in wb[hoja][2]:
                    if hasattr(c.value, "year"):
                        k = c.value.strftime("%Y-%m")
                        vistos[k] = vistos.get(k, 0) + 1
        if not vistos:
            raise SystemExit("No encontre columnas de fecha en las hojas de metas.")
        mes = max(vistos, key=vistos.get)

    partes = []
    for hoja, bloque in HOJAS.items():
        if hoja not in wb.sheetnames:
            avisos.append(f"falta la hoja '{hoja}'")
            continue
        ws = wb[hoja]
        registros = []
        for fila_hdr, titulo in encabezados_de(ws):
            registros += leer_bloque(ws, fila_hdr, titulo, mes, avisos)
        if not registros:
            avisos.append(f"la hoja '{hoja}' no aporto datos")
            continue
        d = (pd.DataFrame(registros)
             .groupby(["FECHA", "CANAL2"], as_index=False).sum(numeric_only=True))
        d["BLOQUE"] = bloque
        partes.append(d)
    wb.close()

    if not partes:
        raise SystemExit("No se pudo leer ninguna hoja de metas.")

    todo = pd.concat(partes, ignore_index=True)
    for c in CLASES:
        if c not in todo.columns:
            todo[c] = 0.0
    todo[CLASES] = todo[CLASES].fillna(0.0)

    # MIXTO queda entero en su bloque y ademas se reparte a Movil y Fijo
    mixto = todo[todo["BLOQUE"] == "MIXTO"].copy()
    piezas = [todo[todo["BLOQUE"] != "MIXTO"].copy()]
    if not mixto.empty:
        mixto["CANAL2"] = mixto["CANAL2"] + " Mixto"
        piezas.append(mixto)
        for destino in ("MOVIL", "FIJO"):
            t = mixto.copy()
            t["BLOQUE"] = destino
            t[CLASES] = t[CLASES] * SPLIT_MIXTO
            piezas.append(t)

    final = (pd.concat(piezas, ignore_index=True)
             .groupby(["FECHA", "BLOQUE", "CANAL2"], as_index=False)[CLASES].sum()
             .sort_values(["FECHA", "BLOQUE", "CANAL2"]))
    final["FECHA"] = pd.to_datetime(final["FECHA"])
    return final, mes, avisos


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit('Uso: python generar_metas.py "Metas y presupuestos.xlsx"')

    nuevas, mes, avisos = normalizar(sys.argv[1],
                                     sys.argv[2] if len(sys.argv) > 2 else None)
    nuevas["MES"] = nuevas["FECHA"].dt.strftime("%Y-%m")

    if os.path.exists(SALIDA):
        viejas = pd.read_excel(SALIDA)
        viejas["FECHA"] = pd.to_datetime(viejas["FECHA"])
        viejas["MES"] = viejas["FECHA"].dt.strftime("%Y-%m")
        for c in CLASES:
            if c not in viejas.columns:
                viejas[c] = 0.0
        viejas = viejas[~viejas["MES"].isin(nuevas["MES"].unique())]
        nuevas = pd.concat([viejas, nuevas], ignore_index=True)

    nuevas = nuevas.drop(columns=["MES"]).sort_values(["FECHA", "BLOQUE", "CANAL2"])
    nuevas.to_excel(SALIDA, index=False)

    print(f"Guardado: {SALIDA}  ({len(nuevas):,} filas)")
    r = nuevas.copy()
    r["MES"] = r["FECHA"].dt.strftime("%Y-%m")
    for m, g in r.groupby("MES"):
        for bloque in ("MOVIL", "FIJO", "MIXTO"):
            b = g[g["BLOQUE"] == bloque]
            if len(b):
                print(f"  {m} {bloque:6} ${b['META_COSTO'].sum():>15,.0f} · "
                      f"{b['META_LEADS'].sum():>9,.0f} leads · "
                      f"{b['META_VENTAS'].sum():>8,.0f} ventas")
    if avisos:
        print("\nAvisos:")
        for a in dict.fromkeys(avisos):
            print("  -", a)
