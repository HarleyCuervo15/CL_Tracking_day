"""
generar_metas.py
-----------------------------------------------------------
Convierte el archivo pesado de PPTo en un metas.xlsx liviano,
ya normalizado, para que la app solo necesite la maqueta.

    python generar_metas.py "09__Tigo___MKT_Digital_Perf____Sep_2026.xlsx"

Corre esto UNA VEZ AL MES, cuando cambie el presupuesto.
Si el metas.xlsx ya existe, los meses nuevos se agregan y los
meses que ya estaban se reemplazan con la version nueva.
"""

import os
import sys

import pandas as pd

HOJA_PPTO = "Desglose presupuesto"
SALIDA = "metas.xlsx"

NOMBRE_PPTO = {"RRSS": "Meta"}
ALIAS_MIXTO = {"Pmax": "Pmax Mixto", "Bing": "Bing Mixto", "Search": "Search Mixto"}
BLOQUES = ["MOVIL", "FIJO", "MIXTO"]


def normalizar(ruta_ppto):
    from openpyxl import load_workbook

    wb = load_workbook(ruta_ppto, read_only=True, data_only=True)
    if HOJA_PPTO not in wb.sheetnames:
        raise SystemExit(f'El archivo no tiene la hoja "{HOJA_PPTO}".')
    filas = list(wb[HOJA_PPTO].iter_rows(max_col=7, values_only=True))
    wb.close()

    cols = [str(c).strip() if c else "" for c in filas[0]]
    m = pd.DataFrame(filas[1:], columns=cols).dropna(subset=["Fecha"])
    m["FECHA"] = pd.to_datetime(m["Fecha"], errors="coerce")
    m = m.dropna(subset=["FECHA"])

    m["BLOQUE"] = m["Tipo"].astype(str).str.upper().str.strip()
    m["CANAL2"] = (m["Canal unificado"].astype(str).str.strip()
                   .map(lambda c: NOMBRE_PPTO.get(c, c)))
    en_mixto = m["BLOQUE"] == "MIXTO"
    m.loc[en_mixto, "CANAL2"] = m.loc[en_mixto, "CANAL2"].map(
        lambda c: ALIAS_MIXTO.get(c, c))

    m = m.rename(columns={"Meta Costo": "META_COSTO", "Meta Lead": "META_LEADS"})
    # El PPTo a veces trae #DIV/0! como texto: se vuelve 0
    for c in ("META_COSTO", "META_LEADS"):
        m[c] = pd.to_numeric(m[c], errors="coerce").fillna(0.0)
    m = m[m["BLOQUE"].isin(BLOQUES)]

    return (m.groupby(["FECHA", "BLOQUE", "CANAL2"], as_index=False)
            [["META_COSTO", "META_LEADS"]].sum()
            .sort_values(["FECHA", "BLOQUE", "CANAL2"]))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit('Uso: python generar_metas.py "archivo de PPTo.xlsx"')

    nuevas = normalizar(sys.argv[1])
    nuevas["MES"] = nuevas["FECHA"].dt.strftime("%Y-%m")

    if os.path.exists(SALIDA):
        viejas = pd.read_excel(SALIDA)
        viejas["FECHA"] = pd.to_datetime(viejas["FECHA"])
        viejas["MES"] = viejas["FECHA"].dt.strftime("%Y-%m")
        # Los meses que vienen en el archivo nuevo pisan a los viejos
        viejas = viejas[~viejas["MES"].isin(nuevas["MES"].unique())]
        nuevas = pd.concat([viejas, nuevas], ignore_index=True)

    nuevas = (nuevas.drop(columns=["MES"])
              .sort_values(["FECHA", "BLOQUE", "CANAL2"]))
    nuevas.to_excel(SALIDA, index=False)

    print(f"Guardado: {SALIDA}")
    print(f"  {len(nuevas):,} filas")
    resumen = nuevas.copy()
    resumen["MES"] = resumen["FECHA"].dt.strftime("%Y-%m")
    for mes, g in resumen.groupby("MES"):
        mf = g[g["BLOQUE"] != "MIXTO"]
        print(f"  {mes}: {g['FECHA'].nunique():2} dias · "
              f"meta Movil+Fijo ${mf['META_COSTO'].sum():>15,.0f} · "
              f"{mf['META_LEADS'].sum():>8,.0f} leads")
