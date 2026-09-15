"""
tracking_cpe.py  ·  v2
-----------------------------------------------------------
Cruza la BASE REAL (que vive dentro del .xlsx de la maqueta) con las
METAS DIARIAS (hoja "Desglose presupuesto" del archivo de PPTo) y arma
el tracking diario + la proyeccion a cierre de mes.

    python tracking_cpe.py maqueta.xlsx ppto.xlsx 2026-09

Genera Tracking_CPE_<mes>.xlsx con:
    MAQUETA    -> los 9 bloques del reporte + resumen del mes
    CURVA      -> dia a dia: meta vs real, % cumplimiento y acumulados
    CONTROL    -> combos que tienen real sin meta o meta sin real
    PARAMETROS -> celdas amarillas editables
    DIARIO     -> real por fecha / bloque / canal
    METAS_DIA  -> meta por fecha / bloque / canal

Metodo de proyeccion: run-rate simple.
    Proyectado = Real acumulado / dias con informacion x dias del mes
Es el mismo que ya usa la hoja "Metas mensuales" del PPTo (real x 30/8).
"""

import calendar
import glob
import os
import sys
from datetime import date

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from extraer_base import extraer

# ----------------------------------------------------------------
# CONFIGURACION
# ----------------------------------------------------------------
# Carpeta donde dejas los dos archivos de entrada.
CARPETA_INPUTS = (r"C:\Users\harley.cuervo\OneDrive - Havas\Escritorio"
                  r"\Proyecto Havas\1. Tigo\CL tipificaciones\Proceso\0. Inputs")

# Carpeta donde se guarda el resultado. Se crea sola si no existe.
CARPETA_OUTPUTS = (r"C:\Users\harley.cuervo\OneDrive - Havas\Escritorio"
                   r"\Proyecto Havas\1. Tigo\CL tipificaciones\Proceso\1. Outputs")

# Como reconocer cada archivo. Si cambia el nombre, ajusta el patron.
PATRON_MAQUETA = "*aqueta*CPE*.xlsx"        # se actualiza a diario
PATRON_PPTO = "*MKT*Digital*Perf*.xlsx"     # se actualiza una vez al mes

AREA = "CANAL ONLINE"
BLOQUES = ["MOVIL", "FIJO", "MIXTO"]
SPLIT_MIXTO = 0.50          # hoja "Split" del PPTo: Movil 0,5 / Fijo 0,5
HOJA_PPTO = "Desglose presupuesto"

# Columnas de volumen que se arrastran hasta el reporte.
# Q_NETO = leads calificados · Q_EMI = venta cantada · Q_TER = venta
METRICAS = ["COSTO", "LEADS", "Q_NETO", "Q_EMI", "Q_TER"]

# Costos unitarios, tal como los define el propio archivo de la maqueta:
#   CPL = COSTO/LEADS · CPL_neto = COSTO/Q_NETO · CPE = COSTO/Q_EMI · CPA = COSTO/Q_TER
# (col_volumen en DIARIO, etiqueta del volumen, etiqueta del costo unitario)
EMBUDO = [("F", "Q Neto", "CPL Neto"),
          ("G", "Q Emi", "CPE"),
          ("H", "Q Ter", "CPA")]

# True  = solo canales con inversion (deja fuera Afiliados, TikTok, Sin clasificar)
# False = todos los canales, aunque no tengan costo (ahi se ven sus ventas)
SOLO_CON_COSTO = False

# CANAL de la base real -> nombre corto
NOMBRE_CANAL = {
    "BING": "Bing", "PMAX": "Pmax", "RRSS": "Meta", "Search": "Search",
    "DISPLAY": "Display", "VIDEO": "Video", "TIKTOK": "TikTok",
}

# "Canal unificado" del PPTo -> mismo nombre corto
NOMBRE_PPTO = {"RRSS": "Meta"}

# El PPTo nombra distinto algunos canales dentro del bloque Mixto
ALIAS_MIXTO = {"Pmax": "Pmax Mixto", "Bing": "Bing Mixto", "Search": "Search Mixto"}

FUENTE = "Arial"
AZUL, CELESTE, AMARILLO, GRIS = "4472C4", "2E9BE0", "FFFF00", "F2F2F2"
VERDE, ROJO = "C6EFCE", "FFC7CE"
GRIS_T = "D9D9D9"


# ----------------------------------------------------------------
# 0. ENCONTRAR LOS ARCHIVOS SOLOS
# ----------------------------------------------------------------
def buscar(carpeta, patron, etiqueta):
    """Devuelve el archivo mas reciente que coincida con el patron."""
    if not os.path.isdir(carpeta):
        raise SystemExit(f"No existe la carpeta:\n  {carpeta}")

    candidatos = [f for f in glob.glob(os.path.join(carpeta, patron))
                  if not os.path.basename(f).startswith("~$")]
    if not candidatos:
        hay = "\n  ".join(sorted(os.path.basename(f)
                                 for f in glob.glob(os.path.join(carpeta, "*.xlsx"))
                                 if not os.path.basename(f).startswith("~$"))) or "(vacia)"
        raise SystemExit(
            f"No encontre el archivo de {etiqueta} con el patron {patron}\n"
            f"En la carpeta hay:\n  {hay}\n"
            f"Renombra el archivo o ajusta el patron arriba en el script.")

    elegido = max(candidatos, key=os.path.getmtime)
    if len(candidatos) > 1:
        print(f"     Ojo: hay {len(candidatos)} archivos de {etiqueta}. "
              f"Uso el mas reciente.")
    return elegido


# ----------------------------------------------------------------
# 1. REAL
# ----------------------------------------------------------------
def construir_diario(df, mes, area=None, split=None, solo_con_costo=None,
                     filtros=None):
    """filtros: {"PRODUCTO": [...], "CANAL_SOLICITUD": [...]} o None.
    Una lista vacia significa 'no filtrar por esa columna'."""
    area = AREA if area is None else area
    split = SPLIT_MIXTO if split is None else split
    solo = SOLO_CON_COSTO if solo_con_costo is None else solo_con_costo

    df = df[(df["MES"] == mes) & (df["AREA"] == area)]
    for col, valores in (filtros or {}).items():
        if valores and col in df.columns:
            df = df[df[col].astype(str).isin(valores)]
    df = df.copy()
    if df.empty:
        return df
    df["CANAL2"] = df["CANAL"].astype(str).map(
        lambda c: NOMBRE_CANAL.get(c, str(c).title()))
    df["TIPO"] = df["TIPO"].astype(str)

    partes = []
    for bloque in ("MOVIL", "FIJO"):
        p = df[df["TIPO"] == bloque].copy()
        p["BLOQUE"] = bloque
        partes.append(p)

    mixto = df[df["TIPO"] == "MIXTO"].copy()
    if not mixto.empty:
        mixto["CANAL2"] = mixto["CANAL2"].map(lambda c: ALIAS_MIXTO.get(c, c + " Mixto"))

        entero = mixto.copy()
        entero["BLOQUE"] = "MIXTO"
        partes.append(entero)

        for bloque in ("MOVIL", "FIJO"):
            trozo = mixto.copy()
            trozo["BLOQUE"] = bloque
            trozo[METRICAS] = trozo[METRICAS] * split
            partes.append(trozo)

    diario = (
        pd.concat(partes, ignore_index=True)
        .groupby(["FECHA", "BLOQUE", "CANAL2"], as_index=False, observed=True)[METRICAS].sum()
        .sort_values(["FECHA", "BLOQUE", "CANAL2"])
    )
    if solo:
        con_costo = diario.groupby("CANAL2")["COSTO"].transform("sum") > 0
        diario = diario[con_costo]
    return diario[diario[METRICAS].abs().sum(axis=1) > 0]


# ----------------------------------------------------------------
# 2. METAS DIARIAS DESDE EL PPTO
# ----------------------------------------------------------------
def leer_metas(ruta_ppto, mes):
    wb = load_workbook(ruta_ppto, read_only=True, data_only=True)
    if HOJA_PPTO not in wb.sheetnames:
        raise SystemExit(f'El archivo de PPTo no tiene la hoja "{HOJA_PPTO}".')

    filas = list(wb[HOJA_PPTO].iter_rows(max_col=7, values_only=True))
    wb.close()

    cols = [str(c).strip() if c else "" for c in filas[0]]
    m = pd.DataFrame(filas[1:], columns=cols).dropna(subset=["Fecha"])
    m["Fecha"] = pd.to_datetime(m["Fecha"], errors="coerce")
    m = m[m["Fecha"].dt.strftime("%Y-%m") == mes].copy()
    if m.empty:
        raise SystemExit(f'La hoja "{HOJA_PPTO}" no tiene metas para {mes}.')

    m["BLOQUE"] = m["Tipo"].astype(str).str.upper().str.strip()
    base = m["Canal unificado"].astype(str).str.strip()
    m["CANAL2"] = base.map(lambda c: NOMBRE_PPTO.get(c, c))
    # Dentro del bloque Mixto el PPTo escribe "Pmax" donde deberia decir "Pmax Mixto"
    en_mixto = m["BLOQUE"] == "MIXTO"
    m.loc[en_mixto, "CANAL2"] = m.loc[en_mixto, "CANAL2"].map(
        lambda c: ALIAS_MIXTO.get(c, c))

    m = m.rename(columns={"Fecha": "FECHA", "Meta Costo": "META_COSTO",
                          "Meta Lead": "META_LEADS"})
    metas = (m.groupby(["FECHA", "BLOQUE", "CANAL2"], as_index=False)
             [["META_COSTO", "META_LEADS"]].sum()
             .sort_values(["FECHA", "BLOQUE", "CANAL2"]))
    return metas[metas["BLOQUE"].isin(BLOQUES)]


def universo(split=None):
    """Bloques que suman el total sin contar dos veces.
    Si Mixto se reparte, ya esta dentro de Movil y Fijo. Si no, va aparte
    y hay que sumarlo."""
    split = SPLIT_MIXTO if split is None else split
    return ["MOVIL", "FIJO"] if split else ["MOVIL", "FIJO", "MIXTO"]


def repartir_metas(metas, split=None):
    """Mixto queda entero en su bloque y ademas se reparte a Movil y Fijo,
    exactamente igual que el real. Con split=0 los bloques quedan puros:
    Movil solo movil, Fijo solo fijo, y Mixto aparte."""
    split = SPLIT_MIXTO if split is None else split
    if metas is None or metas.empty:
        return metas
    cols = [c for c in ("META_COSTO", "META_LEADS", "META_VENTAS")
            if c in metas.columns]

    mixto = metas[metas["BLOQUE"] == "MIXTO"]
    piezas = [metas]
    if not mixto.empty and split:
        for destino in ("MOVIL", "FIJO"):
            t = mixto.copy()
            t["BLOQUE"] = destino
            t[cols] = t[cols] * split
            piezas.append(t)

    return (pd.concat(piezas, ignore_index=True)
            .groupby(["FECHA", "BLOQUE", "CANAL2"], as_index=False)[cols].sum()
            .sort_values(["FECHA", "BLOQUE", "CANAL2"]))


# ----------------------------------------------------------------
# 3. LIBRO
# ----------------------------------------------------------------
def titulo(ws, ini, fin, texto, color):
    ws.merge_cells(f"{ini}:{fin}")
    c = ws[ini]
    c.value, c.font = texto, Font(name=FUENTE, bold=True, size=12, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=color)
    c.alignment = Alignment(horizontal="center", vertical="center")


def encabezado(ws, fila, col_ini, textos):
    for i, t in enumerate(textos):
        c = ws.cell(row=fila, column=col_ini + i, value=t)
        c.font = Font(name=FUENTE, bold=True, size=9)
        c.fill = PatternFill("solid", fgColor=GRIS)
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = Border(bottom=Side(style="thin", color="BFBFBF"))


def volcar(ws, df, columnas, formatos, fila_ini=2):
    encabezado(ws, fila_ini - 1, 1, columnas)
    for i, r in enumerate(df.itertuples(index=False), start=fila_ini):
        for j, (col, fmt) in enumerate(zip(df.columns, formatos), start=1):
            v = getattr(r, col.replace(" ", "_") if " " in col else col, None)
            if v is None:
                v = df.iloc[i - fila_ini][col]
            if hasattr(v, "date"):
                v = v.date()
            c = ws.cell(row=i, column=j, value=v)
            c.number_format = fmt
            c.font = Font(name=FUENTE, size=9)
    ws.freeze_panes = f"A{fila_ini}"
    return len(df) + fila_ini - 1


def construir_libro(diario, metas, mes, salida, split=None):
    split = SPLIT_MIXTO if split is None else split
    uni = universo(split)
    anio, mm = int(mes[:4]), int(mes[5:7])
    dias_mes = calendar.monthrange(anio, mm)[1]
    dias_data = diario["FECHA"].nunique()

    wb = Workbook()

    # -------- DIARIO --------
    wd = wb.active
    wd.title = "DIARIO"
    fin_d = volcar(wd, diario,
                   ["FECHA", "BLOQUE", "CANAL 2", "COSTO", "LEADS",
                    "Q NETO", "Q EMI", "Q TER"],
                   ["yyyy-mm-dd", "@", "@", "$#,##0", "#,##0",
                    "#,##0.0", "#,##0.0", "#,##0.0"])
    for col, w in zip("ABCDEFGH", (12, 10, 20, 16, 10, 11, 11, 11)):
        wd.column_dimensions[col].width = w

    # -------- METAS_DIA --------
    wm = wb.create_sheet("METAS_DIA")
    if "META_VENTAS" not in metas.columns:
        metas = metas.assign(META_VENTAS=0.0)
    metas = metas[["FECHA", "BLOQUE", "CANAL2", "META_COSTO", "META_LEADS",
                   "META_VENTAS"]]
    fin_m = volcar(wm, metas, ["FECHA", "BLOQUE", "CANAL 2", "META COSTO",
                               "META LEADS", "META VENTAS"],
                   ["yyyy-mm-dd", "@", "@", "$#,##0", "#,##0", "#,##0"])
    for col, w in zip("ABCDEF", (12, 10, 20, 16, 12, 13)):
        wm.column_dimensions[col].width = w

    # -------- PARAMETROS --------
    wp = wb.create_sheet("PARAMETROS")
    titulo(wp, "A1", "C1", "PARAMETROS DEL TRACKING", AZUL)
    datos = [("Mes reportado", mes, False), ("Area", AREA, False),
             ("Dias con informacion", dias_data, True),
             ("Dias del mes", dias_mes, True),
             ("Factor de proyeccion", None, False),
             ("% de MIXTO asignado a Movil y a Fijo", SPLIT_MIXTO, True),
             ("Ultima fecha con datos reales", diario["FECHA"].max().date(), False)]
    for i, (etq, val, edit) in enumerate(datos, start=3):
        wp.cell(row=i, column=1, value=etq).font = Font(name=FUENTE, bold=True, size=10)
        c = wp.cell(row=i, column=2, value=val)
        c.font = Font(name=FUENTE, size=10, color="0000FF" if edit else "000000")
        if edit:
            c.fill = PatternFill("solid", fgColor=AMARILLO)
    wp["B7"] = "=B6/B5"
    wp["B7"].number_format = "0.0000"
    wp["C5"] = "Bajalo si el ultimo dia viene incompleto"
    wp["C7"] = "Real acumulado x este factor = proyectado a cierre"
    for c in ("C5", "C7"):
        wp[c].font = Font(name=FUENTE, size=9, italic=True)
    for col, w in zip("ABC", (38, 16, 48)):
        wp.column_dimensions[col].width = w

    # -------- CURVA --------
    wc = wb.create_sheet("CURVA")
    escribir_curva(wc, mes, anio, mm, dias_mes, fin_d, fin_m, uni)

    # -------- CONTROL --------
    wk = wb.create_sheet("CONTROL")
    escribir_control(wk, diario, metas)

    # -------- RESUMEN --------
    wr = wb.create_sheet("RESUMEN")
    escribir_resumen(wr, diario, metas, fin_d, fin_m, mes, uni)

    # -------- MAQUETA --------
    wq = wb.create_sheet("MAQUETA")
    escribir_maqueta(wq, diario, metas, fin_d, fin_m, mes, dias_mes)

    orden = ["RESUMEN", "MAQUETA", "CURVA", "CONTROL", "PARAMETROS",
             "DIARIO", "METAS_DIA"]
    wb._sheets = ([wb[n] for n in orden if n in wb.sheetnames] +
                  [h for h in wb._sheets if h.title not in orden])
    wb.save(salida)
    return dias_data, dias_mes


def escribir_curva(ws, mes, anio, mm, dias_mes, fin_d, fin_m, uni):
    """Vista del dashboard: meta vs real por dia, embudo completo y acumulados."""
    etq = ("Movil + Fijo (Mixto ya viene repartido)" if "MIXTO" not in uni
           else "Movil + Fijo + Mixto")
    titulo(ws, "A1", "P1", f"AVANCE DIARIO {mes}  ·  {etq}", CELESTE)
    encabezado(ws, 2, 1, ["Fecha", "Meta Costo", "Costo", "% Costo",
                          "Meta Lead", "Lead", "CPL",
                          "Q Neto", "CPL Neto", "Q Emi", "CPE", "Q Ter", "CPA",
                          "Costo acum", "Meta acum", "% acum"])

    # Si Mixto va repartido se excluye del total; si no, entra tambien
    filtro = ("MOVIL","FIJO") if "MIXTO" not in uni else ("MOVIL","FIJO","MIXTO")
    def por_bloque(hoja, fin, col, celda_fecha):
        return "+".join(
            f'SUMIFS({hoja}!${col}$2:${col}${fin},'
            f'{hoja}!$A$2:$A${fin},{celda_fecha},'
            f'{hoja}!$B$2:$B${fin},"{b}")' for b in filtro)

    def suma_real(col, fila):
        return "(" + por_bloque("DIARIO", fin_d, col, f"$A{fila}") + ")"

    def suma_meta(col, fila):
        return "(" + por_bloque("METAS_DIA", fin_m, col, f"$A{fila}") + ")"

    for d in range(1, dias_mes + 1):
        r = 2 + d
        ws.cell(row=r, column=1, value=date(anio, mm, d)).number_format = "d mmm yyyy"
        ws.cell(row=r, column=2, value=f"={suma_meta('D', r)}")
        ws.cell(row=r, column=3, value=f"={suma_real('D', r)}")
        ws.cell(row=r, column=4, value=f"=IFERROR(C{r}/B{r},0)")
        ws.cell(row=r, column=5, value=f"={suma_meta('E', r)}")
        ws.cell(row=r, column=6, value=f"={suma_real('E', r)}")
        ws.cell(row=r, column=7, value=f"=IFERROR(C{r}/F{r},0)")

        # Embudo: volumen y su costo unitario, en pares
        for i, (col_src, _, _) in enumerate(EMBUDO):
            c_vol = 8 + i * 2
            L = get_column_letter(c_vol)
            ws.cell(row=r, column=c_vol, value=f"={suma_real(col_src, r)}")
            ws.cell(row=r, column=c_vol + 1, value=f"=IFERROR($C{r}/{L}{r},0)")

        ws.cell(row=r, column=14, value=f"=SUM($C$3:C{r})")
        ws.cell(row=r, column=15, value=f"=SUM($B$3:B{r})")
        ws.cell(row=r, column=16, value=f"=IFERROR(N{r}/O{r},0)")

        for col in (2, 3, 7, 9, 11, 13, 14, 15):
            ws.cell(row=r, column=col).number_format = "$#,##0"
        for col in (5, 6):
            ws.cell(row=r, column=col).number_format = "#,##0"
        for col in (8, 10, 12):
            ws.cell(row=r, column=col).number_format = "#,##0.0"
        for col in (4, 16):
            ws.cell(row=r, column=col).number_format = "0%"
        for col in range(1, 17):
            ws.cell(row=r, column=col).font = Font(name=FUENTE, size=9)

    # ---- TOTAL MES ----
    rt = 3 + dias_mes
    ws.cell(row=rt, column=1, value="TOTAL MES")
    for col in (2, 3, 5, 6, 8, 10, 12):
        L = get_column_letter(col)
        ws.cell(row=rt, column=col, value=f"=SUM({L}3:{L}{rt - 1})")
    ws.cell(row=rt, column=4, value=f"=IFERROR(C{rt}/B{rt},0)")
    ws.cell(row=rt, column=7, value=f"=IFERROR(C{rt}/F{rt},0)")
    for i in range(3):
        c_vol = 8 + i * 2
        L = get_column_letter(c_vol)
        ws.cell(row=rt, column=c_vol + 1, value=f"=IFERROR($C{rt}/{L}{rt},0)")

    # ---- PROYECTADO A CIERRE ----
    rp = rt + 1
    ws.cell(row=rp, column=1, value="PROYECTADO A CIERRE")
    ws.cell(row=rp, column=2, value=f"=B{rt}")
    ws.cell(row=rp, column=3, value=f"=C{rt}*PARAMETROS!$B$7")
    ws.cell(row=rp, column=4, value=f"=IFERROR(C{rp}/B{rp},0)")
    ws.cell(row=rp, column=5, value=f"=E{rt}")
    ws.cell(row=rp, column=6, value=f"=F{rt}*PARAMETROS!$B$7")
    ws.cell(row=rp, column=7, value=f"=IFERROR(C{rp}/F{rp},0)")
    for i in range(3):
        c_vol = 8 + i * 2
        L = get_column_letter(c_vol)
        ws.cell(row=rp, column=c_vol, value=f"={L}{rt}*PARAMETROS!$B$7")
        ws.cell(row=rp, column=c_vol + 1, value=f"=IFERROR($C{rp}/{L}{rp},0)")

    for r in (rt, rp):
        for col in range(1, 17):
            c = ws.cell(row=r, column=col)
            c.font = Font(name=FUENTE, bold=True, size=9)
            c.border = Border(top=Side(style="thin"))
        for col in (2, 3, 7, 9, 11, 13):
            ws.cell(row=r, column=col).number_format = "$#,##0"
        for col in (5, 6):
            ws.cell(row=r, column=col).number_format = "#,##0"
        for col in (8, 10, 12):
            ws.cell(row=r, column=col).number_format = "#,##0.0"
        ws.cell(row=r, column=4).number_format = "0%"

    ws.column_dimensions["A"].width = 14
    for col in "BCDEFGHIJKLMNOP":
        ws.column_dimensions[col].width = 14
    ws.freeze_panes = "B3"


def escribir_resumen(ws, diario, metas, fin_d, fin_m, mes, uni):
    """Hoja RESUMEN GENERAL: 5 tarjetas arriba y dos tablas abajo.
    Todo son formulas contra DIARIO y METAS_DIA, asi que se recalcula
    al cambiar los parametros."""
    MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
             "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    nombre_mes = f"{MESES[int(mes[5:7]) - 1]} {mes[:4]}"

    # ---- helpers de formula ----
    def real(col, bloques):
        col = COL.get(col, col)          # acepta "COSTO" o directamente "D"
        if len(bloques) == 1:
            return (f'SUMIFS(DIARIO!${col}$2:${col}${fin_d},'
                    f'DIARIO!$B$2:$B${fin_d},"{bloques[0]}")')
        partes = [f'SUMIFS(DIARIO!${col}$2:${col}${fin_d},'
                  f'DIARIO!$B$2:$B${fin_d},"{b}")' for b in bloques]
        return "(" + "+".join(partes) + ")"

    def meta(col, bloques):
        col = MCOL.get(col, col)         # acepta "COSTO" o directamente "D"
        partes = [f'SUMIFS(METAS_DIA!${col}$2:${col}${fin_m},'
                  f'METAS_DIA!$B$2:$B${fin_m},"{b}")' for b in bloques]
        return partes[0] if len(partes) == 1 else "(" + "+".join(partes) + ")"

    MF = list(uni)                  # el universo, sin doble conteo
    COL = {"COSTO": "D", "LEADS": "E", "Q_NETO": "F", "Q_EMI": "G", "Q_TER": "H"}
    MCOL = {"COSTO": "D", "LEADS": "E", "VENTAS": "F"}

    # ---- encabezado ----
    ws.merge_cells("A1:F1")
    ws["A1"] = "RESUMEN GENERAL"
    ws["A1"].font = Font(name=FUENTE, bold=True, size=20, color="1F3864")
    ws.merge_cells("A2:F2")
    ws["A2"] = f"Performance CL | Seguimiento · {nombre_mes}"
    ws["A2"].font = Font(name=FUENTE, size=10, color="7F7F7F")
    ws.merge_cells("J1:M1")
    ws["J1"] = '="Fecha de evaluación: "&TEXT(PARAMETROS!$B$9,"DD/MM/YYYY")'
    ws["J1"].font = Font(name=FUENTE, size=10, color="7F7F7F")
    ws["J1"].alignment = Alignment(horizontal="right")

    # Filas donde quedara la tabla de cumplimiento (se escribe mas abajo,
    # pero las formulas pueden apuntar a ella sin problema).
    fila_t = 4
    f0_tab = fila_t + 7
    fh_tab = f0_tab + 1
    r_mov, r_fij = fh_tab + 2, fh_tab + 1 + len(MF)   # primera y ultima del universo
    # Columnas del grupo Ventas y del grupo Presupuesto en esa tabla
    V_META, V_PROY = "K", "L"
    P_META, P_PROY = "B", "C"
    rango_vm = f"${V_META}${r_mov}:${V_META}${r_fij}"
    rango_vp = f"${V_PROY}${r_mov}:${V_PROY}${r_fij}"
    rango_pm = f"${P_META}${r_mov}:${P_META}${r_fij}"
    rango_pp = f"${P_PROY}${r_mov}:${P_PROY}${r_fij}"
    # Solo los bloques con meta de ventas cargada
    v_real = f'SUMIF({rango_vm},">0",{rango_vp})/PARAMETROS!$B$7'
    v_meta = f'SUMIF({rango_vm},">0",{rango_vm})'
    v_proy = f'SUMIF({rango_vm},">0",{rango_vp})'

    # ---- 5 tarjetas ----
    tarjetas = [
        ("INVERSIÓN", [
            ("Real corte", f"={real('COSTO', MF)}", "$#,##0"),
            ("Meta mes", f"={meta('COSTO', MF)}", "$#,##0"),
            ("Proy. cierre", f"={real('COSTO', MF)}*PARAMETROS!$B$7", "$#,##0"),
            ("Cumpl. proy.", None, "0%")]),
        ("LEADS GENERADOS", [
            ("Real corte", f"={real('LEADS', MF)}", "#,##0"),
            ("Meta mes", f"={meta('LEADS', MF)}", "#,##0"),
            ("Proy. cierre", f"={real('LEADS', MF)}*PARAMETROS!$B$7", "#,##0"),
            ("Cumpl. proy.", None, "0%")]),
        ("EMISIONES / NETOS", [
            ("Real corte", f"={real('Q_EMI', MF)}", "#,##0"),
            ("Meta mes", "0", "#,##0"),
            ("Proy. cierre", f"={real('Q_EMI', MF)}*PARAMETROS!$B$7", "#,##0"),
            ("Cumpl. proy.", None, "0%")]),
        ("VENTAS (solo con meta)", [
            ("Real corte", f"={v_real}", "#,##0"),
            ("Meta mes", f"={v_meta}", "#,##0"),
            ("Proy. cierre", f"={v_proy}", "#,##0"),
            ("Cumpl. proy.", None, "0%")]),
    ]

    for i, (nombre, lineas) in enumerate(tarjetas):
        c0 = 1 + i * 3
        li, lf = get_column_letter(c0), get_column_letter(c0 + 2)
        ws.merge_cells(f"{li}{fila_t}:{lf}{fila_t}")
        c = ws[f"{li}{fila_t}"]
        c.value = nombre
        c.font = Font(name=FUENTE, bold=True, size=11, color="1F3864")
        c.fill = PatternFill("solid", fgColor="DEEAF6")
        c.border = Border(left=Side(style="thick", color="1F3864"))
        for j, (etq, formula, fmt) in enumerate(lineas):
            r = fila_t + 1 + j
            ws.cell(row=r, column=c0, value=etq).font = Font(name=FUENTE, size=9)
            cel = ws.cell(row=r, column=c0 + 1)
            if formula is None:      # cumplimiento = proyectado / meta
                Lp = get_column_letter(c0 + 1)
                cel.value = f"=IFERROR({Lp}{r - 1}/{Lp}{r - 2},0)"
            else:
                cel.value = formula
            cel.number_format = fmt
            cel.font = Font(name=FUENTE, bold=(formula is None), size=9)
            for cx in (c0, c0 + 1, c0 + 2):
                ws.cell(row=r, column=cx).fill = PatternFill("solid", fgColor="F2F7FC")
            ws.cell(row=r, column=c0).border = Border(
                left=Side(style="thick", color="1F3864"))

    # Tarjeta de costos
    c0 = 13
    li, lf = get_column_letter(c0), get_column_letter(c0 + 2)
    ws.merge_cells(f"{li}{fila_t}:{lf}{fila_t}")
    c = ws[f"{li}{fila_t}"]
    c.value = "COSTOS"
    c.font = Font(name=FUENTE, bold=True, size=11, color="843C0C")
    c.fill = PatternFill("solid", fgColor="FFF2CC")
    c.border = Border(left=Side(style="thick", color="FFC000"))
    costos = [("CPL Meta", f"=IFERROR({meta('COSTO', MF)}/{meta('LEADS', MF)},0)"),
              ("CPL Proy.", f"=IFERROR({real('COSTO', MF)}/{real('LEADS', MF)},0)"),
              ("CPA Meta", f'=IFERROR(SUMIF({rango_vm},">0",{rango_pm})/{v_meta},0)'),
              ("CPA Proy.", f'=IFERROR(SUMIF({rango_vm},">0",{rango_pp})/{v_proy},0)')]
    for j, (etq, formula) in enumerate(costos):
        r = fila_t + 1 + j
        ws.cell(row=r, column=c0, value=etq).font = Font(name=FUENTE, size=9)
        cel = ws.cell(row=r, column=c0 + 1, value=formula)
        cel.number_format = "$#,##0"
        cel.font = Font(name=FUENTE, size=9)
        for cx in (c0, c0 + 1, c0 + 2):
            ws.cell(row=r, column=cx).fill = PatternFill("solid", fgColor="FFFBF0")
        ws.cell(row=r, column=c0).border = Border(
            left=Side(style="thick", color="FFC000"))

    # ---- tabla: cumplimiento por producto ----
    f0 = f0_tab
    ws[f"A{f0}"] = "Cumplimiento por producto"
    ws[f"A{f0}"].font = Font(name=FUENTE, bold=True, size=12, color="1F3864",
                             underline="single")

    grupos = [("Presupuesto", "COSTO", "COSTO", "$#,##0"),
              ("Leads", "LEADS", "LEADS", "#,##0"),
              ("Emitidos", "Q_EMI", None, "#,##0"),
              ("Ventas", "Q_TER", "VENTAS", "#,##0")]

    fh = fh_tab
    ws.cell(row=fh, column=1, value="Producto")
    for g, (nombre, _, _, _) in enumerate(grupos):
        c0 = 2 + g * 3
        li, lf = get_column_letter(c0), get_column_letter(c0 + 2)
        ws.merge_cells(f"{li}{fh}:{lf}{fh}")
        cel = ws[f"{li}{fh}"]
        cel.value = nombre
        cel.font = Font(name=FUENTE, bold=True, size=10, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor="1F3864")
        cel.alignment = Alignment(horizontal="center")
    ws.cell(row=fh, column=1).font = Font(name=FUENTE, bold=True, size=10,
                                          color="FFFFFF")
    ws.cell(row=fh, column=1).fill = PatternFill("solid", fgColor="1F3864")

    for g in range(len(grupos)):
        for k, sub in enumerate(("Meta", "Proy.", "% cumpl.")):
            cel = ws.cell(row=fh + 1, column=2 + g * 3 + k, value=sub)
            cel.font = Font(name=FUENTE, bold=True, size=8)
            cel.fill = PatternFill("solid", fgColor="DEEAF6")
            cel.alignment = Alignment(horizontal="center")

    etiquetas = {"MOVIL": "MÓVIL", "FIJO": "FIJO", "MIXTO": "MIXTO"}
    for i, bloque in enumerate(BLOQUES):
        r = fh + 2 + i
        ws.cell(row=r, column=1, value=etiquetas[bloque]).font = Font(
            name=FUENTE, bold=True, size=9)
        for g, (_, campo_real, campo_meta, fmt) in enumerate(grupos):
            c0 = 2 + g * 3
            Lm, Lp = get_column_letter(c0), get_column_letter(c0 + 1)
            ws.cell(row=r, column=c0,
                    value=(f"={meta(MCOL[campo_meta], [bloque])}"
                           if campo_meta else "0"))
            ws.cell(row=r, column=c0 + 1,
                    value=f"={real(COL[campo_real], [bloque])}*PARAMETROS!$B$7")
            ws.cell(row=r, column=c0 + 2,
                    value=f"=IFERROR({Lp}{r}/{Lm}{r},0)")
            ws.cell(row=r, column=c0).number_format = fmt
            ws.cell(row=r, column=c0 + 1).number_format = fmt
            ws.cell(row=r, column=c0 + 2).number_format = "0%"
        for cx in range(2, 14):
            ws.cell(row=r, column=cx).font = Font(name=FUENTE, size=9)

    # Total Movil + Fijo
    rt = fh + 2 + len(BLOQUES)
    ws.cell(row=rt, column=1, value="TOTAL").font = Font(name=FUENTE, bold=True, size=9)
    for g, (_, campo_real, campo_meta, fmt) in enumerate(grupos):
        c0 = 2 + g * 3
        Lm, Lp = get_column_letter(c0), get_column_letter(c0 + 1)
        if campo_meta == "VENTAS":
            # Solo los bloques con meta cargada, para no comparar las ventas
            # de todos contra la meta de uno solo.
            ws.cell(row=rt, column=c0, value=f"={v_meta}")
            ws.cell(row=rt, column=c0 + 1, value=f"={v_proy}")
        else:
            ws.cell(row=rt, column=c0,
                    value=f"={meta(MCOL[campo_meta], MF)}" if campo_meta else "0")
            ws.cell(row=rt, column=c0 + 1,
                    value=f"={real(COL[campo_real], MF)}*PARAMETROS!$B$7")
        ws.cell(row=rt, column=c0 + 2, value=f"=IFERROR({Lp}{rt}/{Lm}{rt},0)")
        ws.cell(row=rt, column=c0).number_format = fmt
        ws.cell(row=rt, column=c0 + 1).number_format = fmt
        ws.cell(row=rt, column=c0 + 2).number_format = "0%"
    for cx in range(1, 14):
        c = ws.cell(row=rt, column=cx)
        c.font = Font(name=FUENTE, bold=True, size=9)
        c.border = Border(top=Side(style="thin"))
    ws.cell(row=rt + 1, column=1,
            value=("TOTAL = Móvil + Fijo. Mixto no se suma: ya viene repartido "
                   "dentro de los dos." if "MIXTO" not in MF else
                   "TOTAL = Móvil + Fijo + Mixto. Los bloques están puros, "
                   "Mixto no se reparte.")).font = Font(name=FUENTE, size=8,
                                                    italic=True, color="7F7F7F")

    # ---- tabla: costos unitarios ----
    f1 = rt + 3
    ws[f"A{f1}"] = "Costos"
    ws[f"A{f1}"].font = Font(name=FUENTE, bold=True, size=12, color="C00000",
                             underline="single")

    unit = [("CPL", "LEADS", "LEADS"), ("CPE", "Q_EMI", None),
            ("CPA", "Q_TER", "VENTAS")]
    fh1 = f1 + 1
    cel = ws.cell(row=fh1, column=1, value="Producto")
    cel.font = Font(name=FUENTE, bold=True, size=10, color="FFFFFF")
    cel.fill = PatternFill("solid", fgColor="1F3864")
    for g, (nombre, _, _) in enumerate(unit):
        c0 = 2 + g * 3
        li, lf = get_column_letter(c0), get_column_letter(c0 + 2)
        ws.merge_cells(f"{li}{fh1}:{lf}{fh1}")
        cel = ws[f"{li}{fh1}"]
        cel.value = nombre
        cel.font = Font(name=FUENTE, bold=True, size=10, color="FFFFFF")
        cel.fill = PatternFill("solid", fgColor="1F3864")
        cel.alignment = Alignment(horizontal="center")
        for k, sub in enumerate(("Meta", "Proy.", "Var %")):
            c2 = ws.cell(row=fh1 + 1, column=c0 + k, value=sub)
            c2.font = Font(name=FUENTE, bold=True, size=8)
            c2.fill = PatternFill("solid", fgColor="FCE4E4")
            c2.alignment = Alignment(horizontal="center")

    for i, bloque in enumerate(list(BLOQUES) + ["TOTAL"]):
        r = fh1 + 2 + i
        bl = MF if bloque == "TOTAL" else [bloque]
        ws.cell(row=r, column=1, value=etiquetas.get(bloque, bloque)).font = Font(
            name=FUENTE, bold=True, size=9)
        for g, (_, campo_real, campo_meta) in enumerate(unit):
            c0 = 2 + g * 3
            Lm, Lp = get_column_letter(c0), get_column_letter(c0 + 1)
            if campo_meta == "VENTAS" and bloque == "TOTAL":
                # Solo los bloques con meta de ventas, si no se compara el
                # costo de todos contra las ventas meta de uno solo.
                ws.cell(row=r, column=c0,
                        value=f'=IFERROR(SUMIF({rango_vm},">0",{rango_pm})/{v_meta},0)')
                ws.cell(row=r, column=c0 + 1,
                        value=f'=IFERROR(SUMIF({rango_vm},">0",{rango_pp})/{v_proy},0)')
            else:
                ws.cell(row=r, column=c0,
                        value=(f"=IFERROR({meta('D', bl)}/"
                               f"{meta(MCOL[campo_meta], bl)},0)"
                               if campo_meta else "0"))
                ws.cell(row=r, column=c0 + 1,
                        value=f"=IFERROR({real('D', bl)}/"
                              f"{real(COL[campo_real], bl)},0)")
            ws.cell(row=r, column=c0 + 2,
                    value=f"=IFERROR({Lp}{r}/{Lm}{r}-1,0)")
            ws.cell(row=r, column=c0).number_format = "$#,##0"
            ws.cell(row=r, column=c0 + 1).number_format = "$#,##0"
            ws.cell(row=r, column=c0 + 2).number_format = "+0%;-0%;0%"
            for k in range(3):
                ws.cell(row=r, column=c0 + k).font = Font(
                    name=FUENTE, bold=(bloque == "TOTAL"), size=9)
        if bloque == "TOTAL":
            for cx in range(1, 11):
                ws.cell(row=r, column=cx).border = Border(top=Side(style="thin"))

    ws.cell(row=fh1 + 2 + len(BLOQUES) + 2, column=1,
            value="Emisiones y CPE no tienen meta en el presupuesto: quedan en 0 a "
                  "propósito. La meta de ventas solo viene cargada para Fijo, así que "
                  "la tarjeta de VENTAS y el CPA del TOTAL se calculan únicamente "
                  "sobre los bloques que sí tienen meta.").font = Font(
        name=FUENTE, size=8, italic=True, color="C00000")

    ws.column_dimensions["A"].width = 18
    for cx in range(2, 17):
        ws.column_dimensions[get_column_letter(cx)].width = 13
    ws.sheet_view.showGridLines = False


def escribir_control(ws, diario, metas):
    """Avisa que combos tienen real sin meta o meta sin real."""
    titulo(ws, "A1", "D1", "CONTROL DE CRUCE REAL vs PPTO", AZUL)
    ws["A2"] = ("Revisa esta hoja cada mes. Si aparece algo, es que cambio el nombre de "
                "un canal o que falta cargarlo en el PPTo.")
    ws["A2"].font = Font(name=FUENTE, size=9, italic=True)

    r_real = set(zip(diario["BLOQUE"], diario["CANAL2"]))
    r_meta = set(zip(metas["BLOQUE"], metas["CANAL2"]))
    encabezado(ws, 4, 1, ["BLOQUE", "CANAL 2", "SITUACION", "ACCION SUGERIDA"])

    fila = 5
    for bloque, canal in sorted(r_real - r_meta):
        for col, val in enumerate(
                [bloque, canal, "Tiene inversion real pero no tiene meta en el PPTo",
                 "Cargar la meta o confirmar que no lleva"], start=1):
            c = ws.cell(row=fila, column=col, value=val)
            c.font = Font(name=FUENTE, size=9)
            c.fill = PatternFill("solid", fgColor=ROJO)
        fila += 1
    for bloque, canal in sorted(r_meta - r_real):
        for col, val in enumerate(
                [bloque, canal, "Tiene meta pero todavia no registra inversion",
                 "Verificar si la campana ya arranco"], start=1):
            c = ws.cell(row=fila, column=col, value=val)
            c.font = Font(name=FUENTE, size=9)
            c.fill = PatternFill("solid", fgColor=AMARILLO)
        fila += 1
    if fila == 5:
        c = ws.cell(row=5, column=1, value="Todo cruza. Ningun canal quedo sin meta.")
        c.font = Font(name=FUENTE, size=9)
        c.fill = PatternFill("solid", fgColor=VERDE)

    for col, w in zip("ABCD", (12, 22, 48, 42)):
        ws.column_dimensions[col].width = w


def escribir_maqueta(ws, diario, metas, fin_d, fin_m, mes, dias_mes):
    titulo(ws, "A1", "R1", f"TRACKING Y PROYECCION A CIERRE · {mes}", AZUL)
    ws["A2"] = ("Proyeccion = Real acumulado / dias con informacion x dias del mes "
                "(run-rate simple). 'Meta a hoy' es la meta diaria acumulada hasta "
                "la fecha de corte, para ver el pacing. Parametros en PARAMETROS.")
    ws["A2"].font = Font(name=FUENTE, size=9, italic=True)

    cab = {"COSTOS": ["Canal 2", "Meta Costo mes", "Meta a hoy", "Real Costo",
                      "% a hoy", "Costo proyectado"],
           "LEADS": ["Canal 2", "Meta Leads mes", "Meta a hoy", "Real Leads",
                     "% a hoy", "Leads Proyectados"],
           "CPL": ["Canal 2", "CPL Meta", "CPL", "% Cumplimiento"]}
    ANCHO = {"COSTOS": 6, "LEADS": 6, "CPL": 4}
    INICIO = {"COSTOS": 1, "LEADS": 8, "CPL": 15}     # columnas A, H, O

    fila = 4
    for bloque in BLOQUES:
        canales = sorted(set(diario[diario["BLOQUE"] == bloque]["CANAL2"]) |
                         set(metas[metas["BLOQUE"] == bloque]["CANAL2"]))
        if not canales:
            continue

        for metrica in ("COSTOS", "LEADS", "CPL"):
            c0, ancho = INICIO[metrica], ANCHO[metrica]
            li, lf = get_column_letter(c0), get_column_letter(c0 + ancho - 1)
            titulo(ws, f"{li}{fila}", f"{lf}{fila}", f"{metrica} {bloque}",
                   AZUL if bloque == "MOVIL" else CELESTE)
            encabezado(ws, fila + 1, c0, cab[metrica])

            # Letras de las columnas que necesita el bloque CPL
            L_META_C, L_REAL_C = get_column_letter(2), get_column_letter(4)
            L_META_L, L_REAL_L = get_column_letter(9), get_column_letter(11)
            L_CPLM, L_CPL = get_column_letter(16), get_column_letter(17)

            for k, canal in enumerate(canales):
                r = fila + 2 + k
                cc = get_column_letter(c0)
                ws.cell(row=r, column=c0, value=canal)

                if metrica != "CPL":
                    col = "D" if metrica == "COSTOS" else "E"
                    real = (f'SUMIFS(DIARIO!${col}$2:${col}${fin_d},'
                            f'DIARIO!$B$2:$B${fin_d},"{bloque}",'
                            f'DIARIO!$C$2:$C${fin_d},${cc}{r})')
                    meta_mes = (f'SUMIFS(METAS_DIA!${col}$2:${col}${fin_m},'
                                f'METAS_DIA!$B$2:$B${fin_m},"{bloque}",'
                                f'METAS_DIA!$C$2:$C${fin_m},${cc}{r})')
                    # Meta acumulada solo hasta la fecha de corte
                    meta_hoy = (f'SUMIFS(METAS_DIA!${col}$2:${col}${fin_m},'
                                f'METAS_DIA!$B$2:$B${fin_m},"{bloque}",'
                                f'METAS_DIA!$C$2:$C${fin_m},${cc}{r},'
                                f'METAS_DIA!$A$2:$A${fin_m},"<="&PARAMETROS!$B$9)')
                    Lm, Lh, Lr = (get_column_letter(c0 + 1), get_column_letter(c0 + 2),
                                  get_column_letter(c0 + 3))
                    ws.cell(row=r, column=c0 + 1, value=f"={meta_mes}")
                    ws.cell(row=r, column=c0 + 2, value=f"={meta_hoy}")
                    ws.cell(row=r, column=c0 + 3, value=f"={real}")
                    ws.cell(row=r, column=c0 + 4,
                            value=f"=IFERROR({Lr}{r}/{Lh}{r},0)")
                    ws.cell(row=r, column=c0 + 5,
                            value=f"={Lr}{r}*PARAMETROS!$B$7")
                    fmt = "$#,##0" if metrica == "COSTOS" else "#,##0"
                    for cx in (c0 + 1, c0 + 2, c0 + 3, c0 + 5):
                        ws.cell(row=r, column=cx).number_format = fmt
                    ws.cell(row=r, column=c0 + 4).number_format = "0%"
                else:
                    ws.cell(row=r, column=c0 + 1,
                            value=f"=IFERROR({L_META_C}{r}/{L_META_L}{r},0)")
                    ws.cell(row=r, column=c0 + 2,
                            value=f"=IFERROR({L_REAL_C}{r}/{L_REAL_L}{r},0)")
                    ws.cell(row=r, column=c0 + 3,
                            value=f"=IFERROR({L_CPL}{r}/{L_CPLM}{r},0)")
                    ws.cell(row=r, column=c0 + 1).number_format = "$#,##0"
                    ws.cell(row=r, column=c0 + 2).number_format = "$#,##0"
                    ws.cell(row=r, column=c0 + 3).number_format = "0%"

                for cx in range(c0, c0 + ancho):
                    ws.cell(row=r, column=cx).font = Font(name=FUENTE, size=9)

            rt = fila + 2 + len(canales)
            pp, uu = fila + 2, fila + 1 + len(canales)
            ws.cell(row=rt, column=c0, value="Total")
            if metrica != "CPL":
                fmt = "$#,##0" if metrica == "COSTOS" else "#,##0"
                for cx in (c0 + 1, c0 + 2, c0 + 3, c0 + 5):
                    L = get_column_letter(cx)
                    ws.cell(row=rt, column=cx, value=f"=SUM({L}{pp}:{L}{uu})")
                    ws.cell(row=rt, column=cx).number_format = fmt
                Lh, Lr = get_column_letter(c0 + 2), get_column_letter(c0 + 3)
                ws.cell(row=rt, column=c0 + 4,
                        value=f"=IFERROR({Lr}{rt}/{Lh}{rt},0)")
                ws.cell(row=rt, column=c0 + 4).number_format = "0%"
            else:
                ws.cell(row=rt, column=c0 + 1,
                        value=f"=IFERROR({L_META_C}{rt}/{L_META_L}{rt},0)")
                ws.cell(row=rt, column=c0 + 2,
                        value=f"=IFERROR({L_REAL_C}{rt}/{L_REAL_L}{rt},0)")
                ws.cell(row=rt, column=c0 + 3,
                        value=f"=IFERROR({L_CPL}{rt}/{L_CPLM}{rt},0)")
                ws.cell(row=rt, column=c0 + 1).number_format = "$#,##0"
                ws.cell(row=rt, column=c0 + 2).number_format = "$#,##0"
                ws.cell(row=rt, column=c0 + 3).number_format = "0%"
            for cx in range(c0, c0 + ancho):
                c = ws.cell(row=rt, column=cx)
                c.font = Font(name=FUENTE, bold=True, size=9)
                c.border = Border(top=Side(style="thin"))

        # ---- Tabla de embudo del bloque (sin metas: solo real y proyectado) ----
        f_emb = fila + len(canales) + 4
        titulo(ws, f"A{f_emb}", f"J{f_emb}",
               f"EMBUDO Y COSTOS UNITARIOS {bloque}  ·  sin meta asociada", GRIS_T)
        ws[f"A{f_emb}"].font = Font(name=FUENTE, bold=True, size=11, color="000000")
        encabezado(ws, f_emb + 1, 1,
                   ["Canal 2", "Q Neto", "Q Neto proy", "CPL Neto",
                    "Q Emi", "Q Emi proy", "CPE",
                    "Q Ter", "Q Ter proy", "CPA"])

        for k, canal in enumerate(canales):
            r = f_emb + 2 + k
            ws.cell(row=r, column=1, value=canal)
            costo = (f'SUMIFS(DIARIO!$D$2:$D${fin_d},'
                     f'DIARIO!$B$2:$B${fin_d},"{bloque}",'
                     f'DIARIO!$C$2:$C${fin_d},$A{r})')
            for i, (col_src, _, _) in enumerate(EMBUDO):
                c_vol = 2 + i * 3
                L = get_column_letter(c_vol)
                vol = (f'SUMIFS(DIARIO!${col_src}$2:${col_src}${fin_d},'
                       f'DIARIO!$B$2:$B${fin_d},"{bloque}",'
                       f'DIARIO!$C$2:$C${fin_d},$A{r})')
                ws.cell(row=r, column=c_vol, value=f"={vol}")
                ws.cell(row=r, column=c_vol + 1,
                        value=f"={L}{r}*PARAMETROS!$B$7")
                ws.cell(row=r, column=c_vol + 2,
                        value=f"=IFERROR({costo}/{L}{r},0)")
                ws.cell(row=r, column=c_vol).number_format = "#,##0.0"
                ws.cell(row=r, column=c_vol + 1).number_format = "#,##0.0"
                ws.cell(row=r, column=c_vol + 2).number_format = "$#,##0"
            for cx in range(1, 11):
                ws.cell(row=r, column=cx).font = Font(name=FUENTE, size=9)

        rt = f_emb + 2 + len(canales)
        p0, u0 = f_emb + 2, f_emb + 1 + len(canales)
        ws.cell(row=rt, column=1, value="Total")
        costo_t = (f'SUMIFS(DIARIO!$D$2:$D${fin_d},'
                   f'DIARIO!$B$2:$B${fin_d},"{bloque}")')
        for i in range(3):
            c_vol = 2 + i * 3
            L, L2 = get_column_letter(c_vol), get_column_letter(c_vol + 1)
            ws.cell(row=rt, column=c_vol, value=f"=SUM({L}{p0}:{L}{u0})")
            ws.cell(row=rt, column=c_vol + 1, value=f"=SUM({L2}{p0}:{L2}{u0})")
            ws.cell(row=rt, column=c_vol + 2, value=f"=IFERROR({costo_t}/{L}{rt},0)")
            ws.cell(row=rt, column=c_vol).number_format = "#,##0.0"
            ws.cell(row=rt, column=c_vol + 1).number_format = "#,##0.0"
            ws.cell(row=rt, column=c_vol + 2).number_format = "$#,##0"
        for cx in range(1, 11):
            c = ws.cell(row=rt, column=cx)
            c.font = Font(name=FUENTE, bold=True, size=9)
            c.border = Border(top=Side(style="thin"))

        fila += 2 * len(canales) + 8

    # Resumen del mes (los 3 medidores del dashboard)
    titulo(ws, f"A{fila}", f"D{fila}", "RESUMEN DEL MES (Movil + Fijo)", AZUL)
    encabezado(ws, fila + 1, 1, ["Indicador", "Meta mes", "Real / Proyectado", "%"])
    resumen = [("Costo (real a hoy)", "=CURVA!B{t}", "=CURVA!C{t}", "$"),
               ("Costo (proyectado a cierre)", "=CURVA!B{p}", "=CURVA!C{p}", "$"),
               ("Leads (real a hoy)", "=CURVA!E{t}", "=CURVA!F{t}", "#"),
               ("Leads (proyectado a cierre)", "=CURVA!E{p}", "=CURVA!F{p}", "#"),
               ("CPL (real a hoy)", "=IFERROR(CURVA!B{t}/CURVA!E{t},0)", "=CURVA!G{t}", "$"),
               ("CPL (proyectado a cierre)", "=IFERROR(CURVA!B{p}/CURVA!E{p},0)",
                "=CURVA!G{p}", "$")]
    t_row, p_row = 3 + dias_mes, 4 + dias_mes
    for i, (etq, fm, fr, tipo) in enumerate(resumen):
        r = fila + 2 + i
        ws.cell(row=r, column=1, value=etq)
        ws.cell(row=r, column=2, value=fm.format(t=t_row, p=p_row))
        ws.cell(row=r, column=3, value=fr.format(t=t_row, p=p_row))
        ws.cell(row=r, column=4, value=f"=IFERROR(C{r}/B{r},0)")
        for cx in (2, 3):
            ws.cell(row=r, column=cx).number_format = "$#,##0" if tipo == "$" else "#,##0"
        ws.cell(row=r, column=4).number_format = "0%"
        for cx in range(1, 5):
            ws.cell(row=r, column=cx).font = Font(name=FUENTE, size=9)

    for metrica, c0 in (("COSTOS", 1), ("LEADS", 8), ("CPL", 15)):
        ancho = 6 if metrica != "CPL" else 4
        ws.column_dimensions[get_column_letter(c0)].width = 22
        for k in range(1, ancho):
            ws.column_dimensions[get_column_letter(c0 + k)].width = 16
        if metrica != "CPL":
            ws.column_dimensions[get_column_letter(c0 + ancho)].width = 2


# ----------------------------------------------------------------
if __name__ == "__main__":
    args = sys.argv[1:]

    # Si pasas un mes AAAA-MM lo usa; si no, usa el mes de hoy.
    mes = next((a for a in args if len(a) == 7 and a[4] == "-"),
               date.today().strftime("%Y-%m"))
    rutas = [a for a in args if a != mes]

    if len(rutas) >= 2:
        ruta_maqueta, ruta_ppto = rutas[0], rutas[1]
    else:
        print(f"0/3  Buscando archivos en:\n     {CARPETA_INPUTS}")
        ruta_maqueta = buscar(CARPETA_INPUTS, PATRON_MAQUETA, "maqueta")
        ruta_ppto = buscar(CARPETA_INPUTS, PATRON_PPTO, "PPTo")
        print(f"     Maqueta : {os.path.basename(ruta_maqueta)}")
        print(f"     PPTo    : {os.path.basename(ruta_ppto)}")

    os.makedirs(CARPETA_OUTPUTS, exist_ok=True)
    salida = os.path.join(CARPETA_OUTPUTS, f"Tracking_CPE_{mes}.xlsx")

    print("1/3  Leyendo la base real dentro de la maqueta...")
    base = extraer(ruta_maqueta)
    print(f"     {len(base):,} filas, hasta {base['FECHA'].max().date()}")

    print("2/3  Leyendo las metas diarias del PPTo...")
    metas = leer_metas(ruta_ppto, mes)
    solo_mf = metas[metas["BLOQUE"] != "MIXTO"]
    print(f"     {len(metas):,} filas de meta - {metas['FECHA'].nunique()} dias - "
          f"meta mes Movil+Fijo ${solo_mf['META_COSTO'].sum():,.0f}")

    print("3/3  Armando el tracking...")
    metas = repartir_metas(metas, SPLIT_MIXTO)
    diario = construir_diario(base, mes)
    dd, dm = construir_libro(diario, metas, mes, salida, SPLIT_MIXTO)
    print(f"     {dd} dias con datos de {dm}  ->  factor {dm / dd:.4f}")
    print(f"\nListo: {salida}")
    input("\nPresiona ENTER para cerrar...")
