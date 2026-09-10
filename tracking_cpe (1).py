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
def construir_diario(df, mes, area=None, split=None, solo_con_costo=None):
    area = AREA if area is None else area
    split = SPLIT_MIXTO if split is None else split
    solo = SOLO_CON_COSTO if solo_con_costo is None else solo_con_costo
    df = df[(df["MES"] == mes) & (df["AREA"] == area)].copy()
    if df.empty:
        return df
    df["CANAL2"] = df["CANAL"].map(lambda c: NOMBRE_CANAL.get(c, str(c).title()))

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
        .groupby(["FECHA", "BLOQUE", "CANAL2"], as_index=False)[METRICAS].sum()
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


def construir_libro(diario, metas, mes, salida):
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
    fin_m = volcar(wm, metas, ["FECHA", "BLOQUE", "CANAL 2", "META COSTO", "META LEADS"],
                   ["yyyy-mm-dd", "@", "@", "$#,##0", "#,##0"])
    for col, w in zip("ABCDE", (12, 10, 20, 16, 12)):
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
    escribir_curva(wc, mes, anio, mm, dias_mes, fin_d, fin_m)

    # -------- CONTROL --------
    wk = wb.create_sheet("CONTROL")
    escribir_control(wk, diario, metas)

    # -------- MAQUETA --------
    wq = wb.create_sheet("MAQUETA")
    escribir_maqueta(wq, diario, metas, fin_d, fin_m, mes, dias_mes)

    wb.move_sheet("MAQUETA", offset=-5)
    wb.move_sheet("CURVA", offset=-3)
    wb.save(salida)
    return dias_data, dias_mes


def escribir_curva(ws, mes, anio, mm, dias_mes, fin_d, fin_m):
    """Vista del dashboard: meta vs real por dia, embudo completo y acumulados."""
    titulo(ws, "A1", "P1",
           f"AVANCE DIARIO {mes}  ·  Movil + Fijo (Mixto ya viene repartido)", CELESTE)
    encabezado(ws, 2, 1, ["Fecha", "Meta Costo", "Costo", "% Costo",
                          "Meta Lead", "Lead", "CPL",
                          "Q Neto", "CPL Neto", "Q Emi", "CPE", "Q Ter", "CPA",
                          "Costo acum", "Meta acum", "% acum"])

    no_mix_d = f'DIARIO!$B$2:$B${fin_d},"<>MIXTO"'
    no_mix_m = f'METAS_DIA!$B$2:$B${fin_m},"<>MIXTO"'

    def suma_real(col, fila):
        return (f'SUMIFS(DIARIO!${col}$2:${col}${fin_d},'
                f'DIARIO!$A$2:$A${fin_d},$A{fila},{no_mix_d})')

    for d in range(1, dias_mes + 1):
        r = 2 + d
        ws.cell(row=r, column=1, value=date(anio, mm, d)).number_format = "d mmm yyyy"
        ws.cell(row=r, column=2, value=f'=SUMIFS(METAS_DIA!$D$2:$D${fin_m},'
                                       f'METAS_DIA!$A$2:$A${fin_m},$A{r},{no_mix_m})')
        ws.cell(row=r, column=3, value=f"={suma_real('D', r)}")
        ws.cell(row=r, column=4, value=f"=IFERROR(C{r}/B{r},0)")
        ws.cell(row=r, column=5, value=f'=SUMIFS(METAS_DIA!$E$2:$E${fin_m},'
                                       f'METAS_DIA!$A$2:$A${fin_m},$A{r},{no_mix_m})')
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
    titulo(ws, "A1", "N1", f"TRACKING Y PROYECCION A CIERRE · {mes}", AZUL)
    ws["A2"] = ("Proyeccion = Real acumulado / dias con informacion x dias del mes "
                "(run-rate simple). Metas tomadas del PPTo. Parametros en PARAMETROS.")
    ws["A2"].font = Font(name=FUENTE, size=9, italic=True)

    cab = {"COSTOS": ["Canal 2", "Meta Costo", "Real Costo", "Costo proyectado"],
           "LEADS": ["Canal 2", "Meta Leads", "Real Leads", "Leads Proyectados"],
           "CPL": ["Canal 2", "CPL Meta", "CPL", "% Cumplimiento"]}

    fila = 4
    for bloque in BLOQUES:
        canales = sorted(set(diario[diario["BLOQUE"] == bloque]["CANAL2"]) |
                         set(metas[metas["BLOQUE"] == bloque]["CANAL2"]))
        if not canales:
            continue

        for j, metrica in enumerate(["COSTOS", "LEADS", "CPL"]):
            c0 = 1 + j * 5
            li, lf = get_column_letter(c0), get_column_letter(c0 + 3)
            titulo(ws, f"{li}{fila}", f"{lf}{fila}", f"{metrica} {bloque}",
                   AZUL if bloque == "MOVIL" else CELESTE)
            encabezado(ws, fila + 1, c0, cab[metrica])

            for k, canal in enumerate(canales):
                r = fila + 2 + k
                cc = get_column_letter(c0)
                ws.cell(row=r, column=c0, value=canal)

                col = "D" if metrica != "LEADS" else "E"
                real = (f'SUMIFS(DIARIO!${col}$2:${col}${fin_d},'
                        f'DIARIO!$B$2:$B${fin_d},"{bloque}",'
                        f'DIARIO!$C$2:$C${fin_d},${cc}{r})')
                meta = (f'SUMIFS(METAS_DIA!${col}$2:${col}${fin_m},'
                        f'METAS_DIA!$B$2:$B${fin_m},"{bloque}",'
                        f'METAS_DIA!$C$2:$C${fin_m},${cc}{r})')

                if metrica != "CPL":
                    ws.cell(row=r, column=c0 + 1, value=f"={meta}")
                    ws.cell(row=r, column=c0 + 2, value=f"={real}")
                    ws.cell(row=r, column=c0 + 3,
                            value=f"={get_column_letter(c0 + 2)}{r}*PARAMETROS!$B$7")
                    fmt = "$#,##0" if metrica == "COSTOS" else "#,##0"
                    for cx in (c0 + 1, c0 + 2, c0 + 3):
                        ws.cell(row=r, column=cx).number_format = fmt
                else:
                    ws.cell(row=r, column=c0 + 1, value=f"=IFERROR(B{r}/G{r},0)")
                    ws.cell(row=r, column=c0 + 2, value=f"=IFERROR(C{r}/H{r},0)")
                    ws.cell(row=r, column=c0 + 3, value=f"=IFERROR(M{r}/L{r},0)")
                    ws.cell(row=r, column=c0 + 1).number_format = "$#,##0"
                    ws.cell(row=r, column=c0 + 2).number_format = "$#,##0"
                    ws.cell(row=r, column=c0 + 3).number_format = "0%"

                for cx in range(c0, c0 + 4):
                    ws.cell(row=r, column=cx).font = Font(name=FUENTE, size=9)

            rt = fila + 2 + len(canales)
            p, u = fila + 2, fila + 1 + len(canales)
            ws.cell(row=rt, column=c0, value="Total")
            if metrica != "CPL":
                for cx in (c0 + 1, c0 + 2, c0 + 3):
                    L = get_column_letter(cx)
                    ws.cell(row=rt, column=cx, value=f"=SUM({L}{p}:{L}{u})")
                    ws.cell(row=rt, column=cx).number_format = (
                        "$#,##0" if metrica == "COSTOS" else "#,##0")
            else:
                ws.cell(row=rt, column=c0 + 1, value=f"=IFERROR(B{rt}/G{rt},0)")
                ws.cell(row=rt, column=c0 + 2, value=f"=IFERROR(C{rt}/H{rt},0)")
                ws.cell(row=rt, column=c0 + 3, value=f"=IFERROR(M{rt}/L{rt},0)")
                ws.cell(row=rt, column=c0 + 1).number_format = "$#,##0"
                ws.cell(row=rt, column=c0 + 2).number_format = "$#,##0"
                ws.cell(row=rt, column=c0 + 3).number_format = "0%"
            for cx in range(c0, c0 + 4):
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

    for j in range(3):
        c0 = 1 + j * 5
        ws.column_dimensions[get_column_letter(c0)].width = 24
        for k in (1, 2, 3):
            ws.column_dimensions[get_column_letter(c0 + k)].width = 17
        ws.column_dimensions[get_column_letter(c0 + 4)].width = 2


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
    diario = construir_diario(base, mes)
    dd, dm = construir_libro(diario, metas, mes, salida)
    print(f"     {dd} dias con datos de {dm}  ->  factor {dm / dd:.4f}")
    print(f"\nListo: {salida}")
    input("\nPresiona ENTER para cerrar...")
