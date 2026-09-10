"""
extraer_base.py
-----------------------------------------------------------
Saca la tabla PLANA (fila por fila, con FECHA y CANAL) que vive
dentro del .xlsx de la maqueta.

No hay que borrar filas ni pegar fechas a mano: la consulta ya trae
FECHA, SEMANA, MES, CANAL, TIPO, etc. Lo que ves en Hoja1/Hoja2 es
solo el resumen de una tabla dinamica; el detalle esta en la cache
interna del archivo.

Uso:
    python extraer_base.py archivo.xlsx base_plana.csv
"""

import re
import sys
import zipfile
from xml.etree.ElementTree import iterparse

import pandas as pd

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# Campos que realmente necesitamos para el tracking.
CAMPOS_UTILES = [
    "MES", "SEMANA", "FECHA", "AREA", "TIPO", "CANAL", "CATEGORIA",
    "CAMPANA", "PRODUCTO", "FAMILIA",
    "IMPRESIONES", "CLICK", "LEADS", "COSTO",
    "Q_BRUTO", "Q_NETO", "Q_EMI", "Q_TER",
]


def leer_definicion(zf, ruta_def):
    """Devuelve (nombres_de_campo_en_orden, items_compartidos_por_campo)."""
    xml = zf.read(ruta_def).decode("utf-8")
    bloques = re.split(r"(?=<cacheField name=)", xml)[1:]

    nombres, shared = [], []
    for bloque in bloques:
        nombre = re.match(r'<cacheField name="([^"]+)"', bloque).group(1)
        cabecera = bloque[: bloque.find(">") + 1]

        # Los campos calculados (formula=) no se guardan en los registros.
        if 'databaseField="0"' in cabecera or "formula=" in cabecera:
            continue

        items = []
        for m in re.finditer(r"<(s|n|d|b|m|e)(\s[^>]*)?/>", bloque):
            tipo, attrs = m.group(1), m.group(2) or ""
            v = re.search(r'v="([^"]*)"', attrs)
            if tipo == "m":
                items.append(None)
            elif tipo == "s":
                items.append(_unescape(v.group(1)) if v else "")
            elif tipo == "n":
                items.append(float(v.group(1)) if v else None)
            elif tipo == "d":
                items.append(v.group(1) if v else None)
            else:
                items.append(v.group(1) if v else None)

        nombres.append(nombre)
        shared.append(items)

    return nombres, shared


def _unescape(s):
    return (s.replace("&amp;", "&").replace("&lt;", "<")
             .replace("&gt;", ">").replace("&quot;", '"')
             .replace("&apos;", "'"))


def leer_registros(zf, ruta_rec, nombres, shared, columnas):
    """Recorre la cache en streaming y arma la lista de filas."""
    quiero = [nombres.index(c) for c in columnas]
    filas = []

    with zf.open(ruta_rec) as fh:
        for evento, elem in iterparse(fh, events=("end",)):
            if elem.tag != NS + "r":
                continue

            valores = []
            for hijo in elem:
                etiqueta = hijo.tag[len(NS):]
                v = hijo.get("v")
                if etiqueta == "x":          # indice a la lista compartida
                    valores.append(int(v))
                elif etiqueta == "n":        # numero suelto
                    valores.append(float(v) if v is not None else None)
                elif etiqueta == "m":        # vacio
                    valores.append(None)
                else:                        # s, d, b, e
                    valores.append(v)

            fila = []
            for pos in quiero:
                bruto = valores[pos] if pos < len(valores) else None
                if isinstance(bruto, int) and shared[pos]:
                    fila.append(shared[pos][bruto] if bruto < len(shared[pos]) else None)
                else:
                    fila.append(bruto)
            filas.append(fila)

            elem.clear()

    return filas


def extraer(ruta_xlsx):
    with zipfile.ZipFile(ruta_xlsx) as zf:
        nombres_int = zf.namelist()
        ruta_def = next(n for n in nombres_int if "pivotCacheDefinition" in n and n.endswith(".xml"))
        ruta_rec = next(n for n in nombres_int if "pivotCacheRecords" in n and n.endswith(".xml"))

        nombres, shared = leer_definicion(zf, ruta_def)
        columnas = [c for c in CAMPOS_UTILES if c in nombres]
        filas = leer_registros(zf, ruta_rec, nombres, shared, columnas)

    df = pd.DataFrame(filas, columns=columnas)
    df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
    for c in ["IMPRESIONES", "CLICK", "LEADS", "COSTO", "Q_BRUTO", "Q_NETO", "Q_EMI", "Q_TER"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return df


if __name__ == "__main__":
    entrada = sys.argv[1]
    salida = sys.argv[2] if len(sys.argv) > 2 else "base_plana.csv"

    df = extraer(entrada)
    if salida.lower().endswith(".parquet"):
        df.to_parquet(salida, index=False)
    else:
        df.to_csv(salida, index=False, encoding="utf-8-sig")

    print(f"Filas extraidas : {len(df):,}")
    print(f"Rango de fechas : {df['FECHA'].min().date()} -> {df['FECHA'].max().date()}")
    print(f"Guardado en     : {salida}")
