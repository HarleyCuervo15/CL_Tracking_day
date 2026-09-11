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
from array import array
from xml.etree.ElementTree import iterparse

import numpy as np
import pandas as pd

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# Subir este numero cuando cambien CAMPOS_UTILES o la logica de extraccion.
# La app lo usa como parte de la llave de cache: asi un cambio de codigo
# invalida solito el DataFrame guardado, sin tener que limpiar nada a mano.
VERSION = 2

# Campos que realmente necesitamos para el tracking.
CAMPOS_UTILES = [
    "MES", "SEMANA", "FECHA", "AREA", "TIPO", "CANAL", "CATEGORIA",
    "PRODUCTO", "FAMILIA", "CANAL_SOLICITUD", "FLUJO",
    "LEADS", "COSTO", "Q_BRUTO", "Q_NETO", "Q_EMI", "Q_TER",
]

# Columnas de texto: se guardan como categoria. Son pocos valores distintos
# repetidos 600 mil veces, asi que pasa de ~770 MB a un tercio.
TEXTO = ["MES", "SEMANA", "AREA", "TIPO", "CANAL", "CATEGORIA",
         "PRODUCTO", "FAMILIA", "CANAL_SOLICITUD", "FLUJO"]
NUMERO = ["LEADS", "COSTO", "Q_BRUTO", "Q_NETO", "Q_EMI", "Q_TER"]


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
    """Recorre la cache en streaming y llena un array por columna.

    Guardar codigos enteros en array.array en vez de listas de objetos
    Python baja el pico de memoria de ~670 MB a ~150 MB, que es lo que
    permite correr esto en Streamlit Cloud.
    """
    quiero = [nombres.index(c) for c in columnas]
    es_cat = [bool(shared[p]) for p in quiero]

    # 'i' = codigo de categoria (4 bytes)   'f' = numero (4 bytes)
    buffers = [array("i") if cat else array("f") for cat in es_cat]
    sueltos = {}          # valores que no venian de la lista compartida

    with zf.open(ruta_rec) as fh:
        n = 0
        for _, elem in iterparse(fh, events=("end",)):
            if elem.tag != NS + "r":
                continue

            crudos = []
            for hijo in elem:
                etiqueta = hijo.tag[len(NS):]
                v = hijo.get("v")
                if etiqueta == "x":
                    crudos.append(int(v))
                elif etiqueta == "n":
                    crudos.append(float(v) if v is not None else 0.0)
                elif etiqueta == "m":
                    crudos.append(None)
                else:
                    crudos.append(v)

            for j, pos in enumerate(quiero):
                val = crudos[pos] if pos < len(crudos) else None
                if es_cat[j]:
                    if isinstance(val, int):
                        buffers[j].append(val)
                    else:
                        buffers[j].append(-1)
                        if val is not None:
                            sueltos.setdefault(j, {})[n] = val
                else:
                    buffers[j].append(0.0 if val is None else
                                      (float(val) if not isinstance(val, str) else 0.0))
            n += 1
            elem.clear()

    return buffers, es_cat, sueltos, n


def extraer(ruta_xlsx):
    with zipfile.ZipFile(ruta_xlsx) as zf:
        nombres_int = zf.namelist()
        ruta_def = next(n for n in nombres_int if "pivotCacheDefinition" in n and n.endswith(".xml"))
        ruta_rec = next(n for n in nombres_int if "pivotCacheRecords" in n and n.endswith(".xml"))

        nombres, shared = leer_definicion(zf, ruta_def)
        columnas = [c for c in CAMPOS_UTILES if c in nombres]
        quiero = [nombres.index(c) for c in columnas]
        buffers, es_cat, sueltos, n = leer_registros(zf, ruta_rec, nombres,
                                                     shared, columnas)

    datos = {}
    for j, col in enumerate(columnas):
        if es_cat[j]:
            cats = [("" if x is None else x) for x in shared[quiero[j]]]
            serie = pd.Categorical.from_codes(np.asarray(buffers[j], dtype="int32"),
                                              categories=pd.Index(cats).astype(str))
            serie = pd.Series(serie)
            if j in sueltos:                       # valores fuera de la lista
                for fila, val in sueltos[j].items():
                    serie = serie.cat.add_categories([str(val)]) \
                        if str(val) not in serie.cat.categories else serie
                    serie.iat[fila] = str(val)
            datos[col] = serie
        else:
            datos[col] = pd.Series(np.asarray(buffers[j], dtype="float32"))
        buffers[j] = None

    df = pd.DataFrame(datos)
    del datos, buffers

    df["FECHA"] = pd.to_datetime(df["FECHA"].astype(str), errors="coerce")
    for c in NUMERO:
        if c in df.columns:
            df[c] = df[c].astype("float32")
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
