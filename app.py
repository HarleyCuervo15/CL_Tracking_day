"""
app.py  ·  Tracking CPE
-----------------------------------------------------------
Subes la maqueta y listo. Las metas ya vienen adentro (metas.xlsx).

Para correrla:
    streamlit run app.py
"""

import calendar
import io
import os
import tempfile
from datetime import date

import pandas as pd
import streamlit as st

from extraer_base import VERSION as VERSION_EXTRACTOR, extraer
from tracking_cpe import BLOQUES, METRICAS, construir_diario, construir_libro

ARCHIVO_METAS = "metas.xlsx"

st.set_page_config(page_title="Tracking CPE", page_icon="📊", layout="wide")

# ----------------------------------------------------------------
# ESTILO
# ----------------------------------------------------------------
st.markdown("""
<style>
  /* Solo las tarjetas KPI. El resto lo maneja .streamlit/config.toml,
     que es lo que hace que menus, uploader y tablas se vean bien. */
  .tarjeta { background:#0b3977; border:1px solid #2c659e; border-top:3px solid #3ed598;
             border-radius:10px; padding:14px 16px; height:100%; }
  .tarjeta.gasto  { border-top-color:#ff5a5f; }
  .tarjeta.suelta { border-top-color:#ffb020; }
  .tarjeta h3 { font-size:12px; margin:0 0 8px; color:#c9dcf3; font-weight:600;
                letter-spacing:.3px; }
  .tarjeta .valor { font-size:23px; font-weight:700; line-height:1.15; color:#fff; }
  .tarjeta .sub { font-size:11px; color:#c9dcf3; margin-top:7px; line-height:1.5; }
  .tarjeta .ok    { color:#3ed598; font-weight:700; }
  .tarjeta .mal   { color:#ff5a5f; font-weight:700; }
  .tarjeta .tibio { color:#ffb020; font-weight:700; }
</style>
""", unsafe_allow_html=True)


# ----------------------------------------------------------------
# CARGA (con cache: la extraccion tarda ~30 s)
# ----------------------------------------------------------------
@st.cache_data(show_spinner=False, max_entries=3)
def cargar_base(contenido: bytes, version: int) -> pd.DataFrame:
    """version entra en la llave de cache aunque no se use adentro:
    al subirla, el resultado viejo deja de valer automaticamente."""
    return extraer(io.BytesIO(contenido))


@st.cache_data(show_spinner=False)
def cargar_metas(contenido: bytes | None) -> pd.DataFrame:
    if contenido is None:
        if not os.path.exists(ARCHIVO_METAS):
            return pd.DataFrame(columns=["FECHA", "BLOQUE", "CANAL2",
                                         "META_COSTO", "META_LEADS"])
        m = pd.read_excel(ARCHIVO_METAS)
    else:
        m = pd.read_excel(io.BytesIO(contenido))
    m["FECHA"] = pd.to_datetime(m["FECHA"])
    return m


def revisar_calidad(curva, dias_data, mes_nombre):
    """Dias donde un volumen se desploma frente a la mediana: casi siempre
    es un cargue incompleto, no una caida real."""
    avisos = []
    con_datos = curva.head(dias_data)
    nombres = {"LEADS": "Lead", "Q_NETO": "Q Neto · leads calificados",
               "Q_EMI": "Q Emi · venta cantada", "Q_TER": "Q Ter · venta"}
    for col in ["LEADS", "Q_NETO", "Q_EMI", "Q_TER"]:
        s = con_datos[col]
        if len(s) < 4 or s.median() <= 0:
            continue
        for dia, valor in s[s < s.median() * 0.25].items():
            falta = s.median() - valor
            avisos.append({
                "Métrica": nombres[col],
                "Día": f"{dia} de {mes_nombre}",
                "Ese día": valor,
                "Un día normal": s.median(),
                "Caída": f"-{1 - valor / s.median():.0%}",
                "Faltarían": f"~{falta:,.0f}"})
    return pd.DataFrame(avisos)


def tarjeta(titulo, valor, sub, clase=""):
    st.markdown(f'<div class="tarjeta {clase}"><h3>{titulo}</h3>'
                f'<div class="valor">{valor}</div>'
                f'<div class="sub">{sub}</div></div>', unsafe_allow_html=True)


def pinta(pct, invertido=False):
    """Verde si va bien. En costo y CPL, pasarse es malo."""
    if pct == 0:
        return "sub"
    if invertido:
        return "ok" if pct <= 1.02 else ("tibio" if pct <= 1.10 else "mal")
    return "ok" if pct >= 0.98 else ("tibio" if pct >= 0.85 else "mal")


# ----------------------------------------------------------------
# BARRA LATERAL
# ----------------------------------------------------------------
st.sidebar.title("Tracking CPE")
st.sidebar.caption("Sube la maqueta. Las metas ya están cargadas.")

subida = st.sidebar.file_uploader("Maqueta (.xlsx)", type=["xlsx"])

with st.sidebar.expander("Actualizar metas (1 vez al mes)"):
    st.caption("Solo si cambió el presupuesto. Sube el metas.xlsx nuevo "
               "que sale de generar_metas.py.")
    metas_subidas = st.file_uploader("metas.xlsx", type=["xlsx"], key="metas")

if subida is None:
    st.title("Tracking y proyección a cierre")
    st.info("Sube la maqueta en el panel de la izquierda para empezar.")
    m0 = cargar_metas(None)
    if not m0.empty:
        meses = sorted(m0["FECHA"].dt.strftime("%Y-%m").unique())
        st.success(f"Metas cargadas: {len(m0):,} filas · meses disponibles: "
                   f"{', '.join(meses)}")
    st.stop()

with st.spinner("Leyendo la base que viene dentro de la maqueta (~30 s la primera vez)..."):
    base = cargar_base(subida.getvalue(), VERSION_EXTRACTOR)

metas = cargar_metas(metas_subidas.getvalue() if metas_subidas else None)

meses_base = sorted(base["MES"].dropna().unique())
por_defecto = date.today().strftime("%Y-%m")
idx = meses_base.index(por_defecto) if por_defecto in meses_base else len(meses_base) - 1
mes = st.sidebar.selectbox("Mes", meses_base, index=idx)

areas = sorted(base["AREA"].dropna().unique())
area = st.sidebar.selectbox("Área", areas,
                            index=areas.index("CANAL ONLINE")
                            if "CANAL ONLINE" in areas else 0)

split = st.sidebar.slider("% de Mixto que va a Móvil y a Fijo", 0.0, 1.0, 0.50, 0.05)
solo_con_costo = st.sidebar.checkbox("Solo canales con inversión", value=False,
                                     help="Deja fuera Afiliados, TikTok y Sin clasificar, "
                                          "que generan ventas pero no tienen costo.")

st.sidebar.divider()
st.sidebar.caption("**Filtros** — el presupuesto no se abre por producto "
                   "ni por plataforma, así que al filtrar se ocultan las metas.")

_p = base[(base["MES"] == mes) & (base["AREA"] == area)]


def opciones(col):
    """Valores de una columna. Si la columna no existe (extractor viejo)
    devuelve lista vacia, en vez de tumbar la app."""
    if col not in _p.columns:
        return []
    return sorted(x for x in _p[col].astype(str).unique() if x and x != "nan")


DISPONIBLES = [
    ("PRODUCTO", "Producto", "BAF, TV, VOZ, ALTA SIN_EQ... Vacío = todos"),
    ("CANAL_SOLICITUD", "Plataforma",
     "De dónde entró la solicitud (WEB / APP). Vacío = todas"),
    ("FLUJO", "Flujo", "Magento, C2C, WhatsApp. Vacío = todos"),
]

filtros, faltantes = {}, []
for _col, _etq, _ayuda in DISPONIBLES:
    if _col not in base.columns:
        faltantes.append(_etq)
        continue
    filtros[_col] = st.sidebar.multiselect(_etq, opciones(_col), help=_ayuda)

hay_filtro = any(filtros.values())

if faltantes:
    st.sidebar.error(f"No disponibles: {', '.join(faltantes)}.")
    with st.sidebar.expander("¿Por qué?"):
        st.write(f"**Extractor en uso:** versión {VERSION_EXTRACTOR}")
        st.write("**Columnas que trajo:**")
        st.code(", ".join(base.columns))
        st.write("Si arriba no aparecen CANAL_SOLICITUD ni FLUJO, el "
                 "extraer_base.py del repositorio todavía es el viejo.")

with st.sidebar.expander("Problemas"):
    st.caption(f"Extractor versión {VERSION_EXTRACTOR} · "
               f"{len(base):,} filas · {len(base.columns)} columnas")
    if st.button("Limpiar caché y recargar"):
        st.cache_data.clear()
        st.rerun()

if hay_filtro:
    st.sidebar.warning("El costo no viene marcado por producto ni por flujo: esas "
                       "etiquetas solo existen en las filas de conversión. Al filtrar "
                       "verás volúmenes reales, pero el costo se va a cero.")

diario = construir_diario(base, mes, area, split, solo_con_costo, filtros)
if diario.empty:
    st.error(f"No hay datos para {mes} en el área {area} con los filtros elegidos.")
    st.stop()

anio, mm = int(mes[:4]), int(mes[5:7])
dias_mes = calendar.monthrange(anio, mm)[1]
dias_reales = diario["FECHA"].nunique()

dias_data = st.sidebar.number_input(
    "Días con información", 1, dias_mes, dias_reales,
    help="Bájalo si el último día viene incompleto. La proyección se recalcula sola.")
factor = dias_mes / dias_data
st.sidebar.metric("Factor de proyección", f"{factor:.4f}",
                  help=f"{dias_mes} días del mes ÷ {dias_data} días con info")

metas_mes = metas[metas["FECHA"].dt.strftime("%Y-%m") == mes]
if hay_filtro:
    metas_mes = metas_mes.iloc[0:0]

# ----------------------------------------------------------------
# CALCULOS
# ----------------------------------------------------------------
mf = diario[diario["BLOQUE"] != "MIXTO"]          # Movil + Fijo = universo completo
curva = (mf.groupby(mf["FECHA"].dt.day)[METRICAS].sum()
         .reindex(range(1, dias_mes + 1), fill_value=0))

mfm = metas_mes[metas_mes["BLOQUE"] != "MIXTO"]
if not mfm.empty:
    mcurva = (mfm.groupby(mfm["FECHA"].dt.day)[["META_COSTO", "META_LEADS"]].sum()
              .reindex(range(1, dias_mes + 1), fill_value=0))
else:
    mcurva = pd.DataFrame(0.0, index=range(1, dias_mes + 1),
                          columns=["META_COSTO", "META_LEADS"])
curva = curva.join(mcurva)

real = curva[METRICAS].sum()
meta_costo, meta_leads = curva["META_COSTO"].sum(), curva["META_LEADS"].sum()
proy = real * factor


def div(a, b):
    return a / b if b else 0.0


st.title(f"Tracking y proyección · {mes}")
st.caption(f"{area} · Móvil + Fijo (Mixto repartido al {split:.0%}) · "
           f"{dias_data} de {dias_mes} días · "
           f"último dato {diario['FECHA'].max().date()} · "
           f"proyección run-rate ×{factor:.4f}")

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
if hay_filtro:
    _nombres = {c: e for c, e, _ in DISPONIBLES}
    partes = [f"{_nombres[c].lower()}: {', '.join(v)}"
              for c, v in filtros.items() if v]
    st.info(
        f"**Vista filtrada** — {' · '.join(partes)}. "
        f"Las metas quedan ocultas a propósito: el presupuesto solo está abierto "
        f"por tipo y canal, no por producto ni plataforma. Repartirlo sería inventar "
        f"un número. Aquí ves real y proyectado; para comparar contra meta, quita "
        f"los filtros.")

alertas = revisar_calidad(curva, dias_data, MESES[mm - 1])
if not alertas.empty:
    st.warning(
        f"**Posible dato incompleto — revísalo antes de circular el reporte.** "
        f"Encontré {len(alertas)} día(s) donde un volumen quedó muy por debajo de lo "
        f"normal para ese mismo mes. Si la caída aparece en todos los canales a la vez, "
        f"casi siempre es un cargue que no corrió, no una caída de negocio. "
        f"El problema es que ese hueco arrastra el total del mes hacia abajo y te infla "
        f"el costo unitario proyectado (CPL, CPE o CPA).")
    st.dataframe(alertas.style.format({"Ese día": "{:,.0f}", "Un día normal": "{:,.0f}"}),
                 width="stretch", hide_index=True)
    st.caption("Qué hacer: confirma con quien maneja la fuente si ese día quedó completo. "
               "Si el dato no se puede recuperar, dilo explícitamente cuando presentes "
               "el proyectado.")

# ----------------------------------------------------------------
# TARJETAS
# ----------------------------------------------------------------
st.subheader("Cierre proyectado")
c = st.columns(3)
cpl, cpl_meta = div(real["COSTO"], real["LEADS"]), div(meta_costo, meta_leads)
with c[0]:
    p = div(proy["COSTO"], meta_costo)
    sub = (f"Real hoy ${real['COSTO']:,.0f}" if hay_filtro else
           f"Meta ${meta_costo:,.0f} · <span class='{pinta(p, True)}'>{p:.0%}</span><br>"
           f"Real hoy ${real['COSTO']:,.0f}")
    tarjeta("COSTO", f"${proy['COSTO']:,.0f}", sub, "gasto")
with c[1]:
    p = div(proy["LEADS"], meta_leads)
    sub = (f"Real hoy {real['LEADS']:,.0f}" if hay_filtro else
           f"Meta {meta_leads:,.0f} · <span class='{pinta(p)}'>{p:.0%}</span><br>"
           f"Real hoy {real['LEADS']:,.0f}")
    tarjeta("LEADS", f"{proy['LEADS']:,.0f}", sub)
with c[2]:
    p = div(cpl, cpl_meta)
    sub = ("Con run-rate el CPL proyectado es el mismo" if hay_filtro else
           f"Meta ${cpl_meta:,.0f} · <span class='{pinta(p, True)}'>{p:.0%}</span><br>"
           f"Con run-rate el CPL proyectado es el mismo")
    tarjeta("CPL", f"${cpl:,.0f}", sub, "gasto")

st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
st.caption("Embudo — estas métricas no tienen meta en el presupuesto")
c = st.columns(3)
etiquetas = {"Q_NETO": ("Q NETO · leads calificados", "CPL Neto"),
             "Q_EMI": ("Q EMI · venta cantada", "CPE"),
             "Q_TER": ("Q TER · venta", "CPA")}
for col, (campo, (nombre, costo_unit)) in zip(c, etiquetas.items()):
    with col:
        unit = div(real["COSTO"], real[campo])
        tarjeta(nombre, f"{proy[campo]:,.0f}",
                f"{costo_unit} ${unit:,.0f}<br>Real hoy {real[campo]:,.0f}", "suelta")

# ----------------------------------------------------------------
# PESTAÑAS
# ----------------------------------------------------------------
t1, t2, t3, t4 = st.tabs(["Curva diaria", "Maqueta", "Control", "Descargar"])

with t1:
    v = curva.copy()
    v["% Costo"] = (v["COSTO"] / v["META_COSTO"].replace(0, pd.NA)).fillna(0)
    v["CPL"] = (v["COSTO"] / v["LEADS"].replace(0, pd.NA)).fillna(0)
    v["CPL Neto"] = (v["COSTO"] / v["Q_NETO"].replace(0, pd.NA)).fillna(0)
    v["CPE"] = (v["COSTO"] / v["Q_EMI"].replace(0, pd.NA)).fillna(0)
    v["CPA"] = (v["COSTO"] / v["Q_TER"].replace(0, pd.NA)).fillna(0)
    v["Costo acum"] = v["COSTO"].cumsum()
    v["Meta acum"] = v["META_COSTO"].cumsum()
    v = v.rename(columns={"META_COSTO": "Meta Costo", "COSTO": "Costo",
                          "META_LEADS": "Meta Lead", "LEADS": "Lead",
                          "Q_NETO": "Q Neto", "Q_EMI": "Q Emi", "Q_TER": "Q Ter"})
    v.index.name = "Día"

    st.dataframe(
        v[["Meta Costo", "Costo", "% Costo", "Meta Lead", "Lead", "CPL",
           "Q Neto", "CPL Neto", "Q Emi", "CPE", "Q Ter", "CPA",
           "Costo acum", "Meta acum"]].style.format({
               "Meta Costo": "${:,.0f}", "Costo": "${:,.0f}", "% Costo": "{:.0%}",
               "Meta Lead": "{:,.0f}", "Lead": "{:,.0f}", "CPL": "${:,.0f}",
               "Q Neto": "{:,.0f}", "CPL Neto": "${:,.0f}",
               "Q Emi": "{:,.0f}", "CPE": "${:,.0f}",
               "Q Ter": "{:,.0f}", "CPA": "${:,.0f}",
               "Costo acum": "${:,.0f}", "Meta acum": "${:,.0f}"}),
        width="stretch", height=430)

    st.markdown("**Acumulado: real contra meta**")
    st.line_chart(v[["Costo acum", "Meta acum"]], height=280)

    st.markdown("**Volumen diario**")
    st.bar_chart(v.head(dias_data)[["Lead", "Q Neto", "Q Emi", "Q Ter"]], height=280)

with t2:
    for bloque in BLOQUES:
        d = diario[diario["BLOQUE"] == bloque]
        mb = metas_mes[metas_mes["BLOQUE"] == bloque]
        canales = sorted(set(d["CANAL2"]) | set(mb["CANAL2"]))
        if not canales:
            continue

        r = d.groupby("CANAL2")[METRICAS].sum().reindex(canales, fill_value=0)
        mt = (mb.groupby("CANAL2")[["META_COSTO", "META_LEADS"]].sum()
              .reindex(canales, fill_value=0))

        tabla = pd.DataFrame(index=canales)
        tabla["Meta Costo"] = mt["META_COSTO"]
        tabla["Real Costo"] = r["COSTO"]
        tabla["Costo proy"] = r["COSTO"] * factor
        tabla["Meta Leads"] = mt["META_LEADS"]
        tabla["Real Leads"] = r["LEADS"]
        tabla["Leads proy"] = r["LEADS"] * factor
        tabla["CPL Meta"] = mt["META_COSTO"] / mt["META_LEADS"].replace(0, pd.NA)
        tabla["CPL"] = r["COSTO"] / r["LEADS"].replace(0, pd.NA)
        tabla["% Cumpl"] = tabla["CPL"] / tabla["CPL Meta"]
        for campo, (_, unit) in etiquetas.items():
            tabla[campo.replace("Q_", "Q ").title()] = r[campo]
            tabla[unit] = r["COSTO"] / r[campo].replace(0, pd.NA)
        tabla = tabla.fillna(0)
        tabla.loc["Total"] = tabla.sum()
        for a, b, dest in [("Meta Costo", "Meta Leads", "CPL Meta"),
                           ("Real Costo", "Real Leads", "CPL")]:
            tabla.loc["Total", dest] = div(tabla.loc["Total", a], tabla.loc["Total", b])
        tabla.loc["Total", "% Cumpl"] = div(tabla.loc["Total", "CPL"],
                                            tabla.loc["Total", "CPL Meta"])
        for campo, (_, unit) in etiquetas.items():
            tabla.loc["Total", unit] = div(tabla.loc["Total", "Real Costo"],
                                           tabla.loc["Total", campo.replace("Q_", "Q ").title()])

        st.markdown(f"### {bloque}")
        st.dataframe(tabla.style.format({
            "Meta Costo": "${:,.0f}", "Real Costo": "${:,.0f}", "Costo proy": "${:,.0f}",
            "Meta Leads": "{:,.0f}", "Real Leads": "{:,.0f}", "Leads proy": "{:,.0f}",
            "CPL Meta": "${:,.0f}", "CPL": "${:,.0f}", "% Cumpl": "{:.0%}",
            "Q Neto": "{:,.0f}", "CPL Neto": "${:,.0f}",
            "Q Emi": "{:,.0f}", "CPE": "${:,.0f}",
            "Q Ter": "{:,.0f}", "CPA": "${:,.0f}"}), width="stretch")

with t3:
    r_real = set(zip(diario["BLOQUE"], diario["CANAL2"]))
    r_meta = set(zip(metas_mes["BLOQUE"], metas_mes["CANAL2"]))
    costo_canal = diario.groupby(["BLOQUE", "CANAL2"])["COSTO"].sum()

    revisar, esperado = [], []
    for b, c in sorted(r_real - r_meta):
        if costo_canal.get((b, c), 0) > 0:
            revisar.append({"Bloque": b, "Canal 2": c,
                            "Situación": "Tiene inversión pero no tiene meta en el PPTo",
                            "Qué hacer": "Cargar la meta o confirmar que no lleva"})
        else:
            esperado.append({"Bloque": b, "Canal 2": c,
                             "Situación": "Genera ventas sin costo de medios",
                             "Qué hacer": "Normal: no necesita meta de inversión"})
    for b, c in sorted(r_meta - r_real):
        revisar.append({"Bloque": b, "Canal 2": c,
                        "Situación": "Tiene meta pero todavía no registra inversión",
                        "Qué hacer": "Verificar si la campaña ya arrancó"})

    if revisar:
        st.error(f"{len(revisar)} combinación(es) para revisar.")
        st.dataframe(pd.DataFrame(revisar), width="stretch", hide_index=True)
    else:
        st.success("Todo cruza. Ningún canal con inversión quedó sin meta.")

    if esperado:
        st.caption("Sin meta, pero es lo esperado:")
        st.dataframe(pd.DataFrame(esperado), width="stretch", hide_index=True)

with t4:
    st.write("El Excel trae las mismas hojas del script, con fórmulas vivas.")
    if st.button("Generar Excel", type="primary"):
        with st.spinner("Armando el archivo..."):
            ruta = os.path.join(tempfile.mkdtemp(), f"Tracking_CPE_{mes}.xlsx")
            construir_libro(diario, metas_mes, mes, ruta)
            with open(ruta, "rb") as fh:
                datos = fh.read()
        st.download_button("Descargar", datos, file_name=f"Tracking_CPE_{mes}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet")

    st.divider()
    st.download_button("Descargar la base plana (CSV)",
                       diario.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"base_diaria_{mes}.csv", mime="text/csv")
