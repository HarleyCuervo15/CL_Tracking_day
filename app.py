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
from tracking_cpe import (BLOQUES, METRICAS, construir_diario,
                          construir_libro, limpiar_metas, repartir_metas,
                          universo)

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
def cargar_metas(contenido: bytes | None, huella: tuple = ()) -> pd.DataFrame:
    """huella entra en la llave del cache: si el metas.xlsx del disco cambia,
    el resultado guardado deja de valer solo."""
    if contenido is None:
        if not os.path.exists(ARCHIVO_METAS):
            return pd.DataFrame(columns=["FECHA", "BLOQUE", "CANAL2", "META_COSTO",
                                         "META_LEADS", "META_VENTAS"])
        m = pd.read_excel(ARCHIVO_METAS)
    else:
        m = pd.read_excel(io.BytesIO(contenido))
    m["FECHA"] = pd.to_datetime(m["FECHA"])
    if "META_VENTAS" not in m.columns:
        m["META_VENTAS"] = 0.0
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

if metas_subidas:
    metas = cargar_metas(metas_subidas.getvalue())
else:
    try:
        _st = os.stat(ARCHIVO_METAS)
        huella = (round(_st.st_mtime, 3), _st.st_size)
    except OSError:
        huella = ()
    metas = cargar_metas(None, huella)

meses_base = sorted(base["MES"].dropna().unique())
por_defecto = date.today().strftime("%Y-%m")
idx = meses_base.index(por_defecto) if por_defecto in meses_base else len(meses_base) - 1
mes = st.sidebar.selectbox("Mes", meses_base, index=idx)

areas = sorted(base["AREA"].dropna().unique())
area = st.sidebar.selectbox("Área", areas,
                            index=areas.index("CANAL ONLINE")
                            if "CANAL ONLINE" in areas else 0)

split = st.sidebar.slider(
    "% de Mixto que va a Móvil y a Fijo", 0.0, 1.0, 0.00, 0.05,
    help="0% = cada bloque queda puro y Mixto va aparte. "
         "50% = el reparto que asume tu archivo de presupuesto. "
         "Se aplica igual al real y a las metas.")
if split == 0:
    st.sidebar.caption("Móvil y Fijo quedan puros. Mixto se ve en su propio bloque.")
else:
    st.sidebar.caption(f"Móvil y Fijo incluyen {split:.0%} de Mixto cada uno.")
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

st.sidebar.divider()
bloques_sel = st.sidebar.multiselect(
    "Tipo", BLOQUES, default=BLOQUES,
    help="Móvil y Fijo ya incluyen el Mixto repartido. Este filtro SÍ es "
         "compatible con las metas: el presupuesto está abierto por tipo.")
if not bloques_sel:
    bloques_sel = BLOQUES

if hay_filtro:
    st.sidebar.warning("El costo no viene marcado por producto ni por flujo: esas "
                       "etiquetas solo existen en las filas de conversión. Al filtrar "
                       "verás volúmenes reales, pero el costo se va a cero.")

diario = construir_diario(base, mes, area, split, solo_con_costo, filtros)
if diario.empty:
    st.error(f"No hay datos para {mes} en el área {area} con los filtros elegidos.")
    st.stop()

diario = diario[diario["BLOQUE"].isin(bloques_sel)]
if diario.empty:
    st.error("Los filtros dejaron la selección vacía.")
    st.stop()

anio, mm = int(mes[:4]), int(mes[5:7])
dias_mes = calendar.monthrange(anio, mm)[1]

# Rango de fechas: por defecto del dia 1 al ultimo con datos
tope = diario["FECHA"].max().date()
piso = date(anio, mm, 1)
fin_mes = date(anio, mm, dias_mes)
rango = st.sidebar.date_input("Rango de fechas", value=(piso, tope),
                              min_value=piso, max_value=fin_mes,
                              help="Por defecto va del día 1 al último con datos. "
                                   "Acórtalo para proyectar con el ritmo de los "
                                   "últimos días en vez del promedio del mes.")
if isinstance(rango, (list, tuple)) and len(rango) == 2:
    f_ini, f_fin = rango
else:
    f_ini, f_fin = piso, tope

diario = diario[(diario["FECHA"].dt.date >= f_ini) &
                (diario["FECHA"].dt.date <= f_fin)]
if diario.empty:
    st.error(f"No hay datos entre {f_ini} y {f_fin}.")
    st.stop()

ventana_completa = (f_ini == piso and f_fin == tope)
dias_reales = diario["FECHA"].nunique()

dias_data = st.sidebar.number_input(
    "Días con información", 1, dias_mes, dias_reales,
    help="Bájalo si el último día viene incompleto. La proyección se recalcula sola.")
factor = dias_mes / dias_data
st.sidebar.metric("Factor de proyección", f"{factor:.4f}",
                  help=f"{dias_mes} días del mes ÷ {dias_data} días con info")

# El mismo reparto que se aplica al real, para que los dos lados cuadren
_mes_metas = metas[metas["FECHA"].dt.strftime("%Y-%m") == mes]
_mes_metas, venia_repartido = limpiar_metas(_mes_metas)
metas_mes = repartir_metas(_mes_metas, split)
metas_mes = metas_mes[metas_mes["BLOQUE"].isin(bloques_sel)]
if hay_filtro:
    metas_mes = metas_mes.iloc[0:0]
# Meta del periodo elegido (los dias dentro del rango)
metas_per = metas_mes[(metas_mes["FECHA"].dt.date >= f_ini) &
                      (metas_mes["FECHA"].dt.date <= f_fin)]

# ----------------------------------------------------------------
# CALCULOS
# ----------------------------------------------------------------
# Bloques que suman el total sin contar dos veces: si Mixto se reparte,
# ya esta dentro de Movil y Fijo; si no, entra por su cuenta.
UNI = [b for b in universo(split) if b in bloques_sel] or bloques_sel
mf = diario[diario["BLOQUE"].isin(UNI)]
mfm = metas_mes[metas_mes["BLOQUE"].isin(UNI)]
metas_per = metas_per[metas_per["BLOQUE"].isin(UNI)]

dias_ventana = [d.day for d in pd.date_range(f_ini, f_fin)]
curva = (mf.groupby(mf["FECHA"].dt.day, observed=True)[METRICAS].sum()
         .reindex(dias_ventana, fill_value=0))

if not mfm.empty:
    mcurva = (mfm.groupby(mfm["FECHA"].dt.day)[["META_COSTO", "META_LEADS"]].sum()
              .reindex(dias_ventana, fill_value=0))
else:
    mcurva = pd.DataFrame(0.0, index=dias_ventana,
                          columns=["META_COSTO", "META_LEADS"])
curva = curva.join(mcurva)

real = curva[METRICAS].sum()
# Meta del periodo (para comparar contra lo real) y del mes (contra el proyectado)
meta_costo = curva["META_COSTO"].sum()
meta_leads = curva["META_LEADS"].sum()
meta_mes_costo = mfm["META_COSTO"].sum()
meta_mes_leads = mfm["META_LEADS"].sum()
proy = real * factor


def div(a, b):
    return a / b if b else 0.0


st.title(f"Tracking y proyección · {mes}")
etq_bloque = ("Móvil + Fijo" if set(bloques_sel) == set(BLOQUES)
              else " + ".join(b.title() for b in bloques_sel))
etq_mix = ("Mixto aparte" if split == 0 else f"Mixto repartido al {split:.0%}")
st.caption(f"{area} · {etq_bloque} ({etq_mix}) · "
           f"del {f_ini:%d/%m} al {f_fin:%d/%m} · {dias_data} de {dias_mes} días · "
           f"proyección run-rate ×{factor:.4f}")

if not ventana_completa:
    st.info(f"**Ventana recortada** — estás viendo del {f_ini:%d/%m} al {f_fin:%d/%m}, "
            f"no el mes completo. El proyectado asume que el resto del mes se comporta "
            f"como esta ventana, así que sirve para preguntarte *¿y si el mes siguiera "
            f"al ritmo de estos días?*. Las metas de las tarjetas son las de estos "
            f"mismos días, no las del mes.")

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

if venia_repartido:
    st.warning(
        "**El archivo de metas traía el reparto de Mixto adentro.** Le quité las "
        "filas '… Mixto' de Móvil y Fijo, porque el bloque Mixto ya trae el monto "
        "completo y si no quedaría contado dos veces. Los números de esta pantalla "
        "ya están corregidos. Para dejarlo limpio de raíz, vuelve a generar "
        "metas.xlsx con generar_metas.py y súbelo al repositorio.")

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
cpl = div(real["COSTO"], real["LEADS"])
cpl_meta = div(meta_costo, meta_leads)
with c[0]:
    pp = div(proy["COSTO"], meta_mes_costo)          # proyectado vs meta del mes
    pr = div(real["COSTO"], meta_costo)              # real vs meta del periodo
    sub = (f"Real período ${real['COSTO']:,.0f}" if hay_filtro else
           f"vs meta mes ${meta_mes_costo:,.0f} · "
           f"<span class='{pinta(pp, True)}'>{pp:.0%}</span><br>"
           f"Real período ${real['COSTO']:,.0f} · meta ${meta_costo:,.0f} · "
           f"<span class='{pinta(pr, True)}'>{pr:.0%}</span>")
    tarjeta("COSTO", f"${proy['COSTO']:,.0f}", sub, "gasto")
with c[1]:
    pp = div(proy["LEADS"], meta_mes_leads)
    pr = div(real["LEADS"], meta_leads)
    sub = (f"Real período {real['LEADS']:,.0f}" if hay_filtro else
           f"vs meta mes {meta_mes_leads:,.0f} · "
           f"<span class='{pinta(pp)}'>{pp:.0%}</span><br>"
           f"Real período {real['LEADS']:,.0f} · meta {meta_leads:,.0f} · "
           f"<span class='{pinta(pr)}'>{pr:.0%}</span>")
    tarjeta("LEADS", f"{proy['LEADS']:,.0f}", sub)
with c[2]:
    pr = div(cpl, cpl_meta)
    sub = ("Con run-rate el CPL proyectado es el mismo" if hay_filtro else
           f"Meta del período ${cpl_meta:,.0f} · "
           f"<span class='{pinta(pr, True)}'>{pr:.0%}</span><br>"
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
etq_hoy = "a hoy" if ventana_completa else "período"

t0, t1, t2, t3, t4 = st.tabs(["Resumen", "Curva diaria", "Maqueta",
                              "Control", "Descargar"])

with t0:
    ETQ = {"MOVIL": "MÓVIL", "FIJO": "FIJO", "MIXTO": "MIXTO"}
    usa = UNI

    rb = (diario.groupby("BLOQUE", observed=True)[METRICAS].sum()
          .reindex(bloques_sel, fill_value=0))
    if metas_mes.empty:
        mb = pd.DataFrame(0.0, index=bloques_sel,
                          columns=["META_COSTO", "META_LEADS", "META_VENTAS"])
    else:
        mb = (metas_mes.groupby("BLOQUE")[["META_COSTO", "META_LEADS", "META_VENTAS"]]
              .sum().reindex(bloques_sel, fill_value=0))

    tot_r, tot_m = rb.loc[usa].sum(), mb.loc[usa].sum()
    # Ventas: solo los bloques que sí tienen meta cargada
    con_meta = [b for b in usa if mb.loc[b, "META_VENTAS"] > 0]
    v_meta = mb.loc[con_meta, "META_VENTAS"].sum() if con_meta else 0.0
    v_real = rb.loc[con_meta, "Q_TER"].sum() if con_meta else 0.0
    c_costo_m = mb.loc[con_meta, "META_COSTO"].sum() if con_meta else 0.0
    c_costo_r = rb.loc[con_meta, "COSTO"].sum() if con_meta else 0.0

    st.markdown("#### RESUMEN GENERAL")
    st.caption(f"Performance CL · corte {f_fin:%d/%m/%Y} · "
               f"{' + '.join(ETQ[b] for b in usa)}")

    fichas = [
        ("INVERSIÓN", "$", tot_r["COSTO"], tot_m["META_COSTO"], True),
        ("LEADS GENERADOS", "", tot_r["LEADS"], tot_m["META_LEADS"], False),
        ("EMISIONES / NETOS", "", tot_r["Q_EMI"], 0.0, False),
        ("VENTAS (solo con meta)", "", v_real, v_meta, False),
    ]
    cols = st.columns(5)
    for col, (nombre, sig, r_, m_, gasto) in zip(cols, fichas):
        with col:
            p_ = div(r_ * factor, m_)
            fmt = (lambda x: f"${x:,.0f}") if sig else (lambda x: f"{x:,.0f}")
            meta_txt = (f"Meta mes {fmt(m_)}<br>"
                        f"Cumpl. proy. <span class='{pinta(p_, gasto)}'>{p_:.0%}</span>"
                        if m_ else "Sin meta en el presupuesto")
            tarjeta(nombre, fmt(r_ * factor),
                    f"Real corte {fmt(r_)}<br>{meta_txt}",
                    "gasto" if gasto else ("suelta" if not m_ else ""))
    with cols[4]:
        cplm, cplp = div(tot_m["META_COSTO"], tot_m["META_LEADS"]), \
            div(tot_r["COSTO"], tot_r["LEADS"])
        cpam, cpap = div(c_costo_m, v_meta), div(c_costo_r, v_real)
        tarjeta("COSTOS",
                f"CPL ${cplp:,.0f}",
                f"CPL meta ${cplm:,.0f} · "
                f"<span class='{pinta(div(cplp, cplm), True)}'>"
                f"{div(cplp, cplm):.0%}</span><br>"
                f"CPA ${cpap:,.0f} · meta ${cpam:,.0f} · "
                f"<span class='{pinta(div(cpap, cpam), True)}'>"
                f"{div(cpap, cpam):.0%}</span>", "suelta")

    st.markdown("##### Cumplimiento por producto")
    filas = []
    for b in bloques_sel + ["TOTAL"]:
        r_ = tot_r if b == "TOTAL" else rb.loc[b]
        m_ = tot_m if b == "TOTAL" else mb.loc[b]
        vm = v_meta if b == "TOTAL" else m_["META_VENTAS"]
        vr = v_real if b == "TOTAL" else r_["Q_TER"]
        fila = {"Producto": ETQ.get(b, b)}
        for etq, real_v, meta_v in [
                ("Presupuesto", r_["COSTO"], m_["META_COSTO"]),
                ("Leads", r_["LEADS"], m_["META_LEADS"]),
                ("Emitidos", r_["Q_EMI"], 0.0),
                ("Ventas", vr, vm)]:
            fila[(etq, "Meta")] = meta_v
            fila[(etq, "Proy.")] = real_v * factor
            fila[(etq, "% cumpl.")] = div(real_v * factor, meta_v)
        filas.append(fila)
    tc = pd.DataFrame(filas).set_index("Producto")
    tc.columns = pd.MultiIndex.from_tuples(tc.columns)
    st.dataframe(tc.style.format(
        {c: ("{:.0%}" if c[1] == "% cumpl." else
             ("${:,.0f}" if c[0] == "Presupuesto" else "{:,.0f}"))
         for c in tc.columns}), width="stretch")
    st.caption(
        ("TOTAL = " + " + ".join(ETQ.get(b, b) for b in usa) + ". " +
         ("Mixto no se suma: ya viene repartido dentro de los dos. "
          if "MIXTO" not in usa else "Los bloques están puros. ") +
         "Emisiones no tiene meta en el presupuesto, y la de ventas solo "
         "viene cargada para algunos bloques."))

    st.markdown("##### Costos")
    filas = []
    for b in bloques_sel + ["TOTAL"]:
        r_ = tot_r if b == "TOTAL" else rb.loc[b]
        m_ = tot_m if b == "TOTAL" else mb.loc[b]
        cm = c_costo_m if b == "TOTAL" else m_["META_COSTO"]
        vm = v_meta if b == "TOTAL" else m_["META_VENTAS"]
        cr = c_costo_r if b == "TOTAL" else r_["COSTO"]
        vr = v_real if b == "TOTAL" else r_["Q_TER"]
        fila = {"Producto": ETQ.get(b, b)}
        for etq, meta_v, real_v in [
                ("CPL", div(m_["META_COSTO"], m_["META_LEADS"]),
                 div(r_["COSTO"], r_["LEADS"])),
                ("CPE", 0.0, div(r_["COSTO"], r_["Q_EMI"])),
                ("CPA", div(cm, vm) if vm else 0.0, div(cr, vr) if vr else 0.0)]:
            fila[(etq, "Meta")] = meta_v
            fila[(etq, "Proy.")] = real_v
            fila[(etq, "Var %")] = (real_v / meta_v - 1) if meta_v else 0.0
        filas.append(fila)
    tk = pd.DataFrame(filas).set_index("Producto")
    tk.columns = pd.MultiIndex.from_tuples(tk.columns)
    st.dataframe(tk.style.format(
        {c: ("{:+.0%}" if c[1] == "Var %" else "${:,.0f}") for c in tk.columns}),
        width="stretch")
    st.caption("Con run-rate simple, el costo unitario proyectado es igual al actual: "
               "costo y volumen se multiplican por el mismo factor.")


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
    for bloque in bloques_sel:
        d = diario[diario["BLOQUE"] == bloque]
        mb = metas_mes[metas_mes["BLOQUE"] == bloque]
        canales = sorted(set(d["CANAL2"]) | set(mb["CANAL2"]))
        if not canales:
            continue

        r = d.groupby("CANAL2")[METRICAS].sum().reindex(canales, fill_value=0)
        mt = (mb.groupby("CANAL2")[["META_COSTO", "META_LEADS"]].sum()
              .reindex(canales, fill_value=0))

        # Meta acumulada hasta la fecha de corte (el PPTo es diario)
        mh = (mb[(mb["FECHA"].dt.date >= f_ini) & (mb["FECHA"].dt.date <= f_fin)]
              .groupby("CANAL2")[["META_COSTO", "META_LEADS"]].sum()
              .reindex(canales, fill_value=0))

        tabla = pd.DataFrame(index=canales)
        tabla["Meta Costo"] = mt["META_COSTO"]
        tabla[f"Meta Costo {etq_hoy}"] = mh["META_COSTO"]
        tabla["Real Costo"] = r["COSTO"]
        tabla[f"% Costo {etq_hoy}"] = (r["COSTO"] /
                                       mh["META_COSTO"].replace(0, pd.NA))
        tabla["Costo proy"] = r["COSTO"] * factor
        tabla["% Costo cierre"] = (tabla["Costo proy"] /
                                   mt["META_COSTO"].replace(0, pd.NA))
        tabla["Meta Leads"] = mt["META_LEADS"]
        tabla[f"Meta Leads {etq_hoy}"] = mh["META_LEADS"]
        tabla["Real Leads"] = r["LEADS"]
        tabla[f"% Leads {etq_hoy}"] = (r["LEADS"] /
                                       mh["META_LEADS"].replace(0, pd.NA))
        tabla["Leads proy"] = r["LEADS"] * factor
        tabla["% Leads cierre"] = (tabla["Leads proy"] /
                                   mt["META_LEADS"].replace(0, pd.NA))
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
        for num, den, dest in [
                ("Real Costo", f"Meta Costo {etq_hoy}", f"% Costo {etq_hoy}"),
                ("Costo proy", "Meta Costo", "% Costo cierre"),
                ("Real Leads", f"Meta Leads {etq_hoy}", f"% Leads {etq_hoy}"),
                ("Leads proy", "Meta Leads", "% Leads cierre")]:
            tabla.loc["Total", dest] = div(tabla.loc["Total", num],
                                           tabla.loc["Total", den])
        tabla.loc["Total", "% Cumpl"] = div(tabla.loc["Total", "CPL"],
                                            tabla.loc["Total", "CPL Meta"])
        for campo, (_, unit) in etiquetas.items():
            tabla.loc["Total", unit] = div(tabla.loc["Total", "Real Costo"],
                                           tabla.loc["Total", campo.replace("Q_", "Q ").title()])

        st.markdown(f"### {bloque}")
        st.dataframe(tabla.style.format({
            "Meta Costo": "${:,.0f}", "Real Costo": "${:,.0f}", "Costo proy": "${:,.0f}",
            f"Meta Costo {etq_hoy}": "${:,.0f}", f"% Costo {etq_hoy}": "{:.0%}",
            "% Costo cierre": "{:.0%}",
            "Meta Leads": "{:,.0f}", "Real Leads": "{:,.0f}", "Leads proy": "{:,.0f}",
            f"Meta Leads {etq_hoy}": "{:,.0f}", f"% Leads {etq_hoy}": "{:.0%}",
            "% Leads cierre": "{:.0%}",
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
            construir_libro(diario, metas_mes, mes, ruta, split)
            with open(ruta, "rb") as fh:
                datos = fh.read()
        st.download_button("Descargar", datos, file_name=f"Tracking_CPE_{mes}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet")

    st.divider()
    st.download_button("Descargar la base plana (CSV)",
                       diario.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"base_diaria_{mes}.csv", mime="text/csv")
