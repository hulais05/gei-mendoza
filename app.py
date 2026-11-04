"""
Tablero GEI Mendoza — Minería e Hidrocarburos (ESG Consulting).

streamlit run app.py
"""

import io
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from gei import motor
from gei.factores import UMBRAL_TCO2E
from procesar import procesar

AQUI = Path(__file__).parent
EJEMPLO = AQUI / "Matriz_GEI_Mendoza_EJEMPLO.xlsx"

# Referencias provinciales (fuente: tablero IDE Mendoza / Inventario GEI provincial)
INV_PROVINCIAL_2023 = 10_474_412.31          # tCO2e, inventario provincial 2023
REGISTRO_2023 = {"Hidrocarburos": ("Energía", 4_056_976.90),
                 "Minería": ("Industrial", 1_522_648.10)}   # Registro Provincial 2023 por sector

st.set_page_config(page_title="GEI Mendoza · Minería e Hidrocarburos", page_icon="🌎", layout="wide")

OSCURO = (st.context.theme.type or "dark") != "light"
_UA = st.context.headers.get("User-Agent", "") if st.context.headers else ""
MOVIL = any(k in _UA for k in ("Mobile", "Android", "iPhone"))
C = {  # paleta de referencia validada (dataviz), pasos claro / oscuro
    "a1": "#3987e5" if OSCURO else "#2a78d6",
    "a2": "#d95926" if OSCURO else "#eb6834",
    "prev": "#6b6a64" if OSCURO else "#b9b8b2",
    "txt": "#f0efec" if OSCURO else "#0b0b0b",
    "txt2": "#c3c2b7" if OSCURO else "#52514e",
    "grid": "#2c2c2a" if OSCURO else "#e6e5e0",
    "card": "#161B27" if OSCURO else "#f6f6f4",
    "borde": "#2c3444" if OSCURO else "#e1e0da",
    "alerta": "#e0a030" if OSCURO else "#B45309",
    "ok": "#34b37f" if OSCURO else "#047857",
}


def fmt(x, dec=0):
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def layout(fig, alto=360, leyenda=True):
    fig.update_layout(
        height=alto, margin=dict(l=8, r=24, t=40 if leyenda else 8, b=8), paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)", font=dict(color=C["txt2"], size=13),
        showlegend=leyenda, legend=dict(orientation="h", y=1.0, yanchor="bottom", x=0, font=dict(color=C["txt"])),
        hoverlabel=dict(font_size=13), bargap=0.35, bargroupgap=0.08,
    )
    fig.update_xaxes(gridcolor=C["grid"], zeroline=False, tickformat=",.0f", separatethousands=True)
    fig.update_yaxes(gridcolor="rgba(0,0,0,0)", automargin=True)
    fig.update_layout(separators=",.")
    return fig


st.markdown(f"""
<style>
.tarjetas {{ display:flex; gap:12px; flex-wrap:nowrap; align-items:stretch; margin:4px 0 18px; }}
.tarjetas .kpi {{ flex:1 1 0; min-width:0; height:auto; background:{C['card']};
  border:1px solid {C['borde']}; border-radius:10px; padding:clamp(10px,1vw,14px) clamp(10px,1.1vw,16px); }}
.kpi .t {{ font-size:clamp(11.5px,.85vw,13px); color:{C['txt2']}; margin-bottom:6px; }}
.kpi .v {{ font-size:clamp(20px,1.75vw,28px); font-weight:700; color:{C['txt']}; line-height:1.1; white-space:nowrap; }}
.kpi .p {{ font-size:clamp(11px,.8vw,12.5px); color:{C['txt2']}; margin-top:6px; }}
@media (max-width: 760px) {{ .tarjetas {{ flex-wrap:wrap; }} .tarjetas .kpi {{ min-width:150px; }} }}
.kpi.alerta {{ border-left:4px solid {C['alerta']}; }}
.kpi.ok {{ border-left:4px solid {C['ok']}; }}
[data-testid="stSidebarCollapseButton"], [data-testid="stExpandSidebarButton"] {{ opacity:1 !important; visibility:visible !important; }}
</style>""", unsafe_allow_html=True)


def tarjetas(items):
    html = "".join(f'<div class="kpi {c}"><div class="t">{t}</div><div class="v">{v}</div><div class="p">{p}</div></div>'
                   for t, v, p, c in items)
    st.markdown(f'<div class="tarjetas">{html}</div>', unsafe_allow_html=True)


# ─── Datos ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Datos")
    subida = st.file_uploader("Matriz de la empresa (.xlsx)", type="xlsx",
                              help="Si no subís nada se muestra el caso de ejemplo ficticio.")
    st.caption("Plantilla vacía para completar:")
    st.download_button("Descargar matriz vacía", (AQUI / "Matriz_GEI_Mendoza.xlsx").read_bytes(),
                       "Matriz_GEI_Mendoza.xlsx", width="stretch")


@st.cache_data(show_spinner="Calculando…")
def correr(contenido: bytes):
    with tempfile.TemporaryDirectory() as tmp:
        ruta = Path(tmp) / "matriz.xlsx"
        ruta.write_bytes(contenido)
        em, archivos = procesar(ruta, Path(tmp) / "salidas")
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for f in sorted((Path(tmp) / "salidas").iterdir()):
                z.write(f, f.name)
        from gei import matriz
        datos = matriz.leer(ruta)
    return em, datos, buf.getvalue(), archivos


contenido = subida.getvalue() if subida else EJEMPLO.read_bytes()
try:
    em, datos, zip_salidas, archivos = correr(contenido)
except SystemExit as e:
    st.error(str(e))
    st.stop()

empresa = datos["empresa"].get("Razón social", "—")
anios = [int(a) for a in sorted(em["Año"].unique())]

sectores = sorted(em["Sector"].unique())

# ─── Encabezado y filtros (en la página, no en la barra lateral: en el celular se ven) ──
st.title("Registro GEI Mendoza")
f1, f2 = st.columns([1, 2])
with f1:
    anio = st.segmented_control("Año", anios, default=anios[-1], key="anio") or anios[-1]
with f2:
    sector = st.segmented_control("Negocio", ["Todos"] + sectores, default="Todos", key="sector") or "Todos"

with st.sidebar:
    st.divider()
    st.download_button("⬇ Todo para presentar (zip)", zip_salidas, "GEI_Mendoza_presentacion.zip",
                       type="primary", width="stretch")
    n1 = sum(1 for n, _ in archivos if n.startswith("AnexoI_"))
    n2 = sum(1 for n, _ in archivos if n.startswith("AnexoII_"))
    st.caption(f"Incluye {n1} Anexo I oficiales (Excel), {n2} Anexo II precompletados (Word, lo amarillo lo completa "
               "la empresa), el Resumen Ejecutivo y el Informe GHG Protocol.")

base = em[~em["Informativo"]]
if sector != "Todos":
    base = base[base["Sector"] == sector]
d = base[base["Año"] == anio]
prev = base[base["Año"] == anio - 1]

st.markdown(f"**{empresa}** · {sector if sector != 'Todos' else 'Minería e Hidrocarburos'} · año **{anio}**"
            f"{'  ·  🧪 caso de ejemplo ficticio' if not subida else ''}")

total = d["tCO2e"].sum()
a1 = d[d["Alcance"] == 1]["tCO2e"].sum()
a2 = d[d["Alcance"] == 2]["tCO2e"].sum()
tot_prev = prev["tCO2e"].sum()
var = (total - tot_prev) / tot_prev * 100 if tot_prev else None

por_sector = d.groupby("Sector")["tCO2e"].sum()
superan = [s for s, v in por_sector.items() if v > UMBRAL_TCO2E]
tarjetas([
    ("Emisiones totales", f"{fmt(total)} t", "tCO2e/año · Alcance 1 + 2", ""),
    ("Alcance 1 · directas", f"{fmt(a1)} t", f"{fmt(a1 / total * 100 if total else 0)} % del total", ""),
    ("Alcance 2 · electricidad", f"{fmt(a2)} t", f"{fmt(a2 / total * 100 if total else 0)} % del total", ""),
    (f"Variación vs {anio - 1}", "—" if var is None else f"{'▲' if var > 0 else '▼'} {fmt(abs(var), 1)} %",
     f"{fmt(tot_prev)} t en {anio - 1}" if tot_prev else "sin año anterior cargado", ""),
    ("Umbral 10.000 t/año", "⚠ Supera" if superan else "✓ Debajo",
     ("Segregar por instalación y proceso: " + ", ".join(superan)) if superan
     else "Puede reportar agregado (N1/N2)", "alerta" if superan else "ok"),
])

# ─── 1. De dónde vienen ─────────────────────────────────────────────────────
izq, der = st.container(), st.container()
with izq:
    st.subheader("¿De dónde vienen las emisiones?")
    g = d.groupby(["Actividad", "Alcance"])["tCO2e"].sum().reset_index()
    g["Etiqueta"] = g["Actividad"].str.replace(r"\s*\(.*\)", "", regex=True).str.slice(0, 48)
    if MOVIL:   # en el celular: nombres cortos y solo el %, para que la barra tenga lugar
        g["Etiqueta"] = g["Etiqueta"].map(lambda t: t if len(t) <= 24 else t[:23].rstrip() + "…")
    g = g.groupby(["Etiqueta", "Alcance"])["tCO2e"].sum().reset_index().sort_values("tCO2e")
    g = g[g["tCO2e"] > 0.05]
    fig = go.Figure(go.Bar(
        x=g["tCO2e"], y=g["Etiqueta"], orientation="h",
        marker=dict(color=[C["a1"] if a == 1 else C["a2"] for a in g["Alcance"]], cornerradius=4),
        text=[(f"{fmt(v / total * 100, 1)} %" if MOVIL else f"{fmt(v)} t · {fmt(v / total * 100, 1)} %")
              for v in g["tCO2e"]], textposition="outside",
        textfont=dict(color=C["txt2"], size=11 if MOVIL else 12), cliponaxis=False,
        hovertemplate="<b>%{y}</b><br>%{x:,.1f} tCO2e<extra></extra>"))
    layout(fig, alto=max(300, 34 * len(g) + 40), leyenda=False)
    fig.update_xaxes(range=[0, g["tCO2e"].max() * (1.22 if MOVIL else 1.35)] if len(g) else None,
                     showticklabels=not MOVIL)
    if MOVIL:
        fig.update_yaxes(tickfont=dict(size=11))
    st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})
    st.caption(f"<span style='color:{C['a1']}'>■</span> Alcance 1 (directas)   "
               f"<span style='color:{C['a2']}'>■</span> Alcance 2 (electricidad comprada)", unsafe_allow_html=True)

with der:
    st.divider()
    st.subheader("Por instalación")
    g = d.groupby(["Instalación", "Alcance"])["tCO2e"].sum().unstack(fill_value=0)
    g = g.loc[g.sum(axis=1).sort_values().index]
    fig = go.Figure()
    for alc, col, nom in ((1, C["a1"], "Alcance 1"), (2, C["a2"], "Alcance 2")):
        if alc in g.columns:
            fig.add_bar(x=g[alc], y=g.index, orientation="h", name=nom,
                        marker=dict(color=col, cornerradius=4, line=dict(width=2, color="rgba(0,0,0,0)")),
                        hovertemplate=f"<b>%{{y}}</b><br>{nom}: %{{x:,.0f}} tCO2e<extra></extra>")
    layout(fig, alto=max(260, 70 * len(g) + 60))
    fig.update_layout(barmode="stack")
    st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})

    # Intensidad
    prod = datos["produccion"]
    prod = prod[prod["Año"] == anio]
    filas = []
    for _, p in prod.iterrows():
        e = em[(em["Año"] == anio) & (em["Instalación"] == p["Instalación"]) & (~em["Informativo"])]["tCO2e"].sum()
        if e and p["Cantidad"] and (sector == "Todos" or e and p["Instalación"] in d["Instalación"].values):
            filas.append((f"Intensidad · {p['Producto'].lower()}", f"{fmt(e * 1000 / p['Cantidad'], 0)} kg",
                          f"CO2e por {p['Unidad']} ({p['Instalación']})", ""))
    if filas:
        st.subheader("Intensidad")
        tarjetas(filas[:3])

# ─── 2. Evolución ───────────────────────────────────────────────────────────
st.divider()
st.subheader("Evolución entre períodos (Tabla 5 del Anexo II)")
if len(anios) >= 2:
    g = base.groupby(["Categoría", "Año"])["tCO2e"].sum().unstack(fill_value=0)
    g = g.loc[g[anios[-1]].sort_values().index]
    fig = go.Figure()
    for i, a in enumerate(anios[-2:]):
        fig.add_bar(y=g.index, x=g[a], orientation="h", name=str(a),
                    marker=dict(color=C["prev"] if i == 0 else C["a1"], cornerradius=4),
                    hovertemplate=f"<b>%{{y}}</b><br>{a}: %{{x:,.0f}} tCO2e<extra></extra>")
    layout(fig, alto=max(260, 60 * len(g) + 60))
    fig.update_layout(barmode="group")
    st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})
else:
    st.info("Cargá al menos dos años en la matriz para ver la evolución.")

# ─── 3. En contexto ─────────────────────────────────────────────────────────
st.divider()
st.subheader("La empresa en el contexto de Mendoza")
ctx = [("Del inventario provincial", f"{fmt(total / INV_PROVINCIAL_2023 * 100, 2)} %",
        f"{fmt(total)} t sobre {fmt(INV_PROVINCIAL_2023)} t (inventario provincial 2023)", "")]
for s, v in por_sector.items():
    nombre, ref = REGISTRO_2023[s]
    ctx.append((f"{s}: del sector {nombre} del Registro", f"{fmt(v / ref * 100, 2)} %",
                f"{fmt(v)} t sobre {fmt(ref)} t reportadas en el Registro Provincial 2023", ""))
tarjetas(ctx)
st.caption("Referencia orientativa: compara el año elegido con datos provinciales de 2023 (último publicado). "
           "Fuente: Gobierno de Mendoza, tablero del Inventario y Registro Provincial de GEI.")

# ─── 4. Informe para terceros ───────────────────────────────────────────────
st.divider()
st.subheader("Informe para terceros · GHG Protocol")
from gei import informe_ghg  # noqa: E402

req = informe_ghg.suficiencia(datos, em)
faltan_obl = [r for r in req if r[3] and not r[1]]
faltan_opc = [r for r in req if not r[3] and not r[1]]
ok_n = sum(1 for r in req if r[1])
tarjetas([
    ("Suficiencia de la información", "✓ Completo" if not faltan_obl else f"✗ Faltan {len(faltan_obl)}",
     "Se puede emitir el informe según GHG Protocol" if not faltan_obl
     else "Sale como BORRADOR hasta completar la hoja INFORME", "ok" if not faltan_obl else "alerta"),
    ("Requisitos cumplidos", f"{ok_n} de {len(req)}", "Capítulo 9 del Estándar Corporativo", ""),
    ("Opcionales pendientes", str(len(faltan_opc)), "Suman credibilidad, no bloquean", ""),
])
with st.expander("Ver requisito por requisito", expanded=bool(faltan_obl)):
    st.dataframe(pd.DataFrame([{"": "✓" if c else "✗", "Requisito": t, "Tipo": "Obligatorio" if o else "Opcional",
                                "Detalle": d} for t, c, d, o in req]),
                 width="stretch", hide_index=True)
nombre_inf = next((n for n in zipfile.ZipFile(io.BytesIO(zip_salidas)).namelist() if n.startswith("Informe_GEI")), None)
if nombre_inf:
    st.download_button(f"⬇ {'Informe' if not faltan_obl else 'Borrador de informe'} GHG Protocol (Word)",
                       zipfile.ZipFile(io.BytesIO(zip_salidas)).read(nombre_inf), nombre_inf)

# ─── 5. Detalle ─────────────────────────────────────────────────────────────
st.divider()
with st.expander("Detalle del cálculo (cada fila es lo que va a la Tabla 1 del Anexo II)"):
    t = d[["Instalación", "Proceso", "Categoría", "Gas", "Dato de actividad", "Factor", "Fuente del factor",
           "Fuente del dato", "Método", "tCO2e"]].sort_values("tCO2e", ascending=False)
    st.dataframe(t, width="stretch", hide_index=True,
                 column_config={"tCO2e": st.column_config.NumberColumn(format="%.1f")})
bio = em[(em["Año"] == anio) & em["Informativo"]]
if sector != "Todos":
    bio = bio[bio["Sector"] == sector]
if len(bio):
    st.caption(f"CO2 biogénico del corte de biocombustibles (se informa aparte, no suma): {fmt(bio['kg'].sum() / 1000, 1)} t.")
st.caption("Factores: SNIGEI-AR 2025 (Ambiente de Nación) e IPCC 2006 · PCG AR6 del template oficial (Res. 1/2026) · "
           "Herramienta de ESG Consulting.")
