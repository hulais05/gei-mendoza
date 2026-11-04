"""
Resumen ejecutivo (Word, 2 páginas): lo que lee un gerente o un directorio.

Formato ESG: Calibri, formato directo, navy 1F3864 / azul 2F5496 / grises 404040-595959.
"""

import re
from datetime import date

import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from .factores import FE_RED, FE_RED_ULTIMO_OFICIAL, UMBRAL_TCO2E

NAVY, AZUL, GRIS, GRIS2 = "1F3864", "2F5496", "404040", "595959"
INV_PROVINCIAL_2023 = 10_474_412.31
REGISTRO_2023 = {"Hidrocarburos": ("Energía", 4_056_976.90), "Minería": ("Industrial", 1_522_648.10)}


def _num(x, dec=0):
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _rgb(h):
    return RGBColor.from_string(h)


def _p(doc, texto="", tam=10.5, color=GRIS, negrita=False, antes=0, despues=6, alin=None, italica=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before, pf.space_after, pf.line_spacing = Pt(antes), Pt(despues), 1.1
    if alin:
        p.alignment = alin
    if texto:
        r = p.add_run(texto)
        r.font.name, r.font.size, r.font.color.rgb = "Calibri", Pt(tam), _rgb(color)
        r.bold, r.italic = negrita, italica
    return p


def _run(p, texto, tam=10.5, color=GRIS, negrita=False):
    r = p.add_run(texto)
    r.font.name, r.font.size, r.font.color.rgb, r.bold = "Calibri", Pt(tam), _rgb(color), negrita
    return r


def _titulo(doc, texto):
    _p(doc, texto, tam=13, color=NAVY, negrita=True, antes=12, despues=4)


def _sombrear(cell, hexcolor):
    el = OxmlElement("w:shd")
    el.set(qn("w:val"), "clear")
    el.set(qn("w:fill"), hexcolor)
    cell._tc.get_or_add_tcPr().append(el)


def _anchos(tb, cms):
    grid = tb._tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), cms):
        gc.set(qn("w:w"), str(int(w * 567)))
    tblPr = tb._tbl.tblPr
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)
    for row in tb.rows:
        for c, w in zip(row.cells, cms):
            c.width = Cm(w)


def _tabla(doc, cabecera, filas, cms, total=False):
    tb = doc.add_table(rows=1 + len(filas), cols=len(cabecera))
    tb.style = "Table Grid"
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    for k, h in enumerate(cabecera):
        c = tb.rows[0].cells[k]
        _sombrear(c, NAVY)
        p = c.paragraphs[0]
        _run(p, h, tam=9, color="FFFFFF", negrita=True)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if k else WD_ALIGN_PARAGRAPH.LEFT
    for i, fila in enumerate(filas, 1):
        ultima = total and i == len(filas)
        for k, v in enumerate(fila):
            c = tb.rows[i].cells[k]
            if ultima:
                _sombrear(c, "DEEAF6")
            p = c.paragraphs[0]
            _run(p, str(v), tam=9, color=GRIS, negrita=ultima)
            numerico = re.fullmatch(r"[-+▲▼]?\s?[\d.,]+\s?(%|t|kg.*)?", str(v).strip()) is not None
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if (k and numerico) else (
                WD_ALIGN_PARAGRAPH.CENTER if str(v) in ("✓", "✗") else WD_ALIGN_PARAGRAPH.LEFT)
    _anchos(tb, cms)
    _p(doc, despues=2)
    return tb


def _caja(doc, titulo, cuerpo, fondo="DEEAF6"):
    tb = doc.add_table(rows=1, cols=1)
    tb.style = "Table Grid"
    c = tb.rows[0].cells[0]
    _sombrear(c, fondo)
    p = c.paragraphs[0]
    p.paragraph_format.space_after = Pt(2)
    _run(p, titulo, tam=10.5, color=NAVY, negrita=True)
    for linea in cuerpo:
        q = c.add_paragraph()
        q.paragraph_format.space_after = Pt(2)
        _run(q, linea, tam=10, color=GRIS)
    _anchos(tb, [17])
    _p(doc, despues=2)


def exportar(em, datos, salida, archivos):
    empresa = datos["empresa"]
    e = em[~em["Informativo"]]
    anios = sorted(e["Año"].unique())
    ult, ant = anios[-1], (anios[-2] if len(anios) > 1 else None)
    tot = {a: e[e["Año"] == a]["tCO2e"].sum() for a in anios}

    doc = docx.Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = Cm(2.0)
    sec.top_margin = sec.bottom_margin = Cm(1.8)
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Calibri", Pt(10.5)

    _p(doc, "RESUMEN EJECUTIVO", tam=10, color=AZUL, negrita=True, despues=0)
    _p(doc, "Inventario de Gases de Efecto Invernadero", tam=20, color=NAVY, negrita=True, despues=2)
    _p(doc, f"{empresa.get('Razón social', '')}  ·  Períodos {', '.join(map(str, anios))}  ·  "
            f"Registro Provincial de Emisiones GEI de Mendoza (Res. 1/2026)", tam=10, color=GRIS2, despues=2)
    _p(doc, f"Fecha de emisión: {date.today().strftime('%d/%m/%Y')}", tam=9, color=GRIS2, despues=10)

    # Resultado principal
    var = (tot[ult] - tot[ant]) / tot[ant] * 100 if ant else None
    a1 = e[(e["Año"] == ult) & (e["Alcance"] == 1)]["tCO2e"].sum()
    lineas = [f"En {ult} la empresa emitió {_num(tot[ult])} tCO2e (Alcance 1 + 2): "
              f"{_num(a1)} t directas ({_num(a1 / tot[ult] * 100)} %) y {_num(tot[ult] - a1)} t por electricidad comprada."]
    if ant:
        lineas.append(f"Respecto de {ant} ({_num(tot[ant])} t), las emisiones {'subieron' if var > 0 else 'bajaron'} "
                      f"{_num(abs(var), 1)} %.")
    superan = [s for s, v in e[e["Año"] == ult].groupby("Sector")["tCO2e"].sum().items() if v > UMBRAL_TCO2E]
    if superan:
        lineas.append(f"{' y '.join(superan)} supera{'n' if len(superan) > 1 else ''} el umbral de 10.000 tCO2e/año: "
                      "el reporte debe segregarse por instalación y proceso (así se generó).")
    _caja(doc, "Resultado principal", lineas)

    # Por negocio y año
    _titulo(doc, "Emisiones por negocio y año (tCO2e)")
    filas = []
    for s in sorted(e["Sector"].unique()):
        for a in anios:
            d = e[(e["Sector"] == s) & (e["Año"] == a)]
            t = d["tCO2e"].sum()
            filas.append([s, a, _num(d[d["Alcance"] == 1]["tCO2e"].sum()), _num(d[d["Alcance"] == 2]["tCO2e"].sum()),
                          _num(t), "Supera" if t > UMBRAL_TCO2E else "Debajo"])
    filas.append(["Total empresa", ult, _num(e[(e["Año"] == ult) & (e["Alcance"] == 1)]["tCO2e"].sum()),
                  _num(e[(e["Año"] == ult) & (e["Alcance"] == 2)]["tCO2e"].sum()), _num(tot[ult]), ""])
    _tabla(doc, ["Negocio", "Año", "Alcance 1", "Alcance 2", "Total", "Umbral 10.000 t"], filas,
           [3.6, 1.6, 2.9, 2.9, 3.0, 3.0], total=True)

    # Principales fuentes
    _titulo(doc, f"Dónde están las emisiones ({ult})")
    g = e[e["Año"] == ult].groupby("Actividad")["tCO2e"].sum().sort_values(ascending=False)
    top = g.head(5)
    filas = [[a.split(" (")[0], _num(v), f"{_num(v / tot[ult] * 100, 1)} %"] for a, v in top.items()]
    _tabla(doc, ["Fuente", "tCO2e", "% del total"], filas, [11.0, 3.0, 3.0])
    _p(doc, f"Las cinco primeras fuentes explican el {_num(top.sum() / tot[ult] * 100)} % del total: "
            "es donde una medida de reducción mueve la aguja.", tam=9.5, color=GRIS2, italica=True)

    # Intensidad y contexto
    _titulo(doc, "Intensidad y contexto provincial")
    prod = datos["produccion"]
    filas = []
    for _, p in prod[prod["Año"] == ult].iterrows():
        v = e[(e["Año"] == ult) & (e["Instalación"] == p["Instalación"])]["tCO2e"].sum()
        if v and p["Cantidad"]:
            fila = [f"{p['Producto']} ({p['Instalación']})", f"{_num(v * 1000 / p['Cantidad'])} kg CO2e/{p['Unidad']}"]
            q = prod[(prod["Año"] == ant) & (prod["Instalación"] == p["Instalación"])] if ant else None
            if q is not None and len(q):
                v0 = e[(e["Año"] == ant) & (e["Instalación"] == p["Instalación"])]["tCO2e"].sum()
                fila.append(f"{_num(v0 * 1000 / q['Cantidad'].iloc[0])} kg en {ant}")
            else:
                fila.append("—")
            filas.append(fila)
    if filas:
        _tabla(doc, ["Producto (instalación)", f"Intensidad {ult}", "Año anterior"], filas, [8.0, 4.5, 4.5])
    pr = _p(doc, despues=2)
    _run(pr, "En Mendoza: ", negrita=True, color=NAVY)
    _run(pr, f"la empresa equivale al {_num(tot[ult] / INV_PROVINCIAL_2023 * 100, 2)} % del inventario provincial 2023 "
             f"({_num(INV_PROVINCIAL_2023)} t)")
    for s, v in e[e["Año"] == ult].groupby("Sector")["tCO2e"].sum().items():
        nom, ref = REGISTRO_2023[s]
        _run(pr, f"; {s.lower()}: {_num(v / ref * 100, 2)} % del sector {nom} del Registro 2023")
    _run(pr, ". Referencia orientativa (los datos provinciales son de 2023).")

    # Presentación
    _titulo(doc, "Qué se presenta y cuándo")
    _p(doc, "Vencimiento: 30/11/2026 (períodos 2024 y 2025 en forma conjunta, Res. 1/2026), por el Sistema de "
            "Tickets del Gobierno de Mendoza (MEYA_SOSTENIBILIDAD_GEI). Firma de profesional competente con título "
            "habilitante adjunto.", despues=4)
    for nombre in archivos:
        _p(doc, f"•  {nombre}", tam=9.5, despues=1)

    # Metodología
    _titulo(doc, "Metodología y supuestos")
    fe = FE_RED[FE_RED_ULTIMO_OFICIAL]["CO2"]
    for t in [
        "Límite organizacional: control operacional (GHG Protocol, Estándar Corporativo). Alcances 1 y 2 al 100 %; "
        "Alcance 3 no reportado (opcional).",
        "Factores de emisión: SNIGEI-AR 2025 (Ambiente de Nación) e IPCC 2006. Potenciales de calentamiento AR6 "
        "(GWP100) idénticos a los del template oficial.",
        f"Electricidad: factor de red {_num(fe, 3)} kg CO2/kWh ({FE_RED_ULTIMO_OFICIAL}, último oficial publicado). "
        "Reemplazar cuando se publique el del año del inventario.",
        "CO2 biogénico del corte de biocombustibles: se informa aparte y no suma al total.",
    ]:
        _p(doc, f"•  {t}", tam=9.5, despues=2)

    _caja(doc, "Antes de presentar", [
        "1. Revisión y firma del profesional competente.",
        "2. Completar en el Anexo II lo marcado en amarillo (exclusiones, programas, verificación, compensaciones).",
        "3. Confirmar el factor de red del año y la clasificación de la actividad con la autoridad.",
    ], fondo="FFF2CC")
    _p(doc, "Elaborado por ESG Consulting S.A.S. a partir de los datos de actividad suministrados por la empresa.",
       tam=8.5, color=GRIS2, italica=True, antes=6)
    doc.save(salida)
