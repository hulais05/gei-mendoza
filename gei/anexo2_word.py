"""
Anexo II oficial (ANEXO-II-Metodologia-.docx) precompletado.

Se abre el Word del gobierno y se completa lo que la herramienta ya sabe:
encabezado, límites, año base, Tabla 1 por año y Tabla 5. Lo que solo puede
escribir la empresa (exclusiones, programas, verificación, compensaciones)
queda marcado en AMARILLO con "A COMPLETAR" para que nadie lo presente vacío.

Un Anexo II por inventario (industria × departamento), igual que el Anexo I.
"""

import copy

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor

NEGRO = RGBColor(0, 0, 0)
ORDEN_CAT = ["Combustión fija", "Combustión móvil", "Emisiones de proceso", "Emisiones fugitivas",
             "Electricidad adquirida"]
PENDIENTE = "A COMPLETAR"


def _num(x, dec=1):
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _escribir(par, texto, pendiente=False):
    """Reemplaza el texto del párrafo conservando el formato del primer run."""
    runs = par.runs
    if not runs:
        r = par.add_run(texto)
    else:
        r = runs[0]
        r.text = texto
        for extra in runs[1:]:
            extra._r.getparent().remove(extra._r)
    r.italic = False
    r.font.color.rgb = NEGRO
    if pendiente:
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
        r.bold = True


def _celda(cell, texto, negrita=False, pendiente=False):
    pars = cell.paragraphs
    for p in pars[1:]:
        p._p.getparent().remove(p._p)
    _escribir(pars[0], texto, pendiente)
    pars[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    for r in pars[0].runs:
        r.bold = negrita or pendiente
        r.font.size = docx.shared.Pt(8)


def _borrar(el):
    el.getparent().remove(el)


def _par_que_empieza(doc, prefijo):
    for p in doc.paragraphs:
        if p.text.strip().startswith(prefijo):
            return p
    raise KeyError(prefijo)


def _responder(doc, pregunta, texto, pendiente=False, borrar_ejemplo=True):
    """Completa el primer '_________' que sigue a la pregunta; borra el 'Ejemplo:' intermedio."""
    pars = doc.paragraphs
    i = next(k for k, p in enumerate(pars) if p.text.strip().startswith(pregunta))
    for p in pars[i + 1:]:
        t = p.text.strip()
        if t.startswith("Ejemplo") and borrar_ejemplo:
            _borrar(p._p)
            continue
        if set(t) == {"_"}:
            _escribir(p, texto, pendiente)
            return
    raise KeyError(pregunta)


def _sin_alto(tr):
    for h in tr.findall(qn("w:trPr") + "/" + qn("w:trHeight")):
        h.getparent().remove(h)


def _filas_por_fuente(em_anio):
    """Una fila por fuente (actividad), con los gases juntos: así lo muestra el ejemplo del gobierno."""
    em = em_anio[~em_anio["Informativo"]].copy()
    claves = ["Alcance", "Categoría", "Instalación", "Proceso", "Código"]
    g = em.groupby(claves, sort=False).agg(
        Gas=("Gas", lambda x: " · ".join(dict.fromkeys(x))), tCO2e=("tCO2e", "sum"),
        Factor=("Factor", "first"), FuenteFE=("Fuente del factor", "first"),
        Dato=("Dato de actividad", "first"), FuenteDato=("Fuente del dato", "first")).reset_index()
    g["_o"] = g["Categoría"].map({c: i for i, c in enumerate(ORDEN_CAT)})
    return g.sort_values(["Alcance", "_o", "Instalación", "Proceso"])


def _tabla1(tabla, em_anio):
    filas = tabla.rows
    trpr = filas[0]._tr.get_or_add_trPr()
    if trpr.find(qn("w:tblHeader")) is None:
        trpr.append(OxmlElement("w:tblHeader"))      # el encabezado se repite en cada página
    hdr_a1, molde, hdr_a2 = filas[1]._tr, filas[2]._tr, filas[3]._tr
    molde_a2 = filas[4]._tr
    g = _filas_por_fuente(em_anio)

    def agregar(despues_de, base, vals, negrita=False):
        tr = copy.deepcopy(base)
        _sin_alto(tr)
        despues_de.addnext(tr)
        celdas = [docx.table._Cell(tc, tabla) for tc in tr.findall(qn("w:tc"))]
        for c, v in zip(celdas, vals):
            _celda(c, str(v), negrita=negrita)
        return tr

    def vals(r):
        return [r["Categoría"], r["Instalación"], r["Proceso"], r["Gas"], r["Factor"], r["FuenteFE"],
                r["Dato"], r["FuenteDato"] or "", _num(r["tCO2e"], 2)]

    ultimo = hdr_a1
    for _, r in g[g["Alcance"] == 1].iterrows():
        ultimo = agregar(ultimo, molde, vals(r))
    ultimo = hdr_a2
    for _, r in g[g["Alcance"] == 2].iterrows():
        ultimo = agregar(ultimo, molde_a2, vals(r))
    agregar(ultimo, molde_a2, ["TOTAL (Alcance 1 + 2)"] + [""] * 7 + [_num(g["tCO2e"].sum(), 2)], negrita=True)
    _borrar(molde)
    _borrar(molde_a2)


def _corte_seccion(despues_de_el, sectpr_base, apaisada):
    """Inserta un párrafo vacío que cierra una sección (portrait o landscape) con los mismos encabezados."""
    sp = copy.deepcopy(sectpr_base)
    pg = sp.find(qn("w:pgSz"))
    w, h = int(pg.get(qn("w:w"))), int(pg.get(qn("w:h")))
    if apaisada:
        pg.set(qn("w:w"), str(max(w, h)))
        pg.set(qn("w:h"), str(min(w, h)))
        pg.set(qn("w:orient"), "landscape")
        mar = sp.find(qn("w:pgMar"))
        for lado in ("w:left", "w:right"):
            mar.set(qn(lado), "1134")          # 2 cm
    else:
        pg.set(qn("w:w"), str(min(w, h)))
        pg.set(qn("w:h"), str(max(w, h)))
        if pg.get(qn("w:orient")):
            del pg.attrib[qn("w:orient")]
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    ppr.append(sp)
    p.append(ppr)
    despues_de_el.addnext(p)
    return p


def exportar(plantilla, salida, empresa, instalaciones, em):
    """em: emisiones de UN inventario (una industria y un departamento), todos los años."""
    doc = docx.Document(plantilla)
    anios = sorted(em["Año"].unique())
    inst = instalaciones[instalaciones["Instalación"].isin(em["Instalación"].unique())]

    # Encabezado
    firma = empresa.get("Profesional que firma")
    _escribir(_par_que_empieza(doc, "EMPRESA:"), f"EMPRESA: {empresa.get('Razón social', '')}")
    p = _par_que_empieza(doc, "RESPONSABLE:")
    _escribir(p, f"RESPONSABLE: {firma}" if firma and "completar" not in str(firma) else f"RESPONSABLE: {PENDIENTE}",
              pendiente=not firma or "completar" in str(firma))
    _escribir(_par_que_empieza(doc, "CARGO:"), f"CARGO: {PENDIENTE}", pendiente=True)

    # 1. Límites organizacionales
    _responder(doc, "1.1.", "Control operacional (GHG Protocol, Estándar Corporativo), según Res. 1/2026.")
    _responder(doc, "1.2.", "; ".join(f"{r['Instalación']} ({r['Departamento']}): {r['Descripción'] or ''}".strip()
                                    for _, r in inst.iterrows()))
    _responder(doc, "1.3.", f"{PENDIENTE}: indicar instalaciones excluidas y su justificación, o «Ninguna».",
               pendiente=True)
    # 2. Límites operacionales
    _responder(doc, "2.1.", f"{PENDIENTE}: fuentes de Alcance 1 excluidas y justificación, o «Ninguna». "
                            "El Alcance 3 no se reporta (opcional según Res. 1/2026).", pendiente=True)
    _responder(doc, "2.2.", f"{PENDIENTE}: fuentes de Alcance 2 excluidas, o «Ninguna».", pendiente=True)
    # 3. Año base
    _responder(doc, "Instrucción: Indicar el año base",
               f"{anios[0]}: primer año reportado por la empresa ante el Registro (criterio del Anexo II).")

    # 4. Tabla 1: una por año (se copia la tabla vacía del gobierno antes de llenarla)
    periodo = _par_que_empieza(doc, "Periodo:")
    tabla_vacia = doc.tables[0]._tbl
    molde_par, molde_tbl = copy.deepcopy(periodo._p), copy.deepcopy(tabla_vacia)
    ancla = tabla_vacia
    for k, anio in enumerate(anios):
        if k == 0:
            par_el, tbl_el = periodo._p, tabla_vacia
        else:
            par_el, tbl_el = copy.deepcopy(molde_par), copy.deepcopy(molde_tbl)
            vacio = copy.deepcopy(molde_par)
            for t in vacio.iter(qn("w:t")):
                t.text = ""
            ancla.addnext(vacio)
            vacio.addnext(par_el)
            par_el.addnext(tbl_el)
        _escribir(docx.text.paragraph.Paragraph(par_el, periodo._parent), f"Periodo: 01/01/{anio} al 31/12/{anio}")
        _tabla1(docx.table.Table(tbl_el, periodo._parent), em[em["Año"] == anio])
        ancla = tbl_el

    # El template ya pone la Tabla 1 en una sección apaisada. Su tabla es "flotante"
    # (tblpPr): se le quita para que fluya con el texto y no salte de página sola.
    for t in doc.tables[:len(anios)]:
        for pp in t._tbl.tblPr.findall(qn("w:tblpPr")):
            t._tbl.tblPr.remove(pp)

    # Tabla 2 (ejemplo del gobierno): se quita del documento final
    t2 = _par_que_empieza(doc, "Tabla 2. Registro de Emisiones")
    sig = t2._p.getnext()
    while sig is not None and sig.tag != qn("w:tbl"):
        nxt = sig.getnext()
        _borrar(sig)
        sig = nxt
    if sig is not None:
        _borrar(sig)
    _borrar(t2._p)

    # Tabla 3 (programas), verificación y Tabla 4 (compensaciones): a completar por la empresa
    t3 = next(t for t in doc.tables if t.rows[0].cells[0].text.startswith("Nombre del program"))
    for k, c in enumerate(t3.rows[1].cells):
        _celda(c, f"{PENDIENTE}: programas de reducción, o «Sin programas»" if k == 0 else "", pendiente=k == 0)
    _responder(doc, "6.1.", f"{PENDIENTE} (no obligatorio): entidad verificadora, o «No se realizó verificación externa».",
               pendiente=True)
    _responder(doc, "6.2.", f"{PENDIENTE} (no obligatorio).", pendiente=True)
    t4 = next(t for t in doc.tables if t.rows[0].cells[0].text.startswith("Nombre / Tipo"))
    for k, c in enumerate(t4.rows[1].cells):
        _celda(c, f"{PENDIENTE}, o «No se informan compensaciones»" if k == 0 else "", pendiente=k == 0)

    # Tabla 5: seguimiento por categoría
    t5 = next(t for t in doc.tables if t.rows[0].cells[0].text.startswith("Categorías") and len(t.columns) == 5)
    cols = (list(anios) + [None] * 4)[:4]
    for k, a in enumerate(cols, 1):
        _celda(t5.rows[0].cells[k], f"Valor tCO2e/año Periodo {a}" if a else "Valor tCO2e/año Periodo —", negrita=True)
    d = em[~em["Informativo"]]
    piv = d.pivot_table(index="Categoría", columns="Año", values="tCO2e", aggfunc="sum", fill_value=0)
    cats = [c for c in ORDEN_CAT if c in piv.index] + ["TOTAL"]
    molde5 = t5.rows[2]._tr
    ultimo = t5.rows[1]._tr
    for cat in cats:
        tr = copy.deepcopy(molde5)
        ultimo.addnext(tr)
        ultimo = tr
        celdas = [docx.table._Cell(tc, t5) for tc in tr.findall(qn("w:tc"))]
        _celda(celdas[0], cat, negrita=cat == "TOTAL")
        for k, a in enumerate(cols, 1):
            v = "" if a is None else _num(piv[a].sum() if cat == "TOTAL" else piv.loc[cat, a])
            _celda(celdas[k], v, negrita=cat == "TOTAL")
    _borrar(t5.rows[1]._tr)
    _borrar(molde5)

    # Restos de ejemplos sueltos en la sección de mitigación
    for p in list(doc.paragraphs):
        if p.text.strip().startswith("Ejemplo"):
            _borrar(p._p)

    doc.save(salida)
