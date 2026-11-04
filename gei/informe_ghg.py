"""
Informe de Inventario GEI para terceros según el GHG Protocol
(Estándar Corporativo de Contabilidad y Reporte, capítulo 9 "Reporte de emisiones").

Dos piezas:
  suficiencia()  revisa, requisito por requisito del cap. 9, si la matriz alcanza.
  exportar()     arma el informe en Word. Si falta algún requisito OBLIGATORIO,
                 sale como BORRADOR con los huecos en amarillo y no declara
                 conformidad con el GHG Protocol.
"""

from datetime import date

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.shared import Cm, Pt

from .factores import FE_RED, FE_RED_ULTIMO_OFICIAL, PCG_TEMPLATE, POR_CODIGO
from .resumen import AZUL, GRIS, GRIS2, NAVY, _caja, _num, _p, _run, _tabla, _titulo

GRUPO_GAS = {"CO2": "CO2", "CH4 fósil combustión": "CH4", "CH4 fósil venteado": "CH4", "CH4 orgánico": "CH4",
             "N2O": "N2O", "SF6": "SF6"}
GRUPOS = ["CO2", "CH4", "N2O", "HFCs", "PFCs", "SF6"]
CALIDAD = {"MD": "Alta (medición directa)", "MI": "Alta (medición indirecta)", "N3": "Alta (estimación mejorada)",
           "N2": "Media (estimación intermedia)", "N1": "Básica (estimación básica)"}


def _grupo(gas):
    return GRUPO_GAS.get(gas, "HFCs" if gas.startswith("HFC") else "PFCs")


def _vacio(v):
    return v is None or str(v).strip() == ""


def suficiencia(datos, em):
    """Lista de (requisito, cumple, detalle, obligatorio) según GHG Protocol cap. 9."""
    inf = datos.get("informe", {})
    e = em[~em["Informativo"]]
    anios = sorted(e["Año"].unique())
    req = []

    def campo(nombre, requisito, obligatorio=True):
        v = inf.get(nombre)
        req.append((requisito, not _vacio(v), "Completar en la hoja INFORME: " + nombre if _vacio(v) else str(v)[:90],
                    obligatorio))

    req.append(("Límites organizacionales y enfoque de consolidación", True, "Control operacional", True))
    req.append(("Instalaciones incluidas", len(datos["instalaciones"]) > 0,
                f"{len(datos['instalaciones'])} instalaciones", True))
    campo("Instalaciones excluidas y justificación", "Exclusiones de instalaciones (o «Ninguna»)")
    campo("Fuentes excluidas y justificación", "Exclusiones de fuentes (o «Ninguna»)")
    campo("Alcance 3", "Límites operacionales: tratamiento del Alcance 3")
    req.append(("Período de reporte", len(anios) > 0, ", ".join(map(str, anios)), True))
    req.append(("Emisiones totales de Alcance 1 y 2, por separado", e["Alcance"].isin([1, 2]).any(), "Calculadas", True))
    req.append(("Emisiones por gas (CO2, CH4, N2O, HFCs, PFCs, SF6) en t y tCO2e", True, "Calculadas", True))
    req.append(("CO2 biogénico informado aparte", True, "Calculado", True))
    req.append(("Metodologías y factores con su fuente", True, "Catálogo de factores documentado", True))
    if _vacio(inf.get("Año base")):
        req.append(("Año base elegido", bool(anios), f"{anios[0]} (primer año cargado)" if anios else "—", True))
    else:
        campo("Año base", "Año base elegido")
    campo("Justificación del año base", "Justificación del año base")
    campo("Política de recálculo del año base", "Política de recálculo del año base")
    campo("Cambios estructurales en el período", "Cambios estructurales que obliguen a recalcular (o «Ninguno»)")
    req.append(("Persona de contacto", not _vacio(inf.get("Persona de contacto")),
                inf.get("Persona de contacto") or "Completar en la hoja INFORME", False))
    campo("Causas de los cambios entre años", "Causas de los cambios entre años", obligatorio=len(anios) > 1)
    campo("Programas de reducción", "Programas o estrategias de reducción", obligatorio=False)
    campo("Verificación externa", "Verificación externa", obligatorio=False)
    campo("Compensaciones", "Compensaciones", obligatorio=False)
    req.append(("Indicadores de intensidad", len(datos["produccion"]) > 0,
                "Con datos de PRODUCCION" if len(datos["produccion"]) else "Cargar la hoja PRODUCCION", False))
    return req


def completo(req):
    return all(ok for _, ok, _, obl in req if obl)


def _texto_o_pendiente(doc, valor, que):
    if _vacio(valor):
        p = _p(doc, despues=4)
        r = _run(p, f"A COMPLETAR: {que}", negrita=True)
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    else:
        _p(doc, str(valor), despues=4)


def exportar(em, datos, req, salida):
    empresa = datos["empresa"]
    inf = datos.get("informe", {})
    e = em[~em["Informativo"]].copy()
    e["Grupo"] = e["Gas"].map(_grupo)
    anios = sorted(e["Año"].unique())
    ult = anios[-1]
    ok = completo(req)
    anio_base = int(inf.get("Año base") or anios[0])

    doc = docx.Document()
    sec = doc.sections[0]
    sec.page_height, sec.page_width = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = Cm(2.2)
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    doc.styles["Normal"].font.name = "Calibri"

    # Portada
    _p(doc, despues=60)
    if not ok:
        p = _p(doc, despues=8)
        r = _run(p, "BORRADOR — faltan requisitos obligatorios del GHG Protocol (ver Anexo)", tam=11, negrita=True)
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    _p(doc, "INFORME DE INVENTARIO DE GASES DE EFECTO INVERNADERO", tam=11, color=AZUL, negrita=True, despues=6)
    _p(doc, empresa.get("Razón social", ""), tam=24, color=NAVY, negrita=True, despues=6)
    _p(doc, f"Períodos {', '.join(map(str, anios))}  ·  Alcances 1 y 2", tam=13, color=GRIS, despues=4)
    _p(doc, "Elaborado según el GHG Protocol — Estándar Corporativo de Contabilidad y Reporte (WRI/WBCSD)",
       tam=10.5, color=GRIS2, despues=30)
    if not _vacio(inf.get("Destinatario del informe")):
        _p(doc, f"Preparado para: {inf['Destinatario del informe']}", tam=10.5, color=GRIS, despues=4)
    _p(doc, f"Fecha: {date.today().strftime('%d/%m/%Y')}", tam=10, color=GRIS2, despues=4)
    _p(doc, "Elaborado por ESG Consulting S.A.S. con datos de actividad suministrados por la empresa.",
       tam=9.5, color=GRIS2, italica=True)
    doc.add_page_break()

    tot = {a: e[e["Año"] == a]["tCO2e"].sum() for a in anios}

    # 1. Resumen
    _titulo(doc, "1. Resumen")
    a1 = e[(e["Año"] == ult) & (e["Alcance"] == 1)]["tCO2e"].sum()
    lineas = [f"Emisiones {ult}: {_num(tot[ult])} tCO2e (Alcance 1: {_num(a1)} t; Alcance 2: {_num(tot[ult] - a1)} t)."]
    if anio_base in tot and anio_base != ult:
        v = (tot[ult] - tot[anio_base]) / tot[anio_base] * 100
        lineas.append(f"Respecto del año base {anio_base} ({_num(tot[anio_base])} t): {'+' if v > 0 else ''}{_num(v, 1)} %.")
    top = e[e["Año"] == ult].groupby("Actividad")["tCO2e"].sum().sort_values(ascending=False).head(3)
    lineas.append("Principales fuentes: " + "; ".join(f"{a.split(' (')[0]} ({_num(v / tot[ult] * 100)} %)"
                                                        for a, v in top.items()) + ".")
    _caja(doc, "Resultados principales", lineas)

    # 2. Empresa y límites
    _titulo(doc, "2. Descripción de la empresa y límites del inventario")
    _texto_o_pendiente(doc, inf.get("Descripción de la empresa"), "descripción de la empresa")
    _p(doc, "Límite organizacional: enfoque de control operacional. Se incluye el 100 % de las emisiones de las "
            "instalaciones sobre las que la empresa tiene control operacional.", despues=4)
    filas = [[r["Instalación"], r["Departamento"], r["Descripción"] or ""] for _, r in datos["instalaciones"].iterrows()]
    _tabla(doc, ["Instalación incluida", "Ubicación", "Descripción"], filas, [5.5, 3.0, 8.1])
    _p(doc, "Instalaciones excluidas:", negrita=True, color=NAVY, despues=2)
    _texto_o_pendiente(doc, inf.get("Instalaciones excluidas y justificación"), "instalaciones excluidas o «Ninguna»")
    _p(doc, "Límite operacional: Alcance 1 (emisiones directas) y Alcance 2 (electricidad comprada, método basado "
            "en la ubicación).", despues=2)
    _p(doc, "Alcance 3:", negrita=True, color=NAVY, despues=2)
    _texto_o_pendiente(doc, inf.get("Alcance 3"), "tratamiento del Alcance 3")
    _p(doc, "Fuentes excluidas:", negrita=True, color=NAVY, despues=2)
    _texto_o_pendiente(doc, inf.get("Fuentes excluidas y justificación"), "fuentes excluidas o «Ninguna»")
    _p(doc, f"Período de reporte: años calendario {', '.join(map(str, anios))} (1 de enero al 31 de diciembre).")

    # 3. Emisiones
    _titulo(doc, "3. Emisiones por alcance")
    filas = []
    for a in anios:
        d = e[e["Año"] == a]
        filas.append([a, _num(d[d["Alcance"] == 1]["tCO2e"].sum()), _num(d[d["Alcance"] == 2]["tCO2e"].sum()),
                      _num(tot[a])])
    _tabla(doc, ["Año", "Alcance 1 (tCO2e)", "Alcance 2 (tCO2e)", "Total (tCO2e)"], filas, [3.0, 4.5, 4.5, 4.6])

    _titulo(doc, "4. Emisiones por gas")
    for a in anios:
        d = e[e["Año"] == a]
        g = d.groupby("Grupo").agg(kg=("kg", "sum"), t=("tCO2e", "sum"))
        filas = [[gr, _num(g.loc[gr, "kg"] / 1000, 2) if gr in g.index else "0",
                  _num(g.loc[gr, "t"], 1) if gr in g.index else "0"] for gr in GRUPOS]
        filas.append(["Total", "", _num(d["tCO2e"].sum(), 1)])
        _p(doc, f"Año {a}", negrita=True, color=NAVY, despues=2)
        _tabla(doc, ["Gas", "t de gas", "tCO2e"], filas, [5.0, 5.8, 5.8], total=True)
    bio = em[em["Informativo"]].groupby("Año")["kg"].sum() / 1000
    _p(doc, "CO2 biogénico (informado aparte, fuera de los alcances): " +
       "; ".join(f"{a}: {_num(bio.get(a, 0), 1)} t" for a in anios) +
       ". Proviene del biodiésel y bioetanol incluidos por ley en el gasoil y la nafta.", tam=9.5, despues=6)

    _titulo(doc, f"5. Emisiones por categoría e instalación ({ult})")
    d = e[e["Año"] == ult]
    g = d.groupby(["Instalación", "Categoría"])["tCO2e"].sum().reset_index().sort_values(["Instalación", "tCO2e"],
                                                                                        ascending=[True, False])
    filas = [[r["Instalación"], r["Categoría"], _num(r["tCO2e"], 1), f"{_num(r['tCO2e'] / tot[ult] * 100, 1)} %"]
             for _, r in g.iterrows()]
    _tabla(doc, ["Instalación", "Categoría", "tCO2e", "%"], filas, [6.0, 5.0, 3.0, 2.6])

    # 6. Año base
    _titulo(doc, "6. Año base y evolución")
    _p(doc, f"Año base: {anio_base}.", negrita=True, despues=2)
    _texto_o_pendiente(doc, inf.get("Justificación del año base"), "justificación del año base")
    _p(doc, "Política de recálculo:", negrita=True, color=NAVY, despues=2)
    _texto_o_pendiente(doc, inf.get("Política de recálculo del año base"), "política de recálculo del año base")
    _p(doc, "Cambios estructurales en el período:", negrita=True, color=NAVY, despues=2)
    _texto_o_pendiente(doc, inf.get("Cambios estructurales en el período"), "cambios estructurales o «Ninguno»")
    if len(anios) > 1:
        _p(doc, "Causas de los cambios entre años:", negrita=True, color=NAVY, despues=2)
        _texto_o_pendiente(doc, inf.get("Causas de los cambios entre años"), "qué explica la variación")

    # 7. Intensidad
    prod = datos["produccion"]
    if len(prod):
        _titulo(doc, "7. Indicadores de intensidad")
        filas = []
        for _, p in prod.sort_values(["Instalación", "Año"]).iterrows():
            v = e[(e["Año"] == p["Año"]) & (e["Instalación"] == p["Instalación"])]["tCO2e"].sum()
            if v and p["Cantidad"]:
                filas.append([f"{p['Producto']} ({p['Instalación']})", p["Año"],
                              f"{_num(p['Cantidad'])} {p['Unidad']}", f"{_num(v * 1000 / p['Cantidad'])} kg CO2e/{p['Unidad']}"])
        _tabla(doc, ["Producto", "Año", "Producción", "Intensidad"], filas, [7.0, 2.0, 3.6, 4.0])

    # 8. Metodología
    _titulo(doc, "8. Metodología, factores de emisión y potenciales de calentamiento")
    _p(doc, "Emisiones = dato de actividad × factor de emisión, por gas; luego × potencial de calentamiento global "
            "(GWP100, IPCC AR6). Factores prioritariamente nacionales (SNIGEI-AR 2025, Ambiente de Nación), "
            "y del IPCC 2006 donde no hay factor nacional.", despues=4)
    usados = sorted(e["Código"].unique(), key=lambda c: list(POR_CODIGO).index(c))
    filas = [[POR_CODIGO[c].nombre.split(" (")[0], POR_CODIGO[c].unidad, POR_CODIGO[c].factor_txt,
              POR_CODIGO[c].fuente_fe] for c in usados]
    _tabla(doc, ["Fuente", "Unidad", "Factor", "Referencia"], filas, [4.2, 1.6, 5.6, 5.2])
    fe = FE_RED[FE_RED_ULTIMO_OFICIAL]
    _p(doc, f"Electricidad: factor de red {_num(fe['CO2'], 3)} kg CO2/kWh ({FE_RED_ULTIMO_OFICIAL}, último oficial "
            "publicado en la fecha del informe), salvo que la hoja PARAMETROS indique otro para el año.", tam=9.5)
    gases = sorted({g for g in e["Gas"].unique()})
    _tabla(doc, ["Gas", "GWP100 (AR6)"], [[g, _num(PCG_TEMPLATE[g], 1 if PCG_TEMPLATE[g] % 1 else 0)] for g in gases],
           [8.3, 8.3])

    # 9. Calidad
    _titulo(doc, "9. Calidad de los datos e incertidumbre")
    q = e[e["Año"] == ult].groupby("Método")["tCO2e"].sum()
    filas = [[CALIDAD[m], _num(v), f"{_num(v / tot[ult] * 100, 1)} %"] for m, v in q.sort_values(ascending=False).items()]
    _tabla(doc, ["Calidad del dato (según método)", "tCO2e", "% del total"], filas, [8.6, 4.0, 4.0])
    _p(doc, "Evaluación cualitativa según el método de obtención de cada dato. Las fuentes con estimación básica que "
            "pesan en el total son las primeras candidatas a mejorar (medición o datos de proveedor).", tam=9.5)

    # 10. Gestión
    _titulo(doc, "10. Programas de reducción, verificación y compensaciones")
    for campo, etiqueta in (("Programas de reducción", "Programas de reducción"),
                            ("Verificación externa", "Verificación externa"), ("Compensaciones", "Compensaciones")):
        _p(doc, f"{etiqueta}:", negrita=True, color=NAVY, despues=2)
        _texto_o_pendiente(doc, inf.get(campo), etiqueta.lower())
    _p(doc, "Las emisiones se informan en bruto, sin descontar compensaciones.", tam=9.5)

    _titulo(doc, "11. Contacto")
    _texto_o_pendiente(doc, inf.get("Persona de contacto"), "persona de contacto")

    # Anexo: checklist
    doc.add_page_break()
    _titulo(doc, "Anexo. Cumplimiento de la información requerida (GHG Protocol, cap. 9)")
    filas = [[req_txt, "✓" if cumple else "✗", "Obligatorio" if obl else "Opcional", det] for req_txt, cumple, det, obl in req]
    _tabla(doc, ["Requisito", "", "Tipo", "Detalle"], filas, [6.6, 0.8, 2.2, 7.0])
    if ok:
        _caja(doc, "Declaración", ["El presente inventario fue preparado de acuerdo con el GHG Protocol — Estándar "
                                   "Corporativo de Contabilidad y Reporte, y contiene la información requerida por su "
                                   "capítulo 9 para los Alcances 1 y 2. No fue verificado por un tercero salvo indicación "
                                   "en la sección 10."], fondo="E2EFDA")
    else:
        _caja(doc, "Borrador", ["Faltan requisitos obligatorios (marcados ✗). Hasta completarlos, este documento no "
                                "puede presentarse como conforme al GHG Protocol."], fondo="FFC7CE")
    doc.save(salida)
