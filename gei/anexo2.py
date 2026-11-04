"""
Tablas del Anexo II (metodología): Tabla 1 por año y Tabla 5 (seguimiento entre períodos).

Sale en Excel para copiar al Word del Anexo II. El Word lo arma y firma el
profesional competente; esto le da las tablas ya completas y trazables.
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

NAVY = "1F3864"
_b = Side(style="thin", color="808080")
BORDE = Border(left=_b, right=_b, top=_b, bottom=_b)
ORDEN_CAT = ["Combustión fija", "Combustión móvil", "Emisiones de proceso", "Emisiones fugitivas",
             "Electricidad adquirida"]


def _hdr(ws, fila, cols, anchos=None):
    for c, t in enumerate(cols, 1):
        x = ws.cell(fila, c, t)
        x.font = Font(bold=True, color="FFFFFF")
        x.fill = PatternFill("solid", fgColor=NAVY)
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        x.border = BORDE
        if anchos:
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = anchos[c - 1]


def exportar(em, salida, empresa):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    cols = ["Alcance", "Categoría de fuente", "Complejo / instalación", "Proceso, subproceso, equipo",
            "GEI emitido", "Factor de emisión", "Fuente del factor", "Dato de actividad",
            "Fuente del dato de actividad", "Valor (t gas/año)", "tCO2e/año"]
    for anio in sorted(em["Año"].unique()):
        ws = wb.create_sheet(f"Tabla 1 - {anio}")
        ws.cell(1, 1, f"Tabla 1. Registro de Emisiones — {empresa.get('Razón social', '')}").font = Font(bold=True, size=12)
        ws.cell(2, 1, f"Período: 01/01/{anio} al 31/12/{anio}")
        _hdr(ws, 4, cols, [9, 20, 24, 30, 20, 46, 46, 18, 34, 14, 12])
        d = em[(em["Año"] == anio) & (~em["Informativo"])].copy()
        d["_o"] = d["Categoría"].map({c: i for i, c in enumerate(ORDEN_CAT)})
        d = d.sort_values(["Alcance", "_o", "Instalación", "Proceso"])
        r = 5
        for _, x in d.iterrows():
            vals = [f"Alcance {x['Alcance']}", x["Categoría"], x["Instalación"], x["Proceso"], x["Gas"],
                    x["Factor"], x["Fuente del factor"], x["Dato de actividad"], x["Fuente del dato"],
                    round(x["kg"] / 1000, 4), round(x["tCO2e"], 2)]
            for c, v in enumerate(vals, 1):
                cell = ws.cell(r, c, v)
                cell.border = BORDE
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            r += 1
        ws.cell(r, 10, "TOTAL").font = Font(bold=True)
        ws.cell(r, 11, round(d["tCO2e"].sum(), 2)).font = Font(bold=True)
        bio = em[(em["Año"] == anio) & (em["Informativo"])]["kg"].sum() / 1000
        if bio:
            ws.cell(r + 2, 1, f"CO2 biogénico (corte de biocombustibles), informado aparte y no sumado: {bio:,.1f} t".replace(",", "X").replace(".", ",").replace("X", "."))
        ws.freeze_panes = "A5"

    # Tabla 5
    ws = wb.create_sheet("Tabla 5 - Seguimiento")
    ws.cell(1, 1, "Tabla 5. Seguimiento de emisiones entre períodos (tCO2e/año)").font = Font(bold=True, size=12)
    d = em[~em["Informativo"]]
    p = d.pivot_table(index=["Instalación", "Categoría"], columns="Año", values="tCO2e", aggfunc="sum", fill_value=0)
    anios = list(p.columns)
    cols = ["Complejo / instalación", "Categoría"] + [str(a) for a in anios]
    if len(anios) >= 2:
        cols += [f"Variación {anios[-2]}→{anios[-1]} (%)"]
    _hdr(ws, 3, cols, [26, 24] + [14] * len(anios) + [22])
    r = 4
    for (inst, cat), fila in p.iterrows():
        vals = [inst, cat] + [round(fila[a], 1) for a in anios]
        if len(anios) >= 2:
            ant, act = fila[anios[-2]], fila[anios[-1]]
            vals.append(round((act - ant) / ant * 100, 1) if ant else None)
        for c, v in enumerate(vals, 1):
            ws.cell(r, c, v).border = BORDE
        r += 1
    tot = [round(d[d["Año"] == a]["tCO2e"].sum(), 1) for a in anios]
    ws.cell(r, 1, "TOTAL").font = Font(bold=True)
    for k, v in enumerate(tot):
        ws.cell(r, 3 + k, v).font = Font(bold=True)
    if len(tot) >= 2 and tot[-2]:
        ws.cell(r, 3 + len(tot), round((tot[-1] - tot[-2]) / tot[-2] * 100, 1)).font = Font(bold=True)
    wb.save(salida)
