"""
Matriz de carga (Excel propio de ESG, NO el template oficial).

Hojas:
  EMPRESA        datos del encabezado del Anexo I
  INSTALACIONES  una fila por complejo/instalación (industria, departamento, coordenadas)
  ACTIVIDADES    el corazón: una fila por consumo o actividad, por año
  PRODUCCION     toneladas o m3 producidos (para la intensidad)
  PARAMETROS     factor de red por año (si se conoce el oficial del año)
  CATALOGO       solo lectura: qué significa cada código, unidad y factor
"""

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from .factores import CATALOGO, DEPARTAMENTOS, INDUSTRIAS, METODOS

NAVY, AMARILLO, GRIS, CELESTE = "1F3864", "FFFF99", "F2F2F2", "DDEBF7"
_b = Side(style="thin", color="BFBFBF")
BORDE = Border(left=_b, right=_b, top=_b, bottom=_b)

EMPRESA_CAMPOS = ["Razón social", "CUIT", "Dirección", "Teléfono de contacto", "e-mail de contacto",
                  "Descriptor del sector de la empresa", "Profesional que firma", "Título habilitante"]
INST_COLS = ["Instalación", "Industria (template)", "Departamento", "En ejido urbano", "Ciudad / municipio",
             "Latitud (decimal)", "Longitud (decimal)", "Descripción"]
ACT_COLS = ["Año", "Instalación", "Proceso / equipo", "Código", "Cantidad", "Unidad (auto)", "Método",
            "Fuente del dato", "Parámetro 1", "Parámetro 2", "Parámetro 3", "Notas"]
PROD_COLS = ["Año", "Instalación", "Producto", "Cantidad", "Unidad"]

# Hoja INFORME: lo que el GHG Protocol (Estándar Corporativo, cap. 9) exige y el cálculo no sabe.
# (campo, ayuda). "Ninguna"/"No" son respuestas válidas: lo que no vale es dejarlo vacío.
INFORME_CAMPOS = [
    ("Descripción de la empresa", "Actividad principal, ubicación, tamaño. 2-3 líneas."),
    ("Destinatario del informe", "Opcional: cliente, banco, casa matriz… (se menciona en la portada)."),
    ("Año base", "Año contra el que se compara. Por defecto, el primero cargado."),
    ("Justificación del año base", "Por qué ese año (primer año con datos confiables, primer reporte, etc.)."),
    ("Política de recálculo del año base", "Cuándo se recalcula. Ej.: si adquisiciones, cierres o cambios de método alteran el año base en más de 5 %."),
    ("Cambios estructurales en el período", "Adquisiciones, ventas, cierres, tercerizaciones. «Ninguno» si no hubo."),
    ("Causas de los cambios entre años", "Qué explica la suba o baja: producción, eficiencia, cambio de combustible…"),
    ("Instalaciones excluidas y justificación", "«Ninguna» si están todas."),
    ("Fuentes excluidas y justificación", "Ej.: grupo electrógeno de emergencia < 1 % del total. «Ninguna» si no hay."),
    ("Alcance 3", "«No se reporta» o detalle de las categorías incluidas."),
    ("Programas de reducción", "Medidas en marcha o planificadas. «Sin programas» si no hay."),
    ("Verificación externa", "Entidad y norma, o «No verificado»."),
    ("Compensaciones", "Proyectos, estándar y tCO2e, o «No se informan»."),
    ("Persona de contacto", "Nombre, cargo y correo."),
]


def _hdr(ws, cols, anchos):
    for c, (txt, w) in enumerate(zip(cols, anchos), 1):
        cell = ws.cell(1, c, txt)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDE
        ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = "A2"


def _filas(ws, filas, ncols, desde=2, hasta=None, color=AMARILLO):
    hasta = hasta or max(desde + len(filas) + 30, desde)
    for r in range(desde, hasta):
        for c in range(1, ncols + 1):
            cell = ws.cell(r, c)
            cell.fill = PatternFill("solid", fgColor=color)
            cell.border = BORDE
    for i, fila in enumerate(filas):
        for c, v in enumerate(fila, 1):
            ws.cell(desde + i, c, v)
    return hasta


def generar(ruta, ejemplo=None):
    """Crea la matriz vacía, o con un caso de ejemplo (dict con las mismas claves que leer())."""
    ej = ejemplo or {}
    wb = openpyxl.Workbook()

    # EMPRESA
    ws = wb.active
    ws.title = "EMPRESA"
    _hdr(ws, ["Campo", "Valor"], [36, 60])
    emp = ej.get("empresa", {})
    _filas(ws, [[k, emp.get(k)] for k in EMPRESA_CAMPOS], 2, hasta=2 + len(EMPRESA_CAMPOS))
    for r in range(2, 2 + len(EMPRESA_CAMPOS)):
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=GRIS)

    # INSTALACIONES
    ws = wb.create_sheet("INSTALACIONES")
    _hdr(ws, INST_COLS, [30, 52, 18, 12, 20, 14, 14, 40])
    filas = ej.get("instalaciones", pd.DataFrame(columns=INST_COLS))[INST_COLS].values.tolist()
    fin = _filas(ws, filas, len(INST_COLS))
    lista_ind = '"' + ",".join(INDUSTRIAS)[:250] + '"'
    ws.cell(1, 9, "Industrias válidas del template:").font = Font(bold=True)
    for i, ind in enumerate(INDUSTRIAS, 2):
        ws.cell(i, 9, ind).font = Font(size=8, italic=True)
    ws.column_dimensions["I"].width = 60
    dv = DataValidation(type="list", formula1=f"=$I$2:$I${1 + len(INDUSTRIAS)}", allow_blank=True)
    dv.add(f"B2:B{fin}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="list", formula1='"' + ",".join(DEPARTAMENTOS) + '"', allow_blank=True)
    dv.add(f"C2:C{fin}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="list", formula1='"Si,No"', allow_blank=True)
    dv.add(f"D2:D{fin}")
    ws.add_data_validation(dv)

    # ACTIVIDADES
    ws = wb.create_sheet("ACTIVIDADES")
    _hdr(ws, ACT_COLS, [8, 26, 34, 22, 14, 10, 9, 38, 12, 12, 12, 40])
    df = ej.get("actividades", pd.DataFrame(columns=ACT_COLS))
    filas = [[r.get(c) for c in ACT_COLS] for _, r in df.iterrows()]
    fin = _filas(ws, filas, len(ACT_COLS), hasta=max(len(filas) + 2 + 60, 200))
    for r in range(2, fin):
        ws.cell(r, 6, f'=IFERROR(VLOOKUP(D{r},CATALOGO!$A:$D,4,FALSE),"")').fill = PatternFill("solid", fgColor=GRIS)
        ws.cell(r, 5).number_format = "#,##0.00"
    dv = DataValidation(type="list", formula1=f"=CATALOGO!$A$2:$A${1 + len(CATALOGO)}", allow_blank=True)
    dv.add(f"D2:D{fin}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="list", formula1='"' + ",".join(METODOS) + '"', allow_blank=True)
    dv.add(f"G2:G{fin}")
    ws.add_data_validation(dv)
    dv = DataValidation(type="list", formula1="=INSTALACIONES!$A$2:$A$60", allow_blank=True)
    dv.add(f"B2:B{fin}")
    ws.add_data_validation(dv)

    # PRODUCCION
    ws = wb.create_sheet("PRODUCCION")
    _hdr(ws, PROD_COLS, [8, 26, 30, 16, 12])
    df = ej.get("produccion", pd.DataFrame(columns=PROD_COLS))
    _filas(ws, df[PROD_COLS].values.tolist(), len(PROD_COLS))

    # PARAMETROS
    ws = wb.create_sheet("PARAMETROS")
    _hdr(ws, ["Año", "FE red kg CO2/kWh", "FE red kg CH4/kWh", "FE red kg N2O/kWh", "Fuente"], [8, 18, 18, 18, 60])
    _filas(ws, [], 5, hasta=8)
    ws.cell(10, 1, "Dejar vacío para usar el último factor oficial publicado (SNIGEI-AR 2025, Tabla 5: 2022 = "
                   "0,283 kg CO2/kWh). Cargar acá el del año si CAMMESA o Ambiente de Nación lo publican.").font = Font(italic=True, size=9)

    # INFORME (para el informe a terceros según GHG Protocol)
    ws = wb.create_sheet("INFORME", 1)
    _hdr(ws, ["Campo", "Respuesta de la empresa", "Qué poner"], [36, 70, 70])
    inf = ej.get("informe", {})
    _filas(ws, [[k, inf.get(k), ayuda] for k, ayuda in INFORME_CAMPOS], 3, hasta=2 + len(INFORME_CAMPOS))
    for r in range(2, 2 + len(INFORME_CAMPOS)):
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=GRIS)
        ws.cell(r, 3).fill = PatternFill("solid", fgColor=GRIS)
        ws.cell(r, 3).font = Font(italic=True, size=9)
        for c in (2, 3):
            ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 42

    # CATALOGO
    ws = wb.create_sheet("CATALOGO")
    cols = ["Código", "Actividad", "Sector", "Unidad", "Alcance", "Clasificación en el template", "Factor",
            "Fuente del factor", "Parámetros (en orden)", "Nota"]
    _hdr(ws, cols, [22, 60, 13, 9, 9, 36, 60, 60, 60, 60])
    filas = [[a.codigo, a.nombre, a.sector, a.unidad, a.alcance, a.clasificacion, a.factor_txt, a.fuente_fe,
              "; ".join(f"{n} (defecto {d:g}): {t}" for n, d, t in a.parametros), a.nota] for a in CATALOGO]
    _filas(ws, filas, len(cols), hasta=2 + len(filas), color=CELESTE)
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.protection.sheet = True

    wb.save(ruta)


def leer(ruta):
    wb = openpyxl.load_workbook(ruta, data_only=True)

    def tabla(nombre, cols):
        ws = wb[nombre]
        filas = [list(r[:len(cols)]) for r in ws.iter_rows(min_row=2, values_only=True)]
        df = pd.DataFrame(filas, columns=cols)
        return df.dropna(how="all", subset=[c for c in cols if c != "Unidad (auto)"]).reset_index(drop=True)

    emp = {r[0]: r[1] for r in wb["EMPRESA"].iter_rows(min_row=2, values_only=True) if r[0]}
    inst = tabla("INSTALACIONES", INST_COLS)
    act = tabla("ACTIVIDADES", ACT_COLS)
    act = act.dropna(subset=["Código", "Cantidad"], how="all")
    prod = tabla("PRODUCCION", PROD_COLS)
    fe = {}
    for r in wb["PARAMETROS"].iter_rows(min_row=2, max_row=8, values_only=True):
        if r[0] and r[1] is not None:
            fe[int(r[0])] = {"CO2": float(r[1]), "CH4": float(r[2] or 0), "N2O": float(r[3] or 0)}
    informe = {}
    if "INFORME" in wb.sheetnames:
        informe = {r[0]: r[1] for r in wb["INFORME"].iter_rows(min_row=2, values_only=True)
                   if r[0] and r[1] not in (None, "")}
    return {"empresa": emp, "instalaciones": inst, "actividades": act, "produccion": prod,
            "parametros": {"fe_red": fe}, "informe": informe}
