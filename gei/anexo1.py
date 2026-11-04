"""
Exportación al Anexo I oficial (Template-Planilla-de-Emisiones-de-GEI-Definitivo-V2.xlsx).

Por qué editamos el XML a mano y no con openpyxl:
  openpyxl, al guardar, BORRA las listas desplegables que el template arma con
  "extensiones" de Excel (GEI Emitido, Unidad, Método, Clasificación). El archivo
  quedaría distinto al oficial. Acá se abre el .xlsx como zip, se escriben SOLO
  las celdas verdes (encabezado O2:O16 y columnas M:V de la tabla) y el resto
  del archivo queda byte a byte igual. Al abrirlo, Excel recalcula las columnas
  violetas (tCO2e, clasificación IPCC) con sus propias fórmulas.

Reglas del template que respetamos:
  - un archivo por año (celda O2),
  - un archivo por clasificación de industria y por departamento (Instructivo B10),
  - la tabla tiene 42 filas (20 a 61).
"""

import re
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from .factores import INDUSTRIAS, METODOS

HOJA = "xl/worksheets/sheet2.xml"      # INVENTARIO
FILA_INI, FILA_FIN = 20, 61
COLS = "MNOPQRSTUV"


def a_dms(grados):
    decimas = round(abs(float(grados)) * 36000)      # décimas de segundo, redondeo una sola vez
    d, resto = divmod(decimas, 36000)
    m, s = divmod(resto, 600)
    return f"{d}°{m}'{s / 10:.1f}\"".replace(".", ",")


def _sin_cache(xml):
    """Borra el resultado guardado de las fórmulas de la tabla (filas 20-61) para que se recalculen."""
    def limpiar(m):
        celda = m.group(0)
        if "<f" not in celda:
            return celda
        celda = re.sub(r"<v>.*?</v>", "", celda, flags=re.S)
        return re.sub(r'\st="(str|n|b|e)"', "", celda, count=1)
    def fila(m):
        return re.sub(r"<c r=\"[A-Z]+\d+\"[^>]*?(?:/>|>.*?</c>)", limpiar, m.group(0), flags=re.S)
    return re.sub(r'<row r="(2\d|[3-5]\d|6[01])"[^>]*>.*?</row>', fila, xml, flags=re.S)


def _poner(xml, ref, valor):
    patron = re.compile(r'<c r="%s"([^>]*?)(/>|>(.*?)</c>)' % ref, re.S)
    m = patron.search(xml)
    if not m:
        raise ValueError(f"La celda {ref} no existe en el template: ¿cambió la versión?")
    estilo = re.search(r'\ss="\d+"', m.group(1))
    s = estilo.group(0) if estilo else ""
    if valor is None or valor == "":
        nuevo = f'<c r="{ref}"{s}/>'
    elif isinstance(valor, (int, float)):
        nuevo = f'<c r="{ref}"{s}><v>{repr(float(valor)) if isinstance(valor, float) else valor}</v></c>'
    else:
        nuevo = f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{escape(str(valor))}</t></is></c>'
    return xml[:m.start()] + nuevo + xml[m.end():]


def filas_template(em_grupo, instalaciones):
    """Agrupa las emisiones en filas del template (instalación × proceso × fuente × método × gas)."""
    em = em_grupo[~em_grupo["Informativo"]]
    g = (em.groupby(["Instalación", "Proceso", "Alcance", "Clasificación", "Método", "Gas"], sort=False)["kg"]
           .sum().reset_index())
    inst = instalaciones.set_index("Instalación")
    filas = []
    for _, r in g.iterrows():
        i = inst.loc[r["Instalación"]]
        es_co2 = r["Gas"] == "CO2"
        filas.append([
            r["Instalación"],
            a_dms(i["Latitud (decimal)"]) if i["Latitud (decimal)"] == i["Latitud (decimal)"] else None,
            a_dms(i["Longitud (decimal)"]) if i["Longitud (decimal)"] == i["Longitud (decimal)"] else None,
            int(r["Alcance"]),
            r["Proceso"],
            r["Clasificación"],
            METODOS[r["Método"]],
            r["Gas"],
            "tn/año" if es_co2 else "kg/año",
            round(r["kg"] / 1000, 3) if es_co2 else round(r["kg"], 3),
        ])
    return filas


def exportar(template, salida, empresa, anio, industria, departamento, instalaciones, em_grupo):
    filas = filas_template(em_grupo, instalaciones)
    if len(filas) > FILA_FIN - FILA_INI + 1:
        raise ValueError(f"{len(filas)} filas: el template admite 42. Agrupar procesos o pedir criterio a la autoridad.")
    sector, subsector = INDUSTRIAS[industria]
    inst = instalaciones[instalaciones["Instalación"].isin(em_grupo["Instalación"].unique())]
    ejido = inst["En ejido urbano"].dropna()
    ciudad = inst["Ciudad / municipio"].dropna()

    with zipfile.ZipFile(template) as z:
        partes = {n: z.read(n) for n in z.namelist()}
        infos = {i.filename: i for i in z.infolist()}

    xml = partes[HOJA].decode("utf-8")
    encabezado = {
        "O2": int(anio),
        "O3": empresa.get("Razón social"),
        "O4": str(empresa.get("CUIT") or ""),
        "O5": empresa.get("Dirección"),
        "O6": empresa.get("Teléfono de contacto"),
        "O7": empresa.get("e-mail de contacto"),
        "O9": empresa.get("Descriptor del sector de la empresa") or "Todo",
        "O10": sector,
        "O11": subsector,
        "O12": industria,
        "O13": ", ".join(inst["Instalación"]) if len(inst) > 1 else (inst["Instalación"].iloc[0] if len(inst) else "Todo"),
        "O14": departamento,
        "O15": ejido.iloc[0] if len(ejido) else "No",
        "O16": ciudad.iloc[0] if len(ciudad) else None,
    }
    for ref, v in encabezado.items():
        xml = _poner(xml, ref, v)
    for k in range(FILA_INI, FILA_FIN + 1):
        valores = filas[k - FILA_INI] if k - FILA_INI < len(filas) else [None] * len(COLS)
        for col, v in zip(COLS, valores):
            xml = _poner(xml, f"{col}{k}", v)
    partes[HOJA] = _sin_cache(xml).encode("utf-8")

    wbxml = partes["xl/workbook.xml"].decode("utf-8")
    if "fullCalcOnLoad" not in wbxml:
        wbxml = re.sub(r"<calcPr([^>]*?)/>", r'<calcPr\1 fullCalcOnLoad="1"/>', wbxml, count=1)
    partes["xl/workbook.xml"] = wbxml.encode("utf-8")

    Path(salida).parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(salida, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, datos in partes.items():
            z.writestr(infos[nombre], datos)
    return len(filas)
