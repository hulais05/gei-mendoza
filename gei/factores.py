"""
Catálogo de actividades y factores de emisión — Registro GEI Mendoza.

Cada actividad sabe:
  - en qué unidad se carga el dato de actividad,
  - a qué Alcance pertenece,
  - qué "Clasificación de fuente de emisión" le corresponde en el template
    oficial (Anexo I, Res. 1/2026), con el texto EXACTO de su lista desplegable,
  - cómo convertir el dato de actividad en masa de cada gas (kg/año),
  - de dónde sale cada factor (para la Tabla 1 del Anexo II).

Fuentes principales
  [SNIGEI]  SSAmb (2025). Factores de Emisión SNIGEI-AR para Huella de Carbono.
            Basado en el Informe Nacional de Inventario del 1er IBT (2024).
            https://inventariogei.ambiente.gob.ar/files/Factores_de_Emision_SNIGEIAR_para_HC_2025.pdf
  [IPCC06]  IPCC 2006, Directrices para inventarios nacionales de GEI.
  [AR6]     Potenciales de calentamiento: los mismos que usa el template oficial
            (hoja "Tablas y Relaciones", tabla Gas). Se copian tal cual para que
            nuestro tCO2e coincida con el que calcula la planilla del gobierno.
"""

from dataclasses import dataclass, field
from typing import Callable

# ─── PCG (GWP100) tal como los trae el template oficial V2 ──────────────────
# Nombres EXACTOS de la lista "GEI Emitido" del template.
PCG_TEMPLATE = {
    "CO2": 1,
    "CH4 orgánico": 27,
    "CH4 fósil combustión": 27,
    "CH4 fósil venteado": 29.8,
    "N2O": 273,
    "HFC-32": 770,
    "HFC-125": 3744,
    "HFC-134a": 1526,
    "HFC-143a": 5807,
    "SF6": 25184,
}

# Umbral de la Res. 1/2026: por encima hay que segregar por instalación/proceso.
UMBRAL_TCO2E = 10_000

METODOS = {
    "N1": "Estimación básica",
    "N2": "Estimación intermedia",
    "N3": "Estimación mejorada",
    "MI": "Medición indirecta",
    "MD": "Medición directa",
}

# ─── Constantes físicas ─────────────────────────────────────────────────────
VOL_MOLAR_STD = 0.023645   # m3/mol a 15 °C y 101,325 kPa (m3 estándar argentino)
M_CO2, M_CH4 = 44.01, 16.04  # g/mol
C_A_CO2 = 44.01 / 12.011
FRACCION_C_GASOIL = 0.87   # kg C / kg gasoil (IPCC06 Vol.2 Tabla 1.3: 20,2 kgC/GJ x 43 GJ/t)
EFICIENCIA_ANTORCHA = 0.98  # fracción de hidrocarburo que se quema (IPCC06 / API Compendium)

# ─── Factor de red eléctrica (Alcance 2) ────────────────────────────────────
# [SNIGEI] Tabla 5. Último año oficial publicado = 2022.
# Para 2024 y 2025 se usa 2022 por defecto y la herramienta lo marca como
# "a verificar": si CAMMESA o la SSAmb publican el valor del año, se reemplaza
# en la hoja PARAMETROS de la matriz.
FE_RED = {
    2022: {"CO2": 0.283, "CH4": 0.0000167, "N2O": 0.0000097},
    2021: {"CO2": 0.292, "CH4": 0.0000188, "N2O": 0.0000111},
}
FE_RED_ULTIMO_OFICIAL = 2022

# ─── Refrigerantes: composición en masa de las mezclas ──────────────────────
# El template solo trae gases puros, así que una recarga de R-410A se reporta
# como su HFC-32 + su HFC-125 (composiciones ASHRAE 34; [SNIGEI] Tabla 7).
MEZCLAS_REFRIGERANTES = {
    "R-410A": {"HFC-32": 0.50, "HFC-125": 0.50},
    "R-407C": {"HFC-32": 0.23, "HFC-125": 0.25, "HFC-134a": 0.52},
    "R-404A": {"HFC-125": 0.44, "HFC-143a": 0.52, "HFC-134a": 0.04},
    "R-134a": {"HFC-134a": 1.00},
}


@dataclass
class Actividad:
    codigo: str
    nombre: str
    sector: str                 # "Minería", "Hidrocarburos" o "Ambos"
    unidad: str                 # unidad del dato de actividad
    alcance: int
    clasificacion: str          # texto exacto de la lista del template
    categoria_anexo2: str       # Combustión fija / móvil / Proceso / Fugitivas / Electricidad
    fuente_fe: str              # cita del factor (va a la Tabla 1 del Anexo II)
    factor_txt: str             # el factor escrito para humanos
    calcular: Callable = field(repr=False)   # (cantidad, params, año, fe_red) -> {gas: kg}
    parametros: tuple = ()      # [(nombre, valor por defecto, descripción)]
    nota: str = ""


def _por_unidad(co2, ch4, n2o, gas_ch4="CH4 fósil combustión", co2_bio=0.0):
    """Factores lineales kg/unidad → función de cálculo."""
    def f(q, p, anio, fe_red):
        out = {"CO2": q * co2, gas_ch4: q * ch4, "N2O": q * n2o}
        if co2_bio:
            out["CO2 biogénico (informativo)"] = q * co2_bio
        return out
    return f


def _electricidad(q, p, anio, fe_red):
    fe = fe_red
    return {"CO2": q * fe["CO2"], "CH4 fósil combustión": q * fe["CH4"], "N2O": q * fe["N2O"]}


def _explosivo(q, p, anio, fe_red):
    # Balance de masa: todo el carbono del combustible del explosivo sale como CO2.
    frac_comb = p.get("Fracción de combustible", 0.06)
    return {"CO2": q * frac_comb * FRACCION_C_GASOIL * C_A_CO2}


def _cal(q, p, anio, fe_red):
    return {"CO2": q * 750.0}   # kg CO2 / t cal  [SNIGEI] Tabla 6 (IPCC06, cal alta en calcio)


def _antorcha(q, p, anio, fe_red):
    x_ch4 = p.get("Fracción molar CH4", 0.85)
    x_c2 = p.get("Fracción molar C2+ (como etano)", 0.10)
    x_co2 = p.get("Fracción molar CO2", 0.02)
    mol = q / VOL_MOLAR_STD
    eta = EFICIENCIA_ANTORCHA
    co2 = mol * (eta * (x_ch4 + 2 * x_c2) + x_co2) * M_CO2 / 1000
    ch4 = mol * x_ch4 * (1 - eta) * M_CH4 / 1000
    return {"CO2": co2, "CH4 fósil combustión": ch4}


def _venteo(q, p, anio, fe_red):
    x_ch4 = p.get("Fracción molar CH4", 0.85)
    x_co2 = p.get("Fracción molar CO2", 0.02)
    mol = q / VOL_MOLAR_STD
    return {"CH4 fósil venteado": mol * x_ch4 * M_CH4 / 1000, "CO2": mol * x_co2 * M_CO2 / 1000}


def _ch4_directo(q, p, anio, fe_red):
    return {"CH4 fósil venteado": q}


def _refrigerante(mezcla):
    def f(q, p, anio, fe_red):
        return {gas: q * frac for gas, frac in MEZCLAS_REFRIGERANTES[mezcla].items()}
    return f


def _sf6(q, p, anio, fe_red):
    return {"SF6": q}


SNIGEI_MOV = "SNIGEI-AR 2025, Tabla 2 (gas oil grado 2, blend 7,5 %)"
SNIGEI_FIJ = "SNIGEI-AR 2025, Tabla 4 (instalaciones fijas)"

CATALOGO = [
    # ── Combustión móvil ────────────────────────────────────────────────────
    Actividad("GASOIL_EQUIPOS", "Gasoil — equipos móviles de mina/yacimiento (camiones, palas, cargadoras, perforadoras)",
              "Ambos", "L", 1, "Camiones", "Combustión móvil", SNIGEI_MOV,
              "2,3390 kg CO2/L; 0,00013664 kg CH4/L; 0,00013664 kg N2O/L",
              _por_unidad(2.3390, 0.00013664, 0.00013664, co2_bio=0.1716),
              nota="El CO2 del biodiésel del corte (0,1716 kg/L) es biogénico: se informa aparte y no suma."),
    Actividad("GASOIL_LIVIANOS", "Gasoil — camionetas y vehículos livianos",
              "Ambos", "L", 1, "Vehículos livianos y motos", "Combustión móvil", SNIGEI_MOV,
              "2,3390 kg CO2/L; 0,00013664 kg CH4/L; 0,00013664 kg N2O/L",
              _por_unidad(2.3390, 0.00013664, 0.00013664, co2_bio=0.1716)),
    Actividad("NAFTA_LIVIANOS", "Nafta — vehículos livianos",
              "Ambos", "L", 1, "Vehículos livianos y motos", "Combustión móvil",
              "SNIGEI-AR 2025, Tabla 2 (nafta grado 2, bioetanol 12 %)",
              "1,9693 kg CO2/L; 0,0009567 kg CH4/L; 0,000093756 kg N2O/L",
              _por_unidad(1.9693, 0.0009567, 0.000093756, co2_bio=0.1795)),
    # ── Combustión fija ─────────────────────────────────────────────────────
    Actividad("GASOIL_GENERADORES", "Gasoil — grupos electrógenos, bombas y compresores",
              "Ambos", "L", 1, "motores fijos (bombas, compresores, etc.)", "Combustión fija", SNIGEI_FIJ,
              "2,529 kg CO2/L; 0,0003612 kg CH4/L; 0,0000258 kg N2O/L",
              _por_unidad(2.529, 0.0003612, 0.0000258)),
    Actividad("GN_HORNOS", "Gas natural — hornos y calentadores (incl. horno de cal)",
              "Ambos", "m3", 1, "Calentadores y hornos", "Combustión fija", SNIGEI_FIJ,
              "1,951 kg CO2/m3; 0,0001743 kg CH4/m3; 0,0000033 kg N2O/m3",
              _por_unidad(1.951, 0.0001743, 0.0000033)),
    Actividad("GN_MOTORES", "Gas natural — motocompresores y motogeneradores",
              "Ambos", "m3", 1, "motores fijos (bombas, compresores, etc.)", "Combustión fija", SNIGEI_FIJ,
              "1,951 kg CO2/m3; 0,0001743 kg CH4/m3; 0,0000033 kg N2O/m3",
              _por_unidad(1.951, 0.0001743, 0.0000033)),
    Actividad("GN_CALDERAS", "Gas natural — calderas",
              "Ambos", "m3", 1, "Calderas", "Combustión fija", SNIGEI_FIJ,
              "1,951 kg CO2/m3; 0,0001743 kg CH4/m3; 0,0000033 kg N2O/m3",
              _por_unidad(1.951, 0.0001743, 0.0000033)),
    Actividad("GLP", "GLP — calefacción, cocina de campamento, quemadores",
              "Ambos", "kg", 1, "Quemadores de gas", "Combustión fija", SNIGEI_FIJ,
              "2,901 kg CO2/kg; 0,0002331 kg CH4/kg; 0,0000044 kg N2O/kg",
              _por_unidad(2.901, 0.0002331, 0.0000044)),
    # ── Minería: proceso ────────────────────────────────────────────────────
    Actividad("EXPLOSIVOS", "Explosivos (ANFO / emulsión) — voladuras",
              "Minería", "kg", 1, "Uso de explosivos/detonacines (minas)", "Emisiones de proceso",
              "Balance de masa (IPCC 2006 Vol.2 Cap.1): carbono del combustible del explosivo",
              "CO2 = kg explosivo x fracción de combustible x 0,87 kgC/kg x 44/12 (≈0,19 kg CO2/kg para ANFO 94/6)",
              _explosivo, parametros=(("Fracción de combustible", 0.06, "Gasoil en el explosivo (ANFO típico 0,06)"),),
              nota="Si el gasoil del ANFO ya se cargó como consumo de combustible, no se vuelve a contar acá."),
    Actividad("CAL_CALCINACION", "Calcinación de caliza — producción de cal viva",
              "Minería", "t cal", 1, "Emisiones específicas del proceso", "Emisiones de proceso",
              "SNIGEI-AR 2025, Tabla 6 (IPCC 2006 Vol.3 Cap.2, cal alta en calcio)",
              "750 kg CO2/t cal producida", _cal,
              nota="Solo el CO2 de la descarbonatación. El gas o fuel del horno va como combustión fija."),
    # ── Hidrocarburos: antorcha, venteo, fugas ─────────────────────────────
    Actividad("GAS_ANTORCHA", "Gas quemado en antorcha",
              "Hidrocarburos", "m3", 1, "Antorcha", "Combustión fija",
              "Balance de carbono con composición del gas; eficiencia de quema 98 % (IPCC 2006 / API Compendium)",
              "CO2 = mol x (0,98 x (xCH4 + 2·xC2+) + xCO2) x 44 g/mol; CH4 no quemado = mol x xCH4 x 0,02 x 16 g/mol",
              _antorcha, parametros=(("Fracción molar CH4", 0.85, "Del cromatógrafo"),
                                     ("Fracción molar C2+ (como etano)", 0.10, "Etano y más pesados"),
                                     ("Fracción molar CO2", 0.02, "CO2 que ya trae el gas"))),
    Actividad("GAS_VENTEO", "Gas venteado (sin quemar)",
              "Hidrocarburos", "m3", 1, "Venteo de gas crudo (stacks)", "Emisiones fugitivas",
              "Balance de masa con composición del gas",
              "CH4 = mol x xCH4 x 16 g/mol; CO2 = mol x xCO2 x 44 g/mol",
              _venteo, parametros=(("Fracción molar CH4", 0.85, "Del cromatógrafo"),
                                   ("Fracción molar CO2", 0.02, "CO2 que ya trae el gas"))),
    Actividad("CH4_FUGAS_MEDIDAS", "Fugas de metano medidas o estimadas (campaña LDAR, cámara, balance)",
              "Hidrocarburos", "kg CH4", 1, "Fugas en equipos", "Emisiones fugitivas",
              "Dato propio de la empresa (medición o estimación documentada)", "Masa de CH4 directa",
              _ch4_directo),
    # ── Fugitivas comunes ───────────────────────────────────────────────────
    *[Actividad(f"REFRIG_{m.replace('-', '')}", f"Recarga de refrigerante {m}",
                "Ambos", "kg", 1, "Fugas en equipos de refrigeración", "Emisiones fugitivas",
                "Método de recarga = fuga (GHG Protocol); composición ASHRAE 34 / SNIGEI-AR 2025 Tabla 7",
                " + ".join(f"{int(fr*100)} % {g}" for g, fr in MEZCLAS_REFRIGERANTES[m].items()),
                _refrigerante(m)) for m in MEZCLAS_REFRIGERANTES],
    Actividad("SF6", "Recarga de SF6 en celdas / interruptores de media tensión",
              "Ambos", "kg", 1, "Fugas en equipos", "Emisiones fugitivas",
              "Método de recarga = fuga (GHG Protocol)", "kg SF6 recargado", _sf6),
    # ── Alcance 2 ───────────────────────────────────────────────────────────
    Actividad("ELECTRICIDAD", "Electricidad comprada a la red",
              "Ambos", "kWh", 2, "Electricidad Comprada", "Electricidad adquirida",
              "SNIGEI-AR 2025, Tabla 5 (factor de red, ver hoja PARAMETROS)",
              "kg CO2/kWh del año (por defecto 0,283, dato oficial 2022)", _electricidad),
]

POR_CODIGO = {a.codigo: a for a in CATALOGO}

# ─── Clasificación del encabezado del template (textos exactos) ─────────────
INDUSTRIAS = {
    "Minería de minerales no metálicos (No cementera)": ("Sector Industrial", "Subsector industria minera"),
    "Minería de minerales metálicos": ("Sector Industrial", "Subsector industria minera"),
    "Servicios relacionados con minería": ("Sector Industrial", "Subsector industria minera"),
    "Extracción de petróleo, convencional y no convencional":
        ("Sector Energía", "Subsector explotación, producción, transporte y distribución de hidrocarburos"),
    "Extracción de gas, asociado y no asociado a la extracción de petróleo, convencional y no convencional":
        ("Sector Energía", "Subsector explotación, producción, transporte y distribución de hidrocarburos"),
    "Perforación de pozos petroleros y de gas":
        ("Sector Energía", "Subsector explotación, producción, transporte y distribución de hidrocarburos"),
    "Procesamiento de gas natural":
        ("Sector Energía", "Subsector explotación, producción, transporte y distribución de hidrocarburos"),
}

DEPARTAMENTOS = ["Capital", "General Alvear", "Godoy Cruz", "Guaymallén", "Junín", "La Paz", "Las Heras",
                 "Lavalle", "Luján de Cuyo", "Maipú", "Malargüe", "Rivadavia", "San Carlos", "San Martín",
                 "San Rafael", "Santa Rosa", "Tunuyán", "Tupungato"]
