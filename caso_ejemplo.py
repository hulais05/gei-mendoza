#!/usr/bin/env python3
"""
Caso de ejemplo FICTICIO — no corresponde a ninguna empresa real.

"Empresa Ejemplo S.A." con dos negocios en Mendoza:
  - Minería no metálica: cantera de caliza + planta de cal (Las Heras).
  - Hidrocarburos: yacimiento convencional (Malargüe).
Los consumos son órdenes de magnitud verosímiles, no datos medidos.
"""

import pandas as pd

from gei.matriz import ACT_COLS, generar

EMPRESA = {
    "Razón social": "Empresa Ejemplo S.A. (caso ficticio)",
    "CUIT": "30-00000000-0",
    "Dirección": "Calle Ficticia 123, Mendoza",
    "Teléfono de contacto": "0261 000-0000",
    "e-mail de contacto": "ambiente@ejemplo.com.ar",
    "Descriptor del sector de la empresa": "Minería no metálica y producción de hidrocarburos",
    "Profesional que firma": "(a completar)",
    "Título habilitante": "(a completar)",
}

INSTALACIONES = pd.DataFrame([
    ["Cantera Los Cerrillos", "Minería de minerales no metálicos (No cementera)", "Las Heras", "No", None,
     -32.7650, -68.9850, "Cantera de caliza a cielo abierto, perforación y voladura"],
    ["Planta de Cal Los Cerrillos", "Minería de minerales no metálicos (No cementera)", "Las Heras", "No", None,
     -32.7820, -68.9300, "Hornos verticales a gas natural, molienda y embolsado"],
    ["Yacimiento Cerro Ejemplo", "Extracción de petróleo, convencional y no convencional", "Malargüe", "No", None,
     -35.6200, -69.3800, "Pozos, baterías, antorcha y venteos"],
    ["Planta de Tratamiento Cerro Ejemplo", "Extracción de petróleo, convencional y no convencional", "Malargüe", "No", None,
     -35.6050, -69.3500, "Separación, deshidratación y compresión de gas"],
], columns=["Instalación", "Industria (template)", "Departamento", "En ejido urbano", "Ciudad / municipio",
            "Latitud (decimal)", "Longitud (decimal)", "Descripción"])

# (instalación, proceso, código, cantidad 2024, cantidad 2025, método, fuente, parámetros)
BASE = [
    ("Cantera Los Cerrillos", "Flota de acarreo y carguío", "GASOIL_EQUIPOS", 620_000, 655_000, "N1",
     "Facturas YPF + planilla de despacho de combustible", ()),
    ("Cantera Los Cerrillos", "Perforadoras y grupo electrógeno", "GASOIL_GENERADORES", 48_000, 52_000, "N1",
     "Planilla de despacho de combustible", ()),
    ("Cantera Los Cerrillos", "Voladuras", "EXPLOSIVOS", 240_000, 255_000, "N1",
     "Remitos del proveedor de explosivos (ANFO)", (0.06,)),
    ("Cantera Los Cerrillos", "Campamento y oficinas", "GLP", 14_000, 15_000, "N1", "Facturas de GLP a granel", ()),
    ("Cantera Los Cerrillos", "Trituración primaria", "ELECTRICIDAD", 2_100_000, 2_250_000, "N2",
     "Facturas de la distribuidora", ()),
    ("Cantera Los Cerrillos", "Camionetas de supervisión", "GASOIL_LIVIANOS", 38_000, 40_000, "N1",
     "Tarjetas de combustible de flota", ()),
    ("Planta de Cal Los Cerrillos", "Hornos verticales 1 y 2", "GN_HORNOS", 6_800_000, 7_150_000, "N2",
     "Facturas de la distribuidora de gas + caudalímetro", ()),
    ("Planta de Cal Los Cerrillos", "Hornos verticales 1 y 2", "CAL_CALCINACION", 60_000, 63_500, "N1",
     "Balanza de despacho: t de cal viva producida", ()),
    ("Planta de Cal Los Cerrillos", "Molienda, hidratación y embolsado", "ELECTRICIDAD", 6_900_000, 7_200_000, "N2",
     "Facturas de la distribuidora", ()),
    ("Planta de Cal Los Cerrillos", "Aire acondicionado de sala de control", "REFRIG_R410A", 18, 12, "N1",
     "Órdenes de trabajo de mantenimiento (kg recargados)", ()),
    ("Yacimiento Cerro Ejemplo", "Antorcha batería 1", "GAS_ANTORCHA", 3_100_000, 2_400_000, "N2",
     "Placa orificio de antorcha", (0.85, 0.10, 0.02)),
    ("Yacimiento Cerro Ejemplo", "Venteo de tanques y casing", "GAS_VENTEO", 420_000, 310_000, "N1",
     "Estimación por relación gas-petróleo", (0.85, 0.02)),
    ("Yacimiento Cerro Ejemplo", "Fugas en válvulas y bridas", "CH4_FUGAS_MEDIDAS", 38_000, 26_000, "MI",
     "Campaña LDAR con cámara OGI", ()),
    ("Yacimiento Cerro Ejemplo", "Equipos de workover y pulling", "GASOIL_EQUIPOS", 880_000, 910_000, "N1",
     "Facturas de combustible del contratista", ()),
    ("Planta de Tratamiento Cerro Ejemplo", "Motocompresores", "GN_MOTORES", 5_200_000, 5_400_000, "N2",
     "Medidor de gas combustible", ()),
    ("Planta de Tratamiento Cerro Ejemplo", "Calentadores de crudo", "GN_HORNOS", 1_900_000, 1_850_000, "N2",
     "Medidor de gas combustible", ()),
    ("Planta de Tratamiento Cerro Ejemplo", "Bombas y equipos eléctricos", "ELECTRICIDAD", 19_500_000, 20_100_000, "N2",
     "Facturas CAMMESA / distribuidora", ()),
    ("Planta de Tratamiento Cerro Ejemplo", "Celdas de media tensión", "SF6", 2.5, 0, "N1",
     "Órdenes de trabajo de mantenimiento", ()),
]

filas = []
for anio, k in ((2024, 3), (2025, 4)):
    for b in BASE:
        if b[k] == 0:
            continue
        p = list(b[7]) + [None] * (3 - len(b[7]))
        filas.append({"Año": anio, "Instalación": b[0], "Proceso / equipo": b[1], "Código": b[2], "Cantidad": b[k],
                      "Método": b[5], "Fuente del dato": b[6], "Parámetro 1": p[0], "Parámetro 2": p[1],
                      "Parámetro 3": p[2], "Notas": None})
ACTIVIDADES = pd.DataFrame(filas, columns=ACT_COLS)

PRODUCCION = pd.DataFrame([
    [2024, "Planta de Cal Los Cerrillos", "Cal viva", 60_000, "t"],
    [2025, "Planta de Cal Los Cerrillos", "Cal viva", 63_500, "t"],
    [2024, "Cantera Los Cerrillos", "Caliza extraída", 118_000, "t"],
    [2025, "Cantera Los Cerrillos", "Caliza extraída", 125_000, "t"],
    [2024, "Yacimiento Cerro Ejemplo", "Petróleo", 340_000, "m3"],
    [2025, "Yacimiento Cerro Ejemplo", "Petróleo", 352_000, "m3"],
], columns=["Año", "Instalación", "Producto", "Cantidad", "Unidad"])

INFORME = {
    "Descripción de la empresa": "Empresa mendocina con dos negocios: producción de cal viva a partir de caliza propia "
                                 "(cantera y planta en Las Heras) y producción de petróleo convencional (Malargüe).",
    "Destinatario del informe": "Casa matriz (caso ficticio)",
    "Año base": 2024,
    "Justificación del año base": "Primer año con registros completos de consumo y primer período reportado ante el "
                                  "Registro Provincial de Emisiones GEI de Mendoza.",
    "Política de recálculo del año base": "Se recalcula el año base si adquisiciones, ventas, cierres, tercerizaciones o "
                                          "cambios de metodología modifican sus emisiones en más de un 5 %.",
    "Cambios estructurales en el período": "Ninguno.",
    "Causas de los cambios entre años": "Minería: +5,8 % de producción de cal. Hidrocarburos: menos gas a antorcha y "
                                        "venteo por la puesta en marcha de un compresor de recuperación de gas.",
    "Instalaciones excluidas y justificación": "Ninguna.",
    "Fuentes excluidas y justificación": "Ninguna.",
    "Alcance 3": "No se reporta (opcional según GHG Protocol y Res. 1/2026).",
    "Programas de reducción": "Recuperación de gas de antorcha (en ejecución desde 2025); programa LDAR semestral "
                              "de detección de fugas (en ejecución).",
    "Verificación externa": "No verificado.",
    "Compensaciones": "No se informan.",
    "Persona de contacto": "Responsable de Ambiente — ambiente@ejemplo.com.ar (caso ficticio)",
}

if __name__ == "__main__":
    generar("Matriz_GEI_Mendoza_EJEMPLO.xlsx", {"empresa": EMPRESA, "instalaciones": INSTALACIONES,
                                                 "actividades": ACTIVIDADES, "produccion": PRODUCCION,
                                                 "informe": INFORME})
    generar("Matriz_GEI_Mendoza.xlsx")
    print("Listo: Matriz_GEI_Mendoza_EJEMPLO.xlsx (caso ficticio) y Matriz_GEI_Mendoza.xlsx (vacía)")
