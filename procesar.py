#!/usr/bin/env python3
"""
Uso:  python3 procesar.py Matriz_GEI_Mendoza_EJEMPLO.xlsx [carpeta_salida]

Lee la matriz, calcula y deja en la carpeta de salida:
  - un Anexo I oficial completo por año × industria × departamento,
  - Anexo_II_Tablas.xlsx (Tabla 1 por año + Tabla 5),
  - Emisiones_detalle.csv (todo el cálculo, fila por fila).
"""

import re
import sys
import unicodedata
from pathlib import Path

from gei import anexo1, anexo2, matriz, motor

TEMPLATE = Path(__file__).parent / "plantilla_oficial" / "Template-Planilla-de-Emisiones-de-GEI-Definitivo-V2.xlsx"


def slug(t):
    return re.sub(r"[^A-Za-z0-9]+", "_", unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode())[:40].strip("_")


def procesar(ruta_matriz, salida):
    salida = Path(salida)
    salida.mkdir(parents=True, exist_ok=True)
    datos = matriz.leer(ruta_matriz)
    errores = motor.validar(datos)
    if errores:
        raise SystemExit("Hay que corregir la matriz:\n  - " + "\n  - ".join(errores))
    em = motor.calcular(datos)
    archivos = []
    for (anio, industria, depto), g in em.groupby(["Año", "Industria", "Departamento"]):
        nombre = salida / f"AnexoI_{anio}_{slug(industria)}_{slug(depto)}.xlsx"
        n = anexo1.exportar(TEMPLATE, nombre, datos["empresa"], anio, industria, depto, datos["instalaciones"], g)
        archivos.append((nombre.name, n))
    anexo2.exportar(em, salida / "Anexo_II_Tablas.xlsx", datos["empresa"])
    em.to_csv(salida / "Emisiones_detalle.csv", index=False, encoding="utf-8-sig")
    return em, archivos


if __name__ == "__main__":
    em, archivos = procesar(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "salidas")
    print(motor.resumen(em).round(1).to_string(index=False))
    for n, f in archivos:
        print(f"  {n}: {f} filas")
