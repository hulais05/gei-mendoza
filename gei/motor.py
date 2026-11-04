"""
Motor de cálculo: dato de actividad → kg de cada gas → tCO2e.

Una sola fuente de verdad: el tablero, el Anexo I y el Anexo II leen de acá.
"""

import pandas as pd

from .factores import (FE_RED, FE_RED_ULTIMO_OFICIAL, INDUSTRIAS, METODOS, PCG_TEMPLATE,
                       POR_CODIGO, UMBRAL_TCO2E)

INFORMATIVO = "CO2 biogénico (informativo)"


def factor_red(anio, parametros):
    """Factor de red del año: el que cargó el usuario en PARAMETROS o, si no, el último oficial."""
    if anio in parametros.get("fe_red", {}):
        return parametros["fe_red"][anio], "cargado por el usuario"
    base = FE_RED.get(anio, FE_RED[FE_RED_ULTIMO_OFICIAL])
    origen = "oficial" if anio in FE_RED else f"⚠ se usa {FE_RED_ULTIMO_OFICIAL} (último oficial), verificar"
    return base, origen


def validar(datos):
    """Devuelve una lista de problemas encontrados en la matriz (vacía si está todo bien)."""
    errores = []
    inst = datos["instalaciones"].set_index("Instalación")
    for i, r in datos["actividades"].iterrows():
        fila = i + 2
        if r["Código"] not in POR_CODIGO:
            errores.append(f"ACTIVIDADES fila {fila}: código «{r['Código']}» no existe en el catálogo.")
        if r["Instalación"] not in inst.index:
            errores.append(f"ACTIVIDADES fila {fila}: la instalación «{r['Instalación']}» no está en INSTALACIONES.")
        if r["Método"] not in METODOS:
            errores.append(f"ACTIVIDADES fila {fila}: método «{r['Método']}» (usar N1, N2, N3, MI o MD).")
        if pd.isna(r["Cantidad"]) or r["Cantidad"] < 0:
            errores.append(f"ACTIVIDADES fila {fila}: cantidad vacía o negativa.")
    for nombre, r in inst.iterrows():
        if r["Industria (template)"] not in INDUSTRIAS:
            errores.append(f"INSTALACIONES «{nombre}»: industria no reconocida por el template.")
    return errores


def calcular(datos):
    """Devuelve un DataFrame con una fila por actividad y gas."""
    inst = datos["instalaciones"].set_index("Instalación")
    filas = []
    for _, r in datos["actividades"].iterrows():
        act = POR_CODIGO[r["Código"]]
        anio = int(r["Año"])
        params = {nombre: (r.get(f"Parámetro {k + 1}") if pd.notna(r.get(f"Parámetro {k + 1}")) else defecto)
                  for k, (nombre, defecto, _) in enumerate(act.parametros)}
        fe_red, origen_red = factor_red(anio, datos.get("parametros", {}))
        gases = act.calcular(float(r["Cantidad"]), params, anio, fe_red)
        i = inst.loc[r["Instalación"]]
        factor_txt = act.factor_txt
        if act.codigo == "ELECTRICIDAD":
            factor_txt = f"{fe_red['CO2']:.3f} kg CO2/kWh; {fe_red['CH4']:.7f} kg CH4/kWh; {fe_red['N2O']:.7f} kg N2O/kWh ({origen_red})"
        if act.parametros:
            factor_txt += " | " + "; ".join(f"{k} = {v:g}" for k, v in params.items())
        for gas, kg in gases.items():
            informativo = gas == INFORMATIVO
            filas.append({
                "Año": anio,
                "Instalación": r["Instalación"],
                "Industria": i["Industria (template)"],
                "Sector": "Hidrocarburos" if INDUSTRIAS[i["Industria (template)"]][0] == "Sector Energía" else "Minería",
                "Departamento": i["Departamento"],
                "Proceso": r["Proceso / equipo"],
                "Código": act.codigo,
                "Actividad": act.nombre,
                "Alcance": act.alcance,
                "Clasificación": act.clasificacion,
                "Categoría": act.categoria_anexo2,
                "Gas": gas,
                "kg": kg,
                "tCO2e": 0.0 if informativo else kg * PCG_TEMPLATE[gas] / 1000,
                "Informativo": informativo,
                "Método": r["Método"],
                "Dato de actividad": f"{float(r['Cantidad']):,.0f} {act.unidad}".replace(",", "."),
                "Fuente del dato": r.get("Fuente del dato", ""),
                "Factor": factor_txt,
                "Fuente del factor": act.fuente_fe,
            })
    return pd.DataFrame(filas)


def resumen(em):
    """Totales por año y sector + alerta de umbral."""
    em = em[~em["Informativo"]]
    t = em.groupby(["Año", "Sector", "Alcance"])["tCO2e"].sum().unstack("Alcance", fill_value=0)
    t.columns = [f"Alcance {c}" for c in t.columns]
    t["Total"] = t.sum(axis=1)
    t["Supera 10.000 t"] = t["Total"] > UMBRAL_TCO2E
    return t.reset_index()
