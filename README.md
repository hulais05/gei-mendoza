# GEI Mendoza — Minería e Hidrocarburos

Herramienta de ESG Consulting para el **Registro Provincial de Emisiones GEI de Mendoza**
(Dec. 758/25, Dec. 1338/26, Res. 1/2026). Entrada: consumos de la empresa. Salida: lo que se presenta.

```
Matriz_GEI_Mendoza.xlsx  ──►  motor (factores SNIGEI-AR / IPCC, PCG del template)
   (consumos por año)            │
                                 ├─► Anexo I oficial completo (1 por año × industria × departamento)
                                 ├─► Anexo II: Tabla 1 por año + Tabla 5 (seguimiento)
                                 └─► Emisiones_detalle.csv  (y, próximamente, el tablero Streamlit)
```

## Uso

```bash
python3 caso_ejemplo.py                                  # crea la matriz vacía y la del caso ficticio
python3 procesar.py Matriz_GEI_Mendoza_EJEMPLO.xlsx salidas
```

## Decisiones importantes

- **El Anexo I se escribe editando el XML**, no con openpyxl: openpyxl borra las listas
  desplegables del template al guardar. Solo cambian `sheet2.xml` (celdas verdes) y
  `workbook.xml` (recalcular al abrir). Verificado: las 7 validaciones quedan intactas.
- **Los PCG son los del template** (HFC-32 = 770, CH4 venteado = 29,8…), así nuestro tCO2e
  coincide con el que calcula la planilla del gobierno. Verificado con LibreOffice: diferencia < 0,001 t.
- **Un archivo por industria y por departamento** (Instructivo B10 del template).
- **Factor de red**: el último oficial es 2022 (0,283 kg CO2/kWh, SNIGEI-AR 2025 Tabla 5).
  Para 2024/2025 se usa ese y se marca "verificar"; se reemplaza en la hoja PARAMETROS.
- **CO2 biogénico** del corte de biodiésel/bioetanol: se informa aparte, no suma.
- **Refrigerantes en mezcla** (R-410A, R-404A, R-407C) se reportan por gas puro, porque el
  template solo trae gases puros.

## Preguntas abiertas para la autoridad (capacitación 06-11-2026)

1. Exploración minera: no tiene categoría en la lista de industrias. ¿"Servicios relacionados con minería"?
2. Una calera con cantera: ¿"Minería no metálica" o "Fabricación de cal y productos de yeso" (2A2)?
3. Compensaciones: el template las resta dentro del inventario; el Anexo II las pide aparte (sección 7).
4. Factor de red 2024 y 2025: ¿qué valor de CAMMESA esperan?
5. Si una empresa supera 42 filas en un archivo: ¿agrupar o se puede extender la tabla?

## Datos

El caso de ejemplo es **ficticio**. Con datos de clientes reales, correr **solo en local**.
