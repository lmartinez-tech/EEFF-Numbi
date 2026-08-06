# Numbi EEFF

Aplicación para convertir balances de prueba por tercero en estados financieros,
controles y análisis listos para revisar. La interfaz permite cargar archivos
Excel, clasificar cuentas, consultar resultados por periodo y descargar datos.

## Requisitos

- Node.js `>=22.13.0`
- Python 3 con `openpyxl` para los scripts de análisis de libros
- El runtime de hojas de cálculo de Codex para los scripts `.mjs` de generación
  y renderizado

## Aplicación web

```bash
npm ci
npm run dev
```

La compilación de producción se valida con:

```bash
npm run build
```

El código principal está en `app/`; la lógica de lectura, mapeo y cálculo de los
balances está en `app/lib/financials.ts`.

## Pipeline de estados financieros

Los scripts conservados desde el proyecto original trabajan en este orden:

1. `extract_structure.py` y `extract_ooxml.py` inspeccionan la estructura de los
   libros fuente.
2. `prepare_data.py` normaliza saldos, terceros, periodos y mapeos en
   `analysis_artifacts/normalized_data.json`.
3. `build_eeff.mjs` genera el libro final dentro de `outputs/`.
4. `render_final_check.mjs` produce las vistas de control del resultado.

Los libros fuente no se versionan. Por defecto, los scripts los buscan en
`data/`; también puedes indicar otra carpeta con `EEFF_DATA_DIR` y la plantilla
de referencia con `EEFF_GUIDE_PATH`.

## Contenido migrado

- Aplicación web completa y configuración de Sites
- Scripts de extracción, normalización, construcción e inspección
- Datos normalizados usados por la última generación
- Libro `EEFF_balance_por_tercero.xlsx`, controles y vistas previas en `outputs/`

Las dependencias, cachés, temporales, identificadores de procesos y volcados de
inspección regenerables no forman parte del proyecto versionado.
