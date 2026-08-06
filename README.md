# Numbi EEFF

Aplicación para transformar balances de prueba por tercero descargados de Siigo
en una presentación financiera completa. Todo el procesamiento del archivo ocurre
en el navegador: los libros no se cargan ni se almacenan en un servidor.

## Qué prepara

- Resumen ejecutivo e indicadores de capital de trabajo, liquidez, endeudamiento,
  ROE y margen neto
- Estado de resultados y estado de situación financiera
- Estado de cambios en el patrimonio
- Flujo de efectivo indirecto y partidas pendientes de conciliación
- Cartera, proveedores y otros saldos por tercero
- Apertura mensual de gastos y costos
- Mapeo auditable y controles contables por periodo

La clasificación base reproduce la hoja `Setup` de la macro de referencia. Cuando
una cuenta no existe allí, la aplicación utiliza una sugerencia PUC visible. Las
cuentas materiales que requieren criterio contable quedan marcadas para
confirmación manual; el modelo no crea partidas residuales para cuadrar el balance.

## Formato de entrada

Se aceptan uno o varios archivos `.xlsx` con la estructura del reporte **Balance
de prueba por tercero** de Siigo. La herramienta detecta empresa, NIT y periodo,
y procesa exclusivamente las filas `Auxiliar / Sí` para evitar doble conteo.

Los periodos duplicados se reemplazan y los archivos de un NIT distinto se
rechazan dentro de la misma sesión.

## Desarrollo

Requiere Node.js `>=22.13.0`.

```bash
npm ci
npm run dev
```

Validación:

```bash
npm run lint
npm test
```

El código principal está en `app/`; la lectura, clasificación y cálculo se
encuentran en `app/lib/financials.ts`, y la tabla estática procedente de la macro
en `app/lib/macro-mapping.ts`.

## Scripts conservados

Los scripts del proyecto anterior continúan disponibles para inspección y
generación de libros independientes. Buscan fuentes en `data/` o en las rutas
indicadas por `EEFF_DATA_DIR` y `EEFF_GUIDE_PATH`. Las fuentes, salidas y archivos
temporales no se versionan.
