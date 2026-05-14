# Calculadora Web de Postes, Spans y Midspans

Aplicación estática para GitHub Pages hecha con HTML, CSS y JavaScript puro. No requiere Node.js, backend ni servidor propio.

## Cambio v1.2

Esta versión corrige la importación y separa dos conceptos que no deben mezclarse:

1. **Proposed por span/lado del poste**
   - Se guarda en `SpanSides`.
   - Clave lógica: `spanId + poleId`.
   - No depende del owner/comm.
   - Sirve para el proposed que tú vas a poner por span.

2. **Movimiento de comm existente**
   - Se guarda en `SpanComms`.
   - Clave lógica: `spanId + poleId + owner + wireId`.
   - La columna editable es `Existing HOA Change`.
   - El MR se genera unificado por poste usando esos movimientos.

## Importación desde Excel original

La app lee principalmente:

- `Collection`
  - `Id`
  - `Sequence`
  - `Type`
  - `Tip.display`
  - `Low Power Attachment.display`
  - Si el nombre cambia, busca mínimo columnas que contengan `Low Power Attachment`.

- `Span`
  - `Id`
  - `Span Id`
  - `Span Index`
  - `Span Length.display`
  - `Span Length.bearing.display`
  - `Type`
  - `Linked Collection.Title`
  - `Linked Collection.ID`

- `Span.Wire`
  - `Span Id`
  - `Wire Id`
  - `Owner`
  - `Size`
  - `Construction`
  - `Insulator`
  - `Attachment Height.display`
  - `Mid Span Height.display`

## Reglas actuales

- Se quitó el enfoque de `Backspan` en la interfaz.
- La dirección del span se deduce con `Span Length.bearing.display`.
- Si un span apunta a otro collection que no está, o no se conoce, se crea un poste editable `Unknown-<spanId>`.
- El Low Power se muestra y se puede editar.
- La Altura Max se calcula como `Low Power - clearanceToPower`.
- El MR no se importa desde `Make Ready`; se genera dentro de la app.
- Las notas importadas no se usan como notas principales; las notas son editables dentro de la app.
- `Exportar Excel` descarga un `.xlsx` reimportable con hojas de tablas y una hoja `AppState` con el JSON completo.
- `Exportar JSON` descarga solo el estado completo de la calculadora para poder continuar el trabajo después.
- `Importar JSON` carga un archivo `.json` exportado por la app y restaura postes, spans, movimientos, MR, notas y warnings guardados.

## Botones de importación/exportación

- `Importar Excel crudo`
  - Acepta `.xlsx` o `.csv`.
  - Se usa para cargar el archivo original del levantamiento.
  - Busca las hojas `Collection`, `Span` y `Span.Wire`.
  - Si el archivo incluye otras hojas como `Make Ready`, se conservan como referencia del Excel crudo, pero la app genera su MR propio desde los movimientos editados.

- `Importar JSON`
  - Acepta `.json`.
  - Se usa para continuar un trabajo previamente exportado desde esta calculadora.
  - No reemplaza al Excel crudo; restaura exactamente el estado guardado por la app.

- `Exportar JSON`
  - Descarga un `.json` con todo el estado actual.
  - Sirve como guardado ligero del avance.

- `Exportar Excel`
  - Descarga un `.xlsx` reimportable.
  - Incluye la hoja `AppState` para restaurar exactamente el avance.
  - También incluye hojas legibles como `Poles`, `Spans`, `SpanComms`, `SpanSides`, `MR` y `Warnings`.

## Archivos principales

- `index.html`
- `css/styles.css`
- `js/app.js`
- `js/state.js`
- `js/height-utils.js`
- `js/excel-import.js`
- `js/excel-export.js`
- `js/calculations.js`
- `js/graph.js`
- `js/midspan.js`
- `js/mr-logic.js`
- `js/validations.js`
- `js/floating-calculator.js`
- `libs/xlsx.full.min.js`

> Nota: `libs/xlsx.full.min.js` no es SheetJS completo. Es un lector/exportador XLSX mínimo creado para esta app. Permite exportar un Excel reimportable y leer hojas simples de Excel desde el navegador. Para workbooks muy complejos, fórmulas, estilos avanzados o formatos especiales, conviene reemplazarlo por SheetJS real.

## Cómo usar en local

Abre `index.html` directamente en el navegador.

## Pendientes / notas técnicas

- El lector XLSX incluido es mínimo y está pensado para archivos tabulares simples como los exportados por IKE Office.
- Si un Excel crudo cambia nombres de hojas o encabezados, hay que agregar esos alias en `js/excel-import.js`.
- Si se quiere volver a exportar reportes tipo tabla, puede agregarse otro botón para CSV/XLSX sin cambiar el flujo de guardado JSON.

