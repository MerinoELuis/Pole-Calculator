# Pole Span MR Calculator

Pole Span MR Calculator es una aplicacion web estatica para revisar postes, spans, comunicaciones, midspans, clearances y make ready a partir de archivos Excel de levantamiento. Esta pensada para correr directamente en GitHub Pages con HTML, CSS y JavaScript puro.

La calculadora permite importar datos crudos, editar alturas de comunicaciones existentes, proponer nuevas alturas por span, recalcular midspans afectados por movimientos en ambos extremos del span y generar un MR unificado por poste.

## Funciones principales

- Importa postes desde la hoja `Collection`.
- Importa relaciones entre postes desde la hoja `Span`.
- Importa cables, owners, attachment heights y midspans desde `Span.Wire`.
- Muestra todos los postes en una vista editable.
- Calcula alturas maximas contra Low Power.
- Recalcula midspans cuando se suben o bajan comms en cualquiera de los postes conectados.
- Distingue comms con midspan real de comms que solo sirven como referencia del span conectado.
- Valida clearances de power-comm, comm-comm y bolt-bolt.
- Genera MR por poste a partir de movimientos y proposed.
- Exporta e importa JSON para guardar y continuar el avance.

## Flujo de trabajo

1. Importar el Excel crudo del levantamiento.
2. Revisar postes, spans, power y comms importados.
3. Editar `Low Power`, alturas de comms existentes o proposed por span.
4. Revisar midspans recalculados, warnings y clearances.
5. Revisar el MR generado por poste.
6. Exportar JSON para guardar el avance.
7. Importar ese JSON despues para continuar.

## Hojas del Excel usadas

### Collection

Se usa para crear postes y leer datos generales:

- Pole ID o nombre del poste.
- Altura del poste.
- Low Power.

La columna de Low Power se busca de forma flexible. La app reconoce nombres que contengan `Low Power Attachment`, y tambien fallbacks como `Lowest Power` o `Low Power`.

### Span

Se usa para crear las conexiones entre postes:

- Span ID.
- Poste origen.
- Poste conectado.
- Longitud.
- Bearing.
- Environment.

El bearing se convierte a direccion cardinal para ayudar a leer las relaciones entre postes.

### Span.Wire

Se usa para crear comms y power por span:

- `Owner`.
- `Attachment Height.display`.
- `Mid Span Height.display`.
- `Wire Id`.
- `Size`.
- `Construction`.
- `Insulator`.

El campo Owner/Comm visible sale de la columna `Owner`.

## Reglas de calculo

### Movimientos de comms

Cada comm existente puede tener un `Cambio de HOA`. Cuando se cambia una altura en un extremo del span, el midspan del tramo se recalcula usando la mitad del movimiento.

Ejemplo:

- Si un comm baja de `20'` a `19'`, el movimiento es `-12"` y el midspan baja `6"`.
- Si el comm del otro poste sube de `20'` a `21'`, el movimiento es `+12"` y el midspan sube `6"`.
- Si ambos extremos se mueven, ambos efectos se aplican sobre el midspan.

Los comms que no traen midspan propio pueden mostrarse como `REF`. Eso indica que pertenecen al span conectado, pero el midspan real viene del otro extremo del tramo.

### Proposed por span

La seccion `Proposed por span` se usa solo para spans que tienen midspan real. No repite backspans ni duplica una misma conexion.

El end drop se calcula con el Proposed local y el cambio del otro extremo cuando existe.

### Clearances

Los valores editables de clearances controlan los calculos:

- `Pole · Power-comms`.
- `Pole · Comm-comm`.
- `Pole · Bolt-bolt`.
- `Midspan · Power-comm`.
- `Midspan · Comm-comm`.

Entre comms del poste:

- Si son owners diferentes, se usa `Pole · Comm-comm`.
- Si son el mismo owner, se usa `Pole · Bolt-bolt`.

Para proposed en el poste, tambien se revisa que no quede dentro de una zona no permitida entre bolts existentes. Si los attachments estan demasiado cerca, no se permite colocar proposed en medio de ellos; si hay suficiente separacion, se permite una altura que respete `Pole · Bolt-bolt`.

### Low Power en midspan

La app calcula la altura maxima de comm en midspan usando el power mas bajo del span y el clearance `Midspan · Power-comm`.

Cuando el ajuste de proposed se debe a Low Power en midspan, el MR agrega:

`Ensure min 30" to low power at midspan.`

Si el ajuste ya no es necesario, o si el ajuste fue por comm-comm y no por Low Power, esa nota no se genera.

### Tabla de reglas

| Regla | Valor por defecto | Se aplica a | Resultado |
| --- | --- | --- | --- |
| `Pole · Power-comms` | `40"` | comms y proposed en el poste | Define `Max Height on Pole`. |
| `Pole · Comm-comm` | `12"` | comms de owners diferentes en el poste | Evita comms demasiado juntos en el poste. |
| `Pole · Bolt-bolt` | `4"` | comms del mismo owner y proposed entre sí en el poste | Evita bolts demasiado cercanos. |
| `Midspan · Power-comm` | `30"` | comms y proposed en midspan | Define `Max Height at MS`. |
| `Midspan · Comm-comm` | `4"` | comms y proposed en midspan | Mantiene separación vertical entre cables. |
| `Environment` | según span | comms y proposed en midspan | Evita quedar por debajo del mínimo del entorno. |
| Orden de comms | según alturas | comms existentes del mismo span | Evita que dos comms se crucen entre poste y midspan. |
| Posición `Top Comm` / `Low Comm` | configurable | proposed en poste y midspan | Obliga al proposed a quedar del lado configurado de los comms existentes. |

Para `Proposed` en el poste:

- Siempre debe conservar `Pole · Comm-comm` frente a comms existentes.
- Puede compartir exactamente la misma altura con un attachment existente.
- Entre proposeds distintos del mismo poste se aplica `Pole · Bolt-bolt`.

## Archivos principales

- `index.html`
- `css/styles.css`
- `js/app.js`
- `js/state.js`
- `js/height-utils.js`
- `js/excel-import.js`
- `js/excel-export.js`
- `js/json-export.js`
- `js/calculations.js`
- `js/mr-logic.js`
- `js/validations.js`
- `js/floating-calculator.js`
- `libs/xlsx.full.min.js`

## Uso local

Abre `index.html` en el navegador o publica la carpeta en GitHub Pages.

La aplicacion no requiere backend, servidor propio, Node.js, npm, build step ni frameworks.
