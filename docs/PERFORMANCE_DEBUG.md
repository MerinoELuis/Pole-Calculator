# Performance debug

Performance tracing starts automatically for the current browser session. It is
kept in memory and is copied into the local `Export Debug` JSON as
`performanceTrace`; it is not added to normal saved job data.

To clear the current session samples and start a new measurement window, run:

```js
PoleCalcPerformance.enable();
```

Use the page normally, especially the action that feels slow, then run:

```js
PoleCalcPerformance.report();
```

The report includes count, total, maximum, and last duration for:

- `calc.recalculateAll`
- `ui.render`
- `ui.renderAllPolesWorkspace`
- `ui.renderPoleWorkspace`

To see each individual sample in the console while reproducing the issue:

```js
PoleCalcPerformance.enable({ verbose: true });
```

The most recent 200 samples are retained in `PoleCalcPerformance.snapshot().recent`.
Disable and clear the tracer with:

```js
PoleCalcPerformance.disable();
PoleCalcPerformance.reset();
```
