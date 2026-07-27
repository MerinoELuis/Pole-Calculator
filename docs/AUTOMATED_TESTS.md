# Automated tests

The application remains a static GitHub Pages site with no runtime build step. Node is used only for dependency-free development checks.

## Run everything

```bash
npm test
```

The runner:

1. Executes `node --check` for every file under `js/`.
2. Parses `package.json`, schemas, fixtures and expected JSON files.
3. Runs every `tests/**/*.test.js` file in natural filename order.
4. Returns a non-zero exit code when any check fails.

## Syntax only

```bash
npm run test:syntax
```

## Fixture policy

- Input states belong in `tests/fixtures/`.
- Complete expected exports belong in `tests/expected/`.
- A behavior change must update both fixture and golden output only when the new behavior is explicitly approved.
- UG, PCO, normal aerial, geometry-only, service, DG and multi-span cases should remain represented.

## CI

`.github/workflows/tests.yml` runs `npm test` on pushes and pull requests using Node 24. GitHub Pages continues serving the source files directly.
