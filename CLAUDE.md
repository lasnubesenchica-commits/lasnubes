# Las Nubes — Notas para Claude

## Preferencias del usuario

- **Auto-merge de PRs**: cuando completes un PR (after push + create), mergéalo directamente sin pedir confirmación. El usuario quiere flujo continuo. Para cambios destructivos o riesgosos (force push, drop tablas, etc.) sigue confirmando.

## Arquitectura

- **Frontend**: `dashboard.html` (single-file SPA), servido vía GitHub Pages desde `main`.
- **Backend**: Google Apps Script. El código vive en `apps-script/` y se auto-deploya al mergear a `main` (workflow `.github/workflows/deploy-gas.yml`).
- **Bootstrap del Apps Script**: workflow manual `pull-gas.yml` baja el código actual del Apps Script al repo en una rama `bot/pull-gas-<run_id>`.
- **Web App URL**: hardcodeada en `dashboard.html` como `SHEETS_API_URL`.

## Despliegue

- Cualquier push a `main` que toque `apps-script/**` dispara `Deploy Google Apps Script`. Sube los archivos, crea versión nueva y actualiza el deployment de producción.
- Después de cambios de schema (columnas nuevas en Sheets), recordar al usuario ejecutar manualmente la función de migración correspondiente desde el editor de Apps Script.

## Stack relevante

- Hoja: 24 columnas (la #24 es `Telefono`, agregada en `migrarColumnasV3`).
- Modelo IA: `claude-opus-4-7` para OCR de vouchers/facturas (en `parseVoucherWithClaude` y `parseFacturaEgresoConClaude`).
- API key Anthropic: en Script Properties (`CLAUDE_API_KEY`), nunca en código cliente.
