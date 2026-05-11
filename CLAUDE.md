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

- Hoja `Reservas`: 26 columnas. La #25 es `Tipo` (noche/pasadia/pasadia-largo/early/late, agregada en `migrarColumnasV4`). La #26 es `VoucherURL` (link al voucher en Drive, persistido por `saveVoucherToDrive`). `getOrCreateSheet` auto-asegura ambas columnas en cada llamada.
- Modelo IA: `claude-opus-4-6` para OCR de vouchers/facturas (en `parseVoucherWithClaude` y `parseFacturaEgresoConClaude`). `parseVoucherWithClaude` tiene retry con fallback a Sonnet 4.6.
- API key Anthropic: en Script Properties (`CLAUDE_API_KEY`), nunca en código cliente.

## Modelo de datos para tipos de reserva

Cada reserva se almacena con un rango de bloqueo (checkin/checkout) y un `tipo`. El frontend tiene dos representaciones:

- **Storage** (lo que está en la hoja): rango de días bloqueados.
- **Form / display** (lo que ve el usuario): día real del huésped.

| Tipo            | Storage `checkin`    | Storage `checkout`    | Form `entrada`    | Form `salida`     |
|-----------------|----------------------|-----------------------|-------------------|-------------------|
| `noche`         | día llegada          | día salida            | igual             | igual             |
| `pasadia`       | día pasadía          | día pasadía + 1       | día pasadía       | día pasadía       |
| `pasadia-largo` | día pasadía − 1 (cortesía) | día pasadía + 1 | día pasadía       | día pasadía       |
| `early`         | día llegada − 1 (cortesía) | día salida real | día llegada (9am) | día salida (11am) |
| `late`          | día llegada          | día salida + 1 (cortesía) | día llegada (2pm) | día salida (4pm)  |

Helpers en `dashboard.html`:
- `_formFromStored(r)` y `_storedFromForm(entrada, salida, tipo)`: traducen entre representaciones.
- `reservaDisplayDates(r)`: para mostrar en tablas/popups (devuelve `displayCheckin`, `displayCheckout`).
- `reservaNochesReales(r)`: 0 para pasadías, 1 para early/late, N para noche.
- `reservaDuracionLabel(r)`: "Pasadía" / "1 noche" / "N noches" / etc.
- `getDayKind(r, dateStr)`: `'main'` o `'courtesy'`.

Helpers equivalentes en backend (`Parser.gs`):
- `tipoEmailMeta(r)`: para emails de confirmación/actualización.
- `buildGuiaHTML(cabin, tipo)`: ajusta hora de check-out en la guía según tipo.

## Notas operativas

- **Origen `'Cortesia'`** (sin tilde) — toda referencia debe ser sin tilde. La forma con tilde es solo el label visible en UI.
- **`SIN_PAGO_ORIGINS`**: constante global en dashboard.html con `['Cortesia','Colaboracion','Personal','Abierta']`.
- **Recibo PDF**: generado por `generateReceiptPDF(r)` en backend. Numeración correlativa `LN-NNNN` en Script Properties (`RECEIPT_COUNTER`). Adjuntado al email de confirmación si la reserva tiene voucher. `sendUpdateEmail` también lo adjunta cuando la edición incluye un voucher nuevo.
- **Vouchers en Drive**: subidos por `saveVoucherToDrive` a la carpeta `Las Nubes - Pagos`. URL persistido en columna 26.
- **Recordatorio de check-in** (email automático 1 día antes): trigger `enviarRecordatoriosCheckin` corre diario a las 10am Panamá. Escanea Reservas y manda email a quienes tienen `displayCheckin === mañana` (excluyendo CANCELADA, origen Abierta, sin email). Para activarlo, correr una vez desde el editor: `instalarTriggerRecordatorios()`. Para preview: `enviarRecordatorioPrueba()` (envía al email del usuario que corre el script). Configuración (indicaciones, maps URL) en Script Properties: `CHECKIN_MAPS_URL`, `CHECKIN_INDICACIONES`, `CHECKIN_ACCESO_EXTRA`.
- **Multi-voucher por reserva** (abonos parciales): cuando se sube un segundo voucher en edición, el sistema acumula:
  - `montoVoucher`: suma total ($45 + $45 = $90).
  - `codTransferencia`: lista separada por `|` (ej. `UYAFL-111|UYAFL-222`).
  - col 26 `VoucherURL`: lista separada por `|` (cada subida agrega).
  - `deposit`: si ya hay abono previo y se sube otro voucher, `handleVoucherUpload` lo suma al campo en lugar de sobrescribir.
  El botón "Ver voucher" del detalle abre cada URL en pestaña separada si hay múltiples. `deleteReservation` con `deleteVoucher=true` itera todas las URLs.
