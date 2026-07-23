# Las Nubes — Notas para Claude

## Preferencias del usuario

- **Auto-merge de PRs**: cuando completes un PR (after push + create), mergéalo directamente sin pedir confirmación. El usuario quiere flujo continuo. Para cambios destructivos o riesgosos (force push, drop tablas, etc.) sigue confirmando.
- **Español neutral latinoamericano** en todos los textos visibles al cliente final (Agente de WhatsApp, plantillas, emails, dashboard). NADA de voseo rioplatense: usar `tú`/`tienes`/`puedes`/`dime`/`toca`/`escríbeme` — NO `vos`/`tenés`/`podés`/`decime`/`tocá`/`escribime`. Aplica a código nuevo y a strings que se reescriban; código viejo con voseo se va migrando a medida que se toque.

## Arquitectura

- **Frontend**: `dashboard.html` (single-file SPA), servido vía GitHub Pages desde `main`.
- **Backend**: Google Apps Script. El código vive en `apps-script/` y se auto-deploya al mergear a `main` (workflow `.github/workflows/deploy-gas.yml`).
- **Bootstrap del Apps Script**: workflow manual `pull-gas.yml` baja el código actual del Apps Script al repo en una rama `bot/pull-gas-<run_id>`.
- **Web App URL**: hardcodeada en `dashboard.html` como `SHEETS_API_URL`.

## Despliegue

- Cualquier push a `main` que toque `apps-script/**` dispara `Deploy Google Apps Script`. Sube los archivos, crea versión nueva y actualiza el deployment de producción.
- **Límite de 200 versiones**: Apps Script topa en 200 versiones por proyecto y no se borran por API (solo a mano en el editor). Si se llena, el deploy falla con `RESOURCE_EXHAUSTED`; el usuario limpia versiones viejas desde el editor y re-corre el deploy.
- Después de cambios de schema (columnas nuevas en Sheets), recordar al usuario ejecutar manualmente la función de migración correspondiente desde el editor de Apps Script.

## Malaya Lodge (cabaña referida)

- **No es parte del inventario de Las Nubes**: es una cabaña del cliente Celestino (+507 6542-9927 / malayalodge@gmail.com). Yo cierro reservas como referido y cobro comisión ($10 dom-jue / $20 vie-sáb por noche).
- **Landing público**: `malaya.html` (servido por GitHub Pages como `/malaya.html`). Calendario standalone, no se conecta al sistema de Las Nubes.
- **Sync iCal bidireccional**:
  - **Airbnb → Las Nubes** (pull): trigger `syncMalayaAirbnb` cada 30 min lee el `.ics` público de Airbnb (URL en Script Property `MALAYA_AIRBNB_ICAL`) y refleja los bloqueos en una hoja interna `MalayaIcal`. Cross-checkea las reservas directas pendientes: si pasaron > `MALAYA_GRACE_MINUTES` (default 60) sin que Celestino bloquee en Airbnb → estado `no_bloqueada` y alerta WhatsApp/email al admin.
  - **Las Nubes → Airbnb** (push, vía polling de Airbnb): endpoint público `doGet?action=malayaIcal` en el Web App expone las directas activas como `.ics`. Celestino importa esa URL en Airbnb → Sincronizar calendarios; Airbnb la polla cada pocas horas y bloquea las noches automáticamente. Función en Malaya.gs: `getMalayaIcalFeed()`. Test desde editor: `_testMalayaIcalFeed()`.
- **Hoja `Malaya`**: 14 columnas (id, huésped, teléfono, checkin, checkout, noches, personas, monto total, comisión, origen, estado, airbnb_blocked, fecha_reserva, notas). Estados: `pendiente | confirmada | no_bloqueada | cancelada | completada`.
- **Setup inicial** (correr una vez en el editor): setear Script Properties `MALAYA_AIRBNB_ICAL`, `MALAYA_CELESTINO_PHONE`, `MALAYA_GRACE_MINUTES`, después correr `instalarTriggersMalaya()`.
- **Notificación WA a Celestino con redundancia**: la plantilla `malaya_reserva_celestino` se envía a Celestino y opcionalmente a más contactos (Glorimar, etc.) para reducir el riesgo de que el aviso pase desapercibido. Los extras se configuran en Script Property `MALAYA_EXTRA_NOTIFY_PHONES` como CSV (E.164 sin `+`). Default en código: `50761000079` (Glorimar). El fallback forward-manual al admin sigue disparándose sólo si ningún destinatario recibe la plantilla.
- **Flujo operacional**: admin entra a `malaya.html?admin=1`, selecciona rango en calendario, sube voucher (parseado con Claude, mismo que Las Nubes), ingresa WhatsApp del huésped, click "Bloquear reserva". Tras eso, avisa manualmente por WhatsApp a Celestino y al huésped. Celestino bloquea en Airbnb. Sync verifica. Si Celestino no bloquea → alerta automática.

## Stack relevante

- Hoja `Reservas`: 31 columnas. La #25 es `Tipo` (noche/pasadia/pasadia-largo/early/late, agregada en `migrarColumnasV4`). La #26 es `VoucherURL` (link al voucher en Drive, persistido por `saveVoucherToDrive`). La #27 es `IdHuespedURL`. La #28 es `FechaNacimiento`. La #29 es `CheckoutExtendido` (boolean, cortesía 12:30pm). Las #30 (`HoraEntrada`) y #31 (`HoraSalida`) son overrides horarios opcionales por reserva en formato `HH:MM` (24h); vacío = usa el default del tipo. Cuando están seteados, pisan tanto el default del tipo como `CheckoutExtendido`. `getOrCreateSheet` auto-asegura todas las columnas en cada llamada.
- Hoja `Egresos`: 9 columnas (ID, Fecha, Descripcion, Monto, Categoria, Cabaña, Proveedor, URLFoto, `Item`). La #9 `Item` es para el seguimiento de Suministros (agua, hielo, gas…): solo se persiste cuando la categoría es `Suministros`. `_getEgresoSheetEnsured` (Parser.gs) auto-migra hojas viejas de 8 columnas agregando el header `Item`. La sub-tab **Suministros** en Contabilidad agrupa las compras por `Item` y calcula la cadencia de recompra ("Dura ~X días") con alerta cuando toca recomprar.
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
- **`SIN_PAGO_ORIGINS`**: constante global en dashboard.html con `['Cortesia','Colaboracion','Personal','Mantenimiento','Abierta']`.
- **Origen `'Mantenimiento'`**: bloqueo interno por trabajo de mantenimiento. En el calendario sale 🔧 en vez de iniciales y la celda usa fondo gris (`.disp-day.occupied.maintenance`).
- **Recibo PDF**: generado por `generateReceiptPDF(r)` en backend. Numeración correlativa `LN-NNNN` en Script Properties (`RECEIPT_COUNTER`). Adjuntado al email de confirmación si la reserva tiene voucher. `sendUpdateEmail` también lo adjunta cuando la edición incluye un voucher nuevo.
- **Vouchers en Drive**: subidos por `saveVoucherToDrive` a la carpeta `Las Nubes - Pagos`. URL persistido en columna 26.
- **Recordatorio de check-in** (email automático 1 día antes): trigger `enviarRecordatoriosCheckin` corre diario a las 10am Panamá. Escanea Reservas y manda email a quienes tienen `displayCheckin === mañana` (excluyendo CANCELADA, origen Abierta, sin email). Para activarlo, correr una vez desde el editor: `instalarTriggerRecordatorios()`. Para preview: `enviarRecordatorioPrueba()` (envía al email del usuario que corre el script). Configuración (indicaciones, maps URL) en Script Properties: `CHECKIN_MAPS_URL`, `CHECKIN_WAZE_URL`, `CHECKIN_INDICACIONES`, `CHECKIN_ACCESO_EXTRA`.
- **Multi-voucher por reserva** (abonos parciales): cuando se sube un segundo voucher en edición, el sistema acumula:
  - `montoVoucher`: suma total ($45 + $45 = $90).
  - `codTransferencia`: lista separada por `|` (ej. `UYAFL-111|UYAFL-222`).
  - col 26 `VoucherURL`: lista separada por `|` (cada subida agrega).
  - `deposit`: si ya hay abono previo y se sube otro voucher, `handleVoucherUpload` lo suma al campo en lugar de sobrescribir.
  El botón "Ver voucher" del detalle abre cada URL en pestaña separada si hay múltiples. `deleteReservation` con `deleteVoucher=true` itera todas las URLs.
