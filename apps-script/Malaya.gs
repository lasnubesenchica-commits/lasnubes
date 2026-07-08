/**
 * Malaya Lodge — promoción + tracking de reservas que cierro como referido.
 *
 * Modelo:
 *   - No es mía: la cabaña es del cliente Celestino (+507 6542-9927).
 *   - Yo cierro la reserva por WhatsApp, le paso datos a Celestino, él la
 *     bloquea en Airbnb. Mi comisión: $10 (dom-jue) / $20 (vie-sáb) por noche.
 *   - El calendario público de Malaya se sincroniza con el iCal de Airbnb
 *     cada 30 min para evitar overbooking.
 *
 * Flujo:
 *   1. malaya.html?admin=1: selecciono rango + subo voucher + WA del huésped.
 *   2. saveMalayaReserva → fila en hoja Malaya con estado='pendiente'.
 *   3. Yo aviso a Celestino + huésped manualmente por WhatsApp.
 *   4. Celestino bloquea en Airbnb.
 *   5. syncMalayaAirbnb (cada 30 min): pulla iCal, cross-checkea las
 *      pendientes. Si el iCal incluye las fechas → estado='confirmada'.
 *      Si pasaron > GRACE_MINUTES sin bloqueo → 'no_bloqueada' + alerta WA.
 *
 * Script Properties requeridas:
 *   - MALAYA_AIRBNB_ICAL         : URL .ics del listing en Airbnb.
 *   - MALAYA_CELESTINO_PHONE     : 50765429927 (solo para texto de alertas).
 *   - MALAYA_GRACE_MINUTES       : 60 (opcional, default 60).
 *   - MALAYA_EXTRA_NOTIFY_PHONES : opcional, CSV de teléfonos adicionales
 *       que también reciben la plantilla WA de reserva (redundancia por si
 *       Celestino no la ve). Default: '50761000079' (Glorimar).
 */

const MALAYA_CABIN_NAME    = 'Malaya Lodge';
const MALAYA_TARIFA_LV     = 75;   // domingo a jueves (noche)
const MALAYA_TARIFA_FS     = 100;  // viernes y sábado (noche)
const MALAYA_COMISION_LV   = 10;   // mi comisión dom-jue por noche
const MALAYA_COMISION_FS   = 20;   // mi comisión vie-sáb por noche
const MALAYA_EXTRA_PERSONA = 30;   // cargo por persona extra (3+) por noche
const MALAYA_SHEET_NAME    = 'Malaya';
const MALAYA_ICAL_SHEET    = 'MalayaIcal';

// ─── Sheets ─────────────────────────────────────────────────────

function _malayaSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(MALAYA_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(MALAYA_SHEET_NAME);
    sheet.getRange(1, 1, 1, 16).setValues([[
      'ID', 'Huésped', 'Teléfono', 'Check-in', 'Check-out', 'Noches',
      'Personas', 'Monto total', 'Comisión', 'Origen', 'Estado',
      'Airbnb bloqueado', 'Fecha reserva', 'Notas', 'Voucher URL', 'Email'
    ]]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 16).setFontWeight('bold');
  }
  // Auto-migración: añade cols 15/16 si faltan (hojas creadas antes).
  if (sheet.getLastColumn() < 15) {
    sheet.getRange(1, 15).setValue('Voucher URL').setFontWeight('bold');
  }
  if (sheet.getLastColumn() < 16) {
    sheet.getRange(1, 16).setValue('Email').setFontWeight('bold');
  }
  return sheet;
}

function _malayaIcalSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(MALAYA_ICAL_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(MALAYA_ICAL_SHEET);
    sheet.getRange(1, 1, 1, 4).setValues([[
      'Check-in', 'Check-out', 'Summary', 'UID'
    ]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ─── Helpers de fecha / pricing ─────────────────────────────────

function _malayaAddDays(iso, n) {
  const d = new Date(iso + 'T12:00:00');
  d.setDate(d.getDate() + n);
  return Utilities.formatDate(d, 'America/Panama', 'yyyy-MM-dd');
}

function _malayaNightCount(ciIso, coIso) {
  const a = new Date(ciIso + 'T12:00:00');
  const b = new Date(coIso + 'T12:00:00');
  return Math.round((b - a) / 86400000);
}

// Busca una reserva reciente con MISMO teléfono + mismas fechas dentro de
// una ventana de tiempo (segundos). Se usa como dedupe idempotente en
// saveMalayaReserva para que un retry / double-tap del cliente no cree dos
// filas. Devuelve null si no hay match o el objeto con los campos que la
// respuesta original devolvería (id, noches, montoTotal, comisión, huesped).
function _findRecentMalayaReserva(phone, ciIso, coIso, windowSeconds) {
  const sheet = _malayaSheet();
  if (sheet.getLastRow() < 2) return null;
  const now = Date.now();
  const data = sheet.getDataRange().getValues();
  // Recorremos de abajo hacia arriba: los duplicados son casi siempre las
  // últimas filas insertadas, así que salimos rápido.
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const rowPhone = String(row[2] || '').replace(/\D/g, '');
    if (rowPhone !== phone) continue;
    const rowCi = row[3] instanceof Date ? Utilities.formatDate(row[3], 'America/Panama', 'yyyy-MM-dd') : String(row[3] || '').slice(0,10);
    const rowCo = row[4] instanceof Date ? Utilities.formatDate(row[4], 'America/Panama', 'yyyy-MM-dd') : String(row[4] || '').slice(0,10);
    if (rowCi !== ciIso || rowCo !== coIso) continue;
    const fecha = row[12];
    let ts = null;
    if (fecha instanceof Date) ts = fecha.getTime();
    else if (fecha) {
      const parsed = new Date(String(fecha).replace(' ', 'T') + '-05:00'); // Panamá = UTC-5, sin DST
      if (!isNaN(parsed.getTime())) ts = parsed.getTime();
    }
    if (ts && (now - ts) <= windowSeconds * 1000) {
      return {
        id:         String(row[0] || ''),
        huesped:    String(row[1] || ''),
        noches:     parseInt(row[5], 10) || 0,
        montoTotal: parseFloat(row[7]) || 0,
        comision:   parseFloat(row[8]) || 0
      };
    }
  }
  return null;
}

// Comisión total por reserva: itera cada noche y suma según día de la semana
// del CHECKIN de esa noche. Vie/Sáb → $20, resto → $10.
function _malayaCalculateCommission(ciIso, coIso) {
  let total = 0;
  let current = new Date(ciIso + 'T12:00:00');
  const end   = new Date(coIso + 'T12:00:00');
  while (current < end) {
    const dow = current.getDay();   // 0=dom,5=vie,6=sab
    total += (dow === 5 || dow === 6) ? MALAYA_COMISION_FS : MALAYA_COMISION_LV;
    current.setDate(current.getDate() + 1);
  }
  return total;
}

// Tarifa total: ídem comisión pero con tarifas + cargo por persona extra.
function _malayaCalculateTotal(ciIso, coIso, personas) {
  let total = 0;
  let current = new Date(ciIso + 'T12:00:00');
  const end   = new Date(coIso + 'T12:00:00');
  const extra = Math.max(0, (parseInt(personas, 10) || 2) - 2);
  while (current < end) {
    const dow = current.getDay();
    total += (dow === 5 || dow === 6) ? MALAYA_TARIFA_FS : MALAYA_TARIFA_LV;
    total += extra * MALAYA_EXTRA_PERSONA;
    current.setDate(current.getDate() + 1);
  }
  return total;
}

// ─── iCal parser ────────────────────────────────────────────────
// Parsea el .ics de Airbnb y devuelve un array de { checkin, checkout, uid, summary }.
// Airbnb usa VALUE=DATE: yyyymmdd. La fecha DTEND es EXCLUSIVA (día de checkout).

function _malayaParseIcal(icsText) {
  const events = [];
  const blocks = icsText.split('BEGIN:VEVENT').slice(1);
  blocks.forEach(block => {
    const endIdx = block.indexOf('END:VEVENT');
    const chunk  = endIdx === -1 ? block : block.slice(0, endIdx);
    const dtStart = chunk.match(/DTSTART[^:]*:(\d{8})/);
    const dtEnd   = chunk.match(/DTEND[^:]*:(\d{8})/);
    const uid     = chunk.match(/UID:([^\r\n]+)/);
    const summary = chunk.match(/SUMMARY:([^\r\n]+)/);
    if (!dtStart || !dtEnd) return;
    const toIso = s => s.slice(0,4) + '-' + s.slice(4,6) + '-' + s.slice(6,8);
    events.push({
      checkin:  toIso(dtStart[1]),
      checkout: toIso(dtEnd[1]),
      uid:      uid ? uid[1].trim() : '',
      summary:  summary ? summary[1].trim() : 'Reserved'
    });
  });
  return events;
}

// ─── Sync trigger ───────────────────────────────────────────────

function syncMalayaAirbnb() {
  const props = PropertiesService.getScriptProperties();
  const url   = props.getProperty('MALAYA_AIRBNB_ICAL');
  if (!url) { logDebugEntry('malaya-sync-no-url', {}); return; }
  const graceMin = parseInt(props.getProperty('MALAYA_GRACE_MINUTES'), 10) || 60;

  // Cache-bust: añadimos un timestamp para forzar a Airbnb a servir fresco.
  const sep = url.indexOf('?') === -1 ? '?' : '&';
  const bustUrl = url + sep + '_=' + Date.now();
  let events = [];
  try {
    const res = UrlFetchApp.fetch(bustUrl, { muteHttpExceptions: true });
    if (res.getResponseCode() < 200 || res.getResponseCode() >= 300) {
      logDebugEntry('malaya-sync-FETCH-FAIL', { code: res.getResponseCode() });
      return;
    }
    events = _malayaParseIcal(res.getContentText());
  } catch(e) {
    logDebugEntry('malaya-sync-FETCH-EXCEPTION', { error: e.message });
    return;
  }

  // Sobrescribir hoja MalayaIcal con el snapshot actual.
  const icalSheet = _malayaIcalSheet();
  if (icalSheet.getLastRow() > 1) {
    icalSheet.getRange(2, 1, icalSheet.getLastRow() - 1, 4).clearContent();
  }
  if (events.length) {
    const rows = events.map(e => [e.checkin, e.checkout, e.summary, e.uid]);
    icalSheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }

  // Cross-check: para cada reserva pendiente, marcar como confirmada si el
  // iCal la incluye, o no_bloqueada si pasó el grace period.
  const malayaSheet = _malayaSheet();
  const data = malayaSheet.getDataRange().getValues();
  let transitioned = 0;
  const alerts = [];
  const nowMs = Date.now();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const estado = String(row[10] || '').toLowerCase();
    if (estado !== 'pendiente' && estado !== 'no_bloqueada') continue;
    const origen = String(row[9] || '');
    if (origen !== 'Directa') continue;   // solo verificamos las mías
    const ci = row[3] instanceof Date ? Utilities.formatDate(row[3], 'America/Panama', 'yyyy-MM-dd') : String(row[3] || '').slice(0,10);
    const co = row[4] instanceof Date ? Utilities.formatDate(row[4], 'America/Panama', 'yyyy-MM-dd') : String(row[4] || '').slice(0,10);
    if (!ci || !co) continue;
    const blockedByIcal = events.some(e => e.checkin <= ci && co <= e.checkout);
    if (blockedByIcal) {
      if (estado !== 'confirmada') {
        malayaSheet.getRange(i + 1, 11).setValue('confirmada');
        malayaSheet.getRange(i + 1, 12).setValue(true);
        transitioned++;
      }
    } else {
      // No está bloqueada. ¿Pasó el grace period?
      const reservadaTs = row[12] instanceof Date ? row[12].getTime() : Date.parse(String(row[12] || ''));
      const minutos = reservadaTs ? (nowMs - reservadaTs) / 60000 : 0;
      if (minutos > graceMin && estado !== 'no_bloqueada') {
        malayaSheet.getRange(i + 1, 11).setValue('no_bloqueada');
        alerts.push({
          id: row[0], guest: row[1], phone: row[2],
          checkin: ci, checkout: co, minutos: Math.round(minutos)
        });
      }
    }
  }

  // Alertas (una por reserva que transicionó a no_bloqueada en esta corrida).
  alerts.forEach(a => _malayaAlertNoBloqueada(a));

  logDebugEntry('malaya-sync', {
    icalEvents: events.length, transitioned: transitioned, alerts: alerts.length
  });
}

function _malayaAlertNoBloqueada(a) {
  const celestino = PropertiesService.getScriptProperties().getProperty('MALAYA_CELESTINO_PHONE') || '50765429927';
  const msg =
    '⚠️ *Malaya — falta bloqueo en Airbnb*\n\n' +
    'Tu reserva directa de Malaya pasó ' + a.minutos + ' min sin que aparezca bloqueada en el iCal de Airbnb.\n\n' +
    '👤 ' + a.guest + '\n' +
    '📱 +' + a.phone + '\n' +
    '📅 ' + a.checkin + ' → ' + a.checkout + '\n\n' +
    'Avisa a Celestino (+' + celestino + ') para que la bloquee en Airbnb. Si no, hay riesgo de doble booking.';
  try { sendWhatsAppText(BOT_ADMIN_PHONE, msg); } catch(_) {}
  // Email backup, por las dudas.
  try {
    GmailApp.sendEmail(REPLY_TO_EMAIL,
      '⚠️ Malaya: falta bloqueo en Airbnb — ' + a.guest,
      msg,
      { name: 'Las Nubes Agente' });
  } catch(_) {}
  logDebugEntry('malaya-alert-no-bloqueada', { id: a.id, guest: a.guest });
}

// ─── Verificación diaria 11am (digesto) ─────────────────────────

function verificarMalayaPendientes() {
  const sheet = _malayaSheet();
  const data  = sheet.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < data.length; i++) {
    const estado = String(data[i][10] || '').toLowerCase();
    if (estado !== 'no_bloqueada') continue;
    items.push({
      guest: data[i][1], phone: data[i][2],
      checkin: data[i][3], checkout: data[i][4]
    });
  }
  if (!items.length) return;
  let msg = '📋 *Malaya — reservas sin bloquear en Airbnb (' + items.length + ')*\n\n';
  items.forEach((it, idx) => {
    msg += (idx + 1) + '. ' + it.guest + ' · ' + it.checkin + ' → ' + it.checkout + ' · +' + it.phone + '\n';
  });
  msg += '\nCoordina con Celestino para que las bloquee.';
  try { sendWhatsAppText(BOT_ADMIN_PHONE, msg); } catch(_) {}
}

// ─── Calendar (API para el landing) ─────────────────────────────

// Devuelve fechas ocupadas (de iCal + de mis reservas Directas activas).
// Estructura: { blocked: [{ checkin: 'yyyy-mm-dd', checkout: 'yyyy-mm-dd' }] }
function getMalayaCalendarData() {
  const icalSheet = _malayaIcalSheet();
  const blocked = [];
  if (icalSheet.getLastRow() > 1) {
    const data = icalSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      const ci = data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], 'America/Panama', 'yyyy-MM-dd') : String(data[i][0]).slice(0,10);
      const co = data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], 'America/Panama', 'yyyy-MM-dd') : String(data[i][1]).slice(0,10);
      if (ci && co) blocked.push({ checkin: ci, checkout: co, source: 'airbnb' });
    }
  }
  // Mis reservas directas (pendiente / confirmada / no_bloqueada) cuentan
  // también como ocupadas en el calendario público.
  const malayaSheet = _malayaSheet();
  if (malayaSheet.getLastRow() > 1) {
    const data = malayaSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const estado = String(data[i][10] || '').toLowerCase();
      if (estado === 'cancelada' || estado === 'completada') continue;
      const ci = data[i][3] instanceof Date ? Utilities.formatDate(data[i][3], 'America/Panama', 'yyyy-MM-dd') : String(data[i][3] || '').slice(0,10);
      const co = data[i][4] instanceof Date ? Utilities.formatDate(data[i][4], 'America/Panama', 'yyyy-MM-dd') : String(data[i][4] || '').slice(0,10);
      if (ci && co) blocked.push({ checkin: ci, checkout: co, source: 'directa' });
    }
  }
  return { blocked: blocked };
}

// ─── iCal feed público (para importar en Airbnb) ─────────────────
//
// Devuelve un .ics con todas las reservas directas ACTIVAS de Malaya
// (pendiente / confirmada / no_bloqueada). Airbnb permite importar una URL
// externa de iCal; polla cada pocas horas y bloquea esas fechas para
// evitar double-booking. El endpoint se expone via doGet?action=malayaIcal
// en Parser.gs. La URL es PÚBLICA (sin auth — Airbnb no manda credenciales
// al pollear) → sólo emitimos rango de fechas + summary genérico, cero
// datos personales del huésped.
function getMalayaIcalFeed() {
  const sheet = _malayaSheet();
  const nowUtc = _icalUtcNow();
  const lines = [];
  lines.push('BEGIN:VCALENDAR');
  lines.push('PRODID:-//Las Nubes//Malaya Lodge//ES');
  lines.push('VERSION:2.0');
  lines.push('CALSCALE:GREGORIAN');
  lines.push('METHOD:PUBLISH');
  lines.push('X-WR-CALNAME:Malaya Lodge — Directas Las Nubes');
  lines.push('X-WR-CALDESC:Reservas directas registradas en Las Nubes que Airbnb debe bloquear.');
  lines.push('X-WR-TIMEZONE:America/Panama');

  if (sheet.getLastRow() > 1) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = String(row[0] || '').trim();
      if (!id) continue;
      const estado = String(row[10] || '').toLowerCase().trim();
      if (estado === 'cancelada' || estado === 'completada') continue;
      const ci = row[3] instanceof Date ? Utilities.formatDate(row[3], 'America/Panama', 'yyyy-MM-dd') : String(row[3] || '').slice(0, 10);
      const co = row[4] instanceof Date ? Utilities.formatDate(row[4], 'America/Panama', 'yyyy-MM-dd') : String(row[4] || '').slice(0, 10);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(ci) || !/^\d{4}-\d{2}-\d{2}$/.test(co)) continue;

      // DTSTART inclusivo, DTEND exclusivo (spec de all-day iCal). Una
      // reserva 2026-07-18 → 2026-07-19 emite un bloque de la noche del 18.
      const dtStart = ci.replace(/-/g, '');
      const dtEnd   = co.replace(/-/g, '');
      lines.push('BEGIN:VEVENT');
      lines.push('UID:' + id + '@lasnubes.cloud');
      lines.push('DTSTAMP:' + nowUtc);
      lines.push('DTSTART;VALUE=DATE:' + dtStart);
      lines.push('DTEND;VALUE=DATE:' + dtEnd);
      lines.push('SUMMARY:Reservado (Directa Las Nubes)');
      lines.push('DESCRIPTION:Bloqueo desde el sistema de reservas de Las Nubes.');
      lines.push('STATUS:CONFIRMED');
      lines.push('TRANSP:OPAQUE');
      lines.push('END:VEVENT');
    }
  }
  lines.push('END:VCALENDAR');

  // iCal spec exige CRLF entre líneas.
  const body = lines.join('\r\n') + '\r\n';
  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.ICAL);
}

// Timestamp UTC en formato iCal: YYYYMMDDTHHMMSSZ.
function _icalUtcNow() {
  const d = new Date();
  const p = n => (n < 10 ? '0' + n : String(n));
  return d.getUTCFullYear() + p(d.getUTCMonth() + 1) + p(d.getUTCDate())
       + 'T' + p(d.getUTCHours()) + p(d.getUTCMinutes()) + p(d.getUTCSeconds())
       + 'Z';
}

// ─── Save reserva (API) ─────────────────────────────────────────

function saveMalayaReserva(payload) {
  const guestPhone = String(payload.guestPhone || '').replace(/\D/g, '');
  const checkin    = String(payload.checkin || '').slice(0, 10);
  const checkout   = String(payload.checkout || '').slice(0, 10);
  if (!guestPhone) throw new Error('WhatsApp del huésped es obligatorio');
  if (!checkin || !checkout) throw new Error('Fechas inválidas');
  if (checkout <= checkin)   throw new Error('Check-out debe ser posterior a check-in');

  // Serializamos validate+append para que dos requests concurrentes (o un
  // double-tap del cliente) no puedan ambos "no ver overlap" y ambos insertar.
  // El lock protege TAMBIÉN la lectura del sheet: sin él, la segunda request
  // podría haber leído antes de que la primera hiciera flush del appendRow.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let id, huesped, noches, personas, montoTotal, comision, fechaReserva,
      notas, voucherURL, guestEmail, duplicate = false;
  try {
    // Dedupe idempotente: si en los últimos 120 s ya se guardó una reserva
    // con MISMO teléfono + MISMAS fechas, devolvemos esa en vez de insertar
    // otra. Cubre casos como reintentos del cliente por timeout percibido
    // aunque el primer submit haya llegado ok al servidor.
    const existing = _findRecentMalayaReserva(guestPhone, checkin, checkout, 120);
    if (existing) {
      logDebugEntry('malaya-reserva-DUPLICATE-IGNORED', { id: existing.id, guest: existing.huesped, ci: checkin, co: checkout, phone: guestPhone });
      return {
        ok: true, id: existing.id, noches: existing.noches, montoTotal: existing.montoTotal, comision: existing.comision, duplicate: true
      };
    }

    // Validar disponibilidad (no chocar con iCal ni con otra Directa activa).
    const cal = getMalayaCalendarData();
    for (const b of cal.blocked) {
      // Overlap test: rangos se cruzan si max(start) < min(end).
      if (Math.max(checkin, b.checkin) < Math.min(checkout, b.checkout)) {
        throw new Error('Esas fechas están ocupadas (' + b.source + ').');
      }
    }

    const sheet = _malayaSheet();
    id           = 'MAL-' + Date.now();
    noches       = _malayaNightCount(checkin, checkout);
    personas     = parseInt(payload.personas, 10) || 2;
    montoTotal   = parseFloat(payload.monto) || _malayaCalculateTotal(checkin, checkout, personas);
    comision     = _malayaCalculateCommission(checkin, checkout);
    fechaReserva = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd HH:mm:ss');
    notas        = String(payload.notas || '');
    huesped      = String(payload.guestName || '').trim() || '(sin nombre)';
    voucherURL   = String(payload.voucherURL || '').trim();
    guestEmail   = String(payload.guestEmail || '').trim();

    sheet.appendRow([
      id, huesped, guestPhone, checkin, checkout, noches,
      personas, montoTotal, comision, 'Directa', 'pendiente',
      false, fechaReserva, notas, voucherURL, guestEmail
    ]);
    // Flush inmediato para que si viene un request seguidilla (dedupe check
    // arriba) ya vea esta fila. Sin flush, appendRow puede quedar en cache y
    // el siguiente getDataRange() retornar valores previos.
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  logDebugEntry('malaya-reserva-OK', { id: id, guest: huesped, ci: checkin, co: checkout, monto: montoTotal, comision: comision, voucher: !!voucherURL, email: !!guestEmail });

  // Notificación automática a Celestino (sin ventana de 24h de WhatsApp).
  _emailCelestinoNuevaReserva({
    id: id, huesped: huesped, phone: guestPhone, email: guestEmail,
    checkin: checkin, checkout: checkout, noches: noches,
    personas: personas, montoTotal: montoTotal, comision: comision,
    notas: notas, voucherURL: voucherURL
  });

  // Notificación WhatsApp: si hay plantilla aprobada por Meta, la mando
  // DIRECTA a Celestino (rompe la ventana de 24h). Si no, mando al admin
  // el mensaje listo para reenviar (fallback pre-aprobación).
  _notifyCelestinoWA({
    huesped: huesped, phone: guestPhone, email: guestEmail,
    checkin: checkin, checkout: checkout, noches: noches, personas: personas,
    voucherURL: voucherURL
  });

  return {
    ok: true, id: id, noches: noches, montoTotal: montoTotal, comision: comision
  };
}

// ─── WhatsApp: dispatcher plantilla-o-forward ─────────────────

// Decide cómo entregar la notificación a Celestino:
//   1) Si MALAYA_CELESTINO_TEMPLATE_NAME está seteado en Script Properties
//      → envía la plantilla HSM directo al WhatsApp de Celestino.
//      Además, un aviso corto al admin ("✅ Reserva guardada + notificación
//      enviada a Celestino").
//   2) Si no hay plantilla configurada (aún esperando aprobación de Meta)
//      → cae al flujo actual: mensaje forward-able al admin.
//   3) Si el envío de la plantilla falla (rechazo, cuota, etc.) → también
//      cae al forward-able para no perder la notificación.
function _notifyCelestinoWA(d) {
  const props = PropertiesService.getScriptProperties();
  const templateName   = props.getProperty('MALAYA_CELESTINO_TEMPLATE_NAME');
  const templateLang   = props.getProperty('MALAYA_CELESTINO_TEMPLATE_LANG') || 'es_PA';
  const celestinoPhone = (props.getProperty('MALAYA_CELESTINO_PHONE') || '50765429927').replace(/\D/g, '');
  // Contactos adicionales que también reciben la plantilla (redundancia por si
  // Celestino no la ve). CSV en Script Property MALAYA_EXTRA_NOTIFY_PHONES;
  // default incluye a Glorimar (+507 6100-0079).
  const extraRaw = props.getProperty('MALAYA_EXTRA_NOTIFY_PHONES') || '50761000079';
  const extras = extraRaw.split(',').map(s => s.replace(/\D/g, '')).filter(Boolean);
  const recipients = [celestinoPhone].concat(extras).filter((n, i, arr) => n && arr.indexOf(n) === i);

  if (!templateName || !recipients.length) {
    _whatsappAdminForwardCelestino(d);
    return;
  }

  const vars = _buildCelestinoTemplateVars(d);
  let anyOk = false;
  recipients.forEach(to => {
    try {
      sendWhatsAppTemplate(to, templateName, templateLang, vars);
      logDebugEntry('malaya-celestino-template-OK', { to: to, template: templateName });
      anyOk = true;
    } catch(e) {
      logDebugEntry('malaya-celestino-template-FAIL', { to: to, template: templateName, error: e.message });
    }
  });

  if (anyOk) {
    _whatsappAdminTemplateAck(d, recipients);
  } else {
    // Ningún destinatario recibió la plantilla → aviso manual para no perder la notificación.
    _whatsappAdminForwardCelestino(d);
  }
}

// Variables POSICIONALES {{1}}..{{6}} que rellenan malaya_reserva_celestino.
// Meta NO acepta variables nombradas — devolvemos array ordenado.
// Meta NO permite variables vacías → cada campo tiene un fallback textual.
// Orden: {{1}}=fechas, {{2}}=huesped, {{3}}=telefono, {{4}}=email,
//         {{5}}=personas, {{6}}=voucher.
function _buildCelestinoTemplateVars(d) {
  const fmt = iso => {
    const dt = new Date(iso + 'T12:00:00');
    const M   = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
    const DOW = ['dom','lun','mar','mié','jue','vie','sáb'];
    return DOW[dt.getDay()] + ' ' + dt.getDate() + ' ' + M[dt.getMonth()];
  };
  const ciFmt = fmt(d.checkin);
  const coFmt = fmt(d.checkout);
  const nochesLbl   = d.noches   + ' ' + (d.noches   === 1 ? 'noche'   : 'noches');
  const personasLbl = d.personas + ' ' + (d.personas === 1 ? 'persona' : 'personas');
  return [
    ciFmt + ' → ' + coFmt + ' (' + nochesLbl + ')',   // {{1}} fechas
    d.huesped || '(sin nombre)',                       // {{2}} huésped
    '+' + d.phone,                                     // {{3}} teléfono
    d.email     || 'no proporcionado',                 // {{4}} email
    personasLbl,                                       // {{5}} personas
    d.voucherURL || 'no disponible'                    // {{6}} voucher
  ];
}

// Aviso corto al admin cuando la plantilla se entregó ok. Corto porque el
// admin ya no necesita reenviar nada — solo confirmar que quedó fuera.
function _whatsappAdminTemplateAck(d, recipients) {
  const fmt = iso => {
    const dt = new Date(iso + 'T12:00:00');
    const M = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
    return dt.getDate() + ' ' + M[dt.getMonth()];
  };
  const n = (recipients && recipients.length) || 1;
  const destinoLbl = n === 1
    ? 'Celestino'
    : 'Celestino + ' + (n - 1) + ' contacto' + (n - 1 === 1 ? '' : 's') + ' extra';
  const msg =
    '✅ *Reserva Malaya guardada*\n' +
    'Notificación enviada a ' + destinoLbl + ' via plantilla WhatsApp.\n\n' +
    '👤 ' + d.huesped + '\n' +
    '📅 ' + fmt(d.checkin) + ' → ' + fmt(d.checkout);
  try { sendWhatsAppText(BOT_ADMIN_PHONE, msg); } catch(_) {}
}

// ─── WhatsApp al admin con mensaje listo para reenviar ────────

function _whatsappAdminForwardCelestino(d) {
  const fmt = iso => {
    const dt = new Date(iso + 'T12:00:00');
    const M   = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
    const DOW = ['dom','lun','mar','mié','jue','vie','sáb'];
    return DOW[dt.getDay()] + ' ' + dt.getDate() + ' ' + M[dt.getMonth()];
  };
  const ciFmt  = fmt(d.checkin);
  const coFmt  = fmt(d.checkout);
  const nochesLbl = d.noches + ' ' + (d.noches === 1 ? 'noche' : 'noches');
  const personasLbl = d.personas + ' ' + (d.personas === 1 ? 'persona' : 'personas');

  // Mensaje redactado como si fuera para Celestino — el admin lo reenvía
  // tal cual con un long-press en WhatsApp.
  const voucherLine = d.voucherURL ? '🧾 Voucher: ' + d.voucherURL + '\n' : '';
  const emailLine   = d.email ? '✉ ' + d.email + '\n' : '';
  const forwardable =
    'Hola Celestino! 👋\n\n' +
    'Cerré una reserva directa en Malaya. Por favor bloquéala en Airbnb:\n\n' +
    '📅 ' + ciFmt + ' → ' + coFmt + ' (' + nochesLbl + ')\n' +
    '🕑 Check-in 2pm · Check-out 11am\n' +
    '👤 ' + d.huesped + '\n' +
    '📱 +' + d.phone + '\n' +
    emailLine +
    '👥 ' + personasLbl + '\n' +
    voucherLine + '\n' +
    'Recuerda enviarle al cliente todos los detalles de la reserva e indicaciones de llegada, junto con tu información de contacto para cualquier duda o coordinación el día del check-in.\n\n' +
    'Gracias!';

  // Dos mensajes separados: (1) aviso al admin con instrucción, (2) el
  // mensaje forward-able limpio para que el admin haga long-press → reenviar
  // sin tener que editar nada.
  const header =
    '✅ *Reserva Malaya guardada*\n' +
    'Te mando abajo el mensaje listo para reenviarle a Celestino (long-press → reenviar).';

  try { sendWhatsAppText(BOT_ADMIN_PHONE, header); }      catch(_) {}
  try { sendWhatsAppText(BOT_ADMIN_PHONE, forwardable); } catch(_) {}
}

// ─── Voucher: subir a Drive y devolver URL ─────────────────────

function saveMalayaVoucherToDrive(payload) {
  const FOLDER_NAME = 'Malaya Lodge - Pagos';
  try {
    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    const folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);

    const mimeType = String(payload.mimeType || 'image/jpeg');
    const ext      = mimeType.includes('png') ? '.png'
                   : mimeType.includes('pdf') ? '.pdf'
                   : mimeType.includes('gif') ? '.gif' : '.jpg';
    const nombre   = String(payload.guestName || 'huesped').replace(/\s+/g, '_').slice(0, 20);
    const checkin  = String(payload.checkin || '').slice(0, 10);
    const stamp    = Utilities.formatDate(new Date(), 'America/Panama', 'yyyyMMdd_HHmmss');
    const fileName = 'malaya_' + nombre + '_' + checkin + '_' + stamp + ext;

    const blob = Utilities.newBlob(Utilities.base64Decode(payload.imageBase64), mimeType, fileName);
    const file = folder.createFile(blob);
    file.setDescription([
      'Huésped: '  + (payload.guestName || ''),
      'Entrada: '  + checkin,
      'Salida: '   + (payload.checkout || '').slice(0,10),
      'Monto: $'   + (payload.monto || ''),
      'Registrado: ' + Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd HH:mm')
    ].join('\n'));
    // Cualquiera con el link puede ver (Celestino no tiene cuenta del workspace).
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { ok: true, url: file.getUrl(), id: file.getId() };
  } catch(e) {
    logDebugEntry('malaya-voucher-drive-ERR', { error: e.message });
    return { ok: false, error: e.message };
  }
}

// ─── Email a Celestino al crear reserva directa ─────────────────

function _emailCelestinoNuevaReserva(d) {
  const to = PropertiesService.getScriptProperties().getProperty('MALAYA_CELESTINO_EMAIL') || 'malayalodge@gmail.com';
  const fmt = iso => {
    const dt = new Date(iso + 'T12:00:00');
    const M   = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
    const DOW = ['dom','lun','mar','mié','jue','vie','sáb'];
    return DOW[dt.getDay()] + ' ' + dt.getDate() + ' ' + M[dt.getMonth()];
  };
  const ciFmt = fmt(d.checkin);
  const coFmt = fmt(d.checkout);
  const noches = d.noches + ' ' + (d.noches === 1 ? 'noche' : 'noches');

  const html =
    '<div style="font-family:-apple-system,Segoe UI,Arial,sans-serif; max-width:560px; color:#1a1a1a; line-height:1.5;">' +
      '<h2 style="color:#5a7a4a; margin:0 0 14px;">🌿 Nueva reserva directa en Malaya</h2>' +
      '<p>Hola Celestino, cerré una reserva directa. Por favor <strong>bloquéala en Airbnb</strong> para evitar doble booking.</p>' +
      '<table style="border-collapse:collapse; margin:18px 0; width:100%; font-size:14px;">' +
        '<tr><td style="padding:6px 10px; color:#666; width:130px;">Fechas</td><td style="padding:6px 10px;"><strong>' + ciFmt + ' → ' + coFmt + '</strong> · ' + noches + '</td></tr>' +
        '<tr><td style="padding:6px 10px; color:#666;">Check-in</td><td style="padding:6px 10px;">2:00 pm</td></tr>' +
        '<tr><td style="padding:6px 10px; color:#666;">Check-out</td><td style="padding:6px 10px;">11:00 am</td></tr>' +
        '<tr><td style="padding:6px 10px; color:#666;">Huésped</td><td style="padding:6px 10px;">' + d.huesped + '</td></tr>' +
        '<tr><td style="padding:6px 10px; color:#666;">WhatsApp</td><td style="padding:6px 10px;"><a href="https://wa.me/' + d.phone + '">+' + d.phone + '</a></td></tr>' +
        (d.email ? '<tr><td style="padding:6px 10px; color:#666;">Email</td><td style="padding:6px 10px;"><a href="mailto:' + d.email + '">' + d.email + '</a></td></tr>' : '') +
        '<tr><td style="padding:6px 10px; color:#666;">Personas</td><td style="padding:6px 10px;">' + d.personas + '</td></tr>' +
        '<tr><td style="padding:6px 10px; color:#666;">Total cobrado</td><td style="padding:6px 10px;">$' + Number(d.montoTotal).toFixed(2) + '</td></tr>' +
        '<tr><td style="padding:6px 10px; color:#666;">Comisión</td><td style="padding:6px 10px;">$' + Number(d.comision).toFixed(2) + '</td></tr>' +
        (d.voucherURL ? '<tr><td style="padding:6px 10px; color:#666;">Voucher</td><td style="padding:6px 10px;"><a href="' + d.voucherURL + '">Ver comprobante de pago</a></td></tr>' : '') +
      '</table>' +
      (d.notas ? '<p style="font-size:13px; color:#666;"><strong>Notas:</strong> ' + d.notas + '</p>' : '') +
      '<p style="font-size:13px; color:#888; margin-top:24px;">El sistema verifica cada 30 min si la fecha aparece bloqueada en tu iCal de Airbnb. Si pasan más de 60 min sin bloqueo, recibo una alerta para coordinar contigo.</p>' +
    '</div>';

  const subject = '🌿 Reserva Malaya: ' + ciFmt + ' → ' + coFmt + ' · ' + d.huesped;

  try {
    GmailApp.sendEmail(to, subject, '', {
      htmlBody: html,
      name: 'Las Nubes — Reservas Malaya'
    });
    logDebugEntry('malaya-email-celestino-OK', { to: to, id: d.id });
  } catch(e) {
    logDebugEntry('malaya-email-celestino-ERR', { to: to, id: d.id, error: e.message });
  }
}

// ─── Cancelar reserva ──────────────────────────────────────────

function cancelMalayaReserva(reservaId) {
  if (!reservaId) throw new Error('reservaId requerido');
  const sheet = _malayaSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === reservaId) {
      sheet.getRange(i + 1, 11).setValue('cancelada');
      logDebugEntry('malaya-reserva-cancelada', { id: reservaId });
      return { ok: true };
    }
  }
  throw new Error('Reserva ' + reservaId + ' no encontrada');
}

// ─── Trigger installer (correr UNA VEZ desde editor) ────────────

function instalarTriggersMalaya() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const h = t.getHandlerFunction();
    if (h === 'syncMalayaAirbnb' || h === 'verificarMalayaPendientes') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('syncMalayaAirbnb').timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger('verificarMalayaPendientes').timeBased().everyDays(1).atHour(11).inTimezone('America/Panama').create();
  Logger.log('✓ Triggers Malaya instalados: sync 30 min, verificar 11am diario.');
}

// Test desde editor: renderea el iCal público de Malaya y lo loguea.
// Útil para verificar el contenido antes de compartirle la URL a Celestino.
function _testMalayaIcalFeed() {
  const out = getMalayaIcalFeed();
  Logger.log(out.getContent());
}

// Test desde editor: corre sync manualmente y loguea resultado.
function _testMalayaSync() {
  syncMalayaAirbnb();
  const data = _malayaIcalSheet().getDataRange().getValues();
  Logger.log('iCal events cacheados: ' + (data.length - 1));
  for (let i = 1; i < Math.min(data.length, 20); i++) {
    Logger.log('  ' + data[i][0] + ' → ' + data[i][1] + ' · ' + data[i][2]);
  }
}

// Test de la plantilla malaya_reserva_celestino apuntando a un destino
// custom (útil para probar sin molestar a Celestino). Editá la constante
// TO_PHONE con tu número y correlo desde el editor.
function _testMalayaCelestinoTemplate() {
  const TO_PHONE = '50769812266'; // ← cambiá por TU número (E.164 sin +)

  const props = PropertiesService.getScriptProperties();
  const templateName = props.getProperty('MALAYA_CELESTINO_TEMPLATE_NAME');
  const templateLang = props.getProperty('MALAYA_CELESTINO_TEMPLATE_LANG') || 'es_PA';
  if (!templateName) { Logger.log('✗ Falta MALAYA_CELESTINO_TEMPLATE_NAME en Script Properties'); return; }

  const vars = _buildCelestinoTemplateVars({
    huesped:    'Test Reserva',
    phone:      '50761234567',
    email:      'test@example.com',
    checkin:    '2026-08-15',
    checkout:   '2026-08-17',
    noches:     2,
    personas:   2,
    voucherURL: 'https://drive.google.com/file/d/1abc/view'
  });

  const to = TO_PHONE.replace(/\D/g, '');
  Logger.log('→ Enviando template "' + templateName + '" (' + templateLang + ') a +' + to);
  try {
    const res = sendWhatsAppTemplate(to, templateName, templateLang, vars);
    Logger.log('✓ OK: ' + JSON.stringify(res));
  } catch(e) {
    Logger.log('✗ ERR: ' + e.message);
  }
}
