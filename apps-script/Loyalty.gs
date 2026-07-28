// ═══════════════════════════════════════════════════════════
//  Programa Cliente Fiel — $20 de credito tras la 2da estadia
// ═══════════════════════════════════════════════════════════
//
//  Trigger diario que escanea Reservas y emite el credito a los huespedes
//  que ya completaron 2 estadias y aun no lo recibieron.
//
//  Antes el beneficio era una NOCHE GRATIS. Salio del programa por costo
//  ($90 contra $20 de los otros dos) y porque acumulaba deuda: 10 creditos
//  vivos valian ~$890. Volvio como $20 de credito, que es la misma cifra
//  del cumpleanos y del referido — asi el programa entero se dice en una
//  linea y los mismos 10 creditos valen $200.
//
//  Que esperar: $20 no provoca una segunda visita como lo hacia una noche
//  gratis. El valor de este programa es el CORREO — una excusa para
//  escribirle a alguien que ya vino — mas que el descuento.
//
//  Filtros para que una estadia cuente:
//   - origin = 'Directa'
//   - estadoPago != 'CANCELADA'
//   - tipo en [noche, early, late]
//   - checkout < hoy (ya completada)
//
//  El credito se guarda en hoja LoyaltyCreditos. Cuando admin lo aplica,
//  lo marca como usado (boton en dashboard).
//
//  Vencimiento: 12 meses desde EarnedAt — el mismo que el credito de
//  referido, para que no haya dos reglas.
//
//  Para apagarlo: Script Property LOYALTY_ACTIVO = 'false'.
// ═══════════════════════════════════════════════════════════
//
//  Trigger diario que escanea Reservas y emite credito de cortesia
//  entre semana a los huespedes que ya completaron 2 estadias y aun no
//  recibieron credito.
//
//  Filtros para que una estadia cuente:
//   - origin = 'Directa'
//   - estadoPago != 'CANCELADA'
//   - tipo en [noche, early, late]
//   - checkout < hoy (ya completada)
//
//  El credito se guarda en hoja LoyaltyCreditos. Cuando admin aplica
//  la cortesia, marca el credito como usado (boton en dashboard).
//
//  Vencimiento: 12 meses desde EarnedAt.
//
//  Activacion: ejecutar instalarTriggerLoyalty() una vez.
//  Preview:    ejecutar enviarLoyaltyPrueba() desde el editor.
// ═══════════════════════════════════════════════════════════

const LOYALTY_DIAS_VENCIMIENTO = 365; // 12 meses — igual que el credito de referido
const LOYALTY_MONTO = 20;             // mismo monto que cumpleanos y referido

function _getOrCreateLoyaltySheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let s = ss.getSheetByName('LoyaltyCreditos');
  if (!s) {
    s = ss.insertSheet('LoyaltyCreditos');
    s.getRange(1, 1, 1, 9).setValues([[
      'Email', 'Telefono', 'Nombre',
      'EarnedAt', 'EarnedAfterReservaId',
      'UsedAt', 'UsedInReservaId',
      'ExpiresAt', 'Notas'
    ]]);
    s.getRange(1, 1, 1, 9).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

function _normEmail(e) { return (e || '').toString().toLowerCase().trim(); }
function _normPhone(p) { return (p || '').toString().replace(/\D/g, ''); }

// Devuelve la primera fila de credito que matchea email o telefono, o null.
function _findLoyaltyCredit(email, telefono) {
  const e = _normEmail(email);
  const p = _normPhone(telefono);
  if (!e && !p) return null;
  const s = _getOrCreateLoyaltySheet();
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const re = _normEmail(data[i][0]);
    const rp = _normPhone(data[i][1]);
    if ((e && re === e) || (p && rp === p)) {
      return {
        rowIndex:    i + 1,
        email:       data[i][0],
        telefono:    data[i][1],
        nombre:      data[i][2],
        earnedAt:    data[i][3] instanceof Date ? Utilities.formatDate(data[i][3], 'America/Panama', 'yyyy-MM-dd') : data[i][3],
        earnedAfterReservaId: data[i][4],
        usedAt:      data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], 'America/Panama', 'yyyy-MM-dd') : data[i][5],
        usedInReservaId:      data[i][6],
        expiresAt:   data[i][7] instanceof Date ? Utilities.formatDate(data[i][7], 'America/Panama', 'yyyy-MM-dd') : data[i][7],
        notas:       data[i][8]
      };
    }
  }
  return null;
}

// Cuenta estadias completadas (dormido + directa + no canceladas) por huesped
function _countCompletedNightStaysForGuest(email, telefono, allRows) {
  const e = _normEmail(email);
  const p = _normPhone(telefono);
  if (!e && !p) return { count: 0, latestId: null };
  const today = new Date();
  const todayMs = new Date(Utilities.formatDate(today, 'America/Panama', 'yyyy-MM-dd') + 'T00:00:00-05:00').getTime();
  let count = 0;
  let latestMs = 0;
  let latestId = null;
  for (let i = 1; i < allRows.length; i++) {
    const r = allRows[i];
    if (!r[0]) continue;
    if (!['Directa','Referido'].includes(r[9] || '')) continue;     // origin direct-style
    if ((r[20] || '').toString().toUpperCase() === 'CANCELADA') continue;
    const tipo = (r[24] || 'noche').toString();
    if (tipo === 'pasatarde' || tipo === 'pasadia') continue;
    const re = _normEmail(r[21]);
    const rp = _normPhone(r[23]);
    if (!((e && re === e) || (p && rp === p))) continue;
    if (!r[5]) continue;
    const coStr = r[5] instanceof Date
      ? Utilities.formatDate(r[5], 'America/Panama', 'yyyy-MM-dd')
      : r[5].toString().slice(0, 10);
    const coMs = new Date(coStr + 'T12:00:00-05:00').getTime();
    if (isNaN(coMs) || coMs >= todayMs) continue;
    count++;
    if (coMs > latestMs) { latestMs = coMs; latestId = r[0]; }
  }
  return { count, latestId };
}

// Interruptor de la EMISION. Opt-OUT y no opt-in: el programa esta activo
// salvo que alguien lo apague a proposito. Con opt-in, un proyecto nuevo o una
// property borrada dejaba el programa mudo sin que nadie se enterara.
function _loyaltyActivo() {
  return PropertiesService.getScriptProperties().getProperty('LOYALTY_ACTIVO') !== 'false';
}

function enviarLoyaltyUnlockEmails() {
  if (!_loyaltyActivo()) {
    Logger.log('⊘ Cliente Fiel apagado (LOYALTY_ACTIVO = false) — no se emiten créditos.');
    return;
  }
  const sheet  = getOrCreateSheet();
  _ensureHuespedIdColumns(sheet);
  const data   = sheet.getDataRange().getValues();
  const ls     = _getOrCreateLoyaltySheet();
  const today  = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');

  // Agrupar candidatos unicos por email
  const candidatos = new Map(); // email -> { nombre, telefono, latestId }
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if (!['Directa','Referido'].includes(r[9] || '')) continue;
    if ((r[20] || '').toString().toUpperCase() === 'CANCELADA') continue;
    const tipo = (r[24] || 'noche').toString();
    if (tipo === 'pasatarde' || tipo === 'pasadia') continue;
    const email = _normEmail(r[21]);
    if (!email) continue;  // necesitamos email para enviar
    if (!candidatos.has(email)) {
      candidatos.set(email, { nombre: r[1], telefono: r[23] || '' });
    }
  }

  let issued = 0, skipped = 0, errors = 0;
  candidatos.forEach((info, email) => {
    const existing = _findLoyaltyCredit(email, info.telefono);
    if (existing) { skipped++; return; }
    const stats = _countCompletedNightStaysForGuest(email, info.telefono, data);
    if (stats.count < 2) { skipped++; return; }
    try {
      // Crear credito
      const earnedAt  = today;
      const expiresD  = new Date();
      expiresD.setDate(expiresD.getDate() + LOYALTY_DIAS_VENCIMIENTO);
      const expiresAt = Utilities.formatDate(expiresD, 'America/Panama', 'yyyy-MM-dd');
      ls.appendRow([email, info.telefono, info.nombre, earnedAt, stats.latestId, '', '', expiresAt, '']);
      // Enviar email
      _sendLoyaltyUnlockEmail({ email, nombre: info.nombre, expiresAt });
      issued++;
    } catch(e) {
      errors++;
      Logger.log('⚠ Error loyalty ' + email + ': ' + e.message);
    }
  });
  Logger.log('🌟 Loyalty: ' + issued + ' creditos emitidos · ' + skipped + ' skip · ' + errors + ' errores');
  return { issued, skipped, errors };
}

function _sendLoyaltyUnlockEmail(opts) {
  const props = PropertiesService.getScriptProperties();
  const waNum = (props.getProperty('CONTACT_WHATSAPP_NUMBER') || '50769812266').replace(/\D/g, '');
  const firstName = (opts.nombre || '').split(/\s+/)[0] || opts.nombre;
  const subject = '¡Tu próxima noche en Las Nubes va por la casa!';
  const waText  = encodeURIComponent('Hola! Recibi el correo del programa Cliente Fiel — me interesa usar mi credito de $' + LOYALTY_MONTO + '.');
  const waLink  = 'https://wa.me/' + waNum + '?text=' + waText;
  const html = buildLoyaltyUnlockEmailHTML({
    firstName,
    expiresAt: opts.expiresAt,
    waLink
  });
  GmailApp.sendEmail(opts.email, subject, '', { htmlBody: html, name: 'Las Nubes', replyTo: REPLY_TO_EMAIL });
}

function buildLoyaltyUnlockEmailHTML(opts) {
  const expFmt = _fmtFechaLargaES(opts.expiresAt);
  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>' +
'<body style="margin:0;padding:0;background:#f5f3f0;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;">' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f3f0;padding:32px 16px;"><tr><td align="center">' +
'<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;">' +
'<tr><td style="background:#d4a44c;border-radius:16px 16px 0 0;padding:36px 40px;text-align:center;">' +
'<p style="margin:0 0 6px;font-size:12px;color:rgba(255,255,255,0.8);letter-spacing:2px;text-transform:uppercase;">&#127775; Cliente fiel</p>' +
'<h1 style="margin:0;font-size:32px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></h1>' +
'<p style="margin:10px 0 0;font-size:15px;color:rgba(255,255,255,0.9);">Buenos Aires, Chame · Panamá Oeste</p>' +
'</td></tr>' +
'<tr><td style="background:#ffffff;padding:36px 40px;">' +
'<p style="margin:0 0 16px;font-size:16px;color:#3a3530;line-height:1.6;">Hola <strong>' + (opts.firstName || 'amigo') + '</strong>,</p>' +
'<p style="margin:0 0 16px;font-size:15px;color:#3a3530;line-height:1.7;">Gracias por volver a Las Nubes. Es un placer recibirte como huésped frecuente — y queremos agradecértelo.</p>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#fffbeb;border:1px solid #d4a44c;border-radius:14px;margin:24px 0;">' +
'<tr><td style="padding:24px 28px;text-align:center;">' +
'<p style="margin:0 0 10px;font-size:13px;color:#7a5a1f;letter-spacing:0.05em;text-transform:uppercase;font-weight:600;">Tu regalo</p>' +
'<p style="margin:0 0 14px;font-size:26px;font-weight:300;color:#7a5a1f;font-family:Georgia,serif;">$' + LOYALTY_MONTO + ' de crédito</p>' +
'<p style="margin:0;font-size:13px;color:#7a5a1f;line-height:1.6;">Sobre cualquier cabaña, según disponibilidad.</p>' +
'</td></tr></table>' +
'<p style="margin:0 0 20px;font-size:14px;color:#6b6560;line-height:1.7;">Cómo funciona:</p>' +
'<ul style="margin:0 0 20px;padding-left:20px;color:#6b6560;font-size:14px;line-height:1.7;">' +
'<li>Escribinos por WhatsApp cuando quieras usar tu noche.</li>' +
'<li>Aplica Domingo a Jueves, según disponibilidad.</li>' +
'<li>Vence el <strong>' + expFmt + '</strong>.</li>' +
'<li>1 sola noche por cliente.</li>' +
'<li>No combinable con otras promociones.</li>' +
'</ul>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:20px;"><tr><td align="center">' +
'<a href="' + opts.waLink + '" target="_blank" style="display:inline-block;background:#25d366;color:#ffffff;font-size:15px;font-weight:600;padding:14px 28px;border-radius:10px;text-decoration:none;">&#128172; Usar mi crédito</a>' +
'</td></tr></table>' +
// Los beneficios NO son combinables — decirlo al revés en el email era una
// promesa que el programa no cumple.
'<p style="margin:0;font-size:13px;color:#8a8078;line-height:1.6;text-align:center;">Aplica a noches de domingo a jueves, sin feriados, vísperas de feriado ni vacaciones escolares. Vence a los 12 meses. No se combina con otras promociones.</p>' +
'</td></tr>' +
'<tr><td style="background:#3a3530;border-radius:0 0 16px 16px;padding:24px 40px;text-align:center;">' +
'<p style="margin:0 0 8px;font-size:18px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></p>' +
'<p style="margin:0 0 12px;font-size:11px;color:rgba(255,255,255,0.6);letter-spacing:1px;text-transform:uppercase;">Cabañas en Chicá · Panamá Oeste</p>' +
'<a href="https://wa.me/50769812266" style="color:rgba(255,255,255,0.8);font-size:13px;text-decoration:none;">&#128172; WhatsApp: +507 6981-2266</a>' +
'</td></tr></table></td></tr></table></body></html>';
}

function _fmtFechaLargaES(yyyymmdd) {
  const meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  const m = (yyyymmdd || '').toString().match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return yyyymmdd || '';
  return parseInt(m[3], 10) + ' de ' + meses[parseInt(m[2], 10) - 1] + ' ' + m[1];
}

// Endpoint: GET ?action=getLoyaltyCredits → lista de creditos (admin)
function handleGetLoyaltyCredits(e) {
  try {
    const s = _getOrCreateLoyaltySheet();
    const data = s.getDataRange().getValues();
    const out = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      out.push({
        email:       data[i][0],
        telefono:    data[i][1],
        nombre:      data[i][2],
        earnedAt:    data[i][3] instanceof Date ? Utilities.formatDate(data[i][3], 'America/Panama', 'yyyy-MM-dd') : data[i][3],
        earnedAfterReservaId: data[i][4],
        usedAt:      data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], 'America/Panama', 'yyyy-MM-dd') : data[i][5],
        usedInReservaId:      data[i][6],
        expiresAt:   data[i][7] instanceof Date ? Utilities.formatDate(data[i][7], 'America/Panama', 'yyyy-MM-dd') : data[i][7],
        notas:       data[i][8]
      });
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, credits: out })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Endpoint POST: marca un credito como usado en una reserva
function handleMarkLoyaltyUsed(payload) {
  try {
    const email = payload.email;
    const telefono = payload.telefono;
    const reservaId = payload.reservaId;
    if (!email && !telefono) return _jsonOut({ ok: false, error: 'MISSING_GUEST' });
    if (!reservaId) return _jsonOut({ ok: false, error: 'MISSING_RESERVA' });
    const credit = _findLoyaltyCredit(email, telefono);
    if (!credit) return _jsonOut({ ok: false, error: 'NO_CREDIT' });
    if (credit.usedAt) return _jsonOut({ ok: false, error: 'ALREADY_USED', usedInReservaId: credit.usedInReservaId });
    const s = _getOrCreateLoyaltySheet();
    const today = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');
    s.getRange(credit.rowIndex, 6).setValue(today);
    s.getRange(credit.rowIndex, 7).setValue(reservaId);
    return _jsonOut({ ok: true, usedAt: today, usedInReservaId: reservaId });
  } catch(err) {
    return _jsonOut({ ok: false, error: err.message });
  }
}

// Endpoint POST: deshace 'usada' (admin caso de error)
function handleUnmarkLoyaltyUsed(payload) {
  try {
    const email = payload.email;
    const telefono = payload.telefono;
    if (!email && !telefono) return _jsonOut({ ok: false, error: 'MISSING_GUEST' });
    const credit = _findLoyaltyCredit(email, telefono);
    if (!credit) return _jsonOut({ ok: false, error: 'NO_CREDIT' });
    const s = _getOrCreateLoyaltySheet();
    s.getRange(credit.rowIndex, 6).setValue('');
    s.getRange(credit.rowIndex, 7).setValue('');
    return _jsonOut({ ok: true });
  } catch(err) {
    return _jsonOut({ ok: false, error: err.message });
  }
}

// Quita el trigger diario. Util si se apaga el programa: dejarlo corriendo
// solo genera trabajo que despues hay que deshacer.
function desinstalarTriggerLoyalty() {
  let n = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'enviarLoyaltyUnlockEmails') { ScriptApp.deleteTrigger(t); n++; }
  });
  Logger.log('✓ Triggers de Cliente Fiel eliminados: ' + n);
  Logger.log('  Los créditos ya emitidos siguen válidos hasta su vencimiento.');
}

function instalarTriggerLoyalty() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'enviarLoyaltyUnlockEmails') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarLoyaltyUnlockEmails')
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .inTimezone('America/Panama')
    .create();
  Logger.log('✓ Trigger creado: enviarLoyaltyUnlockEmails diario @ 10am Panama');
}

function enviarLoyaltyPrueba() {
  const myEmail = Session.getActiveUser().getEmail();
  if (!myEmail) { Logger.log('⚠ Email no disponible. Corré desde el editor.'); return; }
  const expiresD = new Date();
  expiresD.setDate(expiresD.getDate() + LOYALTY_DIAS_VENCIMIENTO);
  const expiresAt = Utilities.formatDate(expiresD, 'America/Panama', 'yyyy-MM-dd');
  _sendLoyaltyUnlockEmail({ email: myEmail, nombre: 'María de Prueba', expiresAt });
  Logger.log('✓ Loyalty de prueba enviado a ' + myEmail);
}
