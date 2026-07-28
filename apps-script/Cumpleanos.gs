// ═══════════════════════════════════════════════════════════
//  Programa de cumpleaños — email recordatorio + dedup
// ═══════════════════════════════════════════════════════════
//
//  Trigger diario que escanea Reservas y envia un email a los
//  huespedes cuyo cumpleaños cae exactamente dentro de 30 dias.
//
//  Filtros:
//   - Origin = 'Directa' (no Airbnb)
//   - Tiene email
//   - Tiene FechaNacimiento (col 28) registrada
//   - Aun no se le mando este año (hoja CumpleanosEnviados)
//
//  Descuento: monto FIJO en dolares sobre la noche del cumpleanos.
//  Antes era un porcentaje por dia de semana (20% Dom-Jue / 10% Vie-Sab). Con
//  el tarifario por tipo de dia eso regalaba mas en las noches mas caras: un
//  martes que ademas es vispera de feriado ($135) recibia $27, contra $18 de un
//  martes normal. Un monto fijo no tiene ese sesgo y ademas es comparable con
//  los otros dos beneficios del programa, que ya estaban en dolares.
//  No aplica en feriados, visperas ni vacaciones escolares.
//  Ajustable via Script Properties:
//     CUMPLE_DESCUENTO_DOM_JUE (default 30)
//     CUMPLE_DESCUENTO_VIE_SAB (default 20)
//     CUMPLE_DIAS_ANTES        (default 30)
//
//  Activacion: ejecutar instalarTriggerCumpleanos() una sola vez.
//  Preview:    ejecutar enviarCumpleanosPrueba() desde el editor.
// ═══════════════════════════════════════════════════════════

function _cumpleanosConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    // Monto fijo en dolares. Se sigue leyendo de Script Properties para poder
    // ajustarlo sin deploy; CUMPLE_DESCUENTO_DOM_JUE queda como nombre legado.
    descMonto:  parseInt(props.getProperty('CUMPLE_DESCUENTO_MONTO'), 10) || 25,
    diasAntes:  parseInt(props.getProperty('CUMPLE_DIAS_ANTES'), 10)        || 30
  };
}

function _getOrCreateCumpleanosSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let s = ss.getSheetByName('CumpleanosEnviados');
  if (!s) {
    s = ss.insertSheet('CumpleanosEnviados');
    s.getRange(1, 1, 1, 4).setValues([['Email', 'Year', 'SentAt', 'Nombre']]);
    s.getRange(1, 1, 1, 4).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

function _yaEnviadoCumpleEsteAnio(email, year) {
  const s = _getOrCreateCumpleanosSheet();
  const data = s.getDataRange().getValues();
  const e = (email || '').toString().toLowerCase().trim();
  if (!e) return true; // sin email, lo damos por enviado para skipear
  for (let i = 1; i < data.length; i++) {
    const re = (data[i][0] || '').toString().toLowerCase().trim();
    const ry = parseInt(data[i][1], 10);
    if (re === e && ry === year) return true;
  }
  return false;
}

function _marcarCumpleEnviado(email, year, nombre) {
  const s = _getOrCreateCumpleanosSheet();
  const stamp = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd HH:mm:ss');
  s.appendRow([(email || '').toLowerCase(), year, stamp, nombre || '']);
}

// dobStr en 'YYYY-MM-DD' (o cualquier string con DD/MM o MM/DD). Devuelve { dd, mm } o null.
function _parseDobMonthDay(dobStr) {
  if (!dobStr) return null;
  const s = dobStr.toString().trim();
  // Formato esperado YYYY-MM-DD
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return { dd: parseInt(m[3], 10), mm: parseInt(m[2], 10) };
  // DD/MM/YYYY o DD-MM-YYYY
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (m) return { dd: parseInt(m[1], 10), mm: parseInt(m[2], 10) };
  return null;
}

function _addDays(d, n) {
  const out = new Date(d.getTime());
  out.setDate(out.getDate() + n);
  return out;
}

// Trigger principal — corre diaria 10am Panama.
function enviarRecordatoriosCumpleanos() {
  const cfg = _cumpleanosConfig();
  const sheet = getOrCreateSheet();
  _ensureHuespedIdColumns(sheet);
  const data = sheet.getDataRange().getValues();

  const today  = new Date();
  const target = _addDays(today, cfg.diasAntes);
  const targetMonthDay = { dd: target.getDate(), mm: target.getMonth() + 1 };
  const currentYear = today.getFullYear();

  // Agrupar por email para no mandar duplicados si el huesped tiene varias reservas
  const candidatos = new Map(); // email -> { nombre, dobStr }
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const origin     = (r[9]  || '').toString();
    const estado     = (r[20] || '').toString().toUpperCase();
    const email      = (r[21] || '').toString().toLowerCase().trim();
    const nombre     = (r[1]  || '').toString();
    const fechaNac   = r[27]; // col 28 (1-indexed)

    if (!['Directa','Referido'].includes(origin)) continue;
    if (estado === 'CANCELADA') continue;
    if (!email) continue;
    if (!fechaNac) continue;

    const dobStr = fechaNac instanceof Date
      ? Utilities.formatDate(fechaNac, 'America/Panama', 'yyyy-MM-dd')
      : fechaNac.toString();
    const md = _parseDobMonthDay(dobStr);
    if (!md) continue;
    if (md.dd !== targetMonthDay.dd || md.mm !== targetMonthDay.mm) continue;

    // Solo el primero por email
    if (!candidatos.has(email)) {
      candidatos.set(email, { nombre, dobStr });
    }
  }

  let sent = 0, skipped = 0, errors = 0;
  candidatos.forEach((info, email) => {
    if (_yaEnviadoCumpleEsteAnio(email, currentYear)) {
      skipped++;
      return;
    }
    try {
      _sendCumpleEmail({ email, nombre: info.nombre, dobStr: info.dobStr }, cfg);
      _marcarCumpleEnviado(email, currentYear, info.nombre);
      sent++;
    } catch (e) {
      errors++;
      Logger.log('⚠ Error cumple ' + email + ': ' + e.message);
    }
  });
  Logger.log('🎂 Cumpleaños: ' + sent + ' enviados · ' + skipped + ' ya enviados antes · ' + errors + ' errores');
  return { sent, skipped, errors };
}

function _sendCumpleEmail(huesped, cfg) {
  const props   = PropertiesService.getScriptProperties();
  const waNum   = (props.getProperty('CONTACT_WHATSAPP_NUMBER') || '50769812266').replace(/\D/g, '');
  const md      = _parseDobMonthDay(huesped.dobStr);
  const meses   = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  const fechaCumple = md.dd + ' de ' + (meses[md.mm - 1] || '');
  const firstName = (huesped.nombre || '').split(/\s+/)[0] || huesped.nombre;
  const subject = 'Tu cumpleaños se acerca — un regalo de Las Nubes';
  const waText  = encodeURIComponent('Hola! Vi el correo de cumpleaños — quisiera reservar la noche de mi cumpleaños (' + fechaCumple + ') con el descuento.');
  const waLink  = 'https://wa.me/' + waNum + '?text=' + waText;
  const html = buildCumpleEmailHTML({
    firstName,
    nombre: huesped.nombre,
    fechaCumple,
    descMonto:  cfg.descMonto,
    diasAntes:  cfg.diasAntes,
    waLink
  });
  GmailApp.sendEmail(huesped.email, subject, '', { htmlBody: html, name: 'Las Nubes', replyTo: REPLY_TO_EMAIL });
}

function buildCumpleEmailHTML(opts) {
  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>' +
'<body style="margin:0;padding:0;background:#f5f3f0;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;">' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f3f0;padding:32px 16px;"><tr><td align="center">' +
'<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;">' +
'<tr><td style="background:#6a9e62;border-radius:16px 16px 0 0;padding:36px 40px;text-align:center;">' +
'<p style="margin:0 0 6px;font-size:12px;color:rgba(255,255,255,0.75);letter-spacing:2px;text-transform:uppercase;">&#127874; Regalo de cumpleaños</p>' +
'<h1 style="margin:0;font-size:32px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></h1>' +
'<p style="margin:10px 0 0;font-size:15px;color:rgba(255,255,255,0.9);">Buenos Aires, Chame · Panamá Oeste</p>' +
'</td></tr>' +
'<tr><td style="background:#ffffff;padding:36px 40px;">' +
'<p style="margin:0 0 16px;font-size:16px;color:#3a3530;line-height:1.6;">Hola <strong>' + (opts.firstName || 'amigo') + '</strong>,</p>' +
'<p style="margin:0 0 16px;font-size:15px;color:#3a3530;line-height:1.7;">Tu cumpleaños se acerca — el <strong>' + opts.fechaCumple + '</strong>. En Las Nubes queremos celebrarlo contigo.</p>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0f8ed;border:1px solid #b7d8b1;border-radius:14px;margin:24px 0;">' +
'<tr><td style="padding:24px 28px;">' +
'<p style="margin:0 0 10px;font-size:13px;color:#2c5e22;letter-spacing:0.05em;text-transform:uppercase;font-weight:600;">Tu regalo</p>' +
'<p style="margin:0 0 14px;font-size:26px;font-weight:300;color:#2c5e22;font-family:Georgia,serif;">$' + opts.descMonto + ' off tu noche</p>' +
'<p style="margin:0;font-size:13px;color:#4a6b40;line-height:1.6;">Sobre la noche de tu cumpleaños (' + opts.fechaCumple + '), reservando directo por WhatsApp. Se descuenta de la tarifa vigente ese día.</p>' +
'</td></tr></table>' +
'<p style="margin:0 0 20px;font-size:14px;color:#6b6560;line-height:1.7;">Cómo funciona:</p>' +
'<ul style="margin:0 0 20px;padding-left:20px;color:#6b6560;font-size:14px;line-height:1.7;">' +
'<li>Escríbenos por WhatsApp para coordinar.</li>' +
'<li>Aplica solo la noche de tu cumpleaños — el resto a tarifa normal.</li>' +
'<li>Cualquier día de la semana. No aplica en feriados, vísperas de feriado ni vacaciones escolares.</li>' +
'<li>Sujeto a disponibilidad. Reserva con tiempo.</li>' +
'<li>Solo para reservas directas (no aplica si reservas vía Airbnb).</li>' +
'</ul>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:20px;"><tr><td align="center">' +
'<a href="' + opts.waLink + '" target="_blank" style="display:inline-block;background:#25d366;color:#ffffff;font-size:15px;font-weight:600;padding:14px 28px;border-radius:10px;text-decoration:none;">&#128172; Reservar mi cumpleaños</a>' +
'</td></tr></table>' +
'<p style="margin:0;font-size:13px;color:#8a8078;line-height:1.6;text-align:center;">O abrí el calendario en <a href="https://lasnubes.cloud" style="color:#6a9e62;">lasnubes.cloud</a> para ver disponibilidad.</p>' +
'</td></tr>' +
'<tr><td style="background:#3a3530;border-radius:0 0 16px 16px;padding:24px 40px;text-align:center;">' +
'<p style="margin:0 0 8px;font-size:18px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></p>' +
'<p style="margin:0 0 12px;font-size:11px;color:rgba(255,255,255,0.6);letter-spacing:1px;text-transform:uppercase;">Cabañas en Chicá · Panamá Oeste</p>' +
'<a href="https://wa.me/50769812266" style="color:rgba(255,255,255,0.8);font-size:13px;text-decoration:none;">&#128172; WhatsApp: +507 6981-2266</a>' +
'</td></tr></table></td></tr></table></body></html>';
}

// Ejecutar UNA VEZ desde el editor. Crea el trigger diario.
function instalarTriggerCumpleanos() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'enviarRecordatoriosCumpleanos') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarRecordatoriosCumpleanos')
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .inTimezone('America/Panama')
    .create();
  Logger.log('✓ Trigger creado: enviarRecordatoriosCumpleanos diario @ 10am Panama');
}

// Preview — envia el email al admin con datos de ejemplo.
function enviarCumpleanosPrueba() {
  const myEmail = Session.getActiveUser().getEmail();
  if (!myEmail) {
    Logger.log('⚠ No se pudo obtener email del usuario activo.');
    return;
  }
  const cfg = _cumpleanosConfig();
  const target = _addDays(new Date(), cfg.diasAntes);
  const dobStr = Utilities.formatDate(target, 'America/Panama', '1990-MM-dd');
  _sendCumpleEmail({ email: myEmail, nombre: 'Maria de Prueba', dobStr }, cfg);
  Logger.log('✓ Cumple de prueba enviado a ' + myEmail);
}
