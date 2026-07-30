// ═══════════════════════════════════════════════════════════
//  Programa de Referidos — codigos personales + redenciones
// ═══════════════════════════════════════════════════════════
//
//  Cada huésped con estadia completada (Directa, no cancelada) recibe
//  un código personal LN-XXXXXX vía email. Lo comparte. Cuando un amigo
//  reserva mencionando el código, ambos reciben $20 off.
//
//  Sheets:
//   - Referrals:     Email | Telefono | Nombre | Code | CreatedAt | EmailSentAt
//   - ReferralUses:  Code | ReferrerEmail | UsedByEmail | UsedByReservaId
//                    | UsedAt | ReferrerCreditUsed | ReferrerCreditReservaId
//
//  Trigger diario enviarCodigosReferido(): para cada huesped elegible
//  sin codigo o sin email enviado, crea el codigo (si falta) y envia
//  el email con instrucciones de uso.
//
//  Reward: $20 al referrer (al confirmarse la estadia del referido),
//  $20 al referido (descuento aplicado al primer booking).
//
//  Activacion: ejecutar instalarTriggerReferidos() una vez.
//  Preview:    enviarReferralPrueba()
// ═══════════════════════════════════════════════════════════

const REFERRAL_CODE_ALPHABET = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';  // sin 0/O/I/l/1
const REFERRAL_CODE_LEN      = 6;
const REFERRAL_REWARD_AMOUNT = 20;
// 12 meses, igual que el credito de Cliente Fiel. Antes eran 6 y el otro 12:
// dos vencimientos distintos para el mismo tipo de credito obligaban a mirar
// de cual venia cada uno. Se unifico hacia arriba porque alargar nunca rompe
// una promesa hecha; acortar si.
const REFERRAL_CREDIT_EXPIRY_DAYS = 365;

// El email sale el dia DESPUES del check-out: la estadia esta fresca, ya
// vivieron la experiencia y "Gracias por venir" es cierto. Antes salia con
// checkin <= hoy, o sea el dia de la llegada a las 10am — cuatro horas ANTES de
// que el huesped apareciera, agradeciendole por una visita que todavia no
// ocurria. La plantilla de WhatsApp equivalente se llama `referido_postestadia`:
// el nombre delata cual era la intencion original.
//
// La ventana llega hasta 3 dias para aguantar que el trigger falle un dia sin
// que esos huespedes se queden sin codigo. El guard de emailSentAt evita que
// alguien lo reciba dos veces.
const REFERRAL_DIAS_POST_CHECKOUT_MAX = 3;

// Piso duro: no se le escribe a nadie que se haya ido antes de esta fecha.
//
// Sin esto, la funcion recorria TODA la hoja: en su primera corrida le hubiera
// mandado el email a cada huesped de la historia de Las Nubes con email
// registrado — cientos de golpe, a gente que estuvo hace meses, y con riesgo de
// quemar la cuota diaria de Gmail. La ventana de 3 dias ya lo evita; esta
// constante es la red por si alguien la vuelve a ensanchar sin pensar en esto.
const REFERRAL_NO_ANTES_DE = '2026-07-30';

// Las reglas del programa, en un solo lugar. Antes vivían solo dentro del HTML
// del email; ahora las usan también el bot de WhatsApp y la plantilla, y dos
// listas que dicen lo mismo terminan separándose en la primera corrección.
// `%M%` se reemplaza por el monto para no repetirlo en cada línea.
const REFERRAL_COMO_FUNCIONA = [
  'Comparte el código con tu amigo.',
  'Tu amigo reserva directo por WhatsApp mencionando tu código.',
  'Tu amigo recibe $%M% off en su primera estadía.',
  'Tú recibes $%M% de crédito al confirmarse su estadía. Lo aplicas cuando vuelvas.',
  'Sin tope de referidos. Acumulas crédito por cada amigo.'
];

const REFERRAL_RESTRICCIONES = [
  // Se define por TIPO de día y no por día de semana: un martes feriado o de
  // vacaciones escolares sigue siendo "Dom–Jue", y ahí el descuento caía justo
  // en las noches de mayor demanda.
  'Aplica a noches de domingo a jueves, sin feriados, vísperas de feriado ni vacaciones escolares.',
  'Solo para reservas directas (no Airbnb).',
  'No combinable con tarifa promocional ni otras promociones.',
  'El crédito vence a los 12 meses de la estadía del referido.',
  'Sujeto a disponibilidad.'
];

function _refReglas(lista, monto) {
  return lista.map(function(t) { return t.split('%M%').join(String(monto)); });
}

// Las mismas reglas en texto plano, para WhatsApp.
function referralReglasTexto(codigo, monto) {
  monto = monto || REFERRAL_REWARD_AMOUNT;
  return '🤝 *Programa Amigos*\n\n' +
    'Tu código: *' + codigo + '*\n\n' +
    '*Cómo funciona:*\n' +
    _refReglas(REFERRAL_COMO_FUNCIONA, monto).map(function(t) { return '• ' + t; }).join('\n') +
    '\n\n*Restricciones:*\n' +
    _refReglas(REFERRAL_RESTRICCIONES, monto).map(function(t) { return '• ' + t; }).join('\n');
}

function _getOrCreateReferralsSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let s = ss.getSheetByName('Referrals');
  if (!s) {
    s = ss.insertSheet('Referrals');
    s.getRange(1, 1, 1, 6).setValues([[
      'Email', 'Telefono', 'Nombre', 'Code', 'CreatedAt', 'EmailSentAt'
    ]]);
    s.getRange(1, 1, 1, 6).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

function _getOrCreateReferralUsesSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let s = ss.getSheetByName('ReferralUses');
  if (!s) {
    s = ss.insertSheet('ReferralUses');
    s.getRange(1, 1, 1, 7).setValues([[
      'Code', 'ReferrerEmail', 'UsedByEmail', 'UsedByReservaId',
      'UsedAt', 'ReferrerCreditUsed', 'ReferrerCreditReservaId'
    ]]);
    s.getRange(1, 1, 1, 7).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

function _genReferralCode() {
  let code = '';
  const N = REFERRAL_CODE_ALPHABET.length;
  for (let i = 0; i < REFERRAL_CODE_LEN; i++) {
    code += REFERRAL_CODE_ALPHABET.charAt(Math.floor(Math.random() * N));
  }
  return 'LN-' + code;
}

// Reusa el código si el huésped ya tiene uno; sino genera uno nuevo y único.
function getOrCreateReferralCode(email, telefono, nombre) {
  const e = (email    || '').toString().toLowerCase().trim();
  if (!e) return null;
  const s = _getOrCreateReferralsSheet();
  const data = s.getDataRange().getValues();
  const existingCodes = new Set();
  for (let i = 1; i < data.length; i++) {
    const re = (data[i][0] || '').toString().toLowerCase().trim();
    if (re === e) return data[i][3].toString();
    if (data[i][3]) existingCodes.add(data[i][3].toString());
  }
  let code;
  do { code = _genReferralCode(); } while (existingCodes.has(code));
  const now = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd HH:mm:ss');
  s.appendRow([e, telefono || '', nombre || '', code, now, '']);
  return code;
}

function _findReferralRow(email) {
  const e = (email || '').toString().toLowerCase().trim();
  if (!e) return null;
  const s = _getOrCreateReferralsSheet();
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().toLowerCase().trim() === e) {
      return {
        rowIndex:    i + 1,
        email:       data[i][0],
        telefono:    data[i][1],
        nombre:      data[i][2],
        code:        data[i][3],
        createdAt:   data[i][4] instanceof Date ? Utilities.formatDate(data[i][4], 'America/Panama', 'yyyy-MM-dd') : data[i][4],
        emailSentAt: data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], 'America/Panama', 'yyyy-MM-dd') : data[i][5]
      };
    }
  }
  return null;
}

// Busca el código por TELÉFONO. `getOrCreateReferralCode` exige email y
// devuelve null sin él, y el bot solo conoce el número de quien escribe. Se
// compara por los últimos 8 dígitos para que dé igual si la fila quedó guardada
// como "6981-2266", "+507 6981-2266" o "50769812266".
function findReferralCodeByPhone(telefono) {
  const norm = function(t) {
    const d = String(t || '').replace(/\D/g, '');
    return d.length > 8 ? d.slice(-8) : d;
  };
  const target = norm(telefono);
  if (!target || target.length < 7) return null;
  const data = _getOrCreateReferralsSheet().getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (!data[i][3]) continue;                     // sin código, nada que devolver
    if (norm(data[i][1]) === target) return data[i][3].toString();
  }
  return null;
}

function _findReferralByCode(code) {
  if (!code) return null;
  const c = code.toString().trim().toUpperCase();
  const s = _getOrCreateReferralsSheet();
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][3] || '').toString().trim().toUpperCase() === c) {
      return {
        rowIndex:    i + 1,
        email:       data[i][0],
        telefono:    data[i][1],
        nombre:      data[i][2],
        code:        data[i][3],
        createdAt:   data[i][4] instanceof Date ? Utilities.formatDate(data[i][4], 'America/Panama', 'yyyy-MM-dd') : data[i][4],
        emailSentAt: data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], 'America/Panama', 'yyyy-MM-dd') : data[i][5]
      };
    }
  }
  return null;
}

function _findReferralUsesForReferrer(email) {
  const e = (email || '').toString().toLowerCase().trim();
  if (!e) return [];
  const s = _getOrCreateReferralUsesSheet();
  const data = s.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if ((data[i][1] || '').toString().toLowerCase().trim() === e) {
      out.push({
        rowIndex:                i + 1,
        code:                    data[i][0],
        referrerEmail:           data[i][1],
        usedByEmail:             data[i][2],
        usedByReservaId:         data[i][3],
        usedAt:                  data[i][4] instanceof Date ? Utilities.formatDate(data[i][4], 'America/Panama', 'yyyy-MM-dd') : data[i][4],
        referrerCreditUsed:      data[i][5] === true || data[i][5] === 'TRUE' || data[i][5] === 'true',
        referrerCreditReservaId: data[i][6]
      });
    }
  }
  return out;
}

function _findReferralUseByReservaId(reservaId) {
  if (!reservaId) return null;
  const s = _getOrCreateReferralUsesSheet();
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] && data[i][3].toString() === reservaId.toString()) {
      return {
        rowIndex:                i + 1,
        code:                    data[i][0],
        referrerEmail:           data[i][1],
        usedByEmail:             data[i][2],
        usedByReservaId:         data[i][3],
        usedAt:                  data[i][4] instanceof Date ? Utilities.formatDate(data[i][4], 'America/Panama', 'yyyy-MM-dd') : data[i][4],
        referrerCreditUsed:      data[i][5] === true || data[i][5] === 'TRUE' || data[i][5] === 'true',
        referrerCreditReservaId: data[i][6]
      };
    }
  }
  return null;
}

// Trigger diario @ 10am Panama
function enviarCodigosReferido() {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();

  // Candidatos: huesped Directa/Referido que se fue AYER (o hasta hace 3 dias),
  // no cancelada, tipo dormido, con email.
  const DIA = 24 * 60 * 60 * 1000;
  const todayMs = new Date(Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd') + 'T00:00:00-05:00').getTime();
  const pisoMs  = new Date(REFERRAL_NO_ANTES_DE + 'T00:00:00-05:00').getTime();
  const candidatos = new Map(); // email -> { nombre, telefono }
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if (!['Directa','Referido'].includes(r[9] || '')) continue;
    if ((r[20] || '').toString().toUpperCase() === 'CANCELADA') continue;
    const tipo = (r[24] || 'noche').toString();
    if (tipo === 'pasatarde' || tipo === 'pasadia') continue;
    const email = (r[21] || '').toString().toLowerCase().trim();
    if (!email) continue;
    if (!r[5]) continue;  // sin checkout no podemos decidir
    const coStr = r[5] instanceof Date
      ? Utilities.formatDate(r[5], 'America/Panama', 'yyyy-MM-dd')
      : r[5].toString().slice(0, 10);
    let coMs = new Date(coStr + 'T00:00:00-05:00').getTime();
    if (isNaN(coMs)) continue;
    // `late` bloquea el dia siguiente como cortesia: el huesped se va el dia
    // ANTES del checkout guardado. Sin esto se le escribe un dia tarde.
    if (tipo === 'late') coMs -= DIA;
    if (coMs < pisoMs) continue;                      // se fue antes del piso
    const dias = Math.round((todayMs - coMs) / DIA);
    if (dias < 1 || dias > REFERRAL_DIAS_POST_CHECKOUT_MAX) continue;
    if (!candidatos.has(email)) {
      candidatos.set(email, { nombre: r[1], telefono: r[23] || '' });
    }
  }

  const refs = _getOrCreateReferralsSheet();
  let sent = 0, skipped = 0, errors = 0, wa = 0;
  candidatos.forEach((info, email) => {
    try {
      // Asegura código (si ya existe lo reusa, sino crea fila)
      const code = getOrCreateReferralCode(email, info.telefono, info.nombre);
      const row  = _findReferralRow(email);
      if (!row) { errors++; return; }
      if (row.emailSentAt) { skipped++; return; }  // ya enviado antes
      _sendReferralCodeEmail({ email, nombre: info.nombre, code });
      refs.getRange(row.rowIndex, 6).setValue(Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd HH:mm:ss'));
      sent++;

      // WhatsApp como canal adicional. Va DESPUÉS de marcar el envío y en su
      // propio try: el email es el canal principal y no puede quedar sin marcar
      // porque WhatsApp haya fallado, o el huésped recibiría el mismo email
      // todos los días hasta que la plantilla se apruebe.
      //
      // Mientras `referido_postestadia` esté en revisión, este envío falla y
      // queda en el log; el día que Meta la apruebe empieza a salir solo, sin
      // tocar nada. Por eso está acá y no esperando a la aprobación.
      if (info.telefono) {
        try {
          sendWhatsAppTemplate(info.telefono, 'referido_postestadia', 'es_ES',
            [(info.nombre || '').toString().trim().split(/\s+/)[0] || 'amigo',
             code, String(REFERRAL_REWARD_AMOUNT)],
            null, 'referido_' + code);
          wa++;
        } catch(eWa) {
          Logger.log('· WA referido ' + email + ' no salió (¿plantilla sin aprobar?): ' + eWa.message);
        }
      }
    } catch(e) {
      errors++;
      Logger.log('⚠ Error referido ' + email + ': ' + e.message);
    }
  });
  Logger.log('🤝 Referidos: ' + sent + ' emails · ' + wa + ' WhatsApp · ' +
             skipped + ' ya enviados · ' + errors + ' errores');
  return { sent, wa, skipped, errors };
}

function _sendReferralCodeEmail(opts) {
  const props = PropertiesService.getScriptProperties();
  const waNum = (props.getProperty('CONTACT_WHATSAPP_NUMBER') || '50769812266').replace(/\D/g, '');
  const firstName = (opts.nombre || '').split(/\s+/)[0] || opts.nombre;
  const subject = 'Gracias por venir — tu código de referido Las Nubes';
  const html = buildReferralCodeEmailHTML({
    firstName,
    code:     opts.code,
    amount:   REFERRAL_REWARD_AMOUNT,
    waNumber: waNum
  });
  GmailApp.sendEmail(opts.email, subject, '', { htmlBody: html, name: 'Las Nubes', replyTo: REPLY_TO_EMAIL });
}

function buildReferralCodeEmailHTML(opts) {
  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>' +
'<body style="margin:0;padding:0;background:#f5f3f0;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;">' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f3f0;padding:32px 16px;"><tr><td align="center">' +
'<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;">' +
'<tr><td style="background:#5a85b0;border-radius:16px 16px 0 0;padding:36px 40px;text-align:center;">' +
'<p style="margin:0 0 6px;font-size:12px;color:rgba(255,255,255,0.8);letter-spacing:2px;text-transform:uppercase;">&#129309; Programa Amigos</p>' +
'<h1 style="margin:0;font-size:32px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></h1>' +
'<p style="margin:10px 0 0;font-size:15px;color:rgba(255,255,255,0.9);">Buenos Aires, Chame</p>' +
'</td></tr>' +
'<tr><td style="background:#ffffff;padding:36px 40px;">' +
'<p style="margin:0 0 16px;font-size:16px;color:#3a3530;line-height:1.6;">Hola <strong>' + (opts.firstName || 'amigo') + '</strong>,</p>' +
'<p style="margin:0 0 16px;font-size:15px;color:#3a3530;line-height:1.7;">Gracias por venir a Las Nubes. Si conoces a alguien a quien le gustaría una escapada como la tuya, te dejamos un regalo:</p>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#eef4f8;border:1px solid #a8c5d8;border-radius:14px;margin:24px 0;">' +
'<tr><td style="padding:24px 28px;text-align:center;">' +
'<p style="margin:0 0 8px;font-size:13px;color:#385d7a;letter-spacing:0.05em;text-transform:uppercase;font-weight:600;">Tu código personal</p>' +
'<p style="margin:0 0 10px;font-size:32px;font-weight:500;color:#385d7a;font-family:Georgia,serif;letter-spacing:0.05em;">' + opts.code + '</p>' +
'<p style="margin:0;font-size:13px;color:#385d7a;line-height:1.6;">Compártelo. Cuando tu amigo reserve directo con nosotros mencionando este código, <strong>ambos reciben $' + opts.amount + ' off</strong> en su próxima estadía.</p>' +
'</td></tr></table>' +
'<p style="margin:0 0 14px;font-size:14px;color:#6b6560;line-height:1.7;">Cómo funciona:</p>' +
'<ul style="margin:0 0 20px;padding-left:20px;color:#6b6560;font-size:14px;line-height:1.7;">' +
_refReglas(REFERRAL_COMO_FUNCIONA, opts.amount).map(function(t) {
  // El código en negrita dentro de la línea que lo menciona.
  return '<li>' + t.replace('tu código', '<strong>' + opts.code + '</strong>') + '</li>';
}).join('') +
'</ul>' +
'<p style="margin:0 0 8px;font-size:13px;color:#8a8078;line-height:1.6;font-weight:600;">Restricciones:</p>' +
'<ul style="margin:0 0 20px;padding-left:20px;color:#8a8078;font-size:13px;line-height:1.7;">' +
_refReglas(REFERRAL_RESTRICCIONES, opts.amount).map(function(t) {
  return '<li>' + t.replace('domingo a jueves', '<strong>domingo a jueves</strong>')
                  .replace('12 meses', '<strong>12 meses</strong>') + '</li>';
}).join('') +
'</ul>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:20px;"><tr><td align="center">' +
'<a href="https://wa.me/' + opts.waNumber + '?text=' + encodeURIComponent('Hola! Quiero compartir mi código de referido ' + opts.code) + '" target="_blank" style="display:inline-block;background:#25d366;color:#ffffff;font-size:15px;font-weight:600;padding:14px 28px;border-radius:10px;text-decoration:none;">&#128172; Compartir mi código</a>' +
'</td></tr></table>' +
'</td></tr>' +
'<tr><td style="background:#3a3530;border-radius:0 0 16px 16px;padding:24px 40px;text-align:center;">' +
'<p style="margin:0 0 8px;font-size:18px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></p>' +
'<p style="margin:0 0 12px;font-size:11px;color:rgba(255,255,255,0.6);letter-spacing:1px;text-transform:uppercase;">Cabañas en Chicá</p>' +
'<a href="https://wa.me/50769812266" style="color:rgba(255,255,255,0.8);font-size:13px;text-decoration:none;">&#128172; WhatsApp: +507 6981-2266</a>' +
'</td></tr></table></td></tr></table></body></html>';
}

// Endpoint GET: lista todos los referidos + sus uses (admin)
function handleGetReferrals(e) {
  try {
    const refs = _getOrCreateReferralsSheet();
    const refsData = refs.getDataRange().getValues();
    const referrals = [];
    for (let i = 1; i < refsData.length; i++) {
      if (!refsData[i][0]) continue;
      referrals.push({
        email:       refsData[i][0],
        telefono:    refsData[i][1],
        nombre:      refsData[i][2],
        code:        refsData[i][3],
        createdAt:   refsData[i][4] instanceof Date ? Utilities.formatDate(refsData[i][4], 'America/Panama', 'yyyy-MM-dd') : refsData[i][4],
        emailSentAt: refsData[i][5] instanceof Date ? Utilities.formatDate(refsData[i][5], 'America/Panama', 'yyyy-MM-dd') : refsData[i][5]
      });
    }
    const uses = _getOrCreateReferralUsesSheet();
    const usesData = uses.getDataRange().getValues();
    const referralUses = [];
    for (let i = 1; i < usesData.length; i++) {
      if (!usesData[i][0]) continue;
      referralUses.push({
        code:                    usesData[i][0],
        referrerEmail:           usesData[i][1],
        usedByEmail:             usesData[i][2],
        usedByReservaId:         usesData[i][3],
        usedAt:                  usesData[i][4] instanceof Date ? Utilities.formatDate(usesData[i][4], 'America/Panama', 'yyyy-MM-dd') : usesData[i][4],
        referrerCreditUsed:      usesData[i][5] === true || usesData[i][5] === 'TRUE' || usesData[i][5] === 'true',
        referrerCreditReservaId: usesData[i][6]
      });
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, referrals, referralUses })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Endpoint POST: registra que una reserva uso un codigo de referido.
// payload: { code, usedByEmail, usedByReservaId }
function handleRegisterReferralUse(payload) {
  try {
    const code  = (payload.code || '').toString().trim().toUpperCase();
    const email = (payload.usedByEmail || '').toString().toLowerCase().trim();
    const rid   = payload.usedByReservaId;
    if (!code || !email || !rid) return _jsonOut({ ok: false, error: 'MISSING_PARAMS' });
    const ref = _findReferralByCode(code);
    if (!ref) return _jsonOut({ ok: false, error: 'CODE_NOT_FOUND' });
    if (ref.email && ref.email.toString().toLowerCase() === email) {
      return _jsonOut({ ok: false, error: 'SELF_REFERRAL' });
    }
    // Validar que esa reserva no haya sido registrada ya con otro codigo
    const existing = _findReferralUseByReservaId(rid);
    if (existing) return _jsonOut({ ok: false, error: 'ALREADY_REGISTERED', existing });
    const s = _getOrCreateReferralUsesSheet();
    const now = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');
    s.appendRow([code, ref.email, email, rid, now, false, '']);
    return _jsonOut({ ok: true, referrerEmail: ref.email, referrerNombre: ref.nombre, amount: REFERRAL_REWARD_AMOUNT });
  } catch(err) {
    return _jsonOut({ ok: false, error: err.message });
  }
}

// Endpoint POST: marca un credito del referrer como usado en una reserva.
// payload: { referrerEmail, usedByReservaId (el referido original), referrerCreditReservaId }
function handleMarkReferrerCreditUsed(payload) {
  try {
    const refEmail = (payload.referrerEmail || '').toString().toLowerCase().trim();
    const useRid   = payload.usedByReservaId;  // reserva del referido (identifica la fila)
    const credRid  = payload.referrerCreditReservaId; // donde se aplica el credito ahora
    if (!refEmail || !useRid || !credRid) return _jsonOut({ ok: false, error: 'MISSING_PARAMS' });
    const use = _findReferralUseByReservaId(useRid);
    if (!use) return _jsonOut({ ok: false, error: 'USE_NOT_FOUND' });
    if (use.referrerCreditUsed) return _jsonOut({ ok: false, error: 'ALREADY_USED' });
    // Validar vencimiento (UsedAt + 180 dias)
    if (use.usedAt) {
      const usedMs = new Date(use.usedAt + 'T12:00:00-05:00').getTime();
      const expiresMs = usedMs + REFERRAL_CREDIT_EXPIRY_DAYS * 24 * 60 * 60 * 1000;
      if (Date.now() > expiresMs) return _jsonOut({ ok: false, error: 'EXPIRED', expiredOn: Utilities.formatDate(new Date(expiresMs), 'America/Panama', 'yyyy-MM-dd') });
    }
    const s = _getOrCreateReferralUsesSheet();
    const now = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');
    s.getRange(use.rowIndex, 6).setValue(true);
    s.getRange(use.rowIndex, 7).setValue(credRid);
    return _jsonOut({ ok: true, usedAt: now });
  } catch(err) {
    return _jsonOut({ ok: false, error: err.message });
  }
}

function instalarTriggerReferidos() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'enviarCodigosReferido') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarCodigosReferido')
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .inTimezone('America/Panama')
    .create();
  Logger.log('✓ Trigger creado: enviarCodigosReferido diario @ 10am Panama');
}

function enviarReferralPrueba() {
  const myEmail = Session.getActiveUser().getEmail();
  if (!myEmail) { Logger.log('⚠ Email no disponible.'); return; }
  _sendReferralCodeEmail({ email: myEmail, nombre: 'Maria de Prueba', code: 'LN-' + 'PRUEBA' });
  Logger.log('✓ Referral de prueba enviado a ' + myEmail);
}
