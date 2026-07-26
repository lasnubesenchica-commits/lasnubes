// ═══════════════════════════════════════════════════════════
//  Public reservation link
//
//  Link firmado con HMAC-SHA256 que se le comparte al cliente
//  (via email o WhatsApp) para que vea los detalles de su reserva
//  desde una pagina publica (reserva.html) sin login.
//
//  El link incluye:   ?id=<reservaId>&t=<hmac(id)>
//  El secret se guarda en Script Properties: PUBLIC_LINK_SECRET
//  El base URL se puede sobreescribir con PUBLIC_RESERVA_URL_BASE
//
//  La pagina publica ve datos limpios: nombre (primer nombre),
//  cabaña, fechas, horas, total, tipo. NO ve: telefono, email,
//  voucher, codigo de transferencia, notas internas, monto pagado.
//
//  El codigo del key box se revela solo entre las 9am del check-in
//  y las 11:59pm del dia siguiente al check-out (ventana operativa).
// ═══════════════════════════════════════════════════════════

const PUBLIC_LINK_URL_DEFAULT = 'https://lasnubes.cloud/reserva.html';
const PUBLIC_KEY_BOX_CODE     = '0507';  // ya hardcodeado en buildGuiaHTML

function _publicLinkSecret() {
  const s = PropertiesService.getScriptProperties().getProperty('PUBLIC_LINK_SECRET');
  if (!s) throw new Error('PUBLIC_LINK_SECRET no esta configurado en Script Properties. Generar con: openssl rand -hex 32');
  return s;
}

function _publicLinkBaseUrl() {
  return PropertiesService.getScriptProperties().getProperty('PUBLIC_RESERVA_URL_BASE') || PUBLIC_LINK_URL_DEFAULT;
}

function _signReservaId(id) {
  const sig = Utilities.computeHmacSha256Signature(String(id), _publicLinkSecret());
  // base64 url-safe sin padding
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/, '');
}

function _verifyReservaToken(id, token) {
  if (!id || !token) return false;
  try {
    return _signReservaId(id) === token;
  } catch (e) {
    return false;
  }
}

function getPublicReservaUrl(id) {
  const t = _signReservaId(id);
  return _publicLinkBaseUrl() + '?id=' + encodeURIComponent(String(id)) + '&t=' + t;
}

// El dashboard manda el id placeholder 'preview' cuando la reserva del
// formulario todavía no se guardó (botón "Vista previa"). Ese id se firma
// perfecto, pero no existe como fila en la hoja, así que el link terminaba en
// "Link no válido". Se trata como "todavía sin id".
function _esIdPreview(id) {
  const s = (id == null ? '' : String(id)).trim().toLowerCase();
  return !s || s === 'preview';
}

// Link público de la reserva, o '' si todavía no hay un id real al que apuntar.
// Úsalo en vez de getPublicReservaUrl() en cualquier plantilla (email/WhatsApp):
// así una vista previa omite el botón en lugar de ofrecer un link roto.
function getPublicReservaUrlSafe(id) {
  if (_esIdPreview(id)) return '';
  try {
    return getPublicReservaUrl(id);
  } catch(e) {
    Logger.log('no publicLink para id=' + id + ': ' + e.message);
    return '';
  }
}

// ═══════════════════════════════════════════════════════════
//  Short codes (link mas corto via lookup en sheet ShareLinks)
// ═══════════════════════════════════════════════════════════

// Alfabeto sin caracteres ambiguos (no 0/O/I/l/1)
const SHARE_CODE_ALPHABET = '23456789ABCDEFGHJKLMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz';
const SHARE_CODE_LEN      = 6;

function _shareLinksSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let s = ss.getSheetByName('ShareLinks');
  if (!s) {
    s = ss.insertSheet('ShareLinks');
    s.getRange(1, 1, 1, 3).setValues([['Code', 'ReservaId', 'CreatedAt']]);
    s.getRange(1, 1, 1, 3).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

function _genShareCode() {
  let code = '';
  const N = SHARE_CODE_ALPHABET.length;
  for (let i = 0; i < SHARE_CODE_LEN; i++) code += SHARE_CODE_ALPHABET.charAt(Math.floor(Math.random() * N));
  return code;
}

function getOrCreateShareCode(reservaId) {
  const s = _shareLinksSheet();
  const data = s.getDataRange().getValues();
  const existingCodes = new Set();
  const sid = reservaId.toString();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString() === sid) {
      return data[i][0].toString(); // ya existe
    }
    if (data[i][0]) existingCodes.add(data[i][0].toString());
  }
  // Generar uno unico
  let code;
  do { code = _genShareCode(); } while (existingCodes.has(code));
  const now = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd HH:mm:ss');
  s.appendRow([code, sid, now]);
  return code;
}

function lookupReservaIdByShareCode(code) {
  if (!code) return null;
  const s = _shareLinksSheet();
  const data = s.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === code.toString()) {
      return data[i][1] ? data[i][1].toString() : null;
    }
  }
  return null;
}

function getPublicReservaShortUrl(reservaId) {
  const code = getOrCreateShareCode(reservaId);
  return _publicLinkBaseUrl() + '?c=' + encodeURIComponent(code);
}

// ── Prueba desde el editor ──────────────────────────────────────
// Genera un short link real para la reserva mas reciente (o para el id
// que le pases) y lo loguea, verificando el round-trip codigo→id.
// Uso: correr generarShortLinkPrueba() y mirar el Logger / Ejecuciones.
function generarShortLinkPrueba(reservaId) {
  if (!reservaId) {
    const sheet = getOrCreateSheet();
    const data  = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0]) { reservaId = data[i][0].toString(); break; }
    }
  }
  if (!reservaId) { Logger.log('⚠️ No hay reservas en la hoja.'); return null; }

  const r         = _readReservaById(reservaId);
  const code      = getOrCreateShareCode(reservaId);
  const shortUrl  = _publicLinkBaseUrl() + '?c=' + encodeURIComponent(code);
  let   longUrl   = '';
  try { longUrl = getPublicReservaUrl(reservaId); } catch(_) {}
  const roundTrip = lookupReservaIdByShareCode(code);

  Logger.log('Reserva: ' + reservaId + (r ? ' · ' + (r.name || '?') + ' · ' + (r.cabin || '?') : ' — ⚠️ NO encontrada en Reservas'));
  Logger.log('Short link: ' + shortUrl);
  Logger.log('Long link : ' + longUrl);
  Logger.log('Round-trip codigo→id: ' + roundTrip + (String(roundTrip) === String(reservaId) ? ' ✓' : ' ✗ MISMATCH'));
  return { reservaId: reservaId, shortUrl: shortUrl, longUrl: longUrl, found: !!r, roundTrip: roundTrip };
}

// Hora de inicio/fin del evento de calendario segun tipo (Panama UTC-5).
function _getEventTimes(r) {
  const meta = tipoEmailMeta(r);
  const tipo = meta.tipo;
  let startH, startM, endH, endM;
  if (tipo === 'pasatarde')      { startH = '12'; startM = '30'; endH = '19'; endM = '00'; }
  else if (tipo === 'pasanoche') { startH = '20'; startM = '00'; endH = '12'; endM = '30'; }
  else if (tipo === 'pasadia')   { startH = '09'; startM = '00'; endH = '17'; endM = '00'; }
  else if (tipo === 'early')     { startH = '09'; startM = '00'; endH = '11'; endM = '00'; }
  else if (tipo === 'late')      { startH = '14'; startM = '00'; endH = '16'; endM = '00'; }
  else                           { startH = '14'; startM = '00'; endH = '11'; endM = '00'; }
  return { startDate: meta.displayCheckin, endDate: meta.displayCheckout, startH, startM, endH, endM, tipo };
}

function _buildGcalUrl(r) {
  const et       = _getEventTimes(r);
  const meta     = tipoEmailMeta(r);
  const cabin    = CABIN_NAMES_EMAIL[r.cabin] || r.cabin;
  const title    = encodeURIComponent('Las Nubes — ' + cabin);
  const details  = encodeURIComponent('Reserva en Las Nubes\nCabaña: ' + cabin + '\nPersonas: ' + (r.persons || '') + '\nCheck-in: ' + meta.checkinHora + ' | Check-out: ' + meta.checkoutHora + '\nWhatsApp: +507 6981-2266');
  const loc      = encodeURIComponent('Buenos Aires, Chame, Panamá Oeste');
  const start    = toCalDateTime(et.startDate, et.startH, et.startM);
  const end      = toCalDateTime(et.endDate,   et.endH,   et.endM);
  return 'https://calendar.google.com/calendar/render?action=TEMPLATE&text=' + title + '&dates=' + start + '/' + end + '&details=' + details + '&location=' + loc;
}

function _buildIcsFor(r) {
  const et    = _getEventTimes(r);
  const meta  = tipoEmailMeta(r);
  const cabin = CABIN_NAMES_EMAIL[r.cabin] || r.cabin;
  const start = toCalDateTime(et.startDate, et.startH, et.startM);
  const end   = toCalDateTime(et.endDate,   et.endH,   et.endM);
  const now   = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  return [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Las Nubes//Reserva//ES',
    'BEGIN:VEVENT',
    'DTSTART:' + start,
    'DTEND:' + end,
    'SUMMARY:Las Nubes — ' + cabin,
    'DESCRIPTION:Cabaña: ' + cabin + '\\nPersonas: ' + (r.persons || '') + '\\nCheck-in: ' + meta.checkinHora + ' | Check-out: ' + meta.checkoutHora + '\\nContacto: +507 6981-2266',
    'LOCATION:Buenos Aires\\, Chame\\, Panamá Oeste',
    'DTSTAMP:' + now,
    'UID:lasnubes-' + r.id + '@lasnubes',
    'END:VEVENT',
    'END:VCALENDAR'
  ].join('\r\n');
}

// Pasos de la guia de cabaña. Fuente unica usada por email y pagina publica.
function getCabinGuideSteps(cabin, tipo, checkoutExtendido, horaSalidaCustom) {
  tipo = tipo || 'noche';
  const checkoutTitleMap = {
    'noche':     'Check-out · 11:00 am',
    'pasatarde': 'Salida · 7:00 pm',
    'pasanoche': 'Check-out · 12:30 pm día siguiente',
    'pasadia':   'Salida · 5:00 pm',
    'early':     'Check-out · 11:00 am',
    'late':      'Check-out · 4:00 pm'
  };
  // Cortesia: extender el check-out a 12:30pm cuando aplica (noche, early)
  if (checkoutExtendido && (tipo === 'noche' || tipo === 'early')) {
    checkoutTitleMap.noche = 'Check-out · 12:30 pm (cortesía)';
    checkoutTitleMap.early = 'Check-out · 12:30 pm (cortesía)';
  }
  // Override manual con hora custom de salida (pisa cortesía y default).
  const customSalida = (typeof _normalizeHora === 'function') ? _normalizeHora(horaSalidaCustom) : (horaSalidaCustom || '');
  let checkoutTitle;
  if (customSalida) {
    const isPasadiaTipo = (tipo === 'pasatarde' || tipo === 'pasadia');
    const label = isPasadiaTipo ? 'Salida' : 'Check-out';
    const nextDay = (tipo === 'pasanoche') ? ' día siguiente' : '';
    checkoutTitle = label + ' · ' + _formatHora12(customSalida) + nextDay;
  } else {
    checkoutTitle = checkoutTitleMap[tipo] || checkoutTitleMap.noche;
  }
  const checkoutBody  = 'Dejar la cocina limpia · Llevarse la basura · Cerrar la puerta y dejar la llave dentro del key box.';

  const steps = {
    verde: [
      ['&#128273;', 'Acceso', 'Key Box código <strong>0507</strong>. Dentro está la llave para acceder a la cabaña y un control negro con botones verdes para abrir el portón de la entrada en el caso de que deseen salir del proyecto.'],
      ['&#9728;&#65039;', 'Iluminación', 'Energía solar — luces encienden solas al anochecer.<br><strong>Guirnaldas del baño</strong>: se encienden solas en la noche y se apagan con el botón de encendido/apagado detrás del panel solar encima de la mesa de madera en el baño.<br><strong>Control blanco</strong>: lámparas de cocina y recámara.'],
      ['&#128267;', 'Carga de dispositivos', 'Inversor en la recámara para celulares y dispositivos.'],
      ['&#127859;', 'Cocina', 'Guarda todos los alimentos — no dejar nada expuesto para evitar animalitos.'],
      ['&#127925;', 'Convivencia', 'Música y conversaciones a volumen moderado.'],
      ['&#9989;', checkoutTitle, checkoutBody]
    ],
    azul: [
      ['&#128273;', 'Acceso', 'Key Box código <strong>0507</strong>. Dentro está la llave para acceder a la cabaña y un control negro con botones verdes para abrir el portón de la entrada en el caso de que deseen salir del proyecto. Con la llave abres el candado y luego deslizas la puerta corrediza de metal.'],
      ['&#9728;&#65039;', 'Iluminación', 'Luces del comedor y jardines encienden automáticamente al anochecer.<br><strong>Control blanco</strong> encima de la mesa verde: luces de recámara y baño.'],
      ['&#128267;', 'Carga de dispositivos', 'Powerbank en la recámara para celulares.'],
      ['&#127859;', 'Cocina', 'Guarda todos los alimentos — no dejar nada expuesto para evitar animalitos.'],
      ['&#127925;', 'Convivencia', 'Música y conversaciones a volumen moderado.'],
      ['&#9989;', checkoutTitle, checkoutBody]
    ],
    lila: [
      ['&#128273;', 'Acceso', 'Key Box código <strong>0507</strong>. Dentro está la llave para acceder a la cabaña y un control negro con botones verdes para abrir el portón de la entrada en el caso de que deseen salir del proyecto.'],
      ['&#9728;&#65039;', 'Iluminación', 'Energía solar — guirnaldas del columpio encienden entre 6:30–7:00 pm automáticamente.<br><strong>Control blanco</strong> en la cómoda frente al espejo: luces de recámara, terraza y cocina.'],
      ['&#128267;', 'Carga de dispositivos', 'Inversor en la recámara para celulares y dispositivos.'],
      ['&#127859;', 'Cocina', 'Guarda todos los alimentos — no dejar nada expuesto para evitar animalitos.'],
      ['&#127925;', 'Convivencia', 'Música y conversaciones a volumen moderado.'],
      ['&#9989;', checkoutTitle, checkoutBody]
    ]
  };

  const list = steps[cabin] || steps.verde;
  return list.map(s => ({ icon: s[0], title: s[1], body: s[2] }));
}

// Construye DTO publico a partir de un objeto reservation (formato dashboard).
// Filtra campos sensibles y agrega data de cabin/ubicacion.
function _buildPublicDTO(r) {
  // Reservas Abiertas: sin fechas/cabaña confirmadas. DTO minimo con WA contacto.
  if (r.origin === 'Abierta') {
    const propsA = PropertiesService.getScriptProperties();
    const waNumA = (propsA.getProperty('CONTACT_WHATSAPP_NUMBER') || '50769812266').replace(/\D/g, '');
    const fullNameA   = (r.name || '').toString().trim();
    const primerNombA = fullNameA.split(/\s+/)[0] || fullNameA;
    const waTextA = 'Hola! Quisiera hacer efectiva mi reserva Abierta a nombre de ' + fullNameA + '.';
    return {
      nombre:          primerNombA,
      estado:          'ABIERTA',
      whatsappContact: 'https://wa.me/' + waNumA + '?text=' + encodeURIComponent(waTextA)
    };
  }

  const meta  = tipoEmailMeta(r);
  const tz    = 'America/Panama';
  const props = PropertiesService.getScriptProperties();
  const waNum = (props.getProperty('CONTACT_WHATSAPP_NUMBER') || '50769812266').replace(/\D/g, '');

  // Para mostrar el codigo: haber subido ID (en esta reserva o en cualquier
  // reserva previa del mismo huesped por email/telefono). No hay gate de tiempo:
  // si la identidad esta verificada, el codigo aparece.
  let idUploaded = !!r.idHuespedURL;
  if (!idUploaded) {
    try {
      const prev = _findExistingHuespedId(r.email, r.telefono);
      if (prev && prev.url) idUploaded = true;
    } catch(e) { Logger.log('warn _findExistingHuespedId: ' + e.message); }
  }
  const showKeyBox = idUploaded;

  // Primer nombre solo (privacidad si el link se reenvia)
  const fullName    = (r.name || '').toString().trim();
  const primerNomb  = fullName.split(/\s+/)[0] || fullName;

  // Certificado de regalo: el beneficiario NO debe ver plata (ni tarifa, ni
  // abono, ni saldo). Se anulan los montos antes de armar el payload; reserva.html
  // ya oculta la fila Total y la tarjeta de saldo cuando vienen en cero/null.
  const regalo     = (typeof _parseRegalo === 'function') ? _parseRegalo(r) : { esRegalo: false, de: '', mensaje: '' };
  const total      = regalo.esRegalo ? 0 : (parseFloat(r.amount)  || 0);
  const deposit    = regalo.esRegalo ? 0 : (parseFloat(r.deposit) || 0);
  const saldo      = total - deposit;
  const hasSaldo   = total > 0 && saldo > 0;
  const personas   = parseInt(r.persons, 10) || null;
  const estadoUp   = (r.estadoPago || '').toString().toUpperCase();
  const isCancel   = estadoUp === 'CANCELADA';

  const cabinNombre = CABIN_NAMES_EMAIL[r.cabin] || r.cabinName || r.cabin || '';

  // Calendario: Google Calendar link + ICS para Apple/iCal
  let gcalUrl = '';
  let icsBase64 = '';
  try {
    gcalUrl   = _buildGcalUrl(r);
    icsBase64 = Utilities.base64Encode(_buildIcsFor(r));
  } catch(e) { Logger.log('warn cal links: ' + e.message); }

  // Codigo de referido del huesped (auto-crear si elegible y aún no existe).
  // Eligible: tiene email + reserva Directa + tipo dormido + no cancelada.
  let referralCode = null;
  try {
    const tipoRaw   = (r.tipo || 'noche').toString();
    const isDormido = !(tipoRaw === 'pasatarde' || tipoRaw === 'pasadia');
    const eligible  = !!r.email && ['Directa','Referido'].includes(r.origin || '')
                      && (r.estadoPago || '').toString().toUpperCase() !== 'CANCELADA'
                      && isDormido;
    if (eligible) {
      referralCode = getOrCreateReferralCode(r.email, r.telefono, r.name);
    }
  } catch(e) { Logger.log('warn referral: ' + e.message); }

  // Pasos de la guia de la cabaña
  let cabinGuide = [];
  try {
    cabinGuide = getCabinGuideSteps(r.cabin, meta.tipo, !!r.checkoutExtendido, meta.horaSalidaCustom);
    // Ocultar el codigo del key box dentro de la guia tambien cuando no estamos en ventana operativa
    if (!showKeyBox) {
      const mask = 'El código del Key Box aparece más arriba el día de tu entrada.';
      cabinGuide = cabinGuide.map(s => ({
        icon:  s.icon,
        title: s.title,
        body:  s.body.replace(/Key Box código <strong>\d+<\/strong>\./, mask)
      }));
    }
  } catch(e) { Logger.log('warn cabinGuide: ' + e.message); }

  // Saldo pendiente con WhatsApp pre-armado
  const pagarUrl = hasSaldo
    ? ('https://wa.me/' + waNum + '?text=' + encodeURIComponent('Deseo cancelar el saldo restante de mi reserva del día ' + meta.checkinFmt + ' en la cabaña ' + cabinNombre + '. ¿Me comparte los métodos de pago?'))
    : '';

  return {
    nombre:        primerNomb,
    cabin:         r.cabin || '',
    cabinName:     cabinNombre,
    tipo:          meta.tipo,
    tipoLabel:     meta.estanciaValue, // "Pasatarde" | "Pasadia" | numero | "Noches"
    estanciaLabel: meta.estanciaLabel, // "Noches" | "Estancia"
    checkin:       meta.displayCheckin,
    checkout:      meta.displayCheckout,
    checkinFmt:    meta.checkinFmt,
    checkoutFmt:   meta.checkoutFmt,
    checkinHora:   meta.checkinHora,
    checkoutHora:  meta.checkoutHora,
    personas:      personas,
    total:         total,
    abono:         deposit > 0 ? deposit : null,
    saldo:         hasSaldo ? saldo : null,
    pagarUrl:      pagarUrl,
    estado:        isCancel ? 'CANCELADA' : 'CONFIRMADA',
    keyBoxCode:    showKeyBox ? PUBLIC_KEY_BOX_CODE : null,
    keyBoxFromFmt: meta.checkinFmt,
    idUploaded:    idUploaded,
    gcalUrl:       gcalUrl,
    icsBase64:     icsBase64,
    cabinGuide:    cabinGuide,
    mapsUrl:       props.getProperty('CHECKIN_MAPS_URL')      || 'https://maps.google.com/?q=8.639400,-79.945900',
    wazeUrl:       props.getProperty('CHECKIN_WAZE_URL')      || 'https://waze.com/ul?ll=8.639400,-79.945900&navigate=yes',
    indicaciones:  props.getProperty('CHECKIN_INDICACIONES')  || '',
    accesoExtra:   props.getProperty('CHECKIN_ACCESO_EXTRA')  || '',
    whatsappContact: 'https://wa.me/' + waNum,
    referralCode:    referralCode,
    referralAmount:  20,
    // Regalo: banner "de parte de X" + dedicatoria en la página pública.
    regalo:          regalo.esRegalo ? { de: regalo.de, mensaje: regalo.mensaje } : null
  };
}

// Lee una fila de la hoja Reservas y la convierte al formato { name, cabin, ... }
// que usa el resto del codigo (ej. getReservations en doGet).
function _readReservaById(id) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if (r[0].toString() === id.toString()) {
      return {
        id:          r[0],
        name:        r[1],
        cabinName:   r[2],
        cabin:       r[3],
        checkin:     r[4] instanceof Date ? Utilities.formatDate(r[4], 'America/Panama', 'yyyy-MM-dd') : r[4],
        checkout:    r[5] instanceof Date ? Utilities.formatDate(r[5], 'America/Panama', 'yyyy-MM-dd') : r[5],
        persons:     r[6],
        amount:      r[7],
        deposit:     r[8] || 0,
        origin:      r[9],
        confirmCode: r[10],
        estadoPago:  r[20] || '',
        email:       r[21] || '',
        comentarios: r[22] || '',
        telefono:    r[23] || '',
        tipo:        r[24] || 'noche',
        idHuespedURL: r[26] || '',
        fechaNacimiento: r[27] || '',
        checkoutExtendido: r[28] === true || r[28] === 'TRUE' || r[28] === 'true' || r[28] === 1,
        // Cols 30/31: horas custom. Sin esto, la página pública (reserva.html)
        // mostraba el default del tipo (2pm/11am) en vez del horario especial.
        horaEntrada: (typeof _normalizeHora === 'function') ? _normalizeHora(r[29]) : (r[29] || ''),
        horaSalida:  (typeof _normalizeHora === 'function') ? _normalizeHora(r[30]) : (r[30] || ''),
        // Col 33: certificado de regalo. Sin esto la página pública mostraría
        // la tarifa y el saldo al beneficiario del regalo.
        regalo:      r[32] || '',
        pagador:     r[14] || ''
      };
    }
  }
  return null;
}

// Action handler: GET ?action=getReservaPublic&c=<short> OR ?id=...&t=<hmac>
// Soporta dos formatos:
//   1) Short code (c=): se busca en sheet ShareLinks, link reciente.
//   2) HMAC (id=&t=): formato anterior, sigue funcionando para emails ya enviados.
function handleGetReservaPublic(e) {
  const params = (e && e.parameter) || {};
  let reservaId = null;

  if (params.c) {
    reservaId = lookupReservaIdByShareCode(params.c);
    if (!reservaId) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: 'NOT_FOUND' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } else if (params.id && params.t) {
    if (!_verifyReservaToken(params.id, params.t)) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: 'INVALID_TOKEN' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    reservaId = params.id;
  } else {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'MISSING_PARAMS' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const r = _readReservaById(reservaId);
  if (!r) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'NOT_FOUND' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const dto = _buildPublicDTO(r);
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, reserva: dto }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Action handler: GET ?action=getReservaLink&id=...
// No auth — la URL del API ya es semi-publica y el link solo expone datos
// del viaje (no contacto, no financiero). Aceptable para el modelo actual.
// Diagnostico temporal: busca una reserva por id y reporta que encontro.
// GET ?action=debugFindReserva&id=XXX
function handleDebugFindReserva(e) {
  try {
    const id = e && e.parameter && e.parameter.id;
    if (!id) return _jsonOut({ ok: false, error: 'MISSING_ID' });
    const sheet = getOrCreateSheet();
    const data  = sheet.getDataRange().getValues();
    const idStr = id.toString();
    const out   = { ok: true, idQueried: idStr, totalRows: data.length - 1, exactMatchCol0: null, exactMatchCol10: null, partialMatches: [] };
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      const c0 = r[0];
      const c10 = r[10];
      const c0Str  = c0  != null ? c0.toString()  : '';
      const c10Str = c10 != null ? c10.toString() : '';
      if (c0Str === idStr) {
        out.exactMatchCol0 = { row: i + 1, value: c0Str, type: typeof c0, name: r[1], cabin: r[3], checkin: r[4] instanceof Date ? Utilities.formatDate(r[4],'America/Panama','yyyy-MM-dd') : (r[4]||'').toString() };
      }
      if (c10Str === idStr && !out.exactMatchCol10) {
        out.exactMatchCol10 = { row: i + 1, valueCol0: c0Str, typeCol0: typeof c0, valueCol10: c10Str, name: r[1] };
      }
      // partial matches (id contained in either col)
      if (c0Str && (c0Str.indexOf(idStr) >= 0 || idStr.indexOf(c0Str) >= 0) && c0Str !== idStr) {
        out.partialMatches.push({ row: i + 1, source: 'col0', value: c0Str, type: typeof c0, name: r[1] });
      }
      if (c10Str && (c10Str.indexOf(idStr) >= 0 || idStr.indexOf(c10Str) >= 0) && c10Str !== idStr && c0Str !== idStr) {
        out.partialMatches.push({ row: i + 1, source: 'col10', value: c10Str, type: typeof c10, name: r[1] });
      }
    }
    return _jsonOut(out);
  } catch(err) {
    return _jsonOut({ ok: false, error: err.message, stack: err.stack });
  }
}

function handleGetReservaLink(e) {
  const id = e && e.parameter && e.parameter.id;
  if (!id) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'MISSING_ID' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // Verificar que la reserva existe (no generamos links para ids inventados)
  const r = _readReservaById(id);
  if (!r) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'NOT_FOUND' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const url = getPublicReservaUrl(id);
  // Si el huesped es elegible (Directa, no cancelado, tipo dormido, con email),
  // get-or-create su codigo de referido para incluirlo en el mensaje del admin.
  let referralCode = null;
  try {
    const tipoRaw   = (r.tipo || 'noche').toString();
    const isDormido = !(tipoRaw === 'pasatarde' || tipoRaw === 'pasadia');
    const eligible  = !!r.email && ['Directa','Referido'].includes(r.origin || '')
                      && (r.estadoPago || '').toString().toUpperCase() !== 'CANCELADA'
                      && isDormido;
    if (eligible) referralCode = getOrCreateReferralCode(r.email, r.telefono, r.name);
  } catch(err) { Logger.log('warn referralCode in getReservaLink: ' + err.message); }
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, url: url, id: id, referralCode: referralCode, referralAmount: 20 }))
    .setMimeType(ContentService.MimeType.JSON);
}


// Genera el secret si no existe — correr una sola vez desde el editor.
function inicializarPublicLinkSecret() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('PUBLIC_LINK_SECRET')) {
    Logger.log('✓ PUBLIC_LINK_SECRET ya existe, no se sobreescribe');
    return;
  }
  // Generar 32 bytes random como base64
  const bytes = [];
  for (let i = 0; i < 32; i++) bytes.push(Math.floor(Math.random() * 256));
  const secret = Utilities.base64Encode(bytes);
  props.setProperty('PUBLIC_LINK_SECRET', secret);
  Logger.log('✓ PUBLIC_LINK_SECRET generado');
}
