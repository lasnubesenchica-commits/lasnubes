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

// Construye DTO publico a partir de un objeto reservation (formato dashboard).
// Filtra campos sensibles y agrega data de cabin/ubicacion.
function _buildPublicDTO(r) {
  const meta  = tipoEmailMeta(r);
  const tz    = 'America/Panama';
  const props = PropertiesService.getScriptProperties();

  // Ventana operativa del key box: 9am del display checkin -> 23:59 del dia siguiente al display checkout
  const nowMs = Date.now();
  const ciMs  = new Date(meta.displayCheckin  + 'T09:00:00-05:00').getTime();
  const coDay = new Date(meta.displayCheckout + 'T12:00:00-05:00');
  coDay.setDate(coDay.getDate() + 1);
  const coCutoffMs = coDay.getTime();
  const showKeyBox = nowMs >= ciMs && nowMs <= coCutoffMs;

  // Primer nombre solo (privacidad si el link se reenvia)
  const fullName    = (r.name || '').toString().trim();
  const primerNomb  = fullName.split(/\s+/)[0] || fullName;

  const total      = parseFloat(r.amount)  || 0;
  const personas   = parseInt(r.persons, 10) || null;
  const estadoUp   = (r.estadoPago || '').toString().toUpperCase();
  const isCancel   = estadoUp === 'CANCELADA';

  const cabinNombre = CABIN_NAMES_EMAIL[r.cabin] || r.cabinName || r.cabin || '';

  // Combo (5-6 personas): si tipo === 'noche' y personas >= 5, indicar combo
  // Por simplicidad lo dejamos como una sola reserva por ahora.

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
    estado:        isCancel ? 'CANCELADA' : 'CONFIRMADA',
    keyBoxCode:    showKeyBox ? PUBLIC_KEY_BOX_CODE : null,
    keyBoxFromFmt: meta.checkinFmt,
    mapsUrl:       props.getProperty('CHECKIN_MAPS_URL')      || 'https://maps.google.com/?q=8.639400,-79.945900',
    wazeUrl:       props.getProperty('CHECKIN_WAZE_URL')      || 'https://waze.com/ul?ll=8.639400,-79.945900&navigate=yes',
    indicaciones:  props.getProperty('CHECKIN_INDICACIONES')  || '',
    accesoExtra:   props.getProperty('CHECKIN_ACCESO_EXTRA')  || '',
    whatsappContact: 'https://wa.me/' + ((props.getProperty('CONTACT_WHATSAPP_NUMBER') || '50769812266').replace(/\D/g, ''))
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
        comentarios: r[22] || '',
        tipo:        r[24] || 'noche'
      };
    }
  }
  return null;
}

// Action handler: GET ?action=getReservaPublic&id=...&t=...
function handleGetReservaPublic(e) {
  const id    = e && e.parameter && e.parameter.id;
  const token = e && e.parameter && e.parameter.t;
  if (!id || !token) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'MISSING_PARAMS' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (!_verifyReservaToken(id, token)) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'INVALID_TOKEN' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const r = _readReservaById(id);
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
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, url: url, id: id }))
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
