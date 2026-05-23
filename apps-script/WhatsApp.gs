/**
 * WhatsApp Business Cloud API — sender + helpers
 *
 * Credenciales en Script Properties:
 *   WA_ACCESS_TOKEN         — token permanente del usuario del sistema (Meta Business)
 *   WA_PHONE_NUMBER_ID      — id numerico del numero registrado en WA Business
 *   WA_BUSINESS_ACCOUNT_ID  — id de la WhatsApp Business Account
 */

function _waProps() {
  const p = PropertiesService.getScriptProperties();
  return {
    token:      p.getProperty('WA_ACCESS_TOKEN'),
    phoneId:    p.getProperty('WA_PHONE_NUMBER_ID'),
    businessId: p.getProperty('WA_BUSINESS_ACCOUNT_ID')
  };
}

// Normaliza un telefono al formato E.164 sin "+" (lo que espera la API).
// Acepta "6981-2266", "+507 6981-2266", "507 6981 2266", etc.
// Si tiene 8 digitos asume Panama (prefijo 507).
function _waNormalizePhone(raw) {
  if (!raw) return null;
  const digits = String(raw).replace(/\D/g, '');
  if (digits.length === 8)  return '507' + digits;
  if (digits.length === 11 && digits.indexOf('507') === 0) return digits;
  if (digits.length >= 10)  return digits;
  return null;
}

/**
 * Envia un mensaje de texto plano por WhatsApp Cloud API.
 * Solo funciona dentro de la ventana de 24h tras un mensaje del cliente
 * (mensajes de sesion). Para iniciar conversaciones se necesita una
 * plantilla HSM aprobada — ver sendWhatsAppTemplate().
 *
 * @param {string} toPhone   telefono del destinatario (cualquier formato)
 * @param {string} body      texto del mensaje
 * @return {Object}          response JSON de la API
 */
function sendWhatsAppText(toPhone, body) {
  const cfg = _waProps();
  if (!cfg.token || !cfg.phoneId) throw new Error('WA: faltan credenciales en Script Properties');
  const to = _waNormalizePhone(toPhone);
  if (!to) throw new Error('WA: telefono invalido: ' + toPhone);

  const url = 'https://graph.facebook.com/v21.0/' + cfg.phoneId + '/messages';
  const payload = {
    messaging_product: 'whatsapp',
    to: to,
    type: 'text',
    text: { body: body }
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: { Authorization: 'Bearer ' + cfg.token },
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const text = res.getContentText();
  let json;
  try { json = JSON.parse(text); } catch(_) { json = { raw: text }; }
  logDebugEntry('WA-send-text', { to: to, code: code, ok: code >= 200 && code < 300, id: (json.messages && json.messages[0] && json.messages[0].id) || null, error: json.error || null });
  if (code < 200 || code >= 300) {
    throw new Error('WA send failed HTTP ' + code + ': ' + text.slice(0, 400));
  }
  return json;
}

/**
 * Test: corre desde el editor de Apps Script con tu propio numero.
 * Si funciona, te llega un mensaje al WhatsApp.
 *
 * IMPORTANTE: en modo desarrollo, solo podes enviar a numeros que esten
 * en la lista de "destinatarios de prueba" en Meta Business Manager
 * (WhatsApp > API Setup > seccion "Para").
 */
function testEnviarWhatsAppPrueba() {
  // Cambia este numero al tuyo (el que registraste como destinatario de prueba)
  const numeroDestino = '50769812266';
  const resultado = sendWhatsAppText(numeroDestino, 'Hola, prueba desde Apps Script ✅\n\nSi recibes esto, la API de WhatsApp Cloud esta funcionando.');
  Logger.log('Respuesta: ' + JSON.stringify(resultado));
  return resultado;
}

/**
 * Envia un mensaje basado en plantilla HSM aprobada por Meta.
 * Soporta variables nombradas (parameter_name) — formato nuevo de Meta.
 *
 * @param {string} toPhone        telefono del destinatario
 * @param {string} templateName   nombre exacto de la plantilla aprobada
 * @param {string} languageCode   p.ej. 'es_PA', 'es_MX', 'es_ES'
 * @param {Object} namedParams    { nombre: 'María', cabana: '...', ... }
 * @param {Array}  [buttonParams] opcional, parametros del boton CTA dinamico
 * @return {Object}               response JSON
 */
function sendWhatsAppTemplate(toPhone, templateName, languageCode, namedParams, buttonParams) {
  const cfg = _waProps();
  if (!cfg.token || !cfg.phoneId) throw new Error('WA: faltan credenciales en Script Properties');
  const to = _waNormalizePhone(toPhone);
  if (!to) throw new Error('WA: telefono invalido: ' + toPhone);

  const components = [];
  if (namedParams && Object.keys(namedParams).length > 0) {
    components.push({
      type: 'body',
      parameters: Object.keys(namedParams).map(name => ({
        type: 'text',
        parameter_name: name,
        text: String(namedParams[name])
      }))
    });
  }
  if (buttonParams && buttonParams.length > 0) {
    components.push({
      type: 'button',
      sub_type: 'url',
      index: '0',
      parameters: buttonParams.map(v => ({ type: 'text', text: String(v) }))
    });
  }

  const url = 'https://graph.facebook.com/v21.0/' + cfg.phoneId + '/messages';
  const payload = {
    messaging_product: 'whatsapp',
    to: to,
    type: 'template',
    template: {
      name: templateName,
      language: { code: languageCode || 'es_PA' },
      components: components
    }
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: { Authorization: 'Bearer ' + cfg.token },
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const text = res.getContentText();
  let json;
  try { json = JSON.parse(text); } catch(_) { json = { raw: text }; }
  logDebugEntry('WA-send-template', {
    to: to, template: templateName, code: code, ok: code >= 200 && code < 300,
    id: (json.messages && json.messages[0] && json.messages[0].id) || null,
    error: json.error || null
  });
  if (code < 200 || code >= 300) {
    throw new Error('WA template send failed HTTP ' + code + ': ' + text.slice(0, 400));
  }
  return json;
}

/**
 * Envia la plantilla 'confirmacion_reserva' construyendo los parametros
 * desde un objeto reservation (mismo formato que usa el resto del codigo).
 */
function sendWAReservaConfirmada(reservation) {
  if (!reservation) throw new Error('WA: reservation requerida');
  if (!reservation.telefono) throw new Error('WA: reservation sin telefono');

  const meta = tipoEmailMeta(reservation);
  const CABIN_NAMES = {
    verde: 'Paseo por Las Nubes',
    azul:  'Portal hacia Las Nubes',
    lila:  'Puente entre Las Nubes'
  };
  const cabin = CABIN_NAMES[reservation.cabin] || reservation.cabin || 'Las Nubes';

  // Linea de fechas segun tipo
  let fechas;
  const tipo = meta.tipo;
  if (tipo === 'pasatarde')      fechas = meta.checkinFmt + ' · Pasatarde 12:30pm – 7pm';
  else if (tipo === 'pasadia')   fechas = meta.checkinFmt + ' · Pasadía 9am – 5pm';
  else if (tipo === 'early')     fechas = meta.checkinFmt + ' → ' + meta.checkoutFmt + ' · 1 noche (entrada anticipada 9am)';
  else if (tipo === 'late')      fechas = meta.checkinFmt + ' → ' + meta.checkoutFmt + ' · 1 noche (salida tardía 4pm)';
  else                           fechas = meta.checkinFmt + ' → ' + meta.checkoutFmt + ' · ' + meta.estanciaValue + (meta.estanciaValue === 1 ? ' noche' : ' noches');

  const persons      = parseInt(reservation.persons, 10) || 1;
  const personasStr  = persons + (persons === 1 ? ' persona' : ' personas');
  const total        = (parseFloat(reservation.amount) || 0).toFixed(2);
  const firstName    = ((reservation.name || '').toString().trim().split(/\s+/)[0]) || 'amigo';

  return sendWhatsAppTemplate(reservation.telefono, 'confirmacion_reserva', 'es_PA', {
    nombre: firstName,
    cabana: cabin,
    fechas: fechas,
    personas: personasStr,
    total: total
  });
}

/**
 * Test desde editor: arma una reserva mock y envia la plantilla a
 * tu propio numero. Corre esto despues de que Meta apruebe la plantilla.
 */
function testEnviarPlantillaConfirmacion() {
  const reservaMock = {
    name:     'Carlos Test',
    telefono: '50769812266',
    cabin:    'azul',
    checkin:  '2026-08-15',
    checkout: '2026-08-16',
    persons:  2,
    amount:   180,
    tipo:     'noche'
  };
  const resultado = sendWAReservaConfirmada(reservaMock);
  Logger.log('Respuesta plantilla: ' + JSON.stringify(resultado));
  return resultado;
}

/**
 * Envia un mensaje interactivo con botones de respuesta rapida.
 * Max 3 botones, titulo max 20 chars.
 *
 * @param {string} toPhone
 * @param {string} bodyText    texto principal del mensaje
 * @param {Array}  buttons     [{ id: 'btn_x', title: 'Texto corto' }, ...]
 * @param {string} [headerText] opcional, hasta 60 chars
 * @param {string} [footerText] opcional, hasta 60 chars
 */
function sendWhatsAppButtons(toPhone, bodyText, buttons, headerText, footerText) {
  const cfg = _waProps();
  if (!cfg.token || !cfg.phoneId) throw new Error('WA: faltan credenciales');
  const to = _waNormalizePhone(toPhone);
  if (!to) throw new Error('WA: telefono invalido: ' + toPhone);
  if (!buttons || !buttons.length) throw new Error('WA: sin botones');

  const interactive = {
    type: 'button',
    body: { text: bodyText.slice(0, 1024) },
    action: {
      buttons: buttons.slice(0, 3).map(b => ({
        type: 'reply',
        reply: { id: String(b.id).slice(0, 256), title: String(b.title).slice(0, 20) }
      }))
    }
  };
  if (headerText) interactive.header = { type: 'text', text: headerText.slice(0, 60) };
  if (footerText) interactive.footer = { text: footerText.slice(0, 60) };

  const url = 'https://graph.facebook.com/v21.0/' + cfg.phoneId + '/messages';
  const payload = {
    messaging_product: 'whatsapp',
    to: to,
    type: 'interactive',
    interactive: interactive
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: { Authorization: 'Bearer ' + cfg.token },
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const text = res.getContentText();
  let json;
  try { json = JSON.parse(text); } catch(_) { json = { raw: text }; }
  logDebugEntry('WA-send-buttons', { to: to, code: code, ok: code >= 200 && code < 300, buttons: buttons.map(b => b.id), error: json.error || null });
  if (code < 200 || code >= 300) throw new Error('WA buttons send failed HTTP ' + code + ': ' + text.slice(0, 400));
  return json;
}

