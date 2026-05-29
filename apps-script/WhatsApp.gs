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
  try { logMensaje(to, 'out', 'text', body, (json.messages && json.messages[0] && json.messages[0].id) || null); } catch(_) {}
  return json;
}

// Envia una imagen por URL (la descarga Meta al enviar). caption opcional.
function sendWhatsAppImage(toPhone, imageUrl, caption) {
  const cfg = _waProps();
  if (!cfg.token || !cfg.phoneId) throw new Error('WA: faltan credenciales en Script Properties');
  const to = _waNormalizePhone(toPhone);
  if (!to) throw new Error('WA: telefono invalido: ' + toPhone);

  const url = 'https://graph.facebook.com/v21.0/' + cfg.phoneId + '/messages';
  const image = { link: imageUrl };
  if (caption) image.caption = caption;
  const payload = { messaging_product: 'whatsapp', to: to, type: 'image', image: image };
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
  logDebugEntry('WA-send-image', { to: to, code: code, ok: code >= 200 && code < 300, url: String(imageUrl).slice(0, 120), error: json.error || null });
  if (code < 200 || code >= 300) {
    throw new Error('WA image send failed HTTP ' + code + ': ' + text.slice(0, 300));
  }
  try { logMensaje(to, 'out', 'image', (caption || '[imagen]') + ' · ' + String(imageUrl).slice(0, 100), (json.messages && json.messages[0] && json.messages[0].id) || null); } catch(_) {}
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
 * Test: envia una foto de cabaña (link de Drive) a tu numero para verificar
 * que WhatsApp acepta esas URLs. Requiere que el numero tenga ventana de 24h
 * abierta (escribile algo al Agente antes de correr esto).
 */
function testEnviarFotoCabana() {
  const numeroDestino = '50769812266';
  const url = 'https://lh3.googleusercontent.com/d/1kolAp8PKDO3ws6abcUUfD2hpN_3ZLBjB=w1280';
  const resultado = sendWhatsAppImage(numeroDestino, url, '🏡 Portal hacia Las Nubes (prueba)');
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
function sendWhatsAppTemplate(toPhone, templateName, languageCode, namedParams, buttonParams, quickReplyPayload) {
  const cfg = _waProps();
  if (!cfg.token || !cfg.phoneId) throw new Error('WA: faltan credenciales en Script Properties');
  const to = _waNormalizePhone(toPhone);
  if (!to) throw new Error('WA: telefono invalido: ' + toPhone);

  const components = [];
  // namedParams puede ser:
  //  - Array  → parámetros POSICIONALES ({{1}}, {{2}}, ...)
  //  - Object → parámetros NOMBRADOS ({{nombre}}, {{cabana}}, ...)
  if (Array.isArray(namedParams) && namedParams.length > 0) {
    components.push({
      type: 'body',
      parameters: namedParams.map(v => ({ type: 'text', text: String(v) }))
    });
  } else if (namedParams && !Array.isArray(namedParams) && Object.keys(namedParams).length > 0) {
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
  // Boton quick-reply con payload dinamico (ej. checkout_<reservaId>). El
  // payload vuelve en el webhook cuando el cliente toca el boton.
  if (quickReplyPayload) {
    components.push({
      type: 'button',
      sub_type: 'quick_reply',
      index: '0',
      parameters: [{ type: 'payload', payload: String(quickReplyPayload) }]
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
  try { logMensaje(to, 'out', 'template', templateName + ' · ' + JSON.stringify(namedParams || {}).slice(0, 300), (json.messages && json.messages[0] && json.messages[0].id) || null); } catch(_) {}
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
  const nombre       = (reservation.name || 'amigo').toString().trim();

  // Link público de la reserva. Usamos el HMAC largo (?id=&t=) que está
  // probado en producción; el short link (?c=) crashea en la web app
  // desplegada al leer la hoja ShareLinks, así que no lo usamos.
  let link = '';
  try { link = getPublicReservaUrl(reservation.id); } catch(e) { link = 'https://lasnubes.cloud'; }

  return sendWhatsAppTemplate(reservation.telefono, 'confirmacion_reserva', 'es_PA', {
    nombre:        nombre,
    cabana:        cabin,
    fechas:        fechas,
    personas:      personasStr,
    checkin_hora:  _horaPlantilla(reservation.tipo, 'checkin'),
    checkout_hora: _horaPlantilla(reservation.tipo, 'checkout', reservation.checkoutExtendido),
    link:          link
  }, null, 'consulta_' + reservation.id);   // payload del boton "Consultas y cambios"
}

/**
 * Test desde editor: arma una reserva mock y envia la plantilla a
 * tu propio numero. Corre esto despues de que Meta apruebe la plantilla.
 */
function testEnviarPlantillaConfirmacion() {
  const reservaMock = {
    id:       'TEST-' + Date.now(),
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
 * Descarga una imagen recibida en un webhook de WhatsApp.
 * Meta entrega un id, hay que hacer 2 calls: 1) GET /{id} → URL temporal,
 * 2) GET URL con el bearer token → binario.
 *
 * @param {string} imageId  id del media object (de msg.image.id)
 * @return {?Object}        { base64, mimeType } o null si falla
 */
function fetchWhatsAppImage(imageId) {
  const cfg = _waProps();
  if (!cfg.token) return null;
  const metaRes = UrlFetchApp.fetch('https://graph.facebook.com/v21.0/' + imageId, {
    headers: { Authorization: 'Bearer ' + cfg.token },
    muteHttpExceptions: true
  });
  let meta;
  try { meta = JSON.parse(metaRes.getContentText()); } catch(_) { return null; }
  if (!meta.url) {
    logDebugEntry('WA-image-no-url', { imageId: imageId, raw: metaRes.getContentText().slice(0, 200) });
    return null;
  }
  const imgRes = UrlFetchApp.fetch(meta.url, {
    headers: { Authorization: 'Bearer ' + cfg.token },
    muteHttpExceptions: true
  });
  if (imgRes.getResponseCode() !== 200) {
    logDebugEntry('WA-image-fetch-FAIL', { imageId: imageId, code: imgRes.getResponseCode() });
    return null;
  }
  const blob = imgRes.getBlob();
  return {
    base64:   Utilities.base64Encode(blob.getBytes()),
    mimeType: meta.mime_type || blob.getContentType() || 'image/jpeg'
  };
}

/**
 * Envia un mensaje interactivo de boton CTA URL: muestra un boton que
 * al tocarlo abre la URL en el navegador (o WhatsApp si es wa.me).
 *
 * @param {string} toPhone
 * @param {string} bodyText
 * @param {string} displayText  texto del boton (max 20)
 * @param {string} url          URL absoluta (https://...)
 * @param {string} [headerText] opcional
 * @param {string} [footerText] opcional
 */
function sendWhatsAppCTAUrl(toPhone, bodyText, displayText, url, headerText, footerText) {
  const cfg = _waProps();
  if (!cfg.token || !cfg.phoneId) throw new Error('WA: faltan credenciales');
  const to = _waNormalizePhone(toPhone);
  if (!to) throw new Error('WA: telefono invalido: ' + toPhone);

  const interactive = {
    type: 'cta_url',
    body: { text: bodyText.slice(0, 1024) },
    action: {
      name: 'cta_url',
      parameters: {
        display_text: String(displayText).slice(0, 20),
        url: url
      }
    }
  };
  if (headerText) interactive.header = { type: 'text', text: headerText.slice(0, 60) };
  if (footerText) interactive.footer = { text: footerText.slice(0, 60) };

  const apiUrl = 'https://graph.facebook.com/v21.0/' + cfg.phoneId + '/messages';
  const res = UrlFetchApp.fetch(apiUrl, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      messaging_product: 'whatsapp',
      to: to,
      type: 'interactive',
      interactive: interactive
    }),
    headers: { Authorization: 'Bearer ' + cfg.token },
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const text = res.getContentText();
  let json;
  try { json = JSON.parse(text); } catch(_) { json = { raw: text }; }
  logDebugEntry('WA-send-cta', { to: to, code: code, ok: code >= 200 && code < 300, url: url, error: json.error || null });
  if (code < 200 || code >= 300) throw new Error('WA cta send failed HTTP ' + code + ': ' + text.slice(0, 400));
  try { logMensaje(to, 'out', 'interactive_cta_url', bodyText + ' [btn: ' + displayText + ' → ' + url + ']', (json.messages && json.messages[0] && json.messages[0].id) || null); } catch(_) {}
  return json;
}

/**
 * Envia un mensaje interactivo de lista (hasta 10 opciones agrupadas en
 * secciones). Mucho mejor UX que botones para menus con varias opciones.
 *
 * @param {string} toPhone
 * @param {string} bodyText
 * @param {Array}  sections    [{ title: 'Seccion', rows: [{ id, title, description? }] }]
 * @param {string} buttonText  texto del boton que abre la lista (max 20)
 * @param {string} [headerText] opcional, max 60
 * @param {string} [footerText] opcional, max 60
 */
function sendWhatsAppList(toPhone, bodyText, sections, buttonText, headerText, footerText) {
  const cfg = _waProps();
  if (!cfg.token || !cfg.phoneId) throw new Error('WA: faltan credenciales');
  const to = _waNormalizePhone(toPhone);
  if (!to) throw new Error('WA: telefono invalido: ' + toPhone);

  const interactive = {
    type: 'list',
    body: { text: bodyText.slice(0, 1024) },
    action: {
      button: (buttonText || 'Ver opciones').slice(0, 20),
      sections: (sections || []).map(s => ({
        title: (s.title || '').toString().slice(0, 24),
        rows: (s.rows || []).slice(0, 10).map(r => {
          const row = {
            id: String(r.id).slice(0, 200),
            title: String(r.title).slice(0, 24)
          };
          if (r.description) row.description = String(r.description).slice(0, 72);
          return row;
        })
      }))
    }
  };
  if (headerText) interactive.header = { type: 'text', text: headerText.slice(0, 60) };
  if (footerText) interactive.footer = { text: footerText.slice(0, 60) };

  const url = 'https://graph.facebook.com/v21.0/' + cfg.phoneId + '/messages';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      messaging_product: 'whatsapp',
      to: to,
      type: 'interactive',
      interactive: interactive
    }),
    headers: { Authorization: 'Bearer ' + cfg.token },
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const text = res.getContentText();
  let json;
  try { json = JSON.parse(text); } catch(_) { json = { raw: text }; }
  logDebugEntry('WA-send-list', { to: to, code: code, ok: code >= 200 && code < 300, error: json.error || null });
  if (code < 200 || code >= 300) throw new Error('WA list send failed HTTP ' + code + ': ' + text.slice(0, 400));
  try {
    const summary = bodyText + ' [opts: ' + (sections || []).map(s => (s.rows || []).map(r => r.title).join('/')).join(' | ') + ']';
    logMensaje(to, 'out', 'interactive_list', summary, (json.messages && json.messages[0] && json.messages[0].id) || null);
  } catch(_) {}
  return json;
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
  try {
    const summary = bodyText + ' [btns: ' + buttons.map(b => b.title).join(' / ') + ']';
    logMensaje(to, 'out', 'interactive_buttons', summary, (json.messages && json.messages[0] && json.messages[0].id) || null);
  } catch(_) {}
  return json;
}

