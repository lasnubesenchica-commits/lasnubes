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
  const CABIN_SHORT = { verde: 'Paseo', azul: 'Portal', lila: 'Puente' };

  // Multi-cabaña: override cabana + fechas para reflejar el itinerario.
  const isMulti = reservation.multiCabin && Array.isArray(reservation.multiCabin) && reservation.multiCabin.length >= 2;
  let cabin, fechas;
  if (isMulti) {
    const arr = reservation.multiCabin;
    const cabinTags = arr.map(function(s) { return CABIN_SHORT[s.cabin] || s.cabin; });
    cabin = cabinTags.join(' → ') + ' (multi-cabaña)';
    const totalNights = arr.reduce(function(sum, s) { return sum + (s.nights || 0); }, 0);
    const firstIn = formatDateES(arr[0].checkin);
    const lastOut = formatDateES(arr[arr.length - 1].checkout);
    fechas = firstIn + ' → ' + lastOut + ' · ' + totalNights + ' ' + (totalNights === 1 ? 'noche' : 'noches') + ' en ' + arr.length + ' cabañas';
  } else {
    cabin = CABIN_NAMES[reservation.cabin] || reservation.cabin || 'Las Nubes';
    const tipo = meta.tipo;
    if (tipo === 'pasatarde')      fechas = meta.checkinFmt + ' · Pasatarde 12:30pm – 7pm';
    else if (tipo === 'pasadia')   fechas = meta.checkinFmt + ' · Pasadía 9am – 5pm';
    else if (tipo === 'early')     fechas = meta.checkinFmt + ' → ' + meta.checkoutFmt + ' · 1 noche (entrada anticipada 9am)';
    else if (tipo === 'late')      fechas = meta.checkinFmt + ' → ' + meta.checkoutFmt + ' · 1 noche (salida tardía 4pm)';
    else                           fechas = meta.checkinFmt + ' → ' + meta.checkoutFmt + ' · ' + meta.estanciaValue + (meta.estanciaValue === 1 ? ' noche' : ' noches');
  }

  const persons      = parseInt(reservation.persons, 10) || 1;
  const personasStr  = persons + (persons === 1 ? ' persona' : ' personas');
  const nombre       = (reservation.name || 'amigo').toString().trim();

  // Link público de la reserva. Usamos el HMAC largo (?id=&t=) que está
  // probado en producción; el short link (?c=) crashea en la web app
  // desplegada al leer la hoja ShareLinks, así que no lo usamos.
  // Safe: con el id placeholder de la vista previa el link firmado apuntaría a
  // una fila inexistente → cae al sitio en vez de a "Link no válido".
  const link = getPublicReservaUrlSafe(reservation.id) || 'https://lasnubes.cloud';

  return sendWhatsAppTemplate(reservation.telefono, 'confirmacion_reserva', 'es_PA', {
    nombre:        nombre,
    cabana:        cabin,
    fechas:        fechas,
    personas:      personasStr,
    checkin_hora:  _horaPlantilla(reservation.tipo, 'checkin',  false, reservation.horaEntrada),
    checkout_hora: _horaPlantilla(reservation.tipo, 'checkout', reservation.checkoutExtendido, null, reservation.horaSalida),
    link:          link
  }, null, 'consulta_' + reservation.id);   // payload del boton "Consultas y cambios"
}

/**
 * Envia el aviso de CERTIFICADO DE REGALO al beneficiario.
 *
 * Plantilla 'certificado_regalo' (es_PA), params nombrados:
 *   {{nombre}}   beneficiario (primer nombre)
 *   {{de}}       quien lo regala
 *   {{cabana}}   nombre de la cabaña
 *   {{detalle}}  "12 ago 2026 → 13 ago 2026 · 1 noche" o "Fechas a coordinar"
 *   {{link}}     link publico de la reserva (o lasnubes.cloud)
 *
 * Para darla de alta NO hay que escribirla a mano: correr
 * crearPlantillaRegaloEnMeta() desde el editor. Manda la definicion por API con
 * los nombres de parametro y el boton exactos que espera este codigo, que es
 * justo lo que se rompe al crearla desde la UI. El texto que registra:
 *
 *   🎁 Hola {{nombre}}! Tienes un regalo en Las Nubes de parte de {{de}}.
 *
 *   Cabaña: {{cabana}}
 *   Fechas: {{detalle}}
 *
 *   Todo está cubierto — no tienes nada que pagar.
 *   Aquí puedes ver los detalles: {{link}}
 *
 *   Si necesitas coordinar o ajustar tus fechas, escríbenos al WhatsApp
 *   6981-2266. Estaremos atentos 🙏
 *
 * Ojo con tres cosas, que son las que hacen fallar el envio:
 *   · los parametros son NOMBRADOS, no {{1}}/{{2}};
 *   · `cabana` va sin ñ, igual que la clave que manda el codigo;
 *   · la plantilla NECESITA un boton quick-reply, porque este codigo siempre
 *     manda el payload 'consulta_<id>' (lo atiende _botHandleConsultaReserva).
 *     Sin boton, Meta rechaza el envio por componentes que no coinciden.
 *   · el body no arranca ni termina en variable, que Meta a veces rechaza.
 *
 * APROBADA en Meta (jul 2026): esta es la plantilla que se manda. El fallback a
 * 'confirmacion_reserva' queda como RED DE SEGURIDAD, no como estado esperado —
 * si Meta la pausa o el envio falla, el beneficiario igual recibe su aviso (con
 * el marco de regalo metido en los campos de texto libre, ver
 * REGALO_WA_COORDINAR_CORTO abajo) en vez de quedarse sin nada. Ahora que está
 * aprobada, ver "⚠ certificado_regalo falló" en el log significa que algo se
 * rompió: revisar el motivo con verEstadoPlantillasWA(), no ignorarlo.
 * NUNCA incluye tarifa, abono ni saldo, por ninguna de las dos vías.
 */
// Versión compacta de la invitación para meterla en un param de plantilla.
// Meta rechaza parámetros con saltos de línea o tabs, así que va en una línea.
const REGALO_WA_COORDINAR_CORTO = 'escríbenos al WhatsApp 6981-2266 para verificar disponibilidad y coordinar tus fechas. Estaremos atentos 🙏';

function sendWARegaloCertificado(reservation) {
  if (!reservation) throw new Error('WA: reservation requerida');
  if (!reservation.telefono) throw new Error('WA: reservation sin telefono');

  const g = _parseRegalo(reservation);
  const CABIN_NAMES = {
    verde: 'Paseo por Las Nubes',
    azul:  'Portal hacia Las Nubes',
    lila:  'Puente entre Las Nubes'
  };
  const cabin = CABIN_NAMES[reservation.cabin] || reservation.cabin || 'Las Nubes';

  const fullName = (reservation.name || 'amigo').toString().trim();
  const nombre   = fullName.split(/\s+/)[0] || fullName;
  const de       = g.de || 'alguien que te quiere';

  // Sin fecha (origen Abierta o fechas vacías) → el beneficiario coordina.
  const sinFecha = (reservation.origin === 'Abierta') || !reservation.checkin || !reservation.checkout;
  let detalle;
  if (sinFecha) {
    detalle = 'Fechas a coordinar';
  } else {
    const meta = tipoEmailMeta(reservation);
    if (meta.tipo === 'pasatarde')    detalle = meta.checkinFmt + ' · Pasatarde 12:30pm – 7pm';
    else if (meta.tipo === 'pasadia') detalle = meta.checkinFmt + ' · Pasadía 9am – 5pm';
    else if (meta.isPasadia)          detalle = meta.checkinFmt;
    else detalle = meta.checkinFmt + ' → ' + meta.checkoutFmt +
                   (typeof meta.estanciaValue === 'number'
                     ? ' · ' + meta.estanciaValue + (meta.estanciaValue === 1 ? ' noche' : ' noches')
                     : ' · ' + meta.estanciaValue);
  }

  // Safe: con el id placeholder de la vista previa el link firmado apuntaría a
  // una fila inexistente → cae al sitio en vez de a "Link no válido".
  const link = getPublicReservaUrlSafe(reservation.id) || 'https://lasnubes.cloud';

  try {
    return sendWhatsAppTemplate(reservation.telefono, 'certificado_regalo', 'es_PA', {
      nombre:  nombre,
      de:      de,
      cabana:  cabin,
      detalle: detalle,
      link:    link
    }, null, 'consulta_' + reservation.id);
  } catch(err) {
    // Red de seguridad: 'certificado_regalo' ya está aprobada, así que llegar
    // acá es una anomalía (plantilla pausada, rate limit, red caída), no el
    // camino normal. Reusamos 'confirmacion_reserva' metiendo el marco de
    // regalo en los campos de texto libre para no dejar al beneficiario sin
    // aviso. Sigue sin montos.
    Logger.log('⚠ certificado_regalo falló (' + err.message + ') → fallback confirmacion_reserva');
    const persons     = parseInt(reservation.persons, 10) || 1;
    const personasStr = persons + (persons === 1 ? ' persona' : ' personas');
    return sendWhatsAppTemplate(reservation.telefono, 'confirmacion_reserva', 'es_PA', {
      nombre:        nombre,
      cabana:        '🎁 ' + cabin + ' — regalo de ' + de,
      // El body de 'confirmacion_reserva' es fijo y no trae la invitación a
      // coordinar, así que va aquí (único param de texto libre suficiente).
      fechas:        detalle + ' · ' + REGALO_WA_COORDINAR_CORTO,
      personas:      personasStr,
      checkin_hora:  sinFecha ? 'a coordinar' : _horaPlantilla(reservation.tipo, 'checkin',  false, reservation.horaEntrada),
      checkout_hora: sinFecha ? 'a coordinar' : _horaPlantilla(reservation.tipo, 'checkout', reservation.checkoutExtendido, null, reservation.horaSalida),
      link:          link
    }, null, 'consulta_' + reservation.id);
  }
}

/**
 * Router: manda el aviso correcto segun si la reserva es un regalo o no.
 * Lo usan el dashboard (action sendWAConfirmacion) y la vista previa.
 */
function sendWAAvisoReserva(reservation) {
  return esReservaRegalo(reservation)
    ? sendWARegaloCertificado(reservation)
    : sendWAReservaConfirmada(reservation);
}

/**
 * Test desde editor: manda el certificado de regalo a tu propio numero.
 */
function testEnviarPlantillaRegalo() {
  const reservaMock = {
    id:       'TEST-REGALO-' + Date.now(),
    name:     'Marta Beneficiaria',
    telefono: '50769812266',
    pagador:  'Juan Pagador',
    regalo:   JSON.stringify({ de: 'Juan Pagador', mensaje: 'Feliz cumpleaños, disfrútalo!' }),
    cabin:    'lila',
    checkin:  '2026-08-15',
    checkout: '2026-08-16',
    persons:  2,
    amount:   180,
    tipo:     'noche',
    origin:   'Directa'
  };
  const resultado = sendWARegaloCertificado(reservaMock);
  Logger.log('Respuesta plantilla regalo: ' + JSON.stringify(resultado));
  return resultado;
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
  // `headerText` acepta un string (encabezado de texto, máx 60) o un objeto
  // { imageUrl } para poner una IMAGEN arriba de los botones. Lo segundo evita
  // mandar dos mensajes cuando la foto y las acciones van juntas — el caso de
  // las instrucciones de llegada, donde la foto sirve para confirmar que se
  // llegó a la cabaña correcta.
  if (headerText && typeof headerText === 'object' && headerText.imageUrl) {
    interactive.header = { type: 'image', image: { link: headerText.imageUrl } };
  } else if (headerText) {
    interactive.header = { type: 'text', text: String(headerText).slice(0, 60) };
  }
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


// ═══════════════════════════════════════════════════════════
//  Dar de alta la plantilla 'certificado_regalo' en Meta
// ═══════════════════════════════════════════════════════════
//
// Crearla desde la UI de WhatsApp Manager es donde se rompe todo: hay que
// acordarse de elegir parámetros NOMBRADOS en vez de {{1}}/{{2}}, escribir
// `cabana` sin ñ, y agregar el botón quick-reply. Cualquiera de las tres cosas
// mal y el envío falla con un error de componentes que no dice mucho.
//
// Esta función manda la definición por API con esos tres detalles ya correctos,
// derivados del mismo código que después la envía.
//
// Dry-run por defecto: imprime el JSON y no lo manda, para poder revisarlo (o
// pegarlo a mano si se prefiere). Para crearla de verdad:
//   crearPlantillaRegaloEnMetaAHORA()
function _plantillaRegaloPayload() {
  const body =
    '🎁 Hola {{nombre}}! Tienes un regalo en Las Nubes de parte de {{de}}.\n' +
    '\n' +
    'Cabaña: {{cabana}}\n' +
    'Fechas: {{detalle}}\n' +
    '\n' +
    'Todo está cubierto — no tienes nada que pagar.\n' +
    'Aquí puedes ver los detalles: {{link}}\n' +
    '\n' +
    'Si necesitas coordinar o ajustar tus fechas, escríbenos al WhatsApp ' +
    '6981-2266. Estaremos atentos 🙏';

  return {
    name: 'certificado_regalo',
    language: 'es_PA',
    category: 'UTILITY',
    parameter_format: 'NAMED',
    components: [
      {
        type: 'BODY',
        text: body,
        example: {
          body_text_named_params: [
            { param_name: 'nombre',  example: 'María' },
            { param_name: 'de',      example: 'Josué Rodríguez' },
            { param_name: 'cabana',  example: 'Paseo por Las Nubes' },
            { param_name: 'detalle', example: '12 ago 2026 → 13 ago 2026 · 1 noche' },
            { param_name: 'link',    example: 'https://lasnubes.cloud/reserva.html?id=LN-1234&t=abc123' }
          ]
        }
      },
      {
        type: 'BUTTONS',
        buttons: [{ type: 'QUICK_REPLY', text: 'Consultas y cambios' }]
      }
    ]
  };
}

function crearPlantillaRegaloEnMeta(dryRun) {
  if (dryRun !== false) dryRun = true;
  const cfg = _waProps();
  if (!cfg.token)      { Logger.log('⚠ Falta WA_ACCESS_TOKEN en Script Properties.'); return; }
  if (!cfg.businessId) { Logger.log('⚠ Falta WA_BUSINESS_ACCOUNT_ID en Script Properties.'); return; }

  const payload = _plantillaRegaloPayload();
  Logger.log('═══ ' + (dryRun ? 'DRY-RUN · ' : '') + "PLANTILLA 'certificado_regalo' ═══");
  Logger.log('');
  Logger.log('Así queda el mensaje que va a recibir el beneficiario:');
  Logger.log('');
  payload.components[0].text.split('\n').forEach(l => Logger.log('   ' + (l || ' ')));
  Logger.log('');
  Logger.log('   [ Consultas y cambios ]   ← botón quick-reply');
  Logger.log('');
  Logger.log('Definición que se manda a Meta (WABA ' + cfg.businessId + '):');
  Logger.log(JSON.stringify(payload, null, 2));

  if (dryRun) {
    Logger.log('');
    Logger.log('Nada se envió. Para crearla: crearPlantillaRegaloEnMetaAHORA()');
    return;
  }

  const res = UrlFetchApp.fetch(
    'https://graph.facebook.com/v21.0/' + cfg.businessId + '/message_templates',
    { method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: { Authorization: 'Bearer ' + cfg.token },
      muteHttpExceptions: true });
  const code = res.getResponseCode();
  const txt  = res.getContentText();
  Logger.log('');
  Logger.log('HTTP ' + code);
  Logger.log(txt);
  if (code >= 200 && code < 300) {
    Logger.log('');
    Logger.log('✓ Plantilla enviada a revisión. Meta suele aprobar las UTILITY en minutos.');
    Logger.log('  Cuando quede aprobada, probala con testEnviarPlantillaRegalo().');
    Logger.log('  Mientras esté en revisión el envío sigue cayendo al fallback, que funciona.');
  } else {
    Logger.log('');
    Logger.log('✗ Meta la rechazó. Los motivos habituales:');
    Logger.log('   · ya existe una plantilla con ese nombre en es_PA → borrarla o renombrarla;');
    Logger.log('   · el token no tiene permiso whatsapp_business_management;');
    Logger.log('   · Meta reclasificó la categoría (si la pasa a MARKETING igual sirve).');
  }
}

function crearPlantillaRegaloEnMetaAHORA() { return crearPlantillaRegaloEnMeta(false); }

// ─── listos_para_recibirte (aviso de llegada, 11am) ─────────────────
// Params POSICIONALES para que coincida con enviarAvisoLlegadaHoy, que manda
// un array: [nombre, cabaña, hora de check-in].
function _plantillaLlegadaPayload() {
  const body =
    '¡Hola {{1}}! 🌿\n' +
    '\n' +
    'Hoy te recibimos en {{2}}. El check-in es a partir de las {{3}} — llega a la ' +
    'hora que te quede cómoda, estamos listos para recibirte.\n' +
    '\n' +
    'Antes de subir:\n' +
    '• Trae hielo y tus alimentos (hay cooler grande, no nevera)\n' +
    '• Carga tus equipos en el camino: la energía de la cabaña es solar\n' +
    '• Si eres sensible a los mosquitos, trae repelente\n' +
    '\n' +
    'Cómo llegar: pon "Aires de Chicá" en Waze y te lleva directo al portón.\n' +
    '🗺 https://maps.google.com/?q=8.639400,-79.945900\n' +
    '🚦 https://waze.com/ul?ll=8.639400,-79.945900&navigate=yes\n' +
    '\n' +
    'Cuando estés frente al portón verde, toca el botón de abajo y le avisamos al ' +
    'equipo para que te abran.';
  return {
    name: 'listos_para_recibirte',
    language: 'es_ES',
    category: 'UTILITY',
    components: [
      {
        type: 'BODY',
        text: body,
        example: { body_text: [['Ana', 'Portal hacia Las Nubes', '2:00 pm']] }
      },
      {
        type: 'BUTTONS',
        buttons: [{ type: 'QUICK_REPLY', text: '🚪 He llegado' }]
      }
    ]
  };
}

// Dry-run por defecto: imprime el mensaje tal como lo verá el huésped y el JSON
// que se le manda a Meta, sin crear nada. Para crearla de verdad:
// crearPlantillaLlegadaEnMetaAHORA()
function crearPlantillaLlegadaEnMeta(dryRun) {
  if (dryRun !== false) dryRun = true;
  const cfg = _waProps();
  if (!cfg.token)      { Logger.log('⚠ Falta WA_ACCESS_TOKEN en Script Properties.'); return; }
  if (!cfg.businessId) { Logger.log('⚠ Falta WA_BUSINESS_ACCOUNT_ID en Script Properties.'); return; }

  const payload = _plantillaLlegadaPayload();
  Logger.log('═══ ' + (dryRun ? 'DRY-RUN · ' : '') + "PLANTILLA 'listos_para_recibirte' ═══");
  Logger.log('');
  Logger.log('Así le llega al huésped el día de su check-in, a las 11am:');
  Logger.log('');
  payload.components[0].text
    .replace('{{1}}', 'Ana').replace('{{2}}', 'Portal hacia Las Nubes').replace('{{3}}', '2:00 pm')
    .split('\n').forEach(l => Logger.log('   ' + (l || ' ')));
  Logger.log('');
  Logger.log('   [ 🚪 He llegado ]   ← botón quick-reply');
  Logger.log('');
  Logger.log('Definición que se manda a Meta (WABA ' + cfg.businessId + '):');
  Logger.log(JSON.stringify(payload, null, 2));

  if (dryRun) {
    Logger.log('');
    Logger.log('Nada se envió. Para crearla: crearPlantillaLlegadaEnMetaAHORA()');
    return;
  }

  const res = UrlFetchApp.fetch(
    'https://graph.facebook.com/v21.0/' + cfg.businessId + '/message_templates',
    { method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload),
      headers: { Authorization: 'Bearer ' + cfg.token },
      muteHttpExceptions: true });
  const code = res.getResponseCode();
  const txt  = res.getContentText();
  Logger.log('');
  Logger.log('HTTP ' + code);
  Logger.log(txt);
  if (code >= 200 && code < 300) {
    Logger.log('');
    Logger.log('✓ Plantilla enviada a revisión. Meta suele aprobar las UTILITY en minutos.');
    Logger.log('  Seguí el estado con verEstadoPlantillasWA().');
    Logger.log('  Cuando quede APPROVED:');
    Logger.log('   1) probala con _testAvisoLlegadaAMiNumero()');
    Logger.log('   2) activá el envío diario con instalarTriggerAvisoLlegada()');
  } else {
    Logger.log('');
    Logger.log('✗ Meta la rechazó. Los motivos habituales:');
    Logger.log('   · ya existe una plantilla con ese nombre en es_ES → borrarla o renombrarla;');
    Logger.log('   · el token no tiene permiso whatsapp_business_management;');
    Logger.log('   · Meta reclasificó la categoría (si la pasa a MARKETING igual sirve).');
  }
}

function crearPlantillaLlegadaEnMetaAHORA() { return crearPlantillaLlegadaEnMeta(false); }


// Estado de las plantillas en Meta, sin entrar a WhatsApp Manager.
// Útil sobre todo después de crear una: queda en PENDING y hay que saber
// cuándo pasó a APPROVED para dejar de depender del fallback.
function verEstadoPlantillasWA() {
  const cfg = _waProps();
  if (!cfg.token || !cfg.businessId) {
    Logger.log('⚠ Faltan WA_ACCESS_TOKEN o WA_BUSINESS_ACCOUNT_ID en Script Properties.');
    return;
  }
  const url = 'https://graph.facebook.com/v21.0/' + cfg.businessId +
              '/message_templates?fields=name,language,status,category,rejected_reason&limit=100';
  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + cfg.token }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    Logger.log('HTTP ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 400));
    return;
  }
  const data = JSON.parse(res.getContentText()).data || [];
  const ICONO = { APPROVED: '✅', PENDING: '⏳', REJECTED: '❌', PAUSED: '⏸', DISABLED: '🚫' };
  Logger.log('═══ PLANTILLAS DE WHATSAPP (' + data.length + ') ═══');
  Logger.log('');
  data.sort((a, b) => String(a.name).localeCompare(String(b.name)));
  data.forEach(t => {
    Logger.log((ICONO[t.status] || '·') + ' ' + t.name + '  (' + t.language + ' · ' + t.category + ')'
      + '  → ' + t.status + (t.rejected_reason && t.rejected_reason !== 'NONE'
        ? '  motivo: ' + t.rejected_reason : ''));
  });
  const regalo = data.filter(t => t.name === 'certificado_regalo')[0];
  Logger.log('');
  if (!regalo) {
    Logger.log('· certificado_regalo no aparece todavía. Los envíos de regalo siguen');
    Logger.log('  cayendo al fallback confirmacion_reserva, que funciona.');
  } else if (regalo.status === 'APPROVED') {
    Logger.log('✅ certificado_regalo está aprobada. Probala con testEnviarPlantillaRegalo().');
  } else if (regalo.status === 'REJECTED') {
    Logger.log('❌ certificado_regalo fue rechazada (' + (regalo.rejected_reason || 's/motivo') + ').');
    Logger.log('   Los envíos siguen usando el fallback, así que nada se rompe mientras se resuelve.');
  } else {
    Logger.log('⏳ certificado_regalo sigue en ' + regalo.status + '. Mientras tanto los envíos');
    Logger.log('   caen al fallback confirmacion_reserva, que funciona.');
  }
}
