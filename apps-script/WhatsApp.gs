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
