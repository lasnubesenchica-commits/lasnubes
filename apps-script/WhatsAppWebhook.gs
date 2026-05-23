/**
 * WhatsApp Cloud API — Webhook handler
 *
 * Recibe inbounds de clientes que escriben al numero de Las Nubes Business.
 *
 * Setup en Meta side (https://developers.facebook.com/apps/<APP_ID>/whatsapp-business/wa-settings/):
 *   1. Callback URL: la URL de Web App de este proyecto (misma SHEETS_API_URL).
 *   2. Verify token: el valor de WA_WEBHOOK_VERIFY_TOKEN en Script Properties.
 *   3. Subscribe a field: 'messages'.
 *
 * Flow:
 *   - GET /exec?hub.mode=subscribe&hub.verify_token=XXX&hub.challenge=YYY
 *     → handleWebhookVerify retorna YYY si el token coincide.
 *   - POST /exec con payload de Meta (object='whatsapp_business_account')
 *     → handleWhatsAppWebhook lo procesa.
 */

function handleWebhookVerify(e) {
  const params = (e && e.parameter) || {};
  const mode      = params['hub.mode'];
  const token     = params['hub.verify_token'];
  const challenge = params['hub.challenge'];

  const expectedToken = PropertiesService.getScriptProperties().getProperty('WA_WEBHOOK_VERIFY_TOKEN');
  if (mode === 'subscribe' && token && token === expectedToken) {
    logDebugEntry('WA-webhook-verify-OK', { mode: mode });
    // Meta espera el challenge devuelto como text/plain
    return ContentService.createTextOutput(challenge || '');
  }
  logDebugEntry('WA-webhook-verify-FAIL', { mode: mode, tokenLen: token ? token.length : 0 });
  return ContentService.createTextOutput('forbidden').setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Procesa el payload de webhook de Meta WhatsApp.
 * Estructura tipica:
 *   { object: 'whatsapp_business_account',
 *     entry: [{
 *       id: '<WABA_ID>',
 *       changes: [{
 *         value: {
 *           messaging_product: 'whatsapp',
 *           metadata: {...},
 *           contacts: [{ profile: { name }, wa_id }],
 *           messages: [{ from, id, timestamp, type, text: { body } | ... }]
 *         },
 *         field: 'messages'
 *       }]
 *     }]
 *   }
 *
 * IMPORTANTE: Meta espera respuesta 200 rapida (<5s). Sino reintenta y
 * puede generar duplicados. Procesamos el msg y devolvemos OK.
 */
function handleWhatsAppWebhook(payload) {
  try {
    const entries = payload.entry || [];
    for (const entry of entries) {
      const changes = entry.changes || [];
      for (const change of changes) {
        if (change.field !== 'messages') continue;
        const value = change.value || {};
        const messages = value.messages || [];
        const contacts = value.contacts || [];
        const contactName = (contacts[0] && contacts[0].profile && contacts[0].profile.name) || '';
        for (const msg of messages) {
          processInboundMessage(msg, contactName);
        }
      }
    }
  } catch(err) {
    logDebugEntry('WA-webhook-CRASH', { error: err.message, stack: err.stack ? String(err.stack).slice(0, 400) : '' });
  }
  // Siempre 200 para que Meta no reintente
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Despacha un mensaje inbound al bot.
 * Sprint 1: solo loguea + responde con menu de saludo.
 * Sprints futuros: state machine, NLU, etc.
 */
function processInboundMessage(msg, contactName) {
  const from = msg.from;
  const type = msg.type;
  const text = type === 'text' && msg.text ? msg.text.body : '';
  logDebugEntry('WA-inbound', { from: from, type: type, text: text.slice(0, 200), name: contactName, msgId: msg.id });

  // Sprint 1: auto-reply a cualquier inbound con menu de saludo.
  // Solo respondemos a mensajes de texto por ahora.
  if (type !== 'text') return;

  const firstName = (contactName || '').toString().trim().split(/\s+/)[0] || '';
  const greeting = firstName ? '¡Hola ' + firstName + '! 🌿' : '¡Hola! 🌿';
  const reply =
    greeting + '\n\n' +
    'Soy el asistente de *Las Nubes*. Te puedo ayudar con:\n\n' +
    '1️⃣  Disponibilidad y precios\n' +
    '2️⃣  Cómo llegar\n' +
    '3️⃣  Hablar con una persona\n\n' +
    'Escribime el número de la opción o tu consulta directa.';

  try {
    sendWhatsAppText(from, reply);
  } catch(sendErr) {
    logDebugEntry('WA-reply-FAIL', { from: from, error: sendErr.message });
  }
}
