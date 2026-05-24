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
 * Despacha un mensaje inbound al bot consultor (BotConsultor.gs).
 * El bot maneja state machine, NLU de fechas, availability y handoff.
 */
function processInboundMessage(msg, contactName) {
  const from = msg.from;
  const type = msg.type;
  let text = '';
  let kind = type;
  if (type === 'text' && msg.text) {
    text = msg.text.body || '';
  } else if (type === 'interactive' && msg.interactive) {
    const it = msg.interactive;
    if (it.type === 'button_reply' && it.button_reply) {
      text = it.button_reply.id || it.button_reply.title || '';
      kind = 'button_reply';
    } else if (it.type === 'list_reply' && it.list_reply) {
      text = it.list_reply.id || it.list_reply.title || '';
      kind = 'list_reply';
    }
  } else if (type === 'image' && msg.image) {
    text = msg.image.id || '';
    kind = 'image';
  }
  logDebugEntry('WA-inbound', { from: from, type: type, kind: kind, text: text.slice(0, 200), name: contactName, msgId: msg.id });
  try { logMensaje(from, 'in', kind || type, text || '', msg.id || ''); } catch(_) {}

  // Detectar primer contacto (no existia conversacion previa) para notificar al admin
  let isFirstContact = false;
  try {
    const existingConv = _getConv(from);
    isFirstContact = !existingConv;
  } catch(_) {}

  // Audio/video/sticker: no soportado todavia
  if (!text && type !== 'text' && type !== 'interactive' && type !== 'image') {
    sendWhatsAppText(from, '🤔 Por ahora solo puedo procesar mensajes de texto, botones e imágenes. Escribime tu consulta o "3" para hablar con una persona.');
    return;
  }

  try {
    botHandleMessage(from, text, contactName, kind);
  } catch(err) {
    logDebugEntry('bot-CRASH', { from: from, error: err.message, stack: err.stack ? String(err.stack).slice(0, 400) : '' });
    try { sendWhatsAppText(from, 'Algo salió mal de mi lado. Te derivo con una persona del equipo 🙏'); } catch(_) {}
  }

  // Notificar al admin si es primer contacto (no spam de notificaciones por cada msg)
  if (isFirstContact && from !== BOT_ADMIN_PHONE) {
    try {
      const name = contactName || from;
      const firstMsg = (text || '').slice(0, 200) || '(mensaje sin texto)';
      const dashUrl = 'https://lasnubes.cloud/dashboard.html#bot:' + from;
      const adminMsg =
        '🔔 *Nuevo cliente en el bot*\n\n' +
        '👤 ' + name + '\n' +
        '📱 +' + from + '\n\n' +
        'Me ha consultado:\n_"' + firstMsg + '"_\n\n' +
        '👀 Ver conversación:\n' + dashUrl;
      sendWhatsAppText(BOT_ADMIN_PHONE, adminMsg);
      logDebugEntry('admin-new-lead-notif', { from: from, name: name });
    } catch(notifErr) {
      logDebugEntry('admin-new-lead-notif-FAIL', { error: notifErr.message });
    }
  }
}
