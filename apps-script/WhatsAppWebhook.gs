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
  } else if (type === 'button' && msg.button) {
    // Boton quick-reply de una plantilla (ej. "Ya me retiré" en checkout).
    text = msg.button.payload || msg.button.text || '';
    kind = 'button_reply';
  }
  logDebugEntry('WA-inbound', { from: from, type: type, kind: kind, text: text.slice(0, 200), name: contactName, msgId: msg.id });
  try { logMensaje(from, 'in', kind || type, text || '', msg.id || ''); } catch(_) {}

  // Trackear el último inbound del admin para avisarle 1h antes de que se
  // cierre su ventana de 24h. Sin esto, las alertas de sesión (portón, nuevo
  // cliente, etc.) dejan de entregarse silenciosamente. Verificación en
  // verificarVentanaAdmin (trigger horario).
  if (String(from).replace(/\D/g, '') === '50769812266') {
    try {
      PropertiesService.getScriptProperties().setProperty('ADMIN_LAST_INBOUND_TS', new Date().toISOString());
    } catch(_) {}
  }

  // Erika (limpieza): cualquier mensaje que mande abre la ventana de 24h y le
  // respondemos con el parte de limpieza de hoy. No entra al flujo del Agente
  // ni dispara la notificacion de "nuevo cliente".
  try {
    const limpiezaPhone = (PropertiesService.getScriptProperties().getProperty('LIMPIEZA_PHONE') || '').replace(/\D/g, '');
    if (limpiezaPhone && String(from).replace(/\D/g, '') === limpiezaPhone) {
      sendWhatsAppText(from, _buildLimpiezaMessage());
      logDebugEntry('limpieza-on-demand', { from: from });
      return;
    }
  } catch(err) {
    logDebugEntry('limpieza-on-demand-FAIL', { error: err.message });
  }

  // Detectar primer contacto (no existia conversacion previa) para notificar al admin
  let isFirstContact = false;
  try {
    const existingConv = _getConv(from);
    isFirstContact = !existingConv;
  } catch(_) {}

  // Audio/video/sticker: no soportado todavia
  if (!text && type !== 'text' && type !== 'interactive' && type !== 'image' && type !== 'button') {
    sendWhatsAppText(from, '🤔 Por ahora solo puedo procesar mensajes de texto, botones e imágenes. Escribime tu consulta o "3" para hablar con una persona.');
    return;
  }

  // Debounce de mensajes de TEXTO consecutivos del mismo cliente. Caso real
  // Malu: escribió "me cotiza sábado 13 a domingo 14?" + "para 2 personas" a
  // 5s de distancia, el bot respondió cada uno por separado (no entendí +
  // tarifas) generando ruido y confusión. Buttons / lists / images se
  // procesan al instante (acciones decisivas).
  let combinedText = text;
  if (kind === 'text' && type === 'text') {
    const debouncedText = _debounceInbound(from, text);
    if (debouncedText === null) return;   // otro webhook concurrente toma el control
    combinedText = debouncedText;
  }

  try {
    botHandleMessage(from, combinedText, contactName, kind);
  } catch(err) {
    logDebugEntry('bot-CRASH', { from: from, error: err.message, stack: err.stack ? String(err.stack).slice(0, 400) : '' });
    try { sendWhatsAppText(from, 'Algo salió mal de mi lado. Te derivo con una persona del equipo 🙏'); } catch(_) {}
  }

  // Notificar al admin si es primer contacto (no spam de notificaciones por cada msg)
  // Excluye button_reply: tocar un botón de una plantilla (ej. "Ya me retiré"
  // en checkout) no es un nuevo lead, es un huésped existente respondiendo.
  if (isFirstContact && from !== BOT_ADMIN_PHONE && kind !== 'button_reply') {
    try {
      const name = contactName || from;
      const firstMsg = (text || '').slice(0, 200) || '(mensaje sin texto)';
      const dashUrl = 'https://lasnubes.cloud/dashboard.html?admin=1#bot:' + from;
      const adminMsg =
        '🔔 *El Agente está atendiendo a un nuevo cliente*\n\n' +
        '👤 ' + name + '\n' +
        '📱 +' + from + '\n\n' +
        'Ha consultado:\n_"' + firstMsg + '"_\n\n' +
        '👀 Ver conversación:\n' + dashUrl;
      _botAdminAlert('nuevoCliente', adminMsg);
      logDebugEntry('admin-new-lead-notif', { from: from, name: name });
    } catch(notifErr) {
      logDebugEntry('admin-new-lead-notif-FAIL', { error: notifErr.message });
    }
  }
}

// Debounce de mensajes de texto consecutivos. Cada webhook:
//   1. Agrega su texto a un buffer compartido por teléfono (CacheService).
//   2. Marca su timestamp como "el último arribo".
//   3. Duerme DEBOUNCE_MS y luego compara: si su TS sigue siendo el último,
//      procesa el buffer completo concatenado; si llegó otro mensaje en el
//      ínterin (que sobrescribió el último-TS), retorna null y sale sin
//      procesar — el webhook más reciente se encarga de todo.
// Resultado: ráfagas de mensajes en <2s se concatenan en una sola
// interacción, evitando respuestas duplicadas/contradictorias.
function _debounceInbound(from, text) {
  const DEBOUNCE_MS = 2000;
  const cache       = CacheService.getScriptCache();
  const latestKey   = 'wa-debounce-latest-' + from;
  const bufferKey   = 'wa-debounce-buffer-' + from;
  const ts          = String(Date.now()) + '-' + Math.random().toString(36).slice(2, 8);

  let buffer = [];
  try { buffer = JSON.parse(cache.get(bufferKey) || '[]'); } catch(_) {}
  buffer.push(text || '');
  cache.put(bufferKey, JSON.stringify(buffer), 60);
  cache.put(latestKey, ts, 60);

  Utilities.sleep(DEBOUNCE_MS);

  if (cache.get(latestKey) !== ts) {
    logDebugEntry('debounce-yield', { from: from });
    return null;
  }

  try { buffer = JSON.parse(cache.get(bufferKey) || '[]'); } catch(_) {}
  cache.remove(bufferKey);
  cache.remove(latestKey);

  const combined = buffer.filter(Boolean).join(' ').trim();
  if (buffer.length > 1) {
    logDebugEntry('debounce-merge', { from: from, count: buffer.length, combined: combined.slice(0, 200) });
  }
  return combined;
}
