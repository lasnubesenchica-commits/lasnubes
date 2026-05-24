/**
 * CRM ligero — registro persistente de mensajes inbound/outbound del bot
 * y endpoints para el dashboard.
 *
 * Hoja "Mensajes": Timestamp | Phone | Direction (in/out) | Type | Content | MessageId
 *
 * Diferencia con DebugLog: no tiene cap, crece indefinidamente para
 * mantener historial completo de conversaciones (>300 entradas).
 */

function _mensajesSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Mensajes');
  if (!sheet) {
    sheet = ss.insertSheet('Mensajes');
    sheet.getRange(1, 1, 1, 6).setValues([['Timestamp', 'Phone', 'Direction', 'Type', 'Content', 'MessageId']]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  return sheet;
}

/**
 * Loguea un mensaje al sheet. Llamado desde sendWhatsApp* y
 * processInboundMessage.
 *
 * @param {string} phone     telefono normalizado (formato E.164 sin +)
 * @param {string} direction 'in' o 'out'
 * @param {string} type      'text', 'button_reply', 'list_reply', 'image',
 *                           'interactive_buttons', 'interactive_list',
 *                           'interactive_cta_url', 'template'
 * @param {string} content   texto del msg o id del boton/lista/imagen
 * @param {string} [msgId]   wamid si disponible
 */
function logMensaje(phone, direction, type, content, msgId) {
  try {
    const sheet = _mensajesSheet();
    const ts = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([
      ts,
      String(phone || '').replace(/\D/g, ''),
      direction,
      type || '',
      String(content || '').slice(0, 1500),
      msgId || ''
    ]);
  } catch(e) { /* swallow */ }
}

// ─── Endpoint: lista de conversaciones ─────────────────────────
// GET ?action=getConversaciones
// Devuelve: { ok, conversations: [{ phone, name, step, lastUpdated, context, lastMsg }] }
function handleGetConversaciones(e) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const convSheet = ss.getSheetByName('Conversaciones');
  if (!convSheet || convSheet.getLastRow() < 2) {
    return ContentService.createTextOutput(JSON.stringify({ ok: true, conversations: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const convData = convSheet.getDataRange().getValues();
  // Pre-cargar ultimo mensaje por telefono desde hoja Mensajes (1 sola pasada)
  const lastByPhone = {};
  const msgSheet = ss.getSheetByName('Mensajes');
  if (msgSheet && msgSheet.getLastRow() > 1) {
    const msgData = msgSheet.getDataRange().getValues();
    for (let i = msgData.length - 1; i >= 1; i--) {  // walk backwards, take first hit per phone
      const ph = (msgData[i][1] || '').toString();
      if (!ph || lastByPhone[ph]) continue;
      lastByPhone[ph] = {
        ts: msgData[i][0],
        direction: msgData[i][2],
        type: msgData[i][3],
        content: msgData[i][4]
      };
    }
  }
  const conversations = [];
  for (let i = 1; i < convData.length; i++) {
    const r = convData[i];
    if (!r[0]) continue;
    const phone = r[0].toString();
    let context = {};
    try { context = r[3] ? JSON.parse(r[3]) : {}; } catch(_) {}
    conversations.push({
      phone:       phone,
      name:        r[4] || '',
      step:        r[1] || 'INITIAL',
      lastUpdated: r[2] instanceof Date ? Utilities.formatDate(r[2], 'America/Panama', 'yyyy-MM-dd HH:mm:ss') : (r[2] || '').toString(),
      context:     context,
      lastMsg:     lastByPhone[phone] || null
    });
  }
  // Ordenar por lastUpdated desc
  conversations.sort((a,b) => (b.lastUpdated || '').localeCompare(a.lastUpdated || ''));
  return ContentService.createTextOutput(JSON.stringify({ ok: true, conversations: conversations }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── Endpoint: mensajes de una conversacion ───────────────────
// GET ?action=getMensajes&phone=XXX[&limit=N]
function handleGetMensajes(e) {
  const phone = (e && e.parameter && e.parameter.phone) ? String(e.parameter.phone).replace(/\D/g, '') : null;
  const limit = parseInt((e && e.parameter && e.parameter.limit) || '200', 10);
  if (!phone) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'MISSING_PHONE' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Mensajes');
  if (!sheet || sheet.getLastRow() < 2) {
    return ContentService.createTextOutput(JSON.stringify({ ok: true, messages: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const data = sheet.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if ((data[i][1] || '').toString() !== phone) continue;
    out.push({
      ts:        data[i][0] instanceof Date ? Utilities.formatDate(data[i][0], 'America/Panama', 'yyyy-MM-dd HH:mm:ss') : (data[i][0] || '').toString(),
      direction: data[i][2] || '',
      type:      data[i][3] || '',
      content:   data[i][4] || '',
      msgId:     data[i][5] || ''
    });
  }
  // Solo ultimos `limit` ordenados ascending
  const messages = out.slice(-limit);
  return ContentService.createTextOutput(JSON.stringify({ ok: true, messages: messages }))
    .setMimeType(ContentService.MimeType.JSON);
}
