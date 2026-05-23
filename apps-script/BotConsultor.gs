/**
 * Bot Consultor de Disponibilidad — Sprint 2
 *
 * State machine + NLU + availability check + pricing reply.
 *
 * Flujo basico:
 *   - Cliente escribe "1" o "disponibilidad" → AWAITING_DATES.
 *   - Cliente envia "del 5 al 8 de junio, 2 personas" → parsea con Claude →
 *     chequea calendario → muestra opciones de cabaña con precio.
 *   - Cliente elige cabaña → handoff a humano (Sprint 3 maneja booking).
 *
 * Hoja 'Conversaciones': [Phone, Step, LastUpdated, Context (JSON), Name]
 */

const BOT_TZ = 'America/Panama';

// ─── Tarifas (espejo de index.html — eventualmente leer de Config) ──
const BOT_RATE_WEEKDAY = 90;
const BOT_RATE_WEEKEND = 110;
const BOT_RECARGO_PERSONA_GRANDE = 20;  // Paseo, Puente
const BOT_RECARGO_PERSONA_PORTAL = 10;  // Portal
const BOT_CABIN_NAMES = {
  verde: 'Paseo por Las Nubes',
  azul:  'Portal hacia Las Nubes',
  lila:  'Puente entre Las Nubes'
};
const BOT_CABIN_CAPACITY = { verde: 4, azul: 2, lila: 4 };

function _botToday() {
  return Utilities.formatDate(new Date(), BOT_TZ, 'yyyy-MM-dd');
}

// ─── Conversaciones sheet ──────────────────────────────────────────
function _convSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Conversaciones');
  if (!sheet) {
    sheet = ss.insertSheet('Conversaciones');
    sheet.getRange(1, 1, 1, 5).setValues([['Phone', 'Step', 'LastUpdated', 'Context', 'Name']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function _getConv(phone) {
  const sheet = _convSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === phone) {
      return {
        row:  i + 1,
        step: data[i][1] || 'INITIAL',
        lastUpdated: data[i][2],
        context: data[i][3] ? (function() { try { return JSON.parse(data[i][3]); } catch(_) { return {}; } })() : {},
        name: data[i][4] || ''
      };
    }
  }
  return null;
}

function _saveConv(phone, step, context, name) {
  const sheet = _convSheet();
  const data  = sheet.getDataRange().getValues();
  const now   = Utilities.formatDate(new Date(), BOT_TZ, 'yyyy-MM-dd HH:mm:ss');
  const ctx   = JSON.stringify(context || {});
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === phone) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[phone, step, now, ctx, name || data[i][4] || '']]);
      return;
    }
  }
  sheet.appendRow([phone, step, now, ctx, name || '']);
}

// ─── Keywords y heuristicas ─────────────────────────────────────────
function _isHumanRequest(text) {
  const t = (text || '').toLowerCase();
  return /\b(humano|persona|operador|asesor|cambio|cancelar|cancela|reembolso|atencion|ayuda urg)\b/.test(t);
}

function _looksLikeDateQuery(text) {
  // numeros, fines de semana, nombres de mes, "del .. al .."
  return /\d|fin de sem|finde|del .* al|enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre/i.test(text || '');
}

// ─── NLU con Claude ────────────────────────────────────────────────
function _parseDatesWithClaude(text, today) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) return null;
  const prompt =
    'Hoy es ' + today + ' (timezone America/Panama). Un cliente escribio en espanol:\n\n"' +
    text.replace(/"/g, '\\"') + '"\n\n' +
    'Extrae las fechas de reserva (checkin/checkout) y numero de personas. Devuelve SOLO JSON con esta forma exacta:\n' +
    '{"checkin":"YYYY-MM-DD"|null,"checkout":"YYYY-MM-DD"|null,"persons":N|null,"confidence":0-1}\n\n' +
    'Reglas:\n' +
    '- checkin = dia que llegan\n' +
    '- checkout = dia que se van (mayor a checkin)\n' +
    '- Si solo mencionan 1 fecha y "N noches", calcular checkout = checkin + N\n' +
    '- Si solo mencionan 1 fecha sin noches, asumir 1 noche y checkout = checkin + 1\n' +
    '- persons = null si no se menciona\n' +
    '- confidence 0 a 1: 1 = muy claro, 0 = ambiguo\n' +
    '- Si no podes inferir fechas con confianza > 0.4, devuelve {"checkin":null,"checkout":null,"persons":null,"confidence":0}\n' +
    '- "este finde" / "este fin de semana" = el viernes-domingo mas proximo\n' +
    '- "proximo finde" = el siguiente fin de semana\n' +
    '- "del viernes al domingo" sin mes = el siguiente viernes-domingo\n' +
    '- Output SOLO el JSON, sin texto adicional.';

  try {
    const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 200,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(res.getContentText());
    const raw  = data.content && data.content[0] && data.content[0].text;
    if (!raw) return null;
    const parsed = JSON.parse(raw.trim());
    logDebugEntry('NLU-dates-OK', { text: text.slice(0, 100), parsed: parsed });
    return parsed;
  } catch(err) {
    logDebugEntry('NLU-dates-FAIL', { text: text.slice(0, 100), error: err.message });
    return null;
  }
}

// ─── Availability + pricing ────────────────────────────────────────
function _botCheckAvailability(checkin, checkout) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const occupied = { verde: false, azul: false, lila: false };
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if (r[9] === 'Abierta') continue;
    if ((r[20] || '').toString().toUpperCase() === 'CANCELADA') continue;
    const cabin = r[3];
    if (!occupied.hasOwnProperty(cabin)) continue;
    const ci = r[4] instanceof Date ? Utilities.formatDate(r[4], BOT_TZ, 'yyyy-MM-dd') : (r[4] || '').toString().slice(0, 10);
    const co = r[5] instanceof Date ? Utilities.formatDate(r[5], BOT_TZ, 'yyyy-MM-dd') : (r[5] || '').toString().slice(0, 10);
    if (ci < checkout && co > checkin) occupied[cabin] = true;
  }
  return { verde: !occupied.verde, azul: !occupied.azul, lila: !occupied.lila };
}

function _botPrecioPorNoche(dateStr) {
  const dow = new Date(dateStr + 'T12:00:00').getDay();
  return (dow === 5 || dow === 6) ? BOT_RATE_WEEKEND : BOT_RATE_WEEKDAY;
}

function _botPrecioCabin(cabin, checkin, checkout, personas) {
  let base = 0;
  const start = new Date(checkin + 'T12:00:00');
  const end   = new Date(checkout + 'T12:00:00');
  for (let cur = new Date(start); cur < end; cur.setDate(cur.getDate() + 1)) {
    base += _botPrecioPorNoche(Utilities.formatDate(cur, BOT_TZ, 'yyyy-MM-dd'));
  }
  if (!personas || personas <= 2) return base;
  const recargo = cabin === 'azul' ? BOT_RECARGO_PERSONA_PORTAL : BOT_RECARGO_PERSONA_GRANDE;
  const nights  = Math.round((end - start) / 86400000);
  return base + recargo * (personas - 2) * nights;
}

function _botFmtFecha(iso) {
  const DIAS  = ['dom','lun','mar','mié','jue','vie','sáb'];
  const MESES = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
  const d = new Date(iso + 'T12:00:00');
  return DIAS[d.getDay()] + ' ' + d.getDate() + ' ' + MESES[d.getMonth()];
}

// ─── Main handler ──────────────────────────────────────────────────
function botHandleMessage(from, text, contactName) {
  const conv = _getConv(from) || { step: 'INITIAL', context: {}, name: contactName || '' };

  // Handoff a humano (prioritario)
  if (_isHumanRequest(text) || text.trim() === '3') {
    sendWhatsAppText(from, '👋 Te derivo con una persona del equipo. En breve te respondemos. Si es urgente, llamanos al +507 6981-2266.');
    _saveConv(from, 'HUMAN_HANDOFF', conv.context, contactName);
    try { sendWhatsAppText('50769812266', '🔔 Handoff: ' + (contactName || from) + ' (' + from + ') escribió: "' + text.slice(0, 200) + '"'); } catch(_) {}
    return;
  }

  const t = (text || '').toLowerCase().trim();

  // Opcion 2: como llegar
  if (t === '2' || t.includes('como llegar') || t.includes('cómo llegar') || t.includes('ubicacion') || t.includes('ubicación') || t.includes('direccion') || t.includes('dirección')) {
    const msg =
      '📍 *Las Nubes — Cómo llegar*\n\n' +
      'Buenos Aires, Chame · Panamá Oeste (faldas del cerro Chicá).\n\n' +
      'Por interamericana → entrar por Pío Pío de Bejuco hacia carretera Bejuco–Sorá. En Buenos Aires, dobla a la derecha hacia Chicá. La cabaña queda a 100m.\n\n' +
      'Más fácil: *Waze* → "Aires de Chicá". Te lleva directo al portón verde.\n\n' +
      'Google Maps:\nhttps://maps.google.com/?q=8.639400,-79.945900';
    sendWhatsAppText(from, msg);
    _saveConv(from, 'SHOWED_INFO', conv.context, contactName);
    return;
  }

  // Opcion 1: disponibilidad — pedir fechas
  if (t === '1' || t.includes('disponibilidad') || t.includes('disponible') || t.includes('precios') || t.includes('cuanto cuesta') || t.includes('cuánto cuesta')) {
    sendWhatsAppText(from,
      '¡Genial! 🌿\n\nDecime las *fechas* y cuántas *personas* serían. Por ejemplo:\n\n' +
      '• "del 5 al 8 de junio, 2 personas"\n' +
      '• "viernes a domingo, 4 personas"\n' +
      '• "este fin de semana, 3 personas"'
    );
    _saveConv(from, 'AWAITING_DATES', conv.context, contactName);
    return;
  }

  // Si esta esperando fechas O el texto tiene pinta de fechas, intentar NLU
  if (conv.step === 'AWAITING_DATES' || _looksLikeDateQuery(text)) {
    const parsed = _parseDatesWithClaude(text, _botToday());
    if (parsed && parsed.checkin && parsed.checkout && parsed.confidence > 0.4) {
      const personas = parsed.persons || 2;
      const avail = _botCheckAvailability(parsed.checkin, parsed.checkout);
      const nights = Math.round((new Date(parsed.checkout + 'T12:00:00') - new Date(parsed.checkin + 'T12:00:00')) / 86400000);
      const fechasStr = _botFmtFecha(parsed.checkin) + ' → ' + _botFmtFecha(parsed.checkout) + ' · ' + nights + (nights === 1 ? ' noche' : ' noches');

      const opciones = [];
      ['azul', 'verde', 'lila'].forEach(c => {
        if (!avail[c]) return;
        if (BOT_CABIN_CAPACITY[c] < personas) return;
        const precio = _botPrecioCabin(c, parsed.checkin, parsed.checkout, personas);
        opciones.push('• *' + BOT_CABIN_NAMES[c] + '* — $' + precio.toFixed(2) + ' total');
      });

      if (opciones.length > 0) {
        sendWhatsAppText(from,
          '✅ Disponibilidad para *' + fechasStr + '* (' + personas + (personas === 1 ? ' persona' : ' personas') + '):\n\n' +
          opciones.join('\n') + '\n\n' +
          'Ver fotos y detalles: https://lasnubes.cloud\n\n' +
          '¿Te interesa alguna? Escribime el *nombre de la cabaña* (Paseo / Portal / Puente) o "3" para hablar con una persona.'
        );
        _saveConv(from, 'SHOWING_AVAILABILITY', { dates: parsed, personas: personas, opciones: opciones.length }, contactName);
      } else {
        sendWhatsAppText(from,
          '😔 No tenemos disponibilidad para *' + fechasStr + '* con ' + personas + (personas === 1 ? ' persona' : ' personas') + '.\n\n' +
          'Podés ver el calendario público para sugerirme otras fechas:\nhttps://lasnubes.cloud\n\n' +
          '¿O preferís hablar con una persona? Escribime "3".'
        );
        _saveConv(from, 'NO_AVAILABILITY', { dates: parsed, personas: personas }, contactName);
      }
      return;
    }
    if (parsed && (!parsed.checkin || parsed.confidence <= 0.4)) {
      sendWhatsAppText(from, '🤔 No logré entender las fechas. ¿Podés escribirlas más claras?\n\nEjemplo: "del 5 al 8 de junio, 4 personas".');
      return;
    }
  }

  // Seleccion de cabana tras mostrar disponibilidad → handoff a humano (Sprint 3 hara booking)
  if (conv.step === 'SHOWING_AVAILABILITY') {
    let elegida = null;
    if (/paseo/i.test(text))      elegida = 'verde';
    else if (/portal/i.test(text)) elegida = 'azul';
    else if (/puente/i.test(text)) elegida = 'lila';
    if (elegida) {
      const dates = conv.context && conv.context.dates;
      const personas = conv.context && conv.context.personas;
      sendWhatsAppText(from,
        '🎉 Excelente elección! Te derivo con una persona para coordinar el pago y confirmar tu reserva en *' + BOT_CABIN_NAMES[elegida] + '*.'
      );
      try {
        sendWhatsAppText('50769812266',
          '📅 Nueva consulta de reserva:\n\n' +
          'Cliente: ' + (contactName || from) + ' (' + from + ')\n' +
          'Cabaña elegida: ' + BOT_CABIN_NAMES[elegida] + '\n' +
          'Fechas: ' + (dates ? dates.checkin + ' → ' + dates.checkout : '?') + '\n' +
          'Personas: ' + (personas || '?'));
      } catch(_) {}
      _saveConv(from, 'PENDING_HUMAN_BOOKING', Object.assign({}, conv.context, { cabin: elegida }), contactName);
      return;
    }
  }

  // Fallback
  sendWhatsAppText(from,
    '🤔 No estoy seguro de cómo ayudarte. Probá con:\n\n' +
    '1️⃣  Ver disponibilidad y precios\n' +
    '2️⃣  Cómo llegar\n' +
    '3️⃣  Hablar con una persona'
  );
}
