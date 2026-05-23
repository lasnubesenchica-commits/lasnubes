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

function _botAddDaysISO(iso, n) {
  const d = new Date(iso + 'T12:00:00');
  d.setDate(d.getDate() + n);
  return Utilities.formatDate(d, BOT_TZ, 'yyyy-MM-dd');
}

// Busca hasta 3 fechas cercanas (+/- 10 dias) con al menos 1 cabana libre.
function _botSuggestAlternatives(checkin, checkout, personas) {
  const nights = Math.round((new Date(checkout + 'T12:00:00') - new Date(checkin + 'T12:00:00')) / 86400000);
  const today = _botToday();
  const suggestions = [];
  // Buscar offsets en orden de cercania: +1, -1, +2, -2, ..., +10, -10
  const offsets = [];
  for (let i = 1; i <= 10; i++) { offsets.push(i); offsets.push(-i); }
  for (const offset of offsets) {
    if (suggestions.length >= 3) break;
    const newCheckin  = _botAddDaysISO(checkin, offset);
    const newCheckout = _botAddDaysISO(newCheckin, nights);
    if (newCheckin < today) continue;
    const avail = _botCheckAvailability(newCheckin, newCheckout);
    const cabinsLibres = ['azul', 'verde', 'lila'].filter(c => avail[c] && BOT_CABIN_CAPACITY[c] >= personas);
    if (cabinsLibres.length > 0) {
      suggestions.push({ checkin: newCheckin, checkout: newCheckout, cabinsCount: cabinsLibres.length });
    }
  }
  return suggestions;
}

// ─── Main handler ──────────────────────────────────────────────────
function botHandleMessage(from, text, contactName, kind) {
  const conv = _getConv(from) || { step: 'INITIAL', context: {}, name: contactName || '' };

  // Boton "Ver otras fechas" → vuelve a AWAITING_DATES
  if (kind === 'button_reply' && text === 'try_dates') {
    sendWhatsAppText(from, '🌿 Decime las nuevas fechas:\n\n• "del 5 al 8 de junio, 2 personas"\n• "viernes a domingo, 4 personas"');
    _saveConv(from, 'AWAITING_DATES', conv.context, contactName);
    return;
  }

  // Boton de seleccion de cabana → empieza booking flow
  if (kind === 'button_reply' && /^pick_(verde|azul|lila)$/.test(text)) {
    const elegida = text.split('_')[1];
    return _botStartBooking(from, contactName, conv, elegida);
  }

  // Boton de admin aprobar/rechazar pre-reserva
  if (kind === 'button_reply' && text.indexOf('approve_') === 0) {
    return _botAdminApprove(from, text.replace('approve_', ''));
  }
  if (kind === 'button_reply' && text.indexOf('reject_') === 0) {
    return _botAdminReject(from, text.replace('reject_', ''));
  }

  // Boton "Sugerencia: usar esta fecha"
  if (kind === 'button_reply' && text.indexOf('alt_') === 0) {
    const newCheckin = text.replace('alt_', '');
    const prevDates  = conv.context && conv.context.dates;
    const personas   = (conv.context && conv.context.personas) || 2;
    const nights     = prevDates
      ? Math.round((new Date(prevDates.checkout + 'T12:00:00') - new Date(prevDates.checkin + 'T12:00:00')) / 86400000)
      : 1;
    const newCheckout = _botAddDaysISO(newCheckin, nights);
    return _replyAvailability(from, contactName, conv, newCheckin, newCheckout, personas);
  }

  // Mensaje de imagen → voucher (solo si esta en OFFERING_PAYMENT)
  if (kind === 'image') {
    return _botHandleVoucherImage(from, text, contactName, conv);
  }

  // Email step
  if (conv.step === 'AWAITING_EMAIL') {
    const email = (text || '').trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      sendWhatsAppText(from, '🤔 No parece un email válido. Por favor envíame algo como: nombre@gmail.com');
      return;
    }
    const newCtx = Object.assign({}, conv.context, { email: email });
    _saveConv(from, 'AWAITING_NAME', newCtx, contactName);
    sendWhatsAppText(from, '¡Perfecto! 🌿\n\nÚltimo paso: ¿cuál es tu *nombre completo*?');
    return;
  }

  // Name step → create pre-reservation
  if (conv.step === 'AWAITING_NAME') {
    const fullName = (text || '').trim();
    if (fullName.length < 3 || !/^[a-zA-ZáéíóúñÁÉÍÓÚÑ\s.'\-]+$/.test(fullName)) {
      sendWhatsAppText(from, '🤔 Por favor envíame tu nombre completo (solo letras, sin números).');
      return;
    }
    const newCtx = Object.assign({}, conv.context, { name: fullName });
    return _botCreatePreReservation(from, contactName, newCtx);
  }

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
      return _replyAvailability(from, contactName, conv, parsed.checkin, parsed.checkout, personas);
    }
    if (parsed && (!parsed.checkin || parsed.confidence <= 0.4)) {
      sendWhatsAppText(from, '🤔 No logré entender las fechas. ¿Podés escribirlas más claras?\n\nEjemplo: "del 5 al 8 de junio, 4 personas".');
      return;
    }
  }

  // Fallback de seleccion de cabana por texto (por si el cliente escribe en vez de tocar el boton)
  if (conv.step === 'SHOWING_AVAILABILITY') {
    let elegida = null;
    if (/paseo/i.test(text))      elegida = 'verde';
    else if (/portal/i.test(text)) elegida = 'azul';
    else if (/puente/i.test(text)) elegida = 'lila';
    if (elegida) {
      return botHandleMessage(from, 'pick_' + elegida, contactName, 'button_reply');
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

// ─── Reply helper: muestra disponibilidad con botones, o sugerencias ──
function _replyAvailability(from, contactName, conv, checkin, checkout, personas) {
  const avail  = _botCheckAvailability(checkin, checkout);
  const nights = Math.round((new Date(checkout + 'T12:00:00') - new Date(checkin + 'T12:00:00')) / 86400000);
  const fechasStr = _botFmtFecha(checkin) + ' → ' + _botFmtFecha(checkout) + ' · ' + nights + (nights === 1 ? ' noche' : ' noches');
  const dates = { checkin: checkin, checkout: checkout };

  const opciones = [];
  ['azul', 'verde', 'lila'].forEach(c => {
    if (!avail[c]) return;
    if (BOT_CABIN_CAPACITY[c] < personas) return;
    const precio = _botPrecioCabin(c, checkin, checkout, personas);
    opciones.push({ cabin: c, precio: precio });
  });

  if (opciones.length === 0) {
    // Sin disponibilidad → sugerir fechas cercanas
    const alts = _botSuggestAlternatives(checkin, checkout, personas);
    if (alts.length > 0) {
      const body =
        '😔 No tenemos disponibilidad para *' + fechasStr + '* con ' + personas + (personas === 1 ? ' persona' : ' personas') + '.\n\n' +
        'Pero sí tenemos para estas fechas cercanas:';
      const buttons = alts.slice(0, 3).map(a => ({
        id: 'alt_' + a.checkin,
        title: _botFmtFecha(a.checkin)
      }));
      try {
        sendWhatsAppButtons(from, body, buttons, null, 'Tocá una opción o escribime "3" para hablar con una persona');
      } catch(_) {
        // Fallback a texto si los botones fallan
        sendWhatsAppText(from, body + '\n\n' + alts.map(a => '• ' + _botFmtFecha(a.checkin) + ' → ' + _botFmtFecha(a.checkout)).join('\n') + '\n\nEscribime las fechas que prefieras.');
      }
      _saveConv(from, 'SHOWING_ALTERNATIVES', { dates: dates, personas: personas, alts: alts }, contactName);
      return;
    }
    sendWhatsAppText(from,
      '😔 No tenemos disponibilidad para *' + fechasStr + '* con ' + personas + (personas === 1 ? ' persona' : ' personas') + '.\n\n' +
      'Tampoco tenemos en las fechas cercanas. Podés ver el calendario público para ideas:\nhttps://lasnubes.cloud\n\n' +
      '¿O preferís hablar con una persona? Escribime "3".'
    );
    _saveConv(from, 'NO_AVAILABILITY', { dates: dates, personas: personas }, contactName);
    return;
  }

  // Hay disponibilidad: enviamos cotizacion completa (formato copyPromo) + botones
  const cotizacion = _botCotizacionAvailability(checkin, checkout, opciones, personas);
  sendWhatsAppText(from, cotizacion);

  if (opciones.length === 1) {
    const op = opciones[0];
    try {
      sendWhatsAppButtons(from, '¿Te interesa reservar *' + BOT_CABIN_NAMES[op.cabin] + '*?', [
        { id: 'pick_' + op.cabin, title: '✅ Reservar' },
        { id: 'try_dates',        title: '📅 Otras fechas' }
      ]);
    } catch(_) {
      sendWhatsAppText(from, '¿Te interesa? Escribime "Reservar" o "3" para hablar con una persona.');
    }
  } else {
    const buttons = opciones.map(op => ({
      id: 'pick_' + op.cabin,
      title: BOT_CABIN_NAMES[op.cabin].split(' ')[0]
    }));
    try {
      sendWhatsAppButtons(from, '¿Cuál te interesa?', buttons);
    } catch(_) {
      sendWhatsAppText(from, 'Escribime el nombre de la cabaña (Paseo / Portal / Puente) o "3" para hablar con una persona.');
    }
  }
  _saveConv(from, 'SHOWING_AVAILABILITY', { dates: dates, personas: personas, opciones: opciones.length }, contactName);
}

// ─── Sprint 3: Booking flow ───────────────────────────────────────

const BOT_ADMIN_PHONE = '50769812266';

// Camas por cabana (igual que index.html / dashboard)
const BOT_CABIN_CAMAS = {
  verde: 'La cabaña cuenta con una cama matrimonial queen y un sofá-cama doble.',
  azul:  'La cabaña solo cuenta con una cama matrimonial full. Puede traer colchón inflable.',
  lila:  'La cabaña cuenta con una cama matrimonial queen y una cama auxiliar full.'
};

function _botPaymentInfo() {
  const custom = PropertiesService.getScriptProperties().getProperty('WA_PAYMENT_INFO');
  if (custom) return custom;
  return '*Yappy*: +507 6981-2266\n*ACH*: Banco General · Cuenta 03-91-XX-XXXXXX · A nombre de _[configurar en Script Properties]_';
}

// Secciones comunes (espejo de copyPromo en index.html / admin=1)
function _botSeccionesComunes() {
  return '*Cocina & Alimentación*\n' +
    '• Cocina completamente equipada para preparar sus alimentos\n' +
    '• Área de BBQ disponible\n' +
    '• Incluye café, azúcar y especias básicas\n' +
    '• Cooler grande disponible (no contamos con nevera — traer hielo y alimentos)\n' +
    '• Menú sencillo de comida disponible bajo reserva previa\n' +
    '\n' +
    '*Energía & Conectividad*\n' +
    '• Iluminación 100% solar mediante paneles fotovoltaicos\n' +
    '• Inversor disponible para cargar celulares y dispositivos\n' +
    '• Excelente señal de todas las operadoras\n' +
    '\n' +
    '*Baño & Comodidades*\n' +
    '• Jabón de baño, papel higiénico y toallas limpias incluidos\n' +
    '• Fumigación semanal — se recomienda traer repelente si eres sensible a mosquitos\n' +
    '\n' +
    '*Privacidad*\n' +
    '• Todas las instalaciones son de uso exclusivo para quienes reservan\n' +
    '\n' +
    '*Para Reservar*\n' +
    '• Pago disponible vía Yappy o ACH\n' +
    '• Quedo atento si desea proceder para compartirle las formas de pago';
}

// Texto de cotizacion para 1+ cabanas disponibles (formato copyPromo)
function _botCotizacionAvailability(checkin, checkout, opciones, personas) {
  const nights      = Math.round((new Date(checkout + 'T12:00:00') - new Date(checkin + 'T12:00:00')) / 86400000);
  const fechaIn     = _botFmtFecha(checkin);
  const fechaOut    = _botFmtFecha(checkout);
  const personasLbl = personas + (personas === 1 ? ' persona' : ' personas');

  let intro;
  if (nights === 1) intro = 'Tengo la noche del *' + fechaIn + '* disponible para reserva para ' + personasLbl + '.';
  else              intro = 'Tengo las noches del *' + fechaIn + ' al ' + fechaOut + '* disponibles para reserva para ' + personasLbl + '.';

  let cabinasBlock;
  if (opciones.length === 1) {
    const op = opciones[0];
    cabinasBlock = '🏡 *Cabaña:* ' + BOT_CABIN_NAMES[op.cabin] + '\n';
    if (personas >= 3) cabinasBlock += BOT_CABIN_CAMAS[op.cabin] + '\n';
    cabinasBlock += '💰 *Total:* $' + op.precio.toFixed(2) + '\n';
  } else {
    cabinasBlock = '*Disponibles:*\n';
    opciones.forEach(op => {
      cabinasBlock += '• ' + BOT_CABIN_NAMES[op.cabin] + ' — $' + op.precio.toFixed(2);
      if (BOT_CABIN_CAPACITY[op.cabin] > 2 && personas <= 2) {
        cabinasBlock += ' (hasta ' + BOT_CABIN_CAPACITY[op.cabin] + ' personas)';
      }
      cabinasBlock += '\n';
    });
  }

  return intro + '\n\n' + cabinasBlock +
    '\nCheck in ' + fechaIn + ': 2:00 pm\nCheck out ' + fechaOut + ': 11:00 am\n' +
    (nights > 1 ? (nights + 1) + ' días, ' + nights + ' noches\n' : '') +
    '\n' + _botSeccionesComunes();
}

// Texto de confirmacion al cliente (espejo de _buildClienteShareText en dashboard)
function _botConfirmacionText(reservation, publicUrl, referralCode, referralAmount) {
  const meta = tipoEmailMeta(reservation);
  const CABIN_NAMES_FULL = {
    verde: 'Paseo por Las Nubes',
    azul:  'Portal hacia Las Nubes',
    lila:  'Puente entre Las Nubes'
  };
  const cabin = CABIN_NAMES_FULL[reservation.cabin] || reservation.cabinName || reservation.cabin || '';
  const tipo  = meta.tipo;

  let fechasLine;
  if (tipo === 'pasatarde')      fechasLine = '📅 ' + meta.checkinFmt + ' · Pasatarde 12:30pm – 7pm';
  else if (tipo === 'pasadia')   fechasLine = '📅 ' + meta.checkinFmt + ' · Pasadía 9am – 5pm';
  else if (tipo === 'early')     fechasLine = '📅 ' + meta.checkinFmt + ' → ' + meta.checkoutFmt + ' · 1 noche (entra 9am)';
  else if (tipo === 'late')      fechasLine = '📅 ' + meta.checkinFmt + ' → ' + meta.checkoutFmt + ' · 1 noche (sale 4pm)';
  else                           fechasLine = '📅 ' + meta.checkinFmt + ' → ' + meta.checkoutFmt + ' · ' + meta.estanciaValue + (meta.estanciaValue === 1 ? ' noche' : ' noches');

  const isPasadia = (tipo === 'pasatarde' || tipo === 'pasadia');
  let checkinH = '2:00 pm', checkoutH = '11:00 am';
  if (tipo === 'early') checkinH = '9:00 am';
  if (tipo === 'late')  checkoutH = '4:00 pm';
  if (reservation.checkoutExtendido && (tipo === 'noche' || tipo === 'early')) {
    checkoutH = '12:30 pm (cortesía)';
  }

  let text = '¡Reserva confirmada! 🌿\n\n';
  text += '👤 ' + (reservation.name || '') + '\n';
  text += '🏡 ' + cabin + '\n';
  text += fechasLine + '\n';
  if (reservation.persons) text += '👥 ' + reservation.persons + (reservation.persons == 1 ? ' persona' : ' personas') + '\n';
  if (!isPasadia) text += '\nCheck-in: ' + checkinH + '\nCheck-out: ' + checkoutH + '\n';
  if (reservation.origin === 'Referido') text += '\n🤝 Tarifa pactada con descuento del Programa Amigos.\n';
  if (publicUrl) text += '\nVer detalles e instrucciones:\n' + publicUrl;
  if (referralCode) {
    const amt = referralAmount || 20;
    text += '\n\n🤝 *Programa Amigos de Las Nubes*';
    text += '\nSi durante tu estadía disfrutas la experiencia y deseas compartirla, este es tu código personal: *' + referralCode + '*';
    text += '\n\n• Si un amigo reserva con tu código recibe *$' + amt + ' off*.';
    text += '\n• Y tú *$' + amt + '* para tu próxima visita.';
    text += '\n(Aplica Dom–Jue, reservas directas)';
  }
  text += '\n\n📸 No olvides etiquetarnos en nuestras redes:';
  text += '\nInstagram: https://www.instagram.com/las_nubes_de_chica/';
  text += '\nTikTok: https://www.tiktok.com/@las_nubes_en_chica';
  return text;
}

function _botStartBooking(from, contactName, conv, cabin) {
  const dates    = conv.context && conv.context.dates;
  const personas = (conv.context && conv.context.personas) || 2;
  if (!dates) {
    sendWhatsAppText(from, '🤔 Perdí el contexto de las fechas. ¿Podés decirme de nuevo cuándo querés reservar?');
    _saveConv(from, 'AWAITING_DATES', conv.context, contactName);
    return;
  }
  const precio  = _botPrecioCabin(cabin, dates.checkin, dates.checkout, personas);
  const fechas  = _botFmtFecha(dates.checkin) + ' → ' + _botFmtFecha(dates.checkout);
  sendWhatsAppText(from,
    '🎉 ¡Excelente! Reservando *' + BOT_CABIN_NAMES[cabin] + '* para *' + fechas + '* (' + personas + (personas === 1 ? ' persona' : ' personas') + ').\n\n' +
    '💰 *Total: $' + precio.toFixed(2) + '*\n\n' +
    '💳 *Métodos de pago*\n' + _botPaymentInfo() + '\n\n' +
    'Una vez transferido, *enviame el comprobante como imagen* aquí mismo y procesamos tu reserva.'
  );
  _saveConv(from, 'OFFERING_PAYMENT', Object.assign({}, conv.context, { cabin: cabin, precio: precio }), contactName);
}

function _botHandleVoucherImage(from, imageId, contactName, conv) {
  if (conv.step !== 'OFFERING_PAYMENT' && conv.step !== 'AWAITING_VOUCHER_RETRY') {
    sendWhatsAppText(from, '📷 Recibí tu imagen, pero no estamos en una reserva activa. Si querés reservar, escribime "1" o "disponibilidad".');
    return;
  }
  sendWhatsAppText(from, '⏳ Procesando tu comprobante...');
  const img = fetchWhatsAppImage(imageId);
  if (!img) {
    sendWhatsAppText(from, '⚠️ No pude descargar tu imagen. Probá enviarla de nuevo o escribime "3" para hablar con una persona.');
    return;
  }
  let voucher;
  try {
    const out = parseVoucherWithClaude(img.base64, img.mimeType);
    voucher = JSON.parse(out.getContent());
  } catch(err) {
    logDebugEntry('bot-voucher-OCR-CRASH', { error: err.message });
    sendWhatsAppText(from, '⚠️ No pude leer el voucher. ¿Podés enviarlo más claro o escribime "3" para una persona?');
    return;
  }
  if (!voucher || !voucher.ok || !voucher.codTransferencia) {
    sendWhatsAppText(from,
      '⚠️ No pude leer los datos del voucher. Asegurate que la imagen sea clara y tenga:\n\n' +
      '• Monto\n• Código/referencia\n• Fecha\n\nReenviame la imagen o escribime "3" para una persona.'
    );
    _saveConv(from, 'AWAITING_VOUCHER_RETRY', conv.context, contactName);
    return;
  }
  const monto = parseFloat(voucher.monto) || 0;
  const expectedDeposit = (conv.context && conv.context.precio) ? (conv.context.precio * 0.5) : 0;
  const newCtx = Object.assign({}, conv.context, {
    voucher: {
      monto: monto,
      codTransferencia: voucher.codTransferencia,
      fechaPago: voucher.fechaPago || _botToday(),
      sender:   voucher.sender || ''
    }
  });
  sendWhatsAppText(from,
    '✅ ¡Comprobante recibido!\n\n' +
    '*Remitente:* ' + (voucher.sender || '—') + '\n' +
    '*Monto:* $' + monto.toFixed(2) + '\n' +
    '*Código:* ' + voucher.codTransferencia + '\n\n' +
    'Para finalizar, ¿me podés enviar tu *email*?'
  );
  _saveConv(from, 'AWAITING_EMAIL', newCtx, contactName);
}

function _botCreatePreReservation(from, contactName, ctx) {
  const dates    = ctx.dates;
  const cabin    = ctx.cabin;
  const personas = ctx.personas || 2;
  const email    = ctx.email;
  const fullName = ctx.name;
  const voucher  = ctx.voucher || {};
  const precio   = ctx.precio || _botPrecioCabin(cabin, dates.checkin, dates.checkout, personas);
  const id       = Date.now().toString();
  const today    = _botToday();
  const CABIN_NAMES = {
    verde: 'Paseo por Las Nubes',
    azul:  'Portal hacia Las Nubes',
    lila:  'Puente entre Las Nubes'
  };

  try {
    const sheet = getOrCreateSheet();
    sheet.appendRow([
      id,
      _safeCell(fullName),
      CABIN_NAMES[cabin] || cabin,
      cabin,
      dates.checkin,
      dates.checkout,
      personas,
      precio,
      voucher.monto || 0,
      'Directa',
      id,
      0,                       // serviceFee
      precio,                  // neto
      '',                      // alerta
      _safeCell(fullName),     // pagador
      today,                   // fechaReserva
      '',                      // fechaPago
      0,                       // montoPagado
      _safeCell(voucher.codTransferencia || ''),
      voucher.monto ? '$' + voucher.monto.toFixed(2) : '',
      'PENDIENTE',             // estadoPago → admin debe aprobar
      _safeCell(email),
      _safeCell('🤖 Pre-reserva vía bot WhatsApp · pendiente revisión'),
      _safeCell(from),
      'noche'                  // tipo
    ]);
    logDebugEntry('bot-prereserva-OK', { id: id, name: fullName, cabin: cabin, from: from });
  } catch(err) {
    logDebugEntry('bot-prereserva-FAIL', { error: err.message, stack: err.stack ? String(err.stack).slice(0, 400) : '' });
    sendWhatsAppText(from, '⚠️ Hubo un problema al registrar tu reserva. Te derivo con una persona del equipo.');
    try { sendWhatsAppText(BOT_ADMIN_PHONE, '⚠️ Bot falló al crear pre-reserva de ' + (contactName || from) + ': ' + err.message); } catch(_) {}
    return;
  }

  sendWhatsAppText(from,
    '🎉 ¡Pre-reserva creada!\n\n' +
    'Estamos revisando los datos. En breve te confirmamos tu reserva. ¡Gracias!'
  );

  const fechas = _botFmtFecha(dates.checkin) + ' → ' + _botFmtFecha(dates.checkout);
  const adminMsg =
    '📋 *Nueva pre-reserva via bot*\n\n' +
    '👤 ' + fullName + '\n' +
    '📧 ' + email + '\n' +
    '📱 ' + from + '\n\n' +
    '🏡 ' + (CABIN_NAMES[cabin] || cabin) + '\n' +
    '📅 ' + fechas + '\n' +
    '👥 ' + personas + (personas === 1 ? ' persona' : ' personas') + '\n' +
    '💰 Total: $' + precio.toFixed(2) + '\n\n' +
    '💳 Voucher: $' + (voucher.monto || 0).toFixed(2) + ' (' + (voucher.sender || '?') + ')\n' +
    '#️⃣ Código: ' + (voucher.codTransferencia || '?');
  try {
    sendWhatsAppButtons(BOT_ADMIN_PHONE, adminMsg, [
      { id: 'approve_' + id, title: '✅ Aprobar' },
      { id: 'reject_'  + id, title: '❌ Rechazar' }
    ]);
  } catch(_) {
    sendWhatsAppText(BOT_ADMIN_PHONE, adminMsg + '\n\nResponder "approve_' + id + '" o "reject_' + id + '" para confirmar.');
  }

  _saveConv(from, 'PENDING_REVIEW', Object.assign({}, ctx, { reservaId: id }), contactName);
}

function _botAdminApprove(adminPhone, reservaId) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === reservaId) {
      const row = i + 1;
      sheet.getRange(row, 21).setValue('PAGA');
      const reservation = {
        id:       data[i][0],
        name:     data[i][1],
        cabin:    data[i][3],
        checkin:  data[i][4] instanceof Date ? Utilities.formatDate(data[i][4], BOT_TZ, 'yyyy-MM-dd') : data[i][4],
        checkout: data[i][5] instanceof Date ? Utilities.formatDate(data[i][5], BOT_TZ, 'yyyy-MM-dd') : data[i][5],
        persons:  data[i][6],
        amount:   data[i][7],
        deposit:  data[i][8],
        origin:   data[i][9],
        email:    data[i][21],
        telefono: data[i][23],
        tipo:     data[i][24] || 'noche',
        checkoutExtendido: data[i][28] === true || data[i][28] === 'TRUE' || data[i][28] === 'true' || data[i][28] === 1
      };
      // Construir texto rico (espejo de _buildClienteShareText del dashboard) y enviar
      // como session message — el bot acaba de tener interaccion con el cliente,
      // estamos dentro de la ventana de 24h.
      try {
        let publicUrl = '';
        try { publicUrl = getPublicReservaUrl(reservation.id); } catch(_) {}
        let referralCode = null;
        try {
          const isDormido = !(reservation.tipo === 'pasatarde' || reservation.tipo === 'pasadia');
          if (reservation.email && isDormido) referralCode = getOrCreateReferralCode(reservation.email, reservation.telefono, reservation.name);
        } catch(_) {}
        const texto = _botConfirmacionText(reservation, publicUrl, referralCode, 20);
        sendWhatsAppText(reservation.telefono, texto);
        sendWhatsAppText(adminPhone, '✅ Reserva ' + reservaId + ' aprobada y confirmación enviada al cliente.');
      } catch(err) {
        logDebugEntry('bot-approve-FAIL', { reservaId: reservaId, error: err.message });
        // Fallback al template HSM (caso raro: session fuera de 24h)
        try {
          sendWAReservaConfirmada(reservation);
          sendWhatsAppText(adminPhone, '✅ Reserva ' + reservaId + ' aprobada (template HSM enviado, session habia expirado).');
        } catch(err2) {
          sendWhatsAppText(adminPhone, '⚠️ Reserva ' + reservaId + ' marcada PAGA pero falló envío al cliente: ' + err.message);
        }
      }
      return;
    }
  }
  sendWhatsAppText(adminPhone, '⚠️ No encontré la reserva ' + reservaId);
}

function _botAdminReject(adminPhone, reservaId) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === reservaId) {
      const row = i + 1;
      sheet.getRange(row, 21).setValue('CANCELADA');
      const clientPhone = data[i][23];
      sendWhatsAppText(adminPhone, '❌ Reserva ' + reservaId + ' rechazada y cancelada en el sheet.');
      try {
        sendWhatsAppText(clientPhone, '😔 Hubo un inconveniente con tu reserva. En breve te contactamos para resolverlo.');
      } catch(_) {}
      return;
    }
  }
  sendWhatsAppText(adminPhone, '⚠️ No encontré la reserva ' + reservaId);
}
