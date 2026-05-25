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
const BOT_RECARGO_COMBO_5 = 80;          // 5 personas: Puente + Portal contiguas, por noche
const BOT_RECARGO_COMBO_6 = 100;         // 6 personas
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
  // Agresivo: numeros, dias de la semana, meses, expresiones relativas
  return /\d|fin de sem|finde|del .* al|enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre|lunes|martes|mi[eé]rcoles|jueves|viernes|s[aá]bado|domingo|\bhoy\b|\bma[ñn]ana\b|pasado ma[ñn]ana|esta semana|pr[oó]xima semana|semana que viene|mes que viene|fin de a[ñn]o/i.test(text || '');
}

// Detecta consultas vagas tipo "para julio", "segunda semana de agosto",
// "principios de septiembre", "el mes que viene". En estos casos no
// intentamos cotizar — mandamos al calendario publico para que el cliente
// elija fechas concretas.
function _looksLikeVagueDateQuery(text) {
  const t = (text || '').toLowerCase();
  const MONTHS = '(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)';
  // "primera/segunda/tercera/cuarta/última semana de <mes>"
  if (new RegExp('(primera|segunda|tercera|cuarta|[uú]ltima)\\s+semana\\s+de\\s+' + MONTHS, 'i').test(t)) return true;
  // "principios/mediados/finales de <mes>"
  if (new RegExp('(principios|mediados|fines|finales)\\s+de\\s+' + MONTHS, 'i').test(t)) return true;
  // "para/en/durante <mes>" sin numero de dia ni rango
  const hasMonth   = new RegExp('\\b' + MONTHS + '\\b', 'i').test(t);
  const hasDayNum  = /\b\d{1,2}\b/.test(t);
  const hasDayRange = /del\s+\d+\s+al\s+\d+|al\s+\d+\s+de/i.test(t);
  if (hasMonth && !hasDayNum && !hasDayRange) return true;
  // "proximo mes" / "mes que viene" sin dia especifico
  if (/(pr[oó]xim[oa]|siguiente)\s+mes|mes\s+que\s+viene/i.test(t) && !hasDayNum) return true;
  return false;
}

// Envia link al calendario publico para consultas de fechas vagas.
function _botSendCalendarLink(from, contactName) {
  const body =
    '🗓 Para fechas amplias o flexibles, podés explorar todo el calendario de disponibilidad en nuestra página.\n\n' +
    'Tocá el botón abajo, mirá los días libres y cuando tengas fechas concretas decímelas por aquí (ej: _"del 5 al 8 de julio, 2 personas"_) para cotizar al instante. 🤝';
  try {
    sendWhatsAppCTAUrl(from, body, '📅 Ver calendario', 'https://lasnubes.cloud');
  } catch(_) {
    sendWhatsAppText(from, body + '\n\n👉 https://lasnubes.cloud');
  }
}

// Detecta mensajes tipicos de campanas de Instagram/Facebook:
// "¡Hola! Quiero más información", "Hola, me interesa", etc. En esos casos
// mandamos una bienvenida con pitch breve de las 3 cabanas y amenidades.
function _isCampaignInquiry(text) {
  const t = (text || '').toLowerCase().trim();
  if (/quiero\s+m[aá]s\s+informaci[oó]n/.test(t)) return true;
  if (/^(hola[!,.\s]+)?(quisiera|quiero|necesito|me\s+gustar[ií]a)\s+(m[aá]s\s+)?info(rmaci[oó]n)?/.test(t)) return true;
  if (/^m[aá]s\s+info(rmaci[oó]n)?\b/.test(t)) return true;
  if (/informaci[oó]n\s+por\s+favor/.test(t)) return true;
  if (/^(hola[!,.\s]+)?me\s+interesa(n)?(\s+(saber|conocer|sus|las|tus))?/.test(t)) return true;
  return false;
}

function _botSendCampaignWelcome(from, contactName) {
  const firstName = ((contactName || '').toString().trim().split(/\s+/)[0]) || '';
  const greeting  = firstName ? '¡Hola ' + firstName + '! 🌿' : '¡Hola! 🌿';

  const info =
    greeting + ' Gracias por escribirnos.\n\n' +
    '*Las Nubes* es un refugio de tres cabañas privadas en las faldas del *Cerro Chicá*, a 1h 15min de Ciudad de Panamá. Naturaleza, vistas y privacidad total.\n\n' +
    '🏡 *Nuestras cabañas*\n' +
    '• *Paseo por Las Nubes* — 2 a 4 personas (cama queen + sofá-cama doble)\n' +
    '• *Portal hacia Las Nubes* — 2 personas (cama matrimonial full)\n' +
    '• *Puente entre Las Nubes* — 2 a 4 personas (cama queen + cama auxiliar)\n\n' +
    '✨ *Lo que incluye*\n' +
    '• Cocina equipada con área de BBQ y cooler grande\n' +
    '• Iluminación 100% solar\n' +
    '• Toallas, jabón y papel higiénico incluidos\n' +
    '• Uso exclusivo de las instalaciones — sin vecinos\n\n' +
    '📅 *Para reservar o consultar disponibilidad*, cuéntame *fechas* y *personas*. Por ejemplo:\n' +
    '   _"del 5 al 8 de junio, 2 personas"_\n\n' +
    'Te envío disponibilidad y precio al instante. 🤝';

  sendWhatsAppText(from, info);
  _botSendMainMenu(from, contactName, false, '¿Querés explorar más? Tocá *Ver opciones* abajo 👇');
  _saveConv(from, 'AWAITING_DATES', {}, contactName);
}

// Tarifas base para 2 personas en las 3 cabanas (mismas tarifas, recargo por
// persona extra solo aplica a partir de 3 personas).
function _botSendPricingInfo(from, contactName, conv) {
  const msg =
    '💰 *Tarifas por noche · 2 personas*\n\n' +
    'Mismo precio en las tres cabañas (*Paseo*, *Portal* y *Puente*):\n\n' +
    '• Domingo a jueves: *$' + BOT_RATE_WEEKDAY + '/noche*\n' +
    '• Viernes y sábado: *$' + BOT_RATE_WEEKEND + '/noche*\n\n' +
    '_Para 3 o 4 personas hay un pequeño recargo por persona adicional._\n\n' +
    '¿Querés verificar disponibilidad para alguna fecha? Decime *fechas* y *personas* (ej: _"del 5 al 8 de junio, 2 personas"_) y te cotizo al instante. 🤝';
  sendWhatsAppText(from, msg);
  _saveConv(from, 'AWAITING_DATES', (conv && conv.context) || {}, contactName);
}

// Respuestas puntuales a topics FAQ. Devuelve true si manejo el mensaje.
// Fotos por cabaña (subset de la galería del sitio). lh3 de Drive con
// =w1280 para que se sirvan como JPEG dimensionado (compatible con WhatsApp).
const BOT_CABIN_PHOTOS = {
  verde: [
    'https://lh3.googleusercontent.com/d/1UitkVZ9KCRqvWv5zNueXlfil01FBpDLB=w1280',
    'https://lh3.googleusercontent.com/d/1ETkiIiNih83W0rTMcSndTNbQvDInkM6W=w1280',
    'https://lh3.googleusercontent.com/d/1fCDD95d680rEyUTh3c0moXe3H29nf1vt=w1280',
    'https://lh3.googleusercontent.com/d/1jnNxg_3WXQT0f5DjciDgdhRp8HsJkZbC=w1280'
  ],
  azul: [
    'https://lh3.googleusercontent.com/d/1kolAp8PKDO3ws6abcUUfD2hpN_3ZLBjB=w1280',
    'https://lh3.googleusercontent.com/d/171GVtaWLZAZCqds8yXLOVfVUgbT1URfy=w1280',
    'https://lh3.googleusercontent.com/d/1jn9m_ON3_UnZtq_PRiln9TKHt1c2zimj=w1280',
    'https://lh3.googleusercontent.com/d/1mqwRDpSB5p_6ozufxtC18LQLmDnPTa8j=w1280'
  ],
  lila: [
    'https://lh3.googleusercontent.com/d/1TktbGLMLIXCRLXQh-ctE00wv7NrHjPjg=w1280',
    'https://lh3.googleusercontent.com/d/1xpjjxmuC5nqiF8VsCfwZixRkHLFLPlPR=w1280',
    'https://lh3.googleusercontent.com/d/1D0es4EEzyr10UOr1ohKTGNZpw9AMX8FV=w1280',
    'https://lh3.googleusercontent.com/d/1dPcwrLUoqgonOdZ-Y7hDGoYxjbmU5byo=w1280'
  ]
};

// Detecta a qué cabaña se refiere el texto, por NOMBRE o por color.
//   Paseo / verde · Portal / azul · Puente / lila
function _botDetectCabin(text) {
  const t = (text || '').toLowerCase();
  if (/\bpaseo\b|\bverde\b/.test(t))  return 'verde';
  if (/\bportal\b|\bazul\b/.test(t))  return 'azul';
  if (/\bpuente\b|\blila\b/.test(t))  return 'lila';
  return null;
}

// Envia fotos de una cabaña (o un sampler de las 3 si no se especifica).
function _botSendCabinPhotos(from, contactName, cabinKey) {
  if (cabinKey && BOT_CABIN_PHOTOS[cabinKey]) {
    const fotos = BOT_CABIN_PHOTOS[cabinKey];
    fotos.forEach((url, i) => {
      try { sendWhatsAppImage(from, url, i === 0 ? ('🏡 ' + BOT_CABIN_NAMES[cabinKey]) : ''); } catch(_) {}
    });
    sendWhatsAppText(from,
      'Estas son algunas fotos de *' + BOT_CABIN_NAMES[cabinKey] + '* 🌿\n\n' +
      'Galería completa: https://lasnubes.cloud/#cabanas-' + cabinKey + '\n\n' +
      '¿Querés cotizar para alguna fecha? Decime *fechas* y *personas* 📅'
    );
    logDebugEntry('bot-fotos', { from: from, cabin: cabinKey });
    return;
  }
  // Sin cabaña específica → una foto de cada una + invitación a elegir
  ['verde', 'azul', 'lila'].forEach(c => {
    try { sendWhatsAppImage(from, BOT_CABIN_PHOTOS[c][0], '🏡 ' + BOT_CABIN_NAMES[c]); } catch(_) {}
  });
  sendWhatsAppText(from,
    '🌿 Estas son nuestras tres cabañas. Decime de cuál querés ver *más fotos* (Paseo, Portal o Puente) ' +
    'o mirá la galería completa en https://lasnubes.cloud/#cabanas\n\n' +
    'Y si ya tenés fechas en mente, te cotizo al instante 📅'
  );
  logDebugEntry('bot-fotos', { from: from, cabin: 'all' });
}

// Consultas de acceso: con qué auto se llega, o cómo llegar sin auto.
// Devuelve true si manejó el mensaje. Transporte (bus) se chequea primero.
function _botHandleAccesoQuery(from, contactName, text) {
  const t = (text || '').toLowerCase();

  // Sin auto / transporte público / bus / Albrook
  if (/\b(sin\s+(auto|carro|veh[ií]culo|movilidad)|no\s+tengo\s+(auto|carro|veh[ií]culo|carro)|transporte\s+p[uú]blic|en\s+bus\b|\bbus\b|autob[uú]s|albrook|terminal)\b/.test(t)) {
    sendWhatsAppText(from,
      '🚌 *Cómo llegar sin auto*\n\n' +
      'Desde la terminal de *Albrook* podés tomar cualquier bus hacia el interior y bajarte en el *Pío Pío de Bejuco*. Ahí te recogemos y trasladamos a la cabaña, ida y vuelta, por *$20 adicionales* en tu reserva. 🌿\n\n' +
      '¿Querés cotizar para alguna fecha? Decime *fechas* y *personas* 📅'
    );
    logDebugEntry('bot-acceso', { from: from, tipo: 'sin-auto' });
    return true;
  }

  // Tipo de auto / 4x4 / sedán / estado del camino
  if (/\b(4\s*x\s*4|4x4|cuatro\s+por\s+cuatro|sed[aá]n|tipo\s+de\s+(auto|carro|veh)|cualquier\s+(auto|carro|veh|tipo\s+de\s+auto)|mi\s+(auto|carro)|camino\s+(es|est[aá]|de)|carretera\s+(es|est[aá]|de|en)|c[oó]mo\s+(es|est[aá])\s+(el\s+camino|la\s+carretera|la\s+calle)|se\s+puede\s+(ir|llegar|subir|entrar)\s+en|necesito\s+(un\s+)?4|hace\s+falta\s+(un\s+)?4|sube[ns]?\s+carros?)\b/.test(t)) {
    sendWhatsAppText(from,
      '🚗 *Acceso a Las Nubes*\n\n' +
      'La carretera es de asfalto y en muy buen estado hasta la entrada del proyecto. Luego son unos *5 minutos en camino de tosca fina*. Recibimos *sedanes y todo tipo de autos* sin problema. 🌿\n\n' +
      '¿Querés cotizar para alguna fecha? Decime *fechas* y *personas* 📅'
    );
    logDebugEntry('bot-acceso', { from: from, tipo: 'vehiculo' });
    return true;
  }

  return false;
}

function _botHandleInfoQuery(from, contactName, conv, text) {
  const t = (text || '').toLowerCase();

  // Detectamos si el texto trae marcadores de fecha (mes, rango, relativo) →
  // entonces NO es info generica, es consulta con fechas y la dejamos a NLU.
  const hasMonth = /(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)/i.test(t);
  const hasRange = /del\s+\d+\s+al\s+\d+/i.test(t);
  const hasRelative = /\b(finde|fin\s+de\s+sem|este|pr[oó]xim|siguiente|hoy|ma[ñn]ana|pasado\s+ma[ñn]ana|esta\s+semana|semana\s+que\s+viene|mes\s+que\s+viene)\b/i.test(t);
  const isDatesQuery = hasMonth || hasRange || hasRelative;

  // 0) Fotos / imágenes de cabañas → enviar fotos reales en el chat
  if (/\b(fotos?|im[aá]genes?|im[aá]gen|fotograf[ií]a|mu[eé]stra(me)?|ens[eé][ñn]a(me)?|c[oó]mo\s+(es|son|luce|se\s+ve)|conocer\s+la\s+caba)\b/i.test(t)) {
    let cabinKey = _botDetectCabin(t);
    if (!cabinKey && conv.context && conv.context.cabin) cabinKey = conv.context.cabin;
    _botSendCabinPhotos(from, contactName, cabinKey);
    return true;
  }

  // 1) Precio / tarifa sin fechas → muestra tarifas para 2 pax
  const asksPrice = /\b(precios?|tarifas?|costo|valor|costos?)\b/i.test(t) ||
                    /cu[aá]nto\s+(cuesta|sale|vale|valen|cuestan)/i.test(t);
  if (asksPrice && !isDatesQuery) {
    _botSendPricingInfo(from, contactName, conv);
    return true;
  }

  // 2) FAQ topics — respuesta puntual + invitacion a cotizar
  const tail = '\n\n¿Querés cotizar para alguna fecha? Decime *fechas* y *personas* 📅';
  if (/\b(cocina|bbq|cocinar|parrilla|cooler|nevera|comida|alimentos|equipada)\b/i.test(t)) {
    sendWhatsAppText(from,
      '🍳 *Cocina & alimentación*\n\n' +
      'Cocina completamente equipada y área de BBQ. Incluye café, azúcar y especias básicas. ' +
      'Cooler grande disponible (no contamos con nevera) — te recomendamos traer hielo y los alimentos que vayas a usar.' + tail
    );
    return true;
  }
  if (/\b(energ[ií]a|electric|luz|solar|panel|inversor|cargar|se[ñn]al|internet|wifi)\b/i.test(t)) {
    sendWhatsAppText(from,
      '⚡ *Energía & conectividad*\n\n' +
      'Iluminación 100% solar mediante paneles fotovoltaicos. Inversor disponible para cargar celulares y dispositivos. ' +
      'Excelente señal de todas las operadoras.' + tail
    );
    return true;
  }
  if (/check.?in|checkin|check.?out|checkout|\bhora\s+(de\s+)?(llegada|entrada|salida)|a\s+qu[eé]\s+hora/i.test(t)) {
    sendWhatsAppText(from,
      '🕒 *Horarios*\n\n' +
      '• Check-in: *2:00 pm*\n' +
      '• Check-out: *11:00 am*\n\n' +
      'Si necesitás entrar más temprano o salir más tarde lo coordinamos según disponibilidad.' + tail
    );
    return true;
  }
  if (/\b(toalla|jab[oó]n|papel|ba[ñn]o|amenidades|amenities)\b/i.test(t)) {
    sendWhatsAppText(from,
      '🛁 *Baño & comodidades*\n\n' +
      'Jabón, papel higiénico y toallas limpias incluidos. Fumigamos semanalmente — si sos sensible a mosquitos, te recomendamos traer repelente.' + tail
    );
    return true;
  }
  if (/\b(privacidad|vecinos|exclusiv|compartid)\b/i.test(t)) {
    sendWhatsAppText(from,
      '🌿 *Privacidad*\n\n' +
      'Las tres cabañas son independientes y de uso exclusivo de quienes reservan. Sin vecinos compartidos.' + tail
    );
    return true;
  }
  if (/\b(capacidad|cu[aá]ntas?\s+personas|cu[aá]ntos?\s+(hu[eé]spedes|hu[eé]sped|caben))\b/i.test(t)) {
    sendWhatsAppText(from,
      '👥 *Capacidad por cabaña*\n\n' +
      '• *Paseo por Las Nubes* — 2 a 4 personas (cama queen + sofá-cama doble)\n' +
      '• *Portal hacia Las Nubes* — 2 personas (cama matrimonial full)\n' +
      '• *Puente entre Las Nubes* — 2 a 4 personas (cama queen + cama auxiliar)' + tail
    );
    return true;
  }
  if (/\b(mascot|pet|llevar\s+(mi\s+)?perr|mi\s+gat)\b/i.test(t)) {
    sendWhatsAppText(from,
      '🐾 *Mascotas*\n\n' +
      'Por el momento no recibimos mascotas. Para coordinaciones puntuales podés escribir al equipo y vemos. 🙏'
    );
    return true;
  }

  // Decoracion / sorpresas para cumpleaños / aniversarios
  if (/\b(decoraci[oó]n|decorar|decorad|sorpresa|globos?|flores|cumple|cumplea[ñn]os|aniversario|luna\s+de\s+miel|romant|rom[aá]ntic|propon|pedida\s+de\s+mano|propuesta)\b/i.test(t)) {
    sendWhatsAppText(from,
      '🎉 *Decoración especial para aniversarios y cumpleaños*\n\n' +
      'Realizamos una decoración básica que incluye:\n' +
      '• Arreglo de flores\n' +
      '• Globos\n' +
      '• Letreros de cumpleaños o aniversarios\n' +
      '• Elementos decorativos románticos\n' +
      '• Una botella de espumante 🥂\n\n' +
      '*Costo:* $40 adicional a tu reserva.\n\n' +
      'Si querés agregarlo, decímelo al confirmar las fechas y lo coordinamos. ¿Para cuándo sería?'
    );
    return true;
  }

  // Niños / familia
  if (/\b(ni[ñn]o|ni[ñn]a|hijo|hija|hijos|hijas|beb[eé]|infantil|familia|family|kid|menor(es)?\s+de|aptas?\s+para\s+ni[ñn]os|family\s+friendly|chicos)\b/i.test(t)) {
    sendWhatsAppText(from,
      '👨‍👩‍👧 *Familias con niños*\n\n' +
      '¡Las cabañas son ideales para escapadas familiares! 🌿\n\n' +
      '🛏 *Camas* — Paseo y Puente tienen recámara con cama *queen* + cama auxiliar *doble* debajo. Perfectas para 2 adultos y 2 niños.\n\n' +
      '💰 *Política de niños*\n' +
      '• Menores de *5 años* no pagan.\n' +
      '• De la 3ra persona en adelante (5 años o más): $' + BOT_RECARGO_PERSONA_GRANDE + ' por persona/noche en Paseo y Puente, $' + BOT_RECARGO_PERSONA_PORTAL + ' en Portal.\n\n' +
      'Decime *fechas* y *cuántas personas* (incluí la edad de los niños) y te cotizo al instante 📅'
    );
    return true;
  }

  // Clima / temperatura
  if (/\b(clima|hace\s+(calor|fr[ií]o)|temperatura|qu[eé]\s+tiempo|humedad|caluroso|fresco|ventilador|aire\s+acondicionado|\bac\b)\b/i.test(t)) {
    sendWhatsAppText(from,
      '🌤 *Clima*\n\n' +
      'Las Nubes está en zona de montaña, con clima fresco — unos 4°C por debajo de la temperatura de la ciudad y brisa constante todo el día.\n\n' +
      'Cada cabaña tiene un ventilador pequeño, pero rara vez se usa. Te recomendamos traer una chaqueta liviana para las noches 🧥.\n\n' +
      '¿Querés cotizar para alguna fecha? Decime *fechas* y *personas* 📅'
    );
    return true;
  }

  // Mosquitos / repelente
  if (/\b(mosquit|zancud|insect|bicho|repelente|off\s+spray)\b/i.test(t)) {
    sendWhatsAppText(from,
      '🦟 *Mosquitos*\n\n' +
      'Fumigamos con frecuencia, así que hay muy pocos. La única franja en que molestan un poco es entre las *6 y 7 pm*.\n\n' +
      'Tenemos repelente disponible en la *Tiendita Las Nubes*: OFF spray ($8) o toallitas Family Care ($5).'
    );
    return true;
  }

  // Jacuzzi / piscina / sauna → no tenemos, pero tenemos cascada propia + alternativas
  if (/\b(jacuzzi|piscina|pool|sauna|hot\s+tub|tina\s+caliente)\b/i.test(t)) {
    sendWhatsAppText(from,
      '🌊 No tenemos jacuzzi ni piscina en las cabañas, pero algo mejor: una *cascada propia dentro del proyecto*, en plena naturaleza y exclusiva para nuestros huéspedes. 🌿\n\n' +
      'Y en los alrededores hay opciones increíbles:\n' +
      '• *Los Cajones de Chamé* (~10 min) — piscinas naturales y saltos al agua\n' +
      '• *Cascadas Filipinas* (Sorá, ~20 min) — 7 cascadas encadenadas\n' +
      '• *Cascada Manglarito* (Sorá, ~20 min) — 35m de caída\n' +
      '• *Playas* Coronado y Gorgona (~15-20 min)\n\n' +
      '¿Querés cotizar para alguna fecha? Decime *fechas* y *personas* 📅'
    );
    return true;
  }

  return false;
}

// Decide si vale la pena gastar una llamada a Claude para responder al
// mensaje. Skipea saludos triviales y mensajes muy cortos.
function _botShouldUseSmartFallback(text, conv) {
  const t = (text || '').trim().toLowerCase();
  if (t.length < 6) return false;
  if (/^(hola|hi|hey|buenas|buenos\s+d[ií]as|buenas\s+tardes|buenas\s+noches|gracias|grax|ok|okay|listo|s[ií]|no|claro|perfecto|dale|bien)[!.,\s]*$/i.test(t)) return false;
  return true;
}

// Fallback inteligente: pasa el mensaje del cliente a Claude con todo el
// contexto del negocio en system prompt y devuelve respuesta conversacional.
// Solo corre si ningun handler regex/state-machine matcheo.
function _botSmartFallback(from, contactName, conv, text) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) return false;

  // System prompt base + bloques de conocimiento especifico por topic detectado
  // en el mensaje (actividades, gastronomia, insumos, como llegar).
  const topicCtx = _botKnowledgeTopics(text);
  const systemPrompt = _botKnowledgeBase() + (topicCtx ? '\n\n' + topicCtx : '');

  try {
    const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 450,
        system: systemPrompt,
        messages: [{ role: 'user', content: text }]
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(res.getContentText());
    const reply = data.content && data.content[0] && data.content[0].text;
    if (!reply) {
      logDebugEntry('smart-fallback-NO-REPLY', { text: text.slice(0, 120), raw: JSON.stringify(data).slice(0, 200) });
      return false;
    }
    sendWhatsAppText(from, reply.trim());
    logDebugEntry('smart-fallback-OK', {
      from: from, text: text.slice(0, 120), reply: reply.slice(0, 200),
      inputTokens: data.usage && data.usage.input_tokens, outputTokens: data.usage && data.usage.output_tokens
    });
    _saveConv(from, 'INITIAL', conv.context || {}, contactName);
    return true;
  } catch(err) {
    logDebugEntry('smart-fallback-FAIL', { error: err.message, text: text.slice(0, 120) });
    return false;
  }
}

// System prompt grounded en toda la info del negocio. Cualquier ajuste de
// politica / precios / amenidades se hace aca para que Claude responda
// consistente con los handlers regex.
function _botKnowledgeBase() {
  return (
'Sos el asistente conversacional de *Las Nubes*, un refugio de tres cabañas privadas en las faldas del Cerro Chicá, Panamá. Respondes mensajes de clientes por WhatsApp.\n\n' +
'## CABAÑAS\n' +
'- *Paseo por Las Nubes*: 2-4 personas (cama queen + sofá-cama doble)\n' +
'- *Portal hacia Las Nubes*: 2 personas (cama matrimonial full)\n' +
'- *Puente entre Las Nubes*: 2-4 personas (cama queen + cama auxiliar)\n' +
'Las tres son independientes, privadas y de uso exclusivo de quienes reservan.\n\n' +
'## TARIFAS POR NOCHE (2 personas)\n' +
'- Domingo a jueves: $' + BOT_RATE_WEEKDAY + '/noche (mismo precio en las 3 cabañas)\n' +
'- Viernes y sábado: $' + BOT_RATE_WEEKEND + '/noche\n' +
'- 3-4 personas: recargo pequeño por persona adicional ($' + BOT_RECARGO_PERSONA_PORTAL + ' en Portal, $' + BOT_RECARGO_PERSONA_GRANDE + ' en Paseo/Puente)\n' +
'- Niños menores de 5 años NO pagan\n' +
'- 5-6 personas: combo Puente + Portal (cabañas contiguas), cotización aparte\n\n' +
'## FAMILIAS CON NIÑOS\n' +
'- Paseo y Puente: recámara con cama queen + cama auxiliar doble debajo (ideal para 2 adultos + 2 niños)\n' +
'- Portal: cama matrimonial full (para 2 personas, no incluye espacio para niños extra)\n' +
'- Niños menores de 5 años no pagan\n' +
'- A partir de los 5 años aplican como persona adicional con recargo normal\n\n' +
'## AMENIDADES\n' +
'- *Cascada propia dentro del proyecto*, exclusiva para huéspedes (en plena naturaleza)\n' +
'- Cocina completa + área de BBQ\n' +
'- Cooler grande (NO hay nevera — traer hielo)\n' +
'- Café, azúcar y especias básicas incluidos\n' +
'- Menú simple de comida disponible bajo reserva previa\n' +
'- Iluminación 100% solar + inversor para cargar celulares\n' +
'- Toallas, jabón y papel higiénico incluidos\n' +
'- Excelente señal de todas las operadoras\n' +
'- Fumigamos semanalmente (traer repelente si sos sensible a mosquitos)\n' +
'- NO hay jacuzzi, piscina ni sauna en las cabañas\n\n' +
'## SERVICIOS EXTRAS\n' +
'- *Decoración para aniversarios y cumpleaños* — $40 adicionales. Incluye arreglo de flores, globos, letreros, elementos decorativos románticos y una botella de espumante. Se coordina al confirmar la reserva.\n\n' +
'## CLIMA\n' +
'- Zona de montaña, clima fresco — aprox 4°C por debajo de la temperatura de Ciudad de Panamá\n' +
'- Brisa constante todo el día\n' +
'- Cada cabaña tiene ventilador pequeño (rara vez necesario)\n' +
'- Recomendación: traer chaqueta liviana para las noches\n\n' +
'## MOSQUITOS\n' +
'- Hay muy pocos porque fumigamos con frecuencia\n' +
'- Solo molestan un poco entre las 6 y 7 pm\n' +
'- Repelente disponible en la Tiendita Las Nubes: OFF spray ($8) o toallitas Family Care ($5)\n\n' +
'## UBICACIÓN\n' +
'- Buenos Aires, Chamé · faldas del Cerro Chicá\n' +
'- A 1h 15min de Ciudad de Panamá\n' +
'- Naturaleza, vistas y privacidad total\n\n' +
'## HORARIOS\n' +
'- Check-in: 2:00 pm · Check-out: 11:00 am\n' +
'- Entradas anticipadas / salidas tardías se coordinan según disponibilidad\n\n' +
'## PAGO\n' +
'- Yappy y ACH (transferencia bancaria)\n' +
'- NO aceptamos tarjeta de crédito ni pago contraentrega\n' +
'- Sin voucher/abono, no se confirma la reserva\n\n' +
'## POLÍTICAS\n' +
'- No se reciben mascotas por el momento\n' +
'- Cancelaciones/cambios: coordinar con el equipo (derivá al agente)\n' +
'- Eventos especiales (cumpleaños, aniversarios, lunas de miel): bienvenidos, coordiná con el equipo\n\n' +
'## QUÉ HACER Y QUÉ NO\n' +
'- NO inventes datos, precios ni promos que no estén acá\n' +
'- NO confirmes disponibilidad de fechas específicas — eso requiere consultar sistema, pedí *fechas* y *personas* y avisá que cotizás al instante\n' +
'- Si la pregunta es ambigua o requiere coordinación humana (eventos grandes, cambios de fecha, cancelaciones, descuentos), invitá a tocar *Hablar con un agente* en el menú o escribir "3"\n' +
'- NO ofrezcas descuentos\n\n' +
'## TONO\n' +
'- Cálido, breve, conversacional. Español neutral latinoamericano (usá "tú", no "vos")\n' +
'- Usa *negritas* de WhatsApp para énfasis y emojis con moderación (🌿 🏡 ☕ 🛏 📍)\n' +
'- Máximo 4-5 líneas. Directo al punto\n' +
'- Si la consulta es sobre disponibilidad/precios, terminá con: "Decime *fechas* y *personas* (ej: _\'del 5 al 8 de junio, 2 personas\'_) y te cotizo al instante 🤝"\n' +
'- Si pueden ver fotos: invitá a ver el catálogo en https://lasnubes.cloud\n\n' +
'Respondé directo al mensaje del cliente, sin saludos largos.'
  );
}

// Detecta topics relevantes en el mensaje y devuelve bloques de knowledge
// especifico para apendear al system prompt. Mantenemos los bloques cortos
// (sin inundar el prompt) y solo se agregan los que matchean.
function _botKnowledgeTopics(text) {
  const t = (text || '').toLowerCase();
  const blocks = [];
  if (/\b(actividad|qu[eé]\s+(hacer|hay)|que\s+(hacer|hay)|cerca|excursi[oó]n|pasear|caminar|cascada|playa|tour|senderismo|aventura|cajones|parque|naturaleza|cerro|chic[aá]|coronado|sor[aá]|filipinas|manglarito|gorgona|campana|hike|trail|nadar|surf|mirador)\b/i.test(t)) {
    blocks.push(_botKbActividades());
  }
  if (/\b(gastronom|restaurant|comer|comida\b|cenar|almorzar|desayun|d[oó]nde\s+comer|pizza|sushi|carne|parrilla|bar\b|caf[eé]|peruano|italiano|asi[aá]tico|fusi[oó]n|pollo|ceviche|menu|men[uú])\b/i.test(t)) {
    blocks.push(_botKbGastronomia());
  }
  if (/\b(insumo|tienda|super|supermerc|comprar|provee|provisi[oó]n|aproviv|aprovision|hielo|carb[oó]n|repelente|farmacia|botiqu[ií]n|leña|le[ñn]a|fogata|malvaviscos)\b/i.test(t)) {
    blocks.push(_botKbInsumos());
  }
  if (/(c[oó]mo\s+llegar|como\s+llegar|llegar|ubicaci|d[oó]nde\s+(est[aá]n?|queda)|direcci[oó]n|waze|maps|gps|carretera|interameric|panamericana)/i.test(t)) {
    blocks.push(_botKbComoLlegar());
  }
  return blocks.join('\n\n');
}

function _botKbActividades() {
  return (
'## ACTIVIDADES Y ALREDEDORES (más detalles: https://lasnubes.cloud)\n' +
'En el mismo proyecto y cerca:\n' +
'- *Cascada propia* dentro del proyecto Las Nubes — exclusiva para huéspedes, sin tener que salir.\n' +
'- *Los Cajones de Chamé* (~10 min): cañón con piscinas naturales conectadas y saltos al agua. Snorkel, natación. 4x4 recomendado. Estacionamiento $3. Mejor en temporada seca.\n' +
'- *Parque Nacional Altos de Campana* (~20 min): primer parque nacional de Panamá. Senderismo al Cerro de la Cruz (905m, ~2h ida y vuelta) con vistas al Pacífico y Canal. Entrada ~$5.\n' +
'- *Coronado* (~20 min): playa de arena oscura, surf, supermercados y restaurantes. Mejor entre semana.\n' +
'- *Playa Gorgona* (~15 min): playa tranquila, atardeceres, chiringuitos de mariscos.\n' +
'- *Cascadas Filipinas* (Sorá, ~20 min): sistema de 7 cascadas encadenadas, la segunda de 15m. Nivel medio-alto. Entrada ~$3. 4x4 + jeeps locales disponibles.\n' +
'- *Cascada Manglarito* (Sorá, ~20 min): cascada de 35+m con salto al agua. Mejor con guía local.\n' +
'- *Cascada Nativa* (Sorá, ~20 min): propiedad privada familiar, caminata corta. Accesible. Entrada ~$3.\n' +
'- *Senderismo* por las faldas del Cerro Chicá, vistas panorámicas y fotografía de naturaleza.'
  );
}

function _botKbGastronomia() {
  return (
'## GASTRONOMÍA CERCANA (lista completa con mapas: https://lasnubes.cloud)\n' +
'- *Buenas Pizzas de Sorá* (~20 min): pizzería artesanal con horno de leña, ambiente familiar.\n' +
'- *Pío Pío de Bejuco* (~10 min): pollo asado a la leña — punto rápido al llegar.\n' +
'- *Nación Sushi* (Coronado, ~25 min): sushi fresco, opción de delivery.\n' +
'- *Slabón* (Coronado): restaurante-bar con ambiente animado, ideal para grupos.\n' +
'- *Don Chacho Grill* (Coronado): carnes a la parrilla, familiar.\n' +
'- *Luna Rossa* (Coronado): italiano con pastas y pizzas, cenas en pareja o grupo pequeño.\n' +
'- *Las Bóvedas Fusión* (Coronado): cocina de fusión bien presentada.\n' +
'- *Don Lee* (Coronado): asiático con precios amigables, opción para llevar.\n' +
'- *Nazca 21* (Coronado): peruano con ceviches y tiraditos, ideal para cena especial.\n' +
'- *Coronado zona gastronómica*: muchas opciones — cadenas, cafés de especialidad, bares.'
  );
}

function _botKbInsumos() {
  return (
'## INSUMOS, TIENDAS Y COMPRAS (detalles: https://lasnubes.cloud)\n' +
'- *Tienda de Conveniencia* (a 5 min de las cabañas): hielo, carbón, especias, bebidas — lo básico.\n' +
'- *MiniSuper Buenos Precios* (Bejuco, ~15 min): surtido para aprovisionarse antes de subir.\n' +
'- *Supermercados en Coronado* (~20 min): El Rey, Machetazo, Riba Smith (premium), Super 99 — opciones completas.\n' +
'- *Tiendita Las Nubes* (entrega directo a la cabaña, pago Yappy):\n' +
'  · 🔥 Kit de Fogata (leña + cerillo + palillos + 8 malvaviscos): $10\n' +
'  · 🪨 Bolsa de carbón con cerillo: $5\n' +
'  · 🦟 Repelente OFF spray: $8\n' +
'  · 🧻 Repelente Family Care (toallitas): $5\n' +
'  · 🪥 Kit pasta + cepillo: $5\n' +
'  · 🌸 Toallas sanitarias: $5'
  );
}

function _botKbComoLlegar() {
  return (
'## UBICACIÓN Y CÓMO LLEGAR\n' +
'- Buenos Aires, Chamé · faldas del Cerro Chicá, Panamá.\n' +
'- Desde Ciudad de Panamá: aproximadamente *1h 15min* por la Interamericana.\n' +
'- Indicaciones detalladas, Waze y Google Maps con coordenadas exactas se comparten al confirmar la reserva.\n' +
'- *Acceso en auto*: carretera de asfalto en muy buen estado hasta la entrada del proyecto, luego ~5 min en camino de tosca fina. Se reciben sedanes y todo tipo de autos sin problema (NO hace falta 4x4).\n' +
'- *Sin auto*: desde la terminal de Albrook tomar cualquier bus hacia el interior y bajarse en el Pío Pío de Bejuco; ahí los recogemos y trasladamos a la cabaña ida y vuelta por $20 adicionales en la reserva.'
  );
}

// Keywords que indican cambio/cancelacion → no cotizar, derivar a humano
function _isReservaChangeRequest(text) {
  const t = (text || '').toLowerCase();
  return /cambiar fecha|cambio de fecha|cambio de reserva|cambiar reserva|reagenda|reprograma|cancelar reserva|cancelaci[oó]n|cancelar mi|no podr[eé] ir|no voy a poder|no podemos ir|no vamos a poder|posponer|adelantar mi/.test(t);
}

// ─── NLU con Claude ────────────────────────────────────────────────
function _parseDatesWithClaude(text, today) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) return null;
  const prompt =
    'Hoy es ' + today + ' (timezone America/Panama). Un cliente escribio en espanol:\n\n"' +
    text.replace(/"/g, '\\"') + '"\n\n' +
    'Extrae las fechas de reserva (checkin/checkout), total de personas y niños menores de 5 años (que no pagan). Devuelve SOLO JSON con esta forma exacta:\n' +
    '{"checkin":"YYYY-MM-DD"|null,"checkout":"YYYY-MM-DD"|null,"persons":N|null,"freeChildren":N|null,"confidence":0-1}\n\n' +
    'Reglas:\n' +
    '- checkin = dia que llegan\n' +
    '- checkout = dia que se van (mayor a checkin)\n' +
    '- Si solo mencionan 1 fecha y "N noches", calcular checkout = checkin + N\n' +
    '- Si solo mencionan 1 fecha sin noches, asumir 1 noche y checkout = checkin + 1\n' +
    '- persons = TOTAL de personas (adultos + niños de todas las edades, los menores TAMBIÉN cuentan acá)\n' +
    '- freeChildren = SOLO niños menores de 5 años (bebés, infantes). Si no se menciona edad, dejá en 0\n' +
    '- Ejemplos de freeChildren:\n' +
    '   "2 adultos y 1 bebé" → persons=3, freeChildren=1\n' +
    '   "somos 4: 2 grandes, niños de 3 y 7" → persons=4, freeChildren=1 (solo el de 3 es menor de 5)\n' +
    '   "2 adultos y 2 niños" (sin edad) → persons=4, freeChildren=0 (asumimos que pagan)\n' +
    '   "4 personas" → persons=4, freeChildren=0\n' +
    '   "mi esposa y yo + nuestra hija de 2 años" → persons=3, freeChildren=1\n' +
    '- persons = null si no se menciona\n' +
    '- confidence 0 a 1: 1 = muy claro, 0 = ambiguo\n' +
    '- Si no podes inferir fechas con confianza > 0.4, devuelve {"checkin":null,"checkout":null,"persons":null,"freeChildren":null,"confidence":0}\n' +
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

// Combo Puente + Portal contiguas (5-6 personas). Recargo por noche.
function _botPrecioComboTotal(checkin, checkout, personas) {
  let base = 0;
  const start = new Date(checkin + 'T12:00:00');
  const end   = new Date(checkout + 'T12:00:00');
  let nights = 0;
  for (let cur = new Date(start); cur < end; cur.setDate(cur.getDate() + 1)) {
    base += _botPrecioPorNoche(Utilities.formatDate(cur, BOT_TZ, 'yyyy-MM-dd'));
    nights++;
  }
  const recargo = (personas >= 6) ? BOT_RECARGO_COMBO_6 : BOT_RECARGO_COMBO_5;
  return base + recargo * nights;
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

// ─── Mensaje pre-rellenado del calendario publico ─────────────────
// Detecta el texto que envia el cliente al tocar "Reservar" en
// index.html. Formato (ver _buildClientReservaMessage en index.html):
//   Hola! Deseo reservar:
//
//   *Fecha:* miércoles 27 de mayo · 1 noche
//   *Cabaña:* Paseo por Las Nubes
//   *Personas:* 2
//   *Total:* $90.00
//   ... (secciones informativas) ...
//   Quedo atento a las formas de pago. ¡Gracias!
function _parseClientReservaMessage(text) {
  if (!text) return null;
  if (!/Deseo reservar/i.test(text)) return null;
  if (!/\*Fecha:\*/.test(text) || !/\*Caba.a:\*/.test(text)) return null;

  const cabinMatch = text.match(/\*Caba.a:\*\s*([^\n]+)/);
  const persMatch  = text.match(/\*Personas:\*\s*(\d+)/);
  const fechaMatch = text.match(/\*Fecha:\*\s*([^\n]+)/);
  const totalMatch = text.match(/\*Total:\*\s*\$?([\d,.]+)/);
  if (!cabinMatch || !persMatch || !fechaMatch) return null;

  const cabinStr  = cabinMatch[1];
  let cabin = null, isCombo = false;
  if (/combo/i.test(cabinStr)) isCombo = true;
  else if (/paseo/i.test(cabinStr))  cabin = 'verde';
  else if (/portal/i.test(cabinStr)) cabin = 'azul';
  else if (/puente/i.test(cabinStr)) cabin = 'lila';

  const fechaLine = fechaMatch[1].trim();
  const isPasadia = /pasad[íi]a|pasatarde/i.test(fechaLine);

  return {
    cabin: cabin,
    isCombo: isCombo,
    isPasadia: isPasadia,
    personas: parseInt(persMatch[1], 10),
    fechaLine: fechaLine,
    total: totalMatch ? parseFloat(totalMatch[1].replace(/,/g, '')) : null
  };
}

// Maneja el mensaje pre-rellenado del calendario publico.
// - Combo / pasadia / 5+ personas → handoff al equipo (no se cotiza auto).
// - Noche + cabana single + disponible → directo a OFFERING_PAYMENT.
// - No disponible → flujo normal de disponibilidad con alternativas.
function _botHandleClientReservaMessage(from, contactName, conv, parsed) {
  logDebugEntry('bot-client-reserva-msg', {
    from: from, cabin: parsed.cabin, isCombo: parsed.isCombo,
    isPasadia: parsed.isPasadia, personas: parsed.personas, fechaLine: parsed.fechaLine
  });

  // Combo / pasadia / 5+ personas → handoff
  if (parsed.isCombo || parsed.isPasadia || parsed.personas >= 5) {
    sendWhatsAppText(from,
      '🌿 ¡Recibí tu solicitud! Para este tipo de reserva coordinamos directo con vos. ' +
      'Tocá el botón abajo para escribir al equipo.'
    );
    try {
      sendWhatsAppCTAUrl(from,
        'Te pasamos métodos de pago en un mensaje.',
        '💬 Abrir WhatsApp',
        'https://wa.me/50769812266?text=' + encodeURIComponent('Hola, quiero reservar (vengo del calendario web).')
      );
    } catch(_) {}
    try {
      sendWhatsAppText(BOT_ADMIN_PHONE,
        '⚠️ Solicitud especial vía calendario web:\n' +
        '👤 ' + (contactName || from) + '\n' +
        '📱 +' + from + '\n' +
        '📝 ' + parsed.fechaLine + '\n' +
        '🏡 ' + (parsed.isCombo ? 'Combo' : (parsed.cabin || '?')) + '\n' +
        '👥 ' + parsed.personas
      );
    } catch(_) {}
    _saveConv(from, 'HUMAN_HANDOFF', conv.context || {}, contactName);
    return;
  }

  // Parsear fechas del fechaLine usando NLU existente
  const datesParsed = _parseDatesWithClaude(parsed.fechaLine, _botToday());
  if (!datesParsed || !datesParsed.checkin || !datesParsed.checkout) {
    sendWhatsAppText(from,
      '🌿 ¡Recibí tu solicitud! Para confirmar disponibilidad, recordame las *fechas exactas* ' +
      '(ej: _"del 5 al 8 de junio"_).'
    );
    _saveConv(from, 'AWAITING_DATES', { personas: parsed.personas }, contactName);
    return;
  }

  // Verificar disponibilidad de la cabana solicitada
  const avail = _botCheckAvailability(datesParsed.checkin, datesParsed.checkout);
  if (!parsed.cabin || !avail[parsed.cabin]) {
    // No disponible: caer al flujo normal de disponibilidad (muestra opciones libres + alternativas)
    return _replyAvailability(from, contactName, conv, datesParsed.checkin, datesParsed.checkout, parsed.personas);
  }

  // Disponible: jumpear directo a OFFERING_PAYMENT con formas de pago
  const newCtx = Object.assign({}, conv.context || {}, {
    dates: { checkin: datesParsed.checkin, checkout: datesParsed.checkout },
    personas: parsed.personas
  });
  const fakeConv = { step: 'INITIAL', context: newCtx, name: contactName };
  return _botStartBooking(from, contactName, fakeConv, parsed.cabin);
}

// ─── Main handler ──────────────────────────────────────────────────
function botHandleMessage(from, text, contactName, kind) {
  const conv = _getConv(from) || { step: 'INITIAL', context: {}, name: contactName || '' };

  // Mensaje pre-rellenado desde el calendario publico (index.html → btn "Reservar").
  // Estructura: "Hola! Deseo reservar:" + *Fecha:* + *Cabaña:* + *Personas:* + *Total:*
  // Si lo detectamos, saltamos directo a OFFERING_PAYMENT con las formas de pago.
  if (kind === 'text') {
    const reservaMsg = _parseClientReservaMessage(text);
    if (reservaMsg) {
      return _botHandleClientReservaMessage(from, contactName, conv, reservaMsg);
    }
  }

  // Lead tipico de campana en Instagram/Facebook ("Hola! Quiero más información").
  // Solo lo disparamos si es el primer mensaje (step INITIAL) para no
  // sobreescribir conversaciones en curso.
  if (kind === 'text' && conv.step === 'INITIAL' && _isCampaignInquiry(text)) {
    return _botSendCampaignWelcome(from, contactName);
  }

  // Boton "Ver otras fechas" → vuelve a AWAITING_DATES
  if (kind === 'button_reply' && text === 'try_dates') {
    sendWhatsAppText(from, '🌿 Decime las nuevas fechas:\n\n• "del 5 al 8 de junio, 2 personas"\n• "viernes a domingo, 4 personas"');
    _saveConv(from, 'AWAITING_DATES', conv.context, contactName);
    return;
  }

  // Cambiar cantidad de personas → reconsulta disponibilidad con nuevas N
  if ((kind === 'list_reply' || kind === 'button_reply') && /^persons_(\d)$/.test(text)) {
    const n = parseInt(text.replace('persons_', ''), 10);
    const dates = conv.context && conv.context.dates;
    if (!dates || !dates.checkin || !dates.checkout) {
      sendWhatsAppText(from, '🤔 Perdí el contexto. Decime las fechas otra vez (ej: "del 5 al 8 de junio").');
      _saveConv(from, 'AWAITING_DATES', {}, contactName);
      return;
    }
    return _replyAvailability(from, contactName, { context: conv.context, name: contactName }, dates.checkin, dates.checkout, n);
  }

  // Boton de seleccion de cabana → empieza booking flow
  if (kind === 'button_reply' && /^pick_(verde|azul|lila)$/.test(text)) {
    const elegida = text.split('_')[1];
    return _botStartBooking(from, contactName, conv, elegida);
  }
  // Tambien aceptar list_reply para pick_X
  if (kind === 'list_reply' && /^pick_(verde|azul|lila)$/.test(text)) {
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

  // Boton "Ya me retiré" de la plantilla de check-out → avisar admin + Erika
  if (kind === 'button_reply' && text.indexOf('checkout_') === 0) {
    return _botHandleCheckoutDone(from, contactName, text.replace('checkout_', ''));
  }

  // Boton "Consultas y cambios" de la plantilla de confirmación → CTA a Josh
  if (kind === 'button_reply' && text.indexOf('consulta_') === 0) {
    return _botHandleConsultaReserva(from, contactName, text.replace('consulta_', ''));
  }

  // Menu list/button reply → handler especifico
  if ((kind === 'list_reply' || kind === 'button_reply') && /^menu_/.test(text)) {
    if (text === 'menu_disponibilidad') {
      sendWhatsAppText(from,
        '¡Genial! 🌿 Decime *fechas* y *personas* (ej: _"del 5 al 8 de junio, 2 personas"_).'
      );
      _saveConv(from, 'AWAITING_DATES', conv.context, contactName);
      return;
    }
    if (text === 'menu_he_llegado')   { _botMenuHeLlegado(from, contactName, conv); return; }
    if (text === 'menu_como_llegar')  { _botMenuComoLlegar(from);  _botSendMainMenu(from, contactName, false); _saveConv(from, 'SHOWED_INFO', conv.context, contactName); return; }
    if (text === 'menu_actividades')  { _botMenuActividades(from); _botSendMainMenu(from, contactName, false); _saveConv(from, 'SHOWED_INFO', conv.context, contactName); return; }
    if (text === 'menu_gastronomia')  { _botMenuGastronomia(from); _botSendMainMenu(from, contactName, false); _saveConv(from, 'SHOWED_INFO', conv.context, contactName); return; }
    if (text === 'menu_insumos')      { _botMenuInsumos(from);     _botSendMainMenu(from, contactName, false); _saveConv(from, 'SHOWED_INFO', conv.context, contactName); return; }
    if (text === 'menu_tienda')       { _botMenuTienda(from);      _botSendMainMenu(from, contactName, false); _saveConv(from, 'SHOWED_INFO', conv.context, contactName); return; }
    if (text === 'menu_faq')          { _botMenuFAQ(from);         _botSendMainMenu(from, contactName, false); _saveConv(from, 'SHOWED_INFO', conv.context, contactName); return; }
    if (text === 'menu_agente' || text === 'menu_humano') {
      try {
        sendWhatsAppCTAUrl(from,
          '🙋 ¡Claro! Tocá el botón abajo para escribirle directo a una persona de nuestro equipo por WhatsApp.',
          'Abrir WhatsApp',
          'https://wa.me/50769812266?text=' + encodeURIComponent('Hola, vengo del asistente de Las Nubes 🌿')
        );
      } catch(err) {
        sendWhatsAppText(from, '🙋 Escribinos directo aquí:\nhttps://wa.me/50769812266');
      }
      _saveConv(from, 'HUMAN_HANDOFF', conv.context, contactName);
      try { sendWhatsAppText('50769812266', '🔔 El Agente derivó a un cliente (menú): ' + (contactName || from) + ' (' + from + ')'); } catch(_) {}
      return;
    }
  }

  // Boton "Cancelar" en OFFERING_PAYMENT — libera el estado + muestra menu principal
  if ((kind === 'button_reply' && text === 'cancel_booking') ||
      (conv.step === 'OFFERING_PAYMENT' && /^(cancela|cancelar|atras|atrás)\b/i.test((text || '').trim()))) {
    sendWhatsAppText(from, '👋 Listo, cancelamos esta reserva. ¿En qué más te ayudamos?');
    _saveConv(from, 'INITIAL', {}, contactName);
    _botSendMainMenu(from, contactName, true);
    return;
  }

  // Boton "Sugerencia: usar esta fecha"
  if (kind === 'button_reply' && text.indexOf('alt_') === 0) {
    const newCheckin   = text.replace('alt_', '');
    const prevDates    = conv.context && conv.context.dates;
    const personas     = (conv.context && conv.context.personas) || 2;
    const freeChildren = (conv.context && conv.context.freeChildren) || 0;
    const nights       = prevDates
      ? Math.round((new Date(prevDates.checkout + 'T12:00:00') - new Date(prevDates.checkin + 'T12:00:00')) / 86400000)
      : 1;
    const newCheckout = _botAddDaysISO(newCheckin, nights);
    return _replyAvailability(from, contactName, conv, newCheckin, newCheckout, personas, freeChildren);
  }

  // Mensaje de imagen → voucher (solo si esta en OFFERING_PAYMENT)
  if (kind === 'image') {
    return _botHandleVoucherImage(from, text, contactName, conv);
  }

  // He llegado: esperando nombre del titular para ubicar reserva
  if (conv.step === 'AWAITING_ARRIVAL_NAME') {
    const tName = (text || '').trim();
    if (tName.length < 3) {
      sendWhatsAppText(from, '🤔 Necesito el nombre completo para ubicar la reserva. Probá de nuevo o escribime "agente" para hablar con una persona.');
      return;
    }
    const reservaByName = _botFindReservaByName(tName);
    if (!reservaByName) {
      sendWhatsAppText(from,
        '😔 No encuentro una reserva activa a nombre de *' + tName + '* para hoy.\n\n' +
        'Te derivo con una persona del equipo para resolverlo. Escribime "agente" si querés contactarla directo.'
      );
      try {
        sendWhatsAppText('50769812266',
          '⚠️ Cliente intentó "He llegado" sin match:\n' +
          '📱 +' + from + '\n' +
          '👤 ' + (contactName || '?') + '\n' +
          'Dijo nombre: "' + tName + '"\n\n' +
          'Verificar manualmente.'
        );
      } catch(_) {}
      _saveConv(from, 'INITIAL', {}, contactName);
      return;
    }
    // Encontrada: guardar el telefono en la reserva para futuras consultas
    try {
      const sheet = getOrCreateSheet();
      sheet.getRange(reservaByName.row, 24).setValue(_safeCell(from));
      logDebugEntry('bot-arrival-phone-update', { reservaId: reservaByName.id, telefono: from });
    } catch(updateErr) {
      logDebugEntry('bot-arrival-phone-update-FAIL', { error: updateErr.message });
    }
    return _botSendArrivalInstructions(from, contactName, conv, reservaByName);
  }

  // Email step
  if (conv.step === 'AWAITING_EMAIL') {
    const email = (text || '').trim().toLowerCase();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      sendWhatsAppText(from, '🤔 No parece un email válido. Por favor envíame algo como: nombre@gmail.com');
      return;
    }
    const newCtx = Object.assign({}, conv.context, { email: email });
    // Si ya tenemos el nombre (del voucher OCR), saltar AWAITING_NAME y crear directamente
    if (newCtx.name) {
      return _botCreatePreReservation(from, contactName, newCtx);
    }
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

  // Handoff a humano (prioritario) → CTA URL para abrir WhatsApp del equipo
  if (_isHumanRequest(text) || text.trim() === '3') {
    return botHandleMessage(from, 'menu_agente', contactName, 'list_reply');
  }

  // Cambio de fecha / cancelacion / "no podre ir" → handoff directo al admin
  // (no intenta cotizar aunque haya fechas en el texto).
  if (_isReservaChangeRequest(text)) {
    sendWhatsAppText(from,
      '🙏 Entiendo, te derivo con una persona del equipo para coordinar el cambio o cancelación.\n\n' +
      'Tocá el botón abajo para escribirle directo.'
    );
    try {
      sendWhatsAppCTAUrl(from,
        'Te resolvemos el cambio en un mensaje.',
        '💬 Abrir WhatsApp',
        'https://wa.me/50769812266?text=' + encodeURIComponent('Hola, quiero coordinar un cambio o cancelación de mi reserva.')
      );
    } catch(_) {}
    try {
      sendWhatsAppText('50769812266',
        '⚠️ Cambio/cancelación pedido vía Agente:\n👤 ' + (contactName || from) + '\n📱 +' + from + '\nMensaje: "' + (text || '').slice(0, 200) + '"');
    } catch(_) {}
    _saveConv(from, 'HUMAN_HANDOFF', conv.context || {}, contactName);
    return;
  }

  const t = (text || '').toLowerCase().trim();

  // Acceso: tipo de auto (4x4/sedán) o sin auto (bus/Albrook). Va ANTES del
  // keyword genérico "cómo llegar" para que no lo intercepte el menú de mapas.
  if (_botHandleAccesoQuery(from, contactName, t)) return;

  // Opciones por keywords (compatibilidad: clientes que escriben en vez de tocar)
  if (t === '2' || t.includes('como llegar') || t.includes('cómo llegar') || t.includes('ubicacion') || t.includes('ubicación') || t.includes('direccion') || t.includes('dirección') || t.includes('llegar')) {
    return botHandleMessage(from, 'menu_como_llegar', contactName, 'list_reply');
  }
  if (t.includes('actividad') || t.includes('cascada') || t.includes('playa') || t.includes('que hacer') || t.includes('qué hacer')) {
    return botHandleMessage(from, 'menu_actividades', contactName, 'list_reply');
  }
  if (t.includes('gastrono') || t.includes('restaurant') || t.includes('comer') || t.includes('comida cerca')) {
    return botHandleMessage(from, 'menu_gastronomia', contactName, 'list_reply');
  }
  if (t.includes('he llegado') || t.includes('ya llegue') || t.includes('ya llegué') || t.includes('llegamos') || t.includes('estoy en el porton') || t.includes('estoy en el portón') || t.includes('abrir porton') || t.includes('abrir portón')) {
    return botHandleMessage(from, 'menu_he_llegado', contactName, 'list_reply');
  }
  if (t.includes('hielo') || t.includes('carbon') || t.includes('carbón') || t.includes('tienda cercana') || t.includes('tienda de conv')) {
    return botHandleMessage(from, 'menu_tienda', contactName, 'list_reply');
  }
  if (t.includes('insumo') || t.includes('tiendita') || t.includes('supermercado') || t.includes('compr')) {
    return botHandleMessage(from, 'menu_insumos', contactName, 'list_reply');
  }
  if (t === 'faq' || t.includes('pregunta') || t.includes('duda')) {
    return botHandleMessage(from, 'menu_faq', contactName, 'list_reply');
  }

  // Consultas vagas tipo "para julio", "segunda semana de agosto", "el mes
  // que viene" → mandamos al calendario publico en vez de intentar cotizar.
  if (_looksLikeVagueDateQuery(text)) {
    _botSendCalendarLink(from, contactName);
    _saveConv(from, 'INITIAL', conv.context || {}, contactName);
    return;
  }

  // Info generica (tarifas sin fechas, FAQ topics) → respuesta puntual.
  if (kind === 'text' && _botHandleInfoQuery(from, contactName, conv, text)) {
    return;
  }

  // Date parsing TIENE PRIORIDAD sobre el keyword "disponibilidad".
  // Solo intentamos parsear si el TEXTO tiene pinta de fechas (evita
  // mostrar "no entendi fechas" ante saludos genericos como "Hola").
  // Si el step era AWAITING_DATES pero el cliente cambio de tema (no
  // pinta de fechas), caemos al fallback de menu de bienvenida.
  const midBooking = ['OFFERING_PAYMENT', 'AWAITING_VOUCHER_RETRY', 'AWAITING_EMAIL', 'AWAITING_NAME'].indexOf(conv.step) !== -1;
  if (_looksLikeDateQuery(text)) {
    const parsed = _parseDatesWithClaude(text, _botToday());
    if (parsed && parsed.checkin && parsed.checkout && parsed.confidence > 0.4) {
      const personas     = parsed.persons || 2;
      const freeChildren = parsed.freeChildren || 0;
      if (midBooking) {
        sendWhatsAppText(from, '🔄 Veo que querés cambiar las fechas. Verifico disponibilidad para las nuevas...');
      }
      return _replyAvailability(from, contactName, { step: 'AWAITING_DATES', context: {}, name: contactName }, parsed.checkin, parsed.checkout, personas, freeChildren);
    }
    // Parsing fallo. Si el cliente mencionó personas o noches sueltas
    // (sin fechas concretas), mostramos tarifas como fallback util.
    const hasPersons = /\d+\s*personas?\b|\b(una|dos|tres|cuatro|cinco|seis)\s+personas?\b/i.test(text);
    const hasNoches  = /\d+\s*noches?\b|\b(una|dos|tres|cuatro|cinco)\s+noches?\b/i.test(text);
    if (hasPersons || hasNoches) {
      _botSendPricingInfo(from, contactName, conv);
      return;
    }
    sendWhatsAppText(from, '🤔 No logré entender las fechas. ¿Podés escribirlas más claras?\n\nEjemplo: "del 5 al 8 de junio, 4 personas".');
    return;
  }

  if (t === '1' || t.includes('disponibilidad') || t.includes('disponible') || t.includes('precios') || t.includes('cuanto cuesta') || t.includes('cuánto cuesta') || t.includes('reservar')) {
    return botHandleMessage(from, 'menu_disponibilidad', contactName, 'list_reply');
  }

  // Fallback de seleccion de cabana por texto (por si el cliente escribe en vez de tocar el boton)
  if (conv.step === 'SHOWING_AVAILABILITY') {
    const elegida = _botDetectCabin(text);
    if (elegida) {
      return botHandleMessage(from, 'pick_' + elegida, contactName, 'button_reply');
    }
  }

  // Smart fallback con Claude grounded: si el mensaje vale la pena procesar
  // con LLM (no es un saludo trivial), intentamos respuesta conversacional
  // antes de caer al menu de bienvenida.
  if (kind === 'text' && _botShouldUseSmartFallback(text, conv)) {
    if (_botSmartFallback(from, contactName, conv, text)) return;
  }

  // Fallback final: menu interactivo de bienvenida con instrucciones de
  // reserva al inicio. Resetea state a INITIAL para que conversaciones
  // atrapadas en AWAITING_DATES vuelvan al menu principal.
  _saveConv(from, 'INITIAL', conv.context || {}, contactName);
  _botSendMainMenu(from, contactName, true);
}

// ─── Reply helper: muestra disponibilidad con lista interactiva ──
// 1-4 personas: muestra cabañas individuales libres del tamaño requerido.
// 5+ personas: deriva al equipo (combo no se cotiza automatico desde el bot).
// Al final: lista interactiva con cabañas + opcion para cambiar personas (2,3,4).
function _replyAvailability(from, contactName, conv, checkin, checkout, personas, freeChildren) {
  personas     = personas || 2;
  freeChildren = freeChildren || 0;
  // payingPersons = adultos + niños mayores/iguales a 5. La tarifa base
  // ($90/$110 noche) cubre 2 personas; el recargo solo aplica desde la 3ra
  // pagante. Los niños <5 ocupan espacio en la cabaña (cuentan en personas)
  // pero NO suman al recargo.
  const payingPersons = Math.max(2, personas - freeChildren);
  const dates  = { checkin: checkin, checkout: checkout };

  // 5+ personas (TOTAL) → handoff al equipo (no cotizamos combo automatico desde el bot)
  if (personas >= 5) {
    const fechas = _botFmtFecha(checkin) + ' → ' + _botFmtFecha(checkout);
    sendWhatsAppText(from,
      '👥 Para grupos de *' + personas + ' personas* coordinamos directo con vos para ajustar combo de cabañas y detalles.\n\n' +
      'Tocá el botón abajo para escribir al equipo y resolverlo en un mensaje.'
    );
    try {
      sendWhatsAppCTAUrl(from,
        'Te pasamos cotización y métodos de pago.',
        '💬 Abrir WhatsApp',
        'https://wa.me/50769812266?text=' + encodeURIComponent('Hola, quiero cotizar para ' + personas + ' personas, ' + fechas)
      );
    } catch(_) {}
    try {
      sendWhatsAppText('50769812266',
        '🔔 Consulta de grupo grande vía Agente:\n👤 ' + (contactName || from) + '\n📱 +' + from + '\n📅 ' + fechas + '\n👥 ' + personas + ' personas');
    } catch(_) {}
    _saveConv(from, 'PENDING_HUMAN_BOOKING', { dates: dates, personas: personas, freeChildren: freeChildren }, contactName);
    return;
  }

  const avail  = _botCheckAvailability(checkin, checkout);
  const nights = Math.round((new Date(checkout + 'T12:00:00') - new Date(checkin + 'T12:00:00')) / 86400000);
  const fechasStr = _botFmtFecha(checkin) + ' → ' + _botFmtFecha(checkout) + ' · ' + nights + (nights === 1 ? ' noche' : ' noches');

  const opciones = [];
  ['azul', 'verde', 'lila'].forEach(c => {
    if (!avail[c]) return;
    if (BOT_CABIN_CAPACITY[c] < personas) return;
    // Capacidad usa total; pricing usa solo los pagantes (descuenta niños <5).
    const precio = _botPrecioCabin(c, checkin, checkout, payingPersons);
    opciones.push({ cabin: c, precio: precio });
  });

  if (opciones.length === 0) {
    const alts = _botSuggestAlternatives(checkin, checkout, personas);
    if (alts.length > 0) {
      const body =
        '😔 No tenemos disponibilidad para *' + fechasStr + '* con ' + personas + (personas === 1 ? ' persona' : ' personas') + '.\n\n' +
        'Pero sí tenemos para estas fechas cercanas:';
      const buttons = alts.slice(0, 3).map(a => ({ id: 'alt_' + a.checkin, title: _botFmtFecha(a.checkin) }));
      try {
        sendWhatsAppButtons(from, body, buttons, null, 'Tocá una opción o escribime "agente"');
      } catch(_) {
        sendWhatsAppText(from, body + '\n\n' + alts.map(a => '• ' + _botFmtFecha(a.checkin) + ' → ' + _botFmtFecha(a.checkout)).join('\n') + '\n\nEscribime las fechas que prefieras.');
      }
      _saveConv(from, 'SHOWING_ALTERNATIVES', { dates: dates, personas: personas, freeChildren: freeChildren, alts: alts }, contactName);
      return;
    }
    sendWhatsAppText(from,
      '😔 No tenemos disponibilidad para *' + fechasStr + '* con ' + personas + (personas === 1 ? ' persona' : ' personas') + '.\n\n' +
      'Podés ver el calendario público:\nhttps://lasnubes.cloud\n\n' +
      '¿O preferís hablar con un agente? Tocá *Hablar con un agente* en el menú.'
    );
    _saveConv(from, 'NO_AVAILABILITY', { dates: dates, personas: personas, freeChildren: freeChildren }, contactName);
    return;
  }

  // Cotizacion (formato copyPromo, sin combo)
  const cotizacion = _botCotizacionAvailability(checkin, checkout, opciones, personas, false, freeChildren);
  sendWhatsAppText(from, cotizacion);

  // Botones directos para reservar cada cabaña disponible (max 3).
  // WhatsApp limita title a 20 chars — usamos formato corto sin "por/hacia/entre".
  const CABIN_BUTTON_LABELS = {
    verde: '🏡 Paseo Las Nubes',
    azul:  '🏡 Portal Las Nubes',
    lila:  '🏡 Puente Las Nubes'
  };
  const cabinButtons = opciones.slice(0, 3).map(op => ({
    id: 'pick_' + op.cabin,
    title: CABIN_BUTTON_LABELS[op.cabin] || ('🏡 ' + BOT_CABIN_NAMES[op.cabin].split(' ')[0])
  }));
  try {
    sendWhatsAppButtons(from, '¿Cuál te interesa reservar?', cabinButtons, null, personas + ' personas · ' + nights + (nights === 1 ? ' noche' : ' noches'));
  } catch(_) {
    sendWhatsAppText(from, 'Escribime el nombre de la cabaña (Paseo / Portal / Puente) o "agente" para hablar con una persona.');
  }

  // Despues: lista "Ver opciones" con cambiar personas / otras fechas / agente
  const personaOpts = [2, 3, 4].filter(n => n !== personas);
  const personasRows = personaOpts.map(n => ({
    id: 'persons_' + n,
    title: '👥 ' + n + ' personas'
  }));
  const sections = [];
  if (personasRows.length > 0) {
    sections.push({ title: 'Cambiar personas', rows: personasRows });
  }
  sections.push({
    title: 'Otras opciones',
    rows: [
      { id: 'try_dates',   title: '📅 Otras fechas',         description: 'Cambiar las fechas de la consulta' },
      { id: 'menu_agente', title: '🙋 Hablar con un agente', description: 'WhatsApp del equipo' }
    ]
  });
  try {
    sendWhatsAppList(from, '¿Querés cambiar algo? Tocá ⬇', sections, '📋 Ver opciones');
  } catch(_) { /* ignorable: ya tiene los botones de cabaña */ }

  _saveConv(from, 'SHOWING_AVAILABILITY', { dates: dates, personas: personas, freeChildren: freeChildren, opciones: opciones.length }, contactName);
}

// ─── Menu principal interactivo (lista) ──────────────────────────
// firstTime=true → muestra bienvenida elaborada. firstTime=false → solo "¿Necesitás algo más?"
// customBody (opcional) → reemplaza el cuerpo (uso desde flujos especiales como campaign).
function _botSendMainMenu(from, contactName, firstTime, customBody) {
  const firstName = ((contactName || '').toString().trim().split(/\s+/)[0]) || '';
  const greeting  = firstName ? '¡Hola ' + firstName + '! 🌿' : '¡Hola! 🌿';
  let body;
  if (customBody) {
    body = customBody;
  } else if (firstTime === false) {
    body = '¿Necesitás algo más? Tocá *Ver opciones* abajo 👇';
  } else {
    body = greeting + '\n\n' +
      'Bienvenido a *Las Nubes* — un refugio de tres cabañas en las faldas del Cerro Chicá, a 1h 15min de la ciudad.\n\n' +
      'Para reservar o consultar disponibilidad, cuéntame *fechas* y *personas*. Por ejemplo:\n' +
      '   _"del 5 al 8 de junio, 2 personas"_\n\n' +
      'Te envío disponibilidad y precio al instante, y cerramos la reserva por aquí mismo. 🤝\n\n' +
      '¿Quieres explorar antes — cómo llegar, actividades, fotos o hablar con un agente? Toca *Ver opciones* ⬇';
  }
  const sections = [
    {
      title: 'Reservas',
      rows: [
        { id: 'menu_disponibilidad', title: '📅 Disponibilidad', description: 'Ver fechas libres y precios' },
        { id: 'menu_he_llegado',     title: '🚪 He llegado',     description: 'Estoy en el portón de Las Nubes' }
      ]
    },
    {
      title: 'Sobre Las Nubes',
      rows: [
        { id: 'menu_como_llegar',  title: '📍 Cómo llegar',     description: 'Dirección, Waze, Maps' },
        { id: 'menu_actividades',  title: '🏞 Actividades',     description: 'Cascadas, playas, cerros' },
        { id: 'menu_gastronomia',  title: '🍽 Gastronomía',     description: 'Restaurantes cerca' },
        { id: 'menu_insumos',      title: '🛒 Insumos',         description: 'Tiendita y supermercados' },
        { id: 'menu_tienda',       title: '🧊 Hielo y carbón',  description: 'Tienda a 5 min de la cabaña' }
      ]
    },
    {
      title: 'Ayuda',
      rows: [
        { id: 'menu_faq',     title: '❓ Preguntas frecuentes', description: 'Cocina, energía, check-in' },
        { id: 'menu_agente',  title: '🙋 Hablar con un agente', description: 'Abrir WhatsApp del equipo' }
      ]
    }
  ];
  try {
    sendWhatsAppList(from, body, sections, 'Ver opciones', null, 'Buenos Aires, Chamé · Panamá');
  } catch(err) {
    logDebugEntry('bot-menu-FAIL', { error: err.message });
    sendWhatsAppText(from, body + '\n\nEscribime qué te interesa:\n📅 Disponibilidad · 📍 Cómo llegar · 🏞 Actividades · 🍽 Gastronomía · 🛒 Insumos · 🧊 Hielo y carbón · ❓ FAQ · 🙋 Agente');
  }
}

function _botMenuComoLlegar(from) {
  sendWhatsAppText(from,
    '📍 *Cómo llegar a Las Nubes*\n\n' +
    'Por la carretera Interamericana, entrá por el *Pío Pío de Bejuco* a la carretera Bejuco–Sorá. ' +
    'Al llegar al pueblo de *Buenos Aires*, doblá a la derecha hacia *Chicá*. La cabaña queda a 100 metros.\n\n' +
    '🚗 Lo más fácil: poné en *Waze "Aires de Chicá"* — te lleva directo al portón verde. ' +
    'Cuando llegues, escribime o llamame para abrir y guiarte a la cabaña.\n\n' +
    '🗺 Google Maps:\nhttps://maps.google.com/?q=8.639400,-79.945900\n\n' +
    '🚦 Waze (abre con navegación):\nhttps://waze.com/ul?ll=8.639400,-79.945900&navigate=yes'
  );
}

function _botMenuTienda(from) {
  sendWhatsAppText(from,
    '🧊 *Tienda de conveniencia cercana*\n\n' +
    'Contamos con una tienda a tan solo *5 minutos* de la cabaña. En ella podés encontrar:\n\n' +
    '• Hielo\n• Carbón\n• Especias\n• Bebidas\n• Insumos básicos\n\n' +
    '📍 Ubicación:\nhttps://maps.google.com/?q=8.631809,-79.944489'
  );
}

function _botMenuActividades(from) {
  sendWhatsAppText(from,
    '🏞 *Actividades cerca*\n\n' +
    '• *Cascada Las Nubes* — sendero desde la cabaña 🥾\n' +
    '• *Los Cajones de Chame* — cañón con pozas y saltos (10 min) 🏊\n' +
    '• *Parque Nacional Altos de Campana* — primer parque de Panamá, miradores 🦜\n' +
    '• *Cascadas Filipinas* — 7 cascadas encadenadas 💧\n' +
    '• *Cascada Manglarito* — 35m de caída 💦\n' +
    '• *Cascada Nativa* — acceso fácil, naturaleza solitaria\n' +
    '• *Playa Gorgona* — tranquila, atardeceres (15 min) 🏖\n' +
    '• *Coronado* — playa + restaurantes + comercios (20 min)\n\n' +
    'Fotos, mapas y detalles:\nhttps://lasnubes.cloud#actividades'
  );
}

function _botMenuGastronomia(from) {
  sendWhatsAppText(from,
    '🍽 *Gastronomía cerca*\n\n' +
    'Cerca de las cabañas (5-15 min):\n' +
    '• *Buenas Pizzas de Sorá* — masa fina, horno de leña 🍕\n' +
    '• *Pío Pío de Bejuco* — entrada interamericana 🍗\n' +
    '• *Restaurantes de Coronado* (20 min) — variedad: Las Bóvedas Fusión, Vista del Mar y más 🍴\n\n' +
    'Direcciones, horarios y fotos:\nhttps://lasnubes.cloud#gastronomia'
  );
}

function _botMenuInsumos(from) {
  sendWhatsAppText(from,
    '🛒 *Insumos y compras*\n\n' +
    '🌿 *Tiendita Las Nubes* (te lo llevamos a la cabaña):\n' +
    '• Kit de Fogata $10 (leña, cerillo, palillos, malvaviscos) 🔥\n' +
    '• Bolsa de carbón $5\n' +
    '• Repelente OFF Spray $8\n' +
    '• Repelente Family Care toallitas $5\n' +
    '• Kit pasta y cepillo $5\n' +
    '• Toallas sanitarias $5\n\n' +
    '🛍 *Supermercados cercanos*:\n' +
    '• Tienda de conveniencia (5 min)\n' +
    '• MiniSuper Buenos Precios (Bejuco)\n' +
    '• El Rey, Machetazo, Super 99, Riba Smith (Coronado, 20 min)\n\n' +
    'Más detalles:\nhttps://lasnubes.cloud#insumos'
  );
}

function _botMenuFAQ(from) {
  sendWhatsAppText(from,
    '❓ *Preguntas frecuentes*\n\n' +
    '*¿Tiene cocina equipada?*\n' +
    'Sí, completa + BBQ. Incluye café, azúcar y especias básicas. Cooler grande (no nevera) — trae hielo y alimentos.\n\n' +
    '*¿Cómo es la energía?*\n' +
    '100% solar. Inversor para cargar celulares. Excelente señal de todas las operadoras.\n\n' +
    '*¿Check-in y check-out?*\n' +
    'Entrada: *2:00 pm* · Salida: *11:00 am*\n\n' +
    '*¿Baño?*\n' +
    'Jabón, papel y toallas incluidos. Fumigamos semanal — si sos sensible a mosquitos, trae repelente.\n\n' +
    '*¿Privacidad?*\n' +
    'Sí, toda la cabaña es de uso exclusivo de quienes reservan.\n\n' +
    '*¿Capacidad?*\n' +
    'Portal hasta 2 personas. Paseo y Puente hasta 4 (camas matrimoniales + auxiliar).\n\n' +
    '¿Otra duda? Tocá *Hablar con persona* o escribime "3".'
  );
}

// ─── Find reservation by client phone (for "He llegado") ────────
function _botFindReservaByPhone(phone) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const today = _botToday();
  const normalize = (t) => {
    let d = String(t || '').replace(/\D/g, '');
    if (d.indexOf('507') === 0 && d.length > 8) d = d.substring(3);
    return d;
  };
  const target = normalize(phone);
  if (!target) return null;
  let best = null;
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if ((r[20] || '').toString().toUpperCase() === 'CANCELADA') continue;
    if (r[9] === 'Abierta') continue;
    if (normalize(r[23]) !== target) continue;
    const ci = r[4] instanceof Date ? Utilities.formatDate(r[4], BOT_TZ, 'yyyy-MM-dd') : (r[4] || '').toString().slice(0,10);
    const co = r[5] instanceof Date ? Utilities.formatDate(r[5], BOT_TZ, 'yyyy-MM-dd') : (r[5] || '').toString().slice(0,10);
    if (!ci || !co) continue;
    // Aceptar si hoy esta en [checkin-1, checkout+1]
    const dayBefore = _botAddDaysISO(ci, -1);
    const dayAfter  = _botAddDaysISO(co, 1);
    if (today >= dayBefore && today <= dayAfter) {
      best = {
        id: r[0], name: r[1], cabin: r[3],
        checkin: ci, checkout: co,
        persons: r[6], origin: r[9]
      };
    }
  }
  return best;
}

function _botMenuHeLlegado(from, contactName, conv) {
  const reserva = _botFindReservaByPhone(from);
  if (reserva) {
    return _botSendArrivalInstructions(from, contactName, conv, reserva);
  }
  // No match por telefono → preguntar nombre del titular
  sendWhatsAppText(from,
    '🌿 Recibí tu mensaje.\n\n' +
    'No encuentro una reserva activa con este número para hoy. Decime el *nombre completo del titular* de la reserva para ubicarla en el sistema.\n\n' +
    'Si preferís hablar directo con una persona, escribime "agente".'
  );
  _saveConv(from, 'AWAITING_ARRIVAL_NAME', conv.context || {}, contactName);
}

// Busca reserva por nombre + ventana de fechas activa. Fuzzy match
// (case-insensitive, sin acentos, substring en ambas direcciones).
function _botFindReservaByName(name) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const today = _botToday();
  const normalize = (s) => String(s || '').toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9 ]/g, '').trim();
  const target = normalize(name);
  if (!target || target.length < 3) return null;
  let best = null;
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if ((r[20] || '').toString().toUpperCase() === 'CANCELADA') continue;
    if (r[9] === 'Abierta') continue;
    const storedName = normalize(r[1]);
    if (!storedName) continue;
    if (!storedName.includes(target) && !target.includes(storedName)) continue;
    const ci = r[4] instanceof Date ? Utilities.formatDate(r[4], BOT_TZ, 'yyyy-MM-dd') : (r[4] || '').toString().slice(0,10);
    const co = r[5] instanceof Date ? Utilities.formatDate(r[5], BOT_TZ, 'yyyy-MM-dd') : (r[5] || '').toString().slice(0,10);
    if (!ci || !co) continue;
    const dayBefore = _botAddDaysISO(ci, -1);
    const dayAfter  = _botAddDaysISO(co, 1);
    if (today >= dayBefore && today <= dayAfter) {
      best = {
        row: i + 1,
        id: r[0], name: r[1], cabin: r[3],
        checkin: ci, checkout: co,
        persons: r[6], origin: r[9]
      };
    }
  }
  return best;
}

function _botSendArrivalInstructions(from, contactName, conv, reserva) {
  const cabin    = reserva.cabin;
  const cabinName = BOT_CABIN_NAMES[cabin] || 'Las Nubes';
  const firstName = ((reserva.name || contactName || '').toString().trim().split(/\s+/)[0]) || '';

  let body = '🎉 ¡Bienvenidos a *Las Nubes*';
  if (firstName) body += ', ' + firstName;
  body += '!\n\nYa les abro el portón. Conducen recto y más adelante se van a encontrar con una *calle huella de concreto* — la suben y, cuando termine, toman la siguiente *calle a mano izquierda*.\n\n';

  if (cabin === 'azul') {
    body += 'Unos *25 metros más adelante* van a ver un *tanque de reserva de agua azul*. Estacionan *antes del tanque*, a los laterales de la calle.\n\n' +
            'Al lado del tanque está la *escalera para bajar a la cabaña*. Tiene *techo blanco* y la van a reconocer por los portales con puertas antiguas y la silla colgante.\n\n' +
            'Cualquier dificultad me escriben o llaman. ¡Quedo atento!';
  } else {
    body += 'Apenas doblen, *deténganse y me llaman* para indicarles dónde está la cabaña.\n\nQuedo atento.';
  }
  body += '\n\n📞 +507 6981-2266';

  try {
    sendWhatsAppCTAUrl(from, body, '💬 Escribir al equipo',
      'https://wa.me/50769812266?text=' + encodeURIComponent('Hola, recién llegué a Las Nubes 🌿'));
  } catch(err) {
    sendWhatsAppText(from, body + '\n\n💬 WhatsApp: https://wa.me/50769812266');
  }

  // Notificar al admin para que abra el porton
  const datesStr = _botFmtFecha(reserva.checkin) + ' → ' + _botFmtFecha(reserva.checkout);
  const gatePhone = PropertiesService.getScriptProperties().getProperty('WA_GATE_PHONE') || '+507 6777-5630';
  const adminMsg =
    '🚪 *ABRE EL PORTÓN* — llegó un cliente\n\n' +
    '📞 Llamar: ' + gatePhone + '\n\n' +
    '👤 ' + (reserva.name || contactName || '?') + '\n' +
    '📱 +' + from + '\n' +
    '🏡 ' + cabinName + '\n' +
    '📅 ' + datesStr + '\n' +
    '👥 ' + (reserva.persons || '?') + (reserva.persons === 1 ? ' persona' : ' personas') + '\n\n' +
    '_Notificación automática: el cliente tocó "He llegado" en el Agente._';
  try { sendWhatsAppText('50769812266', adminMsg); } catch(_) {}

  _saveConv(from, 'ARRIVED', Object.assign({}, conv.context || {}, { reservaId: reserva.id }), contactName);
}

// Busca una reserva por su id en la hoja. Devuelve datos basicos o null.
function _botFindReservaById(reservaId) {
  if (!reservaId) return null;
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === reservaId.toString()) {
      const r = data[i];
      const ci = r[4] instanceof Date ? Utilities.formatDate(r[4], BOT_TZ, 'yyyy-MM-dd') : (r[4] || '').toString().slice(0,10);
      const co = r[5] instanceof Date ? Utilities.formatDate(r[5], BOT_TZ, 'yyyy-MM-dd') : (r[5] || '').toString().slice(0,10);
      return {
        id: r[0], name: r[1] || '?', cabin: r[3],
        cabinName: r[2] || BOT_CABIN_NAMES[r[3]] || r[3],
        checkin: ci, checkout: co, tipo: r[24] || 'noche',
        telefono: r[23] || ''
      };
    }
  }
  return null;
}

// El cliente tocó "Consultas y cambios" en la plantilla de confirmación.
// → el Agente le ofrece un botón para escribirle directo a Josh, con un
//   mensaje precargado que incluye su nombre, fecha y cabaña.
function _botHandleConsultaReserva(from, contactName, reservaId) {
  let reserva = _botFindReservaById(reservaId);
  if (!reserva) { try { reserva = _botFindReservaByPhone(from); } catch(_) {} }

  const nombre    = (reserva && reserva.name) || contactName || '';
  const cabinName = (reserva && reserva.cabinName) || (reserva && BOT_CABIN_NAMES[reserva.cabin]) || '';
  const fechaStr  = (reserva && reserva.checkin) ? _botFmtFecha(reserva.checkin) : '';

  let prefill = 'Hola Josh, mi nombre es ' + (nombre || '...');
  if (fechaStr)  prefill += ' y tengo una reserva para el ' + fechaStr;
  else           prefill += ' y tengo una reserva en Las Nubes';
  if (cabinName) prefill += ' en la cabaña ' + cabinName;
  prefill += '. Tengo una consulta respecto a mi reserva, ¿me ayudás?';

  sendWhatsAppText(from, 'Josh te va a asistir con cualquier duda o consulta relacionada a tu reserva. 🌿');
  try {
    sendWhatsAppCTAUrl(from,
      'Tocá el botón para escribirle directo 👇',
      'Escribirle a Josh',
      'https://wa.me/50769812266?text=' + encodeURIComponent(prefill)
    );
  } catch(_) {
    sendWhatsAppText(from, 'Escribile directo aquí:\nhttps://wa.me/50769812266');
  }
  _saveConv(from, 'HUMAN_HANDOFF', (reserva && reserva.id) ? { reservaId: reserva.id } : {}, contactName);
  logDebugEntry('consulta-reserva', { from: from, reservaId: reservaId });
}

// El huesped tocó "Ya me retiré" en la plantilla de check-out.
// → avisa al admin que abra el portón y a Erika que puede limpiar la cabaña.
function _botHandleCheckoutDone(from, contactName, reservaId) {
  let reserva = _botFindReservaById(reservaId);
  // Fallback: si no vino id valido, intentar ubicar por telefono.
  if (!reserva) {
    try { reserva = _botFindReservaByPhone(from); } catch(_) {}
  }
  const cabinName = (reserva && reserva.cabinName) || (reserva && BOT_CABIN_NAMES[reserva.cabin]) || 'una cabaña';
  const guestName = (reserva && reserva.name) || contactName || from;

  // Confirmacion al huésped
  sendWhatsAppText(from,
    '¡Gracias por avisar! 🌿 Ya le avisé al equipo, en un momento te abren el portón. ' +
    '¡Buen viaje y esperamos verte pronto de nuevo en Las Nubes! 🙌'
  );

  // Aviso al admin: abrir portón
  const gatePhone = PropertiesService.getScriptProperties().getProperty('WA_GATE_PHONE') || '+507 6777-5630';
  try {
    sendWhatsAppText(BOT_ADMIN_PHONE,
      '🚪 *ABRE EL PORTÓN — salida*\n\n' +
      '📞 Llamar: ' + gatePhone + '\n\n' +
      '👤 ' + guestName + '\n' +
      '📱 +' + from + '\n' +
      '🏡 ' + cabinName + '\n\n' +
      '_El huésped tocó "Ya me retiré" en el check-out._'
    );
  } catch(_) {}

  // Aviso a Erika (limpieza): cabaña libre para limpiar
  try {
    const limpiezaPhone = PropertiesService.getScriptProperties().getProperty('LIMPIEZA_PHONE');
    if (limpiezaPhone) {
      sendWhatsAppText(limpiezaPhone,
        '🧹 *' + cabinName + '* ya está libre\n\n' +
        guestName + ' acaba de retirarse. Podés proceder con la limpieza cuando puedas. 🌿'
      );
    }
  } catch(_) {}

  logDebugEntry('checkout-done', { from: from, reservaId: reservaId, cabin: (reserva && reserva.cabin) || '?' });
}

const BOT_ADMIN_PHONE = '50769812266';

// Camas por cabana (igual que index.html / dashboard)
const BOT_CABIN_CAMAS = {
  verde: 'La cabaña cuenta con una cama matrimonial queen y un sofá-cama doble.',
  azul:  'La cabaña solo cuenta con una cama matrimonial full. Puede traer colchón inflable.',
  lila:  'La cabaña cuenta con una cama matrimonial queen y una cama auxiliar full.'
};

function _botPaymentInfo() {
  const custom = PropertiesService.getScriptProperties().getProperty('WA_PAYMENT_INFO');
  if (custom) {
    // Soporta saltos de linea escritos como literal "\n" (Script Properties UI
    // los guarda asi cuando los tipeas), o newlines reales si los pegas.
    return custom.replace(/\\n/g, '\n');
  }
  // Default: formato estandar de Las Nubes
  return 'Puede realizar el pago a través de:\n\n' +
    '*Yappy*\n69812266\nJoslyn Lopez\n\n' +
    '*ACH*\nBanco General\nJoslyn Lopez\nCta de Ahorros\n04-99-99-863047-1\n\n' +
    '*Colocar en la sección "Agregar Mensaje" del Yappy o descripción de la transferencia:*\n' +
    '*Nombre Completo*\n*Email*\n*Celular*\n\n' +
    'Quedo atento para proceder a cerrar el espacio de inmediato.';
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
function _botCotizacionAvailability(checkin, checkout, opciones, personas, isCombo, freeChildren) {
  freeChildren = freeChildren || 0;
  const nights      = Math.round((new Date(checkout + 'T12:00:00') - new Date(checkin + 'T12:00:00')) / 86400000);
  const fechaIn     = _botFmtFecha(checkin);
  const fechaOut    = _botFmtFecha(checkout);
  let personasLbl = personas + (personas === 1 ? ' persona' : ' personas');
  if (freeChildren > 0) {
    const childWord = freeChildren === 1 ? 'menor de 5 años (sin cargo)' : 'menores de 5 años (sin cargo)';
    personasLbl += ' (incluye ' + freeChildren + ' ' + childWord + ')';
  }

  let intro;
  if (nights === 1) intro = 'Tengo la noche del *' + fechaIn + '* disponible para reserva para ' + personasLbl + '.';
  else              intro = 'Tengo las noches del *' + fechaIn + ' al ' + fechaOut + '* disponibles para reserva para ' + personasLbl + '.';

  let cabinasBlock;
  if (isCombo) {
    const precio = opciones[0].precio;
    cabinasBlock =
      '🏡 *Combo: Puente entre Las Nubes + Portal hacia Las Nubes*\n' +
      '_(cabañas contiguas, perfectas para grupos)_\n' +
      'Puente: 2 camas matrimoniales (queen y full)\n' +
      'Portal: 1 cama matrimonial full\n' +
      '💰 *Total:* $' + precio.toFixed(2) + '\n';
  } else if (opciones.length === 1) {
    const op = opciones[0];
    cabinasBlock = '🏡 *Cabaña:* ' + BOT_CABIN_NAMES[op.cabin] + '\n';
    if (personas >= 3) cabinasBlock += BOT_CABIN_CAMAS[op.cabin] + '\n';
    cabinasBlock += '💰 *Total:* $' + op.precio.toFixed(2) + '\n';
  } else {
    cabinasBlock = '*Disponibles:*\n';
    opciones.forEach(op => {
      cabinasBlock += '• ' + BOT_CABIN_NAMES[op.cabin] + ' — $' + op.precio.toFixed(2) + '\n';
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
  const dates        = conv.context && conv.context.dates;
  const personas     = (conv.context && conv.context.personas) || 2;
  const freeChildren = (conv.context && conv.context.freeChildren) || 0;
  if (!dates) {
    sendWhatsAppText(from, '🤔 Perdí el contexto de las fechas. ¿Podés decirme de nuevo cuándo querés reservar?');
    _saveConv(from, 'AWAITING_DATES', conv.context, contactName);
    return;
  }
  const payingPersons = Math.max(2, personas - freeChildren);
  const precio  = _botPrecioCabin(cabin, dates.checkin, dates.checkout, payingPersons);
  const fechas  = _botFmtFecha(dates.checkin) + ' → ' + _botFmtFecha(dates.checkout);
  let personasLbl = personas + (personas === 1 ? ' persona' : ' personas');
  if (freeChildren > 0) {
    personasLbl += ' · ' + freeChildren + ' menor' + (freeChildren === 1 ? '' : 'es') + ' de 5 sin cargo';
  }
  const body =
    '🎉 ¡Excelente! Reservando *' + BOT_CABIN_NAMES[cabin] + '* para *' + fechas + '* (' + personasLbl + ').\n\n' +
    '💰 *Total: $' + precio.toFixed(2) + '*\n\n' +
    _botPaymentInfo() + '\n\n' +
    '⚠️ *Importante*: en el detalle de la transferencia colocá tu *nombre completo* y *email* para procesar tu reserva más rápido.\n\n' +
    'Una vez transferido, *enviame el comprobante como imagen* por aquí mismo.';
  try {
    sendWhatsAppButtons(from, body, [
      { id: 'cancel_booking', title: '❌ Cancelar' }
    ]);
  } catch(_) {
    sendWhatsAppText(from, body + '\n\nSi querés cancelar, escribime "cancelar".');
  }
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
  // Recuperar campos extraidos del campo "Mensaje" del voucher (si el cliente los coloco)
  const extractedName  = voucher.nombreCompleto ? voucher.nombreCompleto.toString().trim() : '';
  const extractedEmail = voucher.email ? voucher.email.toString().trim().toLowerCase() : '';
  const newCtx = Object.assign({}, conv.context, {
    voucher: {
      monto: monto,
      codTransferencia: voucher.codTransferencia,
      fechaPago: voucher.fechaPago || _botToday(),
      sender:   voucher.sender || ''
    },
    name:  extractedName  || (conv.context && conv.context.name)  || '',
    email: extractedEmail || (conv.context && conv.context.email) || ''
  });

  // Confirmar el voucher con datos extraidos
  let confirmMsg = '✅ ¡Comprobante recibido!\n\n' +
    '*Remitente:* ' + (voucher.sender || '—') + '\n' +
    '*Monto:* $' + monto.toFixed(2) + '\n' +
    '*Código:* ' + voucher.codTransferencia;
  if (extractedName)  confirmMsg += '\n*Nombre:* ' + extractedName;
  if (extractedEmail) confirmMsg += '\n*Email:* ' + extractedEmail;

  // Decidir proximos pasos segun que datos vienen en el voucher
  if (newCtx.name && newCtx.email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(newCtx.email)) {
    // Tenemos todo → crear pre-reserva directamente
    sendWhatsAppText(from, confirmMsg);
    return _botCreatePreReservation(from, contactName, newCtx);
  }
  if (!newCtx.email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(newCtx.email)) {
    sendWhatsAppText(from, confirmMsg + '\n\nPara finalizar, ¿me podés enviar tu *email*?');
    _saveConv(from, 'AWAITING_EMAIL', newCtx, contactName);
    return;
  }
  // Tenemos email pero falta nombre
  sendWhatsAppText(from, confirmMsg + '\n\n¿Cuál es tu *nombre completo*?');
  _saveConv(from, 'AWAITING_NAME', newCtx, contactName);
}

function _botCreatePreReservation(from, contactName, ctx) {
  const dates        = ctx.dates;
  const cabin        = ctx.cabin;
  const personas     = ctx.personas || 2;
  const freeChildren = ctx.freeChildren || 0;
  const email        = ctx.email;
  const fullName     = ctx.name;
  const voucher      = ctx.voucher || {};
  const skipVoucher  = !!ctx.skipVoucher;
  const payingPersons = Math.max(2, personas - freeChildren);
  const precio   = ctx.precio || _botPrecioCabin(cabin, dates.checkin, dates.checkout, payingPersons);
  const id       = Date.now().toString();
  const today    = _botToday();
  const CABIN_NAMES = {
    verde: 'Paseo por Las Nubes',
    azul:  'Portal hacia Las Nubes',
    lila:  'Puente entre Las Nubes'
  };
  const comentario = skipVoucher
    ? '🤖 Pre-reserva vía Agente WhatsApp · SIN ABONO · coordinar pago · pendiente revisión'
    : '🤖 Pre-reserva vía Agente WhatsApp · pendiente revisión';

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
      _safeCell(comentario),
      _safeCell(from),
      'noche'                  // tipo
    ]);
    logDebugEntry('bot-prereserva-OK', { id: id, name: fullName, cabin: cabin, from: from, skipVoucher: skipVoucher });
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
  const adminHeader = skipVoucher
    ? '📋 *Nueva pre-reserva SIN ABONO vía Agente*\n⚠️ Coordinar el pago manualmente antes de aprobar.\n'
    : '📋 *Nueva pre-reserva vía Agente*\n';
  const voucherBlock = skipVoucher
    ? '💳 _Sin voucher · coordinar pago_'
    : '💳 Voucher: $' + (voucher.monto || 0).toFixed(2) + ' (' + (voucher.sender || '?') + ')\n' +
      '#️⃣ Código: ' + (voucher.codTransferencia || '?');
  const adminMsg =
    adminHeader + '\n' +
    '👤 ' + fullName + '\n' +
    '📧 ' + email + '\n' +
    '📱 ' + from + '\n\n' +
    '🏡 ' + (CABIN_NAMES[cabin] || cabin) + '\n' +
    '📅 ' + fechas + '\n' +
    '👥 ' + personas + (personas === 1 ? ' persona' : ' personas') +
      (freeChildren > 0 ? ' (incluye ' + freeChildren + ' menor' + (freeChildren === 1 ? '' : 'es') + ' de 5)' : '') + '\n' +
    '💰 Total: $' + precio.toFixed(2) + '\n\n' +
    voucherBlock;
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
      // Limpiar el marker de "pendiente revisión" en comentarios
      const prevCmt = (data[i][22] || '').toString();
      const newCmt  = prevCmt.replace(/🤖 Pre-reserva v[ií]a (?:bot|Agente) WhatsApp · pendiente revisi[oó]n\s*\.?\s*/i, '').trim();
      const approvedTag = '✅ Aprobada vía Agente WhatsApp · ' + Utilities.formatDate(new Date(), BOT_TZ, 'yyyy-MM-dd HH:mm');
      sheet.getRange(row, 23).setValue(newCmt ? (newCmt + '\n' + approvedTag) : approvedTag);
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
      const prevCmt = (data[i][22] || '').toString();
      const newCmt  = prevCmt.replace(/🤖 Pre-reserva v[ií]a (?:bot|Agente) WhatsApp · pendiente revisi[oó]n\s*\.?\s*/i, '').trim();
      const rejectedTag = '❌ Rechazada vía Agente WhatsApp · ' + Utilities.formatDate(new Date(), BOT_TZ, 'yyyy-MM-dd HH:mm');
      sheet.getRange(row, 23).setValue(newCmt ? (newCmt + '\n' + rejectedTag) : rejectedTag);
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
