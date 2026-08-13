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
// Fallback si no se puede leer Config. El precio real sale de
// precioNochePublico() en Parser.gs, que lee la hoja y aplica tipo de día,
// promoción vigente y feriados. Estos dos números NO se usan para cotizar
// salvo que la lectura falle.
const BOT_RATE_WEEKDAY = 90;
const BOT_RATE_WEEKEND = 110;

// Texto de tarifas armado desde Config, para que el bot nunca anuncie un
// precio que el calendario no cobra.
function _botTextoTarifas() {
  const t = tarifasPublicas();
  const vig = function(cat, promoCat) {
    const p = t[promoCat] || 0;
    return (p > 0 && t.promoActiva) ? '*$' + p + '* (antes $' + t[cat] + ')' : '*$' + t[cat] + '*';
  };
  let s = '• Domingo a jueves: ' + vig('semana', 'promoSemana') + '/noche\n'
        + '• Viernes: ' + vig('viernes', 'promoViernes') + '/noche\n'
        + '• Sábado: ' + vig('sabado', 'promoSabado') + '/noche\n';
  // Solo se nombran si cuestan distinto: listar categorías que valen lo mismo
  // que un día normal es ruido que además suena a letra chica.
  const extras = [];
  if (t.feriado !== t.semana) extras.push('feriados $' + t.feriado);
  if (t.vispera !== t.semana) extras.push('vísperas de feriado $' + t.vispera);
  if (t.escolar !== t.semana || (t.promoSemana > 0 && t.promoEscolar === 0))
    extras.push('vacaciones escolares $' + t.escolar);
  if (extras.length) s += '\n_' + extras.join(' · ') + '._\n';
  return s;
}
const BOT_RECARGO_PERSONA_GRANDE = 20;  // Paseo, Puente
const BOT_RECARGO_PERSONA_PORTAL = 10;  // Portal
const BOT_RECARGO_COMBO_5 = 80;          // 5 personas: Puente + Portal contiguas, por noche
const BOT_RECARGO_COMBO_6 = 100;         // 6 personas
const BOT_DECOR_FEE = 40;                // decoración especial (cumpleaños/aniversario)
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

// El appendRow final es una condición de carrera: dos ejecuciones concurrentes
// leen la hoja, ninguna encuentra el teléfono, y las DOS agregan la fila. Así
// nacieron las filas repetidas de 50761246512, 50762962863 y 50767765620 —
// creadas con menos de un segundo de diferencia entre sí.
//
// Con la fila duplicada, _getConv devuelve siempre la PRIMERA, así que la
// segunda queda como estado fantasma: si una escritura cayó en una y la lectura
// toma la otra, la conversación parece retroceder de paso.
//
// El lock envuelve leer-buscar-escribir, que es lo que tiene que ser atómico.
function _saveConv(phone, step, context, name) {
  const lock = LockService.getScriptLock();
  let tengoLock = false;
  try { tengoLock = lock.tryLock(10000); } catch(_) {}
  try {
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
  } finally {
    if (tengoLock) { try { lock.releaseLock(); } catch(_) {} }
  }
}

// Limpia las filas repetidas que dejó la carrera de arriba. Conserva la MÁS
// RECIENTE por teléfono —es la que refleja el estado real— y borra el resto.
// Preview por default; borra con limpiarConversacionesDuplicadas(true).
function limpiarConversacionesDuplicadas(aplicar) {
  const sheet = _convSheet();
  const data  = sheet.getDataRange().getValues();
  const porTel = {};
  for (let i = 1; i < data.length; i++) {
    const tel = data[i][0] ? data[i][0].toString() : '';
    if (!tel) continue;
    (porTel[tel] = porTel[tel] || []).push({ fila: i + 1, ts: String(data[i][2] || ''), step: data[i][1] });
  }
  const aBorrar = [];
  Object.keys(porTel).forEach(tel => {
    const fs = porTel[tel];
    if (fs.length < 2) return;
    // Ordena por timestamp; ante empate gana la de más abajo (la más nueva).
    fs.sort((a, b) => a.ts === b.ts ? a.fila - b.fila : (a.ts < b.ts ? -1 : 1));
    const queda = fs[fs.length - 1];
    fs.slice(0, -1).forEach(f => aBorrar.push({ tel: tel, fila: f.fila, step: f.step, ts: f.ts }));
    Logger.log(tel + ': ' + fs.length + ' filas → conservo fila ' + queda.fila + ' (' + queda.step + ' ' + queda.ts + ')');
  });
  if (!aBorrar.length) { Logger.log('✓ Sin duplicados.'); return; }
  Logger.log('Filas a borrar: ' + aBorrar.length);
  aBorrar.forEach(f => Logger.log('  fila ' + f.fila + ' · ' + f.tel + ' · ' + f.step + ' · ' + f.ts));
  if (aplicar !== true) { Logger.log('(preview) Llamar limpiarConversacionesDuplicadas(true) para borrar.'); return; }
  // De abajo hacia arriba: borrar de arriba corre los índices de las demás.
  aBorrar.sort((a, b) => b.fila - a.fila).forEach(f => sheet.deleteRow(f.fila));
  Logger.log('✓ ' + aBorrar.length + ' fila(s) borrada(s).');
}

// Util admin: corrige el estado de una conversación (preserva contexto/nombre).
// Correr desde el editor. Ej: corregirEstadoConversacion('50762879298', 'HUMAN_HANDOFF')
function corregirEstadoConversacion(phone, nuevoStep) {
  const conv = _getConv(String(phone));
  if (!conv) { Logger.log('No existe conversación: ' + phone); return; }
  _saveConv(String(phone), nuevoStep, conv.context, conv.name);
  Logger.log('✓ ' + phone + ': ' + conv.step + ' → ' + nuevoStep);
}

// Yessickam era una consulta de grupo (no un cierre) — sacarla del conteo.
function _fixYessickam() { return corregirEstadoConversacion('50762879298', 'HUMAN_HANDOFF'); }

// Corrige el estado buscando la conversación por nombre (substring, sin
// distinguir mayúsculas). Si hay 0 o más de 1 coincidencia, NO toca nada y
// lo registra en el log para que elijas el teléfono manualmente.
function corregirEstadoPorNombre(nombreParcial, nuevoStep) {
  const sheet = _convSheet();
  const data  = sheet.getDataRange().getValues();
  const q = (nombreParcial || '').toString().toLowerCase().trim();
  if (!q) { Logger.log('Falta el nombre a buscar.'); return; }
  const matches = [];
  for (let i = 1; i < data.length; i++) {
    const name = (data[i][4] || '').toString().toLowerCase();
    if (name.indexOf(q) !== -1) {
      matches.push({ phone: (data[i][0] || '').toString(), name: data[i][4], step: data[i][1] });
    }
  }
  if (matches.length === 0) { Logger.log('Sin coincidencias para "' + nombreParcial + '".'); return; }
  if (matches.length > 1) {
    Logger.log('⚠️ ' + matches.length + ' coincidencias para "' + nombreParcial + '": ' +
      matches.map(m => m.name + ' (' + m.phone + ', ' + m.step + ')').join(' | ') +
      '. No se tocó nada — corré corregirEstadoConversacion(telefono, estado) con el correcto.');
    return;
  }
  corregirEstadoConversacion(matches[0].phone, nuevoStep);
}

// ── Wrappers sin parámetros (ejecutables directo desde el editor) ────
// Yavy: primera venta automática del Agente → confirmada.
function fixYavy()      { return corregirEstadoPorNombre('yav', 'CONFIRMED'); }
// Yessickam: era consulta de grupo → fuera del conteo de reservas.
function fixYessickam() { return corregirEstadoConversacion('50762879298', 'HUMAN_HANDOFF'); }
// Yaviletzy: ya tenía reserva confirmada y al tocar "Consultas y cambios"
// el bot la mandó (incorrectamente) a HUMAN_HANDOFF. Volver a CONFIRMED.
function fixYaviletzy() { return corregirEstadoConversacion('50766866405', 'CONFIRMED'); }

// Diagnóstico rápido del toggle de alertas. Loguea la config actual y
// (si forceOn===true) la resetea a TODO encendido. Útil cuando el admin
// dejó de recibir notificaciones del Agente.
function debugAlertsConfig(forceOn) {
  const cfg = _botGetAlertConfig();
  Logger.log('BOT_ALERTS_CONFIG actual: ' + JSON.stringify(cfg, null, 2));
  if (forceOn === true) {
    _botSetAlertConfig({
      nuevoCliente:      true,
      eligiendoCierre:   true,
      pagando:           true,
      handoff:           true,
      seguimientoDiario: true
    });
    Logger.log('✓ Alertas reseteadas a TODO encendido.');
  }
  return cfg;
}

// Manda un texto al admin para validar entrega de WhatsApp.
// Si NO llega al WhatsApp del admin: problema de Meta (quality / bloqueo).
// Si SÍ llega pero las notificaciones de "nuevo cliente" no salen: el
// problema está aguas arriba (webhook, ruteo). Revisar debugRecentAdminEvents.
function testAdminMessage() {
  const stamp = Utilities.formatDate(new Date(), 'America/Panama', 'HH:mm:ss');
  const msg = '🔧 testAdminMessage @ ' + stamp + ' — si ves esto, la entrega WhatsApp al admin funciona.';
  try {
    sendWhatsAppText(BOT_ADMIN_PHONE, msg);
    Logger.log('✓ Enviado. Verificá si llegó a +' + BOT_ADMIN_PHONE);
  } catch(e) {
    Logger.log('✗ Falló el envío: ' + e.message);
  }
}

// Lee DebugLog y muestra las últimas N entradas relevantes a notificaciones
// admin / nuevo cliente / alertas. Útil para entender si el webhook está
// procesando entrantes y si las alertas se mandan o fallan en Meta.
function debugRecentAdminEvents(limit) {
  const max = parseInt(limit, 10) || 20;
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('DebugLog');
  if (!sheet || sheet.getLastRow() < 2) { Logger.log('Sin DebugLog'); return; }
  const data = sheet.getDataRange().getValues();
  const matches = [];
  const re = /(admin-new-lead|bot-admin-alert|WA-inbound|nuevoCliente|handoff)/i;
  for (let i = data.length - 1; i >= 1 && matches.length < max; i--) {
    const stage = (data[i][1] || '').toString();
    if (re.test(stage)) matches.push({ ts: data[i][0], stage: stage, info: (data[i][2] || '').toString().slice(0, 200) });
  }
  Logger.log('Últimas ' + matches.length + ' entradas relevantes (más nuevas primero):');
  matches.forEach(m => Logger.log('  ' + m.ts + ' · ' + m.stage + ' · ' + m.info));
}

// Dumpea la conversación completa de un teléfono al log del editor para
// poder revisarla offline. Normaliza dígitos (últimos 8 = número local PA).
function dumpConversacion(phone) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Mensajes');
  if (!sheet || sheet.getLastRow() < 2) { Logger.log('Sin hoja Mensajes / vacía'); return; }
  const data = sheet.getDataRange().getValues();
  const target = String(phone || '').replace(/\D/g, '').slice(-8);
  if (!target) { Logger.log('Teléfono inválido'); return; }
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const p = String(data[i][1] || '').replace(/\D/g, '').slice(-8);
    if (p !== target) continue;
    rows.push({
      ts: data[i][0],
      direction: data[i][2],
      type: data[i][3],
      content: String(data[i][4] || '').slice(0, 500)
    });
  }
  if (!rows.length) { Logger.log('Sin mensajes para +' + phone); return; }
  Logger.log('=== Conversación con +' + phone + ' (' + rows.length + ' mensajes) ===');
  rows.forEach(r => {
    const arrow = r.direction === 'in' ? '◀ ' : '▶ ';
    Logger.log(r.ts + ' ' + arrow + '[' + r.type + '] ' + r.content);
  });
}

// Wrapper de un toque para el hilo de Malu (+507 6532-9566)
function dumpMalu() { return dumpConversacion('50765329566'); }

// Prueba alerta_limpieza con varios códigos de idioma para encontrar cuál
// fue aprobado en Meta. El primero que funcione es el correcto para producción.
// Solo Erika recibe el que funciona (los fallidos no entregan nada).
function testAlertaLimpiezaAllLangs() {
  const phone = PropertiesService.getScriptProperties().getProperty('LIMPIEZA_PHONE');
  if (!phone) { Logger.log('✗ LIMPIEZA_PHONE no seteado.'); return; }
  Logger.log('LIMPIEZA_PHONE = ' + phone);
  const params = ['Portal hacia Las Nubes (PRUEBA)', 'María Pérez (PRUEBA)', '🛏 Preparar cama auxiliar para la próxima reserva (3 huéspedes).'];
  const langs  = ['es_ES', 'es_PA', 'es', 'es_MX', 'es_AR'];
  for (const lang of langs) {
    try {
      const res = sendWhatsAppTemplate(phone, 'alerta_limpieza_', lang, params, null, null);
      Logger.log('✅ FUNCIONA en idioma: ' + lang);
      Logger.log('  → Usar lang="' + lang + '" en _botHandleCheckoutDone.');
      return lang;
    } catch(e) {
      const short = e.message.length > 140 ? e.message.slice(0, 140) + '…' : e.message;
      Logger.log('✗ ' + lang + ': ' + short);
    }
  }
  Logger.log('✗ Ningún idioma funcionó. ¿Plantilla pendiente de aprobación o nombre mal escrito?');
}

// Diagnóstico: prueba el envío de alerta_limpieza a LIMPIEZA_PHONE.
// Loguea el valor del Script Property, el resultado del send y cualquier
// error de Meta (típico: nombre/idioma/variables de la plantilla no calzan).
function testAlertaLimpieza() {
  const phone = PropertiesService.getScriptProperties().getProperty('LIMPIEZA_PHONE');
  Logger.log('LIMPIEZA_PHONE = ' + (phone || '(no seteado)'));
  if (!phone) {
    Logger.log('✗ No hay LIMPIEZA_PHONE en Script Properties. Setealo y reintenta.');
    return;
  }
  try {
    const res = sendWhatsAppTemplate(phone, 'alerta_limpieza_', 'es_ES', [
      'Portal hacia Las Nubes (PRUEBA)',
      'María Pérez (PRUEBA)',
      '🛏 Preparar cama auxiliar para la próxima reserva (3 huéspedes).'
    ], null, null);
    Logger.log('✓ Enviado · respuesta Meta: ' + JSON.stringify(res));
  } catch(e) {
    Logger.log('✗ FALLÓ envío: ' + e.message);
    Logger.log('Posibles causas:');
    Logger.log('  - Idioma incorrecto (revisar en Meta: alerta_limpieza puede estar en es / es_PA, no es_ES).');
    Logger.log('  - Nombre de plantilla mal escrito.');
    Logger.log('  - Plantilla aún no aprobada (status "Pending review").');
    Logger.log('  - Número de Erika bloqueado por Meta (calidad).');
  }
}

// Estados de "intención de reserva": el lead está cerrando pero todavía no
// hay reserva ingresada en el sistema.
const _BOT_INTENT_STEPS = [
  'SHOWING_AVAILABILITY','SHOWING_ALTERNATIVES','CHOOSING_DECOR','CHOOSING_CLOSE',
  'OFFERING_PAYMENT','AWAITING_VOUCHER_RETRY','AWAITING_EMAIL','AWAITING_NAME',
  'PENDING_HUMAN_BOOKING','PENDING_REVIEW'
];

// Cuando se ingresa una reserva en el dashboard, si su teléfono matchea una
// conversación del bot en estado de intención, la eleva a CONFIRMED. Así el
// funnel cuenta la venta recién cuando la reserva existe en el sistema
// (clave para el cierre asistido: el lead no cuenta hasta que reservás).
// Devuelve el teléfono de la conversación elevada, o null.
function _botConfirmConversationByPhone(telefono) {
  try {
    const digits = (telefono || '').toString().replace(/\D/g, '');
    if (digits.length < 7) return null;            // teléfono vacío / inválido
    const tail = digits.slice(-8);                 // últimos 8 = número local PA
    const sheet = _convSheet();
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const convPhone = (data[i][0] || '').toString().replace(/\D/g, '');
      if (!convPhone || convPhone.slice(-8) !== tail) continue;
      const step = (data[i][1] || '').toString();
      if (_BOT_INTENT_STEPS.indexOf(step) === -1) continue;
      const phone = (data[i][0] || '').toString();
      let ctx = {};
      try { ctx = data[i][3] ? JSON.parse(data[i][3]) : {}; } catch(_) {}
      _saveConv(phone, 'CONFIRMED', ctx, data[i][4] || '');
      logDebugEntry('bot-conv-confirmed-by-reserva', { phone: phone, from: step });
      return phone;
    }
  } catch(e) {
    logDebugEntry('bot-conv-confirm-FAIL', { error: e.message });
  }
  return null;
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

// Detecta preguntas sobre métodos de pago / Yappy / ACH / transferencia.
// "Para hacer el pago?", "Cómo pago?", "Aceptan tarjeta?", "Por yappi", etc.
function _botLooksLikePaymentQuery(text) {
  const t = (text || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  return /\b(como\s+(pago|paga|se\s+paga|pagar|hago\s+el\s+pago)|para\s+(hacer\s+el\s+)?pag(o|ar)|metodos?\s+de\s+pago|forma[s]?\s+de\s+pago|yappy|yappi|\bach\b|sinpe|transferencia(\s+bancaria)?|deposito(\s+bancario)?|pago\s+contra|tarjeta(\s+de\s+credito)?|paypal|aceptan\s+(tarjeta|efectivo|paypal)|abon(o|ar|ando|amos)|adelanto|adelantar|reservar\s+abon|pagar\s+(completo|todo|el\s+total|en\s+total)|pago\s+(completo|total))\b/.test(t);
}

// Respuesta breve a preguntas de pago sin romper el flujo activo de booking.
// Mantiene el estado: si está en SHOWING_AVAILABILITY / CHOOSING_DECOR /
// CHOOSING_CLOSE, le recordamos que tiene opciones arriba y no le pedimos
// fechas/personas otra vez.
function _botHandlePaymentInfo(from, contactName, conv, text) {
  const inBookingFlow = [
    'SHOWING_AVAILABILITY','SHOWING_ALTERNATIVES',
    'CHOOSING_DECOR','CHOOSING_CLOSE'
  ].indexOf(conv.step) !== -1;

  const t = (text || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  // Pregunta específica por el NÚMERO de Yappy. Caso Malu: Claude alucinó
  // "es el mismo al que escribes por WhatsApp" (falso). Damos el número real.
  const askingYappyNumber = /\b(cual\s+es|cu[aá]l\s+es|numero\s+(de|del|de\s+tu)?|num\s+(de|del)?|pasame|mandame|env[ií]ame|cual\s+el|me\s+das?|tu)\b[^?]{0,30}\byapp[yi]\b|^\s*yapp[yi]\s*\?|\byapp[yi]\s*\?/.test(t);
  // Pregunta sobre abono parcial / pago completo / descuento.
  const askingAbono = /\b(abon(o|ar|ando|amos)|adelanto|adelantar|reservar\s+abon|pagar\s+(completo|todo|el\s+total|en\s+total)|pago\s+(completo|total)|descuento)\b/.test(t);

  let body;
  if (askingAbono) {
    body =
      '💳 *Opciones de pago para tu reserva:*\n\n' +
      '• *Abono del 50%* para apartar la cabaña — el resto antes del día de tu llegada.\n' +
      '• *Pago completo con descuento*: *-$10* si reservas *dom–jue*, *-$20* si reservas *vie–sáb*.\n\n' +
      '*Métodos:* Yappy +507 6981-2266 (Joslyn Lopez) o ACH (Banco General).\n\n' +
      'No aceptamos tarjeta de crédito ni pago contra entrega.';
  } else if (askingYappyNumber) {
    body =
      '💳 *Yappy:* +507 6981-2266 a nombre de *Joslyn Lopez*.\n\n' +
      'También aceptamos *ACH* (Banco General — te paso los datos completos cuando confirmes la reserva).\n\n' +
      'No aceptamos tarjeta de crédito ni pago contra entrega.';
  } else {
    body =
      '💳 *Manejamos dos métodos de pago:*\n\n' +
      '• *Yappy* — +507 6981-2266 (Joslyn Lopez)\n' +
      '• *ACH* (transferencia bancaria — Banco General)\n\n' +
      'No aceptamos tarjeta de crédito ni pago contra entrega. Una vez confirmes la reserva te paso los datos completos. 🤝';
  }

  if (inBookingFlow) {
    body += '\n\n¿Listo para avanzar? Toca una opción arriba ⬆ y seguimos.';
  } else {
    body += '\n\n¿Tienes fechas en mente? Dime *fechas* y *personas* (ej: _"del 5 al 8 de junio, 2 personas"_) y te cotizo al instante.';
  }

  sendWhatsAppText(from, body);
  logDebugEntry('payment-info', { from: from, step: conv.step, inBookingFlow: inBookingFlow, askingYappyNumber: askingYappyNumber, askingAbono: askingAbono });
}

// Envia link al calendario publico para consultas de fechas vagas.
function _botSendCalendarLink(from, contactName) {
  const body =
    '🗓 Para fechas amplias o flexibles, puedes explorar todo el calendario de disponibilidad en nuestra página.\n\n' +
    'Toca el botón abajo, mira los días libres y cuando tengas fechas concretas dímelas por aquí (ej: _"del 5 al 8 de julio, 2 personas"_) para cotizar al instante. 🤝';
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
    '• *Paseo por Las Nubes* — hasta 4 personas (cama king + cama auxiliar twin)\n' +
    '• *Portal hacia Las Nubes* — 2 personas (cama matrimonial full)\n' +
    '• *Puente entre Las Nubes* — hasta 4 personas (cama queen + cama auxiliar full) · la más grande\n\n' +
    '✨ *Lo que incluye*\n' +
    '• Cocina equipada con área de BBQ y cooler grande\n' +
    '• Iluminación 100% solar — no hay luz eléctrica convencional\n' +
    '• Toallas, jabón y papel higiénico incluidos — agua fría\n' +
    '• Uso exclusivo de las instalaciones — sin vecinos\n\n' +
    '📅 *Para reservar o consultar disponibilidad*, cuéntame *fechas* y *personas*. Por ejemplo:\n' +
    '   _"del 5 al 8 de junio, 2 personas"_\n\n' +
    'Te envío disponibilidad y precio al instante. 🤝';

  sendWhatsAppText(from, info);
  _botSendMainMenu(from, contactName, false, '¿Quieres explorar más? Toca *Ver opciones* abajo 👇');
  _saveConv(from, 'AWAITING_DATES', {}, contactName);
}

// Tarifas base para 2 personas en las 3 cabanas (mismas tarifas, recargo por
// persona extra solo aplica a partir de 3 personas).
function _botSendPricingInfo(from, contactName, conv) {
  const msg =
    '💰 *Tarifas por noche · 2 personas*\n\n' +
    'Mismo precio en las tres cabañas (*Paseo*, *Portal* y *Puente*):\n\n' +
    _botTextoTarifas() + '\n' +
    '_Para 3 o 4 personas hay un pequeño recargo por persona adicional._\n\n' +
    '¿Quieres verificar disponibilidad para alguna fecha? Dime *fechas* y *personas* (ej: _"del 5 al 8 de junio, 2 personas"_) y te cotizo al instante. 🤝';
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

// ¿El mensaje sugiere interés en decoración (servicio +$40)?
function _botMentionsDecoracion(text) {
  const t = (text || '').toLowerCase();
  return /\b(decoraci[oó]n|decorar|decorad|sorpresa|globos?|flores|luna\s+de\s+miel|pedida\s+de\s+mano|propuesta|romant|rom[aá]ntic|cumple|cumplea[ñn]os|aniversario)\b/.test(t);
}

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
      '¿Quieres cotizar para alguna fecha? Dime *fechas* y *personas* 📅'
    );
    logDebugEntry('bot-fotos', { from: from, cabin: cabinKey });
    return;
  }
  // Sin cabaña específica → una foto de cada una + invitación a elegir
  ['verde', 'azul', 'lila'].forEach(c => {
    try { sendWhatsAppImage(from, BOT_CABIN_PHOTOS[c][0], '🏡 ' + BOT_CABIN_NAMES[c]); } catch(_) {}
  });
  sendWhatsAppText(from,
    '🌿 Estas son nuestras tres cabañas. Dime de cuál quieres ver *más fotos* (Paseo, Portal o Puente) ' +
    'o mira la galería completa en https://lasnubes.cloud/#cabanas\n\n' +
    'Y si ya tienes fechas en mente, te cotizo al instante 📅'
  );
  logDebugEntry('bot-fotos', { from: from, cabin: 'all' });
}

// Consultas de acceso: con qué auto se llega, o cómo llegar sin auto.
// Devuelve true si manejó el mensaje. Transporte (bus) se chequea primero.
function _botHandleAccesoQuery(from, contactName, text) {
  const t = (text || '').toLowerCase();

  // Sin auto / transporte / bus / Albrook / traslado / pickup
  if (/\b(sin\s+(auto|carro|veh[ií]culo|movilidad)|no\s+tengo\s+(auto|carro|veh[ií]culo|carro)|transporte|traslado|shuttle|pickup|en\s+bus\b|\bbus\b|autob[uú]s|albrook|terminal|recogida|nos\s+recog|me\s+recog)\b/.test(t)) {
    sendWhatsAppText(from,
      '🚌 *Cómo llegar sin auto*\n\n' +
      'Desde la terminal de *Albrook* puedes tomar cualquier bus hacia el interior y bajarte en el *Pío Pío de Bejuco*. Ahí te recogemos y trasladamos a la cabaña, ida y vuelta, por *$20 adicionales* en tu reserva. 🌿\n\n' +
      '¿Quieres cotizar para alguna fecha? Dime *fechas* y *personas* 📅'
    );
    logDebugEntry('bot-acceso', { from: from, tipo: 'sin-auto' });
    return true;
  }

  // Tipo de auto / 4x4 / sedán / estado del camino
  if (/\b(4\s*x\s*4|4x4|cuatro\s+por\s+cuatro|sed[aá]n|tipo\s+de\s+(auto|carro|veh)|cualquier\s+(auto|carro|veh|tipo\s+de\s+auto)|mi\s+(auto|carro)|camino\s+(es|est[aá]|de)|carretera\s+(es|est[aá]|de|en)|c[oó]mo\s+(es|est[aá])\s+(el\s+camino|la\s+carretera|la\s+calle)|se\s+puede\s+(ir|llegar|subir|entrar)\s+en|necesito\s+(un\s+)?4|hace\s+falta\s+(un\s+)?4|sube[ns]?\s+carros?)\b/.test(t)) {
    sendWhatsAppText(from,
      '🚗 *Acceso a Las Nubes*\n\n' +
      'La carretera es de asfalto y en muy buen estado hasta la entrada del proyecto. Luego son unos *5 minutos en camino de tosca fina*. Recibimos *sedanes y todo tipo de autos* sin problema. 🌿\n\n' +
      '¿Quieres cotizar para alguna fecha? Dime *fechas* y *personas* 📅'
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
  const tail = '\n\n¿Quieres cotizar para alguna fecha? Dime *fechas* y *personas* 📅';
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
      'No hay luz eléctrica convencional: la cabaña se ilumina 100% con paneles solares fotovoltaicos. ' +
      'Hay inversor para cargar celulares y dispositivos. ' +
      'Excelente señal de todas las operadoras.' + tail
    );
    return true;
  }
  if (/check.?in|checkin|check.?out|checkout|\bhora\s+(de\s+)?(llegada|entrada|salida)|a\s+qu[eé]\s+hora/i.test(t)) {
    sendWhatsAppText(from,
      '🕒 *Horarios*\n\n' +
      '• Check-in: *2:00 pm*\n' +
      '• Check-out: *11:00 am*\n\n' +
      'Si necesitas entrar más temprano o salir más tarde lo coordinamos según disponibilidad.' + tail
    );
    return true;
  }
  // El agua caliente es de las preguntas más frecuentes y la respuesta es que
  // NO hay, así que tiene que matchear sola: quien pregunta "¿hay agua
  // caliente?" no dice "baño" ni "toalla", y sin estos términos caía al
  // fallback genérico.
  if (/\b(toalla|jab[oó]n|papel|ba[ñn]o|amenidades|amenities|agua\s+caliente|ducha|regadera|calentador)\b/i.test(t)) {
    sendWhatsAppText(from,
      '🛁 *Baño & comodidades*\n\n' +
      'Jabón, papel higiénico y toallas limpias incluidos.\n\n' +
      '💧 *El agua de la ducha es fría* — no contamos con agua caliente en ninguna cabaña.\n\n' +
      'Fumigamos semanalmente — si eres sensible a mosquitos, te recomendamos traer repelente.' + tail
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
      '• *Paseo por Las Nubes* — hasta 4 personas (cama king + cama auxiliar twin)\n' +
      '• *Portal hacia Las Nubes* — 2 personas (cama matrimonial full)\n' +
      '• *Puente entre Las Nubes* — hasta 4 personas (cama queen + cama auxiliar full) · la más grande' + tail
    );
    return true;
  }
  if (/\b(mascot|pet|llevar\s+(mi\s+)?perr|mi\s+gat)\b/i.test(t)) {
    sendWhatsAppText(from,
      '🐾 *Mascotas*\n\n' +
      'Somos *pet friendly* 🐶 Hay un cargo de *$10 por reserva* y aceptamos hasta *2 mascotas*.\n\n' +
      'Para que la estadía sea agradable para todos: no pueden subir a la cama, se mantienen amarradas dentro de los jardines de tu cabaña, y te pedimos especial cuidado con sus necesidades y olores. 🙏'
    );
    return true;
  }

  // Decoracion / sorpresas para cumpleaños / aniversarios.
  // "decoración/globos/flores/sorpresa/propuesta" → siempre (piden el servicio).
  // "cumple/aniversario" SOLO si no hay fechas en el mensaje — si hay fechas,
  // es contexto de una reserva (ej: "mi pareja cumple el 10, quiero esa fecha")
  // y debe ir a cotización, no al pitch de decoración.
  const pideDecoracion  = /\b(decoraci[oó]n|decorar|decorad|sorpresa|globos?|flores|luna\s+de\s+miel|pedida\s+de\s+mano|propuesta|romant|rom[aá]ntic)\b/i.test(t);
  const mencionaOcasion = /\b(cumple|cumplea[ñn]os|aniversario)\b/i.test(t);
  if (pideDecoracion || (mencionaOcasion && !isDatesQuery)) {
    sendWhatsAppText(from,
      '🎉 *Decoración especial para aniversarios y cumpleaños*\n\n' +
      'Realizamos una decoración básica que incluye:\n' +
      '• Arreglo de flores\n' +
      '• Globos\n' +
      '• Letreros de cumpleaños o aniversarios\n' +
      '• Elementos decorativos románticos\n' +
      '• Una botella de espumante 🥂\n\n' +
      '*Costo:* $40 adicional a tu reserva.\n\n' +
      'Si quieres agregarlo, dímelo al confirmar las fechas y lo coordinamos. ¿Para cuándo sería?'
    );
    return true;
  }

  // Niños / familia. Si el mensaje trae fechas, lo dejamos al flujo de
  // cotización (el parser/Claude manejan personas y descuento de menores).
  if (!isDatesQuery && /\b(ni[ñn]o|ni[ñn]a|hijo|hija|hijos|hijas|beb[eé]|infantil|familia|family|kid|menor(es)?\s+de|aptas?\s+para\s+ni[ñn]os|family\s+friendly|chicos)\b/i.test(t)) {
    sendWhatsAppText(from,
      '👨‍👩‍👧 *Familias con niños*\n\n' +
      '¡Las cabañas son ideales para escapadas familiares! 🌿\n\n' +
      '🛏 *Camas* — *Puente* tiene cama *queen* + auxiliar *full*: perfecta para 2 adultos y 2 niños. *Paseo* tiene cama *king* + auxiliar *twin* (individual).\n\n' +
      '💰 *Política de niños*\n' +
      '• Menores de *5 años* no pagan.\n' +
      '• De la 3ra persona en adelante (5 años o más): $' + BOT_RECARGO_PERSONA_GRANDE + ' por persona/noche en Paseo y Puente, $' + BOT_RECARGO_PERSONA_PORTAL + ' en Portal.\n\n' +
      'Dime *fechas* y *cuántas personas* (incluye la edad de los niños) y te cotizo al instante 📅'
    );
    return true;
  }

  // Clima / temperatura
  if (/\b(clima|hace\s+(calor|fr[ií]o)|temperatura|qu[eé]\s+tiempo|humedad|caluroso|fresco|ventilador|aire\s+acondicionado|\bac\b)\b/i.test(t)) {
    sendWhatsAppText(from,
      '🌤 *Clima*\n\n' +
      'Las Nubes está en zona de montaña, con clima fresco — unos 4°C por debajo de la temperatura de la ciudad y brisa constante todo el día.\n\n' +
      'Cada cabaña tiene un ventilador pequeño, pero rara vez se usa. Te recomendamos traer una chaqueta liviana para las noches 🧥.\n\n' +
      '¿Quieres cotizar para alguna fecha? Dime *fechas* y *personas* 📅'
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
      '¿Quieres cotizar para alguna fecha? Dime *fechas* y *personas* 📅'
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
'- *Paseo por Las Nubes*: hasta 4 personas (cama king + cama auxiliar twin). Baño al aire libre entre árboles\n' +
'- *Portal hacia Las Nubes*: 2 personas (cama matrimonial full, sin cama auxiliar). Cocina exterior con vista a las montañas\n' +
'- *Puente entre Las Nubes*: hasta 4 personas (cama queen + cama auxiliar full). La más grande: ~120 m² de jardines y terraza techada apta para camping\n' +
'Las tres son independientes, privadas y de uso exclusivo de quienes reservan.\n\n' +
'## TARIFAS POR NOCHE (2 personas)\n' +
_botTextoTarifas().replace(/\*/g, '').replace(/^•/gm, '-') +
'- Mismo precio en las 3 cabañas\n' +
'- 3-4 personas: recargo pequeño por persona adicional ($' + BOT_RECARGO_PERSONA_PORTAL + ' en Portal, $' + BOT_RECARGO_PERSONA_GRANDE + ' en Paseo/Puente)\n' +
'- Niños menores de 5 años NO pagan\n' +
'- 5-6 personas: combo Puente + Portal (cabañas contiguas), cotización aparte\n\n' +
'## FAMILIAS CON NIÑOS\n' +
'- Puente: cama queen + cama auxiliar full (ideal para 2 adultos + 2 niños)\n- Paseo: cama king + cama auxiliar twin (la auxiliar es individual: 2 adultos + 1 niño)\n' +
'- Portal: cama matrimonial full (para 2 personas, no incluye espacio para niños extra)\n' +
'- Niños menores de 5 años no pagan\n' +
'- A partir de los 5 años aplican como persona adicional con recargo normal\n\n' +
'## AMENIDADES\n' +
'- Cocina completa + área de BBQ\n' +
'- Cooler grande (NO hay nevera — traer hielo)\n' +
'- Café, azúcar y especias básicas incluidos\n' +
'- Menú simple de comida disponible bajo reserva previa\n' +
'- Iluminación 100% solar + inversor para cargar celulares\n' +
'- Toallas, jabón y papel higiénico incluidos\n' +
'- Excelente señal de todas las operadoras\n' +
'- Fumigamos semanalmente (traer repelente si eres sensible a mosquitos)\n' +
'- NO hay agua caliente en ninguna cabaña — el agua de la ducha es fría\n' +
'- NO hay luz eléctrica convencional — la iluminación es 100% solar\n' +
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
'- SÍ se reciben mascotas (pet friendly): $10 por reserva, máximo 2\n' +
'- Reglas: no suben a la cama, amarradas dentro de los jardines de la cabaña, cuidado con necesidades y olores\n' +
'- Cancelaciones/cambios: coordinar con el equipo (derivá al agente)\n' +
'- Eventos especiales (cumpleaños, aniversarios, lunas de miel): bienvenidos, coordiná con el equipo\n\n' +
'## QUÉ HACER Y QUÉ NO\n' +
'- NO inventes datos, precios ni promos que no estén acá\n' +
'- NO confirmes disponibilidad de fechas específicas — eso requiere consultar sistema, pide *fechas* y *personas* y avisa que cotizas al instante\n' +
'- Si la pregunta es ambigua o requiere coordinación humana (eventos grandes, cambios de fecha, cancelaciones, descuentos), invita a tocar *Hablar con un agente* en el menú o escribir "3"\n' +
'- NO ofrezcas descuentos\n' +
'- NO menciones la cascada del proyecto por iniciativa propia. Está en ACTIVIDADES, no en amenidades, y se nombra SOLO si el huésped pregunta por actividades, qué hacer en la zona, o piscina/jacuzzi. Es decisión del anfitrión no ofrecerla sin que la pidan.\n\n' +
'## TONO\n' +
'- Cálido, breve, conversacional. Español neutral latinoamericano (usá "tú", no "vos")\n' +
'- Usa *negritas* de WhatsApp para énfasis y emojis con moderación (🌿 🏡 ☕ 🛏 📍)\n' +
'- Máximo 4-5 líneas. Directo al punto\n' +
'- Si la consulta es sobre disponibilidad/precios, termina con: "Dime *fechas* y *personas* (ej: _\'del 5 al 8 de junio, 2 personas\'_) y te cotizo al instante 🤝"\n' +
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

// Resuelve patrones de día-semana + día-mes SIN nombre de mes ("sábado 13 al
// domingo 14", "viernes 6 a sábado 7"). Caso real Malu: el parser no entendía
// "me cotiza desde sábado 13 a domingo 14?" porque le faltaba el mes.
// Asumimos el mes/año que produzca el match exacto día-semana + día-mes a
// partir de hoy. Devuelve {checkin, checkout, confidence} o null.
function _parseDatesWeekdayNoMonth(text, today) {
  const t = (text || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  const re = /\b(lunes|martes|miercoles|jueves|viernes|sabado|domingo)\s+(\d{1,2})\s*(?:a|al|hasta)\s*(lunes|martes|miercoles|jueves|viernes|sabado|domingo)\s+(\d{1,2})\b/;
  const m = re.exec(t);
  if (!m) return null;
  const DOW = ['domingo','lunes','martes','miercoles','jueves','viernes','sabado'];
  const dow1 = DOW.indexOf(m[1]);
  const dom1 = parseInt(m[2], 10);
  const dow2 = DOW.indexOf(m[3]);
  const dom2 = parseInt(m[4], 10);
  if (dow1 < 0 || dow2 < 0 || !dom1 || !dom2 || dom1 > 31 || dom2 > 31) return null;
  const baseDate = new Date(today + 'T12:00:00');
  // Busca primera fecha futura donde dayOfMonth=dom1 y dayOfWeek=dow1 (próx 90 días).
  for (let offset = 0; offset < 90; offset++) {
    const d = new Date(baseDate);
    d.setDate(d.getDate() + offset);
    if (d.getDate() !== dom1 || d.getDay() !== dow1) continue;
    // Encontramos checkin. Busca checkout (próximos 30 días) con dom2/dow2.
    for (let k = 1; k <= 30; k++) {
      const co = new Date(d);
      co.setDate(co.getDate() + k);
      if (co.getDate() === dom2 && co.getDay() === dow2) {
        return {
          checkin:  Utilities.formatDate(d,  BOT_TZ, 'yyyy-MM-dd'),
          checkout: Utilities.formatDate(co, BOT_TZ, 'yyyy-MM-dd'),
          confidence: 0.9
        };
      }
    }
    return null;   // checkin OK pero checkout no calza → no asumir
  }
  return null;
}

// ─── Parser determinista de fechas explícitas ─────────────────────
// Resuelve formatos numéricos con mes ("del 5 al 8 de junio", "2 de junio",
// "martes 02 junio") sin depender del LLM. Devuelve null si no aplica
// (relativos como "este finde" o si menciona niños → los maneja Claude).
function _parseDatesDeterministic(text, today) {
  const t = (text || '').toLowerCase();
  // Niños → dejar a Claude (maneja descuento de menores de 5)
  if (/\b(ni[ñn]o|ni[ñn]a|beb[eé]|hijo|hija|menor(es)?\s+de)\b/.test(t)) return null;

  // Pattern especial: día-semana + día-mes sin mes ("sábado 13 al domingo 14").
  const weekdayResult = _parseDatesWeekdayNoMonth(text, today);
  if (weekdayResult) return weekdayResult;

  const MESES = {
    enero:0, febrero:1, marzo:2, abril:3, mayo:4, junio:5,
    julio:6, agosto:7, septiembre:8, setiembre:8, octubre:9, noviembre:10, diciembre:11
  };
  const MONTH = '(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)';

  const ymd = (y, mo, d) => y + '-' + String(mo + 1).padStart(2, '0') + '-' + String(d).padStart(2, '0');
  const resolveYear = (mo, d) => {
    const td = new Date(today + 'T12:00:00');
    let y = td.getFullYear();
    const cand = new Date(y, mo, d, 12);
    const todayMid = new Date(td.getFullYear(), td.getMonth(), td.getDate(), 12);
    if (cand < todayMid) y += 1;   // fecha ya pasó este año → próximo año
    return y;
  };
  const validDay = (d) => d >= 1 && d <= 31;

  // Rango mismo mes: "del 5 al 8 de junio" / "del 2 al 3 junio"
  let m = t.match(new RegExp('del?\\s+(\\d{1,2})\\s+al\\s+(\\d{1,2})\\s+(?:de\\s+)?' + MONTH));
  if (m) {
    const d1 = parseInt(m[1], 10), d2 = parseInt(m[2], 10), mo = MESES[m[3]];
    if (validDay(d1) && validDay(d2) && d2 > d1) {
      const y = resolveYear(mo, d1);
      return { checkin: ymd(y, mo, d1), checkout: ymd(y, mo, d2), persons: _botExtractPersons(t), confidence: 1 };
    }
  }

  // "10 y 11 de octubre" → típicamente entra 10, sale 11 (1 noche), igual que
  // el rango. Es ambiguo (a veces quieren ambas noches) → marcamos para
  // aclarar con una nota al cotizar.
  m = t.match(new RegExp('(\\d{1,2})\\s+y\\s+(\\d{1,2})\\s+(?:de\\s+)?' + MONTH));
  if (m) {
    const d1 = parseInt(m[1], 10), d2 = parseInt(m[2], 10), mo = MESES[m[3]];
    if (validDay(d1) && validDay(d2) && d2 > d1) {
      const y = resolveYear(mo, d1);
      return { checkin: ymd(y, mo, d1), checkout: ymd(y, mo, d2), persons: _botExtractPersons(t), confidence: 1, ambiguousNights: true };
    }
  }

  // Fecha única: "[para el] [martes] 02 [de] junio" / "2 de junio"
  m = t.match(new RegExp('(\\d{1,2})\\s+(?:de\\s+)?' + MONTH));
  if (m) {
    const d1 = parseInt(m[1], 10), mo = MESES[m[2]];
    if (validDay(d1)) {
      const y = resolveYear(mo, d1);
      const ci = ymd(y, mo, d1);
      return { checkin: ci, checkout: _botAddDaysISO(ci, 1), persons: _botExtractPersons(t), confidence: 1 };
    }
  }

  return null;
}

// Extrae cantidad de personas de frases como "3 personas", "dos adultos".
function _botExtractPersons(t) {
  const NUM = { una: 1, dos: 2, tres: 3, cuatro: 4, cinco: 5, seis: 6 };
  let m = t.match(/(\d+)\s*(personas?|adultos?|grandes?|pax|hu[eé]spedes?)\b/);
  if (m) return parseInt(m[1], 10);
  m = t.match(/\b(una|dos|tres|cuatro|cinco|seis)\s+(personas?|adultos?|grandes?)\b/);
  if (m) return NUM[m[1]];
  return null;
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

// Desglose noche por noche de un rango: para cada noche, qué cabañas (de
// capacidad suficiente) están libres. Usado para ofrecer noches sueltas
// cuando el rango completo no está disponible (B2).
function _botNightsBreakdown(checkin, checkout, personas) {
  personas = personas || 2;
  const out = [];
  let cur = checkin, guard = 0;
  while (cur < checkout && guard < 60) {
    const next = _botAddDaysISO(cur, 1);
    const avail = _botCheckAvailability(cur, next);
    const freeCabins = ['azul', 'verde', 'lila'].filter(c => avail[c] && BOT_CABIN_CAPACITY[c] >= personas);
    out.push({ night: cur, next: next, free: freeCabins.length > 0, freeCabins: freeCabins });
    cur = next;
    guard++;
  }
  return out;
}

// Delega en el tarifario real (Parser.gs · precioNochePublico). Antes tenía su
// propia tabla de dos precios: en una víspera de feriado cotizaba $90 cuando el
// calendario cobra $135, y durante una promoción cotizaba $90 cuando el precio
// vigente era $75. El bot y el sitio deben decir el mismo número.
function _botPrecioPorNoche(dateStr) {
  try {
    return precioNochePublico(dateStr);
  } catch (e) {
    // Si Config o la hoja de feriados fallan, cotizar de menos sería peor que
    // cotizar el fallback histórico.
    logDebugEntry('bot-precio-fallback', { fecha: dateStr, error: e.message });
    const dow = new Date(dateStr + 'T12:00:00').getDay();
    return (dow === 5 || dow === 6) ? BOT_RATE_WEEKEND : BOT_RATE_WEEKDAY;
  }
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
      'Toca el botón abajo para escribir al equipo.'
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
    // Si el mismo mensaje ya trae fechas (ej: "quiero info, reservar para
    // el 10 y 11 de octubre"), cotizamos directo en vez del pitch genérico.
    if (_looksLikeDateQuery(text) && !_looksLikeVagueDateQuery(text)) {
      let parsed = _parseDatesDeterministic(text, _botToday());
      if (!parsed || !parsed.checkin) parsed = _parseDatesWithClaude(text, _botToday());
      if (parsed && parsed.checkin && parsed.checkout && (parsed.confidence === undefined || parsed.confidence > 0.4)) {
        return _replyAvailability(from, contactName, { step: 'AWAITING_DATES', context: {}, name: contactName },
          parsed.checkin, parsed.checkout, parsed.persons || 2, parsed.freeChildren || 0, parsed.ambiguousNights, _botMentionsDecoracion(text));
      }
    }
    return _botSendCampaignWelcome(from, contactName);
  }

  // Boton "Ver otras fechas" → vuelve a AWAITING_DATES
  if (kind === 'button_reply' && text === 'try_dates') {
    sendWhatsAppText(from, '🌿 Dime las nuevas fechas:\n\n• "del 5 al 8 de junio, 2 personas"\n• "viernes a domingo, 4 personas"');
    _saveConv(from, 'AWAITING_DATES', conv.context, contactName);
    return;
  }

  // Cambiar cantidad de personas → reconsulta disponibilidad con nuevas N
  if ((kind === 'list_reply' || kind === 'button_reply') && /^persons_(\d)$/.test(text)) {
    const n = parseInt(text.replace('persons_', ''), 10);
    const dates = conv.context && conv.context.dates;
    if (!dates || !dates.checkin || !dates.checkout) {
      sendWhatsAppText(from, '🤔 Perdí el contexto. Dime las fechas otra vez (ej: "del 5 al 8 de junio").');
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

  // Boton "Mantener abierta" del recordatorio de ventana 24h al admin.
  // El tap solo sirve para registrar el inbound (que renueva la ventana en
  // ADMIN_LAST_INBOUND_TS). Acá solo confirmamos al admin.

// ─── Anti doble-toque ───────────────────────────────────────────
//
// El guard del webhook (_yaProcesado) deduplica por `wamid` y funciona: en
// todo el historial no hay un solo wamid repetido después de que se desplegó.
// Pero NO cubre este caso, porque son mensajes genuinamente distintos:
//
//   17:13:40  button_reply  acceso_178535…   wamid …A8A42A93
//   17:13:41  button_reply  acceso_178535…   wamid …4C38A6E7A   ← otro id
//
// El huésped tocó el botón dos veces con un segundo de diferencia —lo normal
// cuando la respuesta tarda— y el bot ejecutó la acción dos veces.
//
// Hay una segunda forma del mismo problema: la MISMA acción alcanzable por dos
// payloads distintos. "He llegado" existe como botón de la plantilla
// (`llegada_<id>`) y como ítem del menú (`menu_he_llegado`); tocar los dos
// mandaba las instrucciones de bienvenida dos veces. Por eso la clave se
// normaliza a la ACCIÓN, no al payload.
const _ACCION_COOLDOWN_SEG = 90;

function _accionRepetida(from, accion) {
  if (!from || !accion) return false;
  try {
    const cache = CacheService.getScriptCache();
    const key   = 'wa-accion-' + from + '-' + accion;
    if (cache.get(key)) {
      logDebugEntry('WA-accion-repetida', { from: from, accion: accion });
      return true;
    }
    cache.put(key, '1', _ACCION_COOLDOWN_SEG);
    return false;
  } catch (e) {
    // Ante un fallo del guard se EJECUTA: dejar al huésped sin respuesta es
    // peor que mandarle algo dos veces.
    return false;
  }
}

// Traduce un payload a la acción que dispara. Dos payloads distintos que hacen
// lo mismo comparten clave.
function _accionDePayload(text) {
  const t = String(text || '');
  if (t.indexOf('llegada_') === 0 || t === 'menu_he_llegado') return 'llegada';
  if (t.indexOf('acceso_') === 0)  return 'acceso';
  if (t.indexOf('manual_') === 0)  return 'manual';
  if (t.indexOf('llamar_') === 0)  return 'llamar';
  if (t.indexOf('ubicacion_') === 0) return 'ubicacion';
  if (t.indexOf('referido_') === 0)  return 'referido';
  if (/^menu_/.test(t)) return t;   // cada ítem de menú es su propia acción
  return '';
}

  if (kind === 'button_reply' && text === 'admin_keep_window') {
    try { sendWhatsAppText(from, '✓ Ventana renovada por 24h. Las alertas operativas siguen llegando.'); } catch(_) {}
    return;
  }

  // Doble toque: se ignora la repetición de la MISMA acción dentro del
  // cooldown. Va antes de todo el despacho de botones para cubrirlos a todos
  // de una, en vez de acordarse handler por handler.
  if (kind === 'button_reply' || kind === 'list_reply') {
    const _accion = _accionDePayload(text);
    if (_accion && _accionRepetida(from, _accion)) return;
  }

  // Boton "He llegado" de la plantilla del día de llegada (listos_para_recibirte,
  // 11am). Mismo destino que el "He llegado" del menú: instrucciones de acceso
  // + alerta al admin para que abra el portón.
  if (kind === 'button_reply' && text.indexOf('llegada_') === 0) {
    return _botMenuHeLlegado(from, contactName, conv);
  }

  // Botón "Como funciona" de la plantilla del código de referido. El cuerpo de la
  // plantilla solo trae el titular; el detalle se manda acá, desde la misma
  // fuente que el email (REFERRAL_COMO_FUNCIONA / REFERRAL_RESTRICCIONES).
  if (kind === 'button_reply' && (text.indexOf('referido_') === 0 || /^como\s+funciona$/i.test(text))) {
    return _botHandleReferidoInfo(from, contactName);
  }

  // Botones de las instrucciones de llegada.
  if (kind === 'button_reply' && text.indexOf('acceso_') === 0) {
    return _botHandleCodigoAcceso(from, contactName, conv, text.replace('acceso_', ''));
  }
  if (kind === 'button_reply' && text.indexOf('manual_') === 0) {
    return _botHandleManualCabana(from, contactName, text.replace('manual_', ''));
  }
  if (kind === 'button_reply' && text.indexOf('llamar_') === 0) {
    return _botHandleLlamarJosh(from);
  }

  // Boton "Envíame ubicación" de la plantilla de check-in (recordator_entrada)
  // → mandar ubicación + cómo llegar. Match por payload o por el texto del
  //   botón (fallback si la plantilla se envió sin payload dinámico).
  if (kind === 'button_reply' && (text.indexOf('ubicacion_') === 0 || /env[ií]ame\s+ubicaci[oó]n/i.test(text))) {
    return _botHandleEnviarUbicacion(from, contactName);
  }

  // Menu list/button reply → handler especifico
  if ((kind === 'list_reply' || kind === 'button_reply') && /^menu_/.test(text)) {
    if (text === 'menu_disponibilidad') {
      sendWhatsAppText(from,
        '¡Genial! 🌿 Dime *fechas* y *personas* (ej: _"del 5 al 8 de junio, 2 personas"_).'
      );
      _saveConv(from, 'AWAITING_DATES', conv.context, contactName);
      return;
    }
    if (text === 'menu_he_llegado')   { _botMenuHeLlegado(from, contactName, conv); return; }
    if (text === 'menu_abrir_porton') {
      // Misma lógica que tocar "Abrir el portón" en la plantilla
      // instruccion_checkout. _botHandleCheckoutDone resuelve la reserva
      // por teléfono si no se pasa id.
      return _botHandleCheckoutDone(from, contactName, '');
    }
    if (text === 'menu_como_llegar')  { _botMenuComoLlegar(from);  _botSendMainMenu(from, contactName, false); _saveConv(from, 'SHOWED_INFO', conv.context, contactName); return; }
    if (text === 'menu_actividades')  { _botMenuActividades(from); _botSendMainMenu(from, contactName, false); _saveConv(from, 'SHOWED_INFO', conv.context, contactName); return; }
    if (text === 'menu_gastronomia')  { _botMenuGastronomia(from); _botSendMainMenu(from, contactName, false); _saveConv(from, 'SHOWED_INFO', conv.context, contactName); return; }
    if (text === 'menu_insumos')      { _botMenuInsumos(from);     _botSendMainMenu(from, contactName, false); _saveConv(from, 'SHOWED_INFO', conv.context, contactName); return; }
    if (text === 'menu_tienda')       { _botMenuTienda(from);      _botSendMainMenu(from, contactName, false); _saveConv(from, 'SHOWED_INFO', conv.context, contactName); return; }
    if (text === 'menu_faq')          { _botMenuFAQ(from);         _botSendMainMenu(from, contactName, false); _saveConv(from, 'SHOWED_INFO', conv.context, contactName); return; }
    if (text === 'menu_agente' || text === 'menu_humano') {
      // Enriquecemos con el contexto que haya (cabaña, fechas, personas).
      const ctx       = conv.context || {};
      const cabinName = BOT_CABIN_NAMES[ctx.cabin] || '';
      const dts       = ctx.dates;
      const fechas    = (dts && dts.checkin) ? (_botFmtFecha(dts.checkin) + (dts.checkout ? ' al ' + _botFmtFecha(dts.checkout) : '')) : '';
      const personas  = ctx.personas || '';

      let ctxDesc = '';
      if (cabinName && fechas) ctxDesc = 'Estuve viendo ' + cabinName + ' para el ' + fechas + (personas ? ' (' + personas + ' personas)' : '');
      else if (cabinName)      ctxDesc = 'Estuve viendo ' + cabinName;
      else if (fechas)         ctxDesc = 'Estuve consultando para el ' + fechas + (personas ? ' (' + personas + ' personas)' : '');
      const prefill = 'Hola, vengo del Agente de Las Nubes 🌿. ' + (ctxDesc ? ctxDesc + '. ' : '') + 'Quiero hablar con una persona.';

      try {
        sendWhatsAppCTAUrl(from,
          '🙋 ¡Claro! Toca el botón abajo para escribirle directo a una persona de nuestro equipo por WhatsApp.',
          'Abrir WhatsApp',
          'https://wa.me/50769812266?text=' + encodeURIComponent(prefill)
        );
      } catch(err) {
        sendWhatsAppText(from, '🙋 Escribinos directo aquí:\nhttps://wa.me/50769812266');
      }
      _saveConv(from, 'HUMAN_HANDOFF', conv.context, contactName);
      try {
        let adminMsg = '🔔 *El Agente derivó a un cliente*\n👤 ' + (contactName || from) + '\n📱 +' + from;
        if (cabinName) adminMsg += '\n🏡 ' + cabinName;
        if (fechas)    adminMsg += '\n📅 ' + fechas;
        if (personas)  adminMsg += '\n👥 ' + personas;
        _botAdminAlert('handoff', adminMsg);
      } catch(_) {}
      return;
    }
  }

  // Eleccion de cierre: autoservicio vs asistido (estado CHOOSING_CLOSE)
  // Eleccion de decoración (estado CHOOSING_DECOR)
  if (kind === 'button_reply' && text === 'deco_yes') {
    conv.context = Object.assign({}, conv.context, { decoracion: true });
    return _botShowCloseChoice(from, contactName, conv);
  }
  if (kind === 'button_reply' && text === 'deco_no') {
    conv.context = Object.assign({}, conv.context, { decoracion: false });
    return _botShowCloseChoice(from, contactName, conv);
  }
  if (conv.step === 'CHOOSING_DECOR') {
    const td = (text || '').trim().toLowerCase();
    if (td === '1' || /con\s+decor|s[ií]\b|decora/.test(td))  { conv.context = Object.assign({}, conv.context, { decoracion: true });  return _botShowCloseChoice(from, contactName, conv); }
    if (td === '2' || /sin\s+decor|no\b/.test(td))            { conv.context = Object.assign({}, conv.context, { decoracion: false }); return _botShowCloseChoice(from, contactName, conv); }
  }

  // Eleccion de cierre: autoservicio vs asistido (estado CHOOSING_CLOSE)
  if (kind === 'button_reply' && text === 'close_self')   return _botOfferPayment(from, contactName, conv);
  if (kind === 'button_reply' && text === 'close_asesor') return _botCloseWithAsesor(from, contactName, conv);
  if (conv.step === 'CHOOSING_CLOSE') {
    const tt = (text || '').trim().toLowerCase();
    if (tt === '1' || /reservar\s+ahora|autoservicio|autogest/.test(tt)) return _botOfferPayment(from, contactName, conv);
    if (tt === '2' || /asistid|asesor|josh/.test(tt))                    return _botCloseWithAsesor(from, contactName, conv);
  }

  // Boton "Cancelar" en el cierre/pago — libera el estado + muestra menu principal
  if ((kind === 'button_reply' && text === 'cancel_booking') ||
      ((conv.step === 'OFFERING_PAYMENT' || conv.step === 'CHOOSING_CLOSE' || conv.step === 'CHOOSING_DECOR') && /^(cancela|cancelar|atras|atrás)\b/i.test((text || '').trim()))) {
    sendWhatsAppText(from, '👋 Listo, cancelamos esta reserva. ¿En qué más te ayudamos?');
    _saveConv(from, 'INITIAL', {}, contactName);
    _botSendMainMenu(from, contactName, true);
    return;
  }

  // Boton "Cotiza esa" (noche libre de un rango parcialmente ocupado, B2)
  if (kind === 'button_reply' && /^quote_\d{4}-\d{2}-\d{2}_\d{4}-\d{2}-\d{2}$/.test(text)) {
    const parts = text.replace('quote_', '').split('_');
    const personas     = (conv.context && conv.context.personas) || 2;
    const freeChildren = (conv.context && conv.context.freeChildren) || 0;
    return _replyAvailability(from, contactName, conv, parts[0], parts[1], personas, freeChildren);
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
    // La cédula va primero: si no, una foto enviada durante el flujo de acceso
    // caería en el parser de vouchers y le pediría "monto y referencia".
    if (conv.step === 'AWAITING_CEDULA') {
      return _botHandleCedulaImage(from, text, contactName, conv);
    }
    return _botHandleVoucherImage(from, text, contactName, conv);
  }

  // He llegado: esperando nombre del titular para ubicar reserva
  if (conv.step === 'AWAITING_ARRIVAL_NAME') {
    const tName = (text || '').trim();
    if (tName.length < 3) {
      sendWhatsAppText(from, '🤔 Necesito el nombre completo para ubicar la reserva. Prueba de nuevo o escríbeme "agente" para hablar con una persona.');
      return;
    }
    const reservaByName = _botFindReservaByName(tName);
    if (!reservaByName) {
      sendWhatsAppText(from,
        '😔 No encuentro una reserva activa a nombre de *' + tName + '* para hoy.\n\n' +
        'Te derivo con una persona del equipo para resolverlo. Escribime "agente" si quieres contactarla directo.'
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

  // Insistencia en una fecha sin disponibilidad → derivar a Josh.
  // (El cliente acaba de ver "sin disponibilidad" / alternativas y vuelve a
  // pedir esa misma fecha o muestra apego emocional → oportunidad de cierre
  // humano: lista de espera, mover otra reserva, etc.)
  if ((conv.step === 'NO_AVAILABILITY' || conv.step === 'SHOWING_ALTERNATIVES') && _botIsDateInsistence(text)) {
    return _botHandleDateInsistence(from, contactName, conv);
  }

  // Cambio de fecha / cancelacion / "no podre ir" → handoff directo al admin
  // (no intenta cotizar aunque haya fechas en el texto).
  if (_isReservaChangeRequest(text)) {
    sendWhatsAppText(from,
      '🙏 Entiendo, te derivo con una persona del equipo para coordinar el cambio o cancelación.\n\n' +
      'Toca el botón abajo para escribirle directo.'
    );
    try {
      sendWhatsAppCTAUrl(from,
        'Te resolvemos el cambio en un mensaje.',
        '💬 Abrir WhatsApp',
        'https://wa.me/50769812266?text=' + encodeURIComponent('Hola, quiero coordinar un cambio o cancelación de mi reserva.')
      );
    } catch(_) {}
    _botAdminAlert('handoff',
      '⚠️ Cambio/cancelación pedido vía Agente:\n👤 ' + (contactName || from) + '\n📱 +' + from + '\nMensaje: "' + (text || '').slice(0, 200) + '"');
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

  // Cliente con reserva activa (hoy / mañana / estadía / futura / pasada):
  // si su conversación no está en mitad de un flujo (booking, He llegado,
  // handoff), mostrar el menú contextual personalizado en vez de seguir al
  // parser de fechas o al smart-fallback de Claude. Resuelve casos confusos
  // como "Buenas días joven" o "Tengo una reserva para hoy" → terminaban
  // en "No entendí las fechas" o en respuesta genérica de Claude.
  // Flag once-per-day en context para no re-mandar el menú en cada mensaje.
  const _midFlowSteps = [
    'SHOWING_AVAILABILITY','SHOWING_ALTERNATIVES',
    'CHOOSING_DECOR','CHOOSING_CLOSE',
    'OFFERING_PAYMENT','AWAITING_VOUCHER_RETRY',
    'AWAITING_EMAIL','AWAITING_NAME',
    'PENDING_REVIEW','PENDING_HUMAN_BOOKING',
    'AWAITING_ARRIVAL_NAME','HUMAN_HANDOFF'
  ];
  if (kind === 'text' && _midFlowSteps.indexOf(conv.step) === -1) {
    const _ctxAM = conv.context || {};
    const _todayAM = _botToday();
    if (_ctxAM._arrivalMenuShown !== _todayAM) {
      try {
        const arrival = _botArrivalStatus(from);
        if (arrival) {
          _ctxAM._arrivalMenuShown = _todayAM;
          _saveConv(from, 'SHOWED_INFO', _ctxAM, contactName);
          _botSendMainMenu(from, contactName, true);
          logDebugEntry('arrival-context-menu', { from: from, status: arrival.status, step: conv.step });
          return;
        }
      } catch(_) {}
    }
  }

  // Preguntas sobre métodos de pago durante el flujo de booking. Caso real:
  // cliente en SHOWING_AVAILABILITY preguntó "Para hacer el pago?" y el bot
  // dumpeaba "Dime fechas y personas" rompiendo todo el contexto. Ahora
  // respondemos breve y mantenemos el estado.
  if (kind === 'text' && conv.step !== 'OFFERING_PAYMENT' && _botLooksLikePaymentQuery(text)) {
    _botHandlePaymentInfo(from, contactName, conv, text);
    return;
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
  const midBooking = ['CHOOSING_DECOR', 'CHOOSING_CLOSE', 'OFFERING_PAYMENT', 'AWAITING_VOUCHER_RETRY', 'AWAITING_EMAIL', 'AWAITING_NAME'].indexOf(conv.step) !== -1;
  if (_looksLikeDateQuery(text)) {
    // Primero parser determinista (formatos explícitos con mes); si no
    // aplica, recurrimos al NLU de Claude (relativos, lenguaje natural).
    let parsed = _parseDatesDeterministic(text, _botToday());
    if (!parsed || !parsed.checkin) parsed = _parseDatesWithClaude(text, _botToday());
    if (parsed && parsed.checkin && parsed.checkout && (parsed.confidence === undefined || parsed.confidence > 0.4)) {
      const personas     = parsed.persons || 2;
      const freeChildren = parsed.freeChildren || 0;
      if (midBooking) {
        sendWhatsAppText(from, '🔄 Veo que quieres cambiar las fechas. Verifico disponibilidad para las nuevas...');
      }
      return _replyAvailability(from, contactName, { step: 'AWAITING_DATES', context: {}, name: contactName }, parsed.checkin, parsed.checkout, personas, freeChildren, parsed.ambiguousNights, _botMentionsDecoracion(text));
    }
    // Parsing fallo. Si el cliente mencionó personas o noches sueltas
    // (sin fechas concretas), mostramos tarifas como fallback util.
    const hasPersons = /\d+\s*personas?\b|\b(una|dos|tres|cuatro|cinco|seis)\s+personas?\b/i.test(text);
    const hasNoches  = /\d+\s*noches?\b|\b(una|dos|tres|cuatro|cinco)\s+noches?\b/i.test(text);
    if (hasPersons || hasNoches) {
      _botSendPricingInfo(from, contactName, conv);
      return;
    }
    sendWhatsAppText(from, '🤔 No logré entender las fechas. ¿Puedes escribirlas más claras?\n\nEjemplo: "del 5 al 8 de junio, 4 personas".');
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
function _replyAvailability(from, contactName, conv, checkin, checkout, personas, freeChildren, ambiguousNights, wantsDecoracion) {
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
      'Toca el botón abajo para escribir al equipo y resolverlo en un mensaje.'
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
    // Consulta de grupo grande: es un handoff/derivación, NO un cierre.
    _saveConv(from, 'HUMAN_HANDOFF', { dates: dates, personas: personas, freeChildren: freeChildren }, contactName);
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
    // B2: disponibilidad parcial — si el rango es multi-noche y solo algunas
    // noches están ocupadas, ofrecemos las noches libres en vez de un "no" seco.
    if (nights > 1) {
      const bd = _botNightsBreakdown(checkin, checkout, personas);
      const someFree = bd.some(n => n.free), someOcc = bd.some(n => !n.free);
      if (someFree && someOcc) {
        // bloque contiguo libre más largo
        let bStart = -1, bLen = 0, cStart = -1, cLen = 0;
        bd.forEach((n, i) => {
          if (n.free) { if (cStart < 0) cStart = i; cLen++; if (cLen > bLen) { bLen = cLen; bStart = cStart; } }
          else { cStart = -1; cLen = 0; }
        });
        const freeCi = bd[bStart].night, freeCo = bd[bStart + bLen - 1].next;
        const occ = bd.filter(n => !n.free).map(n => _botFmtFecha(n.night));
        const libreLbl = (bLen === 1)
          ? 'la noche del *' + _botFmtFecha(freeCi) + '*'
          : 'del *' + _botFmtFecha(freeCi) + '* al *' + _botFmtFecha(freeCo) + '*';
        const body =
          '😔 No tengo las *' + nights + ' noches* completas para ' + personas + (personas === 1 ? ' persona' : ' personas') + '.\n\n' +
          'La noche del *' + occ.join('*, *') + '* ya está reservada.\n\n' +
          'Pero sí tengo libre ' + libreLbl + '. ¿Te sirve?';
        try {
          sendWhatsAppButtons(from, body, [
            { id: 'quote_' + freeCi + '_' + freeCo, title: '✅ Cotiza esa' },
            { id: 'menu_agente', title: '🙋 Hablar con Josh' }
          ]);
        } catch(_) {
          sendWhatsAppText(from, body + '\n\nEscríbeme si te sirve o "agente" para hablar con Josh.');
        }
        _saveConv(from, 'SHOWING_ALTERNATIVES', { dates: { checkin: freeCi, checkout: freeCo }, personas: personas, freeChildren: freeChildren }, contactName);
        return;
      }
    }
    const alts = _botSuggestAlternatives(checkin, checkout, personas);
    if (alts.length > 0) {
      const body =
        '😔 No tenemos disponibilidad para *' + fechasStr + '* con ' + personas + (personas === 1 ? ' persona' : ' personas') + '.\n\n' +
        'Pero sí tenemos para estas fechas cercanas:';
      const buttons = alts.slice(0, 3).map(a => ({ id: 'alt_' + a.checkin, title: _botFmtFecha(a.checkin) }));
      try {
        sendWhatsAppButtons(from, body, buttons, null, 'Toca una opción o escríbeme "agente"');
      } catch(_) {
        sendWhatsAppText(from, body + '\n\n' + alts.map(a => '• ' + _botFmtFecha(a.checkin) + ' → ' + _botFmtFecha(a.checkout)).join('\n') + '\n\nEscribime las fechas que prefieras.');
      }
      _saveConv(from, 'SHOWING_ALTERNATIVES', { dates: dates, personas: personas, freeChildren: freeChildren, alts: alts }, contactName);
      return;
    }
    sendWhatsAppText(from,
      '😔 No tenemos disponibilidad para *' + fechasStr + '* con ' + personas + (personas === 1 ? ' persona' : ' personas') + '.\n\n' +
      'Puedes ver el calendario público:\nhttps://lasnubes.cloud\n\n' +
      '¿O prefieres hablar con un agente? Toca *Hablar con un agente* en el menú.'
    );
    _saveConv(from, 'NO_AVAILABILITY', { dates: dates, personas: personas, freeChildren: freeChildren }, contactName);
    return;
  }

  // Cotizacion (formato copyPromo, sin combo)
  const cotizacion = _botCotizacionAvailability(checkin, checkout, opciones, personas, false, freeChildren);
  sendWhatsAppText(from, cotizacion);

  // Nota aclaratoria si la fecha vino en formato ambiguo "10 y 11" (1 noche).
  if (ambiguousNights && nights === 1) {
    sendWhatsAppText(from,
      '📝 Lo coticé como *1 noche* (entras el ' + _botFmtFecha(checkin) + ' y sales el ' + _botFmtFecha(checkout) + '). ' +
      'Si querías quedarte *ambas noches*, dímelo y lo ajusto.'
    );
  }

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
    sendWhatsAppList(from, '¿Quieres cambiar algo? Toca ⬇', sections, '📋 Ver opciones');
  } catch(_) { /* ignorable: ya tiene los botones de cabaña */ }

  _saveConv(from, 'SHOWING_AVAILABILITY', { dates: dates, personas: personas, freeChildren: freeChildren, opciones: opciones.length, wantsDecoracion: !!wantsDecoracion }, contactName);
}

// ─── Menu principal interactivo (lista) ──────────────────────────
// firstTime=true → muestra bienvenida elaborada. firstTime=false → solo "¿Necesitas algo más?"
// customBody (opcional) → reemplaza el cuerpo (uso desde flujos especiales como campaign).
function _botSendMainMenu(from, contactName, firstTime, customBody) {
  let firstName = ((contactName || '').toString().trim().split(/\s+/)[0]) || '';

  // Detectar contexto de reserva (hoy / mañana) para personalizar el saludo
  // y reordenar el menú según relevancia. Si hay customBody, respetar.
  let arrival = null;
  if (!customBody) {
    try { arrival = _botArrivalStatus(from); } catch(_) {}
  }

  // Si hay reserva activa con nombre real, preferirlo sobre el alias de
  // WhatsApp (que puede ser emojis o un display name raro como "💜🥳").
  if (arrival && arrival.reserva && arrival.reserva.name) {
    const reservaFirst = String(arrival.reserva.name).trim().split(/\s+/)[0];
    if (reservaFirst) firstName = reservaFirst;
  }

  const greeting = firstName ? '¡Hola ' + firstName + '! 🌿' : '¡Hola! 🌿';

  let body, sections;
  const isFirst = firstTime !== false;
  if (customBody) {
    body = customBody;
    sections = _botMenuSectionsDefault();
  } else if (arrival && arrival.status === 'hoy') {
    body = _botBuildBodyHoy(arrival.reserva, firstName, isFirst);
    sections = _botMenuSectionsHoy();
  } else if (arrival && arrival.status === 'saliendo') {
    body = _botBuildBodySaliendo(arrival.reserva, firstName, isFirst);
    sections = _botMenuSectionsSaliendo();
  } else if (arrival && arrival.status === 'estadia') {
    body = _botBuildBodyEstadia(arrival.reserva, firstName, isFirst);
    sections = _botMenuSectionsEstadia();
  } else if (arrival && arrival.status === 'manana') {
    body = _botBuildBodyManana(arrival.reserva, firstName, isFirst);
    sections = _botMenuSectionsManana();
  } else if (arrival && arrival.status === 'futura') {
    body = _botBuildBodyFutura(arrival.reserva, firstName, isFirst);
    sections = _botMenuSectionsFutura();
  } else if (arrival && arrival.status === 'pasada') {
    body = _botBuildBodyPasada(arrival.reserva, firstName, isFirst);
    sections = _botMenuSectionsPasada();
  } else if (!isFirst) {
    body = '¿Necesitas algo más? Toca *Ver opciones* abajo 👇';
    sections = _botMenuSectionsDefault();
  } else {
    body = greeting + '\n\n' +
      'Bienvenido a *Las Nubes* — un refugio de tres cabañas en las faldas del Cerro Chicá, a 1h 15min de la ciudad.\n\n' +
      'Para reservar o consultar disponibilidad, cuéntame *fechas* y *personas*. Por ejemplo:\n' +
      '   _"del 5 al 8 de junio, 2 personas"_\n\n' +
      'Te envío disponibilidad y precio al instante, y cerramos la reserva por aquí mismo. 🤝\n\n' +
      '¿Quieres explorar antes — cómo llegar, actividades, fotos o hablar con un agente? Toca *Ver opciones* ⬇';
    sections = _botMenuSectionsDefault();
  }
  try {
    sendWhatsAppList(from, body, sections, 'Ver opciones', null, 'Buenos Aires, Chamé · Panamá');
  } catch(err) {
    logDebugEntry('bot-menu-FAIL', { error: err.message });
    sendWhatsAppText(from, body + '\n\nEscribime qué te interesa:\n📅 Disponibilidad · 📍 Cómo llegar · 🏞 Actividades · 🍽 Gastronomía · 🛒 Insumos · 🧊 Hielo y carbón · ❓ FAQ · 🙋 Agente');
  }
}

// ─── Builders del menú principal por contexto ──────────────────────
function _botMenuSectionsDefault() {
  return [
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
}

function _botMenuSectionsHoy() {
  return [
    {
      title: 'Tu llegada',
      rows: [
        { id: 'menu_he_llegado',  title: '🚪 He llegado',    description: 'Estoy en el portón de entrada' },
        { id: 'menu_como_llegar', title: '📍 Cómo llegar',   description: 'Maps, Waze, indicaciones' }
      ]
    },
    {
      title: 'En la cabaña',
      rows: [
        { id: 'menu_insumos', title: '🛒 Insumos',        description: 'Tiendas para abastecerse' },
        { id: 'menu_tienda',  title: '🧊 Hielo y carbón', description: 'Tienda a 5 min de la cabaña' }
      ]
    },
    {
      title: 'Información',
      rows: [
        { id: 'menu_actividades', title: '🏞 Actividades',          description: 'Cascadas, playas, cerros' },
        { id: 'menu_gastronomia', title: '🍽 Gastronomía',           description: 'Restaurantes cerca' },
        { id: 'menu_faq',         title: '❓ Preguntas frecuentes',  description: 'Cocina, energía, check-out' },
        { id: 'menu_agente',      title: '🙋 Hablar con un agente',  description: 'Abrir WhatsApp del equipo' }
      ]
    }
  ];
}

function _botMenuSectionsManana() {
  return [
    {
      title: 'Para tu llegada',
      rows: [
        { id: 'menu_como_llegar', title: '📍 Cómo llegar',   description: 'Dirección, Waze, Maps' },
        { id: 'menu_insumos',     title: '🛒 Insumos',        description: 'Tiendita y supermercados' },
        { id: 'menu_tienda',      title: '🧊 Hielo y carbón', description: 'Tienda a 5 min de la cabaña' }
      ]
    },
    {
      title: 'Para planear',
      rows: [
        { id: 'menu_actividades', title: '🏞 Actividades',  description: 'Cascadas, playas, cerros' },
        { id: 'menu_gastronomia', title: '🍽 Gastronomía',  description: 'Restaurantes cerca' }
      ]
    },
    {
      title: 'Soporte',
      rows: [
        { id: 'menu_faq',    title: '❓ Preguntas frecuentes', description: 'Acceso, cocina, key box' },
        { id: 'menu_agente', title: '🙋 Hablar con un agente', description: 'Abrir WhatsApp del equipo' }
      ]
    }
  ];
}

function _botMenuSectionsSaliendo() {
  return [
    {
      title: 'Tu salida',
      rows: [
        { id: 'menu_abrir_porton', title: '🚪 Abrir el portón',     description: 'Estoy en el portón de salida' },
        { id: 'menu_faq',          title: '❓ Preguntas frecuentes', description: 'Cocina, basura, key box, check-out' }
      ]
    },
    {
      title: 'Soporte',
      rows: [
        { id: 'menu_agente', title: '🙋 Hablar con un agente', description: 'Abrir WhatsApp del equipo' }
      ]
    }
  ];
}

function _botMenuSectionsEstadia() {
  return [
    {
      title: 'En la cabaña',
      rows: [
        { id: 'menu_tienda',  title: '🧊 Hielo y carbón', description: 'Tienda a 5 min de la cabaña' },
        { id: 'menu_insumos', title: '🛒 Insumos',        description: 'Supermercados cercanos' }
      ]
    },
    {
      title: 'Actividades',
      rows: [
        { id: 'menu_actividades', title: '🏞 Actividades',  description: 'Cascadas, playas, cerros' },
        { id: 'menu_gastronomia', title: '🍽 Gastronomía',  description: 'Restaurantes cerca' }
      ]
    },
    {
      title: 'Soporte',
      rows: [
        { id: 'menu_faq',    title: '❓ Preguntas frecuentes', description: 'Cocina, energía, key box, check-out' },
        { id: 'menu_agente', title: '🙋 Hablar con un agente', description: 'Abrir WhatsApp del equipo' }
      ]
    }
  ];
}

function _botMenuSectionsFutura() {
  return [
    {
      title: 'Para tu llegada',
      rows: [
        { id: 'menu_como_llegar', title: '📍 Cómo llegar',   description: 'Dirección, Waze, Maps' },
        { id: 'menu_insumos',     title: '🛒 Insumos',        description: 'Tiendita y supermercados' },
        { id: 'menu_tienda',      title: '🧊 Hielo y carbón', description: 'Tienda a 5 min de la cabaña' }
      ]
    },
    {
      title: 'Para planear',
      rows: [
        { id: 'menu_actividades', title: '🏞 Actividades',  description: 'Cascadas, playas, cerros' },
        { id: 'menu_gastronomia', title: '🍽 Gastronomía',  description: 'Restaurantes cerca' }
      ]
    },
    {
      title: 'Cambios y soporte',
      rows: [
        { id: 'menu_faq',    title: '❓ Preguntas frecuentes', description: 'Acceso, cocina, key box' },
        { id: 'menu_agente', title: '🙋 Hablar con un agente', description: 'Para cambios o consultas' }
      ]
    }
  ];
}

function _botMenuSectionsPasada() {
  return [
    {
      title: 'Próxima reserva',
      rows: [
        { id: 'menu_disponibilidad', title: '📅 Disponibilidad', description: 'Ver fechas libres y precios' }
      ]
    },
    {
      title: 'Información',
      rows: [
        { id: 'menu_actividades', title: '🏞 Actividades',  description: 'Cascadas, playas, cerros' },
        { id: 'menu_gastronomia', title: '🍽 Gastronomía',  description: 'Restaurantes cerca' }
      ]
    },
    {
      title: 'Soporte',
      rows: [
        { id: 'menu_faq',    title: '❓ Preguntas frecuentes', description: '' },
        { id: 'menu_agente', title: '🙋 Hablar con un agente', description: 'Abrir WhatsApp del equipo' }
      ]
    }
  ];
}

function _botBuildBodyHoy(reserva, firstName, isFirstTime) {
  const greeting = firstName ? '¡Hola ' + firstName + '! 🌿' : '¡Hola! 🌿';
  const fechas   = _fechasRangoCorto(reserva.displayCheckin, reserva.displayCheckout);
  if (!isFirstTime) {
    return '¿Necesitas algo para tu llegada de hoy a *' + reserva.cabinName + '*? Toca *Ver opciones* abajo 👇';
  }
  return greeting + '\n\n' +
    'Hoy te recibimos en *' + reserva.cabinName + '* para tu reserva del ' + fechas + '.\n\n' +
    'Cuando llegues al portón, toca *🚪 He llegado* abajo y te abrimos al instante. 🚪';
}

function _botBuildBodyManana(reserva, firstName, isFirstTime) {
  const greeting = firstName ? '¡Hola ' + firstName + '! 🌿' : '¡Hola! 🌿';
  const fechas   = _fechasRangoCorto(reserva.displayCheckin, reserva.displayCheckout);
  if (!isFirstTime) {
    return '¿Necesitas info para tu llegada de mañana a *' + reserva.cabinName + '*? Toca *Ver opciones* abajo 👇';
  }
  return greeting + '\n\n' +
    'Mañana te recibimos en *' + reserva.cabinName + '* para tu reserva del ' + fechas + '.\n\n' +
    '¿Necesitas info para tu llegada? Toca *Ver opciones* abajo 👇 (cómo llegar, qué llevar, actividades, etc.)\n\n' +
    // El recordatorio que llega mañana es el de las 11am (enviarAvisoLlegadaHoy).
    // Decía "10am" — esa es la hora del recordatorio de HOY, el del día anterior,
    // así que prometía un mensaje que no llegaba.
    '_Mañana a las 11am te escribimos de nuevo con la hora de check-in y el botón para avisarnos cuando llegues._';
}

function _botBuildBodySaliendo(reserva, firstName, isFirstTime) {
  const greeting = firstName ? '¡Hola ' + firstName + '! 🌿' : '¡Hola! 🌿';
  const checkoutHr = _horaPlantilla(reserva.tipo, 'checkout', reserva.checkoutExtendido, null, reserva.horaSalida);
  if (!isFirstTime) {
    return '¿Necesitas algo para tu salida hoy de *' + reserva.cabinName + '*? Toca *Ver opciones* abajo 👇';
  }
  return greeting + '\n\n' +
    'Hoy es tu check-out de *' + reserva.cabinName + '* a las *' + checkoutHr + '*.\n\n' +
    'Antes de salir de la cabaña: deja la cocina ordenada y la llave *y el control ' +
    'del portón* en el key box. 🔑 La basura la puedes dejar, el personal se encarga.\n\n' +
    // "te abrimos al instante" prometía algo que el bot no hace: manda la alerta
    // y el portón lo abre el admin a mano. Mismo arreglo que en la llegada.
    'Cuando estén en el portón de salida, toca *🚪 Abrir el portón* abajo y le avisamos ' +
    'al equipo para que les abra.';
}

function _botBuildBodyEstadia(reserva, firstName, isFirstTime) {
  const greeting = firstName ? '¡Hola ' + firstName + '! 🌿' : '¡Hola! 🌿';
  if (!isFirstTime) {
    return '¿Necesitas algo durante tu estadía en *' + reserva.cabinName + '*? Toca *Ver opciones* abajo 👇';
  }
  return greeting + '\n\n' +
    '¡Esperamos que estés disfrutando tu estadía en *' + reserva.cabinName + '*!\n\n' +
    '¿En qué te puedo ayudar? Toca *Ver opciones* abajo 👇';
}

function _botBuildBodyFutura(reserva, firstName, isFirstTime) {
  const greeting = firstName ? '¡Hola ' + firstName + '! 🌿' : '¡Hola! 🌿';
  const fechas   = _fechasRangoCorto(reserva.displayCheckin, reserva.displayCheckout);
  if (!isFirstTime) {
    return '¿Necesitas info sobre tu reserva en *' + reserva.cabinName + '* (' + fechas + ')? Toca *Ver opciones* abajo 👇';
  }
  return greeting + '\n\n' +
    'Tienes reserva con nosotros para el *' + fechas + '* en *' + reserva.cabinName + '*. 🌿\n\n' +
    '¿Quieres info para tu llegada o ajustar algo? Toca *Ver opciones* abajo 👇';
}

function _botBuildBodyPasada(reserva, firstName, isFirstTime) {
  const greeting = firstName ? '¡Hola ' + firstName + '! 🌿' : '¡Hola! 🌿';
  if (!isFirstTime) {
    return '¿Quieres hacer otra escapada a Las Nubes? Toca *Ver opciones* abajo 👇';
  }
  return greeting + '\n\n' +
    '¡Esperamos que hayas disfrutado tu estadía en *' + reserva.cabinName + '*! 🙌\n\n' +
    '¿Quieres agendar tu próxima escapada? Dime *fechas* y *personas* y te cotizo al instante. 🤝';
}

function _botMenuComoLlegar(from) {
  sendWhatsAppText(from,
    '📍 *Cómo llegar a Las Nubes*\n\n' +
    'Por la carretera Interamericana, entra por el *Pío Pío de Bejuco* a la carretera Bejuco–Sorá. ' +
    'Al llegar al pueblo de *Buenos Aires*, dobla a la derecha hacia *Chicá*. La cabaña queda a 100 metros.\n\n' +
    '🚗 Lo más fácil: pon en *Waze "Aires de Chicá"* — te lleva directo al portón verde. ' +
    'Cuando llegues, escríbeme o llámame para abrir y guiarte a la cabaña.\n\n' +
    '🗺 Google Maps:\nhttps://maps.google.com/?q=8.639400,-79.945900\n\n' +
    '🚦 Waze (abre con navegación):\nhttps://waze.com/ul?ll=8.639400,-79.945900&navigate=yes'
  );
}

function _botMenuTienda(from) {
  sendWhatsAppText(from,
    '🧊 *Tienda de conveniencia cercana*\n\n' +
    'Contamos con una tienda a tan solo *5 minutos* de la cabaña. En ella puedes encontrar:\n\n' +
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
    '¿Otra duda? Toca *Hablar con persona* o escríbeme "3".'
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

// Días entre dos fechas ISO yyyy-MM-dd. Positivo si isoB > isoA.
function _botDaysBetween(isoA, isoB) {
  const a = new Date(isoA + 'T12:00:00');
  const b = new Date(isoB + 'T12:00:00');
  return Math.round((b - a) / 86400000);
}

// Detecta el contexto de reserva del cliente para personalizar el saludo y
// el menú. Devuelve { status, reserva } o null. Status:
//   - 'hoy'      : check-in es hoy (display).
//   - 'manana'   : check-in es mañana (display).
//   - 'saliendo' : check-out es hoy (estadía multi-día que termina hoy).
//   - 'estadia'  : hoy entre check-in y check-out (sin incluir el día de salida).
//   - 'futura'   : check-in en >2 días.
//   - 'pasada'   : check-out fue hace ≤7 días.
// Si hay varias reservas, prioriza: hoy > saliendo > estadia > manana >
// futura > pasada.
function _botArrivalStatus(phone) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const today    = _botToday();
  const tomorrow = _botAddDaysISO(today, 1);
  const normalize = (t) => {
    let d = String(t || '').replace(/\D/g, '');
    if (d.indexOf('507') === 0 && d.length > 8) d = d.substring(3);
    return d;
  };
  const target = normalize(phone);
  if (!target) return null;
  const found = { hoy: null, saliendo: null, estadia: null, manana: null, futura: null, pasada: null };
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if ((r[20] || '').toString().toUpperCase() === 'CANCELADA') continue;
    if (r[9] === 'Abierta') continue;
    if (normalize(r[23]) !== target) continue;
    const ci = r[4] instanceof Date ? Utilities.formatDate(r[4], BOT_TZ, 'yyyy-MM-dd') : (r[4] || '').toString().slice(0,10);
    const co = r[5] instanceof Date ? Utilities.formatDate(r[5], BOT_TZ, 'yyyy-MM-dd') : (r[5] || '').toString().slice(0,10);
    if (!ci || !co) continue;
    const tipo = (r[24] || 'noche').toString();
    // Mismo mapeo storage→display que tipoEmailMeta / _formFromStored.
    let displayCi = ci, displayCo = co;
    if (tipo === 'pasadia')        { displayCi = _botAddDaysISO(ci, 1); displayCo = displayCi; }
    else if (tipo === 'pasatarde') { displayCo = ci; }
    else if (tipo === 'early')     { displayCi = _botAddDaysISO(ci, 1); }
    else if (tipo === 'late')      { displayCo = _botAddDaysISO(co, -1); }
    let status = null;
    if (displayCi === today)                                  status = 'hoy';
    else if (displayCi === tomorrow)                          status = 'manana';
    else if (today === displayCo && today > displayCi)        status = 'saliendo';
    else if (today > displayCi && today < displayCo)          status = 'estadia';
    else if (displayCi > tomorrow)                            status = 'futura';
    else if (displayCo < today && _botDaysBetween(displayCo, today) <= 7) status = 'pasada';
    if (!status) continue;
    const reserva = {
      id: r[0], name: r[1], cabin: r[3],
      checkin: ci, checkout: co,
      persons: r[6], origin: r[9],
      tipo: tipo,
      checkoutExtendido: !!r[28],
      horaEntrada: (typeof _normalizeHora === 'function') ? _normalizeHora(r[29]) : (r[29] || ''),
      horaSalida:  (typeof _normalizeHora === 'function') ? _normalizeHora(r[30]) : (r[30] || ''),
      mascotas:    !!r[33],
      displayCheckin: displayCi, displayCheckout: displayCo,
      cabinName: BOT_CABIN_NAMES[r[3]] || r[3]
    };
    if (!found[status]) {
      found[status] = reserva;
    } else if (status === 'futura' && displayCi < found.futura.displayCheckin) {
      found.futura = reserva;   // próxima futura más cercana
    } else if (status === 'pasada' && displayCo > found.pasada.displayCheckout) {
      found.pasada = reserva;   // estadía pasada más reciente
    }
  }
  const priority = ['hoy', 'saliendo', 'estadia', 'manana', 'futura', 'pasada'];
  for (const s of priority) if (found[s]) return { status: s, reserva: found[s] };
  return null;
}

function _botMenuHeLlegado(from, contactName, conv) {
  const reserva = _botFindReservaByPhone(from);
  if (reserva) {
    return _botSendArrivalInstructions(from, contactName, conv, reserva);
  }
  // No match por telefono → preguntar nombre del titular
  sendWhatsAppText(from,
    '🌿 Recibí tu mensaje.\n\n' +
    'No encuentro una reserva activa con este número para hoy. Dime el *nombre completo del titular* de la reserva para ubicarla en el sistema.\n\n' +
    'Si prefieres hablar directo con una persona, escríbeme "agente".'
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

// Fotos de cada cabaña para confirmar la llegada. Se sirven desde el propio
// repo via GitHub Pages: URL estable, versionada con el código y sin depender
// de que el permiso de un archivo de Drive siga siendo público. Están
// redimensionadas a 1600px de lado mayor (~450 KB) porque en el portón la
// señal es mala y una foto de 9 MB no baja — WhatsApp además topa en 5 MB.
const BOT_CABIN_LLEGADA_FOTO = {
  verde: 'https://lasnubes.cloud/paseo.jpg',
  azul:  'https://lasnubes.cloud/portal.jpg',
  lila:  'https://lasnubes.cloud/puente.jpg'
};

// Desde que doblan a la izquierda hasta la puerta. El tramo anterior es común
// a las tres y vive en _botSendArrivalInstructions.
//
// Paseo y Puente comparten el estacionamiento (el poste de luz) y se separan
// recién al pie de las escaleras rústicas, así que ese trozo está una sola vez.
const _LLEGADA_PARQUEO_POSTE =
  'Deténganse ahí mismo. Unos *10 metros más adelante* pueden estacionar en el ' +
  'lado izquierdo de la calle, *antes del poste de luz*. Es una calle de muy poco ' +
  'tráfico, así que el auto queda bien.\n\n';

function _botLlegadaTramoCabana(cabin) {
  // Portal sigue la misma estructura que las otras dos —detenerse, estacionar
  // con la tranquilidad de que el auto queda bien, bajar, reconocer, key box
  // junto a la puerta— con su referencia propia: el tanque de agua azul en vez
  // del poste de luz. La puerta acá es corrediza de METAL, no blanca.
  if (cabin === 'azul') {
    return 'Deténganse ahí mismo. Unos *25 metros más adelante* verán un *tanque de ' +
           'reserva de agua azul*. Pueden estacionar *antes del tanque*, en los laterales ' +
           'de la calle. Es una calle de muy poco tráfico, así que el auto queda bien.\n\n' +
           'Justo al lado del tanque está la *escalera para bajar a la cabaña*. Tiene un ' +
           '*techo blanco* y la reconocerán por los portales con las puertas antiguas y ' +
           'la silla colgante.\n\n' +
           'Al lado de la *puerta corrediza de metal* está el *key box*.';
  }
  if (cabin === 'lila') {
    return _LLEGADA_PARQUEO_POSTE +
           'Al lado izquierdo del poste de luz verán unas *escaleras rústicas*. ' +
           'Apenas las bajen, la cabaña es la que está *al frente*.\n\n' +
           'La reconocerán por la *terraza con las dos hamacas*, luego la pérgola con ' +
           'la cocina exterior y, al final, la recámara con la *puerta corrediza blanca* ' +
           'donde está el *key box*.';
  }
  // verde (Paseo)
  return _LLEGADA_PARQUEO_POSTE +
         'Al lado izquierdo del poste de luz verán unas *escaleras rústicas*. ' +
         'Apenas las bajen, *doblen INMEDIATAMENTE a mano izquierda*.\n\n' +
         'Después se encontrarán con unas escaleras largas que llevan a la única cabaña ' +
         'que está abajo, incrustada en la montaña. La reconocerán por los *dos columpios* ' +
         'y la *malla suspendida* al frente.\n\n' +
         'Cuando lleguen se van a olvidar de las escaleras: la privacidad es única y la ' +
         'vista, fenomenal. 🌄\n\n' +
         'Al lado de la *puerta corrediza blanca* está el *key box*.';
}

// `opts.preview` corta los efectos hacia afuera: no dispara la alerta de portón
// al admin ni toca el estado de la conversación. Sin esto, revisar los textos
// de las tres cabañas le manda al admin tres "ABRE EL PORTÓN" falsos — y el día
// que llegue uno real, ya aprendió a ignorarlos.
function _botSendArrivalInstructions(from, contactName, conv, reserva, opts) {
  const preview = !!(opts && opts.preview);
  const cabin     = reserva.cabin;
  const cabinName = BOT_CABIN_NAMES[cabin] || 'Las Nubes';
  const firstName = ((reserva.name || contactName || '').toString().trim().split(/\s+/)[0]) || '';

  let body = '🎉 ¡Bienvenidos a *Las Nubes*';
  if (firstName) body += ', ' + firstName;
  // Antes decía "Ya les abro el portón", y el bot NO abre nada: manda la alerta
  // `alerta_porton` al admin, que abre a mano. Si el admin está durmiendo o sin
  // señal, el huésped quedaba parado en el portón creyendo que ya se estaba
  // abriendo. El flujo de salida siempre lo dijo bien ("ya le avisé al equipo");
  // este ahora dice lo mismo.
  // Todo el mensaje va en USTEDES: al portón se llega en grupo, y el texto
  // mezclaba las dos formas en tres líneas seguidas ("para que TE abran",
  // "conducen", "LLAMA a Josh"). Es el término local: una *calle huella* son
  // las dos franjas de concreto; "huella calle" estaba invertido.
  body += '!\n\nYa le avisé al equipo para que les abran el portón. 🚪\n\n' +
          'Apenas se abra, conducen recto y más adelante se encontrarán con una ' +
          '*calle huella de concreto*. Van a subirla y, cuando termine, van a tomar ' +
          'la siguiente *calle a mano izquierda*.\n\n' +
          _botLlegadaTramoCabana(cabin) +
          // Las reglas de mascota van acá y no en el manual solamente: es el
          // momento en que bajan del auto con el perro. Solo si la reserva las
          // tiene, para no darle normas de mascotas a quien no trajo ninguna.
          (reserva.mascotas
            ? '\n\n🐾 *Con tu mascota:* no puede subir a la cama y se mantiene '
              + 'amarrada dentro de los jardines de la cabaña. Cuidado con sus '
              + 'necesidades y olores. 🙏'
            : '') +
          '\n\nCualquier dificultad, llamen a Josh.';

  sendWhatsAppText(from, body);

  // Segundo mensaje: foto de la cabaña como encabezado + las tres acciones.
  // La foto va acá y no en el mensaje de arriba a propósito: si la imagen no
  // baja (señal mala en el portón), las indicaciones ya llegaron igual.
  const foto = BOT_CABIN_LLEGADA_FOTO[cabin];
  const botones = [
    { id: 'acceso_' + reserva.id, title: '🔑 Código de acceso' },
    { id: 'manual_' + reserva.id, title: '📖 Manual de cabaña' },
    { id: 'llamar_' + reserva.id, title: '📞 Llamar a Josh'   }
  ];
  const cuerpoBotones =
    'Esta es *' + cabinName + '*.\n\n' +
    '*Asegúrense de que la cabaña sea la de la foto* antes de entrar. ¿Qué necesitan?';
  try {
    sendWhatsAppButtons(from, cuerpoBotones, botones, foto ? { imageUrl: foto } : null);
  } catch(err) {
    // Sin foto o sin botones (WhatsApp viejo): el link público sigue dando
    // acceso al código y al manual, así que no queda a ciegas.
    logDebugEntry('arrival-buttons-FAIL', { error: err.message, cabin: cabin });
    let publicUrl = '';
    try { publicUrl = getPublicReservaUrl(reserva.id); } catch(_) {}
    sendWhatsAppText(from, cuerpoBotones +
      (publicUrl ? '\n\n🔗 *Código de acceso y manual:*\n' + publicUrl : ''));
  }

  if (preview) {
    logDebugEntry('llegada-preview', { cabin: cabin, to: from });
    return;
  }

  // Notificar al admin via plantilla HSM (alerta_porton) — pasa siempre,
  // sin depender de la ventana de 24h.
  _sendAlertaPorton('entrada', reserva.name || contactName, from, cabinName);

  _saveConv(from, 'ARRIVED', Object.assign({}, conv.context || {}, { reservaId: reserva.id }), contactName);
}

// Busca la próxima reserva no cancelada en una cabaña, con check-in (display)
// hoy o después. Útil para que el alert a Erika anticipe si la próxima reserva
// requiere cama auxiliar.
function _botFindNextReservationForCabin(cabinKey, excludeReservaId) {
  if (!cabinKey) return null;
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const today = _botToday();
  let best = null;
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if (excludeReservaId && r[0] === excludeReservaId) continue;
    if (r[3] !== cabinKey) continue;
    if ((r[20] || '').toString().toUpperCase() === 'CANCELADA') continue;
    if (r[9] === 'Abierta') continue;
    const ci = r[4] instanceof Date ? Utilities.formatDate(r[4], BOT_TZ, 'yyyy-MM-dd') : (r[4] || '').toString().slice(0,10);
    if (!ci) continue;
    const tipo = (r[24] || 'noche').toString();
    let displayCi = ci;
    if (tipo === 'pasadia' || tipo === 'pasadia-largo' || tipo === 'early') displayCi = _botAddDaysISO(ci, 1);
    if (displayCi < today) continue;   // ya pasó
    const next = {
      persons: parseInt(r[6], 10) || 0,
      comentarios: (r[22] || '').toString(),
      // Sin el tipo no se puede decidir la cama auxiliar: una pasadía no la
      // lleva. Ver _botTextoCamaAuxiliar. Y sin horaEntrada, _horaPlantilla
      // devuelve el default del tipo y se pierde el override de la reserva.
      tipo: tipo,
      horaEntrada: (typeof _normalizeHora === 'function') ? _normalizeHora(r[29]) : (r[29] || ''),
      displayCheckin: displayCi
    };
    if (!best || displayCi < best.displayCheckin) best = next;
  }
  return best;
}

// Necesita cama auxiliar la próxima reserva? Personas >= 3 (primario) o los
// comentarios mencionan cama auxiliar/cuna (fallback si el campo de personas
// no refleja la realidad).
function _botNeedsCamaAuxiliar(reserva) {
  if (!reserva) return false;
  if ((reserva.persons || 0) >= 3) return true;
  return _botComentarioPideCama(reserva);
}

// Solo la parte de comentarios. Separada porque en las cabañas sin cama
// auxiliar el motivo cambia el mensaje: 3+ personas es un acuerdo previo
// (el cliente trae colchón), un pedido escrito no lo es.
function _botComentarioPideCama(reserva) {
  const lower = String((reserva && reserva.comentarios) || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  return /cama auxiliar|cama adicional|cama extra|preparar cama|cuna/.test(lower);
}

// Cabañas que NO tienen cama auxiliar. Portal no tiene: cuando una reserva de
// 3+ personas cae ahí es porque el cliente acordó traer colchón inflable, así
// que pedir "preparar cama auxiliar" sería mandar a buscar algo que no existe.
const CABANAS_SIN_CAMA_AUXILIAR = { azul: true };   // azul = Portal hacia Las Nubes

// Estadías de un solo día (pasadía / pasadía largo / pasatarde): nadie duerme,
// así que no llevan cama auxiliar. Ojo, `pasanoche` NO entra acá — ese sí es
// una noche. Y no llevar cama auxiliar no significa no limpiar: usan la
// recámara igual, así que el cambio de sábanas va completo.
function _esEstadiaUnDia(tipo) {
  const t = (tipo || '').toString();
  return t === 'pasadia' || t === 'pasadia-largo' || t === 'pasatarde';
}

// Texto del aviso de cama para una reserva, o '' si no aplica. Fuente única
// para el parte diario de limpieza y para la plantilla alerta_limpieza: si cada
// una decidiera por su cuenta, terminarían diciéndole cosas distintas a Erika
// sobre la misma reserva.
function _botTextoCamaAuxiliar(reserva, cabinKey) {
  if (!reserva) return '';
  if (_esEstadiaUnDia(reserva.tipo)) return '';
  const porPersonas   = (reserva.persons || 0) >= 3;
  const porComentario = _botComentarioPideCama(reserva);
  if (!porPersonas && !porComentario) return '';
  if (!CABANAS_SIN_CAMA_AUXILIAR[cabinKey]) return 'Preparar cama auxiliar.';
  // En una cabaña sin cama auxiliar el motivo importa. Con 3+ personas está
  // acordado de antemano que el cliente trae colchón inflable. Si el pedido
  // sale de un comentario (una cuna, una cama extra) no hay tal acuerdo, así
  // que no se puede afirmar que traiga nada: se avisa para confirmarlo.
  return porPersonas
    ? 'Sin cama auxiliar en esta cabaña: el cliente trae colchón inflable.'
    : 'Ojo: piden cama/cuna en los comentarios y esta cabaña no tiene cama auxiliar. Confírmalo con Josh.';
}

// Texto del {{3}} de la plantilla alerta_limpieza. Siempre devuelve algo
// (Meta no acepta variables vacías) y en una sola línea (Meta rechaza los
// params con saltos de línea).
function _botBuildLimpiezaContextLine(nextReserva, cabinKey) {
  if (!nextReserva) return '📅 Sin próxima reserva agendada por ahora.';
  // Cuándo y a qué hora llega el siguiente. Es lo primero que Erika necesita al
  // recibir esta alerta: sin eso no puede saber si tiene el día entero o dos
  // horas. El parte diario ya lo daba; acá faltaba.
  const cuando = _botCuandoLlega(nextReserva.displayCheckin);
  const hora   = _horaPlantilla(nextReserva.tipo, 'checkin', false, nextReserva.horaEntrada);
  const p      = parseInt(nextReserva.persons, 10) || 0;
  const quien  = p ? ', ' + p + (p === 1 ? ' huésped' : ' huéspedes') : '';
  const base   = 'Próxima reserva: ' + cuando + ' a las ' + hora + quien + '.';
  const cama   = _botTextoCamaAuxiliar(nextReserva, cabinKey);
  return cama ? '🛏 ' + cama + ' ' + base : '📅 ' + base;
}

// "HOY" / "mañana" / "sáb 2 ago". En mayúscula el caso urgente: si llega hoy,
// esa es la única palabra que cambia lo que Erika hace a continuación.
function _botCuandoLlega(iso) {
  if (!iso) return 'sin fecha';
  const hoy = _botToday();
  if (iso === hoy) return 'HOY';
  if (iso === _botAddDaysISO(hoy, 1)) return 'mañana';
  return _botFmtFecha(iso);
}

// Envía la plantilla HSM alerta_porton al admin. Llega siempre, sin importar
// el estado de la ventana de 24h.
function _sendAlertaPorton(direction, guestName, guestPhone, cabinName) {
  try {
    sendWhatsAppTemplate(BOT_ADMIN_PHONE, 'alerta_porton', 'es_ES', [
      String(direction || ''),
      String(guestName || '?'),
      String(guestPhone || '').replace(/\D/g, ''),
      String(cabinName || '?')
    ], null, null);
  } catch(e) {
    logDebugEntry('alerta-porton-FAIL', { direction: direction, error: e.message });
  }
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
        telefono: r[23] || '',
        // Col 27: si ya hay ID subido, el botón "Código de acceso" no vuelve a
        // pedir la cédula. Cols 29/31: el manual muestra la hora de check-out
        // real (cortesía o custom), no el default del tipo.
        idHuespedURL: r[26] || '',
        checkoutExtendido: r[28] === true || r[28] === 'TRUE' || r[28] === 'true' || r[28] === 1,
        horaSalida: (typeof _normalizeHora === 'function') ? _normalizeHora(r[30]) : (r[30] || ''),
        mascotas:   !!r[33]
      };
    }
  }
  return null;
}

// El huésped tocó "Como funciona" en la plantilla del código de referido.
function _botHandleReferidoInfo(from, contactName) {
  // Por teléfono y no por reserva: getOrCreateReferralCode exige email (devuelve
  // null sin él) y la búsqueda por reserva solo alcanza hasta un día después del
  // check-out, pero el botón se puede tocar una semana más tarde.
  let codigo = '';
  try { codigo = findReferralCodeByPhone(from) || ''; } catch(_) {}
  if (!codigo) {
    // Sin match por teléfono no podemos resolver su código, pero las reglas son
    // las mismas para todos: mejor responder eso que un "no te encuentro".
    sendWhatsAppText(from,
      referralReglasTexto('el que te enviamos por correo', REFERRAL_REWARD_AMOUNT));
    return;
  }
  sendWhatsAppText(from, referralReglasTexto(codigo, REFERRAL_REWARD_AMOUNT));
}

// ═══════════════════════════════════════════════════════════
//  Botones de las instrucciones de llegada
//  🔑 Código de acceso · 📖 Manual de cabaña · 📞 Llamar a Josh
// ═══════════════════════════════════════════════════════════

// El id viaja en el payload del botón, pero una reserva vieja o un tap desde
// otro chat pueden dejarlo vacío: se cae al match por teléfono, igual que el
// resto de los handlers de llegada.
function _botResolverReservaLlegada(from, reservaId) {
  let reserva = _botFindReservaById(reservaId);
  if (!reserva) { try { reserva = _botFindReservaByPhone(from); } catch(_) {} }
  return reserva;
}

// El manual sale de getCabinGuideSteps() —la MISMA fuente que la página pública
// y el email— convertido a texto de WhatsApp. Sin fuente única, el manual del
// bot y el de la página se separan en la primera corrección que se haga a uno.
// `omitirAcceso` saca el bloque del key box cuando el manual va pegado al
// mensaje del código: ese mensaje ya dio el código y la regla del control, y
// repetirlo un segundo después es ruido. Si el huésped toca "Manual de cabaña"
// suelto, el bloque va — puede no haber visto el otro mensaje.
function _botManualCabanaTexto(reserva, omitirAcceso) {
  let pasos = getCabinGuideSteps(
    reserva.cabin, reserva.tipo, !!reserva.checkoutExtendido, reserva.horaSalida || '',
    !!reserva.mascotas
  );
  if (omitirAcceso) pasos = pasos.filter(function(p) { return p.title !== 'Acceso'; });
  // getCabinGuideSteps devuelve [{icon, title, body}] con el cuerpo en HTML
  // (lo consume el email). Acá se pasa a texto de WhatsApp.
  const limpiar = (html) => String(html || '')
    .replace(/<br\s*\/?>/gi, '\n')
    // El <a> se convierte en "texto: url". Sin esto el strip se comía la URL y
    // "ver la lista" quedaba como texto muerto.
    .replace(/<a[^>]*href="([^"]*)"[^>]*>(.*?)<\/a>/gi, '$2: $1')
    .replace(/<\/?strong>/gi, '*')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    // Las entidades numéricas (&#9888;) son la forma en que el resto del archivo
    // escribe emoji para que Gmail no los muestre como "??????". En WhatsApp hay
    // que decodificarlas o el huésped lee el código crudo.
    .replace(/&#(\d+);/g, function(_, n) {
      try { return String.fromCodePoint(parseInt(n, 10)); } catch(e) { return ''; }
    })
    .replace(/&amp;/g, '&')
    .trim();
  let txt = '📖 *Manual de ' + (reserva.cabinName || 'la cabaña') + '*\n';
  pasos.forEach(function(p) {
    txt += '\n*' + limpiar(p.title) + '*\n' + limpiar(p.body) + '\n';
  });
  return txt.trim();
}

function _botEnviarManualCabana(from, reserva, omitirAcceso) {
  try {
    sendWhatsAppText(from, _botManualCabanaTexto(reserva, omitirAcceso));
  } catch(err) {
    logDebugEntry('manual-cabana-FAIL', { error: err.message, cabin: reserva && reserva.cabin });
    sendWhatsAppText(from, '⚠️ No pude armar el manual. Llama a Josh al +507 6981-2266 y te ayuda.');
  }
}

function _botHandleManualCabana(from, contactName, reservaId) {
  const reserva = _botResolverReservaLlegada(from, reservaId);
  if (!reserva) {
    sendWhatsAppText(from, '🤔 No encuentro tu reserva para armar el manual. Escríbeme "agente" y te ayuda una persona.');
    return;
  }
  _botEnviarManualCabana(from, reserva);
}

function _botHandleLlamarJosh(from) {
  const tel = '+507 6981-2266';
  try {
    sendWhatsAppCTAUrl(from, '📞 Toca el botón para llamar a Josh.', 'Llamar ahora', 'tel:+50769812266');
  } catch(err) {
    // Meta rechaza tel: en algunos casos. WhatsApp autodetecta el número del
    // cuerpo y lo vuelve tappable, así que el fallback sirve igual.
    logDebugEntry('llamar-josh-cta-FAIL', { error: err.message });
    sendWhatsAppText(from, '📞 Llama a Josh: ' + tel);
  }
}

// ─── Código de acceso (key box) ────────────────────────────────────
// Antes de dar el código pedimos el documento de identidad, igual que la página
// pública. Si la reserva ya lo tiene subido (por la página o por una estadía
// anterior), se salta el trámite y va directo al código.
function _botHandleCodigoAcceso(from, contactName, conv, reservaId) {
  const reserva = _botResolverReservaLlegada(from, reservaId);
  if (!reserva) {
    sendWhatsAppText(from, '🤔 No encuentro tu reserva. Escríbeme "agente" y te ayuda una persona.');
    return;
  }
  if (reserva.idHuespedURL) {
    return _botEnviarCodigoAcceso(from, reserva);
  }
  sendWhatsAppText(from,
    '🔑 *Código de acceso*\n\n' +
    'Para desbloquearlo necesito una foto de tu *cédula o pasaporte*.\n\n' +
    '¿Por qué lo pedimos?\n' +
    '• Validamos la identidad del huésped principal (uso interno).\n' +
    '• *Borramos la foto 60 días después de tu salida.* Solo guardamos tu fecha de nacimiento.\n' +
    '• 🎂 Si tu cumpleaños cae en esta o en futuras estadías, activamos el descuento automáticamente.\n' +
    '• Solo se pide la primera vez. Si vuelves, no te la pedimos de nuevo.\n\n' +
    '📷 Envíamela por aquí como foto. Que se lea bien la *fecha de nacimiento*.'
  );
  _saveConv(from, 'AWAITING_CEDULA',
    Object.assign({}, conv.context || {}, { reservaId: reserva.id }), contactName);
}

function _botEnviarCodigoAcceso(from, reserva) {
  // PUBLIC_KEY_BOX_CODE vive en PublicLink.gs — misma constante que usa la
  // página pública, para que un cambio de código no deje al bot dando el viejo.
  const codigo = (typeof PUBLIC_KEY_BOX_CODE !== 'undefined') ? PUBLIC_KEY_BOX_CODE : '0507';
  sendWhatsAppText(from,
    '🔑 *Código del key box: ' + codigo + '*\n\n' +
    'Dentro está la llave de la cabaña y, en el mismo llavero, un control negro con ' +
    'botones verdes que abre el portón verde de la entrada. Úsenlo si necesitan salir ' +
    '*durante su estadía*.\n\n' +
    '⚠️ *El día del check-out el control se queda.* Déjenlo en el key box junto con la ' +
    'llave — no se lo lleven.\n\n' +
    'Te dejo abajo el manual de la cabaña. ¡Disfruten! 🌿'
  );
  _botEnviarManualCabana(from, reserva, true);   // el acceso ya se explicó arriba
}

function _botHandleCedulaImage(from, imageId, contactName, conv) {
  const reserva = _botResolverReservaLlegada(from, (conv.context || {}).reservaId);
  if (!reserva) {
    sendWhatsAppText(from, '🤔 Perdí el hilo de tu reserva. Escríbeme "agente" y te ayuda una persona.');
    _saveConv(from, 'INITIAL', {}, contactName);
    return;
  }
  sendWhatsAppText(from, '⏳ Revisando tu documento...');
  const img = fetchWhatsAppImage(imageId);
  if (!img) {
    sendWhatsAppText(from, '⚠️ No pude descargar la imagen. ¿La reenvías?');
    return;   // sigue en AWAITING_CEDULA
  }
  const res = guardarIdHuesped(reserva.id, reserva.name, img.base64, img.mimeType);
  if (!res.ok) {
    // DOB_NOT_FOUND cubre los dos casos que importan: foto ilegible y foto que
    // no es un documento (el OCR devuelve la fecha vacía en ambos). El mensaje
    // no acusa a nadie de mandar cualquier cosa: pide una foto mejor.
    sendWhatsAppText(from,
      '🤔 No pude leer un documento de identidad en esa foto.\n\n' +
      'Asegúrate de que sea la *cédula o el pasaporte*, con buena luz, sin reflejos ' +
      'y que se lea la *fecha de nacimiento*.\n\n' +
      'Reenvíala, o escríbeme "agente" si prefieres que te ayude una persona.'
    );
    logDebugEntry('cedula-bot-RECHAZADA', { from: from, reservaId: reserva.id, error: res.error });
    return;   // sigue en AWAITING_CEDULA
  }
  logDebugEntry('cedula-bot-OK', { from: from, reservaId: reserva.id, tipo: res.tipo });
  if (res.warning) {
    try { _botAdminAlert('cedula', '⚠️ ' + res.warning + '\n👤 ' + reserva.name + '\n📱 +' + from); } catch(_) {}
  }
  _saveConv(from, 'ARRIVED',
    Object.assign({}, conv.context || {}, { reservaId: reserva.id }), contactName);
  sendWhatsAppText(from, '✅ ¡Listo, documento recibido!');
  _botEnviarCodigoAcceso(from, reserva);
}

// Preview desde el editor: manda a tu propio número las instrucciones de llegada
// de una cabaña —texto, foto y los tres botones— sin tocar reservas reales.
//
//   _testLlegadaAMiNumero('verde')   → Paseo
//   _testLlegadaAMiNumero('azul')    → Portal
//   _testLlegadaAMiNumero('lila')    → Puente
//
// Es SOLO preview: no dispara la alerta de portón al admin ni toca el estado de
// tu conversación con el bot.
//
// OJO: sirve para revisar los MENSAJES. Los tres botones necesitan una reserva
// activa con tu teléfono para responder algo útil — al tocarlos, el handler cae
// a _botFindReservaByPhone y sin reserva contesta "no encuentro tu reserva".
// Para probar los botones de punta a punta, pasar el id de una reserva real
// como segundo argumento, o crear una reserva de prueba con tu número.
function _testLlegadaAMiNumero(cabinKey, reservaIdOpcional) {
  const phone = PropertiesService.getScriptProperties().getProperty('PREVIEW_NOTIFY_PHONE') || '50769812266';
  const cabin = cabinKey || 'verde';
  const reserva = reservaIdOpcional
    ? _botFindReservaById(reservaIdOpcional)
    : { id: 'preview', name: 'Ana Gómez', cabin: cabin,
        cabinName: BOT_CABIN_NAMES[cabin], tipo: 'noche',
        checkoutExtendido: false, horaSalida: '', idHuespedURL: '', mascotas: false };
  if (!reserva) { Logger.log('✗ No encontré la reserva ' + reservaIdOpcional); return; }
  const conv = { context: {} };
  try {
    _botSendArrivalInstructions(phone, 'Ana', conv, reserva, { preview: true });
    Logger.log('✓ Instrucciones de ' + (reserva.cabinName || cabin) + ' enviadas a ' + phone);
  } catch(e) {
    Logger.log('✗ Falló: ' + e.message);
  }
}

// El desplegable "Ejecutar" del editor solo corre funciones SIN parámetros, así
// que cada cabaña tiene su wrapper para poder probarlas de a una.
function _testLlegadaPaseo()  { _testLlegadaAMiNumero('verde'); }
function _testLlegadaPortal() { _testLlegadaAMiNumero('azul');  }
function _testLlegadaPuente() { _testLlegadaAMiNumero('lila');  }

// Manda las tres, una detrás de otra, para comparar de un vistazo.
function _testLlegadaLasTres() {
  ['azul', 'lila', 'verde'].forEach(function(c) {
    _testLlegadaAMiNumero(c);
    Utilities.sleep(1500);   // que lleguen en orden
  });
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
      'Toca el botón para escribirle directo 👇',
      'Escribirle a Josh',
      'https://wa.me/50769812266?text=' + encodeURIComponent(prefill)
    );
  } catch(_) {
    sendWhatsAppText(from, 'Escribile directo aquí:\nhttps://wa.me/50769812266');
  }
  // Mantener CONFIRMED/ARRIVED — la reserva sigue siendo válida y debe seguir
  // contando como venta aunque el cliente toque "Consultas y cambios". Solo
  // degradar a HUMAN_HANDOFF si el cliente no tenía reserva confirmada.
  const conv = _getConv(from);
  const keep = conv && (conv.step === 'CONFIRMED' || conv.step === 'ARRIVED');
  if (!keep) {
    _saveConv(from, 'HUMAN_HANDOFF', (reserva && reserva.id) ? { reservaId: reserva.id } : {}, contactName);
  }
  logDebugEntry('consulta-reserva', { from: from, reservaId: reservaId, kept: !!keep });
}

// El huésped tocó "Envíame ubicación" en la plantilla de check-in.
// → le mandamos ubicación (Maps/Waze) e indicaciones de cómo llegar.
// No tocamos el estado de la conversación para no interferir con el flujo
// de "He llegado" cuando el huésped escriba al llegar al portón.
function _botHandleEnviarUbicacion(from, contactName) {
  _botMenuComoLlegar(from);
  logDebugEntry('checkin-enviar-ubicacion', { from: from, name: contactName });
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

  // Confirmacion al huésped. Antes cerraba en "en un momento te abren" y ahí
  // terminaba: si nadie abría, el huésped quedaba en el portón sin a quién
  // recurrir.
  //
  // El respaldo es JOSH, no el portón: el portón solo acepta llamadas del número
  // del admin, así que darle al huésped el WA_GATE_PHONE lo mandaba a un número
  // que le iba a colgar. Ese número sigue en el email al admin, que sí puede
  // usarlo.
  const cierre =
    '¡Gracias por avisar! 🌿 Ya le avisé al equipo para que les abran el portón.\n\n' +
    'Si en un par de minutos no se abre, llamen a Josh: *+507 6981-2266*\n\n' +
    '¡Buen viaje y esperamos verlos pronto de nuevo en Las Nubes! 🙌';
  try {
    sendWhatsAppCTAUrl(from, cierre, '📞 Llamar a Josh', 'tel:+50769812266');
  } catch(err) {
    // WhatsApp autodetecta el número del cuerpo y lo vuelve tappable igual.
    logDebugEntry('checkout-cta-josh-FAIL', { error: err.message });
    sendWhatsAppText(from, cierre);
  }

  // Aviso al admin via plantilla HSM (alerta_porton) — pasa siempre,
  // sin depender de la ventana de 24h.
  _sendAlertaPorton('salida', guestName, from, cabinName);

  // Email backup: red de seguridad para el portón. El email siempre llega
  // aunque WhatsApp falle por algún motivo. Se envía como HTML para que iOS
  // renderice los emojis correctamente en la notificación (en texto plano
  // los muestra como "?????" en la preview del lockscreen).
  const gatePhone = PropertiesService.getScriptProperties().getProperty('WA_GATE_PHONE') || '+507 6777-5630';
  try {
    const escape = (s) => String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    const subject = _asuntoEmailSeguro('🚪 ABRE EL PORTÓN — ' + cabinName + ' (' + guestName + ')');
    const htmlBody =
      '<div style="font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif;font-size:15px;color:#222;line-height:1.5;">' +
        '<p style="margin:0 0 12px;">El huésped llegó al portón y pidió que le abran (botón de la plantilla de check-out).</p>' +
        '<p style="margin:0 0 4px;">👤 <b>Huésped:</b> ' + escape(guestName) + '</p>' +
        '<p style="margin:0 0 4px;">📱 <b>Teléfono:</b> +' + escape(from) + '</p>' +
        '<p style="margin:0 0 4px;">🏡 <b>Cabaña:</b> ' + escape(cabinName) + '</p>' +
        '<p style="margin:0 0 12px;">📞 <b>Portón:</b> ' + escape(gatePhone) + '</p>' +
        '<p style="margin:0;color:#888;font-size:13px;">— Agente Las Nubes</p>' +
      '</div>';
    GmailApp.sendEmail(REPLY_TO_EMAIL, subject, '', {
      htmlBody: htmlBody,
      name: 'Las Nubes Agente'
    });
  } catch(e) {
    logDebugEntry('email-portón-FAIL', { from: from, error: e.message });
  }

  // Aviso a Erika via plantilla HSM (alerta_limpieza). {{3}} dice cuándo y a
  // qué hora llega la próxima reserva, y si requiere cama auxiliar — con las
  // reglas por cabaña y por tipo de _botTextoCamaAuxiliar, las mismas que usa
  // el parte diario de las 8am.
  try {
    const limpiezaPhone = PropertiesService.getScriptProperties().getProperty('LIMPIEZA_PHONE');
    if (limpiezaPhone) {
      const next = _botFindNextReservationForCabin(reserva && reserva.cabin, reserva && reserva.id);
      const ctxLine = _botBuildLimpiezaContextLine(next, reserva && reserva.cabin);
      sendWhatsAppTemplate(limpiezaPhone, 'alerta_limpieza_', 'es_ES', [
        cabinName,
        guestName,
        ctxLine
      ], null, null);
    }
  } catch(e) {
    logDebugEntry('alerta-limpieza-FAIL', { error: e.message });
  }

  logDebugEntry('checkout-done', { from: from, reservaId: reservaId, cabin: (reserva && reserva.cabin) || '?' });
}

const BOT_ADMIN_PHONE = '50769812266';

// ─── Config de alertas al admin (toggleable desde el dashboard) ──────
// Cada categoría puede prenderse/apagarse. Se guarda como JSON en la
// Script Property BOT_ALERTS_CONFIG. Las alertas críticas (pre-reserva
// para aprobar, portón de check-out) NO son toggleables y van siempre.
const BOT_ALERT_DEFAULTS = {
  nuevoCliente:      true,   // 🔔 el Agente empieza a atender un nuevo lead
  eligiendoCierre:   true,   // 🤝 el cliente llega a elegir cómo cerrar
  pagando:           true,   // 💳 el cliente entró a formas de pago
  handoff:           true,   // 🙋 derivado a humano (agente / asistido / insiste / cambio)
  seguimientoDiario: true    // 📋 resumen diario 8am de leads para dar seguimiento
};

function _botGetAlertConfig() {
  try {
    const raw   = PropertiesService.getScriptProperties().getProperty('BOT_ALERTS_CONFIG');
    const saved = raw ? JSON.parse(raw) : {};
    return Object.assign({}, BOT_ALERT_DEFAULTS, saved);
  } catch(_) { return Object.assign({}, BOT_ALERT_DEFAULTS); }
}

function _botSetAlertConfig(partial) {
  const merged = Object.assign(_botGetAlertConfig(), partial || {});
  // Solo persistir las claves conocidas (coerce a boolean).
  const clean = {};
  Object.keys(BOT_ALERT_DEFAULTS).forEach(k => { clean[k] = merged[k] !== false; });
  PropertiesService.getScriptProperties().setProperty('BOT_ALERTS_CONFIG', JSON.stringify(clean));
  return clean;
}

// Envía una alerta al admin solo si la categoría está activa.
// categoria: clave de BOT_ALERT_DEFAULTS, o null/'' para alertas siempre-activas.
function _botAdminAlert(categoria, mensaje, buttons) {
  if (categoria) {
    const cfg = _botGetAlertConfig();
    if (cfg[categoria] === false) return false;   // silenciada por el admin
  }
  try {
    if (buttons && buttons.length) sendWhatsAppButtons(BOT_ADMIN_PHONE, mensaje, buttons);
    else sendWhatsAppText(BOT_ADMIN_PHONE, mensaje);
    return true;
  } catch(e) {
    logDebugEntry('bot-admin-alert-FAIL', { categoria: categoria || '(siempre)', error: e.message });
    return false;
  }
}

// Alerta: el cliente llegó a elegir cómo cerrar (decoración / autoservicio vs asistido).
function _botNotifyEligiendoCierre(from, contactName, ctx) {
  try {
    const cabinName = BOT_CABIN_NAMES[ctx.cabin] || (ctx.cabin || '?');
    const dts       = ctx.dates;
    const fechas    = (dts && dts.checkin) ? (_botFmtFecha(dts.checkin) + ' → ' + _botFmtFecha(dts.checkout)) : '?';
    const personas  = ctx.personas || '?';
    const dashUrl   = 'https://lasnubes.cloud/dashboard.html?admin=1#bot:' + from;
    _botAdminAlert('eligiendoCierre',
      '🤝 *Cliente eligiendo cómo cerrar*\n\n' +
      '👤 ' + (contactName || from) + '\n' +
      '📱 +' + from + '\n' +
      '🏡 ' + cabinName + '\n' +
      '📅 ' + fechas + '\n' +
      '👥 ' + personas + '\n\n' +
      '_Está a un paso de reservar. 👀 ' + dashUrl + '_');
  } catch(_) {}
}

// Alerta: el cliente entró a formas de pago (autoservicio).
function _botNotifyPagando(from, contactName, ctx) {
  try {
    const cabinName = BOT_CABIN_NAMES[ctx.cabin] || (ctx.cabin || '?');
    const dts       = ctx.dates;
    const fechas    = (dts && dts.checkin) ? (_botFmtFecha(dts.checkin) + ' → ' + _botFmtFecha(dts.checkout)) : '?';
    const precioStr = (typeof ctx.precio === 'number') ? ('$' + ctx.precio.toFixed(2)) : '?';
    const dashUrl   = 'https://lasnubes.cloud/dashboard.html?admin=1#bot:' + from;
    _botAdminAlert('pagando',
      '💳 *Cliente en formas de pago*\n\n' +
      '👤 ' + (contactName || from) + '\n' +
      '📱 +' + from + '\n' +
      '🏡 ' + cabinName + '\n' +
      '📅 ' + fechas + '\n' +
      '💰 ' + precioStr + '\n\n' +
      '_Le mostré las formas de pago (autoservicio). 👀 ' + dashUrl + '_');
  } catch(_) {}
}

// Camas por cabana (igual que index.html / dashboard).
// Ojo: `verde` decía "queen y un sofá-cama doble" mientras el sitio y el
// dashboard dicen "king + auxiliar twin". Un sofá-cama doble duerme 2 y una
// twin duerme 1, así que la diferencia no es de redacción: alguien podía
// reservar Paseo para 2 adultos y 2 niños y no tener dónde acostar al cuarto.
const BOT_CABIN_CAMAS = {
  verde: 'La cabaña cuenta con una cama king y una cama auxiliar twin (individual).',
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
  // Default: formato estandar de Las Nubes.
  //
  // OJO CON LA CUENTA DEL ACH. Acá estaba la cuenta PERSONAL (a nombre de
  // Joslyn Lopez) mientras el admin, respondiendo a mano, mandaba a la de Las
  // Nubes (Iris Albelo). O sea que había dos instrucciones de pago circulando y
  // una de ellas —la automatizada, la que corre sola— cobraba en la cuenta
  // equivocada.
  //
  // No es cosmético: en feb-jul 2026, $10,378 de $25,497 en pagos de huéspedes
  // (41%) entraron por la cuenta personal en vez de la del negocio, y eso hace
  // que la contabilidad y el banco no cuadren nunca. Ver "Conciliación
  // bancaria" en CLAUDE.md.
  //
  // El orden también cambió: primero DÓNDE pagar y después qué escribir en el
  // mensaje. Al revés, el huésped lee la instrucción antes de saber a qué se
  // refiere.
  return 'Puede realizar el pago a través de:\n\n' +
    '*Yappy*\n69812266\nJoslyn Lopez\n\n' +
    '*ACH*\nBanco General\nIris Albelo\nCta de Ahorros\n04-99-99-818911-2\n\n' +
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
    '• Iluminación 100% solar — no hay luz eléctrica convencional\n' +
    '• Inversor disponible para cargar celulares y dispositivos\n' +
    '• Excelente señal de todas las operadoras\n' +
    '\n' +
    '*Baño & Comodidades*\n' +
    '• Jabón de baño, papel higiénico y toallas limpias incluidos\n• Agua fría — no contamos con agua caliente\n' +
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
  // Horas efectivas vía _horaPlantilla: respeta HoraEntrada/HoraSalida custom
  // (ej. noche que entra 12:30pm y sale 8am) y el guard un-solo-día. Antes
  // estaban hardcodeadas (2pm/11am).
  const checkinH  = _horaPlantilla(tipo, 'checkin',  reservation.checkoutExtendido, reservation.horaEntrada);
  const checkoutH = _horaPlantilla(tipo, 'checkout', reservation.checkoutExtendido, reservation.horaEntrada, reservation.horaSalida);

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
    // La elegibilidad se define por TIPO de día y no por día de semana: un
    // martes feriado o de vacaciones escolares seguía siendo "Dom–Jue", y ahí
    // el descuento caía justo en las noches de mayor demanda.
    text += '\n(Domingo a jueves, sin feriados ni vacaciones escolares · reservas directas)';
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
    sendWhatsAppText(from, '🤔 Perdí el contexto de las fechas. ¿Puedes decirme de nuevo cuándo quieres reservar?');
    _saveConv(from, 'AWAITING_DATES', conv.context, contactName);
    return;
  }
  const payingPersons = Math.max(2, personas - freeChildren);
  const basePrecio    = _botPrecioCabin(cabin, dates.checkin, dates.checkout, payingPersons);
  const ctx = Object.assign({}, conv.context, { cabin: cabin, basePrecio: basePrecio });

  // Si hubo señal de decoración y aún no decidió → preguntar con/sin antes del cierre.
  if (ctx.wantsDecoracion && ctx.decoracion === undefined) {
    return _botAskDecoracion(from, contactName, { context: ctx, name: contactName });
  }
  return _botShowCloseChoice(from, contactName, { context: ctx, name: contactName });
}

// Pregunta si agrega decoración especial (+$40) antes de elegir el cierre.
function _botAskDecoracion(from, contactName, conv) {
  const ctx   = conv.context || {};
  const cabin = ctx.cabin;
  const base  = ctx.basePrecio || 0;
  const fechas = _botFmtFecha(ctx.dates.checkin) + ' → ' + _botFmtFecha(ctx.dates.checkout);
  const body =
    '🎉 Reservando *' + BOT_CABIN_NAMES[cabin] + '* para *' + fechas + '*.\n\n' +
    '¿Quieres agregar nuestra *decoración especial*? Incluye flores, globos, letrero, detalles románticos y una botella de espumante 🥂.\n\n' +
    '💰 Sin decoración: *$' + base.toFixed(2) + '*\n' +
    '🎉 Con decoración: *$' + (base + BOT_DECOR_FEE).toFixed(2) + '* (+$' + BOT_DECOR_FEE + ')';
  try {
    sendWhatsAppButtons(from, body, [
      { id: 'deco_yes',       title: '🎉 Con decoración' },
      { id: 'deco_no',        title: 'Sin decoración' },
      { id: 'cancel_booking', title: '❌ Cancelar' }
    ]);
  } catch(_) {
    sendWhatsAppText(from, body + '\n\nEscribe "1" para *con decoración*, "2" para *sin*, o "cancelar".');
  }
  if (!ctx.cierreAlertSent) { _botNotifyEligiendoCierre(from, contactName, ctx); ctx.cierreAlertSent = true; }
  _saveConv(from, 'CHOOSING_DECOR', ctx, contactName);
}

// Muestra la elección de cierre (autoservicio vs asistido) con el total final.
function _botShowCloseChoice(from, contactName, conv) {
  const ctx          = conv.context || {};
  const cabin        = ctx.cabin;
  const dates        = ctx.dates;
  const personas     = ctx.personas || 2;
  const freeChildren = ctx.freeChildren || 0;
  const base         = (typeof ctx.basePrecio === 'number') ? ctx.basePrecio : (ctx.precio || 0);
  const deco         = !!ctx.decoracion;
  const total        = base + (deco ? BOT_DECOR_FEE : 0);
  const fechas       = _botFmtFecha(dates.checkin) + ' → ' + _botFmtFecha(dates.checkout);
  let personasLbl = personas + (personas === 1 ? ' persona' : ' personas');
  if (freeChildren > 0) personasLbl += ' · ' + freeChildren + ' menor' + (freeChildren === 1 ? '' : 'es') + ' de 5 sin cargo';

  let totalLine = '💰 *Total: $' + total.toFixed(2) + '*';
  if (deco) totalLine += ' _(incluye decoración especial 🎉)_';

  const body =
    '🎉 ¡Excelente! Reservando *' + BOT_CABIN_NAMES[cabin] + '* para *' + fechas + '* (' + personasLbl + ').\n\n' +
    totalLine + '\n\n' +
    '¿Cómo quieres cerrar tu reserva?\n\n' +
    '⚡ *Reservar Ahora*: te autogestionas — te comparto las formas de pago, subes el comprobante aquí mismo, tu reserva queda registrada y la confirmamos en breve por email y por este WhatsApp.\n\n' +
    '🙋 *Reservar Asistido*: te transfiero con Josh para que te asista durante todo el proceso de reserva.';
  try {
    sendWhatsAppButtons(from, body, [
      { id: 'close_self',     title: 'Reservar Ahora ⚡' },
      { id: 'close_asesor',   title: 'Reservar Asistido 🙋' },
      { id: 'cancel_booking', title: '❌ Cancelar' }
    ]);
  } catch(_) {
    sendWhatsAppText(from, body + '\n\nEscribe "1" para *Reservar Ahora*, "2" para *Asistido*, o "cancelar".');
  }
  const ctxClose = Object.assign({}, ctx, { precio: total });
  if (!ctxClose.cierreAlertSent) { _botNotifyEligiendoCierre(from, contactName, ctxClose); ctxClose.cierreAlertSent = true; }
  _saveConv(from, 'CHOOSING_CLOSE', ctxClose, contactName);
}

// Autoservicio: muestra formas de pago y pide el comprobante.
function _botOfferPayment(from, contactName, conv) {
  const cabin = conv.context && conv.context.cabin;
  if (!cabin) {
    sendWhatsAppText(from, '🤔 Perdí el contexto de la reserva. Dime de nuevo las fechas y personas, por favor.');
    _saveConv(from, 'AWAITING_DATES', {}, contactName);
    return;
  }
  const body =
    '🌿 ¡Perfecto! Cerramos *' + BOT_CABIN_NAMES[cabin] + '*.\n\n' +
    _botPaymentInfo() + '\n\n' +
    '⚠️ *Importante*: en el detalle de la transferencia coloca tu *nombre completo* y *email* para procesar tu reserva más rápido.\n\n' +
    'Una vez transferido, *envíame el comprobante como imagen* por aquí mismo.';
  try {
    sendWhatsAppButtons(from, body, [{ id: 'cancel_booking', title: '❌ Cancelar' }]);
  } catch(_) {
    sendWhatsAppText(from, body + '\n\nSi quieres cancelar, escribe "cancelar".');
  }
  const ctxPay = conv.context || {};
  if (!ctxPay.pagandoAlertSent) { _botNotifyPagando(from, contactName, ctxPay); ctxPay.pagandoAlertSent = true; }
  _saveConv(from, 'OFFERING_PAYMENT', ctxPay, contactName);
}

// Asistido: conecta con Josh para cerrar + alerta al admin (lead caliente).
function _botCloseWithAsesor(from, contactName, conv) {
  const ctx       = conv.context || {};
  const cabinName = BOT_CABIN_NAMES[ctx.cabin] || '';
  const dates     = ctx.dates;
  const fechas    = (dates && dates.checkin) ? (_botFmtFecha(dates.checkin) + ' al ' + _botFmtFecha(dates.checkout)) : '';
  const personas  = ctx.personas || '';

  const deco = !!ctx.decoracion;
  let prefill = 'Hola Josh, quiero reservar';
  if (cabinName) prefill += ' ' + cabinName;
  if (fechas)    prefill += ' del ' + fechas;
  if (personas)  prefill += ' (' + personas + ' personas)';
  if (deco)      prefill += ' con decoración especial';
  prefill += '. ¿Me ayudas a cerrar la reserva?';

  sendWhatsAppText(from, 'Genial 🌿 Josh te va a acompañar para cerrar tu reserva.');
  try {
    sendWhatsAppCTAUrl(from, 'Toca el botón para escribirle 👇', 'Escribir a Josh',
      'https://wa.me/50769812266?text=' + encodeURIComponent(prefill));
  } catch(_) {
    sendWhatsAppText(from, 'Escríbele directo:\nhttps://wa.me/50769812266');
  }
  const precioStr = (typeof ctx.precio === 'number') ? ('$' + ctx.precio.toFixed(2)) : '?';
  _botAdminAlert('handoff',
    '🔥 *Lead listo para cerrar (asistido)*\n\n' +
    '👤 ' + (contactName || from) + '\n' +
    '📱 +' + from + '\n' +
    '🏡 ' + (cabinName || '?') + '\n' +
    '📅 ' + (fechas || '?') + '\n' +
    '👥 ' + (personas || '?') + '\n' +
    (deco ? '🎉 Con decoración especial (+$' + BOT_DECOR_FEE + ')\n' : '') +
    '💰 ' + precioStr + '\n\n' +
    '_El cliente eligió "Reservar Asistido" tras cotizar._');
  _saveConv(from, 'PENDING_HUMAN_BOOKING', ctx, contactName);
}

// ¿El cliente insiste en una fecha (tras ver que no hay disponibilidad)?
function _botIsDateInsistence(text) {
  const t = (text || '').toLowerCase();
  return /\b(esa\s+fecha|esas\s+fechas|quiero\s+esa|necesito\s+esa|tiene\s+que\s+ser|s[ií]\s+o\s+s[ií]|justo\s+esa|esa\s+es\s+la|no\s+puede\s+ser\s+otra|s[oó]lo\s+esa|solo\s+esa|es\s+importante|importante)\b/.test(t)
      || /\b(cumple|cumplea[ñn]os|aniversario|luna\s+de\s+miel)\b/.test(t);
}

// Insistencia en fecha ocupada → empatía + conectar con Josh + alerta al admin.
function _botHandleDateInsistence(from, contactName, conv) {
  const ctx    = conv.context || {};
  const dates  = ctx.dates;
  const fechas = (dates && dates.checkin) ? (_botFmtFecha(dates.checkin) + (dates.checkout ? ' → ' + _botFmtFecha(dates.checkout) : '')) : '';
  const personas = ctx.personas || '';

  let prefill = 'Hola Josh, quiero reservar';
  if (fechas)   prefill += ' para el ' + fechas;
  if (personas) prefill += ' (' + personas + ' personas)';
  prefill += ' — sé que aparece sin disponibilidad pero esa fecha es importante para mí. ¿Hay alguna opción?';

  sendWhatsAppText(from,
    '🌿 Entiendo que esa fecha es especial para ti. Déjame conectarte con Josh para ver opciones — a veces se puede coordinar algo (lista de espera, fechas muy cercanas, etc.).'
  );
  try {
    sendWhatsAppCTAUrl(from, 'Toca el botón para escribirle 👇', 'Escribir a Josh',
      'https://wa.me/50769812266?text=' + encodeURIComponent(prefill));
  } catch(_) {
    sendWhatsAppText(from, 'Escríbele directo:\nhttps://wa.me/50769812266');
  }
  _botAdminAlert('handoff',
    '⭐ *Cliente insiste en fecha sin disponibilidad*\n\n' +
    '👤 ' + (contactName || from) + '\n' +
    '📱 +' + from + '\n' +
    '📅 ' + (fechas || '?') + '\n' +
    '👥 ' + (personas || '?') + '\n\n' +
    '_Quiere esa fecha específica. Ver si se puede acomodar (lista de espera / mover otra reserva)._');
  _saveConv(from, 'HUMAN_HANDOFF', ctx, contactName);
}


function _botHandleVoucherImage(from, imageId, contactName, conv) {
  if (conv.step !== 'OFFERING_PAYMENT' && conv.step !== 'AWAITING_VOUCHER_RETRY') {
    sendWhatsAppText(from, '📷 Recibí tu imagen, pero no estamos en una reserva activa. Si quieres reservar, escríbeme "1" o "disponibilidad".');
    return;
  }
  sendWhatsAppText(from, '⏳ Procesando tu comprobante...');
  const img = fetchWhatsAppImage(imageId);
  if (!img) {
    sendWhatsAppText(from, '⚠️ No pude descargar tu imagen. Prueba enviarla de nuevo o escríbeme "3" para hablar con una persona.');
    return;
  }
  let voucher;
  try {
    const out = parseVoucherWithClaude(img.base64, img.mimeType);
    voucher = JSON.parse(out.getContent());
  } catch(err) {
    logDebugEntry('bot-voucher-OCR-CRASH', { error: err.message });
    sendWhatsAppText(from, '⚠️ No pude leer el voucher. ¿Puedes enviarlo más claro o escríbeme "3" para una persona?');
    return;
  }
  if (!voucher || !voucher.ok || !voucher.codTransferencia) {
    sendWhatsAppText(from,
      '⚠️ No pude leer los datos del voucher. Asegurate que la imagen sea clara y tenga:\n\n' +
      '• Monto\n• Código/referencia\n• Fecha\n\nReenviame la imagen o escríbeme "3" para una persona.'
    );
    _saveConv(from, 'AWAITING_VOUCHER_RETRY', conv.context, contactName);
    return;
  }
  const monto = parseFloat(voucher.monto) || 0;
  // Guardar la imagen en Drive AHORA: los bytes solo existen en este punto del
  // flujo (el contexto de la conversación se persiste en una hoja, así que no
  // puede cargar un base64). Antes se descartaban y la reserva quedaba con el
  // código pero sin archivo — "sin archivo" en el modal.
  const voucherUrl = _botGuardarVoucherEnDrive(img, voucher, contactName || from);
  // Recuperar campos extraidos del campo "Mensaje" del voucher (si el cliente los coloco)
  const extractedName  = voucher.nombreCompleto ? voucher.nombreCompleto.toString().trim() : '';
  const extractedEmail = voucher.email ? voucher.email.toString().trim().toLowerCase() : '';
  const newCtx = Object.assign({}, conv.context, {
    voucher: {
      monto: monto,
      // Mismo saneo que el dashboard (handleVoucherUpload): sin el '#' inicial.
      codTransferencia: (voucher.codTransferencia || '').toString().replace(/^#/, ''),
      fechaPago: voucher.fechaPago || _botToday(),
      sender:   voucher.sender || '',
      url:      voucherUrl || ''
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
    sendWhatsAppText(from, confirmMsg + '\n\nPara finalizar, ¿me puedes enviar tu *email*?');
    _saveConv(from, 'AWAITING_EMAIL', newCtx, contactName);
    return;
  }
  // Tenemos email pero falta nombre
  sendWhatsAppText(from, confirmMsg + '\n\n¿Cuál es tu *nombre completo*?');
  _saveConv(from, 'AWAITING_NAME', newCtx, contactName);
}

// Sube a Drive el voucher que el cliente mandó por WhatsApp y devuelve su URL
// ('' si algo falla — nunca debe tumbar el flujo de la reserva).
//
// Se apoya en saveVoucherToDrive para reusar el nombrado y la descripción del
// archivo. La reserva todavía no existe como fila, así que ese paso interno no
// encuentra nada que actualizar y solo loguea un aviso; el archivo igual queda
// creado y su URL se guarda en el contexto para escribirla al crear la fila.
// El `confirmCode` va con el código de transferencia para que la descripción lo
// contenga y diagnosticarVouchersSinArchivo() pueda matchearlo después.
function _botGuardarVoucherEnDrive(img, voucher, quien) {
  try {
    if (!img || !img.base64) return '';
    const cod = (voucher && voucher.codTransferencia || '').toString().replace(/^#/, '');
    const fake = {
      id:          '',
      name:        (voucher && voucher.sender) || quien || 'huesped',
      cabin:       'wa',
      checkin:     (voucher && voucher.fechaPago) || '',
      checkout:    '',
      amount:      (voucher && voucher.monto) || '',
      deposit:     (voucher && voucher.monto) || '',
      confirmCode: cod,
      origin:      'Directa (Agente WhatsApp)'
    };
    const out  = saveVoucherToDrive(fake, img.base64, img.mimeType, 'voucher-wa.jpg', {
      monto: (voucher && voucher.monto) || 0,
      cod:   cod,
      fecha: (voucher && voucher.fechaPago) || ''
    });
    const data = JSON.parse(out.getContent());
    if (data && data.ok && data.fileUrl) {
      logDebugEntry('bot-voucher-drive-OK', { cod: cod, url: data.fileUrl });
      return data.fileUrl;
    }
    logDebugEntry('bot-voucher-drive-FAIL', { cod: cod, resp: JSON.stringify(data).slice(0, 200) });
  } catch(err) {
    logDebugEntry('bot-voucher-drive-CRASH', { error: err.message });
  }
  return '';
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
  const decoracion = !!ctx.decoracion;
  let comentario = skipVoucher
    ? '🤖 Pre-reserva vía Agente WhatsApp · SIN ABONO · coordinar pago · pendiente revisión'
    : '🤖 Pre-reserva vía Agente WhatsApp · pendiente revisión';
  if (decoracion) comentario += ' · 🎉 CON DECORACIÓN ESPECIAL (+$' + BOT_DECOR_FEE + ')';

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
    // Col 26 (VoucherURL) + col 32 (VouchersMeta): el archivo se subió a Drive
    // cuando llegó la imagen; acá recién existe la fila donde anotarlo.
    if (voucher.url) {
      try {
        const fila = sheet.getLastRow();
        sheet.getRange(fila, 26).setValue(voucher.url);
        sheet.getRange(fila, 32).setValue(JSON.stringify([{
          monto: parseFloat(voucher.monto) || 0,
          cod:   voucher.codTransferencia || '',
          fecha: voucher.fechaPago || '',
          url:   voucher.url
        }]));
      } catch(vErr) {
        logDebugEntry('bot-prereserva-voucherurl-FAIL', { id: id, error: vErr.message });
      }
    }
    logDebugEntry('bot-prereserva-OK', { id: id, name: fullName, cabin: cabin, from: from, skipVoucher: skipVoucher, voucherUrl: voucher.url || null });
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
    (decoracion ? '🎉 *CON DECORACIÓN ESPECIAL* (+$' + BOT_DECOR_FEE + ') — preparar\n' : '') +
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
        checkoutExtendido: data[i][28] === true || data[i][28] === 'TRUE' || data[i][28] === 'true' || data[i][28] === 1,
        horaEntrada: (typeof _normalizeHora === 'function') ? _normalizeHora(data[i][29]) : (data[i][29] || ''),
        horaSalida:  (typeof _normalizeHora === 'function') ? _normalizeHora(data[i][30]) : (data[i][30] || ''),
        mascotas:    !!data[i][33]
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
      // Elevar el estado de la conversación: PENDING_REVIEW → CONFIRMED
      try {
        const conv = _getConv(reservation.telefono.toString());
        if (conv) _saveConv(reservation.telefono.toString(), 'CONFIRMED', conv.context, conv.name);
      } catch(_) {}
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
      // Bajar el estado de la conversación: PENDING_REVIEW → REJECTED
      try {
        const conv = _getConv((clientPhone || '').toString());
        if (conv) _saveConv((clientPhone || '').toString(), 'REJECTED', conv.context, conv.name);
      } catch(_) {}
      return;
    }
  }
  sendWhatsAppText(adminPhone, '⚠️ No encontré la reserva ' + reservaId);
}
