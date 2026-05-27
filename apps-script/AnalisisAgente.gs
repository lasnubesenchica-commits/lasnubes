/**
 * Análisis del Agente de WhatsApp.
 *
 * analizarAgente(): arma el funnel real desde la hoja Conversaciones,
 * muestrea los mensajes de clientes desde Mensajes, y se lo pasa a Claude
 * para un análisis enfocado en dos objetivos:
 *   1. Manejar el "ruido" (curiosos sin intención de reservar) sin cargar al admin.
 *   2. Convertir reservas (funnel).
 *
 * Correr desde el editor y leer el Registro de ejecución (Logger).
 */

// Mapeo de step → etapa del funnel.
const _FUNNEL_STAGES = {
  INITIAL:                'A. Llegó / saludo',
  SHOWED_INFO:            'B. Consultó info (curioso)',
  AWAITING_DATES:         'C. Pidió/iba a dar fechas',
  SHOWING_AVAILABILITY:   'D. Cotizó (vio disponibilidad)',
  SHOWING_ALTERNATIVES:   'D. Cotizó (alternativas)',
  NO_AVAILABILITY:        'D. Cotizó (sin disponibilidad)',
  CHOOSING_DECOR:         'E. Eligió cabaña (eligiendo decoración)',
  CHOOSING_CLOSE:         'E. Eligió cabaña (eligiendo cierre)',
  OFFERING_PAYMENT:       'E. Autoservicio / formas de pago',
  AWAITING_VOUCHER_RETRY: 'E. Reintentando voucher',
  AWAITING_EMAIL:         'F. Dando datos (email)',
  AWAITING_NAME:          'F. Dando datos (nombre)',
  PENDING_REVIEW:         'G. Pre-reserva creada',
  CONFIRMED:              'G. Reserva confirmada ✓',
  REJECTED:               'X. Pre-reserva rechazada',
  ARRIVED:                'H. Llegó a la cabaña',
  HUMAN_HANDOFF:          'X. Derivado a humano',
  PENDING_HUMAN_BOOKING:  'E. Cierre asistido (esperando ingreso de reserva)',
  AWAITING_ARRIVAL_NAME:  'H. He llegado (verificando)'
};

function analizarAgente() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // ── 1) Funnel desde Conversaciones ──────────────────────────────
  const convSheet = ss.getSheetByName('Conversaciones');
  const stepCount = {};
  const stageCount = {};
  let totalConv = 0;
  let conGanasDeFecha = 0;   // dieron o iban a dar fechas
  let cotizaron = 0;          // llegaron a ver disponibilidad
  let avanzaronPago = 0;      // eligieron cabaña / formas de pago en adelante
  let preReservas = 0;
  let handoffs = 0;
  const firstMsgByPhone = {};

  if (convSheet && convSheet.getLastRow() > 1) {
    const data = convSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const phone = (data[i][0] || '').toString();
      if (!phone) continue;
      totalConv++;
      const step = (data[i][1] || 'INITIAL').toString();
      stepCount[step] = (stepCount[step] || 0) + 1;
      const stage = _FUNNEL_STAGES[step] || ('? ' + step);
      stageCount[stage] = (stageCount[stage] || 0) + 1;
      if (['AWAITING_DATES','SHOWING_AVAILABILITY','SHOWING_ALTERNATIVES','NO_AVAILABILITY','CHOOSING_DECOR','CHOOSING_CLOSE','OFFERING_PAYMENT','AWAITING_VOUCHER_RETRY','AWAITING_EMAIL','AWAITING_NAME','PENDING_HUMAN_BOOKING','PENDING_REVIEW','CONFIRMED','ARRIVED'].indexOf(step) !== -1) conGanasDeFecha++;
      if (['SHOWING_AVAILABILITY','SHOWING_ALTERNATIVES','NO_AVAILABILITY','CHOOSING_DECOR','CHOOSING_CLOSE','OFFERING_PAYMENT','AWAITING_VOUCHER_RETRY','AWAITING_EMAIL','AWAITING_NAME','PENDING_HUMAN_BOOKING','PENDING_REVIEW','CONFIRMED','ARRIVED'].indexOf(step) !== -1) cotizaron++;
      if (['CHOOSING_DECOR','CHOOSING_CLOSE','OFFERING_PAYMENT','AWAITING_VOUCHER_RETRY','AWAITING_EMAIL','AWAITING_NAME','PENDING_HUMAN_BOOKING','PENDING_REVIEW','CONFIRMED','ARRIVED'].indexOf(step) !== -1) avanzaronPago++;
      if (['PENDING_REVIEW','CONFIRMED','ARRIVED'].indexOf(step) !== -1) preReservas++;
      if (['HUMAN_HANDOFF'].indexOf(step) !== -1) handoffs++;
    }
  }

  // ── 2) Mensajes de clientes (inbound) desde Mensajes ─────────────
  const msgSheet = ss.getSheetByName('Mensajes');
  const inbound = [];                 // últimos N mensajes de cliente
  if (msgSheet && msgSheet.getLastRow() > 1) {
    const m = msgSheet.getDataRange().getValues();
    for (let i = 1; i < m.length; i++) {
      const phone = (m[i][1] || '').toString();
      const dir   = (m[i][2] || '').toString();
      const type  = (m[i][3] || '').toString();
      const content = (m[i][4] || '').toString();
      if (dir !== 'in') continue;
      if (!content) continue;
      // primer mensaje por teléfono (el "opener" — clave para objetivo 1)
      if (!firstMsgByPhone[phone]) firstMsgByPhone[phone] = content.slice(0, 200);
      inbound.push(content.slice(0, 200));
    }
  }
  const openers = Object.keys(firstMsgByPhone).map(p => firstMsgByPhone[p]);
  const muestraInbound = inbound.slice(-200);   // últimos 200

  // ── 3) Reporte de funnel (determinista) ──────────────────────────
  let report = '═══ FUNNEL DEL AGENTE ═══\n';
  report += 'Total conversaciones: ' + totalConv + '\n';
  report += '  • Pidieron/dieron fechas: ' + conGanasDeFecha + ' (' + _pct(conGanasDeFecha, totalConv) + ')\n';
  report += '  • Cotizaron (vieron disponibilidad): ' + cotizaron + ' (' + _pct(cotizaron, totalConv) + ')\n';
  report += '  • Avanzaron a pago/datos: ' + avanzaronPago + '\n';
  report += '  • Pre-reservas creadas: ' + preReservas + '\n';
  report += '  • Derivados a humano: ' + handoffs + '\n';
  report += '\nPor estado:\n';
  Object.keys(stageCount).sort().forEach(s => { report += '  ' + s + ': ' + stageCount[s] + '\n'; });
  Logger.log(report);

  // ── 4) Análisis cualitativo con Claude ───────────────────────────
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) { Logger.log('Sin CLAUDE_API_KEY — solo funnel.'); return report; }

  const systemPrompt =
    'Sos analista de growth de "Las Nubes", 3 cabañas en Panamá con un agente de WhatsApp que atiende leads. ' +
    'Hay DOS objetivos del agente:\n' +
    '1) Manejar el RUIDO: curiosos que solo preguntan sobre las cabañas sin intención real de reservar — el agente debe atenderlos solo y liberar al dueño.\n' +
    '2) CONVERTIR reservas (funnel).\n\n' +
    'Te paso (a) el funnel actual y (b) mensajes reales de clientes. Devolvé un análisis accionable y conciso en español con esta estructura:\n' +
    '## Lectura del funnel (dónde se cae)\n' +
    '## Objetivo 1 — Ruido/curiosos: qué preguntan, qué tan bien se auto-maneja, qué handlers o respuestas faltan\n' +
    '## Objetivo 2 — Conversión: dónde y por qué se pierde la intención, fricciones detectadas\n' +
    '## 5 acciones concretas priorizadas (handlers nuevos, cambios de copy, momentos de handoff, etc.)\n' +
    'Sé específico, citá ejemplos reales de los mensajes. No inventes datos que no estén.';

  const userContent =
    'FUNNEL:\n' + report + '\n\n' +
    'PRIMEROS MENSAJES de cada lead (openers, ' + openers.length + '):\n' +
    openers.map((s, i) => (i+1) + '. ' + s).join('\n') + '\n\n' +
    'MUESTRA de mensajes de clientes (últimos ' + muestraInbound.length + '):\n' +
    muestraInbound.map((s, i) => '- ' + s).join('\n');

  try {
    const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 2000,
        system: systemPrompt,
        messages: [{ role: 'user', content: userContent }]
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(res.getContentText());
    const analisis = data.content && data.content[0] && data.content[0].text;
    if (analisis) {
      Logger.log('\n═══ ANÁLISIS CLAUDE ═══\n' + analisis);
      report += '\n\n═══ ANÁLISIS CLAUDE ═══\n' + analisis;
    } else {
      Logger.log('Claude no devolvió análisis: ' + JSON.stringify(data).slice(0, 300));
    }
  } catch(err) {
    Logger.log('Error llamando a Claude: ' + err.message);
  }

  return report;
}

function _pct(n, total) {
  if (!total) return '0%';
  return Math.round((n / total) * 100) + '%';
}

// Opcional: envía el reporte al WhatsApp del admin (puede ser largo, se trunca).
function analizarAgenteYEnviar() {
  const report = analizarAgente();
  try { sendWhatsAppText('50769812266', '📊 *Análisis del Agente*\n\n' + report.slice(0, 3500)); } catch(e) { Logger.log(e.message); }
}
