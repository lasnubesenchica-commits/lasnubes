/**
 * Recordatorios diarios al ADMIN via WhatsApp.
 *
 * - 11:00 am: resumen de entradas y salidas del dia (similar al
 *   "Reservas de hoy" del dashboard), con links wa.me para escribir
 *   a cada cliente con un tap.
 * - 9:00 am: alerta si alguna reserva del dia tiene servicios
 *   especiales mencionados en comentarios (decoración, traslado,
 *   cumpleaños, menú, etc.).
 * - 8:00 am: recordatorio de limpieza a la Sra que limpia (numero en
 *   Script Property LIMPIEZA_PHONE), con instrucciones por cabaña.
 *
 * Setup: correr UNA VEZ desde el editor `instalarTriggersAdminReminders()`
 * y `instalarTriggerLimpieza()` (este ultimo tras setear LIMPIEZA_PHONE).
 */

// Reusa BOT_ADMIN_PHONE, BOT_TZ, BOT_CABIN_NAMES, _botAddDaysISO de BotConsultor.gs

// Keywords (sin acentos, lowercase) que disparan alerta de servicios especiales
const ADMIN_SPECIAL_KEYWORDS = [
  'decoracion', 'decorar', 'decoraciones', 'decorad',
  'traslado', 'transporte', 'shuttle', 'recoger',
  'cumpleanos', 'aniversario', 'celebrac',
  'menu', 'comida especial', 'cena especial', 'desayuno especial',
  'servicio especial', 'servicio adicional', 'pedido especial',
  'torta', 'pastel', 'cake',
  'flores', 'globos', 'rosas', 'ramo',
  'aeropuerto', 'vuelo'
];

function _adminHasSpecialKw(text) {
  if (!text) return false;
  const lower = String(text).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  return ADMIN_SPECIAL_KEYWORDS.some(k => lower.indexOf(k) >= 0);
}

function _adminWaLink(rawPhone) {
  if (!rawPhone) return null;
  const digits = String(rawPhone).replace(/\D/g, '');
  let waNum;
  if (digits.length === 8) waNum = '507' + digits;
  else if (digits.length === 11 && digits.indexOf('507') === 0) waNum = digits;
  else if (digits.length >= 10) waNum = digits;
  else return null;
  return 'https://wa.me/' + waNum;
}

// Formatea telefono panameno legible: +507 XXXX-XXXX. WhatsApp autodetecta
// este formato y lo vuelve tappable (abre options: chat/llamar).
function _adminFormatPhone(rawPhone) {
  if (!rawPhone) return null;
  const digits = String(rawPhone).replace(/\D/g, '');
  if (digits.length === 8) return '+507 ' + digits.slice(0, 4) + '-' + digits.slice(4);
  if (digits.length === 11 && digits.indexOf('507') === 0) {
    const rest = digits.slice(3);
    return '+507 ' + rest.slice(0, 4) + '-' + rest.slice(4);
  }
  if (digits.length >= 10) return '+' + digits;
  return null;
}

// Devuelve [{ kind: 'entrada'|'salida'|'pasadia', reserva }] para targetDate.
// Mirror simplificado de renderReservasHoy en dashboard.html.
function _adminGetMovimientosDia(targetDate) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    const origin = r[9];
    if (origin === 'Abierta') continue;
    const estado = (r[20] || '').toString().toUpperCase();
    if (estado === 'CANCELADA') continue;
    const ci = r[4] instanceof Date ? Utilities.formatDate(r[4], BOT_TZ, 'yyyy-MM-dd') : (r[4] || '').toString().slice(0,10);
    const co = r[5] instanceof Date ? Utilities.formatDate(r[5], BOT_TZ, 'yyyy-MM-dd') : (r[5] || '').toString().slice(0,10);
    if (!ci || !co) continue;
    const tipo = r[24] || 'noche';

    // Display dates segun tipo (mirror de _formFromStored)
    let displayCi = ci, displayCo = co;
    if (tipo === 'pasatarde') { displayCo = ci; }
    else if (tipo === 'pasadia') { displayCi = _botAddDaysISO(ci, 1); displayCo = displayCi; }
    else if (tipo === 'early')  { displayCi = _botAddDaysISO(ci, 1); }
    else if (tipo === 'late')   { displayCo = _botAddDaysISO(co, -1); }

    const reserva = {
      id: r[0], name: r[1] || '?',
      cabin: r[3], cabinName: r[2] || BOT_CABIN_NAMES[r[3]] || r[3],
      persons: r[6] || '?',
      origin: origin || '',
      telefono: r[23] || '',
      comentarios: (r[22] || '').toString().trim(),
      tipo: tipo,
      displayCi: displayCi,
      displayCo: displayCo
    };

    if (displayCi === targetDate && displayCo === targetDate) {
      items.push({ kind: 'pasadia', reserva: reserva });
    } else {
      if (displayCi === targetDate) items.push({ kind: 'entrada', reserva: reserva });
      if (displayCo === targetDate) items.push({ kind: 'salida',  reserva: reserva });
    }
  }
  return items;
}

function _adminFmtFechaLarga(iso) {
  const dias = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
  const meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  const d = new Date(iso + 'T12:00:00');
  return dias[d.getDay()] + ' ' + d.getDate() + ' de ' + meses[d.getMonth()];
}

// ─── Trigger 11am: resumen del dia ────────────────────────────────
function enviarRecordatorioAdminReservasHoy() {
  const today    = Utilities.formatDate(new Date(), BOT_TZ, 'yyyy-MM-dd');
  const items    = _adminGetMovimientosDia(today);
  const fechaLbl = _adminFmtFechaLarga(today);

  if (items.length === 0) {
    sendWhatsAppText(BOT_ADMIN_PHONE, '🌿 *Reservas de hoy* — ' + fechaLbl + '\n\nNo hay entradas ni salidas. Día tranquilo ✨');
    logDebugEntry('admin-recordatorio-11am', { count: 0 });
    return;
  }

  const order = { entrada: 0, pasadia: 1, salida: 2 };
  items.sort((a,b) => order[a.kind] - order[b.kind]);

  let msg = '🌿 *Reservas de hoy* — ' + fechaLbl + '\n';

  const entradas = items.filter(i => i.kind === 'entrada' || i.kind === 'pasadia');
  const salidas  = items.filter(i => i.kind === 'salida');

  if (entradas.length > 0) {
    msg += '\n📥 *ENTRAN* (' + entradas.length + ')';
    entradas.forEach(it => {
      const r = it.reserva;
      msg += '\n\n🏡 ' + r.cabinName;
      msg += '\n👤 ' + r.name + ' · ' + r.persons + (r.persons == 1 ? ' persona' : ' personas');
      if (r.origin && r.origin !== 'Airbnb') msg += ' · ' + r.origin;
      const phone = _adminFormatPhone(r.telefono);
      if (phone) msg += '\n💬 ' + phone;
      if (it.kind === 'pasadia') {
        msg += '\n⏰ ' + (r.tipo === 'pasadia' ? '9am – 5pm' : '12:30pm – 7pm');
      } else if (r.tipo === 'early') {
        msg += '\n⏰ Entrada anticipada 9am';
      }
      if (r.comentarios) msg += '\n📝 ' + r.comentarios;
    });
  }

  if (salidas.length > 0) {
    msg += '\n\n📤 *SALEN* (' + salidas.length + ')';
    salidas.forEach(it => {
      const r = it.reserva;
      msg += '\n\n🏡 ' + r.cabinName;
      msg += '\n👤 ' + r.name + ' · ' + r.persons + (r.persons == 1 ? ' persona' : ' personas');
      const phone = _adminFormatPhone(r.telefono);
      if (phone) msg += '\n💬 ' + phone;
      if (r.tipo === 'late') msg += '\n⏰ Sale 4pm (late check-out)';
    });
  }

  sendWhatsAppText(BOT_ADMIN_PHONE, msg);
  logDebugEntry('admin-recordatorio-11am', { count: items.length, entradas: entradas.length, salidas: salidas.length });
}

// ─── Trigger 9am: servicios especiales del dia ──────────────────
function enviarRecordatorioServiciosEspeciales() {
  const today = Utilities.formatDate(new Date(), BOT_TZ, 'yyyy-MM-dd');
  const items = _adminGetMovimientosDia(today);
  const conServicios = items
    .filter(i => i.kind === 'entrada' || i.kind === 'pasadia')
    .filter(i => _adminHasSpecialKw(i.reserva.comentarios));

  if (conServicios.length === 0) {
    logDebugEntry('admin-recordatorio-9am', { count: 0 });
    return;
  }

  const fechaLbl = _adminFmtFechaLarga(today);
  let msg = '🎁 *Servicios especiales hoy* — ' + fechaLbl + '\n\n' +
    'Estas reservas tienen notas que requieren coordinación:';
  conServicios.forEach(it => {
    const r = it.reserva;
    msg += '\n\n🏡 ' + r.cabinName;
    msg += '\n👤 ' + r.name;
    const phone = _adminFormatPhone(r.telefono);
    if (phone) msg += '\n💬 ' + phone;
    msg += '\n📝 ' + r.comentarios;
  });

  sendWhatsAppText(BOT_ADMIN_PHONE, msg);
  logDebugEntry('admin-recordatorio-9am', { count: conServicios.length });
}

// ─── Limpieza del dia ──────────────────────────────────────────
// Mensaje orientado a limpieza, por cabaña, segun el estado de hoy:
//  - Salida hoy           → 🧹 limpiar + cambiar sábanas (urgente si entra alguien hoy)
//  - Estadía multi-noche  → ✅ no limpiar (huésped sigue)
//  - Entra hoy, vacía ayer→ 👀 no cambiar sábanas, solo verificar
//  - Sin actividad        → ⚪ libre
// Se envia: (a) por trigger 8am a LIMPIEZA_PHONE, y (b) on-demand cuando
// ese numero le escribe al Agente (ver WhatsAppWebhook).
function _buildLimpiezaMessage(greeting) {
  const today     = Utilities.formatDate(new Date(), BOT_TZ, 'yyyy-MM-dd');
  const yesterday = _botAddDaysISO(today, -1);
  const fechaLbl  = _adminFmtFechaLarga(today);

  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();

  const CABINS = [
    { key: 'verde', name: 'Paseo por Las Nubes' },
    { key: 'azul',  name: 'Portal hacia Las Nubes' },
    { key: 'lila',  name: 'Puente entre Las Nubes' }
  ];
  const byCabin = { verde: [], azul: [], lila: [] };
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue;
    if (r[9] === 'Abierta') continue;
    if ((r[20] || '').toString().toUpperCase() === 'CANCELADA') continue;
    const cabin = r[3];
    if (!byCabin[cabin]) continue;
    const ci = r[4] instanceof Date ? Utilities.formatDate(r[4], BOT_TZ, 'yyyy-MM-dd') : (r[4] || '').toString().slice(0,10);
    const co = r[5] instanceof Date ? Utilities.formatDate(r[5], BOT_TZ, 'yyyy-MM-dd') : (r[5] || '').toString().slice(0,10);
    if (!ci || !co) continue;
    byCabin[cabin].push({ name: r[1] || '?', ci: ci, co: co });
  }

  let limpiar = 0;
  const lines = CABINS.map(cab => {
    const res = byCabin[cab.key];
    const checkoutToday = res.find(x => x.co === today);
    const checkinToday  = res.find(x => x.ci === today);
    const occLastNight  = res.find(x => x.ci <= yesterday && x.co > yesterday);
    const occTonight    = res.find(x => x.ci <= today && x.co > today);

    let line = '🏡 *' + cab.name + '*\n';
    if (checkoutToday) {
      limpiar++;
      line += '🧹 *LIMPIAR* — salió ' + checkoutToday.name + '. Cambiar sábanas y dejar la cabaña lista.';
      if (checkinToday) {
        line += '\n⚠️ ¡Hoy mismo llega ' + checkinToday.name + '! Dejarla lista a tiempo.';
      }
    } else if (occTonight && occLastNight) {
      line += '✅ *No limpiar* — ' + occTonight.name + ' sigue hospedado (estadía de varias noches).';
    } else if (checkinToday && !occLastNight) {
      limpiar++;
      line += '👀 Llega ' + checkinToday.name + ' hoy. Anoche estuvo vacía: no hace falta cambiar sábanas, solo pasá a verificar que todo esté en orden.';
    } else if (occTonight) {
      line += '✅ *No limpiar* — ocupada.';
    } else {
      line += '⚪ Libre, sin actividad hoy.';
    }
    return line;
  });

  const saludo = greeting || '¡Hola Erika! 🌿';
  const intro = limpiar > 0
    ? saludo + ' Acá está la limpieza de hoy:'
    : saludo + ' Hoy no hay cabañas que limpiar. Igual te dejo el estado de cada una:';
  return '🧹 *Limpieza de hoy* — ' + fechaLbl + '\n\n' + intro + '\n\n' + lines.join('\n\n');
}

// Trigger 8am: envia el parte de limpieza a LIMPIEZA_PHONE.
function enviarRecordatorioLimpieza() {
  const phone = PropertiesService.getScriptProperties().getProperty('LIMPIEZA_PHONE');
  if (!phone) { logDebugEntry('recordatorio-limpieza-no-phone', {}); return; }
  sendWhatsAppText(phone, _buildLimpiezaMessage('¡Buenos días, Erika! 🌿'));
  logDebugEntry('recordatorio-limpieza', { phone: phone });
}

// ─── Trigger 9am: instrucciones de check-out al huésped que sale hoy ──
// Envia la plantilla 'instrucciones_checkout' (con boton quick-reply
// "Ya me retiré") a cada reserva que sale hoy y tiene teléfono. Al tocar
// el boton, el Agente avisa al admin (portón) y a Erika (limpieza).
//
// REQUISITO de la plantilla en Meta:
//   - Nombre: instrucciones_checkout · idioma "English" (en)
//   - Body con variables posicionales {{1}} (nombre) y {{2}} (cabaña)
//   - 1 boton de Respuesta rápida (ej. "Ya me retiré")
function enviarRecordatoriosCheckout() {
  const today = Utilities.formatDate(new Date(), BOT_TZ, 'yyyy-MM-dd');
  const salidas = _adminGetMovimientosDia(today).filter(i => i.kind === 'salida');

  let enviados = 0;
  salidas.forEach(it => {
    const r = it.reserva;
    if (!r.telefono) return;                 // sin teléfono no podemos contactar
    if (r.origin === 'Airbnb') return;        // Airbnb gestiona su propio canal
    const firstName  = (r.name || '').toString().trim().split(/\s+/)[0] || 'amigo';
    const cabinName  = r.cabinName || BOT_CABIN_NAMES[r.cabin] || r.cabin;
    try {
      sendWhatsAppTemplate(
        r.telefono,
        'instrucciones_checkout',
        'en',                        // la plantilla quedó registrada como "English"
        [firstName, cabinName],      // posicionales: {{1}} nombre, {{2}} cabaña
        null,
        'checkout_' + r.id           // payload del boton quick-reply
      );
      enviados++;
    } catch(err) {
      logDebugEntry('checkout-template-FAIL', { id: r.id, error: err.message });
    }
  });
  logDebugEntry('recordatorios-checkout', { salidas: salidas.length, enviados: enviados });
}

// ─── Setup ────────────────────────────────────────────────────────
// Correr UNA VEZ desde el editor para instalar los triggers diarios.
function instalarTriggersAdminReminders() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    const h = t.getHandlerFunction();
    if (h === 'enviarRecordatorioAdminReservasHoy' || h === 'enviarRecordatorioServiciosEspeciales') {
      ScriptApp.deleteTrigger(t);
    }
  }
  ScriptApp.newTrigger('enviarRecordatorioServiciosEspeciales')
    .timeBased().everyDays(1).atHour(9).inTimezone(BOT_TZ).create();
  ScriptApp.newTrigger('enviarRecordatorioAdminReservasHoy')
    .timeBased().everyDays(1).atHour(11).inTimezone(BOT_TZ).create();
  Logger.log('✓ Triggers instalados:\n - 9am: enviarRecordatorioServiciosEspeciales\n - 11am: enviarRecordatorioAdminReservasHoy');
}

// Instala (o reinstala) el trigger diario 8am de limpieza. Correr UNA VEZ
// desde el editor DESPUÉS de configurar la Script Property LIMPIEZA_PHONE.
function instalarTriggerLimpieza() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'enviarRecordatorioLimpieza') ScriptApp.deleteTrigger(t);
  }
  ScriptApp.newTrigger('enviarRecordatorioLimpieza')
    .timeBased().everyDays(1).atHour(8).inTimezone(BOT_TZ).create();
  Logger.log('✓ Trigger de limpieza instalado: 8am diario → enviarRecordatorioLimpieza');
}

// Instala (o reinstala) el trigger diario 9am de check-out al huésped.
// Correr UNA VEZ tras aprobar la plantilla instrucciones_checkout con boton.
function instalarTriggerCheckout() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'enviarRecordatoriosCheckout') ScriptApp.deleteTrigger(t);
  }
  ScriptApp.newTrigger('enviarRecordatoriosCheckout')
    .timeBased().everyDays(1).atHour(9).inTimezone(BOT_TZ).create();
  Logger.log('✓ Trigger de check-out instalado: 9am diario → enviarRecordatoriosCheckout');
}

// Funciones para probar manualmente desde el editor
function _testRecordatorio11am()  { return enviarRecordatorioAdminReservasHoy(); }
function _testRecordatorio9am()   { return enviarRecordatorioServiciosEspeciales(); }
function _testRecordatorioLimpieza() { return enviarRecordatorioLimpieza(); }
function _testCheckout()          { return enviarRecordatoriosCheckout(); }

// Test SEGURO: envía la plantilla de checkout a TU número (no a clientes).
// OJO: si tocas el botón "Ya me retiré", se dispara el aviso real al admin
// y a Erika (es el flujo normal del botón).
function _testCheckoutAMiNumero() {
  const r = sendWhatsAppTemplate(
    '50769812266',
    'instrucciones_checkout',
    'en',
    ['Josh (PRUEBA)', 'Portal hacia Las Nubes'],
    null,
    'checkout_TEST'
  );
  Logger.log('Enviado: ' + JSON.stringify(r));
  return r;
}

// ─── Tests SEGUROS a tu número (no escriben a clientes) ──────────────
// Plantillas en español (es). Datos de prueba. Solo para ver el render.
function _testCheckinAMiNumero() {
  const r = sendWhatsAppTemplate('50769812266', 'recordatorio_checkin', 'es',
    ['María (PRUEBA)', 'Portal hacia Las Nubes', 'vie 5 al dom 7 de junio']);
  Logger.log('checkin: ' + JSON.stringify(r));
  return r;
}

function _testSaldoAMiNumero() {
  const r = sendWhatsAppTemplate('50769812266', 'recordatorio_saldo', 'es',
    ['María (PRUEBA)', 'Portal hacia Las Nubes', 'vie 5 al dom 7 de junio', '90.00']);
  Logger.log('saldo: ' + JSON.stringify(r));
  return r;
}

function _testReferidoAMiNumero() {
  const r = sendWhatsAppTemplate('50769812266', 'referido_postestadia', 'es',
    ['María (PRUEBA)', 'LN-AB12CD']);
  Logger.log('referido: ' + JSON.stringify(r));
  return r;
}

// confirmacion_reserva: named params + botón "Consultas y cambios" (es_PA).
function _testConfirmacionAMiNumero() {
  const r = sendWAReservaConfirmada({
    id:       'TEST-' + Date.now(),
    name:     'María Pérez (PRUEBA)',
    telefono: '50769812266',
    cabin:    'azul',
    checkin:  '2026-08-15',
    checkout: '2026-08-17',
    persons:  2,
    amount:   180,
    tipo:     'noche'
  });
  Logger.log('confirmacion: ' + JSON.stringify(r));
  return r;
}

// ─── Seguimiento diario de leads (8am) ───────────────────────────────
// Resumen de las conversaciones que quedaron en "eligiendo cierre" o
// "pagando" el día anterior, para que el admin les dé seguimiento personal
// y trate de cerrar la venta. Se manda como texto de sesión al admin.
const _SEGUIMIENTO_STEPS = {
  CHOOSING_DECOR:         '🤝 Eligiendo decoración',
  CHOOSING_CLOSE:         '🤝 Eligiendo cierre',
  OFFERING_PAYMENT:       '💳 Pagando (formas de pago)',
  AWAITING_VOUCHER_RETRY: '💳 Reintentando voucher',
  PENDING_HUMAN_BOOKING:  '🙋 Cierre asistido (pendiente de ingresar)'
};

function enviarSeguimientoDiario() {
  if (_botGetAlertConfig().seguimientoDiario === false) {
    logDebugEntry('seguimiento-diario', { skip: 'desactivado' });
    return;
  }
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Conversaciones');
  if (!sheet || sheet.getLastRow() < 2) return;

  // "Ayer" en zona horaria de Panamá (yyyy-MM-dd).
  const ayer    = new Date(Date.now() - 24 * 60 * 60 * 1000);
  const ayerStr = Utilities.formatDate(ayer, BOT_TZ, 'yyyy-MM-dd');

  const data  = sheet.getDataRange().getValues();
  const leads = [];
  for (let i = 1; i < data.length; i++) {
    const phone = (data[i][0] || '').toString();
    const step  = (data[i][1] || '').toString();
    if (!phone || !_SEGUIMIENTO_STEPS[step]) continue;
    const lastUpd = (data[i][2] || '').toString();      // "yyyy-MM-dd HH:mm:ss"
    if (lastUpd.slice(0, 10) !== ayerStr) continue;     // solo del día anterior
    let ctx = {};
    try { ctx = data[i][3] ? JSON.parse(data[i][3]) : {}; } catch(_) {}
    leads.push({ phone: phone, step: step, name: (data[i][4] || '').toString(), ctx: ctx });
  }

  if (!leads.length) {
    logDebugEntry('seguimiento-diario', { ayer: ayerStr, leads: 0 });
    return;   // nada que reportar → no molestamos al admin
  }

  let msg = '📋 *Seguimiento de leads — ' + _botFmtFecha(ayerStr) + '*\n' +
            'Quedaron a un paso de reservar ayer. Escríbeles para cerrar la venta 👇\n';
  leads.forEach((l, idx) => {
    const cabin  = BOT_CABIN_NAMES[l.ctx.cabin] || l.ctx.cabin || '';
    const dts    = l.ctx.dates;
    const fechas = (dts && dts.checkin) ? (_botFmtFecha(dts.checkin) + ' → ' + _botFmtFecha(dts.checkout)) : '';
    msg += '\n' + (idx + 1) + '. *' + (l.name || ('+' + l.phone)) + '*\n' +
           '   ' + _SEGUIMIENTO_STEPS[l.step] + '\n' +
           (cabin  ? '   🏡 ' + cabin + '\n'  : '') +
           (fechas ? '   📅 ' + fechas + '\n' : '') +
           '   💬 https://wa.me/' + l.phone + '\n';
  });
  msg += '\n_' + leads.length + ' lead' + (leads.length === 1 ? '' : 's') + ' para seguimiento._';

  try { sendWhatsAppText(BOT_ADMIN_PHONE, msg); }
  catch(e) { logDebugEntry('seguimiento-diario-FAIL', { error: e.message }); }
  logDebugEntry('seguimiento-diario', { ayer: ayerStr, leads: leads.length });
}

// Instala (o reinstala) el trigger diario 8am del resumen de seguimiento.
// Correr UNA VEZ desde el editor.
function instalarTriggerSeguimiento() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'enviarSeguimientoDiario') ScriptApp.deleteTrigger(t);
  }
  ScriptApp.newTrigger('enviarSeguimientoDiario')
    .timeBased().everyDays(1).atHour(8).inTimezone(BOT_TZ).create();
  Logger.log('✓ Trigger de seguimiento instalado: 8am diario → enviarSeguimientoDiario');
}

// Test manual desde el editor: arma y envía el resumen de AYER al admin.
function _testSeguimientoDiario() { return enviarSeguimientoDiario(); }
