/**
 * Recordatorios diarios al ADMIN via WhatsApp.
 *
 * - 11:00 am: resumen de entradas y salidas del dia (similar al
 *   "Reservas de hoy" del dashboard), con links wa.me para escribir
 *   a cada cliente con un tap.
 * - 9:00 am: alerta si alguna reserva del dia tiene servicios
 *   especiales mencionados en comentarios (decoración, traslado,
 *   cumpleaños, menú, etc.).
 *
 * Setup: correr UNA VEZ desde el editor `instalarTriggersAdminReminders()`.
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

// Funciones para probar manualmente desde el editor
function _testRecordatorio11am() { return enviarRecordatorioAdminReservasHoy(); }
function _testRecordatorio9am()  { return enviarRecordatorioServiciosEspeciales(); }
