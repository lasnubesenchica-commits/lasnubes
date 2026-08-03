// ═══════════════════════════════════════════════════════════
//  LAS NUBES · Gmail → Google Sheets Auto-Sync
//  Pega este código en Google Apps Script (script.google.com)
// ═══════════════════════════════════════════════════════════

// ─── CONFIGURACIÓN ──────────────────────────────────────────
const SHEET_ID   = '15CJLNbyq8BQns8ybMDsZ_QAA-wWE5sT9GCRzwum2Ujo';
const SHEET_NAME = 'Reservas';

const CABINS = {
  'Paseo por Las Nubes':   'verde',
  'Portal hacia Las Nubes': 'azul',
  'Puente entre Las Nubes': 'lila'
};

// Ventana de Gmail de los syncs que corren por TRIGGER (cada 15 min los de
// reservas y alteraciones, cada hora los de pagos y cancelaciones).
//
// Estuvo en 365d y eso hacía que cada corrida recorriera un año de hilos: aunque
// `processed` evita re-parsear, el `thread.getMessages()` y el `msg.getId()` son
// un viaje a Gmail POR HILO, así que el costo crece con el histórico y no con lo
// nuevo. Con tres triggers cada 15 minutos, las ejecuciones interactivas
// quedaban esperando en cola y el calendario tardaba hasta un minuto en cargar.
//
// 30 días es holgado: los triggers corren cada 15 minutos, así que un email
// tendría que pasar un mes entero sin procesarse para caerse de la ventana.
//
// La red de seguridad NO cambia: `reconciliarReservasAirbnb()` en Cleanup.gs
// barre 400 días comparando TODOS los emails contra la hoja y carga los que
// falten. Si los triggers estuvieran caídos más de un mes, esa es la que
// recupera lo perdido — conviene correrla de vez en cuando.
const SYNC_VENTANA = 'newer_than:30d';

// ─── FUNCIÓN PRINCIPAL ──────────────────────────────────────
function syncAirbnbReservations() {
  const sheet = getOrCreateSheet();
  const processed = getProcessedIds(sheet);

  const threads = GmailApp.search(
    'from:automated@airbnb.com subject:"Reserva confirmada:" ' + SYNC_VENTANA
  );

  let added = 0;
  // Fuera del loop: estaba releyendo la hoja Blacklist completa por cada email
  // nuevo que parseaba. Es la misma lista para toda la corrida.
  const blacklisted = getBlacklistedCodes();

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const msgId = msg.getId();
      if (processed.has(msgId)) return;

      const body = msg.getPlainBody();
      const msgDate = Utilities.formatDate(msg.getDate(), 'America/Panama', 'yyyy-MM-dd');
      const reservation = parseAirbnbEmail(body, msgId, msgDate);

      if (reservation) {
        const code = (reservation.confirmCode || '').toString().trim();
        if (blacklisted.has(code)) {
          Logger.log(`⛔ Blacklisted, ignorando: ${reservation.name} (${code})`);
        } else if (code && processed.has(code)) {
          // Guard anti-re-inserción. `processed` sale de las filas que HAY en la
          // hoja (ids + CodConfirmacion), así que al borrar una fila su msgId
          // desaparece del set y este email vuelve a verse como "nuevo": el sync
          // recreaba la reserva desde el email ORIGINAL, con el monto de antes de
          // cualquier modificación. Pasó con Yarisel Rangel (HMBBQXQ3QD): se borró
          // la fila duplicada y volvió con $149 en vez de los $169 vigentes.
          // Si ya existe una fila con ese código de confirmación, no se re-inserta.
          Logger.log(`↩ Ya hay una reserva con el código ${code}, no se re-inserta: ${reservation.name}`);
        } else {
          appendReservation(sheet, reservation);
          if (code) processed.add(code);   // evita duplicar dentro de esta misma corrida
          added++;
          Logger.log(`✓ Reserva agregada: ${reservation.name} - ${reservation.cabin}`);
        }
      }
    });
  });

  Logger.log(`Sync completado. ${added} reservas nuevas agregadas.`);
}

// ─── SYNC RESERVAS ACTUALIZADAS ────────────────────────────
// ─── SYNC ALTERACIONES (cambios de reserva) ────────────────
//
// Airbnb manda DOS emails por cada cambio, y ninguno alcanza solo:
//
//   1) La SOLICITUD. Asunto "<Huésped> quiere hacer un cambio en su reserva".
//      Trae el DETALLE ("VIAJEROS ORIGINALES 3 viajeros, 1 mascota /
//      VIAJEROS SOLICITADOS 4 viajeros") pero NO el código HM: solo un id de
//      alteración que no sirve para cruzar con la hoja.
//      OJO: el X-Template de este email tiene DOS formas según la época —
//      STAY_RESERVATION_ALTERATION_REQUESTED y, en minúsculas,
//      reservation/alteration/alteration_requested. Por eso se busca por
//      ASUNTO; filtrar por header perdería la mitad de los emails en silencio.
//
//   2) X-Template: ALTERATION_ACCEPTED
//      Asunto "Reserva actualizada". Trae el CÓDIGO HM en el link al
//      itinerario, pero NO dice qué cambió ("ya actualizamos el itinerario").
//
// Así que hay que EMPAREJARLOS por huésped + cabaña, tomando la solicitud
// pendiente más reciente anterior a la confirmación. Además la solicitud sola no
// debe tocar nada: el anfitrión puede rechazarla.
//
// OJO — una misma reserva puede tener VARIAS alteraciones a lo largo del tiempo,
// y una solicitud rechazada (o vencida) se queda 'solicitada' para siempre. Sin
// más cuidado, la confirmación de la alteración B terminaría emparejada con la
// solicitud A y aplicaría fechas que nadie pidió. Caso real: HMP5R5WYAF
// (Kenneth) tiene una confirmación del 29-mar-2026 y una solicitud del
// 6-may-2026 que nunca se aceptó (canceló el 8). Por eso:
//   · solo se empareja dentro de ALT_VENTANA_DIAS (Airbnb vence la solicitud del
//     huésped a las 72 h, así que una semana es de sobra);
//   · las solicitudes que pasan esa ventana sin confirmarse se cierran como
//     'sin_confirmacion' y dejan de ser candidatas;
//   · antes de escribir fechas se verifica que la fila tenga HOY las FECHAS
//     ORIGINALES del email. Si no coinciden, el emparejamiento es dudoso y se
//     marca para revisar en vez de pisar el rango bueno.
// La confirmación NO trae el id de alteración (solo el código HM), así que
// emparejar por id no es una opción.
//
// La versión anterior de esta función solo escribía '⚠ Verificar' en la columna
// Alerta y nunca aplicaba el cambio — por eso Yarisel quedó con 3 personas y
// $149 después de pasar a 4 viajeros y $169.
const ALT_VENTANA_DIAS = 7;

function syncAirbnbUpdates() {
  const sheet = getOrCreateSheet();
  const ss    = SpreadsheetApp.openById(SHEET_ID);

  let altSheet = ss.getSheetByName('Alteraciones');
  if (!altSheet) {
    altSheet = ss.insertSheet('Alteraciones');
    const h = ['FechaSolicitud', 'MsgIdSolicitud', 'Huesped', 'Cabana', 'AlteracionId',
               'Cambios', 'Estado', 'FechaAceptada', 'MsgIdAceptada', 'CodConfirmacion', 'Aplicado'];
    altSheet.getRange(1, 1, 1, h.length).setValues([h]);
    altSheet.getRange(1, 1, 1, h.length).setFontWeight('bold');
  }
  const altData  = altSheet.getDataRange().getValues();
  const yaSolic  = new Set(altData.slice(1).map(r => (r[1] || '').toString()).filter(Boolean));
  const yaAcept  = new Set(altData.slice(1).map(r => (r[8] || '').toString()).filter(Boolean));

  const norm = t => (t || '').toString().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, ' ').trim();
  const cabinDe = body => {
    for (const [cn, code] of Object.entries(CABINS)) {
      if (body.toUpperCase().indexOf(cn.toUpperCase()) >= 0) return { code: code, nombre: cn };
    }
    return { code: '', nombre: '' };
  };

  // ── Paso 1: solicitudes ────────────────────────────────────
  let nuevasSolic = 0;
  GmailApp.search('from:automated@airbnb.com subject:"quiere hacer un cambio" ' + SYNC_VENTANA)
    .forEach(th => th.getMessages().forEach(msg => {
      const msgId = msg.getId();
      if (yaSolic.has(msgId)) return;
      const subject = msg.getSubject();
      const body    = msg.getPlainBody();
      const bodyN   = body.replace(/\s+/g, ' ');
      const nombreM = subject.match(/^(.+?)\s+quiere hacer un cambio/i);
      const huesped = nombreM ? nombreM[1].trim() : '';
      const cab     = cabinDe(body);
      const altM    = bodyN.match(/reservation\/alteration\/(\d+)/i);
      const cambios = _airbnbCambiosSolicitados(bodyN);
      altSheet.appendRow([
        Utilities.formatDate(msg.getDate(), 'America/Panama', 'yyyy-MM-dd HH:mm'),
        msgId, huesped, cab.nombre, altM ? altM[1] : '',
        JSON.stringify(cambios), 'solicitada', '', '', '', ''
      ]);
      yaSolic.add(msgId);
      nuevasSolic++;
      Logger.log('◷ Cambio SOLICITADO por ' + huesped + ' (' + cab.nombre + '): '
        + cambios.map(c => c.que + ' ' + c.antes + ' → ' + c.despues).join(' | '));
    }));

  // ── Paso 2: confirmaciones ─────────────────────────────────
  const filas   = altSheet.getDataRange().getValues();      // releer con las nuevas
  const resData = sheet.getDataRange().getValues();
  let aplicados = 0, sinPareja = 0, sinFila = 0;

  GmailApp.search('from:automated@airbnb.com subject:"Reserva actualizada" ' + SYNC_VENTANA)
    .forEach(th => th.getMessages().forEach(msg => {
      const msgId = msg.getId();
      if (yaAcept.has(msgId)) return;
      const body  = msg.getPlainBody();
      const bodyN = body.replace(/\s+/g, ' ');
      const fecha = Utilities.formatDate(msg.getDate(), 'America/Panama', 'yyyy-MM-dd HH:mm');
      const codM  = bodyN.match(/reservations\/details\/([A-Z0-9]{10})/i)
                 || bodyN.match(/\b(HM[A-Z0-9]{8})\b/);
      const cod   = codM ? codM[1].toUpperCase() : '';
      const nomM  = bodyN.match(/reserva con\s+(.+?)\s+se ha actualizado/i);
      const huesped = nomM ? nomM[1].trim() : '';
      const cab   = cabinDe(body);

      // Emparejar con la solicitud pendiente más reciente del mismo huésped y
      // cabaña, anterior a esta confirmación.
      // Si Airbnb mandó "%{GUEST_NAME}" sin sustituir, el nombre no sirve: se
      // cae a cercanía temporal y solo se acepta si hay UN candidato.
      const sinNombre = _altNombreInutil(huesped);
      let mejor = -1, candidatos = 0;
      for (let k = 1; k < filas.length; k++) {
        if ((filas[k][6] || '').toString() !== 'solicitada') continue;
        if (!sinNombre && norm(filas[k][2]) !== norm(huesped)) continue;
        if (cab.nombre && norm(filas[k][3]) && norm(filas[k][3]) !== norm(cab.nombre)) continue;
        // _altTs: la celda puede volver como Date, no como el string que escribimos.
        if (_altTs(filas[k][0]) > fecha) continue;
        const gapMin = _altDiasEntre(filas[k][0], fecha) * 1440;
        if (sinNombre ? gapMin > ALT_FALLBACK_MIN : gapMin > ALT_VENTANA_DIAS * 1440) continue;
        candidatos++;
        if (mejor < 0 || _altTs(filas[k][0]) > _altTs(filas[mejor][0])) mejor = k;
      }
      if (sinNombre && candidatos !== 1) {
        Logger.log('⚠ Confirmación sin nombre utilizable (' + cod + '): '
          + candidatos + ' candidato(s) por cercanía; no se arriesga a emparejar.');
        mejor = -1;
      }

      if (mejor < 0) {
        sinPareja++;
        Logger.log('⚠ Confirmación de ' + huesped + ' (' + cod + ') sin solicitud emparejable '
          + 'en los ' + ALT_VENTANA_DIAS + ' días previos. Se marca la reserva para revisar a mano.');
        _altMarcarRevisar(sheet, resData, cod, 'Airbnb confirmó un cambio pero no se encontró el detalle.');
        altSheet.appendRow([fecha, '', huesped, cab.nombre, '', '[]', 'aceptada_sin_detalle',
          fecha, msgId, cod, 'no']);
        yaAcept.add(msgId);
        return;
      }

      const cambios = (function() { try { return JSON.parse(filas[mejor][5] || '[]'); } catch(_) { return []; } })();
      const res = _altAplicarCambios(sheet, resData, cod, cambios, fecha);
      altSheet.getRange(mejor + 1, 7).setValue('aceptada');
      altSheet.getRange(mejor + 1, 8).setValue(fecha);
      altSheet.getRange(mejor + 1, 9).setValue(msgId);
      altSheet.getRange(mejor + 1, 10).setValue(cod);
      altSheet.getRange(mejor + 1, 11).setValue(res.aplicado ? 'si' : 'parcial');
      yaAcept.add(msgId);
      if (res.filaEncontrada) aplicados++; else sinFila++;
      Logger.log((res.aplicado ? '✓ ' : '◐ ') + cod + ' (' + huesped + '): ' + res.detalle);
    }));

  // ── Paso 3: cerrar solicitudes colgadas ────────────────────
  // Una solicitud que pasó la ventana sin confirmarse fue rechazada o venció.
  // Si se queda 'solicitada' es una mina para la próxima confirmación.
  const ahora = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd HH:mm');
  const finales = altSheet.getDataRange().getValues();
  let vencidas = 0;
  for (let k = 1; k < finales.length; k++) {
    if ((finales[k][6] || '').toString() !== 'solicitada') continue;
    if (!_altTs(finales[k][0])) continue;                    // sin fecha no se cierra a ciegas
    if (_altDiasEntre(finales[k][0], ahora) <= ALT_VENTANA_DIAS) continue;
    altSheet.getRange(k + 1, 7).setValue('sin_confirmacion');
    vencidas++;
  }

  Logger.log('Alteraciones · ' + nuevasSolic + ' solicitud(es) nueva(s) · ' + aplicados
    + ' aplicada(s) · ' + sinPareja + ' sin detalle · ' + sinFila + ' sin fila en Reservas · '
    + vencidas + ' solicitud(es) cerrada(s) sin confirmar');
}

// Timestamp de la hoja → 'yyyy-MM-dd HH:mm'.
//
// IMPRESCINDIBLE. Escribimos la fecha como string con Utilities.formatDate, pero
// Sheets la reconoce como fecha y la guarda como VALOR: al releer con
// getValues() vuelve un objeto Date, no el string. Y ahí las dos defensas del
// emparejamiento fallaban en silencio:
//   · el orden se comparaba como string, y "Sat Mar 28 2026 22:40:00 GMT-0500"
//     es MAYOR que "2026-03-28 22:43" (la 'S' pesa más que el '2'), así que toda
//     solicitud parecía posterior a su confirmación → par descartado;
//   · _altDiasEntre no podía parsear ese formato → Infinity → fuera de ventana.
// Resultado: en el backfill inicial ninguna de las 26 alteraciones históricas se
// emparejó (todas quedaron 'sin_confirmacion' / 'aceptada_sin_detalle') aunque
// varias tenían la confirmación a 2 o 3 minutos de la solicitud.
function _altTs(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd HH:mm');
  return String(v).trim();
}

// Días entre dos timestamps. Acepta Date o string. Si alguno no parsea devuelve
// Infinity, así que el que compara descarta el par en vez de tomarlo por bueno.
function _altDiasEntre(a, b) {
  const p = s => {
    const t = _altTs(s).replace(' ', 'T');
    if (!t) return null;
    const d = new Date(t.length === 10 ? t + 'T00:00:00' : t);
    return isNaN(d.getTime()) ? null : d.getTime();
  };
  const ta = p(a), tb = p(b);
  if (ta === null || tb === null) return Infinity;
  return Math.abs(tb - ta) / 86400000;
}

// Extrae los bloques "X ORIGINALES … X SOLICITADOS" del cuerpo normalizado.
// En el texto plano los encabezados vienen en mayúsculas ("VIAJEROS ORIGINALES"),
// y el género de la segunda palabra cambia (solicitadOS / solicitadAS).
function _airbnbCambiosSolicitados(bodyN) {
  const out = [];
  const re = /([A-ZÁÉÍÓÚÑ]{4,})\s+ORIGINAL(?:ES)?\s+(.+?)\s+\1\s+SOLICITAD[AO]S?\s+(.+?)(?=\s+[A-ZÁÉÍÓÚÑ]{4,}\s+ORIGINAL|\s+Si aceptas|\s+Ir a la solicitud|$)/g;
  let m;
  while ((m = re.exec(bodyN)) !== null) {
    out.push({ que: m[1], antes: m[2].trim(), despues: m[3].trim() });
  }
  return out;
}

// Rango de fechas de un bloque "FECHAS ORIGINALES / SOLICITADAS".
// Formato real del email (confirmado con el caso Kenneth):
//     "9 de may. de 2026 - 10 de may. de 2026"
// `parseFecha` NO sirve acá: su regex es `(\d{1,2})\s+(mes)` y acá entre el día
// y el mes va un "de", así que devolvía null y ningún cambio de fecha se
// aplicaba nunca. Además el email trae el AÑO explícito, que es mejor que
// deducirlo de la fecha del mensaje (un cambio puede saltar de año).
function _altRangoFechas(txt, baseISO) {
  const t = String(txt == null ? '' : txt).replace(/\s+/g, ' ').trim();
  if (!t) return null;
  const MES = '(ene|feb|mar|abr|may|jun|jul|ago|sep|set|oct|nov|dic)[a-z]*\\.?';

  const fechas = [];
  const uno = new RegExp('(\\d{1,2})\\s+(?:de\\s+)?' + MES + '(?:\\s+de\\s+(\\d{4}))?', 'gi');
  let m;
  while ((m = uno.exec(t)) !== null) fechas.push(_altArmarFecha(m[1], m[2], m[3], baseISO));

  // Variante compacta "9 - 10 de may. de 2026": el primer día no lleva mes.
  if (fechas.length === 1) {
    const comp = t.match(new RegExp(
      '^(\\d{1,2})\\s*[-–—]\\s*(\\d{1,2})\\s+(?:de\\s+)?' + MES + '(?:\\s+de\\s+(\\d{4}))?$', 'i'));
    if (comp) {
      const a = _altArmarFecha(comp[1], comp[3], comp[4], baseISO);
      const b = _altArmarFecha(comp[2], comp[3], comp[4], baseISO);
      if (a && b) fechas.splice(0, 1, a, b);
    }
  }

  if (fechas.length < 2 || !fechas[0] || !fechas[1]) return null;
  if (fechas[1] <= fechas[0]) return null;          // check-out tiene que ser después
  return { ci: fechas[0], co: fechas[1] };
}

function _altArmarFecha(dia, mesTxt, anio, baseISO) {
  const d = parseInt(dia, 10);
  if (!(d >= 1 && d <= 31)) return null;
  const mm = monthToNum(String(mesTxt || '').slice(0, 3));
  if (!mm) return null;
  let y = anio ? parseInt(anio, 10) : 0;
  if (!y) y = _anioParaEstadia_(parseInt(mm, 10), d, baseISO) || getCurrentYear(parseInt(mm, 10));
  if (!(y >= 2000 && y <= 2100)) return null;
  return y + '-' + mm + '-' + (d < 10 ? '0' + d : '' + d);
}

// Encabezado del bloque de cambio, en MAYÚSCULAS y SIN TILDES.
//
// Airbnb cambió el rótulo con el tiempo: los emails de 2025 dicen "HUÉSPEDES
// ORIGINALES" y los de 2026 "VIAJEROS ORIGINALES". El código comparaba contra
// 'HUESPED' sin tilde, así que 'HUÉSPEDES'.indexOf('HUESPED') daba -1 y las 9
// alteraciones de 2025 caían al branch de "anotar para revisar": el cambio de
// personas quedaba registrado pero nunca se aplicaba.
function _altNormQue(s) {
  return String(s == null ? '' : s).toUpperCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '');
}

// Airbnb a veces manda la confirmación con la variable de plantilla SIN
// sustituir: el asunto trae "%{GUEST_NAME}" en vez del nombre. Ahí el huésped no
// sirve para emparejar y hay que caer a la cercanía temporal.
function _altNombreInutil(s) {
  const t = String(s == null ? '' : s).trim();
  return !t || /%\s*\{|GUEST_NAME/i.test(t);
}
const ALT_FALLBACK_MIN = 120;   // minutos, solo para el fallback sin nombre

// Celda de fecha de la hoja → 'yyyy-MM-dd' (puede venir Date o string).
function _altISO(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd');
  return String(v).trim().slice(0, 10);
}

// Busca la fila por código y le escribe un aviso en Alerta (col 14).
function _altMarcarRevisar(sheet, resData, cod, texto) {
  if (!cod) return false;
  for (let i = 1; i < resData.length; i++) {
    if ((resData[i][10] || '').toString().trim().toUpperCase() !== cod) continue;
    const prev = (resData[i][13] || '').toString().trim();
    const nota = '⚠ ' + texto;
    if (prev.indexOf(nota) < 0) sheet.getRange(i + 1, 14).setValue(prev ? prev + ' | ' + nota : nota);
    return true;
  }
  return false;
}

// Aplica a la reserva los cambios que se pueden interpretar sin ambigüedad.
// VIAJEROS se aplica solo (es un número). FECHAS se intenta parsear. Cualquier
// otra cosa —precio incluido, que ninguno de los dos emails trae— se deja
// anotada para que el admin decida: preferimos avisar antes que inventar plata.
function _altAplicarCambios(sheet, resData, cod, cambios, fechaConf) {
  const detalle = [];
  let fila = -1;
  for (let i = 1; i < resData.length; i++) {
    if ((resData[i][10] || '').toString().trim().toUpperCase() === cod) { fila = i + 1; break; }
  }
  if (fila < 0) {
    return { filaEncontrada: false, aplicado: false,
      detalle: 'no hay fila con ese código todavía; queda registrado en Alteraciones.' };
  }

  let todoAplicado = cambios.length > 0;
  const pendientes = [];
  cambios.forEach(c => {
    const que = _altNormQue(c.que);          // sin tildes: "HUÉSPEDES" → "HUESPEDES"
    if (que.indexOf('VIAJERO') === 0 || que.indexOf('HUESPED') === 0) {
      const n = parseInt((c.despues.match(/(\d+)/) || [])[1], 10);
      if (n > 0) {
        sheet.getRange(fila, 7).setValue(n);
        detalle.push('personas → ' + n);
        return;
      }
    }
    if (que.indexOf('FECHA') === 0) {
      const base    = fechaConf.slice(0, 10);
      const antes   = _altRangoFechas(c.antes, base);
      const despues = _altRangoFechas(c.despues, base);
      if (despues) {
        // Leer las fechas EN VIVO: si en esta misma corrida ya se aplicó otra
        // alteración a la fila, `resData` quedó viejo.
        const act   = sheet.getRange(fila, 5, 1, 2).getValues()[0];
        const ciAct = _altISO(act[0]), coAct = _altISO(act[1]);

        if (ciAct === despues.ci && coAct === despues.co) {
          detalle.push('fechas ya estaban en ' + despues.ci + ' a ' + despues.co);
          return;
        }
        // Solo pisar el rango si la fila tiene HOY las fechas que el email da
        // por originales. Si no, el email puede estar mal emparejado (ver el
        // comentario de syncAirbnbUpdates) y es preferible avisar.
        if (!antes || (ciAct === antes.ci && coAct === antes.co)) {
          sheet.getRange(fila, 5).setValue(despues.ci);
          sheet.getRange(fila, 6).setValue(despues.co);
          detalle.push('fechas → ' + despues.ci + ' a ' + despues.co);
          return;
        }
        todoAplicado = false;
        pendientes.push('FECHAS: la reserva está en ' + ciAct + '→' + coAct
          + ' pero el email cambia ' + antes.ci + '→' + antes.co
          + ' por ' + despues.ci + '→' + despues.co + ' (no coinciden, no se tocó)');
        return;
      }
    }
    todoAplicado = false;
    pendientes.push(c.que + ': ' + c.antes + ' → ' + c.despues);
  });

  if (pendientes.length) {
    _altMarcarRevisar(sheet, resData, cod, 'Cambio aceptado en Airbnb, revisar a mano: ' + pendientes.join(' ; '));
    detalle.push('pendiente de revisar: ' + pendientes.join(' ; '));
  }
  // Dejar constancia en Comentarios de qué cambió y cuándo.
  const nota = '🔄 ' + fechaConf + ' Airbnb aplicó un cambio: '
    + cambios.map(c => c.que + ' ' + c.antes + ' → ' + c.despues).join(' | ');
  try {
    const cAct = (resData[fila - 1][22] || '').toString().trim();
    if (cAct.indexOf(nota) < 0) sheet.getRange(fila, 23).setValue(cAct ? cAct + '\n' + nota : nota);
  } catch(_) {}

  return { filaEncontrada: true, aplicado: todoAplicado,
    detalle: detalle.length ? detalle.join(' · ') : 'nada que aplicar automáticamente.' };
}

// ─── PARSER DEL EMAIL ───────────────────────────────────────
// Monto de un email de Airbnb a número. El viejo `replace(',', '.')` reemplazaba
// SOLO la primera coma, así que "1,234.56" quedaba "1.234.56" → parseFloat 1.23.
// Acepta formato US (1,234.56) y europeo (1.234,56): el ÚLTIMO separador es el
// decimal cuando aparecen los dos.
function _montoAirbnbANumero(raw) {
  let t = String(raw == null ? '' : raw).trim().replace(/\s/g, '');
  // Sacar separadores colgando en los extremos (ej. "92,10." con el punto de la
  // oración), que corren el cálculo de cuál es el decimal.
  t = t.replace(/^[.,]+/, '').replace(/[.,]+$/, '');
  if (!t) return 0;
  const iComa = t.lastIndexOf(','), iPunto = t.lastIndexOf('.');
  if (iComa >= 0 && iPunto >= 0) {
    const decSep  = iComa > iPunto ? ',' : '.';
    const milSep  = decSep === ',' ? '.' : ',';
    t = t.split(milSep).join('');
    t = t.replace(decSep, '.');
  } else if (iComa >= 0) {
    // Solo comas: es decimal si quedan exactamente 2 dígitos después.
    t = (t.length - iComa - 1) === 2 ? t.replace(/,/g, '.') : t.split(',').join('');
  } else if ((t.match(/\./g) || []).length > 1) {
    t = t.split('.').join('');                 // varios puntos = separador de miles
  } else if (iPunto >= 0 && (t.length - iPunto - 1) === 3) {
    t = t.split('.').join('');                 // "1.234" = mil doscientos treinta y cuatro
  }
  const n = parseFloat(t);
  return isNaN(n) ? 0 : n;
}

// Monto que sigue a una etiqueta en el email de Airbnb ("Total (USD)",
// "Ganas"…), tolerando las DOS posiciones del símbolo de moneda.
//
// Airbnb cambió la plantilla: antes escribía "$243,00" y ahora "243,00 $" (con
// un espacio duro, =A0 en el iso-8859-1 del email). El regex viejo exigía el
// "$" ANTES de los dígitos:
//
//     /Total\s*\(USD\)[^$]*\$\s*([\d,.]+)/
//
// así que con la plantilla nueva no matcheaba y `amount` quedaba en 0 — sin
// alerta, porque el aviso "⚠ Monto no detectado" solo se agrega si además
// faltan otros datos. Resultado: TODA reserva de Airbnb creada después del
// cambio entró con monto cero, en silencio. Se detectó con Everett Richardson
// (HMSPN4MZFH): la app decía $243.00 y la hoja tenía $0.
function _montoEtiquetadoAirbnb_(body, etiquetaRegex) {
  // `[^\n$]{0,40}?` no cruza saltos de línea ni otro "$": evita que la etiqueta
  // se enganche con el monto de la fila siguiente si la suya viniera vacía.
  const re = new RegExp(etiquetaRegex + '[^\\n$]{0,40}?(?:\\$\\s*([\\d.,]+)|([\\d.,]+)\\s*\\$)', 'i');
  const m = String(body || '').match(re);
  if (!m) return 0;
  return _montoAirbnbANumero(m[1] || m[2]);
}

// Un `confirmCode` nunca puede ser un id interno. El dashboard manda
// `confirmCode: id` (correcto en reservas Directas, donde el código ES el id),
// pero en una reserva de Airbnb sin código el id llega como `airbnb_<msgId>` y
// terminaba escrito en la columna CodConfirmacion. Esa fila ya no cruza con
// ningún payout —sale "sin cobrar"— y la reconciliación cree que el HM real
// falta e inserta un DUPLICADO. Pasó con Yarisel, Yuliany y Kj Thomas.
function _sanitizeConfirmCode(cod) {
  const c = (cod == null ? '' : cod).toString().trim();
  if (!c) return '';
  if (/^(airbnb_|csv_)/i.test(c)) return '';
  return c;
}

function parseAirbnbEmail(body, msgId, msgDate) {
  try {
    const nameMatch = body.match(/([A-ZÁÉÍÓÚÜÑ][a-záéíóúüñ]+(?:\s+[A-ZÁÉÍÓÚÜÑ][a-záéíóúüñ]+){1,3})\s+\[https:\/\/www\.airbnb/);
    const name = nameMatch ? nameMatch[1].trim() : extractFallbackName(body);

    // OJO: antes esto arrancaba en 'verde', así que un email cuyo nombre de
    // cabaña no matcheara entraba a Paseo EN SILENCIO. Fue lo que puso a Mario
    // De León en verde cuando Airbnb lo tenía en azul, y de ahí el choque
    // fantasma con Dianeth. Ahora se deja sin cabaña y se marca la fila: mejor
    // que salte a la vista que ensuciar la ocupación de una cabaña real.
    let cabin = '';
    let cabinName = '';
    for (const [cn, code] of Object.entries(CABINS)) {
      if (body.toUpperCase().includes(cn.toUpperCase())) {
        cabin = code;
        cabinName = cn;
        break;
      }
    }

    let checkin = null, checkout = null;

    const fechasMatch = body.match(/Llegada\s+Salida\s+(\w+,?\s+\d{1,2}\s+\w+)\s+(\w+,?\s+\d{1,2}\s+\w+)/i);
    if (fechasMatch) {
      checkin  = parseFecha(fechasMatch[1], msgDate);
      checkout = parseFecha(fechasMatch[2], msgDate);
    }

    if (!checkin || !checkout) {
      const llegadaSalidaBlock = body.match(/Llegada[\s\S]{0,80}?Salida[\s\S]{0,200}?(\w+,?\s+\d{1,2}\s+\w+)[\s\S]{0,80}?(\w+,?\s+\d{1,2}\s+\w+)/i);
      if (llegadaSalidaBlock) {
        const d1 = parseFecha(llegadaSalidaBlock[1], msgDate);
        const d2 = parseFecha(llegadaSalidaBlock[2], msgDate);
        if (d1 && d2 && d1 !== d2) {
          checkin  = d1;
          checkout = d2;
        }
      }
    }

    const nochesMatch = body.match(/por\s+(\d+)\s+noche/i);
    const nochesEsperadas = nochesMatch ? parseInt(nochesMatch[1]) : null;

    if (checkin && checkout && nochesEsperadas) {
      const d1 = new Date(checkin + 'T12:00:00');
      const d2 = new Date(checkout + 'T12:00:00');
      const nochesActuales = Math.round((d2 - d1) / 86400000);
      if (nochesActuales !== nochesEsperadas) {
        Logger.log(`⚠ Noches calculadas (${nochesActuales}) ≠ noches en email (${nochesEsperadas}). Corrigiendo checkout.`);
        const corrected = new Date(d1);
        corrected.setDate(corrected.getDate() + nochesEsperadas);
        checkout = corrected.toISOString().slice(0, 10);
      }
    }

    if (!checkin || !checkout) {
      const allDates = [...body.matchAll(/\b(\d{1,2})\s+(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)\b/gi)]
        .filter(m => parseInt(m[1]) >= 1 && parseInt(m[1]) <= 31)
        .map(m => m[0]);

      const uniqueDates = [...new Set(allDates)];

      if (uniqueDates.length >= 2) {
        checkin  = parseFecha(uniqueDates[0], msgDate);
        checkout = parseFecha(uniqueDates[1], msgDate);
        if (checkin && checkout && checkin === checkout) {
          const d = new Date(checkout + 'T12:00:00');
          d.setDate(d.getDate() + (nochesEsperadas || 1));
          checkout = d.toISOString().slice(0, 10);
        }
      } else if (uniqueDates.length === 1) {
        checkin = parseFecha(uniqueDates[0], msgDate);
        if (checkin) {
          const d = new Date(checkin + 'T12:00:00');
          d.setDate(d.getDate() + (nochesEsperadas || 1));
          checkout = d.toISOString().slice(0, 10);
        }
      }
    }

    if (!checkin) {
      const subjectMatch = body.match(/LLEGA EL\s+(\d{1,2}\s+\w+)/i);
      if (subjectMatch) {
        checkin = parseFecha(subjectMatch[1], msgDate);
        if (checkin) {
          const d = new Date(checkin + 'T12:00:00');
          d.setDate(d.getDate() + 1);
          checkout = d.toISOString().slice(0, 10);
        }
      }
    }

    // Antes solo contaba adultos, así que "2 adultos, 2 niños" entraba como 2.
    // Los bebés no se cuentan (Airbnb tampoco los cuenta para la capacidad).
    const adultosM = body.match(/(\d+)\s+adulto/i);
    const ninosM   = body.match(/(\d+)\s+ni[ñn]o/i);
    const adultos  = adultosM ? parseInt(adultosM[1], 10) : 0;
    const ninos    = ninosM   ? parseInt(ninosM[1],   10) : 0;
    const persons  = (adultos + ninos) || 2;

    const amount = _montoEtiquetadoAirbnb_(body, 'Total\\s*\\(USD\\)');

    // Código de confirmación: varios patrones antes de rendirse. Con el código
    // vacío la reserva no puede cruzarse nunca con un payout, y la
    // reconciliación cree que el HM real falta e inserta un DUPLICADO.
    let confirmCode = '';
    const patronesCod = [
      /\/details\/([A-Z0-9]{10})/,          // link a la reserva
      /\/reservation\/itinerary\?[^\s]*code=([A-Z0-9]{10})/i,
      /\b(HM[A-Z0-9]{8})\b/,                // el código suelto en el cuerpo
      /c[oó]digo de confirmaci[oó]n[:\s]+([A-Z0-9]{10})/i
    ];
    for (let p = 0; p < patronesCod.length; p++) {
      const m = body.match(patronesCod[p]);
      if (m && m[1]) { confirmCode = m[1].toUpperCase(); break; }
    }

    Logger.log(`  Nombre: ${name}`);
    Logger.log(`  Cabaña: ${cabinName} (${cabin})`);
    Logger.log(`  Entrada: ${checkin} | Salida: ${checkout}`);
    Logger.log(`  Personas: ${persons} | Monto: $${amount} | Código: ${confirmCode}`);

    if (!name || !checkin || !checkout) {
      Logger.log('⚠ No se pudieron extraer datos completos del email');
      Logger.log(`  → name="${name}" checkin="${checkin}" checkout="${checkout}"`);
      return null;
    }

    // Marcas para que un dato faltante se vea en el dashboard en vez de quedar
    // como un valor plausible pero inventado.
    const avisos = [];
    if (!cabin)       avisos.push('⚠ Cabaña no detectada — asignar a mano');
    if (!confirmCode) avisos.push('⚠ Sin código de confirmación — pegar el HM de Airbnb');
    if (!amount)      avisos.push('⚠ Monto no detectado');
    if (avisos.length) Logger.log('  ' + avisos.join(' | '));

    return {
      id: msgId,
      name,
      cabin,
      cabinName,
      checkin,
      checkout,
      persons,
      amount,
      deposit: 0,
      origin: 'Airbnb',
      confirmCode,
      alerta: avisos.join(' | '),
      bookingDate: msgDate || Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd'),
      addedAt: new Date().toISOString()
    };

  } catch(e) {
    Logger.log('Error parseando email: ' + e.message);
    return null;
  }
}

// ─── EXTRACTOR DE FECHAS ────────────────────────────────────
// Deduce el AÑO de una estadía a partir de la fecha en que se hizo la reserva.
// Regla: la estadía es SIEMPRE posterior (o igual) a la fecha de reserva, así
// que se toma el año de la fecha base y se suma 1 solo si ese día ya pasó.
// Es determinístico: no depende de cuándo corra el sync.
function _anioParaEstadia_(monthNum, dayNum, baseISO) {
  const base = new Date(baseISO + 'T00:00:00');
  if (isNaN(base.getTime())) return null;
  let y = base.getFullYear();
  if (new Date(y, monthNum - 1, dayNum, 12, 0, 0) < base) y++;
  return y;
}

// `baseISO` = fecha del email (yyyy-MM-dd). Si no se pasa, cae al método viejo
// basado en la fecha de HOY — que es justamente el que produce años erróneos
// cuando el sync corre tarde o se reprocesa un email viejo (ej. re-parsear en
// julio una estadía de marzo devolvía el año siguiente).
function parseFecha(str, baseISO) {
  if (!str) return null;
  const match = str.match(/(\d{1,2})\s+(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)/i);
  if (!match) return null;
  const day   = match[1].padStart(2, '0');
  const month = monthToNum(match[2]);
  let year = null;
  if (baseISO) year = _anioParaEstadia_(parseInt(month, 10), parseInt(match[1], 10), baseISO);
  if (!year)   year = getCurrentYear(parseInt(month));
  return `${year}-${month}-${day}`;
}

function monthToNum(m) {
  const map = {
    ene:'01', feb:'02', mar:'03', abr:'04', may:'05', jun:'06',
    jul:'07', ago:'08', sep:'09', oct:'10', nov:'11', dic:'12'
  };
  return map[m.toLowerCase()] || '01';
}

function getCurrentYear(monthNum) {
  const now = new Date();
  const currentMonth = now.getMonth() + 1;
  const diff = currentMonth - monthNum;
  return diff > 3 ? now.getFullYear() + 1 : now.getFullYear();
}

function extractFallbackName(body) {
  const m = body.match(/CONFIRMADA!\s+([A-ZÁÉÍÓÚÜÑ][A-ZÁÉÍÓÚÜÑa-záéíóúüñ]+)\s+LLEGA/i);
  return m ? m[1] : 'Huésped Airbnb';
}

// ─── GOOGLE SHEET ───────────────────────────────────────────
function getOrCreateSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 33).setValues([[
      'ID', 'Nombre', 'Cabaña', 'CabañaCodigo',
      'Entrada', 'Salida', 'Personas',
      'Monto', 'Abono', 'Origen', 'CodConfirmacion',
      'ServiceFee', 'Neto', 'Alerta', 'Pagador', 'FechaReserva',
      'FechaPago', 'MontoPagado', 'CodTransferencia', 'MontoVoucher', 'EstadoPago',
      'Email', 'Comentarios', 'Telefono', 'Tipo', 'VoucherURL',
      'IdHuespedURL', 'FechaNacimiento', 'CheckoutExtendido',
      'HoraEntrada', 'HoraSalida', 'VouchersMeta', 'Regalo'
    ]]);
    sheet.getRange(1, 1, 1, 33).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    // Auto-asegurar columnas Tipo (25), VoucherURL (26), IdHuespedURL (27),
    // FechaNacimiento (28), CheckoutExtendido (29), HoraEntrada (30),
    // HoraSalida (31), VouchersMeta (32), Regalo (33).
    // Idempotente: cada llamada rellena lo que falte.
    if (sheet.getLastColumn() < 33) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (!headers.includes('Tipo')) {
        sheet.getRange(1, 25).setValue('Tipo');
        sheet.getRange(1, 25).setFontWeight('bold');
      }
      if (!headers.includes('VoucherURL')) {
        sheet.getRange(1, 26).setValue('VoucherURL');
        sheet.getRange(1, 26).setFontWeight('bold');
      }
      if (!headers.includes('IdHuespedURL')) {
        sheet.getRange(1, 27).setValue('IdHuespedURL');
        sheet.getRange(1, 27).setFontWeight('bold');
      }
      if (!headers.includes('FechaNacimiento')) {
        sheet.getRange(1, 28).setValue('FechaNacimiento');
        sheet.getRange(1, 28).setFontWeight('bold');
      }
      if (!headers.includes('CheckoutExtendido')) {
        sheet.getRange(1, 29).setValue('CheckoutExtendido');
        sheet.getRange(1, 29).setFontWeight('bold');
      }
      if (!headers.includes('HoraEntrada')) {
        sheet.getRange(1, 30).setValue('HoraEntrada');
        sheet.getRange(1, 30).setFontWeight('bold');
      }
      if (!headers.includes('HoraSalida')) {
        sheet.getRange(1, 31).setValue('HoraSalida');
        sheet.getRange(1, 31).setFontWeight('bold');
      }
      if (!headers.includes('VouchersMeta')) {
        sheet.getRange(1, 32).setValue('VouchersMeta');
        sheet.getRange(1, 32).setFontWeight('bold');
      }
      if (!headers.includes('Regalo')) {
        sheet.getRange(1, 33).setValue('Regalo');
        sheet.getRange(1, 33).setFontWeight('bold');
      }
    }
  }

  return sheet;
}

// ─── CONFIG (Tarifas) ────────────────────────────────────────
// Cada `getDataRange().getValues()` es un viaje al backend de Sheets, y Config
// se lee varias veces en una sola petición: la migración de tarifas, después
// `tarifasPublicas()`, después el handler de getTarifas. Se memoiza por
// EJECUCIÓN — las globales de Apps Script se reinician en cada invocación, así
// que no hay riesgo de servir un valor viejo a la petición siguiente. Cualquier
// escritura sobre Config tiene que llamar a `_invalidarConfigCache()`.
let _configRowsCache = null;
// Mismo memo por ejecución para la hoja Feriados: en una sola petición se pide
// desde getTarifas y desde el motor de precios. Se invalida al sembrar y al
// guardar fechas desde el modal de Tarifas. Declarado acá arriba porque
// getOrCreateFeriados —que lo limpia— está antes que getFechasEspeciales.
let _feriadosCache = null;
// Las filas crudas de Feriados, para que getFechasEspeciales no repita la
// lectura que ya hizo getOrCreateFeriados para decidir si sembrar.
let _feriadosRowsCache = null;
function _configRows(cfg) {
  if (!_configRowsCache) _configRowsCache = cfg.getDataRange().getValues();
  return _configRowsCache;
}
function _invalidarConfigCache() { _configRowsCache = null; _tarifasCache = null; }

function getOrCreateConfig() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let cfg  = ss.getSheetByName('Config');
  if (!cfg) {
    cfg = ss.insertSheet('Config');
    cfg.getRange(1, 1, 1, 5).setValues([['Clave', 'Valor', 'Descripcion', 'UltimaActualizacion', 'ActualizadoPor']]);
    cfg.getRange(1, 1, 1, 5).setFontWeight('bold');
    cfg.setFrozenRows(1);
    const hoy = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');
    cfg.appendRow(['weekday',     90,      'Tarifa Dom-Jue',             hoy, 'sistema']);
    cfg.appendRow(['weekend',     110,     'Tarifa Vie-Sab (legado)',    hoy, 'sistema']);
    cfg.appendRow(['promo',       75,      'Tarifa promocional Dom-Jue', hoy, 'sistema']);
    cfg.appendRow(['promoActive', 'false', 'Promo activa (true/false)',  hoy, 'sistema']);
    Logger.log('✓ Hoja Config creada con valores por defecto');
  }
  _migrarTarifasPorTipoDia_(cfg);
  return cfg;
}

// Tarifas por tipo de día. Antes viernes y sábado compartían `weekend` y el
// promocional solo existía para entre semana; ahora cada tipo tiene su precio
// normal y su promocional, y se agregan feriado y víspera de feriado.
//
// Se siembra desde lo que ya había —viernes y sábado heredan `weekend`— para
// que el precio publicado no cambie el día que se despliega esto. `weekend`
// queda en la hoja como legado: ya nadie lo lee, pero borrarlo sería tirar el
// valor del que salieron los nuevos.
// [clave, heredaDe, descripción, valorPorDefecto]
// `heredaDe` toma el valor de otra clave ya existente; si es null se usa el
// cuarto elemento, o 0.
const _TARIFAS_NUEVAS = [
  ['viernes',       'weekend', 'Tarifa viernes'],
  ['sabado',        'weekend', 'Tarifa sábado'],
  ['promoViernes',  null,      'Tarifa promocional viernes (0 = sin promo)'],
  ['promoSabado',   null,      'Tarifa promocional sábado (0 = sin promo)'],
  ['vispera',       'weekend', 'Tarifa víspera de feriado'],
  ['promoVispera',  null,      'Tarifa promocional víspera (0 = sin promo)'],
  ['feriado',       'weekend', 'Tarifa feriado'],
  ['promoFeriado',  null,      'Tarifa promocional feriado (0 = sin promo)'],
  ['escolar',       'weekend', 'Tarifa vacaciones escolares'],
  ['promoEscolar',  null,      'Tarifa promocional vacaciones escolares (0 = sin promo)'],
  // Cuántos meses dura la promo contando desde el ACTUAL. 0 = sin límite.
  // Es una cantidad y no una fecha de fin a propósito: así la ventana se corre
  // sola al pasar el mes, en vez de vencer y obligar a reconfigurar.
  ['promoMeses',    null,      'Meses de promo desde el actual (0 = sin límite)'],

  // Estadías cortas y recargos. Estaban en el código del frontend y getTarifas
  // ya las leía, pero la hoja nunca creaba la fila: sin fila, saveTarifas no
  // tenía dónde escribir y el valor por defecto ganaba siempre. Se siembran con
  // exactamente lo que estaba hardcodeado, así el precio no cambia.
  ['pasatarde',               null, 'Tarifa pasatarde 12:30pm–7pm',        60],
  ['pasanoche',               null, 'Tarifa pasanoche 8pm–12:30pm',        75],
  ['pasadia',                 null, 'Tarifa pasadía 9am–5pm',              75],
  ['recargoPasatardePersona', null, 'Recargo pasatarde por persona (3ra+)', 20],
  ['recargoPasanochePersona', null, 'Recargo pasanoche por persona (3ra+)', 25],
  ['recargoPasadiaPersona',   null, 'Recargo pasadía por persona (3ra+)',   25],
  ['recargoPersonaGrande',    null, 'Recargo por persona Paseo/Puente',     20],
  ['recargoPersonaPortal',    null, 'Recargo por persona Portal',           10],
  // Descuento del combo Puente+Portal sobre el precio de las dos por separado.
  // Reemplaza a recargoCombo5/6, que eran montos planos: no seguían la tarifa
  // de la noche, así que solo cuadraban en un tipo de día.
  ['comboDescuento',          null, 'Descuento combo por noche',            30],
  // Early check-in y late check-out son una NOCHE con horas extra, así que se
  // cobran como recargo sobre el hospedaje y no como tarifa propia: heredan
  // solas el precio del tipo de día, la promo y el recargo por persona.
  // Por estadía, no por noche — la cortesía es solo el primer o último día.
  ['recargoEarly',            null, 'Recargo early check-in (por estadía)',  25],
  ['recargoLate',             null, 'Recargo late check-out (por estadía)',  25],

  // Malaya Lodge. Estaban escritas DOS veces en el HTML del dashboard: en el
  // bloque visible de tarifas y, por separado, dentro de calcTarifa. Subir el
  // precio obligaba a acordarse de los dos lugares, y olvidarse de uno hacía
  // que el bloque mostrara un número y el modal cotizara otro. Se siembran con
  // exactamente lo que estaba hardcodeado, así el precio no cambia.
  ['malayaSemana',        null, 'Malaya · tarifa Dom-Jue',            75],
  ['malayaFinde',         null, 'Malaya · tarifa Vie-Sáb',           100],
  ['malayaPersonaExtra',  null, 'Malaya · recargo por persona (3ra+)', 30]
];

function _migrarTarifasPorTipoDia_(cfg) {
  const rows = _configRows(cfg);
  const map  = {};
  for (let i = 1; i < rows.length; i++) {
    const k = rows[i][0] ? rows[i][0].toString().trim() : '';
    if (k) map[k] = rows[i][1];
  }
  const hoy = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');
  let añadidas = 0;
  _TARIFAS_NUEVAS.forEach(function(t) {
    const clave = t[0], hereda = t[1], desc = t[2], porDefecto = t[3];
    if (map.hasOwnProperty(clave)) return;
    const valor = hereda ? (parseFloat(map[hereda]) || 110) : (porDefecto || 0);
    cfg.appendRow([clave, valor, desc, hoy, 'sistema']);
    añadidas++;
  });
  if (añadidas) {
    _invalidarConfigCache();   // se agregaron filas: lo memoizado quedó corto
    Logger.log('✓ Config: ' + añadidas + ' tarifas por tipo de día agregadas');
  }
}

// ─── PRECIO PÚBLICO POR NOCHE ────────────────────────────────
// ESPEJO de getPrecio() en index.html. Si se toca la lógica de precios hay que
// tocar los dos: uno corre en el navegador y el otro en Apps Script, no hay
// forma de compartir el código. La alternativa era peor — el bot tenía su
// propia tabla ($90 entre semana / $110 fin de semana) y por eso cotizaba $90
// en una víspera de feriado que el calendario cobra a $135.
//
// Reglas, iguales que en el front:
//  · una fecha puede caer en varias categorías (un sábado que además es
//    víspera) y GANA LA MÁS ALTA;
//  · el promocional de una categoría solo aplica si está activo, si vale > 0 y
//    si la fecha cae en la ventana de meses vigente.

let _tarifasCache = null;   // memo por ejecución; el bot lee esto por mensaje
function tarifasPublicas() {
  if (_tarifasCache) return _tarifasCache;
  const rows = _configRows(getOrCreateConfig());
  const m = {};
  for (let i = 1; i < rows.length; i++) {
    const k = rows[i][0] ? rows[i][0].toString().trim() : '';
    if (k) m[k] = rows[i][1];
  }
  const n = (k, d) => { const v = parseFloat(m[k]); return isNaN(v) ? d : v; };
  _tarifasCache = {
    semana:  n('weekday', 90),  promoSemana:  n('promo', 0),
    viernes: n('viernes', 110), promoViernes: n('promoViernes', 0),
    sabado:  n('sabado', 110),  promoSabado:  n('promoSabado', 0),
    vispera: n('vispera', 110), promoVispera: n('promoVispera', 0),
    feriado: n('feriado', 110), promoFeriado: n('promoFeriado', 0),
    escolar: n('escolar', 110), promoEscolar: n('promoEscolar', 0),
    promoActiva: m['promoActive'] === true || String(m['promoActive']).toLowerCase() === 'true',
    promoMeses:  parseInt(m['promoMeses'], 10) || 0
  };
  return _tarifasCache;
}

// Último mes 'yyyy-MM' con promo, o '' si no hay límite. Se calcula desde hoy,
// así la ventana se corre sola al cambiar de mes.
function _promoUltimoMesBackend() {
  const t = tarifasPublicas();
  if (!t.promoMeses || t.promoMeses <= 0) return '';
  const d = new Date();
  d.setDate(1);
  d.setMonth(d.getMonth() + t.promoMeses - 1);
  return Utilities.formatDate(d, 'America/Panama', 'yyyy-MM');
}

function _promoVigenteBackend(dateStr) {
  const t = tarifasPublicas();
  if (!t.promoActiva) return false;
  const hasta = _promoUltimoMesBackend();
  if (!hasta) return true;
  const mes = String(dateStr).slice(0, 7);
  const hoyMes = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM');
  return mes >= hoyMes && mes <= hasta;
}

function precioNochePublico(dateStr) {
  const t = tarifasPublicas();
  const esp = getFechasEspeciales();
  const dow = new Date(dateStr + 'T12:00:00').getDay();
  const cats = [dow === 5 ? 'viernes' : dow === 6 ? 'sabado' : 'semana'];
  if (esp.feriados[dateStr]) cats.push('feriado');
  // La víspera se ata solo a feriados; dentro de una semana escolar cada día
  // sería víspera del siguiente y la categoría perdería sentido.
  const manana = Utilities.formatDate(
    new Date(new Date(dateStr + 'T12:00:00').getTime() + 86400000), 'America/Panama', 'yyyy-MM-dd');
  if (esp.feriados[manana]) cats.push('vispera');
  if (esp.escolares[dateStr]) cats.push('escolar');

  const PROMO = { semana:'promoSemana', viernes:'promoViernes', sabado:'promoSabado',
                  vispera:'promoVispera', feriado:'promoFeriado', escolar:'promoEscolar' };
  let max = 0;
  cats.forEach(function(c) {
    const base  = t[c] || t.semana;
    const promo = t[PROMO[c]] || 0;
    const val = (promo > 0 && _promoVigenteBackend(dateStr)) ? promo : base;
    if (val > max) max = val;
  });
  return max;
}

// ─── FERIADOS ────────────────────────────────────────────────
// Hoja editable: el admin puede borrar los que no le interesan (un feriado no
// siempre llena la cabaña) y agregar fechas propias — un puente local, una
// feria de la zona. Por eso la lista vive en la hoja y no en el código.
//
// Se siembra con los feriados nacionales de Panamá del año en curso y el
// siguiente. Los móviles (carnaval y viernes santo) se calculan desde la fecha
// de Pascua, no se hardcodean: hardcodearlos significa que en enero de cada año
// las tarifas de carnaval quedan mal sin que nadie se entere.
function getOrCreateFeriados() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh   = ss.getSheetByName('Feriados');
  if (!sh) {
    sh = ss.insertSheet('Feriados');
    sh.getRange(1, 1, 1, 3).setValues([['Fecha', 'Nombre', 'Tipo']]);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  // Columna Tipo para hojas que se crearon antes: `feriado` | `escolar`. Vacío
  // se lee como feriado, que es lo único que existía.
  if (sh.getLastColumn() < 3) sh.getRange(1, 3).setValue('Tipo');
  // Las tres siembras solo necesitan saber si ya existe su fila centinela, así
  // que comparten UNA lectura en vez de hacer una cada una. Antes esto costaba
  // tres viajes a Sheets en cada petición, más el de getFechasEspeciales: cuatro
  // para leer la misma hoja de ~60 filas.
  const rows = sh.getDataRange().getValues();
  _feriadosRowsCache = rows;
  const año = new Date().getFullYear();
  let escribio = false;
  escribio = _sembrarFeriados_(sh, año, rows)     || escribio;
  escribio = _sembrarFeriados_(sh, año + 1, rows) || escribio;
  escribio = _sembrarEscolares_(sh, rows)         || escribio;
  if (escribio) { _feriadosCache = null; _feriadosRowsCache = null; }
  return sh;
}

// Recesos escolares de MEDUCA. Solo 2026: el calendario de cada año se publica
// alrededor de septiembre del anterior, así que hardcodear 2027 hoy sería
// inventar fechas. Cuando salga, se cargan desde la pestaña de Tarifas.
const _RECESOS_ESCOLARES = {
  2026: [
    ['2026-06-01', '2026-06-05', 'Receso 1er trimestre'],
    ['2026-09-07', '2026-09-11', 'Receso 2do trimestre']
  ]
};

// `rows` viene de quien llama para no releer la hoja. Devuelve true si escribió.
function _sembrarEscolares_(sh, rows) {
  let escribio = false;
  Object.keys(_RECESOS_ESCOLARES).forEach(function(año) {
    const marca = '__escolar__' + año;
    for (let i = 1; i < rows.length; i++) {
      if ((rows[i][0] || '').toString().trim() === marca) return;
    }
    const filas = [];
    _RECESOS_ESCOLARES[año].forEach(function(r) {
      const desde = new Date(r[0] + 'T12:00:00'), hasta = new Date(r[1] + 'T12:00:00');
      for (let d = new Date(desde); d <= hasta; d.setDate(d.getDate() + 1)) {
        filas.push([_isoDia_(d), r[2], 'escolar']);
      }
    });
    filas.push([marca, 'no borrar — evita resembrar ' + año, '']);
    sh.getRange(sh.getLastRow() + 1, 1, filas.length, 3).setValues(filas);
    escribio = true;
    Logger.log('✓ Recesos escolares ' + año + ' sembrados (' + (filas.length - 1) + ' días)');
  });
  return escribio;
}

// Domingo de Pascua (algoritmo gregoriano anónimo, Meeus/Jones/Butcher).
function _pascua_(año) {
  const a = año % 19, b = Math.floor(año / 100), c = año % 100;
  const d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4), k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const mes = Math.floor((h + l - 7 * m + 114) / 31);
  const dia = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(año, mes - 1, dia);
}

function _isoDia_(d) {
  return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
}
function _masDias_(d, n) { const x = new Date(d.getTime()); x.setDate(x.getDate() + n); return x; }

function feriadosPanama(año) {
  const p = _pascua_(año);
  return [
    [año + '-01-01', 'Año Nuevo'],
    [año + '-01-09', 'Día de los Mártires'],
    [_isoDia_(_masDias_(p, -48)), 'Lunes de Carnaval'],
    [_isoDia_(_masDias_(p, -47)), 'Martes de Carnaval'],
    [_isoDia_(_masDias_(p,  -2)), 'Viernes Santo'],
    [año + '-05-01', 'Día del Trabajador'],
    [año + '-11-02', 'Día de Difuntos'],
    [año + '-11-03', 'Separación de Panamá de Colombia'],
    [año + '-11-04', 'Día de los Símbolos Patrios'],
    [año + '-11-05', 'Día de Colón'],
    [año + '-11-10', 'Primer Grito de Independencia'],
    [año + '-11-28', 'Independencia de España'],
    [año + '-12-08', 'Día de la Madre'],
    [año + '-12-25', 'Navidad']
  ];
}

// Agrega los que falten SIN tocar los existentes: si el admin borró un feriado
// a propósito, volver a sembrarlo cada vez que corre esto anularía su decisión.
// Por eso solo se siembra un año una vez — se marca con la fila `__sembrado__`.
// `rows` viene de quien llama para no releer la hoja. Devuelve true si escribió.
function _sembrarFeriados_(sh, año, rows) {
  const marca = '__sembrado__' + año;
  for (let i = 1; i < rows.length; i++) {
    if ((rows[i][0] || '').toString().trim() === marca) return false;
  }
  const filas = feriadosPanama(año).map(function(f) { return [f[0], f[1]]; });
  filas.push([marca, 'no borrar — evita resembrar ' + año]);
  sh.getRange(sh.getLastRow() + 1, 1, filas.length, 2).setValues(filas);
  Logger.log('✓ Feriados ' + año + ' sembrados (' + (filas.length - 1) + ')');
  return true;
}

// Dos mapas separados —{fecha: nombre}— en vez de uno con el tipo adentro: el
// frontend consulta "¿es feriado?" y "¿es vacaciones escolares?" por separado
// para decidir la tarifa, y un solo mapa lo obligaría a mirar el tipo en cada
// lectura. Ignora las filas centinela.
function getFechasEspeciales() {
  if (_feriadosCache) return _feriadosCache;
  const sh = getOrCreateFeriados();
  const rows = _feriadosRowsCache || sh.getDataRange().getValues();
  const feriados = {}, escolares = {};
  for (let i = 1; i < rows.length; i++) {
    let f = rows[i][0];
    if (!f) continue;
    if (f instanceof Date) f = Utilities.formatDate(f, 'America/Panama', 'yyyy-MM-dd');
    f = f.toString().trim();
    if (f.indexOf('__') === 0) continue;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(f)) continue;
    const tipo = (rows[i][2] || 'feriado').toString().trim().toLowerCase();
    (tipo === 'escolar' ? escolares : feriados)[f] = (rows[i][1] || '').toString();
  }
  _feriadosCache = { feriados: feriados, escolares: escolares };
  return _feriadosCache;
}

// Compat: quedaba usada por getTarifas antes de separar los tipos.
function getFeriadosSet() { return getFechasEspeciales().feriados; }

// ─── BLACKLIST ───────────────────────────────────────────────
function getOrCreateBlacklist() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let blSheet = ss.getSheetByName('Blacklist');
  if (!blSheet) {
    blSheet = ss.insertSheet('Blacklist');
    blSheet.getRange(1, 1).setValue('CodConfirmacion');
    blSheet.getRange(1, 2).setValue('Motivo');
    blSheet.getRange(1, 3).setValue('Fecha');
    blSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    blSheet.setFrozenRows(1);
  }
  return blSheet;
}

function getBlacklistedCodes() {
  const blSheet = getOrCreateBlacklist();
  const data    = blSheet.getDataRange().getValues();
  const codes   = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) codes.add(data[i][0].toString().trim());
  }
  return codes;
}

function addToBlacklist(confirmCode, motivo) {
  if (!confirmCode) return;
  const blSheet = getOrCreateBlacklist();
  const existing = getBlacklistedCodes();
  if (existing.has(confirmCode.toString().trim())) return;
  const fecha = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');
  blSheet.appendRow([confirmCode.toString(), motivo || 'Cancelada', fecha]);
  Logger.log('Blacklist: agregado ' + confirmCode);
}

function normalizeId(id) {
  return id ? id.toString().replace(/^airbnb_/, '') : '';
}

function getProcessedIds(sheet) {
  const data = sheet.getDataRange().getValues();
  const ids  = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      ids.add(data[i][0].toString());
      ids.add(normalizeId(data[i][0]));
    }
    if (data[i][10]) ids.add(data[i][10].toString());
  }
  return ids;
}

function verificarDuplicados() {
  const sheet   = getOrCreateSheet();
  const data    = sheet.getDataRange().getValues();
  const codigos = {};
  let duplicados = 0;
  for (let i = 1; i < data.length; i++) {
    const codigo = data[i][10];
    if (!codigo) continue;
    if (codigos[codigo]) {
      duplicados++;
      Logger.log('DUPLICADO: ' + data[i][1] + ' (' + codigo + ') - filas ' + codigos[codigo] + ' y ' + (i + 1));
    } else {
      codigos[codigo] = i + 1;
    }
  }
  Logger.log('Total filas: ' + (data.length - 1) + ' | Duplicados: ' + duplicados);
}

function eliminarDuplicados() {
  const sheet   = getOrCreateSheet();
  const data    = sheet.getDataRange().getValues();
  const vistos  = new Set();
  const rowsToDelete = [];

  for (let i = 1; i < data.length; i++) {
    const codigo = data[i][10] ? data[i][10].toString() : '';
    const msgId  = data[i][0]  ? data[i][0].toString()  : '';
    const key    = codigo || msgId;
    if (!key) continue;
    if (vistos.has(key)) {
      rowsToDelete.push(i);
      Logger.log('Eliminando duplicado fila ' + (i+1) + ': ' + data[i][1] + ' (' + codigo + ')');
    } else {
      vistos.add(key);
    }
  }

  rowsToDelete.sort((a, b) => b - a);
  rowsToDelete.forEach(i => sheet.deleteRow(i + 1));
  Logger.log('✓ ' + rowsToDelete.length + ' duplicados eliminados. Filas restantes: ' + (data.length - 1 - rowsToDelete.length));
}

function appendReservation(sheet, r) {
  sheet.appendRow([
    r.id, r.name, r.cabinName, r.cabin,
    r.checkin, r.checkout, r.persons,
    r.amount, r.deposit, r.origin, r.confirmCode,
    r.serviceFee || 0,
    r.neto || r.amount,
    r.alerta || '',
    r.name,
    r.bookingDate || Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd'),
    '',
    0,
    r.codTransferencia || '',
    r.montoVoucher     || '',
    r.estadoPago       || '',
    r.email            || '',
    r.comentarios      || '',
    r.telefono         || ''
  ]);
}

// ═══════════════════════════════════════════════════════════
//  Drive Screenshots → Google Sheets
// ═══════════════════════════════════════════════════════════

const DRIVE_FOLDER_NAME    = 'Las Nubes - Reservas Directas';
const FACTURAS_FOLDER_NAME = 'Las Nubes - Facturas';

function getClaudeApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!key) throw new Error('CLAUDE_API_KEY no configurada en Propiedades del script');
  return key;
}

const PROCESSED_SHEET = 'DriveProcessed';

function syncDriveScreenshots() {
  const folder = getDriveFolder();
  if (!folder) {
    Logger.log('⚠ Carpeta "' + DRIVE_FOLDER_NAME + '" no encontrada en Drive.');
    return;
  }

  const sheet        = getOrCreateSheet();
  const processedIds = getProcessedDriveIds();
  const files        = folder.getFiles();
  let added = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileId = file.getId();

    if (processedIds.has(fileId)) continue;
    const mime = file.getMimeType();
    if (!mime.startsWith('image/')) {
      markDriveFileProcessed(fileId, file.getName(), 'SKIPPED - not image');
      continue;
    }

    Logger.log('📷 Procesando: ' + file.getName());

    try {
      const reservation = parseScreenshotWithClaude(file);

      if (reservation) {
        const existing = getProcessedIds(sheet);
        if (!existing.has(reservation.id)) {
          appendReservation(sheet, reservation);
          added++;
          Logger.log('✓ Reserva directa registrada: ' + reservation.name + ' · ' + reservation.cabinName);
        }
        markDriveFileProcessed(fileId, file.getName(), 'OK · ' + reservation.name);
      } else {
        markDriveFileProcessed(fileId, file.getName(), 'ERROR - no se pudo parsear');
        Logger.log('⚠ No se pudo extraer reserva de: ' + file.getName());
      }
    } catch(e) {
      markDriveFileProcessed(fileId, file.getName(), 'ERROR - ' + e.message);
      Logger.log('❌ Error procesando ' + file.getName() + ': ' + e.message);
    }
  }

  Logger.log('Drive sync completado. ' + added + ' reservas directas agregadas.');
}

function parseScreenshotWithClaude(file) {
  const blob    = file.getBlob();
  const base64  = Utilities.base64Encode(blob.getBytes());
  const mime    = file.getMimeType();

  const prompt = `Analiza este screenshot del calendario de Airbnb y extrae la información de la reserva bloqueada.

El screenshot muestra un calendario con:
- Panel inferior izquierdo: "Bloqueadas por ti" con nombre del cliente y nombre de la cabaña
- Panel inferior derecho: "Precio por noche" con el monto
- Un botón negro en el calendario que muestra la fecha o rango de fechas seleccionado

Extrae exactamente:
1. NOMBRE: El nombre completo del cliente (primera y segunda línea del panel "Bloqueadas por ti", ignorar "Cancelado")
2. CABAÑA: El nombre de la cabaña (puede ser "Paseo por Las Nubes", "Portal hacia Las Nubes", o "Puente entre Las Nubes")
3. FECHAS: Del botón negro en el calendario:
   - Si dice una sola fecha como "19 feb" → checkin = esa fecha, checkout = día siguiente
   - Si dice un rango como "3 – 4 de abr" → checkin = primer día, checkout = último día + 1
4. PRECIO: El monto que aparece en "Precio por noche" (es el precio neto por noche)
5. AÑO: Infiere el año basándote en el mes visible en el calendario. El año actual es 2026.

Responde ÚNICAMENTE con un objeto JSON válido, sin texto adicional, sin markdown:
{
  "nombre": "Nombre Completo",
  "cabana": "Nombre exacto de la cabaña",
  "checkin": "YYYY-MM-DD",
  "checkout": "YYYY-MM-DD",
  "noches": 1,
  "precioPorNoche": 90.00,
  "total": 90.00
}`;

  const payload = {
    model: 'claude-opus-4-6',
    max_tokens: 300,
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: mime, data: base64 } },
        { type: 'text', text: prompt }
      ]
    }]
  };

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': getClaudeApiKey(),
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());

  if (!result.content || !result.content[0]) {
    Logger.log('Claude API error: ' + response.getContentText());
    return null;
  }

  const text = result.content[0].text.trim();
  Logger.log('Claude respuesta: ' + text);

  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) return null;

  const data = JSON.parse(jsonMatch[0]);

  if (!data.nombre || !data.checkin || !data.checkout) return null;

  let cabin = 'verde';
  let cabinName = data.cabana || '';
  for (const [cn, code] of Object.entries(CABINS)) {
    if (cabinName.toLowerCase().includes(cn.toLowerCase()) ||
        cn.toLowerCase().includes(cabinName.toLowerCase().split(' ')[0])) {
      cabin = code;
      cabinName = cn;
      break;
    }
  }

  const reservationId = 'drive_' + data.checkin + '_' + data.nombre.replace(/\s+/g, '').toLowerCase().slice(0, 10);

  return {
    id:          reservationId,
    name:        data.nombre,
    cabin:       cabin,
    cabinName:   cabinName,
    checkin:     data.checkin,
    checkout:    data.checkout,
    persons:     2,
    amount:      data.total || data.precioPorNoche,
    deposit:     data.total || data.precioPorNoche,
    origin:      'WhatsApp',
    confirmCode: '',
    serviceFee:  0,
    neto:        data.total || data.precioPorNoche
  };
}

function getDriveFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : null;
}

function getProcessedDriveIds() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(PROCESSED_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PROCESSED_SHEET);
    sheet.getRange(1,1,1,3).setValues([['FileID','Nombre','Estado']]);
    sheet.getRange(1,1,1,3).setFontWeight('bold');
  }
  const data = sheet.getDataRange().getValues();
  const ids  = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) ids.add(data[i][0].toString());
  }
  return ids;
}

function markDriveFileProcessed(fileId, fileName, status) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(PROCESSED_SHEET) || ss.insertSheet(PROCESSED_SHEET);
  sheet.appendRow([fileId, fileName, status, new Date().toISOString()]);
}

// ═══════════════════════════════════════════════════════════
//  Sync Pagos Airbnb
// ═══════════════════════════════════════════════════════════

// Extrae un monto "$1.234,56 USD" de una línea RESPETANDO EL SIGNO.
// Airbnb usa montos negativos para los "Ajustes de la resolución" (reembolsos
// al huésped), ej. "-$99,00 USD". El regex anterior solo capturaba los dígitos,
// así que un ajuste de -$99 se acreditaba como +$99 y el código quedaba
// sobrecontado por el DOBLE del ajuste (ej. HMWC8N44BR: real $68.31 → $266.31,
// descuadrando ese payout en $198). Devuelve null si la línea no trae monto.
function _montoUSDConSigno_(linea) {
  const s = String(linea || '');
  const m = s.match(/(-\s*)?\$\s*(-\s*)?([\d.]+,\d{2})\s*\)?\s*USD/);
  if (!m) return null;
  const v = parseFloat(m[3].replace(/\./g, '').replace(',', '.'));
  if (isNaN(v)) return null;
  // Negativo si hay "-" antes o después del "$", o si viene entre paréntesis.
  const neg = !!m[1] || !!m[2] || /\(\s*-?\s*\$/.test(s);
  return neg ? -v : v;
}

function syncAirbnbPayouts() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  let pagosSheet = ss.getSheetByName('Pagos');
  if (!pagosSheet) {
    pagosSheet = ss.insertSheet('Pagos');
    const headers = ['FechaCobro', 'EmailId', 'MontoTotal', 'ComisionWU', 'MontoNeto', 'ConfirmCodes', 'MontosPorCodigo'];
    pagosSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    pagosSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  } else {
    const lastCol = pagosSheet.getLastColumn();
    if (lastCol < 6) {
      pagosSheet.getRange(1, 6).setValue('ConfirmCodes');
      pagosSheet.getRange(1, 6).setFontWeight('bold');
    }
    if (lastCol < 7) {
      pagosSheet.getRange(1, 7).setValue('MontosPorCodigo');
      pagosSheet.getRange(1, 7).setFontWeight('bold');
    }
  }

  const pagosData    = pagosSheet.getDataRange().getValues().slice(1);
  const processedIds = new Set(pagosData.map(r => r[1].toString()));

  const threads = GmailApp.search(
    'from:automated@airbnb.com subject:"Te hemos enviado un cobro" ' + SYNC_VENTANA
  );

  let added = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const msgId = msg.getId();
      if (processedIds.has(msgId)) {
        Logger.log('~ Ya procesado: ' + msgId);
        return;
      }

      const payDate = Utilities.formatDate(msg.getDate(), 'America/Panama', 'yyyy-MM-dd');
      const body    = msg.getPlainBody();
      const lines   = body.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);

      let montoTotal = 0;
      for (let i = 0; i < lines.length; i++) {
        if (lines[i].indexOf('Total pagado') < 0 && lines[i].indexOf('Total paid') < 0) continue;
        const candidates = [lines[i]];
        if (i + 1 < lines.length) candidates.push(lines[i + 1]);
        for (const c of candidates) {
          const m = c.match(/\$([\d.]+,\d{2})\s*USD/);
          if (m) {
            montoTotal = parseFloat(m[1].replace(/\./g, '').replace(',', '.')) || 0;
            break;
          }
        }
        if (montoTotal > 0) break;
      }

      if (montoTotal === 0) {
        Logger.log('⚠ No se encontró "Total pagado", usando fallback para: ' + msgId);
        for (let i = 0; i < lines.length; i++) {
          const m = lines[i].match(/\$([\d.,]+)\s*USD$/);
          if (!m) continue;
          const v = parseFloat(m[1].replace(/\./g, '').replace(',', '.')) || 0;
          if (v > 0 && v < 2000) montoTotal += v;
        }
        montoTotal = Math.round(montoTotal * 100) / 100;
      }

      if (montoTotal === 0) {
        Logger.log('⚠ No se encontró monto en email: ' + msgId + ' (' + payDate + ')');
        return;
      }

      const comisionWU = 10;
      const montoNeto  = montoTotal;

      const confirmCodes    = [];
      const montosPorCodigo = {};

      for (let i = 0; i < lines.length; i++) {
        const codeMatch = lines[i].match(/^(HM[A-Z0-9]{8})$/);
        if (!codeMatch) continue;
        const code = codeMatch[1];

        let monto = 0;
        for (let j = i - 1; j >= Math.max(0, i - 6); j--) {
          if (lines[j].indexOf('Total pagado') >= 0 || lines[j].indexOf('Total paid') >= 0) continue;
          const v = _montoUSDConSigno_(lines[j]);
          // `!== 0` (no `> 0`): un ajuste negativo es un monto válido y debe
          // acumularse con su signo para que el código quede neteado.
          if (v !== null && v !== 0) { monto = v; break; }
        }

        montosPorCodigo[code] = parseFloat(((montosPorCodigo[code] || 0) + monto).toFixed(2));
        if (!confirmCodes.includes(code)) confirmCodes.push(code);
      }

      const confirmCodesStr    = confirmCodes.join(',');
      const montosPorCodigoStr = confirmCodes
        .map(c => c + ':' + (montosPorCodigo[c] || 0).toFixed(2))
        .join(',');

      const sumaIndividual = Object.values(montosPorCodigo).reduce((s, v) => s + v, 0);
      Logger.log('✓ Email ' + payDate + ' (' + msgId + '): Total $' + montoTotal.toFixed(2) +
        ' | Suma individual $' + sumaIndividual.toFixed(2) +
        ' | ' + confirmCodes.length + ' códigos');

      pagosSheet.appendRow([payDate, msgId, montoTotal, comisionWU, montoNeto, confirmCodesStr, montosPorCodigoStr]);
      processedIds.add(msgId);
      added++;
    });
  });

  Logger.log('syncAirbnbPayouts completado. Nuevos cobros: ' + added);
}

function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('syncAirbnbReservations')
    .timeBased().everyMinutes(15).create();

  ScriptApp.newTrigger('syncAirbnbUpdates')
    .timeBased().everyMinutes(15).create();

  ScriptApp.newTrigger('syncDriveScreenshots')
    .timeBased().everyMinutes(15).create();

  ScriptApp.newTrigger('syncCompleto')
    .timeBased().everyHours(1).create();

  Logger.log('✓ Triggers instalados.');
}

function normalizarFecha(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  const s = val.toString().trim();
  if (!s || s === 'null') return null;
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const d = new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

function limpiarFechasInvalidas() {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const COL_CHECKIN  = 4;
  const COL_CHECKOUT = 5;
  const rowsToDelete = [];

  for (let i = 1; i < data.length; i++) {
    const checkin  = normalizarFecha(data[i][COL_CHECKIN]);
    const checkout = normalizarFecha(data[i][COL_CHECKOUT]);
    let invalid = false, reason = '';
    if (!checkin || !checkout) {
      invalid = true;
      reason  = `fecha no parseable`;
    }
    if (!invalid && checkout <= checkin) {
      invalid = true;
      reason  = `checkout <= checkin`;
    }
    if (invalid) rowsToDelete.push({ rowIndex: i, reason, name: data[i][1], code: data[i][10] });
  }

  rowsToDelete.sort((a, b) => b.rowIndex - a.rowIndex);
  rowsToDelete.forEach(r => sheet.deleteRow(r.rowIndex + 1));
  Logger.log(`✓ ${rowsToDelete.length} filas con fechas inválidas eliminadas.`);
}

// ─── Egresos: helpers de hoja ────────────────────────────────
// Devuelve la hoja Egresos asegurando el header correcto de 9 columnas
// (incluyendo 'Item', añadida para el seguimiento de Suministros). Migra
// hojas viejas de 8 columnas agregando el header 'Item' en la col 9.
function _getEgresoSheetEnsured(ss) {
  const HEADERS = ['ID','Fecha','Descripcion','Monto','Categoria','Cabaña','Proveedor','URLFoto','Item','FechaFin','MontosItem','CantidadesItem'];
  let sheet = ss.getSheetByName('Egresos');
  if (!sheet) {
    sheet = ss.insertSheet('Egresos');
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1,1,1,HEADERS.length).setFontWeight('bold');
    return sheet;
  }
  const firstRow = sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),1)).getValues()[0];
  // Header ausente/corrupto → reescribir completo.
  if (!firstRow[0] || firstRow[0].toString().toLowerCase() !== 'id') {
    sheet.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1,1,1,HEADERS.length).setFontWeight('bold');
    return sheet;
  }
  // Auto-migración: añadir columnas nuevas al final si faltan.
  const lowerHeaders = firstRow.map(h => (h||'').toString().toLowerCase().trim());
  if (lowerHeaders.indexOf('item') === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('Item').setFontWeight('bold');
  }
  // Releer por si acabamos de agregar Item.
  let cur = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(h => (h||'').toString().toLowerCase().trim());
  if (cur.indexOf('fechafin') === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('FechaFin').setFontWeight('bold');
  }
  // Releer por si acabamos de agregar FechaFin.
  cur = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(h => (h||'').toString().toLowerCase().trim());
  if (cur.indexOf('montositem') === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('MontosItem').setFontWeight('bold');
  }
  // Releer por si acabamos de agregar MontosItem.
  cur = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(h => (h||'').toString().toLowerCase().trim());
  if (cur.indexOf('cantidadesitem') === -1) {
    sheet.getRange(1, sheet.getLastColumn() + 1).setValue('CantidadesItem').setFontWeight('bold');
  }
  return sheet;
}

// Índice 1-based de una columna por nombre (case-insensitive) en Egresos (0 si no existe).
function _egresoColIndex(sheet, name) {
  const headers = sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),1)).getValues()[0];
  const target = name.toLowerCase().trim();
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] && headers[i].toString().toLowerCase().trim() === target) return i + 1;
  }
  return 0;
}
// Compat: se usaba _egresoItemColIndex en varios lugares.
function _egresoItemColIndex(sheet) { return _egresoColIndex(sheet, 'item'); }

// ── Dedup de egresos (firma de canasta + códigos de referencia) ──────
// Normaliza una descripción para comparar: minúsculas, sin acentos, sin
// códigos de referencia [XXXX], espacios colapsados.
function _egresoNormDesc(s) {
  return (s || '').toString().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\[[^\]]*\]/g, ' ').replace(/\s+/g, ' ').trim();
}
// Extrae códigos de referencia de una descripción: [WLPTY-41210925], [0000012525].
function _egresoRefCodes(desc) {
  const out = [];
  const re = /\[([A-Za-z0-9]{4,}-[A-Za-z0-9]{4,}|[0-9]{6,})\]/g;
  let m; while ((m = re.exec((desc || '').toString()))) out.push(m[1].toUpperCase());
  return out;
}
// Firma de canasta: multiset ordenado de (descNorm|monto). Ignora fecha,
// proveedor y códigos de referencia (el proveedor suele venir con grafías
// distintas — "PriceSmart Panama" vs "PriceSmart Panama S.A."). Una canasta de
// ≥2 items idénticos en descripción y monto es señal fuerte de la misma factura.
function _egresoBasketSig(items) {
  return items.map(it =>
    _egresoNormDesc(it.desc) + '|' + ((Math.round((parseFloat(it.amount) || 0) * 100) / 100).toFixed(2))
  ).sort().join('##');
}
// Busca si el batch entrante ya existe en la hoja. Devuelve el motivo (string) o
// null. Detecta: (1) un código de referencia ya presente; (2) una canasta
// idéntica (mismo proveedor, mismos items+montos) sin importar la fecha (≥2 items).
function _egresoBatchDuplicate(egresoSheet, items) {
  if (!items || !items.length) return null;
  const data = egresoSheet.getDataRange().getValues();
  const hdr  = data[0].map(h => (h || '').toString().toLowerCase().trim());
  const iId = hdr.indexOf('id'), iDesc = hdr.indexOf('descripcion'), iMonto = hdr.indexOf('monto'), iProv = hdr.indexOf('proveedor');
  if (iId < 0 || iDesc < 0) return null;

  const existingCodes = {};   // codigo -> id
  const groups = {};          // prefijo de id -> [{desc, amount, proveedor}]
  for (let r = 1; r < data.length; r++) {
    const id = (data[r][iId] || '').toString();
    if (!id) continue;
    const desc = (data[r][iDesc] || '').toString();
    _egresoRefCodes(desc).forEach(c => { existingCodes[c] = id; });
    const pfx = id.replace(/_\d+$/, '');
    (groups[pfx] = groups[pfx] || []).push({ desc: desc, amount: data[r][iMonto], proveedor: iProv >= 0 ? data[r][iProv] : '' });
  }

  // 1) Código de referencia ya presente.
  for (let k = 0; k < items.length; k++) {
    const codes = _egresoRefCodes(items[k].desc);
    for (let c = 0; c < codes.length; c++) {
      if (existingCodes[codes[c]]) return 'código de referencia [' + codes[c] + '] ya registrado';
    }
  }
  // 2) Canasta idéntica (solo si ≥2 items, para no bloquear pagos recurrentes de 1 línea).
  if (items.length >= 2) {
    const inSig = _egresoBasketSig(items);
    for (const pfx in groups) {
      if (groups[pfx].length === items.length && _egresoBasketSig(groups[pfx]) === inSig) {
        return 'factura idéntica ya registrada (' + ((items[0].proveedor) || '') + ', ' + items.length + ' items)';
      }
    }
  }
  return null;
}

// Hoja con la lista de keywords a trackear en el seguimiento de Suministros.
// Una columna 'Keyword' — el frontend maneja la lista y la reescribe entera.
function _getSuministrosItemsSheet(ss) {
  let sh = ss.getSheetByName('SuministrosItems');
  if (!sh) {
    sh = ss.insertSheet('SuministrosItems');
    sh.getRange(1,1,1,7).setValues([['Keyword','NoTimeline','Reventa','PrecioVenta','Receta','Venta','CostoUnitario']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // Semilla inicial de keywords sugeridos.
    const seed = ['Papel','Café','Agua','Gas','Hielo','Cloro','Desinfectante','Detergente','Jabón','Kerosene'];
    sh.getRange(2,1,seed.length,1).setValues(seed.map(k => [k]));
    return sh;
  }
  // Auto-migración: asegurar la columna 2 NoTimeline (keywords que NO van en el
  // timeline, ej. Delivery — no es un insumo consumible, solo se ve en Detalle).
  const hdr = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),7)).getValues()[0];
  if ((hdr[1]||'').toString().toLowerCase().trim() !== 'notimeline') {
    sh.getRange(1,2).setValue('NoTimeline').setFontWeight('bold');
  }
  // Columna 3: keywords cuyos insumos se REVENDEN en la tiendita. Marca de
  // qué compras sale el costo de la mercadería vendida.
  if ((hdr[2]||'').toString().toLowerCase().trim() !== 'reventa') {
    sh.getRange(1,3).setValue('Reventa').setFontWeight('bold');
  }
  // Columna 4: precio al que se VENDE el item en la tiendita. El costo sale de
  // los egresos; el precio de venta no existía en ningún lado y había que
  // teclearlo en cada venta.
  if ((hdr[3]||'').toString().toLowerCase().trim() !== 'precioventa') {
    sh.getRange(1,4).setValue('PrecioVenta').setFontWeight('bold');
  }
  // Columna 5: RECETA — JSON { keyword_componente_normalizada: cantidad }.
  // Para lo que se ARMA en vez de comprarse hecho: un "Kit de Fogata" nunca va a
  // aparecer en una factura, su costo son sus componentes (leña + malvaviscos +
  // cerillos). Vacío = el item se compra tal cual se vende y su costo sale
  // directo de los egresos que matchean su keyword.
  if ((hdr[4]||'').toString().toLowerCase().trim() !== 'receta') {
    sh.getRange(1,5).setValue('Receta').setFontWeight('bold');
  }
  // Columna 6: VENTA — es un PRODUCTO de la tiendita (lo que se le cobra al
  // huésped), no un insumo. Se marca sola al importar de la web. Es el otro
  // lado de Reventa(🛒), que ahora significa "insumo que puede entrar en un
  // combo". Son roles distintos: el carbón a granel es insumo, la "Bolsa de
  // Carbón" de $5 es producto. Mezclarlos hacía que el desplegable de
  // componentes ofreciera kits y que la lista de venta mostrara insumos.
  if ((hdr[5]||'').toString().toLowerCase().trim() !== 'venta') {
    sh.getRange(1,6).setValue('Venta').setFontWeight('bold');
  }
  // Columna 7: COSTO UNITARIO FIJO, puesto a mano. Manda sobre el derivado de
  // las compras. Existe para lo que no llega como factura con su keyword: la
  // mano de obra de recolectar la leña es un costo real del kit, pero no hay
  // ningún egreso que diga "leña" y del que se pueda deducir un precio por
  // unidad. Vacío = se deriva de las compras, como siempre.
  if ((hdr[6]||'').toString().toLowerCase().trim() !== 'costounitario') {
    sh.getRange(1,7).setValue('CostoUnitario').setFontWeight('bold');
  }
  return sh;
}

// ═══════════════════════════════════════════════════════════
//  doGet — endpoint principal
// ═══════════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════════
//  Control de acceso del Web App
// ═══════════════════════════════════════════════════════════
// El despliegue es ANYONE_ANONYMOUS y corre como el dueño, así que cualquiera
// con la URL podía leer todas las reservas —con emails, teléfonos y montos—, la
// contabilidad y los chats del bot, y además crear o borrar reservas. Y la URL
// está escrita en index.html, que es público: basta con ver el código fuente
// del calendario.
//
// Ahora todo exige la clave ADMIN_KEY (Script Properties) salvo lo que abajo
// está en la lista blanca, que son las acciones de las páginas públicas y las
// que pollea Airbnb.
const ACCIONES_PUBLICAS = {
  getTarifas:        true,   // calendario público: precios
  getReservaPublic:  true,   // página del huésped — valida su propio short code
  uploadHuespedId:   true,   // el huésped sube su cédula desde esa página
  getMalayaCalendar: true,   // calendario público de Malaya
  getIcal:           true,   // lo pollea Airbnb
  malayaIcal:        true    // lo pollea Airbnb
};

function _accesoPermitido(action, params, payload) {
  const esperada = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');

  // Sin ADMIN_KEY configurada no se bloquea nada. Desplegar esto no puede dejar
  // al admin afuera de su propio dashboard: la protección se enciende sola en el
  // momento en que se crea la propiedad.
  if (!esperada) {
    if (action && !ACCIONES_PUBLICAS[action]) {
      logDebugEntry('acceso-SIN-CLAVE', { action: action, nota: 'ADMIN_KEY no configurada — todo abierto' });
    }
    return true;
  }

  if (ACCIONES_PUBLICAS[action]) return true;

  // El calendario público pide las reservas con scope=public, que solo devuelve
  // fechas, cabaña y tipo. Sin ese scope la MISMA acción devuelve emails,
  // teléfonos y montos, así que ahí sí hace falta la clave.
  if ((!action || action === 'getReservations') && params && params.scope === 'public') return true;

  const dada = (params && params.k) || (payload && payload.k) || '';
  return String(dada) === String(esperada);
}

function _denegado(action) {
  logDebugEntry('acceso-DENEGADO', { action: action || '(default)' });
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: 'UNAUTHORIZED' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  // ── WHATSAPP WEBHOOK VERIFY ─────────────────────────────────
  // Meta hace GET con hub.mode=subscribe&hub.verify_token=...&hub.challenge=...
  if (e && e.parameter && e.parameter['hub.mode']) {
    return handleWebhookVerify(e);
  }

  if (!_accesoPermitido(action, e && e.parameter, null)) return _denegado(action);

  try {
    // ── SAVE TARIFAS (via GET params para evitar redirect 302) ──
    if (action === 'saveTarifas') {
      const p   = e.parameter;
      const cfg = getOrCreateConfig();
      const rows = cfg.getDataRange().getValues();
      const hoy  = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');
      // Solo se escriben las claves que vinieron en la request: mandar un
      // default por cada una ausente pisaría tarifas que el admin no tocó.
      // promoActive solo si vino en la request. Antes se escribía SIEMPRE, así que
      // un guardado parcial —por ejemplo el de las tarifas de Malaya, que no
      // manda ese campo— apagaba la promoción de Las Nubes sin que nadie lo
      // pidiera.
      const updates = {};
      if (p.promoActive !== undefined) {
        updates.promoActive = (p.promoActive === 'true' || p.promoActive === '1') ? 'true' : 'false';
      }
      ['weekday','weekend','promo','viernes','sabado','vispera','feriado','escolar',
       'promoViernes','promoSabado','promoVispera','promoFeriado','promoEscolar',
       'promoMeses',
       'pasatarde','pasanoche','pasadia',
       'recargoPasatardePersona','recargoPasanochePersona','recargoPasadiaPersona',
       'recargoPersonaGrande','recargoPersonaPortal','comboDescuento',
       'malayaSemana','malayaFinde','malayaPersonaExtra',
       'recargoEarly','recargoLate'].forEach(function(k) {
        if (p[k] !== undefined && p[k] !== '' && !isNaN(parseFloat(p[k]))) updates[k] = parseFloat(p[k]);
      });
      for (let i = 1; i < rows.length; i++) {
        const clave = rows[i][0] ? rows[i][0].toString().trim() : '';
        if (updates.hasOwnProperty(clave)) {
          cfg.getRange(i + 1, 2).setValue(updates[clave]);
          cfg.getRange(i + 1, 4).setValue(hoy);
          cfg.getRange(i + 1, 5).setValue('admin');
        }
      }
      _invalidarConfigCache();
      Logger.log('✓ Tarifas guardadas via GET: ' + JSON.stringify(updates));
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── FERIADOS (alta/baja desde el modal de Tarifas) ────────
    // Se manda el DIFF, no la lista entera: los nombres son largos y treinta
    // feriados en la URL rozan el límite. Con add/del la request queda corta
    // sin importar cuántos haya cargados.
    //   add=2026-07-04|Feria de Chame;2026-07-05|Otro
    //   del=2026-11-02,2026-11-05
    if (action === 'saveFeriados') {
      const sh = getOrCreateFeriados();
      const rows = sh.getDataRange().getValues();
      const p = e.parameter;

      const borrar = {};
      (p.del || '').split(',').forEach(function(f) { f = f.trim(); if (f) borrar[f] = true; });

      let quitados = 0;
      // De abajo hacia arriba: borrar hacia adelante corre las filas y se
      // saltearía la siguiente.
      for (let i = rows.length - 1; i >= 1; i--) {
        let f = rows[i][0];
        if (f instanceof Date) f = Utilities.formatDate(f, 'America/Panama', 'yyyy-MM-dd');
        f = (f || '').toString().trim();
        if (borrar[f]) { sh.deleteRow(i + 1); quitados++; }
      }

      // La fila centinela __sembrado__<año> sobrevive a los borrados, así que
      // un feriado que el admin quita no vuelve en la próxima siembra.
      const nuevas = [];
      (p.add || '').split(';').forEach(function(item) {
        const parte = item.split('|');
        const fecha = (parte[0] || '').trim();
        if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) return;
        // fecha|nombre|tipo — el tipo es opcional y cae a feriado.
        const tipo = (parte[2] || 'feriado').trim().toLowerCase() === 'escolar' ? 'escolar' : 'feriado';
        nuevas.push([fecha, (parte[1] || 'Feriado').trim(), tipo]);
      });
      if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, 3).setValues(nuevas);

      _feriadosCache = null; _feriadosRowsCache = null;
      Logger.log('✓ Fechas especiales: +' + nuevas.length + ' / -' + quitados);
      const _fe2 = getFechasEspeciales();
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, agregados: nuevas.length, quitados: quitados,
                                           feriados: _fe2.feriados, escolares: _fe2.escolares }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── MALAYA iCAL FEED (para importar en Airbnb) ────────────
    // Devuelve un .ics con todas las reservas directas activas de Malaya.
    // Airbnb polla esta URL cada pocas horas y bloquea esas fechas para
    // evitar double-booking. La URL es pública (necesita serlo para que
    // Airbnb la lea sin auth) — no expone datos del huésped, sólo rango
    // de fechas y un summary genérico.
    if (action === 'malayaIcal') {
      return getMalayaIcalFeed();
    }

    // ── GET TARIFAS ───────────────────────────────────────────
    if (action === 'getTarifas') {
      const cfg  = getOrCreateConfig();
      const rows = cfg.getDataRange().getValues().slice(1);
      const map  = {};
      rows.forEach(r => {
        const key = r[0] ? r[0].toString().trim() : '';
        if (key) map[key] = r[1];
      });
      const _fe = getFechasEspeciales();
      const tarifas = {
        weekday:     parseFloat(map['weekday'])    || 90,
        weekend:     parseFloat(map['weekend'])    || 110,
        promo:       parseFloat(map['promo'])      || 75,
        promoActive: map['promoActive'] === true || map['promoActive'] === 'true' || map['promoActive'] === 1,
        // Tarifas planas para ventanas cortas (2 personas base).
        pasatarde:          parseFloat(map['pasatarde'])            || 60,
        pasanoche:          parseFloat(map['pasanoche'])            || 75,
        pasadia:            parseFloat(map['pasadia'])              || 75,
        // Recargos por persona extra (desde la 3ra).
        recargoPersonaGrande:   parseFloat(map['recargoPersonaGrande'])   || 20,
        recargoPersonaPortal:   parseFloat(map['recargoPersonaPortal'])   || 10,
        comboDescuento:         parseFloat(map['comboDescuento'])         || 0,
        recargoEarly:           parseFloat(map['recargoEarly'])           || 0,
        recargoLate:            parseFloat(map['recargoLate'])            || 0,
        recargoPasatardePersona: parseFloat(map['recargoPasatardePersona']) || 20,
        recargoPasanochePersona: parseFloat(map['recargoPasanochePersona']) || 25,
        recargoPasadiaPersona:   parseFloat(map['recargoPasadiaPersona'])   || 25,
        // Malaya: el dashboard las usa para el bloque de tarifas y para cotizar.
        malayaSemana:       parseFloat(map['malayaSemana'])       || 75,
        malayaFinde:        parseFloat(map['malayaFinde'])        || 100,
        malayaPersonaExtra: parseFloat(map['malayaPersonaExtra']) || 30,
        // Tarifas por tipo de día. `weekend` queda de fallback para instalaciones
        // que todavía no corrieron la migración.
        viernes:      parseFloat(map['viernes'])      || parseFloat(map['weekend']) || 110,
        sabado:       parseFloat(map['sabado'])       || parseFloat(map['weekend']) || 110,
        vispera:      parseFloat(map['vispera'])      || parseFloat(map['weekend']) || 110,
        feriado:      parseFloat(map['feriado'])      || parseFloat(map['weekend']) || 110,
        // Los promocionales admiten 0 = "sin promo para este tipo de día", así
        // que NO llevan `||` con un default: eso convertiría el 0 en un precio.
        promoViernes: parseFloat(map['promoViernes']) || 0,
        promoSabado:  parseFloat(map['promoSabado'])  || 0,
        promoVispera: parseFloat(map['promoVispera']) || 0,
        promoFeriado: parseFloat(map['promoFeriado']) || 0,
        escolar:      parseFloat(map['escolar'])      || parseFloat(map['weekend']) || 110,
        promoEscolar: parseFloat(map['promoEscolar']) || 0,
        promoMeses:   parseInt(map['promoMeses'], 10)  || 0,
        feriados:     _fe.feriados,
        escolares:    _fe.escolares,
        // Interruptor global de las ventanas cortas en el calendario público
        // (pasatarde/pasanoche en los huecos que dejan early/late/pasadía).
        // Default ENCENDIDO: solo se apaga poniendo 'false' en la hoja Config,
        // así una fila ausente no desactiva la función en silencio.
        ventanasPublicas: !(map['ventanasPublicas'] === false ||
                            String(map['ventanasPublicas']).toLowerCase() === 'false' ||
                            map['ventanasPublicas'] === 0)
      };
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, tarifas }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── GET EGRESOS ───────────────────────────────────────────
    if (action === 'getEgresos') {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const egresoSheet = ss.getSheetByName('Egresos');
      if (!egresoSheet) {
        return ContentService
          .createTextOutput(JSON.stringify({ ok: true, egresos: [] }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const eData = egresoSheet.getDataRange().getValues();
      const headers = eData[0].map(h => h.toString().toLowerCase().trim());
      const idx = {
        id:        headers.indexOf('id'),
        fecha:     headers.indexOf('fecha'),
        desc:      headers.indexOf('descripcion'),
        monto:     headers.indexOf('monto'),
        cat:       headers.indexOf('categoria'),
        cabin:     headers.findIndex(h => h.includes('caba')),
        proveedor: headers.indexOf('proveedor'),
        urlFoto:   headers.findIndex(h => h.includes('url') || h.includes('foto')),
        item:      headers.indexOf('item'),
        fechaFin:  headers.indexOf('fechafin'),
        montosItem: headers.indexOf('montositem'),
        cantidadesItem: headers.indexOf('cantidadesitem')
      };
      const egresos = eData.slice(1)
        .filter(r => r[idx.fecha] && r[idx.monto])
        .map(r => {
          const fecha = r[idx.fecha] instanceof Date
            ? Utilities.formatDate(r[idx.fecha], 'America/Panama', 'yyyy-MM-dd')
            : r[idx.fecha].toString().slice(0,10);
          const id = idx.id >= 0 && r[idx.id] ? r[idx.id].toString() : ('sheet_egr_' + fecha + '_' + Math.random().toString(36).slice(2,6));
          return {
            id,
            date:      fecha,
            desc:      idx.desc      >= 0 ? (r[idx.desc]      || '') : '',
            amount:    parseFloat(r[idx.monto]) || 0,
            cat:       idx.cat       >= 0 ? (r[idx.cat]       || 'Otro') : 'Otro',
            cabin:     idx.cabin     >= 0 ? (r[idx.cabin]      || '') : '',
            proveedor: idx.proveedor >= 0 ? (r[idx.proveedor]  || '') : '',
            urlFoto:   idx.urlFoto   >= 0 ? (r[idx.urlFoto]    || '') : '',
            item:      idx.item      >= 0 ? (r[idx.item]       || '') : '',
            fechaFin:  idx.fechaFin  >= 0 ? (r[idx.fechaFin] instanceof Date
                          ? Utilities.formatDate(r[idx.fechaFin], 'America/Panama', 'yyyy-MM-dd')
                          : (r[idx.fechaFin] || '').toString().slice(0,10)) : '',
            montosItem: idx.montosItem >= 0 ? (r[idx.montosItem] || '').toString() : '',
            cantidadesItem: idx.cantidadesItem >= 0 ? (r[idx.cantidadesItem] || '').toString() : '',
            fromSheets: true
          };
        });
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, egresos }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── GET SUMINISTROS KEYWORDS (lista de items a trackear) ──
    if (action === 'getSuministrosItems') {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sh = _getSuministrosItemsSheet(ss);
      const rows = sh.getLastRow() > 1 ? sh.getRange(2,1,sh.getLastRow()-1,7).getValues() : [];
      const keywords = [];
      const noTimeline = [];
      const reventa = [];
      const venta = [];
      const precios = {};
      const recetas = {};
      const costos = {};
      const esFlag = f => (f === true || f === 'TRUE' || f === 'true' || f === 1 || f === '1');
      rows.forEach(r => {
        const kw = (r[0]||'').toString().trim();
        if (!kw) return;
        keywords.push(kw);
        if (esFlag(r[1])) noTimeline.push(kw);
        if (esFlag(r[2])) reventa.push(kw);
        if (esFlag(r[5])) venta.push(kw);
        const pv = parseFloat(r[3]);
        if (!isNaN(pv) && pv > 0) precios[kw] = pv;
        const cu = parseFloat(r[6]);
        if (!isNaN(cu) && cu > 0) costos[kw] = cu;
        // Receta corrupta = item sin receta, no un error que tumbe la carga
        // entera de keywords (sin ellas el timeline queda vacío).
        const rec = (r[4]||'').toString().trim();
        if (rec) {
          try { const o = JSON.parse(rec); if (o && typeof o === 'object' && Object.keys(o).length) recetas[kw] = o; }
          catch(_) {}
        }
      });
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, keywords, noTimeline, reventa, venta, precios, recetas, costos }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── GET PAGOS ─────────────────────────────────────────────
    if (action === 'getPagos') {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const pagosSheet = ss.getSheetByName('Pagos');
      if (!pagosSheet) {
        return ContentService
          .createTextOutput(JSON.stringify({ ok: true, pagos: [] }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const pData = pagosSheet.getDataRange().getValues().slice(1);
      const pagos = pData
        .filter(r => r[0] && r[1])
        .map(r => ({
          fechaCobro:      r[0] instanceof Date ? Utilities.formatDate(r[0], 'America/Panama', 'yyyy-MM-dd') : r[0].toString().slice(0,10),
          emailId:         r[1].toString(),
          montoTotal:      parseFloat(r[2]) || 0,
          comisionWU:      parseFloat(r[3]) || 0,
          montoNeto:       parseFloat(r[4]) || 0,
          confirmCodes:    r[5] ? r[5].toString().split(',').map(c => c.trim()).filter(Boolean) : [],
          montosPorCodigo: r[6] ? (function(s) {
            const map = {};
            s.toString().split(',').forEach(function(pair) {
              const parts = pair.trim().split(':');
              if (parts.length === 2) map[parts[0].trim()] = parseFloat(parts[1]) || 0;
            });
            return map;
          })(r[6]) : {}
        }));
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, pagos }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── GET ICAL ──────────────────────────────────────────────
    if (action === 'getIcal') {
      const cabin = e.parameter.cabin || '';
      return getIcalForCabin(cabin);
    }

    // ── PUBLIC LINK ───────────────────────────────────────────
    if (action === 'getReservaPublic') return handleGetReservaPublic(e);
    if (action === 'getReservaLink')   return handleGetReservaLink(e);
    if (action === 'getLoyaltyCredits') return handleGetLoyaltyCredits(e);
    if (action === 'getReferrals')      return handleGetReferrals(e);
    if (action === 'debugFindReserva')  return handleDebugFindReserva(e);
    if (action === 'getDebugLog')       return handleGetDebugLog(e);
    if (action === 'getConversaciones') return handleGetConversaciones(e);
    if (action === 'getMensajes')       return handleGetMensajes(e);
    if (action === 'getTienditaVentas')
      return ContentService
        .createTextOutput(JSON.stringify(Object.assign({ ok: true }, getTienditaVentas())))
        .setMimeType(ContentService.MimeType.JSON);
    if (action === 'getPrestamos')
      return ContentService
        .createTextOutput(JSON.stringify(Object.assign({ ok: true }, getPrestamosData())))
        .setMimeType(ContentService.MimeType.JSON);
    if (action === 'getMalayaCalendar') {
      // ?admin=1 incluye datos de las reservas directas (id, huésped, teléfono,
      // etc.) para poblar el modal de edición. Sin admin=1, solo devuelve
      // rangos ocupados anónimos (para el calendario público).
      const isAdmin = e && e.parameter && (e.parameter.admin === '1' || e.parameter.admin === 'true');
      const data = getMalayaCalendarData(isAdmin);
      return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
    }
    // No va en ACCIONES_PUBLICAS: son mis números, exige clave.
    if (action === 'getMalayaGanancias')
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, ganancias: getMalayaGanancias() }))
        .setMimeType(ContentService.MimeType.JSON);
    if (action === 'getBotAlertConfig')
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, config: _botGetAlertConfig() }))
        .setMimeType(ContentService.MimeType.JSON);

    // ── GET RESERVATIONS (default) ────────────────────────────
    const sheet = getOrCreateSheet();
    const data  = sheet.getDataRange().getValues();
    const rows  = data.slice(1);
    const scopePublic = e && e.parameter && e.parameter.scope === 'public';

    // Modo publico: solo campos necesarios para mostrar disponibilidad en el
    // calendario. Evita exponer datos sensibles (email, telefono, montos, etc.)
    // y reduce el payload ~80%.
    if (scopePublic) {
      const reservations = rows
        .filter(r => r[0])
        .map(r => ({
          cabin:      r[3],
          checkin:    r[4] instanceof Date ? Utilities.formatDate(r[4], 'America/Panama', 'yyyy-MM-dd') : r[4],
          checkout:   r[5] instanceof Date ? Utilities.formatDate(r[5], 'America/Panama', 'yyyy-MM-dd') : r[5],
          origin:     r[9],
          estadoPago: r[20] || '',
          tipo:       r[24] || 'noche',
          // Horas reales de la reserva. El calendario público las necesita para
          // calcular qué ventana corta cabe en el hueco que deja una reserva
          // que rompe el patrón 2pm–11am; sin ellas caería al default del tipo
          // e ignoraría los overrides. No son datos sensibles.
          checkoutExtendido: !!r[28],
          horaEntrada:       _normalizeHora(r[29]),
          horaSalida:        _normalizeHora(r[30])
        }));
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, reservations }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const reservations = rows
      .filter(r => r[0])
      .map(r => ({
        id:          r[0],
        name:        r[1],
        cabinName:   r[2],
        cabin:       r[3],
        checkin:     r[4] instanceof Date ? Utilities.formatDate(r[4], 'America/Panama', 'yyyy-MM-dd') : r[4],
        checkout:    r[5] instanceof Date ? Utilities.formatDate(r[5], 'America/Panama', 'yyyy-MM-dd') : r[5],
        persons:     r[6],
        amount:      r[7],
        deposit:     r[8] || 0,
        origin:      r[9],
        confirmCode: r[10],
        serviceFee:  r[11] || 0,
        neto:        r[12] || r[7],
        alerta:           r[13] || '',
        pagador:          r[14] || '',
        fechaReserva:     r[15] || '',
        fechaPago:        r[16] instanceof Date ? Utilities.formatDate(r[16], 'America/Panama', 'yyyy-MM-dd') : (r[16] || ''),
        montoPagado:      r[17] || 0,
        codTransferencia: r[18] || '',
        montoVoucher:     r[19] || '',
        estadoPago:       r[20] || '',
        email:            r[21] || '',
        comentarios:      r[22] || '',
        telefono:         r[23] || '',
        tipo:             r[24] || 'noche',
        voucherURL:       r[25] || '',
        idHuespedURL:     r[26] || '',
        fechaNacimiento:  r[27] instanceof Date ? Utilities.formatDate(r[27], 'America/Panama', 'yyyy-MM-dd') : (r[27] || ''),
        checkoutExtendido: r[28] === true || r[28] === 'TRUE' || r[28] === 'true' || r[28] === 1,
        horaEntrada:      _normalizeHora(r[29]),
        horaSalida:       _normalizeHora(r[30]),
        // JSON array con un entry por voucher subido: {monto, cod, fecha, url}.
        // Vacío en reservas viejas — el frontend degrada al par
        // VoucherURL / CodTransferencia.
        vouchersMeta:     r[31] || '',
        // JSON {de, mensaje} cuando la reserva es un certificado de regalo.
        // Vacío = reserva normal. Ver _parseRegalo().
        regalo:           r[32] || ''
      }));

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, reservations }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ═══════════════════════════════════════════════════════════
//  doPost — endpoint principal
// ═══════════════════════════════════════════════════════════
function _testSaveReserva() {
  const testId = 'TEST_' + Date.now();
  const fakePayload = {
    postData: {
      contents: JSON.stringify({
        action: 'saveReservation',
        reservation: {
          id: testId,
          name: 'TEST DESDE EDITOR',
          email: 'test@example.com',
          telefono: '6000-0000',
          pagador: 'TEST',
          cabin: 'verde',
          checkin: '2026-12-01',
          checkout: '2026-12-02',
          persons: 2,
          origin: 'Directa',
          deposit: 0,
          amount: 90,
          confirmCode: testId,
          serviceFee: 0,
          neto: 90,
          fechaReserva: '2026-05-18',
          codTransferencia: '',
          montoVoucher: '',
          estadoPago: 'PENDIENTE',
          comentarios: 'test desde editor',
          tipo: 'noche'
        }
      })
    }
  };
  Logger.log('=== _testSaveReserva START · testId=' + testId + ' ===');
  const out = doPost(fakePayload);
  Logger.log('=== Response: ' + out.getContent() + ' ===');

  // Verificar inmediatamente si quedo en la hoja
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  let found = false;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] && data[i][0].toString() === testId) {
      found = true;
      Logger.log('  ✓ Row encontrado en fila ' + (i + 1) + ': ' + JSON.stringify(data[i].slice(0, 5)));
      break;
    }
  }
  if (!found) Logger.log('  ✗ Row NO se encontro en la hoja despues del save');
  Logger.log('=== _testSaveReserva END ===');
}

function doPost(e) {
  const _ts0 = Date.now();
  let _action = '?', _id = '?', _payload = null;
  try {
    const contentLen  = e && e.postData ? (e.postData.contents || '').length : 0;
    const contentType = e && e.postData ? (e.postData.type || '?') : 'no-postData';
    _payload = JSON.parse(e.postData.contents);
    const payload = _payload;

    // ── WHATSAPP WEBHOOK INBOUND ─────────────────────────────
    // Meta envia POST con payload { object: 'whatsapp_business_account', entry: [...] }
    if (payload && payload.object === 'whatsapp_business_account') {
      logDebugEntry('doPost-IN', { action: 'wa-webhook', contentLen, contentType });
      return handleWhatsAppWebhook(payload);
    }

    _action = payload.action;
    _id     = payload.reservation ? payload.reservation.id : (payload.id || '');
    logDebugEntry('doPost-IN', { action: _action, id: _id, contentLen, contentType });
    const action  = payload.action;

    if (!_accesoPermitido(action, e && e.parameter, payload)) return _denegado(action);

    // ── SAVE TARIFAS ─────────────────────────────────────────
    if (action === 'saveTarifas') {
      const t   = payload.tarifas;
      const cfg = getOrCreateConfig();
      const rows = cfg.getDataRange().getValues();
      const hoy  = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');

      const updates = {
        weekday:     parseFloat(t.weekday)    || 90,
        weekend:     parseFloat(t.weekend)    || 110,
        promo:       parseFloat(t.promo)      || 75,
        promoActive: t.promoActive ? 'true' : 'false',
        // Tipos de día. Los promocionales pueden ser 0 (= sin promo), así que
        // se leen con `|| 0` y no con un default de precio.
        viernes:      parseFloat(t.viernes)      || 110,
        sabado:       parseFloat(t.sabado)       || 110,
        vispera:      parseFloat(t.vispera)      || 110,
        feriado:      parseFloat(t.feriado)      || 110,
        promoViernes: parseFloat(t.promoViernes) || 0,
        promoSabado:  parseFloat(t.promoSabado)  || 0,
        promoVispera: parseFloat(t.promoVispera) || 0,
        promoFeriado: parseFloat(t.promoFeriado) || 0,
        escolar:      parseFloat(t.escolar)      || 110,
        promoEscolar: parseFloat(t.promoEscolar) || 0,
        promoMeses:   parseInt(t.promoMeses, 10) || 0
      };

      for (let i = 1; i < rows.length; i++) {
        const clave = rows[i][0] ? rows[i][0].toString().trim() : '';
        if (updates.hasOwnProperty(clave)) {
          cfg.getRange(i + 1, 2).setValue(updates[clave]);
          cfg.getRange(i + 1, 4).setValue(hoy);
          cfg.getRange(i + 1, 5).setValue('admin');
        }
      }
      _invalidarConfigCache();
      Logger.log('✓ Tarifas guardadas: ' + JSON.stringify(updates));
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── SAVE RESERVATION ─────────────────────────────────────
    if (action === 'saveReservation') {
      const r     = payload.reservation;
      Logger.log('▶ saveReservation entry · id=' + (r && r.id) + ' name=' + (r && r.name) + ' confirmCode=' + (r && r.confirmCode));
      const sheet = getOrCreateSheet();
      Logger.log('  sheet name=' + sheet.getName() + ' lastRow=' + sheet.getLastRow() + ' lastCol=' + sheet.getLastColumn());

      const existing = getProcessedIds(sheet);
      const key = r.confirmCode || r.id;
      Logger.log('  dedup key=' + key + ' existing.size=' + existing.size + ' has?=' + existing.has(String(key)));
      if (key && existing.has(key.toString())) {
        Logger.log('  → DUPLICATE skip (id/confirmCode match)');
        logDebugEntry('saveReservation-DUPE', { id: r.id, key: key });
        return ContentService
          .createTextOutput(JSON.stringify({ ok: true, status: 'duplicate' }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      // Dedup semántico: el front genera un id nuevo con Date.now() en cada
      // submit; si el form se reabre y se vuelve a guardar, el id es distinto
      // y el check por id no atrapa el duplicado. Acá buscamos en TODA la
      // hoja si ya existe una reserva con mismo cliente (name normalizado) +
      // cabaña + checkin + checkout. No es esperable que un mismo cliente
      // reserve dos veces la misma cabaña en los mismos días.
      try {
        const data = sheet.getDataRange().getValues();
        const normName = String(r.name || '').toLowerCase().replace(/\s+/g, ' ').trim();
        const ciNew = String(r.checkin || '').slice(0, 10);
        const coNew = String(r.checkout || '').slice(0, 10);
        const cabinKey = (r.cabin || '').toString();
        if (normName && ciNew && coNew && cabinKey) {
          for (let i = 1; i < data.length; i++) {
            const rr = data[i];
            if (!rr[0]) continue;
            // Skip canceladas — esa fila quedó "muerta", una nueva con misma
            // info es legítima.
            if ((rr[20] || '').toString().toUpperCase() === 'CANCELADA') continue;
            const rowName = String(rr[1] || '').toLowerCase().replace(/\s+/g, ' ').trim();
            if (rowName !== normName) continue;
            if (String(rr[3] || '') !== cabinKey) continue;
            const rowCi = rr[4] instanceof Date ? Utilities.formatDate(rr[4], 'America/Panama', 'yyyy-MM-dd') : String(rr[4] || '').slice(0,10);
            const rowCo = rr[5] instanceof Date ? Utilities.formatDate(rr[5], 'America/Panama', 'yyyy-MM-dd') : String(rr[5] || '').slice(0,10);
            if (rowCi !== ciNew || rowCo !== coNew) continue;
            Logger.log('  → DUPLICATE skip (semántico: name+cabin+fechas)');
            logDebugEntry('saveReservation-DUPE-SEM', { id: r.id, matchedId: rr[0], name: r.name, cabin: cabinKey });
            return ContentService
              .createTextOutput(JSON.stringify({ ok: true, status: 'duplicate', matchedId: rr[0] }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      } catch(dupErr) {
        Logger.log('  dedup semántico fallo: ' + dupErr.message);
      }

      const CABIN_NAMES = {
        verde: 'Paseo por Las Nubes',
        azul:  'Portal hacia Las Nubes',
        lila:  'Puente entre Las Nubes'
      };

      const amount_    = parseFloat(r.amount)    || 0;
      const deposit_   = parseFloat(r.deposit)   || 0;
      const serviceFee_= parseFloat(r.serviceFee)|| 0;
      const neto_      = parseFloat(r.neto)       || amount_;
      const persons_   = parseInt(r.persons)      || 1;
      const today_     = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');

      const rowToAppend = [
        r.id,
        _safeCell(r.name),
        CABIN_NAMES[r.cabin] || r.cabin,
        r.cabin,
        r.checkin,
        r.checkout,
        persons_,
        amount_,
        deposit_,
        r.origin    || 'Directa',
        _sanitizeConfirmCode(r.confirmCode) || '',
        serviceFee_,
        neto_,
        '',
        _safeCell(r.pagador || r.name),
        r.fechaReserva || today_,
        '',
        0,
        _safeCell(r.codTransferencia || ''),
        r.montoVoucher || '',
        r.estadoPago   || '',
        _safeCell(r.email || ''),
        _safeCell(r.comentarios || ''),
        _safeCell(r.telefono || ''),
        r.tipo         || 'noche'
      ];
      Logger.log('  appendRow length=' + rowToAppend.length + ' id=' + rowToAppend[0] + ' name=' + rowToAppend[1]);
      try {
        sheet.appendRow(rowToAppend);
        Logger.log('  ✓ appendRow OK · newLastRow=' + sheet.getLastRow());
        logDebugEntry('saveReservation-OK', { id: r.id, row: sheet.getLastRow(), name: r.name, cabin: r.cabin, checkin: r.checkin });
        // Col 29 (CheckoutExtendido) — persistir el flag de cortesia si aplica
        if (r.checkoutExtendido) {
          sheet.getRange(sheet.getLastRow(), 29).setValue(true);
        }
        // Col 26 (VoucherURL) — persistir si viene en el payload. Usado por el
        // flujo multi-cabaña donde el voucher se sube una vez y su URL se
        // comparte entre las reservas hermanas.
        if (r.voucherURL) {
          sheet.getRange(sheet.getLastRow(), 26).setValue(r.voucherURL);
        }
        // Cols 30/31 (HoraEntrada/HoraSalida) — horas custom que overridean
        // los defaults del tipo. Vacío = usa el horario estándar del tipo.
        const hi = _normalizeHora(r.horaEntrada);
        const ho = _normalizeHora(r.horaSalida);
        if (hi) sheet.getRange(sheet.getLastRow(), 30).setValue(hi);
        if (ho) sheet.getRange(sheet.getLastRow(), 31).setValue(ho);
        // Col 33 (Regalo) — JSON {de, mensaje} si es certificado de regalo.
        const regaloCell = _serializeRegalo(r.regalo, r.pagador);
        if (regaloCell) sheet.getRange(sheet.getLastRow(), 33).setValue(regaloCell);
      } catch(appendErr) {
        Logger.log('  ✗ appendRow THREW: ' + appendErr.message + ' stack: ' + appendErr.stack);
        logDebugEntry('saveReservation-FAIL', { id: r.id, error: appendErr.message, stack: appendErr.stack ? String(appendErr.stack).slice(0, 400) : '' });
        throw appendErr;
      }

      // Si el teléfono matchea un lead del Agente que estaba cerrando, marcar
      // su conversación como CONFIRMED (ahora sí existe la reserva → cuenta).
      try { _botConfirmConversationByPhone(r.telefono); } catch(_) {}

      if (payload.fechaAnterior) {
        const fa   = payload.fechaAnterior;
        const nota = '📅 Entrada: ' + fa.checkin + ' → ' + r.checkin
                   + '  |  Salida: ' + fa.checkout + ' → ' + r.checkout;
        const allData = sheet.getDataRange().getValues();
        for (let i = allData.length - 1; i >= 1; i--) {
          if (allData[i][0] && allData[i][0].toString() === r.id.toString()) {
            const cell = sheet.getRange(i + 1, 14);
            const prev = cell.getValue();
            cell.setValue(prev ? prev + ' | ' + nota : nota);
            cell.setBackground('#E3F2FD');
            cell.setFontColor('#1565C0');
            sheet.getRange(i + 1, 1, 1, 13).setBackground('#E8F4FD');
            break;
          }
        }
      }

      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, status: 'saved' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── UPDATE RESERVATION ────────────────────────────────────
    if (action === 'updateReservation') {
      const r     = payload.reservation;
      const sheet = getOrCreateSheet();
      const data  = sheet.getDataRange().getValues();

      const CABIN_NAMES = {
        verde: 'Paseo por Las Nubes',
        azul:  'Portal hacia Las Nubes',
        lila:  'Puente entre Las Nubes'
      };

      const amount_    = parseFloat(r.amount)    || 0;
      const deposit_   = parseFloat(r.deposit)   || 0;
      const serviceFee_= parseFloat(r.serviceFee)|| 0;
      const neto_      = parseFloat(r.neto)       || amount_;
      const persons_   = parseInt(r.persons)      || 1;
      const today_     = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');

      const stripId = id => id.toString().replace(/^(airbnb_)+/, '');
      const rawId = stripId(r.id);

      for (let i = 1; i < data.length; i++) {
        const sheetId = data[i][0] ? stripId(data[i][0].toString()) : '';
        if (sheetId && sheetId === rawId) {
          const row = i + 1;
          sheet.getRange(row, 1, 1, 25).setValues([[
            r.id,
            _safeCell(r.name),
            CABIN_NAMES[r.cabin] || r.cabin,
            r.cabin,
            r.checkin,
            r.checkout,
            persons_,
            amount_,
            deposit_,
            r.origin    || 'Directa',
            _sanitizeConfirmCode(r.confirmCode) || _sanitizeConfirmCode(data[i][10]) || r.id,
            serviceFee_,
            neto_,
            data[i][13] || '',
            _safeCell(r.pagador || r.name),
            r.fechaReserva || today_,
            data[i][16] || '',
            data[i][17] || 0,
            _safeCell(r.codTransferencia || data[i][18] || ''),
            r.montoVoucher || data[i][19] || '',
            (function() {
              const existingEstado = data[i][20] ? data[i][20].toString().trim() : '';
              const newEstado      = r.estadoPago ? r.estadoPago.toString().trim() : '';
              if (existingEstado === 'CANCELADA') return 'CANCELADA';
              if (existingEstado === 'PAGA' && newEstado === 'PENDIENTE') return 'PAGA';
              return newEstado || existingEstado || '';
            })(),
            _safeCell(r.email || data[i][21] || ''),
            _safeCell(r.comentarios || data[i][22] || ''),
            _safeCell(r.telefono || data[i][23] || ''),
            r.tipo         || data[i][24] || 'noche'
          ]]);
          // Col 29 (CheckoutExtendido) — actualizar el flag de cortesia
          sheet.getRange(row, 29).setValue(r.checkoutExtendido ? true : false);
          // Cols 30/31 (HoraEntrada/HoraSalida) — horas custom del check-in
          // / check-out. Setear siempre (incluso a vacío) para permitir
          // borrar el override al editar.
          sheet.getRange(row, 30).setValue(_normalizeHora(r.horaEntrada));
          sheet.getRange(row, 31).setValue(_normalizeHora(r.horaSalida));
          // Col 33 (Regalo) — setear siempre para permitir desmarcar el regalo
          // al editar. Vacío = reserva normal.
          sheet.getRange(row, 33).setValue(_serializeRegalo(r.regalo, r.pagador));

          if (payload.fechaAnterior) {
            const fa   = payload.fechaAnterior;
            const stamp = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');
            const nota  = '📅 Entrada: ' + fa.checkin + ' → ' + r.checkin
                        + '  |  Salida: ' + fa.checkout + ' → ' + r.checkout
                        + ' (' + stamp + ')';
            const alertaCell = sheet.getRange(row, 14);
            const prev = alertaCell.getValue();
            alertaCell.setValue(prev ? prev + ' | ' + nota : nota);
          }

          return ContentService
            .createTextOutput(JSON.stringify({ ok: true, status: 'updated' }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, status: 'not_found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── DELETE RESERVATION ────────────────────────────────────
    if (action === 'deleteReservation') {
      const stripId = s => s.toString().replace(/^(airbnb_|csv_)+/g, '');
      const id          = payload.id          ? stripId(payload.id)          : '';
      const confirmCode = payload.confirmCode ? stripId(payload.confirmCode) : '';
      const deleteVoucher = payload.deleteVoucher === true;
      const sheet = getOrCreateSheet();
      const data  = sheet.getDataRange().getValues();

      for (let i = data.length - 1; i >= 1; i--) {
        const rowId   = data[i][0]  ? stripId(data[i][0].toString())  : '';
        const rowCode = data[i][10] ? stripId(data[i][10].toString()) : '';
        const match   = (id && rowId === id) || (confirmCode && rowCode === confirmCode);
        if (match) {
          if (deleteVoucher) {
            try {
              // 1) Vía precisa: usar VoucherURL (col 26) y extraer fileId(s).
              // Soporta multiples URLs separadas por '|' (acumuladas via saveVoucherToDrive).
              const voucherUrl = data[i][25] ? data[i][25].toString() : '';
              const urls = voucherUrl.split('|').map(s => s.trim()).filter(Boolean);
              const ids = urls
                .map(u => { const mm = u.match(/\/d\/([A-Za-z0-9_-]+)/); return mm ? mm[1] : null; })
                .filter(Boolean);
              if (ids.length > 0) {
                for (const fid of ids) {
                  try {
                    DriveApp.getFileById(fid).setTrashed(true);
                    Logger.log('✓ Voucher trasheado por URL: ' + fid);
                  } catch(idErr) {
                    Logger.log('⚠ No se pudo trashear por fileId ' + fid + ': ' + idErr.message);
                  }
                }
              } else {
                // 2) Fallback fuzzy: buscar por código de transferencia / nombre del huésped
                const codTransf = data[i][18] ? data[i][18].toString() : '';
                const nombre    = data[i][1]  ? data[i][1].toString().replace(/\s+/g,'_').slice(0,20) : '';
                const folders   = DriveApp.getFoldersByName(VOUCHER_FOLDER_NAME);
                if (folders.hasNext()) {
                  const folder = folders.next();
                  const files  = folder.getFiles();
                  while (files.hasNext()) {
                    const file = files.next();
                    const desc = file.getDescription() || '';
                    const fname = file.getName() || '';
                    if ((codTransf && desc.includes(codTransf)) ||
                        (nombre && fname.toLowerCase().includes(nombre.toLowerCase()))) {
                      file.setTrashed(true);
                    }
                  }
                }
              }
            } catch(driveErr) {
              Logger.log('⚠ No se pudo eliminar voucher: ' + driveErr.message);
            }
          }
          sheet.deleteRow(i + 1);
          return ContentService
            .createTextOutput(JSON.stringify({ ok: true, status: 'deleted' }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, status: 'not_found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── CANCEL RESERVATION ────────────────────────────────────
    if (action === 'cancelReservation') {
      const stripId = s => s.toString().replace(/^(airbnb_|csv_)+/g, '');
      const id          = payload.id          ? stripId(payload.id)          : '';
      const confirmCode = payload.confirmCode ? stripId(payload.confirmCode) : '';
      const sheet = getOrCreateSheet();
      const data  = sheet.getDataRange().getValues();

      for (let i = data.length - 1; i >= 1; i--) {
        const rowId   = data[i][0]  ? stripId(data[i][0].toString())  : '';
        const rowCode = data[i][10] ? stripId(data[i][10].toString()) : '';
        const match   = (id && rowId === id) || (confirmCode && rowCode === confirmCode);
        if (match) {
          const numCols = sheet.getLastColumn();
          if (numCols < 14) {
            sheet.getRange(1, 14).setValue('Alerta');
            sheet.getRange(1, 14).setFontWeight('bold');
          }
          const cell  = sheet.getRange(i + 1, 14);
          const stamp = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');
          const prev  = cell.getValue();
          const nota  = '❌ Cancelado ' + stamp;
          cell.setValue(prev ? prev + ' | ' + nota : nota);
          cell.setBackground('#FDECEA');
          cell.setFontColor('#B71C1C');
          sheet.getRange(i + 1, 1, 1, 14).setFontLine('line-through');
          sheet.getRange(i + 1, 1, 1, 14).setBackground('#FFF5F5');
          sheet.getRange(i + 1, 21).setValue('CANCELADA');
          const cCode = data[i][10] ? data[i][10].toString() : '';
          if (cCode) addToBlacklist(cCode, 'Cancelada ' + stamp);
          return ContentService
            .createTextOutput(JSON.stringify({ ok: true, status: 'cancelled' }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, status: 'not_found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── SAVE SUMINISTROS KEYWORDS (reemplaza la lista completa) ──
    if (action === 'saveSuministrosItems') {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sh = _getSuministrosItemsSheet(ss);
      // Limpiar filas de datos y reescribir la lista completa que manda el front.
      if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,7).clearContent();
      const kws = (payload.keywords || []).map(k => (k||'').toString().trim()).filter(Boolean);
      const norm = k => (k||'').toString().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g,'').trim();
      // Keywords marcadas "solo detalle / no timeline" (ej. Delivery), sin acentos.
      const noSet = {};
      (payload.noTimeline || []).forEach(k => { noSet[norm(k)] = 1; });
      // Keywords de REVENTA: sus compras son el costo de la tiendita.
      const revSet = {};
      (payload.reventa || []).forEach(k => { revSet[norm(k)] = 1; });
      // Productos de la tiendita (los que llegan del import de la web).
      const ventaSet = {};
      (payload.venta || []).forEach(k => { ventaSet[norm(k)] = 1; });
      // Costo unitario fijo por keyword (mano de obra, insumos sin factura).
      const costos = {};
      Object.keys(payload.costos || {}).forEach(k => {
        const v = parseFloat(payload.costos[k]);
        if (!isNaN(v) && v > 0) costos[norm(k)] = v;
      });
      // Precio de venta por keyword (normalizada, para que no dependa de acentos).
      const precios = {};
      Object.keys(payload.precios || {}).forEach(k => {
        const v = parseFloat(payload.precios[k]);
        if (!isNaN(v) && v > 0) precios[norm(k)] = v;
      });
      // Recetas: { keyword: { componente_normalizado: cantidad } }. Se guardan
      // como JSON en la col 5. Las cantidades <= 0 se descartan acá para que una
      // receta no llegue a la hoja con componentes que no aportan nada.
      const recetas = {};
      Object.keys(payload.recetas || {}).forEach(k => {
        const src = payload.recetas[k] || {};
        const limpia = {};
        Object.keys(src).forEach(c => {
          const q = parseFloat(src[c]);
          if (!isNaN(q) && q > 0) limpia[norm(c)] = q;
        });
        if (Object.keys(limpia).length) recetas[norm(k)] = limpia;
      });
      if (kws.length) {
        sh.getRange(2,1,kws.length,7).setValues(
          kws.map(k => [k, noSet[norm(k)] ? true : false, revSet[norm(k)] ? true : false,
                        precios[norm(k)] || '',
                        recetas[norm(k)] ? JSON.stringify(recetas[norm(k)]) : '',
                        ventaSet[norm(k)] ? true : false,
                        costos[norm(k)] || '']));
      }
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, count: kws.length }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── TIENDITA ──────────────────────────────────────────────
    if (action === 'saveTienditaVenta' || action === 'deleteTienditaVenta'
        || action === 'saveTienditaVoucher') {
      try {
        let r;
        if (action === 'saveTienditaVenta')       r = saveTienditaVenta(payload);
        else if (action === 'deleteTienditaVenta')r = deleteTienditaVenta(payload.ventaId);
        else                                      r = saveTienditaVoucherToDrive(payload);
        return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ── PRÉSTAMOS A COLABORADORES ─────────────────────────────
    if (action === 'savePrestamo' || action === 'savePrestamoAbono'
        || action === 'deletePrestamoAbono' || action === 'deletePrestamo'
        || action === 'savePrestamoFactura') {
      try {
        let r;
        if (action === 'savePrestamo')            r = savePrestamo(payload);
        else if (action === 'savePrestamoAbono')  r = savePrestamoAbono(payload);
        else if (action === 'deletePrestamoAbono')r = deletePrestamoAbono(payload.abonoId);
        else if (action === 'deletePrestamo')     r = deletePrestamo(payload.prestamoId);
        else                                      r = savePrestamoFacturaToDrive(payload);
        return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ── SAVE EGRESOS (multi-item) ─────────────────────────────
    if (action === 'saveEgresos') {
      const items = payload.egresos;
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const egresoSheet = _getEgresoSheetEnsured(ss);
      // Rechazo server-side de duplicados (salvo que el usuario fuerce). Atrapa
      // la MISMA factura re-subida con otra fecha/otro código.
      if (!payload.force) {
        const dupReason = _egresoBatchDuplicate(egresoSheet, items);
        if (dupReason) {
          return ContentService
            .createTextOutput(JSON.stringify({ ok: false, status: 'duplicate', reason: dupReason }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      const saved = [];
      items.forEach(eg => {
        egresoSheet.appendRow([
          eg.id        || '',
          eg.date      || '',
          eg.desc      || '',
          parseFloat(eg.amount) || 0,
          eg.cat       || 'Otro',
          eg.cabin     || '',
          eg.proveedor || '',
          eg.urlFoto   || '',
          eg.item      || ''
        ]);
        saved.push(eg.id);
      });
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, saved }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── SAVE EGRESO (singular) ────────────────────────────────
    if (action === 'saveEgreso') {
      const eg = payload.egreso;
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const egresoSheet = _getEgresoSheetEnsured(ss);
      egresoSheet.appendRow([
        eg.id        || '',
        eg.date      || '',
        eg.desc      || '',
        parseFloat(eg.amount) || 0,
        eg.cat       || 'Otro',
        eg.cabin     || '',
        eg.proveedor || '',
        eg.urlFoto   || '',
        eg.item      || ''
      ]);
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, status: 'saved' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── DELETE EGRESO ─────────────────────────────────────────
    if (action === 'deleteEgreso') {
      const eg = payload.egreso;
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const egresoSheet = ss.getSheetByName('Egresos');
      if (!egresoSheet) {
        return ContentService
          .createTextOutput(JSON.stringify({ ok: false, error: 'No Egresos sheet' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const eData = egresoSheet.getDataRange().getValues();
      for (let i = eData.length - 1; i >= 1; i--) {
        const rowId = eData[i][0] ? eData[i][0].toString() : '';
        const rowDate = eData[i][1] instanceof Date
          ? Utilities.formatDate(eData[i][1], 'America/Panama', 'yyyy-MM-dd')
          : eData[i][1].toString().slice(0,10);
        const matchById   = eg.id && rowId === eg.id;
        const matchByData = !eg.id && rowDate === eg.date && eData[i][2] === eg.desc && parseFloat(eData[i][3]) === parseFloat(eg.amount);
        if (matchById || matchByData) {
          const urlFoto = eData[i][7] ? eData[i][7].toString() : '';
          let driveDeleted = false;
          if (urlFoto) {
            try {
              const match = urlFoto.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) ||
                            urlFoto.match(/[?&]id=([a-zA-Z0-9_-]+)/);
              if (match) {
                DriveApp.getFileById(match[1]).setTrashed(true);
                driveDeleted = true;
              }
            } catch(driveErr) {
              Logger.log('⚠ No se pudo eliminar foto: ' + driveErr.message);
            }
          }
          egresoSheet.deleteRow(i + 1);
          return ContentService
            .createTextOutput(JSON.stringify({ ok: true, status: 'deleted', driveDeleted }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, status: 'not_found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── UPDATE EGRESO ─────────────────────────────────────────
    if (action === 'updateEgreso') {
      const eg = payload.egreso;
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const sheet = ss.getSheetByName('Egresos');
      if (!sheet) return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: 'No Egresos sheet' }))
        .setMimeType(ContentService.MimeType.JSON);
      // Asegurar columnas Item/FechaFin/MontosItem/CantidadesItem si el update las va a escribir.
      if (eg.item !== undefined || eg.fechaFin !== undefined || eg.montosItem !== undefined || eg.cantidadesItem !== undefined) { _getEgresoSheetEnsured(ss); }
      const itemCol       = eg.item       !== undefined ? _egresoColIndex(sheet, 'item')       : 0;
      const fechaFinCol   = eg.fechaFin   !== undefined ? _egresoColIndex(sheet, 'fechafin')   : 0;
      const montosItemCol = eg.montosItem !== undefined ? _egresoColIndex(sheet, 'montositem') : 0;
      const cantItemCol   = eg.cantidadesItem !== undefined ? _egresoColIndex(sheet, 'cantidadesitem') : 0;
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString() === eg.id) {
          if (eg.cat       !== undefined) sheet.getRange(i+1, 5).setValue(eg.cat);
          if (eg.desc      !== undefined) sheet.getRange(i+1, 3).setValue(eg.desc);
          // Fecha y monto faltaban: eran los dos campos que más se corrigen
          // (una factura cargada con la fecha de hoy, un OCR que leyó de más)
          // y no había forma de arreglarlos sin borrar y volver a cargar.
          if (eg.fecha     !== undefined) sheet.getRange(i+1, 2).setValue(eg.fecha);
          if (eg.monto     !== undefined) sheet.getRange(i+1, 4).setValue(parseFloat(eg.monto) || 0);
          if (eg.proveedor !== undefined) sheet.getRange(i+1, 7).setValue(eg.proveedor);
          if (eg.cabin     !== undefined) sheet.getRange(i+1, 6).setValue(eg.cabin);
          if (eg.item      !== undefined && itemCol > 0)       sheet.getRange(i+1, itemCol).setValue(eg.item);
          if (eg.fechaFin  !== undefined && fechaFinCol > 0)   sheet.getRange(i+1, fechaFinCol).setValue(eg.fechaFin || '');
          if (eg.montosItem !== undefined && montosItemCol > 0) sheet.getRange(i+1, montosItemCol).setValue(eg.montosItem || '');
          if (eg.cantidadesItem !== undefined && cantItemCol > 0) sheet.getRange(i+1, cantItemCol).setValue(eg.cantidadesItem || '');
          return ContentService
            .createTextOutput(JSON.stringify({ ok: true, status: 'updated' }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, status: 'not_found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── SAVE FOTO EGRESO TO DRIVE ─────────────────────────────
    if (action === 'saveFotoEgreso') {
      const { imageBase64, mimeType, egresoId, fecha } = payload;
      return saveFotoEgreso(imageBase64, mimeType, egresoId, fecha);
    }

    // ── UPDATE EGRESO FOTO ────────────────────────────────────
    if (action === 'updateEgresoFoto') {
      const { id, urlFoto } = payload;
      const ss = SpreadsheetApp.openById(SHEET_ID);
      const egresoSheet = ss.getSheetByName('Egresos');
      if (!egresoSheet) return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: 'No sheet' }))
        .setMimeType(ContentService.MimeType.JSON);
      const data = egresoSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][0].toString() === id) {
          egresoSheet.getRange(i+1, 8).setValue(urlFoto);
          return ContentService
            .createTextOutput(JSON.stringify({ ok: true }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: 'not_found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── PARSE FACTURA EGRESO ──────────────────────────────────
    if (action === 'parseFacturaEgreso') {
      return parseFacturaEgresoConClaude(payload.imageBase64, payload.mimeType);
    }

    // ── PARSE VOUCHER ─────────────────────────────────────────
    if (action === 'parseVoucher') {
      return parseVoucherWithClaude(payload.imageBase64, payload.mimeType);
    }

    // ── SAVE VOUCHER TO DRIVE ─────────────────────────────────
    if (action === 'saveVoucherToDrive') {
      return saveVoucherToDrive(payload.reservation, payload.imageBase64, payload.mimeType, payload.fileName, payload.voucherMeta);
    }

    // ── MALAYA: SAVE RESERVA ─────────────────────────────────
    // Reserva directa de Malaya Lodge (cabaña referida). Detalles en Malaya.gs.
    if (action === 'saveMalayaReserva') {
      try {
        const result = saveMalayaReserva(payload);
        return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
      } catch(e) {
        return ContentService.createTextOutput(JSON.stringify({ ok: false, error: e.message })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ── MALAYA: SUBIR VOUCHER A DRIVE ──────────────────────────
    if (action === 'saveMalayaVoucher') {
      try {
        const result = saveMalayaVoucherToDrive(payload);
        return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
      } catch(e) {
        return ContentService.createTextOutput(JSON.stringify({ ok: false, error: e.message })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ── MALAYA: FORZAR SYNC AHORA ──────────────────────────────
    // Dispara syncMalayaAirbnb on-demand y devuelve el snapshot del iCal
    // resultante. Útil para diagnosticar desde el admin panel sin tener
    // que abrir el editor.
    if (action === 'syncMalayaNow') {
      try {
        syncMalayaAirbnb(true);   // el admin pidió el sync: saltea el guard de cadencia
        const sheet = _malayaIcalSheet();
        const data  = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues() : [];
        const events = data.map(r => ({
          checkin:  r[0] instanceof Date ? Utilities.formatDate(r[0], 'America/Panama', 'yyyy-MM-dd') : String(r[0]).slice(0,10),
          checkout: r[1] instanceof Date ? Utilities.formatDate(r[1], 'America/Panama', 'yyyy-MM-dd') : String(r[1]).slice(0,10),
          summary:  String(r[2] || ''),
          // El UID es el ÚNICO dato que distingue un bloqueo propio de Celestino
          // del eco de nuestro feed: Airbnb rotula los dos como "Airbnb (Not
          // available)". Va al panel para no depender del editor de Apps Script.
          uid:      String(r[3] || ''),
          tipo:     _malayaClasificarEvento({ summary: String(r[2] || ''), uid: String(r[3] || '') })
        }));
        const url = PropertiesService.getScriptProperties().getProperty('MALAYA_AIRBNB_ICAL') || null;
        // Lo que de verdad se quiere saber después de sincronizar: si mis
        // fechas quedaron bloqueadas. El listado de eventos crudos solo lo
        // dice si uno se pone a cruzarlo a mano.
        const bloqueos = getMalayaEstadoBloqueos();
        return ContentService.createTextOutput(JSON.stringify({
          ok: true, eventCount: events.length, events: events,
          icalUrlConfigured: !!url,
          bloqueos: bloqueos.reservas, tiposEventos: bloqueos.tiposEventos
        })).setMimeType(ContentService.MimeType.JSON);
      } catch(e) {
        return ContentService.createTextOutput(JSON.stringify({ ok: false, error: e.message })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ── MALAYA: CANCEL RESERVA ─────────────────────────────────
    if (action === 'cancelMalayaReserva') {
      try {
        const result = cancelMalayaReserva(payload.reservaId);
        return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
      } catch(e) {
        return ContentService.createTextOutput(JSON.stringify({ ok: false, error: e.message })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ── MALAYA: UPDATE RESERVA (editar fecha/personas/monto/notas) ──
    if (action === 'updateMalayaReserva') {
      try {
        const result = updateMalayaReserva(payload);
        return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
      } catch(e) {
        return ContentService.createTextOutput(JSON.stringify({ ok: false, error: e.message })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ── UPLOAD HUESPED ID (publico, gating del key box) ───────
    if (action === 'uploadHuespedId') {
      return handleUploadHuespedId(payload);
    }

    // ── LOYALTY (admin) ───────────────────────────────────────
    if (action === 'markLoyaltyUsed')   return handleMarkLoyaltyUsed(payload);
    if (action === 'unmarkLoyaltyUsed') return handleUnmarkLoyaltyUsed(payload);

    // ── REFERIDOS (admin) ─────────────────────────────────────
    if (action === 'registerReferralUse')   return handleRegisterReferralUse(payload);
    if (action === 'markReferrerCreditUsed') return handleMarkReferrerCreditUsed(payload);

    // ── SEND EMAILS ───────────────────────────────────────────
    if (action === 'sendCancellationEmail') return sendCancellationEmail(payload.reservation);
    if (action === 'sendConfirmationEmail') return sendConfirmationEmail(payload.reservation, payload.voucherBase64, payload.voucherMimeType);
    // Vista previa: envía la MISMA confirmación (email + WhatsApp) al admin, no al
    // cliente. Sirve para revisar cómo le llegará al huésped antes de mandarla.
    if (action === 'previewConfirmacionSelf') {
      const r = payload.reservation || {};
      const ch = payload.channels || { email: true, whatsapp: true };
      const result = {};
      if (ch.email) {
        try {
          const clone = Object.assign({}, r, { email: REPLY_TO_EMAIL });
          sendConfirmationEmail(clone, payload.voucherBase64, payload.voucherMimeType, '👁 Vista previa · ');
          result.email = { ok: true, to: REPLY_TO_EMAIL };
        } catch(e) { result.email = { ok: false, error: e.message }; }
      }
      if (ch.whatsapp) {
        try {
          // Celular PERSONAL del admin para vistas previas (NO el número emisor
          // del bot). Configurable por Script Property; default 6981-2266.
          const previewPhone = PropertiesService.getScriptProperties().getProperty('PREVIEW_NOTIFY_PHONE') || '50769812266';
          const clone = Object.assign({}, r, { telefono: previewPhone });
          sendWAAvisoReserva(clone);
          result.whatsapp = { ok: true, to: previewPhone };
        } catch(e) { result.whatsapp = { ok: false, error: e.message }; }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, result: result }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (action === 'sendUpdateEmail')       return sendUpdateEmail(payload.reservation, payload.voucherBase64, payload.voucherMimeType);
    if (action === 'sendCheckinReminder')   return sendCheckinReminderEmail(payload.reservation);

    // ── BOT CRM ───────────────────────────────────────────────
    if (action === 'deleteConversation') return handleDeleteConversation(payload);
    if (action === 'saveBotAlertConfig') {
      const cfg = _botSetAlertConfig(payload.config || {});
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, config: cfg }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── SEND WHATSAPP ─────────────────────────────────────────
    if (action === 'sendWAConfirmacion') {
      try {
        const result = sendWAAvisoReserva(payload.reservation);
        return ContentService
          .createTextOutput(JSON.stringify({ ok: true, result: result }))
          .setMimeType(ContentService.MimeType.JSON);
      } catch(waErr) {
        return ContentService
          .createTextOutput(JSON.stringify({ ok: false, error: waErr.message }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ── SYNC PAGOS Y ESTADOS ──────────────────────────────────
    if (action === 'syncPayoutsYEstados') {
      try {
        syncAirbnbPayouts();
        actualizarEstadoPagoAirbnb();
        return ContentService
          .createTextOutput(JSON.stringify({ ok: true, status: 'pagos_sincronizados' }))
          .setMimeType(ContentService.MimeType.JSON);
      } catch(e) {
        return ContentService
          .createTextOutput(JSON.stringify({ ok: false, error: e.toString() }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: 'Unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    logDebugEntry('doPost-CRASH', { action: _action, id: _id, error: err.message, stack: err.stack ? String(err.stack).slice(0, 400) : '' });
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── DEBUG LOG ───────────────────────────────────────────────
// Persiste eventos importantes a una hoja 'DebugLog' para diagnosticar
// problemas que no se ven en los logs de Apps Script (ej. fallos en iOS).
// Escapa valores que empiezan con caracteres que Sheets interpreta como
// formula (=, +, -, @). Sin esto, appendRow con telefono "+507 ..." hace
// que Sheets intente evaluar "=+507 ..." como expresion → error o timeout
// → respuesta HTTP 404 desde el frontend de Google Apps Engine.
function _safeCell(v) {
  if (v === null || v === undefined || v === '') return v;
  const s = String(v);
  const c = s.charAt(0);
  if (c === '=' || c === '+' || c === '-' || c === '@') return "'" + s;
  return s;
}

function logDebugEntry(stage, info) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('DebugLog');
    if (!sheet) {
      sheet = ss.insertSheet('DebugLog');
      sheet.getRange(1, 1, 1, 3).setValues([['Timestamp', 'Stage', 'Info']]);
      sheet.setFrozenRows(1);
    }
    const ts = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd HH:mm:ss');
    const infoStr = (typeof info === 'object') ? JSON.stringify(info).slice(0, 800) : String(info);
    sheet.appendRow([ts, stage, infoStr]);
    // Cap at last 300 entries (header + 300)
    const last = sheet.getLastRow();
    if (last > 301) sheet.deleteRows(2, last - 301);
  } catch(_e) { /* swallow */ }
}

function handleGetDebugLog(e) {
  const limit = parseInt((e && e.parameter && e.parameter.limit) || '40', 10);
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('DebugLog');
  if (!sheet || sheet.getLastRow() < 2) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, entries: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const last     = sheet.getLastRow();
  const startRow = Math.max(2, last - limit + 1);
  const numRows  = last - startRow + 1;
  const data     = sheet.getRange(startRow, 1, numRows, 3).getValues();
  const entries  = data.map(r => ({ ts: r[0], stage: r[1], info: r[2] })).reverse();
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, entries: entries }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════
//  Backfill & Cleanup utilities
// ═══════════════════════════════════════════════════════════

function backfillFechaReserva() {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();

  if (!data[0][15] || data[0][15].toString().trim() === '') {
    sheet.getRange(1, 16).setValue('FechaReserva');
    sheet.getRange(1, 16).setFontWeight('bold');
  }

  const pendingCodes = new Set();
  for (let i = 1; i < data.length; i++) {
    const origin      = data[i][9]  ? data[i][9].toString().trim()  : '';
    const confirmCode = data[i][10] ? data[i][10].toString().trim() : '';
    const existingFecha = data[i][15] ? data[i][15].toString().trim() : '';
    if (origin === 'Airbnb' && confirmCode && existingFecha === '') {
      pendingCodes.add(confirmCode);
    }
  }

  Logger.log('Códigos Airbnb sin fecha en el Sheet: ' + pendingCodes.size);
  if (pendingCodes.size === 0) return;

  const emailDateMap = {};
  pendingCodes.forEach(code => {
    const threads = GmailApp.search(
      'from:automated@airbnb.com subject:"Reserva confirmada:" "' + code + '"'
    );
    threads.forEach(thread => {
      thread.getMessages().forEach(msg => {
        if (msg.getPlainBody().includes(code)) {
          emailDateMap[code] = Utilities.formatDate(msg.getDate(), 'America/Panama', 'yyyy-MM-dd');
        }
      });
    });
  });

  let updated = 0;
  for (let i = 1; i < data.length; i++) {
    const existingFecha = data[i][15] ? data[i][15].toString().trim() : '';
    if (existingFecha !== '') continue;
    const confirmCode = data[i][10] ? data[i][10].toString().trim() : '';
    if (confirmCode && emailDateMap[confirmCode]) {
      sheet.getRange(i + 1, 16).setValue(emailDateMap[confirmCode]);
      updated++;
    }
  }
  Logger.log('Backfill completado. Actualizadas: ' + updated);
}

function eliminarDuplicadosViejos() {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();

  const IDS_TO_DELETE = new Set([
    "1771865495959","1771865622532","1771865651637","1771865804032",
    "1771865901592","1771866028468","1771866069021","1771866471217",
    "1771866496624","1771866542519","1771866639302","1771866666078",
    "1771866696882","1771866826182","1771866891933","airbnb_1771866891933",
    "1771867189585","1771867355750","1771867531677","airbnb_1771867531677",
    "1771867595457"
  ]);

  const rowsToDelete = [];
  for (let i = 1; i < data.length; i++) {
    const id = data[i][0] ? data[i][0].toString().trim() : '';
    if (IDS_TO_DELETE.has(id)) rowsToDelete.push(i + 1);
  }

  rowsToDelete.sort((a, b) => b - a);
  rowsToDelete.forEach(rowNum => sheet.deleteRow(rowNum));
  Logger.log('✓ ' + rowsToDelete.length + ' filas eliminadas.');
}

function migrarColumnas() {
  const sheet = getOrCreateSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let cambios = 0;
  ['MontoVoucher','EstadoPago'].forEach(h => {
    if (!headers.includes(h)) {
      const col = headers.length + 1;
      sheet.getRange(1, col).setValue(h);
      sheet.getRange(1, col).setFontWeight('bold');
      headers.push(h);
      cambios++;
    }
  });
  Logger.log(cambios ? '✓ ' + cambios + ' columna(s) agregada(s)' : '✓ Columnas ya existían');
}

function migrarColumnasV2() {
  const sheet = getOrCreateSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let cambios = 0;
  ['Email','Comentarios'].forEach(h => {
    if (!headers.includes(h)) {
      const col = headers.length + 1;
      sheet.getRange(1, col).setValue(h);
      sheet.getRange(1, col).setFontWeight('bold');
      headers.push(h);
      cambios++;
    }
  });
  Logger.log(cambios ? '✓ Migración V2 — ' + cambios + ' columna(s)' : '✓ Columnas ya existían');
}

function migrarColumnasV3() {
  const sheet = getOrCreateSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let cambios = 0;
  ['Telefono'].forEach(h => {
    if (!headers.includes(h)) {
      const col = headers.length + 1;
      sheet.getRange(1, col).setValue(h);
      sheet.getRange(1, col).setFontWeight('bold');
      headers.push(h);
      cambios++;
    }
  });
  Logger.log(cambios ? '✓ Migración V3 — ' + cambios + ' columna(s)' : '✓ Columnas V3 ya existían');
}

// V4: agrega columna Tipo (noche | pasadia | pasadia-largo | early | late) en col 25.
// V5 (rename): pasadia (12:30-7pm) -> pasatarde, pasadia-largo (9am-5pm) -> pasadia. Ver migrarTiposV5.
// Default 'noche' SOLO en celdas vacías (preserva valores ya escritos por saveReservation).
function migrarColumnasV4() {
  const sheet = getOrCreateSheet();
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  const headers = sheet.getRange(1, 1, 1, Math.max(lastCol, 25)).getValues()[0];

  const TIPO_COL = 25;
  // Si Tipo ya está en alguna columna, no tocamos nada
  if (headers.includes('Tipo')) {
    Logger.log('✓ Columna Tipo ya existía en col ' + (headers.indexOf('Tipo') + 1));
    return;
  }
  // Forzar header en col 25 (no usar headers.length+1 porque appendRow pudo haber extendido el sheet)
  sheet.getRange(1, TIPO_COL).setValue('Tipo');
  sheet.getRange(1, TIPO_COL).setFontWeight('bold');

  if (lastRow >= 2) {
    const range = sheet.getRange(2, TIPO_COL, lastRow - 1, 1);
    const existing = range.getValues();
    let updated = 0;
    for (let i = 0; i < existing.length; i++) {
      if (!existing[i][0] || existing[i][0].toString().trim() === '') {
        existing[i][0] = 'noche';
        updated++;
      }
    }
    if (updated > 0) range.setValues(existing);
    Logger.log('✓ Migración V4 — header en col 25, ' + updated + ' filas con default "noche" (preservadas las que ya tenían valor)');
    return;
  }
  Logger.log('✓ Migración V4 — header Tipo agregado, sin filas existentes');
}

// V5: renombrar tipos de reserva para coherencia semántica.
//   pasadia (12:30-7pm)  -> pasatarde
//   pasadia-largo (9-5)  -> pasadia
// Idempotente: usa Script Property MIGRATION_V5_TIPO_RENAME='done'.
function migrarTiposV5() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('MIGRATION_V5_TIPO_RENAME') === 'done') {
    Logger.log('✓ V5 ya aplicada, skip');
    return { skipped: true };
  }
  const sheet = getOrCreateSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    props.setProperty('MIGRATION_V5_TIPO_RENAME', 'done');
    Logger.log('✓ V5 — sin filas que migrar');
    return { migrated: 0 };
  }
  const TIPO_COL = 25;
  const range = sheet.getRange(2, TIPO_COL, lastRow - 1, 1);
  const values = range.getValues();
  let count = 0;
  const out = values.map(row => {
    const t = (row[0] || '').toString();
    if (t === 'pasadia-largo') { count++; return ['pasadia']; }
    if (t === 'pasadia')       { count++; return ['pasatarde']; }
    return row;
  });
  range.setValues(out);
  props.setProperty('MIGRATION_V5_TIPO_RENAME', 'done');
  Logger.log('✓ V5 — ' + count + ' filas migradas (pasadia→pasatarde, pasadia-largo→pasadia)');
  return { migrated: count };
}

function limpiarDuplicadosAirbnb() {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const idsLimpios = new Set();
  for (let i = 1; i < data.length; i++) {
    const id = data[i][0] ? data[i][0].toString() : '';
    if (!id.startsWith('airbnb_')) idsLimpios.add(id);
  }
  const filasAEliminar = [];
  for (let i = 1; i < data.length; i++) {
    const id = data[i][0] ? data[i][0].toString() : '';
    if (id.startsWith('airbnb_') && idsLimpios.has(id.replace(/^airbnb_/, ''))) {
      filasAEliminar.push(i + 1);
    }
  }
  filasAEliminar.reverse().forEach(row => sheet.deleteRow(row));
  Logger.log('✓ ' + filasAEliminar.length + ' duplicados airbnb_ eliminados');
}

// ═══════════════════════════════════════════════════════════
//  Email helpers
// ═══════════════════════════════════════════════════════════

const CABIN_NAMES_EMAIL = {
  verde: 'Paseo por Las Nubes',
  azul:  'Portal hacia Las Nubes',
  lila:  'Puente entre Las Nubes'
};

const CABIN_COLORS_EMAIL = {
  verde: '#6a9e62',
  azul:  '#5a85b0',
  lila:  '#8060b0'
};

const REPLY_TO_EMAIL = 'lasnubesenchica@gmail.com';

function formatDateES(dateStr) {
  const parts = dateStr.toString().slice(0,10).split('-');
  const y = parseInt(parts[0]), m = parseInt(parts[1]), d = parseInt(parts[2]);
  const months = ['enero','febrero','marzo','abril','mayo','junio',
                  'julio','agosto','septiembre','octubre','noviembre','diciembre'];
  const days   = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'];
  const date   = new Date(y, m - 1, d);
  return days[date.getDay()] + ' ' + d + ' de ' + months[m-1] + ' de ' + y;
}

function nightCount(checkin, checkout) {
  const a = new Date(checkin.toString().slice(0,10)  + 'T12:00:00');
  const b = new Date(checkout.toString().slice(0,10) + 'T12:00:00');
  return Math.round((b - a) / (1000 * 60 * 60 * 24));
}

// Devuelve metadata visual del email según tipo de reserva (noche/pasatarde/pasadia/early/late).
// Las fechas/horas mostradas reflejan la realidad del huésped (no el rango bloqueado en hoja).
// Normaliza una hora "HH:MM" recibida desde el frontend/sheet a un string
// consistente. Si viene vacío/inválido/malformado, devuelve '' (= sin
// override). Acepta también horas guardadas por Sheets como Date (raro pero
// posible si el usuario formatea la celda).
function _normalizeHora(raw) {
  if (raw === null || raw === undefined || raw === '') return '';
  if (raw instanceof Date) {
    return Utilities.formatDate(raw, 'America/Panama', 'HH:mm');
  }
  const s = String(raw).trim();
  if (!s) return '';
  const m = s.match(/^(\d{1,2}):(\d{2})/);
  if (!m) return '';
  const h = parseInt(m[1], 10);
  const mn = parseInt(m[2], 10);
  if (isNaN(h) || h < 0 || h > 23 || isNaN(mn) || mn < 0 || mn > 59) return '';
  return (h < 10 ? '0' + h : String(h)) + ':' + (mn < 10 ? '0' + mn : String(mn));
}

// "16:30" → "4:30 pm", "20:00" → "8:00 pm", "09:00" → "9:00 am".
function _formatHora12(hhmm) {
  const s = _normalizeHora(hhmm);
  if (!s) return '';
  const parts = s.split(':');
  let h = parseInt(parts[0], 10);
  const mn = parseInt(parts[1], 10) || 0;
  const ampm = h >= 12 ? 'pm' : 'am';
  h = h % 12; if (h === 0) h = 12;
  return h + ':' + (mn < 10 ? '0' + mn : String(mn)) + ' ' + ampm;
}

// ═══════════════════════════════════════════════════════════
//  Certificados de regalo
// ═══════════════════════════════════════════════════════════
//
// Un certificado de regalo es una reserva normal (con su `tipo`, su `origen` y
// su plata registrada como ingreso) donde QUIEN PAGA y QUIEN DISFRUTA son
// personas distintas:
//   - `pagador`             → quien compró el regalo
//   - `name`/`email`/`telefono` → el beneficiario, a quien le llegan los avisos
// La marca vive en la col 33 `Regalo` como JSON `{de, mensaje}`; vacío = reserva
// normal. `de` es el nombre que se muestra como remitente del regalo (default:
// `pagador`) y `mensaje` es la dedicatoria opcional.
//
// Es un flag ORTOGONAL al `tipo` y al `origen` a propósito: un regalo sigue
// siendo una noche/pasadía y sigue entrando por Directa/Referido. Combinado con
// origen `Abierta` da el caso "regalo sin fecha, a coordinar por el beneficiario".

// Normaliza la celda 33 a { esRegalo, de, mensaje }. Tolera JSON corrupto y
// acepta que `regalo` venga ya como objeto (payloads del dashboard).
function _parseRegalo(r) {
  const vacio = { esRegalo: false, de: '', mensaje: '' };
  if (!r) return vacio;
  let obj = r.regalo;
  if (!obj) return vacio;
  if (typeof obj === 'string') {
    const raw = obj.trim();
    if (!raw) return vacio;
    try { obj = JSON.parse(raw); } catch(_) { return vacio; }
  }
  if (!obj || typeof obj !== 'object') return vacio;
  // `esRegalo:false` explícito gana sobre la presencia del objeto.
  if (obj.esRegalo === false) return vacio;
  return {
    esRegalo: true,
    de:      (obj.de      || r.pagador || '').toString().trim(),
    mensaje: (obj.mensaje || '').toString().trim()
  };
}

// Inverso de _parseRegalo: arma la celda 33. Devuelve '' si no es regalo.
function _serializeRegalo(regalo, pagadorFallback) {
  const g = _parseRegalo({ regalo: regalo, pagador: pagadorFallback });
  if (!g.esRegalo) return '';
  return JSON.stringify({ de: g.de, mensaje: g.mensaje });
}

function esReservaRegalo(r) {
  return _parseRegalo(r).esRegalo;
}

function tipoEmailMeta(r) {
  const tipo = (r.tipo || 'noche').toString();
  const checkinStored  = r.checkin instanceof Date  ? Utilities.formatDate(r.checkin,  'America/Panama', 'yyyy-MM-dd') : r.checkin.toString().slice(0,10);
  const checkoutStored = r.checkout instanceof Date ? Utilities.formatDate(r.checkout, 'America/Panama', 'yyyy-MM-dd') : r.checkout.toString().slice(0,10);
  const addDaysISO = (s, n) => {
    const d = new Date(s + 'T12:00:00');
    d.setDate(d.getDate() + n);
    return Utilities.formatDate(d, 'America/Panama', 'yyyy-MM-dd');
  };
  let displayCheckin  = checkinStored;
  let displayCheckout = checkoutStored;
  let checkinHora     = 'a partir de las 2:00 pm';
  let checkoutHora    = 'antes de las 11:00 am';
  let estanciaLabel   = 'Noches';
  let estanciaValue;
  if (tipo === 'pasatarde') {
    displayCheckout = checkinStored;
    checkinHora     = 'a partir de las 12:30 pm';
    checkoutHora    = 'salida 7:00 pm';
    estanciaLabel   = 'Estancia';
    estanciaValue   = 'Pasatarde';
  } else if (tipo === 'pasanoche') {
    // Entra 8pm día X, sale 12:30pm día X+1. Storage checkin=X, checkout=X+1
    // (idéntico a noche). Solo difiere en horario y label.
    displayCheckout = checkoutStored;
    checkinHora     = 'a partir de las 8:00 pm';
    checkoutHora    = 'antes de las 12:30 pm';
    estanciaLabel   = 'Estancia';
    estanciaValue   = 'Pasanoche';
  } else if (tipo === 'pasadia') {
    displayCheckin  = addDaysISO(checkinStored, 1);
    displayCheckout = displayCheckin;
    checkinHora     = 'a partir de las 9:00 am';
    checkoutHora    = 'salida 5:00 pm';
    estanciaLabel   = 'Estancia';
    estanciaValue   = 'Pasadía';
  } else if (tipo === 'early') {
    displayCheckin  = addDaysISO(checkinStored, 1);
    checkinHora     = 'a partir de las 9:00 am';
    estanciaValue   = 1;
  } else if (tipo === 'late') {
    displayCheckout = addDaysISO(checkoutStored, -1);
    checkoutHora    = 'antes de las 4:00 pm';
    estanciaValue   = 1;
  } else {
    estanciaValue   = nightCount(checkinStored, checkoutStored);
  }
  // Check-out extendido (cortesia): override 11:00am a 12:30pm cuando aplica
  // Solo aplica para tipos donde el checkout default es 11:00am (noche, early).
  // No tiene sentido para pasatarde/pasadia (hora fija) ni late (ya es 4pm).
  const isExtended = !!r.checkoutExtendido;
  if (isExtended && (tipo === 'noche' || tipo === 'early')) {
    checkoutHora = 'antes de las 12:30 pm (cortesía)';
  }
  // Override manual con horas custom por reserva (columnas HoraEntrada /
  // HoraSalida en la hoja). Pisa el default del tipo y también la cortesía
  // — la intención explícita del admin siempre gana.
  const horaEntradaCustom = _normalizeHora(r.horaEntrada);
  let   horaSalidaCustom  = _normalizeHora(r.horaSalida);
  // Guard: en pasatarde/pasadía entrada y salida caen el mismo día, así que la
  // salida DEBE ser posterior a la entrada. Una hora-salida custom inválida
  // (ej. 8am en un pasatarde que entra 12:30pm) se ignora — deja el default del
  // tipo ('salida 7:00 pm' / 'salida 5:00 pm') en vez de un checkout imposible.
  if (horaSalidaCustom && (tipo === 'pasatarde' || tipo === 'pasadia')) {
    const _toMin = s => { const p = String(s).split(':'); return (parseInt(p[0],10)||0)*60 + (parseInt(p[1],10)||0); };
    const entradaRef = horaEntradaCustom || (tipo === 'pasatarde' ? '12:30' : '09:00');
    if (_toMin(horaSalidaCustom) <= _toMin(entradaRef)) horaSalidaCustom = '';
  }
  if (horaEntradaCustom) checkinHora  = 'a partir de las ' + _formatHora12(horaEntradaCustom);
  if (horaSalidaCustom)  checkoutHora = 'antes de las ' + _formatHora12(horaSalidaCustom);
  return {
    tipo,
    displayCheckin,
    displayCheckout,
    checkinFmt:  formatDateES(displayCheckin),
    checkoutFmt: formatDateES(displayCheckout),
    checkinHora,
    checkoutHora,
    estanciaLabel,
    estanciaValue,
    checkoutExtendido: isExtended,
    horaEntradaCustom,
    horaSalidaCustom,
    isPasadia: tipo === 'pasatarde' || tipo === 'pasadia'
  };
}

function toCalDate(dateStr) {
  return dateStr.toString().slice(0,10).replace(/-/g, '');
}

function toCalDateTime(dateStr, hour, minute) {
  const d = new Date(dateStr.toString().slice(0,10) + 'T' + hour + ':' + minute + ':00-05:00');
  return Utilities.formatDate(d, 'UTC', "yyyyMMdd'T'HHmmss'Z'");
}

function toGCalDateTime(dateStr, hour, minute) {
  return toCalDate(dateStr) + 'T' + hour + minute + '00';
}

function googleCalLink(reservation) {
  const cabin   = CABIN_NAMES_EMAIL[reservation.cabin] || reservation.cabin;
  const title   = encodeURIComponent('Las Nubes — ' + cabin);
  const details = encodeURIComponent('Reserva en Las Nubes\nCabaña: ' + cabin + '\nPersonas: ' + reservation.persons + '\nCheck-in: 2:00 pm | Check-out: 11:00 am\nWhatsApp: +507 6981-2266');
  const loc     = encodeURIComponent('Buenos Aires, Chame, Panamá Oeste');
  const start   = toGCalDateTime(reservation.checkin,  '14', '00');
  const end     = toGCalDateTime(reservation.checkout, '11', '00');
  return 'https://calendar.google.com/calendar/render?action=TEMPLATE&text=' + title + '&dates=' + start + '/' + end + '&details=' + details + '&location=' + loc;
}

function icsContent(reservation) {
  const cabin = CABIN_NAMES_EMAIL[reservation.cabin] || reservation.cabin;
  const start = toCalDateTime(reservation.checkin,  '14', '00');
  const end   = toCalDateTime(reservation.checkout, '11', '00');
  const now   = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  return [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Las Nubes//Reserva//ES',
    'BEGIN:VEVENT',
    'DTSTART:' + start,
    'DTEND:' + end,
    'SUMMARY:Las Nubes — ' + cabin,
    'DESCRIPTION:Cabaña: ' + cabin + '\\nPersonas: ' + reservation.persons + '\\nCheck-in: 2:00 pm | Check-out: 11:00 am\\nContacto: +507 6981-2266',
    'LOCATION:Buenos Aires\\, Chame\\, Panamá Oeste',
    'DTSTAMP:' + now,
    'UID:lasnubes-' + reservation.id + '@lasnubes',
    'END:VEVENT',
    'END:VCALENDAR'
  ].join('\r\n');
}

function buildGuiaHTML(cabin, tipo, checkoutExtendido, horaSalidaCustom) {
  // Fuente unica: getCabinGuideSteps() en PublicLink.gs
  var list = getCabinGuideSteps(cabin, tipo, !!checkoutExtendido, horaSalidaCustom || '');
  var rows = '';
  for (var i = 0; i < list.length; i++) {
    var s = list[i];
    var isLast = (i === list.length - 1);
    var border = isLast ? '' : 'border-bottom:1px solid #f0ede8;';
    rows +=
      '<tr><td style="padding:12px 0;' + border + 'vertical-align:top;">' +
        '<table cellpadding="0" cellspacing="0" width="100%"><tr>' +
          '<td style="width:28px;font-size:18px;vertical-align:top;padding-top:1px;">' + s.icon + '</td>' +
          '<td>' +
            '<p style="margin:0 0 2px;font-size:13px;font-weight:700;color:#3a3530;">' + s.title + '</p>' +
            '<p style="margin:0;font-size:13px;color:#6b6560;line-height:1.6;">' + s.body + '</p>' +
          '</td>' +
        '</tr></table>' +
      '</td></tr>';
  }
  return '' +
    '<h2 style="margin:0 0 6px;font-size:17px;font-weight:600;color:#3a3530;">&#128273; Guía de acceso a tu cabaña</h2>' +
    '<p style="margin:0 0 16px;font-size:13px;color:#8a8078;">Todo lo que necesitas para entrar y disfrutar desde el primer momento.</p>' +
    '<table width="100%" cellpadding="0" cellspacing="0" style="background:#f9f8f6;border:1px solid #e8e4de;border-radius:12px;margin-bottom:24px;"><tr><td style="padding:8px 20px;">' +
    '<table width="100%" cellpadding="0" cellspacing="0">' + rows + '</table>' +
    '</td></tr></table>';
}

function buildEmailHTMLAbierta(r) {
  const name = (r.name || '').toString();
  const waUrl = 'https://wa.me/50769812266?text=' + encodeURIComponent('Hola! Quisiera hacer efectiva mi reserva Abierta a nombre de ' + name + '.');
  const publicLink = getPublicReservaUrlSafe(r.id);
  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>' +
'<body style="margin:0;padding:0;background:#f5f3f0;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;">' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f3f0;padding:32px 16px;"><tr><td align="center">' +
'<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;">' +
'<tr><td style="background:#6a9e62;border-radius:16px 16px 0 0;padding:36px 40px;text-align:center;">' +
'<p style="margin:0 0 6px;font-size:12px;color:rgba(255,255,255,0.75);letter-spacing:2px;text-transform:uppercase;">Reserva Abierta</p>' +
'<h1 style="margin:0;font-size:32px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></h1>' +
'<p style="margin:10px 0 0;font-size:15px;color:rgba(255,255,255,0.9);">Buenos Aires, Chame · Panamá Oeste</p>' +
'</td></tr>' +
'<tr><td style="background:#ffffff;padding:36px 40px;">' +
'<p style="margin:0 0 18px;font-size:16px;color:#3a3530;line-height:1.6;">Hola <strong>' + name + '</strong>, tu reserva quedó registrada en concepto <strong>Abierta</strong>, sin fecha confirmada todavía.</p>' +
(publicLink ? '<p style="margin:0 0 18px;font-size:13px;color:#6b6560;">&#128279; <a href="' + publicLink + '" target="_blank" style="color:#6a9e62;text-decoration:none;font-weight:500;border-bottom:1px solid #6a9e62;">Ver detalles de tu reserva</a></p>' : '') +
'<p style="margin:0 0 24px;font-size:14px;color:#6b6560;line-height:1.7;">Cuando tengas las fechas listas, escríbenos por WhatsApp y la hacemos efectiva. Coordinamos disponibilidad, cabaña y detalles del check-in.</p>' +
'<table cellpadding="0" cellspacing="0" style="margin-bottom:28px;"><tr><td>' +
'<a href="' + waUrl + '" target="_blank" style="display:inline-block;background:#25d366;color:#ffffff;font-size:14px;font-weight:600;padding:12px 24px;border-radius:8px;text-decoration:none;">&#128172; Escribir por WhatsApp</a>' +
'</td></tr></table>' +
'<p style="margin:0 0 8px;font-size:13px;color:#6b6560;line-height:1.6;">O al número: <strong>+507 6981-2266</strong></p>' +
'<hr style="border:none;border-top:1px solid #e8e4de;margin:28px 0;">' +
'<p style="margin:0 0 6px;font-size:13px;color:#6b6560;line-height:1.6;">Mientras tanto, te invitamos a conocer las cabañas y la experiencia en nuestro sitio:</p>' +
'<a href="https://lasnubes.cloud" target="_blank" style="color:#6a9e62;font-size:13px;font-weight:500;text-decoration:none;border-bottom:1px solid #6a9e62;">lasnubes.cloud</a>' +
'</td></tr>' +
'<tr><td style="background:#3a3530;border-radius:0 0 16px 16px;padding:24px 40px;text-align:center;">' +
'<p style="margin:0 0 8px;font-size:18px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></p>' +
'<p style="margin:0 0 12px;font-size:12px;color:rgba(255,255,255,0.6);">Buenos Aires, Chame · En las faldas de Chicá · Panamá Oeste</p>' +
'<a href="https://wa.me/50769812266" style="color:rgba(255,255,255,0.8);font-size:13px;text-decoration:none;">&#128172; WhatsApp: +507 6981-2266</a>' +
'</td></tr></table></td></tr></table></body></html>';
}

// Número de contacto y cierre que se repiten en el certificado de regalo y en
// la plantilla de WhatsApp. Fuente única para no desincronizar los canales.
const REGALO_WA_NUMERO = '+507 6981-2266';
// El cierre se arma por canal y SIN el emoji: en el htmlBody de un email los
// emoji crudos salen como "??????" en Gmail (el resto de las plantillas de este
// archivo ya usa entidades numéricas por eso mismo). En WhatsApp, en cambio, el
// carácter va tal cual. Ver _regaloCoordinarHTML() y REGALO_WA_COORDINAR_CORTO.
const REGALO_COORDINAR_TXT = 'Para verificar disponibilidad y coordinar las fechas de tu reserva escríbenos al WhatsApp ' + REGALO_WA_NUMERO + '. Estaremos atentos';
// Con fechas ya puestas no tiene sentido pedirle "coordinar las fechas", así que
// la invitación se reformula sin perder el cierre.
const REGALO_DUDAS_TXT = 'Si necesitas ajustar las fechas o tienes cualquier duda, escríbenos al WhatsApp ' + REGALO_WA_NUMERO + '. Estaremos atentos';

function _regaloCoordinarHTML(sinFecha) {
  const txt = sinFecha ? REGALO_COORDINAR_TXT : REGALO_DUDAS_TXT;
  // &#128591; = 🙏
  return '<p style="margin:0 0 16px;font-size:13px;color:#6b6560;line-height:1.7;">' + txt + ' &#128591;</p>';
}

// Escapa texto que escribió el admin antes de meterlo en el HTML de un email
// (dedicatoria, nombres). Sin esto un "<" en la dedicatoria rompe el layout.
function _escHtmlEmail(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Convierte los caracteres fuera del BMP (emoji) a entidades numéricas HTML.
// Gmail muestra el emoji crudo del htmlBody como "??????", así que se pasa todo
// el HTML por acá antes de enviarlo — cubre también los emoji que el admin
// escriba en la dedicatoria. Los acentos (BMP) no se tocan.
function _emojiAEntidades(html) {
  return String(html == null ? '' : html)
    .replace(/[\u{10000}-\u{10FFFF}]/gu, ch => '&#' + ch.codePointAt(0) + ';');
}

// Email de CERTIFICADO DE REGALO para el beneficiario.
//
// Deliberadamente NO lleva ningún monto: ni tarifa, ni abono, ni saldo, ni
// recibo adjunto (eso lo bloquea sendConfirmationEmail). Quien recibe un regalo
// no debe enterarse de lo que costó.
//
// Tampoco repite el manual completo de la cabaña: es un certificado, no la
// confirmación operativa. Cuando el beneficiario define fechas, el admin manda
// la confirmación normal (que sí trae guía de acceso, tiendita, etc.).
function buildEmailHTMLRegalo(r) {
  const g       = _parseRegalo(r);
  const cabin   = CABIN_NAMES_EMAIL[r.cabin] || r.cabin || '';
  const color   = CABIN_COLORS_EMAIL[r.cabin] || '#6a9e62';
  const oro     = '#b5893f';
  const nombre  = (r.name || '').toString().trim();
  const primer  = nombre.split(/\s+/)[0] || nombre;
  // Sin fecha definida: el regalo se emite y el beneficiario coordina después.
  const sinFecha = (r.origin === 'Abierta') || !r.checkin || !r.checkout;
  const meta     = sinFecha ? null : tipoEmailMeta(r);
  const personas = parseInt(r.persons, 10) || 0;

  const waRedimir = 'https://wa.me/50769812266?text=' + encodeURIComponent(
    'Hola! Tengo un certificado de regalo de Las Nubes a nombre de ' + nombre + ' y quisiera coordinar las fechas.');
  const publicLink = getPublicReservaUrlSafe(r.id);

  // Qué incluye el regalo — mismas filas que la confirmación pero SIN "Total".
  const filaEstancia = meta
    ? '<tr>' +
      '<td style="padding:20px 24px;border-top:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Check-in</p><p style="margin:0;font-size:15px;color:#3a3530;font-weight:500;">' + meta.checkinFmt + '</p><p style="margin:4px 0 0;font-size:12px;color:#8a8078;">' + meta.checkinHora + '</p></td>' +
      '<td style="padding:20px 24px;border-top:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Check-out</p><p style="margin:0;font-size:15px;color:#3a3530;font-weight:500;">' + meta.checkoutFmt + '</p><p style="margin:4px 0 0;font-size:12px;color:#8a8078;">' + meta.checkoutHora + '</p></td>' +
      '</tr><tr>' +
      '<td colspan="2" style="padding:20px 24px;border-top:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">' + meta.estanciaLabel + '</p><p style="margin:0;font-size:17px;font-weight:600;color:#3a3530;">' + meta.estanciaValue + '</p></td>' +
      '</tr>'
    : '<tr><td colspan="2" style="padding:20px 24px;border-top:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Fechas</p><p style="margin:0;font-size:15px;color:#3a3530;font-weight:500;">A coordinar</p><p style="margin:4px 0 0;font-size:12px;color:#8a8078;">Escríbenos cuando tengas tus fechas listas.</p></td></tr>';

  // Todo el HTML pasa por _emojiAEntidades: en Gmail un emoji crudo en el
  // htmlBody sale como "??????".
  return _emojiAEntidades('<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>' +
'<body style="margin:0;padding:0;background:#f5f3f0;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;">' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f3f0;padding:32px 16px;"><tr><td align="center">' +
'<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;">' +
'<tr><td style="background:' + oro + ';border-radius:16px 16px 0 0;padding:36px 40px;text-align:center;">' +
'<p style="margin:0 0 6px;font-size:12px;color:rgba(255,255,255,0.8);letter-spacing:2px;text-transform:uppercase;">&#127873; Certificado de regalo</p>' +
'<h1 style="margin:0;font-size:32px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></h1>' +
'<p style="margin:10px 0 0;font-size:15px;color:rgba(255,255,255,0.9);">Buenos Aires, Chame · Panamá Oeste</p>' +
'</td></tr>' +
'<tr><td style="background:#ffffff;padding:36px 40px;">' +
'<p style="margin:0 0 8px;font-size:16px;color:#3a3530;line-height:1.6;">Hola <strong>' + _escHtmlEmail(primer) + '</strong>,</p>' +
'<p style="margin:0 0 22px;font-size:16px;color:#3a3530;line-height:1.6;">Te regalaron una estadía en Las Nubes. Este correo es tu certificado &mdash; guárdalo.</p>' +
// Tarjeta del regalo: de parte de + dedicatoria
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#fdf8f0;border:1px solid #ecd9b8;border-left:4px solid ' + oro + ';border-radius:10px;margin-bottom:24px;"><tr><td style="padding:20px 24px;">' +
'<p style="margin:0 0 4px;font-size:11px;color:#8a6a2f;text-transform:uppercase;letter-spacing:0.08em;font-weight:600;">De parte de</p>' +
'<p style="margin:0;font-size:19px;font-weight:600;color:#3a3530;font-family:Georgia,serif;">' + _escHtmlEmail(g.de || 'alguien que te quiere') + '</p>' +
(g.mensaje
  ? '<p style="margin:14px 0 0;padding:12px 16px;background:#ffffff;border:1px solid #ecd9b8;border-radius:8px;font-size:14px;color:#6b6560;line-height:1.65;font-style:italic;">&ldquo;' + _escHtmlEmail(g.mensaje) + '&rdquo;</p>'
  : '') +
'</td></tr></table>' +
// Qué incluye
'<p style="margin:0 0 12px;font-size:14px;font-weight:600;color:#3a3530;">Tu regalo incluye</p>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f9f8f6;border-radius:12px;border:1px solid #e8e4de;margin-bottom:24px;">' +
'<tr>' +
'<td style="padding:20px 24px;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Cabaña</p><p style="margin:0;font-size:17px;font-weight:600;color:' + color + ';">' + (cabin || 'Las Nubes') + '</p></td>' +
'<td style="padding:20px 24px;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Personas</p><p style="margin:0;font-size:17px;font-weight:600;color:#3a3530;">' + (personas || '—') + '</p></td>' +
'</tr>' + filaEstancia +
'</table>' +
'<p style="margin:0 0 20px;font-size:13px;color:#6b6560;line-height:1.7;">&#10003; Todo está cubierto por quien te lo regaló. No tienes nada que pagar.</p>' +
// CTA
(sinFecha
  ? '<p style="margin:0 0 12px;font-size:14px;font-weight:600;color:#3a3530;">Cómo hacerlo efectivo</p>' +
    _regaloCoordinarHTML(true) +
    '<table cellpadding="0" cellspacing="0" style="margin-bottom:24px;"><tr><td>' +
    '<a href="' + waRedimir + '" target="_blank" style="display:inline-block;background:#25d366;color:#ffffff;font-size:14px;font-weight:600;padding:12px 24px;border-radius:8px;text-decoration:none;">&#128172; Coordinar mis fechas</a>' +
    '</td></tr></table>'
  : (publicLink
      ? '<table cellpadding="0" cellspacing="0" style="margin-bottom:20px;"><tr><td>' +
        '<a href="' + publicLink + '" target="_blank" style="display:inline-block;background:' + color + ';color:#ffffff;font-size:14px;font-weight:600;padding:12px 24px;border-radius:8px;text-decoration:none;">&#128279; Ver los detalles de tu estadía</a>' +
        '</td></tr></table>'
      : '') +
    _regaloCoordinarHTML(false)) +
'<hr style="border:none;border-top:1px solid #e8e4de;margin:28px 0;">' +
// Teaser de la experiencia — sin precios
'<h2 style="margin:0 0 16px;font-size:17px;font-weight:600;color:#3a3530;">Lo que te espera</h2>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:16px;"><tr><td style="vertical-align:top;padding-right:12px;font-size:22px;width:36px;">&#9728;&#65039;</td><td><p style="margin:0 0 3px;font-size:14px;font-weight:600;color:#3a3530;">Cabaña solar, privacidad total</p><p style="margin:0;font-size:13px;color:#6b6560;line-height:1.6;">Energía 100% solar y uso exclusivo de todas las instalaciones para ti y tus acompañantes.</p></td></tr></table>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:16px;"><tr><td style="vertical-align:top;padding-right:12px;font-size:22px;width:36px;">&#127859;</td><td><p style="margin:0 0 3px;font-size:14px;font-weight:600;color:#3a3530;">Cocina equipada y área de BBQ</p><p style="margin:0;font-size:13px;color:#6b6560;line-height:1.6;">Con café, azúcar, especias básicas y cooler grande. Solo traes hielo y tus alimentos.</p></td></tr></table>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:20px;"><tr><td style="vertical-align:top;padding-right:12px;font-size:22px;width:36px;">&#127956;&#65039;</td><td><p style="margin:0 0 3px;font-size:14px;font-weight:600;color:#3a3530;">Montaña, cascadas y playa cerca</p><p style="margin:0;font-size:13px;color:#6b6560;line-height:1.6;">Cascada dentro del proyecto, Los Cajones de Chame, Cerro Campana y las playas de Coronado a 15–20 minutos.</p></td></tr></table>' +
'<a href="https://lasnubes.cloud" target="_blank" style="color:' + color + ';font-size:13px;font-weight:500;text-decoration:none;border-bottom:1px solid ' + color + ';">Conocer Las Nubes &rarr;</a>' +
'</td></tr>' +
'<tr><td style="background:#3a3530;border-radius:0 0 16px 16px;padding:24px 40px;text-align:center;">' +
'<p style="margin:0 0 8px;font-size:18px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></p>' +
'<p style="margin:0 0 12px;font-size:12px;color:rgba(255,255,255,0.6);">Buenos Aires, Chame · En las faldas de Chicá · Panamá Oeste</p>' +
'<a href="https://wa.me/50769812266" style="color:rgba(255,255,255,0.8);font-size:13px;text-decoration:none;">&#128172; WhatsApp: +507 6981-2266</a>' +
'</td></tr></table></td></tr></table></body></html>');
}

// Banner para reservas multi-cabaña: se inserta arriba del bloque de detalles
// cuando r.multiCabin trae un array con las N estadías hermanas.
function _multiCabinBannerHTML(r) {
  if (!r || !r.multiCabin || !Array.isArray(r.multiCabin) || r.multiCabin.length < 2) return '';
  const total = parseFloat(r.amount) || 0;
  let rows = '';
  r.multiCabin.forEach(function(s, i) {
    const color = CABIN_COLORS_EMAIL[s.cabin] || '#6a9e62';
    const ciFmt = formatDateES(s.checkin);
    const coFmt = formatDateES(s.checkout);
    const nightsLbl = (s.nights === 1 ? '1 noche' : s.nights + ' noches');
    const amtLbl = s.amount ? ' · $' + parseFloat(s.amount).toFixed(2) : '';
    rows += '<tr>'
         +  '<td style="padding:8px 0;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:0.06em;width:36px;vertical-align:top;">' + (i+1) + '</td>'
         +  '<td style="padding:8px 0;font-size:14px;color:#3a3530;vertical-align:top;">'
         +    '<span style="display:inline-block;width:8px;height:8px;background:' + color + ';border-radius:50%;margin-right:8px;vertical-align:middle;"></span>'
         +    '<strong style="color:' + color + ';">' + (s.cabinName || s.cabin) + '</strong>'
         +    '<span style="color:#8a8078;font-size:12px;margin-left:8px;">' + nightsLbl + amtLbl + '</span>'
         +    '<br><span style="font-size:12px;color:#6b6560;padding-left:18px;">' + ciFmt + ' → ' + coFmt + '</span>'
         +  '</td>'
         +  '</tr>';
  });
  return '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fef9f0;border:1px solid #e8dfc9;border-left:4px solid #d97706;border-radius:10px;margin-bottom:20px;">'
       +   '<tr><td style="padding:18px 22px;">'
       +     '<p style="margin:0 0 4px;font-size:12px;color:#8a6000;text-transform:uppercase;letter-spacing:0.08em;font-weight:600;">Estadía multi-cabaña</p>'
       +     '<p style="margin:0 0 12px;font-size:14px;color:#3a3530;line-height:1.5;">Reservaste <strong>' + r.multiCabin.length + ' cabañas</strong> consecutivas con un pago único de <strong>$' + total.toFixed(2) + '</strong>. Cambiarás de cabaña según este itinerario:</p>'
       +     '<table width="100%" cellpadding="0" cellspacing="0">' + rows + '</table>'
       +     '<p style="margin:12px 0 0;font-size:12px;color:#6b6560;line-height:1.5;">El día del cambio te esperamos con la siguiente cabaña lista. Si tienes dudas, escríbenos al WhatsApp.</p>'
       +   '</td></tr>'
       + '</table>';
}

function buildEmailHTML(r) {
  // El regalo se evalúa ANTES de Abierta: un regalo sin fecha debe recibir el
  // certificado, no el email genérico de reserva abierta.
  if (esReservaRegalo(r)) return buildEmailHTMLRegalo(r);
  if (r.origin === 'Abierta') return buildEmailHTMLAbierta(r);
  const cabin       = CABIN_NAMES_EMAIL[r.cabin] || r.cabin;
  const color       = CABIN_COLORS_EMAIL[r.cabin] || '#6a9e62';
  const meta        = tipoEmailMeta(r);
  const checkinFmt  = meta.checkinFmt;
  const checkoutFmt = meta.checkoutFmt;
  const gcalUrl     = googleCalLink(r);
  const amount      = parseFloat(r.amount)  || 0;
  const deposit     = parseFloat(r.deposit) || 0;
  const saldo       = (amount - deposit).toFixed(2);
  const hasSaldo    = amount > 0 && parseFloat(saldo) > 0;
  const ics         = icsContent(r);
  const icsB64      = Utilities.base64Encode(ics);
  const icsUri      = 'data:text/calendar;base64,' + icsB64;
  const pagarUrl    = 'https://wa.me/50769812266?text=' + encodeURIComponent('Deseo cancelar el saldo restante de mi reserva del día ' + meta.checkinFmt + ' en la cabaña ' + cabin + '. ¿Me comparte los métodos de pago?');
  const publicLink = getPublicReservaUrlSafe(r.id);

  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>' +
'<body style="margin:0;padding:0;background:#f5f3f0;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;">' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f3f0;padding:32px 16px;"><tr><td align="center">' +
'<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;">' +
'<tr><td style="background:' + color + ';border-radius:16px 16px 0 0;padding:36px 40px;text-align:center;">' +
'<p style="margin:0 0 6px;font-size:12px;color:rgba(255,255,255,0.75);letter-spacing:2px;text-transform:uppercase;">Confirmación de reserva</p>' +
'<h1 style="margin:0;font-size:32px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></h1>' +
'<p style="margin:10px 0 0;font-size:15px;color:rgba(255,255,255,0.9);">Buenos Aires, Chame · Panamá Oeste</p>' +
'</td></tr>' +
'<tr><td style="background:#ffffff;padding:36px 40px;">' +
'<p style="margin:0 0 18px;font-size:16px;color:#3a3530;line-height:1.6;">Hola <strong>' + r.name + '</strong>, tu reserva ha sido confirmada. ¡Te esperamos en Las Nubes!</p>' +
_multiCabinBannerHTML(r) +
(r.origin === 'Referido' ? '<p style="margin:0 0 18px;font-size:13px;color:#385d7a;background:#eef4f8;border-left:3px solid #5a85b0;padding:10px 14px;border-radius:6px;">&#129309; Tarifa pactada con descuento del <strong>Programa Amigos</strong>.</p>' : '') +
(publicLink ? '<p style="margin:0 0 24px;font-size:13px;color:#6b6560;">&#128279; <a href="' + publicLink + '" target="_blank" style="color:' + color + ';text-decoration:none;font-weight:500;border-bottom:1px solid ' + color + ';">Ver detalles online</a> &mdash; este link tambi&eacute;n se puede compartir por WhatsApp.</p>' : '') +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f9f8f6;border-radius:12px;border:1px solid #e8e4de;margin-bottom:24px;">' +
'<tr>' +
'<td style="padding:20px 24px;border-bottom:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Cabaña</p><p style="margin:0;font-size:17px;font-weight:600;color:' + color + ';">' + cabin + '</p></td>' +
'<td style="padding:20px 24px;border-bottom:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Personas</p><p style="margin:0;font-size:17px;font-weight:600;color:#3a3530;">' + r.persons + '</p></td>' +
'</tr><tr>' +
'<td style="padding:20px 24px;border-bottom:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Check-in</p><p style="margin:0;font-size:15px;color:#3a3530;font-weight:500;">' + checkinFmt + '</p><p style="margin:4px 0 0;font-size:12px;color:#8a8078;">' + meta.checkinHora + '</p></td>' +
'<td style="padding:20px 24px;border-bottom:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Check-out</p><p style="margin:0;font-size:15px;color:#3a3530;font-weight:500;">' + checkoutFmt + '</p><p style="margin:4px 0 0;font-size:12px;color:#8a8078;">' + meta.checkoutHora + '</p></td>' +
'</tr><tr>' +
'<td style="padding:20px 24px;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">' + meta.estanciaLabel + '</p><p style="margin:0;font-size:17px;font-weight:600;color:#3a3530;">' + meta.estanciaValue + '</p></td>' +
'<td style="padding:20px 24px;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Total</p><p style="margin:0;font-size:17px;font-weight:600;color:#3a3530;">$' + amount.toFixed(2) + '</p>' + (deposit > 0 ? '<p style="margin:4px 0 0;font-size:12px;color:#8a8078;">Abono: $' + deposit.toFixed(2) + '</p>' : '') + '</td>' +
'</tr>' + (hasSaldo ? '<tr><td colspan="2" style="padding:18px 24px;background:#fff8e1;border-top:1px solid #e8e4de;border-radius:0 0 12px 12px;"><p style="margin:0 0 4px;font-size:13px;color:#8a6000;">&#9888; <strong>Saldo pendiente: $' + saldo + '</strong></p><p style="margin:0 0 12px;font-size:12px;color:#8a6000;line-height:1.5;">Te pedimos cancelar el saldo antes del día de tu reserva.</p><a href="' + pagarUrl + '" target="_blank" style="display:inline-block;background:#25d366;color:#ffffff;font-size:13px;font-weight:600;padding:9px 20px;border-radius:8px;text-decoration:none;">&#128172; Pagar saldo</a></td></tr>' : '') +
'</table>' +
'<p style="margin:0 0 12px;font-size:14px;font-weight:600;color:#3a3530;">Agregar a tu calendario</p>' +
'<table cellpadding="0" cellspacing="0" style="margin-bottom:32px;"><tr>' +
'<td style="padding-right:10px;"><a href="' + gcalUrl + '" target="_blank" style="display:inline-block;background:#4285f4;color:#ffffff;font-size:13px;font-weight:500;padding:10px 20px;border-radius:8px;text-decoration:none;">&#128197; Google Calendar</a></td>' +
'<td><a href="' + icsUri + '" style="display:inline-block;background:#3a3530;color:#ffffff;font-size:13px;font-weight:500;padding:10px 20px;border-radius:8px;text-decoration:none;">&#127822; Apple / Outlook</a></td>' +
'</tr></table>' +
'<hr style="border:none;border-top:1px solid #e8e4de;margin:0 0 28px;">' +
'<h2 style="margin:0 0 20px;font-size:17px;font-weight:600;color:#3a3530;">Lo que necesitas saber</h2>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:18px;"><tr><td style="vertical-align:top;padding-right:12px;font-size:22px;width:36px;">&#127859;</td><td><p style="margin:0 0 3px;font-size:14px;font-weight:600;color:#3a3530;">Cocina y alimentación</p><p style="margin:0;font-size:13px;color:#6b6560;line-height:1.6;">La cabaña cuenta con cocina completamente equipada y área de BBQ. Incluye café, azúcar, especias básicas y un cooler grande (no contamos con nevera). Solo trae hielo y tus alimentos.</p></td></tr></table>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:18px;"><tr><td style="vertical-align:top;padding-right:12px;font-size:22px;width:36px;">&#9728;&#65039;</td><td><p style="margin:0 0 3px;font-size:14px;font-weight:600;color:#3a3530;">Electricidad solar</p><p style="margin:0;font-size:13px;color:#6b6560;line-height:1.6;">La cabaña no cuenta con luz eléctrica convencional. Está completamente iluminada a través de paneles solares. Contamos con inversor para cargar celulares y la señal de las telefónicas es excelente.</p></td></tr></table>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:18px;"><tr><td style="vertical-align:top;padding-right:12px;font-size:22px;width:36px;">&#128703;</td><td><p style="margin:0 0 3px;font-size:14px;font-weight:600;color:#3a3530;">Baño</p><p style="margin:0;font-size:13px;color:#6b6560;line-height:1.6;">El baño cuenta con jabón, papel higiénico y toallas limpias. Fumigamos semanalmente, pero si eres sensible a los mosquitos, te recomendamos traer repelente.</p></td></tr></table>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:18px;"><tr><td style="vertical-align:top;padding-right:12px;font-size:22px;width:36px;">&#128274;</td><td><p style="margin:0 0 3px;font-size:14px;font-weight:600;color:#3a3530;">Privacidad total</p><p style="margin:0;font-size:13px;color:#6b6560;line-height:1.6;">Todas las instalaciones de la cabaña son de uso exclusivo de quienes la reservan.</p></td></tr></table>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:8px;"><tr><td style="vertical-align:top;padding-right:12px;font-size:22px;width:36px;">&#128506;&#65039;</td><td><p style="margin:0 0 3px;font-size:14px;font-weight:600;color:#3a3530;">Cómo llegar</p><p style="margin:0 0 8px;font-size:13px;color:#6b6560;line-height:1.6;">Por carretera interamericana, entrar por el Pío Pío de Bejuco hacia carretera Bejuco–Sorá. Al llegar al pueblo de Buenos Aires, doblar a la derecha hacia el pueblo de Chicá. La cabaña queda a 100 metros.</p><p style="margin:0 0 12px;font-size:13px;color:#6b6560;line-height:1.6;">La manera más fácil es colocar en <strong>Waze: &quot;Aires de Chicá&quot;</strong>. Te llevará directo al portón verde.</p><a href="https://maps.google.com/?q=8.639400,-79.945900" target="_blank" style="display:inline-block;background:#f0ede8;color:#3a3530;font-size:13px;font-weight:500;padding:8px 16px;border-radius:8px;text-decoration:none;border:1px solid #e8e4de;">&#128205; Ver en Google Maps</a></td></tr></table>' +
'<hr style="border:none;border-top:1px solid #e8e4de;margin:24px 0;">' +
buildGuiaHTML(r.cabin, meta.tipo, !!r.checkoutExtendido, meta.horaSalidaCustom) +
'<hr style="border:none;border-top:1px solid #e8e4de;margin:24px 0;">' +
'<h2 style="margin:0 0 6px;font-size:17px;font-weight:600;color:#3a3530;">&#127978; Tiendita Las Nubes</h2>' +
'<p style="margin:0 0 16px;font-size:13px;color:#8a8078;">Tenemos insumos disponibles — te los llevamos directo a la cabaña.</p>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f9f8f6;border:1px solid #e8e4de;border-radius:12px;margin-bottom:16px;"><tr><td style="padding:16px 20px;"><table width="100%" cellpadding="0" cellspacing="0">' +
'<tr style="border-bottom:1px solid #e8e4de;"><td style="padding:8px 0;"><span style="font-size:13px;font-weight:600;color:#3a3530;">&#128293; Kit de Fogata</span><br><span style="font-size:12px;color:#8a8078;">Leña · cerillo · palillos · malvaviscos</span></td><td style="text-align:right;font-size:14px;font-weight:700;color:#2e7d32;vertical-align:middle;">$10.00</td></tr>' +
'<tr style="border-bottom:1px solid #e8e4de;"><td style="padding:8px 0;"><span style="font-size:13px;font-weight:600;color:#3a3530;">&#129704; Bolsa de Carbón</span><br><span style="font-size:12px;color:#8a8078;">Con cerillo especial</span></td><td style="text-align:right;font-size:14px;font-weight:700;color:#2e7d32;vertical-align:middle;">$5.00</td></tr>' +
'<tr style="border-bottom:1px solid #e8e4de;"><td style="padding:8px 0;"><span style="font-size:13px;font-weight:600;color:#3a3530;">&#129432; Repelente OFF Spray</span></td><td style="text-align:right;font-size:14px;font-weight:700;color:#2e7d32;vertical-align:middle;">$8.00</td></tr>' +
'<tr style="border-bottom:1px solid #e8e4de;"><td style="padding:8px 0;"><span style="font-size:13px;font-weight:600;color:#3a3530;">&#129432; Repelente Family Care</span><br><span style="font-size:12px;color:#8a8078;">Toallitas</span></td><td style="text-align:right;font-size:14px;font-weight:700;color:#2e7d32;vertical-align:middle;">$5.00</td></tr>' +
'<tr style="border-bottom:1px solid #e8e4de;"><td style="padding:8px 0;"><span style="font-size:13px;font-weight:600;color:#3a3530;">&#129461; Kit Pasta y Cepillo</span></td><td style="text-align:right;font-size:14px;font-weight:700;color:#2e7d32;vertical-align:middle;">$5.00</td></tr>' +
'<tr><td style="padding:8px 0;"><span style="font-size:13px;font-weight:600;color:#3a3530;">&#127803; Toallas Sanitarias</span></td><td style="text-align:right;font-size:14px;font-weight:700;color:#2e7d32;vertical-align:middle;">$5.00</td></tr>' +
'</table></td></tr><tr><td style="background:#e8f5e9;border-top:1px solid #c8e6c9;border-radius:0 0 12px 12px;padding:12px 20px;"><p style="margin:0;font-size:13px;color:#2e7d32;">&#128156; Paga por <strong>Yappy</strong> al mismo número y te lo llevamos a la cabaña.</p></td></tr></table>' +
'<a href="https://wa.me/50769812266?text=Hola!%20Quisiera%20pedir%20de%20la%20Tiendita%20Las%20Nubes%20%F0%9F%8C%BF" target="_blank" style="display:inline-block;background:#25d366;color:#fff;font-size:13px;font-weight:600;padding:9px 18px;border-radius:8px;text-decoration:none;margin-bottom:24px;">&#128172; Pedir por WhatsApp</a>' +
'<hr style="border:none;border-top:1px solid #e8e4de;margin:0 0 24px;">' +
'<h2 style="margin:0 0 6px;font-size:17px;font-weight:600;color:#3a3530;">&#127752; Insumos y alimentos</h2>' +
'<p style="margin:0 0 16px;font-size:13px;color:#6b6560;line-height:1.6;">A 5 minutos tienes una tienda de conveniencia. En Bejuco el MiniSuper Buenos Precios, y a 20 min en Coronado: El Rey, Machetazo, Riba Smith y Super 99.</p>' +
'<a href="https://lasnubes.cloud/#insumos" target="_blank" style="display:inline-block;background:#5a85b0;color:#ffffff;font-size:13px;font-weight:600;padding:10px 20px;border-radius:8px;text-decoration:none;margin-bottom:24px;">&#128722; Ver guía de insumos</a>' +
'<hr style="border:none;border-top:1px solid #e8e4de;margin:0 0 24px;">' +
'<h2 style="margin:0 0 6px;font-size:17px;font-weight:600;color:#3a3530;">&#127956;&#65039; Actividades y alrededores</h2>' +
'<p style="margin:0 0 16px;font-size:13px;color:#6b6560;line-height:1.6;">Cascada dentro del proyecto (20 min), Los Cajones de Chame (10 min), Cerro Campana (20 min), Playa Gorgona y Coronado (15–20 min) y cascadas de Sorá (20 min).</p>' +
'<a href="https://lasnubes.cloud/#actividades" target="_blank" style="display:inline-block;background:#4a7c5f;color:#ffffff;font-size:13px;font-weight:600;padding:10px 20px;border-radius:8px;text-decoration:none;">&#127758; Ver guía de actividades</a>' +

// Gastronomía existía solo en el menú del bot: quien nunca le escribiera no se
// enteraba de que hay una guía de restaurantes.
'<hr style="border:none;border-top:1px solid #e8e4de;margin:24px 0;">' +
'<h2 style="margin:0 0 6px;font-size:17px;font-weight:600;color:#3a3530;">&#127869;&#65039; Dónde comer</h2>' +
'<p style="margin:0 0 16px;font-size:13px;color:#6b6560;line-height:1.6;">Restaurantes y fondas cerca, con direcciones, horarios y fotos.</p>' +
'<a href="https://lasnubes.cloud/#gastronomia" target="_blank" style="display:inline-block;background:#a0673f;color:#ffffff;font-size:13px;font-weight:600;padding:10px 20px;border-radius:8px;text-decoration:none;">&#127860; Ver guía de gastronomía</a>' +
'</td></tr>' +
'<tr><td style="background:#3a3530;border-radius:0 0 16px 16px;padding:24px 40px;text-align:center;">' +
'<p style="margin:0 0 8px;font-size:18px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></p>' +
'<p style="margin:0 0 12px;font-size:12px;color:rgba(255,255,255,0.6);">Buenos Aires, Chame · En las faldas de Chicá · Panamá Oeste</p>' +
'<a href="https://wa.me/50769812266" style="color:rgba(255,255,255,0.8);font-size:13px;text-decoration:none;">&#128172; WhatsApp: +507 6981-2266</a>' +
'</td></tr></table></td></tr></table></body></html>';
}

function sendConfirmationEmail(reservation, voucherBase64, voucherMimeType, subjectPrefix) {
  try {
    const email = reservation.email;
    if (!email) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'No hay email de huésped' })).setMimeType(ContentService.MimeType.JSON);
    const cabin   = CABIN_NAMES_EMAIL[reservation.cabin] || reservation.cabin;
    const isAbierta = reservation.origin === 'Abierta';
    const isRegalo  = esReservaRegalo(reservation);
    const subject = (subjectPrefix || '') + (isRegalo
      ? '🎁 Tienes un regalo en Las Nubes'
      : (isAbierta
        ? '📌 Reserva Abierta — Las Nubes'
        : '✅ Confirmación de reserva — ' + cabin + ' · Las Nubes'));

    // Adjuntar recibo PDF si la reserva tiene voucher (codTransferencia o montoVoucher).
    // En un REGALO no se adjunta nada: ni recibo ni voucher — el beneficiario no
    // debe ver cuánto pagó quien se lo regaló.
    const attachments = [];
    const montoVoucherNum = reservation.montoVoucher ? parseFloat(String(reservation.montoVoucher).replace(/[^\d.]/g, '')) || 0 : 0;
    const hasVoucher = !isRegalo && !!(reservation.codTransferencia || montoVoucherNum > 0);
    Logger.log('🧾 sendConfirmationEmail · hasVoucher=' + hasVoucher + ' codT=' + (reservation.codTransferencia || 'null') + ' montoVoucher=' + (reservation.montoVoucher || 'null') + ' voucherBytes=' + (voucherBase64 ? voucherBase64.length : 0));
    const debug = { hasVoucher, codT: reservation.codTransferencia || null, montoVoucher: reservation.montoVoucher || null };
    if (hasVoucher && !isAbierta) {
      try {
        const receipt = generateReceiptPDF(reservation);
        attachments.push(receipt.blob);
        Logger.log('📄 Recibo ' + receipt.number + ' adjuntado (' + (receipt.blob.getBytes().length) + ' bytes)');
        debug.receipt = { number: receipt.number, bytes: receipt.blob.getBytes().length };
      } catch(rcpErr) {
        Logger.log('⚠ No se pudo generar recibo PDF: ' + rcpErr + '\n' + (rcpErr && rcpErr.stack ? rcpErr.stack : ''));
        debug.receiptError = rcpErr.toString();
      }
    }

    // Adjuntar voucher original (imagen) si el dashboard lo envió
    if (voucherBase64 && voucherMimeType && !isRegalo) {
      try {
        const ext = voucherMimeType.includes('png') ? '.png' : voucherMimeType.includes('gif') ? '.gif' : voucherMimeType.includes('webp') ? '.webp' : '.jpg';
        const safeName = ((reservation.name || 'huesped').toString())
          .replace(/\s+/g, '_').replace(/[^A-Za-z0-9_]/g, '').slice(0, 40) || 'huesped';
        const voucherBlob = Utilities.newBlob(Utilities.base64Decode(voucherBase64), voucherMimeType, 'Voucher_' + safeName + ext);
        attachments.push(voucherBlob);
        Logger.log('🧾 Voucher imagen adjuntada (' + voucherBlob.getBytes().length + ' bytes)');
        debug.voucherAttached = voucherBlob.getBytes().length;
      } catch(vErr) {
        Logger.log('⚠ No se pudo adjuntar voucher imagen: ' + vErr);
        debug.voucherAttachError = vErr.toString();
      }
    }

    GmailApp.sendEmail(email, subject, '', {
      htmlBody:    buildEmailHTML(reservation),
      attachments: attachments,
      name:        'Las Nubes',
      replyTo:     REPLY_TO_EMAIL
    });
    Logger.log('📧 Confirmación enviada a: ' + email);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, status: 'email_sent', _debug: debug, version: 'recibo-v2' })).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: e.toString(), stack: e.stack })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ═══════════════════════════════════════════════════════════
//  Recibos PDF
// ═══════════════════════════════════════════════════════════

const RECIBOS_FOLDER_NAME = 'Recibos Las Nubes';
const RECIBO_LOGO_URL     = 'https://lasnubes.cloud/logo-black.png';

// Ejecutar UNA VEZ desde el editor para retro-poblar la columna VoucherURL
// usando el codTransferencia (description) y el nombre (filename) como criterio.
// Solo actualiza filas con codTransferencia y VoucherURL vacío.
function repoblarVoucherURLs() {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const folders = DriveApp.getFoldersByName(VOUCHER_FOLDER_NAME);
  if (!folders.hasNext()) {
    Logger.log('⚠ Carpeta "' + VOUCHER_FOLDER_NAME + '" no existe.');
    return;
  }
  const folder = folders.next();

  // Pre-cargar archivos en memoria (una sola pasada por la carpeta)
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    files.push({ name: f.getName().toLowerCase(), desc: f.getDescription() || '', url: f.getUrl() });
  }
  Logger.log('📁 ' + files.length + ' archivos en ' + VOUCHER_FOLDER_NAME);

  let updated = 0, skipped = 0, notFound = 0;
  for (let i = 1; i < data.length; i++) {
    const nombre      = data[i][1]  ? data[i][1].toString()              : '';
    const codTransf   = data[i][18] ? data[i][18].toString().trim()      : '';
    const existingURL = data[i][25] ? data[i][25].toString().trim()      : '';
    if (existingURL) { skipped++; continue; }
    if (!codTransf && !nombre) continue;

    let match = null;
    if (codTransf) {
      match = files.find(f => f.desc.includes(codTransf));
    }
    if (!match && nombre) {
      const nombreSafe = nombre.replace(/\s+/g, '_').slice(0, 20).toLowerCase();
      if (nombreSafe) match = files.find(f => f.name.includes(nombreSafe));
    }

    if (match) {
      sheet.getRange(i + 1, 26).setValue(match.url);
      updated++;
    } else if (codTransf) {
      notFound++;
    }
  }
  Logger.log('✓ Migración VoucherURL: ' + updated + ' actualizados, ' + skipped + ' ya tenían URL, ' + notFound + ' con codTransf sin match.');
}

// Ejecutar UNA VEZ desde el editor para otorgar los permisos necesarios
// (DocumentApp, DriveApp, UrlFetchApp, LockService) que usa generateReceiptPDF.
// Después de aceptar la consola OAuth, el web app podrá generar recibos PDF.
function autorizarPermisosRecibo() {
  PropertiesService.getScriptProperties().getProperty('RECEIPT_COUNTER');
  const lock = LockService.getScriptLock();
  try { lock.tryLock(100); lock.releaseLock(); } catch(_) {}
  const doc = DocumentApp.create('test-permisos-recibo-' + Date.now());
  DocumentApp.openById(doc.getId()); // abrir explicito
  DriveApp.getFileById(doc.getId()).setTrashed(true);
  try { UrlFetchApp.fetch(RECIBO_LOGO_URL).getBlob(); } catch(_) {}
  Logger.log('✓ Permisos OK. Recibos PDF deberían funcionar a partir de ahora.');
}

function nextReceiptNumber() {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const cur = parseInt(props.getProperty('RECEIPT_COUNTER') || '0', 10);
    const next = cur + 1;
    props.setProperty('RECEIPT_COUNTER', String(next));
    return next;
  } finally {
    try { lock.releaseLock(); } catch(_) {}
  }
}

function getOrCreateRecibosFolder() {
  const folders = DriveApp.getFoldersByName(RECIBOS_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(RECIBOS_FOLDER_NAME);
}

// Alinear horizontalmente todos los párrafos dentro de una celda de tabla.
function _alignCell(cell, alignment) {
  for (let i = 0; i < cell.getNumChildren(); i++) {
    const ch = cell.getChild(i);
    if (ch.getType && ch.getType() === DocumentApp.ElementType.PARAGRAPH) {
      ch.asParagraph().setAlignment(alignment);
    }
  }
}

function generateReceiptPDF(r) {
  const numStr   = 'LN-' + String(nextReceiptNumber()).padStart(4, '0');
  const meta     = tipoEmailMeta(r);
  const cabin    = CABIN_NAMES_EMAIL[r.cabin] || r.cabin;
  const today    = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');
  const todayFmt = formatDateES(today);

  // Texto de estancia según tipo
  let estanciaText;
  if (meta.tipo === 'pasadia') {
    estanciaText = meta.checkinFmt + ' · Pasadía 9am – 5pm';
  } else if (meta.tipo === 'pasatarde') {
    estanciaText = meta.checkinFmt + ' · Pasatarde 12:30pm – 7pm';
  } else if (meta.tipo === 'early') {
    estanciaText = meta.checkinFmt + ' → ' + meta.checkoutFmt + ' · 1 noche (entra 9am)';
  } else if (meta.tipo === 'late') {
    estanciaText = meta.checkinFmt + ' → ' + meta.checkoutFmt + ' · 1 noche (sale 4pm)';
  } else {
    const n = meta.estanciaValue;
    estanciaText = meta.checkinFmt + ' → ' + meta.checkoutFmt + ' · ' + n + (n === 1 ? ' noche' : ' noches');
  }

  const amount       = parseFloat(r.amount)  || 0;
  const deposit      = parseFloat(r.deposit) || 0;
  const monto        = deposit > 0 ? deposit : amount;
  const saldo        = (amount - deposit).toFixed(2);
  const ref          = (r.codTransferencia || '').toString().trim();
  let metodo         = 'Otro';
  if (r.origin === 'Airbnb') metodo = 'Airbnb';
  else if (ref)              metodo = 'Yappy / Pago digital';
  else if (r.origin === 'Cortesia' || r.origin === 'Colaboracion' || r.origin === 'Personal' || r.origin === 'Mantenimiento') metodo = 'Sin cobro';

  // Crear Doc temporal
  const doc  = DocumentApp.create('Recibo ' + numStr + ' (temp)');
  const body = doc.getBody();
  body.setMarginTop(50).setMarginBottom(50).setMarginLeft(60).setMarginRight(60);

  // Logo (best-effort: si falla, sigue sin logo)
  try {
    const logoBlob = UrlFetchApp.fetch(RECIBO_LOGO_URL).getBlob();
    const logoP    = body.appendParagraph('');
    logoP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    const img = logoP.appendInlineImage(logoBlob);
    img.setWidth(90).setHeight(90);
  } catch(e) {
    Logger.log('Logo recibo error: ' + e);
  }

  // Encabezado
  body.appendParagraph('Las Nubes')
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setFontSize(26).setBold(true).setForegroundColor('#3a3530');
  body.appendParagraph('Buenos Aires, Chame · Panamá Oeste')
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setFontSize(10).setForegroundColor('#8a8078');
  body.appendParagraph(' ').editAsText().setFontSize(6); // espaciador
  body.appendHorizontalRule();

  // Cabecera del recibo (label izq, número/fecha der)
  const headerTbl = body.appendTable([['RECIBO DE PAGO', 'N°  ' + numStr]]);
  headerTbl.setBorderWidth(0);
  headerTbl.getCell(0,0).editAsText().setFontSize(11).setBold(true).setForegroundColor('#3a3530');
  _alignCell(headerTbl.getCell(0,1), DocumentApp.HorizontalAlignment.RIGHT);
  headerTbl.getCell(0,1).editAsText().setFontSize(11).setBold(true).setForegroundColor('#3a3530');
  const dateP = body.appendParagraph('Emitido ' + todayFmt);
  dateP.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  dateP.editAsText().setFontSize(9).setBold(false).setForegroundColor('#8a8078');

  body.appendHorizontalRule();

  // Recibido de
  body.appendParagraph('RECIBIDO DE').editAsText().setFontSize(9).setBold(false).setForegroundColor('#8a8078');
  body.appendParagraph(r.name || '—').editAsText().setFontSize(15).setBold(true).setForegroundColor('#3a3530');

  // Por concepto de
  body.appendParagraph(' ').editAsText().setFontSize(4);
  body.appendParagraph('POR CONCEPTO DE').editAsText().setFontSize(9).setBold(false).setForegroundColor('#8a8078');
  body.appendParagraph('Reserva — ' + cabin).editAsText().setFontSize(12).setBold(true).setForegroundColor('#3a3530');
  body.appendParagraph(estanciaText).editAsText().setFontSize(11).setBold(false).setForegroundColor('#6b6560');

  body.appendHorizontalRule();

  // Detalle de pago (tabla 2 cols)
  const payRows = [
    ['Monto recibido',  '$ ' + monto.toFixed(2) + ' USD'],
    ['Método de pago',  metodo]
  ];
  if (ref) payRows.push(['Referencia', ref]);
  payRows.push(['Saldo pendiente', '$ ' + saldo + (parseFloat(saldo) > 0 ? '' : '  (saldo cero)')]);

  const payTbl = body.appendTable(payRows);
  payTbl.setBorderWidth(0);
  for (let i = 0; i < payRows.length; i++) {
    const labelCell = payTbl.getCell(i, 0);
    const valueCell = payTbl.getCell(i, 1);
    labelCell.editAsText().setFontSize(10).setBold(false).setForegroundColor('#8a8078');
    valueCell.editAsText().setFontSize(11).setBold(true).setForegroundColor('#3a3530');
    _alignCell(valueCell, DocumentApp.HorizontalAlignment.RIGHT);
  }

  body.appendHorizontalRule();

  // Footer
  body.appendParagraph('Gracias por elegir Las Nubes')
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setFontSize(11).setBold(false).setItalic(true).setForegroundColor('#6b6560');
  body.appendParagraph('WhatsApp +507 6981-2266 · lasnubes.cloud')
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .editAsText().setFontSize(9).setBold(false).setItalic(false).setForegroundColor('#8a8078');

  doc.saveAndClose();

  // Exportar a PDF
  const safeName = ((r.name || 'huesped').toString())
    .replace(/\s+/g, '_')
    .replace(/[^A-Za-z0-9_]/g, '')
    .slice(0, 40) || 'huesped';
  const pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
  pdfBlob.setName('Recibo_' + numStr + '_' + safeName + '.pdf');

  // Guardar copia en carpeta archivada
  try {
    const folder = getOrCreateRecibosFolder();
    folder.createFile(pdfBlob.copyBlob().setName(pdfBlob.getName()));
  } catch(e) {
    Logger.log('Error guardando recibo en Drive: ' + e);
  }

  // Trash temp doc
  try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch(_) {}

  return { blob: pdfBlob, number: numStr };
}

function buildCancellationEmailHTML(r) {
  const cabin       = CABIN_NAMES_EMAIL[r.cabin] || r.cabin;
  const color       = CABIN_COLORS_EMAIL[r.cabin] || '#6a9e62';
  const checkinStr  = r.checkin  instanceof Date ? Utilities.formatDate(r.checkin,  'America/Panama', 'yyyy-MM-dd') : r.checkin.toString().slice(0,10);
  const checkoutStr = r.checkout instanceof Date ? Utilities.formatDate(r.checkout, 'America/Panama', 'yyyy-MM-dd') : r.checkout.toString().slice(0,10);
  const nights      = nightCount(checkinStr, checkoutStr);

  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"></head>' +
'<body style="margin:0;padding:0;background:#f5f3f0;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;">' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f3f0;padding:32px 16px;"><tr><td align="center">' +
'<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;">' +
'<tr><td style="background:#5a5550;border-radius:16px 16px 0 0;padding:36px 40px;text-align:center;">' +
'<p style="margin:0 0 6px;font-size:12px;color:rgba(255,255,255,0.65);letter-spacing:2px;text-transform:uppercase;">Cancelación de reserva</p>' +
'<h1 style="margin:0;font-size:32px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></h1>' +
'</td></tr>' +
'<tr><td style="background:#ffffff;padding:36px 40px;">' +
'<p style="margin:0 0 24px;font-size:16px;color:#3a3530;line-height:1.6;">Hola <strong>' + r.name + '</strong>, te confirmamos que tu reserva ha sido cancelada.</p>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f9f8f6;border-radius:12px;border:1px solid #e8e4de;margin-bottom:28px;">' +
'<tr><td colspan="2" style="padding:16px 24px;border-bottom:1px solid #e8e4de;"><p style="margin:0 0 2px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Reserva cancelada</p><p style="margin:0;font-size:16px;font-weight:600;color:#5a5550;text-decoration:line-through;">' + cabin + '</p></td></tr>' +
'<tr><td style="padding:16px 24px;border-bottom:1px solid #e8e4de;"><p style="margin:0 0 3px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Check-in</p><p style="margin:0;font-size:14px;color:#8a8078;text-decoration:line-through;">' + formatDateES(checkinStr) + '</p></td>' +
'<td style="padding:16px 24px;border-bottom:1px solid #e8e4de;"><p style="margin:0 0 3px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Check-out</p><p style="margin:0;font-size:14px;color:#8a8078;text-decoration:line-through;">' + formatDateES(checkoutStr) + '</p></td></tr>' +
'<tr><td style="padding:16px 24px;"><p style="margin:0 0 3px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Noches</p><p style="margin:0;font-size:14px;color:#8a8078;text-decoration:line-through;">' + nights + '</p></td>' +
'<td style="padding:16px 24px;"><p style="margin:0 0 3px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Personas</p><p style="margin:0;font-size:14px;color:#8a8078;text-decoration:line-through;">' + (r.persons || '—') + '</p></td></tr>' +
'</table>' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#fdf8f0;border-radius:12px;border:1px solid #f0e8d8;margin-bottom:24px;"><tr><td style="padding:20px 24px;"><p style="margin:0 0 8px;font-size:14px;font-weight:600;color:#3a3530;">¿Tienes otras fechas en mente?</p><p style="margin:0;font-size:13px;color:#6b6560;line-height:1.6;">Nos encantaría tenerte en Las Nubes. Escríbenos por WhatsApp y revisamos disponibilidad.</p></td></tr></table>' +
'<a href="https://wa.me/50769812266?text=' + encodeURIComponent('Hola! Me gustaría consultar disponibilidad para una nueva fecha en Las Nubes.') + '" target="_blank" style="display:inline-block;background:' + color + ';color:#ffffff;font-size:14px;font-weight:500;padding:12px 28px;border-radius:10px;text-decoration:none;">&#128172; Consultar nueva fecha</a>' +
'</td></tr>' +
'<tr><td style="background:#3a3530;border-radius:0 0 16px 16px;padding:24px 40px;text-align:center;">' +
'<p style="margin:0 0 8px;font-size:18px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></p>' +
'<a href="https://wa.me/50769812266" style="color:rgba(255,255,255,0.8);font-size:13px;text-decoration:none;">&#128172; WhatsApp: +507 6981-2266</a>' +
'</td></tr></table></td></tr></table></body></html>';
}

function sendCancellationEmail(reservation) {
  try {
    const email = reservation.email;
    if (!email) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'No hay email de huésped' })).setMimeType(ContentService.MimeType.JSON);
    const cabin   = CABIN_NAMES_EMAIL[reservation.cabin] || reservation.cabin;
    const subject = '❌ Cancelación de reserva — ' + cabin + ' · Las Nubes';
    GmailApp.sendEmail(email, subject, '', { htmlBody: buildCancellationEmailHTML(reservation), name: 'Las Nubes', replyTo: REPLY_TO_EMAIL });
    Logger.log('📧 Cancelación enviada a: ' + email);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, status: 'cancellation_email_sent' })).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: e.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function buildUpdateEmailHTML(reservation, cabin, color, checkinFmt, checkoutFmt, nights, amount, deposit, saldo, hasSaldo, comentarios, gcalUrl) {
  const ics    = icsContent(reservation);
  const icsB64 = Utilities.base64Encode(ics);
  const icsUri = 'data:text/calendar;base64,' + icsB64;
  // Recalcular usando el tipo de reserva (overrides los valores pasados en posicionales legacy)
  const meta     = tipoEmailMeta(reservation);
  checkinFmt     = meta.checkinFmt;
  checkoutFmt    = meta.checkoutFmt;
  hasSaldo       = amount > 0 && parseFloat(saldo) > 0;
  const pagarUrl = 'https://wa.me/50769812266?text=' + encodeURIComponent('Deseo cancelar el saldo restante de mi reserva del día ' + meta.checkinFmt + ' en la cabaña ' + cabin + '. ¿Me comparte los métodos de pago?');
  const publicLink = getPublicReservaUrlSafe(reservation.id);

  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"></head>' +
'<body style="margin:0;padding:0;background:#f5f3f0;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;">' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f3f0;padding:32px 16px;"><tr><td align="center">' +
'<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;">' +
'<tr><td style="background:' + color + ';border-radius:16px 16px 0 0;padding:36px 40px;text-align:center;">' +
'<p style="margin:0 0 6px;font-size:12px;color:rgba(255,255,255,0.75);letter-spacing:2px;text-transform:uppercase;">Actualización de reserva</p>' +
'<h1 style="margin:0;font-size:32px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></h1>' +
'<p style="margin:10px 0 0;font-size:15px;color:rgba(255,255,255,0.9);">Buenos Aires, Chame · Panamá Oeste</p>' +
'</td></tr>' +
'<tr><td style="background:#ffffff;padding:36px 40px;">' +
'<p style="margin:0 0 18px;font-size:16px;color:#3a3530;line-height:1.6;">Hola <strong>' + reservation.name + '</strong>, los datos de tu reserva han sido actualizados.</p>' +
(publicLink ? '<p style="margin:0 0 24px;font-size:13px;color:#6b6560;">&#128279; <a href="' + publicLink + '" target="_blank" style="color:' + color + ';text-decoration:none;font-weight:500;border-bottom:1px solid ' + color + ';">Ver detalles actualizados online</a> &mdash; el link te muestra siempre la informaci&oacute;n vigente.</p>' : '') +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f9f8f6;border-radius:12px;border:1px solid #e8e4de;margin-bottom:24px;">' +
'<tr><td style="padding:20px 24px;border-bottom:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Cabaña</p><p style="margin:0;font-size:17px;font-weight:600;color:' + color + ';">' + cabin + '</p></td>' +
'<td style="padding:20px 24px;border-bottom:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Personas</p><p style="margin:0;font-size:17px;font-weight:600;color:#3a3530;">' + reservation.persons + '</p></td></tr>' +
'<tr><td style="padding:20px 24px;border-bottom:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Check-in</p><p style="margin:0;font-size:15px;color:#3a3530;font-weight:500;">' + checkinFmt + '</p><p style="margin:4px 0 0;font-size:12px;color:#8a8078;">' + meta.checkinHora + '</p></td>' +
'<td style="padding:20px 24px;border-bottom:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Check-out</p><p style="margin:0;font-size:15px;color:#3a3530;font-weight:500;">' + checkoutFmt + '</p><p style="margin:4px 0 0;font-size:12px;color:#8a8078;">' + meta.checkoutHora + '</p></td></tr>' +
'<tr><td style="padding:20px 24px;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">' + meta.estanciaLabel + '</p><p style="margin:0;font-size:17px;font-weight:600;color:#3a3530;">' + meta.estanciaValue + '</p></td>' +
'<td style="padding:20px 24px;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Total</p><p style="margin:0;font-size:17px;font-weight:600;color:#3a3530;">$' + amount.toFixed(2) + '</p>' + (deposit > 0 ? '<p style="margin:4px 0 0;font-size:12px;color:#8a8078;">Abono: $' + deposit.toFixed(2) + '</p>' : '') + '</td></tr>' +
(hasSaldo ? '<tr><td colspan="2" style="padding:18px 24px;background:#fff8e1;border-top:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:13px;color:#8a6000;">&#9888; <strong>Saldo pendiente: $' + saldo + '</strong></p><p style="margin:0 0 12px;font-size:12px;color:#8a6000;line-height:1.5;">Te pedimos cancelar el saldo antes del día de tu reserva.</p><a href="' + pagarUrl + '" target="_blank" style="display:inline-block;background:#25d366;color:#ffffff;font-size:13px;font-weight:600;padding:9px 20px;border-radius:8px;text-decoration:none;">&#128172; Pagar saldo</a></td></tr>' : '') +
'</table>' +
(comentarios ? '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff8e1;border-radius:12px;padding:20px;margin-bottom:24px;"><tr><td><p style="margin:0 0 8px;font-size:12px;color:#999;text-transform:uppercase;">Comentarios</p><p style="margin:0;font-size:14px;color:#555;line-height:1.6;">' + comentarios + '</p></td></tr></table>' : '') +
'<p style="margin:0 0 12px;font-size:14px;font-weight:600;color:#3a3530;">Actualizar tu calendario</p>' +
'<table cellpadding="0" cellspacing="0" style="margin-bottom:32px;"><tr>' +
'<td style="padding-right:10px;"><a href="' + gcalUrl + '" target="_blank" style="display:inline-block;background:#4285f4;color:#ffffff;font-size:13px;font-weight:500;padding:10px 20px;border-radius:8px;text-decoration:none;">&#128197; Google Calendar</a></td>' +
'<td><a href="' + icsUri + '" style="display:inline-block;background:#3a3530;color:#ffffff;font-size:13px;font-weight:500;padding:10px 20px;border-radius:8px;text-decoration:none;">&#127822; Apple / Outlook</a></td>' +
'</tr></table>' +
'<hr style="border:none;border-top:1px solid #e8e4de;margin:0 0 28px;">' +
buildGuiaHTML(reservation.cabin, meta.tipo, !!reservation.checkoutExtendido, meta.horaSalidaCustom) +
'</td></tr>' +
'<tr><td style="background:#3a3530;border-radius:0 0 16px 16px;padding:24px 40px;text-align:center;">' +
'<p style="margin:0 0 8px;font-size:18px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></p>' +
'<a href="https://wa.me/50769812266" style="color:rgba(255,255,255,0.8);font-size:13px;text-decoration:none;">&#128172; WhatsApp: +507 6981-2266</a>' +
'</td></tr></table></td></tr></table></body></html>';
}

function sendUpdateEmail(reservation, voucherBase64, voucherMimeType) {
  try {
    const email = reservation.email;
    if (!email) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'No hay email de huésped' })).setMimeType(ContentService.MimeType.JSON);
    const cabin       = CABIN_NAMES_EMAIL[reservation.cabin] || reservation.cabin;
    const color       = CABIN_COLORS_EMAIL[reservation.cabin] || '#6a9e62';
    const checkinStr  = reservation.checkin.toString().slice(0,10);
    const checkoutStr = reservation.checkout.toString().slice(0,10);
    const amount      = parseFloat(reservation.amount)  || 0;
    const deposit     = parseFloat(reservation.deposit) || 0;
    const saldo       = (amount - deposit).toFixed(2);
    const hasSaldo    = deposit > 0 && parseFloat(saldo) > 0;
    // En un REGALO no se manda el email de actualización (habla de montos,
    // abonos y saldo): se reenvía el certificado, que no lleva plata.
    const isRegalo    = esReservaRegalo(reservation);
    const subject     = isRegalo
      ? '🎁 Tu regalo en Las Nubes — actualizado'
      : 'Actualización de reserva — ' + cabin + ' · Las Nubes';
    const html = isRegalo
      ? buildEmailHTMLRegalo(reservation)
      : buildUpdateEmailHTML(reservation, cabin, color, formatDateES(checkinStr), formatDateES(checkoutStr), nightCount(checkinStr, checkoutStr), amount, deposit, saldo, hasSaldo, reservation.comentarios || '', googleCalLink(reservation));

    // Si la edicion incluye un voucher nuevo, generar recibo PDF y adjuntar
    // (mismo flujo que sendConfirmationEmail). Nunca en un regalo.
    const attachments = [];
    if (voucherBase64 && voucherMimeType && !isRegalo) {
      try {
        const receipt = generateReceiptPDF(reservation);
        attachments.push(receipt.blob);
        Logger.log('📄 Recibo ' + receipt.number + ' adjuntado en update');
      } catch(rcpErr) {
        Logger.log('⚠ No se pudo generar recibo PDF en update: ' + rcpErr);
      }
      try {
        const ext = voucherMimeType.includes('png') ? '.png' : voucherMimeType.includes('gif') ? '.gif' : voucherMimeType.includes('webp') ? '.webp' : '.jpg';
        const safeName = ((reservation.name || 'huesped').toString())
          .replace(/\s+/g, '_').replace(/[^A-Za-z0-9_]/g, '').slice(0, 40) || 'huesped';
        const voucherBlob = Utilities.newBlob(Utilities.base64Decode(voucherBase64), voucherMimeType, 'Voucher_' + safeName + ext);
        attachments.push(voucherBlob);
      } catch(vErr) {
        Logger.log('⚠ No se pudo adjuntar voucher imagen en update: ' + vErr);
      }
    }

    GmailApp.sendEmail(email, subject, '', { htmlBody: html, attachments: attachments, name: 'Las Nubes', replyTo: REPLY_TO_EMAIL });
    Logger.log('📧 Update enviado a: ' + email + ' (adjuntos: ' + attachments.length + ')');
    return ContentService.createTextOutput(JSON.stringify({ ok: true, status: 'update_email_sent', attachments: attachments.length })).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: e.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ═══════════════════════════════════════════════════════════
//  Recordatorio de check-in (email automatico 1 dia antes)
// ═══════════════════════════════════════════════════════════

// Configuracion editable via Script Properties (sin tocar codigo).
// Defaults razonables para que el preview se vea completo aunque no
// se hayan configurado todas las claves.
function getCheckinReminderConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    mapsUrl:       props.getProperty('CHECKIN_MAPS_URL')       || 'https://maps.google.com/?q=8.639400,-79.945900',
    wazeUrl:       props.getProperty('CHECKIN_WAZE_URL')       || 'https://waze.com/ul?ll=8.639400,-79.945900&navigate=yes',
    indicaciones:  props.getProperty('CHECKIN_INDICACIONES')   || 'Por carretera interamericana, entra por el Pío Pío de Bejuco hacia carretera Bejuco–Sorá. Al llegar al pueblo de Buenos Aires, dobla a la derecha hacia Chicá. La cabaña queda a 100 metros. En Waze busca <strong>"Aires de Chicá"</strong> y te llevará directo al portón verde.',
    accesoExtra:   props.getProperty('CHECKIN_ACCESO_EXTRA')   || ''
  };
}

function buildCheckinReminderHTML(r, meta, config) {
  const cabin   = CABIN_NAMES_EMAIL[r.cabin] || r.cabin;
  const color   = CABIN_COLORS_EMAIL[r.cabin] || '#6a9e62';
  const amount  = parseFloat(r.amount)  || 0;
  const deposit = parseFloat(r.deposit) || 0;
  const saldo   = (amount - deposit).toFixed(2);
  // En un CERTIFICADO DE REGALO con fecha definida este recordatorio sí se
  // dispara (el regalo "Abierta" nunca llega acá porque no tiene check-in).
  // El beneficiario no debe ver saldo: lo pagó quien se lo regaló.
  const hasSaldo = !esReservaRegalo(r) && amount > 0 && parseFloat(saldo) > 0;
  const pagarUrl = 'https://wa.me/50769812266?text=' + encodeURIComponent('Hola! Quiero cancelar el saldo restante de mi reserva del ' + meta.checkinFmt + ' en la cabaña ' + cabin + '.');
  const llegadaUrl = 'https://wa.me/50769812266?text=' + encodeURIComponent('Hola! Soy ' + (r.name || 'huésped') + ', llego mañana a Las Nubes (' + cabin + ').');

  // Instrucciones de acceso por cabaña (key box code 0507 — mismo que buildGuiaHTML)
  const accesoTexto = 'Key Box código <strong>0507</strong>' +
    (r.cabin === 'azul'  ? ' &middot; luego desliza la puerta corrediza de metal' : '') +
    '.';
  const accesoExtraHtml = config.accesoExtra
    ? '<p style="margin:6px 0 0;font-size:12px;color:#8a8078;line-height:1.5;">' + config.accesoExtra + '</p>'
    : '';

  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"></head>' +
'<body style="margin:0;padding:0;background:#f5f3f0;font-family:\'Helvetica Neue\',Helvetica,Arial,sans-serif;">' +
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f5f3f0;padding:32px 16px;"><tr><td align="center">' +
'<table width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;">' +

// HERO
'<tr><td style="background:' + color + ';border-radius:16px 16px 0 0;padding:36px 40px;text-align:center;">' +
'<p style="margin:0 0 6px;font-size:12px;color:rgba(255,255,255,0.75);letter-spacing:2px;text-transform:uppercase;">Te esperamos mañana</p>' +
'<h1 style="margin:0;font-size:32px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></h1>' +
'<p style="margin:12px 0 0;font-size:14px;color:rgba(255,255,255,0.9);">' + meta.checkinFmt + ' &middot; ' + meta.checkinHora + '</p>' +
'</td></tr>' +

// CUERPO
'<tr><td style="background:#ffffff;padding:32px 40px;">' +
'<p style="margin:0 0 24px;font-size:16px;color:#3a3530;line-height:1.6;">Hola <strong>' + (r.name || '') + '</strong>, falta solo un día para tu llegada a <strong style="color:' + color + ';">' + cabin + '</strong>. Aquí lo esencial para tu check-in.</p>' +

// RESUMEN (cabaña / personas / llegada / salida)
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f9f8f6;border-radius:12px;border:1px solid #e8e4de;margin-bottom:20px;">' +
'<tr><td style="padding:18px 22px;border-bottom:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Cabaña</p><p style="margin:0;font-size:16px;font-weight:600;color:' + color + ';">' + cabin + '</p></td>' +
'<td style="padding:18px 22px;border-bottom:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Personas</p><p style="margin:0;font-size:16px;font-weight:600;color:#3a3530;">' + (r.persons || '—') + '</p></td></tr>' +
'<tr><td style="padding:18px 22px;border-bottom:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Llegada</p><p style="margin:0;font-size:14px;color:#3a3530;font-weight:500;">' + meta.checkinFmt + '</p><p style="margin:4px 0 0;font-size:12px;color:#8a8078;">' + meta.checkinHora + '</p></td>' +
'<td style="padding:18px 22px;border-bottom:1px solid #e8e4de;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">Salida</p><p style="margin:0;font-size:14px;color:#3a3530;font-weight:500;">' + meta.checkoutFmt + '</p><p style="margin:4px 0 0;font-size:12px;color:#8a8078;">' + meta.checkoutHora + '</p></td></tr>' +
'<tr><td colspan="2" style="padding:14px 22px;"><p style="margin:0 0 4px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;">' + meta.estanciaLabel + '</p><p style="margin:0;font-size:15px;font-weight:600;color:#3a3530;">' + meta.estanciaValue + '</p></td></tr>' +
'</table>' +

// CÓMO LLEGAR
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#eaf4e6;border-radius:12px;border:1px solid #d0e6c6;margin-bottom:20px;"><tr><td style="padding:20px 22px;">' +
'<p style="margin:0 0 6px;font-size:11px;color:#4a7340;text-transform:uppercase;letter-spacing:1px;font-weight:700;">&#128205; Cómo llegar</p>' +
'<p style="margin:0 0 12px;font-size:14px;color:#3a3530;line-height:1.5;"><strong>Buenos Aires, Chame</strong> &middot; Panamá Oeste</p>' +
'<p style="margin:0 0 14px;font-size:13px;color:#5a5550;line-height:1.6;">' + config.indicaciones + '</p>' +
'<table cellpadding="0" cellspacing="0"><tr>' +
'<td style="padding-right:8px;"><a href="' + config.mapsUrl + '" target="_blank" style="display:inline-block;background:#4285f4;color:#ffffff;font-size:13px;font-weight:500;padding:10px 18px;border-radius:8px;text-decoration:none;">&#128506; Google Maps</a></td>' +
'<td><a href="' + config.wazeUrl + '" target="_blank" style="display:inline-block;background:#33ccff;color:#ffffff;font-size:13px;font-weight:500;padding:10px 18px;border-radius:8px;text-decoration:none;">&#128663; Waze</a></td>' +
'</tr></table>' +
'</td></tr></table>' +

// ACCESO
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#fdf8f0;border-radius:12px;border:1px solid #f0e8d8;margin-bottom:20px;"><tr><td style="padding:18px 22px;">' +
'<p style="margin:0 0 6px;font-size:11px;color:#8a6000;text-transform:uppercase;letter-spacing:1px;font-weight:700;">&#128273; Acceso</p>' +
'<p style="margin:0;font-size:13px;color:#3a3530;line-height:1.6;">' + accesoTexto + '</p>' +
accesoExtraHtml +
'</td></tr></table>' +

// SALDO PENDIENTE
(hasSaldo ? '<table width="100%" cellpadding="0" cellspacing="0" style="background:#fff3cd;border-radius:12px;border:1px solid #ffe6a3;margin-bottom:20px;"><tr><td style="padding:18px 22px;">' +
'<p style="margin:0 0 4px;font-size:13px;color:#8a6000;font-weight:700;">&#9888; Saldo pendiente: $' + saldo + '</p>' +
'<p style="margin:0 0 12px;font-size:12px;color:#8a6000;line-height:1.5;">Te pedimos cancelar el saldo antes de tu llegada para agilizar el check-in.</p>' +
'<a href="' + pagarUrl + '" target="_blank" style="display:inline-block;background:#25d366;color:#ffffff;font-size:13px;font-weight:600;padding:9px 20px;border-radius:8px;text-decoration:none;">&#128172; Coordinar pago</a>' +
'</td></tr></table>' : '') +

// QUÉ LLEVAR
'<table width="100%" cellpadding="0" cellspacing="0" style="background:#f9f8f6;border-radius:12px;border:1px solid #e8e4de;margin-bottom:20px;"><tr><td style="padding:18px 22px;">' +
'<p style="margin:0 0 8px;font-size:11px;color:#8a8078;text-transform:uppercase;letter-spacing:1px;font-weight:700;">&#127890; Te recomendamos llevar</p>' +
'<p style="margin:0;font-size:13px;color:#5a5550;line-height:1.7;">Hielo y tus alimentos &middot; Repelente (si eres sensible a los mosquitos) &middot; Ropa de abrigo (refresca de noche) &middot; Calzado cómodo</p>' +
'</td></tr></table>' +

// EMERGENCIAS
'<p style="margin:0 0 8px;font-size:13px;color:#3a3530;text-align:center;">¿Algún cambio o duda?</p>' +
'<p style="margin:0 0 24px;text-align:center;"><a href="' + llegadaUrl + '" target="_blank" style="display:inline-block;background:#25d366;color:#ffffff;font-size:14px;font-weight:600;padding:11px 24px;border-radius:10px;text-decoration:none;">&#128172; Escríbenos por WhatsApp</a></p>' +

'<p style="margin:0;font-size:14px;color:#6b6560;line-height:1.6;text-align:center;font-style:italic;">¡Nos vemos mañana!</p>' +

// Las tres guías van ACÁ y no en el email de confirmación: ese llega al
// reservar, a veces semanas antes, cuando nadie está pensando en dónde cenar.
// Este llega la víspera, mientras arman la maleta y deciden qué hacer — y hasta
// hoy no enlazaba a nada.
'<hr style="border:none;border-top:1px solid #e8e4de;margin:24px 0 20px;">' +
'<p style="margin:0 0 14px;font-size:14px;font-weight:600;color:#3a3530;text-align:center;">Para ir preparando</p>' +
'<table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center">' +
'<a href="https://lasnubes.cloud/#insumos" target="_blank" style="display:inline-block;background:#5a85b0;color:#ffffff;font-size:13px;font-weight:600;padding:10px 18px;border-radius:8px;text-decoration:none;margin:0 4px 8px;">&#128722; Qué llevar</a>' +
'<a href="https://lasnubes.cloud/#actividades" target="_blank" style="display:inline-block;background:#4a7c5f;color:#ffffff;font-size:13px;font-weight:600;padding:10px 18px;border-radius:8px;text-decoration:none;margin:0 4px 8px;">&#127966;&#65039; Actividades</a>' +
'<a href="https://lasnubes.cloud/#gastronomia" target="_blank" style="display:inline-block;background:#a0673f;color:#ffffff;font-size:13px;font-weight:600;padding:10px 18px;border-radius:8px;text-decoration:none;margin:0 4px 8px;">&#127869;&#65039; Dónde comer</a>' +
'</td></tr></table>' +
'</td></tr>' +

// FOOTER
'<tr><td style="background:#3a3530;border-radius:0 0 16px 16px;padding:24px 40px;text-align:center;">' +
'<p style="margin:0 0 8px;font-size:18px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></p>' +
'<a href="https://wa.me/50769812266" style="color:rgba(255,255,255,0.8);font-size:13px;text-decoration:none;">&#128172; WhatsApp +507 6981-2266</a>' +
'</td></tr>' +

'</table></td></tr></table></body></html>';
}

function sendCheckinReminderEmail(reservation) {
  try {
    const email = reservation.email;
    if (!email) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'No hay email de huésped' })).setMimeType(ContentService.MimeType.JSON);
    const cabin   = CABIN_NAMES_EMAIL[reservation.cabin] || reservation.cabin;
    const meta    = tipoEmailMeta(reservation);
    const config  = getCheckinReminderConfig();
    const subject = 'Te esperamos mañana — ' + cabin + ' · Las Nubes';
    const html    = buildCheckinReminderHTML(reservation, meta, config);
    GmailApp.sendEmail(email, subject, '', { htmlBody: html, name: 'Las Nubes', replyTo: REPLY_TO_EMAIL });
    Logger.log('📧 Recordatorio check-in enviado a: ' + email);
    return ContentService.createTextOutput(JSON.stringify({ ok: true, status: 'reminder_sent' })).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: e.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// "vie 5 al dom 7 de junio" (o "vie 5 de junio" si es un solo día).
// Para el {{3}} de la plantilla recordator_entrada.
function _fechasRangoCorto(checkinStr, checkoutStr) {
  const dias  = ['dom','lun','mar','mié','jue','vie','sáb'];
  const meses = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  const parse = s => { const p = s.toString().slice(0,10).split('-'); return new Date(+p[0], +p[1]-1, +p[2]); };
  const a = parse(checkinStr), b = parse(checkoutStr);
  const ini = dias[a.getDay()] + ' ' + a.getDate();
  if (a.getTime() === b.getTime()) return ini + ' de ' + meses[a.getMonth()];
  const fin = dias[b.getDay()] + ' ' + b.getDate();
  if (a.getMonth() === b.getMonth()) return ini + ' al ' + fin + ' de ' + meses[b.getMonth()];
  return ini + ' de ' + meses[a.getMonth()] + ' al ' + fin + ' de ' + meses[b.getMonth()];
}

// Hora corta para plantillas de WhatsApp ("2:00 pm", "11:00 am", "9:00 am",
// "5:00 pm", "12:30 pm (cortesía)", etc.) según el tipo de reserva.
// kind: 'checkin' | 'checkout'. Si se pasan horas custom (HoraEntrada /
// HoraSalida por reserva), pisan al default del tipo y a la cortesía.
function _horaPlantilla(tipo, kind, checkoutExtendido, horaEntradaCustom, horaSalidaCustom) {
  const t = (tipo || 'noche').toString();
  if (kind === 'checkin') {
    const custom = _normalizeHora(horaEntradaCustom);
    if (custom) return _formatHora12(custom);
    if (t === 'pasatarde') return '12:30 pm';
    if (t === 'pasanoche') return '8:00 pm';
    if (t === 'pasadia' || t === 'early') return '9:00 am';
    return '2:00 pm';                          // noche, late
  }
  // checkout
  let custom = _normalizeHora(horaSalidaCustom);
  // Guard: en pasatarde/pasadía la salida debe ser posterior a la entrada; una
  // hora-salida custom inválida (ej. 8am con entrada 12:30pm) se ignora.
  if (custom && (t === 'pasatarde' || t === 'pasadia')) {
    const _toMin = s => { const p = String(s).split(':'); return (parseInt(p[0],10)||0)*60 + (parseInt(p[1],10)||0); };
    const entradaRef = _normalizeHora(horaEntradaCustom) || (t === 'pasatarde' ? '12:30' : '09:00');
    if (_toMin(custom) <= _toMin(entradaRef)) custom = '';
  }
  if (custom) return _formatHora12(custom);
  if (t === 'pasatarde') return '7:00 pm';
  if (t === 'pasanoche') return '12:30 pm';
  if (t === 'pasadia')   return '5:00 pm';
  if (t === 'late')      return '4:00 pm';
  // noche / early: 11am, salvo cortesía 12:30 pm
  if (checkoutExtendido) return '12:30 pm (cortesía)';
  return '11:00 am';
}

// Trigger diario @ 10am Panama. Escanea Reservas y manda recordatorio
// a quienes hacen check-in REAL (display) mañana.
// Envía por WhatsApp (plantilla recordator_entrada, si hay teléfono y no es
// Airbnb) y por email (si hay email). Excluye CANCELADA y origen Abierta.
function enviarRecordatoriosCheckin() {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, 'America/Panama', 'yyyy-MM-dd');

  let sent = 0, skipped = 0, errors = 0, waSent = 0, waErrors = 0;
  for (let i = 1; i < data.length; i++) {
    const r = {
      id:         data[i][0],
      name:       data[i][1],
      cabin:      data[i][3],
      checkin:    data[i][4],
      checkout:   data[i][5],
      persons:    data[i][6],
      amount:     data[i][7],
      deposit:    data[i][8],
      origin:     data[i][9],
      estadoPago: data[i][20] || '',
      email:      data[i][21] || '',
      telefono:   data[i][23] || '',
      tipo:       data[i][24] || 'noche',
      // Cols 29/30/31: cortesía + horas custom. Sin esto, un check-in a las
      // 12:30pm (ej. pasatarde convertido a noche) se anunciaría al default 2pm.
      checkoutExtendido: data[i][28] === true || data[i][28] === 'TRUE' || data[i][28] === 'true' || data[i][28] === 1,
      horaEntrada: _normalizeHora(data[i][29]),
      horaSalida:  _normalizeHora(data[i][30]),
      // Col 33: sin esto el recordatorio de un regalo con fecha le mostraría el
      // saldo al beneficiario (ver buildCheckinReminderHTML).
      regalo:      data[i][32] || '',
      pagador:     data[i][14] || ''
    };
    if (!r.checkin) continue;
    if (r.estadoPago === 'CANCELADA') continue;
    if (r.origin === 'Abierta') continue;
    const meta = tipoEmailMeta(r);
    if (meta.displayCheckin !== tomorrowStr) continue;

    // WhatsApp: plantilla recordator_entrada (botón "Envíame ubicación").
    if (r.telefono && r.origin !== 'Airbnb') {
      try {
        const firstName  = (r.name || '').toString().trim().split(/\s+/)[0] || 'amigo';
        const cabinName  = BOT_CABIN_NAMES[r.cabin] || r.cabin;
        const fechas     = _fechasRangoCorto(meta.displayCheckin, meta.displayCheckout);
        const checkinHr  = _horaPlantilla(r.tipo, 'checkin', r.checkoutExtendido, r.horaEntrada);
        sendWhatsAppTemplate(r.telefono, 'recordator_entrada', 'es_ES',
          [firstName, cabinName, fechas, checkinHr], null, 'ubicacion_' + r.id);
        waSent++;
      } catch(e) {
        waErrors++;
        Logger.log('⚠ WA check-in ' + r.name + ': ' + e);
      }
    }

    // Email (canal complementario).
    if (!r.email) { skipped++; Logger.log('⊘ Sin email: ' + r.name); continue; }
    try {
      sendCheckinReminderEmail(r);
      sent++;
    } catch(e) {
      errors++;
      Logger.log('⚠ Error con ' + r.name + ' (' + r.email + '): ' + e);
    }
  }
  Logger.log('✓ Recordatorios check-in: WhatsApp ' + waSent + ' (' + waErrors + ' err) · email ' + sent + ' (' + skipped + ' sin email, ' + errors + ' err)');
}

// Ejecutar UNA VEZ desde el editor de Apps Script para activar el envio diario.
// Idempotente: borra triggers previos de la misma funcion antes de crear el nuevo.
function instalarTriggerRecordatorios() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'enviarRecordatoriosCheckin') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('enviarRecordatoriosCheckin')
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .inTimezone('America/Panama')
    .create();
  Logger.log('✓ Trigger creado: enviarRecordatoriosCheckin diario @ 10am America/Panama');
}

// Ejecutar desde el editor para previsualizar el email — se envia al
// correo del usuario que corre el script (no al cliente real).
function enviarRecordatorioPrueba() {
  const myEmail = Session.getActiveUser().getEmail();
  if (!myEmail) {
    Logger.log('⚠ No se pudo obtener el email del usuario activo. Ejecuta desde el editor de Apps Script.');
    return;
  }
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const dayAfter = new Date();
  dayAfter.setDate(dayAfter.getDate() + 2);
  const sample = {
    id:         'PREVIEW-' + Date.now(),
    name:       'Juan de Prueba',
    email:      myEmail,
    cabin:      'verde',
    checkin:    Utilities.formatDate(tomorrow, 'America/Panama', 'yyyy-MM-dd'),
    checkout:   Utilities.formatDate(dayAfter, 'America/Panama', 'yyyy-MM-dd'),
    persons:    2,
    amount:     180,
    deposit:    90,
    origin:     'Directa',
    estadoPago: 'ABONADO',
    telefono:   '+507 1234-5678',
    tipo:       'noche'
  };
  sendCheckinReminderEmail(sample);
  Logger.log('✓ Recordatorio de prueba enviado a ' + myEmail);
}

// ═══════════════════════════════════════════════════════════
//  Vouchers & Fotos
// ═══════════════════════════════════════════════════════════

const VOUCHER_FOLDER_NAME = 'Las Nubes - Pagos';

function saveFotoEgreso(imageBase64, mimeType, egresoId, fecha) {
  try {
    const rootFolders = DriveApp.getFoldersByName(FACTURAS_FOLDER_NAME);
    const rootFolder  = rootFolders.hasNext() ? rootFolders.next() : DriveApp.createFolder(FACTURAS_FOLDER_NAME);
    const mes         = fecha ? fecha.slice(0,7) : Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM');
    const subFolders  = rootFolder.getFoldersByName(mes);
    const subFolder   = subFolders.hasNext() ? subFolders.next() : rootFolder.createFolder(mes);
    const ext         = mimeType === 'image/png' ? 'png' : mimeType === 'image/webp' ? 'webp' : 'jpg';
    const fileName    = 'factura_' + (egresoId || 'egr') + '.' + ext;
    const blob        = Utilities.newBlob(Utilities.base64Decode(imageBase64), mimeType, fileName);
    const file        = subFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const url = 'https://drive.google.com/file/d/' + file.getId() + '/view';
    return ContentService.createTextOutput(JSON.stringify({ ok: true, url, fileId: file.getId() })).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: e.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Red de seguridad para fechas de facturas: los recibos panameños usan
// DD/MM/AAAA, pero a veces el OCR las lee como MM/DD y produce una fecha
// FUTURA (imposible en una factura real, ej. "08/07/26" = 8-jul leído como
// 7-ago). Si la fecha parseada es posterior a hoy y al intercambiar día↔mes
// queda una fecha de calendario válida y pasada, se corrige. Si no se puede
// corregir con seguridad, se deja tal cual. Determinístico (no depende de TZ).
function _corregirFechaFuturaFactura(fechaStr) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fechaStr || '')) return fechaStr || '';
  const hoy = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');
  if (fechaStr <= hoy) return fechaStr; // pasada o de hoy → OK
  const yr = parseInt(fechaStr.slice(0, 4), 10);
  const mm = parseInt(fechaStr.slice(5, 7), 10);
  const dd = parseInt(fechaStr.slice(8, 10), 10);
  // Intercambiar: el día viejo pasa a mes, el mes viejo pasa a día.
  const newMonth = dd, newDay = mm;
  if (newMonth < 1 || newMonth > 12) return fechaStr; // día viejo no es un mes válido
  const bis = (yr % 4 === 0 && yr % 100 !== 0) || yr % 400 === 0;
  const diasMes = [31, bis ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  if (newDay < 1 || newDay > diasMes[newMonth - 1]) return fechaStr; // día inválido para ese mes
  const swapped = fechaStr.slice(0, 4) + '-' +
    ('0' + newMonth).slice(-2) + '-' + ('0' + newDay).slice(-2);
  if (swapped <= hoy) {
    Logger.log('⚠ Fecha de factura futura ' + fechaStr + ' → ' + swapped + ' (corregida día↔mes)');
    return swapped;
  }
  return fechaStr; // el swap también da futuro → no se puede corregir
}

function parseFacturaEgresoConClaude(imageBase64, mimeType) {
  try {
    const prompt = `Analiza esta imagen. Puede ser una factura comercial, un comprobante de Yappy, una transferencia ACH, un recibo o cualquier documento de pago.

Extrae la siguiente información según el tipo de documento:

PARA COMPROBANTES YAPPY (tienen logo "yappy", fondo azul, campo "Mensaje" y "Confirmación"):
- proveedor: el nombre de la persona que RECIBE el pago (campo "Enviado a")
- fecha: la fecha del pago (formato YYYY-MM-DD)
- monto: el monto numérico sin símbolo (ej: 130.00)
- items: un array con UN solo ítem donde:
    - desc: lo que dice el campo "Mensaje" del voucher
    - monto: el mismo monto total
    - categoria: infiere según el mensaje (ej: "limpieza cabañas" → "Limpieza", "materiales" → "Suministros", etc.)
- numFactura: el código de confirmación (ej: UYAFL-85668143, sin el #)

PARA FACTURAS COMERCIALES TRADICIONALES:
- proveedor: nombre del emisor de la factura
- fecha: fecha de la factura (formato YYYY-MM-DD)
- monto: monto total
- items: array de líneas de la factura, cada una con:
    - desc: descripción del ítem
    - monto: monto numérico del ítem
    - categoria: una de estas opciones según el tipo: Limpieza, Mantenimiento, Suministros, Servicios, General, Otro
- numFactura: número de factura si existe

PARA TRANSFERENCIAS ACH / BANCARIAS:
- proveedor: nombre del destinatario
- fecha: fecha de la transferencia (formato YYYY-MM-DD)
- monto: monto numérico
- items: UN ítem con la descripción/concepto de la transferencia
- numFactura: número de referencia si existe

Responde ÚNICAMENTE con JSON válido, sin texto adicional ni markdown:
{
  "proveedor": "",
  "fecha": "",
  "monto": 0,
  "numFactura": "",
  "items": [
    { "desc": "", "monto": 0, "categoria": "General" }
  ]
}

IMPORTANTE — FORMATO DE FECHA (Panamá): los recibos y facturas usan formato DÍA/MES/AÑO (DD/MM/AAAA). Si ves la fecha como número tipo "08/07/26", "08-07-2026" o "080726", interprétala como día 08, mes 07 → 2026-07-08 (8 de julio), NUNCA como 7 de agosto. Además, una factura real jamás tiene fecha futura: si tu lectura da una fecha posterior a hoy, casi seguro invertiste el día y el mes.`;

    const payload = {
      model: 'claude-opus-4-6',
      max_tokens: 600,
      messages: [{ role: 'user', content: [
        { type: 'image', source: { type: 'base64', media_type: mimeType, data: imageBase64 } },
        { type: 'text', text: prompt }
      ]}]
    };

    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': getClaudeApiKey(), 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const result = JSON.parse(response.getContentText());
    if (!result.content || !result.content[0]) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Claude API error' })).setMimeType(ContentService.MimeType.JSON);
    const text = result.content[0].text.trim();
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'No JSON in response' })).setMimeType(ContentService.MimeType.JSON);
    const data = JSON.parse(jsonMatch[0]);
    const fechaSegura = _corregirFechaFuturaFactura(data.fecha || '');
    return ContentService.createTextOutput(JSON.stringify({ ok: true, proveedor: data.proveedor || '', fecha: fechaSegura, monto: data.monto || 0, numFactura: data.numFactura || '', items: Array.isArray(data.items) ? data.items : [] })).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: e.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function parseVoucherWithClaude(imageBase64, mimeType) {
  try {
    const prompt = `Lee este voucher de pago (Yappy, Banco General, ACH u otro) y devuelve EXCLUSIVAMENTE este JSON, sin markdown ni texto antes o después:

{
  "sender": "<nombre del REMITENTE — quien envió el dinero>",
  "monto": <número decimal sin símbolo $>,
  "fechaPago": "<YYYY-MM-DD>",
  "codTransferencia": "<código sin # inicial>",
  "mensaje": "<texto literal del campo Mensaje/Concepto/Comentario si existe, si no null>",
  "nombreCompleto": "<nombre de una persona dentro del Mensaje, si lo hay; si no null>",
  "email": "<email dentro del Mensaje o null>",
  "telefono": "<XXXX-XXXX dentro del Mensaje o null>"
}

Cómo encontrar cada campo:

SENDER — el REMITENTE (quien hace el pago, NO el destinatario):
  - En Yappy: texto grande arriba "<Nombre> te envió". Ej: "Alejandra D. te envió" → "Alejandra D.". "Jose Zaldaña te envió" → "Jose Zaldaña". Devuélvelo EXACTO como aparece (con iniciales y puntos si los tiene).
  - En Banco General / ACH / otros: campo "Cliente", "Remitente", "Ordenante", "Origen", "De".
  - NO confundas con "Enviado a" / "Beneficiario" / "Joslyn Lopez" / "Las Nubes" — eso es el DESTINATARIO y NO va aquí.
  - Este campo casi siempre está visible. Solo devuelve "" si la imagen está ilegible.

MONTO — número grande con "$" (Yappy: bajo el sender; otros: campo "Monto/Total/Importe"). Devuelve número JSON: 75.00, no "$75.00".

FECHA PAGO — campo "Fecha". Yappy formato "5 may 2026 - 12:05 p.m." → "2026-05-05". Meses: ene=01 feb=02 mar=03 abr=04 may=05 jun=06 jul=07 ago=08 sep=09 oct=10 nov=11 dic=12. Si no hay año, usa 2026. Zona Panamá.

CÓDIGO TRANSFERENCIA — campo "Confirmación" / "Referencia". Yappy formato "#XXXXX-NNNNNNNN". DEVUÉLVELO SIN el "#".

MENSAJE — solo si hay un campo "Mensaje", "Concepto", "Comentario", "Descripción" o "Detalle". Si no existe ese campo en el voucher, null.

NOMBRE COMPLETO — nombre de una persona DENTRO del mensaje (ej: "Reserva Karina" → "Karina", "Pago Juan Pérez" → "Juan Pérez", "ManuelFlores" → "Manuel Flores"). Acepta nombres simples y nombres pegados sin espacio (separa "ManuelFlores" → "Manuel Flores"). Excluye palabras tipo "abono", "pago", "reserva", "saldo", "anticipo", "deposito", "transferencia", fechas, montos, números → null. Es el nombre del HUÉSPED, distinto del SENDER.

EMAIL — cualquier dirección de correo (algo@dominio.xxx) que aparezca en el mensaje. Devuélvela tal cual, en minúsculas. Si no hay, null.

TELÉFONO — cualquier número panameño de 8 dígitos en el mensaje, ya sea con guión ("6300-4489"), sin guión ("63004489") o con espacios ("6300 4489"). Devuélvelo SIEMPRE en formato XXXX-XXXX (insertando el guión si no lo tiene). NO confundas con el monto, fechas, ni el código de confirmación. Si no hay número de 8 dígitos, null.

Responde solo el JSON.`;

    const buildPayload = (model) => ({
      model,
      max_tokens: 500,
      messages: [{ role: 'user', content: [
        { type: 'image', source: { type: 'base64', media_type: mimeType, data: imageBase64 } },
        { type: 'text', text: prompt }
      ]}]
    });

    const callClaude = (payload) => UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': getClaudeApiKey(), 'anthropic-version': '2023-06-01' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    // Retry con backoff ante 429/529 (overloaded). Tras 2 intentos en opus, cae a sonnet.
    const attempts = [
      { model: 'claude-opus-4-6', wait: 0    },
      { model: 'claude-opus-4-6', wait: 1500 },
      { model: 'claude-sonnet-4-6', wait: 1500 }
    ];

    let response, rawBody, status, result, lastModel;
    for (const a of attempts) {
      if (a.wait) Utilities.sleep(a.wait);
      lastModel = a.model;
      response  = callClaude(buildPayload(a.model));
      status    = response.getResponseCode();
      rawBody   = response.getContentText();
      try { result = JSON.parse(rawBody); } catch(_) { result = null; }
      if (status === 200 && result && result.content && result.content[0]) break;
      Logger.log('Voucher attempt ' + a.model + ' status ' + status + ': ' + rawBody.slice(0, 300));
      if (status !== 429 && status !== 529 && status < 500) break; // error no-transitorio: no reintentar
    }

    if (!result || !result.content || !result.content[0]) {
      const errMsg = (result && result.error && result.error.message) ? result.error.message : 'Claude API error';
      return ContentService.createTextOutput(JSON.stringify({
        ok: false,
        error: errMsg,
        _debug: { status, rawBody, payloadModel: lastModel }
      })).setMimeType(ContentService.MimeType.JSON);
    }
    const text = result.content[0].text.trim();
    Logger.log('Voucher Claude raw response: ' + text);
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'No JSON' })).setMimeType(ContentService.MimeType.JSON);
    const data = JSON.parse(jsonMatch[0]);
    Logger.log('Voucher parsed JSON: ' + JSON.stringify(data));

    // Normalizaciones defensivas (por si Claude devuelve string con "$" o "#")
    const codTransferencia = (data.codTransferencia || '').toString().replace(/^#/, '').trim();
    let monto = data.monto;
    if (typeof monto === 'string') {
      monto = parseFloat(monto.replace(/[^\d.,-]/g, '').replace(',', '.')) || 0;
    } else {
      monto = parseFloat(monto) || 0;
    }

    return ContentService.createTextOutput(JSON.stringify({
      ok:               true,
      codTransferencia: codTransferencia,
      monto:            monto,
      sender:           data.sender         || '',
      fechaPago:        data.fechaPago      || '',
      mensaje:          data.mensaje        || '',
      nombreCompleto:   data.nombreCompleto || null,
      email:            data.email          || null,
      telefono:         data.telefono       || null,
      _debug: {
        rawText:  text,
        parsed:   data,
        promptVersion: 'v3-2026-05-06'
      }
    })).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: e.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// Agrega un entry a la columna 32 (VouchersMeta) de `row`, preservando lo que
// ya hubiera. Formato: JSON array de {monto, cod, fecha, url}. Tolerante a
// celdas vacías o con JSON corrupto (arranca de cero en ese caso).
function _appendVoucherMeta(sheet, row, celdaActual, fileUrl, voucherMeta) {
  try {
    let lista = [];
    const raw = celdaActual ? celdaActual.toString().trim() : '';
    if (raw) {
      try {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed)) lista = parsed;
      } catch(_) { lista = []; }
    }
    const m = voucherMeta || {};
    const monto = (function() {
      const n = parseFloat(String(m.monto == null ? '' : m.monto).replace(/[^\d.]/g, ''));
      return isNaN(n) ? 0 : n;
    })();
    lista.push({
      monto: monto,
      cod:   (m.cod   || '').toString().trim(),
      fecha: (m.fecha || '').toString().trim(),
      url:   fileUrl || ''
    });
    sheet.getRange(row, 32).setValue(JSON.stringify(lista));
  } catch(e) {
    Logger.log('⚠ No se pudo persistir VouchersMeta: ' + e);
  }
}

function saveVoucherToDrive(reservation, imageBase64, mimeType, fileName, voucherMeta) {
  try {
    let folder;
    const folders = DriveApp.getFoldersByName(VOUCHER_FOLDER_NAME);
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(VOUCHER_FOLDER_NAME);

    const cabin    = reservation.cabin || 'cabin';
    const nombre   = (reservation.name || 'huesped').replace(/\s+/g, '_').slice(0, 20);
    const checkin  = (reservation.checkin || '').toString().slice(0, 10);
    const checkout = (reservation.checkout || '').toString().slice(0, 10);
    const ext      = mimeType.includes('png') ? '.png' : mimeType.includes('gif') ? '.gif' : '.jpg';
    const timestamp = Utilities.formatDate(new Date(), 'America/Panama', 'yyyyMMdd_HHmmss');
    const archName  = cabin + '_' + nombre + '_' + checkin + '_' + timestamp + ext;

    const nights = (() => {
      try {
        const a = new Date(checkin + 'T12:00:00');
        const b = new Date(checkout + 'T12:00:00');
        return Math.round((b - a) / 86400000);
      } catch(e) { return '?'; }
    })();
    const cabinName = { verde: 'Paseo por Las Nubes', azul: 'Portal hacia Las Nubes', lila: 'Puente entre Las Nubes' }[cabin] || cabin;

    const descripcion = [
      'Huésped: '   + (reservation.name     || ''),
      'Cabaña: '    + cabinName,
      'Entrada: '   + checkin,
      'Salida: '    + checkout,
      'Noches: '    + nights,
      'Monto: $'    + (reservation.amount   || ''),
      'Abono: $'    + (reservation.deposit  || ''),
      'Cód. Pago: ' + (reservation.confirmCode || reservation.id || ''),
      'Origen: '    + (reservation.origin   || ''),
      'Registrado: '+ Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd HH:mm')
    ].join('\n');

    const bytes = Utilities.base64Decode(imageBase64);
    const blob  = Utilities.newBlob(bytes, mimeType, archName);
    const file  = folder.createFile(blob);
    file.setDescription(descripcion);
    const fileUrl = file.getUrl();

    // Persistir el URL en la columna 26 (VoucherURL) de la fila correspondiente.
    // Si ya hay URL(s), agregar la nueva separada por '|' (multiples vouchers por reserva).
    try {
      const sheet = getOrCreateSheet();
      const data  = sheet.getDataRange().getValues();
      const stripId = id => id ? id.toString().replace(/^(airbnb_)+/, '') : '';
      const targetId = stripId(reservation.id);
      for (let i = 1; i < data.length; i++) {
        const rowId = stripId(data[i][0]);
        if (rowId && rowId === targetId) {
          const existing = data[i][25] ? data[i][25].toString().trim() : '';
          const merged   = existing ? (existing + '|' + fileUrl) : fileUrl;
          sheet.getRange(i + 1, 26).setValue(merged);
          // Col 32 (VouchersMeta) — un entry por voucher, en el MISMO orden que
          // las URLs de la col 26. Se escribe aquí (y solo aquí) para que ambas
          // listas no se puedan desalinear.
          _appendVoucherMeta(sheet, i + 1, data[i][31], fileUrl, voucherMeta);
          break;
        }
      }
    } catch(persistErr) {
      Logger.log('⚠ No se pudo persistir VoucherURL en hoja: ' + persistErr);
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true, fileUrl: fileUrl, fileName: archName })).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: e.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ═══════════════════════════════════════════════════════════
//  Pagos, estados y cancelaciones
// ═══════════════════════════════════════════════════════════

function actualizarEstadoPagoAirbnb() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const pagosSheet = ss.getSheetByName('Pagos');
  if (!pagosSheet) { Logger.log('⚠ Hoja Pagos no existe'); return; }

  const reservasSheet = getOrCreateSheet();
  const pagosData     = pagosSheet.getDataRange().getValues().slice(1);
  const resData       = reservasSheet.getDataRange().getValues();

  const pagoMap = {};
  pagosData.forEach(row => {
    const fechaPago  = row[0] ? (row[0] instanceof Date ? Utilities.formatDate(row[0], 'America/Panama', 'yyyy-MM-dd') : row[0].toString().slice(0,10)) : '';
    const codesStr   = row[5] ? row[5].toString() : '';
    const montosStr  = row[6] ? row[6].toString() : '';
    if (!codesStr) return;
    const codes  = codesStr.split(',');
    const montos = {};
    montosStr.split(',').forEach(pair => {
      const [c, m] = pair.split(':');
      if (c && m) montos[c.trim()] = parseFloat(m) || 0;
    });
    codes.forEach(code => {
      code = code.trim();
      if (!code) return;
      const montoEste = montos[code] || 0;
      if (!pagoMap[code]) {
        pagoMap[code] = { fechaPago, montoPagado: montoEste };
      } else {
        pagoMap[code].montoPagado = parseFloat((pagoMap[code].montoPagado + montoEste).toFixed(2));
        if (fechaPago > pagoMap[code].fechaPago) pagoMap[code].fechaPago = fechaPago;
      }
    });
  });

  let actualizados = 0;
  for (let i = 1; i < resData.length; i++) {
    const confirmCode    = resData[i][10] ? resData[i][10].toString().trim() : '';
    const origin         = resData[i][9]  ? resData[i][9].toString().trim()  : '';
    if (origin !== 'Airbnb' || !confirmCode) continue;
    const pago = pagoMap[confirmCode];
    if (!pago) continue;
    const estadoActual    = resData[i][20] ? resData[i][20].toString().trim() : '';
    const fechaPagoActual = resData[i][16] ? resData[i][16].toString().trim() : '';
    const montoActual     = parseFloat(resData[i][17]) || 0;
    // CANCELADA no se pisa. En una cancelación con penalización Airbnb IGUAL
    // paga, así que el código aparece en un payout y esta función volvía a
    // marcar PAGA en CADA corrida de syncCompleto, deshaciendo lo que había
    // escrito syncCancelacionesAirbnb. Aneea y Jennifer se marcaron CANCELADA a
    // mano y a las pocas horas ya estaban PAGA otra vez. Los datos del cobro sí
    // se guardan —la plata entró— pero el estado se preserva.
    const esCancelada = estadoActual === 'CANCELADA';
    // Antes bastaba con estar PAGA para saltear la fila. Pero Airbnb puede
    // pagar una misma reserva en VARIOS payouts (estadías largas): al llegar el
    // 2º cobro la reserva ya estaba PAGA y `MontoPagado` se quedaba con el
    // parcial (ej. LARS $47.10 de $271.24, GORAN $84.60 de $213.45). Ahora solo
    // se saltea si además el monto ya coincide con la suma acumulada.
    if ((estadoActual === 'PAGA' || esCancelada) && fechaPagoActual &&
        Math.abs(montoActual - pago.montoPagado) < 0.01) continue;
    const row = i + 1;
    reservasSheet.getRange(row, 17).setValue(pago.fechaPago);
    reservasSheet.getRange(row, 18).setValue(pago.montoPagado);
    if (!esCancelada) reservasSheet.getRange(row, 21).setValue('PAGA');
    actualizados++;
  }
  Logger.log('actualizarEstadoPagoAirbnb completado. Actualizadas: ' + actualizados);
}

function syncPayoutsYEstados() {
  syncAirbnbPayouts();
  actualizarEstadoPagoAirbnb();
}

function syncCancelacionesAirbnb() {
  const ss            = SpreadsheetApp.openById(SHEET_ID);
  const reservasSheet = getOrCreateSheet();

  let cancelSheet = ss.getSheetByName('Cancelaciones');
  if (!cancelSheet) {
    cancelSheet = ss.insertSheet('Cancelaciones');
    const headers = ['FechaCancelacion', 'EmailId', 'CodConfirmacion', 'Huesped', 'Cabana', 'FechasReserva', 'CanceladoPor', 'PenalizacionCobrada'];
    cancelSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    cancelSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  const cancelData   = cancelSheet.getDataRange().getValues().slice(1);
  const processedIds = new Set(cancelData.map(r => r[1].toString()));

  // La búsqueda vieja era subject:"Cancelada: reserva". En Gmail los dos puntos
  // son separador de operador y dentro de una frase entrecomillada el parseo es
  // impredecible, así que la query podía no matchear nada. Se busca por palabras
  // sueltas, que además cubre las variantes en inglés.
  const threads = GmailApp.search(
    'from:automated@airbnb.com (subject:Cancelada OR subject:cancelled OR subject:canceled) ' + SYNC_VENTANA
  );

  const resData = reservasSheet.getDataRange().getValues();
  const headers = resData[0].map(h => h.toString().trim());
  const colCod    = headers.indexOf('CodConfirmacion');
  const colEstado = headers.indexOf('EstadoPago');
  const colNombre = headers.indexOf('Nombre');
  const colCabana = headers.indexOf('Cabaña');

  if (colCod < 0 || colEstado < 0) { Logger.log('⚠ Columnas no encontradas'); return; }

  const codeToRow = {};
  for (let i = 1; i < resData.length; i++) {
    const code = resData[i][colCod] ? resData[i][colCod].toString().trim() : '';
    if (code) codeToRow[code] = i + 1;
  }

  let cancelados = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const msgId = msg.getId();
      if (processedIds.has(msgId)) return;

      const subject = msg.getSubject();
      const body    = msg.getPlainBody();

      let confirmCode = '';
      const m1 = subject.match(/reserva\s+([A-Z0-9]{8,12})\s/i);
      if (m1) { confirmCode = m1[1].toUpperCase(); }
      else {
        const m2 = subject.match(/HM[A-Z0-9]{6,10}/i);
        if (m2) { confirmCode = m2[0].toUpperCase(); }
        else {
          const m3 = body.match(/reserva\s+([A-Z0-9]{8,12})\s/i) || body.match(/HM[A-Z0-9]{6,10}/i);
          if (m3) confirmCode = (m3[1] || m3[0]).toUpperCase();
        }
      }

      if (!confirmCode) { Logger.log('⚠ No se pudo extraer código: ' + subject); return; }

      // OJO: la heurística vieja probaba /host/i contra el CUERPO, y todos los
      // enlaces de Airbnb contienen "hosting/reservations/details/...", así que
      // CUALQUIER cancelación del huésped se atribuía al anfitrión.
      // El header X-Template es inequívoco
      // (CANCELLATIONS_RESERVATION_CANCELED_BY_GUEST_TO_HOST); el cuerpo queda
      // como respaldo, ahora con frases que no aparecen en las URLs.
      let canceladoPor = '';
      try {
        const tpl = (msg.getRawContent().match(/^X-Template:\s*(.+)$/mi) || [])[1] || '';
        if      (/BY_GUEST/i.test(tpl))  canceladoPor = 'Huésped';
        else if (/BY_HOST/i.test(tpl))   canceladoPor = 'Anfitrión';
        else if (/BY_AIRBNB/i.test(tpl)) canceladoPor = 'Airbnb';
      } catch(_) {}
      // El texto plano de Airbnb viene con saltos de línea DENTRO de las frases
      // ("tu cobro se ha\nactualizado a $92,10"), así que cualquier regex de
      // varias palabras tiene que correr sobre el cuerpo con los espacios
      // normalizados. Si no, no matchea nada y falla en silencio.
      const bodyN = body.replace(/\s+/g, ' ');

      if (!canceladoPor) {
        if      (/\b(has cancelado|cancelaste|you cancelled|you canceled)\b/i.test(bodyN)) canceladoPor = 'Anfitrión';
        else if (/airbnb\s+(ha\s+)?cancel/i.test(bodyN))                                   canceladoPor = 'Airbnb';
        else canceladoPor = 'Huésped';
      }

      // "De acuerdo con tu política de cancelación, tu cobro se ha actualizado a
      // $92,10." Es la penalización que Airbnb igual paga — el dato que tuvimos
      // que deducir a mano para Aneea y Jennifer. Ojo con el decimal europeo.
      // El grupo tiene que terminar en dígito: con [\d.,]+ se comía el punto
      // final de la oración y "92,10." se interpretaba como 9210.
      const cobroM = bodyN.match(/cobro se ha actualizado a\s*\$\s*([\d.,]*\d)/i);
      const penalizacion = cobroM ? _montoAirbnbANumero(cobroM[1]) : null;

      const fechasMatch   = subject.match(/\(del?\s+(.+?)\)/i);
      const fechasReserva = fechasMatch ? fechasMatch[1] : '';
      const fechaCancelacion = Utilities.formatDate(msg.getDate(), 'America/Panama', 'yyyy-MM-dd');

      const rowNum  = codeToRow[confirmCode];
      let huesped = '', cabana = '';

      if (rowNum) {
        huesped = resData[rowNum-1][colNombre] ? resData[rowNum-1][colNombre].toString() : '';
        cabana  = colCabana >= 0 && resData[rowNum-1][colCabana] ? resData[rowNum-1][colCabana].toString() : '';
        const estadoActual = resData[rowNum-1][colEstado] ? resData[rowNum-1][colEstado].toString().trim() : '';
        if (estadoActual !== 'CANCELADA') {
          reservasSheet.getRange(rowNum, colEstado + 1).setValue('CANCELADA');
          cancelados++;
        }
        // Dejar el motivo y la penalización en Comentarios (col 23). NO se toca
        // monto ni neto a propósito: el ajuste de plata queda a criterio del
        // admin. Contabilidad ya cuenta lo de Airbnb desde la hoja Pagos, así
        // que el ingreso no se pierde por marcarla cancelada.
        const nota = '✕ Cancelada por ' + canceladoPor + ' el ' + fechaCancelacion
          + (penalizacion != null
              ? '. Airbnb pagó la penalización: $' + penalizacion.toFixed(2) + '.'
              : '. Sin penalización indicada en el email.');
        try {
          const cAct = (resData[rowNum-1][22] || '').toString().trim();
          if (cAct.indexOf('✕ Cancelada por') < 0) {
            reservasSheet.getRange(rowNum, 23).setValue(cAct ? cAct + '\n' + nota : nota);
          }
        } catch(cErr) { Logger.log('⚠ No se pudo escribir el comentario de cancelación: ' + cErr); }
        Logger.log('✕ ' + confirmCode + ' (' + huesped + ') → CANCELADA · ' + nota);
      } else {
        // La cancelación puede llegar antes de que la reserva se haya
        // sincronizado. Queda registrada en la hoja para no perderla.
        Logger.log('⚠ ' + confirmCode + ': cancelación registrada pero no hay fila con ese código todavía.');
      }

      cancelSheet.appendRow([fechaCancelacion, msgId, confirmCode, huesped, cabana, fechasReserva, canceladoPor,
        penalizacion != null ? penalizacion : '']);
      processedIds.add(msgId);
    });
  });

  Logger.log('Cancelaciones procesadas: ' + cancelados);
}

function syncCompleto() {
  syncAirbnbPayouts();
  actualizarEstadoPagoAirbnb();
  syncCancelacionesAirbnb();
  Logger.log('✓ syncCompleto finalizado');
}

// ═══════════════════════════════════════════════════════════
//  iCal export
// ═══════════════════════════════════════════════════════════

function getIcalForCabin(cabin) {
  try {
    const sheet = getOrCreateSheet();
    const data  = sheet.getDataRange().getValues();
    const rows  = data.slice(1).filter(r => r[0]);

    // El criterio va al revés que antes: TODO lo que ocupa la cabaña bloquea,
    // salvo lo que explícitamente no ocupa. La lista blanca vieja era
    // ['directa','cortesia','colaboracion'] y dejaba tres agujeros graves:
    //
    //  · `mantenimiento` y `personal` no se le informaban a Airbnb, así que
    //    Airbnb podía vender una noche que ya estaba tomada por trabajo o por
    //    uso propio.
    //  · `referido` tampoco, aunque ocupa la cabaña igual.
    //  · `airbnb` se excluía por parecer redundante —Airbnb ya conoce sus
    //    propias reservas— y ahí estuvo el daño real: una reserva de origen
    //    Airbnb cuyas FECHAS Airbnb no conoce. Mairanis no pudo venir el 19-21
    //    jun, se le honró el crédito para el 7-9 ago y esa fila quedó con origen
    //    Airbnb (el pago entró por ahí). El feed la omitió, Airbnb siguió
    //    mostrando agosto libre y Tesis reservó el 8 → doble booking.
    //
    // La asimetría decide: bloquear de más cuesta, como mucho, una noche vacía
    // que se libera en la próxima corrida; bloquear de menos cuesta dos huéspedes
    // pagando la misma noche.
    //
    // También se excluyen las CANCELADA, que antes seguían bloqueando el
    // calendario para siempre.
    const ORIGENES_SIN_FECHA = ['abierta'];   // reserva sin fechas: nada que bloquear
    const reservas = rows.filter(r => {
      const cabinCode = r[3] ? r[3].toString().toLowerCase() : '';
      const origen    = r[9] ? r[9].toString().toLowerCase().trim() : '';
      const estado    = r[20] ? r[20].toString().trim().toUpperCase() : '';
      if (!cabin ? false : cabinCode !== cabin.toLowerCase()) return false;
      if (estado === 'CANCELADA') return false;
      if (ORIGENES_SIN_FECHA.indexOf(origen) >= 0) return false;
      return true;
    });

    const now = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
    const CABIN_DISPLAY = {
      verde: 'Paseo por Las Nubes',
      azul:  'Portal hacia Las Nubes',
      lila:  'Puente entre Las Nubes'
    };
    const calName = CABIN_DISPLAY[cabin.toLowerCase()] || 'Las Nubes';

    const events = reservas.map(r => {
      const id       = r[0].toString();
      const nombre   = r[1] ? r[1].toString() : 'Huésped';
      const checkin  = r[4] instanceof Date ? Utilities.formatDate(r[4], 'America/Panama', 'yyyy-MM-dd') : r[4].toString().slice(0, 10);
      const checkout = r[5] instanceof Date ? Utilities.formatDate(r[5], 'America/Panama', 'yyyy-MM-dd') : r[5].toString().slice(0, 10);
      const origen   = r[9] ? r[9].toString() : '';
      if (!checkin || !checkout) return null;
      const dtStart = checkin.replace(/-/g, '')  + 'T190000Z';
      const dtEnd   = checkout.replace(/-/g, '') + 'T160000Z';
      return [
        'BEGIN:VEVENT',
        'DTSTART:' + dtStart,
        'DTEND:'   + dtEnd,
        'SUMMARY:Las Nubes - ' + nombre,
        'DESCRIPTION:Reserva Las Nubes\\n' + calName + '\\nOrigen: ' + origen,
        'UID:lasnubes-ical-' + id + '@lasnubes.pa',
        'DTSTAMP:' + now,
        'STATUS:CONFIRMED',
        'TRANSP:OPAQUE',
        'END:VEVENT'
      ].join('\r\n');
    }).filter(Boolean);

    const ical = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//Las Nubes//Reservas//ES',
      'CALSCALE:GREGORIAN',
      'METHOD:PUBLISH',
      'X-WR-CALNAME:Las Nubes - ' + calName,
      'X-WR-TIMEZONE:America/Panama',
      ...events,
      'END:VCALENDAR'
    ].join('\r\n');

    return ContentService.createTextOutput(ical).setMimeType(ContentService.MimeType.ICAL);
  } catch(e) {
    return ContentService.createTextOutput('BEGIN:VCALENDAR\r\nVERSION:2.0\r\nEND:VCALENDAR').setMimeType(ContentService.MimeType.ICAL);
  }
}

// ═══════════════════════════════════════════════════════════
//  Diagnóstico y reparación (ejecutar manualmente)
// ═══════════════════════════════════════════════════════════

function repararFechaPagoFaltante() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const pagosSheet    = ss.getSheetByName('Pagos');
  const reservasSheet = getOrCreateSheet();
  if (!pagosSheet) { Logger.log('⚠ Hoja Pagos no existe'); return; }

  const pagosData = pagosSheet.getDataRange().getValues().slice(1);
  const resData   = reservasSheet.getDataRange().getValues();

  const pagoMap = {};
  pagosData.forEach(row => {
    const fechaPago = row[0] ? (row[0] instanceof Date ? Utilities.formatDate(row[0], 'America/Panama', 'yyyy-MM-dd') : row[0].toString().slice(0,10)) : '';
    const codesStr  = row[5] ? row[5].toString() : '';
    const montosStr = row[6] ? row[6].toString() : '';
    if (!codesStr) return;
    const codes  = codesStr.split(',');
    const montos = {};
    montosStr.split(',').forEach(pair => {
      const [c, m] = pair.split(':');
      if (c && m) montos[c.trim()] = parseFloat(m) || 0;
    });
    codes.forEach(code => {
      code = code.trim();
      if (!code) return;
      if (!pagoMap[code]) pagoMap[code] = { fechaPago, montoPagado: montos[code] || 0 };
      else {
        pagoMap[code].montoPagado = parseFloat((pagoMap[code].montoPagado + (montos[code] || 0)).toFixed(2));
        if (fechaPago > pagoMap[code].fechaPago) pagoMap[code].fechaPago = fechaPago;
      }
    });
  });

  let reparadas = 0;
  for (let i = 1; i < resData.length; i++) {
    const confirmCode    = resData[i][10] ? resData[i][10].toString().trim() : '';
    const origin         = resData[i][9]  ? resData[i][9].toString().trim()  : '';
    const estadoActual   = resData[i][20] ? resData[i][20].toString().trim() : '';
    const fechaPagoActual = resData[i][16] ? resData[i][16].toString().trim() : '';
    if (origin !== 'Airbnb') continue;
    if (estadoActual !== 'PAGA') continue;
    if (fechaPagoActual && /^\d{4}-\d{2}-\d{2}$/.test(fechaPagoActual)) continue;
    const pago = pagoMap[confirmCode];
    if (!pago) continue;
    reservasSheet.getRange(i + 1, 17).setValue(pago.fechaPago);
    reservasSheet.getRange(i + 1, 18).setValue(pago.montoPagado);
    reparadas++;
  }
  Logger.log('Reparadas: ' + reparadas);
}

function diagnosticoPagosAirbnb() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const pagosSheet = ss.getSheetByName('Pagos');
  if (!pagosSheet) { Logger.log('⚠ Hoja Pagos no existe'); return; }

  const pagosData = pagosSheet.getDataRange().getValues().slice(1);
  const pagoMap   = {};

  pagosData.forEach((row, idx) => {
    const fechaPago  = row[0] ? (row[0] instanceof Date ? Utilities.formatDate(row[0], 'America/Panama', 'yyyy-MM-dd') : row[0].toString().slice(0,10)) : '';
    const codesStr   = row[5] ? row[5].toString() : '';
    const montosStr  = row[6] ? row[6].toString() : '';
    Logger.log('Pago #' + (idx+1) + ' | ' + fechaPago + ' | $' + row[2] + ' | Códigos: ' + (codesStr || '(ninguno)'));
    if (!codesStr) return;
    const montos = {};
    montosStr.split(',').forEach(pair => {
      const [c, m] = pair.split(':');
      if (c && m) montos[c.trim()] = parseFloat(m) || 0;
    });
    codesStr.split(',').forEach(code => {
      code = code.trim();
      if (!code) return;
      if (!pagoMap[code]) pagoMap[code] = { fechaPago, montoPagado: montos[code] || 0 };
      else {
        pagoMap[code].montoPagado += montos[code] || 0;
        if (fechaPago > pagoMap[code].fechaPago) pagoMap[code].fechaPago = fechaPago;
      }
    });
  });

  const resSheet = getOrCreateSheet();
  const resData  = resSheet.getDataRange().getValues().slice(1);
  const airbnbRes = resData.filter(r => r[9] && r[9].toString() === 'Airbnb');

  let matches = 0, sinMatch = 0;
  airbnbRes.forEach(r => {
    const code = r[10] ? r[10].toString().trim() : '';
    if (pagoMap[code]) { matches++; Logger.log('✓ ' + r[1] + ' | ' + code); }
    else { sinMatch++; Logger.log('✗ ' + r[1] + ' | ' + code); }
  });
  Logger.log('Con pago: ' + matches + ' | Sin pago: ' + sinMatch);
}

function corregirMontoPagadoPeter() {
  const reservasSheet = getOrCreateSheet();
  const data = reservasSheet.getDataRange().getValues();
  const TARGET_CODE = 'HMSKS5DS9Q';
  for (let i = 1; i < data.length; i++) {
    if (data[i][10] && data[i][10].toString().trim() === TARGET_CODE) {
      reservasSheet.getRange(i + 1, 17).setValue('2026-02-28');
      reservasSheet.getRange(i + 1, 18).setValue(261.90);
      Logger.log('✓ Peter corregido');
      return;
    }
  }
  Logger.log('⚠ Código no encontrado');
}