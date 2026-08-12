/**
 * CONCILIACIÓN BANCARIA
 * =====================
 *
 * Cruza los movimientos reales del banco contra lo que el sistema tiene
 * registrado (`Egresos`), y responde dos preguntas que hasta ahora no tenían
 * dónde contestarse:
 *
 *   1. ¿Qué salió del banco y NO está registrado como egreso?
 *   2. ¿Qué está registrado como egreso y NO salió del banco?
 *
 * Sin esto la contabilidad describe un negocio parecido al real pero no el
 * real: durante feb-jul 2026 el modelo daba ~$1,450/mes de sobrante mientras
 * las cuentas terminaban el mes en $7.25 (31-may) y $0.63 (30-jun). La
 * diferencia no era una sino muchas, y ninguna se podía ver sin cruzar.
 *
 * ─── Por qué NO reusa Prestamos.gs ─────────────────────────────────────
 *
 * La forma es parecida (un total, una lista de movimientos, saldo derivado),
 * pero `getAbonosPrestamoEnRango` RESTA los abonos de los egresos del período.
 * Eso es correcto para un adelanto a un colaborador y es veneno acá: la cuota
 * mensual del terreno es una salida real de plata, y netearla contra los
 * egresos inflaría la utilidad en silencio. Hoja aparte, a propósito.
 *
 * ─── Por qué NO reusa BancoGeneral_Module.gs ───────────────────────────
 *
 * Aquel módulo BORRA y reescribe `BG_Movimientos` en cada importación, así que
 * no acumula histórico ni deduplica, y clasifica con IA en categorías genéricas
 * de finanzas personales en vez de cruzar contra los egresos que ya existen.
 * Resuelve otro problema. Se deja donde está.
 */

const BANCO_SHEET = 'BancoMovimientos';
const BANCO_COLS = [
  'ID', 'Cuenta', 'Fecha', 'Descripcion', 'Monto', 'Saldo',
  'Huella', 'MatchTipo', 'MatchID', 'Revisado', 'Nota'
];

// Las cuentas NO se hardcodean: este repo es público. Van en Script Properties
// como CSV `clave:numero`, ej. `lasnubes:04999...,personal:04999...`.
// Sin configurar, la importación igual funciona (la cuenta se toma del payload),
// pero se pierde el guard de identidad de abajo.
const BANCO_CUENTAS_PROP = 'BANCO_CUENTAS';

function _bancoSheet() {
  return _hojaConCabecera(BANCO_SHEET, BANCO_COLS);
}

function _bancoCuentasConfig() {
  const raw = String(PropertiesService.getScriptProperties()
                     .getProperty(BANCO_CUENTAS_PROP) || '').trim();
  const out = {};
  if (!raw) return out;
  raw.split(',').forEach(par => {
    const i = par.indexOf(':');
    if (i <= 0) return;
    const clave = par.slice(0, i).trim().toLowerCase();
    const num = par.slice(i + 1).trim();
    if (clave && num) out[clave] = num;
  });
  return out;
}

function _bancoFecha(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}

function _bancoNum(v) {
  if (typeof v === 'number') return v;
  // Formato US (1,234.56) y europeo (1.234,56). Un replace(',','.') pelado
  // convierte 1,234.56 en 1.23 — ya pasó en el parser de Airbnb.
  let s = String(v == null ? '' : v).replace(/[^\d.,\-()]/g, '').trim();
  if (!s) return 0;
  let neg = /^\(.*\)$/.test(s) || s.indexOf('-') >= 0;
  s = s.replace(/[()\-]/g, '');
  const ultComa = s.lastIndexOf(','), ultPunto = s.lastIndexOf('.');
  if (ultComa >= 0 && ultPunto >= 0) {
    // El separador decimal es el que va más a la derecha.
    if (ultComa > ultPunto) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if (ultComa >= 0) {
    // Una coma sola: decimal si deja 1-2 dígitos detrás, si no es de miles.
    s = (s.length - ultComa - 1 <= 2) ? s.replace(',', '.') : s.replace(/,/g, '');
  }
  const n = parseFloat(s);
  if (!isFinite(n)) return 0;
  return neg ? -Math.abs(n) : n;
}

// La huella incluye el SALDO a propósito. Dos movimientos del mismo día por el
// mismo monto son indistinguibles por (fecha, monto, descripción) —pasa seguido
// con los Yappy—, pero su saldo corriente siempre difiere. Sin el saldo, el
// segundo se descartaba como duplicado y el mes quedaba corto.
function _bancoHuella(cuenta, fecha, desc, monto, saldo) {
  return [
    String(cuenta || '').toLowerCase(),
    fecha,
    Math.round(_bancoNum(monto) * 100),
    Math.round(_bancoNum(saldo) * 100),
    String(desc || '').replace(/\s+/g, ' ').trim().toUpperCase().slice(0, 60)
  ].join('|');
}

/**
 * Guard de identidad de cuenta.
 *
 * Un estado de cuenta NUNCA se lista a sí mismo como contraparte. Si las filas
 * que dicen ser de la cuenta X mencionan el número de X, la etiqueta está mal.
 *
 * No es hipotético: los dos exports de agosto-2026 traían las cuentas
 * INVERTIDAS entre sí, y eso hizo contar 8 cuotas del terreno donde había 6 y
 * atribuirle a un tercero $8,113 de aportes que eran transferencias de la
 * propia cuenta hacia sí misma. Un error de etiqueta que se propaga a todas las
 * conclusiones, y que esta comprobación de dos líneas atrapa al importar.
 */
function _bancoVerificarIdentidad(cuenta, filas) {
  const cfg = _bancoCuentasConfig();
  const propio = cfg[String(cuenta || '').toLowerCase()];
  if (!propio) return null;                      // sin config no hay nada que verificar
  const hits = filas.filter(f => String(f.desc || '').indexOf(propio) >= 0).length;
  if (hits === 0) return null;
  return 'La cuenta "' + cuenta + '" se menciona a sí misma como contraparte en '
       + hits + ' de ' + filas.length + ' movimientos. Es casi seguro que el '
       + 'archivo corresponde a OTRA cuenta. Verifica la etiqueta antes de importar.';
}

// ─── Lectura del XLSX de Banco General ──────────────────────────
//
// El export de BG ("Últimos movimientos", hoja `BGPExcelReport`) trae:
//
//     fila 3  Cuenta:VISADEBITO 04-99-99-XXXXXX-X
//     fila 7  Fecha | Referencia | Descripción | Monto | Saldo total
//     fila 8+ los movimientos, del más nuevo al más viejo
//
// Se parsea en el servidor y no en el navegador a propósito: `dashboard.html`
// no carga ni un solo script externo, y meterle un parser de XLSX de un CDN al
// panel de administración por una pantalla es un mal negocio. Drive convierte
// el archivo y `SpreadsheetApp` lo lee, que es lo que ya hay disponible.

/**
 * El archivo dice de qué cuenta es. Esto es MEJOR que el guard de identidad:
 * el guard detecta la etiqueta mal puesta, esto hace que no haya etiqueta que
 * poner. Devuelve solo dígitos ('04-99-99-123456-7' → '0499991234567').
 */
function _bancoCuentaDelArchivo(filas) {
  for (let i = 0; i < Math.min(filas.length, 12); i++) {
    const txt = filas[i].join(' ');
    const m = txt.match(/Cuenta\s*:?\s*[A-Za-zÁÉÍÓÚÑ ]*([\d\-]{10,})/);
    if (m) {
      const dig = m[1].replace(/\D/g, '');
      if (dig.length >= 10) return dig;
    }
  }
  return '';
}

// De número de cuenta a la clave configurada ('lasnubes' / 'personal').
function _bancoClaveDeNumero(numero) {
  if (!numero) return '';
  const cfg = _bancoCuentasConfig();
  const claves = Object.keys(cfg);
  for (let i = 0; i < claves.length; i++) {
    if (String(cfg[claves[i]]).replace(/\D/g, '') === numero) return claves[i];
  }
  return '';
}

/**
 * Convierte el XLSX a hoja de Google, lo lee y borra el temporal.
 *
 * Las columnas se ubican por NOMBRE de encabezado, no por posición: si BG
 * agrega una columna, un índice fijo empieza a leer el campo de al lado sin
 * avisar. El `finally` borra el temporal aunque el parseo explote — si no,
 * cada import fallido deja basura en Drive para siempre.
 */
function _bancoParsearXlsx(base64, mimeType) {
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(
    bytes, mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'bg-import.xlsx');

  let tempId = '';
  try {
    const creado = Drive.Files.create(
      { name: 'BG_Temp_' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS }, blob);
    tempId = creado.id;
    const valores = SpreadsheetApp.openById(tempId).getSheets()[0]
                    .getDataRange().getValues()
                    .map(r => r.map(c => (c == null ? '' : c)));

    const cuentaNum = _bancoCuentaDelArchivo(
      valores.map(r => r.map(c => String(c))));

    // Encabezado = la primera fila que tiene Fecha y Monto.
    let hdr = -1, col = {};
    for (let i = 0; i < valores.length; i++) {
      const norm = valores[i].map(c => String(c).normalize('NFD')
                     .replace(/[\u0300-\u036f]/g, '').toLowerCase().trim());
      if (norm.indexOf('fecha') >= 0 && norm.indexOf('monto') >= 0) {
        hdr = i;
        norm.forEach((h, j) => {
          if (h === 'fecha') col.fecha = j;
          else if (h === 'monto') col.monto = j;
          else if (h.indexOf('descrip') === 0) col.desc = j;
          else if (h.indexOf('saldo') === 0) col.saldo = j;
          else if (h.indexOf('referencia') === 0) col.ref = j;
        });
        break;
      }
    }
    if (hdr < 0) throw new Error('No encontré la fila de encabezados (Fecha / Monto). ¿Es el export de Banco General?');
    if (col.desc == null || col.saldo == null) {
      throw new Error('El archivo no tiene columnas Descripción y/o Saldo.');
    }

    const filas = [];
    for (let i = hdr + 1; i < valores.length; i++) {
      const r = valores[i];
      // La celda Fecha trae hora, pero es el timestamp de EXPORTACIÓN (idéntico
      // en todas las filas del archivo), no la hora del movimiento. Se descarta.
      const f = _bancoFecha(r[col.fecha]);
      const d = String(r[col.desc] || '').replace(/\s+/g, ' ').trim();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(f) || !d) continue;
      filas.push({ fecha: f, desc: d, monto: _bancoNum(r[col.monto]), saldo: _bancoNum(r[col.saldo]) });
    }
    return { filas: filas, cuentaNum: cuentaNum, clave: _bancoClaveDeNumero(cuentaNum) };
  } finally {
    if (tempId) { try { Drive.Files.remove(tempId); } catch (_) {} }
  }
}

/**
 * Sube un XLSX de BG y lo importa. Con `dryRun` no escribe nada: devuelve
 * cuántos entrarían nuevos y cuántos ya están, para poder mirar antes de
 * aplicar.
 */
function importarBancoXlsx(payload) {
  if (!payload || !payload.base64) return { ok: false, error: 'Falta el archivo' };
  let p;
  try {
    p = _bancoParsearXlsx(payload.base64, payload.mimeType);
  } catch (e) {
    return { ok: false, error: e.message };
  }
  if (!p.filas.length) return { ok: false, error: 'El archivo no tenía movimientos legibles' };

  // La cuenta la manda el archivo. Solo se cae a la del payload si el export no
  // la trae o si todavía no está mapeada en BANCO_CUENTAS.
  const cuenta = p.clave || String(payload.cuenta || '').trim().toLowerCase();
  if (!cuenta) {
    return {
      ok: false, sinMapear: true, cuentaNum: p.cuentaNum, filas: p.filas.length,
      error: 'El archivo es de la cuenta ' + (p.cuentaNum || '(no declarada)')
           + ', que no está en BANCO_CUENTAS. Agrégala como `clave:' + p.cuentaNum + '`.'
    };
  }
  if (payload.cuenta && p.clave && String(payload.cuenta).toLowerCase() !== p.clave) {
    return {
      ok: false, necesitaConfirmacion: true, cuentaDetectada: p.clave,
      error: 'Elegiste "' + payload.cuenta + '" pero el archivo declara ser de "'
           + p.clave + '" (' + p.cuentaNum + '). Se importa como "' + p.clave + '".'
    };
  }

  const desde = p.filas.reduce((a, f) => (!a || f.fecha < a) ? f.fecha : a, '');
  const hasta = p.filas.reduce((a, f) => (f.fecha > a) ? f.fecha : a, '');

  if (payload.dryRun) {
    const sheet = _bancoSheet();
    const ya = {};
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).getValues()
           .forEach(r => { if (r[0]) ya[String(r[0])] = true; });
    }
    const vistas = {};
    let nuevos = 0, dup = 0;
    p.filas.forEach(f => {
      const h = _bancoHuella(cuenta, f.fecha, f.desc, f.monto, f.saldo);
      if (ya[h] || vistas[h]) dup++; else { vistas[h] = true; nuevos++; }
    });
    return {
      ok: true, dryRun: true, cuenta: cuenta, cuentaNum: p.cuentaNum,
      total: p.filas.length, importados: nuevos, duplicados: dup,
      desde: desde, hasta: hasta
    };
  }

  const r = importarMovimientosBanco({ cuenta: cuenta, filas: p.filas, force: payload.force });
  if (r.ok) { r.cuentaNum = p.cuentaNum; r.desde = desde; r.hasta = hasta; }
  return r;
}

// ─── Importación ────────────────────────────────────────────────
//
// Acumula. No borra lo anterior: el histórico es lo que permite comparar meses,
// y un import que reescribe la hoja pierde toda la conciliación ya hecha a mano.
// Idempotente por huella, así que re-subir el mismo archivo no duplica nada.
function importarMovimientosBanco(payload) {
  const cuenta = String((payload && payload.cuenta) || '').trim().toLowerCase();
  if (!cuenta) return { ok: false, error: 'Falta la cuenta' };
  const filas = (payload && payload.filas) || [];
  if (!filas.length) return { ok: false, error: 'No hay movimientos que importar' };

  const norm = filas.map(f => ({
    fecha: _bancoFecha(f.fecha),
    desc: String(f.desc || f.descripcion || '').replace(/\s+/g, ' ').trim(),
    monto: _bancoNum(f.monto),
    saldo: _bancoNum(f.saldo)
  })).filter(f => f.fecha && f.desc);

  if (!norm.length) return { ok: false, error: 'Ninguna fila tenía fecha y descripción' };

  const aviso = _bancoVerificarIdentidad(cuenta, norm);
  if (aviso && !(payload && payload.force)) {
    return { ok: false, error: aviso, necesitaConfirmacion: true };
  }

  const sheet = _bancoSheet();
  const existentes = {};
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).getValues()
         .forEach(r => { if (r[0]) existentes[String(r[0])] = true; });
  }

  const nuevas = [];
  let dup = 0;
  norm.forEach((f, i) => {
    const huella = _bancoHuella(cuenta, f.fecha, f.desc, f.monto, f.saldo);
    if (existentes[huella]) { dup++; return; }
    existentes[huella] = true;                   // también deduplica dentro del mismo archivo
    nuevas.push([
      'BM-' + Date.now() + '-' + i, cuenta, f.fecha, f.desc, f.monto, f.saldo,
      huella, '', '', '', ''
    ]);
  });

  if (nuevas.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, nuevas.length, BANCO_COLS.length)
         .setValues(nuevas);
  }

  logDebugEntry('banco-import', { cuenta: cuenta, nuevas: nuevas.length, dup: dup });
  return {
    ok: true, cuenta: cuenta, importados: nuevas.length, duplicados: dup,
    total: norm.length, aviso: aviso || null
  };
}

// ─── Lectura ────────────────────────────────────────────────────

function _bancoLeer(desde, hasta) {
  const sheet = _bancoSheet();
  if (sheet.getLastRow() < 2) return [];
  return sheet.getDataRange().getValues().slice(1)
    .filter(r => r[0])
    .map((r, i) => ({
      id: String(r[0]), cuenta: String(r[1] || ''), fecha: _bancoFecha(r[2]),
      desc: String(r[3] || ''), monto: _bancoNum(r[4]), saldo: _bancoNum(r[5]),
      huella: String(r[6] || ''), matchTipo: String(r[7] || ''),
      matchId: String(r[8] || ''), revisado: String(r[9] || ''),
      nota: String(r[10] || ''), fila: i + 2
    }))
    .filter(m => (!desde || m.fecha >= desde) && (!hasta || m.fecha <= hasta));
}

// ─── Cruce contra Egresos ───────────────────────────────────────
//
// Solo se cruzan los DÉBITOS: un egreso es plata que salió. Los créditos
// (ingresos) se cruzan contra reservas y payouts, que es otro problema y tiene
// su propia función.

const _CONC_STOP = {
  'BANCA': 1, 'MOVIL': 1, 'LINEA': 1, 'TRANSFERENCIA': 1, 'YAPPY': 1, 'BG': 1,
  'A': 1, 'DE': 1, 'POR': 1, 'EL': 1, 'LA': 1, 'LOS': 1, 'LAS': 1, 'DEL': 1,
  'ENTRE': 1, 'CUENTAS': 1, 'CUENTA': 1, 'PAGO': 1, 'ACH': 1, 'TDC': 1
};

function _concTokens(s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase().replace(/[^A-Z0-9 ]/g, ' ')
    .split(/\s+/)
    .filter(t => t.length >= 4 && !_CONC_STOP[t]);
}

/**
 * Puntaje de un candidato. El monto tiene que calzar exacto: un egreso por otro
 * monto no es el mismo egreso, y aflojar eso llena el reporte de falsos
 * positivos que después hay que deshacer a mano.
 *
 * Devuelve null si no es candidato.
 */
function _concPuntaje(mov, egr) {
  if (Math.abs(Math.abs(mov.monto) - Math.abs(egr.amount)) > 0.01) return null;
  const dias = Math.abs(_concDias(mov.fecha, egr.date));
  if (dias > 5) return null;                     // más lejos, el monto igual es coincidencia

  let p = 100 - dias * 8;                        // la cercanía manda
  const tm = _concTokens(mov.desc);
  const te = _concTokens(egr.desc + ' ' + (egr.proveedor || ''));
  let comunes = 0;
  tm.forEach(t => { if (te.indexOf(t) >= 0) comunes++; });
  p += comunes * 15;
  return { puntaje: p, dias: dias, tokens: comunes };
}

function _concDias(a, b) {
  const pa = String(a).split('-'), pb = String(b).split('-');
  if (pa.length !== 3 || pb.length !== 3) return 9999;
  const da = Date.UTC(+pa[0], +pa[1] - 1, +pa[2]);
  const db = Date.UTC(+pb[0], +pb[1] - 1, +pb[2]);
  return Math.round((da - db) / 86400000);
}

/**
 * Corre el cruce y devuelve el resultado SIN escribir. Se persiste aparte
 * (`guardarConciliacion`) para que el admin pueda revisar antes de fijar los
 * matches: un cruce automático que se auto-guarda es imposible de auditar.
 *
 * El emparejamiento es 1-a-1 y greedy por puntaje: una factura del banco no
 * puede pagar dos egresos, ni un egreso aparecer dos veces en el banco.
 */
function conciliarConEgresos(desde, hasta) {
  const movs = _bancoLeer(desde, hasta).filter(m => m.monto < 0 && !m.matchId);

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const eSheet = ss.getSheetByName('Egresos');
  const egresos = [];
  if (eSheet && eSheet.getLastRow() > 1) {
    const data = eSheet.getDataRange().getValues();
    const head = data[0].map(h => String(h).toLowerCase().trim());
    const iId = head.indexOf('id'), iF = head.indexOf('fecha'),
          iD = head.indexOf('descripcion'), iM = head.indexOf('monto'),
          iC = head.indexOf('categoria'), iP = head.indexOf('proveedor');
    data.slice(1).forEach(r => {
      if (!r[iF] || !r[iM]) return;
      const f = _bancoFecha(r[iF]);
      if (desde && f < desde) return;
      if (hasta && f > hasta) return;
      egresos.push({
        id: String(r[iId] || ''), date: f, desc: String(r[iD] || ''),
        amount: _bancoNum(r[iM]), cat: String(r[iC] || ''),
        proveedor: iP >= 0 ? String(r[iP] || '') : ''
      });
    });
  }

  // Todos los pares candidatos, mejor puntaje primero.
  const pares = [];
  movs.forEach(m => egresos.forEach(e => {
    const p = _concPuntaje(m, e);
    if (p) pares.push({ mov: m, egr: e, puntaje: p.puntaje, dias: p.dias, tokens: p.tokens });
  }));
  pares.sort((a, b) => b.puntaje - a.puntaje);

  const movUsado = {}, egrUsado = {}, matches = [];
  pares.forEach(par => {
    if (movUsado[par.mov.id] || egrUsado[par.egr.id]) return;
    movUsado[par.mov.id] = true; egrUsado[par.egr.id] = true;
    matches.push(par);
  });

  const sinRegistrar = movs.filter(m => !movUsado[m.id]);
  const sinBanco = egresos.filter(e => !egrUsado[e.id]);

  return {
    ok: true,
    matches: matches.map(p => ({
      movId: p.mov.id, movFecha: p.mov.fecha, movDesc: p.mov.desc, movMonto: p.mov.monto,
      egresoId: p.egr.id, egresoFecha: p.egr.date, egresoDesc: p.egr.desc,
      egresoCat: p.egr.cat, puntaje: p.puntaje, dias: p.dias, tokens: p.tokens,
      // Un match con 0 tokens en común y varios días de diferencia calza solo
      // por monto: es el que hay que mirar antes de aceptar.
      dudoso: (p.tokens === 0 && p.dias >= 2)
    })),
    // Plata que SALIÓ del banco y no está en la contabilidad. Es el hallazgo
    // que importa: son gastos reales que nadie registró.
    sinRegistrar: sinRegistrar.map(m => ({
      movId: m.id, fecha: m.fecha, desc: m.desc, monto: m.monto, cuenta: m.cuenta
    })),
    // Registrados pero sin respaldo en el banco. O se pagaron en efectivo / por
    // otra cuenta, o están cargados de más.
    sinBanco: sinBanco.map(e => ({
      egresoId: e.id, fecha: e.date, desc: e.desc, monto: e.amount, cat: e.cat
    })),
    resumen: {
      movimientos: movs.length, egresos: egresos.length,
      conciliados: matches.length,
      montoSinRegistrar: +sinRegistrar.reduce((s, m) => s + Math.abs(m.monto), 0).toFixed(2),
      montoSinBanco: +sinBanco.reduce((s, e) => s + Math.abs(e.amount), 0).toFixed(2)
    }
  };
}

// Fija los matches aceptados. Recibe [{movId, egresoId}].
function guardarConciliacion(payload) {
  const lista = (payload && payload.matches) || [];
  if (!lista.length) return { ok: true, guardados: 0 };
  const sheet = _bancoSheet();
  const movs = _bancoLeer(null, null);
  const porId = {};
  movs.forEach(m => { porId[m.id] = m; });

  let n = 0;
  lista.forEach(x => {
    const m = porId[String(x.movId || '')];
    if (!m) return;
    sheet.getRange(m.fila, 8, 1, 2).setValues([['egreso', String(x.egresoId || '')]]);
    n++;
  });
  logDebugEntry('banco-conciliado', { n: n });
  return { ok: true, guardados: n };
}

// ─── Resumen por cuenta ─────────────────────────────────────────
//
// Responde "¿cuánta plata hay y por dónde entra?", que es la pregunta que no
// tenía panel. El dato que destapó el problema de fondo: durante feb-jul 2026 el
// 41% de los Yappy de huéspedes ($10,378 de $25,497) caía en la cuenta personal
// y no en la de Las Nubes, así que ninguna métrica del dashboard cuadraba con
// el banco.
/**
 * Cuál fue el ÚLTIMO movimiento de un día.
 *
 * La celda Fecha del export trae hora, pero es el timestamp de exportación
 * —idéntico en todas las filas—, así que dentro de un mismo día no hay con qué
 * ordenar. Confiar en el orden de lectura tampoco sirve: depende de en qué
 * orden se hayan importado los archivos, y dos exports con rangos distintos
 * dejan el mismo día en posiciones distintas.
 *
 * El orden se reconstruye con la CADENA DE SALDOS: para un movimiento `r`, el
 * saldo justo antes de él es `r.saldo - r.monto`. El último del día es el único
 * cuyo saldo no es el "saldo previo" de ningún otro del grupo. Verificado
 * contra los extractos reales: la cadena encaja en 1023 de 1023 filas.
 *
 * Si el grupo es ambiguo (saldos repetidos), se cae al orden de lectura en vez
 * de inventar un ganador.
 */
function _bancoUltimoDelDia(rows) {
  if (rows.length === 1) return rows[0];
  const previos = {};
  rows.forEach(r => { previos[Math.round((r.saldo - r.monto) * 100)] = true; });
  const cands = rows.filter(r => !previos[Math.round(r.saldo * 100)]);
  return cands.length === 1 ? cands[0] : rows[rows.length - 1];
}

function getCuentasResumen(desde, hasta) {
  const movs = _bancoLeer(desde, hasta);
  const porCuenta = {};

  movs.forEach(m => {
    const c = m.cuenta || '(sin cuenta)';
    if (!porCuenta[c]) {
      porCuenta[c] = {
        cuenta: c, entra: 0, sale: 0, movimientos: 0,
        yappyHuespedes: 0, yappyCount: 0, saldoFinal: null, ultimaFecha: '',
        meses: {}, dias: {}
      };
    }
    const a = porCuenta[c];
    a.movimientos++;
    if (m.monto > 0) a.entra += m.monto; else a.sale += m.monto;

    // Un Yappy entrante de una persona es, casi siempre, un huésped pagando.
    if (m.monto > 0 && /^YAPPY\s+BG\s+DE/i.test(m.desc)) {
      a.yappyHuespedes += m.monto; a.yappyCount++;
    }

    (a.dias[m.fecha] = a.dias[m.fecha] || []).push(m);

    const mes = m.fecha.slice(0, 7);
    if (!a.meses[mes]) a.meses[mes] = { mes: mes, entra: 0, sale: 0, saldoFinal: null, ultima: '' };
    const mm = a.meses[mes];
    if (m.monto > 0) mm.entra += m.monto; else mm.sale += m.monto;
    if (m.fecha > mm.ultima) mm.ultima = m.fecha;
    if (m.fecha > a.ultimaFecha) a.ultimaFecha = m.fecha;
  });

  const cuentas = Object.keys(porCuenta).map(k => {
    const a = porCuenta[k];
    a.entra = +a.entra.toFixed(2); a.sale = +a.sale.toFixed(2);
    a.neto = +(a.entra + a.sale).toFixed(2);
    a.yappyHuespedes = +a.yappyHuespedes.toFixed(2);

    // El saldo de cierre es el del último movimiento del último día con
    // movimientos — no el de la última fila leída.
    const ultDia = a.dias[a.ultimaFecha];
    a.saldoFinal = ultDia ? _bancoUltimoDelDia(ultDia).saldo : null;

    a.meses = Object.keys(a.meses).sort().map(mk => {
      const mm = a.meses[mk];
      mm.entra = +mm.entra.toFixed(2); mm.sale = +mm.sale.toFixed(2);
      mm.neto = +(mm.entra + mm.sale).toFixed(2);
      const d = a.dias[mm.ultima];
      mm.saldoFinal = d ? _bancoUltimoDelDia(d).saldo : null;
      return mm;
    });
    delete a.dias;
    return a;
  }).sort((a, b) => b.entra - a.entra);

  const totalYappy = cuentas.reduce((s, c) => s + c.yappyHuespedes, 0);
  const ln = cuentas.filter(c => c.cuenta === 'lasnubes')
                    .reduce((s, c) => s + c.yappyHuespedes, 0);

  return {
    ok: true,
    cuentas: cuentas,
    // Cuánto del ingreso de huéspedes NO entró por la cuenta del negocio.
    // Mientras esto no sea 0, el dashboard y el banco van a discrepar siempre.
    mezcla: {
      totalYappy: +totalYappy.toFixed(2),
      enLasNubes: +ln.toFixed(2),
      fueraDeLasNubes: +(totalYappy - ln).toFixed(2),
      pctFuera: totalYappy > 0 ? +(((totalYappy - ln) / totalYappy) * 100).toFixed(1) : 0
    },
    saldoTotal: +cuentas.reduce((s, c) => s + (c.saldoFinal || 0), 0).toFixed(2)
  };
}

// ─── Atajos para el editor ──────────────────────────────────────
//
// El editor de Apps Script corre las funciones SIN argumentos, así que toda
// función con parámetros necesita un runner sin parámetros. Mismo patrón que
// la sección equivalente de Cleanup.gs.
//
// Sirven para cargar los extractos ANTES de que exista la pantalla: se suben
// los .xlsx a Drive, se pegan los IDs acá y se corre. La cuenta la detecta el
// propio archivo, así que no hay nada más que configurar.
//
// Cómo sacar el ID: abrir el archivo en Drive → la URL trae /d/<ID>/view.

const IMPORTAR_BANCO_FILE_IDS = [
  // 'ID_DEL_XLSX_1',
  // 'ID_DEL_XLSX_2',
];

/**
 * Importa uno o varios .xlsx de Banco General que ya estén en Drive.
 * `dryRun` en true no escribe: solo dice qué entraría.
 */
function importarBancoDesdeDrive(fileIds, dryRun) {
  const ids = (fileIds && fileIds.length) ? fileIds : IMPORTAR_BANCO_FILE_IDS;
  if (!ids.length) {
    Logger.log('No hay IDs. Pegá los IDs de Drive en IMPORTAR_BANCO_FILE_IDS.');
    return;
  }
  ids.forEach(id => {
    let nombre = id;
    try {
      const file = DriveApp.getFileById(String(id).trim());
      nombre = file.getName();
      const r = importarBancoXlsx({
        base64: Utilities.base64Encode(file.getBlob().getBytes()),
        mimeType: file.getMimeType(),
        dryRun: !!dryRun
      });
      if (!r.ok) {
        Logger.log('✗ ' + nombre + ' → ' + r.error);
      } else {
        Logger.log('✓ ' + nombre + ' → cuenta "' + r.cuenta + '" (' + r.cuentaNum + ')  '
                 + r.importados + ' nuevos, ' + r.duplicados + ' ya estaban'
                 + (r.desde ? '  [' + r.desde + ' .. ' + r.hasta + ']' : '')
                 + (dryRun ? '   (SIMULACIÓN, no se escribió nada)' : ''));
      }
    } catch (e) {
      Logger.log('✗ ' + nombre + ' → ' + e.message);
    }
  });
  if (!dryRun) Logger.log('\n' + JSON.stringify(getCuentasResumen(null, null).cuentas
    .map(c => ({ cuenta: c.cuenta, movs: c.movimientos, saldo: c.saldoFinal })), null, 2));
}

// Simulación: no escribe. Correr esta PRIMERO.
function importarBancoReporte() { importarBancoDesdeDrive(null, true); }
// Escribe de verdad.
function importarBancoESCRIBIR() { importarBancoDesdeDrive(null, false); }

// Test desde el editor.
function _testConciliacion() {
  Logger.log(JSON.stringify(getCuentasResumen(null, null), null, 2));
}
