// ═══════════════════════════════════════════════════════════
//  Limpieza puntual de egresos duplicados (detectados a mano)
//
//  Duplicados identificados (facturas idénticas cargadas 2 veces):
//   1) PriceSmart — 16 items idénticos, cargados con 1.6s de diferencia
//      pero con fecha distinta (5 may vs 5 jun). Total $187.78 c/u.
//      Se conserva el set del 5-may; se borra el del 5-jun.
//   2) Do It Center (7-ago) — 5 items idénticos, cargados ~110s aparte.
//      Se conserva el primero; se borra el segundo.
//   3) Rafael Rodriguez limpieza (16-mar) — MISMO código de referencia
//      [WLPTY-41210925], $150. Se conserva el primero; se borra el segundo.
//
//  NO incluido (revisar a mano): Rafael 9-feb dos veces $165 "limpieza,
//  gasolina y kit de fogata" — fotos distintas, podría ser legítimo.
//
//  USO: correr primero `previewEgresosDuplicados()` para ver qué se borraría,
//  verificar en el log, y luego `borrarEgresosDuplicados()` para borrar.
// ═══════════════════════════════════════════════════════════

// IDs de las COPIAS a borrar (se conserva la otra copia de cada par).
function _idsDuplicadosABorrar() {
  const ids = [];
  // 1) PriceSmart set del 5-jun (se conserva el del 5-may, egr_1784770562996_*)
  for (let i = 0; i <= 15; i++) ids.push('egr_1784770564618_' + i);
  // 2) Do It Center 2do set (se conserva egr_1784764917414_*)
  for (let i = 0; i <= 4; i++)  ids.push('egr_1784765027583_' + i);
  // 3) Rafael limpieza 16-mar duplicado (se conserva egr_1773696796311_0)
  ids.push('egr_1773950891537_0');
  return ids;
}

// Muestra en el log las filas que se borrarían (sin borrar nada). Correr esto
// PRIMERO para verificar.
function previewEgresosDuplicados() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Egresos');
  if (!sheet) { Logger.log('⚠️ No existe la hoja Egresos'); return; }
  const data  = sheet.getDataRange().getValues();
  const idSet = {}; _idsDuplicadosABorrar().forEach(id => idSet[id] = true);

  let encontrados = 0, suma = 0;
  Logger.log('─── VISTA PREVIA: filas a borrar ───');
  for (let i = 1; i < data.length; i++) {
    const id = (data[i][0] || '').toString();
    if (idSet[id]) {
      encontrados++;
      const monto = parseFloat(data[i][3]) || 0;
      suma += monto;
      Logger.log('• ' + id + ' | ' + data[i][1] + ' | ' + data[i][2] + ' | $' + monto.toFixed(2) + ' | ' + data[i][6]);
    }
  }
  Logger.log('─── ' + encontrados + ' de ' + _idsDuplicadosABorrar().length +
             ' IDs encontrados · total a eliminar $' + suma.toFixed(2) + ' ───');
  Logger.log('Si se ve bien, corre borrarEgresosDuplicados().');
  return { encontrados: encontrados, total: suma };
}

// Borra las filas duplicadas de la hoja Egresos (de abajo hacia arriba para no
// desfasar índices). Correr previewEgresosDuplicados() antes.
function borrarEgresosDuplicados() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Egresos');
  if (!sheet) { Logger.log('⚠️ No existe la hoja Egresos'); return; }
  const data  = sheet.getDataRange().getValues();
  const idSet = {}; _idsDuplicadosABorrar().forEach(id => idSet[id] = true);

  const filas = []; // números de fila (1-based) a borrar
  let suma = 0;
  for (let i = 1; i < data.length; i++) {
    const id = (data[i][0] || '').toString();
    if (idSet[id]) { filas.push(i + 1); suma += (parseFloat(data[i][3]) || 0); }
  }
  filas.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));
  Logger.log('✓ ' + filas.length + ' filas duplicadas eliminadas · $' + suma.toFixed(2) + ' recuperados.');
  return { eliminadas: filas.length, total: suma };
}


// ═══════════════════════════════════════════════════════════
//  CORRECCIONES DE DATOS EN LA HOJA `Reservas` (auditoría jul-2026)
//
//  A) Multi-cabaña con voucher duplicado
//     Antes del fix, `_saveMultiCabinReservation` copiaba el voucher COMPLETO
//     en cada reserva hermana. Un pago único de $150 repartido en 2 cabañas
//     quedaba como MontoVoucher=$150 en AMBAS filas → se leía como $300
//     cobrados (inflaba "recibido" en Ingresos y tapaba saldos pendientes).
//     Este fix reparte el voucher en proporción al Monto de cada hermana.
//
//  B) Airbnb sin monto registrado
//     Hay reservas Airbnb con Monto/Neto en 0. Para las que YA tienen
//     MontoPagado (el depósito real de Airbnb) se reconstruye el bruto:
//         bruto = MontoPagado / (1 - tasa)      tasa: 3% si checkin < 24-dic-2025, 15.5% después
//     Verificado contra el histórico: la mediana MontoPagado/Neto es 0.845,
//     exactamente 1 - 15.5%, y los 7 casos reconstruyen a montos redondos.
//     Las que NO tienen MontoPagado (reservas futuras aún no cobradas) se
//     listan aparte: hay que traer el monto desde Airbnb a mano.
//
//  USO: correr primero el `preview...` de cada bloque, revisar el log y
//  después el `corregir...`.
// ═══════════════════════════════════════════════════════════

// Índices 0-based en la hoja Reservas (mismo mapeo que getReservations).
var _R = { ID:0, ENTRADA:4, MONTO:7, ORIGEN:9, COD:10, NETO:12, COD_TRANSF:18,
           FECHAPAGO:16, MONTOPAGADO:17, MONTOVOUCHER:19, ESTADO:20, COMENT:22 };

function _cleanMoney_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  var s = String(v).replace(/[^0-9.\-]/g, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

// Reparte `target` entre `weights` proporcionalmente; el último absorbe el redondeo.
function _repartir_(target, weights) {
  var tot = weights.reduce(function(a, b) { return a + b; }, 0);
  var out = weights.map(function(w) {
    return tot > 0 ? Math.round(target * w / tot * 100) / 100 : Math.round(target / weights.length * 100) / 100;
  });
  var s = Math.round(out.reduce(function(a, b) { return a + b; }, 0) * 100) / 100;
  out[out.length - 1] = Math.round((out[out.length - 1] + (target - s)) * 100) / 100;
  return out;
}

// Detecta grupos multi-cabaña (IDs `MC-<ts>-<n>`) cuyo voucher está duplicado.
// Devuelve [{ gid, filas:[{row, id, monto, voucher}], pagoReal, nuevos:[] }]
function _detectarMultiCabanaDuplicada_() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();
  var grupos = {};
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][_R.ID] || '');
    if (id.indexOf('MC-') !== 0) continue;
    var gid = id.substring(0, id.lastIndexOf('-'));
    if (!grupos[gid]) grupos[gid] = [];
    grupos[gid].push({
      row: i + 1, id: id,
      monto: _cleanMoney_(data[i][_R.MONTO]),
      voucher: _cleanMoney_(data[i][_R.MONTOVOUCHER]),
      coment: String(data[i][_R.COMENT] || '')
    });
  }
  var out = [];
  Object.keys(grupos).forEach(function(gid) {
    var fs = grupos[gid];
    if (fs.length < 2) return;
    var conVoucher = fs.filter(function(f) { return f.voucher > 0; });
    if (conVoucher.length < 2) return;                        // no hay duplicación
    var distintos = {}; conVoucher.forEach(function(f) { distintos[f.voucher.toFixed(2)] = 1; });
    if (Object.keys(distintos).length !== 1) return;          // ya están repartidos
    // Pago real: el comentario "[... pago único $X]" manda; si no, el valor repetido.
    var m = /pago único \$([0-9.]+)/.exec(fs[0].coment || '');
    var pagoReal = m ? parseFloat(m[1]) : conVoucher[0].voucher;
    var nuevos = _repartir_(pagoReal, fs.map(function(f) { return f.monto; }));
    out.push({ gid: gid, filas: fs, pagoReal: pagoReal, nuevos: nuevos });
  });
  return out;
}

function previewFixVoucherMultiCabana() {
  var grupos = _detectarMultiCabanaDuplicada_();
  if (!grupos.length) { Logger.log('✓ No hay multi-cabañas con voucher duplicado.'); return; }
  Logger.log('=== MULTI-CABAÑA con voucher duplicado: ' + grupos.length + ' grupo(s) ===');
  grupos.forEach(function(g) {
    var suma = g.filas.reduce(function(s, f) { return s + f.voucher; }, 0);
    Logger.log('Grupo ' + g.gid + ' — pago real $' + g.pagoReal.toFixed(2) +
               ' | hoy suma $' + suma.toFixed(2) + ' (inflado $' + (suma - g.pagoReal).toFixed(2) + ')');
    g.filas.forEach(function(f, i) {
      Logger.log('   fila ' + f.row + ' ' + f.id + '  Monto=$' + f.monto.toFixed(2) +
                 '  MontoVoucher: $' + f.voucher.toFixed(2) + ' → $' + g.nuevos[i].toFixed(2));
    });
  });
  Logger.log('Revisá y luego corré corregirVoucherMultiCabana()');
}

function corregirVoucherMultiCabana() {
  var sheet  = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var grupos = _detectarMultiCabanaDuplicada_();
  if (!grupos.length) { Logger.log('✓ Nada que corregir.'); return; }
  var n = 0;
  grupos.forEach(function(g) {
    g.filas.forEach(function(f, i) {
      sheet.getRange(f.row, _R.MONTOVOUCHER + 1).setValue('$' + g.nuevos[i].toFixed(2));
      Logger.log('✏️  fila ' + f.row + ' ' + f.id + ' → $' + g.nuevos[i].toFixed(2));
      n++;
    });
  });
  SpreadsheetApp.flush();
  Logger.log('✓ ' + n + ' fila(s) corregida(s) en ' + grupos.length + ' grupo(s).');
}

// ── B) Airbnb sin monto ───────────────────────────────────────
function _tasaAirbnb_(checkinISO) {
  // < 24-dic-2025 → plan host-only 3% · ≥ → split fee 15.5%
  return (checkinISO && checkinISO < '2025-12-24') ? 0.03 : 0.155;
}

function _detectarAirbnbSinMonto_() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();
  var conDato = [], sinDato = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][_R.ORIGEN] || '') !== 'Airbnb') continue;
    if (String(data[i][_R.ESTADO] || '') === 'CANCELADA') continue;
    if (_cleanMoney_(data[i][_R.MONTO]) > 0 || _cleanMoney_(data[i][_R.NETO]) > 0) continue;
    var pagado = _cleanMoney_(data[i][_R.MONTOPAGADO]);
    var ci = data[i][_R.ENTRADA];
    var ciISO = (ci instanceof Date) ? Utilities.formatDate(ci, 'America/Panama', 'yyyy-MM-dd')
                                     : String(ci || '').slice(0, 10);
    var item = { row: i + 1, cod: String(data[i][_R.COD] || ''), nombre: String(data[i][1] || ''),
                 checkin: ciISO, pagado: pagado };
    if (pagado > 0) {
      var tasa = _tasaAirbnb_(ciISO);
      item.tasa  = tasa;
      item.bruto = Math.round(pagado / (1 - tasa) * 100) / 100;
      conDato.push(item);
    } else {
      sinDato.push(item);
    }
  }
  return { conDato: conDato, sinDato: sinDato };
}

function previewFixAirbnbMontos() {
  var d = _detectarAirbnbSinMonto_();
  Logger.log('=== AIRBNB con Monto/Neto en 0 ===');
  Logger.log('A) Reconstruibles desde MontoPagado: ' + d.conDato.length);
  d.conDato.forEach(function(x) {
    Logger.log('   fila ' + x.row + ' ' + x.cod + ' ' + x.nombre +
               ' | pagado $' + x.pagado.toFixed(2) + ' · tasa ' + (x.tasa * 100).toFixed(1) + '%' +
               ' → Monto/Neto = $' + x.bruto.toFixed(2));
  });
  Logger.log('B) SIN dato de monto (traer de Airbnb a mano): ' + d.sinDato.length);
  d.sinDato.forEach(function(x) {
    Logger.log('   fila ' + x.row + ' ' + x.cod + ' ' + x.nombre + ' | check-in ' + x.checkin);
  });
  Logger.log('Revisá y luego corré corregirAirbnbMontos() (solo toca el grupo A).');
}

function corregirAirbnbMontos() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var d = _detectarAirbnbSinMonto_();
  if (!d.conDato.length) { Logger.log('✓ Nada que reconstruir.'); return; }
  d.conDato.forEach(function(x) {
    sheet.getRange(x.row, _R.MONTO + 1).setValue(x.bruto);
    sheet.getRange(x.row, _R.NETO  + 1).setValue(x.bruto);
    Logger.log('✏️  fila ' + x.row + ' ' + x.cod + ' → Monto/Neto = $' + x.bruto.toFixed(2));
  });
  SpreadsheetApp.flush();
  Logger.log('✓ ' + d.conDato.length + ' reserva(s) Airbnb actualizada(s).');
  if (d.sinDato.length) {
    Logger.log('⚠️  Faltan ' + d.sinDato.length + ' sin monto (reservas futuras). Cargalas a mano:');
    d.sinDato.forEach(function(x) { Logger.log('   fila ' + x.row + ' ' + x.cod + ' ' + x.nombre); });
  }
}

// ── C) Reporte (solo lectura): vouchers > monto que NO son multi-cabaña ──
// Casos ambiguos que requieren criterio humano: pueden ser sobrepago, propina,
// o un voucher que cubre 2 reservas cargadas por separado. NO se tocan.
function reporteVouchersInflados() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();
  var n = 0, exceso = 0;
  Logger.log('=== Vouchers > Monto (revisar a mano) ===');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][_R.ORIGEN] || '') === 'Airbnb') continue;
    if (String(data[i][_R.ESTADO] || '') === 'CANCELADA') continue;
    var monto = _cleanMoney_(data[i][_R.MONTO]);
    var vou   = _cleanMoney_(data[i][_R.MONTOVOUCHER]);
    if (monto <= 0 || vou <= monto + 0.005) continue;
    var esMC = String(data[i][_R.ID] || '').indexOf('MC-') === 0;
    n++; exceso += vou - monto;
    Logger.log('   fila ' + (i + 1) + ' ' + String(data[i][1] || '').slice(0, 26) +
               ' | Monto $' + monto.toFixed(2) + ' · Voucher $' + vou.toFixed(2) +
               ' · exceso $' + (vou - monto).toFixed(2) + (esMC ? '  [MULTI-CABAÑA]' : ''));
  }
  Logger.log('Total: ' + n + ' reserva(s), exceso $' + exceso.toFixed(2));
}

// ── D) Igualar Monto al Voucher (vouchers inflados NO multi-cabaña) ──
//  Para reservas donde el huésped pagó MÁS de lo que dice el Monto. Casi
//  siempre significa que la reserva valía lo del voucher pero se cargó
//  incompleta (ej. un pago que cubría 2 noches/cabañas y se registró una sola,
//  o un early check-in donde el pago incluía también la noche).
//  Deja la reserva cuadrada: Monto = Neto = Voucher y estado PAGA.
//
//  ⚠️ OJO: si algún día un huésped deja PROPINA o hace un sobrepago genuino,
//  este script lo convertiría en ingreso. Correr SIEMPRE el preview primero y
//  verificar caso por caso con reporteVouchersInflados().

function _detectarMontoMenorQueVoucher_() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][_R.ORIGEN] || '') === 'Airbnb') continue;
    if (String(data[i][_R.ESTADO] || '') === 'CANCELADA') continue;
    if (String(data[i][_R.ID] || '').indexOf('MC-') === 0) continue;  // esas van por el bloque A
    var monto = _cleanMoney_(data[i][_R.MONTO]);
    var vou   = _cleanMoney_(data[i][_R.MONTOVOUCHER]);
    if (monto <= 0 || vou <= monto + 0.005) continue;
    out.push({
      row: i + 1,
      nombre: String(data[i][1] || ''),
      monto: monto,
      neto: _cleanMoney_(data[i][_R.NETO]),
      voucher: vou,
      estado: String(data[i][_R.ESTADO] || '')
    });
  }
  return out;
}

function previewIgualarMontoAlVoucher() {
  var items = _detectarMontoMenorQueVoucher_();
  if (!items.length) { Logger.log('✓ No hay reservas con Monto < Voucher.'); return; }
  var extra = 0;
  Logger.log('=== Igualar Monto al Voucher: ' + items.length + ' reserva(s) ===');
  items.forEach(function(x) {
    extra += x.voucher - x.monto;
    Logger.log('   fila ' + x.row + ' ' + x.nombre.slice(0, 26) +
               ' | Monto $' + x.monto.toFixed(2) + ' → $' + x.voucher.toFixed(2) +
               ' · Neto $' + x.neto.toFixed(2) + ' → $' + x.voucher.toFixed(2) +
               ' · Estado ' + (x.estado || '(vacío)') + ' → PAGA');
  });
  Logger.log('Ingreso adicional a reconocer: +$' + extra.toFixed(2));
  Logger.log('Si todos son cargas incompletas (no propinas), corré igualarMontoAlVoucher()');
}

function igualarMontoAlVoucher() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var items = _detectarMontoMenorQueVoucher_();
  if (!items.length) { Logger.log('✓ Nada que corregir.'); return; }
  var extra = 0;
  items.forEach(function(x) {
    sheet.getRange(x.row, _R.MONTO  + 1).setValue(x.voucher);
    sheet.getRange(x.row, _R.NETO   + 1).setValue(x.voucher);
    sheet.getRange(x.row, _R.ESTADO + 1).setValue('PAGA');
    extra += x.voucher - x.monto;
    Logger.log('✏️  fila ' + x.row + ' ' + x.nombre.slice(0, 26) + ' → $' + x.voucher.toFixed(2) + ' PAGA');
  });
  SpreadsheetApp.flush();
  Logger.log('✓ ' + items.length + ' reserva(s) igualada(s). Ingreso adicional: +$' + extra.toFixed(2));
}

// ═══════════════════════════════════════════════════════════
//  E) CORRECCIONES PENDIENTES DE VOUCHER (auditoría 26-jul-2026)
//
//  Casos donde el MontoVoucher quedó inflado. En los 4 el monto realmente
//  recibido por esa reserva es su propio `Monto` (y ya coincide con el `Abono`
//  que se corrigió a mano) — lo que falta es bajar el voucher, porque
//  `montoRecibido()` le da PRIORIDAD al voucher sobre el abono:
//      mv > 0 ? mv : deposit      → mientras el voucher esté inflado, el abono
//                                   corregido a mano NO tiene efecto.
//
//   1-2) KEITLYN (multi-cabaña MC-1783376570702): pago único $150 en 2 cabañas.
//        El voucher se copió completo en ambas filas → $75 c/u.
//   3)   Yaricel Peralta: 2 reservas mismas fechas (Puente + Portal, $75 c/u),
//        un solo voucher de $150 cargado en la fila de Puente → $75.
//   4)   Gisela Nuñez: mismo voucher subido dos veces (codTransferencia trae
//        `LVMXW-69443745` repetido) y el sistema acumuló → $300 en vez de $200.
//        Se baja a $200 y se deduplican los códigos.
//
//  Es idempotente: si una fila ya está corregida, la saltea.
//  USO: previewCorreccionesVoucherPendientes() → corregirVouchersPendientes()
// ═══════════════════════════════════════════════════════════

function _idsVoucherPendientes_() {
  return ['MC-1783376570702-1', 'MC-1783376570702-2', '1774551742898', '1776205252012'];
}

// Deduplica la lista `A|B|A` → `A|B` conservando el orden.
function _dedupCods_(s) {
  var parts = String(s || '').split('|');
  var seen = {}, out = [];
  for (var i = 0; i < parts.length; i++) {
    var p = parts[i].trim();
    if (!p) continue;
    var k = p.replace(/^#/, '').toUpperCase();
    if (seen[k]) continue;
    seen[k] = 1; out.push(p);
  }
  return out.join('|');
}

function _analizarVoucherPendientes_() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();
  var want  = {}; _idsVoucherPendientes_().forEach(function(id) { want[id] = true; });
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var id = String(data[i][_R.ID] || '');
    if (!want[id]) continue;
    var monto   = _cleanMoney_(data[i][_R.MONTO]);
    var voucher = _cleanMoney_(data[i][_R.MONTOVOUCHER]);
    var cods    = String(data[i][_R.COD_TRANSF] || '');
    var codsNew = _dedupCods_(cods);
    out.push({
      row: i + 1, id: id, nombre: String(data[i][1] || ''),
      cabana: String(data[i][2] || ''),
      monto: monto, voucher: voucher, abono: _cleanMoney_(data[i][8]),
      nuevoVoucher: monto,                       // el recibido real de ESTA reserva
      cods: cods, codsNew: codsNew,
      yaOk: voucher <= monto + 0.005,
      cambiaCods: codsNew !== cods
    });
  }
  return out;
}

function previewCorreccionesVoucherPendientes() {
  var items = _analizarVoucherPendientes_();
  if (!items.length) { Logger.log('⚠️ No se encontró ninguna de las filas objetivo (¿IDs cambiados?).'); return; }
  var falta = 0;
  Logger.log('=== Correcciones de voucher pendientes ===');
  items.forEach(function(x) {
    if (x.yaOk && !x.cambiaCods) {
      Logger.log('   ✓ fila ' + x.row + ' ' + x.nombre.slice(0, 22) + ' — ya corregida (voucher $' + x.voucher.toFixed(2) + ')');
      return;
    }
    falta++;
    Logger.log('   fila ' + x.row + ' ' + x.nombre.slice(0, 22) + ' · ' + x.cabana.slice(0, 18));
    if (!x.yaOk) {
      Logger.log('        MontoVoucher $' + x.voucher.toFixed(2) + ' → $' + x.nuevoVoucher.toFixed(2) +
                 '   (Monto $' + x.monto.toFixed(2) + ' · Abono $' + x.abono.toFixed(2) + ')');
    }
    if (x.cambiaCods) Logger.log('        códigos: ' + x.cods + ' → ' + x.codsNew);
  });
  Logger.log(falta ? ('Pendientes: ' + falta + '. Corré corregirVouchersPendientes()') : '✓ Todo corregido.');
}

function corregirVouchersPendientes() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var items = _analizarVoucherPendientes_();
  if (!items.length) { Logger.log('⚠️ No se encontró ninguna de las filas objetivo.'); return; }
  var n = 0, ajuste = 0;
  items.forEach(function(x) {
    var toco = false;
    if (!x.yaOk) {
      sheet.getRange(x.row, _R.MONTOVOUCHER + 1).setValue('$' + x.nuevoVoucher.toFixed(2));
      ajuste += x.voucher - x.nuevoVoucher;
      toco = true;
    }
    if (x.cambiaCods) { sheet.getRange(x.row, _R.COD_TRANSF + 1).setValue(x.codsNew); toco = true; }
    if (toco) { n++; Logger.log('✏️  fila ' + x.row + ' ' + x.nombre.slice(0, 22) + ' → voucher $' + x.nuevoVoucher.toFixed(2)); }
  });
  SpreadsheetApp.flush();
  Logger.log('✓ ' + n + ' fila(s) corregida(s). Ingreso sobrecontado eliminado: $' + ajuste.toFixed(2));
}

// ═══════════════════════════════════════════════════════════
//  F) CARGAR RESERVAS AIRBNB FALTANTES (auditoría 26-jul-2026)
//
//  El cruce del export oficial de Airbnb contra la hoja `Reservas` mostró
//  reservas que SÍ fueron cobradas (aparecen en la hoja `Pagos`) pero que
//  nunca se registraron. Todas son de "Paseo por Las Nubes", lo que sugiere
//  que el parser de emails falló con ese formato.
//
//  Los montos salen del export de Airbnb:
//    Monto/Neto = "Ingresos brutos"  (lo que pagó el huésped)
//    MontoPagado = "Monto"           (lo que Airbnb depositó, ya neto de su comisión)
//
//  Es idempotente: saltea los códigos que ya existan.
//  USO: previewReservasAirbnbFaltantes() → insertarReservasAirbnbFaltantes()
// ═══════════════════════════════════════════════════════════

function _reservasAirbnbFaltantes_() {
  // cod, huésped, cabaña, código de cabaña, entrada, salida, noches,
  // bruto (lo que pagó el huésped), neto (lo que depositó Airbnb),
  // fechaReserva, fechaPago (payout donde se cobró)
  return [
    { cod:'HMQRAW8CA5', nombre:'Anna Van Mondfrans', cabana:'Paseo por Las Nubes', code:'verde',
      entrada:'2026-01-07', salida:'2026-01-10', noches:3, bruto:261.00, neto:253.17,
      fechaReserva:'2025-11-15', fechaPago:'2026-01-15' },
    { cod:'HMBACCYQM9', nombre:'Yuliany Guerrero',   cabana:'Paseo por Las Nubes', code:'verde',
      entrada:'2026-03-16', salida:'2026-03-18', noches:2, bruto:178.20, neto:150.58,
      fechaReserva:'2026-02-22', fechaPago:'2026-03-26' },
    { cod:'HMD4ESKMPX', nombre:'Kj Thomas',          cabana:'Paseo por Las Nubes', code:'verde',
      entrada:'2026-03-18', salida:'2026-03-21', noches:3, bruto:276.30, neto:233.47,
      fechaReserva:'2026-02-22', fechaPago:'2026-03-26' },
    { cod:'HMBBQXQ3QD', nombre:'Yarisel Rangel',     cabana:'Paseo por Las Nubes', code:'verde',
      entrada:'2026-06-05', salida:'2026-06-06', noches:1, bruto:169.00, neto:142.80,
      fechaReserva:'2026-05-27', fechaPago:'2026-06-17' },
    { cod:'HMEQYEZS85', nombre:'Mairanis Lopez',     cabana:'Paseo por Las Nubes', code:'verde',
      entrada:'2026-06-19', salida:'2026-06-21', noches:2, bruto:196.20, neto:165.79,
      fechaReserva:'2026-05-18', fechaPago:'2026-06-22' }
  ];
}

function _codsExistentes_() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();
  const set   = {};
  for (var i = 1; i < data.length; i++) {
    var c = String(data[i][_R.COD] || '').trim();
    if (c) set[c] = true;
  }
  return set;
}

function previewReservasAirbnbFaltantes() {
  const ex = _codsExistentes_();
  const items = _reservasAirbnbFaltantes_();
  var nuevos = 0, bruto = 0;
  Logger.log('=== Reservas Airbnb faltantes ===');
  items.forEach(function(x) {
    if (ex[x.cod]) { Logger.log('   ✓ ' + x.cod + ' ya existe — se saltea'); return; }
    nuevos++; bruto += x.bruto;
    Logger.log('   + ' + x.cod + ' ' + x.nombre + ' · ' + x.cabana +
               ' · ' + x.entrada + '→' + x.salida + ' (' + x.noches + 'n)' +
               ' · bruto $' + x.bruto.toFixed(2) + ' · depositado $' + x.neto.toFixed(2) +
               ' · payout ' + x.fechaPago);
  });
  Logger.log(nuevos ? ('A insertar: ' + nuevos + ' reserva(s), bruto $' + bruto.toFixed(2) +
                       '. Corré insertarReservasAirbnbFaltantes()')
                    : '✓ No hay nada que insertar.');
}

function insertarReservasAirbnbFaltantes() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const ex    = _codsExistentes_();
  const nCols = sheet.getLastColumn();
  var insertados = 0;
  _reservasAirbnbFaltantes_().forEach(function(x) {
    if (ex[x.cod]) { Logger.log('~ ' + x.cod + ' ya existe, se saltea'); return; }
    var fila = new Array(nCols);
    for (var k = 0; k < nCols; k++) fila[k] = '';
    fila[_R.ID]           = Utilities.getUuid().replace(/-/g, '').substring(0, 16);
    fila[1]               = x.nombre;                 // Nombre
    fila[2]               = x.cabana;                 // Cabaña
    fila[3]               = x.code;                   // CabañaCodigo
    fila[_R.ENTRADA]      = x.entrada;                // Entrada
    fila[5]               = x.salida;                 // Salida
    fila[6]               = 2;                        // Personas (Airbnb no lo trae acá)
    fila[_R.MONTO]        = x.bruto;                  // Monto = lo que pagó el huésped
    fila[8]               = 0;                        // Abono
    fila[_R.ORIGEN]       = 'Airbnb';
    fila[_R.COD]          = x.cod;                    // CodConfirmacion
    fila[11]              = 0;                        // Service Fee
    fila[_R.NETO]         = x.bruto;                  // Neto (convención: Monto == Neto en Airbnb)
    fila[15]              = x.fechaReserva;           // FechaReserva
    fila[_R.FECHAPAGO]    = x.fechaPago;
    fila[_R.MONTOPAGADO]  = x.neto;                   // lo realmente depositado por Airbnb
    fila[_R.ESTADO]       = 'PAGA';
    fila[24]              = 'noche';                  // Tipo
    fila[_R.COMENT]       = '[Cargada por auditoría 26-jul-2026 desde el export de Airbnb]';
    sheet.appendRow(fila);
    insertados++;
    Logger.log('✏️  insertada ' + x.cod + ' ' + x.nombre + ' $' + x.bruto.toFixed(2));
  });
  SpreadsheetApp.flush();
  Logger.log('✓ ' + insertados + ' reserva(s) insertada(s).');
}

// ═══════════════════════════════════════════════════════════
//  G) DIAGNÓSTICO DEL PARSER DE RESERVAS AIRBNB
//
//  Vuelca EXACTAMENTE lo que ve `syncAirbnbReservations`: el `getPlainBody()`
//  numerado línea por línea, y el resultado de `parseAirbnbEmail()` sobre ese
//  mismo texto. Sirve para ver por qué un email no generó reserva sin depender
//  de copiar/pegar (que rompe los saltos de línea, de los que el parser depende).
//
//  USO en el editor de Apps Script:
//     diagnosticarEmailReserva('Yuliany')        // busca por nombre del huésped
//     diagnosticarEmailReserva('HMBACCYQM9')     // o por código
//     listarEmailsReservaAirbnb()                // lista los últimos y cuáles ya están en la hoja
//
//  El log de Apps Script trunca mensajes largos: si el body es muy grande,
//  usá diagnosticarEmailReserva(x, true) para volcar SOLO las líneas clave.
// ═══════════════════════════════════════════════════════════

function _buscarMsgsReserva_(query) {
  var threads = GmailApp.search('from:automated@airbnb.com subject:"Reserva confirmada:" newer_than:400d');
  var out = [];
  threads.forEach(function(t) {
    t.getMessages().forEach(function(m) {
      var hay = m.getSubject() + '\n' + m.getPlainBody();
      if (!query || hay.toUpperCase().indexOf(String(query).toUpperCase()) >= 0) out.push(m);
    });
  });
  return out;
}

function listarEmailsReservaAirbnb() {
  var msgs = _buscarMsgsReserva_(null);
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var cods = {};
  for (var i = 1; i < data.length; i++) {
    var c = String(data[i][_R.COD] || '').trim(); if (c) cods[c] = true;
  }
  Logger.log('=== Emails "Reserva confirmada" (últimos 400d): ' + msgs.length + ' ===');
  msgs.forEach(function(m) {
    var body = m.getPlainBody();
    var cm = body.match(/\b(HM[A-Z0-9]{8})\b/);
    var cod = cm ? cm[1] : '(sin código)';
    var enHoja = cods[cod] ? '✓ en hoja' : '✗ FALTA EN LA HOJA';
    var parsed = null;
    try { parsed = parseAirbnbEmail(body, m.getId(),
            Utilities.formatDate(m.getDate(), 'America/Panama', 'yyyy-MM-dd')); } catch (e) {}
    Logger.log('  ' + Utilities.formatDate(m.getDate(), 'America/Panama', 'yyyy-MM-dd') +
               ' ' + cod + ' ' + enHoja +
               ' | parse: ' + (parsed ? ('OK ' + parsed.cabin + ' ' + parsed.checkin + '→' + parsed.checkout + ' $' + parsed.amount) : 'NULL ⚠') +
               ' | ' + m.getSubject().slice(0, 60));
  });
}

function diagnosticarEmailReserva(query, soloClaves) {
  var msgs = _buscarMsgsReserva_(query);
  if (!msgs.length) { Logger.log('⚠️ No se encontró ningún email que contenga: ' + query); return; }
  Logger.log('=== ' + msgs.length + ' email(s) para "' + query + '" ===');
  msgs.forEach(function(m, idx) {
    var body = m.getPlainBody();
    var fecha = Utilities.formatDate(m.getDate(), 'America/Panama', 'yyyy-MM-dd');
    Logger.log('\n───────── EMAIL ' + (idx + 1) + ' ─────────');
    Logger.log('Asunto: ' + m.getSubject());
    Logger.log('Fecha : ' + fecha + '  | msgId: ' + m.getId());
    Logger.log('Largo body: ' + body.length + ' chars');

    // Qué extrae cada regex del parser (para ver cuál falla)
    Logger.log('\n--- Lo que detecta cada regex ---');
    var nameMatch = body.match(/([A-ZÁÉÍÓÚÜÑ][a-záéíóúüñ]+(?:\s+[A-ZÁÉÍÓÚÜÑ][a-záéíóúüñ]+){1,3})\s+\[https:\/\/www\.airbnb/);
    Logger.log('  nombre (regex principal): ' + (nameMatch ? nameMatch[1] : 'NO MATCH → usa fallback'));
    var cab = 'NINGUNA (default verde)';
    for (var cn in CABINS) { if (body.toUpperCase().indexOf(cn.toUpperCase()) >= 0) { cab = cn + ' → ' + CABINS[cn]; break; } }
    Logger.log('  cabaña: ' + cab);
    var fm = body.match(/Llegada\s+Salida\s+(\w+,?\s+\d{1,2}\s+\w+)\s+(\w+,?\s+\d{1,2}\s+\w+)/i);
    Logger.log('  fechas "Llegada Salida": ' + (fm ? (fm[1] + ' | ' + fm[2]) : 'NO MATCH'));
    var nm = body.match(/por\s+(\d+)\s+noche/i);
    Logger.log('  noches: ' + (nm ? nm[1] : 'NO MATCH'));
    var tm = body.match(/Total\s*\(USD\)[^\$]*\$\s*([\d,\.]+)/i);
    Logger.log('  total: ' + (tm ? tm[1] : 'NO MATCH ⚠'));
    var cm = body.match(/\b(HM[A-Z0-9]{8})\b/);
    Logger.log('  código: ' + (cm ? cm[1] : 'NO MATCH'));

    var parsed = null, err = '';
    try { parsed = parseAirbnbEmail(body, m.getId(), fecha); } catch (e) { err = e.message; }
    Logger.log('\n--- Resultado de parseAirbnbEmail ---');
    Logger.log(parsed ? JSON.stringify(parsed) : ('NULL ⚠' + (err ? ' · error: ' + err : '')));

    // Body línea por línea (lo que realmente ve el parser)
    var lines = body.split(/\r?\n/);
    Logger.log('\n--- BODY (' + lines.length + ' líneas, numeradas) ---');
    for (var i = 0; i < lines.length; i++) {
      var l = lines[i];
      if (soloClaves) {
        var rel = /HM[A-Z0-9]{8}|Llegada|Salida|noche|Total|USD|Nubes|adulto|\d{1,2}\s+(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)/i.test(l);
        if (!rel) continue;
      }
      Logger.log(String(i + 1).padStart(3, ' ') + '| ' + l);
    }
  });
}

// ═══════════════════════════════════════════════════════════
//  H) RECONCILIACIÓN DE RESERVAS AIRBNB (red de seguridad)
//
//  `syncAirbnbReservations` solo mira una ventana de emails recientes. Si una
//  reserva no se registra dentro de esa ventana (trigger caído, error puntual),
//  el email queda fuera del rango y NUNCA se reintenta: se pierde para siempre.
//  Esto compara TODOS los emails de reserva contra la hoja y reporta/carga los
//  que falten, sin depender de la ventana del sync.
//
//  Ignora los códigos en Blacklist (cancelados) y los que ya están en la hoja.
//  Por default solo mira estadías del AÑO EN CURSO (filtra por check-in, no por
//  fecha del email, para que entre una reserva hecha en 2025 con estadía 2026).
//  USO: reconciliarReservasAirbnb()                    → reporta (año en curso)
//       reconciliarReservasAirbnb(true)                → además las inserta
//       reconciliarReservasAirbnb(false, '2025-01-01') → otro corte de fecha
// ═══════════════════════════════════════════════════════════

function _codsBlacklist_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName('Blacklist');
  var out = {};
  if (!sh) return out;
  var d = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    var c = String(d[i][0] || '').trim();
    if (c) out[c.replace(/\.0$/, '')] = true;
  }
  return out;
}

function reconciliarReservasAirbnb(insertar, desdeISO) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var enHoja = _codsExistentes_();
  var black  = _codsBlacklist_();
  var threads = GmailApp.search('from:automated@airbnb.com subject:"Reserva confirmada:" newer_than:400d');

  // Filtro por fecha de ESTADÍA (check-in), no por fecha del email: así entra
  // una reserva hecha en 2025 cuya estadía cae en 2026. Default = 1-ene del año
  // en curso, para no arrastrar el histórico previo al arranque de la hoja.
  var desde = desdeISO || (new Date().getFullYear() + '-01-01');

  var faltan = [], yaEstan = 0, blacklisted = 0, noParse = 0, viejas = 0;
  threads.forEach(function(t) {
    t.getMessages().forEach(function(m) {
      var body = m.getPlainBody();
      var cm = body.match(/\b(HM[A-Z0-9]{8})\b/);
      var cod = cm ? cm[1] : null;
      if (!cod) return;
      if (enHoja[cod]) { yaEstan++; return; }
      if (black[cod])  { blacklisted++; return; }
      var fecha = Utilities.formatDate(m.getDate(), 'America/Panama', 'yyyy-MM-dd');
      var r = null;
      try { r = parseAirbnbEmail(body, m.getId(), fecha); } catch (e) {}
      if (!r) { noParse++; Logger.log('⚠ no parsea: ' + cod + ' (' + fecha + ')'); return; }
      if (String(r.checkin || '') < desde) { viejas++; return; }   // fuera del período pedido
      faltan.push({ cod: cod, emailFecha: fecha, r: r });
    });
  });

  Logger.log('=== Reconciliación de reservas Airbnb (estadías desde ' + desde + ') ===');
  Logger.log('  ya en la hoja: ' + yaEstan + ' | en blacklist: ' + blacklisted +
             ' | no parsean: ' + noParse + ' | anteriores a ' + desde + ': ' + viejas +
             ' | FALTAN: ' + faltan.length);
  faltan.sort(function(a, b) { return a.emailFecha < b.emailFecha ? -1 : 1; });
  faltan.forEach(function(x) {
    Logger.log('   ' + x.emailFecha + ' ' + x.cod + ' ' + String(x.r.name).slice(0, 22) +
               ' · ' + x.r.cabin + ' · ' + x.r.checkin + '→' + x.r.checkout + ' · $' + x.r.amount);
  });

  if (!insertar) {
    Logger.log(faltan.length ? 'Para insertarlas: reconciliarReservasAirbnb(true)' : '✓ Nada que reconciliar.');
    return;
  }
  var n = 0;
  faltan.forEach(function(x) {
    try { appendReservation(sheet, x.r); n++; Logger.log('✏️  insertada ' + x.cod); }
    catch (e) { Logger.log('✗ error insertando ' + x.cod + ': ' + e.message); }
  });
  SpreadsheetApp.flush();
  Logger.log('✓ ' + n + ' reserva(s) insertada(s).');
}

// ── I) Reserva guardada con ID sintético en vez del código de Airbnb ──
//  Lourdes Morales quedó con ID `airbnb_19f3fae432b7c17d` en vez de
//  `HMKFCXHCQJ`, así que no cruzaba con la hoja Pagos ni con el email.
//  Genérico: recibe pares {idActual, codReal} y escribe CodConfirmacion.
function _codsSinteticosACorregir_() {
  return [ { idActual: 'airbnb_19f3fae432b7c17d', codReal: 'HMKFCXHCQJ', quien: 'Lourdes Morales' } ];
}

function previewCorregirCodsSinteticos() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();
  _codsSinteticosACorregir_().forEach(function(x) {
    var hallado = false;
    for (var i = 1; i < data.length; i++) {
      var id  = String(data[i][_R.ID]  || '').trim();
      var cod = String(data[i][_R.COD] || '').trim();
      if (id !== x.idActual && cod !== x.idActual) continue;
      hallado = true;
      Logger.log('  fila ' + (i + 1) + ' ' + x.quien + ' | CodConfirmacion: "' + cod + '" → "' + x.codReal + '"' +
                 (cod === x.codReal ? '  (ya corregida)' : ''));
    }
    if (!hallado) Logger.log('  ⚠ no se encontró la fila de ' + x.quien + ' (' + x.idActual + ')');
  });
}

function corregirCodsSinteticos() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var data  = sheet.getDataRange().getValues();
  var n = 0;
  _codsSinteticosACorregir_().forEach(function(x) {
    for (var i = 1; i < data.length; i++) {
      var id  = String(data[i][_R.ID]  || '').trim();
      var cod = String(data[i][_R.COD] || '').trim();
      if ((id !== x.idActual && cod !== x.idActual) || cod === x.codReal) continue;
      sheet.getRange(i + 1, _R.COD + 1).setValue(x.codReal);
      n++;
      Logger.log('✏️  fila ' + (i + 1) + ' ' + x.quien + ' → ' + x.codReal);
    }
  });
  SpreadsheetApp.flush();
  Logger.log('✓ ' + n + ' código(s) corregido(s). Corré actualizarEstadoPagoAirbnb() para traer su monto.');
}

// ═══════════════════════════════════════════════════════════
//  Vouchers que muestran "sin archivo" en el modal
// ═══════════════════════════════════════════════════════════
//
// Una reserva puede tener CodTransferencia (el OCR leyó el voucher) y aun así
// no tener VoucherURL. Este diagnóstico separa las dos causas posibles:
//
//   A) El archivo NUNCA se subió a Drive. Pasa con las reservas que entran por
//      el Agente de WhatsApp: el bot descarga la imagen, la manda a Claude y
//      guarda el código, pero descarta los bytes — nunca llama a
//      saveVoucherToDrive. Se reconocen porque el código conserva el '#' que
//      el dashboard sí quita (handleVoucherUpload hace .replace(/^#/, '')).
//
//   B) El archivo SÍ está en Drive pero la URL no quedó en la hoja (la subida
//      falló a mitad, o es anterior a que existiera la columna). Esas son
//      recuperables: `repoblarVoucherURLs()` las matchea por código en la
//      descripción del archivo y rellena la columna.
//
// Solo reporta, no escribe nada.
function diagnosticarVouchersSinArchivo() {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();

  // ── Índice de Drive ────────────────────────────────────────────────────
  // OJO con la llave: la descripción que escribe saveVoucherToDrive NO lleva el
  // código de transferencia. Lleva 'Cód. Pago: ' + (confirmCode || id), y en las
  // reservas del dashboard confirmCode ES el id de la reserva. Buscar por código
  // de transferencia no matchea casi nunca y da un falso "no está en Drive".
  // Se indexa por id de reserva, por huésped+entrada y por huésped solo.
  const norm = s => (s || '').toString().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9]+/g, ' ').trim();
  const byPago = {}, byNameDate = {}, byName = {};
  const archivos = [];
  const folders = DriveApp.getFoldersByName(VOUCHER_FOLDER_NAME);
  if (folders.hasNext()) {
    const it = folders.next().getFiles();
    while (it.hasNext()) {
      const f    = it.next();
      const desc = f.getDescription() || '';
      const get  = re => { const m = desc.match(re); return m ? m[1].trim() : ''; };
      const info = {
        url:     f.getUrl(),
        nombre:  f.getName(),
        huesped: get(/Hu[ée]sped:\s*([^\n]*)/i),
        entrada: get(/Entrada:\s*([^\n]*)/i),
        codPago: get(/C[óo]d\.\s*Pago:\s*([^\n]*)/i)
      };
      archivos.push(info);
      if (info.codPago) (byPago[info.codPago] = byPago[info.codPago] || []).push(info);
      const nn = norm(info.huesped);
      if (nn) {
        (byName[nn] = byName[nn] || []).push(info);
        if (info.entrada) (byNameDate[nn + '|' + info.entrada] = byNameDate[nn + '|' + info.entrada] || []).push(info);
      }
    }
  }
  Logger.log('📁 ' + archivos.length + ' archivos en "' + VOUCHER_FOLDER_NAME + '"');

  const split = s => (s || '').toString().split('|').map(x => x.trim()).filter(Boolean);
  const iso = v => v instanceof Date ? Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd') : (v || '').toString().slice(0, 10);

  // URLs ya referenciadas en la hoja → para detectar archivos huérfanos
  const usadas = {};
  for (let i = 1; i < data.length; i++) split(data[i][25]).forEach(u => { usadas[u] = true; });

  const rec = { okCompleto: 0, sinCodigo: 0, recuperable: [], nuncaSubido: [], desalineado: [] };

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[_R.ID]) continue;
    const origen = (r[_R.ORIGEN] || '').toString().trim();
    if (['Directa', 'Referido'].indexOf(origen) < 0) continue;
    const cods = split(r[_R.COD_TRANSF]);
    const urls = split(r[25]);
    if (!cods.length) { rec.sinCodigo++; continue; }
    if (urls.length >= cods.length) { rec.okCompleto++; continue; }

    const id      = r[_R.ID].toString();
    const nombre  = r[1];
    const entrada = iso(r[_R.ENTRADA]);
    const nn      = norm(nombre);

    // Buscar el archivo por las llaves que SÍ existen en la descripción
    let hallados = byPago[id] || byNameDate[nn + '|' + entrada] || [];
    let via = byPago[id] ? 'id' : (hallados.length ? 'huesped+entrada' : '');
    if (!hallados.length && byName[nn] && byName[nn].length === 1) {
      hallados = byName[nn]; via = 'huesped (probable)';
    }
    // Solo cuenta como recuperable si el archivo no está ya enlazado en otra fila
    const nuevos = hallados.filter(f => !usadas[f.url]);

    const info = {
      fila: i + 1, id: id, nombre: nombre, entrada: entrada,
      cods: cods.join(', '), urls: urls.length,
      conHash: cods.some(c => c.indexOf('#') === 0)
    };
    if (nuevos.length) {
      info.via = via; info.archivos = nuevos.map(f => f.nombre).join(', ');
      info.urlsDrive = nuevos.map(f => f.url).join(' | ');
      rec.recuperable.push(info);
    } else if (urls.length > 0) rec.desalineado.push(info);
    else rec.nuncaSubido.push(info);
  }

  const huerfanos = archivos.filter(f => !usadas[f.url]);

  Logger.log('');
  Logger.log('═══ VOUCHERS SIN ARCHIVO ═══');
  Logger.log('✓ Con todas sus URLs: ' + rec.okCompleto + '  ·  sin código (nada que buscar): ' + rec.sinCodigo);
  Logger.log('📎 Archivos en Drive sin enlazar a ninguna fila: ' + huerfanos.length + ' de ' + archivos.length);
  Logger.log('');
  Logger.log('🟢 CAUSA B — recuperables (el archivo SÍ está en Drive): ' + rec.recuperable.length);
  rec.recuperable.forEach(x => Logger.log('   fila ' + x.fila + ' · ' + x.nombre + ' · ' + x.entrada
    + ' · match por ' + x.via + ' → ' + x.archivos));
  if (rec.recuperable.length) Logger.log('   → Corré vincularVouchersHuerfanos() para escribir esas URLs.');
  Logger.log('');
  Logger.log('🔴 CAUSA A — nunca se subió el archivo: ' + rec.nuncaSubido.length);
  rec.nuncaSubido.forEach(x => Logger.log('   fila ' + x.fila + ' · ' + x.nombre + ' · ' + x.entrada + ' · ' + x.cods
    + (x.conHash ? '  [# → bot de WhatsApp]' : '')));
  Logger.log('   de esas, ' + rec.nuncaSubido.filter(x => x.conHash).length + ' entraron por el Agente de WhatsApp.');
  Logger.log('');
  Logger.log('🟡 Desalineadas (tienen alguna URL pero menos que códigos): ' + rec.desalineado.length);
  rec.desalineado.forEach(x => Logger.log('   fila ' + x.fila + ' · ' + x.nombre + ' · ' + x.cods + ' · urls=' + x.urls));
  return rec;
}

// ═══════════════════════════════════════════════════════════
//  Pagos registrados SOLO en el campo Abono (sin rastro de voucher)
// ═══════════════════════════════════════════════════════════
//
// Caso Severino Villarreal (jun-2026): la fila tenía Monto $110 y **Abono $110**
// pero `codTransferencia`, `MontoVoucher`, `VoucherURL` y `VouchersMeta` vacíos.
// El pago existe —el abono es la evidencia— pero el voucher no quedó en ningún
// lado, así que la reserva no aparece en "Por cobrar" (Contabilidad deriva el
// cobro del abono) y a la vez salía "○ Pendiente" (la columna `EstadoPago` solo
// se recalculaba al subir un voucher). Dos superficies, dos respuestas opuestas.
//
// `diagnosticarVouchersSinArchivo()` NO los ve: descarta las filas sin código de
// transferencia antes de mirar Drive, y estas justamente no lo tienen. Por eso
// esta función va por el otro lado — busca en Drive por id de reserva y por
// huésped+entrada — y separa dos desenlaces distintos:
//
//   • está en Drive  → recuperable: el archivo se subió, la URL no se escribió.
//                      Se enlaza con `vincularPagosSinRastroESCRIBIR()`.
//   • no está        → el voucher nunca llegó a Drive. El pago igual es real
//                      (el abono lo respalda); lo que falta es el comprobante.
//
// ── Carga retroactiva ──────────────────────────────────────────────────────
// La primera corrida (jul-2026) devolvió 41 filas, y 34 eran ruido: las filas
// 115-148, todas creadas en una ventana de 17 horas del 4-5 de marzo con
// entradas de FEBRERO — estadías que ya habían pasado cuando se cargaron. Ahí
// nunca hubo voucher que subir: el pago ocurrió antes de que la reserva
// existiera en el sistema. Reportarlas junto a las que sí perdieron el archivo
// convierte el log en una pared de 41 líneas donde las 7 que importan se
// pierden. Se separan comparando el timestamp del `id` (los ids del dashboard
// son `Date.now()`) contra el check-in: creada DESPUÉS de la entrada = carga
// retroactiva, no es un archivo perdido.
//
// Solo lee y reporta. Para escribir: vincularPagosSinRastroESCRIBIR().
function diagnosticarPagosSinRastro() {
  return _pagosSinRastro_(true);
}

function vincularPagosSinRastro(dryRun) {
  if (dryRun === undefined) dryRun = true;   // default seguro: el editor corre sin argumentos
  return _pagosSinRastro_(dryRun);
}

function vincularPagosSinRastroESCRIBIR() {
  return _pagosSinRastro_(false);
}

// Fecha en que se CREÓ la fila, deducida del id. Los ids del dashboard son
// `Date.now()` (13 dígitos); los de Airbnb y los hex no dicen nada y devuelven ''.
function _idCreadoISO_(id) {
  const s = (id || '').toString().trim();
  if (!/^\d{13}$/.test(s)) return '';
  const n = parseInt(s, 10);
  // Sanidad: entre 2020 y 2100, por si algún id numérico no fuera un timestamp.
  if (n < 1577854800000 || n > 4102462800000) return '';
  return Utilities.formatDate(new Date(n), 'America/Panama', 'yyyy-MM-dd');
}

function _pagosSinRastro_(dryRun) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const split = s => (s || '').toString().split('|').map(x => x.trim()).filter(Boolean);
  const iso   = v => v instanceof Date ? Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd') : (v || '').toString().slice(0, 10);
  const norm  = s => (s || '').toString().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z0-9]+/g, ' ').trim();

  // ── Índice de Drive (mismas llaves que diagnosticarVouchersSinArchivo) ──
  const byPago = {}, byNameDate = {}, usadas = {};
  for (let i = 1; i < data.length; i++) split(data[i][25]).forEach(u => { usadas[u] = true; });
  let totalArchivos = 0;
  const folders = DriveApp.getFoldersByName(VOUCHER_FOLDER_NAME);
  if (folders.hasNext()) {
    const it = folders.next().getFiles();
    while (it.hasNext()) {
      const f    = it.next();
      const desc = f.getDescription() || '';
      const get  = re => { const m = desc.match(re); return m ? m[1].trim() : ''; };
      const info = { url: f.getUrl(), nombre: f.getName(),
                     huesped: get(/Hu[ée]sped:\s*([^\n]*)/i),
                     entrada: get(/Entrada:\s*([^\n]*)/i),
                     codPago: get(/C[óo]d\.\s*Pago:\s*([^\n]*)/i) };
      totalArchivos++;
      if (info.codPago) (byPago[info.codPago] = byPago[info.codPago] || []).push(info);
      const nn = norm(info.huesped);
      if (nn && info.entrada) (byNameDate[nn + '|' + info.entrada] = byNameDate[nn + '|' + info.entrada] || []).push(info);
    }
  }

  Logger.log('');
  Logger.log('═══ ' + (dryRun ? 'DRY-RUN · ' : '') + 'PAGOS SIN RASTRO DE VOUCHER ═══');
  Logger.log('📁 ' + totalArchivos + ' archivos en "' + VOUCHER_FOLDER_NAME + '"');
  Logger.log('');

  const enDrive = [], sinArchivo = [], retro = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[_R.ID]) continue;
    const origen = (r[_R.ORIGEN] || '').toString().trim();
    if (['Directa', 'Referido'].indexOf(origen) < 0) continue;
    if ((r[_R.ESTADO] || '').toString().toUpperCase() === 'CANCELADA') continue;

    // Rastro de voucher: cualquiera de las cuatro columnas alcanza.
    const tieneRastro = split(r[_R.COD_TRANSF]).length || split(r[25]).length
                     || _cleanMoney_(r[_R.MONTOVOUCHER]) > 0
                     || (r[31] || '').toString().trim();
    if (tieneRastro) continue;

    const abono = _cleanMoney_(r[8]);
    if (abono <= 0) continue;   // sin abono no hay pago registrado; no es este caso

    const id      = r[_R.ID].toString();
    const nombre  = r[1];
    const entrada = iso(r[_R.ENTRADA]);
    const total   = _cleanMoney_(r[_R.NETO]) || _cleanMoney_(r[_R.MONTO]);
    const hall    = (byPago[id] || byNameDate[norm(nombre) + '|' + entrada] || [])
                      .filter(f => !usadas[f.url]);
    const creada = _idCreadoISO_(id);
    const item = { fila: i + 1, id: id, nombre: nombre, entrada: entrada,
                   total: total, abono: abono, creada: creada,
                   estado: (r[_R.ESTADO] || '').toString() || '(vacío)',
                   via: byPago[id] ? 'id' : (hall.length ? 'huesped+entrada' : ''),
                   archivos: hall };
    if (hall.length) enDrive.push(item);
    // Fila creada DESPUÉS del check-in: la estadía ya había pasado, el pago
    // ocurrió antes de que la reserva existiera. No hay archivo perdido.
    else if (creada && creada > entrada) retro.push(item);
    else sinArchivo.push(item);
  }

  const linea = x => '   fila ' + x.fila + ' · ' + x.entrada + ' · ' + x.nombre
    + ' · abono $' + x.abono.toFixed(2) + ' de $' + x.total.toFixed(2)
    + ' · EstadoPago=' + x.estado;

  Logger.log('🟢 EL ARCHIVO ESTÁ EN DRIVE (' + enDrive.length + ') — solo falta enlazarlo');
  enDrive.forEach(x => {
    Logger.log(linea(x) + ' · match por ' + x.via);
    x.archivos.forEach(f => Logger.log('        ' + f.nombre));
  });
  if (!enDrive.length) Logger.log('   (ninguno)');

  Logger.log('');
  Logger.log('🔴 FALTA EL COMPROBANTE (' + sinArchivo.length + ') — se esperaba voucher y no está');
  Logger.log('   Reservas cargadas ANTES de la estadía: el pago pasó por el sistema y el');
  Logger.log('   archivo se perdió. El pago es real (el abono lo respalda); hay que re-subirlo.');
  sinArchivo.forEach(x => Logger.log(linea(x) + ' · creada ' + (x.creada || '?')));
  if (!sinArchivo.length) Logger.log('   (ninguno)');

  Logger.log('');
  Logger.log('⚪ CARGA RETROACTIVA (' + retro.length + ') — nunca hubo voucher que subir');
  Logger.log('   Filas creadas DESPUÉS del check-in: la estadía ya había pasado y el pago');
  Logger.log('   ocurrió fuera del sistema. No es un archivo perdido, no hay nada que hacer.');
  if (retro.length) {
    const porDia = {};
    retro.forEach(x => { porDia[x.creada || '?'] = (porDia[x.creada || '?'] || 0) + 1; });
    Object.keys(porDia).sort().forEach(d => Logger.log('   ' + d + ' · ' + porDia[d] + ' fila(s) cargadas'));
    Logger.log('   Rango de filas: ' + retro[0].fila + '–' + retro[retro.length - 1].fila
      + ' · $' + retro.reduce((s, x) => s + x.abono, 0).toFixed(2) + ' en abonos.');
  } else {
    Logger.log('   (ninguna)');
  }

  // ── Escritura ──
  let n = 0;
  if (enDrive.length) Logger.log('');
  enDrive.forEach(x => {
    const f = x.archivos[0];   // el más probable; si hay varios se enlaza uno solo
    if (x.archivos.length > 1) {
      Logger.log('   ⚠ fila ' + x.fila + ' tiene ' + x.archivos.length + ' archivos candidatos; se enlaza el primero.');
    }
    Logger.log((dryRun ? '[dry] ' : '✓ ') + 'fila ' + x.fila + ' · ' + x.nombre + ' → ' + f.nombre);
    if (!dryRun) {
      sheet.getRange(x.fila, 26).setValue(f.url);
      // VouchersMeta: el desglose pago por pago. Con un solo voucher, su monto ES
      // el abono registrado — no hay nada que repartir.
      sheet.getRange(x.fila, 32).setValue(JSON.stringify([
        { monto: x.abono, cod: '', fecha: x.entrada, url: f.url }
      ]));
    }
    n++;
  });
  if (!dryRun) SpreadsheetApp.flush();
  Logger.log('');
  Logger.log((dryRun ? '[dry-run] ' : '') + n + ' fila(s) ' + (dryRun ? 'se enlazarían' : 'enlazadas')
    + ' · ' + sinArchivo.length + ' a re-subir a mano · ' + retro.length + ' retroactivas (ignorar).');
  return { enDrive: enDrive.length, sinArchivo: sinArchivo.length, retro: retro.length };
}


// ═══════════════════════════════════════════════════════════
//  EstadoPago desincronizado con el abono
// ═══════════════════════════════════════════════════════════
//
// `EstadoPago` es una columna GUARDADA y solo se recalculaba al subir un voucher.
// Un pago registrado a mano en el campo Abono la dejaba en PENDIENTE para
// siempre, aunque el abono cubriera el total — mientras Contabilidad, que deriva
// el cobro de `montoRecibido()`, ya la daba por cobrada.
//
// El dashboard ahora deriva el estado en pantalla, así que esto es para dejar la
// HOJA coherente: los reportes y cualquier lectura externa leen la columna cruda.
//
// Solo SUBE el estado (pendiente → abonado → paga): nunca degrada un PAGA puesto
// a mano ni toca CANCELADA. Dry-run por defecto.
function sincronizarEstadoPagoConAbono(dryRun) {
  if (dryRun === undefined) dryRun = true;
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const iso   = v => v instanceof Date ? Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd') : (v || '').toString().slice(0, 10);
  const rank  = { '': 0, 'PENDIENTE': 1, 'ABONADO': 2, 'PAGA': 3 };

  Logger.log('');
  Logger.log('═══ ' + (dryRun ? 'DRY-RUN · ' : '') + 'SINCRONIZAR EstadoPago CON EL ABONO ═══');
  let n = 0, total = 0;
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[_R.ID]) continue;
    const origen = (r[_R.ORIGEN] || '').toString().trim();
    if (['Directa', 'Referido'].indexOf(origen) < 0) continue;
    const actual = (r[_R.ESTADO] || '').toString().trim().toUpperCase();
    if (actual === 'CANCELADA') continue;

    const tot = _cleanMoney_(r[_R.NETO]) || _cleanMoney_(r[_R.MONTO]);
    if (tot <= 0) continue;
    // Mismo criterio que montoRecibido() en el dashboard: el voucher manda; si no
    // hay voucher, el abono es lo recibido.
    const mv  = _cleanMoney_(r[_R.MONTOVOUCHER]);
    const rec = mv > 0 ? mv : _cleanMoney_(r[8]);

    const derivado = rec >= tot - 0.01 ? 'PAGA' : (rec > 0.01 ? 'ABONADO' : 'PENDIENTE');
    if ((rank[derivado] || 0) <= (rank[actual] || 0)) continue;

    Logger.log((dryRun ? '[dry] ' : '✓ ') + 'fila ' + (i + 1) + ' · ' + iso(r[_R.ENTRADA])
      + ' · ' + r[1] + ' · recibido $' + rec.toFixed(2) + ' de $' + tot.toFixed(2)
      + ' · ' + (actual || '(vacío)') + ' → ' + derivado);
    if (!dryRun) sheet.getRange(i + 1, 21).setValue(derivado);
    n++; total += rec;
  }
  if (!dryRun) SpreadsheetApp.flush();
  Logger.log('');
  Logger.log((dryRun ? '[dry-run] ' : '') + n + ' fila(s) ' + (dryRun ? 'se corregirían' : 'corregidas')
    + ' · $' + total.toFixed(2) + ' ya recibidos que la columna daba por pendientes.');
  return n;
}

function sincronizarEstadoPagoConAbonoESCRIBIR() {
  return sincronizarEstadoPagoConAbono(false);
}


// ═══════════════════════════════════════════════════════════
//  Reservas de Airbnb sin monto en la hoja
// ═══════════════════════════════════════════════════════════
//
// Caso Everett Richardson (HMSPN4MZFH, 26-29 jul 2026): la app de Airbnb dice
// $243.00 brutos / $205.33 de payout, pero la fila quedó con Monto y Neto en
// cero. En el dashboard eso se veía como "$0.00", indistinguible de una
// cortesía; ahora dice "sin monto", y esta función busca de dónde sacarlo.
//
// Ojo con la diferencia: que el payout no haya llegado es NORMAL —el neto de
// Airbnb se llena al cobrarse— pero el BRUTO se conoce desde el email de la
// reserva. Sin él, la reserva vale $0 en "Ingresos del período", diluye el ADR
// y resta a "Noches vendidas".
//
// Para cada fila sin monto busca su email por el código HM y extrae el Total
// con el mismo parser que usa el sync (`_montoAirbnbANumero`), así se ve qué
// habría que cargar antes de escribir nada.
//
// Solo lee y reporta. Para escribir: corregirMontosAirbnbDesdeEmailESCRIBIR().
function diagnosticarAirbnbSinMonto() {
  return _airbnbMontosDesdeEmail_(true);
}

function corregirMontosAirbnbDesdeEmailESCRIBIR() {
  return _airbnbMontosDesdeEmail_(false);
}

function _airbnbMontosDesdeEmail_(dryRun) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const iso   = v => v instanceof Date ? Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd') : (v || '').toString().slice(0, 10);
  const norm  = t => (t || '').toString().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, ' ').trim();

  // ── Índice de emails, UNA sola pasada ──────────────────────────────
  // El código HM NO se puede poner en la query de Gmail: viene dentro de una
  // URL (/details/HMXXXXXXXX) y Gmail no lo indexa como palabra. Buscar
  // `"HMSPN4MZFH"` devolvía cero para las nueve filas. El sync nunca lo hizo
  // así — trae todos los "Reserva confirmada:" y filtra el cuerpo en JS.
  // Además se indexa una vez y no una búsqueda por fila: nueve barridos de
  // Gmail es lento y arriesga la cuota.
  const porCod = {}, porNombre = {};
  let nMsgs = 0;
  GmailApp.search('from:automated@airbnb.com subject:"Reserva confirmada:" newer_than:400d')
    .forEach(t => t.getMessages().forEach(m => {
      const body = m.getPlainBody() || '';
      nMsgs++;
      // Mismo extractor que el sync: tolera "$243,00" y "243,00 $". Airbnb
      // cambió a la segunda forma y por eso estas filas entraron en cero.
      const monto = _montoEtiquetadoAirbnb_(body, 'Total\\s*\\(USD\\)');
      if (monto <= 0) return;
      // "Ganas" es lo que queda después de la comisión de Airbnb: es el neto
      // esperado de esta reserva, mejor dato que el bruto para el Neto.
      const ganas = _montoEtiquetadoAirbnb_(body, '\\bGanas\\b');
      const info = { monto: monto, ganas: ganas, asunto: m.getSubject().slice(0, 46) };
      const cm = body.match(/\b(HM[A-Z0-9]{8})\b/);
      if (cm) porCod[cm[1].toUpperCase()] = info;
      // Fallback por nombre: hay filas sin código HM (el id quedó de timestamp).
      const nm = m.getSubject().match(/Reserva confirmada:\s*(.+?)\s+llega/i);
      if (nm) {
        const k = norm(nm[1]);
        if (k) (porNombre[k] = porNombre[k] || []).push(info);
      }
    }));

  const casos = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[_R.ID]) continue;
    if (String(r[_R.ORIGEN] || '').trim() !== 'Airbnb') continue;
    if ((r[_R.ESTADO] || '').toString().toUpperCase() === 'CANCELADA') continue;
    if (_cleanMoney_(r[_R.MONTO]) > 0 || _cleanMoney_(r[_R.NETO]) > 0) continue;
    casos.push({ fila: i + 1, nombre: (r[1] || '').toString().replace(/\s+/g, ' ').trim(),
                 entrada: iso(r[_R.ENTRADA]),
                 cod: (r[_R.COD] || '').toString().trim().toUpperCase(),
                 pagado: _cleanMoney_(r[_R.MONTOPAGADO]) });
  }

  Logger.log('');
  Logger.log('═══ ' + (dryRun ? 'DRY-RUN · ' : '') + 'AIRBNB SIN MONTO EN LA HOJA ═══');
  Logger.log('📧 ' + nMsgs + ' emails "Reserva confirmada:" revisados · '
    + Object.keys(porCod).length + ' con código y monto');
  if (!casos.length) { Logger.log('✓ Todas las reservas de Airbnb activas tienen monto.'); return 0; }
  Logger.log(casos.length + ' fila(s) con Monto y Neto en cero.');
  Logger.log('');

  let n = 0, sinEmail = 0;
  casos.forEach(c => {
    let hit = /^HM[A-Z0-9]{8}$/.test(c.cod) ? porCod[c.cod] : null;
    let via = hit ? 'código' : '';
    if (!hit) {
      // Sin código utilizable, se cae al nombre — solo si es inequívoco.
      const cand = porNombre[norm(c.nombre)] || [];
      if (cand.length === 1) { hit = cand[0]; via = 'nombre (único)'; }
      else if (cand.length > 1) via = 'nombre ambiguo (' + cand.length + ' emails)';
    }
    const base = 'fila ' + c.fila + ' · ' + c.entrada + ' · ' + c.nombre
               + ' · ' + (c.cod || '(SIN CÓDIGO HM)');
    if (!hit) {
      sinEmail++;
      Logger.log('   ✗ ' + base + ' → sin email' + (via ? ' · ' + via : ''));
      return;
    }
    // El Neto es lo que realmente entra. Si el email trae "Ganas" se usa eso —
    // ya viene sin la comisión de Airbnb; si no, se cae al bruto para que la
    // reserva al menos no valga cero. `actualizarEstadoPagoAirbnb` lo pisa con
    // el neto real cuando el payout llega.
    const neto = hit.ganas > 0 ? hit.ganas : hit.monto;
    Logger.log((dryRun ? '   [dry] ' : '   ✓ ') + base + ' → bruto $' + hit.monto.toFixed(2)
      + (hit.ganas > 0 ? ' · neto $' + hit.ganas.toFixed(2) : ' · sin "Ganas", neto = bruto')
      + '   [por ' + via + ']');
    if (!dryRun) {
      sheet.getRange(c.fila, _R.MONTO + 1).setValue(hit.monto);
      if (c.pagado <= 0) sheet.getRange(c.fila, _R.NETO + 1).setValue(neto);
    }
    n++;
  });
  if (!dryRun) SpreadsheetApp.flush();
  Logger.log('');
  Logger.log((dryRun ? '[dry-run] ' : '') + n + (dryRun ? ' recuperables del email' : ' corregidas')
    + ' · ' + sinEmail + ' sin email que las respalde (cargar a mano).');
  return n;
}


// ═══════════════════════════════════════════════════════════
//  Voucher compartido cargado COMPLETO en cada cabaña
// ═══════════════════════════════════════════════════════════
//
// Un pago único que cubre dos cabañas tiene que repartirse entre las hermanas:
// `_saveMultiCabinReservation` lo prorratea, pero cuando el voucher se sube a
// mano —una reserva por vez— es fácil escribir el TOTAL del comprobante en cada
// fila. El sistema entonces lee el doble de lo cobrado: `montoRecibido()` da
// prioridad a `MontoVoucher`, así que dos filas de $75 con "$150" cada una se
// leen como $300 recibidos por una reserva de $150.
//
// Al re-subir los comprobantes perdidos de 2026 pasó en 4 filas ($315 de más).
//
// La huella es inconfundible: **varias filas comparten el mismo código de
// transferencia** (mismo comprobante) y alguna tiene `MontoVoucher` mayor que
// su propio `Monto`. El código se compara sin el `#` que a veces antepone el
// OCR y sin distinguir mayúsculas.
//
// La corrección reparte el voucher en proporción al monto de cada fila, igual
// que `_saveMultiCabinReservation`: la última absorbe el redondeo para que la
// suma dé exactamente el total del comprobante. También reescribe el monto en
// `VouchersMeta`, que es de donde el modal saca el desglose pago por pago.
//
// Solo lee y reporta. Para escribir: repartirVoucherCompartidoESCRIBIR().
function repartirVoucherCompartido(dryRun) {
  if (dryRun === undefined) dryRun = true;
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const iso   = v => v instanceof Date ? Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd') : (v || '').toString().slice(0, 10);
  // '#OEKKW-52729303' y 'OEKKW-52729303' son el MISMO comprobante: el OCR a
  // veces se trae el '#' del recibo. Sin normalizar, las hermanas no se agrupan
  // y el reparto no se detecta.
  const normCod = c => (c || '').toString().trim().replace(/^#+/, '').toUpperCase();

  // Agrupar filas por código de transferencia. Una fila con varios códigos
  // (abonos parciales, separados por '|') no entra: ahí el voucher ya es
  // acumulado y repartirlo requiere saber a qué pago corresponde cada parte.
  const grupos = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[_R.ID]) continue;
    const origen = (r[_R.ORIGEN] || '').toString().trim();
    if (['Directa', 'Referido'].indexOf(origen) < 0) continue;
    if ((r[_R.ESTADO] || '').toString().toUpperCase() === 'CANCELADA') continue;
    const cods = (r[_R.COD_TRANSF] || '').toString().split('|').map(normCod).filter(Boolean);
    if (cods.length !== 1) continue;
    (grupos[cods[0]] = grupos[cods[0]] || []).push({
      fila: i + 1, nombre: r[1], entrada: iso(r[_R.ENTRADA]), cabana: r[3],
      monto: _cleanMoney_(r[_R.MONTO]), voucher: _cleanMoney_(r[_R.MONTOVOUCHER]),
      meta: (r[31] || '').toString()
    });
  }

  Logger.log('');
  Logger.log('═══ ' + (dryRun ? 'DRY-RUN · ' : '') + 'VOUCHER COMPARTIDO MAL REPARTIDO ═══');
  let n = 0, exceso = 0;
  Object.keys(grupos).sort().forEach(cod => {
    const g = grupos[cod];
    if (g.length < 2) return;
    const totalMonto   = g.reduce((s, x) => s + x.monto, 0);
    const totalVoucher = g.reduce((s, x) => s + x.voucher, 0);
    // Sin exceso no hay nada que arreglar: o ya está repartido, o el grupo son
    // pagos legítimamente distintos que casualmente comparten código.
    if (totalVoucher <= totalMonto + 0.5) return;
    // El comprobante real es el MAYOR de los montos escritos: es el que alguien
    // copió entero. Los demás son porciones o copias del mismo número.
    const real = g.reduce((mx, x) => Math.max(mx, x.voucher), 0);
    if (totalMonto <= 0) {
      Logger.log('   ⚠ ' + cod + ': montos en cero, no se puede prorratear. Revisar a mano.');
      return;
    }
    const partes = _repartir_(real, g.map(x => x.monto));

    Logger.log('');
    Logger.log('📄 ' + cod + ' · comprobante $' + real.toFixed(2)
      + ' · ' + g.length + ' cabañas por $' + totalMonto.toFixed(2)
      + ' · registrado $' + totalVoucher.toFixed(2)
      + '  (sobra $' + (totalVoucher - real).toFixed(2) + ')');
    exceso += totalVoucher - real;
    g.forEach((x, k) => {
      Logger.log((dryRun ? '   [dry] ' : '   ✓ ') + 'fila ' + x.fila + ' · ' + x.entrada + ' · '
        + x.nombre + ' (' + x.cabana + ') · $' + x.voucher.toFixed(2) + ' → $' + partes[k].toFixed(2));
      if (!dryRun) {
        sheet.getRange(x.fila, 20).setValue('$' + partes[k].toFixed(2));
        // VouchersMeta guarda el desglose pago por pago; si no se corrige, el
        // modal sigue mostrando el monto viejo en "Pagos registrados".
        if (x.meta) {
          try {
            const arr = JSON.parse(x.meta);
            if (Array.isArray(arr) && arr.length === 1) {
              arr[0].monto = partes[k];
              sheet.getRange(x.fila, 32).setValue(JSON.stringify(arr));
            }
          } catch (_) { /* meta corrupta: se deja como está, el monto ya se corrigió */ }
        }
      }
      n++;
    });
  });
  if (!dryRun) SpreadsheetApp.flush();
  Logger.log('');
  Logger.log((dryRun ? '[dry-run] ' : '') + n + ' fila(s) ' + (dryRun ? 'se corregirían' : 'corregidas')
    + ' · $' + exceso.toFixed(2) + ' que el sistema contaba como cobrados de más.');
  return n;
}

function repartirVoucherCompartidoESCRIBIR() {
  return repartirVoucherCompartido(false);
}


// Escribe las URLs que diagnosticarVouchersSinArchivo() encontró en Drive.
// Solo toca filas cuya columna VoucherURL está VACÍA y cuyo archivo no está
// enlazado en ninguna otra fila. Idempotente.
// Timestamp que saveVoucherToDrive pone al final del nombre del archivo
// (cabin_nombre_checkin_YYYYMMDD_HHMMSS.ext). Sirve para quedarse con la subida
// más reciente cuando hay varias.
function _vhTimestamp(nombreArchivo) {
  const m = (nombreArchivo || '').match(/(\d{8}_\d{6})/);
  return m ? m[1] : '';
}

// Escribe las URLs que diagnosticarVouchersSinArchivo() encontró en Drive.
//
// Solo toca filas cuya columna VoucherURL está VACÍA, saltea los matches
// "probables" (los deja para revisión manual) y **enlaza como máximo tantos
// archivos como códigos de transferencia tenga la fila**: varias reservas
// tienen 2-3 archivos con el mismo id porque el voucher se re-subió, y escribir
// todos haría que el modal muestre "Pagos registrados (3)" para un solo pago.
// Se prefieren las subidas más recientes. Idempotente.
//
// Correr primero con vincularVouchersHuerfanos(true) para ver qué haría.
function vincularVouchersHuerfanos(dryRun) {
  // Default SEGURO: sin argumentos hace dry-run. El editor de Apps Script corre
  // las funciones sin parámetros, así que un default "escribir" convertía el
  // botón Run en una escritura silenciosa. Para escribir de verdad:
  // vincularVouchersHuerfanosESCRIBIR().
  if (dryRun === undefined) dryRun = true;
  const rec   = diagnosticarVouchersSinArchivo();
  const sheet = getOrCreateSheet();
  let n = 0, sobrantes = 0;
  Logger.log('');
  Logger.log('═══ ' + (dryRun ? 'DRY-RUN · ' : '') + 'VINCULAR VOUCHERS ═══');
  rec.recuperable.forEach(x => {
    if (x.via === 'huesped (probable)') {
      Logger.log('… fila ' + x.fila + ' (' + x.nombre + ') match solo PROBABLE, se salta. Revisar a mano: ' + x.urlsDrive);
      return;
    }
    const todos  = x.urlsDrive.split(' | ');
    const nombres = x.archivos.split(', ');
    // Ordenar por timestamp del nombre, más reciente primero
    const orden = nombres.map((nom, k) => ({ nom: nom, url: todos[k], ts: _vhTimestamp(nom) }))
      .sort((a, b) => b.ts < a.ts ? -1 : (b.ts > a.ts ? 1 : 0));
    // Cupo = cuántos archivos le faltan a esta fila. Cuando ya tiene alguna URL
    // (2 códigos y 1 archivo, p.ej.) se AGREGA la que falta en vez de saltearla:
    // saltearlas dejaba el pago sin archivo justamente en las filas mixtas.
    const cupo = Math.max(0, x.cods.split(',').length - x.urls);
    if (cupo === 0) return;
    const elegidos = orden.slice(0, cupo);
    if (orden.length > cupo) {
      sobrantes++;
      Logger.log('   ⚠ fila ' + x.fila + ' (' + x.nombre + ') tiene ' + orden.length
        + ' archivos para ' + cupo + ' pago(s); se enlaza(n) el/los más reciente(s). Resto: '
        + orden.slice(cupo).map(f => f.nom).join(', '));
    }
    // Preservar lo que ya estuviera en la celda y anexar lo nuevo.
    const previas = (sheet.getRange(x.fila, 26).getValue() || '').toString()
      .split('|').map(u => u.trim()).filter(Boolean);
    const valor = previas.concat(elegidos.map(f => f.url)).join('|');
    Logger.log((dryRun ? '[dry] ' : '✓ ') + 'fila ' + x.fila + ' · ' + x.nombre
      + (previas.length ? ' (+' + elegidos.length + ' a ' + previas.length + ' existente(s))' : '')
      + ' → ' + elegidos.map(f => f.nom).join(', '));
    if (!dryRun) sheet.getRange(x.fila, 26).setValue(valor);
    n++;
  });
  if (!dryRun) SpreadsheetApp.flush();
  Logger.log('');
  Logger.log((dryRun ? '[dry-run] ' : '') + n + ' fila(s) ' + (dryRun ? 'se vincularían' : 'vinculadas')
    + ' · ' + sobrantes + ' con archivos de sobra (revisar si hubo re-subidas).');
  return n;
}

// ═══════════════════════════════════════════════════════════
//  ¿Por qué una reserva de Airbnb sale "sin cobrar"?
// ═══════════════════════════════════════════════════════════
//
// Recorre la cadena completa y dice en qué eslabón se cortó:
//
//   email de payout  →  hoja Pagos  →  actualizarEstadoPagoAirbnb  →  Reservas
//
// `query` puede ser el nombre del huésped o el código HM… Ej:
//   diagnosticarPagoAirbnb('Yarisel')
//   diagnosticarPagoAirbnb('HMBBQXQ3QD')
//
// Solo lee y reporta, no escribe nada.
function diagnosticarPagoAirbnb(query) {
  const q = (query || '').toString().trim().toUpperCase();
  if (!q) { Logger.log('Pasá un nombre o un código HM…'); return; }
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  const resS = getOrCreateSheet();
  const data = resS.getDataRange().getValues();
  const iso  = v => v instanceof Date ? Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd') : (v || '').toString().slice(0, 10);

  // ── 1. La(s) fila(s) en Reservas
  const filas = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[_R.ID]) continue;
    const nombre = (r[1] || '').toString().toUpperCase();
    const cod    = (r[_R.COD] || '').toString().trim().toUpperCase();
    if (nombre.indexOf(q) < 0 && cod !== q) continue;
    filas.push({ fila: i + 1, nombre: r[1], cod: (r[_R.COD] || '').toString().trim(),
      origen: r[_R.ORIGEN], entrada: iso(r[_R.ENTRADA]), salida: iso(r[5]),
      monto: r[_R.MONTO], neto: r[_R.NETO],
      fechaPago: iso(r[16]), montoPagado: r[17], estado: r[20], alerta: r[13] });
  }
  Logger.log('═══ DIAGNÓSTICO DE PAGO · "' + query + '" ═══');
  if (!filas.length) { Logger.log('✗ No hay ninguna fila en Reservas que matchee.'); return; }
  filas.forEach(f => {
    Logger.log('');
    Logger.log('📋 RESERVAS fila ' + f.fila + ' · ' + f.nombre);
    Logger.log('   código=' + (f.cod || '(vacío)') + ' · origen=' + f.origen
      + ' · estadía ' + f.entrada + '→' + f.salida);
    Logger.log('   monto=' + f.monto + ' · neto=' + f.neto
      + ' · MontoPagado=' + (f.montoPagado || 0) + ' · FechaPago=' + (f.fechaPago || '(vacía)')
      + ' · EstadoPago=' + (f.estado || '(vacío)'));
    if (f.alerta) Logger.log('   alerta: ' + f.alerta);
  });

  // ── 2. La hoja Pagos
  const pagosS = ss.getSheetByName('Pagos');
  if (!pagosS) { Logger.log(''); Logger.log('✗ La hoja Pagos no existe. Corré syncAirbnbPayouts().'); return; }
  const pagos = pagosS.getDataRange().getValues().slice(1);
  const cods  = filas.map(f => f.cod.toUpperCase()).filter(Boolean);
  const hits  = [];
  pagos.forEach((row, k) => {
    const codes = (row[5] || '').toString().toUpperCase();
    const codsDeEsteRow = cods.filter(c => codes.indexOf(c) >= 0);
    if (codsDeEsteRow.length || (q.indexOf('HM') === 0 && codes.indexOf(q) >= 0)) {
      hits.push({ fila: k + 2, fechaCobro: iso(row[0]), montoNeto: row[4],
        codes: (row[5] || '').toString(), montos: (row[6] || '').toString(),
        // Qué códigos de LAS filas encontradas aparecen en este payout. Sin esto
        // el veredicto atribuía el cobro a todas las filas del resultado, y con
        // una reserva duplicada (una con código sintético y otra con el HM real)
        // decía que la sintética "solo necesita actualizarse" — falso, su código
        // no está en ningún payout.
        paraCods: codsDeEsteRow });
    }
  });
  Logger.log('');
  if (!hits.length) {
    Logger.log('💸 PAGOS: ✗ ningún payout menciona ese código.');
    Logger.log('   → El email de cobro no se sincronizó. Corré syncAirbnbPayouts() y después');
    Logger.log('     actualizarEstadoPagoAirbnb(). Si sigue sin aparecer, el email puede estar');
    Logger.log('     fuera de la ventana SYNC_VENTANA o no ser de automated@airbnb.com.');
  } else {
    hits.forEach(h => {
      Logger.log('💸 PAGOS fila ' + h.fila + ' · cobro ' + h.fechaCobro + ' · neto ' + h.montoNeto);
      const detalle = h.montos.split(',').filter(p => cods.some(c => p.toUpperCase().indexOf(c) >= 0));
      Logger.log('   monto asignado a este código: ' + (detalle.join(' ') || '(no desglosado)'));
    });
    Logger.log('   → El pago SÍ está registrado. Si la reserva sigue sin FechaPago, corré');
    Logger.log('     actualizarEstadoPagoAirbnb(): matchea Reservas.CodConfirmacion contra');
    Logger.log('     Pagos.ConfirmCodes, así que revisá que el código de la fila sea idéntico.');
  }

  // ── 3. Veredicto
  Logger.log('');
  Logger.log('🔎 VEREDICTO');
  filas.forEach(f => {
    if ((f.origen || '').toString().trim() !== 'Airbnb') {
      Logger.log('   fila ' + f.fila + ': origen es "' + f.origen + '", no "Airbnb" → actualizarEstadoPagoAirbnb la ignora.');
    } else if (!f.cod) {
      Logger.log('   fila ' + f.fila + ': no tiene CodConfirmacion → imposible cruzarla con Pagos.');
    } else if (!hits.some(h => h.paraCods.indexOf(f.cod.toUpperCase()) >= 0)) {
      Logger.log('   fila ' + f.fila + ': su código (' + f.cod + ') NO aparece en ningún payout.');
      if (f.cod.toLowerCase().indexOf('airbnb_') === 0) {
        Logger.log('      ↳ es un código SINTÉTICO (airbnb_…), no el HM real de Airbnb.');
        const hermana = filas.filter(o => o.fila !== f.fila && o.entrada === f.entrada && o.salida === f.salida);
        if (hermana.length) {
          Logger.log('      ↳ y hay otra fila con la MISMA estadía: fila ' + hermana.map(o => o.fila + ' (' + o.cod + ')').join(', '));
          Logger.log('      ↳ DUPLICADA: la buena es la del código HM. Esta habría que borrarla o cancelarla.');
        }
      }
    } else if (!f.fechaPago) {
      Logger.log('   fila ' + f.fila + ': el pago está en Pagos pero la reserva no se actualizó → correr actualizarEstadoPagoAirbnb().');
    } else {
      Logger.log('   fila ' + f.fila + ': ya tiene FechaPago ' + f.fechaPago + ' y MontoPagado ' + f.montoPagado + '.');
    }
  });
  return { filas: filas, pagos: hits };
}


// ═══════════════════════════════════════════════════════════
//  Atajos para el editor de Apps Script
// ═══════════════════════════════════════════════════════════
//
// El editor corre las funciones SIN argumentos (no hay dónde escribirlos), así
// que las que necesitan uno se manejan desde acá: editá la constante y corré la
// función de abajo, que aparece en el desplegable como cualquier otra.

// ── ¿Por qué una reserva de Airbnb sale sin cobrar? ──────────
// Poné el nombre del huésped o el código HM… y corré diagnosticarPagoAqui().
var DIAG_PAGO_QUERY = 'Yarisel';
function diagnosticarPagoAqui() {
  return diagnosticarPagoAirbnb(DIAG_PAGO_QUERY);
}

// ── Vincular vouchers huérfanos ──────────────────────────────
// vincularVouchersHuerfanos() ya hace dry-run por defecto (no escribe).
// Esta es la que escribe de verdad.
function vincularVouchersHuerfanosESCRIBIR() {
  return vincularVouchersHuerfanos(false);
}

// ── Pagos que solo están en el campo Abono ───────────────────
// Los que vincularVouchersHuerfanos() no ve porque no tienen código de
// transferencia. Reporte: diagnosticarPagosSinRastro().
// Escribe: vincularPagosSinRastroESCRIBIR().
//
// Y para dejar la columna EstadoPago coherente con el abono:
// sincronizarEstadoPagoConAbono() (dry-run) / …ESCRIBIR().

// ── Diagnóstico del email de una reserva ─────────────────────
var DIAG_EMAIL_QUERY = '';
function diagnosticarEmailAqui() {
  if (!DIAG_EMAIL_QUERY) { Logger.log('Editá DIAG_EMAIL_QUERY arriba con el nombre o código a buscar.'); return; }
  return diagnosticarEmailReserva(DIAG_EMAIL_QUERY, false);
}

// ── Reconciliar reservas de Airbnb ───────────────────────────
// Solo reporta las faltantes; para insertarlas, reconciliarAirbnbINSERTAR().
// Puesto en 2025 para recuperar a Michelle (HM5N4SPEJR, 7-8 ago 2025), que el
// reporte de alteraciones destapó como faltante: está en el export de Airbnb
// ($81.00) pero no tiene fila en Reservas.
var RECONCILIAR_DESDE = '2025-01-01';   // vacío = año en curso
function reconciliarAirbnbReporte() {
  return reconciliarReservasAirbnb(false, RECONCILIAR_DESDE || undefined);
}
function reconciliarAirbnbINSERTAR() {
  return reconciliarReservasAirbnb(true, RECONCILIAR_DESDE || undefined);
}

// ═══════════════════════════════════════════════════════════
//  Reservas de Airbnb duplicadas (sintética + HM real)
// ═══════════════════════════════════════════════════════════
//
// Cuando Airbnb notifica un cambio (fecha o monto), la fila vieja puede quedar
// con un código SINTÉTICO (`airbnb_<hash>`, generado por el sync cuando no pudo
// leer el HM del email) y con el monto anterior. Después, una reconciliación
// como insertarReservasAirbnbFaltantes() ve que el HM real "falta" —porque
// busca por código, no por fechas— e inserta una fila NUEVA. Resultado: dos
// filas para la misma estadía, la vieja sin cobrar y con el monto viejo.
//
// Esa vieja es la que aparece en "por cobrar" aunque Airbnb ya pagó.
//
// Agrupa por cabaña + entrada + salida y reporta los grupos con más de una fila.
// Solo lee y reporta.
function buscarReservasAirbnbDuplicadas() {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const iso   = v => v instanceof Date ? Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd') : (v || '').toString().slice(0, 10);
  const esSintetico = c => /^airbnb_/i.test(c) || !/^HM/i.test(c);

  const grupos = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[_R.ID]) continue;
    if ((r[_R.ORIGEN] || '').toString().trim() !== 'Airbnb') continue;
    const estado = (r[20] || '').toString().trim().toUpperCase();
    if (estado === 'CANCELADA') continue;
    const key = (r[3] || '') + '|' + iso(r[_R.ENTRADA]) + '|' + iso(r[5]);
    (grupos[key] = grupos[key] || []).push({
      fila: i + 1, nombre: r[1], cabana: r[3],
      entrada: iso(r[_R.ENTRADA]), salida: iso(r[5]),
      cod: (r[_R.COD] || '').toString().trim(),
      monto: r[_R.MONTO], montoPagado: r[17], fechaPago: iso(r[16]), estado: estado
    });
  }

  const dups = Object.keys(grupos).filter(k => grupos[k].length > 1).map(k => grupos[k]);
  Logger.log('═══ RESERVAS DE AIRBNB DUPLICADAS ═══');
  if (!dups.length) { Logger.log('✓ No hay estadías de Airbnb con más de una fila activa.'); return []; }

  let conSintetica = 0, mismoCodigo = 0, choqueFechas = 0;
  dups.forEach(g => {
    Logger.log('');
    Logger.log('⚠ ' + g[0].cabana + ' · ' + g[0].entrada + '→' + g[0].salida + ' · ' + g.length + ' filas:');
    g.forEach(f => {
      const tag = esSintetico(f.cod) ? '  ← código SINTÉTICO' : '';
      Logger.log('   fila ' + f.fila + ' · ' + f.nombre + ' · cod=' + (f.cod || '(vacío)')
        + ' · monto=' + f.monto + ' · pagado=' + (f.montoPagado || 0)
        + ' · ' + (f.fechaPago || 'sin fecha de pago') + ' · ' + (f.estado || '(sin estado)') + tag);
    });
    const buena = g.filter(f => !esSintetico(f.cod) && f.fechaPago);
    const mala  = g.filter(f => esSintetico(f.cod) && !f.fechaPago);
    // Distinguir los dos problemas que caen en el mismo grupo:
    //  · MISMO código HM en las dos filas → es LA MISMA reserva cargada dos
    //    veces. Duplicado inequívoco, y encima el ingreso se cuenta doble.
    //  · Códigos HM DISTINTOS → son dos reservas reales de Airbnb. Que choquen
    //    en cabaña+fechas casi siempre significa que una tiene las FECHAS mal
    //    en la hoja (año equivocado por el bug viejo de parseFecha, o una
    //    modificación de Airbnb que no se reflejó), no que sobre una fila.
    const codsHM = g.map(f => f.cod.toUpperCase()).filter(c => /^HM/.test(c));
    const mismoCod = codsHM.length === g.length && new Set(codsHM).size === 1;
    if (buena.length === 1 && mala.length) {
      conSintetica++;
      Logger.log('   → La buena es la fila ' + buena[0].fila + ' (' + buena[0].cod + ', cobrada).');
      Logger.log('     Sobra: fila ' + mala.map(f => f.fila).join(', ') + '. Borrala o marcala CANCELADA.');
    } else if (mismoCod) {
      mismoCodigo++;
      Logger.log('   → DUPLICADO SEGURO: las ' + g.length + ' filas tienen el MISMO código ('
        + codsHM[0] + '), o sea la misma reserva cargada dos veces.');
      Logger.log('     ⚠ El ingreso se está contando DOBLE. Dejá una sola: la del monto que');
      Logger.log('       coincide con Airbnb (revisá el export) y borrá la otra.');
    } else {
      choqueFechas++;
      Logger.log('   → NO es un duplicado: son reservas DISTINTAS (' + codsHM.join(' vs ') + ').');
      Logger.log('     Dos reservas no pueden ocupar la misma cabaña la misma noche, así que');
      Logger.log('     casi seguro una tiene las FECHAS mal en la hoja. Comparalas contra el');
      Logger.log('     export de Airbnb y corregí la fila equivocada (no borres ninguna).');
    }
  });
  Logger.log('');
  Logger.log('RESUMEN · ' + dups.length + ' grupo(s):');
  Logger.log('   ' + conSintetica + ' con fila sintética sobrante (borrar la sintética)');
  Logger.log('   ' + mismoCodigo + ' con el MISMO código HM repetido (duplicado real, ingreso doble)');
  Logger.log('   ' + choqueFechas + ' con códigos distintos (NO son duplicados: fechas mal en una fila)');
  return dups;
}

// ── Borrar filas de Reservas por número ──────────────────────
// Poné acá los números de fila que buscarReservasAirbnbDuplicadas() marcó como
// sobrantes y corré borrarFilasReservas() (dry-run) para revisar, después
// borrarFilasReservasESCRIBIR().
//
// Borra de MAYOR a MENOR a propósito: al eliminar una fila todas las de abajo
// suben un lugar, así que borrar en orden ascendente hace que los números
// siguientes apunten a la fila equivocada.
// Cada entrada puede ser un número de fila o —mejor— un objeto
// { fila, cod, monto } para que el script VERIFIQUE que la fila sigue siendo la
// que se quiere borrar. Los números de fila se corren si alguien borró algo en
// el medio, y borrar la fila equivocada acá es irreversible.
//
// Cargado con lo que devolvió buscarReservasAirbnbDuplicadas() (jul-2026):
//   69, 98   → filas con código sintético; la buena es la del HM
//   480,482,490 → mismo código HM que su gemela, con el monto inflado
// Duplicados pendientes de borrar. Los 5 de la tanda anterior (Yuliany, Kj,
// Dalit, Anna, Sabrina) ya se borraron el 26-jul-2026, así que salieron de acá.
//
// Queda la fila vieja de Mairanis: tiene el código sintético
// `airbnb_19e3be8ea9ba8b8e` —un msgId de Gmail, o sea que nació de un email de
// Airbnb— y por eso la reconciliación no encontró HMEQYEZS85 y lo insertó otra
// vez. Mientras exista, cualquier reconciliación futura puede volver a
// duplicarla. Está como origen "Abierta", así que no ocupa ni suma ingreso: el
// riesgo de borrarla es nulo. La fila 531 conserva el código HM real y el payout.
//
// Número de fila confirmado por verificarSaludAirbnb() el 26-jul-2026 a las
// 18:25, DESPUÉS de mover a Frank a Malaya y de borrar los 5 duplicados. Si se
// borra o inserta cualquier fila antes de la 316, volver a correr el chequeo:
// el guard incluye `id`, así que un índice corrido se rechaza en vez de borrar
// la fila equivocada.
var FILAS_A_BORRAR = [
  { fila: 316, id: 'airbnb_19e3be8ea9ba8b8e', cod: 'airbnb_19e3be8ea9ba8b8e', monto: 196.2 }
];


function borrarFilasReservas(dryRun) {
  if (dryRun === undefined) dryRun = true;   // default seguro (ver runners)
  const items = (FILAS_A_BORRAR || [])
    .map(x => (typeof x === 'object' ? x : { fila: x }))
    .filter(x => x.fila > 1)
    .sort((a, b) => b.fila - a.fila);
  if (!items.length) { Logger.log('FILAS_A_BORRAR está vacío. Editá la constante arriba.'); return 0; }
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  Logger.log('═══ ' + (dryRun ? 'DRY-RUN · ' : '') + 'BORRAR FILAS DE RESERVAS ═══');
  let n = 0, rechazadas = 0;
  items.forEach(it => {
    const f = it.fila;
    const r = data[f - 1];
    if (!r || !r[_R.ID]) { Logger.log('   ⚠ fila ' + f + ': vacía o inexistente, se salta.'); return; }
    // Verificación: si los índices se corrieron, el contenido no coincide y no
    // se borra nada. Vale más abortar que borrar una reserva buena.
    const codOk   = it.cod   === undefined || String(r[_R.COD] || '').trim() === String(it.cod).trim();
    const montoOk = it.monto === undefined || Math.abs((parseFloat(r[_R.MONTO]) || 0) - it.monto) < 0.01;
    // El `id` es el único discriminador fiable entre dos filas DUPLICADAS: las
    // dos de Yuliany comparten código (HMBACCYQM9) y monto ($178.20), así que
    // cod+monto no alcanza para saber cuál es cuál y un índice corrido borraría
    // la buena. Siempre poner `id` cuando se borra un duplicado.
    const idOk    = it.id    === undefined || String(r[_R.ID]  || '').trim() === String(it.id).trim();
    if (!codOk || !montoOk || !idOk) {
      rechazadas++;
      Logger.log('   ✗ fila ' + f + ': NO coincide con lo esperado, se salta.');
      Logger.log('      esperaba id=' + it.id + ' cod=' + it.cod + ' monto=' + it.monto
        + '  ·  encontró id=' + (r[_R.ID] || '(vacío)')
        + ' cod=' + (r[_R.COD] || '(vacío)') + ' monto=' + r[_R.MONTO] + ' (' + r[1] + ')');
      Logger.log('      → los números de fila se corrieron. Volvé a correr buscarReservasAirbnbDuplicadas().');
      return;
    }
    Logger.log((dryRun ? '   [dry] ' : '   ✓ ') + 'fila ' + f + ' · ' + r[1] + ' · ' + r[3]
      + ' · cod=' + (r[_R.COD] || '(vacío)') + ' · monto=' + r[_R.MONTO]
      + ' · pagado=' + (r[17] || 0));
    if (!dryRun) sheet.deleteRow(f);
    n++;
  });
  if (rechazadas) Logger.log('   ⚠ ' + rechazadas + ' fila(s) rechazadas por no coincidir.');
  if (!dryRun) SpreadsheetApp.flush();
  Logger.log('');
  Logger.log((dryRun ? '[dry-run] ' : '') + n + ' fila(s) ' + (dryRun ? 'se borrarían' : 'borradas')
    + ' (de mayor a menor, para no correr los índices).');
  return n;
}

function borrarFilasReservasESCRIBIR() {
  return borrarFilasReservas(false);
}

// ═══════════════════════════════════════════════════════════
//  Corregir reservas de Airbnb contra el export oficial
// ═══════════════════════════════════════════════════════════
//
// Filas cuyas fechas o montos no coinciden con lo que reporta Airbnb. Se
// busca por CÓDIGO, no por número de fila: los números se corren cada vez que
// se borra algo, el código no.
//
// Valores tomados del export "airbnb_01_202607_2026.csv" (jul-2026):
//   HMP5R5WYAF Kenneth        → la hoja tenía 04-04; Airbnb dice 05-09
//   HMD8PXQWMT Maria Celeste  → la hoja tenía 04-24; Airbnb dice 05-29
//   HMXZRKWEP8 Daniel Yanez   → la hoja tenía 2026; Airbnb dice 2025
//   HMJQRY88C3 Jennifer       → idem, año equivocado
//   HM3H889PFJ Aneea          → fechas OK, pero el monto era 99 y Airbnb dice 49.50
//
// Las tres primeras son el clásico choque de cabaña+noche que NO era duplicado:
// la fila tenía las fechas mal y por eso pisaba a otra reserva real.
function _correccionesAirbnb_() {
  return [
    { cod: 'HMP5R5WYAF', quien: 'Kenneth',       checkin: '2026-05-09', checkout: '2026-05-10' },
    { cod: 'HMD8PXQWMT', quien: 'Maria Celeste', checkin: '2026-05-29', checkout: '2026-05-30' },
    { cod: 'HMXZRKWEP8', quien: 'Daniel Yanez',  checkin: '2025-12-30', checkout: '2025-12-31' },
    // Mismo caso que Aneea, pero con política firme: la penalización fue del
    // 100% ($99 bruto / $83.65 neto, idéntico a una estadía normal), así que acá
    // no hay monto que corregir — solo sacarla de la ocupación. Daniel Yanez
    // reservó esa misma noche el 30-dic, el día del check-in, al liberarse.
    { cod: 'HMJQRY88C3', quien: 'Jennifer', checkin: '2025-12-30', checkout: '2025-12-31',
      estadoPago: 'CANCELADA',
      comentario: 'Cancelada por la huésped. Sin reembolso por política firme, Airbnb pagó la penalización completa ($83.65 neto). La fecha la retomó Daniel Yanez (HMXZRKWEP8).' },
    // Aneea canceló. La política no le daba reembolso, así que Airbnb igual pagó
    // la penalización ($49.50 bruto / $41.83 neto, la mitad de la tarifa) y
    // después liberó la fecha, que tomó Jakdiel. Por eso las dos aparecen la
    // misma noche en el export: la de Aneea es un cobro, no una estadía.
    // CANCELADA la saca de la ocupación sin borrar el ingreso: Contabilidad
    // calcula lo de Airbnb desde la hoja Pagos (los payouts), no desde las
    // reservas. Sí baja $41.83 en el widget de Ingresos, que suma reserva por
    // reserva y excluye las canceladas.
    { cod: 'HM3H889PFJ', quien: 'Aneea', monto: 49.50, neto: 41.83, estadoPago: 'CANCELADA',
      comentario: 'Cancelada por la huésped. Sin reembolso por política, Airbnb pagó la penalización ($41.83 neto). La fecha la retomó Jakdiel Moreno (HM54W53MJ9).' },
    // Export 2025. NO eran un choque de fechas: están en cabañas DISTINTAS y la
    // hoja tenía a Mario en la de Dianeth. Además ambos montos venían inflados y
    // figuraban sin cobrar, aunque el payout del 24-sep-2025 los pagó (ese
    // payout cuadra exacto: Σ netos = 719.73, y el export 2025 no lleva comisión
    // de Western Union).
    { cod: 'HMN4XNR44T', quien: 'Mario De León', cabin: 'azul',  monto: 85.50, neto: 82.93,
      fechaPago: '2025-09-24', montoPagado: 82.93, estadoPago: 'PAGA' },
    { cod: 'HMNYXJNQTH', quien: 'Dianeth Rueda', cabin: 'verde', monto: 90.00, neto: 87.30,
      fechaPago: '2025-09-24', montoPagado: 87.30, estadoPago: 'PAGA' },
    // Otra víctima del bug viejo de parseFecha: la hoja lo tenía en dic-2026
    // cuando el export dice dic-2025. Lo destapó el reporte de alteraciones, que
    // mostró una reserva de 2026-12 colgada de un cambio de nov-2025.
    { cod: 'HMQSNRTXYZ', quien: 'Gilberto Alexander Mitchell Rios',
      checkin: '2025-12-20', checkout: '2025-12-21', monto: 121.50, neto: 117.85 },

    // ── Las 9 reservas de 2025 que faltaban en la hoja ──────────
    // La reconciliación las encontró (todas reservadas entre el 22-jun y el
    // 19-jul-2025, cuando el sync todavía no estaba corriendo) y las inserta
    // desde el EMAIL. Ese monto viene inflado exactamente un 14.12%: el email
    // de 2025 traía el total que paga el huésped —con su tarifa de servicio—
    // mientras que el bruto del anfitrión es el del export. En los emails de
    // 2026 el formato cambió y el monto ya coincide (Mairanis, HMEQYEZS85,
    // entra correcta y por eso no está en esta tabla).
    // Además, en 4 de ellas el parser no reconoció el nombre del anuncio y
    // quedaron sin cabaña; el export las pone todas en Portal (azul).
    // Correr esto DESPUÉS de reconciliarAirbnbINSERTAR().
    // El cobro va acá y no lo resuelve actualizarEstadoPagoAirbnb: los payouts
    // de mediados de 2025 quedaron fuera de la ventana de syncAirbnbPayouts, así
    // que sus códigos no están en la hoja Pagos y estas 9 se verían "por cobrar"
    // para siempre ($838.50 de bruto). La columna `Fecha` del export ES la fecha
    // de payout: se validó contra Dianeth y Mario, que ya teníamos verificados
    // en 2025-09-24 con $87.30 y $82.93.
    { cod: 'HM2JSXHFE8', quien: 'Karldave',                 cabin: 'azul',  monto:  85.50, neto:  82.93,
      fechaPago: '2025-07-01', montoPagado:  82.93, estadoPago: 'PAGA' },
    { cod: 'HMNP53QXR8', quien: 'Manuel',                   cabin: 'verde', monto:  85.50, neto:  82.93,
      fechaPago: '2025-07-01', montoPagado:  82.93, estadoPago: 'PAGA' },
    { cod: 'HMPWNTCTZH', quien: 'Shelsy',                   cabin: 'verde', monto:  85.50, neto:  82.93,
      fechaPago: '2025-07-17', montoPagado:  82.93, estadoPago: 'PAGA' },
    { cod: 'HMH822D3SA', quien: 'Barbara',                  cabin: 'verde', monto:  90.00, neto:  87.30,
      fechaPago: '2025-07-17', montoPagado:  87.30, estadoPago: 'PAGA' },
    { cod: 'HMTJNZMTAR', quien: 'Stephanie',                cabin: 'verde', monto:  81.00, neto:  78.57,
      fechaPago: '2025-08-27', montoPagado:  78.57, estadoPago: 'PAGA' },
    { cod: 'HM5N4SPEJR', quien: 'Michelle Brockmann',       cabin: 'verde', monto:  81.00, neto:  78.57,
      fechaPago: '2025-08-27', montoPagado:  78.57, estadoPago: 'PAGA' },
    { cod: 'HMBX5SARFH', quien: 'Sandy Silvera Santamaria', cabin: 'azul',  monto: 110.00, neto: 106.70,
      fechaPago: '2025-08-27', montoPagado: 106.70, estadoPago: 'PAGA' },
    { cod: 'HM8AX95KQ8', quien: 'Edmundo Alexander',        cabin: 'azul',  monto: 110.00, neto: 106.70,
      fechaPago: '2025-09-01', montoPagado: 106.70, estadoPago: 'PAGA' },
    { cod: 'HMMFD55MMP', quien: 'Ana',                      cabin: 'azul',  monto: 110.00, neto: 106.70,
      fechaPago: '2025-10-09', montoPagado: 106.70, estadoPago: 'PAGA' },

    // Mairanis entró con el monto bien (el email de 2026 ya trae el bruto) pero
    // el neto quedó copiado del bruto. La comisión de 2026 es 15.5%:
    // 196.20 − 30.41 = 165.79, que es lo que pagó el payout del 22-jun.
    { cod: 'HMEQYEZS85', quien: 'Mairanis Lopez', neto: 165.79 },

    // Kenneth canceló el 8-may-2026, un día antes del check-in, y Airbnb pagó la
    // penalización ($92.10). syncCancelacionesAirbnb lo había marcado CANCELADA,
    // pero actualizarEstadoPagoAirbnb se lo revirtió a PAGA en la corrida
    // siguiente (el bug que se arregló hoy). Como el email de cancelación ya
    // figura procesado en la hoja Cancelaciones, no se va a volver a marcar
    // solo: hay que escribirlo desde acá. Al quedar CANCELADA se resuelve además
    // el choque del 9-may en azul con la directa de David Sauceda, que tomó esa
    // noche al liberarse — el mismo patrón de Aneea y Jennifer.
    // ── Filas de 2025 que quedaron SIN CABAÑA ───────────────────
    // Las destapó verificarSaludAirbnb(). Sin cabaña no aparecen en la ocupación
    // de ninguna, así que el calendario muestra esas noches libres y se pueden
    // vender dos veces — es exactamente lo que pasó con Maria Celeste y Mairanis.
    // El export las pone a las seis en Portal (azul). Falta ABRAHAM
    // (HM9DNW4QZ2), que no aparece en ningún export: revisar a mano.
    { cod: 'HM25N8TDMC', quien: 'Nadia',     cabin: 'azul', monto:  90.00, neto:  87.30,
      fechaPago: '2025-11-17', montoPagado:  87.30, estadoPago: 'PAGA' },
    { cod: 'HMWTRTFQFD', quien: 'Jair',      cabin: 'azul', monto:  90.00, neto:  87.30,
      fechaPago: '2025-09-24', montoPagado:  87.30, estadoPago: 'PAGA' },
    { cod: 'HM49BTRKJY', quien: 'Ricardo',   cabin: 'azul', monto:  85.50, neto:  82.93,
      fechaPago: '2025-09-24', montoPagado:  82.93, estadoPago: 'PAGA' },
    { cod: 'HMKD48NX88', quien: 'Monserrat', cabin: 'azul', monto:  90.00, neto:  87.30,
      fechaPago: '2025-10-09', montoPagado:  87.30, estadoPago: 'PAGA' },
    { cod: 'HMYFQ2WYP5', quien: 'John',      cabin: 'azul', monto: 180.00, neto: 174.60,
      fechaPago: '2025-09-01', montoPagado: 174.60, estadoPago: 'PAGA' },
    { cod: 'HMB248EDD9', quien: 'Giselle',   cabin: 'azul', monto:  85.50, neto:  82.93,
      checkin: '2025-08-03', checkout: '2025-08-04',
      fechaPago: '2025-08-27', montoPagado:  82.93, estadoPago: 'PAGA' },

    // ── Dos montos viejos confirmados contra el export ───────────
    // No los agarra corregirMontosInfladosAirbnb() porque su ratio no es el
    // 1.1412 de la tarifa de servicio: acá el monto quedó por DEBAJO de lo que
    // Airbnb pagó, así que hubo un cargo extra.
    //   Zulay   HM83ZQKJKW · subió a 4 viajeros → bruto real $165.30 (export).
    //   Nidemis HMPW3RYYNE · el export tiene DOS líneas, $99.00 de la reserva y
    //           $30.00 de una resolución. El bruto es $99.00; el neto se queda en
    //           $113.65 porque la resolución se paga sin comisión, y por eso el
    //           chequeo la va a seguir marcando como 🟡 — está bien así.
    // Otros tres que la derivación no puede resolver, tomados del export:
    //   Peter  HMSKS5DS9Q · el payout se partió en DOS líneas ($200.81 + $69.19)
    //          por ser estadía larga → bruto real $270.00, neto $261.90.
    //   Aitza  HMC3T2BQFJ · su alteración la movió del 30-dic al 31-ene y la
    //          tarifa SUBIÓ: la hoja tiene $99.00 y el bruto real es $109.00.
    //   Ivania HMN2FDYTRE · mismo caso al revés, la fecha nueva era más barata:
    //          la hoja tiene $112.98 y el bruto real es $81.00.
    // Carla Bonilla (HMWC8N44BR) NO va acá: su monto de $198.00 ya es correcto y
    // su neto de $68.31 refleja un ajuste de resolución de −$99.00. El chequeo la
    // va a marcar 🟡 para siempre y está bien — es un reembolso real.
    { cod: 'HMSKS5DS9Q', quien: 'Peter',  monto: 270.00, neto: 261.90 },
    { cod: 'HMC3T2BQFJ', quien: 'Aitza',  monto: 109.00, neto:  92.10 },
    { cod: 'HMN2FDYTRE', quien: 'Ivania', monto:  81.00, neto:  78.57 },
    { cod: 'HM83ZQKJKW', quien: 'Zulay Valles',   monto: 165.30 },
    { cod: 'HMPW3RYYNE', quien: 'Nidemis Mordock', monto:  99.00 },

    // ── Cobros de 2025 que quedaron fuera de la ventana ──────────
    // 24 reservas de 2025 cuyo payout quedó fuera de la ventana de
    // syncAirbnbPayouts, así que sus códigos no están en la hoja Pagos y
    // actualizarEstadoPagoAirbnb no puede resolverlas. Fecha de payout y neto
    // tomados del export; el monto también, porque el email de 2025 traía el
    // total del huésped (+14.12%).
    { cod: 'HMZRX9FHMK', quien: 'VANESSA', monto: 90.00, neto: 87.30,
      fechaPago: '2025-09-01', montoPagado: 87.30, estadoPago: 'PAGA' },
    { cod: 'HMZ8DB8RN8', quien: 'CESAR', monto: 100.00, neto: 97.00,
      fechaPago: '2025-09-01', montoPagado: 97.00, estadoPago: 'PAGA' },
    { cod: 'HMSANQ9R83', quien: 'VICTOR', monto: 100.00, neto: 97.00,
      fechaPago: '2025-09-24', montoPagado: 97.00, estadoPago: 'PAGA' },
    { cod: 'HMCST9QNTR', quien: 'YOLANY', monto: 100.00, neto: 97.00,
      fechaPago: '2025-09-24', montoPagado: 97.00, estadoPago: 'PAGA' },
    { cod: 'HMEZFFAKTN', quien: 'DANIELLE', monto: 81.00, neto: 78.57,
      fechaPago: '2025-09-24', montoPagado: 78.57, estadoPago: 'PAGA' },
    { cod: 'HMQP9AQ8SR', quien: 'MEYBOLL', monto: 90.00, neto: 87.30,
      fechaPago: '2025-10-09', montoPagado: 87.30, estadoPago: 'PAGA' },
    { cod: 'HMXAFEWYWM', quien: 'DANIEL', monto: 171.00, neto: 165.87,
      fechaPago: '2025-10-09', montoPagado: 165.87, estadoPago: 'PAGA' },
    { cod: 'HM3RWMYDN8', quien: 'Nery Elenna', monto: 90.00, neto: 87.30,
      fechaPago: '2025-10-09', montoPagado: 87.30, estadoPago: 'PAGA' },
    { cod: 'HMWXAQPAYW', quien: 'MELANIE', monto: 85.50, neto: 82.93,
      fechaPago: '2025-10-09', montoPagado: 82.93, estadoPago: 'PAGA' },
    { cod: 'HMK3JF2A3Q', quien: 'EVELYN', monto: 90.00, neto: 87.30,
      fechaPago: '2025-10-09', montoPagado: 87.30, estadoPago: 'PAGA' },
    { cod: 'HMSTWZA223', quien: 'GLADYS', monto: 90.00, neto: 87.30,
      fechaPago: '2025-10-21', montoPagado: 87.30, estadoPago: 'PAGA' },
    { cod: 'HMEBRF3HTF', quien: 'ESTEFANIS', monto: 81.00, neto: 78.57,
      fechaPago: '2025-10-21', montoPagado: 78.57, estadoPago: 'PAGA' },
    { cod: 'HM22CJE5JT', quien: 'MARISABEL', monto: 110.00, neto: 106.70,
      fechaPago: '2025-10-21', montoPagado: 106.70, estadoPago: 'PAGA' },
    { cod: 'HMCETSPRH4', quien: 'LIAM', monto: 180.00, neto: 174.60,
      fechaPago: '2025-10-21', montoPagado: 174.60, estadoPago: 'PAGA' },
    { cod: 'HMPDHZATAK', quien: 'ALEXANDRA', monto: 200.00, neto: 194.00,
      fechaPago: '2025-10-21', montoPagado: 194.00, estadoPago: 'PAGA' },
    { cod: 'HMTPP2WZDH', quien: 'JENNY', monto: 81.00, neto: 78.57,
      fechaPago: '2025-10-21', montoPagado: 78.57, estadoPago: 'PAGA' },
    { cod: 'HMPDFDDC4F', quien: 'JIAN', monto: 90.00, neto: 87.30,
      fechaPago: '2025-10-27', montoPagado: 87.30, estadoPago: 'PAGA' },
    { cod: 'HMW955DM43', quien: 'ALEX', monto: 90.00, neto: 87.30,
      fechaPago: '2025-10-30', montoPagado: 87.30, estadoPago: 'PAGA' },
    { cod: 'HMENBPD9YX', quien: 'CLAUDIA', monto: 104.50, neto: 101.36,
      fechaPago: '2025-11-17', montoPagado: 101.36, estadoPago: 'PAGA' },
    { cod: 'HMC4BDZKW5', quien: 'TANIA', monto: 90.00, neto: 87.30,
      fechaPago: '2025-11-17', montoPagado: 87.30, estadoPago: 'PAGA' },
    { cod: 'HMKKP5MKJ9', quien: 'PAOLA', monto: 90.00, neto: 87.30,
      fechaPago: '2025-11-17', montoPagado: 87.30, estadoPago: 'PAGA' },
    { cod: 'HMH53QPTWY', quien: 'Julio César', monto: 162.00, neto: 157.14,
      fechaPago: '2025-11-17', montoPagado: 157.14, estadoPago: 'PAGA' },
    { cod: 'HMSYJC95FS', quien: 'JOEL', monto: 99.00, neto: 96.03,
      fechaPago: '2025-11-17', montoPagado: 96.03, estadoPago: 'PAGA' },
    { cod: 'HMB59DDMES', quien: 'MARIBEL', monto: 110.00, neto: 106.70,
      fechaPago: '2025-11-17', montoPagado: 106.70, estadoPago: 'PAGA' },

    // ── Las 5 cancelaciones que la fila no reflejaba ─────────────
    // Las destapó el chequeo G de verificarSaludAirbnb() cruzando la hoja
    // Cancelaciones contra Reservas. Se parten en dos grupos según el export:
    //
    //   CON penalización cobrada (aparecen en el export, Airbnb pagó):
    //     Lena  HMF4R5HH49 · $75.29 neto, payout 14-abr-2026
    //     Paige HMMP54CHQ3 · $87.30 neto, payout 15-ene-2026 (su monto venía
    //           inflado: $103.74 en la hoja contra $90.00 de bruto real)
    //
    //   SIN cobro (no están en ningún export → cancelación libre, $0):
    //     Anna HMBKBXNSRR · Colton HMZZHYA985 · Abraham HM9DNW4QZ2
    //
    // A los tres sin cobro no se les toca el monto: al quedar CANCELADA salen de
    // la ocupación y de "por cobrar", que es lo que corresponde — son $548.10
    // que nunca se debieron. Abraham queda sin cabaña y así está bien: no está
    // en el export porque nunca se concretó, y una cancelada no ocupa nada.
    // ── Mairanis: crédito honrado en otras fechas ────────────────
    // No pudo venir el 19-21 jun por salud. Airbnb liberó el calendario y esas
    // dos noches se revendieron directo (Santiago el 19, Marilyn el 20). El
    // crédito se honró para el 7-9 ago y esa estadía ya está registrada aparte
    // (fila 395, $0, "Pagada el 20 de junio via Airbnb").
    // Por eso la fila del código HM va CANCELADA: es el mecanismo que la saca de
    // la ocupación —si no, esas noches figuran vendidas dos veces— sin perder el
    // ingreso, que Contabilidad toma de la hoja Pagos ($165.79 del payout del
    // 22-jun). Ojo: el widget de Ingresos sí resta ese monto, porque suma reserva
    // por reserva y excluye las canceladas.
    { cod: 'HMEQYEZS85', quien: 'Mairanis Lopez', neto: 165.79, estadoPago: 'CANCELADA',
      comentario: 'No pudo venir por salud. Airbnb liberó el 19-21 jun y esas noches se revendieron directo (Santiago 19, Marilyn 20). El crédito se honró para el 7-9 ago, registrado aparte. Airbnb pagó $165.79 (payout del 22-jun).' },

    { cod: 'HMF4R5HH49', quien: 'Lena Wiedmann', neto: 75.29, estadoPago: 'CANCELADA',
      comentario: 'Cancelada. Airbnb pagó la penalización ($75.29 neto, payout del 14-abr-2026).' },
    { cod: 'HMMP54CHQ3', quien: 'Paige', monto: 90.00, neto: 87.30, estadoPago: 'CANCELADA',
      comentario: 'Cancelada. Airbnb pagó la penalización ($87.30 neto, payout del 15-ene-2026).' },
    { cod: 'HMBKBXNSRR', quien: 'Anna',    estadoPago: 'CANCELADA',
      comentario: 'Cancelada sin cobro: no figura en el export de Airbnb, no hubo penalización.' },
    { cod: 'HMZZHYA985', quien: 'Colton',  estadoPago: 'CANCELADA',
      comentario: 'Cancelada sin cobro: no figura en el export de Airbnb, no hubo penalización.' },
    { cod: 'HM9DNW4QZ2', quien: 'Abraham', estadoPago: 'CANCELADA',
      comentario: 'Cancelada sin cobro: no figura en el export de Airbnb, no hubo penalización. Quedó sin cabaña asignada y no hay con qué determinarla; al estar cancelada no afecta la ocupación.' },

    { cod: 'HMP5R5WYAF', quien: 'Kenneth', estadoPago: 'CANCELADA',
      comentario: 'Cancelada por el huésped el 8-may-2026. Sin reembolso por política, Airbnb pagó la penalización ($92.10). La noche la retomó David Sauceda (directa).' }
  ];
}

function corregirReservasAirbnbDesdeExport(dryRun) {
  if (dryRun === undefined) dryRun = true;   // default seguro
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const iso   = v => v instanceof Date ? Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd') : (v || '').toString().slice(0, 10);
  Logger.log('═══ ' + (dryRun ? 'DRY-RUN · ' : '') + 'CORREGIR RESERVAS CONTRA EL EXPORT ═══');

  let n = 0, sinHallar = 0, yaOk = 0;
  _correccionesAirbnb_().forEach(c => {
    let fila = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][_R.COD] || '').trim().toUpperCase() === c.cod) { fila = i + 1; break; }
    }
    if (fila < 0) { sinHallar++; Logger.log('   ⚠ ' + c.cod + ' (' + c.quien + '): no se encontró ninguna fila con ese código.'); return; }
    const r = data[fila - 1];
    const cambios = [];
    if (c.checkin  && iso(r[_R.ENTRADA]) !== c.checkin)  cambios.push({ col: 5,  de: iso(r[_R.ENTRADA]), a: c.checkin,  que: 'entrada' });
    if (c.checkout && iso(r[5])          !== c.checkout) cambios.push({ col: 6,  de: iso(r[5]),          a: c.checkout, que: 'salida'  });
    if (c.monto != null && Math.abs((parseFloat(r[_R.MONTO]) || 0) - c.monto) > 0.01)
      cambios.push({ col: 8,  de: r[_R.MONTO], a: c.monto, que: 'monto' });
    if (c.neto  != null && Math.abs((parseFloat(r[_R.NETO])  || 0) - c.neto)  > 0.01)
      cambios.push({ col: 13, de: r[_R.NETO],  a: c.neto,  que: 'neto'  });
    // Cabaña: se escriben las DOS columnas, el nombre visible (3) y el código (4).
    if (c.cabin && String(r[3] || '').trim() !== c.cabin) {
      const NOMBRES = { verde: 'Paseo por Las Nubes', azul: 'Portal hacia Las Nubes', lila: 'Puente entre Las Nubes' };
      cambios.push({ col: 3, de: r[2], a: NOMBRES[c.cabin] || c.cabin, que: 'cabaña (nombre)' });
      cambios.push({ col: 4, de: r[3], a: c.cabin,                     que: 'cabaña (código)' });
    }
    if (c.fechaPago   && iso(r[16]) !== c.fechaPago)
      cambios.push({ col: 17, de: iso(r[16]) || '(vacía)', a: c.fechaPago, que: 'fecha de pago' });
    if (c.montoPagado != null && Math.abs((parseFloat(r[17]) || 0) - c.montoPagado) > 0.01)
      cambios.push({ col: 18, de: r[17] || 0, a: c.montoPagado, que: 'monto pagado' });
    if (c.estadoPago  && String(r[20] || '').trim() !== c.estadoPago)
      cambios.push({ col: 21, de: r[20] || '(vacío)', a: c.estadoPago, que: 'estado de pago' });
    // Comentario: se anexa (no pisa lo que el admin haya escrito) y solo si no
    // está ya, para que correr el script dos veces no lo duplique.
    if (c.comentario) {
      const actual = String(r[22] || '').trim();
      if (actual.indexOf(c.comentario) < 0) {
        cambios.push({ col: 23, de: actual || '(vacío)', a: actual ? actual + '\n' + c.comentario : c.comentario, que: 'comentario' });
      }
    }

    if (!cambios.length) { yaOk++; Logger.log('   ✓ ' + c.cod + ' (' + c.quien + ') fila ' + fila + ': ya está correcta.'); return; }
    Logger.log((dryRun ? '   [dry] ' : '   ✓ ') + c.cod + ' (' + c.quien + ') fila ' + fila + ':');
    cambios.forEach(x => Logger.log('        ' + x.que + ': ' + x.de + ' → ' + x.a));
    if (!dryRun) cambios.forEach(x => sheet.getRange(fila, x.col).setValue(x.a));
    n++;
  });
  if (!dryRun) SpreadsheetApp.flush();
  Logger.log('');
  Logger.log((dryRun ? '[dry-run] ' : '') + n + ' fila(s) ' + (dryRun ? 'se corregirían' : 'corregidas')
    + ' · ' + yaOk + ' ya estaban bien · ' + sinHallar + ' sin encontrar.');
  return n;
}

function corregirReservasAirbnbDesdeExportESCRIBIR() {
  return corregirReservasAirbnbDesdeExport(false);
}

// ═══════════════════════════════════════════════════════════
//  Reparar el emparejamiento de Alteraciones
// ═══════════════════════════════════════════════════════════
//
// El primer backfill de syncAirbnbUpdates() corrió con un bug: Sheets convierte
// el string 'yyyy-MM-dd HH:mm' que escribimos en FechaSolicitud a un valor de
// fecha, y al releerlo volvía un Date. La comparación de orden se hacía como
// string —y "Sat Mar 28 2026 22:40:00 GMT-0500" es MAYOR que "2026-03-28
// 22:43"— así que toda solicitud parecía posterior a su confirmación y el par
// se descartaba. Ninguna de las 26 alteraciones históricas se emparejó, aunque
// varias tenían la confirmación 2 o 3 minutos después de la solicitud.
//
// El bug ya está arreglado (_altTs en Parser.gs), pero syncAirbnbUpdates no
// reprocesa: los msgId quedaron registrados en la hoja. Esta función re-empareja
// usando los datos que YA están en Alteraciones —no vuelve a Gmail— y muestra
// exactamente qué cambiaría en cada reserva.
//
// Corré primero repararAlteracionesReporte() (no escribe nada) y revisá el log.
// Solo si el reporte se ve bien, corré repararAlteracionesAPLICAR().
//
// OJO al revisar: los cambios de FECHAS tienen red de seguridad (solo se
// escriben si la fila tiene hoy las fechas originales del email), pero los de
// VIAJEROS no la tienen — son un número y se escriben directo. Si un huésped
// pidió varios cambios de viajeros y alguno se rechazó, revisá ese caso a mano
// antes de aplicar.
function repararAlteracionesHuerfanas(dryRun) {
  if (dryRun !== false) dryRun = true;

  const ss       = SpreadsheetApp.openById(SHEET_ID);
  const altSheet = ss.getSheetByName('Alteraciones');
  if (!altSheet) { Logger.log('No existe la hoja Alteraciones. Nada que reparar.'); return 0; }

  const filas    = altSheet.getDataRange().getValues();
  const resSheet = getOrCreateSheet();
  const resData  = resSheet.getDataRange().getValues();

  const norm = t => (t || '').toString().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, ' ').trim();

  // Solicitudes huérfanas y confirmaciones que quedaron sin detalle.
  const solicitudes = [], confirmaciones = [];
  for (let k = 1; k < filas.length; k++) {
    const estado = (filas[k][6] || '').toString();
    if (filas[k][1] && (estado === 'solicitada' || estado === 'sin_confirmacion')) solicitudes.push(k);
    else if (filas[k][8] && estado === 'aceptada_sin_detalle') confirmaciones.push(k);
  }
  Logger.log((dryRun ? '[DRY-RUN · no escribe nada] ' : '[APLICANDO] ')
    + solicitudes.length + ' solicitud(es) huérfana(s) · '
    + confirmaciones.length + ' confirmación(es) sin detalle');
  Logger.log('');

  confirmaciones.sort((a, b) => _altTs(filas[a][7]).localeCompare(_altTs(filas[b][7])));

  const usada = {};
  let pares = 0, sinPareja = 0, aplicados = 0, pendientesSinFila = 0;

  confirmaciones.forEach(ci => {
    const fconf = _altTs(filas[ci][7]) || _altTs(filas[ci][0]);
    const cod   = (filas[ci][9] || '').toString().trim().toUpperCase();
    const hue   = norm(filas[ci][2]);
    const cab   = norm(filas[ci][3]);

    // "%{GUEST_NAME}": Airbnb no sustituyó la variable de plantilla. Sin nombre
    // se empareja por cercanía temporal, y solo si hay UN candidato.
    const sinNombre = _altNombreInutil(filas[ci][2]);
    let mejor = -1, candidatos = 0;
    solicitudes.forEach(si => {
      if (usada[si]) return;
      if (!sinNombre && norm(filas[si][2]) !== hue) return;
      const cabS = norm(filas[si][3]);
      if (cab && cabS && cabS !== cab) return;
      const fs = _altTs(filas[si][0]);
      if (!fs || fs > fconf) return;
      const gapMin = _altDiasEntre(fs, fconf) * 1440;
      if (sinNombre ? gapMin > ALT_FALLBACK_MIN : gapMin > ALT_VENTANA_DIAS * 1440) return;
      candidatos++;
      if (mejor < 0 || fs > _altTs(filas[mejor][0])) mejor = si;
    });

    if (sinNombre && candidatos !== 1) {
      Logger.log('⚠ ' + (cod || '(sin código)') + ': confirmación sin nombre utilizable y '
        + candidatos + ' candidato(s) por cercanía. No se arriesga a emparejar.');
      mejor = -1;
    } else if (sinNombre && mejor >= 0) {
      Logger.log('(nombre sin sustituir en el email; emparejado por cercanía temporal)');
    }

    if (mejor < 0) {
      sinPareja++;
      Logger.log('· ' + (cod || '(sin código)') + ' — ' + (filas[ci][2] || '?')
        + ' · ' + fconf + ': sigue sin solicitud emparejable.');
      return;
    }

    usada[mejor] = true;
    pares++;
    let cambios = [];
    try { cambios = JSON.parse(filas[mejor][5] || '[]'); } catch (_) {}

    const gap = _altDiasEntre(_altTs(filas[mejor][0]), fconf);
    Logger.log('■ ' + cod + ' — ' + (filas[ci][2] || '?') + ' (' + (filas[ci][3] || 'sin cabaña') + ')');
    Logger.log('    solicitud ' + _altTs(filas[mejor][0]) + '  →  confirmada ' + fconf
      + '   (' + (gap < 1 ? Math.round(gap * 1440) + ' min' : gap.toFixed(1) + ' días') + ')');
    _altPreviewCambios(resData, cod, cambios).forEach(l => Logger.log('    ' + l));

    if (!dryRun) {
      const res = _altAplicarCambios(resSheet, resData, cod, cambios, fconf);
      if (!res.filaEncontrada) {
        // Todavía no existe la reserva (ej. HM5N4SPEJR / Michelle, que falta
        // importar). NO se consume el par: si lo marcáramos procesado, cuando la
        // fila aparezca la alteración ya no se reintentaría nunca.
        usada[mejor] = false;
        pares--;
        pendientesSinFila++;
        Logger.log('    → sin fila con ese código; el par queda pendiente para reintentar.');
      } else {
        altSheet.getRange(mejor + 1, 7).setValue('aceptada');
        altSheet.getRange(mejor + 1, 8).setValue(fconf);
        altSheet.getRange(mejor + 1, 9).setValue(filas[ci][8]);
        altSheet.getRange(mejor + 1, 10).setValue(cod);
        altSheet.getRange(mejor + 1, 11).setValue(res.aplicado ? 'si' : 'parcial');
        altSheet.getRange(ci + 1, 7).setValue('emparejada');   // deja de contar como huérfana
        aplicados++;
        Logger.log('    → ' + res.detalle);
      }
    }
    Logger.log('');
  });

  if (!dryRun) SpreadsheetApp.flush();
  Logger.log('─────────────────────────────────────────');
  Logger.log((dryRun ? '[dry-run] ' : '') + pares + ' par(es) '
    + (dryRun ? 'se emparejarían' : 'emparejados')
    + (dryRun ? '' : ' · ' + aplicados + ' aplicado(s) a Reservas')
    + ' · ' + sinPareja + ' confirmación(es) siguen sin pareja'
    + (pendientesSinFila ? ' · ' + pendientesSinFila + ' pendiente(s) porque falta la reserva' : '') + '.');
  if (dryRun) Logger.log('Nada se escribió. Si el reporte se ve bien: repararAlteracionesAPLICAR()');
  return pares;
}

// Qué haría _altAplicarCambios, sin escribir. Devuelve líneas para el log.
function _altPreviewCambios(resData, cod, cambios) {
  const out = [];
  let fila = -1;
  for (let i = 1; i < resData.length; i++) {
    if ((resData[i][10] || '').toString().trim().toUpperCase() === cod) { fila = i; break; }
  }
  if (fila < 0) return ['⚠ no hay fila en Reservas con el código ' + cod + ' — no se aplicaría nada.'];

  const ciAct = _altISO(resData[fila][4]);
  const coAct = _altISO(resData[fila][5]);
  const perAct = resData[fila][6];
  out.push('hoy la reserva dice: ' + ciAct + ' → ' + coAct + ' · ' + perAct + ' persona(s)');

  if (!cambios.length) { out.push('sin detalle de cambios; no se aplicaría nada.'); return out; }

  cambios.forEach(c => {
    const que = _altNormQue(c.que);          // sin tildes: "HUÉSPEDES" → "HUESPEDES"
    if (que.indexOf('VIAJERO') === 0 || que.indexOf('HUESPED') === 0) {
      const n = parseInt((String(c.despues).match(/(\d+)/) || [])[1], 10);
      if (n > 0) {
        out.push(String(perAct) === String(n)
          ? '✓ personas: ya está en ' + n + ' (no cambia nada)'
          : '→ personas: ' + perAct + ' ⇒ ' + n + '   [sin red de seguridad: se escribe directo]');
        return;
      }
    }
    if (que.indexOf('FECHA') === 0) {
      const antes   = _altRangoFechas(c.antes,   _altISO(new Date()));
      const despues = _altRangoFechas(c.despues, _altISO(new Date()));
      if (!despues) { out.push('⚠ FECHAS ilegibles (' + c.despues + '): quedaría anotado para revisar.'); return; }
      if (ciAct === despues.ci && coAct === despues.co) {
        out.push('✓ fechas: ya están en ' + despues.ci + ' → ' + despues.co + ' (no cambia nada)');
      } else if (!antes || (ciAct === antes.ci && coAct === antes.co)) {
        out.push('→ fechas: ' + ciAct + ' → ' + coAct + '  ⇒  ' + despues.ci + ' → ' + despues.co);
      } else {
        out.push('⊘ fechas: NO se tocarían. El email parte de ' + antes.ci + ' → ' + antes.co
          + ' y la fila dice ' + ciAct + ' → ' + coAct + '. Se marcaría para revisar.');
      }
      return;
    }
    out.push('· ' + c.que + ': ' + c.antes + ' ⇒ ' + c.despues + '   [se anota, no se aplica]');
  });
  return out;
}

function repararAlteracionesReporte() { return repararAlteracionesHuerfanas(true); }
function repararAlteracionesAPLICAR() { return repararAlteracionesHuerfanas(false); }

// ═══════════════════════════════════════════════════════════
//  Chequeo de salud de los datos de Airbnb
// ═══════════════════════════════════════════════════════════
//
// Contesta "¿está todo bien?" con invariantes en vez de con una revisión a ojo.
// No escribe nada nunca: solo reporta. Correr después de cualquier tanda de
// arreglos y de vez en cuando (o dejarlo en un trigger semanal).
//
// Cada chequeo nació de un problema real de la auditoría de jul-2026, así que si
// alguno vuelve a dar positivo es una regresión concreta, no una sospecha.
// Hallazgos ya analizados que son correctos y no van a cambiar. El reporte los
// cuenta y los resume al pie con su motivo, en vez de repetirlos como si fueran
// problemas nuevos. Para dejar de aceptar uno, borrar su línea.
var EXCEPCIONES_SALUD = [
  { que: 'alteración aceptada sin aplicar', clave: 'HMD4ESKMPX',
    motivo: 'Kj tuvo 3 alteraciones encadenadas (16-18 → 17-19 → 18-21 mar). La del medio se rehúsa a propósito: la fila ya avanzó al rango final.' },
  { que: 'confirmación sin pareja', clave: 'HMJN58AQTF',
    motivo: 'La solicitud de Kaz quedó a 9 días de su confirmación, fuera de la ventana de 7. Emparejarla sería adivinar.' },
  { que: 'código raro', clave: '1782149615204',
    motivo: 'Fila creada a mano para el crédito honrado de Mairanis (7-9 ago). Nunca va a tener código HM porque no nació de una reserva de Airbnb.' },
  { que: 'neto no cuadra con la comisión', clave: 'Nidemis',
    motivo: 'Su payout incluye una resolución de +$30.00 que se paga sin comisión, así que el neto ($113.65) supera al bruto ($99.00). Es correcto.' },
  { que: 'neto no cuadra con la comisión', clave: 'Carla Bonilla',
    motivo: 'Ajuste de resolución de -$99.00 sobre una reserva de $198.00: Airbnb terminó pagando $68.31. Es un reembolso real.' }
];

function verificarSaludAirbnb() {
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  const hoja = getOrCreateSheet();
  const data = hoja.getDataRange().getValues();
  const hoy  = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');

  const iso = v => v instanceof Date
    ? Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd')
    : String(v || '').trim().slice(0, 10);
  const num = v => _cleanMoney_(v);

  // `impacto` (opcional) es el monto en juego, para sumarlo por categoría: en un
  // hallazgo sistémico de 167 filas el total importa más que el detalle.
  const problemas = [];   // { sev, que, detalle, impacto }
  const P = (sev, que, detalle, impacto) =>
    problemas.push({ sev: sev, que: que, detalle: detalle, impacto: impacto || 0 });

  // Filas de Airbnb activas (con su número de fila real).
  const filas = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (String(r[_R.ORIGEN] || '').trim() !== 'Airbnb') continue;
    filas.push({ n: i + 1, r: r });
  }

  // ── A) Código de confirmación válido ───────────────────────
  // Una fila sin HM real no cruza con los payouts, sale "sin cobrar" y hace que
  // la reconciliación inserte un duplicado (pasó con Yarisel, Yuliany, Kj).
  filas.forEach(f => {
    const c = String(f.r[_R.COD] || '').trim();
    if (!c)                       P('alta', 'sin código HM',      'fila ' + f.n + ' · ' + f.r[1]);
    else if (/^(airbnb_|csv_)/i.test(c)) P('alta', 'código sintético', 'fila ' + f.n + ' · ' + f.r[1] + ' · ' + c);
    else if (!/^HM[A-Z0-9]{8}$/i.test(c)) P('media', 'código raro',  'fila ' + f.n + ' · ' + f.r[1] + ' · ' + c);
  });

  // ── B) Códigos HM repetidos ────────────────────────────────
  // Mismo código en dos filas = ingreso contado DOBLE.
  const porCod = {};
  filas.forEach(f => {
    const c = String(f.r[_R.COD] || '').trim().toUpperCase();
    if (/^HM[A-Z0-9]{8}$/.test(c)) (porCod[c] = porCod[c] || []).push(f);
  });
  Object.keys(porCod).forEach(c => {
    if (porCod[c].length < 2) return;
    P('alta', 'código HM duplicado', c + ' en filas ' + porCod[c].map(f => f.n).join(', ')
      + ' (' + porCod[c].map(f => '$' + num(f.r[_R.MONTO])).join(' / ') + ')');
  });

  // ── B2) Códigos sintéticos en CUALQUIER origen ─────────────
  // Un `airbnb_<hash>` es la semilla de un duplicado futuro: la reconciliación
  // busca por código, no encuentra el HM real y lo inserta de nuevo. Pasó con
  // Mairanis, cuya fila vieja estaba como origen "Abierta" con código sintético
  // y por eso se escapó de los chequeos A y B, que solo mira(ba)n origen Airbnb.
  for (let i = 1; i < data.length; i++) {
    const c = String(data[i][_R.COD] || '').trim();
    if (!/^(airbnb_|csv_)/i.test(c)) continue;
    if (String(data[i][_R.ORIGEN] || '').trim() === 'Airbnb') continue;   // ya lo vio el chequeo A
    // Solo importa si el sufijo parece un msgId de Gmail (lleva letras
    // hexadecimales): eso significa que la fila SÍ nació de un email de Airbnb y
    // su HM real está sin registrar, así que la reconciliación lo va a insertar
    // otra vez. Un sufijo de puros dígitos es un timestamp local —hay 21 filas
    // Directa/Cortesía así, de una migración vieja— que no cruza con ningún HM y
    // no puede duplicar nada.
    if (!/[a-f]/i.test(c.replace(/^(airbnb_|csv_)+/i, ''))) continue;
    P('alta', 'código sintético de un email de Airbnb (semilla de duplicado)',
      'fila ' + (i + 1) + ' · ' + data[i][1] + ' · origen ' + (data[i][_R.ORIGEN] || '?') + ' · ' + c);
  }

  // ── B3) Misma persona, mismas fechas, misma cabaña ─────────
  // Atrapa el duplicado que B no ve porque una de las dos filas no tiene código
  // HM válido. Se compara sin importar el origen, que es justo donde se escondía
  // la fila de Mairanis.
  const porEstadia = {};
  for (let i = 1; i < data.length; i++) {
    const nom = String(data[i][1] || '').toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z]+/g, '');
    const ci = iso(data[i][_R.ENTRADA]), co = iso(data[i][5]), cab = String(data[i][3] || '').trim();
    if (!nom || !ci || !co) continue;
    const k = nom + '|' + cab + '|' + ci + '|' + co;
    (porEstadia[k] = porEstadia[k] || []).push(i + 1);
  }
  Object.keys(porEstadia).forEach(k => {
    if (porEstadia[k].length < 2) return;
    const p = k.split('|');
    P('alta', 'misma persona y fechas en dos filas', p[0] + ' · ' + (p[1] || 'sin cabaña')
      + ' · ' + p[2] + '→' + p[3] + ' · filas ' + porEstadia[k].join(', '));
  });

  // ── C) Cabaña asignada ─────────────────────────────────────
  // Sin cabaña la reserva no aparece en la ocupación de ninguna, y esa noche
  // queda libre en el calendario para venderla de nuevo. Las CANCELADAS se
  // saltean: no ocupan nada, así que su cabaña es irrelevante (Abraham no está
  // en ningún export justamente porque fue una cancelación sin cobro).
  filas.forEach(f => {
    if (String(f.r[_R.ESTADO] || '').trim().toUpperCase() === 'CANCELADA') return;
    if (!String(f.r[2] || '').trim() || !String(f.r[3] || '').trim())
      P('alta', 'sin cabaña', 'fila ' + f.n + ' · ' + f.r[1] + ' · ' + iso(f.r[_R.ENTRADA]));
  });

  // ── D) Neto coherente con el monto ─────────────────────────
  // La comisión de Airbnb es 3% hasta el 23-dic-2025 y 15.5% desde el 24.
  // Un neto IGUAL al monto significa que nunca se calculó (le pasó a Mairanis).
  //
  // Solo se mira si la fila YA tiene payout. Sin cobro registrado el Neto es un
  // campo dormido: el widget de Ingresos lee `fechaPago ? (neto||amount) : 0`,
  // así que no se usa, y cuando la plata llega lo llenan
  // actualizarEstadoPagoAirbnb + recalcularNetosAirbnb con el dato real. Marcarlo
  // antes reportaba 10 filas que no tenían nada que arreglar —3 de ellas
  // canceladas sin cobro, donde el neto no significa nada— y enterraba el único
  // hallazgo real del reporte.
  filas.forEach(f => {
    const mo = num(f.r[_R.MONTO]), ne = num(f.r[_R.NETO]);
    if (!mo || !ne) return;
    if (num(f.r[_R.MONTOPAGADO]) <= 0) return;                       // espera payout
    if (String(f.r[_R.ESTADO] || '').trim().toUpperCase() === 'CANCELADA') return;
    const ref  = iso(f.r[15]) || iso(f.r[_R.ENTRADA]);
    const pct  = ref && ref < '2025-12-24' ? 0.03 : 0.155;
    const esp  = Math.round(mo * (1 - pct) * 100) / 100;
    if (Math.abs(ne - mo) < 0.01 && mo > 0)
      P('media', 'neto igual al monto (nunca se calculó la comisión)',
        'fila ' + f.n + ' · ' + f.r[1] + ' · $' + mo + ' → debería ser ~$' + esp.toFixed(2),
        mo - esp);
    else if (Math.abs(ne - esp) > 1.00)
      P('baja', 'neto no cuadra con la comisión', 'fila ' + f.n + ' · ' + f.r[1]
        + ' · monto $' + mo + ' neto $' + ne + ' esperado ~$' + esp.toFixed(2), ne - esp);
  });

  // ── E) Estadías ya pasadas sin cobro registrado ────────────
  filas.forEach(f => {
    const co  = iso(f.r[5]);
    const est = String(f.r[_R.ESTADO] || '').trim().toUpperCase();
    if (!co || co >= hoy) return;
    if (est === 'PAGA' || est === 'CANCELADA') return;
    P('media', 'estadía pasada sin cobro registrado', 'fila ' + f.n + ' · ' + f.r[1]
      + ' · salió ' + co + ' · estado "' + (est || 'vacío') + '" · $' + num(f.r[_R.MONTO]),
      num(f.r[_R.MONTO]));
  });

  // ── F) Alteraciones aceptadas que no se aplicaron ───────────
  const alt = ss.getSheetByName('Alteraciones');
  if (alt) {
    const ad = alt.getDataRange().getValues();
    for (let k = 1; k < ad.length; k++) {
      const estado = String(ad[k][6] || '');
      if (estado === 'aceptada' && String(ad[k][10] || '') !== 'si')
        P('media', 'alteración aceptada sin aplicar', (ad[k][9] || '?') + ' · ' + (ad[k][2] || '?')
          + ' · ' + _altTs(ad[k][0]) + ' · ' + String(ad[k][5] || '').slice(0, 90));
      if (estado === 'aceptada_sin_detalle')
        P('baja', 'confirmación sin pareja', (ad[k][9] || '?') + ' · ' + (ad[k][2] || '?')
          + ' · ' + _altTs(ad[k][7]));
    }
  } else P('baja', 'falta la hoja Alteraciones', 'nunca corrió syncAirbnbUpdates()');

  // ── G) Cancelaciones que la fila no refleja ────────────────
  // Este es el canario del bug en que actualizarEstadoPagoAirbnb pisaba el
  // CANCELADA en cada corrida. Si vuelve a aparecer, volvió la regresión.
  const can = ss.getSheetByName('Cancelaciones');
  if (can) {
    const cd = can.getDataRange().getValues();
    for (let k = 1; k < cd.length; k++) {
      const cod = String(cd[k][2] || '').trim().toUpperCase();
      if (!cod) continue;
      (porCod[cod] || []).forEach(f => {
        if (String(f.r[_R.ESTADO] || '').trim().toUpperCase() !== 'CANCELADA')
          P('alta', 'cancelada en Airbnb pero no en la fila', cod + ' · fila ' + f.n + ' · '
            + f.r[1] + ' · estado "' + (f.r[_R.ESTADO] || 'vacío') + '"');
      });
    }
  }

  // ── H) Dos noches vendidas en la misma cabaña ──────────────
  // Solo entre reservas de tipo `noche` (o sin tipo), donde el rango es
  // inequívoco. Las pasadías/pasatardes bloquean días de cortesía a propósito,
  // así que compararlas produciría falsas alarmas.
  const noches = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const est = String(r[_R.ESTADO] || '').trim().toUpperCase();
    const org = String(r[_R.ORIGEN] || '').trim();
    const tipo = String(r[24] || '').trim().toLowerCase();
    if (est === 'CANCELADA' || org === 'Abierta') continue;
    if (tipo && tipo !== 'noche') continue;
    const cab = String(r[3] || '').trim();
    let ci = iso(r[_R.ENTRADA]), co = iso(r[5]);
    if (!cab || !ci || !co || co <= ci) continue;
    let d = new Date(ci + 'T12:00:00');
    const fin = new Date(co + 'T12:00:00');
    while (d < fin) {
      const k = cab + '|' + Utilities.formatDate(d, 'America/Panama', 'yyyy-MM-dd');
      (noches[k] = noches[k] || []).push('fila ' + (i + 1) + ' ' + String(r[1]).slice(0, 18)
        + ' [' + (org || '?') + ']');
      d = new Date(d.getTime() + 86400000);
    }
  }
  Object.keys(noches).sort().forEach(k => {
    if (noches[k].length < 2) return;
    P('alta', 'noche vendida dos veces', k.replace('|', ' · ') + ' → ' + noches[k].join('  vs  '));
  });

  // ── Excepciones aceptadas ──────────────────────────────────
  // Hallazgos ya analizados que son correctos y no van a cambiar. Sin esto el
  // reporte semanal repite las mismas 5 líneas para siempre, uno aprende a
  // ignorarlo entero y el día que aparezca algo real se pasa por alto.
  // Se cuentan y se resumen al pie con su motivo, no desaparecen.
  //
  // `clave` es un fragmento del detalle. Si se omite, se acepta la categoría
  // completa. Para dejar de aceptar algo, basta con borrar su línea de acá.
  const aceptadas = [], reales = [];
  problemas.forEach(p => {
    const ex = EXCEPCIONES_SALUD.filter(e =>
      p.que.indexOf(e.que) >= 0 && (!e.clave || p.detalle.indexOf(e.clave) >= 0))[0];
    if (ex) aceptadas.push({ p: p, motivo: ex.motivo }); else reales.push(p);
  });
  problemas.length = 0;
  reales.forEach(p => problemas.push(p));

  // ── Reporte ────────────────────────────────────────────────
  const orden = { alta: 0, media: 1, baja: 2 };
  problemas.sort((a, b) => orden[a.sev] - orden[b.sev] || a.que.localeCompare(b.que));
  const cuenta = { alta: 0, media: 0, baja: 0 };
  problemas.forEach(p => cuenta[p.sev]++);

  Logger.log('═══ SALUD DE LOS DATOS DE AIRBNB · ' + hoy + ' ═══');
  Logger.log(filas.length + ' reservas de Airbnb revisadas');
  Logger.log('');
  const pieAceptadas = () => {
    if (!aceptadas.length) return;
    Logger.log('');
    Logger.log('─────────────────────────────────────────');
    Logger.log('Además hay ' + aceptadas.length + ' hallazgo(s) ya revisado(s) y aceptado(s):');
    aceptadas.forEach(a => {
      Logger.log('   · ' + a.p.que + ' — ' + a.p.detalle.slice(0, 70));
      Logger.log('     ' + a.motivo);
    });
    Logger.log('Se listan acá para no olvidarlos, pero no son problemas.');
    Logger.log('(se configuran en EXCEPCIONES_SALUD, arriba de esta función)');
  };

  if (!problemas.length) {
    Logger.log('✅ Sin novedades. Los 8 chequeos pasaron:');
    Logger.log('   código HM válido · sin duplicados · con cabaña · neto coherente');
    Logger.log('   estadías pasadas cobradas · alteraciones aplicadas');
    Logger.log('   cancelaciones reflejadas · ninguna noche vendida dos veces');
    pieAceptadas();
    return 0;
  }
  // Se listan hasta MAX_DETALLE por categoría. Un hallazgo sistémico de 167
  // filas no se lee fila por fila: ahí lo que importa es el total, y listarlas
  // todas sepulta los 4 hallazgos que sí hay que atender uno por uno.
  const MAX_DETALLE = 8;
  const ICONO = { alta: '🔴', media: '🟠', baja: '🟡' };
  const grupos = [];
  problemas.forEach(p => {
    let g = grupos.filter(x => x.que === p.que)[0];
    if (!g) { g = { que: p.que, sev: p.sev, items: [], impacto: 0 }; grupos.push(g); }
    g.items.push(p.detalle);
    g.impacto += p.impacto;
  });

  grupos.forEach(g => {
    Logger.log('');
    Logger.log(ICONO[g.sev] + ' ' + g.que.toUpperCase() + '  (' + g.items.length + ')'
      + (Math.abs(g.impacto) >= 1 ? '  ·  en juego: $' + g.impacto.toFixed(2) : ''));
    g.items.slice(0, MAX_DETALLE).forEach(d => Logger.log('     ' + d));
    if (g.items.length > MAX_DETALLE)
      Logger.log('     … y ' + (g.items.length - MAX_DETALLE) + ' más (mismo patrón)');
  });

  Logger.log('');
  Logger.log('─────────────────────────────────────────');
  Logger.log('🔴 ' + cuenta.alta + ' grave(s) · 🟠 ' + cuenta.media + ' media(s) · 🟡 ' + cuenta.baja + ' leve(s)');
  Logger.log('🔴 = afecta plata u ocupación, revisar uno por uno.');
  Logger.log('🟠 = revisar el patrón; si son muchas suele ser una condición histórica.');
  Logger.log('🟡 = informativo.');
  pieAceptadas();
  return problemas.length;
}

// Chequeo semanal: lunes a las 8am Panamá. El resultado queda en el log de
// ejecuciones del editor (Ejecuciones → verificarSaludAirbnb).
function instalarTriggerSaludAirbnb() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'verificarSaludAirbnb') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('verificarSaludAirbnb')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .inTimezone('America/Panama')
    .create();
  Logger.log('✓ Trigger creado: verificarSaludAirbnb los lunes @ 8am America/Panama');
}

// ═══════════════════════════════════════════════════════════
//  Mover una reserva de Las Nubes a la hoja Malaya
// ═══════════════════════════════════════════════════════════
//
// Caso Frank Guilarte (29-30 may 2026, azul, $110): fue un doble booking por
// error y la reserva se le cedió a Malaya, pasándole el pago. Para Las Nubes es
// como si nunca hubiera entrado — pero sí existió como referido, así que va a la
// hoja Malaya en vez de borrarse sin más.
//
// Dry-run por defecto: sin argumentos NO escribe (borra una fila de Reservas).
// Para ejecutar de verdad: moverAMalayaESCRIBIR().
var MOVER_A_MALAYA = '1779577897186';   // id o CodConfirmacion de la reserva

function moverReservaAMalaya(codOrId, dryRun) {
  if (dryRun !== false) dryRun = true;
  const clave = String(codOrId || MOVER_A_MALAYA || '').trim();
  if (!clave) { Logger.log('Editá MOVER_A_MALAYA arriba con el id o código de la reserva.'); return 0; }

  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const iso   = v => v instanceof Date
    ? Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd')
    : String(v || '').trim().slice(0, 10);

  let fila = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][_R.ID] || '').trim() === clave ||
        String(data[i][_R.COD] || '').trim() === clave) { fila = i + 1; break; }
  }
  if (fila < 0) { Logger.log('⚠ No se encontró ninguna reserva con id/código ' + clave); return 0; }

  const r  = data[fila - 1];
  const ci = iso(r[_R.ENTRADA]), co = iso(r[5]);
  if (!ci || !co) { Logger.log('⚠ La fila ' + fila + ' no tiene fechas; no se puede migrar.'); return 0; }

  const noches   = Math.round((new Date(co + 'T12:00:00') - new Date(ci + 'T12:00:00')) / 86400000);
  const comision = _malayaCalculateCommission(ci, co);
  const monto    = _cleanMoney_(r[_R.MONTO]);

  Logger.log('═══ ' + (dryRun ? 'DRY-RUN · ' : '') + 'MOVER A MALAYA ═══');
  Logger.log('  Reservas fila ' + fila + ' · ' + r[1] + ' · ' + r[2]);
  Logger.log('    ' + ci + ' → ' + co + '  (' + noches + ' noche' + (noches === 1 ? '' : 's') + ')');
  Logger.log('    monto $' + monto.toFixed(2) + ' · personas ' + (r[6] || '?')
    + ' · tel ' + (r[23] || '(sin teléfono)') + ' · ' + (r[21] || '(sin email)'));
  Logger.log('    comisión calculada: $' + comision.toFixed(2)
    + '  ← si esta cesión NO se cobró como referido, poné 0 a mano en la hoja Malaya');
  Logger.log('  → se agrega a la hoja Malaya como "completada" y se BORRA de Reservas.');

  if (dryRun) {
    Logger.log('');
    Logger.log('Nada se escribió. Para ejecutar: moverAMalayaESCRIBIR()');
    return 0;
  }

  // Mismo orden de columnas que usa _malayaCreateReserva al hacer appendRow.
  _malayaSheet().appendRow([
    'LN-' + String(r[_R.ID] || '').trim(),        // id, prefijado para saber de dónde vino
    String(r[1] || '(sin nombre)'),               // huésped
    String(r[23] || ''),                          // teléfono
    ci, co, noches,
    r[6] || 2,                                    // personas
    monto, comision,
    'Referido', 'completada',
    true,                                         // airbnb_blocked: ya pasó, no hay nada que verificar
    iso(r[15]) || ci,                             // fecha_reserva
    'Cedida a Malaya: doble booking por error en ' + (r[2] || 'Las Nubes')
      + ' el ' + ci + '. El pago del huésped ($' + monto.toFixed(2) + ') se le pasó a Malaya. '
      + 'Venía de Reservas con id ' + (r[_R.ID] || '?') + '.',
    String(r[25] || ''),                          // voucherURL
    String(r[21] || '')                           // email
  ]);
  sheet.deleteRow(fila);
  SpreadsheetApp.flush();
  Logger.log('');
  Logger.log('✓ Agregada a Malaya y borrada de Reservas (fila ' + fila + ').');
  Logger.log('  OJO: los números de fila de Reservas se corrieron. Volvé a correr');
  Logger.log('  verificarSaludAirbnb() antes de usar cualquier lista de filas.');
  return 1;
}

function moverAMalayaESCRIBIR() { return moverReservaAMalaya(MOVER_A_MALAYA, false); }

// ═══════════════════════════════════════════════════════════
//  Recalcular el Neto de las reservas de Airbnb
// ═══════════════════════════════════════════════════════════
//
// En 159 filas el Neto es IGUAL al Monto, o sea que la comisión de Airbnb nunca
// se restó. Importa porque el widget de Ingresos usa `neto || amount`, así que
// sobrestima lo cobrado en ~$2.5k.
//
// No se estima con un porcentaje: cuando la fila tiene `MontoPagado`, ESE es el
// neto real —lo que Airbnb depositó— y es dato duro del payout. Las filas sin
// MontoPagado se dejan como están y se reportan aparte: adivinarlas con un 3% o
// 15.5% sería meter un número inventado en la contabilidad.
//
// Dry-run por defecto. Para escribir: recalcularNetosAirbnbESCRIBIR().
function recalcularNetosAirbnb(dryRun) {
  if (dryRun !== false) dryRun = true;
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();

  const arreglables = [], sinDato = [], raras = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (String(r[_R.ORIGEN] || '').trim() !== 'Airbnb') continue;
    const monto  = _cleanMoney_(r[_R.MONTO]);
    const neto   = _cleanMoney_(r[_R.NETO]);
    const pagado = _cleanMoney_(r[_R.MONTOPAGADO]);
    if (!monto || Math.abs(neto - monto) > 0.01) continue;   // el neto ya está distinto: no se toca
    if (pagado <= 0) { sinDato.push({ fila: i + 1, quien: r[1], monto: monto }); continue; }
    // Un pagado MAYOR que el monto significa que la tarifa subió (ej. Zulay, que
    // pasó a 4 viajeros) y el monto quedó viejo. El neto igual se corrige, pero
    // se avisa: ahí lo que está mal es el monto.
    if (pagado > monto + 0.01) raras.push({ fila: i + 1, quien: r[1], monto: monto, pagado: pagado });
    arreglables.push({ fila: i + 1, quien: r[1], monto: monto, pagado: pagado });
  }

  Logger.log('═══ ' + (dryRun ? 'DRY-RUN · ' : '') + 'RECALCULAR NETO DE AIRBNB ═══');
  Logger.log('');
  let dif = 0;
  arreglables.forEach(x => { dif += x.monto - x.pagado; });
  Logger.log(arreglables.length + ' fila(s) con MontoPagado: el neto pasa a ese valor exacto.');
  Logger.log('  El total cobrado baja $' + dif.toFixed(2) + ' — no es plata perdida, es la');
  Logger.log('  comisión de Airbnb que antes no se estaba restando.');
  arreglables.slice(0, 10).forEach(x => Logger.log('     fila ' + x.fila + ' · ' + x.quien
    + ' · neto $' + x.monto.toFixed(2) + ' → $' + x.pagado.toFixed(2)));
  if (arreglables.length > 10) Logger.log('     … y ' + (arreglables.length - 10) + ' más');

  if (raras.length) {
    Logger.log('');
    Logger.log('⚠ ' + raras.length + ' fila(s) donde lo pagado SUPERA el monto: ahí el que está mal');
    Logger.log('  es el monto, no el neto (suele ser una tarifa que subió por huésped extra).');
    raras.forEach(x => Logger.log('     fila ' + x.fila + ' · ' + x.quien
      + ' · monto $' + x.monto.toFixed(2) + ' < pagado $' + x.pagado.toFixed(2)));
  }

  if (sinDato.length) {
    Logger.log('');
    Logger.log('· ' + sinDato.length + ' fila(s) sin MontoPagado: NO se tocan. Sin el payout no hay');
    Logger.log('  con qué saber el neto real, y estimarlo con un % sería inventar un número.');
  }

  if (dryRun) {
    Logger.log('');
    Logger.log('Nada se escribió. Para ejecutar: recalcularNetosAirbnbESCRIBIR()');
    return 0;
  }
  arreglables.forEach(x => sheet.getRange(x.fila, _R.NETO + 1).setValue(x.pagado));
  SpreadsheetApp.flush();
  Logger.log('');
  Logger.log('✓ ' + arreglables.length + ' neto(s) actualizado(s).');
  return arreglables.length;
}

function recalcularNetosAirbnbESCRIBIR() { return recalcularNetosAirbnb(false); }

// ═══════════════════════════════════════════════════════════
//  Montos inflados de 2025 · derivados del payout
// ═══════════════════════════════════════════════════════════
//
// El email de Airbnb de 2025 traía el total que paga el HUÉSPED (con su tarifa
// de servicio), no el bruto del anfitrión: exactamente +14.12%. Una vez que el
// Neto tiene el valor real del payout, el bruto verdadero se puede DERIVAR sin
// necesidad del export:  bruto = neto / (1 − comisión).
//
// Se corrige SOLO cuando el ratio monto/derivado da 1.1412 clavado. Esa firma es
// autoevidente: es la tarifa de servicio del huésped y nada más da ese número.
//
// Todo lo demás se reporta pero NO se toca, porque la derivación se rompe si el
// payout incluye una resolución. Caso testigo: Nidemis Mordock tiene DOS líneas
// en el export —$99.00 de la reserva y $30.00 de una resolución— y el $30 no
// lleva comisión. La fórmula habría dado $134.50 cuando el bruto real es $99.
//
// Dry-run por defecto. Para escribir: corregirMontosInfladosESCRIBIR().
const TARIFA_SERVICIO_HUESPED_2025 = 1.1412;

function corregirMontosInfladosAirbnb(dryRun) {
  if (dryRun !== false) dryRun = true;
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const iso = v => v instanceof Date
    ? Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd')
    : String(v || '').trim().slice(0, 10);

  const corregir = [], revisar = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (String(r[_R.ORIGEN] || '').trim() !== 'Airbnb') continue;
    const monto  = _cleanMoney_(r[_R.MONTO]);
    const pagado = _cleanMoney_(r[_R.MONTOPAGADO]);
    if (!monto || pagado <= 0) continue;

    // La comisión NO se elige por fecha. Elegirla así fallaba: a Tiffany le tomó
    // 15.5% por su FechaPago y derivaba $103.31, cuando la verdad es 3% y $90.00.
    // Se prueban las dos tasas conocidas y gana la que produzca la firma exacta
    // del 1.1412; si ninguna la produce, no se toca. Así el dato decide en vez de
    // una heurística de fechas, y de paso queda auto-verificado: que el ratio dé
    // 1.1412 con una tasa válida es evidencia de que ESA es la tasa aplicada.
    const PCTS = [0.03, 0.155];
    let elegido = null, yaEstaBien = false;
    for (let k = 0; k < PCTS.length; k++) {
      const der = Math.round(pagado / (1 - PCTS[k]) * 100) / 100;
      if (Math.abs(monto - der) < 0.5) { yaEstaBien = true; break; }
      if (Math.abs(monto / der - TARIFA_SERVICIO_HUESPED_2025) < 0.002)
        elegido = { derivado: der, pct: PCTS[k], ratio: monto / der };
    }
    if (yaEstaBien) continue;

    if (elegido) {
      corregir.push({ fila: i + 1, quien: r[1], monto: monto, pagado: pagado,
                      derivado: elegido.derivado, ratio: elegido.ratio, pct: elegido.pct });
    } else {
      // Se muestran las dos derivaciones posibles para que el admin vea que
      // ninguna cuadra y por qué no se tocó.
      revisar.push({ fila: i + 1, quien: r[1], monto: monto, pagado: pagado,
                     der3: Math.round(pagado / 0.97 * 100) / 100,
                     der15: Math.round(pagado / 0.845 * 100) / 100 });
    }
  }

  Logger.log('═══ ' + (dryRun ? 'DRY-RUN · ' : '') + 'MONTOS INFLADOS DE AIRBNB ═══');
  Logger.log('');
  Logger.log(corregir.length + ' fila(s) con la firma exacta del +14.12% (tarifa de servicio del');
  Logger.log('huésped). El bruto real se deriva del payout, no se estima:');
  corregir.slice(0, 12).forEach(x => Logger.log('     fila ' + x.fila + ' · ' + x.quien
    + ' · $' + x.monto.toFixed(2) + ' → $' + x.derivado.toFixed(2)
    + '   (neto $' + x.pagado.toFixed(2) + ' ÷ ' + (1 - x.pct).toFixed(3) + ')'));
  if (corregir.length > 12) Logger.log('     … y ' + (corregir.length - 12) + ' más');

  if (revisar.length) {
    Logger.log('');
    Logger.log('⚠ ' + revisar.length + ' fila(s) que NO se tocan: el ratio no es 1.1412, así que la');
    Logger.log('  derivación no es de fiar (un cargo de resolución en el payout la rompe).');
    revisar.forEach(x => Logger.log('     fila ' + x.fila + ' · ' + x.quien
      + ' · monto $' + x.monto.toFixed(2) + ' · neto $' + x.pagado.toFixed(2)
      + '  →  con 3% daría $' + x.der3.toFixed(2) + ', con 15.5% daría $' + x.der15.toFixed(2)
      + ' · ninguna da la firma 1.1412'));
  }

  if (dryRun) {
    Logger.log('');
    Logger.log('Nada se escribió. Para ejecutar: corregirMontosInfladosESCRIBIR()');
    return 0;
  }
  corregir.forEach(x => sheet.getRange(x.fila, _R.MONTO + 1).setValue(x.derivado));
  SpreadsheetApp.flush();
  Logger.log('');
  Logger.log('✓ ' + corregir.length + ' monto(s) corregido(s).');
  return corregir.length;
}

function corregirMontosInfladosESCRIBIR() { return corregirMontosInfladosAirbnb(false); }

// ═══════════════════════════════════════════════════════════
//  Liberar las fechas de los créditos abiertos
// ═══════════════════════════════════════════════════════════
//
// Origen `Abierta` significa CRÉDITO SIN FECHA: el huésped pagó, no pudo venir y
// en vez de pedir devolución dejó la fecha a coordinar. Pero varias de esas filas
// conservaron la fecha original, y eso cuesta plata de dos formas:
//
//   · esas noches figuran vendidas y NO se pueden vender, aunque el huésped ya
//     avisó que no viene;
//   · inflan la ocupación y los ingresos del mes en que nunca hubo nadie.
//
// El dashboard ya trata `Abierta` como sin fecha (el modal oculta los campos y
// la ocupación filtra por checkin/checkout), así que vaciar las fechas es lo que
// alinea el dato con lo que el sistema ya asume. NO se toca el monto ni el
// voucher: son el respaldo del crédito.
//
// El modal ya no puede volver a crear este estado (al elegir Abierta vacía las
// fechas), así que esto es una migración de una sola vez.
//
// Dry-run por defecto. Para ejecutar: liberarFechasAbiertasESCRIBIR()
function liberarFechasAbiertas(dryRun) {
  if (dryRun !== false) dryRun = true;
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const iso = v => v instanceof Date
    ? Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd')
    : String(v || '').trim().slice(0, 10);
  const hoy = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');

  const conFecha = [], sinFecha = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (String(r[_R.ORIGEN] || '').trim() !== 'Abierta') continue;
    const ci = iso(r[_R.ENTRADA]), co = iso(r[5]);
    const item = { fila: i + 1, quien: r[1], cabana: r[2], ci: ci, co: co,
                   monto: _cleanMoney_(r[_R.MONTO]), futura: ci >= hoy };
    if (ci || co) conFecha.push(item); else sinFecha.push(item);
  }

  Logger.log('═══ ' + (dryRun ? 'DRY-RUN · ' : '') + 'LIBERAR FECHAS DE CRÉDITOS ABIERTOS ═══');
  Logger.log('');
  Logger.log(conFecha.length + ' crédito(s) con fecha puesta · ' + sinFecha.length + ' ya sin fecha (correctos)');
  if (!conFecha.length) {
    Logger.log('');
    Logger.log('✅ Nada que liberar: todos los créditos abiertos ya están sin fecha.');
    return 0;
  }
  Logger.log('');
  let noches = 0, plata = 0, futuras = 0;
  conFecha.forEach(x => {
    const n = (x.ci && x.co)
      ? Math.round((new Date(x.co + 'T12:00:00') - new Date(x.ci + 'T12:00:00')) / 86400000) : 0;
    noches += n; plata += x.monto;
    if (x.futura) futuras++;
    Logger.log('   fila ' + x.fila + ' · ' + x.quien + ' · ' + (x.cabana || 'sin cabaña'));
    Logger.log('        ' + (x.ci || '?') + ' → ' + (x.co || '?') + '  (' + n + ' noche' + (n === 1 ? '' : 's') + ')'
      + ' · $' + x.monto.toFixed(2)
      + (x.futura ? '  ← A FUTURO: al liberarla podés vender esas noches' : '  · ya pasó'));
  });
  Logger.log('');
  Logger.log('Total: ' + noches + ' noche(s) que dejan de figurar ocupadas · $' + plata.toFixed(2)
    + ' que pasan de "ingreso" a "crédito por usar".');
  if (futuras) Logger.log(futuras + ' de esos créditos son de fechas FUTURAS: esas noches vuelven a estar a la venta.');

  if (dryRun) {
    Logger.log('');
    Logger.log('Nada se escribió. Para ejecutar: liberarFechasAbiertasESCRIBIR()');
    return 0;
  }
  conFecha.forEach(x => {
    sheet.getRange(x.fila, _R.ENTRADA + 1).setValue('');
    sheet.getRange(x.fila, 6).setValue('');
    const cAct = String(data[x.fila - 1][_R.COMENT] || '').trim();
    const nota = '📅 Fechas liberadas (' + hoy + '): crédito abierto, estadía a coordinar. '
      + 'Fechas originales ' + (x.ci || '?') + ' → ' + (x.co || '?') + '.';
    if (cAct.indexOf('Fechas liberadas') < 0) {
      sheet.getRange(x.fila, _R.COMENT + 1).setValue(cAct ? cAct + '\n' + nota : nota);
    }
  });
  SpreadsheetApp.flush();
  Logger.log('');
  Logger.log('✓ ' + conFecha.length + ' crédito(s) liberado(s). El monto y el voucher quedaron intactos.');
  Logger.log('  Cuando el huésped elija fecha: poné las fechas y cambiá el origen a Directa.');
  return conFecha.length;
}

function liberarFechasAbiertasESCRIBIR() { return liberarFechasAbiertas(false); }

// ═══════════════════════════════════════════════════════════
//  Borrar créditos abiertos vencidos
// ═══════════════════════════════════════════════════════════
//
// Un crédito que el huésped nunca usó y ya venció. Se borra la fila en vez de
// dejarla: sin fecha, sin monto y sin ocupación no aparece en ninguna vista, así
// que quedaría como fila huérfana que nadie puede interpretar después.
//
// Se identifican por CONTENIDO (huésped + fecha original + monto), no por número
// de fila: los índices se corren con cada borrado y una lista de filas queda
// obsoleta al primer cambio. El guard exige que la fila sea `Abierta` y que los
// tres campos coincidan, así que si algo no cuadra no borra nada.
//
// OJO: si el crédito vencido TIENE monto, borrarlo pierde el rastro de una plata
// que sí entró. En ese caso conviene reconocerlo como ingreso (el huésped lo
// perdió) en vez de borrarlo — hablalo antes de agregarlo a esta lista.
var CREDITOS_VENCIDOS = [
  // Crédito de marzo-2026 que nunca se usó. Sin monto registrado, así que no hay
  // plata que reconocer: se borra sin más.
  { quien: 'Daniela Silva', checkin: '2026-03-07', monto: 0 }
];

function borrarCreditosVencidos(dryRun) {
  if (dryRun !== false) dryRun = true;
  if (!CREDITOS_VENCIDOS.length) { Logger.log('CREDITOS_VENCIDOS está vacío.'); return 0; }

  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const iso = v => v instanceof Date
    ? Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd')
    : String(v || '').trim().slice(0, 10);
  const norm = t => String(t || '').trim().toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '');

  Logger.log('═══ ' + (dryRun ? 'DRY-RUN · ' : '') + 'BORRAR CRÉDITOS VENCIDOS ═══');
  Logger.log('');
  const aBorrar = [];
  CREDITOS_VENCIDOS.forEach(c => {
    const hits = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (String(r[_R.ORIGEN] || '').trim() !== 'Abierta') continue;
      if (norm(r[1]) !== norm(c.quien)) continue;
      if (iso(r[_R.ENTRADA]) !== c.checkin) continue;
      if (Math.abs(_cleanMoney_(r[_R.MONTO]) - c.monto) > 0.01) continue;
      hits.push({ fila: i + 1, quien: r[1], cabana: r[2], ci: iso(r[_R.ENTRADA]) });
    }
    if (hits.length === 1) {
      aBorrar.push(hits[0]);
      Logger.log('   ✓ ' + c.quien + ' · ' + c.checkin + ' · $' + c.monto.toFixed(2)
        + '  → fila ' + hits[0].fila + ' (' + hits[0].cabana + ')');
    } else {
      Logger.log('   ✗ ' + c.quien + ' · ' + c.checkin + ' · $' + c.monto.toFixed(2)
        + '  → ' + hits.length + ' coincidencia(s), no se toca.');
      if (hits.length > 1) Logger.log('      filas: ' + hits.map(h => h.fila).join(', '));
      else Logger.log('      ¿ya se borró, o le cambiaron el origen / la fecha / el monto?');
    }
  });

  if (!aBorrar.length) {
    Logger.log('');
    Logger.log('Nada para borrar.');
    return 0;
  }
  if (dryRun) {
    Logger.log('');
    Logger.log('Nada se escribió. Para ejecutar: borrarCreditosVencidosESCRIBIR()');
    return 0;
  }
  // De mayor a menor para no correr los índices durante el borrado.
  aBorrar.sort((a, b) => b.fila - a.fila).forEach(x => sheet.deleteRow(x.fila));
  SpreadsheetApp.flush();
  Logger.log('');
  Logger.log('✓ ' + aBorrar.length + ' crédito(s) vencido(s) borrado(s).');
  return aBorrar.length;
}

function borrarCreditosVencidosESCRIBIR() { return borrarCreditosVencidos(false); }


// ═══════════════════════════════════════════════════════════════
//  RESCATAR UNA RESERVA QUE NUNCA LLEGÓ A LA HOJA
// ═══════════════════════════════════════════════════════════════
//
// Cuando el POST de `saveReservation` se pierde, la reserva queda viva solo en
// el `localStorage` del navegador del admin: se ve en el dashboard, no está en
// la hoja, y la cabaña se sigue publicando como libre. (Caso testigo: 12-ago-2026,
// un pasatarde en Paseo; `DebugLog` tenía los otros cinco `saveReservation` de
// esos días con su fila y el de esta no.)
//
// Recrearla desde el dashboard NO sirve si al huésped ya se le mandó el email o
// el WhatsApp de confirmación: el link público es `reserva.html?id=<ID>&t=<firma>`
// y el ID es un timestamp que se genera de nuevo en cada alta. Con un ID nuevo,
// el link que el huésped ya tiene queda apuntando a una reserva inexistente.
//
// Esta función inserta la fila CON EL ID ORIGINAL, así el link que ya circula
// sigue resolviendo. Idempotente: si el ID ya está, no hace nada.
//
// Cómo se usa:
//   1. Completar RESCATE_RESERVA con los datos del dashboard.
//   2. Correr `rescatarReservaReporte()` — no escribe, muestra qué haría.
//   3. Si se ve bien, `rescatarReservaESCRIBIR()`.

const RESCATE_RESERVA = {
  id:          '1786572168213',        // El del link que ya tiene el huésped. NO inventar uno nuevo.
  name:        'Víctor Bazan y Patty Recuero',
  cabin:       'verde',                // verde | azul | lila
  tipo:        'pasatarde',
  // FECHAS DE STORAGE, no las del formulario. Para pasatarde es día → día+1
  // (el form muestra el mismo día en entrada y salida, pero se guarda el rango).
  checkin:     '2026-08-13',
  checkout:    '2026-08-14',
  persons:     2,
  amount:      80,
  deposit:     80,
  origin:      'Directa',
  estadoPago:  'PAGA',
  codTransferencia: 'ZHWVH-37194891',
  montoVoucher:     80,
  bookingDate: '2026-08-12',
  comentarios: 'En colaboración con Malaya.\nClientes pasaran la mañana en Malaya. 50.00 Las Nubes / 30.00 Malaya',
  email:       '',                     // ← COMPLETAR: el mismo al que se envió la confirmación
  telefono:    '',                     // ← COMPLETAR
  horaEntrada: '',                     // vacío = default del tipo (12:30pm para pasatarde)
  horaSalida:  ''                      // vacío = default del tipo (7:00pm)
};

function rescatarReserva(dryRun) {
  if (dryRun === undefined) dryRun = true;   // el default es el seguro
  const r = RESCATE_RESERVA;
  const CABIN_NAMES = {
    verde: 'Paseo por Las Nubes',
    azul:  'Portal hacia Las Nubes',
    lila:  'Puente entre Las Nubes'
  };

  if (!r.id)                       { Logger.log('✗ Falta el id.'); return 0; }
  if (!CABIN_NAMES[r.cabin])       { Logger.log('✗ Cabaña inválida: ' + r.cabin); return 0; }
  if (!r.checkin || !r.checkout)   { Logger.log('✗ Faltan fechas de storage.'); return 0; }
  if (r.checkout <= r.checkin)     { Logger.log('✗ La salida debe ser posterior a la entrada. ¿Usaste las fechas del formulario en vez de las de storage?'); return 0; }

  const sheet = getOrCreateSheet();          // asegura las 33 columnas
  const data  = sheet.getDataRange().getValues();

  // Idempotencia: si el ID ya está, no se duplica.
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(r.id).trim()) {
      Logger.log('✓ La reserva ya está en la hoja (fila ' + (i + 1) + '). No se hace nada.');
      return 0;
    }
  }

  // Aviso —no bloqueo— si esas noches ya están ocupadas en esa cabaña: puede ser
  // legítimo (un pasatarde convive con una pasanoche) pero conviene mirarlo.
  const choques = [];
  for (let i = 1; i < data.length; i++) {
    const f = data[i];
    if (String(f[3]) !== r.cabin) continue;
    if (String(f[20] || '').toUpperCase() === 'CANCELADA') continue;
    const ci = _rrFecha(f[4]), co = _rrFecha(f[5]);
    if (!ci || !co) continue;
    if (ci < r.checkout && r.checkin < co) {
      choques.push('fila ' + (i + 1) + ': ' + f[1] + ' · ' + ci + ' → ' + co + ' · ' + (f[24] || 'noche'));
    }
  }

  Logger.log('── Rescate de reserva ' + r.id + ' ──');
  Logger.log('   ' + r.name + ' · ' + CABIN_NAMES[r.cabin] + ' · ' + (r.tipo || 'noche'));
  Logger.log('   storage ' + r.checkin + ' → ' + r.checkout + ' · ' + r.persons + ' pers · $' + r.amount);
  Logger.log('   link que se preserva: ' + getPublicReservaUrl(r.id));
  if (!r.email)    Logger.log('   ⚠ sin email — completá RESCATE_RESERVA.email si querés que reciba recordatorios');
  if (!r.telefono) Logger.log('   ⚠ sin teléfono');
  if (choques.length) {
    Logger.log('   ⚠ solapa con:');
    choques.forEach(c => Logger.log('      ' + c));
  }

  if (dryRun) {
    Logger.log('');
    Logger.log('Nada se escribió. Para ejecutar: rescatarReservaESCRIBIR()');
    return 0;
  }

  const fila = [
    r.id, r.name, CABIN_NAMES[r.cabin], r.cabin,
    r.checkin, r.checkout, r.persons || 1,
    r.amount || 0, r.deposit || 0, r.origin || 'Directa',
    r.id,                                   // CodConfirmacion: mismo id, como hace saveReservation
    0, r.amount || 0, '',                   // Service Fee, Neto, Alerta
    r.name,                                 // Pagador
    r.bookingDate || Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd'),
    '', 0,                                  // FechaPago, MontoPagado
    r.codTransferencia || '', r.montoVoucher || '', r.estadoPago || '',
    r.email || '', r.comentarios || '', r.telefono || '',
    r.tipo || 'noche'
  ];
  sheet.appendRow(fila);
  const nuevaFila = sheet.getLastRow();

  // Horas custom (cols 30/31) solo si se pidieron: vacías = default del tipo.
  if (r.horaEntrada) sheet.getRange(nuevaFila, 30).setValue(r.horaEntrada);
  if (r.horaSalida)  sheet.getRange(nuevaFila, 31).setValue(r.horaSalida);

  _invalidarCacheReservasPublicas();   // que el calendario público la vea ya
  logDebugEntry('rescatarReserva', { id: r.id, row: nuevaFila, name: r.name, cabin: r.cabin, checkin: r.checkin });
  SpreadsheetApp.flush();

  Logger.log('');
  Logger.log('✓ Insertada en la fila ' + nuevaFila + '.');
  Logger.log('  El link del huésped sigue funcionando: ' + getPublicReservaUrl(r.id));
  Logger.log('  Verificá el calendario público en ~1 min.');
  return 1;
}

function _rrFecha(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}

function rescatarReservaReporte()  { return rescatarReserva(true);  }
function rescatarReservaESCRIBIR() { return rescatarReserva(false); }
