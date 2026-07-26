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
