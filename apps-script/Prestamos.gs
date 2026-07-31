/**
 * PRÉSTAMOS A COLABORADORES
 * =========================
 *
 * Adelantos que se le hacen a un colaborador (una compra, una transferencia) y
 * que se recuperan descontando de lo que se le paga después.
 *
 * Dos formas de cobro, mismo modelo:
 *   - `por_pago`   : monto fijo cada vez que se le paga. Ej. el centro de
 *                    lavado: el pago semanal de limpieza es $175, se le
 *                    entregan $135, y esos $40 abonan al saldo.
 *   - `por_evento` : monto fijo por cada evento. Ej. las piezas de la moto:
 *                    cada delivery vale $4 y en vez de transferirlos se
 *                    acreditan al saldo.
 *
 * ─── Cómo entra en la contabilidad ─────────────────────────────────────
 *
 * El desembolso ya está registrado como egreso (categoría `Prestamo`), así que
 * la salida de plata está contada.
 *
 * Los abonos NO son ingresos: son plata que nunca salió. El egreso semanal de
 * limpieza se registra por la tarifa completa ($175) aunque solo se hayan
 * entregado $135, así que sin corregir nada el gasto queda inflado en $40 cada
 * semana. Por eso Contabilidad resta los abonos del período:
 *
 *     egresos netos = egresos brutos − abonos de préstamos
 *
 * Sumarlos a Ingresos habría sido el otro camino, pero eso inventa facturación
 * que no existe y ensucia toda la métrica de ventas.
 */

const PRESTAMOS_SHEET = 'Prestamos';
const PRESTAMO_ABONOS_SHEET = 'PrestamoAbonos';

const PRESTAMOS_COLS = [
  'ID', 'Fecha', 'Beneficiario', 'Concepto', 'MontoTotal',
  'TipoCobro', 'MontoPorCobro', 'Frecuencia', 'EgresoID',
  'URLFactura', 'Estado', 'Notas'
];
const PRESTAMO_ABONOS_COLS = ['ID', 'PrestamoID', 'Fecha', 'Monto', 'Nota', 'EgresoID'];

function _prestamosSheet() {
  return _hojaConCabecera(PRESTAMOS_SHEET, PRESTAMOS_COLS);
}
function _prestamoAbonosSheet() {
  return _hojaConCabecera(PRESTAMO_ABONOS_SHEET, PRESTAMO_ABONOS_COLS);
}

// Crea la hoja si falta y le asegura las columnas. Igual que getOrCreateSheet
// para Reservas: si mañana se agrega una columna, las hojas viejas la reciben
// sola en el próximo acceso en vez de romper por índice corrido.
function _hojaConCabecera(nombre, cols) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const ancho = Math.max(sheet.getLastColumn(), 1);
  const cab = sheet.getRange(1, 1, 1, ancho).getValues()[0].map(String);
  if (cab.length < cols.length) {
    sheet.getRange(1, cab.length + 1, 1, cols.length - cab.length)
         .setValues([cols.slice(cab.length)]);
  }
  return sheet;
}

function _pFecha(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'America/Panama', 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}

// ─── Lectura ────────────────────────────────────────────────────
//
// Devuelve los préstamos con su saldo ya calculado. El saldo NO se guarda en la
// hoja: se deriva de los abonos en cada lectura. Un total guardado y una lista
// de abonos son dos fuentes para el mismo número, y tarde o temprano se separan
// (basta con borrar un abono a mano y el saldo queda mintiendo).
function getPrestamosData() {
  const abonosPorPrestamo = {};
  const aSheet = _prestamoAbonosSheet();
  const abonos = [];
  if (aSheet.getLastRow() > 1) {
    aSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (!r[0]) return;
      const a = {
        id: String(r[0]), prestamoId: String(r[1] || ''), fecha: _pFecha(r[2]),
        monto: parseFloat(r[3]) || 0, nota: String(r[4] || ''), egresoId: String(r[5] || '')
      };
      abonos.push(a);
      abonosPorPrestamo[a.prestamoId] = (abonosPorPrestamo[a.prestamoId] || 0) + a.monto;
    });
  }

  const pSheet = _prestamosSheet();
  const prestamos = [];
  if (pSheet.getLastRow() > 1) {
    pSheet.getDataRange().getValues().slice(1).forEach(r => {
      if (!r[0]) return;
      const id = String(r[0]);
      const total = parseFloat(r[4]) || 0;
      const abonado = +(abonosPorPrestamo[id] || 0).toFixed(2);
      prestamos.push({
        id: id, fecha: _pFecha(r[1]), beneficiario: String(r[2] || ''),
        concepto: String(r[3] || ''), montoTotal: total,
        tipoCobro: String(r[5] || 'por_pago'), montoPorCobro: parseFloat(r[6]) || 0,
        frecuencia: String(r[7] || ''), egresoId: String(r[8] || ''),
        urlFactura: String(r[9] || ''), estado: String(r[10] || 'activo'),
        notas: String(r[11] || ''),
        abonado: abonado,
        saldo: +(total - abonado).toFixed(2),
        // Cuántos cobros más faltan al ritmo configurado. Sirve para decir
        // "faltan 4 pagos" en vez de solo un saldo suelto.
        cobrosRestantes: (parseFloat(r[6]) > 0)
          ? Math.ceil(Math.max(0, total - abonado) / parseFloat(r[6])) : null
      });
    });
  }
  prestamos.sort((a, b) => (a.fecha < b.fecha ? 1 : -1));
  return { prestamos: prestamos, abonos: abonos };
}

// Total abonado dentro de un rango de fechas. Lo usa Contabilidad para netear
// contra los egresos del período. `desde`/`hasta` en ISO, ambos inclusive;
// vacíos = sin límite.
function getAbonosPrestamoEnRango(desde, hasta) {
  const sheet = _prestamoAbonosSheet();
  if (sheet.getLastRow() < 2) return 0;
  let total = 0;
  sheet.getDataRange().getValues().slice(1).forEach(r => {
    if (!r[0]) return;
    const f = _pFecha(r[2]);
    if (desde && f < desde) return;
    if (hasta && f > hasta) return;
    total += parseFloat(r[3]) || 0;
  });
  return +total.toFixed(2);
}

// ─── Escritura ──────────────────────────────────────────────────

function savePrestamo(payload) {
  const sheet = _prestamosSheet();
  const id = String(payload.id || ('PR-' + Date.now()));
  const fila = [
    id,
    payload.fecha || Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd'),
    String(payload.beneficiario || '').trim(),
    String(payload.concepto || '').trim(),
    parseFloat(payload.montoTotal) || 0,
    String(payload.tipoCobro || 'por_pago'),
    parseFloat(payload.montoPorCobro) || 0,
    String(payload.frecuencia || ''),
    String(payload.egresoId || ''),
    String(payload.urlFactura || ''),
    String(payload.estado || 'activo'),
    String(payload.notas || '')
  ];

  const existente = _filaPorId(sheet, id);
  if (existente > 0) {
    sheet.getRange(existente, 1, 1, fila.length).setValues([fila]);
  } else {
    sheet.appendRow(fila);
  }
  logDebugEntry('prestamo-guardado', { id: id, benef: fila[2], monto: fila[4] });
  return { ok: true, id: id };
}

function savePrestamoAbono(payload) {
  const prestamoId = String(payload.prestamoId || '');
  if (!prestamoId) return { ok: false, error: 'Falta prestamoId' };
  const monto = parseFloat(payload.monto) || 0;
  if (monto <= 0) return { ok: false, error: 'El abono debe ser mayor a 0' };

  const sheet = _prestamoAbonosSheet();
  const id = String(payload.id || ('AB-' + Date.now()));
  const fila = [
    id, prestamoId,
    payload.fecha || Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd'),
    monto, String(payload.nota || ''), String(payload.egresoId || '')
  ];
  const existente = _filaPorId(sheet, id);
  if (existente > 0) sheet.getRange(existente, 1, 1, fila.length).setValues([fila]);
  else sheet.appendRow(fila);

  // Si el abono terminó de cubrir el total, el préstamo se marca saldado solo.
  // Se recalcula desde los abonos, no desde un contador aparte.
  const data = getPrestamosData();
  const p = data.prestamos.find(x => x.id === prestamoId);
  if (p && p.saldo <= 0.009 && p.estado !== 'saldado') {
    const pSheet = _prestamosSheet();
    const fila2 = _filaPorId(pSheet, prestamoId);
    if (fila2 > 0) pSheet.getRange(fila2, 11).setValue('saldado');
  }
  return { ok: true, id: id, saldo: p ? p.saldo : null, saldado: !!(p && p.saldo <= 0.009) };
}

function deletePrestamoAbono(abonoId) {
  const sheet = _prestamoAbonosSheet();
  const fila = _filaPorId(sheet, String(abonoId || ''));
  if (fila <= 0) return { ok: false, error: 'Abono no encontrado' };
  sheet.deleteRow(fila);
  return { ok: true };
}

// Borra el préstamo y TODOS sus abonos. Dejar abonos huérfanos haría que
// getAbonosPrestamoEnRango siguiera neteando contra los egresos plata de un
// préstamo que ya no existe.
function deletePrestamo(prestamoId) {
  const id = String(prestamoId || '');
  const aSheet = _prestamoAbonosSheet();
  if (aSheet.getLastRow() > 1) {
    const vals = aSheet.getDataRange().getValues();
    for (let i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][1] || '') === id) aSheet.deleteRow(i + 1);
    }
  }
  const pSheet = _prestamosSheet();
  const fila = _filaPorId(pSheet, id);
  if (fila > 0) pSheet.deleteRow(fila);
  return { ok: true };
}

function _filaPorId(sheet, id) {
  if (!id || sheet.getLastRow() < 2) return -1;
  const vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]) === id) return i + 2;
  }
  return -1;
}

// ─── Factura del préstamo a Drive ───────────────────────────────
// Misma carpeta que los vouchers de pago: son comprobantes del mismo tipo y
// tenerlos separados solo obliga a recordar en cuál buscar.
// Además de guardar el archivo, LEE la factura con Claude y devuelve los datos
// para prellenar el formulario. Reusa parseFacturaEgresoConClaude, el mismo
// parser de las facturas de egresos: entiende facturas comerciales, Yappy y
// transferencias ACH, así que no hace falta un prompt aparte.
//
// Si el OCR falla, el archivo igual queda subido: el adjunto es lo importante y
// los campos siempre se pueden escribir a mano.
function savePrestamoFacturaToDrive(payload) {
  try {
    const carpeta = _carpetaPagos();
    const bytes = Utilities.base64Decode(payload.base64);
    const nombre = 'prestamo-' + (payload.prestamoId || Date.now()) + '-'
                 + Utilities.formatDate(new Date(), 'America/Panama', 'yyyyMMdd-HHmmss')
                 + (String(payload.mimeType || '').indexOf('png') >= 0 ? '.png' : '.jpg');
    const blob = Utilities.newBlob(bytes, payload.mimeType || 'image/jpeg', nombre);
    const file = carpeta.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    let datos = null;
    try {
      const p = parseFacturaEgresoConClaude(payload.base64, payload.mimeType || 'image/jpeg');
      if (p && !p.error) {
        // El concepto sale del detalle de la factura, que es lo que describe QUÉ
        // se compró; el proveedor solo dice a quién se le pagó.
        const primerItem = (p.items && p.items.length) ? p.items[0] : null;
        datos = {
          monto:     parseFloat(p.monto) || 0,
          fecha:     p.fecha || '',
          proveedor: p.proveedor || '',
          concepto:  (primerItem && primerItem.desc) || p.proveedor || '',
          numFactura: p.numFactura || ''
        };
      }
    } catch (_) { /* sin OCR: se llena a mano */ }

    return { ok: true, url: file.getUrl(), datos: datos };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function _carpetaPagos() {
  const nombre = 'Las Nubes - Pagos';
  const it = DriveApp.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : DriveApp.createFolder(nombre);
}

// Test desde el editor.
function _testPrestamos() {
  Logger.log(JSON.stringify(getPrestamosData(), null, 2));
}
