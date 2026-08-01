/**
 * TIENDITA DE LAS NUBES
 * =====================
 *
 * Insumos que se le venden al huésped (hielo, carbón, kit de fogata, repelente…).
 *
 * ─── De dónde sale cada número ─────────────────────────────────────────
 *
 * COSTO: no hay una hoja de compras aparte. Las compras ya están en `Egresos`,
 * y qué parte de cada factura es mercadería de reventa ya se sabe: son las
 * keywords de `SuministrosItems` marcadas con la columna `Reventa`. Ese
 * mecanismo ya reparte facturas compartidas (`MontosItem`), así que un Yappy
 * de gas + carbón puede tener $18 asignados a carbón sin inventar nada nuevo.
 * El frontend hace ese cruce, que es donde vive el matcher de keywords.
 *
 * VENTAS: sí necesitan hoja propia (`TienditaVentas`) — un cobro al huésped no
 * es un egreso ni una reserva, no tenía dónde vivir.
 *
 * GANANCIA = ventas − costo de las compras de reventa del período.
 *
 * Ojo con la comparación: en un período corto el costo y la venta no se
 * corresponden (comprás una caja de carbón en enero y la vendés hasta marzo).
 * El margen es confiable en ventanas largas; el frontend lo advierte.
 */

const TIENDITA_SHEET = 'TienditaVentas';
const TIENDITA_COLS = [
  'ID', 'Fecha', 'Item', 'Cantidad', 'PrecioUnitario', 'MontoTotal',
  'ReservaID', 'Huesped', 'VoucherURL', 'Notas'
];

function _tienditaSheet() {
  // Reusa el helper de Prestamos.gs: mismo patrón de hoja auto-migrada.
  return _hojaConCabecera(TIENDITA_SHEET, TIENDITA_COLS);
}

function getTienditaVentas() {
  const sheet = _tienditaSheet();
  const ventas = [];
  if (sheet.getLastRow() > 1) {
    sheet.getDataRange().getValues().slice(1).forEach(r => {
      if (!r[0]) return;
      ventas.push({
        id: String(r[0]),
        fecha: _pFecha(r[1]),
        item: String(r[2] || ''),
        cantidad: parseFloat(r[3]) || 1,
        precioUnitario: parseFloat(r[4]) || 0,
        monto: parseFloat(r[5]) || 0,
        reservaId: String(r[6] || ''),
        huesped: String(r[7] || ''),
        voucherURL: String(r[8] || ''),
        notas: String(r[9] || '')
      });
    });
  }
  ventas.sort((a, b) => (a.fecha < b.fecha ? 1 : -1));
  return { ventas: ventas };
}

function saveTienditaVenta(payload) {
  const sheet = _tienditaSheet();
  const id = String(payload.id || ('TV-' + Date.now()));
  const cantidad = parseFloat(payload.cantidad) || 1;
  const unit     = parseFloat(payload.precioUnitario) || 0;
  // El total se DERIVA de cantidad × precio unitario, salvo que venga explícito.
  // Guardar los tres campos sin relación es pedir que se contradigan.
  const total = (payload.monto !== undefined && payload.monto !== null && payload.monto !== '')
    ? (parseFloat(payload.monto) || 0)
    : +(cantidad * unit).toFixed(2);

  const fila = [
    id,
    payload.fecha || Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd'),
    String(payload.item || '').trim(),
    cantidad, unit, total,
    String(payload.reservaId || ''),
    String(payload.huesped || ''),
    String(payload.voucherURL || ''),
    String(payload.notas || '')
  ];
  const existente = _filaPorId(sheet, id);
  if (existente > 0) sheet.getRange(existente, 1, 1, fila.length).setValues([fila]);
  else sheet.appendRow(fila);
  logDebugEntry('tiendita-venta', { id: id, item: fila[2], monto: total });
  return { ok: true, id: id, monto: total };
}

function deleteTienditaVenta(ventaId) {
  const sheet = _tienditaSheet();
  const fila = _filaPorId(sheet, String(ventaId || ''));
  if (fila <= 0) return { ok: false, error: 'Venta no encontrada' };
  sheet.deleteRow(fila);
  return { ok: true };
}

// El voucher del cliente va a la misma carpeta de Drive que el resto de los
// comprobantes: son todos pagos, y separarlos solo obliga a recordar dónde
// buscar cada uno.
function saveTienditaVoucherToDrive(payload) {
  try {
    const carpeta = _carpetaPagos();
    const nombre = 'tiendita-' + (payload.ventaId || Date.now()) + '-'
                 + Utilities.formatDate(new Date(), 'America/Panama', 'yyyyMMdd-HHmmss')
                 + (String(payload.mimeType || '').indexOf('png') >= 0 ? '.png' : '.jpg');
    const blob = Utilities.newBlob(Utilities.base64Decode(payload.base64),
                                   payload.mimeType || 'image/jpeg', nombre);
    const file = carpeta.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Lee el voucher del cliente para prellenar monto y fecha. Es el mismo
    // parser de los vouchers de reserva, que ya entiende Yappy y ACH.
    let datos = null;
    try {
      const v = parseVoucherWithClaude(payload.base64, payload.mimeType || 'image/jpeg');
      if (v && !v.error) {
        datos = {
          monto: parseFloat(String(v.monto || '').replace(/[^0-9.]/g, '')) || 0,
          fecha: v.fechaPago || '',
          remitente: v.sender || '',
          // El campo "Mensaje" del Yappy suele traer qué se compró ("kit de
          // fogata", "2 hielos"). El frontend lo cruza con los items para
          // preseleccionar el correcto.
          mensaje: v.mensaje || ''
        };
      }
    } catch (_) { /* sin OCR: se llena a mano */ }

    return { ok: true, url: file.getUrl(), datos: datos };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function _testTiendita() {
  Logger.log(JSON.stringify(getTienditaVentas(), null, 2));
}
