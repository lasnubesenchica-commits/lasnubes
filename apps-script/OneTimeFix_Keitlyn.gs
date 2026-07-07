/**
 * ONE-TIME FIX: reconstruir hermana 1 de la reserva multi-cabaña de
 * KEITLYN RANGEL SANCHEZ que se perdió por el bug de fallo silencioso
 * en _saveMultiCabinReservation (arreglado en PR #341).
 *
 * Contexto: la reserva original era Puente (07/07) + Portal (08/07),
 * pago único de $150. Solo la hermana 2 (Portal) quedó en el Sheet.
 * Este script inserta la hermana 1 (Puente) con los mismos metadatos.
 *
 * USO: correr una única vez desde el editor de Apps Script. Después de
 * confirmar que la fila aparece, eliminar este archivo del repo.
 */

function fixKeitlynMultiCabinHermana1() {
  const sheet = getOrCreateSheet();
  const HERMANA1_ID = 'MC-1783376570702-1';

  // Guard: si ya existe (por si se ejecuta 2 veces), no duplicar.
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === HERMANA1_ID) {
      const msg = '✗ Ya existe MC-1783376570702-1 en la fila ' + (i + 1) + '. No se hace nada.';
      Logger.log(msg);
      return msg;
    }
  }

  // Datos de hermana 1 basados en hermana 2 (que sí se guardó en fila 442
  // del CSV) — mismo huésped, mismo voucher, mismo pago único. Solo
  // cambian cabaña, fechas y el marcador 1/2 en el comentario.
  const rowToAppend = [
    HERMANA1_ID,                                                 // 1  ID
    'KEITLYN RANGEL SANCHEZ',                                    // 2  Nombre
    'Puente entre Las Nubes',                                    // 3  Cabaña (display)
    'lila',                                                      // 4  CabañaCodigo
    '2026-07-07',                                                // 5  Entrada
    '2026-07-08',                                                // 6  Salida
    2,                                                           // 7  Personas
    75,                                                          // 8  Monto (mitad del total)
    150,                                                         // 9  Abono (deposit total del par — la hermana 1 en el flujo original recibía el deposit completo)
    'Directa',                                                   // 10 Origen
    HERMANA1_ID,                                                 // 11 CodConfirmacion
    0,                                                           // 12 Service Fee
    75,                                                          // 13 Neto
    '',                                                          // 14 Alerta
    'KEITLYN RANGEL SANCHEZ',                                    // 15 Pagador
    '2026-07-06',                                                // 16 FechaReserva
    '',                                                          // 17 FechaPago
    0,                                                           // 18 MontoPagado
    '',                                                          // 19 codTransferencia
    '$150.00',                                                   // 20 MontoVoucher (total del pago único)
    'PAGA',                                                      // 21 EstadoPago
    'Keitlyns1516@gmail.com',                                    // 22 Email
    'Clientes llegaran despues de las 7:30pm\n[Multi-cabaña 1/2 · Puente + Portal · pago único $150.00]', // 23 Comentarios
    '507 6385-3681',                                             // 24 Telefono
    'noche'                                                      // 25 Tipo
    // 26 VoucherURL — se deja vacía. Si querés vincular el voucher que
    //    tenés guardado en Drive, editarlo a mano en la hoja después.
  ];

  sheet.appendRow(rowToAppend);
  const newRow = sheet.getLastRow();
  Logger.log('✓ Hermana 1 insertada en fila ' + newRow + ' con ID ' + HERMANA1_ID);

  logDebugEntry('fixKeitlynMultiCabinHermana1-OK', {
    id: HERMANA1_ID,
    row: newRow,
    cabin: 'lila',
    checkin: '2026-07-07',
    checkout: '2026-07-08'
  });

  return '✓ OK — fila ' + newRow + ' agregada. Revisá el Sheet.';
}
