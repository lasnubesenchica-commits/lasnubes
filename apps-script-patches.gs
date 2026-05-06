// ═══════════════════════════════════════════════════════════
//  PATCHES PARA Apps Script — Las Nubes
//  Aplicar en el editor: reemplaza las funciones existentes
//  con las versiones de abajo. Las "INLINE PATCHES" son
//  cambios pequeños dentro de funciones que se mantienen
//  casi iguales.
//
//  Después de pegar todo: ejecutar UNA VEZ migrarColumnasV3()
//  para añadir el header "Telefono" a la hoja Reservas.
// ═══════════════════════════════════════════════════════════


// ─── REEMPLAZAR ENTERA: getOrCreateSheet ────────────────────
//  (Cambia: schema pasa de 23 a 24 columnas — agrega 'Telefono')
function getOrCreateSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 24).setValues([[
      'ID', 'Nombre', 'Cabaña', 'CabañaCodigo',
      'Entrada', 'Salida', 'Personas',
      'Monto', 'Abono', 'Origen', 'CodConfirmacion',
      'ServiceFee', 'Neto', 'Alerta', 'Pagador', 'FechaReserva',
      'FechaPago', 'MontoPagado', 'CodTransferencia', 'MontoVoucher', 'EstadoPago',
      'Email', 'Comentarios', 'Telefono'
    ]]);
    sheet.getRange(1, 1, 1, 24).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}


// ─── REEMPLAZAR ENTERA: appendReservation ───────────────────
//  (Cambia: agrega r.telefono como columna 24)
function appendReservation(sheet, r) {
  sheet.appendRow([
    r.id, r.name, r.cabinName, r.cabin,
    r.checkin, r.checkout, r.persons,
    r.amount, r.deposit, r.origin, r.confirmCode,
    r.serviceFee || 0,
    r.neto || r.amount,
    '',
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


// ─── AGREGAR NUEVA: migrarColumnasV3 ────────────────────────
//  Ejecutar una vez desde el editor para añadir el header
//  "Telefono" a hojas existentes.
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


// ─── REEMPLAZAR ENTERA: parseVoucherWithClaude ──────────────
//  (Cambia: prompt y respuesta JSON ahora extraen sender,
//  fechaPago, mensaje, nombreCompleto, email y telefono.
//  El monto se devuelve como número, código sin el "#".)
function parseVoucherWithClaude(imageBase64, mimeType) {
  try {
    const prompt = `Analiza este voucher de pago de Yappy y extrae los siguientes datos.

Datos visibles directamente en el voucher:
1. CÓDIGO DE CONFIRMACIÓN: campo "Confirmación" — devuelve SIN el "#" inicial (ej: "NBQEV-06253256").
2. MONTO: el monto principal numérico, sin "$" ni texto (ej: 90.00).
3. SENDER: el nombre del remitente que aparece arriba en "X te envió" (ej: "Ana G." o "Alejandra D.").
4. FECHA DEL PAGO: en formato YYYY-MM-DD. Asume zona horaria de Panamá. Infiere el año basándote en la fecha visible (año actual 2026 si no hay otro indicio).
5. MENSAJE: el texto literal del campo "Mensaje" si existe (puede no existir, en ese caso usa null).

Datos parseados del campo MENSAJE (si está presente):
6. NOMBRE COMPLETO: si el mensaje contiene un nombre completo del cliente (al menos nombre + apellido), devuélvelo. Si solo hay un nombre o no hay nombre, usa null.
7. EMAIL: si el mensaje contiene una dirección de correo, devuélvela. Si no, null.
8. TELÉFONO: si el mensaje contiene un número de teléfono panameño (8 dígitos, con o sin guión), devuélvelo en formato XXXX-XXXX. Si no, null.

Si algún campo opcional no está presente o no estás seguro, usa null. Nunca inventes datos.

Responde ÚNICAMENTE con JSON válido, sin markdown ni texto adicional:
{
  "codTransferencia": "NBQEV-06253256",
  "monto": 90.00,
  "sender": "Ana G.",
  "fechaPago": "2026-05-04",
  "mensaje": "abono",
  "nombreCompleto": null,
  "email": null,
  "telefono": null
}`;

    const payload = {
      model: 'claude-opus-4-6',
      max_tokens: 400,
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
    if (!result.content || !result.content[0]) {
      Logger.log('Voucher Claude API error: ' + response.getContentText());
      return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'Claude API error' })).setMimeType(ContentService.MimeType.JSON);
    }
    const text = result.content[0].text.trim();
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'No JSON' })).setMimeType(ContentService.MimeType.JSON);
    const data = JSON.parse(jsonMatch[0]);

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
      telefono:         data.telefono       || null
    })).setMimeType(ContentService.MimeType.JSON);
  } catch(e) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: e.message })).setMimeType(ContentService.MimeType.JSON);
  }
}


// ═══════════════════════════════════════════════════════════
//  INLINE PATCHES — cambios pequeños dentro de funciones existentes
// ═══════════════════════════════════════════════════════════


// ─── INLINE PATCH 1: doGet → mapeo de reservations ─────────
//  Dentro de doGet(), en el bloque que devuelve reservations
//  por defecto (justo antes del catch final). Buscar:
//
//      comentarios:      r[22] || ''
//      }));
//
//  y reemplazar por:
//
//      comentarios:      r[22] || '',
//      telefono:         r[23] || ''
//      }));


// ─── INLINE PATCH 2: doPost → action 'saveReservation' ─────
//  Dentro de doPost(), en el handler de action === 'saveReservation',
//  el bloque sheet.appendRow([...]) tiene 23 valores. Buscar la
//  última línea antes del cierre "]);":
//
//      r.comentarios  || ''
//      ]);
//
//  y reemplazar por:
//
//      r.comentarios  || '',
//      r.telefono     || ''
//      ]);


// ─── INLINE PATCH 3: doPost → action 'updateReservation' ───
//  Dentro de doPost(), en el handler de action === 'updateReservation',
//  el bloque tiene:
//
//      sheet.getRange(row, 1, 1, 23).setValues([[
//        ...
//        r.comentarios  || data[i][22] || ''
//      ]]);
//
//  Cambiar dos cosas:
//    1) (row, 1, 1, 23)  →  (row, 1, 1, 24)
//    2) Después de la línea de r.comentarios, agregar coma y línea:
//       r.telefono     || data[i][23] || ''
//
//  Resultado:
//
//      sheet.getRange(row, 1, 1, 24).setValues([[
//        ...
//        r.comentarios  || data[i][22] || '',
//        r.telefono     || data[i][23] || ''
//      ]]);


// ═══════════════════════════════════════════════════════════
//  CHECKLIST DE DESPLIEGUE
// ═══════════════════════════════════════════════════════════
//
//  1. Pegar getOrCreateSheet, appendReservation,
//     migrarColumnasV3 y parseVoucherWithClaude (reemplazos completos).
//  2. Aplicar los 3 INLINE PATCHES (doGet + doPost x2).
//  3. Guardar el script.
//  4. Ejecutar manualmente migrarColumnasV3() una vez.
//  5. Crear/Re-publicar la Web App si fue necesario.
//  6. Probar subiendo un voucher Yappy desde el dashboard;
//     verificar que se auto-llenan huésped, pagador, email,
//     teléfono y abono. Tip: probar con voucher cuyo "Mensaje"
//     contenga "Juan Pérez | juan@ejemplo.com | 6000-0000".
//
// ═══════════════════════════════════════════════════════════
