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

// ─── FUNCIÓN PRINCIPAL ──────────────────────────────────────
function syncAirbnbReservations() {
  const sheet = getOrCreateSheet();
  const processed = getProcessedIds(sheet);

  const threads = GmailApp.search(
    'from:automated@airbnb.com subject:"Reserva confirmada:" newer_than:90d'
  );

  let added = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const msgId = msg.getId();
      if (processed.has(msgId)) return;

      const body = msg.getPlainBody();
      const msgDate = Utilities.formatDate(msg.getDate(), 'America/Panama', 'yyyy-MM-dd');
      const reservation = parseAirbnbEmail(body, msgId, msgDate);

      if (reservation) {
        const blacklisted = getBlacklistedCodes();
        const code = reservation.confirmCode || '';
        if (blacklisted.has(code.toString().trim())) {
          Logger.log(`⛔ Blacklisted, ignorando: ${reservation.name} (${code})`);
        } else {
          appendReservation(sheet, reservation);
          added++;
          Logger.log(`✓ Reserva agregada: ${reservation.name} - ${reservation.cabin}`);
        }
      }
    });
  });

  Logger.log(`Sync completado. ${added} reservas nuevas agregadas.`);
}

// ─── SYNC RESERVAS ACTUALIZADAS ────────────────────────────
function syncAirbnbUpdates() {
  const sheet = getOrCreateSheet();

  const threads = GmailApp.search(
    'from:automated@airbnb.com subject:"Reserva actualizada" newer_than:90d'
  );

  const ss            = SpreadsheetApp.openById(SHEET_ID);
  let updSheet        = ss.getSheetByName('UpdatesProcessed');
  if (!updSheet) {
    updSheet = ss.insertSheet('UpdatesProcessed');
    updSheet.getRange(1,1).setValue('MsgId');
  }
  const updData       = updSheet.getDataRange().getValues();
  const processedUpds = new Set(updData.slice(1).map(r => r[0].toString()));

  let flagged = 0;

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const msgId = msg.getId();
      if (processedUpds.has(msgId)) return;

      const body = msg.getPlainBody();

      const codeMatch = body.match(/reservations\/details\/([A-Z0-9]{8,12})\?/);
      if (!codeMatch) {
        Logger.log('⚠ Update sin código identificable: ' + msgId);
        updSheet.appendRow([msgId]);
        return;
      }
      const confirmCode = codeMatch[1];

      const data = sheet.getDataRange().getValues();
      let foundRow = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][10] && data[i][10].toString() === confirmCode) {
          foundRow = i + 1;
          break;
        }
      }

      if (foundRow > 0) {
        const numCols = sheet.getLastColumn();
        if (numCols < 14) {
          sheet.getRange(1, 14).setValue('Alerta');
          sheet.getRange(1, 14).setFontWeight('bold');
        }
        const alertCell = sheet.getRange(foundRow, 14);
        alertCell.setValue('⚠ Verificar');
        alertCell.setBackground('#FFF3CD');
        alertCell.setFontColor('#856404');
        sheet.getRange(foundRow, 1, 1, 13).setBackground('#FFFDE7');

        Logger.log('⚠ Marcada para verificación: ' + confirmCode + ' (fila ' + foundRow + ')');
        flagged++;
      } else {
        Logger.log('⚠ Código ' + confirmCode + ' no encontrado en el Sheet');
      }

      updSheet.appendRow([msgId]);
    });
  });

  Logger.log('Updates procesados. ' + flagged + ' reservas marcadas para verificación.');
}

// ─── PARSER DEL EMAIL ───────────────────────────────────────
function parseAirbnbEmail(body, msgId, msgDate) {
  try {
    const nameMatch = body.match(/([A-ZÁÉÍÓÚÜÑ][a-záéíóúüñ]+(?:\s+[A-ZÁÉÍÓÚÜÑ][a-záéíóúüñ]+){1,3})\s+\[https:\/\/www\.airbnb/);
    const name = nameMatch ? nameMatch[1].trim() : extractFallbackName(body);

    let cabin = 'verde';
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
      checkin  = parseFecha(fechasMatch[1]);
      checkout = parseFecha(fechasMatch[2]);
    }

    if (!checkin || !checkout) {
      const llegadaSalidaBlock = body.match(/Llegada[\s\S]{0,80}?Salida[\s\S]{0,200}?(\w+,?\s+\d{1,2}\s+\w+)[\s\S]{0,80}?(\w+,?\s+\d{1,2}\s+\w+)/i);
      if (llegadaSalidaBlock) {
        const d1 = parseFecha(llegadaSalidaBlock[1]);
        const d2 = parseFecha(llegadaSalidaBlock[2]);
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
        checkin  = parseFecha(uniqueDates[0]);
        checkout = parseFecha(uniqueDates[1]);
        if (checkin && checkout && checkin === checkout) {
          const d = new Date(checkout + 'T12:00:00');
          d.setDate(d.getDate() + (nochesEsperadas || 1));
          checkout = d.toISOString().slice(0, 10);
        }
      } else if (uniqueDates.length === 1) {
        checkin = parseFecha(uniqueDates[0]);
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
        checkin = parseFecha(subjectMatch[1]);
        if (checkin) {
          const d = new Date(checkin + 'T12:00:00');
          d.setDate(d.getDate() + 1);
          checkout = d.toISOString().slice(0, 10);
        }
      }
    }

    const personsMatch = body.match(/(\d+)\s+adulto/i);
    const persons = personsMatch ? parseInt(personsMatch[1]) : 2;

    const totalMatch = body.match(/Total\s*\(USD\)[^\$]*\$\s*([\d,\.]+)/i);
    const amount = totalMatch
      ? parseFloat(totalMatch[1].replace(',', '.'))
      : 0;

    const codeMatch = body.match(/\/details\/([A-Z0-9]{10})/);
    const confirmCode = codeMatch ? codeMatch[1] : '';

    Logger.log(`  Nombre: ${name}`);
    Logger.log(`  Cabaña: ${cabinName} (${cabin})`);
    Logger.log(`  Entrada: ${checkin} | Salida: ${checkout}`);
    Logger.log(`  Personas: ${persons} | Monto: $${amount} | Código: ${confirmCode}`);

    if (!name || !checkin || !checkout) {
      Logger.log('⚠ No se pudieron extraer datos completos del email');
      Logger.log(`  → name="${name}" checkin="${checkin}" checkout="${checkout}"`);
      return null;
    }

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
      bookingDate: msgDate || Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd'),
      addedAt: new Date().toISOString()
    };

  } catch(e) {
    Logger.log('Error parseando email: ' + e.message);
    return null;
  }
}

// ─── EXTRACTOR DE FECHAS ────────────────────────────────────
function parseFecha(str) {
  if (!str) return null;
  const match = str.match(/(\d{1,2})\s+(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)/i);
  if (!match) return null;
  const day   = match[1].padStart(2, '0');
  const month = monthToNum(match[2]);
  const year  = getCurrentYear(parseInt(month));
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
    sheet.getRange(1, 1, 1, 26).setValues([[
      'ID', 'Nombre', 'Cabaña', 'CabañaCodigo',
      'Entrada', 'Salida', 'Personas',
      'Monto', 'Abono', 'Origen', 'CodConfirmacion',
      'ServiceFee', 'Neto', 'Alerta', 'Pagador', 'FechaReserva',
      'FechaPago', 'MontoPagado', 'CodTransferencia', 'MontoVoucher', 'EstadoPago',
      'Email', 'Comentarios', 'Telefono', 'Tipo', 'VoucherURL'
    ]]);
    sheet.getRange(1, 1, 1, 26).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    // Auto-asegurar columnas Tipo (25) y VoucherURL (26)
    if (sheet.getLastColumn() < 26) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (!headers.includes('Tipo')) {
        sheet.getRange(1, 25).setValue('Tipo');
        sheet.getRange(1, 25).setFontWeight('bold');
      }
      if (!headers.includes('VoucherURL')) {
        sheet.getRange(1, 26).setValue('VoucherURL');
        sheet.getRange(1, 26).setFontWeight('bold');
      }
    }
  }

  return sheet;
}

// ─── CONFIG (Tarifas) ────────────────────────────────────────
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
    cfg.appendRow(['weekend',     110,     'Tarifa Vie-Sab',             hoy, 'sistema']);
    cfg.appendRow(['promo',       75,      'Tarifa promocional Dom-Jue', hoy, 'sistema']);
    cfg.appendRow(['promoActive', 'false', 'Promo activa (true/false)',  hoy, 'sistema']);
    Logger.log('✓ Hoja Config creada con valores por defecto');
  }
  return cfg;
}

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
    'from:automated@airbnb.com subject:"Te hemos enviado un cobro" newer_than:365d'
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
          const m = lines[j].match(/\$([\d.]+,\d{2})\s*USD/);
          if (m) {
            monto = parseFloat(m[1].replace(/\./g, '').replace(',', '.')) || 0;
            if (monto > 0) break;
          }
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

// ═══════════════════════════════════════════════════════════
//  doGet — endpoint principal
// ═══════════════════════════════════════════════════════════
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  try {
    // ── SAVE TARIFAS (via GET params para evitar redirect 302) ──
    if (action === 'saveTarifas') {
      const p   = e.parameter;
      const cfg = getOrCreateConfig();
      const rows = cfg.getDataRange().getValues();
      const hoy  = Utilities.formatDate(new Date(), 'America/Panama', 'yyyy-MM-dd');
      const updates = {
        weekday:     parseFloat(p.weekday)  || 90,
        weekend:     parseFloat(p.weekend)  || 110,
        promo:       parseFloat(p.promo)    || 75,
        promoActive: (p.promoActive === 'true' || p.promoActive === '1') ? 'true' : 'false'
      };
      for (let i = 1; i < rows.length; i++) {
        const clave = rows[i][0] ? rows[i][0].toString().trim() : '';
        if (updates.hasOwnProperty(clave)) {
          cfg.getRange(i + 1, 2).setValue(updates[clave]);
          cfg.getRange(i + 1, 4).setValue(hoy);
          cfg.getRange(i + 1, 5).setValue('admin');
        }
      }
      Logger.log('✓ Tarifas guardadas via GET: ' + JSON.stringify(updates));
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
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
      const tarifas = {
        weekday:     parseFloat(map['weekday'])    || 90,
        weekend:     parseFloat(map['weekend'])    || 110,
        promo:       parseFloat(map['promo'])      || 75,
        promoActive: map['promoActive'] === true || map['promoActive'] === 'true' || map['promoActive'] === 1
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
        urlFoto:   headers.findIndex(h => h.includes('url') || h.includes('foto'))
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
            fromSheets: true
          };
        });
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, egresos }))
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
    if (action === 'debugShareLinks')  return handleDebugShareLinks(e);

    // ── GET RESERVATIONS (default) ────────────────────────────
    const sheet = getOrCreateSheet();
    const data  = sheet.getDataRange().getValues();
    const rows  = data.slice(1);

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
        voucherURL:       r[25] || ''
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
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;

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
        promoActive: t.promoActive ? 'true' : 'false'
      };

      for (let i = 1; i < rows.length; i++) {
        const clave = rows[i][0] ? rows[i][0].toString().trim() : '';
        if (updates.hasOwnProperty(clave)) {
          cfg.getRange(i + 1, 2).setValue(updates[clave]);
          cfg.getRange(i + 1, 4).setValue(hoy);
          cfg.getRange(i + 1, 5).setValue('admin');
        }
      }
      Logger.log('✓ Tarifas guardadas: ' + JSON.stringify(updates));
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── SAVE RESERVATION ─────────────────────────────────────
    if (action === 'saveReservation') {
      const r     = payload.reservation;
      const sheet = getOrCreateSheet();

      const existing = getProcessedIds(sheet);
      const key = r.confirmCode || r.id;
      if (key && existing.has(key.toString())) {
        return ContentService
          .createTextOutput(JSON.stringify({ ok: true, status: 'duplicate' }))
          .setMimeType(ContentService.MimeType.JSON);
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

      sheet.appendRow([
        r.id,
        r.name,
        CABIN_NAMES[r.cabin] || r.cabin,
        r.cabin,
        r.checkin,
        r.checkout,
        persons_,
        amount_,
        deposit_,
        r.origin    || 'Directa',
        r.confirmCode || '',
        serviceFee_,
        neto_,
        '',
        r.pagador || r.name,
        r.fechaReserva || today_,
        '',
        0,
        r.codTransferencia || '',
        r.montoVoucher || '',
        r.estadoPago   || '',
        r.email        || '',
        r.comentarios  || '',
        r.telefono     || '',
        r.tipo         || 'noche'
      ]);

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
            r.name,
            CABIN_NAMES[r.cabin] || r.cabin,
            r.cabin,
            r.checkin,
            r.checkout,
            persons_,
            amount_,
            deposit_,
            r.origin    || 'Directa',
            r.confirmCode || r.id,
            serviceFee_,
            neto_,
            data[i][13] || '',
            r.pagador   || r.name,
            r.fechaReserva || today_,
            data[i][16] || '',
            data[i][17] || 0,
            r.codTransferencia || data[i][18] || '',
            r.montoVoucher || data[i][19] || '',
            (function() {
              const existingEstado = data[i][20] ? data[i][20].toString().trim() : '';
              const newEstado      = r.estadoPago ? r.estadoPago.toString().trim() : '';
              if (existingEstado === 'CANCELADA') return 'CANCELADA';
              if (existingEstado === 'PAGA' && newEstado === 'PENDIENTE') return 'PAGA';
              return newEstado || existingEstado || '';
            })(),
            r.email        || data[i][21] || '',
            r.comentarios  || data[i][22] || '',
            r.telefono     || data[i][23] || '',
            r.tipo         || data[i][24] || 'noche'
          ]]);

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

    // ── SAVE EGRESOS (multi-item) ─────────────────────────────
    if (action === 'saveEgresos') {
      const items = payload.egresos;
      const ss = SpreadsheetApp.openById(SHEET_ID);
      let egresoSheet = ss.getSheetByName('Egresos');
      if (!egresoSheet) {
        egresoSheet = ss.insertSheet('Egresos');
        egresoSheet.getRange(1,1,1,8).setValues([['ID','Fecha','Descripcion','Monto','Categoria','Cabaña','Proveedor','URLFoto']]);
        egresoSheet.getRange(1,1,1,8).setFontWeight('bold');
      } else {
        const firstRow = egresoSheet.getRange(1,1,1,8).getValues()[0];
        if (!firstRow[0] || firstRow[0].toString().toLowerCase() !== 'id') {
          egresoSheet.getRange(1,1,1,8).setValues([['ID','Fecha','Descripcion','Monto','Categoria','Cabaña','Proveedor','URLFoto']]);
          egresoSheet.getRange(1,1,1,8).setFontWeight('bold');
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
          eg.urlFoto   || ''
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
      let egresoSheet = ss.getSheetByName('Egresos');
      if (!egresoSheet) {
        egresoSheet = ss.insertSheet('Egresos');
        egresoSheet.getRange(1,1,1,8).setValues([['ID','Fecha','Descripcion','Monto','Categoria','Cabaña','Proveedor','URLFoto']]);
        egresoSheet.getRange(1,1,1,8).setFontWeight('bold');
      }
      egresoSheet.appendRow([
        eg.id        || '',
        eg.date      || '',
        eg.desc      || '',
        parseFloat(eg.amount) || 0,
        eg.cat       || 'Otro',
        eg.cabin     || '',
        eg.proveedor || '',
        eg.urlFoto   || ''
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
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString() === eg.id) {
          if (eg.cat       !== undefined) sheet.getRange(i+1, 5).setValue(eg.cat);
          if (eg.desc      !== undefined) sheet.getRange(i+1, 3).setValue(eg.desc);
          if (eg.proveedor !== undefined) sheet.getRange(i+1, 7).setValue(eg.proveedor);
          if (eg.cabin     !== undefined) sheet.getRange(i+1, 6).setValue(eg.cabin);
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
      return saveVoucherToDrive(payload.reservation, payload.imageBase64, payload.mimeType, payload.fileName);
    }

    // ── SEND EMAILS ───────────────────────────────────────────
    if (action === 'sendCancellationEmail') return sendCancellationEmail(payload.reservation);
    if (action === 'sendConfirmationEmail') return sendConfirmationEmail(payload.reservation, payload.voucherBase64, payload.voucherMimeType);
    if (action === 'sendUpdateEmail')       return sendUpdateEmail(payload.reservation, payload.voucherBase64, payload.voucherMimeType);
    if (action === 'sendCheckinReminder')   return sendCheckinReminderEmail(payload.reservation);

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
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
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

function buildGuiaHTML(cabin, tipo) {
  // Fuente unica: getCabinGuideSteps() en PublicLink.gs
  var list = getCabinGuideSteps(cabin, tipo);
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

function buildEmailHTML(r) {
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
  let publicLink    = '';
  try { if (r.id) publicLink = getPublicReservaUrl(r.id); } catch(e) { Logger.log('no publicLink: ' + e.message); }

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
buildGuiaHTML(r.cabin, meta.tipo) +
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
'<a href="http://lasnubes.cloud/#insumos" target="_blank" style="display:inline-block;background:#5a85b0;color:#ffffff;font-size:13px;font-weight:600;padding:10px 20px;border-radius:8px;text-decoration:none;margin-bottom:24px;">&#128722; Ver guía de insumos</a>' +
'<hr style="border:none;border-top:1px solid #e8e4de;margin:0 0 24px;">' +
'<h2 style="margin:0 0 6px;font-size:17px;font-weight:600;color:#3a3530;">&#127956;&#65039; Actividades y alrededores</h2>' +
'<p style="margin:0 0 16px;font-size:13px;color:#6b6560;line-height:1.6;">Cascada dentro del proyecto (20 min), Los Cajones de Chame (10 min), Cerro Campana (20 min), Playa Gorgona y Coronado (15–20 min) y cascadas de Sorá (20 min).</p>' +
'<a href="http://lasnubes.cloud/#actividades" target="_blank" style="display:inline-block;background:#4a7c5f;color:#ffffff;font-size:13px;font-weight:600;padding:10px 20px;border-radius:8px;text-decoration:none;">&#127758; Ver guía de actividades</a>' +
'</td></tr>' +
'<tr><td style="background:#3a3530;border-radius:0 0 16px 16px;padding:24px 40px;text-align:center;">' +
'<p style="margin:0 0 8px;font-size:18px;font-weight:300;color:#ffffff;font-family:Georgia,serif;">Las <em>Nubes</em></p>' +
'<p style="margin:0 0 12px;font-size:12px;color:rgba(255,255,255,0.6);">Buenos Aires, Chame · En las faldas de Chicá · Panamá Oeste</p>' +
'<a href="https://wa.me/50769812266" style="color:rgba(255,255,255,0.8);font-size:13px;text-decoration:none;">&#128172; WhatsApp: +507 6981-2266</a>' +
'</td></tr></table></td></tr></table></body></html>';
}

function sendConfirmationEmail(reservation, voucherBase64, voucherMimeType) {
  try {
    const email = reservation.email;
    if (!email) return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'No hay email de huésped' })).setMimeType(ContentService.MimeType.JSON);
    const cabin   = CABIN_NAMES_EMAIL[reservation.cabin] || reservation.cabin;
    const subject = '✅ Confirmación de reserva — ' + cabin + ' · Las Nubes';

    // Adjuntar recibo PDF si la reserva tiene voucher (codTransferencia o montoVoucher)
    const attachments = [];
    const montoVoucherNum = reservation.montoVoucher ? parseFloat(String(reservation.montoVoucher).replace(/[^\d.]/g, '')) || 0 : 0;
    const hasVoucher = !!(reservation.codTransferencia || montoVoucherNum > 0);
    Logger.log('🧾 sendConfirmationEmail · hasVoucher=' + hasVoucher + ' codT=' + (reservation.codTransferencia || 'null') + ' montoVoucher=' + (reservation.montoVoucher || 'null') + ' voucherBytes=' + (voucherBase64 ? voucherBase64.length : 0));
    const debug = { hasVoucher, codT: reservation.codTransferencia || null, montoVoucher: reservation.montoVoucher || null };
    if (hasVoucher) {
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
    if (voucherBase64 && voucherMimeType) {
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
  else if (r.origin === 'Cortesia' || r.origin === 'Colaboracion' || r.origin === 'Personal') metodo = 'Sin cobro';

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
  let publicLink = '';
  try { if (reservation.id) publicLink = getPublicReservaUrl(reservation.id); } catch(e) { Logger.log('no publicLink: ' + e.message); }

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
buildGuiaHTML(reservation.cabin, meta.tipo) +
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
    const subject     = 'Actualización de reserva — ' + cabin + ' · Las Nubes';
    const html = buildUpdateEmailHTML(reservation, cabin, color, formatDateES(checkinStr), formatDateES(checkoutStr), nightCount(checkinStr, checkoutStr), amount, deposit, saldo, hasSaldo, reservation.comentarios || '', googleCalLink(reservation));

    // Si la edicion incluye un voucher nuevo, generar recibo PDF y adjuntar
    // (mismo flujo que sendConfirmationEmail).
    const attachments = [];
    if (voucherBase64 && voucherMimeType) {
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
  const hasSaldo = amount > 0 && parseFloat(saldo) > 0;
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

// Trigger diario @ 10am Panama. Escanea Reservas y manda recordatorio
// a quienes hacen check-in REAL (display) mañana.
// Excluye CANCELADA, origen Abierta, y reservas sin email.
function enviarRecordatoriosCheckin() {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = Utilities.formatDate(tomorrow, 'America/Panama', 'yyyy-MM-dd');

  let sent = 0, skipped = 0, errors = 0;
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
      tipo:       data[i][24] || 'noche'
    };
    if (!r.checkin) continue;
    if (r.estadoPago === 'CANCELADA') continue;
    if (r.origin === 'Abierta') continue;
    const meta = tipoEmailMeta(r);
    if (meta.displayCheckin !== tomorrowStr) continue;
    if (!r.email) { skipped++; Logger.log('⊘ Sin email: ' + r.name); continue; }
    try {
      sendCheckinReminderEmail(r);
      sent++;
    } catch(e) {
      errors++;
      Logger.log('⚠ Error con ' + r.name + ' (' + r.email + '): ' + e);
    }
  }
  Logger.log('✓ Recordatorios: ' + sent + ' enviados, ' + skipped + ' sin email, ' + errors + ' errores');
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
}`;

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
    return ContentService.createTextOutput(JSON.stringify({ ok: true, proveedor: data.proveedor || '', fecha: data.fecha || '', monto: data.monto || 0, numFactura: data.numFactura || '', items: Array.isArray(data.items) ? data.items : [] })).setMimeType(ContentService.MimeType.JSON);
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

function saveVoucherToDrive(reservation, imageBase64, mimeType, fileName) {
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
    if (estadoActual === 'PAGA' && fechaPagoActual) continue;
    const row = i + 1;
    reservasSheet.getRange(row, 17).setValue(pago.fechaPago);
    reservasSheet.getRange(row, 18).setValue(pago.montoPagado);
    reservasSheet.getRange(row, 21).setValue('PAGA');
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
    const headers = ['FechaCancelacion', 'EmailId', 'CodConfirmacion', 'Huesped', 'Cabana', 'FechasReserva', 'CanceladoPor'];
    cancelSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    cancelSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  const cancelData   = cancelSheet.getDataRange().getValues().slice(1);
  const processedIds = new Set(cancelData.map(r => r[1].toString()));

  const threads = GmailApp.search(
    'from:automated@airbnb.com (subject:"Cancelada: reserva" OR subject:"Reservation cancelled") newer_than:365d'
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

      let canceladoPor = 'Huésped';
      if (/host|anfitri[oó]n|you cancelled|t[uú] cancel/i.test(body)) canceladoPor = 'Anfitrión';
      else if (/airbnb cancelled|airbnb cancel/i.test(body)) canceladoPor = 'Airbnb';

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
      }

      cancelSheet.appendRow([fechaCancelacion, msgId, confirmCode, huesped, cabana, fechasReserva, canceladoPor]);
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

    const ORIGENES_ICAL = ['directa', 'cortesia', 'colaboracion'];
    const reservas = rows.filter(r => {
      const cabinCode = r[3] ? r[3].toString().toLowerCase() : '';
      const origen    = r[9] ? r[9].toString().toLowerCase() : '';
      const cabinOk   = !cabin || cabinCode === cabin.toLowerCase();
      const origenOk  = ORIGENES_ICAL.includes(origen);
      return cabinOk && origenOk;
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