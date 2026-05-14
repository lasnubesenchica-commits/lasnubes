// ═══════════════════════════════════════════════════════════
//  ID del huésped — upload, OCR, gating del key box
// ═══════════════════════════════════════════════════════════
//
//  El huésped sube foto de su cédula desde reserva.html. La foto se
//  guarda en Drive y la fecha de nacimiento se extrae con Claude.
//  Una vez subido, el código del key box queda desbloqueado en la
//  pagina pública.
//
//  La foto se borra automáticamente 60 dias después del checkout
//  (trigger borrarIdsHuespedesViejos). La fecha de nacimiento se
//  conserva para activar promo de cumpleaños en futuras visitas.
//
//  Columnas en hoja Reservas:
//    27 - IdHuespedURL    (URL en Drive)
//    28 - FechaNacimiento (YYYY-MM-DD)
// ═══════════════════════════════════════════════════════════

const HUESPED_ID_FOLDER_NAME = 'Las Nubes - IDs Huéspedes';
const COL_ID_HUESPED_URL     = 27;
const COL_FECHA_NACIMIENTO   = 28;

function _huespedIdFolder() {
  const folders = DriveApp.getFoldersByName(HUESPED_ID_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(HUESPED_ID_FOLDER_NAME);
}

// Garantiza que existan las columnas 27/28 en Reservas
function _ensureHuespedIdColumns(sheet) {
  if (sheet.getLastColumn() < COL_FECHA_NACIMIENTO) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes('IdHuespedURL')) {
      sheet.getRange(1, COL_ID_HUESPED_URL).setValue('IdHuespedURL');
      sheet.getRange(1, COL_ID_HUESPED_URL).setFontWeight('bold');
    }
    if (!headers.includes('FechaNacimiento')) {
      sheet.getRange(1, COL_FECHA_NACIMIENTO).setValue('FechaNacimiento');
      sheet.getRange(1, COL_FECHA_NACIMIENTO).setFontWeight('bold');
    }
  }
}

// OCR del ID con Claude — extrae fecha de nacimiento, nombre y tipo de documento.
function parseIdHuespedConClaude(imageBase64, mimeType) {
  try {
    const prompt = `Lee este documento de identidad (cédula panameña, pasaporte, cédula de extranjero E, o licencia de conducir) y devuelve EXCLUSIVAMENTE este JSON, sin markdown ni texto antes o después:

{
  "fechaNacimiento": "<YYYY-MM-DD>",
  "nombreCompleto": "<nombre tal como aparece en el documento>",
  "tipoDocumento": "<cedula | pasaporte | cedula-e | licencia | otro>",
  "numeroDocumento": "<numero limpio sin espacios ni guiones>"
}

Reglas:
- FECHA DE NACIMIENTO: campo "Fecha de Nacimiento", "Date of Birth", "Nacido", "Nac." o equivalente. Formatos posibles: DD-MM-YYYY, MM/DD/YYYY, DD MMM YYYY. Convertilo SIEMPRE a YYYY-MM-DD. Meses en español: ene=01 feb=02 mar=03 abr=04 may=05 jun=06 jul=07 ago=08 sep=09 oct=10 nov=11 dic=12. Si no puedes leerla con seguridad, devuelve "" (string vacío).
- NOMBRE COMPLETO: el nombre que aparece en el documento. Pasaporte: campo "Apellidos" + "Nombres". Cédula: nombre tal como aparece. Devuélvelo con capitalización normal (Juan Pérez García, no JUAN PEREZ GARCIA).
- TIPO DOCUMENTO: cedula (cédula panameña típica), pasaporte, cedula-e (cédula de extranjero residente E-XX-XXXX), licencia (licencia de conducir), otro si no es claro.
- NÚMERO DOCUMENTO: el número de identificación. Limpio, sin espacios ni guiones. Si es pasaporte, su número.

Si no puedes leer el documento con calidad suficiente, devuelve todos los campos como "" excepto tipoDocumento="otro".

Responde solo el JSON.`;

    const buildPayload = (model) => ({
      model,
      max_tokens: 400,
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

    const attempts = [
      { model: 'claude-opus-4-6', wait: 0 },
      { model: 'claude-opus-4-6', wait: 1500 },
      { model: 'claude-sonnet-4-6', wait: 1500 }
    ];

    let response, rawBody, status, result;
    for (const a of attempts) {
      if (a.wait) Utilities.sleep(a.wait);
      response = callClaude(buildPayload(a.model));
      status   = response.getResponseCode();
      rawBody  = response.getContentText();
      try { result = JSON.parse(rawBody); } catch(_) { result = null; }
      if (status === 200 && result && result.content && result.content[0]) break;
      Logger.log('Id OCR attempt ' + a.model + ' status ' + status + ': ' + rawBody.slice(0, 200));
      if (status !== 429 && status !== 529 && status < 500) break;
    }

    if (!result || !result.content || !result.content[0]) {
      return { ok: false, error: 'OCR_FAILED' };
    }
    let text = result.content[0].text || '';
    text = text.trim().replace(/^```(?:json)?\s*/i, '').replace(/```\s*$/, '');
    let parsed;
    try { parsed = JSON.parse(text); } catch(e) {
      Logger.log('Id OCR JSON parse fail: ' + text.slice(0, 300));
      return { ok: false, error: 'JSON_PARSE' };
    }
    const dob = (parsed.fechaNacimiento || '').toString().trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dob)) {
      return { ok: false, error: 'DOB_NOT_FOUND', parsed: parsed };
    }
    return { ok: true, parsed: parsed };
  } catch(err) {
    Logger.log('parseIdHuespedConClaude error: ' + err.message);
    return { ok: false, error: err.message };
  }
}

// Busca si el huésped (por email o teléfono) ya subió ID en alguna reserva previa.
// Devuelve { url, dob } o null.
function _findExistingHuespedId(email, telefono) {
  if (!email && !telefono) return null;
  const sheet = getOrCreateSheet();
  _ensureHuespedIdColumns(sheet);
  const data  = sheet.getDataRange().getValues();
  const e = (email    || '').toString().toLowerCase().trim();
  const t = (telefono || '').toString().replace(/\D/g, '');
  for (let i = data.length - 1; i >= 1; i--) {  // empezar de atrás (mas reciente)
    const url = data[i][COL_ID_HUESPED_URL - 1];
    const dob = data[i][COL_FECHA_NACIMIENTO - 1];
    if (!url && !dob) continue;
    const re = (data[i][21] || '').toString().toLowerCase().trim();
    const rt = (data[i][23] || '').toString().replace(/\D/g, '');
    if ((e && re === e) || (t && rt === t)) {
      return { url: url || '', dob: dob || '' };
    }
  }
  return null;
}

// Action: subir ID del huésped desde reserva.html
// POST { action:'uploadHuespedId', c:'<short>', imageBase64, mimeType }
//   (o id+t para backward compat)
function handleUploadHuespedId(payload) {
  try {
    let reservaId;
    if (payload.c) {
      reservaId = lookupReservaIdByShareCode(payload.c);
      if (!reservaId) return _jsonOut({ ok: false, error: 'NOT_FOUND' });
    } else if (payload.id && payload.t) {
      if (!_verifyReservaToken(payload.id, payload.t)) return _jsonOut({ ok: false, error: 'INVALID_TOKEN' });
      reservaId = payload.id;
    } else {
      return _jsonOut({ ok: false, error: 'MISSING_PARAMS' });
    }
    const r = _readReservaById(reservaId);
    if (!r) return _jsonOut({ ok: false, error: 'NOT_FOUND' });

    if (!payload.imageBase64 || !payload.mimeType) {
      return _jsonOut({ ok: false, error: 'MISSING_IMAGE' });
    }

    // OCR primero — si no se puede leer la DOB, rechazar y pedir foto mas clara.
    const ocr = parseIdHuespedConClaude(payload.imageBase64, payload.mimeType);
    if (!ocr.ok) {
      return _jsonOut({ ok: false, error: ocr.error || 'OCR_FAILED' });
    }
    const dob = ocr.parsed.fechaNacimiento;

    // Subir a Drive
    const folder    = _huespedIdFolder();
    const ext       = payload.mimeType.includes('png') ? '.png' : payload.mimeType.includes('webp') ? '.webp' : '.jpg';
    const safeName  = (r.name || 'huesped').toString().replace(/\s+/g, '_').replace(/[^A-Za-z0-9_]/g, '').slice(0, 30);
    const timestamp = Utilities.formatDate(new Date(), 'America/Panama', 'yyyyMMdd_HHmmss');
    const archName  = 'ID_' + safeName + '_' + reservaId + '_' + timestamp + ext;
    const blob      = Utilities.newBlob(Utilities.base64Decode(payload.imageBase64), payload.mimeType, archName);
    const file      = folder.createFile(blob);
    file.setDescription([
      'Reserva ID: '    + reservaId,
      'Huésped: '       + (r.name || ''),
      'Subido: '        + timestamp,
      'DOB extraída: '  + dob,
      'Nombre OCR: '    + (ocr.parsed.nombreCompleto || ''),
      'Tipo: '          + (ocr.parsed.tipoDocumento  || '')
    ].join('\n'));
    const url = file.getUrl();

    // Escribir en Reservas
    const sheet = getOrCreateSheet();
    _ensureHuespedIdColumns(sheet);
    const data  = sheet.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === reservaId.toString()) { rowIdx = i + 1; break; }
    }
    if (rowIdx > 0) {
      sheet.getRange(rowIdx, COL_ID_HUESPED_URL).setValue(url);
      sheet.getRange(rowIdx, COL_FECHA_NACIMIENTO).setValue(dob);
    }

    // Si hay alerta de nombre que no coincide, log para que el admin lo vea
    const nameOcr = (ocr.parsed.nombreCompleto || '').toLowerCase();
    const nameRes = (r.name || '').toLowerCase();
    let warning = null;
    if (nameOcr && nameRes && !_namesMatch(nameOcr, nameRes)) {
      warning = 'Nombre del ID (' + ocr.parsed.nombreCompleto + ') no coincide con el de la reserva (' + r.name + ')';
      Logger.log('⚠ ' + warning);
    }

    return _jsonOut({
      ok: true,
      dob: dob,
      nombreOcr: ocr.parsed.nombreCompleto || '',
      tipo: ocr.parsed.tipoDocumento || '',
      warning: warning
    });
  } catch(err) {
    Logger.log('handleUploadHuespedId error: ' + err.message);
    return _jsonOut({ ok: false, error: err.message });
  }
}

function _jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Match suave: si al menos 2 palabras del nombre OCR aparecen en el nombre de reserva (o viceversa).
function _namesMatch(a, b) {
  const norm = (s) => s.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').split(/\s+/).filter(w => w.length > 2);
  const A = norm(a), B = norm(b);
  let common = 0;
  for (const w of A) if (B.indexOf(w) >= 0) common++;
  return common >= 2 || (A.length === 1 && B.indexOf(A[0]) >= 0);
}

// Trigger diario: borra de Drive los IDs de huéspedes cuyo checkout + 60 días ya pasó.
// La fecha de nacimiento en Reservas se conserva.
function borrarIdsHuespedesViejos() {
  const sheet = getOrCreateSheet();
  _ensureHuespedIdColumns(sheet);
  const data  = sheet.getDataRange().getValues();
  const tz    = 'America/Panama';
  const now   = new Date();
  const cutoffMs = now.getTime() - 60 * 24 * 60 * 60 * 1000;

  let deleted = 0, skipped = 0;
  for (let i = 1; i < data.length; i++) {
    const url = (data[i][COL_ID_HUESPED_URL - 1] || '').toString();
    if (!url) continue;
    const checkoutRaw = data[i][5];
    if (!checkoutRaw) continue;
    const checkoutStr = checkoutRaw instanceof Date
      ? Utilities.formatDate(checkoutRaw, tz, 'yyyy-MM-dd')
      : checkoutRaw.toString().slice(0, 10);
    const checkoutMs = new Date(checkoutStr + 'T12:00:00-05:00').getTime();
    if (isNaN(checkoutMs)) continue;
    if (checkoutMs > cutoffMs) { skipped++; continue; }

    // Extraer file ID del URL (formato https://drive.google.com/file/d/<ID>/view o similar)
    const m = url.match(/\/file\/d\/([A-Za-z0-9_-]+)/) || url.match(/[?&]id=([A-Za-z0-9_-]+)/);
    if (!m) { Logger.log('No file ID en URL: ' + url); continue; }
    try {
      const file = DriveApp.getFileById(m[1]);
      file.setTrashed(true);
      sheet.getRange(i + 1, COL_ID_HUESPED_URL).setValue('');
      deleted++;
    } catch(e) {
      Logger.log('No se pudo borrar ' + m[1] + ': ' + e.message);
    }
  }
  Logger.log('✓ borrarIdsHuespedesViejos — ' + deleted + ' borrados, ' + skipped + ' aún en ventana');
  return { deleted, skipped };
}

function instalarTriggerBorradoIds() {
  // Eliminar triggers previos de la misma función
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'borrarIdsHuespedesViejos') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('borrarIdsHuespedesViejos')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .inTimezone('America/Panama')
    .create();
  Logger.log('✓ Trigger diario instalado: borrarIdsHuespedesViejos @ 3am Panama');
}
