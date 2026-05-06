// ============================================================
//  MÓDULO: Banco General — Análisis de Cuenta de Ahorros
//  Sistema: Las Nubes System
//  Archivo: BancoGeneral_Module.gs
// ============================================================
//
//  SHEETS REQUERIDAS (se crean automáticamente):
//    - BG_Movimientos   → datos crudos del XLSX de Banco General
//    - BG_Clasificado   → transacciones con categoría AI
//    - BG_Dashboard     → P&L mensual, gráficas, alertas
//
//  CONFIGURACIÓN INICIAL:
//    1. Agrega tu API key de Anthropic en las propiedades del script:
//       Extensions → Apps Script → Project Settings → Script Properties
//       Key: ANTHROPIC_API_KEY   Value: sk-ant-...
//    2. Ejecuta setupBGSheets() una sola vez para crear las sheets
//    3. Para importar movimientos: corre importarMovimientosBG()
//       (el script te pedirá seleccionar el archivo XLSX)
// ============================================================

const ANTHROPIC_KEY_PROP = 'ANTHROPIC_API_KEY';

// Categorías disponibles para clasificación
const CATEGORIAS = [
  'Alimentación',
  'Transporte / Combustible',
  'Servicios públicos',
  'Salud',
  'Educación',
  'Entretenimiento',
  'Tecnología / Suscripciones',
  'Publicidad / Marketing',
  'Operaciones Las Nubes',
  'Transferencia recibida',
  'Transferencia enviada',
  'Retiro / ATM',
  'Compras / Retail',
  'Restaurantes / Comida',
  'Otros',
];

// ─────────────────────────────────────────────
//  SETUP: Crear sheets si no existen
// ─────────────────────────────────────────────
function setupBGSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _ensureSheet(ss, 'BG_Movimientos', [
    ['Fecha', 'Referencia', 'Descripción', 'Monto', 'Saldo Total', 'Tipo']
  ]);

  _ensureSheet(ss, 'BG_Clasificado', [
    ['Fecha', 'Descripción', 'Monto', 'Tipo', 'Categoría', 'Subcategoría', 'Nota AI', 'Confianza']
  ]);

  _ensureSheet(ss, 'BG_Dashboard', []);

  SpreadsheetApp.getUi().alert('✅ Sheets creadas correctamente.\n\nPróximo paso: corre importarMovimientosBG() para cargar tu XLSX.');
}

function _ensureSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers.length > 0) {
      sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
      sheet.getRange(1, 1, 1, headers[0].length)
        .setBackground('#1a1a2e')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// ─────────────────────────────────────────────
//  IMPORTAR XLSX de Banco General
//  Llama a esta función después de subir el archivo a Drive
// ─────────────────────────────────────────────
function importarMovimientosBG() {
  const ui = SpreadsheetApp.getUi();

  // Pedir el ID del archivo de Drive
  const response = ui.prompt(
    'Importar Movimientos',
    'Pega el ID del archivo XLSX de Banco General en Google Drive:\n(Abre el archivo en Drive → URL → el código entre /d/ y /view)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const fileId = response.getResponseText().trim();
  if (!fileId) {
    ui.alert('❌ No ingresaste un ID de archivo.');
    return;
  }

  try {
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    _parsearYCargarXLSX(blob);
  } catch (e) {
    ui.alert('❌ Error al acceder al archivo:\n' + e.message + '\n\nVerifica que el ID sea correcto y que tengas acceso.');
  }
}

// ─────────────────────────────────────────────
//  ALTERNATIVA: Importar desde CSV pegado manualmente
//  Útil si prefieres copiar/pegar los datos directamente
// ─────────────────────────────────────────────
function importarMovimientosDesdeSheet() {
  // Esta función asume que ya copiaste los datos del XLSX
  // directamente en la sheet BG_Movimientos (columnas A-F)
  // y solo necesitas normalizar/limpiar los datos

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BG_Movimientos');

  if (!sheet) {
    SpreadsheetApp.getUi().alert('❌ Sheet BG_Movimientos no existe. Corre setupBGSheets() primero.');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('❌ No hay datos en BG_Movimientos. Pega los movimientos primero.');
    return;
  }

  // Leer datos existentes y agregar columna Tipo
  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  const tipoCol = [];

  data.forEach((row, i) => {
    const monto = parseFloat(row[3]) || 0;
    tipoCol.push([monto >= 0 ? 'Crédito' : 'Débito']);
  });

  sheet.getRange(2, 6, tipoCol.length, 1).setValues(tipoCol);

  // Colorear débitos en rojo claro, créditos en verde claro
  data.forEach((row, i) => {
    const monto = parseFloat(row[3]) || 0;
    const color = monto >= 0 ? '#e6f4ea' : '#fce8e6';
    sheet.getRange(i + 2, 1, 1, 6).setBackground(color);
  });

  SpreadsheetApp.getUi().alert(`✅ ${data.length} movimientos normalizados.\n\nAhora corre clasificarConIA() para categorizar.`);
}

// ─────────────────────────────────────────────
//  PARSER XLSX via Sheets API (conversión automática)
// ─────────────────────────────────────────────
function _parsearYCargarXLSX(blob) {
  // Convertir XLSX a Google Sheets temporalmente para leer datos
  const resource = {
    title: 'BG_Temp_Import',
    mimeType: MimeType.GOOGLE_SHEETS,
  };

  const tempFile = Drive.Files.insert(resource, blob, { convert: true });
  const tempSS = SpreadsheetApp.openById(tempFile.id);
  const tempSheet = tempSS.getSheets()[0];
  const allData = tempSheet.getDataRange().getValues();

  // Limpiar archivo temporal
  DriveApp.getFileById(tempFile.id).setTrashed(true);

  // Encontrar fila de headers (buscar "Fecha")
  let headerRow = -1;
  for (let i = 0; i < allData.length; i++) {
    if (allData[i][1] === 'Fecha') {
      headerRow = i;
      break;
    }
  }

  if (headerRow === -1) {
    SpreadsheetApp.getUi().alert('❌ No se encontró la fila de headers en el archivo.');
    return;
  }

  // Extraer transacciones
  const movimientos = [];
  for (let i = headerRow + 1; i < allData.length; i++) {
    const row = allData[i];
    const fecha = row[1];
    const referencia = row[2];
    const descripcion = row[3];
    const monto = row[4];
    const saldo = row[5];

    if (!fecha || !descripcion || monto === '') continue;

    const montoNum = parseFloat(monto) || 0;
    const fechaStr = fecha instanceof Date
      ? Utilities.formatDate(fecha, 'America/Panama', 'yyyy-MM-dd')
      : String(fecha).slice(0, 10);

    movimientos.push([
      fechaStr,
      String(referencia || ''),
      String(descripcion),
      montoNum,
      parseFloat(saldo) || 0,
      montoNum >= 0 ? 'Crédito' : 'Débito'
    ]);
  }

  if (movimientos.length === 0) {
    SpreadsheetApp.getUi().alert('❌ No se encontraron transacciones en el archivo.');
    return;
  }

  // Volcar en BG_Movimientos
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('BG_Movimientos') || _ensureSheet(ss, 'BG_Movimientos', [
    ['Fecha', 'Referencia', 'Descripción', 'Monto', 'Saldo Total', 'Tipo']
  ]);

  // Limpiar datos previos (mantener header)
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }

  sheet.getRange(2, 1, movimientos.length, 6).setValues(movimientos);

  // Colorear filas
  movimientos.forEach((row, i) => {
    const color = row[3] >= 0 ? '#e6f4ea' : '#fce8e6';
    sheet.getRange(i + 2, 1, 1, 6).setBackground(color);
  });

  SpreadsheetApp.getUi().alert(
    `✅ ${movimientos.length} movimientos importados exitosamente.\n\nAhora corre clasificarConIA() para categorizar con inteligencia artificial.`
  );
}

// ─────────────────────────────────────────────
//  CLASIFICACIÓN CON IA (Claude API)
//  Procesa en lotes de 20 para no exceder límites
// ─────────────────────────────────────────────
function clasificarConIA() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(ANTHROPIC_KEY_PROP);
  if (!apiKey) {
    SpreadsheetApp.getUi().alert(
      '❌ No se encontró la API key de Anthropic.\n\n' +
      'Ve a Extensions → Apps Script → Project Settings → Script Properties\n' +
      'y agrega:\n  Key: ANTHROPIC_API_KEY\n  Value: sk-ant-...'
    );
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = ss.getSheetByName('BG_Movimientos');
  const dstSheet = ss.getSheetByName('BG_Clasificado') || _ensureSheet(ss, 'BG_Clasificado', [
    ['Fecha', 'Descripción', 'Monto', 'Tipo', 'Categoría', 'Subcategoría', 'Nota AI', 'Confianza']
  ]);

  if (!srcSheet || srcSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('❌ No hay datos en BG_Movimientos. Importa primero.');
    return;
  }

  const data = srcSheet.getRange(2, 1, srcSheet.getLastRow() - 1, 6).getValues();

  // Limpiar BG_Clasificado
  if (dstSheet.getLastRow() > 1) {
    dstSheet.deleteRows(2, dstSheet.getLastRow() - 1);
  }

  const BATCH_SIZE = 20;
  const resultados = [];

  for (let i = 0; i < data.length; i += BATCH_SIZE) {
    const lote = data.slice(i, i + BATCH_SIZE);
    const clasificados = _clasificarLote(lote, apiKey);
    resultados.push(...clasificados);

    // Pequeña pausa para no saturar la API
    if (i + BATCH_SIZE < data.length) {
      Utilities.sleep(500);
    }
  }

  // Escribir resultados
  if (resultados.length > 0) {
    dstSheet.getRange(2, 1, resultados.length, 8).setValues(resultados);

    // Colorear por tipo
    resultados.forEach((row, i) => {
      const color = row[2] >= 0 ? '#e6f4ea' : '#fce8e6';
      dstSheet.getRange(i + 2, 1, 1, 8).setBackground(color);
    });
  }

  SpreadsheetApp.getUi().alert(
    `✅ ${resultados.length} transacciones clasificadas.\n\nAhora corre generarDashboardBG() para ver tu reporte.`
  );
}

function _clasificarLote(lote, apiKey) {
  const transacciones = lote.map((row, i) => ({
    id: i,
    fecha: row[0],
    descripcion: row[2],
    monto: row[3],
    tipo: row[5]
  }));

  const prompt = `Eres un asistente financiero experto en finanzas personales en Panamá.
Clasifica cada transacción bancaria de una cuenta de ahorros personal.

Categorías disponibles:
${CATEGORIAS.map((c, i) => `${i + 1}. ${c}`).join('\n')}

Para cada transacción devuelve un objeto JSON con:
- id: el mismo id que recibiste
- categoria: exactamente una de las categorías listadas
- subcategoria: subcategoría más específica (texto libre, máximo 3 palabras)
- nota: insight breve útil (máximo 10 palabras)
- confianza: "Alta", "Media" o "Baja"

CONTEXTO IMPORTANTE:
- "YAPPY BG A" = transferencia enviada vía Yappy
- "YAPPY BG DE" = transferencia recibida vía Yappy  
- "BANCA MOVIL TRANSFERENCIA DE" = transferencia recibida
- "FACEBK" = publicidad en Facebook/Meta
- "APPLE.COM/BILL" = suscripción Apple
- "TERPEL" = combustible/gasolina
- El negocio "Las Nubes" es un complejo de cabañas turísticas en Panamá

Responde SOLO con un array JSON válido, sin texto adicional, sin backticks.

Transacciones:
${JSON.stringify(transacciones, null, 2)}`;

  try {
    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: 'claude-opus-4-5',
        max_tokens: 2000,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    });

    const json = JSON.parse(response.getContentText());
    if (json.error) throw new Error(json.error.message);

    const texto = json.content[0].text.trim();
    const clasificaciones = JSON.parse(texto);

    // Mapear de vuelta a filas
    return lote.map((row, i) => {
      const cl = clasificaciones.find(c => c.id === i) || {};
      return [
        row[0],                          // Fecha
        row[2],                          // Descripción
        row[3],                          // Monto
        row[5],                          // Tipo
        cl.categoria || 'Otros',         // Categoría
        cl.subcategoria || '',           // Subcategoría
        cl.nota || '',                   // Nota AI
        cl.confianza || 'Media'          // Confianza
      ];
    });

  } catch (e) {
    Logger.log('Error en lote: ' + e.message);
    // Si falla el lote, devolver sin categorizar
    return lote.map(row => [
      row[0], row[2], row[3], row[5],
      'Otros', '', 'Error de clasificación', 'Baja'
    ]);
  }
}

// ─────────────────────────────────────────────
//  DASHBOARD: P&L mensual + alertas
// ─────────────────────────────────────────────
function generarDashboardBG() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clSheet = ss.getSheetByName('BG_Clasificado');

  if (!clSheet || clSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('❌ No hay datos en BG_Clasificado. Corre clasificarConIA() primero.');
    return;
  }

  const dashSheet = ss.getSheetByName('BG_Dashboard') || ss.insertSheet('BG_Dashboard');
  dashSheet.clearContents();
  dashSheet.clearFormats();

  const data = clSheet.getRange(2, 1, clSheet.getLastRow() - 1, 8).getValues();

  // Calcular métricas
  const metricas = _calcularMetricas(data);

  // Escribir dashboard
  _escribirDashboard(dashSheet, metricas);

  ss.setActiveSheet(dashSheet);
  SpreadsheetApp.getUi().alert('✅ Dashboard generado correctamente.');
}

function _calcularMetricas(data) {
  const porMes = {};
  const porCategoria = {};
  const suscripciones = [];
  const recurrentes = {};

  data.forEach(row => {
    const fecha = String(row[0]);
    const mes = fecha.slice(0, 7); // yyyy-MM
    const monto = parseFloat(row[2]) || 0;
    const tipo = row[3];
    const categoria = row[4];
    const descripcion = row[1];

    // Por mes
    if (!porMes[mes]) porMes[mes] = { ingresos: 0, gastos: 0, neto: 0 };
    if (monto >= 0) porMes[mes].ingresos += monto;
    else porMes[mes].gastos += Math.abs(monto);
    porMes[mes].neto += monto;

    // Por categoría
    if (monto < 0) {
      if (!porCategoria[categoria]) porCategoria[categoria] = 0;
      porCategoria[categoria] += Math.abs(monto);
    }

    // Detectar suscripciones (Apple, Netflix, Spotify, Facebook Ads, etc.)
    const esSuscripcion = /APPLE|NETFLIX|SPOTIFY|FACEBK|GOOGLE|AMAZON|MICROSOFT|DISNEY/i.test(descripcion);
    if (esSuscripcion && monto < 0) {
      suscripciones.push({ descripcion: descripcion.slice(0, 40), monto: Math.abs(monto), fecha });
    }

    // Detectar pagos recurrentes por descripción similar
    const clave = descripcion.slice(0, 20).toUpperCase();
    if (monto < 0) {
      if (!recurrentes[clave]) recurrentes[clave] = { count: 0, total: 0, monto: Math.abs(monto) };
      recurrentes[clave].count++;
      recurrentes[clave].total += Math.abs(monto);
    }
  });

  // Top 5 categorías de gasto
  const topCategorias = Object.entries(porCategoria)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 8);

  // Pagos recurrentes (aparecen 2+ veces)
  const alertasRecurrentes = Object.entries(recurrentes)
    .filter(([k, v]) => v.count >= 2)
    .sort((a, b) => b[1].total - a[1].total)
    .slice(0, 5);

  return { porMes, topCategorias, suscripciones, alertasRecurrentes, totalTransacciones: data.length };
}

function _escribirDashboard(sheet, m) {
  const COL_HEADER_BG = '#1a1a2e';
  const COL_HEADER_FG = '#ffffff';
  const COL_SECTION = '#16213e';
  const COL_INGRESO = '#137333';
  const COL_GASTO = '#c5221f';

  let fila = 1;

  // ── TÍTULO ──
  sheet.getRange(fila, 1, 1, 6).merge()
    .setValue('📊 DASHBOARD FINANCIERO — BANCO GENERAL')
    .setBackground(COL_HEADER_BG).setFontColor(COL_HEADER_FG)
    .setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  fila += 2;

  // ── P&L POR MES ──
  sheet.getRange(fila, 1, 1, 5).merge()
    .setValue('ESTADO DE RESULTADOS MENSUAL')
    .setBackground(COL_SECTION).setFontColor(COL_HEADER_FG)
    .setFontWeight('bold');
  fila++;

  const headersPlL = ['Mes', 'Ingresos', 'Gastos', 'Neto', 'Estado'];
  sheet.getRange(fila, 1, 1, 5).setValues([headersPlL])
    .setBackground('#4a4e69').setFontColor('#ffffff').setFontWeight('bold');
  fila++;

  const meses = Object.keys(m.porMes).sort().reverse();
  meses.forEach(mes => {
    const d = m.porMes[mes];
    const estado = d.neto >= 0 ? '✅ Positivo' : '⚠️ Negativo';
    sheet.getRange(fila, 1, 1, 5).setValues([[
      mes,
      `$${d.ingresos.toFixed(2)}`,
      `$${d.gastos.toFixed(2)}`,
      `$${d.neto.toFixed(2)}`,
      estado
    ]]);
    sheet.getRange(fila, 4).setFontColor(d.neto >= 0 ? COL_INGRESO : COL_GASTO).setFontWeight('bold');
    fila++;
  });
  fila++;

  // ── TOP CATEGORÍAS DE GASTO ──
  sheet.getRange(fila, 1, 1, 3).merge()
    .setValue('TOP CATEGORÍAS DE GASTO')
    .setBackground(COL_SECTION).setFontColor(COL_HEADER_FG).setFontWeight('bold');
  fila++;

  sheet.getRange(fila, 1, 1, 3).setValues([['Categoría', 'Total Gastado', '% del Total']])
    .setBackground('#4a4e69').setFontColor('#ffffff').setFontWeight('bold');
  fila++;

  const totalGastos = m.topCategorias.reduce((s, [, v]) => s + v, 0);
  m.topCategorias.forEach(([cat, total]) => {
    const pct = totalGastos > 0 ? ((total / totalGastos) * 100).toFixed(1) + '%' : '0%';
    sheet.getRange(fila, 1, 1, 3).setValues([[cat, `$${total.toFixed(2)}`, pct]]);
    fila++;
  });
  fila++;

  // ── ALERTAS: SUSCRIPCIONES ──
  sheet.getRange(fila, 1, 1, 3).merge()
    .setValue('⚠️ ALERTAS: SUSCRIPCIONES DETECTADAS')
    .setBackground('#7b2d00').setFontColor('#ffffff').setFontWeight('bold');
  fila++;

  if (m.suscripciones.length === 0) {
    sheet.getRange(fila, 1).setValue('No se detectaron suscripciones.');
    fila++;
  } else {
    sheet.getRange(fila, 1, 1, 3).setValues([['Descripción', 'Monto', 'Fecha']])
      .setBackground('#4a4e69').setFontColor('#ffffff').setFontWeight('bold');
    fila++;
    m.suscripciones.forEach(s => {
      sheet.getRange(fila, 1, 1, 3).setValues([[s.descripcion, `$${s.monto.toFixed(2)}`, s.fecha]]);
      fila++;
    });
  }
  fila++;

  // ── ALERTAS: PAGOS RECURRENTES ──
  sheet.getRange(fila, 1, 1, 3).merge()
    .setValue('🔁 PAGOS RECURRENTES (2+ veces)')
    .setBackground('#1e3a5f').setFontColor('#ffffff').setFontWeight('bold');
  fila++;

  if (m.alertasRecurrentes.length === 0) {
    sheet.getRange(fila, 1).setValue('No se detectaron pagos recurrentes.');
    fila++;
  } else {
    sheet.getRange(fila, 1, 1, 3).setValues([['Descripción', 'Veces', 'Total Acumulado']])
      .setBackground('#4a4e69').setFontColor('#ffffff').setFontWeight('bold');
    fila++;
    m.alertasRecurrentes.forEach(([desc, v]) => {
      sheet.getRange(fila, 1, 1, 3).setValues([[
        desc, v.count, `$${v.total.toFixed(2)}`
      ]]);
      fila++;
    });
  }
  fila++;

  // ── RESUMEN TOTAL ──
  sheet.getRange(fila, 1, 1, 3).merge()
    .setValue(`Total transacciones analizadas: ${m.totalTransacciones}`)
    .setFontStyle('italic').setFontColor('#666666');

  // Ajustar columnas
  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 130);
}

// ─────────────────────────────────────────────
//  REPORTE AI NARRATIVO (texto libre de Claude)
//  Genera un análisis en lenguaje natural de tus finanzas
// ─────────────────────────────────────────────
function generarReporteNarrativoIA() {
  const apiKey = PropertiesService.getScriptProperties().getProperty(ANTHROPIC_KEY_PROP);
  if (!apiKey) {
    SpreadsheetApp.getUi().alert('❌ Falta la API key de Anthropic. Ver instrucciones en setupBGSheets().');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clSheet = ss.getSheetByName('BG_Clasificado');
  if (!clSheet || clSheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('❌ Corre clasificarConIA() primero.');
    return;
  }

  const data = clSheet.getRange(2, 1, clSheet.getLastRow() - 1, 8).getValues();
  const metricas = _calcularMetricas(data);

  const resumen = {
    meses: metricas.porMes,
    topGastos: metricas.topCategorias,
    suscripciones: metricas.suscripciones,
    recurrentes: metricas.alertasRecurrentes,
    totalTransacciones: metricas.totalTransacciones
  };

  const prompt = `Eres un asesor financiero personal experto. Analiza estos datos de una cuenta de ahorros personal en Panamá y genera un reporte narrativo breve en español.

Datos financieros:
${JSON.stringify(resumen, null, 2)}

El reporte debe incluir:
1. Resumen ejecutivo (2-3 oraciones)
2. Análisis del flujo de ingresos vs gastos por mes
3. Top 3 insights sobre patrones de gasto
4. Top 3 recomendaciones concretas de optimización
5. Una alerta si hay algo preocupante

Sé directo, específico con los números, y usa un tono profesional pero amigable. Máximo 400 palabras.`;

  try {
    const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: 'claude-opus-4-5',
        max_tokens: 1000,
        messages: [{ role: 'user', content: prompt }]
      }),
      muteHttpExceptions: true
    });

    const json = JSON.parse(response.getContentText());
    if (json.error) throw new Error(json.error.message);

    const reporte = json.content[0].text;

    // Mostrar en popup y también en la sheet de dashboard
    const dashSheet = ss.getSheetByName('BG_Dashboard');
    if (dashSheet) {
      const lastRow = dashSheet.getLastRow() + 2;
      dashSheet.getRange(lastRow, 1, 1, 5).merge()
        .setValue('🤖 ANÁLISIS IA — REPORTE NARRATIVO')
        .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
      dashSheet.getRange(lastRow + 1, 1, 1, 5).merge()
        .setValue(reporte)
        .setWrap(true);
      dashSheet.setRowHeight(lastRow + 1, 300);
    }

    SpreadsheetApp.getUi().alert('✅ Reporte generado. Revisa la sección final del Dashboard.');

  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Error al generar reporte: ' + e.message);
  }
}

// ─────────────────────────────────────────────
//  MENÚ PERSONALIZADO
// ─────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏦 Banco General')
    .addItem('1. Setup inicial (crear sheets)', 'setupBGSheets')
    .addSeparator()
    .addItem('2. Importar XLSX desde Drive', 'importarMovimientosBG')
    .addItem('2b. Normalizar datos pegados', 'importarMovimientosDesdeSheet')
    .addSeparator()
    .addItem('3. Clasificar con IA', 'clasificarConIA')
    .addItem('4. Generar Dashboard', 'generarDashboardBG')
    .addItem('5. Reporte narrativo IA', 'generarReporteNarrativoIA')
    .addToUi();
}