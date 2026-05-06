// ═══════════════════════════════════════════════════════════════════
//  LAS NUBES · Importar gastos publicitarios Instagram / Meta
//  RUTINA PUNTUAL — ejecutar UNA SOLA VEZ desde el editor de Apps Script
//  Período: 1 feb 2026 – 5 mar 2026
// ═══════════════════════════════════════════════════════════════════

function importarGastosInstagram() {

  // ── Datos del CSV (fecha D/M/YYYY, monto USD) ──────────────────
  const transacciones = [
    { fecha: '4/3/2026',  monto: 6.32  },
    { fecha: '4/3/2026',  monto: 2.50  },
    { fecha: '4/3/2026',  monto: 1.25  },
    { fecha: '3/3/2026',  monto: 10.00 },
    { fecha: '2/3/2026',  monto: 10.00 },
    { fecha: '1/3/2026',  monto: 10.00 },
    { fecha: '28/2/2026', monto: 10.00 },
    { fecha: '27/2/2026', monto: 10.00 },
    { fecha: '26/2/2026', monto: 10.00 },
    { fecha: '25/2/2026', monto: 10.00 },
    { fecha: '24/2/2026', monto: 10.00 },
    { fecha: '23/2/2026', monto: 10.00 },
    { fecha: '22/2/2026', monto: 10.00 },
    { fecha: '21/2/2026', monto: 10.00 },
    { fecha: '21/2/2026', monto: 10.00 },
    { fecha: '12/2/2026', monto: 10.00 },
    { fecha: '11/2/2026', monto: 10.00 },
    { fecha: '10/2/2026', monto: 10.00 },
    { fecha: '9/2/2026',  monto: 10.00 },
    { fecha: '8/2/2026',  monto: 10.00 },
    { fecha: '8/2/2026',  monto: 10.00 },
    { fecha: '7/2/2026',  monto: 9.71  },
    { fecha: '6/2/2026',  monto: 10.00 },
    { fecha: '5/2/2026',  monto: 10.00 },
    { fecha: '3/2/2026',  monto: 10.00 },
    { fecha: '2/2/2026',  monto: 10.00 },
    { fecha: '1/2/2026',  monto: 10.00 },
    { fecha: '1/2/2026',  monto: 10.00 },
  ];

  // ── Convertir fecha D/M/YYYY → YYYY-MM-DD ─────────────────────
  function parseFecha(str) {
    const parts = str.split('/');
    const d = parts[0].padStart(2, '0');
    const m = parts[1].padStart(2, '0');
    const y = parts[2];
    return y + '-' + m + '-' + d;
  }

  // ── Obtener/validar hoja Egresos ───────────────────────────────
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName('Egresos');

  if (!sheet) {
    sheet = ss.insertSheet('Egresos');
    sheet.getRange(1,1,1,8).setValues([['ID','Fecha','Descripcion','Monto','Categoria','Cabaña','Proveedor','URLFoto']]);
    sheet.getRange(1,1,1,8).setFontWeight('bold');
    Logger.log('✓ Hoja Egresos creada');
  }

  // ── Verificar que no estén ya importados (evitar duplicados) ───
  const existingData = sheet.getDataRange().getValues();
  const existingIds  = new Set(existingData.slice(1).map(r => r[0].toString()));

  // ── Insertar filas ──────────────────────────────────────────────
  let insertados = 0;
  let omitidos   = 0;

  transacciones.forEach((t, i) => {
    const fecha = parseFecha(t.fecha);
    // ID único basado en fecha y posición para evitar duplicados en re-ejecución
    const id = 'meta_ig_' + fecha.replace(/-/g,'') + '_' + String(i).padStart(2,'0');

    if (existingIds.has(id)) {
      omitidos++;
      Logger.log('⏩ Omitido (ya existe): ' + id);
      return;
    }

    sheet.appendRow([
      id,                          // ID
      fecha,                       // Fecha (YYYY-MM-DD)
      'Pauta Instagram / Meta Ads',// Descripcion
      t.monto,                     // Monto
      'Publicidad',                // Categoria
      '',                          // Cabaña (gasto general, no por cabaña)
      'Meta Platforms Ireland',    // Proveedor
      ''                           // URLFoto
    ]);

    insertados++;
  });

  // ── Resumen ────────────────────────────────────────────────────
  const total = transacciones.reduce((s, t) => s + t.monto, 0).toFixed(2);
  const msg = `✅ Importación completa\n\nInsertados: ${insertados}\nOmitidos:   ${omitidos}\nTotal:      $${total} USD\nPeríodo:    1 feb – 5 mar 2026`;

  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}