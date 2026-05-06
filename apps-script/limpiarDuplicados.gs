// ─── LIMPIAR DUPLICADOS ──────────────────────────────────────
// Ejecuta esta función UNA SOLA VEZ para eliminar duplicados.
// Lógica: si hay dos filas con el mismo código de confirmación,
// conserva la que empieza con "csv_" y elimina las demás.
function limpiarDuplicados() {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();

  // Columnas (0-indexed): 0=ID, 1=Nombre, 2=CabañaNombre, 3=CabañaCodigo,
  // 4=Entrada, 5=Salida, 6=Personas, 7=Monto, 8=Seña, 9=Origen, 10=CodConfirmacion

  const COL_ID   = 0;  // ID interno (csv_HM... o 19c...)
  const COL_CODE = 10; // Código de confirmación Airbnb (HM82DXFMYN)

  // Paso 1: mapear código de confirmación → filas que lo tienen
  // (ignorar fila de encabezado)
  const codeMap = {}; // confirmCode → [ {rowIndex, id} ]
  for (let i = 1; i < data.length; i++) {
    const confirmCode = data[i][COL_CODE]?.toString().trim();
    const id          = data[i][COL_ID]?.toString().trim();
    if (!confirmCode) continue;
    if (!codeMap[confirmCode]) codeMap[confirmCode] = [];
    codeMap[confirmCode].push({ rowIndex: i, id });
  }

  // Paso 2: identificar filas a eliminar
  // Si hay más de una fila con el mismo confirmCode:
  //   - Conservar la que tiene ID que empieza con "csv_"
  //   - Marcar las demás para eliminar
  const rowsToDelete = new Set();

  for (const [code, entries] of Object.entries(codeMap)) {
    if (entries.length <= 1) continue; // Sin duplicado, ignorar

    const csvEntry   = entries.find(e => e.id.startsWith('csv_'));
    const otherEntries = entries.filter(e => !e.id.startsWith('csv_'));

    if (csvEntry) {
      // Hay una versión CSV — eliminar las otras
      otherEntries.forEach(e => rowsToDelete.add(e.rowIndex));
    } else {
      // No hay versión CSV — conservar la primera, eliminar el resto
      entries.slice(1).forEach(e => rowsToDelete.add(e.rowIndex));
    }
  }

  Logger.log(`Filas duplicadas encontradas: ${rowsToDelete.size}`);

  if (rowsToDelete.size === 0) {
    SpreadsheetApp.getUi().alert('✓ No se encontraron duplicados. El Sheet está limpio.');
    return;
  }

  // Paso 3: eliminar de abajo hacia arriba para no desplazar índices
  const sortedRows = [...rowsToDelete].sort((a, b) => b - a);
  sortedRows.forEach(rowIndex => {
    sheet.deleteRow(rowIndex + 1); // +1 porque Sheets es 1-indexed
  });

  const msg = `✓ Limpieza completa.\n${rowsToDelete.size} filas duplicadas eliminadas.\nSe conservaron las versiones del CSV histórico.`;
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}