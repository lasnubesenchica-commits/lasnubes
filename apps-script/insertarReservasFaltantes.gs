/**
 * insertarReservasFaltantes()
 * 
 * Inserta las 2 reservas Airbnb faltantes del pago de Feb 28
 * y las marca directamente como PAGA con FechaPago = 2026-02-28
 * 
 * Ejecutar UNA SOLA VEZ desde el editor de Apps Script.
 */
function insertarReservasFaltantes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Reservas");
  
  if (!sheet) {
    Logger.log("❌ No se encontró la hoja 'Reservas'");
    return;
  }

  // Obtener headers para mapear columnas dinámicamente
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = {};
  headers.forEach((h, i) => { col[h.trim()] = i + 1; }); // 1-indexed

  Logger.log("Columnas detectadas: " + JSON.stringify(col));

  // Verificar que los códigos no existan ya
  const allData = sheet.getDataRange().getValues();
  const codConfIdx = col['CodConfirmacion'] - 1;
  const existingCodes = allData.map(r => r[codConfIdx]);

  const reservas = [
    {
      // HM8BH8C99E — Dalit Nathanson — Paseo por Las Nubes
      ID:              generateId(),
      Nombre:          "Dalit Nathanson",
      Cabaña:          "Paseo por Las Nubes",
      CabaňaCodigo:    "verde",
      Entrada:         "2026-02-19",
      Salida:          "2026-02-21",
      Personas:        1,
      Monto:           200,
      Abono:           0,
      Origen:          "Airbnb",
      CodConfirmacion: "HM8BH8C99E",
      "Service Fee":   6,
      Neto:            194,
      Alerta:          "",
      Pagador:         "",
      FechaReserva:    "2025-11-28",
      FechaPago:       "2026-02-28",
      MontoPagado:     194,
      codTransferencia:"",
      MontoVoucher:    0,
      EstadoPago:      "PAGA",
      Email:           "",
      Comentarios:     "Insertada manualmente - pago Feb 28 / email Nov 28 2025"
    },
    {
      // HMTKZY43CH — Sabrina Otto — Portal hacia Las Nubes
      ID:              generateId(),
      Nombre:          "Sabrina Otto",
      Cabaña:          "Portal hacia Las Nubes",
      CabaňaCodigo:    "azul",
      Entrada:         "2026-02-14",
      Salida:          "2026-02-16",
      Personas:        2,
      Monto:           200,
      Abono:           0,
      Origen:          "Airbnb",
      CodConfirmacion: "HMTKZY43CH",
      "Service Fee":   6,
      Neto:            194,
      Alerta:          "",
      Pagador:         "",
      FechaReserva:    "2025-11-02",
      FechaPago:       "2026-02-28",
      MontoPagado:     194,
      codTransferencia:"",
      MontoVoucher:    0,
      EstadoPago:      "PAGA",
      Email:           "",
      Comentarios:     "Insertada manualmente - pago Feb 28 / email Nov 2 2025"
    }
  ];

  let insertadas = 0;
  let omitidas = 0;

  reservas.forEach(r => {
    if (existingCodes.includes(r.CodConfirmacion)) {
      Logger.log("⚠️ OMITIDA (ya existe): " + r.CodConfirmacion);
      omitidas++;
      return;
    }

    // Construir la fila en el orden de los headers
    const fila = headers.map(h => {
      const key = h.trim();
      return (key in r) ? r[key] : "";
    });

    sheet.appendRow(fila);
    Logger.log("✅ INSERTADA: " + r.CodConfirmacion + " — " + r.Nombre);
    insertadas++;
  });

  Logger.log("─────────────────────────────");
  Logger.log("Insertadas: " + insertadas);
  Logger.log("Omitidas (duplicados): " + omitidas);

  if (insertadas > 0) {
    SpreadsheetApp.getUi().alert(
      "✅ Listo!\n\n" +
      "Reservas insertadas: " + insertadas + "\n" +
      "Omitidas (ya existían): " + omitidas + "\n\n" +
      "Ambas marcadas como PAGA con FechaPago = 2026-02-28"
    );
  } else {
    SpreadsheetApp.getUi().alert(
      "⚠️ No se insertó nada.\n\nAmbos códigos ya existen en la hoja."
    );
  }
}

/** Genera un ID único estilo UUID corto (misma lógica que tu sistema) */
function generateId() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 16);
}