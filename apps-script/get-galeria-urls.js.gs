// ═══════════════════════════════════════════════════════════════════
//  LAS NUBES · Obtener URLs de galería desde Google Drive
//  Ejecutar UNA VEZ desde Apps Script — copia el output al HTML
// ═══════════════════════════════════════════════════════════════════

function getGaleriaUrls() {
  const FOLDERS = {
    verde: '1dZK7juPyPbzKhkwr3jzPEr8pqiatTs17',
    azul:  '1wq9k7Bgvc0sbyHi3HsJWlcRGd-SNBucO',
    lila:  '1GvNAxrL2iCbp_237-7SLIQD9xfXG2EhA'
  };

  const CABIN_NAMES = {
    verde: 'Paseo por Las Nubes',
    azul:  'Portal hacia Las Nubes',
    lila:  'Puente entre Las Nubes'
  };

  const result = {};

  Object.entries(FOLDERS).forEach(([cabin, folderId]) => {
    const folder = DriveApp.getFolderById(folderId);
    const files  = folder.getFiles();
    const imgs   = [];

    while (files.hasNext()) {
      const file = files.next();
      const mime = file.getMimeType();
      // Only images
      if (!mime.startsWith('image/')) continue;
      const id   = file.getId();
      imgs.push({
        id,
        name: file.getName(),
        thumb: `https://drive.google.com/thumbnail?id=${id}&sz=w800`,
        full:  `https://drive.google.com/thumbnail?id=${id}&sz=w1600`
      });
    }

    result[cabin] = imgs;
    Logger.log(`${CABIN_NAMES[cabin]}: ${imgs.length} imágenes`);
  });

  // ── Output: JavaScript array to paste into HTML ───────────────
  const lines = ['const GALERIA = {'];
  Object.entries(result).forEach(([cabin, imgs]) => {
    lines.push(`  ${cabin}: [`);
    imgs.forEach(img => {
      lines.push(`    { id:'${img.id}', name:'${img.name}', thumb:'${img.thumb}', full:'${img.full}' },`);
    });
    lines.push(`  ],`);
  });
  lines.push('};');

  const output = lines.join('\n');
  Logger.log('\n\n=== COPIA ESTO EN EL HTML ===\n\n' + output);

  // Also show in UI
  SpreadsheetApp.getUi().alert(
    'Galería generada ✓\n\nRevisa los Registros de ejecución (View → Execution log)\npara copiar el array GALERIA completo.'
  );
}