function getActividadesImgs() {
  const folder = DriveApp.getFolderById('1DyomVd_7W4WWgMK1fawG2I0n_xIyAoNr');
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType().startsWith('image/'))
      Logger.log(f.getName() + ' | ' + f.getId());
  }
}