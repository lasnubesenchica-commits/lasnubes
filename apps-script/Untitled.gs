function revocarYReautorizar() {
  const token = ScriptApp.getOAuthToken();
  UrlFetchApp.fetch('https://oauth2.googleapis.com/revoke?token=' + token, {
    method: 'post',
    muteHttpExceptions: true
  });
  Logger.log('Token revocado. Ejecuta cualquier función para reautorizar.');
}