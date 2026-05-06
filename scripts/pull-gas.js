#!/usr/bin/env node
/**
 * pull-gas.js
 *
 * Baja el contenido del proyecto Apps Script a la carpeta apps-script/.
 * Pensado para correr una vez (bootstrap) o cuando alguien edita
 * directamente en el editor de Google y necesita resincronizar.
 *
 * Uso local:
 *   GOOGLE_CLIENT_ID=... GOOGLE_CLIENT_SECRET=... GOOGLE_REFRESH_TOKEN=... \
 *     GAS_SCRIPT_ID=... node scripts/pull-gas.js
 *
 * Uso en GitHub Actions: ver .github/workflows/pull-gas.yml
 *   (workflow_dispatch — corre manual desde la pestana Actions).
 */

const { google } = require('googleapis');
const fs   = require('fs');
const path = require('path');

const GAS_DIR = path.join(process.cwd(), 'apps-script');

function getAuthClient() {
  const clientId     = process.env.GOOGLE_CLIENT_ID;
  const clientSecret = process.env.GOOGLE_CLIENT_SECRET;
  const refreshToken = process.env.GOOGLE_REFRESH_TOKEN;
  if (!clientId || !clientSecret || !refreshToken) {
    throw new Error('Faltan: GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_REFRESH_TOKEN');
  }
  const auth = new google.auth.OAuth2(clientId, clientSecret);
  auth.setCredentials({ refresh_token: refreshToken });
  return auth;
}

function fileExtFor(type) {
  if (type === 'JSON')      return '.json';
  if (type === 'SERVER_JS') return '.gs';
  if (type === 'HTML')      return '.html';
  return null;
}

async function main() {
  const scriptId = process.env.GAS_SCRIPT_ID;
  if (!scriptId) throw new Error('Falta GAS_SCRIPT_ID');

  const auth = getAuthClient();
  const api  = google.script({ version: 'v1', auth });

  console.log(`Bajando contenido de Apps Script ${scriptId}...`);
  const res = await api.projects.getContent({ scriptId });
  const files = res.data.files || [];

  fs.mkdirSync(GAS_DIR, { recursive: true });

  let written = 0, skipped = 0;
  for (const file of files) {
    const ext = fileExtFor(file.type);
    if (!ext) {
      console.log(`  - Saltando ${file.name} (tipo ${file.type})`);
      skipped++;
      continue;
    }
    const outPath = path.join(GAS_DIR, file.name + ext);
    fs.writeFileSync(outPath, file.source);
    console.log(`  OK ${file.name}${ext} (${file.source.length} bytes)`);
    written++;
  }
  console.log(`\nListo. ${written} archivo(s) escritos en ${GAS_DIR}/, ${skipped} saltados.`);
}

main().catch(err => {
  console.error('Error:', err.message);
  if (err.response && err.response.data && err.response.data.error) {
    console.error('API:', err.response.data.error);
  }
  process.exit(1);
});
