#!/usr/bin/env node
/**
 * deploy-gas.js
 *
 * Sube los archivos de apps-script/ al proyecto Google Apps Script de
 * Las Nubes, crea una version nueva y actualiza el deployment de
 * produccion (web app) automaticamente.
 *
 * Secrets requeridos en GitHub Actions:
 *   GOOGLE_CLIENT_ID      OAuth2 Client ID (Web application)
 *   GOOGLE_CLIENT_SECRET  OAuth2 Client Secret
 *   GOOGLE_REFRESH_TOKEN  Refresh token (via OAuth Playground)
 *   GAS_SCRIPT_ID         Script ID del proyecto Apps Script
 *   GAS_DEPLOYMENT_ID     Deployment ID del web app de produccion
 *                         (opcional — si no, se busca el activo)
 */

const { google } = require('googleapis');
const fs   = require('fs');
const path = require('path');

const GAS_DIR = path.join(process.cwd(), 'apps-script');
const DASHBOARD_HTML = path.join(process.cwd(), 'dashboard.html');

function extractDeploymentIdFromDashboard() {
  if (!fs.existsSync(DASHBOARD_HTML)) return null;
  const content = fs.readFileSync(DASHBOARD_HTML, 'utf8');
  const match = content.match(/script\.google\.com\/macros\/s\/([A-Za-z0-9_-]+)\/exec/);
  return match ? match[1] : null;
}

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

function readGasFiles(dir) {
  const files = [];
  for (const entry of fs.readdirSync(dir).sort()) {
    const fp = path.join(dir, entry);
    if (!fs.statSync(fp).isFile()) continue;
    const ext  = path.extname(entry).toLowerCase();
    const name = path.basename(entry, ext);
    const source = fs.readFileSync(fp, 'utf8');
    if (entry === 'appsscript.json') {
      files.push({ name: 'appsscript', type: 'JSON', source });
    } else if (ext === '.gs' || ext === '.js') {
      files.push({ name, type: 'SERVER_JS', source });
    } else if (ext === '.html') {
      files.push({ name, type: 'HTML', source });
    }
  }
  return files;
}

async function deploymentExists(api, scriptId, deploymentId) {
  try {
    await api.projects.deployments.get({ scriptId, deploymentId });
    return true;
  } catch (err) {
    if (err.response && err.response.status === 404) return false;
    throw err;
  }
}

async function findWebAppDeploymentId(api, scriptId) {
  const res = await api.projects.deployments.list({ scriptId });
  for (const dep of (res.data.deployments || [])) {
    if (!dep.deploymentConfig || !dep.deploymentConfig.versionNumber) continue;
    for (const ep of (dep.entryPoints || [])) {
      if (ep.entryPointType === 'WEB_APP') return dep.deploymentId;
    }
  }
  return null;
}

const LIMITE_VERSIONES = 200;

// Lista TODAS las versiones (la API pagina de a 200 por defecto).
async function listarVersiones(api, scriptId) {
  const todas = [];
  let pageToken;
  do {
    const res = await api.projects.versions.list({ scriptId, pageSize: 200, pageToken });
    (res.data.versions || []).forEach(v => todas.push(v));
    pageToken = res.data.nextPageToken;
  } while (pageToken);
  return todas;
}

// Qué versiones están tomadas por un deployment. Una versión en uso NO se puede
// borrar, así que sin este dato el usuario intenta borrar las más viejas y se
// topa con que justo esas son las que están pinneadas.
async function versionesEnUso(api, scriptId) {
  const res = await api.projects.deployments.list({ scriptId, pageSize: 200 });
  const enUso = new Map();
  (res.data.deployments || []).forEach(d => {
    const v = d.deploymentConfig && d.deploymentConfig.versionNumber;
    if (v) enUso.set(v, (enUso.get(v) || 0) + 1);
  });
  return enUso;
}

// El error crudo de Google no dice ni cuántas versiones hay, ni cuáles se
// pueden borrar, ni que el CÓDIGO ya se subió. Sin eso el mensaje parece un
// fallo de red y no una tarea concreta.
async function explicarLimiteVersiones(api, scriptId) {
  console.error('');
  console.error('════════════════════════════════════════════════════════════');
  console.error(' TOPE DE VERSIONES DE APPS SCRIPT');
  console.error('════════════════════════════════════════════════════════════');
  console.error(' El CÓDIGO YA SE SUBIÓ (ver "OK Codigo actualizado" arriba),');
  console.error(' pero no se pudo crear una versión nueva, así que el');
  console.error(' deployment sigue sirviendo la versión anterior. La Web App');
  console.error(' responde "Unknown action" a todo lo que sea nuevo.');
  console.error('');
  try {
    const versiones = await listarVersiones(api, scriptId);
    const enUso = await versionesEnUso(api, scriptId);
    const nums = versiones.map(v => v.versionNumber).sort((a, b) => a - b);
    const borrables = nums.filter(n => !enUso.has(n));
    console.error(` Versiones en el proyecto : ${versiones.length} / ${LIMITE_VERSIONES}`);
    if (nums.length) console.error(` Rango                    : v${nums[0]} … v${nums[nums.length - 1]}`);
    console.error(` En uso por un deployment : ${enUso.size}  (NO se pueden borrar)`);
    console.error(` Borrables                : ${borrables.length}`);
    if (borrables.length) {
      const muestra = borrables.slice(0, 20).map(n => 'v' + n).join(', ');
      console.error(` Las más viejas borrables : ${muestra}${borrables.length > 20 ? ', …' : ''}`);
    } else {
      console.error(' ⚠ Ninguna versión está libre: hay un deployment por versión.');
      console.error('   Hay que archivar deployments viejos antes de poder borrar.');
    }
  } catch (e) {
    console.error(` (no se pudo listar el detalle: ${e.message})`);
  }
  console.error('');
  console.error(' QUÉ HACER:');
  console.error(`   1. Abrir  https://script.google.com/home/projects/${scriptId}/versions`);
  console.error('   2. Seleccionar versiones viejas y borrarlas (se pueden varias a la vez).');
  console.error('   3. Volver a correr este workflow: Actions → Deploy Google Apps Script → Run workflow.');
  console.error('════════════════════════════════════════════════════════════');
  console.error('');
}

async function avisarSiCercaDelLimite(api, scriptId, versionNumber) {
  try {
    // El número de versión es incremental, así que sirve de estimador barato
    // sin pedir la lista entera en cada deploy.
    const restantes = LIMITE_VERSIONES - (await listarVersiones(api, scriptId)).length;
    if (restantes <= 20) {
      console.warn(`AVISO Quedan ${restantes} versiones antes del tope de ${LIMITE_VERSIONES}.`);
      console.warn(`      Conviene limpiar: https://script.google.com/home/projects/${scriptId}/versions`);
    }
  } catch (_) { /* el aviso es cortesía: nunca debe romper un deploy exitoso */ }
}

async function main() {
  const scriptId = process.env.GAS_SCRIPT_ID;
  if (!scriptId) throw new Error('Falta GAS_SCRIPT_ID');

  if (!fs.existsSync(GAS_DIR)) {
    throw new Error(`Directorio no encontrado: ${GAS_DIR}\nCorre primero "npm run pull:gas" para bajar el codigo del Apps Script.`);
  }

  const files = readGasFiles(GAS_DIR);
  if (files.length === 0) throw new Error(`Sin archivos .gs/.js/.html en ${GAS_DIR}/`);

  console.log(`Subiendo ${files.length} archivo(s) a Apps Script ${scriptId}:`);
  files.forEach(f => console.log(`  - ${f.name}  [${f.type}]`));

  const auth = getAuthClient();
  const api  = google.script({ version: 'v1', auth });

  // 1. Subir codigo
  await api.projects.updateContent({
    scriptId,
    requestBody: { scriptId, files }
  });
  console.log('OK Codigo actualizado');

  // 2. Crear nueva version
  let versionNumber;
  try {
    const verRes = await api.projects.versions.create({
      scriptId,
      requestBody: { description: `Auto-deploy ${new Date().toISOString()}` }
    });
    versionNumber = verRes.data.versionNumber;
    console.log(`OK Version ${versionNumber} creada`);
  } catch (err) {
    const apiErr = err.response && err.response.data && err.response.data.error;
    if (apiErr && apiErr.status === 'RESOURCE_EXHAUSTED') {
      await explicarLimiteVersiones(api, scriptId);
    }
    throw err;
  }

  // Aviso preventivo: sin esto el tope aparece de golpe, con el deploy ya roto.
  await avisarSiCercaDelLimite(api, scriptId, versionNumber);

  // 3. Determinar lista de deployments a actualizar.
  // Prioridad:
  //   a) deployment ID extraido de dashboard.html (fuente de verdad: el URL que usa el frontend)
  //   b) GAS_DEPLOYMENT_ID secret (legacy)
  //   c) primer web app deployment encontrado
  // Si dashboard.html y el secret apuntan a IDs distintos, actualizamos AMBOS para evitar drift.
  const dashboardDeploymentId = extractDeploymentIdFromDashboard();
  if (dashboardDeploymentId) {
    console.log(`-> Deployment ID en dashboard.html: ${dashboardDeploymentId}`);
  } else {
    console.warn('AVISO No pude extraer deployment ID de dashboard.html');
  }

  const candidates = new Set();
  if (dashboardDeploymentId) candidates.add(dashboardDeploymentId);
  if (process.env.GAS_DEPLOYMENT_ID) candidates.add(process.env.GAS_DEPLOYMENT_ID);

  const validIds = [];
  for (const id of candidates) {
    if (await deploymentExists(api, scriptId, id)) {
      validIds.push(id);
    } else {
      console.warn(`AVISO deployment ${id} no existe en el script, se omite`);
    }
  }

  if (validIds.length === 0) {
    const fallbackId = await findWebAppDeploymentId(api, scriptId);
    if (!fallbackId) throw new Error('No se encontro un web app deployment activo');
    console.log(`-> Fallback web app encontrado: ${fallbackId}`);
    validIds.push(fallbackId);
  }

  for (const deploymentId of validIds) {
    await api.projects.deployments.update({
      scriptId, deploymentId,
      requestBody: {
        deploymentConfig: {
          scriptId,
          versionNumber,
          manifestFileName: 'appsscript',
          description: `Auto-deploy ${new Date().toISOString()}`
        }
      }
    });
    console.log(`OK Deployment ${deploymentId} -> version ${versionNumber}`);
  }

  console.log('Despliegue completado.');
}

main().catch(err => {
  console.error('Error:', err.message);
  if (err.response && err.response.data && err.response.data.error) {
    console.error('API:', err.response.data.error);
  }
  process.exit(1);
});
