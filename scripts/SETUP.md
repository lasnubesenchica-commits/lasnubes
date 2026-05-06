# Setup — Auto-deploy de Apps Script

## 1. Secrets de GitHub

Repo → Settings → Secrets and variables → Actions → New repository secret.

| Secret | De donde sale |
|---|---|
| `GOOGLE_CLIENT_ID` | Google Cloud Console → APIs & Services → Credentials → OAuth 2.0 Client IDs (tipo "Web application"). Si reusas las de balanceclip, son los mismos valores. |
| `GOOGLE_CLIENT_SECRET` | Mismo OAuth Client ID — el secret que aparece junto al Client ID. |
| `GOOGLE_REFRESH_TOKEN` | Generado via [OAuth 2.0 Playground](https://developers.google.com/oauthplayground/). Scope: `https://www.googleapis.com/auth/script.projects` y `https://www.googleapis.com/auth/script.deployments`. |
| `GAS_SCRIPT_ID` | URL del editor de Apps Script: `https://script.google.com/.../projects/<SCRIPT_ID>/edit`. |
| `GAS_DEPLOYMENT_ID` | Editor → Deploy → Manage deployments → copia el "Deployment ID" del web app de produccion. (Opcional — si no lo pones el script busca el web app activo.) |

## 2. APIs habilitadas

En Google Cloud Console del proyecto OAuth, habilitar:

- Apps Script API (`https://console.cloud.google.com/apis/library/script.googleapis.com`)

## 3. Bootstrap (una sola vez)

Una vez configurados los secrets:

1. GitHub → Actions → "Pull Google Apps Script" → Run workflow.
2. El workflow baja el codigo actual del Apps Script a `apps-script/` en una rama nueva (`bot/pull-gas-<run_id>`).
3. Abre el PR sugerido y mergealo a `main`. A partir de ahi, el `.gs` vive en el repo.

## 4. Flujo regular

- Editas archivos en `apps-script/` y push a `main`.
- El workflow "Deploy Google Apps Script" se dispara, sube el codigo y actualiza el deployment de produccion.

## 5. Troubleshooting

- **403 / token expired**: regenera `GOOGLE_REFRESH_TOKEN` desde OAuth Playground.
- **404 deploymentId**: borra el secret `GAS_DEPLOYMENT_ID`; el script auto-detecta el web app activo.
- **Apps Script API not enabled**: habilitala en el proyecto de Google Cloud asociado al OAuth Client.
