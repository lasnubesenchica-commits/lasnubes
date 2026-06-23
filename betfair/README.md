# Betfair — bot de apuestas (investigación + automatización)

Sistema para automatizar estrategias de apuestas sobre la cuenta de Betfair vía
la API oficial. Construcción por etapas, con backtesting riguroso antes de
arriesgar capital.

## Estado

- [x] **Backtest fútbol** (`backtest/lay_the_draw.py`, `explore_strategies.py`) —
      conclusión: sin edge atractivo robusto. Ver `RESULTS.md`, `STRATEGIES.md`.
- [x] **Backtest tennis** (`backtest/tennis_explore.py`) — **hallazgo robusto**:
      lay al favorito corto en Masters 1000. Ver `TENNIS.md`.
- [x] **Cliente Betfair** (login + lecturas) y **capturador de cuotas del exchange**.
- [ ] Validación con cuotas reales del exchange (paper trading una tanda de Masters).
- [ ] Bot dirigido por eventos (modo live, con gestión de bankroll).

## Estrategias capturadas (paper trading con cuotas reales)

Pasivas, pre-partido (automatizables; sin in-play ni delay):

- **🎾 Tennis — Lay al favorito en Masters 1000.** Banda de cuota **[1.20, 1.50]**.
  Backtest (Pinnacle 2023-25): ~+10.6% ROI, consistente. Detalle en `TENNIS.md`.
- **⚽ Fútbol — Lay the Draw.** Banda de cuota del empate **[3.00, 3.70]**.
  Mundial 2026 ahora + ligas top al reiniciar. El backtest histórico no fue
  rentable de media; capturamos datos reales para reevaluar al cabo de la temporada.

El capturador registra ambos con el mismo motor y un esquema común. ~75% de los
lays individuales pierden poco / ~25% ganan más (alta varianza, exige bankroll).

## Estructura

```
betfair/
  backtest/        # análisis histórico (fútbol y tennis) + resultados (.md)
  bot/
    config.py          # endpoints + config desde variables de entorno
    betfair_client.py  # login (cert/interactivo) + Betting/Account JSON-RPC
    parsing.py         # interpreta respuestas de la API (testeable offline)
    strategy.py        # selector compartido: make_pick(sport, ...) tennis/fútbol
  capture/
    capture.py         # captura genérica (--sport tennis|football) + paper P&L
    export_json.py     # CSV -> JSON para el dashboard (agregados, equity)
    make_demo.py       # genera datasets demo desde los backtests
  tests/
    test_strategy.py   # tests offline (sin red)
data/                  # JSON publicado que lee el dashboard (vivo + demo)
```

## Configuración

1. Instala dependencias: `pip install -r betfair/requirements.txt`
2. Copia `betfair/.env.example` a `.env` y rellena credenciales (ver abajo).
3. **(Opcional, recomendado) certificado** para login desatendido:
   ```bash
   openssl req -newkey rsa:2048 -nodes -keyout client-2048.key \
       -x509 -days 3650 -out client-2048.crt
   ```
   Sube el `.crt` en tu cuenta Betfair (My Account → Security → Automated betting)
   y apunta `BETFAIR_CERT_FILE` / `BETFAIR_KEY_FILE` al `.crt`/`.key`.
   Sin certificado, el cliente usa login interactivo (usuario+clave).

> Las credenciales y el `.key` NUNCA se commitean (`.gitignore` los excluye) y en
> GitHub Actions/VPS van como **Secrets**. La **delayed key** basta para CAPTURAR
> datos; para apostar real hace falta la **live key**.

## Capturar cuotas reales del exchange (validación)

`record` guarda la cuota de LAY de la selección (favorito en tennis, empate en
fútbol) de los partidos que arrancan pronto; `settle` liquida contra el resultado
y calcula el P&L del paper trading. **Sólo lee; no apuesta.**

```bash
python3 -m betfair.capture.capture record --sport tennis     # Masters 1000
python3 -m betfair.capture.capture record --sport football   # Mundial + ligas
python3 -m betfair.capture.capture settle --sport football
```

Salida en `betfair/data/<deporte>_capture.csv` (se publica para el dashboard).

## Prueba de conexión (smoke test)

Verifica credenciales y conexión (sólo lectura): loguea, muestra el saldo y lista
los próximos partidos de tennis marcando Masters y favoritos en banda.

```bash
python3 -m betfair.tools.check_connection
```

## Dashboard (`betfair.html`)

Página estática servida por GitHub Pages (`lasnubes.cloud/betfair.html`). **Solo
lectura**: selector de deporte (🎾 Tennis / ⚽ Football), KPIs, curva de equity,
resultados jornada a jornada, capturas y resumen del backtest. Nunca coloca
apuestas ni maneja credenciales.

- En vivo: `data/masters.json` (tennis), `data/football.json` (fútbol).
- Demo: `data/masters_demo.json`, `data/football_demo.json` — para ver el
  dashboard funcionando ya. Regenerar:
  ```bash
  python3 -m betfair.capture.make_demo --sport tennis   --src <dir_tennis> --season 2025
  python3 -m betfair.capture.make_demo --sport football --src <dir_fd>     --season 2425
  ```

El JSON lo genera `export_json.py` (P&L, ROI sobre liability, por día, equity).

## Captura automática en GitHub Actions

`.github/workflows/betfair-capture.yml` corre cada 20 min, captura tennis y fútbol,
y publica `betfair/data/*.json` en el repo para que Pages los sirva. **No apuesta.**

- Para que el cron dispare, el workflow debe estar en `main`. En una rama de
  feature, úsalo con **Run workflow** (manual).
- Secrets requeridos: `BETFAIR_APP_KEY`, `BETFAIR_USERNAME`, `BETFAIR_PASSWORD`.
  Opcionales (cert): `BETFAIR_CERT_B64`, `BETFAIR_KEY_B64` (el `.crt`/`.key` en
  base64, p.ej. `base64 -w0 client-2048.crt`).

## Tests

```bash
python3 -m betfair.tests.test_strategy      # lógica de estrategia + parsing (offline)
```
