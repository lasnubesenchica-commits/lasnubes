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

## Estrategia activa: LAY al favorito en Masters 1000

Pasiva, pre-partido (automatizable; sin in-play ni delay):

- **Mercado:** Match Odds de tennis ATP, torneos **Masters 1000**.
- **Acción:** **LAY** al jugador favorito (menor cuota).
- **Filtro:** cuota del favorito en **[1.20, 1.50]**.
- **Backtest (Pinnacle, 2023-25):** ~+10.6% ROI (5% comisión), consistente las 3
  temporadas. Específica de Masters 1000 (no Grand Slam ni ATP 250/500).
- **Caveats:** validar con cuotas reales del exchange; ~75% de apuestas pierden
  poco / ~25% ganan más (alta varianza, exige bankroll). Detalle en `TENNIS.md`.

## Estructura

```
betfair/
  backtest/        # análisis histórico (fútbol y tennis) + resultados (.md)
  bot/
    config.py          # endpoints + config desde variables de entorno
    betfair_client.py  # login (cert/interactivo) + Betting/Account JSON-RPC
    parsing.py         # interpreta respuestas de la API (testeable offline)
    strategy.py        # regla lay-favorito-Masters (compartida)
  capture/
    capture_masters.py # captura cuotas de cierre del exchange + paper P&L
  tests/
    test_strategy.py   # tests offline (sin red)
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

Corre `record` periódicamente durante un torneo Masters (p.ej. cada 5 min vía
cron) para guardar la cuota de cierre del favorito; luego `settle` para liquidar
contra el resultado y ver el P&L del paper trading. **Sólo lee; no apuesta.**

```bash
# Captura partidos que arrancan en los próximos 15 min
python3 -m betfair.capture.capture_masters record --within-min 15

# Liquida los ya jugados y muestra el P&L acumulado
python3 -m betfair.capture.capture_masters settle
```

Salida incremental en `capture/masters_capture.csv` (ignorado por git).

## Prueba de conexión (smoke test)

Verifica credenciales y conexión (sólo lectura): loguea, muestra el saldo y lista
los próximos partidos de tennis marcando Masters y favoritos en banda.

```bash
python3 -m betfair.tools.check_connection
```

## Captura automática en GitHub Actions

`.github/workflows/betfair-capture.yml` corre el capturador cada 15 min en los
meses de Masters 1000 (mar, abr, may, ago, oct, nov) y guarda el CSV como
artifact. **No coloca apuestas.**

- Para que el cron dispare, el workflow debe estar en `main`. En una rama de
  feature, úsalo con **Run workflow** (manual).
- Secrets requeridos: `BETFAIR_APP_KEY`, `BETFAIR_USERNAME`, `BETFAIR_PASSWORD`.
  Opcionales (cert): `BETFAIR_CERT_B64`, `BETFAIR_KEY_B64` (el `.crt`/`.key` en
  base64, p.ej. `base64 -w0 client-2048.crt`).

## Tests

```bash
python3 -m betfair.tests.test_strategy      # lógica de estrategia + parsing (offline)
```
