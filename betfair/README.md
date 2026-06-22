# Betfair — Bot semanal "Lay the Draw"

Sistema para automatizar una estrategia de apuestas semanal sobre la cuenta de
Betfair vía la API oficial. En construcción por etapas.

## Estado

- [x] **Backtest** de la estrategia con datos históricos reales (`backtest/`).
- [ ] Esqueleto del bot (login por certificado, lectura de mercados) — simulación.
- [ ] Motor de reglas compartido (mismo que el backtest).
- [ ] Flujo semi-automático con aprobación en GitHub Actions.
- [ ] Integración live (cuota real + colocación de apuestas).
- [ ] Etapa 2: research previo + automatización total.

## Estrategia (definición acordada)

**Lay the Draw puro** (dejar correr hasta el final):

- Mercado: Match Odds, selección **Empate (The Draw)**.
- Acción: **LAY** (apostamos a que NO hay empate).
- Filtro: cuota del empate en **[3.00, 3.70]**.
- Sin cierre in-play: se deja correr hasta el resultado final.
- Stake: configurable (por defecto 1 unidad de backer stake por apuesta).
- Comisión Betfair: configurable (por defecto 5%).
- Cadencia: semanal.

## Backtest

`backtest/lay_the_draw.py` descarga los CSV de
[football-data.co.uk](https://www.football-data.co.uk/) (resultado final + cuotas
de cierre) y simula la regla exacta, reportando por liga y temporada: nº de
apuestas que pasan el filtro, % de empates, cuota promedio, P&L y ROI.

```bash
# Descarga directa (requiere acceso de red a football-data.co.uk)
python3 betfair/backtest/lay_the_draw.py

# Con CSVs locales (si la red está restringida): bajar los CSV a una carpeta
# y nombrarlos <LIGA>_<TEMP>.csv, p.ej. E0_2425.csv
python3 betfair/backtest/lay_the_draw.py --data-dir ./csv

# Parametrizable
python3 betfair/backtest/lay_the_draw.py --min 3.0 --max 3.7 \
    --commission 0.05 --leagues E0,SP1,F1 --seasons 2324,2425,2526
```

Ligas (códigos football-data.co.uk): `E0` Premier League, `SP1` La Liga,
`D1` Bundesliga, `F1` Ligue 1, `I1` Serie A, `N1` Eredivisie, `E1` Championship,
`P1` Primeira Liga. Temporadas: `2425` = 2024-25, etc.

> **Nota sobre el % de la liga vs. el filtro:** el % de empates *dentro del
> filtro 3.0–3.7* es más alto que el promedio de la liga, porque el filtro
> selecciona partidos donde el empate es más probable. El backtest mide
> justamente la tasa dentro del filtro, que es lo que determina la rentabilidad.

## API de Betfair (notas)

- App Key delayed (gratis) ya creada — para desarrollo y simulación.
- Login no interactivo con certificado SSL para que el bot opere solo.
- Credenciales (key, certificado, password) van como **GitHub Secrets**, nunca en
  el código ni en el cliente.
