# Resultados del backtest — Lay the Draw

Datos: football-data.co.uk (cuotas de cierre `AvgCD`, resultado final `FTR`).
Temporadas 2023-24, 2024-25, 2025-26. Lay puro (dejar correr), comisión 5%,
1 unidad de backer stake por apuesta. P&L en unidades; ROI sobre stake.

> ⚠️ La temporada 2025-26 puede estar parcial en el mirror para alguna liga.

## Estrategia original tal cual (5 ligas, cuota empate 3.0–3.7)

| Liga | Apuestas | % Empates | Cuota prom | P&L | ROI |
|---|---:|---:|---:|---:|---:|
| Premier League | 419 | 28.2% | 3.45 | −2.46u | −0.6% |
| La Liga | 574 | 31.9% | 3.31 | −49.19u | −8.6% |
| Bundesliga | 356 | 31.5% | 3.48 | −45.05u | −12.7% |
| **Ligue 1** | 470 | **26.8%** | 3.41 | **+23.29u** | **+5.0%** |
| Serie A | 599 | 30.9% | 3.31 | −32.34u | −5.4% |
| **TOTAL** | **2418** | **29.9%** | 3.38 | **−105.75u** | **−4.4%** |

**La estrategia, aplicada a las 5 ligas, PIERDE** (−4.4%). El punto de equilibrio
está en ~28.6% de empates dentro del filtro; el global (29.9%) queda por encima.

## Hallazgos clave

1. **La premisa de la Premier League ya no aplica.** La EPL pasó de "pocos
   empates" a ~28% en el filtro; apenas breakeven (mejorando: −11.6% → +0.7% →
   +5.2% por temporada).
2. **Ligue 1 es la única consistentemente rentable** (+5.0%, 470 apuestas), por
   su baja tasa de empates (26.8%). 2024-25 fue +14.3%.
3. **La Liga, Bundesliga y Serie A pierden** de forma clara. No usar.
4. **El borde está en la parte BAJA del rango de cuotas.** Restringiendo a
   Ligue 1 + Premier League:

   | Rango cuota empate | Apuestas | % Emp | P&L | ROI |
   |---|---:|---:|---:|---:|
   | **3.0–3.3** | 218 | 25.2% | **+33.62u** | **+15.4%** |
   | 3.0–3.5 | 572 | 27.6% | +25.07u | +4.4% |
   | 3.0–3.7 | 889 | 27.4% | +20.83u | +2.3% |
   | 3.2–3.7 | 800 | 28.0% | −2.29u | −0.3% |
   | 3.3–3.7 | 686 | 28.0% | −8.29u | −1.2% |
   | 3.4–3.7 | 506 | 27.5% | −4.81u | −1.0% |

   Es decir, el mercado parece **sobrevalorar el empate cuando su cuota es baja
   (3.0–3.3)**: los empates ocurren menos (25.2%) de lo que la cuota implica
   (30–33%). Ahí está el edge.

## Caveats (importantes antes de apostar real)

- **Cuotas de bookie ≠ Betfair exchange.** Usamos la cuota de cierre promedio de
  casas como proxy. En Betfair laiarías a cuotas algo distintas (normalmente un
  poco más altas en el exchange → mayor liability). Hay que revalidar con cuotas
  reales del exchange.
- **Comisión 5%** asumida; tu tarifa real puede variar (2–5% según base rate).
- **Riesgo de sobreajuste.** Cuanto más se afina el sub-rango (3.0–3.3), más se
  ajusta al ruido del pasado. El +15.4% es prometedor pero con muestra menor
  (218). El veredicto sólido es: **Ligue 1, y borde concentrado en cuotas bajas.**
- El edge es **fino**: requiere disciplina, registro de cada apuesta, y validación
  hacia adelante (paper trading) antes de subir stakes.

## Recomendación

1. Operar **Ligue 1** (núcleo), opcionalmente **Premier League**.
2. Filtro de cuota del empate **3.0–3.3** (o configurable; revalidar en vivo).
3. **No** operar La Liga, Bundesliga ni Serie A con esta regla.
4. Arrancar en **paper trading / dry-run** una temporada para validar con cuotas
   reales de Betfair antes de arriesgar capital.

## Reproducir

```bash
python3 betfair/backtest/lay_the_draw.py --data-dir <carpeta_csv> \
    --leagues F1,E0 --seasons 2324,2425,2526 --min 3.0 --max 3.3
```
