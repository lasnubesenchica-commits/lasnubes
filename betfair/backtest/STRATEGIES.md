# Exploración de estrategias — ¿hay algo de 8-10% ROI?

Escaneo sistemático de back/lay en 1X2 y Over/Under 2.5 por rango de cuota,
usando **cuotas de cierre del Betfair Exchange** (`BFE*`, las reales para operar),
comisión 5%, ROI normalizado a 1 unidad de capital en riesgo por apuesta.
Script: `explore_strategies.py`.

## Respuesta corta

**No.** En las cuotas de cierre del exchange (que son muy eficientes), **no existe
una estrategia robusta de 8-10% ROI** en esta data con reglas simples por rango de
cuota. Los buckets que tocaron +8-13% **no sobreviven** la prueba de consistencia
entre temporadas: son ruido / sobreajuste.

## Método y por qué importa

- **Cuotas del exchange, no de bookie.** Las columnas `BFE*` son del propio Betfair
  Exchange. Crucial: con cuotas de **bookie** (que llevan margen), cualquier
  estrategia de *lay* se ve falsamente rentable porque el margen juega a favor del
  layer. El "lay the draw +15%" del primer backtest estaba inflado por eso; en
  cuotas reales del exchange ese edge se diluye.
- **Riesgo normalizado.** Back y lay se comparan a 1 unidad de capital en riesgo
  (en lay, liability = 1 → backer stake = 1/(cuota−1)). En un exchange justo, back
  y lay dan ambos ≈ −comisión. Un bug inicial pagaba el lay como back e inventaba
  ROIs de +300%; corregido.
- **Cobertura:** el exchange (`BFE`) está en 2024-25 y 2025-26 (64% de partidos).
  2023-24 no trae exchange → solo 2 temporadas para validar. Muestra corta.

## Prueba de consistencia entre temporadas (cuotas BFE)

| Estrategia (cuota) | 2024-25 | 2025-26 | Veredicto |
|---|---:|---:|---|
| BACK_AWAY 2.0–2.5 | +13.5% | −0.6% | ❌ espejismo de 1 temporada |
| BACK_DRAW 3.0–3.5 | +2.3% | −2.6% | ❌ cambia de signo |
| LAY_UNDER25 2.0–2.5 | +2.5% | −0.4% | ⚠️ marginal |
| **BACK_HOME 1.0–1.5** | +1.5% | +7.2% | ✅ positivo ambas (~+3.8% prom) |
| **LAY_OVER25 2.0–2.5** | +3.9% | +2.6% | ✅ positivo ambas (~+3.2% prom) |

Una estrategia real debe ganar en cada temporada por separado. Solo dos lo logran,
y ambas con ROI **modesto (~3-4%), no 8-10%**, y con lógica económica conocida:

1. **Back a favoritos fuertes (local, cuota ≤ 1.5):** el clásico *favorite-longshot
   bias* — los grandes favoritos están levemente infravalorados. ROI ~4%, pero
   requiere mucho volumen (cuotas bajas) y el capital en riesgo por apuesta es alto.
2. **Lay del Over 2.5 (cuota ~2.0-2.5):** equivale a apostar a "pocos goles". Los
   apostadores recreativos tienden a respaldar goles, lo que abarata el over y deja
   valor en layarlo. ROI ~3%.

## Conclusión honesta sobre el 8-10%

Las cuotas de **cierre** ya incorporan casi toda la información: son lo más difícil
de batir. Un ROI sostenido de 8-10% **no sale de minar buckets de cuotas de cierre**;
sale de una de estas vías (todas = "research previo", la Etapa 2 ya prevista):

- **Closing Line Value (CLV):** apostar HORAS antes del cierre, a precios más blandos
  que el cierre. Si consigues sistemáticamente mejor precio que el de cierre, ganas.
- **Modelo predictivo propio** (xG, forma, lesiones, clima) que detecte mispricing
  puntual, en vez de reglas fijas por cuota.
- **Mercados menos eficientes:** ligas menores, in-play, props — menos liquidez,
  más sesgos explotables (pero más varianza y límites de stake).

## Recomendación

- Para un bot **simple y honesto**: combinar las dos señales robustas (back favorito
  fuerte + lay over 2.5) da ~3-4% esperado, no 8-10%. Útil como base de bajo riesgo.
- Para apuntar a 8-10%: pasar a la **Etapa 2** (modelo + CLV), validando siempre en
  **paper trading** antes de capital real. Aquí el edge es de modelado, no de reglas
  por cuota de cierre.
- **Cuidado con el sobreajuste:** se escanearon ~90 buckets; encontrar algunos de
  +10% por azar es lo esperado. Solo cuenta lo que sobrevive out-of-sample.

## Reproducir

```bash
python3 betfair/backtest/explore_strategies.py --data-dir <fd> --min-n 150 --top 30
```
