# Tennis — estrategia pasiva candidata

Datos: tennis-data.co.uk (ATP), temporadas 2023-2025. Cuotas de cierre de
**Pinnacle** (`PSW/PSL`, las más sharp). Script: `tennis_explore.py`.
Se reconstruye favorito/underdog por la cuota (favorito = menor cuota).

## Contexto: por qué NO la estrategia in-play del usuario

La estrategia manual original (layear el break point 0-40 / 15-40 y cerrar en el
salto del mercado) es **trading reactivo punto a punto**: requiere ejecutar al
instante (apuesta agresiva → delay de 5s de Betfair) y un feed de marcador más
rápido que el mercado (la Streaming API solo da precios). **No es automatizable
como bot retail.** Por eso buscamos una estrategia PASIVA (pre-partido).

## Hallazgo: LAY al favorito corto en Masters 1000

**Layear (apostar contra) al favorito cuando su cuota está en ~1.2-1.5, en torneos
Masters 1000.** Pre-partido, sin in-play.

| Serie | ROI (lay fav 1.2-1.5, com. 5%) | Por temporada | Veredicto |
|---|---:|---|---|
| **Masters 1000** | **+10.6%** (n=579) | +8% / +7% / +18% | ✅ consistente 3/3 |
| ATP 500 | +4.6% (n=392) | +1% / −1% / +13% | ⚠️ irregular |
| ATP 250 | −4.5% (n=800) | +9% / −20% / −12% | ❌ |
| Grand Slam | −6.8% (n=453) | +4% / −11% / −18% | ❌ (best-of-5) |

### Por qué tiene sentido económico
- **Best-of-3** (Masters) = más varianza → el favorito pierde más seguido de lo
  que su cuota corta implica. En **best-of-5** (Grand Slam) el favorito es fiable
  → layearlo pierde (y los datos lo confirman: −6.8%).
- **Máximo perfil** (Masters) = mucho dinero recreativo respaldando a las
  estrellas → su cuota queda demasiado corta → valor en layarlas.
- Que funcione donde la lógica predice (Masters) y NO en ATP250/Grand Slam reduce
  la sospecha de azar: hay un mecanismo, no solo un número bonito.

### Robustez
- **Estable ante la banda**: 1.2-1.4, 1.2-1.45, 1.2-1.5, 1.15-1.5 → todas +9% a
  +14%. No es un bucket afinado al milímetro.
- **Aguanta comisión**: +10.6% al 5%, **+9.3% al 6.5%** (base Betfair), +8% al 8%.
- Positivo en las 3 temporadas por separado.

### Caveats (leer antes de apostar real)
1. **Cuotas de Pinnacle, no de Betfair exchange.** El margen de bookie infla el
   lay ~1.5-2%. Estimación realista en Betfair: **~+6-9%** tras comisión. Hay que
   **revalidar con cuotas reales del exchange** (el dataset de tennis no las trae).
2. **Una serie elegida entre varias** vía exploración → confirmar con
   **forward-testing** (paper trading) antes de capital real.
3. **Solo 3 temporadas** (~580 partidos). Muestra decente, no enorme.
4. **Perfil de varianza alto**: ~75% de las apuestas individuales PIERDEN (liability
   chico, porque la cuota del favorito es baja) y ~25% GANAN (más grande). Es
   +EV pero psicológicamente duro; exige disciplina de bankroll y aguantar rachas.
   El liability por apuesta es bajo (0.2-0.5× el backer stake).

## Por qué esto SÍ es automatizable (a diferencia del break-point)
- **Pre-partido**: no hay delay in-play ni carrera por el marcador.
- **Volumen manejable**: ~190 apuestas/año, concentradas en las ~9 semanas de
  Masters 1000 (Indian Wells, Miami, Montecarlo, Madrid, Roma, Canadá, Cincinnati,
  Shanghái, París). Encaja en un bot dirigido por eventos.
- **Liquidez**: los Masters tienen mercados líquidos en Betfair para layear.

## Siguiente paso recomendado
1. Validar la señal con **cuotas reales del Betfair Exchange** (capturar precios de
   cierre de Masters durante una temporada, o vía API histórica si se consigue).
2. **Paper trading** una tanda de Masters antes de arriesgar capital.
3. Si holdea: bot dirigido por eventos que, en semanas de Masters 1000, detecte
   partidos con favorito a cuota 1.2-1.5 y coloque el lay pre-partido, con gestión
   de stake/bankroll para la varianza.

## Reproducir
```bash
python3 betfair/backtest/tennis_explore.py --data-dir <tennis> --series "Masters 1000" --min-n 50
```
