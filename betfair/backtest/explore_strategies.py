#!/usr/bin/env python3
"""
Explorador sistematico de estrategias sobre datos historicos de football-data.co.uk.

Usa las cuotas de CIERRE del BETFAIR EXCHANGE (columnas BFE*) cuando estan
disponibles -> los numeros son directamente relevantes para operar en Betfair
(sin el margen de bookie que infla artificialmente las estrategias de lay).
Cae a la cuota promedio de mercado (Avg*, con margen) solo si falta la del exchange.

Escanea, para cada (mercado, lado, seleccion), el ROI realizado por rango de
cuota, apostando flat 1 unidad de stake, con comision configurable sobre la
ganancia neta. Reporta un leaderboard ordenado por ROI, filtrando por muestra
minima para evitar conclusiones de poco soporte.

ADVERTENCIA METODOLOGICA: escanear muchas combinaciones y quedarse con la mejor
es sobreajuste. Trata cada hallazgo como HIPOTESIS, no como certeza; valida con
muestra grande, logica economica y forward-testing.

Uso:
    python3 explore_strategies.py --data-dir ./fd
    python3 explore_strategies.py --data-dir ./fd --commission 0.05 --min-n 150
    python3 explore_strategies.py --data-dir ./fd --leagues F1,E0
"""

from __future__ import annotations

import argparse
import csv
import io
import os
from dataclasses import dataclass

LEAGUES = ["E0", "SP1", "D1", "F1", "I1"]
SEASONS = ["2324", "2425", "2526"]
LEAGUE_NAMES = {"E0": "Premier League", "SP1": "La Liga", "D1": "Bundesliga",
                "F1": "Ligue 1", "I1": "Serie A"}


def _f(row, *cols):
    """Primer valor float valido (>1) entre las columnas dadas, o None."""
    for c in cols:
        v = (row.get(c) or "").strip()
        if v:
            try:
                x = float(v)
                if x > 1.0:
                    return x
            except ValueError:
                pass
    return None


def _goals(row):
    try:
        return int(row["FTHG"]) + int(row["FTAG"])
    except (ValueError, KeyError):
        return None


# Cada "mercado" devuelve (cuota, gano_la_apuesta) o None si faltan datos.
# Lado BACK: gano si ocurre la seleccion. Lado LAY: gano si NO ocurre.
def market_defs():
    """Lista de (id, descripcion, funcion_seleccion). Prefiere exchange BFE*."""
    defs = []

    def back(name, desc, odds_cols, win_fn):
        defs.append((name, desc, "back", odds_cols, win_fn))

    def lay(name, desc, odds_cols, happen_fn):
        # en lay ganamos si la seleccion NO ocurre
        defs.append((name, desc, "lay", odds_cols, lambda r: (not happen_fn(r))
                     if happen_fn(r) is not None else None))

    # ---- 1X2 ----
    back("BACK_HOME", "Back local", ("BFECH", "AvgCH"), lambda r: r["FTR"] == "H")
    back("BACK_AWAY", "Back visitante", ("BFECA", "AvgCA"), lambda r: r["FTR"] == "A")
    back("BACK_DRAW", "Back empate", ("BFECD", "AvgCD"), lambda r: r["FTR"] == "D")
    lay("LAY_HOME", "Lay local", ("BFECH", "AvgCH"), lambda r: r["FTR"] == "H")
    lay("LAY_AWAY", "Lay visitante", ("BFECA", "AvgCA"), lambda r: r["FTR"] == "A")
    lay("LAY_DRAW", "Lay empate", ("BFECD", "AvgCD"), lambda r: r["FTR"] == "D")

    # ---- Over / Under 2.5 ----
    def over(r):
        g = _goals(r)
        return None if g is None else g >= 3

    def under(r):
        g = _goals(r)
        return None if g is None else g <= 2

    back("BACK_OVER25", "Back +2.5 goles", ("BFEC>2.5", "AvgC>2.5"), over)
    back("BACK_UNDER25", "Back -2.5 goles", ("BFEC<2.5", "AvgC<2.5"), under)
    lay("LAY_OVER25", "Lay +2.5 goles", ("BFEC>2.5", "AvgC>2.5"), over)
    lay("LAY_UNDER25", "Lay -2.5 goles", ("BFEC<2.5", "AvgC<2.5"), under)
    return defs


@dataclass
class Stat:
    n: int = 0
    wins: int = 0
    pnl: float = 0.0
    odds_sum: float = 0.0
    exch: int = 0  # cuantas usaron cuota del exchange (no fallback)

    def add(self, odds, won, commission, used_exchange, side):
        # ROI normalizado a 1 unidad de CAPITAL EN RIESGO por apuesta, para que
        # back y lay sean comparables. En un exchange justo, ambos dan ~ -comision.
        self.n += 1
        self.odds_sum += odds
        self.exch += 1 if used_exchange else 0
        if side == "back":
            # stake = 1; gano (cuota-1) menos comision, o pierdo 1
            if won:
                self.wins += 1
                self.pnl += (odds - 1.0) * (1.0 - commission)
            else:
                self.pnl -= 1.0
        else:  # lay: fijamos liability = 1 => backer stake = 1/(cuota-1)
            backer_stake = 1.0 / (odds - 1.0)
            if won:  # la seleccion NO ocurrio: el lay gana el backer stake
                self.wins += 1
                self.pnl += backer_stake * (1.0 - commission)
            else:    # la seleccion ocurrio: el lay pierde el liability (=1)
                self.pnl -= 1.0

    @property
    def roi(self):
        return self.pnl / self.n if self.n else 0.0

    @property
    def winrate(self):
        return self.wins / self.n if self.n else 0.0

    @property
    def avg_odds(self):
        return self.odds_sum / self.n if self.n else 0.0


def bucket_of(odds, side):
    """Bin de cuota de ancho 0.25 (back) o por liability (lay)."""
    lo = int(odds / 0.25) * 0.25
    return (round(lo, 2), round(lo + 0.25, 2))


def load(league, season, data_dir):
    for name in (f"{league}_{season}.csv", f"{league}-{season}.csv"):
        p = os.path.join(data_dir, name)
        if os.path.exists(p):
            with open(p, "r", encoding="latin-1") as fh:
                return list(csv.DictReader(io.StringIO(fh.read())))
    return []


def run(args):
    leagues = args.leagues.split(",")
    seasons = args.seasons.split(",")
    defs = market_defs()
    commission = args.commission

    # stats[(market, bucket)] = Stat
    stats: dict[tuple, Stat] = {}

    for league in leagues:
        for season in seasons:
            for row in load(league, season, args.data_dir):
                if (row.get("FTR") or "").strip() not in ("H", "D", "A"):
                    continue
                for name, desc, side, odds_cols, win_fn in defs:
                    won = win_fn(row)
                    if won is None:
                        continue
                    # cuota de la SELECCION (exchange preferida)
                    exch_odds = _f(row, odds_cols[0])
                    odds = exch_odds if exch_odds else _f(row, *odds_cols)
                    if odds is None:
                        continue
                    if side == "lay" and odds > args.max_lay_odds:
                        continue  # liability demasiado alto
                    b = bucket_of(odds, side)
                    key = (name, desc, side, b)
                    stats.setdefault(key, Stat()).add(
                        odds, won, commission, used_exchange=bool(exch_odds),
                        side=side)

    rows = []
    for (name, desc, side, b), s in stats.items():
        if s.n < args.min_n:
            continue
        rows.append((s.roi, name, desc, side, b, s))
    rows.sort(reverse=True)

    print(f"\nEXPLORACION DE ESTRATEGIAS | comision {commission:.0%} | "
          f"muestra min {args.min_n} | cuotas: Betfair Exchange (fallback Avg)\n")
    hdr = (f"{'Estrategia':<16} {'Rango cuota':<13} {'N':>5} {'%Exch':>6} "
           f"{'Acierto':>8} {'CuotaPr':>8} {'ROI':>8}")
    print(hdr)
    print("-" * len(hdr))
    shown = 0
    for roi, name, desc, side, b, s in rows:
        flag = ""
        if roi > 0:
            flag = "  <== positivo"
        pct_exch = s.exch / s.n * 100
        print(f"{name:<16} {b[0]:>5.2f}-{b[1]:<6.2f} {s.n:>5} {pct_exch:>5.0f}% "
              f"{s.winrate*100:>7.1f}% {s.avg_odds:>8.2f} {roi*100:>+7.1f}%{flag}")
        shown += 1
        if shown >= args.top:
            break
    print(f"\n(mostrando top {shown} buckets por ROI; total con muestra suficiente: {len(rows)})")
    print("RECORDATORIO: ROI positivo aqui = HIPOTESIS. Sospecha de sobreajuste si "
          "la muestra es chica o no hay logica economica.")


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--data-dir", required=True)
    p.add_argument("--commission", type=float, default=0.05)
    p.add_argument("--min-n", type=int, default=150)
    p.add_argument("--max-lay-odds", type=float, default=6.0)
    p.add_argument("--top", type=int, default=30)
    p.add_argument("--leagues", default=",".join(LEAGUES))
    p.add_argument("--seasons", default=",".join(SEASONS))
    run(p.parse_args())


if __name__ == "__main__":
    main()
