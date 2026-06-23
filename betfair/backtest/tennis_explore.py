#!/usr/bin/env python3
"""
Explorador de estrategias PASIVAS (pre-partido) de tennis sobre datos de
tennis-data.co.uk (ATP). Usa cuotas de cierre de Pinnacle (PSW/PSL), las mas
"sharp" del mercado. Filtra por serie (Grand Slam / Masters 1000).

Las cuotas vienen por ganador (W) / perdedor (L); reconstruimos favorito vs
underdog por la cuota (favorito = menor cuota). Probamos back/lay a cada rol por
rango de cuota del favorito, con ROI normalizado a 1u de riesgo y comision.

Incluye ROI por temporada para separar señal de ruido (una estrategia real debe
ganar los 3 años, no solo uno).

CAVEAT: Pinnacle es bookie (margen ~2-3%). Para BACK, el margen juega EN CONTRA,
asi que un ROI de back positivo es señal real (en Betfair, con cuotas algo mas
altas, seria similar o levemente mejor). Para LAY, el margen infla un poco; restar
~1-1.5% para aproximar Betfair.

Uso:
    python3 tennis_explore.py --data-dir ./tennis --series "Grand Slam,Masters 1000"
"""
from __future__ import annotations
import argparse, csv, io, os
from dataclasses import dataclass, field

SEASONS = ["2023", "2024", "2025"]


def fnum(r, c):
    v = (r.get(c) or "").strip()
    try:
        x = float(v); return x if x > 1.0 else None
    except ValueError:
        return None


@dataclass
class Stat:
    n: int = 0
    wins: int = 0
    pnl: float = 0.0
    odds_sum: float = 0.0
    per_season: dict = field(default_factory=lambda: {s: [0, 0.0] for s in SEASONS})

    def add(self, season, odds, won, commission, side):
        self.n += 1; self.odds_sum += odds
        if side == "back":
            p = (odds - 1.0) * (1.0 - commission) if won else -1.0
        else:  # lay, liability=1 => backer stake = 1/(odds-1)
            p = (1.0 / (odds - 1.0)) * (1.0 - commission) if won else -1.0
        if won:
            self.wins += 1
        self.pnl += p
        if season in self.per_season:
            self.per_season[season][0] += 1
            self.per_season[season][1] += p

    @property
    def roi(self): return self.pnl / self.n if self.n else 0.0
    @property
    def winrate(self): return self.wins / self.n if self.n else 0.0
    @property
    def avg_odds(self): return self.odds_sum / self.n if self.n else 0.0
    def season_roi(self, s):
        n, pnl = self.per_season[s]
        return (pnl / n if n else 0.0), n


def load(data_dir):
    rows = []
    for s in SEASONS:
        p = os.path.join(data_dir, f"{s}.csv")
        if os.path.exists(p):
            for r in csv.DictReader(io.StringIO(open(p, encoding="latin-1").read())):
                r["_season"] = s
                rows.append(r)
    return rows


def run(args):
    series_ok = {x.strip() for x in args.series.split(",")} if args.series else None
    commission = args.commission
    rows = load(args.data_dir)

    # buckets de cuota del FAVORITO
    edges = [1.0, 1.2, 1.4, 1.6, 1.8, 2.0]
    def bucket(o):
        for i in range(len(edges) - 1):
            if edges[i] <= o < edges[i + 1]:
                return (edges[i], edges[i + 1])
        return None  # favorito con cuota >=2 (pick'em) lo ignoramos

    # strategies: (id, side, role)  role = fav|dog
    strategies = [("BACK_FAV", "back", "fav"), ("LAY_FAV", "lay", "fav"),
                  ("BACK_DOG", "back", "dog"), ("LAY_DOG", "lay", "dog")]
    stats: dict = {}

    used = 0
    for r in rows:
        if series_ok and (r.get("Series") or "").strip() not in series_ok:
            continue
        psw, psl = fnum(r, "PSW"), fnum(r, "PSL")
        if psw is None or psl is None or psw == psl:
            continue
        fav_won = psw < psl
        fav_odds = min(psw, psl); dog_odds = max(psw, psl)
        b = bucket(fav_odds)
        if b is None:
            continue
        used += 1
        for name, side, role in strategies:
            if role == "fav":
                odds = fav_odds; won = fav_won if side == "back" else (not fav_won)
            else:
                odds = dog_odds; won = (not fav_won) if side == "back" else fav_won
            if side == "lay" and odds > args.max_lay_odds:
                continue
            key = (name, b)
            stats.setdefault(key, Stat()).add(r["_season"], odds, won, commission, side)

    rows_out = []
    for (name, b), s in stats.items():
        if s.n < args.min_n:
            continue
        rows_out.append((s.roi, name, b, s))
    rows_out.sort(reverse=True)

    title = args.series if args.series else "TODAS las series"
    print(f"\nTENNIS pasivo | series: {title} | comision {commission:.0%} | "
          f"cuotas Pinnacle | partidos usados: {used}\n")
    hdr = (f"{'Estrategia':<10} {'CuotaFav':<11} {'N':>5} {'Win%':>6} {'CuotaPr':>8} "
           f"{'2023':>8} {'2024':>8} {'2025':>8} {'TOTAL':>8}")
    print(hdr); print("-" * len(hdr))
    for roi, name, b, s in rows_out:
        def sc(season):
            r_, n_ = s.season_roi(season)
            return f"{r_*100:+5.1f}%" if n_ >= 25 else "  --  "
        consistent = all(s.season_roi(x)[0] > 0 for x in SEASONS if s.season_roi(x)[1] >= 25)
        flag = "  <== consistente" if (roi > 0 and consistent) else ""
        print(f"{name:<10} {b[0]:>4.1f}-{b[1]:<6.1f} {s.n:>5} {s.winrate*100:>5.1f}% "
              f"{s.avg_odds:>8.2f} {sc('2023'):>8} {sc('2024'):>8} {sc('2025'):>8} "
              f"{roi*100:>+7.1f}%{flag}")
    print("\nCONSISTENTE = ROI positivo en cada temporada con muestra suficiente (>=25). "
          "Esos son los unicos candidatos creibles; el resto es probable ruido.")


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--data-dir", required=True)
    p.add_argument("--series", default="Grand Slam,Masters 1000")
    p.add_argument("--commission", type=float, default=0.05)
    p.add_argument("--min-n", type=int, default=80)
    p.add_argument("--max-lay-odds", type=float, default=8.0)
    run(p.parse_args())


if __name__ == "__main__":
    main()
