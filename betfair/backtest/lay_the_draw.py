#!/usr/bin/env python3
"""
Backtest de la estrategia "Lay the Draw" (puro: dejar correr hasta el final)
usando datos historicos reales de football-data.co.uk.

Estrategia
----------
Para cada partido cuya CUOTA DE CIERRE DEL EMPATE caiga dentro del rango
[min_draw_odds, max_draw_odds], colocamos un LAY al empate (apostamos a que NO
habra empate) y dejamos correr la apuesta hasta el resultado final.

- Si el partido NO termina en empate  -> GANAMOS el backer's stake (menos comision).
- Si el partido termina en empate      -> PERDEMOS el liability = stake * (cuota - 1).

Unidad de apuesta: 1 unidad de "backer's stake" por seleccion.
  * Ganancia por acierto:  +(1 - comision) unidades
  * Perdida por empate:    -(cuota - 1)   unidades   (este es el capital real en riesgo)

Asi, P&L se reporta en "unidades" donde 1 unidad = el stake que laias por apuesta.
OJO: el capital real arriesgado por apuesta es el liability (cuota-1), ~2.0-2.7 unidades.

Fuente de datos
---------------
football-data.co.uk publica un CSV por liga y temporada con resultado final y
cuotas (incluida la cuota del empate de varios bookies). Usamos la cuota de
CIERRE (closing) como mejor proxy de lo que ofreceria Betfair cerca del kickoff.

Uso
---
    python3 lay_the_draw.py
    python3 lay_the_draw.py --min 3.0 --max 3.7 --commission 0.05
    python3 lay_the_draw.py --leagues E0,SP1,D1,F1,I1 --seasons 2324,2425,2526
    python3 lay_the_draw.py --data-dir ./csv   # usa CSVs locales en vez de descargar

Si la red esta bloqueada (allowlist del entorno), descarga los CSV a mano desde
https://www.football-data.co.uk/  y apunta --data-dir a la carpeta.
"""

from __future__ import annotations

import argparse
import csv
import io
import os
import sys
import urllib.request
from dataclasses import dataclass, field

# --- Configuracion de ligas y temporadas -----------------------------------

# Codigos de division de football-data.co.uk
LEAGUE_NAMES = {
    "E0": "Premier League (ENG)",
    "E1": "Championship (ENG)",
    "SP1": "La Liga (ESP)",
    "D1": "Bundesliga (GER)",
    "F1": "Ligue 1 (FRA)",
    "I1": "Serie A (ITA)",
    "N1": "Eredivisie (NED)",
    "P1": "Primeira Liga (POR)",
}

# Codigos de temporada: "2425" = temporada 2024-25
SEASON_NAMES = {
    "2223": "2022-23",
    "2324": "2023-24",
    "2425": "2024-25",
    "2526": "2025-26",
}

BASE_URL = "https://www.football-data.co.uk/mmz4281/{season}/{league}.csv"

# Prioridad de columnas para la cuota del empate.
# Preferimos cuotas de CIERRE (closing, sufijo C) y promedios de mercado, que son
# el mejor proxy de Betfair. Caemos a cuotas pre-partido si no hay closing.
DRAW_ODDS_COLUMNS = [
    "AvgCD",   # promedio de mercado, cierre
    "PSCD",    # Pinnacle, cierre (sharp)
    "B365CD",  # Bet365, cierre
    "MaxCD",   # mejor cuota de mercado, cierre
    "AvgD",    # promedio de mercado, pre-partido
    "PSD",     # Pinnacle, pre-partido
    "B365D",   # Bet365, pre-partido
    "BbAvD",   # promedio Betbrain (temporadas viejas)
]


# --- Modelo de resultados ----------------------------------------------------

@dataclass
class Bucket:
    """Acumulador de resultados para una liga/temporada (o total)."""
    bets: int = 0
    wins: int = 0          # partidos NO empatados (ganamos el lay)
    draws: int = 0         # partidos empatados (perdemos el liability)
    pnl: float = 0.0       # ganancia/perdida en unidades de backer stake
    odds_sum: float = 0.0  # para promediar la cuota del empate apostada

    def add(self, drew: bool, draw_odds: float, commission: float) -> None:
        self.bets += 1
        self.odds_sum += draw_odds
        if drew:
            self.draws += 1
            self.pnl -= (draw_odds - 1.0)        # perdemos el liability
        else:
            self.wins += 1
            self.pnl += (1.0 - commission)        # ganamos el stake menos comision

    @property
    def draw_rate(self) -> float:
        return self.draws / self.bets if self.bets else 0.0

    @property
    def avg_odds(self) -> float:
        return self.odds_sum / self.bets if self.bets else 0.0

    @property
    def roi(self) -> float:
        # ROI sobre el total apostado (1 unidad por apuesta => pnl / bets)
        return self.pnl / self.bets if self.bets else 0.0


# --- Logica de la estrategia (reutilizable por el bot) -----------------------

def pick_draw_odds(row: dict) -> float | None:
    """Devuelve la cuota del empate segun la prioridad de columnas, o None."""
    for col in DRAW_ODDS_COLUMNS:
        val = row.get(col, "").strip()
        if val:
            try:
                odds = float(val)
                if odds > 1.0:
                    return odds
            except ValueError:
                continue
    return None


def is_lay_candidate(draw_odds: float, min_odds: float, max_odds: float) -> bool:
    """Regla de seleccion: laiamos el empate si su cuota esta en el rango."""
    return min_odds <= draw_odds <= max_odds


# --- Carga de datos ----------------------------------------------------------

def load_csv_text(league: str, season: str, data_dir: str | None) -> str | None:
    """Lee el CSV desde --data-dir (<league>_<season>.csv o <league>.csv) o lo descarga."""
    if data_dir:
        for name in (f"{league}_{season}.csv", f"{season}_{league}.csv", f"{league}.csv"):
            path = os.path.join(data_dir, name)
            if os.path.exists(path):
                with open(path, "r", encoding="latin-1") as fh:
                    return fh.read()
        return None
    url = BASE_URL.format(season=season, league=league)
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            return resp.read().decode("latin-1")
    except Exception as exc:  # noqa: BLE001
        print(f"  ! no se pudo obtener {league} {season}: {exc}", file=sys.stderr)
        return None


def parse_rows(text: str) -> list[dict]:
    reader = csv.DictReader(io.StringIO(text))
    rows = []
    for row in reader:
        # FTR = Full Time Result: H / D / A. Filas validas tienen resultado.
        if row.get("FTR", "").strip() in ("H", "D", "A"):
            rows.append(row)
    return rows


# --- Backtest ----------------------------------------------------------------

def run_backtest(args) -> None:
    leagues = args.leagues.split(",")
    seasons = args.seasons.split(",")

    grand = Bucket()
    per_league: dict[str, Bucket] = {lg: Bucket() for lg in leagues}
    cells: dict[tuple[str, str], Bucket] = {}

    print(f"\nEstrategia: LAY THE DRAW (dejar correr) | cuota empate "
          f"[{args.min:.2f}, {args.max:.2f}] | comision {args.commission:.0%}\n")

    for league in leagues:
        for season in seasons:
            text = load_csv_text(league, season, args.data_dir)
            if not text:
                continue
            cell = Bucket()
            for row in parse_rows(text):
                odds = pick_draw_odds(row)
                if odds is None or not is_lay_candidate(odds, args.min, args.max):
                    continue
                drew = row["FTR"].strip() == "D"
                cell.add(drew, odds, args.commission)
                per_league[league].add(drew, odds, args.commission)
                grand.add(drew, odds, args.commission)
            cells[(league, season)] = cell

    _print_report(leagues, seasons, cells, per_league, grand, args)


def _fmt_row(label: str, b: Bucket) -> str:
    if not b.bets:
        return f"{label:<26} {'-':>6} {'-':>6} {'-':>9} {'-':>9} {'-':>9}"
    return (f"{label:<26} {b.bets:>6} {b.draw_rate*100:>5.1f}% "
            f"{b.avg_odds:>8.2f} {b.pnl:>+8.2f}u {b.roi*100:>+7.1f}%")


def _print_report(leagues, seasons, cells, per_league, grand, args) -> None:
    header = (f"{'Liga / Temporada':<26} {'Aptas':>6} {'%Emp':>6} "
              f"{'CuotaPr':>9} {'P&L':>9} {'ROI':>8}")
    print(header)
    print("-" * len(header))

    for league in leagues:
        any_data = any(cells.get((league, s), Bucket()).bets for s in seasons)
        if not any_data:
            print(f"{LEAGUE_NAMES.get(league, league):<26} (sin datos)")
            continue
        print(f"\n{LEAGUE_NAMES.get(league, league)}")
        for season in seasons:
            cell = cells.get((league, season), Bucket())
            print("  " + _fmt_row(SEASON_NAMES.get(season, season), cell))
        print("  " + _fmt_row("TOTAL liga", per_league[league]))

    print("\n" + "=" * len(header))
    print(_fmt_row("TOTAL GLOBAL", grand))
    print("=" * len(header))

    # Interpretacion rapida
    if grand.bets:
        # Probabilidad de empate de breakeven para la cuota promedio apostada
        o = grand.avg_odds
        c = args.commission
        p_be = (1 - c) / ((o - 1) + (1 - c))
        print(f"\nLectura:")
        print(f"  - Apuestas que pasaron el filtro: {grand.bets}")
        print(f"  - Empates (perdidas): {grand.draws} ({grand.draw_rate*100:.1f}%)")
        print(f"  - Punto de equilibrio: necesitas <{p_be*100:.1f}% de empates "
              f"en el filtro para ganar (a cuota promedio {o:.2f}).")
        verdict = "RENTABLE" if grand.pnl > 0 else "PERDEDORA"
        print(f"  - Veredicto historico: {verdict} "
              f"(P&L {grand.pnl:+.2f}u, ROI {grand.roi*100:+.1f}% sobre stake).")
        print(f"  - Recuerda: el capital real en riesgo por apuesta es el liability "
              f"(~{o-1:.2f}u), no 1u.")
    else:
        print("\nNo se proceso ninguna apuesta. Revisa --data-dir o el acceso de red.")


def main() -> None:
    p = argparse.ArgumentParser(description="Backtest Lay the Draw (football-data.co.uk)")
    p.add_argument("--min", type=float, default=3.0, help="cuota minima del empate")
    p.add_argument("--max", type=float, default=3.7, help="cuota maxima del empate")
    p.add_argument("--commission", type=float, default=0.05, help="comision Betfair (0.05 = 5%)")
    p.add_argument("--leagues", default="E0,SP1,D1,F1,I1,N1",
                   help="codigos de liga separados por coma")
    p.add_argument("--seasons", default="2223,2324,2425,2526",
                   help="temporadas separadas por coma (ej. 2425 = 2024-25)")
    p.add_argument("--data-dir", default=None,
                   help="carpeta con CSVs locales; si se omite, descarga de la web")
    args = p.parse_args()
    run_backtest(args)


if __name__ == "__main__":
    main()
