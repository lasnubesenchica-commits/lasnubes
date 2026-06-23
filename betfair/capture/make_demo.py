#!/usr/bin/env python3
"""Genera datasets DEMO para el dashboard a partir de los backtests.

  - tennis  : Masters 1000 de una temporada (tennis-data.co.uk, cuotas Pinnacle).
  - football: Lay the Draw en ligas de una temporada (football-data.co.uk, cuotas
              de cierre del exchange BFE si están, si no promedio de mercado).

Producen el MISMO esquema genérico que consume el dashboard, marcado source="demo"
y claramente etiquetado como ilustrativo (no es dinero real).

    python3 -m betfair.capture.make_demo --sport tennis   --src <dir> --season 2025
    python3 -m betfair.capture.make_demo --sport football --src <dir> --season 2425
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
import json
import os

from ..bot.strategy import LayDrawConfig, LayFavConfig, settle_lay_pnl
from .export_json import BACKER_STAKE, COMMISSION, assemble

TENNIS_LEAGUES = None
FOOTBALL_LEAGUES = ["E0", "SP1", "D1", "F1", "I1"]


def _date(s, fmts):
    s = (s or "").strip()
    for fmt in fmts:
        try:
            return dt.datetime.strptime(s[:10], fmt).strftime("%Y-%m-%dT12:00:00Z")
        except ValueError:
            continue
    return ""


def _num(r, *cols):
    for c in cols:
        v = (r.get(c) or "").strip()
        if v:
            try:
                x = float(v)
                if x > 1:
                    return x
            except ValueError:
                pass
    return None


def tennis_captures(path, cfg: LayFavConfig):
    rows = list(csv.DictReader(io.StringIO(open(path, encoding="latin-1").read())))
    caps = []
    for r in rows:
        if (r.get("Series") or "").strip() != "Masters 1000":
            continue
        psw, psl = _num(r, "PSW"), _num(r, "PSL")
        if not psw or not psl or psw == psl:
            continue
        fav_won = psw < psl
        fav_odds, dog_odds = min(psw, psl), max(psw, psl)
        winner, loser = r.get("Winner", "").strip(), r.get("Loser", "").strip()
        fav, dog = (winner, loser) if fav_won else (loser, winner)
        in_band = cfg.min_fav_odds <= fav_odds <= cfg.max_fav_odds
        when = _date(r.get("Date"), ("%Y-%m-%d", "%d/%m/%Y"))
        caps.append({
            "start_time": when, "competition": r.get("Tournament", "Masters 1000"),
            "event": f"{fav} v {dog}", "selection": fav,
            "lay_price": round(fav_odds, 2), "detail": f"vs {dog} @ {round(dog_odds,2)}",
            "in_band": int(in_band),
            "won": int(not fav_won),
            "pnl_units": (settle_lay_pnl(fav_odds, BACKER_STAKE, fav_won, COMMISSION)
                          if in_band else None)})
    return caps


def football_captures(src_dir, season, cfg: LayDrawConfig):
    caps = []
    for lg in FOOTBALL_LEAGUES:
        path = os.path.join(src_dir, f"{lg}_{season}.csv")
        if not os.path.exists(path):
            continue
        for r in csv.DictReader(io.StringIO(open(path, encoding="latin-1").read())):
            if (r.get("FTR") or "").strip() not in ("H", "D", "A"):
                continue
            draw = _num(r, "BFECD", "AvgCD", "B365CD")
            if draw is None:
                continue
            drew = r["FTR"].strip() == "D"
            in_band = cfg.min_draw_odds <= draw <= cfg.max_draw_odds
            home = _num(r, "BFECH", "AvgCH", "B365CH")
            away = _num(r, "BFECA", "AvgCA", "B365CA")
            when = _date(r.get("Date"), ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d"))
            caps.append({
                "start_time": when, "competition": r.get("Div", lg),
                "event": f"{r.get('HomeTeam','')} v {r.get('AwayTeam','')}",
                "selection": "Empate", "lay_price": round(draw, 2),
                "detail": f"casa {home} · fuera {away}",
                "in_band": int(in_band),
                "won": int(not drew),
                "pnl_units": (settle_lay_pnl(draw, BACKER_STAKE, drew, COMMISSION)
                              if in_band else None)})
    return caps


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--sport", choices=["tennis", "football"], required=True)
    p.add_argument("--src", required=True, help="carpeta con los CSV del backtest")
    p.add_argument("--season", required=True, help="tennis: 2025 | football: 2425")
    p.add_argument("--out", required=True)
    args = p.parse_args()

    if args.sport == "tennis":
        caps = tennis_captures(os.path.join(args.src, f"{args.season}.csv"), LayFavConfig())
        note = (f"DEMO: backtest Masters 1000 {args.season} (cuotas Pinnacle, no "
                f"Betfair exchange). Solo ilustrativo.")
    else:
        caps = football_captures(args.src, args.season, LayDrawConfig())
        note = (f"DEMO: Lay the Draw en ligas top {args.season} (cuotas de cierre). "
                f"Estrategia de referencia, no necesariamente rentable. Ilustrativo.")

    data = assemble(caps, args.sport, source="demo", note=note)
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, "w") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)
    s = data["summary"]
    print(f"[{args.sport}] demo {args.out}: {len(caps)} partidos, "
          f"{s['settled_bets']} apuestas, P&L {s['pnl']}u, ROI {s['roi']}%")


if __name__ == "__main__":
    main()
