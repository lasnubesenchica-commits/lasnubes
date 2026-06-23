#!/usr/bin/env python3
"""Convierte el CSV de capturas en el JSON que consume el dashboard betfair.html.

Calcula agregados: P&L total, ROI, win rate, desglose jornada a jornada y curva de
equity. Reutilizable por el generador de demo.

    python3 -m betfair.capture.export_json --in betfair/data/masters_capture.csv \
        --out betfair/data/masters.json
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import os

from ..bot.strategy import LayFavConfig


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def parse_csv_rows(path: str) -> list[dict]:
    """Lee el CSV de capturas y lo normaliza a dicts de captura."""
    if not os.path.exists(path):
        return []
    with open(path, newline="") as fh:
        raw = list(csv.DictReader(fh))
    out = []
    for r in raw:
        out.append({
            "captured_at": r.get("captured_at", ""),
            "competition": r.get("competition", ""),
            "event": r.get("event", ""),
            "start_time": r.get("start_time", ""),
            "is_masters": int(r.get("is_masters") or 0),
            "fav_name": r.get("fav_name", ""),
            "fav_lay_price": _num(r.get("fav_lay_price")),
            "dog_name": r.get("dog_name", ""),
            "dog_lay_price": _num(r.get("dog_lay_price")),
            "in_band": int(r.get("in_band") or 0),
            "fav_won": (int(r["fav_won"]) if r.get("fav_won") not in (None, "") else None),
            "pnl_units": _num(r.get("pnl_units")),
        })
    return out


def summarize(captures: list[dict], cfg: LayFavConfig) -> dict:
    """Agregados para el dashboard a partir de las apuestas liquidadas."""
    bets = [c for c in captures
            if c["is_masters"] and c["in_band"] and c["pnl_units"] is not None]
    bets.sort(key=lambda c: (c["start_time"] or ""))

    by_day: dict[str, dict] = {}
    for b in bets:
        day = (b["start_time"] or "")[:10]
        d = by_day.setdefault(day, {"date": day, "bets": 0, "pnl": 0.0})
        d["bets"] += 1
        d["pnl"] = round(d["pnl"] + b["pnl_units"], 2)

    days = sorted(by_day.values(), key=lambda d: d["date"])
    cum = 0.0
    equity = []
    for d in days:
        cum = round(cum + d["pnl"], 2)
        d["cum_pnl"] = cum
        equity.append({"date": d["date"], "cum_pnl": cum})

    nbets = len(bets)
    total_pnl = round(sum(b["pnl_units"] for b in bets), 2)
    # ROI sobre el CAPITAL EN RIESGO (liability), igual que el backtest:
    # liability de cada lay = backer_stake * (cuota - 1).
    total_liability = sum(cfg.backer_stake * ((b["fav_lay_price"] or 1) - 1)
                          for b in bets)
    wins = sum(1 for b in bets if b["fav_won"] == 0)    # lay gana si el fav PIERDE
    losses = sum(1 for b in bets if b["fav_won"] == 1)
    return {
        "settled_bets": nbets,
        "pnl": total_pnl,
        "liability": round(total_liability, 2),
        "roi": (round(total_pnl / total_liability * 100, 1) if total_liability else None),
        "wins": wins, "losses": losses,
        "win_rate": (round(wins / nbets * 100, 1) if nbets else None),
        "by_day": days, "equity": equity,
    }


def assemble(captures: list[dict], cfg: LayFavConfig, *, source: str = "live") -> dict:
    return {
        "updated_at": dt.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "source": source,
        "config": {"min_fav_odds": cfg.min_fav_odds, "max_fav_odds": cfg.max_fav_odds,
                   "backer_stake": cfg.backer_stake, "commission": cfg.commission},
        "captures": captures,
        "summary": summarize(captures, cfg),
    }


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--in", dest="inp", default="betfair/data/masters_capture.csv")
    p.add_argument("--out", default="betfair/data/masters.json")
    args = p.parse_args()
    cfg = LayFavConfig()
    data = assemble(parse_csv_rows(args.inp), cfg, source="live")
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, "w") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)
    s = data["summary"]
    print(f"escrito {args.out}: {len(data['captures'])} capturas, "
          f"{s['settled_bets']} apuestas liquidadas, P&L {s['pnl']}")


if __name__ == "__main__":
    main()
