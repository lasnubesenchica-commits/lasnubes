#!/usr/bin/env python3
"""Convierte el CSV de capturas (genérico) en el JSON que consume betfair.html.

Esquema de captura (común a tennis y fútbol):
  start_time, competition, event, selection, lay_price, detail, in_band,
  won (1=el lay ganó), pnl_units.

Calcula agregados: P&L, ROI sobre el capital en riesgo (liability), win rate,
desglose jornada a jornada y curva de equity. Reutilizable por los generadores
de demo.

    python3 -m betfair.capture.export_json --sport football \
        --in betfair/data/football_capture.csv --out betfair/data/football.json
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import os

STRATEGY_META = {
    "tennis": {"label": "Lay al favorito · Masters 1000", "band": [1.2, 1.5]},
    "football": {"label": "Lay the Draw", "band": [3.0, 3.7]},
}
BACKER_STAKE = 2.0
COMMISSION = 0.05


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def parse_csv_rows(path: str) -> list[dict]:
    if not os.path.exists(path):
        return []
    with open(path, newline="") as fh:
        raw = list(csv.DictReader(fh))
    out = []
    for r in raw:
        out.append({
            "start_time": r.get("start_time", ""),
            "competition": r.get("competition", ""),
            "event": r.get("event", ""),
            "selection": r.get("selection", ""),
            "lay_price": _num(r.get("lay_price")),
            "detail": r.get("detail", ""),
            "in_band": int(r.get("in_band") or 0),
            "won": (int(r["won"]) if r.get("won") not in (None, "") else None),
            "pnl_units": _num(r.get("pnl_units")),
        })
    return out


def summarize(captures: list[dict], backer_stake: float = BACKER_STAKE) -> dict:
    bets = [c for c in captures if c["in_band"] and c["pnl_units"] is not None]
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
    total_liability = sum(backer_stake * ((b["lay_price"] or 1) - 1) for b in bets)
    wins = sum(1 for b in bets if b["won"] == 1)
    losses = sum(1 for b in bets if b["won"] == 0)
    return {
        "settled_bets": nbets,
        "pnl": total_pnl,
        "liability": round(total_liability, 2),
        "roi": (round(total_pnl / total_liability * 100, 1) if total_liability else None),
        "wins": wins, "losses": losses,
        "win_rate": (round(wins / nbets * 100, 1) if nbets else None),
        "by_day": days, "equity": equity,
    }


def assemble(captures, sport, *, source="live", note=None) -> dict:
    meta = STRATEGY_META[sport]
    data = {
        "updated_at": dt.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "sport": sport, "source": source,
        "strategy": meta["label"],
        "config": {"band": meta["band"], "backer_stake": BACKER_STAKE,
                   "commission": COMMISSION},
        "captures": captures,
        "summary": summarize(captures),
    }
    if note:
        data["demo_note"] = note
    return data


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--sport", choices=["tennis", "football"], required=True)
    p.add_argument("--in", dest="inp", required=True)
    p.add_argument("--out", required=True)
    args = p.parse_args()
    data = assemble(parse_csv_rows(args.inp), args.sport, source="live")
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, "w") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)
    s = data["summary"]
    print(f"[{args.sport}] {args.out}: {len(data['captures'])} capturas, "
          f"{s['settled_bets']} apuestas, P&L {s['pnl']}, ROI {s['roi']}")


if __name__ == "__main__":
    main()
