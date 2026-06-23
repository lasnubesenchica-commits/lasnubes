#!/usr/bin/env python3
"""Genera un dataset DEMO para el dashboard a partir del backtest de tennis.

Toma los Masters 1000 de una temporada (datos tennis-data.co.uk, cuotas Pinnacle),
aplica la regla lay-favorito-1.2-1.5 y produce el MISMO JSON que consume el
dashboard, marcado como source="demo". Sirve para ver el dashboard funcionando
antes del primer Masters real, claramente etiquetado como demostración.

    python3 -m betfair.capture.make_demo --tennis-dir <dir> --season 2025 \
        --out betfair/data/masters_demo.json
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
import json
import os

from ..bot.strategy import LayFavConfig, settle_lay_pnl
from .export_json import assemble


def _parse_date(s: str) -> str:
    s = (s or "").strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d/%m/%y"):
        try:
            return dt.datetime.strptime(s[:10], fmt).strftime("%Y-%m-%dT12:00:00Z")
        except ValueError:
            continue
    return ""


def _num(r, c):
    v = (r.get(c) or "").strip()
    try:
        x = float(v)
        return x if x > 1 else None
    except ValueError:
        return None


def build_demo_captures(path: str, cfg: LayFavConfig) -> list[dict]:
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
        fav_name, dog_name = (winner, loser) if fav_won else (loser, winner)
        in_band = cfg.min_fav_odds <= fav_odds <= cfg.max_fav_odds
        pnl = (settle_lay_pnl(fav_odds, cfg.backer_stake, fav_won, cfg.commission)
               if in_band else None)
        caps.append({
            "captured_at": _parse_date(r.get("Date", "")),
            "competition": r.get("Tournament", "Masters 1000"),
            "event": f"{fav_name} v {dog_name}",
            "start_time": _parse_date(r.get("Date", "")),
            "is_masters": 1,
            "fav_name": fav_name, "fav_lay_price": round(fav_odds, 2),
            "dog_name": dog_name, "dog_lay_price": round(dog_odds, 2),
            "in_band": int(in_band),
            "fav_won": int(fav_won), "pnl_units": pnl,
        })
    return caps


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--tennis-dir", required=True)
    p.add_argument("--season", default="2025")
    p.add_argument("--out", default="betfair/data/masters_demo.json")
    args = p.parse_args()
    cfg = LayFavConfig()
    caps = build_demo_captures(os.path.join(args.tennis_dir, f"{args.season}.csv"), cfg)
    data = assemble(caps, cfg, source="demo")
    data["demo_note"] = (f"DEMO: backtest Masters 1000 {args.season} con cuotas de "
                         f"cierre de Pinnacle (no Betfair exchange). Solo ilustrativo.")
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, "w") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)
    s = data["summary"]
    print(f"demo {args.out}: {len(caps)} partidos Masters, {s['settled_bets']} "
          f"apuestas in-band, P&L {s['pnl']}u, ROI {s['roi']}%")


if __name__ == "__main__":
    main()
