#!/usr/bin/env python3
"""Capturador genérico de cuotas del Betfair Exchange (tennis y fútbol).

Modos:
  - record : busca mercados Match Odds que arrancan pronto, identifica la
             selección que la estrategia laiaría (favorito en tennis, empate en
             fútbol) y guarda su mejor cuota de LAY.
  - settle : re-lee los mercados ya jugados, detecta el ganador y calcula el P&L
             del paper trading (lay gana si la selección laeada NO ganó).

Sólo LEE de Betfair. NO coloca apuestas. Un CSV por deporte.

    python3 -m betfair.capture.capture record  --sport football --within-min 20
    python3 -m betfair.capture.capture settle  --sport football
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import logging
import os

from ..bot import config as C
from ..bot import parsing
from ..bot.betfair_client import BetfairClient
from ..bot.strategy import (LayDrawConfig, LayFavConfig, is_masters_competition,
                            make_pick, matches_keywords, settle_lay_pnl)

log = logging.getLogger("capture")

FIELDS = ["captured_at", "market_id", "sport", "competition", "event",
          "start_time", "selection", "selection_id", "lay_price", "detail",
          "in_band", "won", "pnl_units"]

SPORT_EVENT_TYPE = {"tennis": C.TENNIS_EVENT_TYPE_ID, "football": C.SOCCER_EVENT_TYPE_ID}
DEFAULT_OUT = {"tennis": "betfair/data/masters_capture.csv",
               "football": "betfair/data/football_capture.csv"}


def _iso(t: dt.datetime) -> str:
    return t.strftime("%Y-%m-%dT%H:%M:%SZ")


def _read(path):
    if not os.path.exists(path):
        return []
    with open(path, newline="") as fh:
        return list(csv.DictReader(fh))


def _write(path, rows):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(rows)


def _keep_competition(sport, comp, tcfg, fcfg) -> bool:
    if sport == "tennis":
        return is_masters_competition(comp, tcfg)          # sólo Masters 1000
    return matches_keywords(comp, fcfg.competition_keywords)  # Mundial + ligas


def cmd_record(client, sport, tcfg, fcfg, args):
    now = dt.datetime.utcnow()
    markets = client.list_market_catalogue(
        event_type_id=SPORT_EVENT_TYPE[sport], market_type_codes=["MATCH_ODDS"],
        from_iso=_iso(now), to_iso=_iso(now + dt.timedelta(minutes=args.within_min)),
        max_results=args.max_markets)
    if not markets:
        log.info("[%s] sin mercados arrancando en %d min", sport, args.within_min)
        return

    books = {b["marketId"]: b for b in
             client.list_market_book([m["marketId"] for m in markets])}
    rows = _read(args.out)
    seen = {r["market_id"] for r in rows}
    added = 0
    for m in markets:
        mid = m["marketId"]
        if mid in seen:
            continue
        comp = (m.get("competition") or {}).get("name", "")
        if not _keep_competition(sport, comp, tcfg, fcfg):
            continue
        names = parsing.runner_names(m)
        prices = parsing.best_lay_prices(books.get(mid, {}), names)
        pick = make_pick(sport, prices, tcfg, fcfg)
        if pick is None:
            continue
        row = {f: "" for f in FIELDS}
        row.update({
            "captured_at": _iso(now), "market_id": mid, "sport": sport,
            "competition": comp, "event": (m.get("event") or {}).get("name", ""),
            "start_time": m.get("marketStartTime", ""),
            "selection": pick.selection, "selection_id": pick.selection_id,
            "lay_price": pick.lay_price, "detail": pick.detail,
            "in_band": int(pick.in_band)})
        rows.append(row)
        added += 1
    _write(args.out, rows)
    log.info("[%s] +%d mercados (total %d) en %s", sport, added, len(rows), args.out)


def cmd_settle(client, sport, tcfg, fcfg, args):
    cfg = tcfg if sport == "tennis" else fcfg
    rows = _read(args.out)
    pending = [r for r in rows if not r.get("won")]
    if not pending:
        log.info("[%s] nada pendiente de liquidar", sport)
        return
    books = {b["marketId"]: b for b in
             client.list_market_book([r["market_id"] for r in pending])}
    settled = 0
    for r in pending:
        b = books.get(r["market_id"])
        if not b or b.get("status") != "CLOSED":
            continue
        winner = parsing.winner_selection_id(b)
        if winner is None:
            continue
        selection_won = (str(winner) == str(r["selection_id"]))
        lay_won = not selection_won
        r["won"] = int(lay_won)
        if int(r.get("in_band") or 0):
            r["pnl_units"] = settle_lay_pnl(float(r["lay_price"]), cfg.backer_stake,
                                            selection_won, cfg.commission)
        settled += 1
    _write(args.out, rows)
    log.info("[%s] liquidados %d mercados", sport, settled)


def cmd_diagnose(client, sport, tcfg, fcfg, args):
    """Lista los próximos mercados (incl. en juego) y qué se capturaría. No escribe."""
    now = dt.datetime.utcnow()
    frm = now - dt.timedelta(hours=2)   # incluir partidos ya en curso
    markets = client.list_market_catalogue(
        event_type_id=SPORT_EVENT_TYPE[sport], market_type_codes=["MATCH_ODDS"],
        from_iso=_iso(frm), to_iso=_iso(now + dt.timedelta(minutes=args.within_min)),
        max_results=args.max_markets)
    log.info("[%s] %d mercados en ventana (-2h .. +%dmin)", sport, len(markets), args.within_min)
    if not markets:
        return
    books = {b["marketId"]: b for b in
             client.list_market_book([m["marketId"] for m in markets])}
    for m in sorted(markets, key=lambda x: x.get("marketStartTime", "")):
        comp = (m.get("competition") or {}).get("name", "")
        names = parsing.runner_names(m)
        prices = parsing.best_lay_prices(books.get(m["marketId"], {}), names)
        pick = make_pick(sport, prices, tcfg, fcfg)
        kw = _keep_competition(sport, comp, tcfg, fcfg)
        tag = "CAPTURA" if (kw and pick) else ("comp-ok" if kw else "-")
        sel = (f"{pick.selection}@{pick.lay_price} "
               f"{'EN-BANDA' if pick.in_band else 'fuera-banda'}") if pick else "(sin pick)"
        log.info("  [%-7s] %s | %-30s | %-28s | %s", tag,
                 (m.get("marketStartTime", "") or "")[:16], comp[:30],
                 ((m.get("event") or {}).get("name", "") or "")[:28], sel)


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
    p = argparse.ArgumentParser()
    p.add_argument("mode", choices=["record", "settle", "diagnose"])
    p.add_argument("--sport", choices=["tennis", "football"], default="tennis")
    p.add_argument("--within-min", type=int, default=20)
    p.add_argument("--max-markets", type=int, default=200)
    p.add_argument("--out", default=None)
    args = p.parse_args()
    if args.out is None:
        args.out = DEFAULT_OUT[args.sport]

    tcfg, fcfg = LayFavConfig(), LayDrawConfig()
    client = BetfairClient(C.BetfairConfig.from_env())
    client.login()
    if args.mode == "record":
        cmd_record(client, args.sport, tcfg, fcfg, args)
    elif args.mode == "settle":
        cmd_settle(client, args.sport, tcfg, fcfg, args)
    else:
        cmd_diagnose(client, args.sport, tcfg, fcfg, args)


if __name__ == "__main__":
    main()
