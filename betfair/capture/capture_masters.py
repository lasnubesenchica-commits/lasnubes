#!/usr/bin/env python3
"""Capturador de cuotas reales del Betfair Exchange para validar la estrategia.

Dos modos:
  - record : busca partidos de tennis Match Odds que arrancan pronto, y guarda la
             mejor cuota de LAY de cada jugador (la "cuota de cierre" del exchange).
             Pensado para correr periódicamente durante torneos Masters 1000.
  - settle : re-lee los mercados ya guardados, detecta el ganador y calcula el P&L
             que habría dado el lay al favorito (paper trading).

Sólo LEE de Betfair (sirve la delayed key para datos). NO coloca apuestas.

Salida: CSV incremental en --out (default capture/masters_capture.csv).

Uso:
    python3 -m betfair.capture.capture_masters record --within-min 15
    python3 -m betfair.capture.capture_masters settle
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import logging
import os

from ..bot import config as C
from ..bot.betfair_client import BetfairClient
from ..bot import parsing
from ..bot.strategy import LayFavConfig, select_lay_favorite, settle_lay_pnl, is_masters_competition

log = logging.getLogger("capture")
FIELDS = ["captured_at", "market_id", "competition", "event", "start_time",
          "is_masters", "fav_selection_id", "fav_name", "fav_lay_price",
          "dog_name", "dog_lay_price", "in_band", "winner_selection_id",
          "fav_won", "pnl_units"]


def _iso(t: dt.datetime) -> str:
    return t.strftime("%Y-%m-%dT%H:%M:%SZ")


def _read_rows(path):
    if not os.path.exists(path):
        return []
    with open(path, newline="") as fh:
        return list(csv.DictReader(fh))


def _write_rows(path, rows):
    with open(path, "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=FIELDS)
        w.writeheader()
        w.writerows(rows)


def cmd_record(client: BetfairClient, cfg: LayFavConfig, args):
    now = dt.datetime.utcnow()
    to = now + dt.timedelta(minutes=args.within_min)
    markets = client.list_market_catalogue(
        event_type_id=C.TENNIS_EVENT_TYPE_ID, market_type_codes=["MATCH_ODDS"],
        from_iso=_iso(now), to_iso=_iso(to), max_results=args.max_markets)
    if not markets:
        log.info("no hay mercados de tennis arrancando en %d min", args.within_min)
        return

    books = {b["marketId"]: b for b in
             client.list_market_book([m["marketId"] for m in markets])}
    existing = _read_rows(args.out)
    seen = {r["market_id"] for r in existing}
    added = 0
    for m in markets:
        mid = m["marketId"]
        if mid in seen:
            continue
        comp = (m.get("competition") or {}).get("name", "")
        event = (m.get("event") or {}).get("name", "")
        names = parsing.runner_names(m)
        prices = parsing.best_lay_prices(books.get(mid, {}), names)
        if len(prices) != 2 or any(p.lay_price is None for p in prices):
            continue
        fav, dog = sorted(prices, key=lambda p: p.lay_price)
        sel = select_lay_favorite(prices, cfg)
        row = {f: "" for f in FIELDS}
        row.update({
            "captured_at": _iso(now), "market_id": mid, "competition": comp,
            "event": event, "start_time": m.get("marketStartTime", ""),
            "is_masters": int(is_masters_competition(comp, cfg)),
            "fav_selection_id": fav.selection_id, "fav_name": fav.name,
            "fav_lay_price": fav.lay_price, "dog_name": dog.name,
            "dog_lay_price": dog.lay_price, "in_band": int(sel is not None)})
        existing.append(row)
        added += 1
    _write_rows(args.out, existing)
    log.info("registrados %d mercados nuevos (total %d) en %s",
             added, len(existing), args.out)


def cmd_settle(client: BetfairClient, cfg: LayFavConfig, args):
    rows = _read_rows(args.out)
    pending = [r for r in rows if not r.get("winner_selection_id")]
    if not pending:
        log.info("nada pendiente de liquidar")
        return
    books = {b["marketId"]: b for b in
             client.list_market_book([r["market_id"] for r in pending])}
    settled = 0
    for r in pending:
        b = books.get(r["market_id"])
        if not b or b.get("status") != "CLOSED":
            continue
        win = parsing.winner_selection_id(b)
        if win is None:
            continue
        fav_won = (str(win) == str(r["fav_selection_id"]))
        r["winner_selection_id"] = win
        r["fav_won"] = int(fav_won)
        if int(r.get("in_band") or 0):
            r["pnl_units"] = settle_lay_pnl(
                float(r["fav_lay_price"]), cfg.backer_stake, fav_won, cfg.commission)
        settled += 1
    _write_rows(args.out, rows)
    _report(rows, cfg)
    log.info("liquidados %d mercados", settled)


def _report(rows, cfg: LayFavConfig):
    bets = [r for r in rows if int(r.get("in_band") or 0)
            and int(r.get("is_masters") or 0) and r.get("pnl_units") != ""]
    if not bets:
        print("\n(aún no hay apuestas Masters in-band liquidadas)")
        return
    pnl = sum(float(r["pnl_units"]) for r in bets)
    staked = cfg.backer_stake * len(bets)
    print(f"\n=== Paper trading: LAY favorito Masters {cfg.min_fav_odds}-{cfg.max_fav_odds} ===")
    print(f"  apuestas: {len(bets)} | P&L: {pnl:+.2f} u | "
          f"ROI sobre stake: {pnl/staked*100:+.1f}%")


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(message)s")
    p = argparse.ArgumentParser()
    p.add_argument("mode", choices=["record", "settle"])
    p.add_argument("--within-min", type=int, default=15,
                   help="captura mercados que arrancan dentro de N minutos")
    p.add_argument("--max-markets", type=int, default=100)
    p.add_argument("--out", default=os.path.join(os.path.dirname(__file__),
                                                 "masters_capture.csv"))
    args = p.parse_args()

    cfg = LayFavConfig()
    client = BetfairClient(C.BetfairConfig.from_env())
    client.login()
    if args.mode == "record":
        cmd_record(client, cfg, args)
    else:
        cmd_settle(client, cfg, args)


if __name__ == "__main__":
    main()
