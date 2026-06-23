#!/usr/bin/env python3
"""Smoke test de conexión a Betfair (sólo lectura, no apuesta).

Verifica: login, saldo de la cuenta, y lista los próximos mercados de tennis
Match Odds marcando cuáles son Masters 1000 y cuáles tienen un favorito dentro
de la banda de la estrategia.

Carga credenciales de variables de entorno o de un archivo .env local.

    python3 -m betfair.tools.check_connection
"""
from __future__ import annotations

import datetime as dt
import os

from ..bot import config as C
from ..bot import parsing
from ..bot.betfair_client import BetfairClient
from ..bot.strategy import LayFavConfig, is_masters_competition, select_lay_favorite


def load_dotenv(path: str = ".env") -> None:
    """Carga un .env simple sin dependencias (KEY=VALUE por línea)."""
    if not os.path.exists(path):
        return
    for line in open(path):
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())


def _iso(t: dt.datetime) -> str:
    return t.strftime("%Y-%m-%dT%H:%M:%SZ")


def main() -> None:
    load_dotenv()
    cfg_bf = C.BetfairConfig.from_env()
    client = BetfairClient(cfg_bf)

    print("Logueando…", "(certificado)" if cfg_bf.has_cert else "(interactivo)")
    client.login()
    print("  sesión OK")

    funds = client.get_account_funds()
    print(f"  saldo disponible: {funds.get('availableToBetBalance')} | "
          f"exposición: {funds.get('exposure')}")

    now = dt.datetime.utcnow()
    mkts = client.list_market_catalogue(
        event_type_id=C.TENNIS_EVENT_TYPE_ID, market_type_codes=["MATCH_ODDS"],
        from_iso=_iso(now), to_iso=_iso(now + dt.timedelta(hours=48)),
        max_results=20)
    print(f"\npróximos {len(mkts)} mercados de tennis (48h):")

    cfg = LayFavConfig()
    books = {b["marketId"]: b for b in
             client.list_market_book([m["marketId"] for m in mkts])} if mkts else {}
    for m in mkts:
        comp = (m.get("competition") or {}).get("name", "")
        names = parsing.runner_names(m)
        prices = parsing.best_lay_prices(books.get(m["marketId"], {}), names)
        fav = select_lay_favorite(prices, cfg)
        tag = "MASTERS" if is_masters_competition(comp, cfg) else "-"
        favtxt = (f"-> LAY {fav.name} @ {fav.lay_price}" if fav
                  else "(sin favorito en banda)")
        event = (m.get("event") or {}).get("name", "")[:34]
        print(f"  [{tag:7}] {m.get('marketStartTime',''):20} {event:34} {favtxt}")

    print("\nConexión verificada ✔  (no se colocó ninguna apuesta)")


if __name__ == "__main__":
    main()
