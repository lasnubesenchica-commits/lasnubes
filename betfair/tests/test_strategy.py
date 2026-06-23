#!/usr/bin/env python3
"""Tests offline de la lógica de estrategia y del parsing de respuestas Betfair.

No tocan la red: usan JSON de muestra con la forma real de listMarketCatalogue /
listMarketBook. Correr:  python3 -m betfair.tests.test_strategy
"""
from __future__ import annotations

from ..bot import parsing
from ..bot.strategy import (LayFavConfig, RunnerPrice, select_lay_favorite,
                            settle_lay_pnl, liability, is_masters_competition)

# --- Muestras con la forma real de la API ------------------------------------
CATALOGUE_ITEM = {
    "marketId": "1.234", "marketName": "Match Odds",
    "marketStartTime": "2026-03-10T18:00:00Z",
    "competition": {"id": "1", "name": "ATP Indian Wells Masters 1000"},
    "event": {"id": "9", "name": "Alcaraz v Struff"},
    "runners": [
        {"selectionId": 1001, "runnerName": "Carlos Alcaraz"},
        {"selectionId": 1002, "runnerName": "Jan-Lennard Struff"},
    ],
}
BOOK_ITEM = {
    "marketId": "1.234", "status": "OPEN",
    "runners": [
        {"selectionId": 1001, "ex": {"availableToLay": [
            {"price": 1.30, "size": 500}, {"price": 1.31, "size": 200}]}},
        {"selectionId": 1002, "ex": {"availableToLay": [
            {"price": 4.50, "size": 120}]}},
    ],
}
BOOK_SETTLED = {
    "marketId": "1.234", "status": "CLOSED",
    "runners": [
        {"selectionId": 1001, "status": "LOSER"},
        {"selectionId": 1002, "status": "WINNER"},
    ],
}


def check(name, cond):
    print(f"  {'OK ' if cond else 'FAIL'} {name}")
    assert cond, name


def main():
    cfg = LayFavConfig()
    print("parsing:")
    names = parsing.runner_names(CATALOGUE_ITEM)
    check("nombres por selectionId", names[1001] == "Carlos Alcaraz")
    prices = parsing.best_lay_prices(BOOK_ITEM, names)
    check("mejor lay del favorito = 1.30",
          next(p.lay_price for p in prices if p.selection_id == 1001) == 1.30)
    check("ganador del settled = 1002",
          parsing.winner_selection_id(BOOK_SETTLED) == 1002)

    print("estrategia:")
    check("Indian Wells reconocido como Masters",
          is_masters_competition(CATALOGUE_ITEM["competition"]["name"], cfg))
    check("ATP 250 cualquiera NO es Masters",
          not is_masters_competition("ATP Estoril Open", cfg))

    fav = select_lay_favorite(prices, cfg)
    check("selecciona al favorito (Alcaraz)", fav and fav.selection_id == 1001)
    check("favorito en banda 1.2-1.5", fav.lay_price == 1.30)

    # favorito demasiado corto -> fuera de banda
    short = [RunnerPrice(1, "A", 1.05), RunnerPrice(2, "B", 12.0)]
    check("favorito 1.05 queda fuera de banda", select_lay_favorite(short, cfg) is None)
    # partido parejo -> favorito > 1.5 -> fuera
    even = [RunnerPrice(1, "A", 1.80), RunnerPrice(2, "B", 2.10)]
    check("favorito 1.80 queda fuera de banda", select_lay_favorite(even, cfg) is None)

    print("liquidación / P&L:")
    # liability de lay 1.30 con backer stake 2.0 = 2*0.30 = 0.60
    check("liability correcto", liability(1.30, 2.0) == 0.60)
    # fav PERDIÓ -> lay gana backer stake menos comisión: 2*(1-0.05)=1.90
    check("P&L cuando el favorito pierde (ganamos)",
          settle_lay_pnl(1.30, 2.0, fav_won=False, commission=0.05) == 1.90)
    # fav GANÓ -> perdemos el liability: -0.60
    check("P&L cuando el favorito gana (perdemos)",
          settle_lay_pnl(1.30, 2.0, fav_won=True, commission=0.05) == -0.60)

    print("\nTODOS LOS TESTS PASARON ✔")


if __name__ == "__main__":
    main()
