"""Helpers para interpretar respuestas de la API (sin red, testeables offline)."""
from __future__ import annotations

from .strategy import RunnerPrice


def runner_names(catalogue_market: dict) -> dict[int, str]:
    """selectionId -> nombre del jugador, desde un item de listMarketCatalogue."""
    return {r["selectionId"]: r.get("runnerName", str(r["selectionId"]))
            for r in catalogue_market.get("runners", [])}


def best_lay_prices(market_book: dict, names: dict[int, str]) -> list[RunnerPrice]:
    """Mejor cuota de LAY por runner, desde un item de listMarketBook.

    Estructura Betfair: runner['ex']['availableToLay'] = [{'price':..,'size':..}, ...]
    (ordenado del mejor al peor). Tomamos el primero.
    """
    out = []
    for r in market_book.get("runners", []):
        sid = r["selectionId"]
        lays = (r.get("ex") or {}).get("availableToLay") or []
        price = lays[0]["price"] if lays else None
        out.append(RunnerPrice(selection_id=sid,
                               name=names.get(sid, str(sid)), lay_price=price))
    return out


def winner_selection_id(market_book: dict) -> int | None:
    """selectionId del ganador en un mercado liquidado (status == 'WINNER')."""
    for r in market_book.get("runners", []):
        if r.get("status") == "WINNER":
            return r["selectionId"]
    return None
