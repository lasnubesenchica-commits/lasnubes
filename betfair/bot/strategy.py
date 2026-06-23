"""Lógica de selección de la estrategia: LAY al favorito corto en Masters 1000.

Pura y sin dependencias de red para poder testearla offline. La misma regla la
usan el capturador (para etiquetar candidatos), el paper trading y, más adelante,
el bot real.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from . import config as C


@dataclass(frozen=True)
class LayFavConfig:
    min_fav_odds: float = 1.2      # banda de cuota del favorito (validada en backtest)
    max_fav_odds: float = 1.5
    backer_stake: float = 2.0      # stake que laias; liability = stake*(cuota-1)
    commission: float = 0.05
    masters_keywords: tuple = field(default=C.MASTERS_1000_KEYWORDS)


@dataclass(frozen=True)
class RunnerPrice:
    selection_id: int
    name: str
    lay_price: float | None        # mejor cuota disponible para LAY


def is_masters_competition(competition_name: str, cfg: LayFavConfig) -> bool:
    name = (competition_name or "").lower()
    return any(kw.lower() in name for kw in cfg.masters_keywords)


def select_lay_favorite(runners: list[RunnerPrice], cfg: LayFavConfig) -> RunnerPrice | None:
    """Devuelve el favorito a layear si su cuota está en la banda, o None.

    El favorito es el de MENOR cuota de lay. Requiere exactamente un partido a 2
    jugadores con precios disponibles para ambos.
    """
    priced = [r for r in runners if r.lay_price and r.lay_price > 1.0]
    if len(priced) != 2:
        return None
    fav = min(priced, key=lambda r: r.lay_price)
    if cfg.min_fav_odds <= fav.lay_price <= cfg.max_fav_odds:
        return fav
    return None


def liability(price: float, backer_stake: float) -> float:
    """Capital en riesgo de un lay = backer_stake * (cuota - 1)."""
    return round(backer_stake * (price - 1.0), 2)


def settle_lay_pnl(price: float, backer_stake: float, fav_won: bool,
                   commission: float) -> float:
    """P&L de un lay al favorito: gana el stake (menos comisión) si el fav PIERDE;
    pierde el liability si el fav GANA."""
    if fav_won:
        return -liability(price, backer_stake)
    return round(backer_stake * (1.0 - commission), 2)
