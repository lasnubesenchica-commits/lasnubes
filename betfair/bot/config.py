"""Configuración y endpoints de la API de Betfair.

Credenciales y App Key se leen de variables de entorno (nunca en el código).
En GitHub Actions / VPS van como Secrets. Para correr local, usa un archivo .env
(ver .env.example) cargado por tu shell o por python-dotenv.
"""
from __future__ import annotations

import os
from dataclasses import dataclass

# --- Endpoints (exchange global). Para .com.au/regional, ajustar dominios. ---
CERT_LOGIN_URL = "https://identitysso-cert.betfair.com/api/certlogin"
INTERACTIVE_LOGIN_URL = "https://identitysso.betfair.com/api/login"
KEEPALIVE_URL = "https://identitysso.betfair.com/api/keepAlive"
LOGOUT_URL = "https://identitysso.betfair.com/api/logout"
BETTING_URL = "https://api.betfair.com/exchange/betting/json-rpc/v1"
ACCOUNT_URL = "https://api.betfair.com/exchange/account/json-rpc/v1"

# Event type ids de Betfair
TENNIS_EVENT_TYPE_ID = "2"
SOCCER_EVENT_TYPE_ID = "1"

# En el Match Odds de fútbol de Betfair, "The Draw" tiene este selectionId fijo.
DRAW_SELECTION_ID = 58805

# Competiciones de fútbol a capturar (keywords, case-insensitive). El Mundial
# está activo ahora; las ligas se suman al reiniciar. Editar para ampliar/filtrar.
FOOTBALL_COMPETITION_KEYWORDS = (
    "World Cup", "FIFA World Cup", "Premier League", "La Liga", "Primera",
    "Bundesliga", "Serie A", "Ligue 1", "Champions League", "Europa League",
    "Eredivisie", "Primeira", "Championship",
)

# Torneos Masters 1000 (ATP). Patrones de nombre para reconocerlos en la API;
# el naming de Betfair varía, por eso usamos contains (case-insensitive).
MASTERS_1000_KEYWORDS = (
    "Indian Wells", "Miami", "Monte Carlo", "Monte-Carlo", "Madrid",
    "Rome", "Italian Open", "Canadian", "Toronto", "Montreal",
    "Cincinnati", "Shanghai", "Paris Masters",
)


@dataclass
class BetfairConfig:
    app_key: str
    username: str
    password: str
    cert_file: str | None = None   # ruta al .crt (login no interactivo)
    key_file: str | None = None    # ruta al .key
    timeout: int = 15

    @classmethod
    def from_env(cls) -> "BetfairConfig":
        def need(name):
            v = os.environ.get(name)
            if not v:
                raise RuntimeError(f"Falta variable de entorno {name}")
            return v
        return cls(
            app_key=need("BETFAIR_APP_KEY"),
            username=need("BETFAIR_USERNAME"),
            password=need("BETFAIR_PASSWORD"),
            cert_file=os.environ.get("BETFAIR_CERT_FILE"),
            key_file=os.environ.get("BETFAIR_KEY_FILE"),
            timeout=int(os.environ.get("BETFAIR_TIMEOUT", "15")),
        )

    @property
    def has_cert(self) -> bool:
        return bool(self.cert_file and self.key_file)
