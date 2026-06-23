"""Cliente mínimo de la API de Betfair (login + Betting/Account JSON-RPC).

Soporta login no interactivo por certificado (recomendado para bots) y login
interactivo (usuario+clave) para arrancar rápido. Sólo depende de `requests`.

Las funciones de lectura (listMarketCatalogue, listMarketBook, getAccountFunds)
son suficientes para el capturador y el paper trading. `place_orders` se incluye
para la fase de apuestas reales pero el paper trading NO la usa.
"""
from __future__ import annotations

import logging
from typing import Any

import requests

from . import config as C

log = logging.getLogger("betfair")


class BetfairError(RuntimeError):
    pass


class BetfairClient:
    def __init__(self, cfg: C.BetfairConfig):
        self.cfg = cfg
        self.session_token: str | None = None
        self._http = requests.Session()

    # ----------------------------- Login -----------------------------------
    def login(self) -> str:
        """Inicia sesión (certificado si está configurado, si no interactivo)."""
        if self.cfg.has_cert:
            return self._login_cert()
        return self._login_interactive()

    def _login_cert(self) -> str:
        headers = {"X-Application": self.cfg.app_key,
                   "Content-Type": "application/x-www-form-urlencoded"}
        resp = self._http.post(
            C.CERT_LOGIN_URL,
            data={"username": self.cfg.username, "password": self.cfg.password},
            cert=(self.cfg.cert_file, self.cfg.key_file),
            headers=headers, timeout=self.cfg.timeout)
        resp.raise_for_status()
        data = resp.json()
        if data.get("loginStatus") != "SUCCESS":
            raise BetfairError(f"certlogin falló: {data.get('loginStatus')}")
        self.session_token = data["sessionToken"]
        log.info("login por certificado OK")
        return self.session_token

    def _login_interactive(self) -> str:
        headers = {"X-Application": self.cfg.app_key, "Accept": "application/json",
                   "Content-Type": "application/x-www-form-urlencoded"}
        resp = self._http.post(
            C.INTERACTIVE_LOGIN_URL,
            data={"username": self.cfg.username, "password": self.cfg.password},
            headers=headers, timeout=self.cfg.timeout)
        resp.raise_for_status()
        data = resp.json()
        if data.get("status") != "SUCCESS":
            raise BetfairError(f"login falló: {data.get('status')} {data.get('error')}")
        self.session_token = data["token"]
        log.info("login interactivo OK")
        return self.session_token

    def keep_alive(self) -> None:
        self._http.post(C.KEEPALIVE_URL, headers=self._auth_headers(),
                        timeout=self.cfg.timeout)

    # --------------------------- JSON-RPC core ------------------------------
    def _auth_headers(self) -> dict:
        if not self.session_token:
            raise BetfairError("no hay sesión; llama a login() primero")
        return {"X-Application": self.cfg.app_key,
                "X-Authentication": self.session_token,
                "Content-Type": "application/json", "Accept": "application/json"}

    def _rpc(self, url: str, api: str, method: str, params: dict) -> Any:
        body = {"jsonrpc": "2.0", "method": f"{api}/v1.0/{method}",
                "params": params, "id": 1}
        resp = self._http.post(url, json=body, headers=self._auth_headers(),
                               timeout=self.cfg.timeout)
        resp.raise_for_status()
        payload = resp.json()
        if "error" in payload:
            raise BetfairError(f"{method} error: {payload['error']}")
        return payload["result"]

    def _betting(self, method: str, params: dict) -> Any:
        return self._rpc(C.BETTING_URL, "SportsAPING", method, params)

    # ----------------------------- Lecturas ---------------------------------
    def list_competitions(self, event_type_id: str) -> list[dict]:
        return self._betting("listCompetitions",
                             {"filter": {"eventTypeIds": [event_type_id]}})

    def list_market_catalogue(self, *, event_type_id: str,
                              market_type_codes: list[str], max_results: int = 100,
                              from_iso: str | None = None, to_iso: str | None = None,
                              text_query: str | None = None) -> list[dict]:
        mfilter: dict = {"eventTypeIds": [event_type_id],
                         "marketTypeCodes": market_type_codes}
        if from_iso or to_iso:
            mfilter["marketStartTime"] = {k: v for k, v in
                                          (("from", from_iso), ("to", to_iso)) if v}
        if text_query:
            mfilter["textQuery"] = text_query
        return self._betting("listMarketCatalogue", {
            "filter": mfilter, "maxResults": str(max_results),
            "marketProjection": ["COMPETITION", "EVENT", "MARKET_START_TIME",
                                 "RUNNER_DESCRIPTION"]})

    def list_market_book(self, market_ids: list[str]) -> list[dict]:
        return self._betting("listMarketBook", {
            "marketIds": market_ids,
            "priceProjection": {"priceData": ["EX_BEST_OFFERS"],
                                "virtualise": True}})

    def get_account_funds(self) -> dict:
        return self._rpc(C.ACCOUNT_URL, "AccountAPING", "getAccountFunds", {})

    # --------------------- Apuestas (fase real, NO paper) -------------------
    def place_lay(self, market_id: str, selection_id: int, price: float,
                  backer_stake: float, persistence: str = "LAPSE") -> dict:
        """Coloca un LAY. Úsese sólo en modo live; el paper trading NO llama aquí."""
        return self._betting("placeOrders", {
            "marketId": market_id,
            "instructions": [{
                "selectionId": selection_id, "side": "LAY", "orderType": "LIMIT",
                "limitOrder": {"size": round(backer_stake, 2),
                               "price": price, "persistenceType": persistence}}]})
