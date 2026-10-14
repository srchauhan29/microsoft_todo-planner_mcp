"""Minimal async Microsoft Graph client with retry, paging and ETag support."""

from __future__ import annotations

import asyncio
import logging
from typing import Any

import anyio
import httpx

from .auth import TokenProvider
from .config import Settings

log = logging.getLogger(__name__)

_RETRY_STATUS = {429, 502, 503, 504}
_MAX_RETRIES = 4


class GraphError(RuntimeError):
    def __init__(self, status: int, code: str, message: str, request_id: str | None) -> None:
        self.status, self.code, self.request_id = status, code, request_id
        super().__init__(f"Graph {status} {code}: {message}" + (f" (request-id {request_id})" if request_id else ""))


class GraphClient:
    def __init__(self, settings: Settings, tokens: TokenProvider) -> None:
        self._base = settings.graph_base
        self._tokens = tokens
        self._http = httpx.AsyncClient(timeout=httpx.Timeout(30.0))

    async def aclose(self) -> None:
        await self._http.aclose()

    async def request(
        self,
        method: str,
        path: str,
        *,
        json: dict[str, Any] | None = None,
        params: dict[str, Any] | None = None,
        etag: str | None = None,
        return_representation: bool = False,
        token: str | None = None,
    ) -> dict[str, Any] | None:
        url = path if path.startswith("https://") else f"{self._base}{path}"
        if not url.startswith(self._base):
            # Only ever send the bearer token to Microsoft Graph.
            raise ValueError(f"Refusing to call non-Graph URL: {url}")

        for attempt in range(_MAX_RETRIES + 1):
            bearer = token or await anyio.to_thread.run_sync(self._tokens.get_token)
            headers = {"Authorization": f"Bearer {bearer}", "Accept": "application/json"}
            if etag:
                headers["If-Match"] = etag
            if return_representation:
                headers["Prefer"] = "return=representation"

            resp = await self._http.request(method, url, json=json, params=params, headers=headers)
            if resp.status_code in _RETRY_STATUS and attempt < _MAX_RETRIES:
                delay = float(resp.headers.get("Retry-After", 2 ** attempt))
                log.warning("Graph %s %s -> %s, retrying in %.1fs", method, path, resp.status_code, delay)
                await asyncio.sleep(min(delay, 30))
                continue
            break

        log.info("Graph %s %s -> %s", method, path.split("?")[0], resp.status_code)
        if resp.status_code >= 400:
            try:
                err = resp.json().get("error", {})
            except ValueError:
                err = {}
            raise GraphError(
                resp.status_code,
                err.get("code", "UnknownError"),
                err.get("message", resp.text[:300]),
                resp.headers.get("request-id"),
            )
        if resp.status_code == 204 or not resp.content:
            return None
        return resp.json()

    async def get(self, path: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        return await self.request("GET", path, params=params) or {}

    async def get_all(self, path: str, params: dict[str, Any] | None = None, max_items: int = 500) -> list[dict]:
        items: list[dict] = []
        page = await self.get(path, params)
        while True:
            items.extend(page.get("value", []))
            next_link = page.get("@odata.nextLink")
            if not next_link or len(items) >= max_items:
                break
            page = await self.get(next_link)
        return items[:max_items]

    async def post(self, path: str, body: dict[str, Any]) -> dict[str, Any]:
        return await self.request("POST", path, json=body) or {}

    async def patch(self, path: str, body: dict[str, Any], etag: str | None = None) -> dict[str, Any] | None:
        return await self.request("PATCH", path, json=body, etag=etag, return_representation=etag is not None)

    async def delete(self, path: str, etag: str | None = None) -> None:
        await self.request("DELETE", path, etag=etag)
