"""Cliente básico y de solo lectura para Odoo mediante JSON-RPC."""

from __future__ import annotations

import itertools
import json
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


class OdooJsonRpcError(RuntimeError):
    """Error de autenticación, comunicación o respuesta de Odoo."""


class OdooJsonRpcClient:
    def __init__(self, url: str, database: str, username: str, password: str) -> None:
        self.url = url.rstrip("/")
        self.database = database
        self.username = username
        self.password = password
        self.user_id: int | None = None
        self._request_ids = itertools.count(1)

    def _request(self, service: str, method: str, *arguments: Any) -> Any:
        payload = {
            "jsonrpc": "2.0",
            "method": "call",
            "params": {
                "service": service,
                "method": method,
                "args": list(arguments),
            },
            "id": next(self._request_ids),
        }
        request = Request(
            f"{self.url}/jsonrpc",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urlopen(request, timeout=45) as response:
                result = json.load(response)
        except HTTPError as exc:
            raise OdooJsonRpcError(f"Odoo respondió HTTP {exc.code}.") from exc
        except URLError as exc:
            raise OdooJsonRpcError(f"No se pudo conectar a Odoo: {exc.reason}") from exc

        if "error" in result:
            error = result["error"]
            message = error.get("data", {}).get("message") or error.get("message", "Error desconocido")
            raise OdooJsonRpcError(f"Odoo devolvió un error: {message}")
        return result["result"]

    def authenticate(self) -> int:
        user_id = self._request("common", "login", self.database, self.username, self.password)
        if not user_id:
            raise OdooJsonRpcError("Odoo rechazó las credenciales.")
        self.user_id = int(user_id)
        return self.user_id

    def execute_kw(
        self,
        model: str,
        method: str,
        args: list[Any] | None = None,
        kwargs: dict[str, Any] | None = None,
    ) -> Any:
        """Ejecuta una operación de Odoo una vez autenticado."""
        if self.user_id is None:
            self.authenticate()
        return self._request(
            "object",
            "execute_kw",
            self.database,
            self.user_id,
            self.password,
            model,
            method,
            args or [],
            kwargs or {},
        )
