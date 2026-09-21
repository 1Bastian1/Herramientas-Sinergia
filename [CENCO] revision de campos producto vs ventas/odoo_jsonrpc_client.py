"""Cliente mínimo para consumir una instancia de Odoo mediante JSON-RPC.

Configuración requerida (variables de entorno):
    ODOO_URL=https://mi-odoo.example.com
    ODOO_DB=mi_base
    ODOO_USERNAME=usuario@example.com
    ODOO_PASSWORD=contraseña

Ejemplo:
    python odoo_jsonrpc_client.py
"""

from __future__ import annotations

import itertools
import json
import os
import getpass
from importlib import import_module
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


class OdooJsonRpcError(RuntimeError):
    """Error devuelto por Odoo o por la comunicación JSON-RPC."""


class OdooClient:
    def __init__(self, url: str, database: str, username: str, password: str) -> None:
        self.url = url.rstrip("/")
        self.database = database
        self.username = username
        self.password = password
        self.uid: int | None = None
        self._request_ids = itertools.count(1)

    def _call(self, service: str, method: str, *args: Any) -> Any:
        payload = {
            "jsonrpc": "2.0",
            "method": "call",
            "params": {"service": service, "method": method, "args": list(args)},
            "id": next(self._request_ids),
        }
        request = Request(
            f"{self.url}/jsonrpc",
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )

        try:
            # Las agregaciones de grandes volúmenes pueden tardar más que una
            # consulta normal; el límite evita abortarlas prematuramente.
            with urlopen(request, timeout=120) as response:
                data = json.load(response)
        except HTTPError as exc:
            raise OdooJsonRpcError(f"HTTP {exc.code} al conectar con Odoo") from exc
        except URLError as exc:
            raise OdooJsonRpcError(f"No se pudo conectar a Odoo: {exc.reason}") from exc

        if "error" in data:
            error = data["error"]
            details = error.get("data", {}).get("message") or error.get("message")
            raise OdooJsonRpcError(f"Error JSON-RPC: {details}")
        return data["result"]

    def authenticate(self) -> int:
        uid = self._call("common", "login", self.database, self.username, self.password)
        if not uid:
            raise OdooJsonRpcError("Autenticación rechazada: revisa base de datos, usuario y contraseña.")
        self.uid = int(uid)
        return self.uid

    def execute_kw(
        self, model: str, method: str, args: list[Any] | None = None, kwargs: dict[str, Any] | None = None
    ) -> Any:
        if self.uid is None:
            self.authenticate()
        return self._call(
            "object",
            "execute_kw",
            self.database,
            self.uid,
            self.password,
            model,
            method,
            args or [],
            kwargs or {},
        )


def client_from_environment() -> OdooClient:
    required = ("ODOO_URL", "ODOO_DB", "ODOO_USERNAME", "ODOO_PASSWORD")
    local = {}
    try:
        local = vars(import_module("odoo_local_config"))
    except ModuleNotFoundError:
        pass
    values = {name: os.environ.get(name) or local.get(name, "") for name in required}
    prompts = {
        "ODOO_URL": "URL de Odoo (ej. https://empresa.odoo.com): ",
        "ODOO_DB": "Base de datos de Odoo: ",
        "ODOO_USERNAME": "Usuario de Odoo: ",
    }
    for name, prompt in prompts.items():
        if not values[name]:
            values[name] = input(prompt).strip()
    if not values["ODOO_PASSWORD"]:
        values["ODOO_PASSWORD"] = getpass.getpass("Contraseña de Odoo: ")
    missing = [name for name, value in values.items() if not value]
    if missing:
        raise ValueError(f"Faltan datos de conexión: {', '.join(missing)}")
    return OdooClient(
        values["ODOO_URL"],
        values["ODOO_DB"],
        values["ODOO_USERNAME"],
        values["ODOO_PASSWORD"],
    )


if __name__ == "__main__":
    odoo = client_from_environment()
    user_id = odoo.authenticate()
    print(f"Conectado correctamente. UID: {user_id}")

    # Ejemplo: leer los primeros cinco partners. Ajusta permisos/modelo según tu instancia.
    partners = odoo.execute_kw(
        "res.partner",
        "search_read",
        [[]],
        {"fields": ["name", "email"], "limit": 5},
    )
    print(json.dumps(partners, indent=2, ensure_ascii=False))
