"""Solicita las credenciales de Odoo y verifica una conexión JSON-RPC."""

from __future__ import annotations

from getpass import getpass

from odoo_jsonrpc import OdooJsonRpcClient, OdooJsonRpcError


def main() -> None:
    print("Conexión a Odoo por JSON-RPC")
    url = input("URL de Odoo (ej.: https://empresa.odoo.com): ").strip()
    database = input("Base de datos: ").strip()
    username = input("Usuario: ").strip()
    password = getpass("Contraseña: ")

    if not all((url, database, username, password)):
        raise ValueError("URL, base de datos, usuario y contraseña son obligatorios.")

    client = OdooJsonRpcClient(url, database, username, password)
    user_id = client.authenticate()
    print(f"Conexión correcta. ID de usuario de Odoo: {user_id}")
    print("No se modificó ningún dato.")


if __name__ == "__main__":
    try:
        main()
    except OdooJsonRpcError as error:
        print(f"No fue posible conectar: {error}")
        raise SystemExit(1)
