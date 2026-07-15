"""Almacenamiento local y cifrado de las credenciales del RPA.

Las credenciales se guardan en un archivo JSON junto a la aplicación. La
contraseña nunca se guarda en claro: se cifra con DPAPI (ver core/dpapi.py), que
ata el cifrado a la cuenta de usuario de Windows (solo el mismo usuario en la
misma máquina puede descifrarla, sin que la app maneje ninguna llave propia).
"""

from __future__ import annotations

import json
import os

from . import dpapi, rutas

RUTA = os.path.join(rutas.DATOS, "credenciales_rpa.json")


def guardar(usuario: str, contrasena: str) -> None:
    """Guarda usuario (en claro) y contraseña (cifrada) en el archivo local."""
    datos = {"usuario": usuario, "contrasena": dpapi.cifrar(contrasena)}
    with open(RUTA, "w", encoding="utf-8") as fh:
        json.dump(datos, fh)


def cargar() -> tuple[str, str] | None:
    """Devuelve (usuario, contraseña) si hay credenciales guardadas y se pueden
    descifrar; None si no hay archivo o no se puede leer/descifrar."""
    if not os.path.exists(RUTA):
        return None
    try:
        with open(RUTA, encoding="utf-8") as fh:
            datos = json.load(fh)
        return datos.get("usuario", ""), dpapi.descifrar(datos["contrasena"])
    except (OSError, ValueError, KeyError):
        return None


def borrar() -> None:
    """Elimina las credenciales guardadas (al desmarcar 'recordar')."""
    try:
        os.remove(RUTA)
    except FileNotFoundError:
        pass
