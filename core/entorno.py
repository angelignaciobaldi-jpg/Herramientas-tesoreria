"""Acceso centralizado a configuración sensible vía variables de entorno.

Reúne en un solo lugar la lectura de secretos (p. ej. el PAT de GitHub que usa
el AutoUpdater). Prioriza las variables del sistema operativo; si existe un
archivo `.env` junto a la aplicación (o en la raíz del proyecto durante el
desarrollo), lo carga SIN pisar las variables ya definidas en el sistema.

No requiere dependencias externas: incluye un mini-parser de `.env`.

Nota: las credenciales del SIPP NO van aquí; esas se guardan cifradas por
usuario con DPAPI (ver core/credenciales.py). Este módulo es para secretos de
configuración de la app (tokens de servicios, etc.).
"""

from __future__ import annotations

import os
import sys

from core import rutas

# Nombres de las variables que maneja la app.
VAR_GITHUB_PAT = "QUETZALTIC_GITHUB_PAT"

_cargado = False


class FaltaVariableEntorno(Exception):
    """No se encontró una variable de entorno requerida."""

    def __init__(self, nombre: str):
        super().__init__(
            f"Falta la variable de entorno '{nombre}'. Defínela en el sistema o "
            f"en un archivo .env junto a la aplicación."
        )


def _rutas_env() -> list[str]:
    """Ubicaciones candidatas del .env, en orden de prioridad:
      1. rutas.DATOS  -> %LOCALAPPDATA%\\... (empaquetado) o raíz del proyecto (dev).
      2. carpeta del .exe instalado (empaquetado): .env colocado en la instalación.
    Las variables del sistema operativo siguen teniendo prioridad sobre ambas."""
    rutas_posibles = [os.path.join(rutas.DATOS, ".env")]
    if getattr(sys, "frozen", False):
        rutas_posibles.append(os.path.join(rutas.INSTALL, ".env"))
    return rutas_posibles


def _cargar_archivo(ruta: str) -> None:
    if not os.path.exists(ruta):
        return
    with open(ruta, encoding="utf-8") as fh:
        for linea in fh:
            linea = linea.strip()
            if not linea or linea.startswith("#") or "=" not in linea:
                continue
            clave, _, valor = linea.partition("=")
            # Quita comillas envolventes opcionales del valor.
            valor = valor.strip().strip('"').strip("'")
            os.environ.setdefault(clave.strip(), valor)


def cargar_env(ruta: str | None = None, forzar: bool = False) -> None:
    """Carga pares CLAVE=VALOR desde el/los archivo(s) .env, si existen. No
    sobrescribe variables ya presentes en el entorno (estas tienen prioridad)."""
    global _cargado
    if _cargado and not forzar:
        return
    rutas_a_cargar = [ruta] if ruta else _rutas_env()
    for r in rutas_a_cargar:
        _cargar_archivo(r)
    _cargado = True


def obtener(nombre: str, por_defecto: str | None = None, requerido: bool = False) -> str | None:
    """Lee una variable de entorno (cargando el .env la primera vez). Si
    `requerido` y no existe (o está vacía), lanza FaltaVariableEntorno."""
    cargar_env()
    valor = os.environ.get(nombre, por_defecto)
    if requerido and not valor:
        raise FaltaVariableEntorno(nombre)
    return valor


def github_pat(requerido: bool = True) -> str | None:
    """PAT de GitHub para acceder al repositorio privado (AutoUpdater)."""
    return obtener(VAR_GITHUB_PAT, requerido=requerido)
