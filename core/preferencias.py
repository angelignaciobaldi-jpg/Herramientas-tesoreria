"""Preferencias del usuario que se recuerdan entre sesiones (clave -> lista).

Pensado para guardar/cargar selecciones que el usuario repite, como los filtros
de Empresa y Tipo de Solicitud de la pantalla de Dispersión (No Pemex). Se
guarda como un JSON en la carpeta escribible de datos (rutas.DATOS)."""

from __future__ import annotations

import json
import os

from core import rutas

_RUTA = os.path.join(rutas.DATOS, "preferencias.json")


def _cargar_todo() -> dict:
    try:
        with open(_RUTA, encoding="utf-8") as fh:
            datos = json.load(fh)
        return datos if isinstance(datos, dict) else {}
    except (OSError, json.JSONDecodeError):
        return {}


def cargar_lista(clave: str) -> list[str]:
    """Devuelve la lista guardada bajo `clave` (o [] si no hay/está corrupta)."""
    valor = _cargar_todo().get(clave)
    return [str(v) for v in valor] if isinstance(valor, list) else []


def guardar_lista(clave: str, valores: list[str]) -> None:
    """Guarda `valores` bajo `clave` (mezclando con el resto de preferencias)."""
    guardar_valor(clave, list(valores))


def cargar_valor(clave: str, defecto=None):
    """Devuelve el valor guardado bajo `clave` (o `defecto` si no existe)."""
    valor = _cargar_todo().get(clave)
    return defecto if valor is None else valor


def guardar_valor(clave: str, valor) -> None:
    """Guarda un valor JSON-serializable bajo `clave` (mezcla con el resto)."""
    datos = _cargar_todo()
    datos[clave] = valor
    os.makedirs(os.path.dirname(_RUTA), exist_ok=True)
    with open(_RUTA, "w", encoding="utf-8") as fh:
        json.dump(datos, fh, ensure_ascii=False, indent=2)
