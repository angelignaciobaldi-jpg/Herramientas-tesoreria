"""Smoke test de imports.

Importa el entrypoint `app` y TODOS los submódulos de `core/` y `ui/` para
detectar módulos faltantes o imports rotos ANTES de compilar/publicar.

Atrapa el caso que provocó el crash "cannot import name 'X' from 'core'": un
módulo referenciado por el código pero no versionado en el repo. Sin este
chequeo, ese error solo se veía como un diálogo de crash al abrir la app YA
instalada en la máquina del usuario.

Sale con código 1 (falla el job de CI) si cualquier módulo no importa; 0 si todos
importan. No requiere binarios externos (Chromium/Tesseract): solo valida que los
paquetes de Python resuelvan.
"""

from __future__ import annotations

import importlib
import os
import pkgutil
import sys

# Asegura que la RAÍZ del proyecto esté en sys.path (al correr como
# `python scripts/smoke_import.py`, sys.path[0] es 'scripts/', no la raíz), para
# que 'app', 'core' y 'ui' se resuelvan sin importar desde dónde se invoque.
_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)


def _importar_paquete(nombre: str) -> list[tuple[str, str]]:
    """Importa el paquete y, recursivamente, todos sus submódulos. Devuelve la
    lista de (módulo, error) que fallaron."""
    fallos: list[tuple[str, str]] = []
    try:
        paquete = importlib.import_module(nombre)
    except Exception as exc:  # noqa: BLE001 — se reporta, no se propaga
        return [(nombre, repr(exc))]
    for mod in pkgutil.walk_packages(paquete.__path__, paquete.__name__ + "."):
        try:
            importlib.import_module(mod.name)
        except Exception as exc:  # noqa: BLE001
            fallos.append((mod.name, repr(exc)))
    return fallos


def main() -> int:
    fallos: list[tuple[str, str]] = []
    # Entrypoint (sus imports perezosos se cubren al recorrer core/ y ui/).
    try:
        importlib.import_module("app")
    except Exception as exc:  # noqa: BLE001
        fallos.append(("app", repr(exc)))
    # Todos los submódulos de la app: aquí se disparan los imports internos (p.
    # ej. ui.dispersion_no_pemex -> from core import preferencias, conciliacion…).
    for paquete in ("core", "ui"):
        fallos.extend(_importar_paquete(paquete))

    if fallos:
        print("SMOKE TEST FALLO - modulos que no importan:")
        for nombre, err in fallos:
            print(f"  - {nombre}: {err}")
        return 1
    print("SMOKE TEST OK - todos los modulos de app/core/ui importan.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
