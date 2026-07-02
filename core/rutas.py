"""Rutas base de la aplicación, válidas tanto en desarrollo como empaquetada.

Distingue tres ubicaciones:
  - BUNDLE   : archivos empaquetados de solo lectura (tessdata, Imagenes). Al
               estar congelado con PyInstaller viven en la carpeta temporal de
               extracción (sys._MEIPASS).
  - DATOS    : archivos escribibles/del usuario (base de datos, caché,
               credenciales, .env, el Excel de cuentas). Al estar congelado se
               usa %LOCALAPPDATA%\\... — NO la carpeta del .exe, que suele estar
               en Archivos de Programa (solo lectura para usuarios estándar).
  - INSTALL  : carpeta donde está el .exe instalado (solo lectura en uso normal).
En desarrollo, las tres apuntan a la carpeta del proyecto.
"""

from __future__ import annotations

import os
import sys

_PROYECTO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_SUBCARPETA_DATOS = os.path.join("Quetzaltic Solutions", "Herramientas de Tesoreria")

if getattr(sys, "frozen", False):  # ejecutándose como .exe (PyInstaller)
    BUNDLE = getattr(sys, "_MEIPASS", _PROYECTO)
    INSTALL = os.path.dirname(sys.executable)
    # Datos del usuario en un sitio escribible (no en Archivos de Programa).
    _base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    DATOS = os.path.join(_base, _SUBCARPETA_DATOS)
    os.makedirs(DATOS, exist_ok=True)
else:  # ejecutándose como script de Python
    BUNDLE = _PROYECTO
    INSTALL = _PROYECTO
    DATOS = _PROYECTO
