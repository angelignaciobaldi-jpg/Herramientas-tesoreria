"""Rastro de pasos para diagnosticar cuelgues en la máquina del usuario.

Existe por un caso concreto: la aplicación se congela al abrir el navegador de
archivos en una sola máquina, hay que matarla desde el Administrador de tareas y
no deja NI UN traceback —porque no hay excepción, se queda bloqueada—. Sin poder
reproducirlo, la única forma de saber dónde se detiene es que ella misma vaya
dejando marcas al pasar.

La idea es simple y la restricción también: cada marca se ESCRIBE Y SE CIERRA en
el momento. Si el archivo quedara abierto con buffer, lo último —justo lo que
interesa— se perdería al matar el proceso, que es exactamente como termina este
fallo.

El archivo vive junto a los datos del usuario (`rutas.DATOS/diagnostico.log`), se
recorta solo al llegar a 1 MB y nunca propaga un error: registrar un paso jamás
puede ser la causa de que falle la operación que está registrando.
"""

from __future__ import annotations

import datetime
import os
import threading

from . import rutas

RUTA = os.path.join(rutas.DATOS, "diagnostico.log")

# Un tamaño en el que caben miles de pasos y que igual no estorba si se queda
# olvidado. Al pasarlo se rota a '.1' y se empieza de nuevo.
_TOPE_BYTES = 1_000_000

_candado = threading.Lock()


def registrar(etapa: str, detalle: str = "") -> None:
    """Anota un paso, con hora e hilo. Best-effort: nunca lanza."""
    try:
        ahora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S.%f")[:-3]
        linea = "{} [{}] {}{}\n".format(
            ahora, threading.current_thread().name, etapa,
            " · {}".format(detalle) if detalle else "")
        with _candado:
            _rotar_si_hace_falta()
            # Se abre, se escribe y se cierra en cada paso, a propósito: ver el
            # encabezado del módulo.
            with open(RUTA, "a", encoding="utf-8") as f:
                f.write(linea)
    except Exception:  # noqa: BLE001 — el diagnóstico jamás puede romper nada
        pass


def _rotar_si_hace_falta() -> None:
    try:
        if os.path.getsize(RUTA) < _TOPE_BYTES:
            return
    except OSError:
        return
    try:
        os.replace(RUTA, RUTA + ".1")
    except OSError:
        pass


def marcar_arranque(version: str = "") -> None:
    """Separador al abrir la app, para distinguir una sesión de la siguiente."""
    registrar("=" * 20 + " arranque " + (version or "") + " " + "=" * 20)
