"""Archivos copiados en el Explorador, leídos del portapapeles de Windows.

Sirve para no depender del diálogo de selección de archivos, que en la máquina de
tesorería se queda colgado. Además el propio proyecto ya tenía la nota: el
diálogo de multiselección «falla al elegir muchos archivos con nombres largos»
(ver `ui/alta_beneficiarios.py`), y los reportes de saldos son justo eso —una
docena de archivos con nombres kilométricos que pone el portal—.

Cuando alguien copia archivos en el Explorador, Windows deja en el portapapeles
un `CF_HDROP`: la MISMA estructura que viaja en un arrastre. Aquí se lee con
ctypes, que ya es como el proyecto habla con Win32 (ver `core/win_titlebar.py`).

Es una lectura BAJO DEMANDA —se ejecuta cuando el usuario pega, y nada más—, lo
que la distingue de intentar el arrastre de verdad: aquello obliga a colgarse del
procedimiento de ventana de Flutter, y cada mensaje del sistema tendría que tomar
el GIL. En una app que se pasa segundos parseando Excel, eso congelaría la
ventana en vez de arreglarla.
"""

from __future__ import annotations

import os
import sys

CF_HDROP = 15

_disponible = sys.platform == "win32"

if _disponible:
    import ctypes
    from ctypes import wintypes

    _u32 = ctypes.WinDLL("user32", use_last_error=True)
    _k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    _shell = ctypes.WinDLL("shell32", use_last_error=True)

    # Declarar los tipos NO es opcional: sin esto ctypes asume que todo devuelve
    # un int de 32 bits, trunca el puntero del portapapeles en Python de 64 y la
    # primera lectura revienta con una violación de acceso.
    _u32.IsClipboardFormatAvailable.argtypes = [wintypes.UINT]
    _u32.OpenClipboard.argtypes = [wintypes.HWND]
    _u32.GetClipboardData.argtypes = [wintypes.UINT]
    _u32.GetClipboardData.restype = wintypes.HANDLE
    _k32.GlobalLock.argtypes = [wintypes.HANDLE]
    _k32.GlobalLock.restype = ctypes.c_void_p
    _k32.GlobalUnlock.argtypes = [wintypes.HANDLE]
    _shell.DragQueryFileW.argtypes = [ctypes.c_void_p, wintypes.UINT,
                                      wintypes.LPWSTR, wintypes.UINT]
    _shell.DragQueryFileW.restype = wintypes.UINT


def hay_archivos() -> bool:
    """Si el portapapeles trae archivos copiados. No lo abre."""
    if not _disponible:
        return False
    try:
        return bool(_u32.IsClipboardFormatAvailable(CF_HDROP))
    except Exception:  # noqa: BLE001 — leer el portapapeles nunca es crítico
        return False


def archivos() -> list[str]:
    """Las rutas de los archivos copiados, o lista vacía.

    Solo devuelve ARCHIVOS: si se copió una carpeta, se descarta —quien quiera
    una carpeta completa tiene el botón que la pide—. Y nunca lanza: que el
    portapapeles traiga algo raro no puede tumbar la pantalla."""
    if not hay_archivos():
        return []
    # El portapapeles es un recurso global del sistema y puede estar tomado por
    # otra aplicación; si no abre, se devuelve vacío y el usuario reintenta.
    if not _u32.OpenClipboard(None):
        return []
    try:
        mango = _u32.GetClipboardData(CF_HDROP)
        if not mango:
            return []
        puntero = _k32.GlobalLock(mango)
        if not puntero:
            return []
        try:
            # 0xFFFFFFFF pide la CUENTA en vez de un elemento concreto.
            cuantos = _shell.DragQueryFileW(puntero, 0xFFFFFFFF, None, 0)
            rutas = []
            for i in range(cuantos):
                largo = _shell.DragQueryFileW(puntero, i, None, 0) + 1
                buf = ctypes.create_unicode_buffer(largo)
                _shell.DragQueryFileW(puntero, i, buf, largo)
                if buf.value and os.path.isfile(buf.value):
                    rutas.append(buf.value)
            return rutas
        finally:
            _k32.GlobalUnlock(mango)
    except Exception:  # noqa: BLE001 — se trata como «no había nada que pegar»
        return []
    finally:
        _u32.CloseClipboard()
