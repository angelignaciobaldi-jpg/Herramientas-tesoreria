"""Cifrado local con DPAPI (Data Protection API de Windows).

DPAPI ata el cifrado a la cuenta de usuario de Windows: solo el mismo usuario en
la misma máquina puede descifrar, y sin que la app maneje (ni guarde) una llave
propia. Ideal para secretos EN REPOSO en el equipo (contraseña del RPA, token de
los microservicios, etc.).

Se usa vía ctypes para no agregar dependencias; solo funciona en Windows, que es
la plataforma de la herramienta. En otras plataformas lanza al usarse.
"""

from __future__ import annotations

import base64
import sys

if sys.platform == "win32":
    import ctypes
    import ctypes.wintypes

    class _DATA_BLOB(ctypes.Structure):
        """Estructura DATA_BLOB que DPAPI usa para entrada y salida."""

        _fields_ = [
            ("cbData", ctypes.wintypes.DWORD),
            ("pbData", ctypes.POINTER(ctypes.c_char)),
        ]

    def _a_blob(datos: bytes) -> _DATA_BLOB:
        buffer = ctypes.create_string_buffer(datos, len(datos))
        return _DATA_BLOB(
            len(datos), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_char)))

    def _de_blob(blob: _DATA_BLOB) -> bytes:
        datos = ctypes.string_at(blob.pbData, blob.cbData)
        ctypes.windll.kernel32.LocalFree(blob.pbData)  # DPAPI reserva la memoria
        return datos

    def cifrar(texto: str) -> str:
        """Cifra `texto` con DPAPI y lo devuelve en base64 (ASCII)."""
        entrada = _a_blob(texto.encode("utf-8"))
        salida = _DATA_BLOB()
        if not ctypes.windll.crypt32.CryptProtectData(
            ctypes.byref(entrada), None, None, None, None, 0,
            ctypes.byref(salida)
        ):
            raise OSError("No se pudo cifrar con DPAPI.")
        return base64.b64encode(_de_blob(salida)).decode("ascii")

    def descifrar(b64: str) -> str:
        """Descifra un texto cifrado con `cifrar` (base64 -> texto claro)."""
        entrada = _a_blob(base64.b64decode(b64))
        salida = _DATA_BLOB()
        if not ctypes.windll.crypt32.CryptUnprotectData(
            ctypes.byref(entrada), None, None, None, None, 0,
            ctypes.byref(salida)
        ):
            raise OSError("No se pudo descifrar con DPAPI.")
        return _de_blob(salida).decode("utf-8")

else:  # fuera de Windows: stubs (la app se distribuye solo para Windows)

    def cifrar(texto: str) -> str:
        raise RuntimeError("DPAPI solo está disponible en Windows.")

    def descifrar(b64: str) -> str:
        raise RuntimeError("DPAPI solo está disponible en Windows.")
