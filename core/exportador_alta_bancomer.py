"""Generación del TXT de alta de cuentas en el portal Bancomer (ancho fijo).

Distinto del TXT de dispersión: este sirve para dar de alta las cuentas de los
beneficiarios en el portal. Cada registro es una línea de 177 caracteres
terminada en CRLF:

    pos   0 ancho 18 : CLABE
    pos  18 ancho  3 : moneda -> 'MXP'
    pos  21 ancho 16 : monto (límite de la cuenta), fijo, ceros a la izquierda
    pos  37 ancho 30 : nombre del beneficiario (izquierda, espacios)
    pos  67 ancho 30 : nombre del beneficiario (repetido)
    pos  97 ancho 80 : correo de notificación (izquierda, espacios)

El monto es un límite estándar para el alta (no es un pago), por eso es fijo.
"""

from __future__ import annotations

import re
import unicodedata

from .exportador import _ascii_banco

MONEDA = "MXP"
# Límite estándar del alta (16 caracteres): 100000.00 con ceros a la izquierda.
MONTO_FIJO = "0000000100000.00"
ANCHO_NOMBRE = 30
ANCHO_EMAIL = 80
FIN_LINEA = "\r\n"


def _campo_nombre(texto: str) -> str:
    return _ascii_banco(texto)[:ANCHO_NOMBRE].ljust(ANCHO_NOMBRE)


def _campo_email(correo: str) -> str:
    """Correo en mayúsculas y ASCII, justificado a la izquierda con espacios.
    Conserva los caracteres válidos de un correo (@ . _ % + -)."""
    sin_acentos = unicodedata.normalize("NFKD", correo or "").encode("ascii", "ignore").decode()
    limpio = re.sub(r"[^A-Z0-9@._%+\-]", "", sin_acentos.upper())
    return limpio[:ANCHO_EMAIL].ljust(ANCHO_EMAIL)


def linea_registro(clabe: str, beneficiario: str, correo: str) -> str:
    """Construye la línea de ancho fijo (177) para el alta de una cuenta."""
    clabe = re.sub(r"\D", "", clabe or "")
    return (
        clabe
        + MONEDA
        + MONTO_FIJO
        + _campo_nombre(beneficiario)
        + _campo_nombre(beneficiario)
        + _campo_email(correo)
    )


def generar_txt(registros: list[tuple[str, str, str]]) -> str:
    """Genera el contenido completo del TXT de alta.

    Args:
        registros: lista de tuplas (clabe, beneficiario, correo).
    """
    return "".join(linea_registro(*r) + FIN_LINEA for r in registros)
