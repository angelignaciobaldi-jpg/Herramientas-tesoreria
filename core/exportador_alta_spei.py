"""Generación del TXT de alta de cuentas SPEI (otros bancos) — ancho fijo.

Para las cuentas que NO son Bancomer (012), el alta se da por SPEI con un
formato distinto al de Bancomer: es el mismo de la dispersión (código de banco
+ CLABE + moneda + monto + nombre x2 + tipo de cuenta '40') seguido del correo.

Cada registro es una línea de 182 caracteres terminada en CRLF:

    pos   0 ancho  3 : código de banco (3 primeros dígitos de la CLABE)
    pos   3 ancho 18 : CLABE
    pos  21 ancho  3 : moneda -> 'MXP'
    pos  24 ancho 16 : monto, ceros a la izquierda, 2 decimales
    pos  40 ancho 30 : nombre del beneficiario
    pos  70 ancho 30 : nombre del beneficiario (repetido)
    pos 100 ancho  2 : tipo de cuenta -> '40' (CLABE)
    pos 102 ancho 80 : correo de notificación (izquierda, espacios)
"""

from __future__ import annotations

from .exportador import linea_registro as _linea_dispersion
from .exportador_alta_bancomer import _campo_email

FIN_LINEA = "\r\n"


def linea_registro(clabe: str, monto: float | None, beneficiario: str, correo: str) -> str:
    """Construye la línea de ancho fijo (182) para el alta SPEI de una cuenta."""
    # La parte inicial (102) es idéntica a la dispersión: el alias es el mismo
    # nombre del beneficiario. Después va el correo (80).
    return _linea_dispersion(clabe, monto, beneficiario, beneficiario) + _campo_email(correo)


def generar_txt(registros: list[tuple[str, float | None, str, str]]) -> str:
    """Genera el contenido completo del TXT de alta SPEI.

    Args:
        registros: lista de tuplas (clabe, monto, beneficiario, correo).
    """
    return "".join(linea_registro(*r) + FIN_LINEA for r in registros)
