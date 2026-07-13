"""Generación de archivos TXT de dispersión de DEVOLUCIONES.

Soporta dos formatos (el usuario elige el banco):

Banregio (separado por comas, 119 caracteres por línea):
    NNNNN,S,CLABE+2esp(20),monto(16),0.00(16),concepto(40),fecha(15)
    - NNNNN  : consecutivo de 5 dígitos (00001, 00002, ...)
    - CLABE  : 18 dígitos justificados a la izquierda, rellenado a 20 con espacios
    - monto  : 16 caracteres, ceros a la izquierda, 2 decimales
    - 0.00   : segundo importe, siempre cero
    - concepto: 40 caracteres, izquierda
    - fecha  : DDMMYYYY justificada a la derecha en 15 caracteres

Bancomer (ancho fijo, 131 caracteres por línea):
    PSC + CLABE_benef(18) + CLABE_origen(18) + MXP + monto(16) + nombre(30)
        + 40 + codigo_banco(3) + concepto(30 der.) + folio(8)

Ambos archivos usan salto de línea LF y SIN salto al final.
"""

from __future__ import annotations

import re
import unicodedata

FIN_LINEA = "\n"


def _ascii(texto: str, mayusculas: bool = False) -> str:
    s = unicodedata.normalize("NFKD", texto or "").encode("ascii", "ignore").decode()
    if mayusculas:
        s = s.upper()
    s = re.sub(r"[^A-Za-z0-9 ]", " ", s)
    return " ".join(s.split())


def banco_formato(texto: str) -> str | None:
    """Formato de layout según el banco que menciona `texto` (p. ej. el nombre de la
    cuenta de origen). Devuelve 'banregio' o 'bancomer', o None si el banco no tiene
    un formato de generación soportado por la app."""
    t = _ascii(texto, mayusculas=True)
    if "BANREGIO" in t:
        return "banregio"
    if "BBVA" in t or "BANCOMER" in t:
        return "bancomer"
    return None


def _monto16(monto: float | None) -> str:
    s = f"{float(monto or 0):.2f}"
    return s.rjust(16, "0")[-16:]


def _digitos(clabe: str) -> str:
    return re.sub(r"\D", "", clabe or "")


# ============================================================ Banregio
def linea_banregio(consecutivo: int, clabe: str, monto: float | None,
                   concepto: str, fecha_ddmmyyyy: str) -> str:
    return (
        f"{consecutivo:05d},S,"
        + _digitos(clabe).ljust(20)[:20] + ","
        + _monto16(monto) + ","
        + "0000000000000.00,"
        + _ascii(concepto)[:40].ljust(40) + ","
        + fecha_ddmmyyyy.strip().rjust(15)
    )


def generar_banregio(registros: list[tuple], fecha_ddmmyyyy: str) -> str:
    """registros: lista de (clabe, monto, beneficiario, concepto)."""
    lineas = [
        linea_banregio(i, clabe, monto, concepto, fecha_ddmmyyyy)
        for i, (clabe, monto, _benef, concepto) in enumerate(registros, start=1)
    ]
    return FIN_LINEA.join(lineas)


# ============================================================ Bancomer
def linea_bancomer(clabe_benef: str, clabe_origen: str, monto: float | None,
                   nombre: str, concepto: str, folio: str) -> str:
    cb = _digitos(clabe_benef)
    return (
        "PSC"
        + cb
        + _digitos(clabe_origen)
        + "MXP"
        + _monto16(monto)
        + _ascii(nombre)[:30].ljust(30)
        + "40"
        + cb[:3]
        + _ascii(concepto)[:30].rjust(30)
        + folio
    )


def generar_bancomer(registros: list[tuple], clabe_origen: str, folio: str) -> str:
    """registros: lista de (clabe, monto, beneficiario, concepto)."""
    lineas = [
        linea_bancomer(clabe, clabe_origen, monto, benef, concepto, folio)
        for (clabe, monto, benef, concepto) in registros
    ]
    return FIN_LINEA.join(lineas)
