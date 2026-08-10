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

Bancomer / BBVA (ancho fijo). El prefijo distingue el tipo de pago según si la
cuenta DESTINO es del MISMO banco que la de origen o de OTRO banco.

    Otro banco (SPEI) -> PSC, 131 caracteres. Layout oficial de BBVA Net Cash
    ("Layout Importación de Grupos - Traspasos y/o Pagos Interbancarios",
    act. 06/07/2019). Las posiciones del PDF son las del layout base (128) y aquí
    van corridas 3 por la clave de pago que exige el archivo MIXTO:

        pos    campo (nombre del PDF)             tipo  long
        1-3    Clave de Pago Bnc                  AL     3   'PSC'
        4-21   Asunto Beneficiario                N     18   CLABE destino
        22-39  Asunto Ordenante                   N     18   CLABE cargo
        40-42  Divisa de la Operación             AL     3   'MXP'
        43-58  Importe de la Operación            M     16   ceros izq., ####.dd
        59-88  Titular Asunto Beneficiario        AL    30   izq., espacios
        89-90  Tipo de Cuenta                     N      2   '40' CLABE / '03' débito
        91-93  Número de Banco del Beneficiario   N      3   oficial Banxico
        94-123 Motivo de Pago                     AL    30
        124-130 Referencia Numérica               N      7   ceros izq.
        131    Disponibilidad                     AL     1   'H' SPEI / 'M' CECOBAN

    Mismo banco (Bancomer 012) -> PTC, 88 caracteres (sin nombre, tipo, código,
    referencia ni disponibilidad). OJO: ese layout NO viene en el PDF de
    interbancarios; se dedujo de un archivo de referencia y está sin verificar
    contra documentación oficial.
        PTC + CLABE_benef(18) + CLABE_origen(18) + MXP + monto(16) + concepto(30 der.)

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


def _texto(texto: str) -> str:
    """Sanea nombre/concepto para el layout del banco: conserva letras (con acentos),
    dígitos, espacio y los signos , . - / que el banco admite (p. ej. 'S.A. DE C.V.'
    o 'Z5728847,Z6419734'); colapsa espacios. Lo no representable en latin-1 —cómo se
    escribe el archivo— se normaliza a ASCII para no romper el guardado."""
    s = "".join(
        ch if (ch.isalnum() or ch in " ,.-/") else " "
        for ch in (texto or "")
    )
    try:
        s.encode("latin-1")
    except UnicodeEncodeError:
        s = unicodedata.normalize("NFKD", s).encode("latin-1", "ignore").decode("latin-1")
    # Solo se recortan los extremos; los espacios internos se conservan tal cual
    # (el layout del banco no los colapsa).
    return s.strip()


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


def _referencia7(folio: str) -> str:
    """Campo 9 del layout BBVA, 'Referencia Numérica': 7 dígitos, ceros a la
    izquierda ('762' -> '0000762'). Es de ancho fijo; sin rellenar, la línea sale
    corta y el banco rechaza el archivo."""
    return _digitos(folio).rjust(7, "0")[-7:]


# Campo 10 del layout BBVA, 'Disponibilidad' (1 posición, cierra el registro):
#   H = mismo día vía SPEI      M = día siguiente vía CECOBAN
# Los pagos interbancarios de la app salen el mismo día.
DISPONIBILIDAD_MISMO_DIA = "H"
DISPONIBILIDAD_DIA_SIGUIENTE = "M"


# Código de banco (3 primeros dígitos de la CLABE) de Bancomer, el propio banco
# que dispersa: si el destino también es 012, el pago es a MISMO banco.
BANCO_BANCOMER = "012"


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
    base = (
        cb
        + _digitos(clabe_origen)
        + "MXP"
        + _monto16(monto)
    )
    concepto_fmt = _texto(concepto)[:30].rjust(30)
    if cb[:3] == BANCO_BANCOMER:
        # Mismo banco (Bancomer): pago a terceros PTC, formato corto (sin nombre,
        # tipo de cuenta, código de banco ni folio).
        return "PTC" + base + concepto_fmt
    # Otro banco: pago SPEI PSC, formato completo (131 caracteres).
    return (
        "PSC"
        + base
        + _texto(nombre)[:30].ljust(30)
        + "40"
        + cb[:3]
        + concepto_fmt
        + _referencia7(folio)
        + DISPONIBILIDAD_MISMO_DIA
    )


def generar_bancomer(registros: list[tuple], clabe_origen: str, folio: str) -> str:
    """registros: lista de (clabe, monto, beneficiario, concepto)."""
    lineas = [
        linea_bancomer(clabe, clabe_origen, monto, benef, concepto, folio)
        for (clabe, monto, benef, concepto) in registros
    ]
    return FIN_LINEA.join(lineas)
