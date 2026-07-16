"""Conversión de números a letra (español) para cheques y documentos.

Genera la representación en palabras de un monto, en MAYÚSCULAS y con el formato
clásico de cheque mexicano, p. ej.:

    1234.50  ->  "UN MIL DOSCIENTOS TREINTA Y CUATRO PESOS 50/100 M.N."

Notas de formato (decisiones del negocio):
  - El millar exacto se escribe "UN MIL" (no el estándar "MIL"), como en el
    ejemplo del cheque.
  - Los centavos van como fracción de dos dígitos "NN/100".
  - La unidad final se apocopa antes del sustantivo: 1 -> "UN PESO",
    21 -> "VEINTIÚN PESOS", 31 -> "TREINTA Y UN PESOS".

El módulo es puro (sin dependencias externas) para poder probarlo y reutilizarlo.
"""

from __future__ import annotations

from decimal import ROUND_HALF_UP, Decimal

# Unidades 0..29 (los "dieci-"/"veinti-" son una sola palabra en español).
_UNIDADES = [
    "CERO", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO",
    "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISÉIS",
    "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIUNO", "VEINTIDÓS",
    "VEINTITRÉS", "VEINTICUATRO", "VEINTICINCO", "VEINTISÉIS", "VEINTISIETE",
    "VEINTIOCHO", "VEINTINUEVE",
]
# Decenas exactas (30, 40, ... 90); se unen con "Y" a la unidad.
_DECENAS = {
    30: "TREINTA", 40: "CUARENTA", 50: "CINCUENTA", 60: "SESENTA",
    70: "SETENTA", 80: "OCHENTA", 90: "NOVENTA",
}
# Centenas exactas (100 se maneja aparte: CIEN vs CIENTO).
_CENTENAS = {
    1: "CIENTO", 2: "DOSCIENTOS", 3: "TRESCIENTOS", 4: "CUATROCIENTOS",
    5: "QUINIENTOS", 6: "SEISCIENTOS", 7: "SETECIENTOS", 8: "OCHOCIENTOS",
    9: "NOVECIENTOS",
}

# Sustantivo (singular, plural) y sufijo por moneda.
_MONEDAS = {
    "MXN": ("PESO", "PESOS", "M.N."),
    "USD": ("DÓLAR", "DÓLARES", "USD"),
}


def _centenas(n: int) -> str:
    """Palabras de un número de 0..999 (sin apócope)."""
    if n == 0:
        return ""
    if n == 100:
        return "CIEN"
    partes: list[str] = []
    centena, resto = divmod(n, 100)
    if centena:
        partes.append(_CENTENAS[centena])
    if resto:
        if resto < 30:
            partes.append(_UNIDADES[resto])
        else:
            decena, unidad = divmod(resto, 10)
            if unidad:
                partes.append(f"{_DECENAS[decena * 10]} Y {_UNIDADES[unidad]}")
            else:
                partes.append(_DECENAS[decena * 10])
    return " ".join(partes)


def numero_a_letras(entero: int) -> str:
    """Convierte un entero (0..999,999,999) a palabras en MAYÚSCULAS.

    El millar exacto se escribe "UN MIL" (formato de cheque). Lanza ValueError si
    el número es negativo o excede el rango soportado.
    """
    if entero < 0:
        raise ValueError("El número no puede ser negativo.")
    if entero > 999_999_999:
        raise ValueError("El número excede el rango soportado (máx. 999,999,999).")
    if entero == 0:
        return "CERO"

    millones, resto = divmod(entero, 1_000_000)
    miles, cientos = divmod(resto, 1_000)
    partes: list[str] = []

    if millones:
        if millones == 1:
            partes.append("UN MILLÓN")
        else:  # apócope antes del sustantivo: "VEINTIÚN MILLONES"
            partes.append(f"{_apocopar_unidad(numero_a_letras(millones))} MILLONES")
    if miles:
        # "UN MIL" (decisión de negocio), "DOS MIL", ... "DOSCIENTOS MIL". La
        # unidad se apocopa antes de "MIL": 1 -> "UN", 21 -> "VEINTIÚN".
        partes.append(f"{_apocopar_unidad(numero_a_letras(miles))} MIL")
    if cientos:
        partes.append(_centenas(cientos))
    return " ".join(partes)


def _apocopar_unidad(texto: str) -> str:
    """Apócope de la unidad final antes de un sustantivo masculino: 'UNO' -> 'UN'
    y 'VEINTIUNO' -> 'VEINTIÚN' (con acento). No toca 'TREINTA', etc."""
    if texto == "UNO":
        return "UN"
    if texto.endswith("VEINTIUNO"):
        return texto[:-len("VEINTIUNO")] + "VEINTIÚN"
    if texto.endswith(" UNO"):  # p. ej. "TREINTA Y UNO" -> "TREINTA Y UN"
        return texto[:-len(" UNO")] + " UN"
    return texto


def monto_en_letras(monto: float, moneda: str = "MXN") -> str:
    """Monto con letra en formato de cheque, p. ej.:

        1234.50 -> "UN MIL DOSCIENTOS TREINTA Y CUATRO PESOS 50/100 M.N."

    `moneda` es "MXN" (PESOS … M.N.) o "USD" (DÓLARES … USD). Los centavos van
    como "NN/100". Lanza ValueError si el monto es negativo o la moneda no existe.
    """
    if monto < 0:
        raise ValueError("El monto no puede ser negativo.")
    singular, plural, sufijo = _MONEDAS.get(
        (moneda or "").upper(), _MONEDAS["MXN"])

    # Redondeo a 2 decimales HALF-UP con Decimal (correcto para dinero y sin el
    # sesgo binario del float: 12.345 -> "12.35"). str(monto) da el decimal exacto.
    centavos_totales = int(
        (Decimal(str(monto)).quantize(Decimal("0.01"), ROUND_HALF_UP)) * 100)
    entero, centavos = divmod(centavos_totales, 100)

    letras = _apocopar_unidad(numero_a_letras(entero))
    sustantivo = singular if entero == 1 else plural
    return f"{letras} {sustantivo} {centavos:02d}/100 {sufijo}"
