"""Generación del archivo de ALTA de beneficiarios para el portal Banregio.

Replica la estructura del archivo 'ALTABANREGIO 1.xls' (formato Excel .xls):

    Cuenta | NumBanco | NomBanco | Descripción | RFC | Correo

  - Cuenta      : CLABE de 18 dígitos
  - NumBanco    : 3 primeros dígitos de la CLABE (código de banco)
  - NomBanco    : nombre del banco según ese código
  - Descripción : nombre del beneficiario
  - RFC         : (no se captura; se deja vacío)
  - Correo      : email de notificación
"""

from __future__ import annotations

import re
import unicodedata

import xlwt

# Código de banco (3 dígitos de la CLABE) -> nombre para el portal Banregio.
NOMBRES_BANCO = {
    "002": "BANAMEX", "006": "BANCOMEXT", "009": "BANOBRAS", "012": "BANCOMER",
    "014": "SANTANDER", "019": "BANJERCITO", "021": "HSBC", "030": "BANBAJIO",
    "036": "INBURSA", "042": "MIFEL", "044": "SCOTIABANK", "058": "BANREGIO",
    "059": "INVEX", "060": "BANSI", "062": "AFIRME", "072": "BANORTE",
    "106": "BANK OF AMERICA", "108": "MUFG", "110": "JP MORGAN", "112": "MONEX",
    "113": "VE POR MAS", "124": "CITI MEXICO", "127": "AZTECA", "128": "INTERCAM", "129": "BARCLAYS",
    "130": "COMPARTAMOS", "132": "MULTIVA", "133": "ACTINVER", "137": "BANCOPPEL",
    "138": "ABC CAPITAL", "140": "CONSUBANCO", "143": "CIBANCO", "147": "BANKAOOL",
    "148": "PAGATODO", "150": "INMOBILIARIO", "151": "DONDE", "152": "BANCREA",
    "154": "BANCO COVALTO", "155": "ICBC", "156": "SABADELL", "166": "BIENESTAR",
    "168": "HIPOTECARIA FEDERAL", "600": "MONEXCB", "601": "GBM", "602": "MASARI",
    "605": "VALUE", "608": "VECTOR", "616": "FINAMEX", "617": "VALMEX",
    "620": "PROFUTURO", "630": "INTERCAM", "631": "CI BOLSA", "634": "FINCOMUN",
    "646": "STP", "652": "CREDICAPITAL", "653": "KUSPIT", "656": "UNAGRA",
    "659": "ASP INTEGRA", "670": "LIBERTAD", "677": "CAJA POP MEXICANA",
    "684": "OPM", "706": "ARCUS", "710": "NVIO", "722": "MERCADO PAGO",
    "723": "CUENCA", "728": "SPIN BY OXXO",
}


def nombre_banco(clabe: str) -> str:
    return NOMBRES_BANCO.get(clabe[:3], "") if clabe else ""


def _ascii(texto: str) -> str:
    s = unicodedata.normalize("NFKD", texto or "").encode("ascii", "ignore").decode()
    return " ".join(s.upper().split())


def generar(ruta: str, registros: list[tuple], hoja: str = "ALTA") -> None:
    """Crea el archivo .xls de alta.

    Args:
        ruta: ruta destino .xls
        registros: lista de (clabe, beneficiario, email)
        hoja: nombre de la hoja
    """
    wb = xlwt.Workbook(encoding="utf-8")
    ws = wb.add_sheet(hoja)

    encabezados = ["Cuenta", "NumBanco", "NomBanco", "Descripción", "RFC", "Correo"]
    for col, titulo in enumerate(encabezados):
        ws.write(0, col, titulo)

    for i, (clabe, beneficiario, email) in enumerate(registros, start=1):
        clabe = re.sub(r"\D", "", clabe or "")
        ws.write(i, 0, clabe)                  # Cuenta (CLABE como texto)
        ws.write(i, 1, clabe[:3])              # NumBanco
        ws.write(i, 2, nombre_banco(clabe))    # NomBanco
        ws.write(i, 3, _ascii(beneficiario))   # Descripción (beneficiario)
        ws.write(i, 4, "")                      # RFC (no se captura)
        ws.write(i, 5, (email or "").strip())  # Correo

    wb.save(ruta)
