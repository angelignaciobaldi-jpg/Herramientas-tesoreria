"""Lector del reporte 'Cuentas Bancarias' para conciliar con los estados de cuenta.

El reporte (Excel) trae, por cada proveedor, su CLABE (columna 'Cuenta'), el
nombre del beneficiario, una descripción, el banco, el RFC y el correo. Se usa
para complementar/validar los registros extraídos por OCR: si la CLABE de un
estado de cuenta aparece en el reporte, se toman de ahí el nombre y el correo,
que son la fuente autorizada.

El formato del archivo trae filas de metadatos arriba (título, fechas, etc.),
así que la fila de encabezados se localiza buscando las columnas 'Beneficiario'
y 'Cuenta' en lugar de asumir una posición fija.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

import openpyxl

# Palabras que no distinguen a un beneficiario (sufijos legales, artículos) y
# que se ignoran al comparar nombres para conciliar por nombre.
_RUIDO_NOMBRE = {
    "SA", "SAB", "SC", "RL", "CV", "DE", "DEL", "LA", "LAS", "LOS", "EL", "Y",
    "E", "I", "S", "A", "B", "C", "V", "COMPANY", "CIA", "GRUPO",
}


@dataclass
class CuentaReporte:
    clabe: str
    beneficiario: str
    descripcion: str
    banco: str
    rfc: str
    correo: str
    tipo: str = ""  # tipo de beneficiario (Proveedor/Deudor/Acreedor Diverso)


def _norm(valor) -> str:
    return str(valor).strip().lower() if valor is not None else ""


def _solo_digitos(valor) -> str:
    return re.sub(r"\D", "", str(valor)) if valor is not None else ""


def leer(ruta: str) -> dict[str, CuentaReporte]:
    """Lee el reporte y devuelve un diccionario {CLABE -> CuentaReporte}.

    Lanza ValueError si el Excel no tiene el formato esperado (sin columnas
    'Beneficiario' y 'Cuenta').
    """
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    try:
        ws = wb.active
        filas = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()

    # Localiza la fila de encabezados (la que tiene 'Beneficiario' y 'Cuenta')
    # y mapea cada título a su número de columna.
    col: dict[str, int] = {}
    idx_enc: int | None = None
    for i, fila in enumerate(filas):
        titulos = {_norm(v) for v in fila}
        if "beneficiario" in titulos and "cuenta" in titulos:
            idx_enc = i
            col = {_norm(v): j for j, v in enumerate(fila) if _norm(v)}
            break
    if idx_enc is None:
        raise ValueError(
            "El Excel no tiene el formato del reporte 'Cuentas Bancarias' "
            "(no se encontraron las columnas 'Beneficiario' y 'Cuenta')."
        )

    def campo(fila, *nombres: str) -> str:
        for n in nombres:
            j = col.get(n)
            if j is not None and j < len(fila) and fila[j] is not None:
                return str(fila[j]).strip()
        return ""

    catalogo: dict[str, CuentaReporte] = {}
    for fila in filas[idx_enc + 1:]:
        clabe = _solo_digitos(campo(fila, "cuenta"))
        if len(clabe) != 18:  # ignora totales, cuentas extranjeras y filas vacías
            continue
        catalogo[clabe] = CuentaReporte(
            clabe=clabe,
            beneficiario=campo(fila, "beneficiario"),
            descripcion=campo(fila, "descripción", "descripcion"),
            banco=campo(fila, "nombre banco"),
            rfc=campo(fila, "rfc"),
            correo=campo(fila, "correo"),
            tipo=campo(fila, "tipo beneficiario", "tipo de beneficiario", "tipo"),
        )
    return catalogo


# --- Conciliación por nombre --------------------------------------------
def _tokens_nombre(nombre: str) -> set[str]:
    """Normaliza un nombre a un conjunto de palabras significativas (sin acentos,
    mayúsculas, sin puntuación ni sufijos legales). El orden no importa, de modo
    que 'JORGE MONTERO GAZCON' y 'Montero Gazcon Jorge' den el mismo conjunto."""
    s = unicodedata.normalize("NFKD", nombre or "").encode("ascii", "ignore").decode()
    s = re.sub(r"[^A-Za-z0-9 ]", " ", s).upper()
    return {t for t in s.split() if len(t) > 1 and t not in _RUIDO_NOMBRE}


def buscar_por_nombre(catalogo: dict[str, CuentaReporte], nombre: str) -> CuentaReporte | None:
    """Busca en el reporte una cuenta cuyo nombre coincida con `nombre`.

    Compara por conjunto de palabras (independiente del orden) y exige una
    coincidencia fuerte (al menos 2 palabras en común y 80% de cobertura del
    nombre más corto). Devuelve None si no hay coincidencia clara o si hay
    empate entre cuentas distintas (para no adivinar)."""
    objetivo = _tokens_nombre(nombre)
    if len(objetivo) < 2:
        return None
    puntajes: dict[str, tuple[float, CuentaReporte]] = {}
    for cuenta in catalogo.values():
        mejor = 0.0
        for variante in (cuenta.beneficiario, cuenta.descripcion):
            toks = _tokens_nombre(variante)
            comunes = objetivo & toks
            if len(toks) < 2 or len(comunes) < 2:
                continue
            mejor = max(mejor, len(comunes) / min(len(objetivo), len(toks)))
        if mejor > 0:
            puntajes[cuenta.clabe] = (mejor, cuenta)
    if not puntajes:
        return None
    tope = max(s for s, _ in puntajes.values())
    if tope < 0.8:
        return None
    ganadores = [c for s, c in puntajes.values() if s == tope]
    return ganadores[0] if len(ganadores) == 1 else None
