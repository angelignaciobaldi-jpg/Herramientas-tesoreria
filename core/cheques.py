"""Modelo de dominio de cheques y preparación de sus datos de impresión.

Separa el "qué imprimir" de la interfaz. Trabaja con LISTAS de cheques desde el
diseño: la UI de la primera vuelta captura uno solo, pero el flujo que venga del
RPA (una solicitud autorizada -> varios cheques) reutilizará `preparar_cheques`
sin cambios.

La impresión real (posicionar los campos sobre el cheque físico) NO se resuelve
aquí todavía; `LAYOUTS_BANCO` es el catálogo del selector y se ampliará con las
coordenadas por banco cuando se cablee `core/impresion.py`.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass

from . import numero_letras

# Coordenadas de impresión (en MILÍMETROS, origen = esquina superior izquierda de
# la hoja) de cada campo, por layout de banco. Convención de cada campo:
#   x, y        -> posición; Y es el BORDE SUPERIOR del texto (imprime hacia abajo).
#   alineacion  -> ancla horizontal: 'izq' (x = borde izq), 'der' (x = borde der),
#                  'centro' (x = centro del texto).
#   ancho_max   -> ancho máximo (mm) para recortar/ajustar textos largos.
# Además, cada layout admite una llave opcional (NO es un campo):
#   margen_superior -> mm que BAJA todo el cheque en la impresión vertical (más
#                      margen = más abajo; menos = más arriba). Sirve para acomodar
#                      el layout sobre el cheque físico. Default 0 si no se indica.
# Se miden con la hoja de calibración (Configuración -> Impresión de cheques). Las
# consumirá la impresión real (aún por cablear con core/impresion.py).
LAYOUTS: dict[str, dict] = {
    "BBVA Bancomer": {
        "margen_superior": 10,
        "beneficiario": {"x": 13, "y": 27, "alineacion": "izq", "ancho_max": 110},
        "beneficiario_2": {"x": 13, "y": 25, "alineacion": "izq", "ancho_max": 110},
        "fecha": {"x": 95, "y":19, "alineacion": "izq", "ancho_max": 30},
        "monto_numero": {"x": 160, "y": 27, "alineacion": "der", "ancho_max": 30},
        "monto_letras": {"x": 18, "y": 35, "alineacion": "izq", "ancho_max": 117},
        "monto_letras_2": {"x": 18, "y": 33, "alineacion": "izq", "ancho_max": 117},
    },
    "Banregio": {
        "margen_superior": 10,
        "beneficiario": {"x": 10, "y": 27, "alineacion": "izq", "ancho_max": 100},
        "beneficiario_2": {"x": 10, "y": 25, "alineacion": "izq", "ancho_max": 100},
        "fecha": {"x": 155, "y": 14, "alineacion": "der", "ancho_max": 50},
        "monto_numero": {"x": 155, "y": 29, "alineacion": "der", "ancho_max": 30},
        "monto_letras": {"x": 10, "y": 37, "alineacion": "izq", "ancho_max": 145},
        "monto_letras_2": {"x": 10, "y": 35, "alineacion": "izq", "ancho_max": 145},
    },
}

# Bancos disponibles en el selector: los que tienen un layout definido.
LAYOUTS_BANCO = list(LAYOUTS.keys())


def layout_banco(banco: str) -> dict[str, dict] | None:
    """Coordenadas de impresión del banco `banco` (o None si no tiene layout)."""
    return LAYOUTS.get(banco)


@dataclass
class Cheque:
    """Datos de un cheque a imprimir (capturados o tomados de una solicitud)."""

    beneficiario: str
    monto: float
    fecha: datetime.date
    moneda: str = "MXN"


@dataclass
class DatosImpresionCheque:
    """Información ya formateada, lista para mostrarse/imprimirse."""

    beneficiario: str
    monto_texto: str      # p. ej. "$1,234.50"
    monto_letras: str     # p. ej. "UN MIL ... PESOS 50/100 M.N."
    fecha_texto: str      # "DD/MM/AAAA"
    moneda: str
    banco: str


def _fmt_monto(monto: float | None) -> str:
    """Monto como moneda con 2 decimales (p. ej. "$1,234.50")."""
    return f"${(monto or 0):,.2f}"


def preparar_cheque(cheque: Cheque, banco: str) -> DatosImpresionCheque:
    """Arma los datos de impresión de UN cheque (monto en número y en letra,
    fecha formateada y layout de banco)."""
    return DatosImpresionCheque(
        beneficiario=cheque.beneficiario.strip().upper(),
        monto_texto=_fmt_monto(cheque.monto),
        monto_letras=numero_letras.monto_en_letras(cheque.monto, cheque.moneda),
        fecha_texto=cheque.fecha.strftime("%d/%m/%Y"),
        moneda=cheque.moneda,
        banco=banco,
    )


def preparar_cheques(
    cheques: list[Cheque], banco: str) -> list[DatosImpresionCheque]:
    """Prepara los datos de impresión de una LISTA de cheques (todos con el mismo
    layout de banco). Punto único que consumen la UI y, a futuro, el flujo del RPA."""
    return [preparar_cheque(c, banco) for c in cheques]
