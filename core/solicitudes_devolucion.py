"""Solicitudes de devolución autorizadas (consulta al SIPP).

La conexión con el SIPP todavía está en construcción, así que este módulo sirve
solicitudes FICTICIAS para poder desarrollar y probar el flujo completo:

    consultar por empresa(s) -> seleccionar -> asignar cuenta origen de pago
    -> generar un TXT por empresa origen

Cuando la API del SIPP esté lista, basta reimplementar `empresas()` y
`consultar()` para que hablen con el portal: el resto de la herramienta trabaja
contra la dataclass `SolicitudDevolucion`, así que la interfaz NO cambia.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class SolicitudDevolucion:
    """Una solicitud de devolución de un cliente, tal como la entrega el SIPP."""

    folio: str          # folio de la solicitud en el SIPP
    empresa: str        # empresa que registró la solicitud
    cliente: str        # beneficiario de la devolución (a quien se le paga)
    clabe: str          # CLABE destino (18 dígitos)
    banco: str          # banco del beneficiario (informativo)
    monto: float        # importe a devolver
    concepto: str       # concepto / referencia del pago
    fecha: str          # fecha de la solicitud (DD/MM/AAAA)
    estatus: str = "Autorizada"


def _con_digito(pre17: str) -> str:
    """Completa una CLABE de 17 dígitos con su dígito de control, para que los
    datos ficticios pasen las mismas validaciones que los reales."""
    pesos = (3, 7, 1)
    suma = sum((int(d) * pesos[i % 3]) % 10 for i, d in enumerate(pre17[:17]))
    return pre17[:17] + str((10 - (suma % 10)) % 10)


# ===================== DATOS FICTICIOS (placeholder del SIPP) =====================
# TODO(SIPP): sustituir por la consulta real a la API. Mientras tanto, esta
# semilla permite probar el flujo con varias empresas y montos distintos.
_EMPRESAS = [
    "ABASTECEDORA DEL NORTE",
    "COMBUSTIBLES DEL PACIFICO",
    "PETROIL SERVICIOS",
    "QUETZALTIC SOLUTIONS",
]

# (folio, índice de empresa, cliente, CLABE de 17 díg., monto, concepto, fecha)
# El banco NO se captura aquí: se deduce de los 3 primeros dígitos de la CLABE
# (012=BBVA, 014=Santander, 002=Banamex, 021=HSBC, 044=Scotiabank, 058=Banregio,
# 072=Banorte), igual que hace la herramienta. Así los datos son coherentes.
_SEMILLA = [
    ("DEV-2026-0001", 0, "ARMANDO LOPEZ RODRIGUEZ",   "01218001523004947", 12500.00, "Devolucion saldo a favor",      "02/07/2026"),
    ("DEV-2026-0002", 0, "MARIA FERNANDA SOTO CRUZ",  "01418004512300981",  3480.50, "Devolucion pago duplicado",     "02/07/2026"),
    ("DEV-2026-0003", 0, "GRUPO INDUSTRIAL DEL BAJIO","07218007700123455", 58900.00, "Devolucion garantia",           "03/07/2026"),
    ("DEV-2026-0004", 1, "JOSE ANTONIO MEZA GARZA",   "05818009900456781",  9750.25, "Devolucion anticipo",           "03/07/2026"),
    ("DEV-2026-0005", 1, "TRANSPORTES DEL GOLFO SA",  "00218001100998877", 21300.00, "Devolucion flete cancelado",    "04/07/2026"),
    ("DEV-2026-0006", 1, "LAURA PATRICIA RIOS DIAZ",  "01218003400221199",  1580.75, "Devolucion saldo a favor",      "04/07/2026"),
    ("DEV-2026-0007", 2, "COMERCIALIZADORA ANAHUAC",  "01418006600334422", 44100.00, "Devolucion nota de credito",    "05/07/2026"),
    ("DEV-2026-0008", 2, "RICARDO ELIAS MONTES VEGA", "02118002200776655",  6200.00, "Devolucion deposito garantia",  "05/07/2026"),
    ("DEV-2026-0009", 2, "SERVICIOS LOGISTICOS OMEGA","04418005500887711", 17840.90, "Devolucion pago en exceso",     "06/07/2026"),
    ("DEV-2026-0010", 3, "ANA KAREN BARRIOS SALDANA", "01218007700112233",  2990.00, "Devolucion saldo a favor",      "06/07/2026"),
    ("DEV-2026-0011", 3, "DISTRIBUIDORA SIERRA ALTA", "00218008800556644", 33250.40, "Devolucion cancelacion pedido", "07/07/2026"),
    ("DEV-2026-0012", 3, "MIGUEL ANGEL TORRES LUNA",  "07218002200998811",  7410.00, "Devolucion anticipo",           "07/07/2026"),
]


def _ficticias() -> list[SolicitudDevolucion]:
    from .catalogo_bancos import banco_desde_clabe  # import local: evita ciclos

    solicitudes = []
    for folio, idx_emp, cliente, pre17, monto, concepto, fecha in _SEMILLA:
        clabe = _con_digito(pre17)
        solicitudes.append(
            SolicitudDevolucion(
                folio=folio,
                empresa=_EMPRESAS[idx_emp],
                cliente=cliente,
                clabe=clabe,
                banco=banco_desde_clabe(clabe) or "",
                monto=monto,
                concepto=concepto,
                fecha=fecha,
            )
        )
    return solicitudes


# ============================== API del módulo ==============================
def empresas() -> list[str]:
    """Empresas que tienen solicitudes para consultar (para el filtro)."""
    return sorted({s.empresa for s in _ficticias()})


def consultar(empresas_sel: list[str] | None = None) -> list[SolicitudDevolucion]:
    """Solicitudes AUTORIZADAS de las empresas indicadas (todas si viene vacío).

    PLACEHOLDER: hoy devuelve las solicitudes ficticias. Cuando la API del SIPP
    exista, esta función hará la consulta real y devolverá las mismas dataclasses,
    sin que la interfaz tenga que cambiar.
    """
    autorizadas = [
        s for s in _ficticias() if s.estatus.strip().lower().startswith("autoriz")
    ]
    if not empresas_sel:
        return autorizadas
    quiero = {e.strip().lower() for e in empresas_sel}
    return [s for s in autorizadas if s.empresa.strip().lower() in quiero]
