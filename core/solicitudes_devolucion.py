"""Solicitudes de devolución autorizadas del SIPP (consulta a la API real).

Consulta el microservicio (core/api.solicitudes_devolucion) por empresa —usando
su id— paginando para traer todas, y mapea el JSON del servicio a la dataclass
`SolicitudDevolucion`, que es lo que consume la pantalla de devoluciones.

Las empresas y sus IDs NO viven aquí: la fuente única es `ui/comun.EMPRESAS`. La
pantalla resuelve nombre -> id y llama a `consultar()` con el id (así este módulo
de core no depende de la interfaz).
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from core import api
from core.catalogo_bancos import banco_desde_clabe


@dataclass
class SolicitudDevolucion:
    """Una solicitud de devolución de un cliente, tal como la entrega el SIPP."""

    folio: str          # folio de la solicitud en el SIPP (id_Solicitud)
    empresa: str        # empresa que registró la solicitud
    cliente: str        # beneficiario de la devolución (a quien se le paga)
    clabe: str          # CLABE destino (18 dígitos)
    banco: str          # banco del beneficiario (informativo)
    monto: float        # importe a devolver
    concepto: str       # concepto / referencia del pago
    fecha: str          # fecha de la solicitud (DD/MM/AAAA)
    estatus: str = "Autorizada"


# --- Mapeo del JSON del servicio a la dataclass --------------------------
def _texto(valor) -> str:
    return "" if valor is None else str(valor).strip()


def _monto(valor) -> float:
    try:
        return float(valor)
    except (TypeError, ValueError):
        return 0.0


def _fecha(valor) -> str:
    """'2026-02-11T11:28:09.910Z' -> '11/02/2026'. Si no parsea, deja el texto."""
    s = _texto(valor)
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s)
    return f"{m.group(3)}/{m.group(2)}/{m.group(1)}" if m else s


def _mapear(registro: dict, empresa: str) -> SolicitudDevolucion:
    """Convierte un registro del servicio en una SolicitudDevolucion. El `empresa`
    (nombre legible) se pasa aparte porque el servicio se consulta por id y no lo
    devuelve en cada registro."""
    return SolicitudDevolucion(
        folio=_texto(registro.get("id_Solicitud")),
        empresa=empresa,
        # El beneficiario es el cliente (razón social); 'nb_Solicitante' es el
        # empleado interno que la registró, no a quien se le paga.
        cliente=_texto(registro.get("de_Cliente_RazonSocial")
                       or registro.get("nb_Solicitante")),
        clabe=re.sub(r"\D", "", _texto(registro.get("nu_CuentaBancaria")))[:18],
        banco=_texto(registro.get("nb_Banco")),
        monto=_monto(registro.get("im_ImporteDevolucion")),
        concepto=_texto(registro.get("de_Comentarios")),
        fecha=_fecha(registro.get("fh_Solicitud")),
        estatus=_texto(registro.get("de_EstatusSolicitud")) or "Autorizada",
    )


# --- Consulta ------------------------------------------------------------
def consultar(
    id_empresa: int | str,
    empresa: str = "",
    fecha_inicio: str | None = None,
    fecha_fin: str | None = None,
) -> list[SolicitudDevolucion]:
    """Trae TODAS las solicitudes de devolución de una empresa (por su id),
    paginando el servicio hasta agotar las páginas.

    - `id_empresa`: identificador de la empresa (idEmpresa que espera la API).
    - `empresa`: nombre legible para etiquetar los resultados (no lo devuelve la
      API; lo pone la pantalla desde el id que consultó).
    - `fecha_inicio` / `fecha_fin`: rango 'YYYY-MM-DD' opcional.

    Propaga las excepciones de `core.api` (ApiSinConexion / ErrorRespuestaApi) si
    falla; la pantalla decide cómo avisarle al usuario.
    """
    solicitudes: list[SolicitudDevolucion] = []
    page = 1
    while True:
        respuesta = api.solicitudes_devolucion(
            id_empresa, fecha_inicio=fecha_inicio, fecha_fin=fecha_fin,
            page=page, page_size=100,
        )
        cuerpo = respuesta if isinstance(respuesta, dict) else {}
        datos = cuerpo.get("data") or []
        for registro in datos:
            if isinstance(registro, dict):
                solicitudes.append(_mapear(registro, empresa))
        meta = cuerpo.get("meta") or {}
        total_paginas = meta.get("totalPages") or 1
        if not datos or page >= total_paginas:
            break
        page += 1
    return solicitudes


# --- Deudores diversos (colaboradores / otros deudores) ------------------
def _mapear_deudor(registro: dict, empresa: str) -> SolicitudDevolucion:
    """Convierte un registro de deudor diverso en una SolicitudDevolucion.

    El folio y el concepto NO vienen del servicio para estos pagos: quedan
    vacíos y EDITABLES en la pantalla. El banco se deduce de la CLABE."""
    # El servicio manda la CLABE en 'nu_ClabeInterBancaria' (con B mayúscula); se
    # aceptan variantes por si acaso.
    clabe = re.sub(r"\D", "", _texto(
        registro.get("nu_ClabeInterBancaria")
        or registro.get("nu_ClabeInterbancaria")
        or registro.get("un_ClabeInterbancaria")))[:18]
    return SolicitudDevolucion(
        folio="",  # editable en la pantalla para estos tipos de pago
        empresa=empresa,
        cliente=_texto(registro.get("nb_Empleado") or registro.get("de_DeudorDiverso")),
        clabe=clabe,
        banco=banco_desde_clabe(clabe) or "",
        monto=_monto(registro.get("im_Total")),
        concepto=_texto(registro.get("de_Comentarios") or ""),  # editable
        fecha=_fecha(registro.get("fh_Solicitud") or ""),
        estatus="Deudor diverso",
    )


def consultar_deudores(
    id_empresa: int | str,
    empresa: str = "",
    tipo: int = 1,
) -> list[SolicitudDevolucion]:
    """Trae los pagos/devoluciones pendientes a deudores diversos de una empresa.

    - `id_empresa`: id de la empresa (idEmpresa que espera la API).
    - `empresa`: nombre legible para etiquetar los resultados.
    - `tipo`: 1 = 'Devoluciones pendientes', 0 = 'Pagos pendientes'.

    Pagina hasta agotar. Propaga las excepciones de core.api si falla.
    """
    solicitudes: list[SolicitudDevolucion] = []
    page = 1
    while True:
        respuesta = api.pagos_deudores_diversos(
            id_empresa, tipo, page=page, page_size=100)
        cuerpo = respuesta if isinstance(respuesta, dict) else {}
        datos = cuerpo.get("data") or []
        for registro in datos:
            if isinstance(registro, dict):
                solicitudes.append(_mapear_deudor(registro, empresa))
        meta = cuerpo.get("meta") or {}
        total_paginas = meta.get("totalPages") or 1
        if not datos or page >= total_paginas:
            break
        page += 1
    return solicitudes
