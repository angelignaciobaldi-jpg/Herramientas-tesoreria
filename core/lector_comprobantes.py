"""Lectura LOCAL de comprobantes de pago BBVA Net Cash (sin red ni tokens).

Complementa a `core.api.leer_comprobantes_pagos` (el extractor remoto). Existe por
dos razones:

  1. El extractor devuelve `cuenta_origen`, `cuenta_destino` e `importe`, pero NO
     la **fecha de aplicación**, que es con la que se arma la Referencia que el RPA
     escribe en el SIPP (formato AAAAMMDD).
  2. No depende de red ni de credenciales, así que el casado sigue funcionando
     aunque el token del extractor caduque.

Devuelve la MISMA forma que el extractor (`documento_lectura`, `cuenta_origen`,
`cuenta_destino`, `importe`) para que `core.comprobantes` case igual venga de donde
venga la lectura, y agrega los campos propios del comprobante.

Los comprobantes de BBVA Net Cash traen capa de texto (no hacen falta OCR ni
rasterizado). Vienen en dos variantes con campos distintos:

  - *Pago Mismo Banco*: trae 'Motivo de pago' y 'Folio único'. NO trae 'Referencia'
    ni 'Clave de rastreo'.
  - *Pago Interbancario*: además trae 'Referencia', 'Clave de rastreo',
    'Folio interbancario' y 'Banco beneficiario'.

Por eso los campos exclusivos del interbancario se devuelven vacíos en el otro
caso, y NINGUNO se usa para casar: las 3 reglas se apoyan solo en lo que traen las
dos variantes (cuentas e importe).

Ojo con la cuenta de retiro: viene como NÚMERO DE CUENTA con ceros a la izquierda
('000000000117421184'), no como CLABE. Ver `core.comprobantes.Objetivo.origenes`.
"""

from __future__ import annotations

import os
import re
import unicodedata

import pymupdf

# Etiquetas del comprobante -> clave del dict. Se buscan SIN acentos ni caja (el
# PDF trae 'depósito'/'aplicación'), y el valor es lo que sigue a ':' en la línea.
_ETIQUETAS = {
    "tipo de operacion": "tipo_operacion",
    "descripcion": "descripcion",
    "importe": "importe",
    "cuenta de retiro": "cuenta_origen",
    "cuenta de deposito": "cuenta_destino",
    "fecha de aplicacion": "fecha_aplicacion",
    "fecha de creacion": "fecha_creacion",
    "referencia": "referencia",
    "clave de rastreo": "clave_rastreo",
    "folio interbancario": "folio_interbancario",
    "folio de firma": "folio_firma",
    "folio unico": "folio_unico",
    "banco beneficiario": "banco_beneficiario",
    "concepto de pago": "concepto_pago",
    "motivo de pago": "motivo_pago",
    "estado": "estado",
}

# Un comprobante se considera aplicado solo con este estado; cualquier otro
# ('En proceso', 'Cancelado'…) no debería subirse al SIPP como pago hecho.
ESTADO_APLICADO = "operado"


class ErrorLectura(Exception):
    """No se pudo leer el comprobante (PDF ilegible, protegido o sin texto)."""


def _sin_acentos(texto: str) -> str:
    base = unicodedata.normalize("NFKD", texto or "")
    return "".join(c for c in base if not unicodedata.combining(c))


def _clave_etiqueta(texto: str) -> str:
    """Normaliza una etiqueta para buscarla en `_ETIQUETAS`."""
    return " ".join(_sin_acentos(texto).lower().split())


def _a_importe(texto: str) -> float | None:
    """'3,227.00' -> 3227.0. None si no parece un importe."""
    limpio = re.sub(r"[^\d.]", "", (texto or "").replace(",", ""))
    try:
        return float(limpio) if limpio else None
    except ValueError:
        return None


def _campos_de_texto(texto: str) -> dict:
    """Extrae los pares 'Etiqueta: valor' de la capa de texto de una página.

    Se queda con la PRIMERA aparición de cada etiqueta: 'Titular de la cuenta'
    aparece dos veces (retiro y depósito) y el resto una sola, así que quedarse con
    la primera evita que un valor pise a otro."""
    campos: dict[str, str] = {}
    for linea in texto.splitlines():
        if ":" not in linea:
            continue
        etiqueta, _, valor = linea.partition(":")
        clave = _ETIQUETAS.get(_clave_etiqueta(etiqueta))
        if clave and clave not in campos:
            campos[clave] = valor.strip()
    return campos


def leer_pdf(ruta_pdf: str) -> list[dict]:
    """Lee un comprobante y devuelve UNA lectura por página con texto aprovechable.

    Cada lectura trae las claves del extractor (`documento_lectura`,
    `cuenta_origen`, `cuenta_destino`, `importe`) más los campos del comprobante.
    `pagina` es 1-based. Lanza `ErrorLectura` si el PDF no se puede abrir."""
    nombre = os.path.basename(ruta_pdf)
    try:
        doc = pymupdf.open(ruta_pdf)
    except Exception as exc:  # noqa: BLE001 — se traduce a un error propio
        raise ErrorLectura(f"«{nombre}»: no se pudo abrir el PDF ({exc}).") from exc
    lecturas: list[dict] = []
    try:
        if doc.needs_pass:
            raise ErrorLectura(f"«{nombre}»: el PDF está protegido con contraseña.")
        for i, pagina in enumerate(doc, start=1):
            campos = _campos_de_texto(pagina.get_text("text") or "")
            # Sin importe ni cuentas no hay nada que casar: no es un comprobante.
            if not campos.get("importe") and not campos.get("cuenta_destino"):
                continue
            lecturas.append({
                "documento_lectura": nombre,
                "pagina": i,
                "cuenta_origen": campos.get("cuenta_origen", ""),
                "cuenta_destino": campos.get("cuenta_destino", ""),
                "importe": _a_importe(campos.get("importe", "")),
                "fecha_aplicacion": campos.get("fecha_aplicacion", ""),
                "fecha_creacion": campos.get("fecha_creacion", ""),
                "referencia": campos.get("referencia", ""),
                "clave_rastreo": campos.get("clave_rastreo", ""),
                "folio_interbancario": campos.get("folio_interbancario", ""),
                "folio_firma": campos.get("folio_firma", ""),
                "folio_unico": campos.get("folio_unico", ""),
                "banco_beneficiario": campos.get("banco_beneficiario", ""),
                "concepto": campos.get("concepto_pago") or campos.get("motivo_pago", ""),
                "tipo_operacion": campos.get("tipo_operacion", ""),
                "estado": campos.get("estado", ""),
            })
    finally:
        doc.close()
    return lecturas


def leer_varios(rutas_pdf: list[str]) -> tuple[list[dict], list[tuple[str, str]]]:
    """Lee varios comprobantes. Devuelve `(lecturas, errores)`, donde `errores` es
    una lista de `(ruta, motivo)`: un PDF ilegible no debe abortar el lote."""
    lecturas: list[dict] = []
    errores: list[tuple[str, str]] = []
    for ruta in rutas_pdf:
        try:
            leidas = leer_pdf(ruta)
        except ErrorLectura as exc:
            errores.append((ruta, str(exc)))
            continue
        if not leidas:
            errores.append((ruta, "no se encontraron datos de pago en el PDF."))
            continue
        lecturas.extend(leidas)
    return lecturas, errores


def esta_aplicado(lectura: dict) -> bool:
    """True si el comprobante reporta la operación como aplicada ('Operado')."""
    return _clave_etiqueta(lectura.get("estado", "")) == ESTADO_APLICADO


def referencia_aaaammdd(lectura: dict) -> str:
    """Referencia que el RPA escribe en el SIPP: la FECHA DE APLICACIÓN del
    comprobante como AAAAMMDD ('04/08/2026' -> '20260804').

    Se usa la fecha de aplicación —y no la de la dispersión ni la del día en que
    corre el RPA— porque es por comprobante: sigue siendo correcta aunque los
    comprobantes se suban días después o el lote abarque varios días. Si no viene,
    se cae a la fecha de creación; '' si no hay ninguna (el RPA debe avisar en vez
    de escribir una referencia inventada)."""
    for clave in ("fecha_aplicacion", "fecha_creacion"):
        m = re.match(r"\s*(\d{2})/(\d{2})/(\d{4})", lectura.get(clave, "") or "")
        if m:
            return f"{m.group(3)}{m.group(2)}{m.group(1)}"
    return ""


def fecha_aplicacion_ddmmaaaa(lectura: dict) -> str:
    """Fecha de aplicación del comprobante como 'DD/MM/AAAA', que es el formato en
    que el SIPP captura la Fecha de Devolución.

    Es la fecha en que el banco aplicó el pago, no la de hoy: el portal prellena
    ese campo con el día en que se captura, así que sin fijarlo una devolución
    subida días después quedaría registrada con la fecha equivocada.

    Devuelve '' si el comprobante no la trae (quien llame debe dejar entonces lo
    que el portal haya prellenado, en vez de inventar una fecha).
    """
    for clave in ("fecha_aplicacion", "fecha_creacion"):
        m = re.match(r"\s*(\d{2})/(\d{2})/(\d{4})", lectura.get(clave, "") or "")
        if m:
            return f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
    return ""
