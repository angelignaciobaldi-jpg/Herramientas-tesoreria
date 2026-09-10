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




# --- Formato Banregio ----------------------------------------------------
# El comprobante de Banregio no usa "Etiqueta: valor": aplana un layout de DOS
# COLUMNAS, así que cada etiqueta queda pegada a su valor pero unas veces ANTES
# y otras DESPUÉS. Por eso no se asume la dirección: se miran los dos vecinos y
# se toma el que tenga la forma esperada (un importe donde va un importe, algo
# con dígitos donde va una cuenta).
_ETIQUETAS_BANREGIO = ("Cantidad a Transferir", "Número de referencia",
                       "Cuenta Origen", "Cuenta Destino")

# TODAS las etiquetas del formato. Se usan para descartar candidatos: al mirar
# los vecinos de una etiqueta, el de un lado suele ser OTRA etiqueta, y sin esta
# lista se colaba como valor (el concepto salía "Cantidad a Transferir").
_ROTULOS_BANREGIO = frozenset(_clave_etiqueta(e) for e in (
    "Tipo de Transferencia", "Cuenta Origen", "Cuenta Destino",
    "Cantidad a Transferir", "Concepto de pago", "Número de referencia",
    "Quien autoriza", "Quien solicita", "Recibo de la solicitud",
    "Fecha solicita", "Banco", "Verificador", "Datos de tu operación",
))

_MESES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}


def es_banregio(texto: str) -> bool:
    """True si la página tiene la pinta del comprobante de Banregio."""
    plano = _clave_etiqueta(texto)
    return any(_clave_etiqueta(e) in plano for e in _ETIQUETAS_BANREGIO)


def _fecha_larga(texto: str) -> str:
    """'7 mayo 2026 - 11:45 a. m.' -> '07/05/2026'. '' si no parsea.

    Banregio escribe la fecha con el mes en palabras, no en dígitos.
    """
    m = re.search(r"(\d{1,2})\s+([A-Za-zÁÉÍÓÚáéíóú]+)\s+(\d{4})", texto or "")
    if not m:
        return ""
    mes = _MESES.get(_clave_etiqueta(m.group(2)))
    if not mes:
        return ""
    return f"{int(m.group(1)):02d}/{mes:02d}/{m.group(3)}"


def _vecino(lineas: list, etiqueta: str, valido) -> str:
    """Valor pegado a `etiqueta`: se prueba la línea de ARRIBA y la de ABAJO y
    se devuelve la primera que pase `valido`. Así da igual de qué lado quedó al
    aplanarse el PDF."""
    objetivo = _clave_etiqueta(etiqueta)
    for i, linea in enumerate(lineas):
        if _clave_etiqueta(linea) != objetivo:
            continue
        for j in (i - 1, i + 1):
            if not 0 <= j < len(lineas):
                continue
            candidato = lineas[j]
            if _clave_etiqueta(candidato) in _ROTULOS_BANREGIO:
                continue   # el vecino es otra etiqueta, no un valor
            if valido(candidato):
                return candidato.strip()
    return ""


def _cuenta_de(texto: str) -> str:
    """Cuenta dentro de 'NOMBRE - 137456101298786280' o 'NOMBRE - *0011'.

    Se queda con el ÚLTIMO tramo tras el guion: los nombres de empresa traen
    guiones propios ("S.A. DE C.V. - *0011").
    """
    cola = str(texto or "").rsplit("-", 1)[-1].strip()
    return cola if re.search(r"\d", cola) else ""


def _campos_banregio(texto: str) -> dict:
    """Extrae los datos de pago de un comprobante Banregio.

    Devuelve las mismas claves que el formato BBVA para que el casado y la
    referencia funcionen igual, vengan del banco que vengan.
    """
    lineas = [l.strip() for l in (texto or "").splitlines() if l.strip()]
    tiene_digitos = lambda s: bool(re.search(r"\d", s))  # noqa: E731
    es_importe = lambda s: bool(re.match(r"^\$?\s*[\d,]+\.\d{2}$", s.strip()))  # noqa: E731

    origen = _cuenta_de(_vecino(lineas, "Cuenta Origen", tiene_digitos))
    destino = _cuenta_de(_vecino(lineas, "Cuenta Destino", tiene_digitos))
    importe = _vecino(lineas, "Cantidad a Transferir", es_importe)
    referencia = _vecino(lineas, "Número de referencia",
                         lambda s: s.strip().isdigit())
    # Banregio no rotula "fecha de aplicación": la operación es del mismo día
    # hábil (SPEI), así que la fecha de la solicitud ES la de aplicación.
    fecha = _fecha_larga(_vecino(lineas, "Fecha solicita",
                                 lambda s: bool(_fecha_larga(s))))
    concepto = _vecino(lineas, "Concepto de pago",
                       lambda s: bool(s.strip()) and not tiene_digitos(s))
    banco = _vecino(lineas, "Banco", lambda s: s.strip().isalpha())
    return {
        "cuenta_origen": origen,
        "cuenta_destino": destino,
        "importe": importe,
        "fecha_aplicacion": fecha,
        "referencia": referencia,
        "concepto_pago": concepto,
        "banco_beneficiario": banco,
        "tipo_operacion": _vecino(lineas, "Tipo de Transferencia",
                                  lambda s: bool(s.strip())),
    }

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
            texto = pagina.get_text("text") or ""
            # Cada banco arma el comprobante a su manera; se detecta el formato y
            # se normaliza a las MISMAS claves, para que el casado y la referencia
            # no tengan que saber de qué banco vino.
            if es_banregio(texto):
                campos, emisor = _campos_banregio(texto), "Banregio"
            else:
                campos, emisor = _campos_de_texto(texto), "BBVA"
            # Sin importe ni cuentas no hay nada que casar: no es un comprobante.
            if not campos.get("importe") and not campos.get("cuenta_destino"):
                continue
            lecturas.append({
                "documento_lectura": nombre,
                "emisor": emisor,
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
