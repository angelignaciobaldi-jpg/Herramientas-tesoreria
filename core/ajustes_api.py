"""Configuración de acceso a los microservicios: URL base y token de servicio.

Centraliza DÓNDE viven la URL y el token en la app final y cómo se resuelven,
con estas prioridades:

  URL base: preferencia local (Configuración)  ->  variable de entorno  ->  ''.
  Token:    DPAPI local (Configuración)         ->  variable de entorno.

- La **URL no es secreta** (se ve en el tráfico): se guarda como preferencia en
  claro (core/preferencias.py) y puede fijarse desde Configuración sin recompilar.
- El **token SÍ es sensible**: se guarda cifrado con DPAPI (core/dpapi.py), atado
  a la cuenta de Windows. NUNCA se guarda en claro, ni se sube al repo, ni se
  empaqueta con el .exe. Se captura por máquina desde Configuración.
- Las variables de entorno (`.env`) siguen sirviendo para DESARROLLO; en producción
  se usa el store local (o, mejor aún a futuro, un token por sesión vía login).

`core/api.py` consume `base_url()` y `token()` de aquí.
"""

from __future__ import annotations

import json
import os

from . import dpapi, entorno, preferencias, rutas

# Clave de la URL base en preferencias (texto claro; la URL no es secreta).
_CLAVE_URL = "api_base_url"
# Archivo con el token cifrado con DPAPI (por usuario/máquina, fuera del repo).
_RUTA_TOKEN = os.path.join(rutas.DATOS, "token_api.json")

# --- API extractor de PDF (comprobantes de pago) -------------------------
# API INDEPENDIENTE del SIPP: su propia URL (preferencia, texto claro) y su propio
# token (Bearer) cifrado con DPAPI. Sin URL por defecto: si no está configurada, se
# considera ausente (la UI lo avisa).
_CLAVE_URL_EXTRACTOR = "extractor_api_base_url"
_RUTA_TOKEN_EXTRACTOR = os.path.join(rutas.DATOS, "token_extractor_api.json")

# URL base QUEMADA por defecto. La URL no es secreta (se ve en el tráfico), así
# que se deja fija en el código para que la app funcione sin configurar nada. Es
# solo el ÚLTIMO recurso: la preferencia local (Configuración) y la variable de
# entorno la SOBRESCRIBEN, útil para apuntar a otro entorno (dev/staging/prod)
# sin recompilar. Para cambiar el default de fábrica, se edita esta constante.
# De momento (integración local) apunta al backend en localhost.
_URL_POR_DEFECTO = "http://localhost:3001"


# --- URL base ------------------------------------------------------------
def base_url(requerido: bool = False) -> str | None:
    """URL base (sin '/' final): preferencia local -> variable de entorno -> URL
    quemada por defecto. Lanza entorno.FaltaVariableEntorno solo si `requerido` y
    no hay NINGUNA (incluida la de fábrica, que normalmente siempre está)."""
    url = (preferencias.cargar_valor(_CLAVE_URL) or "").strip()
    if not url:
        url = entorno.api_base_url() or ""
    if not url:
        url = _URL_POR_DEFECTO
    url = url.rstrip("/")
    if requerido and not url:
        raise entorno.FaltaVariableEntorno(entorno.VAR_API_BASE_URL)
    return url or None


def guardar_base_url(url: str) -> None:
    """Fija la URL base (preferencia local). Cadena vacía -> queda sin fijar."""
    preferencias.guardar_valor(_CLAVE_URL, (url or "").strip().rstrip("/"))


# --- Token (cifrado con DPAPI) -------------------------------------------
def token() -> str | None:
    """Token de servicio: store DPAPI local -> variable de entorno (dev). None si
    no hay ninguno."""
    guardado = _cargar_token()
    if guardado:
        return guardado
    return entorno.api_token(requerido=False)


def hay_token_local() -> bool:
    """True si hay un token guardado localmente (cifrado)."""
    return _cargar_token() is not None


def guardar_token(valor: str) -> None:
    """Guarda el token cifrado con DPAPI. Vacío -> borra el token guardado."""
    valor = (valor or "").strip()
    if not valor:
        borrar_token()
        return
    with open(_RUTA_TOKEN, "w", encoding="utf-8") as fh:
        json.dump({"token": dpapi.cifrar(valor)}, fh)


def borrar_token() -> None:
    """Elimina el token guardado localmente."""
    try:
        os.remove(_RUTA_TOKEN)
    except FileNotFoundError:
        pass


def _cargar_token() -> str | None:
    """Lee y descifra el token local; None si no hay o no se puede descifrar."""
    if not os.path.exists(_RUTA_TOKEN):
        return None
    try:
        with open(_RUTA_TOKEN, encoding="utf-8") as fh:
            datos = json.load(fh)
        return dpapi.descifrar(datos["token"]) or None
    except (OSError, ValueError, KeyError):
        return None


# --- API extractor: URL base ---------------------------------------------
def base_url_extractor(requerido: bool = False) -> str | None:
    """URL base de la API extractor (sin '/' final): preferencia local -> variable de
    entorno. SIN default de fábrica: None si no está configurada. Lanza
    entorno.FaltaVariableEntorno solo si `requerido` y no hay ninguna."""
    url = (preferencias.cargar_valor(_CLAVE_URL_EXTRACTOR) or "").strip()
    if not url:
        url = entorno.extractor_api_base_url() or ""
    url = url.rstrip("/")
    if requerido and not url:
        raise entorno.FaltaVariableEntorno(entorno.VAR_EXTRACTOR_API_BASE_URL)
    return url or None


def guardar_base_url_extractor(url: str) -> None:
    """Fija la URL base de la API extractor (preferencia local). Vacía -> sin fijar."""
    preferencias.guardar_valor(
        _CLAVE_URL_EXTRACTOR, (url or "").strip().rstrip("/"))


# --- API extractor: token (Bearer, cifrado con DPAPI) --------------------
def token_extractor() -> str | None:
    """Bearer token de la API extractor: store DPAPI local -> variable de entorno
    (dev). None si no hay ninguno."""
    guardado = _cargar_token_extractor()
    if guardado:
        return guardado
    return entorno.extractor_api_token(requerido=False)


def hay_token_extractor_local() -> bool:
    """True si hay un token del extractor guardado localmente (cifrado)."""
    return _cargar_token_extractor() is not None


def guardar_token_extractor(valor: str) -> None:
    """Guarda el token del extractor cifrado con DPAPI. Vacío -> borra el guardado."""
    valor = (valor or "").strip()
    if not valor:
        borrar_token_extractor()
        return
    with open(_RUTA_TOKEN_EXTRACTOR, "w", encoding="utf-8") as fh:
        json.dump({"token": dpapi.cifrar(valor)}, fh)


def borrar_token_extractor() -> None:
    """Elimina el token del extractor guardado localmente."""
    try:
        os.remove(_RUTA_TOKEN_EXTRACTOR)
    except FileNotFoundError:
        pass


def _cargar_token_extractor() -> str | None:
    """Lee y descifra el token local del extractor; None si no hay o no se descifra."""
    if not os.path.exists(_RUTA_TOKEN_EXTRACTOR):
        return None
    try:
        with open(_RUTA_TOKEN_EXTRACTOR, encoding="utf-8") as fh:
            datos = json.load(fh)
        return dpapi.descifrar(datos["token"]) or None
    except (OSError, ValueError, KeyError):
        return None
