"""Cliente HTTP central para los microservicios de la app.

Concentra en un solo lugar la mecánica de las peticiones a los endpoints
(construcción de URL, JSON, cabeceras, autenticación, TLS, timeouts y manejo de
errores), para que el resto de módulos solo llamen `api.get(...)`, `api.post(...)`,
etc. y no repitan el plumbing de `urllib`.

Sin dependencias externas: usa solo la librería estándar (`urllib`), igual que
`core/tipo_cambio.py` y `core/auto_updater.py`. La URL base y el token se resuelven
en `core/ajustes_api.py` (store local en Configuración -> variable de entorno), y
se pueden sobrescribir por llamada (por si hay varios microservicios con distinta
base). El token se guarda cifrado con DPAPI, nunca en claro ni en el .exe.

Uso típico:

    from core import api
    datos = api.get("/proveedores", params={"empresa": 8})
    creado = api.post("/dispersiones", json_body={"folio": 123})

Errores (todos derivan de `ErrorApi`):
  - `ApiSinConexion`    -> no se pudo conectar (red caída, DNS, timeout, TLS).
  - `ErrorRespuestaApi` -> el servicio respondió con un status >= 400
    (trae `.status` y `.datos` con el cuerpo del error ya parseado).

Las funciones NO capturan estos errores: el módulo que llama decide cómo
reaccionar (avisar al usuario, reintentar, caer a un valor por defecto, etc.).
"""

from __future__ import annotations

import json
import ssl
import urllib.error
import urllib.parse
import urllib.request
from typing import Any

from core import ajustes_api

# Timeout por defecto (segundos) de cada petición.
TIMEOUT = 30
# Identifica a la app en el header User-Agent.
_UA = "HerramientasTesoreria/1.0"


class ErrorApi(Exception):
    """Error base al consultar un microservicio."""


class ApiSinConexion(ErrorApi):
    """No se pudo establecer la conexión (red caída, DNS, timeout, TLS)."""


class ErrorRespuestaApi(ErrorApi):
    """El servicio respondió con un status HTTP de error (>= 400)."""

    def __init__(self, status: int, mensaje: str, datos: Any = None):
        super().__init__(f"HTTP {status}: {mensaje}".strip())
        self.status = status
        self.datos = datos  # cuerpo del error ya parseado (dict/list/str o None)


def _parse_cuerpo(cuerpo: bytes) -> Any:
    """Decodifica el cuerpo de la respuesta: JSON si se puede, si no texto plano,
    y None si viene vacío."""
    if not cuerpo:
        return None
    texto = cuerpo.decode("utf-8", "replace")
    try:
        return json.loads(texto)
    except ValueError:
        return texto


def _construir_url(ruta: str, params: dict | None, base_url: str | None) -> str:
    """Arma la URL final. `ruta` puede ser absoluta (http...) o relativa a la base.
    `params` se añaden como query string (los valores None se omiten)."""
    if ruta.startswith(("http://", "https://")):
        url = ruta
    else:
        base = base_url if base_url is not None else ajustes_api.base_url(
            requerido=True)
        url = base.rstrip("/") + "/" + ruta.lstrip("/")
    if params:
        limpio = {k: v for k, v in params.items() if v is not None}
        if limpio:
            sep = "&" if "?" in url else "?"
            url += sep + urllib.parse.urlencode(limpio, doseq=True)
    return url


def _cabeceras(json_body: Any, headers: dict | None,
               token: str | None) -> dict[str, str]:
    """Cabeceras por defecto (Accept/User-Agent, Content-Type si hay cuerpo JSON,
    Authorization si hay token), fusionadas con las que pase quien llama."""
    cab = {"Accept": "application/json", "User-Agent": _UA}
    if json_body is not None:
        cab["Content-Type"] = "application/json"
    tok = token if token is not None else ajustes_api.token()
    if tok:
        cab["Authorization"] = f"Bearer {tok}"
    if headers:
        cab.update(headers)
    return cab


def _enviar(req: urllib.request.Request, timeout: float) -> Any:
    """Ejecuta la petición y devuelve el cuerpo parseado. Reintenta el TLS sin
    verificar si la cadena de certificados falla (algunos servidores la traen
    incompleta), igual que core/tipo_cambio.py. Traduce los fallos a ErrorApi."""
    ultimo: Exception | None = None
    for ctx in (ssl.create_default_context(), ssl._create_unverified_context()):
        try:
            with urllib.request.urlopen(req, context=ctx, timeout=timeout) as r:
                return _parse_cuerpo(r.read())
        except urllib.error.HTTPError as exc:
            # Respuesta HTTP con status de error: NO se reintenta (es válida).
            datos = _parse_cuerpo(exc.read())
            raise ErrorRespuestaApi(exc.code, exc.reason or "", datos) from exc
        except (urllib.error.URLError, TimeoutError, ssl.SSLError) as exc:
            ultimo = exc  # problema de conexión/TLS: se prueba el siguiente ctx
    raise ApiSinConexion(
        f"No se pudo conectar con el servicio: {ultimo}") from ultimo


def solicitar(
    metodo: str,
    ruta: str,
    *,
    params: dict | None = None,
    json_body: Any = None,
    headers: dict | None = None,
    base_url: str | None = None,
    token: str | None = None,
    timeout: float = TIMEOUT,
) -> Any:
    """Realiza una petición al microservicio y devuelve el cuerpo ya parseado
    (dict/list si es JSON, texto si no, o None si viene vacío).

    - `metodo`: 'GET', 'POST', 'PUT', 'PATCH', 'DELETE'.
    - `ruta`: relativa a la base configurada (p. ej. '/proveedores') o absoluta.
    - `params`: query string (los valores None se omiten).
    - `json_body`: se serializa a JSON como cuerpo de la petición.
    - `headers`: cabeceras extra (se fusionan con las por defecto).
    - `base_url`/`token`: sobrescriben la configuración por llamada.

    Lanza `ApiSinConexion` o `ErrorRespuestaApi` (ambas `ErrorApi`) ante fallos.
    """
    url = _construir_url(ruta, params, base_url)
    datos = json.dumps(json_body).encode("utf-8") if json_body is not None else None
    req = urllib.request.Request(
        url, data=datos, headers=_cabeceras(json_body, headers, token),
        method=metodo.upper())
    return _enviar(req, timeout)


# --- Atajos por verbo ----------------------------------------------------
def get(ruta: str, **kwargs) -> Any:
    """GET a `ruta`. Ver `solicitar` para los parámetros."""
    return solicitar("GET", ruta, **kwargs)


def post(ruta: str, json_body: Any = None, **kwargs) -> Any:
    """POST a `ruta` con cuerpo JSON. Ver `solicitar`."""
    return solicitar("POST", ruta, json_body=json_body, **kwargs)


def put(ruta: str, json_body: Any = None, **kwargs) -> Any:
    """PUT a `ruta` con cuerpo JSON. Ver `solicitar`."""
    return solicitar("PUT", ruta, json_body=json_body, **kwargs)


def patch(ruta: str, json_body: Any = None, **kwargs) -> Any:
    """PATCH a `ruta` con cuerpo JSON. Ver `solicitar`."""
    return solicitar("PATCH", ruta, json_body=json_body, **kwargs)


def delete(ruta: str, **kwargs) -> Any:
    """DELETE a `ruta`. Ver `solicitar`."""
    return solicitar("DELETE", ruta, **kwargs)


# --- Endpoints concretos -------------------------------------------------
# A medida que se definan los microservicios, sus funciones específicas van aquí
# (p. ej. `def proveedores(empresa): return get("/proveedores", params={...})`),
# de modo que las pantallas llamen a una función con nombre en vez de a una ruta
# suelta. Por ahora solo está lista la mecánica (get/post/...).
