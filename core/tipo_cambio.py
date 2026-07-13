"""Tipo de cambio USD (FIX) publicado por el DOF, con caché en memoria.

El Diario Oficial de la Federación publica el tipo de cambio del dólar (indicador
158). Se consulta su página de detalle (server-side, sin JS) con un rango de fechas
y se toma el valor de la fila ``<tr class="Celda 1">`` (2ª columna = valor).

El valor se guarda en una caché de módulo (`_CACHE`) que vive mientras la app esté
abierta: así no se consulta el DOF en cada operación. Para forzar una relectura,
llamar con ``refrescar=True`` o reiniciar la app.
"""

from __future__ import annotations

import datetime
import re
import ssl
import urllib.parse
import urllib.request

# Indicador del DOF para el tipo de cambio USD.
_COD_USD = 158
_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
       "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")
_TIMEOUT = 30

# Caché en memoria: {fecha_dd/mm/aaaa consultada -> valor}. Persiste hasta cerrar
# la app. Es de módulo a propósito (una sola consulta por sesión).
_CACHE: dict[str, float] = {}


class TipoCambioNoDisponible(Exception):
    """No se pudo obtener el tipo de cambio del DOF (sin conexión o sin dato)."""


def _url(dfecha: str, hfecha: str) -> str:
    d = urllib.parse.quote(dfecha, safe="")
    h = urllib.parse.quote(hfecha, safe="")
    return (f"https://www.dof.gob.mx/indicadores_detalle.php?"
            f"cod_tipo_indicador={_COD_USD}&dfecha={d}&hfecha={h}")


def _descargar(url: str) -> str:
    req = urllib.request.Request(url, headers={"User-Agent": _UA})
    # Algunos servidores gubernamentales traen la cadena de certificados
    # incompleta; se intenta con verificación y, si falla, sin ella.
    ultimo: Exception | None = None
    for ctx in (ssl.create_default_context(), ssl._create_unverified_context()):
        try:
            with urllib.request.urlopen(req, context=ctx, timeout=_TIMEOUT) as r:
                return r.read().decode("utf-8", "replace")
        except Exception as exc:  # noqa: BLE001 — se reintenta / se reporta
            ultimo = exc
    raise TipoCambioNoDisponible(
        f"No se pudo consultar el DOF: {ultimo}") from ultimo


def _filas_celda1(html: str) -> list[list[str]]:
    """Filas ``<tr class="Celda 1">`` como listas de textos de sus celdas."""
    filas = re.findall(
        r'<tr[^>]*class="[^"]*Celda\s*1[^"]*"[^>]*>(.*?)</tr>', html, re.S | re.I)
    resultado: list[list[str]] = []
    for row in filas:
        celdas = re.findall(r"<td[^>]*>(.*?)</td>", row, re.S | re.I)
        textos = [re.sub(r"<[^>]+>", "", c).replace("&nbsp;", " ").strip()
                  for c in celdas]
        resultado.append(textos)
    return resultado


def _a_float(texto: str) -> float | None:
    limpio = re.sub(r"[^\d.]", "", (texto or "").replace(",", ""))
    try:
        return float(limpio) if limpio else None
    except ValueError:
        return None


def _consultar(dfecha: datetime.date, hfecha: datetime.date) -> float | None:
    """Consulta el DOF en el rango [dfecha, hfecha] y devuelve el valor de la
    ÚLTIMA fila (la fecha más reciente con dato) o None si no hay."""
    html = _descargar(_url(dfecha.strftime("%d/%m/%Y"), hfecha.strftime("%d/%m/%Y")))
    filas = _filas_celda1(html)
    for celdas in reversed(filas):  # la más reciente primero
        if len(celdas) > 1:
            valor = _a_float(celdas[1])
            if valor:
                return valor
    return None


def tipo_cambio_usd(
    fecha: datetime.date | None = None, refrescar: bool = False,
) -> float:
    """Tipo de cambio USD del DOF para `fecha` (por defecto, AYER).

    Usa la caché en memoria salvo que `refrescar` sea True. Si el día pedido no
    tiene dato publicado (fin de semana/festivo), amplía la búsqueda hacia atrás y
    toma el más reciente disponible. Lanza TipoCambioNoDisponible si no se obtiene.
    """
    if fecha is None:
        fecha = datetime.date.today() - datetime.timedelta(days=1)
    clave = fecha.strftime("%d/%m/%Y")
    if not refrescar and clave in _CACHE:
        return _CACHE[clave]

    # 1) El día pedido (dfecha = hfecha), como en el ejemplo del DOF.
    valor = _consultar(fecha, fecha)
    # 2) Respaldo: rango de los últimos días hasta hoy (cubre fines de semana o
    #    festivos, tomando el valor publicado más reciente).
    if valor is None:
        valor = _consultar(fecha - datetime.timedelta(days=6),
                           datetime.date.today())
    if valor is None:
        raise TipoCambioNoDisponible(
            "El DOF no devolvió un tipo de cambio para la fecha solicitada.")
    _CACHE[clave] = valor
    return valor


def limpiar_cache() -> None:
    """Vacía la caché en memoria (p. ej. para pruebas)."""
    _CACHE.clear()
