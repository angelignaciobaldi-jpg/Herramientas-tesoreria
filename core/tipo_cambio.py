"""Tipo de cambio USD (FIX) publicado por el DOF, con caché en memoria.

El Diario Oficial de la Federación publica el tipo de cambio del dólar (indicador
158). Se consulta su página de detalle (server-side, sin JS) con un rango de fechas
y se toma el valor de la fila ``<tr class="Celda 1">`` (2ª columna = valor).

Por defecto se toma el valor del DÍA HÁBIL ANTERIOR: el día previo de martes a
viernes, y el VIERNES pasado cuando hoy es lunes (ver `fecha_referencia`).

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

# Días hacia atrás que se recorren cuando la fecha pedida no tiene publicación
# (día inhábil). No se lleva un calendario de días festivos a propósito: el DOF
# simplemente no publica esos días, así que basta con tomar la última fila con
# valor. Un calendario propio habría que actualizarlo cada año y se desfasaría de
# la realidad. 15 días cubren con holgura el hueco más largo (Semana Santa, que
# deja sin publicación de miércoles a lunes) y cualquier puente o suspensión
# extraordinaria.
_DIAS_RESPALDO = 15

# Caché en memoria: {fecha_dd/mm/aaaa consultada -> (valor, fecha_dof)}. Persiste
# hasta cerrar la app. Es de módulo a propósito (una sola consulta por sesión).
# `fecha_dof` es la fecha de la fila del DOF de la que salió el valor (puede
# diferir de la consultada si esa cayó en fin de semana/festivo).
_CACHE: dict[str, tuple[float, str]] = {}


class TipoCambioNoDisponible(Exception):
    """No se pudo obtener el tipo de cambio del DOF (sin conexión o sin dato)."""


def fecha_referencia(hoy: datetime.date | None = None) -> datetime.date:
    """Fecha del DOF con la que se toma el tipo de cambio: el DÍA HÁBIL ANTERIOR.

    De martes a viernes es simplemente el día anterior. Los LUNES se retrocede
    hasta el VIERNES previo (p. ej. el lunes 10/08/2026 usa el del viernes
    07/08/2026): el DOF no publica en fin de semana, así que el valor vigente
    sigue siendo el del cierre de la semana pasada.
    """
    hoy = hoy or datetime.date.today()
    # weekday(): lunes=0 ... domingo=6. Solo el lunes salta el fin de semana.
    return hoy - datetime.timedelta(days=3 if hoy.weekday() == 0 else 1)


def _url(dfecha: str, hfecha: str) -> str:
    # Las diagonales de la fecha van LITERALES (safe="/"), no como %2F: el DOF
    # dejó de aceptar la forma codificada y responde con una redirección a
    # Error.php, que no trae ninguna fila de datos.
    d = urllib.parse.quote(dfecha, safe="/")
    h = urllib.parse.quote(hfecha, safe="/")
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


def _fecha_dof_texto(texto: str) -> str:
    """Normaliza la fecha de la 1ª columna del DOF a 'DD/MM/AAAA'. El DOF la
    publica con guiones ('07-08-2026'), pero el resto de la app trabaja con
    diagonales. Si no se reconoce, devuelve el texto tal cual."""
    crudo = (texto or "").strip()
    candidato = crudo.replace("-", "/").replace(".", "/")
    try:
        datetime.datetime.strptime(candidato, "%d/%m/%Y")
    except ValueError:
        return crudo
    return candidato


def _a_float(texto: str) -> float | None:
    limpio = re.sub(r"[^\d.]", "", (texto or "").replace(",", ""))
    try:
        return float(limpio) if limpio else None
    except ValueError:
        return None


def _consultar(
    dfecha: datetime.date, hfecha: datetime.date,
) -> tuple[float, str] | None:
    """Consulta el DOF en el rango [dfecha, hfecha] y devuelve `(valor, fecha_dof)`
    de la ÚLTIMA fila (la fecha más reciente con dato) o None si no hay. `fecha_dof`
    es la fecha de esa fila (1ª columna), normalizada a 'DD/MM/AAAA'."""
    html = _descargar(_url(dfecha.strftime("%d/%m/%Y"), hfecha.strftime("%d/%m/%Y")))
    filas = _filas_celda1(html)
    for celdas in reversed(filas):  # la más reciente primero
        if len(celdas) > 1:
            valor = _a_float(celdas[1])
            if valor:
                return valor, _fecha_dof_texto(celdas[0])
    return None


def tipo_cambio_usd_detalle(
    fecha: datetime.date | None = None, refrescar: bool = False,
) -> tuple[float, str]:
    """Tipo de cambio USD del DOF para `fecha` (por defecto, el DÍA HÁBIL ANTERIOR;
    ver `fecha_referencia`), junto con la fecha de publicación del DOF de la que
    salió: `(valor, fecha_dof)`.

    Usa la caché en memoria salvo que `refrescar` sea True. Si el día pedido cayó
    en DÍA INHÁBIL (festivo, puente o suspensión) y no tiene dato publicado, se
    recorre hacia atrás y se toma el último valor publicado antes de esa fecha;
    por eso `fecha_dof` puede diferir de `fecha`, y es la que hay que mostrarle al
    usuario. Lanza TipoCambioNoDisponible si no se obtiene.
    """
    if fecha is None:
        fecha = fecha_referencia()
    clave = fecha.strftime("%d/%m/%Y")
    if not refrescar and clave in _CACHE:
        return _CACHE[clave]

    # 1) El día pedido (dfecha = hfecha), como en el ejemplo del DOF.
    dato = _consultar(fecha, fecha)
    # 2) Respaldo para días inhábiles: rango de los días PREVIOS a la fecha pedida,
    #    tomando el valor publicado más reciente. El rango cierra en `fecha`, nunca
    #    en hoy: si mirara hacia adelante, un lunes cuyo viernes fue festivo
    #    terminaría tomando el valor publicado HOY.
    if dato is None:
        dato = _consultar(fecha - datetime.timedelta(days=_DIAS_RESPALDO), fecha)
    if dato is None:
        raise TipoCambioNoDisponible(
            f"El DOF no publicó tipo de cambio entre el "
            f"{(fecha - datetime.timedelta(days=_DIAS_RESPALDO)):%d/%m/%Y} y el "
            f"{fecha:%d/%m/%Y}.")
    _CACHE[clave] = dato
    return dato


def tipo_cambio_usd(
    fecha: datetime.date | None = None, refrescar: bool = False,
) -> float:
    """Tipo de cambio USD del DOF (solo el valor). Ver `tipo_cambio_usd_detalle`
    si además se necesita la fecha de publicación del DOF."""
    return tipo_cambio_usd_detalle(fecha, refrescar)[0]


def limpiar_cache() -> None:
    """Vacía la caché en memoria (p. ej. para pruebas)."""
    _CACHE.clear()
