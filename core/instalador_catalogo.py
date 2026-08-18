"""Instalación TRANSACCIONAL de un Excel de catálogo (cuentas bancarias, cuentas de
dispersión…).

Los catálogos se actualizan copiando el Excel que elige el usuario sobre el que vive
junto al .exe. La copia tiene que ser todo-o-nada: si el archivo nuevo resulta
ilegible, hay que quedarse con el anterior. Aquí vive esa mecánica, compartida, para
no tener que acertarla dos veces.

Tres casos que la versión anterior no cubría y que sí ocurren:

  * **Elegir el archivo desde su propia ruta de instalación.** Es lo natural cuando
    alguien acaba de editarlo ahí: no es un error, es pedir que se relea. Copiar un
    archivo sobre sí mismo lanza `shutil.SameFileError`, y el rollback posterior
    fallaba tapando la causa real.

  * **El destino abierto en Excel.** Excel bloquea el archivo contra escritura, así
    que la copia moría con un `[Errno 13] Permission denied` que no le dice nada al
    usuario. Ahora se detecta ANTES de tocar nada y se explica qué hacer.

  * **Un rollback que falla.** Si la restauración no se podía escribir (justo lo que
    pasa con el archivo abierto en Excel), el `finally` borraba el respaldo de todos
    modos: el destino quedaba a medias y sin vuelta atrás. Ahora el respaldo SOLO se
    borra cuando ya no hace falta.
"""

from __future__ import annotations

import os
import shutil


class CatalogoBloqueado(PermissionError):
    """El archivo de destino no se puede escribir (típicamente, abierto en Excel)."""


class RespaldoConservado(RuntimeError):
    """La instalación falló y el rollback tampoco pudo escribir. Se conserva el
    respaldo en `ruta_respaldo` para restaurarlo a mano."""

    def __init__(self, mensaje: str, ruta_respaldo: str):
        super().__init__(mensaje)
        self.ruta_respaldo = ruta_respaldo


def _mismo_archivo(a: str, b: str) -> bool:
    """True si las dos rutas apuntan al MISMO archivo (aunque se escriban distinto:
    mayúsculas, rutas relativas, 8.3 de Windows…). False si alguna no existe."""
    try:
        return os.path.samefile(a, b)
    except OSError:
        return False


def _escribible(ruta: str) -> bool:
    """True si se puede abrir para escritura. Un archivo abierto en Excel existe y
    se puede LEER, pero no escribir: por eso no basta con os.access."""
    if not os.path.exists(ruta):
        return True  # todavía no está: lo que manda es el permiso de la carpeta
    try:
        with open(ruta, "r+b"):
            return True
    except OSError:
        return False


def instalar(ruta_origen: str, ruta_destino: str, leer, validar,
             ruta_cache: str | None = None):
    """Instala `ruta_origen` como `ruta_destino` y devuelve lo que `leer` obtuvo.

    `leer(ruta)` devuelve el catálogo ya parseado; `validar(catalogo)` debe LANZAR la
    excepción propia del módulo si no se reconoce. Si `ruta_cache` se provee, se borra
    al terminar bien (para que la próxima consulta relea el archivo nuevo).

    Lanza `CatalogoBloqueado` si el destino está abierto, `RespaldoConservado` si
    además falló el rollback, o lo que lance `validar`.
    """
    os.makedirs(os.path.dirname(ruta_destino), exist_ok=True)

    # Caso 1: ya ES el archivo instalado -> no hay nada que copiar, solo releer.
    # Se atiende primero porque funciona incluso con el archivo abierto en Excel
    # (leerlo sí se puede; escribirlo no).
    if _mismo_archivo(ruta_origen, ruta_destino):
        catalogo = leer(ruta_destino)
        validar(catalogo)
        _borrar_cache(ruta_cache)
        return catalogo

    # Caso 2: destino bloqueado -> se corta ANTES de respaldar ni copiar, para no
    # dejar estado a medias.
    if not _escribible(ruta_destino):
        raise CatalogoBloqueado(
            f"«{os.path.basename(ruta_destino)}» está abierto en Excel (o en otro "
            "programa). Ciérralo e intenta de nuevo.")

    respaldo = None
    conservar_respaldo = False
    if os.path.exists(ruta_destino):
        respaldo = ruta_destino + ".bak"
        shutil.copyfile(ruta_destino, respaldo)
    try:
        shutil.copyfile(ruta_origen, ruta_destino)
        catalogo = leer(ruta_destino)
        validar(catalogo)
    except Exception as exc:
        try:
            if respaldo is not None:
                shutil.copyfile(respaldo, ruta_destino)
            elif os.path.exists(ruta_destino):
                os.remove(ruta_destino)
        except OSError as fallo:
            # El respaldo NO se borra: es lo único que queda del archivo bueno. Hace
            # falta la bandera porque el `finally` se ejecuta también cuando este
            # `except` relanza, y sin ella lo borraría justo en el peor momento.
            conservar_respaldo = True
            raise RespaldoConservado(
                f"No se pudo instalar el catálogo ({exc}) y tampoco restaurar el "
                f"anterior ({fallo}). Se dejó una copia del archivo bueno en "
                f"«{os.path.basename(respaldo or '')}»: renómbrala para recuperarlo.",
                respaldo or "") from exc
        raise
    finally:
        # Se borra solo si el destino quedó consistente (instalado o restaurado).
        if (respaldo is not None and not conservar_respaldo
                and os.path.exists(respaldo)):
            try:
                os.remove(respaldo)
            except OSError:
                pass
    _borrar_cache(ruta_cache)
    return catalogo


def _borrar_cache(ruta_cache: str | None) -> None:
    """Invalida el caché en disco. Si no se puede borrar no pasa nada: la próxima
    lectura lo sobrescribe."""
    if not ruta_cache:
        return
    try:
        if os.path.exists(ruta_cache):
            os.remove(ruta_cache)
    except OSError:
        pass
