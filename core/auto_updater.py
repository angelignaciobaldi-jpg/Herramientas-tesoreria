"""AutoUpdater: actualización automática desde releases PRIVADAS de GitHub.

Pensado para una app empaquetada con PyInstaller (modo `sys.frozen`) e instalada
con Inno Setup. Flujo:

  1. Consulta la última release: GET /repos/{owner}/{repo}/releases/latest, con el
     Token PAT en el header `Authorization: Bearer <TOKEN>`.
  2. Compara la versión local (core.version.__version__) contra el `tag_name`.
  3. Si hay una más nueva, ubica el asset `Instalador_Quetzaltic.exe`, toma su id
     y lo descarga por la API de assets privados
     (GET /repos/{owner}/{repo}/releases/assets/{asset_id}) con
     `Accept: application/octet-stream`. Lo guarda como `nuevo_instalador.exe`.
  4. Escribe un .bat temporal que espera ~3 s (a que la app cierre con
     sys.exit()), corre el instalador en modo silencioso de Inno Setup
     (`/VERYSILENT /SUPPRESSMSGBOXES /NORESTART`) y luego se borra a sí mismo y
     al instalador. El .bat se lanza desacoplado para sobrevivir al cierre.

Sin dependencias externas: usa solo la librería estándar (urllib), apto para
empaquetar con PyInstaller.

Nota de seguridad: en una app distribuida, el PAT queda accesible (binario o
config). Usa un token de mínimo alcance (solo lectura de este repo). El token se
toma de la variable de entorno QUETZALTIC_GITHUB_PAT (ver core/entorno.py), que
puede definirse en el sistema o en un archivo .env junto a la app; evita
embeberlo en el código.
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import tempfile
import urllib.error
import urllib.parse
import urllib.request

from core import entorno, rutas
from core.version import __version__ as VERSION_ACTUAL

# --- Configuración del repositorio (privado) ---
OWNER = "angelignaciobaldi-jpg"
REPO = "Herramientas-tesoreria"
NOMBRE_ASSET = "Instalador_Quetzaltic.exe"
NOMBRE_DESCARGA = "nuevo_instalador.exe"
NOMBRE_BAT = "actualizar_quetzaltic.bat"

API = "https://api.github.com"
API_VERSION = "2022-11-28"
USER_AGENT = "Quetzaltic-AutoUpdater"


class ErrorActualizacion(Exception):
    """Falla durante la búsqueda, descarga o aplicación de la actualización."""


class _RedireccionSinAuth(urllib.request.HTTPRedirectHandler):
    """Quita el header Authorization cuando GitHub redirige el asset a otro host.

    La API de assets responde con un 302 hacia una URL firmada (S3) que rechaza
    el header `Authorization`. Hay que dejar de mandarlo al cambiar de dominio."""

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        nueva = super().redirect_request(req, fp, code, msg, headers, newurl)
        if nueva is not None:
            host_orig = urllib.parse.urlsplit(req.full_url).hostname
            host_nuevo = urllib.parse.urlsplit(newurl).hostname
            if host_orig != host_nuevo:
                for cabecera in list(nueva.headers):
                    if cabecera.lower() == "authorization":
                        del nueva.headers[cabecera]
        return nueva


class AutoUpdater:
    """Comprueba, descarga y aplica actualizaciones desde releases de GitHub."""

    def __init__(
        self,
        token: str | None = None,
        owner: str = OWNER,
        repo: str = REPO,
        version_actual: str = VERSION_ACTUAL,
        nombre_asset: str = NOMBRE_ASSET,
        timeout: int = 30,
    ):
        # Por defecto, el PAT se toma del entorno (QUETZALTIC_GITHUB_PAT).
        token = token or entorno.github_pat(requerido=False)
        if not token:
            raise ErrorActualizacion(
                "Falta el Token PAT para acceder al repo privado. Define la "
                "variable de entorno QUETZALTIC_GITHUB_PAT (o un archivo .env)."
            )
        self.token = token
        self.owner = owner
        self.repo = repo
        self.version_actual = version_actual
        self.nombre_asset = nombre_asset
        self.timeout = timeout
        self._opener = urllib.request.build_opener(_RedireccionSinAuth)

    # ------------------------------------------------------- peticiones
    def _headers(self, accept: str) -> dict[str, str]:
        return {
            "Authorization": f"Bearer {self.token}",
            "Accept": accept,
            "X-GitHub-Api-Version": API_VERSION,
            "User-Agent": USER_AGENT,
        }

    def _pedir(self, url: str, accept: str) -> bytes:
        req = urllib.request.Request(url, headers=self._headers(accept))
        try:
            with self._opener.open(req, timeout=self.timeout) as resp:
                return resp.read()
        except urllib.error.HTTPError as exc:
            raise ErrorActualizacion(
                f"GitHub respondió {exc.code} al pedir {url}: {exc.reason}"
            ) from exc
        except urllib.error.URLError as exc:
            raise ErrorActualizacion(f"No se pudo conectar a GitHub: {exc.reason}") from exc

    # --------------------------------------------------- chequeo de versión
    def obtener_release_latest(self) -> dict:
        """Devuelve el JSON de la última release del repo."""
        url = f"{API}/repos/{self.owner}/{self.repo}/releases/latest"
        datos = self._pedir(url, "application/vnd.github+json")
        return json.loads(datos.decode("utf-8"))

    def _id_asset(self, release: dict) -> int | None:
        """Busca el id del asset por nombre dentro de la release."""
        for asset in release.get("assets", []):
            if asset.get("name") == self.nombre_asset:
                return asset.get("id")
        return None

    @staticmethod
    def _normalizar(version: str) -> tuple[int, ...]:
        """Convierte 'v1.2.3' / '1.2.3' en (1, 2, 3) para comparar numéricamente."""
        partes = re.findall(r"\d+", version or "")
        return tuple(int(p) for p in partes) or (0,)

    def hay_version_mas_nueva(self, tag_name: str) -> bool:
        """True si `tag_name` (remoto) es mayor que la versión local."""
        remota = self._normalizar(tag_name)
        local = self._normalizar(self.version_actual)
        largo = max(len(remota), len(local))
        remota += (0,) * (largo - len(remota))
        local += (0,) * (largo - len(local))
        return remota > local

    # ------------------------------------------------------ descarga
    def descargar_asset(self, asset_id: int, destino: str | None = None) -> str:
        """Descarga el asset privado como `nuevo_instalador.exe`. Devuelve la ruta."""
        if destino is None:
            destino = os.path.join(self._dir_temporal(), NOMBRE_DESCARGA)
        destino = os.path.abspath(destino)
        url = f"{API}/repos/{self.owner}/{self.repo}/releases/assets/{asset_id}"
        contenido = self._pedir(url, "application/octet-stream")
        try:
            with open(destino, "wb") as fh:
                fh.write(contenido)
        except OSError as exc:
            raise ErrorActualizacion(f"No se pudo guardar el instalador: {exc}") from exc
        return destino

    # ------------------------------------------- aplicar y reiniciar
    def aplicar_y_salir(self, ruta_instalador: str) -> None:
        """Escribe el .bat, lo lanza desacoplado y cierra la app (sys.exit()).

        El .bat espera ~3 s, corre el instalador mostrando su barra de progreso
        (/SILENT), REINICIA la app ya actualizada y se autolimpia."""
        ruta_instalador = os.path.abspath(ruta_instalador)
        ruta_bat = os.path.join(self._dir_temporal(), NOMBRE_BAT)
        # Ruta de la app a reiniciar tras instalar (el mismo .exe en ejecución;
        # la actualización lo sobrescribe en el sitio).
        exe = os.path.abspath(sys.executable)
        dir_exe = os.path.dirname(exe)
        # /SILENT (no /VERYSILENT) para que Inno muestre su ventana de progreso y
        # el usuario vea que se está instalando. 'ping -n 4' da ~3 s de espera de
        # forma fiable en un proceso sin consola (a diferencia de 'timeout').
        contenido = (
            "@echo off\r\n"
            "rem Espera ~3 s a que la aplicacion termine de cerrarse.\r\n"
            "ping 127.0.0.1 -n 4 >nul\r\n"
            f'"{ruta_instalador}" /SILENT /SUPPRESSMSGBOXES /NORESTART\r\n'
            "rem Reinicia la app ya actualizada.\r\n"
            f'start "" /D "{dir_exe}" "{exe}"\r\n'
            f'del "{ruta_instalador}"\r\n'
            'del "%~f0"\r\n'
        )
        try:
            with open(ruta_bat, "w", encoding="ascii") as fh:
                fh.write(contenido)
        except OSError as exc:
            raise ErrorActualizacion(f"No se pudo crear el script de actualización: {exc}") from exc

        flags = 0
        if os.name == "nt":
            flags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP
        subprocess.Popen(
            ["cmd", "/c", ruta_bat],
            creationflags=flags,
            close_fds=True,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        sys.exit(0)

    # ---------------------------------------------------- orquestación
    def buscar_y_descargar(self, al_iniciar_descarga=None) -> str | None:
        """Comprueba y, si hay versión nueva, descarga el instalador y devuelve su
        ruta — SIN aplicarlo ni cerrar la app. Devuelve None si ya está al día.

        Al no llamar a sys.exit, es seguro ejecutarlo en un hilo (p. ej. con
        asyncio.to_thread) para no congelar la interfaz durante la descarga; luego
        el llamador aplica con aplicar_y_salir() en el hilo principal.

        `al_iniciar_descarga(tag)`: callback opcional justo antes de descargar.
        Lanza ErrorActualizacion ante cualquier problema (red, asset ausente…)."""
        release = self.obtener_release_latest()
        tag = release.get("tag_name", "")
        if not self.hay_version_mas_nueva(tag):
            return None
        # Anti-bucle: si esta MISMA release ya se intentó aplicar y aun así la
        # versión local no avanzó (p. ej. el instalador de la release trae una
        # versión más vieja que su propio tag), no reintentar: se caería en un
        # bucle infinito de "actualizar" en cada arranque.
        if self._tag_ya_aplicado(tag):
            return None
        asset_id = self._id_asset(release)
        if asset_id is None:
            raise ErrorActualizacion(
                f"La release {tag} no incluye el asset '{self.nombre_asset}'."
            )
        if al_iniciar_descarga is not None:
            al_iniciar_descarga(tag)
        ruta = self.descargar_asset(asset_id)
        # Marca la release ANTES de aplicar (aplicar_y_salir no retorna). Si el
        # instalador de esa release trae una versión más vieja que su tag, en el
        # siguiente arranque la versión seguirá sin avanzar; el guard de arriba
        # detectará que este tag ya se intentó y NO volverá a aplicarlo (evita el
        # bucle infinito de actualización).
        self._marcar_tag_aplicado(tag)
        return ruta

    def buscar_y_actualizar(self, al_iniciar_descarga=None) -> bool:
        """Comprueba, y si hay versión nueva descarga y aplica (no retorna: la
        app se cierra y se reinicia sola). Devuelve False si ya está actualizada."""
        ruta = self.buscar_y_descargar(al_iniciar_descarga)
        if ruta is None:
            return False
        self.aplicar_y_salir(ruta)  # no retorna: cierra la app y la reinicia
        return True

    def hay_actualizacion(self) -> str | None:
        """Devuelve el tag de la última release si es MÁS NUEVA que la versión
        instalada (y no se marcó ya como aplicada, para no ofrecer una que
        entraría en bucle); si no, None. Solo CONSULTA, no descarga nada. Útil
        para avisar al usuario y que él decida cuándo aplicarla."""
        release = self.obtener_release_latest()
        tag = release.get("tag_name", "")
        if not self.hay_version_mas_nueva(tag) or self._tag_ya_aplicado(tag):
            return None
        return tag

    # -------------------------------------------------- estado anti-bucle
    @staticmethod
    def _ruta_estado() -> str:
        """Archivo (en datos del usuario, que el instalador NO sobrescribe) donde
        se recuerda el último tag que se intentó aplicar."""
        return os.path.join(rutas.DATOS, "actualizador_estado.json")

    def _tag_ya_aplicado(self, tag: str) -> bool:
        try:
            with open(self._ruta_estado(), encoding="utf-8") as fh:
                return json.load(fh).get("ultimo_tag_aplicado") == tag
        except (OSError, json.JSONDecodeError):
            return False

    def _marcar_tag_aplicado(self, tag: str) -> None:
        try:
            os.makedirs(os.path.dirname(self._ruta_estado()), exist_ok=True)
            with open(self._ruta_estado(), "w", encoding="utf-8") as fh:
                json.dump({"ultimo_tag_aplicado": tag}, fh)
        except OSError:
            pass  # si no se puede escribir, el peor caso es reintentar una vez

    # ------------------------------------------------------- utilidades
    @staticmethod
    def _dir_temporal() -> str:
        """Carpeta escribible para el instalador y el .bat.

        Con `sys.frozen` la app suele vivir en Archivos de Programa (solo lectura
        sin elevación), así que se usa siempre %TEMP%, que es escribible. Se deja
        registrado el modo empaquetado para depuración."""
        _ = getattr(sys, "frozen", False)  # ejecutándose como .exe de PyInstaller
        return tempfile.gettempdir()
