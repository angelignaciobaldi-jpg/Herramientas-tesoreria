"""RPA del SIPP: inicio de sesión y selección de empresa/sucursal (Playwright).

Encapsula en una sola clase reutilizable (`SesionSipp`) el arranque del
navegador, el login en el portal y la selección de empresa y sucursal (selects
tipo "chosen", el plugin de jQuery). La idea es que cualquier módulo que
necesite operar el SIPP reuse esta clase en lugar de duplicar la automatización.

Nota sobre los localizadores: por ser un portal autenticado no se pudo
inspeccionar el DOM real al escribir esto. Los localizadores siguen patrones
estándar (formulario de acceso + estructura del plugin jQuery Chosen) y están
centralizados en constantes/métodos para poder afinarlos tras una corrida de
verificación. Se priorizan localizadores orientados al usuario
(get_by_placeholder / get_by_role / get_by_label / texto) con respaldos por CSS.

Uso típico:

    from core import credenciales
    from core.rpa_sipp import SesionSipp

    async with SesionSipp(headless=False) as sipp:
        usuario, contrasena = credenciales.cargar()
        await sipp.login(usuario, contrasena)
        await sipp.seleccionar_empresa_sucursal("Abastecedora", "Corporativo")
"""

from __future__ import annotations

import asyncio
import concurrent.futures
import os
import re
import sys
import threading

from playwright.async_api import (
    Browser,
    BrowserContext,
    Locator,
    Page,
    Playwright,
    TimeoutError as PlaywrightTimeoutError,
    async_playwright,
)

import glob

from core import rutas


def _ruta_navegadores() -> str:
    """Carpeta (escribible) donde vive Chromium en la app empaquetada."""
    return os.path.join(rutas.DATOS, "ms-playwright")


def _hay_chromium(base: str) -> bool:
    """True si ya hay un Chromium instalado en `base`."""
    return bool(glob.glob(os.path.join(base, "chromium-*", "**", "chrome.exe"), recursive=True))


def necesita_navegador() -> bool:
    """True si la app está empaquetada y aún falta descargar Chromium (primera
    vez). Permite a la interfaz mostrar un aviso antes de la descarga."""
    if not getattr(sys, "frozen", False):
        return False
    return not _hay_chromium(_ruta_navegadores())


async def asegurar_navegador() -> None:
    """En la app empaquetada (sys.frozen): fija PLAYWRIGHT_BROWSERS_PATH a una
    carpeta escribible del usuario y, si Chromium no está, lo descarga (la primera
    vez, requiere internet). En desarrollo no hace nada (usa la instalación normal
    de Playwright).

    El driver (node) viene empaquetado (--collect-all playwright); con él se
    invoca la descarga, igual que haría 'playwright install'."""
    if not getattr(sys, "frozen", False):
        return
    destino = _ruta_navegadores()
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = destino  # antes de usar Playwright
    if _hay_chromium(destino):
        return
    os.makedirs(destino, exist_ok=True)
    # Descarga Chromium (sin headless shell) con el driver node empaquetado.
    from playwright._impl._driver import compute_driver_executable, get_driver_env

    node, cli = compute_driver_executable()
    entorno_driver = {**os.environ, **get_driver_env()}
    entorno_driver["PLAYWRIGHT_BROWSERS_PATH"] = destino
    try:
        proc = await asyncio.create_subprocess_exec(
            node, cli, "install", "chromium", "--no-shell",
            env=entorno_driver,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.STDOUT,
        )
        await proc.communicate()
    except Exception as exc:  # noqa: BLE001 — se reporta como ErrorSipp
        raise ErrorSipp(
            "No se pudo descargar el navegador (Chromium): %s" % exc
        ) from exc
    if not _hay_chromium(destino):
        raise ErrorSipp(
            "No se pudo preparar el navegador (Chromium). Revisa la conexión a "
            "internet e inténtalo de nuevo."
        )


class ErrorSipp(Exception):
    """Falla esperada del RPA del SIPP (login fallido, elemento ausente, etc.)."""


class SesionSipp:
    """Maneja una sesión automatizada del SIPP: navegador, login y selección
    de empresa/sucursal. Pensada para reusarse desde distintos módulos."""

    # --- URLs ---
    BASE_URL = "https://dev.sipp.petroil.dev"
    URL_LOGIN = BASE_URL + "/"
    URL_CONFIG_SESION = BASE_URL + "/index.cfm#/configuracionsession"
    URL_DASHBOARD_TESOR = BASE_URL + "/#/DashboardTesor"

    # --- Tiempos de espera (ms) ---
    TIMEOUT_NAV = 30_000        # navegación / carga de página
    TIMEOUT_ELEMENTO = 10_000   # aparición de un elemento
    TIMEOUT_LOGIN_OK = 5_000    # confirmación de inicio de sesión (requisito: 5 s)

    def __init__(self, headless: bool = False, slow_mo: int = 0):
        self.headless = headless
        self.slow_mo = slow_mo
        self._pw: Playwright | None = None
        self.browser: Browser | None = None
        self.context: BrowserContext | None = None
        self.page: Page | None = None

    # ------------------------------------------------------ ciclo de vida
    async def iniciar(self) -> "SesionSipp":
        """Arranca Playwright, el navegador y una pestaña limpia."""
        await asegurar_navegador()
        self._pw = await async_playwright().start()
        self.browser = await self._pw.chromium.launch(
            headless=self.headless, slow_mo=self.slow_mo,
        )
        self.context = await self.browser.new_context()
        self.page = await self.context.new_page()
        # Trae la ventana/pestaña al frente para que quede enfocada al abrir.
        await self.page.bring_to_front()
        return self

    async def cerrar(self) -> None:
        """Cierra todo de forma segura (idempotente)."""
        if self.context is not None:
            await self.context.close()
        if self.browser is not None:
            await self.browser.close()
        if self._pw is not None:
            await self._pw.stop()
        self._pw = self.browser = self.context = self.page = None

    async def __aenter__(self) -> "SesionSipp":
        return await self.iniciar()

    async def __aexit__(self, *_exc) -> None:
        await self.cerrar()

    # ------------------------------------------------------------ login
    async def login(self, usuario: str, contrasena: str) -> None:
        """Inicia sesión en el portal. Lanza ErrorSipp si faltan credenciales o
        no se confirma el acceso al panel interno en 5 s."""
        if not usuario or not contrasena:
            raise ErrorSipp("Faltan credenciales para iniciar sesión en el SIPP.")
        page = self._exigir_pagina()
        await page.goto(self.URL_LOGIN, wait_until="domcontentloaded", timeout=self.TIMEOUT_NAV)

        # Localizadores verificados contra el DOM real (login.html). Se priorizan
        # los orientados al usuario; el #id queda como respaldo estable.
        campo_usuario = await self._primer_visible(
            [
                page.get_by_placeholder("Usuario", exact=True),
                page.locator("#nb_Usuario"),
                page.get_by_role("textbox", name=re.compile("usuario", re.I)),
            ],
            "campo de usuario",
        )
        await campo_usuario.fill(usuario)

        # Hay varios campos de contraseña (un modal de cambio de contraseña está
        # oculto); _primer_visible se queda con el visible del formulario.
        campo_contrasena = await self._primer_visible(
            [
                page.get_by_placeholder("Contraseña", exact=True),
                page.locator("input[type='password']:visible"),
                page.locator("input[type='password']").first,
            ],
            "campo de contraseña",
        )
        await campo_contrasena.fill(contrasena)

        # Ojo: existe también "Usuario único Petroil" (#btnLoginOAuth); aquí se
        # usa el login con credenciales (#btnLogin, "Iniciar Sesión").
        boton = await self._primer_visible(
            [
                page.get_by_role("button", name=re.compile(r"iniciar sesi", re.I)),
                page.locator("#btnLogin"),
            ],
            "botón de iniciar sesión",
        )
        await boton.click()

        await self._verificar_login()

    async def _verificar_login(self) -> None:
        """Confirma que se entró al panel interno. El acceso correcto redirige
        de login.html a index.cfm; además se busca un elemento exclusivo de la
        sesión iniciada. Lanza ErrorSipp si no se confirma en 5 s."""
        page = self._exigir_pagina()
        # Señal principal: salir de login.html hacia la aplicación (index.cfm).
        try:
            await page.wait_for_url(
                re.compile(r"index\.cfm", re.I), timeout=self.TIMEOUT_LOGIN_OK,
            )
            return
        except PlaywrightTimeoutError:
            pass
        # Respaldo: algún elemento exclusivo del panel interno.
        señales = [
            page.get_by_text(re.compile("bienvenid", re.I)),
            page.get_by_role("link", name=re.compile("salir|cerrar sesi|logout", re.I)),
            page.get_by_role("button", name=re.compile("salir|cerrar sesi|logout", re.I)),
            page.locator("nav, .navbar, #menu, .sidebar, .main-menu").first,
        ]
        try:
            await self._primer_visible(señales, "panel interno", timeout=self.TIMEOUT_LOGIN_OK)
        except ErrorSipp as exc:
            raise ErrorSipp(
                "No se confirmó el inicio de sesión en el SIPP: no se llegó al panel "
                "interno en 5 s. Revisa las credenciales o el localizador de éxito."
            ) from exc

    # ------------------------------------------- empresa / sucursal (chosen)
    async def seleccionar_empresa_sucursal(
        self, empresa: str, sucursal: str, guardar: bool = True,
    ) -> None:
        """Va a la pantalla de configuración de sesión, elige empresa y sucursal
        en los selects 'chosen' y, si `guardar` es True, pulsa "Guardar" para
        que la selección quede activa en la sesión."""
        page = self._exigir_pagina()
        await page.goto(
            self.URL_CONFIG_SESION, wait_until="domcontentloaded", timeout=self.TIMEOUT_NAV,
        )
        # Es una SPA por hash: esperar a que rendericen los selects 'chosen'.
        try:
            await page.locator(".chosen-container").first.wait_for(
                state="visible", timeout=self.TIMEOUT_ELEMENTO,
            )
        except PlaywrightTimeoutError as exc:
            raise ErrorSipp(
                "No se cargaron los selects de empresa/sucursal (chosen) en la "
                "pantalla de configuración de sesión."
            ) from exc

        empresa_cont = await self._contenedor_chosen("Empresa", indice=0)
        await self._seleccionar_chosen(empresa_cont, empresa)

        # La plaza/sucursal arranca deshabilitada (clase 'chosen-disabled') y sus
        # opciones se cargan al elegir empresa: hay que esperar a que se habilite.
        # La etiqueta real es "Plaza (Sucursal):" -> basta con buscar "Sucursal".
        sucursal_cont = await self._contenedor_chosen("Sucursal", indice=1)
        await self._esperar_chosen_habilitado(sucursal_cont)
        await self._seleccionar_chosen(sucursal_cont, sucursal)

        if guardar:
            boton_guardar = await self._primer_visible(
                [
                    page.get_by_role("button", name=re.compile(r"^\s*guardar\s*$", re.I)),
                    page.locator("button:has-text('Guardar')").first,
                ],
                "botón Guardar de la configuración de sesión",
            )
            await boton_guardar.click()

    async def _contenedor_chosen(self, etiqueta: str, indice: int) -> Locator:
        """Devuelve el contenedor .chosen-container del select rotulado con
        `etiqueta`; si no lo halla por etiqueta, cae al de posición `indice`."""
        page = self._exigir_pagina()
        por_etiqueta = page.locator(
            "xpath=//label[contains(normalize-space(.), '%s')]"
            "/following::div[contains(@class,'chosen-container')][1]" % etiqueta
        )
        if await por_etiqueta.count() and await por_etiqueta.first.is_visible():
            return por_etiqueta.first
        contenedores = page.locator(".chosen-container")
        if await contenedores.count() > indice:
            return contenedores.nth(indice)
        raise ErrorSipp("No se encontró el select '%s' (chosen)." % etiqueta)

    async def _seleccionar_chosen(self, contenedor: Locator, texto: str) -> None:
        """Abre un select 'chosen', filtra por `texto` y elige la opción."""
        await contenedor.scroll_into_view_if_needed()
        await contenedor.click()  # abre el desplegable
        # Espera a que el desplegable de ESTE chosen quede abierto y activo.
        drop = contenedor.locator(".chosen-drop")
        try:
            await drop.wait_for(state="visible", timeout=self.TIMEOUT_ELEMENTO)
        except PlaywrightTimeoutError as exc:
            raise ErrorSipp("No se abrió el desplegable del select.") from exc
        # El desplegable incluye un buscador que filtra las opciones al teclear.
        buscador = contenedor.locator(
            ".chosen-search input, input.chosen-search-input"
        ).first
        try:
            await buscador.fill(texto, timeout=2_000)
        except PlaywrightTimeoutError:
            pass  # algunos 'chosen' no tienen buscador; se elige directo
        opcion = contenedor.locator(
            ".chosen-results li.active-result", has_text=texto,
        ).first
        try:
            await opcion.click(timeout=self.TIMEOUT_ELEMENTO)
        except PlaywrightTimeoutError as exc:
            raise ErrorSipp(
                "No se encontró la opción '%s' en el select." % texto
            ) from exc

    async def _esperar_chosen_habilitado(self, contenedor: Locator, timeout: int | None = None) -> None:
        """Espera a que un 'chosen' deje de estar deshabilitado (p. ej. la plaza,
        que se habilita al elegir empresa)."""
        limite = (timeout or self.TIMEOUT_ELEMENTO) / 1000
        fin = asyncio.get_event_loop().time() + limite
        while True:
            clase = await contenedor.get_attribute("class") or ""
            if "chosen-disabled" not in clase:
                return
            if asyncio.get_event_loop().time() >= fin:
                raise ErrorSipp(
                    "La plaza/sucursal no se habilitó tras elegir la empresa."
                )
            await asyncio.sleep(0.25)

    # -------------------------------------------------- navegación / menús
    async def ir_a_registrar_dispersion_no_pemex(self) -> None:
        """Va al Dashboard de Tesorería y abre el menú 'Acciones' para elegir
        'Registrar Dispersión (No Pemex)'."""
        page = self._exigir_pagina()
        # 'domcontentloaded' (no 'networkidle'): el dashboard es una SPA con
        # peticiones en segundo plano que casi nunca dejan la red en reposo, lo
        # que dispararía un delay enorme. La espera real la hace elegir_en_menu,
        # que aguarda a que aparezca el menú 'Acciones' (auto-waiting).
        await page.goto(
            self.URL_DASHBOARD_TESOR, wait_until="domcontentloaded", timeout=self.TIMEOUT_NAV,
        )
        await self.elegir_en_menu("Acciones", "Registrar Dispersión (No Pemex)")

    async def elegir_en_menu(self, menu: str, opcion: str) -> None:
        """Abre un menú desplegable de la navbar (por su texto) y elige una de
        sus opciones (por su texto). Reutilizable para cualquier menú/opción.

        Usa localizadores por rol y EXACTOS: el texto exacto evita coincidencias
        parciales (p. ej. 'Acciones' contenido en 'Refacciones'/'Fracciones')."""
        page = self._exigir_pagina()
        # Abre el menú (el toggle suele ser un <a>; se admite también <button>).
        toggle = await self._primer_visible(
            [
                page.get_by_role("link", name=menu, exact=True),
                page.get_by_role("button", name=menu, exact=True),
            ],
            "menú '%s'" % menu,
        )
        await toggle.click()
        # Elige la opción ya desplegada.
        opcion_loc = page.get_by_role("link", name=opcion, exact=True)
        try:
            await opcion_loc.first.wait_for(state="visible", timeout=self.TIMEOUT_ELEMENTO)
        except PlaywrightTimeoutError as exc:
            raise ErrorSipp(
                "No apareció la opción '%s' en el menú '%s'." % (opcion, menu)
            ) from exc
        await opcion_loc.first.click()

    # --------------------------------------------------------- utilidades
    async def _primer_visible(
        self, locators: list[Locator], descripcion: str, timeout: int | None = None,
    ) -> Locator:
        """Devuelve el primer locator de la lista que se vuelva visible dentro
        del `timeout`. Lanza ErrorSipp si ninguno aparece."""
        limite = (timeout or self.TIMEOUT_ELEMENTO) / 1000
        fin = asyncio.get_event_loop().time() + limite
        while True:
            for loc in locators:
                try:
                    if await loc.count() and await loc.first.is_visible():
                        return loc.first
                except Exception:  # noqa: BLE001 — un locator inválido no debe abortar
                    continue
            if asyncio.get_event_loop().time() >= fin:
                raise ErrorSipp("No se encontró %s en la página." % descripcion)
            await asyncio.sleep(0.25)

    def _exigir_pagina(self) -> Page:
        if self.page is None:
            raise ErrorSipp("La sesión no está iniciada; llama a iniciar() primero.")
        return self.page


class BucleRpa:
    """Bucle de asyncio en un hilo dedicado para correr el RPA.

    Sirve para dos cosas al integrarlo con una GUI (Flet):
      - No congelar la interfaz: el navegador se opera en otro hilo.
      - En Windows, Playwright necesita un ProactorEventLoop para lanzar el
        navegador (subprocesos); `new_event_loop()` lo provee por defecto.

    Todas las corrutinas enviadas corren en el MISMO bucle/hilo, requisito de
    Playwright (sus objetos quedan atados al loop donde se crearon).
    """

    def __init__(self):
        self._loop = asyncio.new_event_loop()
        self._hilo = threading.Thread(target=self._run, name="rpa-loop", daemon=True)
        self._hilo.start()

    def _run(self) -> None:
        asyncio.set_event_loop(self._loop)
        self._loop.run_forever()

    def enviar(self, coro) -> "concurrent.futures.Future":
        """Programa una corrutina en el bucle y devuelve un Future. Desde un
        manejador async de Flet: `await asyncio.wrap_future(bucle.enviar(coro))`."""
        return asyncio.run_coroutine_threadsafe(coro, self._loop)

    def cerrar(self) -> None:
        """Detiene el bucle (el hilo es daemon, así que muere con la app)."""
        self._loop.call_soon_threadsafe(self._loop.stop)


async def _demo() -> None:
    """Prueba manual: login + selección con las credenciales guardadas.

    Ejecutar con:  python -m core.rpa_sipp
    """
    from core import credenciales

    datos = credenciales.cargar()
    if datos is None:
        raise SystemExit(
            "No hay credenciales guardadas. Captúralas en el menú Configuración."
        )
    usuario, contrasena = datos
    async with SesionSipp(headless=False, slow_mo=120) as sipp:
        await sipp.login(usuario, contrasena)
        await sipp.seleccionar_empresa_sucursal("Abastecedora", "Corporativo")
        await sipp.ir_a_registrar_dispersion_no_pemex()
        print("Login, selección y pantalla de Registrar Dispersión (No Pemex) OK.")
        await asyncio.sleep(3)


if __name__ == "__main__":
    asyncio.run(_demo())
