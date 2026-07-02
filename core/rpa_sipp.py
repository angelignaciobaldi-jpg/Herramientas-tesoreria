"""RPA del SIPP: inicio de sesión y selección de empresa/sucursal (Playwright).

Encapsula en una sola clase reutilizable (`SesionSipp`) el arranque del
navegador, el login en el portal y la selección de empresa y sucursal (selects
tipo "chosen", el plugin de jQuery). La idea es que cualquier módulo que
necesite operar el SIPP reuse esta clase en lugar de duplicar la automatización.

Nota sobre los localizadores: se priorizan los orientados al usuario
(get_by_placeholder / get_by_role / get_by_label / texto) con respaldos por CSS.
Excepción: los selects de empresa/sucursal de la pantalla de configuración de
sesión son selects de AngularJS decorados con el plugin 'chosen'. No tienen id
y la pantalla repite las etiquetas "Empresa:"/"Sucursal:" (un bloque-resumen y
el formulario real), por lo que buscarlos por etiqueta es ambiguo. Se
identifican por su binding único y estable (ng-model="id_Empresa" /
"id_Sucursal") y se opera el <select> nativo directamente (ver
_elegir_opcion_chosen), que es la fuente de verdad que AngularJS termina
guardando.

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
import json
import os
import re
import sys
import threading
from dataclasses import dataclass
from datetime import datetime
from typing import Callable
from urllib.parse import unquote

# Carpeta del proyecto (dos niveles arriba de este archivo: core/ -> raíz). Los
# diagnósticos del RPA se guardan aquí para tenerlos siempre a mano en desarrollo.
_PROYECTO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

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


def _higienizar_nombre(texto: str) -> str:
    """Limpia un texto para usarlo como parte de un nombre de archivo en Windows
    (quita caracteres no permitidos y espacios sobrantes)."""
    limpio = re.sub(r'[<>:"/\\|?*]', "", texto).strip()
    return re.sub(r"\s+", " ", limpio)


def _ruta_unica(ruta: str) -> str:
    """Si `ruta` ya existe, agrega ' (n)' antes de la extensión hasta hallar un
    nombre libre. Evita sobrescribir reportes al descargar varias combinaciones."""
    if not os.path.exists(ruta):
        return ruta
    base, ext = os.path.splitext(ruta)
    n = 2
    while os.path.exists(f"{base} ({n}){ext}"):
        n += 1
    return f"{base} ({n}){ext}"


class ErrorSipp(Exception):
    """Falla esperada del RPA del SIPP (login fallido, elemento ausente, etc.)."""


# JS que elige una opción de un <select> de AngularJS decorado con 'chosen'.
# Recibe {ngModel, texto}: localiza el <select> por su ng-model, busca la opción
# cuyo texto coincida con `texto` (sin acentos ni mayúsculas; exacta, luego
# "empieza con", luego "contiene"), fija el valor y notifica el cambio tanto a
# AngularJS (evento 'change' -> actualiza el ng-model y dispara la recarga de
# sucursales) como al plugin 'chosen' ('chosen:updated' -> refresca el widget).
# Devuelve {ok, elegido, widget} o {ok:false, motivo, disponibles}.
_JS_ELEGIR_OPCION = r"""(args) => {
    const {ngModel, texto} = args;
    const norm = s => (s || '')
        .normalize('NFD').replace(/[̀-ͯ]/g, '')
        .replace(/\s+/g, ' ').trim().toLowerCase();
    const sel = document.querySelector('select[ng-model="' + ngModel + '"]');
    if (!sel) return {ok: false, motivo: 'select-no-encontrado'};
    const objetivo = norm(texto);
    // Ignora el placeholder ("Seleccionar", value '' o '0').
    const opts = Array.from(sel.options).filter(o => o.value !== '' && o.value !== '0');
    let opt = opts.find(o => norm(o.textContent) === objetivo)
           || opts.find(o => norm(o.textContent).startsWith(objetivo))
           || opts.find(o => norm(o.textContent).includes(objetivo));
    if (!opt) return {ok: false, motivo: 'opcion-no-encontrada',
                      disponibles: opts.map(o => o.textContent.trim())};
    sel.value = opt.value;
    const jq = window.jQuery || window.$;
    if (jq) { try { jq(sel).val(opt.value).trigger('change').trigger('chosen:updated'); } catch (e) {} }
    // El 'change' nativo asegura que AngularJS actualice el modelo y su cascada.
    sel.dispatchEvent(new Event('change', {bubbles: true}));
    // Texto que muestra el widget 'chosen' (contenedor hermano del <select>),
    // útil para verificar que la selección se reflejó en la interfaz.
    let widget = '';
    const cont = sel.nextElementSibling;
    if (cont && cont.classList && cont.classList.contains('chosen-container')) {
        const span = cont.querySelector('.chosen-single > span');
        if (span) widget = span.textContent.trim();
    }
    return {ok: true, elegido: opt.textContent.trim(), widget: widget};
}"""


@dataclass
class FiltrosSolicitudPago:
    """Filtros del modal 'Agregar Facturas/Solicitudes de Pago'.

    Todos los campos son opcionales: si vienen vacíos/None, ese filtro NO se toca
    (se deja como está en pantalla). Pensado para que la UI los rellene con lo que
    el usuario capture y deje en blanco lo que no quiera filtrar."""

    empresa: str | None = None
    proveedor: str | None = None
    cuenta_bancaria: str | None = None
    fecha_inicio: str | None = None
    fecha_fin: str | None = None
    folio_solicitud: str | None = None
    tipo_solicitud: str | None = None


class SesionSipp:
    """Maneja una sesión automatizada del SIPP: navegador, login y selección
    de empresa/sucursal. Pensada para reusarse desde distintos módulos."""

    # --- URLs ---
    # Sistema productivo: nuestras actividades (consultar/descargar anexos) no
    # alteran registros reales, así que se opera directo sobre producción.
    # BASE_URL = "https://dev.sipp.petroil.dev"
    BASE_URL = "https://sipp.petroil.com.mx"
    URL_LOGIN = BASE_URL + "/login.html"
    URL_CONFIG_SESION = BASE_URL + "/index.cfm#/configuracionsession"
    URL_DASHBOARD_TESOR = BASE_URL + "/#/DashboardTesor"
    URL_REPORTE_CUENTAS = BASE_URL + "/index.cfm#/ProveedoresCuentasBancariasReporte"

    # --- Tiempos de espera (ms) ---
    TIMEOUT_NAV = 30_000        # navegación / carga de página
    TIMEOUT_ELEMENTO = 10_000   # aparición de un elemento
    TIMEOUT_LOGIN_OK = 5_000    # confirmación de inicio de sesión (requisito: 5 s)

    def __init__(self, headless: bool = False, slow_mo: int = 0, zoom: float = 0.8):
        self.headless = headless
        self.slow_mo = slow_mo
        # Factor de escala de la ventana (< 1 = zoom out). Con la ventana al 100%
        # los últimos registros del reporte quedaban debajo del borde visible y no
        # se podían pulsar; al reducir el zoom, todo cabe y el clic funciona.
        self.zoom = zoom
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
            args=[
                "--start-maximized",
                f"--force-device-scale-factor={self.zoom}",
            ],
        )
        # no_viewport: usa el tamaño real de la ventana maximizada (sin viewport
        # fijo), para que el zoom reducido se traduzca en más espacio útil.
        self.context = await self.browser.new_context(no_viewport=True)
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
    # ng-model de cada select (identificador único y estable en el DOM real).
    _NG_MODEL = {"Empresa": "id_Empresa", "Sucursal": "id_Sucursal"}

    async def seleccionar_empresa_sucursal(
        self, empresa: str, sucursal: str, guardar: bool = True,
    ) -> None:
        """Va a la pantalla de configuración de sesión, elige empresa y sucursal
        y, si `guardar` es True, pulsa "Guardar" para que la selección quede
        activa en la sesión."""
        page = self._exigir_pagina()
        await self._ir_a_ruta_spa(
            self.URL_CONFIG_SESION,
            page.locator(".chosen-container").first,
            "No se cargaron los selects de empresa/sucursal (chosen) en la "
            "pantalla de configuración de sesión. Se guardó un diagnóstico "
            "(captura + HTML) en la carpeta '_diagnostico_rpa' del proyecto.",
            "config_sesion",
        )

        try:
            await self._elegir_opcion_chosen("Empresa", empresa)
            # Al cambiar la empresa, el portal recarga las sucursales por AJAX
            # (red-bind ...@change) y 'chosen' tarda en reconstruir su lista. Con
            # esperar_opcion=True se reintenta hasta que la sucursal buscada
            # exista, en vez de elegir sobre una lista a medio cargar.
            await self._elegir_opcion_chosen("Sucursal", sucursal, esperar_opcion=True)
        except ErrorSipp:
            # No se pudo elegir empresa/sucursal: guarda el DOM real para revisar.
            await self._capturar_diagnostico("seleccion_empresa_sucursal")
            raise

        if guardar:
            boton_guardar = await self._primer_visible(
                [
                    page.get_by_role("button", name=re.compile(r"^\s*guardar\s*$", re.I)),
                    page.locator("button:has-text('Guardar')").first,
                ],
                "botón Guardar de la configuración de sesión",
            )
            await boton_guardar.click()

    async def _elegir_opcion_chosen(
        self, etiqueta: str, texto: str, esperar_opcion: bool = False,
    ) -> None:
        """Elige `texto` en el select de `etiqueta` ('Empresa'/'Sucursal').

        Opera el <select> nativo por su ng-model (fuente de verdad de AngularJS)
        en vez de pelear con el widget 'chosen': evita la ambigüedad de las
        etiquetas duplicadas y la condición de carrera con el filtrado del
        plugin. Si `esperar_opcion` es True, reintenta mientras la opción aún no
        aparezca (útil para la sucursal, que se carga tras elegir empresa). Si el
        camino por JS no logra elegir, cae al respaldo por interfaz.
        """
        page = self._exigir_pagina()
        ng_model = self._NG_MODEL[etiqueta]
        fin = asyncio.get_event_loop().time() + self.TIMEOUT_ELEMENTO / 1000
        ultimo: dict = {}
        while True:
            ultimo = await page.evaluate(
                _JS_ELEGIR_OPCION, {"ngModel": ng_model, "texto": texto}
            )
            if ultimo.get("ok"):
                return
            # 'opcion-no-encontrada'/'select-no-encontrado' pueden ser transitorios
            # mientras Angular carga las sucursales: se reintenta hasta el timeout.
            recuperable = esperar_opcion and ultimo.get("motivo") in (
                "opcion-no-encontrada", "select-no-encontrado",
            )
            if not recuperable or asyncio.get_event_loop().time() >= fin:
                break
            await asyncio.sleep(0.25)

        # Respaldo: operar el widget 'chosen' como un usuario (abrir + teclear).
        if await self._elegir_opcion_chosen_ui(etiqueta, texto):
            return

        disponibles = ultimo.get("disponibles")
        detalle = ""
        if disponibles:
            muestra = ", ".join(disponibles[:8]) + ("…" if len(disponibles) > 8 else "")
            detalle = " Opciones disponibles: " + muestra
        raise ErrorSipp(
            "No se pudo elegir '%s' en el select de %s.%s" % (texto, etiqueta, detalle)
        )

    async def _contenedor_chosen(self, etiqueta: str) -> Locator:
        """Contenedor .chosen-container del select de `etiqueta`, ubicado por el
        ng-model del <select>. El plugin 'chosen' inserta el contenedor como
        hermano inmediato del <select>, así que el selector adyacente es exacto
        y no depende de las etiquetas (que están duplicadas en la pantalla)."""
        page = self._exigir_pagina()
        ng_model = self._NG_MODEL[etiqueta]
        cont = page.locator("select[ng-model='%s'] + .chosen-container" % ng_model)
        if await cont.count():
            return cont.first
        raise ErrorSipp("No se encontró el select de %s (chosen)." % etiqueta)

    async def _elegir_opcion_chosen_ui(self, etiqueta: str, texto: str) -> bool:
        """Respaldo por interfaz: abre el widget 'chosen', teclea el texto (con
        eventos de teclado reales, que es lo que dispara su filtrado) y hace clic
        en la opción, prefiriendo la coincidencia exacta. Devuelve True si logró
        elegir; False si no (para que el llamador reporte el error principal)."""
        try:
            contenedor = await self._contenedor_chosen(etiqueta)
        except ErrorSipp:
            return False
        try:
            await contenedor.scroll_into_view_if_needed()
            await contenedor.click()  # abre el desplegable
            drop = contenedor.locator(".chosen-drop")
            await drop.wait_for(state="visible", timeout=self.TIMEOUT_ELEMENTO)
            buscador = contenedor.locator(
                "input.chosen-search-input, .chosen-search input"
            ).first
            try:
                await buscador.click()
                # press_sequentially emite keydown/keyup reales (a diferencia de
                # fill), necesarios para que 'chosen' filtre las opciones.
                await buscador.press_sequentially(texto, delay=40, timeout=3_000)
            except PlaywrightTimeoutError:
                pass  # algunos 'chosen' no tienen buscador; se elige directo
            exacta = contenedor.locator(
                ".chosen-results li.active-result",
                has_text=re.compile(r"^\s*%s\s*$" % re.escape(texto), re.I),
            )
            opcion = exacta.first if await exacta.count() else contenedor.locator(
                ".chosen-results li.active-result", has_text=texto,
            ).first
            await opcion.click(timeout=self.TIMEOUT_ELEMENTO)
            return True
        except PlaywrightTimeoutError:
            return False

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

    # --------------------------- reporte de cuentas bancarias (anexos)
    async def ir_a_reporte_cuentas_bancarias(self) -> None:
        """Navega a 'Proveedores > Cuentas Bancarias > Reporte' y espera a que
        cargue el filtro de tipo de beneficiario."""
        page = self._exigir_pagina()
        await self._ir_a_ruta_spa(
            self.URL_REPORTE_CUENTAS,
            page.locator("#id_TipoReporte_chosen"),
            "No se cargó la pantalla 'Reporte de Cuentas Bancarias' (no apareció "
            "el filtro de tipo de beneficiario). Se guardó un diagnóstico "
            "(captura + HTML) en la carpeta '_diagnostico_rpa' del proyecto.",
            "reporte_cuentas",
        )

    async def filtrar_reporte_cuentas_bancarias(
        self, tipo_beneficiario: str, fecha_inicio: str, fecha_fin: str,
    ) -> None:
        """En el reporte, elige el tipo de beneficiario (select 'chosen'
        múltiple), captura el rango de fechas y aplica el filtro (lupa)."""
        page = self._exigir_pagina()

        # 1) Tipo de beneficiario (chosen múltiple: #id_TipoReporte_chosen).
        cont = page.locator("#id_TipoReporte_chosen")
        await cont.scroll_into_view_if_needed()
        await cont.click()
        buscador = cont.locator("input.chosen-search-input").first
        try:
            await buscador.fill(tipo_beneficiario, timeout=2_000)
        except PlaywrightTimeoutError:
            pass
        opcion = cont.locator(
            ".chosen-results li.active-result", has_text=tipo_beneficiario,
        ).first
        try:
            await opcion.click(timeout=self.TIMEOUT_ELEMENTO)
        except PlaywrightTimeoutError as exc:
            raise ErrorSipp(
                "No se encontró la opción '%s' en el tipo de beneficiario." % tipo_beneficiario
            ) from exc

        # 2) Rango de fechas (campos con máscara dd/MM/yyyy).
        await self._llenar_fecha("#fh_inicial", fecha_inicio)
        await self._llenar_fecha("#fh_fin", fecha_fin)

        # 3) Aplicar el filtro con el botón de lupa del reporte. En el DOM real es
        # <button ng-click="listar()" class="btn-buscar25p"> (¡ojo!: NO el #lupa
        # del navbar, que es el buscador rápido 'ctrl+b' y no aplica el filtro).
        try:
            boton_buscar = await self._primer_visible(
                [
                    page.locator("button[ng-click='listar()']"),
                    page.locator("button.btn-buscar25p"),
                    page.locator("[ng-click*='listar']"),
                ],
                "botón de lupa (aplicar filtro) del reporte",
            )
            await boton_buscar.click()
        except ErrorSipp:
            # No se ubicó el botón: guarda el DOM real del reporte para afinar.
            await self._capturar_diagnostico("reporte_filtro")
            raise
        # La grid recarga por AJAX; se da un respiro para que pinten los datos.
        await page.wait_for_timeout(2_000)

    async def _llenar_fecha(self, selector: str, valor: str) -> None:
        """Captura una fecha en un campo con máscara dd/MM/yyyy. Recibe la fecha
        en cualquier forma (DD/MM/AAAA o DDMMAAAA): teclea solo los 8 dígitos y
        la máscara agrega las diagonales. Cierra el calendario emergente."""
        page = self._exigir_pagina()
        digitos = re.sub(r"\D", "", valor or "")[:8]
        campo = page.locator(selector)
        await campo.click()
        await page.keyboard.press("Control+A")
        await page.keyboard.press("Delete")
        await campo.type(digitos, delay=30)
        await page.keyboard.press("Escape")  # cierra el datepicker si se abrió

    async def descargar_anexos(
        self, carpeta_destino: str,
        progreso: Callable[[int, int], None] | None = None,
    ) -> tuple[int, int]:
        """Descarga los anexos de TODOS los registros de la grid.

        Cada fila con archivo tiene un botón 'Ver archivo' (ver-archivo >
        button.blue-sip); las filas sin archivo lo traen oculto (ng-hide). Al
        pulsarlo se abre el modal 'Visor de Documentos', cuyo enlace 'Aquí'
        dispara la descarga; luego se cierra el modal y se pasa al siguiente.

        La grid (ngGrid) virtualiza filas: solo mantiene en el DOM las visibles y
        las recicla al hacer scroll. Por eso se recorre el viewport de arriba a
        abajo (ver _recorrer_filas), identificando cada fila por una clave única
        para no repetirla ni saltarla.

        Si se pasa `progreso`, se le llama como progreso(hechos, total) al
        empezar y tras cada registro, para reflejar el avance en la interfaz.

        Devuelve cuántos archivos se descargaron.
        """
        page = self._exigir_pagina()
        os.makedirs(carpeta_destino, exist_ok=True)
        botones_sel = "ver-archivo:not(.ng-hide) button.blue-sip"
        if await page.locator(botones_sel).count() == 0:
            if progreso:
                progreso(0, 0)
            return 0, 0

        # Viewport scrollable que contiene los botones (grid virtualizado ngGrid).
        # Si no se identifica, se opera solo sobre lo visible (sin scroll).
        viewport = page.locator(botones_sel).first.locator(
            "xpath=ancestor::div[contains(@class,'ngViewport')][1]"
        )
        if await viewport.count() == 0:
            viewport = None

        # Pasada previa: cuenta las filas con archivo recorriendo todo el grid,
        # para que la barra de progreso sea exacta (y verifica que el scroll
        # funciona antes de empezar a descargar).
        total = len(await self._recorrer_filas(viewport, botones_sel, None))
        if progreso:
            progreso(0, total)

        descargados = 0
        danados = 0
        usados: set[str] = set()  # nombres ya usados (para no sobrescribir)

        async def descargar(boton: Locator, _clave: str, hechos: int) -> None:
            nonlocal descargados, danados
            beneficiario = await self._beneficiario_de_fila(boton)
            estado = await self._descargar_un_anexo(
                boton, carpeta_destino, hechos - 1, beneficiario, usados,
            )
            if estado == "ok":
                descargados += 1
            elif estado == "invalido":
                danados += 1
            if progreso:
                progreso(min(hechos, total), total)

        await self._recorrer_filas(viewport, botones_sel, descargar)
        if progreso:
            progreso(total, total)
        return descargados, danados

    async def _recorrer_filas(
        self, viewport: Locator | None, botones_sel: str, accion,
    ) -> set[str]:
        """Recorre TODAS las filas del grid virtualizado haciendo scroll por el
        `viewport`, procesando cada fila nueva una sola vez (identificada por una
        clave única de sus celdas). Por cada fila nueva con archivo llama a
        `accion(boton, clave, hechos)`; si `accion` es None, solo las cuenta.
        Devuelve el conjunto de claves vistas. Si `viewport` es None, procesa solo
        las filas visibles (sin scroll)."""
        page = self._exigir_pagina()
        vistas: set[str] = set()
        if viewport is not None:
            try:
                await self._scroll_a(viewport, 0)  # arranca desde arriba
            except Exception:  # noqa: BLE001
                viewport = None
        while True:
            # Procesa (una a una, re-escaneando) las filas nuevas en esta posición;
            # re-escanear evita perder filas si el DOM se recompone tras descargar.
            while await self._procesar_visibles(botones_sel, vistas, accion):
                pass
            if viewport is None:
                break
            try:
                estado = await viewport.evaluate(
                    "v => ({top: v.scrollTop, ch: v.clientHeight, sh: v.scrollHeight})"
                )
            except Exception:  # noqa: BLE001
                break
            if estado["top"] + estado["ch"] >= estado["sh"] - 2:
                break  # se llegó al fondo
            antes = estado["top"]
            await self._scroll_a(viewport, antes + int(estado["ch"] * 0.85))
            despues = await viewport.evaluate("v => v.scrollTop")
            if despues <= antes:
                break  # no se pudo avanzar más (fin o bloqueado)
        return vistas

    async def _procesar_visibles(self, botones_sel: str, vistas: set[str], accion) -> bool:
        """Procesa la PRIMERA fila nueva (con archivo) visible en la posición
        actual del grid y devuelve True (para volver a escanear). Devuelve False
        si no hay ninguna fila nueva visible. `accion(boton, clave, hechos)` puede
        ser None (solo contar)."""
        page = self._exigir_pagina()
        botones = page.locator(botones_sel)
        n = await botones.count()
        for i in range(n):
            boton = botones.nth(i)
            try:
                if not await boton.is_visible():
                    continue
                clave = await self._clave_de_fila(boton)
                if not clave or clave in vistas:
                    continue
                vistas.add(clave)
                if accion is not None:
                    await accion(boton, clave, len(vistas))
                return True
            except Exception:  # noqa: BLE001 — una fila que falle no aborta el resto
                await self._cerrar_modal_visor()
                return True
        return False

    async def _scroll_a(self, viewport: Locator, y: int) -> None:
        """Fija el scroll vertical del viewport y avisa a ngGrid (evento scroll)
        para que renderice las filas de esa posición."""
        await viewport.evaluate(
            "(v, y) => { v.scrollTop = y; v.dispatchEvent(new Event('scroll')); }", y,
        )
        await self._exigir_pagina().wait_for_timeout(350)

    async def _clave_de_fila(self, boton: Locator) -> str:
        """Clave única de la fila del botón: el texto de todas sus celdas unido.
        Sirve para no descargar dos veces la misma fila al reciclarse el DOM."""
        try:
            fila = boton.locator("xpath=ancestor::*[contains(@class,'ngRow')][1]")
            textos = await fila.locator(".ngCellText").all_inner_texts()
            return " | ".join(t.strip() for t in textos)
        except Exception:  # noqa: BLE001
            return ""

    async def _beneficiario_de_fila(self, boton: Locator) -> str:
        """Lee el nombre del beneficiario de la fila del botón 'Ver archivo'. En
        la grid del reporte es la 2ª columna ('Beneficiario'). Devuelve '' si no
        se puede leer (para caer a otro nombre)."""
        try:
            fila = boton.locator("xpath=ancestor::*[contains(@class,'ngRow')][1]")
            celda = fila.locator(".ngCellText").nth(1)
            return (await celda.inner_text()).strip()
        except Exception:  # noqa: BLE001
            return ""

    async def _descargar_un_anexo(
        self, boton: Locator, carpeta_destino: str, indice: int,
        beneficiario: str = "", usados: set[str] | None = None,
    ) -> str:
        """Abre el 'Visor de Documentos' de un registro (modal #idFile_), descarga
        su anexo (validándolo) y cierra el modal.

        El visor muestra el archivo en un <iframe src='...storage...'> (el archivo
        EXACTO que se ve en el SIPP) y ofrece un enlace 'Descargar Aqui' (#bajar >
        a) a 'downloadFile.cfm'. Se baja directo con las cookies de la sesión
        (context.request), probando primero el src del visor y luego el enlace:
        algunos 'downloadFile.cfm' bajan dañados aunque el visor sí muestre bien el
        documento. Se valida cada fuente y se guarda la primera que dé un archivo
        íntegro; si ninguna sirve, se recurre al clic. El archivo se nombra con el
        `beneficiario` (si viene) más la extensión real.

        Devuelve el estado: 'ok' (guardado válido) | 'invalido' (bajó dañado y se
        apartó en '_no_validos') | 'fallo' (no se pudo obtener el archivo)."""
        page = self._exigir_pagina()
        await boton.scroll_into_view_if_needed()
        await boton.click()  # abre el modal 'Visor de Documentos' (#idFile_)

        modal = page.locator("#idFile_")
        enlace = modal.locator("#bajar a, a[href*='downloadFile']").first
        try:
            await enlace.wait_for(state="visible", timeout=self.TIMEOUT_ELEMENTO)
        except PlaywrightTimeoutError:
            # No apareció el visor/enlace esperado: guarda el DOM para revisar.
            await self._capturar_diagnostico("visor_anexo")
            await self._cerrar_modal_visor()
            return "fallo"

        try:
            # Fuentes del archivo, en orden: el src del visor (lo que muestra el
            # SIPP) y el enlace de descarga. URLs absolutas ya resueltas.
            src = None
            try:
                src = await modal.locator("iframe").first.evaluate("f => f.src")
            except Exception:  # noqa: BLE001
                pass
            href = None
            try:
                href = await enlace.evaluate("a => a.href")
            except Exception:  # noqa: BLE001
                pass
            urls = [u for u in (src, href) if u]

            nombre = await self._nombre_anexo(
                modal, href or src or "", indice, beneficiario,
                usados if usados is not None else set(),
            )
            validos, ultimo = await self._bajar_bytes(urls)
            if validos is not None:
                with open(os.path.join(carpeta_destino, nombre), "wb") as fh:
                    fh.write(validos)
                estado = "ok"
            else:
                # Ninguna URL dio un archivo válido: respaldo por clic.
                estado = await self._descargar_por_click(enlace, carpeta_destino, nombre)
                if estado == "fallo" and ultimo is not None:
                    self._guardar_no_valido(carpeta_destino, nombre, ultimo)
                    estado = "invalido"
        finally:
            await self._cerrar_modal_visor()
        return estado

    # Firmas (magic bytes) de los formatos de anexo esperados: imagen o PDF.
    _FIRMAS_ARCHIVO = (
        b"\xff\xd8\xff",          # JPEG
        b"\x89PNG\r\n\x1a\n",     # PNG
        b"%PDF",                   # PDF
        b"II*\x00", b"MM\x00*",   # TIFF (little/big endian)
        b"GIF87a", b"GIF89a",     # GIF
        b"BM",                      # BMP
    )

    @staticmethod
    def _contenido_valido(datos: bytes) -> bool:
        """True si los bytes descargados son una imagen o PDF real; False si vienen
        vacíos, muy cortos o como página de error HTML (respuesta inválida)."""
        if not datos or len(datos) < 100:
            return False
        if datos.lstrip()[:1] == b"<":  # HTML/XML: página de error o de login
            return False
        return any(datos.startswith(f) for f in SesionSipp._FIRMAS_ARCHIVO)

    async def _bajar_bytes(
        self, urls: list[str], reintentos: int = 3,
    ) -> tuple[bytes | None, bytes | None]:
        """Baja bytes VÁLIDOS (imagen/PDF) de la primera de `urls` que funcione,
        con las cookies de la sesión y reintentando ante fallos transitorios (baja
        vacío o como página de error). Devuelve (bytes_validos | None,
        ultimo_no_valido | None) — el segundo sirve para apartar el archivo dañado
        si ninguna fuente dio un archivo íntegro."""
        ultimo: bytes | None = None
        for url in urls:
            for intento in range(reintentos):
                try:
                    resp = await self.context.request.get(url, timeout=self.TIMEOUT_NAV)
                    if resp.ok:
                        datos = await resp.body()
                        if self._contenido_valido(datos):
                            return datos, None
                        if datos:
                            ultimo = datos
                except Exception:  # noqa: BLE001 — se reintenta / se prueba otra url
                    pass
                if intento < reintentos - 1:
                    await self._exigir_pagina().wait_for_timeout(700)
        return None, ultimo

    def _guardar_no_valido(self, carpeta_destino: str, nombre: str, datos: bytes) -> None:
        """Guarda un anexo que no pasó la validación en la subcarpeta '_no_validos'
        (para poder revisarlo sin que estorbe en la carga). Best-effort."""
        try:
            carpeta = os.path.join(carpeta_destino, "_no_validos")
            os.makedirs(carpeta, exist_ok=True)
            with open(os.path.join(carpeta, nombre), "wb") as fh:
                fh.write(datos)
        except Exception:  # noqa: BLE001
            pass

    async def _descargar_por_click(
        self, enlace: Locator, carpeta_destino: str, nombre: str,
    ) -> str:
        """Respaldo: hace clic en el enlace y captura la descarga, validándola.
        Cierra pestañas emergentes si el enlace abrió el archivo en una pestaña
        nueva. Devuelve 'ok' | 'invalido' | 'fallo'."""
        page = self._exigir_pagina()
        try:
            async with page.expect_download(timeout=self.TIMEOUT_NAV) as info:
                await enlace.click()
            descarga = await info.value
            destino = os.path.join(carpeta_destino, nombre or descarga.suggested_filename)
            await descarga.save_as(destino)
            try:
                with open(destino, "rb") as fh:
                    datos = fh.read()
            except Exception:  # noqa: BLE001
                datos = b""
            if self._contenido_valido(datos):
                return "ok"
            # No es válido: se aparta en '_no_validos' y se quita del destino.
            try:
                os.remove(destino)
            except Exception:  # noqa: BLE001
                pass
            if datos:
                self._guardar_no_valido(carpeta_destino, nombre, datos)
                return "invalido"
            return "fallo"
        except PlaywrightTimeoutError:
            for extra in list(self.context.pages)[1:]:
                try:
                    await extra.close()
                except Exception:  # noqa: BLE001
                    pass
            return "fallo"

    async def _nombre_anexo(
        self, modal: Locator, url: str, indice: int,
        beneficiario: str, usados: set[str],
    ) -> str:
        """Nombre con el que guardar el anexo: el nombre del beneficiario más la
        extensión real del archivo. Si no hay beneficiario, usa el nombre
        'amigable' del encabezado del visor, luego el parámetro n= del href y, en
        último caso, un genérico por índice. Evita sobrescribir agregando (2),
        (3)... cuando el nombre ya se usó (p. ej. beneficiario repetido)."""
        # Nombre amigable del encabezado ('Visor de Documentos - <archivo.ext>').
        titulo = ""
        try:
            t = (await modal.locator("#h4").inner_text()).strip()
            if " - " in t:
                titulo = t.split(" - ", 1)[1].strip()
        except Exception:  # noqa: BLE001
            pass

        # Extensión: del título o del parámetro n= del href.
        ext = os.path.splitext(titulo)[1]
        if not ext:
            m = re.search(r"[?&]n=([^&]+)", url or "")
            if m:
                ext = os.path.splitext(unquote(m.group(1)))[1]
        if not ext or len(ext) > 6:
            ext = ".pdf"

        # Base del nombre: beneficiario > amigable(sin ext) > n=(sin ext) > índice.
        base = self._sanear_nombre(beneficiario) if beneficiario else ""
        if not base and titulo:
            base = self._sanear_nombre(os.path.splitext(titulo)[0])
        if not base:
            m = re.search(r"[?&]n=([^&]+)", url or "")
            if m:
                base = self._sanear_nombre(os.path.splitext(unquote(m.group(1)))[0])
        if not base:
            base = f"anexo_{indice + 1}"

        nombre = base + ext
        n = 2
        while nombre.lower() in usados:
            nombre = f"{base} ({n}){ext}"
            n += 1
        usados.add(nombre.lower())
        return nombre

    @staticmethod
    def _sanear_nombre(nombre: str) -> str:
        """Quita de un nombre de archivo los caracteres no válidos en Windows."""
        limpio = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", nombre).strip().strip(".")
        return limpio or "anexo"

    async def _cerrar_modal_visor(self) -> None:
        """Cierra el modal 'Visor de Documentos' (#idFile_). El cierre real es un
        botón con role='closeModal' que invoca fileClose(); si no se puede pulsar,
        se llama a fileClose() directo. Best-effort: nunca propaga errores."""
        page = self._exigir_pagina()
        modal = page.locator("#idFile_")
        try:
            if not await modal.count():
                return
        except Exception:  # noqa: BLE001
            return
        cerrar = page.locator(
            "#idFile_ [role='closeModal'], #idFile_ .btn-cerrar25p, "
            "#idFile_ button[onclick*='fileClose']"
        ).first
        try:
            if await cerrar.count():
                await cerrar.click()
            else:
                await page.evaluate("if (window.fileClose) window.fileClose();")
            await modal.wait_for(state="hidden", timeout=3_000)
            return
        except Exception:  # noqa: BLE001 — se intenta el cierre por función
            pass
        try:
            await page.evaluate("if (window.fileClose) window.fileClose();")
            await modal.wait_for(state="hidden", timeout=3_000)
        except Exception:  # noqa: BLE001
            pass

    async def descargar_anexos_proveedores(
        self, fecha_inicio: str, fecha_fin: str, carpeta_destino: str,
        tipo_beneficiario: str = "Proveedores",
        progreso: Callable[[int, int], None] | None = None,
    ) -> tuple[int, int]:
        """Flujo completo (asumiendo sesión ya iniciada y empresa/sucursal
        seleccionadas): abre el reporte, filtra por tipo de beneficiario y rango
        de fechas, y descarga todos los anexos. Devuelve (descargados, dañados).

        `progreso` (opcional) se reenvía a descargar_anexos para reportar avance.
        Los archivos se guardan en una subcarpeta 'Estados de cuenta proveedores
        - <rango de fechas>' dentro de `carpeta_destino`."""
        await self.ir_a_reporte_cuentas_bancarias()
        await self.filtrar_reporte_cuentas_bancarias(tipo_beneficiario, fecha_inicio, fecha_fin)
        # El grid arranca mostrando solo 10 registros: se sube el tamaño de página
        # para traer TODOS los del filtro a una sola página antes de recorrerla.
        await self._maximizar_pagina()
        carpeta = os.path.join(
            carpeta_destino, self._nombre_carpeta(fecha_inicio, fecha_fin),
        )
        return await self.descargar_anexos(carpeta, progreso=progreso)

    async def _maximizar_pagina(self) -> None:
        """Sube el tamaño de página del grid (select ng-model='pagingOptions.
        pageSize', que arranca en 10) para que todos los registros del filtro
        queden en una sola página. Elige la opción más chica que cubra el total
        (leído del pie 'Elementos Totales: N'); si no lo puede leer, usa la mayor
        disponible. Best-effort: si algo falla, deja el tamaño actual."""
        page = self._exigir_pagina()
        select = page.locator("select[ng-model='pagingOptions.pageSize']").first
        try:
            if not await select.count():
                return
            valores = await select.locator("option").evaluate_all(
                "opts => opts.map(o => o.value)"
            )
            opciones = sorted({int(v) for v in valores if v and v.isdigit()})
            if not opciones:
                return
            total = await self._total_registros_footer()
            objetivo = next((n for n in opciones if n >= total), opciones[-1]) if total \
                else opciones[-1]
            actual = await select.input_value()
            if actual and actual.isdigit() and int(actual) >= objetivo:
                return  # el tamaño actual ya alcanza
            await select.select_option(value=str(objetivo))
            # ngGrid recarga la página (puede pedir datos al servidor): espera.
            await page.wait_for_timeout(2_500)
        except Exception:  # noqa: BLE001 — si no se puede, se queda como está
            pass

    async def _total_registros_footer(self) -> int:
        """Lee el total de registros del pie del grid ('Elementos Totales: N').
        Devuelve 0 si no se puede leer."""
        page = self._exigir_pagina()
        try:
            txt = await page.locator(".ngFooterTotalItems").first.inner_text()
            # 'Elementos Totales: 470' (puede traer también '(Artículos
            # Mostrando: 10)'): se toma el número que sigue a 'Totales'.
            m = re.search(r"totales?\D*([\d.,]+)", txt, re.I)
            if m:
                digitos = re.sub(r"\D", "", m.group(1))
                return int(digitos) if digitos else 0
            nums = [int(n) for n in re.findall(r"\d+", txt)]
            return max(nums) if nums else 0
        except Exception:  # noqa: BLE001
            return 0

    @staticmethod
    def _nombre_carpeta(fecha_inicio: str, fecha_fin: str) -> str:
        """Nombre de la subcarpeta de descarga: 'Estados de cuenta proveedores -
        DD-MM-AAAA a DD-MM-AAAA'."""
        def fmt(f: str) -> str:
            d = re.sub(r"\D", "", f or "")[:8]
            return f"{d[0:2]}-{d[2:4]}-{d[4:8]}" if len(d) == 8 else (f or "")
        return SesionSipp._sanear_nombre(
            f"Estados de cuenta proveedores - {fmt(fecha_inicio)} a {fmt(fecha_fin)}"
        )

    # ------------------------------------------ modal Agregar Solicitudes
    # ng-model de cada control del modal (verificados contra el DOM real).
    _NG_EMPRESA = "filtroSPP_D.id_Empresa"
    _NG_PROVEEDOR = "filtroSPP_D.id_Proveedor"
    _NG_CUENTA = "filtroSPP_D.id_CuentaBancaria"
    _NG_FECHA_INI = "dt_fh_Inicio"
    _NG_FECHA_FIN = "dt_fh_Fin"
    _NG_FOLIO = "filtroSPP_D.fl_FolioSolicitud"
    _NG_TIPO = "filtroSPP_D.id_TipoSolicitudPago"
    # ng-click de los botones de acción del modal (verificados contra el DOM).
    _NG_BUSCAR = "getSolicitudesPagoProveedores_Dispersion()"
    _NG_REPORTE = "generarReporteSolicitudesPagoProveedores_Dispersion()"
    # Nombre de la función del backend que atiende la búsqueda; aparece en la URL
    # del XHR (cfproxy.cfc?_=<func>...). Sirve para esperar la respuesta real.
    _FUNC_BUSCAR = "getSolicitudesPagoProveedores_Dispersion"

    async def abrir_modal_agregar_solicitudes(self) -> None:
        """Pulsa 'Agregar Facturas/Solicitudes de Pago' y espera el modal."""
        page = self._exigir_pagina()
        boton = page.get_by_text("Agregar Facturas/Solicitudes de Pago")
        try:
            await boton.first.click(timeout=self.TIMEOUT_ELEMENTO)
        except PlaywrightTimeoutError as exc:
            raise ErrorSipp(
                "No se encontró el botón 'Agregar Facturas/Solicitudes de Pago'."
            ) from exc
        # El modal está listo cuando aparece el primer filtro (la empresa).
        try:
            await page.locator(f'[ng-model="{self._NG_EMPRESA}"]').first.wait_for(
                state="attached", timeout=self.TIMEOUT_ELEMENTO,
            )
        except PlaywrightTimeoutError as exc:
            raise ErrorSipp("No abrió el modal de Agregar Facturas/Solicitudes de Pago.") from exc

    async def aplicar_filtros_solicitudes(self, filtros: "FiltrosSolicitudPago") -> None:
        """Abre el modal y aplica los filtros capturados. Atajo para el caso de
        una sola pasada; si vas a iterar varias combinaciones reutilizando el
        mismo modal, usa abrir_modal_agregar_solicitudes() + fijar_filtros()."""
        await self.abrir_modal_agregar_solicitudes()
        await self.fijar_filtros(filtros)

    async def fijar_filtros(self, filtros: "FiltrosSolicitudPago") -> None:
        """Aplica solo los filtros que el usuario haya capturado sobre el modal YA
        abierto; los campos vacíos se dejan tal cual (no se tocan). Reutilizable
        para fijar distintas combinaciones sin reabrir el modal."""
        if filtros.empresa:
            await self._set_combo(self._NG_EMPRESA, filtros.empresa)
        if filtros.proveedor:
            await self._set_combo(self._NG_PROVEEDOR, filtros.proveedor)
        if filtros.cuenta_bancaria:
            await self._set_combo(self._NG_CUENTA, filtros.cuenta_bancaria)
        if filtros.fecha_inicio:
            await self._set_fecha(self._NG_FECHA_INI, filtros.fecha_inicio)
        if filtros.fecha_fin:
            await self._set_fecha(self._NG_FECHA_FIN, filtros.fecha_fin)
        if filtros.folio_solicitud:
            await self._set_input(self._NG_FOLIO, filtros.folio_solicitud)
        if filtros.tipo_solicitud:
            await self._set_combo(self._NG_TIPO, filtros.tipo_solicitud)

    async def _set_combo(self, ng_model: str, texto: str) -> None:
        """Elige `texto` en un combo del modal. Los combos son <select> con el
        plugin 'chosen' (el <select> queda oculto y se opera la UI de chosen)."""
        page = self._exigir_pagina()
        select = page.locator(f'[ng-model="{ng_model}"]').first
        # 'chosen' inserta su contenedor como hermano inmediato del <select>.
        chosen = select.locator(
            "xpath=following-sibling::*[contains(@class,'chosen-container')][1]"
        )
        if await chosen.count():
            await self._seleccionar_chosen(chosen.first, texto)
        else:
            await select.select_option(label=texto)

    async def _set_input(self, ng_model: str, valor: str) -> None:
        """Escribe `valor` en un input de texto simple del modal.

        Se filtra por ':visible' porque el modal repite varios ng-model en
        paneles ocultos (ng-hide); hay que tomar la instancia visible."""
        page = self._exigir_pagina()
        campo = page.locator(f'[ng-model="{ng_model}"]:visible').first
        try:
            await campo.fill(valor, timeout=3_000)
        except Exception:  # noqa: BLE001 — fallback: setear por JS y avisar a Angular
            await campo.evaluate(
                "(el, v) => { el.value = v;"
                " el.dispatchEvent(new Event('input', {bubbles:true}));"
                " el.dispatchEvent(new Event('change', {bubbles:true})); }",
                valor,
            )

    async def _set_fecha(self, ng_model: str, valor: str) -> None:
        """Escribe una fecha (DD/MM/AAAA) en el input con máscara (ui-mask).

        Se usa fill, que enfoca el campo SIN dar un clic real: así nunca se abre
        el selector de calendario (un clic normal sí lo abre) y, a la vez, dispara
        el evento 'input' con el que Angular registra el valor (verificado: el
        modelo ng-model queda actualizado).

        Se filtra por ':visible' porque el modal repite el ng-model de fecha en un
        panel oculto (0×0); hay que tomar el input visible, no el primero del DOM."""
        page = self._exigir_pagina()
        campo = page.locator(f'[ng-model="{ng_model}"]:visible').first
        await campo.fill(valor)

    # --------------------------------------------- buscar / descargar reporte
    async def buscar_solicitudes(self) -> int:
        """Pulsa 'Buscar' y espera la RESPUESTA del servidor (no 'networkidle',
        que en esta SPA tardaba ~10 s por el polling de fondo). Devuelve cuántos
        resultados trajo la búsqueda (0 = tabla vacía / sin resultados).

        El conteo se lee del propio payload del XHR (JSON con QUERY.DATA), lo que
        es inmediato y fiable: no depende del DOM ni de filas de una búsqueda
        anterior."""
        page = self._exigir_pagina()
        boton = page.locator(f'[ng-click="{self._NG_BUSCAR}"]:visible').first
        try:
            async with page.expect_response(
                lambda r: self._FUNC_BUSCAR in r.url, timeout=self.TIMEOUT_NAV,
            ) as info:
                await boton.click(timeout=self.TIMEOUT_ELEMENTO)
            resp = await info.value
        except PlaywrightTimeoutError as exc:
            raise ErrorSipp(
                "No respondió la búsqueda de solicitudes (Buscar)."
            ) from exc
        try:
            datos = json.loads(await resp.text())
            return len(datos.get("QUERY", {}).get("DATA", []))
        except Exception:  # noqa: BLE001 — si no se pudo leer, cae al DOM
            return await self._contar_filas_dom()

    async def _contar_filas_dom(self) -> int:
        """Respaldo: cuenta filas de datos de la tabla de resultados del modal
        (la tabla visible cuyo encabezado incluye 'Folio Factura')."""
        page = self._exigir_pagina()
        return await page.evaluate(
            r"""() => {
              const t = [...document.querySelectorAll('table')].find(
                t => t.offsetParent &&
                     /Folio\s*Factura/i.test((t.querySelector('thead')||{}).textContent||''));
              if (!t) return 0;
              return [...t.querySelectorAll('tbody tr')].filter(
                tr => tr.offsetParent && tr.querySelector('input[type=checkbox]')).length;
            }"""
        )

    async def descargar_reporte_excel(
        self, ruta_destino: str | None = None, prefijo: str | None = None,
    ) -> str:
        """Pulsa 'Generar XLS', confirma el aviso y guarda el Excel descargado.

        El botón 'Generar XLS' no descarga directamente: abre un modal de
        confirmación (HTML) con un botón 'Aceptar'; la descarga arranca al
        aceptarlo. Devuelve la ruta del archivo guardado. Si `ruta_destino` es
        None, lo guarda en una carpeta escribible (DATOS/descargas) con el nombre
        sugerido por el servidor (con `prefijo` opcional delante) y, si ya existe,
        agrega un sufijo numérico para no sobrescribir. Solo se descarga: NO se
        procesa el contenido (eso queda para más adelante)."""
        page = self._exigir_pagina()
        boton = page.locator(f'[ng-click="{self._NG_REPORTE}"]:visible').first
        try:
            await boton.click(timeout=self.TIMEOUT_ELEMENTO)
        except PlaywrightTimeoutError as exc:
            raise ErrorSipp("No se encontró el botón 'Generar XLS'.") from exc

        # Confirmar el aviso (modal HTML): el 'Aceptar' dispara la descarga.
        try:
            async with page.expect_download(timeout=self.TIMEOUT_NAV) as info:
                await self.confirmar_aviso()
            descarga = await info.value
        except PlaywrightTimeoutError as exc:
            raise ErrorSipp(
                "No se generó/descargó el reporte de solicitudes (Excel)."
            ) from exc
        if ruta_destino is None:
            carpeta = os.path.join(rutas.DATOS, "descargas")
            os.makedirs(carpeta, exist_ok=True)
            nombre = descarga.suggested_filename
            if prefijo:
                nombre = f"{_higienizar_nombre(prefijo)} - {nombre}"
            ruta_destino = _ruta_unica(os.path.join(carpeta, nombre))
        else:
            carpeta = os.path.dirname(ruta_destino)
            if carpeta:
                os.makedirs(carpeta, exist_ok=True)
        await descarga.save_as(ruta_destino)
        return ruta_destino

    async def generar_reporte_solicitudes(
        self, ruta_destino: str | None = None, prefijo: str | None = None,
    ) -> str | None:
        """Atajo: pulsa 'Buscar' y, SOLO si hay resultados, descarga el Excel.

        Si la búsqueda no trae resultados (tabla vacía), NO genera el XLS (se
        evita el aviso 'No se Encontraron Registros') y devuelve None. Si sí hay,
        devuelve la ruta del archivo descargado."""
        if await self.buscar_solicitudes() == 0:
            return None
        return await self.descargar_reporte_excel(ruta_destino, prefijo=prefijo)

    async def confirmar_aviso(self) -> None:
        """Pulsa 'Aceptar' en el modal de confirmación (HTML, no nativo) que el
        portal muestra antes de ciertas acciones. Reutilizable para cualquier
        aviso con un botón 'Aceptar'."""
        page = self._exigir_pagina()
        aceptar = await self._primer_visible(
            [
                page.get_by_role("button", name=re.compile(r"^\s*aceptar\s*$", re.I)),
                page.locator("button.btn-primary:has-text('Aceptar')"),
            ],
            "botón 'Aceptar' del aviso de confirmación",
        )
        await aceptar.click(timeout=self.TIMEOUT_ELEMENTO)

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

    async def _ir_a_ruta_spa(
        self, url: str, ancla: Locator, error: str, diagnostico: str,
    ) -> None:
        """Navega a una ruta de la SPA (AngularJS + ui-router, ruteo por hash) y
        espera a que aparezca `ancla` (un elemento propio de esa vista).

        Ojo: en esta SPA cambiar SOLO el hash (goto/recarga) no dispara la
        transición de estado de ui-router; la vista (ui-view) se queda en la
        anterior (p. ej. la de bienvenida). La forma fiable es navegar como un
        usuario: disparar el ng-click del enlace del menú y forzar $location con
        un digest. Si aun así no aparece la vista, se intenta el respaldo duro
        (goto + recarga) y, en última instancia, se guarda diagnóstico y se lanza
        ErrorSipp."""
        page = self._exigir_pagina()
        ruta = "/" + url.split("#/", 1)[1] if "#/" in url else "/"

        # 1) Vía Angular: enlace del menú + $location (lo que hace un usuario).
        await self._navegar_angular(ruta)
        try:
            await ancla.wait_for(state="visible", timeout=self.TIMEOUT_ELEMENTO)
            return
        except PlaywrightTimeoutError:
            pass

        # 2) Respaldo: navegación dura (goto + recarga en la ruta destino).
        await page.goto(url, wait_until="domcontentloaded", timeout=self.TIMEOUT_NAV)
        try:
            await ancla.wait_for(state="visible", timeout=self.TIMEOUT_ELEMENTO)
            return
        except PlaywrightTimeoutError:
            pass
        await page.reload(wait_until="domcontentloaded", timeout=self.TIMEOUT_NAV)
        try:
            await ancla.wait_for(state="visible", timeout=self.TIMEOUT_ELEMENTO)
        except PlaywrightTimeoutError as exc:
            await self._capturar_diagnostico(diagnostico)
            raise ErrorSipp(error) from exc

    async def _navegar_angular(self, ruta: str) -> None:
        """Cambia la ruta de la SPA como lo haría un usuario: dispara el ng-click
        del enlace del menú (href='#<ruta>', que además guarda la selección) y,
        por si el clic no basta, fuerza el cambio con el servicio $location de
        AngularJS y un digest, que es lo que realmente activa la transición de
        ui-view. Best-effort: si Angular no está accesible, no hace nada (queda el
        respaldo de goto/recarga)."""
        page = self._exigir_pagina()
        try:
            await page.evaluate(
                """(ruta) => {
                    // 1) Disparar el enlace real del menú (incluye su ng-click).
                    var a = document.querySelector("a[href='#" + ruta + "']");
                    if (a) { a.click(); }
                    // 2) Asegurar el cambio de ruta vía $location + digest.
                    try {
                        if (window.angular) {
                            var inj = angular.element(document.body).injector()
                                   || angular.element(document.documentElement).injector();
                            if (inj) {
                                inj.get('$location').url(ruta);
                                inj.get('$rootScope').$apply();
                            }
                        }
                    } catch (e) { /* queda el respaldo de goto/recarga */ }
                    return true;
                }""",
                ruta,
            )
        except Exception:  # noqa: BLE001 — si falla, queda el respaldo
            pass
        # Da un respiro a Angular para resolver la ruta y renderizar la vista
        # (incluida la inicialización del plugin 'chosen' de los selects).
        await page.wait_for_timeout(1_000)

    def _exigir_pagina(self) -> Page:
        if self.page is None:
            raise ErrorSipp("La sesión no está iniciada; llama a iniciar() primero.")
        return self.page

    async def _capturar_diagnostico(self, prefijo: str) -> None:
        """Guarda una captura de pantalla + el HTML actual de la página en la
        carpeta '_diagnostico_rpa' (junto a la app). Sirve para inspeccionar el
        DOM real cuando un localizador falla. Best-effort: nunca propaga errores
        para no enmascarar la falla original que se está reportando."""
        if self.page is None:
            return
        carpeta = os.path.join(_PROYECTO, "_diagnostico_rpa")
        marca = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = os.path.join(carpeta, f"{prefijo}_{marca}")
        try:
            os.makedirs(carpeta, exist_ok=True)
            await self.page.screenshot(path=base + ".png", full_page=True)
            html = await self.page.content()
            with open(base + ".html", "w", encoding="utf-8") as fh:
                fh.write(html)
        except Exception:  # noqa: BLE001 — el diagnóstico no debe romper nada
            pass


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
