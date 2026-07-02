"""Pantalla: Dispersión (No Pemex).

Ofrece los controles para operar el RPA del SIPP: un botón que intercala
Iniciar / Pausar / Reanudar y un botón para detener (abortar) la ejecución. Al
iniciar, hace login en el SIPP y selecciona la empresa/sucursal con
core.rpa_sipp.

Las credenciales de inicio de sesión se capturan en el menú "Configuración" de
la barra superior (ver ui/configuracion.py); aquí se leen desde ahí al arrancar.

El RPA corre en un bucle de asyncio en un hilo aparte (BucleRpa) para no
congelar la interfaz y para que Playwright pueda lanzar el navegador en Windows.
Los métodos _pausar_rpa / _reanudar_rpa quedan como puntos de conexión para
cuando exista el proceso de dispersión que se pueda pausar.
"""

from __future__ import annotations

import asyncio

import flet as ft
from playwright.async_api import TimeoutError as PlaywrightTimeoutError

from core.rpa_sipp import (
    BucleRpa,
    ErrorSipp,
    SesionSipp,
    asegurar_navegador,
    necesita_navegador,
)
from ui.comun import GRIS, ROJO, VERDE, tarjeta


class SeccionDispersionNoPemex:
    """Pestaña para operar el RPA de dispersión (No Pemex)."""

    # Empresa y plaza/sucursal donde opera el RPA. Por ahora fijas; más adelante
    # pueden capturarse en la UI o en Configuración.
    EMPRESA = "Abastecedora"
    SUCURSAL = "Corporativo"

    def __init__(self, app):
        self.app = app
        self.page = app.page
        # Estado de la ejecución: "detenido" | "ejecutando" | "pausado".
        self.estado = "detenido"
        self.sesion: SesionSipp | None = None
        self.bucle: BucleRpa | None = None
        self.contenido = self._construir()

    # ------------------------------------------------------------ UI
    def _construir(self) -> ft.Control:
        # Botón que intercala Iniciar / Pausar / Reanudar según el estado.
        self.btn_iniciar = ft.FilledButton(
            content="Iniciar", icon=ft.Icons.PLAY_ARROW, on_click=self._iniciar_pausar,
        )
        # Detener solo se habilita mientras el RPA está en marcha o en pausa.
        self.btn_detener = ft.OutlinedButton(
            content="Detener", icon=ft.Icons.STOP, on_click=self._detener,
            disabled=True,
        )
        self.txt_estado = ft.Text("Estado: Detenido", size=13, color=GRIS)
        # Aviso + barra (indeterminada) para la descarga del navegador la 1ra vez.
        self.txt_install = ft.Text(
            "Instalando componentes extra, espere un momento…",
            size=13, color=GRIS, visible=False,
        )
        self.barra_install = ft.ProgressBar(width=320, visible=False)
        control_card = tarjeta(
            "Control de ejecución",
            ft.Column(
                [
                    ft.Row([self.btn_iniciar, self.btn_detener], spacing=10, wrap=True),
                    self.txt_estado,
                    self.txt_install,
                    self.barra_install,
                ],
                spacing=12,
            ),
        )
        return ft.Column(
            [control_card],
            spacing=14, scroll=ft.ScrollMode.AUTO, expand=True,
        )

    # ----------------------------------------------------- ejecución
    async def _iniciar_pausar(self, _e) -> None:
        """Único botón que arranca, pausa o reanuda según el estado actual."""
        if self.estado == "detenido":
            usuario, contrasena = self.app.config.credenciales()
            if not usuario or not contrasena:
                self.app.avisar(
                    "Captura usuario y contraseña en Configuración.", ROJO)
                return
            # Login + selección tardan unos segundos: bloquea el botón y avisa.
            self.estado = "ejecutando"
            self.btn_iniciar.disabled = True
            self.txt_estado.value = "Estado: Iniciando sesión en el SIPP…"
            self.page.update()
            try:
                # Primera vez: descarga el navegador mostrando aviso + barra.
                if necesita_navegador():
                    self._mostrar_instalacion(True)
                    self.page.update()
                    try:
                        await self._correr(asegurar_navegador())
                    finally:
                        self._mostrar_instalacion(False)
                        self.page.update()
                await self._arrancar_rpa(usuario, contrasena)
            except ErrorSipp as exc:
                await self._abortar_por_error(f"No se pudo iniciar el RPA: {exc}")
                return
            except PlaywrightTimeoutError:
                # Timeout de navegación: casi siempre es la conexión o que el
                # portal del SIPP tardó/no respondió, no un fallo del RPA.
                await self._abortar_por_error(
                    "La página del SIPP tardó demasiado en responder. Suele ser por "
                    "una conexión a internet lenta o inestable (o el portal está "
                    "caído/muy lento). Revisa tu conexión e inténtalo de nuevo."
                )
                return
            except Exception as exc:  # noqa: BLE001 — se reporta al usuario
                await self._abortar_por_error(f"Error inesperado al iniciar el RPA: {exc}")
                return
            self.btn_iniciar.disabled = False
            self.app.avisar(
                "RPA iniciado: en la pantalla de Registrar Dispersión (No Pemex).", VERDE)
        elif self.estado == "ejecutando":
            self.estado = "pausado"
            await self._pausar_rpa()
        else:  # pausado
            self.estado = "ejecutando"
            await self._reanudar_rpa()
        self._refrescar_controles()

    async def _detener(self, _e) -> None:
        """Aborta la ejecución del RPA y vuelve al estado inicial."""
        self.estado = "detenido"
        await self._detener_rpa()
        self.app.avisar("RPA detenido.", ROJO)
        self._refrescar_controles()

    def _mostrar_instalacion(self, visible: bool) -> None:
        """Muestra/oculta el aviso y la barra de descarga del navegador."""
        self.txt_install.visible = visible
        self.barra_install.visible = visible

    async def _abortar_por_error(self, mensaje: str) -> None:
        """Cierra la sesión a medias, vuelve a 'detenido' y reporta el error."""
        self._mostrar_instalacion(False)
        await self._detener_rpa()
        self.estado = "detenido"
        self.btn_iniciar.disabled = False
        self._refrescar_controles()
        self.app.avisar(mensaje, ROJO)

    def _refrescar_controles(self) -> None:
        """Ajusta etiquetas, íconos y disponibilidad de los botones al estado."""
        if self.estado == "ejecutando":
            self.btn_iniciar.content = "Pausar"
            self.btn_iniciar.icon = ft.Icons.PAUSE
            self.btn_detener.disabled = False
            self.txt_estado.value = "Estado: En ejecución"
        elif self.estado == "pausado":
            self.btn_iniciar.content = "Reanudar"
            self.btn_iniciar.icon = ft.Icons.PLAY_ARROW
            self.btn_detener.disabled = False
            self.txt_estado.value = "Estado: En pausa"
        else:  # detenido
            self.btn_iniciar.content = "Iniciar"
            self.btn_iniciar.icon = ft.Icons.PLAY_ARROW
            self.btn_detener.disabled = True
            self.txt_estado.value = "Estado: Detenido"
        self.page.update()

    # ------------------------------------------------- hooks del RPA
    async def _correr(self, coro):
        """Corre una corrutina del RPA en el bucle del hilo dedicado y espera su
        resultado sin congelar la interfaz."""
        if self.bucle is None:
            self.bucle = BucleRpa()
        return await asyncio.wrap_future(self.bucle.enviar(coro))

    async def _arrancar_rpa(self, usuario: str, contrasena: str) -> None:
        """Abre el navegador, inicia sesión y selecciona empresa/sucursal."""
        self.sesion = SesionSipp(headless=False)
        sesion = self.sesion

        async def flujo() -> None:
            await sesion.iniciar()
            await sesion.login(usuario, contrasena)
            await sesion.seleccionar_empresa_sucursal(self.EMPRESA, self.SUCURSAL)
            await sesion.ir_a_registrar_dispersion_no_pemex()

        await self._correr(flujo())

    async def _pausar_rpa(self) -> None:
        # Aún no hay un proceso de dispersión en curso que pausar; el punto de
        # conexión queda listo para cuando exista.
        ...

    async def _reanudar_rpa(self) -> None:
        ...

    async def _detener_rpa(self) -> None:
        """Cierra el navegador y libera la sesión (best-effort)."""
        if self.sesion is None:
            return
        sesion = self.sesion
        self.sesion = None
        try:
            await self._correr(sesion.cerrar())
        except Exception:  # noqa: BLE001 — el cierre no debe propagar errores
            pass
