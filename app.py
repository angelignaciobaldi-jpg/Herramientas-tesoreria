"""Herramienta Integral de Tesorería — punto de entrada (shell).

Arma la ventana: encabezado con logo y botón de modo claro/oscuro, y las
pestañas. Cada pestaña es una pantalla en su propio módulo, para que se pueda
trabajar en colaboración sin pisarse:

    ui/alta_beneficiarios.py  -> pestaña "Alta de beneficiarios"
    ui/devoluciones.py        -> pestaña "Generar dispersión devoluciones"
    ui/comun.py               -> constantes y utilidades compartidas

El shell expone a cada pantalla: page, picker (diálogos de archivo) y avisar().
"""

from __future__ import annotations

import os
import sys

import flet as ft

from core import db, ocr, preferencias, rutas
from ui.alta_beneficiarios import SeccionAltaBeneficiarios
from ui.configuracion import SeccionConfiguracion
from ui.devoluciones import SeccionDevoluciones
from ui.dispersion_no_pemex import SeccionDispersionNoPemex


class AppTesoreria: 
    """Shell de la aplicación: ventana, encabezado, tema y pestañas."""

    def __init__(self, page: ft.Page):
        self.page = page
        self.picker = ft.FilePicker()
        page.services.append(self.picker)
        self._construir()

    # Servicio compartido por todas las pantallas: aviso tipo snackbar.
    def avisar(self, mensaje: str, color: str | None = None) -> None:
        self.page.show_dialog(
            ft.SnackBar(content=ft.Text(mensaje, color=ft.Colors.WHITE), bgcolor=color)
        )

    def _construir(self) -> None:
        # Cada pantalla construye su propio contenido.
        self.config = SeccionConfiguracion(self)
        self.alta = SeccionAltaBeneficiarios(self)
        self.devoluciones = SeccionDevoluciones(self)
        self.dispersion_no_pemex = SeccionDispersionNoPemex(self)

        # Área de contenido: las tres pantallas viven aquí; solo se muestra la
        # activa (se alterna 'visible'), en vez de un TabBarView de Material.
        self._secciones = [
            self.alta.contenido,
            self.devoluciones.contenido,
            self.dispersion_no_pemex.contenido,
        ]
        for i, seccion in enumerate(self._secciones):
            seccion.visible = i == 0
        self._area = ft.Column(self._secciones, expand=True)

        # Encabezado: logo (izq), navegación (centro, scroll horizontal) y
        # botones de configuración/tema (der).
        self.logo = ft.Image(
            src="Imagenes/Quetzaltic Texto negro.png",
            height=58, fit=ft.BoxFit.CONTAIN,
            error_content=ft.Text("Quetzaltic Solutions", weight=ft.FontWeight.BOLD, size=20),
        )
        self.btn_config = ft.IconButton(
            icon=ft.Icons.SETTINGS, tooltip="Configuración", on_click=self.config.abrir,
        )
        self.btn_tema = ft.IconButton(
            icon=ft.Icons.DARK_MODE, tooltip="Modo oscuro", on_click=self._alternar_tema,
        )
        encabezado = ft.Row(
            [
                self.logo,
                self._construir_nav(),
                ft.Row([self.btn_config, self.btn_tema], tight=True),
            ],
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=12,
        )

        # Pie de página: crédito fijo abajo, centrado y discreto.
        pie = ft.Container(
            content=ft.Text(
                "Quetzaltic Solutions - 2026",
                size=12,
                color=ft.Colors.ON_SURFACE_VARIANT,
                text_align=ft.TextAlign.CENTER,
            ),
            alignment=ft.Alignment(0, 0),
            padding=6,
        )

        self.page.add(encabezado, self._area, pie)
        # El redimensionado afecta la tabla de la pantalla de alta.
        self.page.on_resize = self.alta._on_resize
        # Ya con la página construida, se cargan los registros guardados.
        self.alta.cargar_desde_db()

    # ------------------------------------------------------ navegación
    def _construir_nav(self) -> ft.Control:
        """Barra de navegación propia dentro del encabezado. Va en un contenedor
        que se expande al centro y hace scroll horizontal si no caben todas las
        opciones (la app es solo para PC, pero la ventana puede achicarse)."""
        self._nav_activa = 0
        self._nav_items: list[dict] = []
        definiciones = [
            ("Alta de beneficiarios", ft.Icons.ACCOUNT_BALANCE),
            ("Generar dispersión devoluciones", ft.Icons.CURRENCY_EXCHANGE),
            ("Dispersión (No Pemex)", ft.Icons.PAYMENTS),
        ]
        controles = []
        for idx, (texto, icono) in enumerate(definiciones):
            ico = ft.Icon(icono, size=18)
            txt = ft.Text(texto, size=13, no_wrap=True)
            cont = ft.Container(
                content=ft.Row(
                    [ico, txt], spacing=8, tight=True,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                ),
                padding=ft.Padding.symmetric(horizontal=16, vertical=12),
                border_radius=8,
                on_click=lambda _e, i=idx: self._seleccionar_nav(i),
                on_hover=lambda e, i=idx: self._hover_nav(i, e.data == "true"),
                animate=ft.Animation(160, ft.AnimationCurve.EASE_OUT),
            )
            self._nav_items.append(
                {"container": cont, "icono": ico, "texto": txt, "hover": False})
            self._estilo_nav(idx)
            controles.append(cont)
        fila = ft.Row(
            controles, scroll=ft.ScrollMode.AUTO, spacing=6,
            alignment=ft.MainAxisAlignment.CENTER,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        # El contenedor se expande para ocupar el hueco entre logo y botones.
        return ft.Container(content=fila, expand=True)

    def _estilo_nav(self, idx: int) -> None:
        """Aplica el estilo del ítem según si está activo o con hover: el activo
        o el que recibe hover se resaltan con color de acento y un subrayado
        inferior. El borde siempre mide 3px (transparente cuando no se resalta)
        para que no salte el layout."""
        item = self._nav_items[idx]
        activo = idx == self._nav_activa
        resaltar = activo or item["hover"]
        color = ft.Colors.PRIMARY if resaltar else ft.Colors.ON_SURFACE_VARIANT
        item["icono"].color = color
        item["texto"].color = color
        item["texto"].weight = ft.FontWeight.BOLD if activo else ft.FontWeight.W_500
        item["container"].border = ft.Border(
            bottom=ft.BorderSide(
                3, ft.Colors.PRIMARY if resaltar else ft.Colors.TRANSPARENT))

    def _hover_nav(self, idx: int, dentro: bool) -> None:
        self._nav_items[idx]["hover"] = dentro
        self._estilo_nav(idx)
        self._nav_items[idx]["container"].update()

    def _seleccionar_nav(self, idx: int) -> None:
        if idx == self._nav_activa:
            return
        anterior = self._nav_activa
        self._nav_activa = idx
        self._estilo_nav(anterior)
        self._estilo_nav(idx)
        # Muestra solo la pantalla elegida.
        for i, seccion in enumerate(self._secciones):
            seccion.visible = i == idx
        self._nav_items[anterior]["container"].update()
        self._nav_items[idx]["container"].update()
        self._area.update()

    def _alternar_tema(self, _e) -> None:
        """Cambia entre modo claro y oscuro (y ajusta el logo y el ícono)."""
        oscuro = self.page.theme_mode != ft.ThemeMode.DARK
        self.page.theme_mode = ft.ThemeMode.DARK if oscuro else ft.ThemeMode.LIGHT
        self.logo.src = (
            "Imagenes/Quetzaltic Texto Blanco .png" if oscuro
            else "Imagenes/Quetzaltic Texto negro.png"
        )
        self.btn_tema.icon = ft.Icons.LIGHT_MODE if oscuro else ft.Icons.DARK_MODE
        self.btn_tema.tooltip = "Modo claro" if oscuro else "Modo oscuro"
        self.page.update()


def _pantalla_actualizando(page: ft.Page, mensaje: str) -> None:
    """Muestra una pantalla centrada de 'actualizando' (con spinner) para que el
    usuario no vea una ventana en blanco mientras se busca/descarga/instala."""
    page.controls.clear()
    page.add(
        ft.Container(
            content=ft.Column(
                [
                    ft.ProgressRing(width=46, height=46, stroke_width=4),
                    ft.Text("Actualizando Herramientas Tesorería",
                            size=20, weight=ft.FontWeight.BOLD,
                            text_align=ft.TextAlign.CENTER),
                    ft.Text(mensaje, size=14, text_align=ft.TextAlign.CENTER,
                            color=ft.Colors.ON_SURFACE_VARIANT),
                ],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=18,
            ),
            alignment=ft.Alignment(0, 0),
            expand=True,
        )
    )
    page.update()


def _buscar_actualizaciones(page: ft.Page) -> bool:
    """Solo en la app empaquetada: busca una versión más nueva en GitHub y, si la
    hay, la descarga y la instala mostrando una pantalla clara; la app se cierra y
    se reinicia sola al terminar. Devuelve True si va a actualizar. Cualquier fallo
    (sin red, sin PAT, etc.) se ignora para no impedir el arranque."""
    if not getattr(sys, "frozen", False):
        return False  # en desarrollo no se autoactualiza
    try:
        from core import entorno
        from core.auto_updater import AutoUpdater

        if not entorno.github_pat(requerido=False):
            return False  # sin PAT configurado: se omite el chequeo
        _pantalla_actualizando(page, "Buscando actualizaciones…")
        actualizo = AutoUpdater().buscar_y_actualizar(
            al_iniciar_descarga=lambda tag: _pantalla_actualizando(
                page,
                f"Descargando e instalando la versión {tag}.\n"
                "La aplicación se reiniciará automáticamente al terminar.",
            )
        )
        if not actualizo:
            page.controls.clear()  # ya está al día: limpia el splash y sigue
            page.update()
        return actualizo
    except Exception:  # noqa: BLE001 — el updater nunca debe tumbar el arranque
        page.controls.clear()
        page.update()
        return False


_CLAVE_VENTANA = "ventana"


def _restaurar_ventana(page: ft.Page) -> None:
    """Aplica el tamaño/posición/maximizado guardados de la última sesión. La
    primera vez (sin estado guardado) abre la ventana maximizada."""
    est = preferencias.cargar_valor(_CLAVE_VENTANA)
    if isinstance(est, dict) and est:
        if est.get("maximized"):
            page.window.maximized = True
        else:
            if est.get("width"):
                page.window.width = est["width"]
            if est.get("height"):
                page.window.height = est["height"]
            if est.get("left") is not None:
                page.window.left = est["left"]
            if est.get("top") is not None:
                page.window.top = est["top"]
    else:
        page.window.maximized = True  # primera vez: maximizada


def _vigilar_ventana(page: ft.Page) -> None:
    """Guarda el estado de la ventana al redimensionar/mover/maximizar, para
    restaurarlo en el próximo arranque. No usa prevent_close (evita el riesgo de
    bloquear el cierre): guarda en los eventos ya 'terminados'."""
    def guardar() -> None:
        est = {
            "width": page.window.width,
            "height": page.window.height,
            "left": page.window.left,
            "top": page.window.top,
            "maximized": bool(page.window.maximized),
        }
        # Maximizada: el ancho/alto/pos serían los de pantalla completa; conserva
        # los últimos valores 'normales' para poder restaurar un tamaño sensato.
        if est["maximized"]:
            prev = preferencias.cargar_valor(_CLAVE_VENTANA)
            if isinstance(prev, dict):
                for k in ("width", "height", "left", "top"):
                    if prev.get(k) is not None:
                        est[k] = prev[k]
        preferencias.guardar_valor(_CLAVE_VENTANA, est)

    def on_event(e) -> None:
        if e.type in (
            ft.WindowEventType.RESIZED, ft.WindowEventType.MOVED,
            ft.WindowEventType.MAXIMIZE, ft.WindowEventType.UNMAXIMIZE,
            ft.WindowEventType.RESTORE,
        ):
            guardar()

    page.window.on_event = on_event


def main(page: ft.Page) -> None:
    page.title = "Herramienta Integral de Tesorería"
    # Icono de la ventana / barra de tareas (mismo que el del escritorio).
    # Ruta relativa al assets_dir (rutas.BUNDLE), igual que el logo del encabezado.
    page.window.icon = "Imagenes/icon.ico"
    page.padding = 18
    page.theme_mode = ft.ThemeMode.LIGHT
    # Barras de scroll siempre visibles e interactivas (no solo al hacer hover),
    # para que el usuario pueda desplazarse en tablas anchas sin adivinar.
    _barra = ft.ScrollbarTheme(
        thumb_visibility=True, track_visibility=True, thickness=12, interactive=True,
    )
    page.theme = ft.Theme(scrollbar_theme=_barra)
    page.dark_theme = ft.Theme(scrollbar_theme=_barra)
    # Restaura el estado de la ventana de la última sesión (o maximizada la 1ra
    # vez) y empieza a recordarlo al redimensionar/mover/maximizar.
    _restaurar_ventana(page)
    _vigilar_ventana(page)

    # Chequeo de actualización con feedback visible (solo empaquetada). Si hay
    # una nueva versión, la app se cierra/reinicia y no se sigue construyendo.
    if _buscar_actualizaciones(page):
        return

    db.inicializar()

    if not ocr.tesseract_disponible():
        page.show_dialog(
            ft.SnackBar(
                content=ft.Text(
                    "No se encontró el motor Tesseract. Los PDF con texto se leerán igual, "
                    "pero los documentos escaneados no podrán procesarse por OCR."
                ),
                bgcolor=ft.Colors.AMBER_800,
            )
        )

    AppTesoreria(page)


if __name__ == "__main__":
    # Al anclar la app a la barra de tareas, Windows crea el acceso directo con
    # el "Iniciar en" vacío, por lo que arranca con el directorio de trabajo en
    # System32 (a diferencia del acceso del escritorio, que arranca en {app}).
    # Eso dejaba la pantalla en blanco al resolverse algo contra el CWD. Fijar
    # el CWD a la carpeta de la app hace que todos los lanzadores (pin, menú
    # inicio, jump list) se comporten igual que el acceso del escritorio.
    try:
        os.chdir(rutas.INSTALL)
    except OSError:
        pass
    ft.run(main, assets_dir=rutas.BUNDLE)
