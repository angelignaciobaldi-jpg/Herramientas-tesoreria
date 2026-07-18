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

import asyncio
import os
import sys
import traceback

import flet as ft

# Solo se importa al arranque lo IMPRESCINDIBLE y estable (flet + rutas). El resto
# de módulos (core.db/ocr/preferencias y las pantallas de ui) se importan de forma
# PEREZOSA dentro de las funciones, para que un módulo roto no impida cargar app.py
# y, sobre todo, no impida que corra el auto-updater (que podría traer la corrección).
from core import rutas


class AppTesoreria: 
    """Shell de la aplicación: ventana, encabezado, tema y pestañas."""

    def __init__(self, page: ft.Page):
        self.page = page
        self.picker = ft.FilePicker()
        page.services.append(self.picker)
        self._picker_al_frente()
        self._construir()

    # ------------------------------------------------ diálogos de archivo
    def _picker_al_frente(self) -> None:
        """Envuelve los métodos del FilePicker (elegir/guardar archivo o carpeta)
        para que la ventana del sistema aparezca SIEMPRE por encima de la app: si
        no, en Windows el diálogo nativo suele abrirse detrás de la ventana y el
        usuario cree que no pasó nada. Se pone la ventana en 'topmost' mientras el
        diálogo está abierto (el diálogo es hijo de la ventana, así sube con ella)
        y se restaura al cerrarse."""
        for nombre in ("pick_files", "get_directory_path", "save_file"):
            original = getattr(self.picker, nombre, None)
            if callable(original):
                setattr(self.picker, nombre, self._envolver_al_frente(original))

    def _envolver_al_frente(self, original):
        async def envuelto(*args, **kwargs):
            self._fijar_topmost(True)
            try:
                return await original(*args, **kwargs)
            finally:
                self._fijar_topmost(False)
        return envuelto

    def _fijar_topmost(self, valor: bool) -> None:
        """Pone/quita el 'siempre encima' de la ventana (best-effort: nunca debe
        romper la apertura del diálogo)."""
        try:
            self.page.window.always_on_top = valor
            self.page.update()
        except Exception:  # noqa: BLE001 — el traer-al-frente no es crítico
            pass

    def abrir_en_sistema(self, ruta: str) -> None:
        """Abre un archivo o carpeta en el programa predeterminado (Explorador/visor)
        y lo trae AL FRENTE de la app. En Windows, una ventana recién lanzada no puede
        robar el foco por el 'foreground lock'; además, si la app quedó 'siempre
        encima', el Explorador saldría detrás. Por eso: (1) se quita el topmost y (2)
        se autoriza a la ventana que se abre a tomar el foco (AllowSetForegroundWindow).
        Best-effort: si algo falla, se abre igual con os.startfile."""
        self._fijar_topmost(False)  # por si un diálogo previo lo dejó activo
        if sys.platform == "win32":
            try:
                import ctypes
                # ASFW_ANY (-1): permite que el próximo proceso lanzado tome el foco.
                ctypes.windll.user32.AllowSetForegroundWindow(-1)
            except Exception:  # noqa: BLE001 — el traer-al-frente no es crítico
                pass
        try:
            os.startfile(ruta)  # noqa: S606 — abre en el visor/Explorador predeterminado
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.avisar(f"No se pudo abrir: {exc}", ft.Colors.RED_700)

    # Servicio compartido por todas las pantallas: aviso tipo snackbar. Opcional:
    # un botón de acción (p. ej. "Abrir") con su callback, y una duración mayor
    # para dar tiempo a hacer clic.
    def avisar(self, mensaje: str, color: str | None = None,
               accion: str | None = None, on_accion=None, duracion=None) -> None:
        barra = ft.SnackBar(
            content=ft.Text(mensaje, color=ft.Colors.WHITE), bgcolor=color)
        if accion:
            barra.action = accion
            barra.on_action = on_accion
        if duracion is not None:
            barra.duration = duracion
        self.page.show_dialog(barra)

    def _construir(self) -> None:
        # Import perezoso de las pantallas (ver nota de imports arriba): si una
        # pantalla estuviera rota, el error se contiene en _arrancar_app (que lo
        # muestra en una pantalla clara) en vez de tumbar todo el proceso.
        from ui.alta_beneficiarios import SeccionAltaBeneficiarios
        from ui.cheques import SeccionCheques
        from ui.configuracion import SeccionConfiguracion
        from ui.devoluciones import SeccionDevoluciones
        from ui.dispersion_no_pemex import SeccionDispersionNoPemex

        # Cada pantalla construye su propio contenido.
        self.config = SeccionConfiguracion(self)
        self.alta = SeccionAltaBeneficiarios(self)
        self.devoluciones = SeccionDevoluciones(self)
        self.dispersion_no_pemex = SeccionDispersionNoPemex(self)
        self.cheques = SeccionCheques(self)

        # Área de contenido: las tres pantallas viven aquí; solo se muestra la
        # activa (se alterna 'visible'), en vez de un TabBarView de Material.
        self._secciones = [
            self.alta.contenido,
            self.devoluciones.contenido,
            self.dispersion_no_pemex.contenido,
            self.cheques.contenido,
        ]
        for i, seccion in enumerate(self._secciones):
            seccion.visible = i == 0
        self._area = ft.Column(self._secciones, expand=True)

        # Encabezado: logo (izq), navegación (centro, scroll horizontal) y
        # botones de configuración/tema (der).
        # Estado inicial del tema (recordado entre sesiones, leído en main()): el
        # logo y el botón se inicializan acorde a claro/oscuro.
        oscuro = self.page.theme_mode == ft.ThemeMode.DARK
        self.logo = ft.Image(
            src=self._logo_src(oscuro),
            height=58, fit=ft.BoxFit.CONTAIN,
            error_content=ft.Text("Quetzaltic Solutions", weight=ft.FontWeight.BOLD, size=20),
        )
        # Botón de actualización: busca releases nuevas en GitHub y, si hay,
        # ofrece aplicarla (con reinicio). Se "prende" (naranja) cuando la
        # revisión en segundo plano detecta una versión pendiente.
        self._tag_disponible: str | None = None
        self.btn_actualizar = ft.IconButton(
            icon=ft.Icons.SYSTEM_UPDATE, tooltip="Buscar actualizaciones",
            on_click=self._buscar_actualizacion_manual,
        )
        self.btn_config = ft.IconButton(
            icon=ft.Icons.SETTINGS, tooltip="Configuración", on_click=self.config.abrir,
        )
        self.btn_tema = ft.IconButton(
            icon=ft.Icons.LIGHT_MODE if oscuro else ft.Icons.DARK_MODE,
            tooltip="Modo claro" if oscuro else "Modo oscuro",
            on_click=self._alternar_tema,
        )
        encabezado = ft.Row(
            [
                self.logo,
                self._construir_nav(),
                ft.Row([self.btn_actualizar, self.btn_config, self.btn_tema], tight=True),
            ],
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
            spacing=12,
        )

        # Pie de página: crédito fijo abajo, centrado y discreto. Padding vertical
        # mínimo para que ocupe poca altura.
        pie = ft.Container(
            content=ft.Text(
                "Quetzaltic Solutions - 2026",
                size=11,
                color=ft.Colors.ON_SURFACE_VARIANT,
                text_align=ft.TextAlign.CENTER,
            ),
            alignment=ft.Alignment(0, 0),
            padding=ft.Padding.symmetric(horizontal=6, vertical=1),
        )

        # Quita el splash de arranque justo al pintar la app real (evita ver el
        # splash y la app apilados, y evita un blanco intermedio).
        self.page.controls.clear()
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
            ("Cheques", ft.Icons.REQUEST_QUOTE),
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

    @staticmethod
    def _logo_src(oscuro: bool) -> str:
        return ("Imagenes/Quetzaltic Texto Blanco .png" if oscuro
                else "Imagenes/Quetzaltic Texto negro.png")

    def _alternar_tema(self, _e) -> None:
        """Cambia entre modo claro y oscuro (ajusta logo e ícono) y RECUERDA la
        elección para el próximo arranque."""
        oscuro = self.page.theme_mode != ft.ThemeMode.DARK
        self.page.theme_mode = ft.ThemeMode.DARK if oscuro else ft.ThemeMode.LIGHT
        self.logo.src = self._logo_src(oscuro)
        self.btn_tema.icon = ft.Icons.LIGHT_MODE if oscuro else ft.Icons.DARK_MODE
        self.btn_tema.tooltip = "Modo claro" if oscuro else "Modo oscuro"
        _guardar_tema_oscuro(oscuro)
        self.page.update()

    # ----------------------------------------------------- actualizaciones
    def marcar_actualizacion_disponible(self, tag: str) -> None:
        """Prende el botón de actualización (lo detectó la revisión en 2º plano):
        lo resalta en naranja y ajusta el tooltip. NO instala nada."""
        self._tag_disponible = tag
        self.btn_actualizar.icon_color = ft.Colors.AMBER
        self.btn_actualizar.tooltip = (
            f"Actualización disponible ({tag}) — haz clic para aplicarla"
        )
        self.page.update()

    async def _buscar_actualizacion_manual(self, _e=None) -> None:
        """Al pulsar el botón: revisa GitHub (en un hilo, sin congelar la UI). Si
        hay versión nueva, ofrece aplicarla; si no, avisa que ya está al día."""
        self.btn_actualizar.disabled = True
        self.page.update()
        tag = await asyncio.to_thread(_comprobar_update_sync)
        self.btn_actualizar.disabled = False
        self.page.update()
        if not tag:
            self.avisar("Ya tienes la última versión instalada.", ft.Colors.GREEN_700)
            return
        self.marcar_actualizacion_disponible(tag)
        self._dialogo_actualizacion(tag)

    def _dialogo_actualizacion(self, tag: str) -> None:
        """Diálogo que confirma aplicar la actualización, avisando del reinicio."""
        def aplicar(_e=None) -> None:
            self.page.pop_dialog()
            self.page.run_task(self._aplicar_update, tag)

        self.page.show_dialog(
            ft.AlertDialog(
                modal=True,
                title=ft.Row(
                    [ft.Icon(ft.Icons.SYSTEM_UPDATE, color=ft.Colors.PRIMARY),
                     ft.Text("Actualización disponible", weight=ft.FontWeight.BOLD)],
                    spacing=10,
                ),
                content=ft.Text(
                    f"Hay una nueva versión ({tag}).\n\n"
                    "Al aplicarla, la aplicación se cerrará y se volverá a abrir "
                    "automáticamente ya actualizada. Guarda tus pendientes antes "
                    "de continuar.",
                ),
                actions=[
                    ft.TextButton("Ahora no", on_click=lambda e: self.page.pop_dialog()),
                    ft.FilledButton("Aplicar actualización", icon=ft.Icons.SYSTEM_UPDATE,
                                    on_click=aplicar),
                ],
                actions_alignment=ft.MainAxisAlignment.END,
            )
        )

    async def _aplicar_update(self, tag: str) -> None:
        """Descarga el instalador (en un hilo, con un diálogo de progreso) y lo
        aplica: la app se cierra y se reinicia ya actualizada. Si la descarga
        falla, se conserva la app y su estado (solo avisa)."""
        from core.auto_updater import AutoUpdater

        progreso = ft.AlertDialog(
            modal=True,
            content=ft.Row(
                [
                    ft.ProgressRing(width=28, height=28, stroke_width=3),
                    ft.Text(f"Descargando e instalando la versión {tag}…\n"
                            "La aplicación se reiniciará al terminar."),
                ],
                spacing=16, tight=True,
            ),
        )
        self.page.show_dialog(progreso)
        self.page.update()
        try:
            ruta = await asyncio.to_thread(AutoUpdater().buscar_y_descargar)
        except Exception as exc:  # noqa: BLE001 — se reporta, la app sigue viva
            self.page.pop_dialog()
            self.avisar(f"No se pudo actualizar: {exc}", ft.Colors.RED_700)
            return
        if ruta is None:
            self.page.pop_dialog()
            self.avisar("Ya tienes la última versión instalada.", ft.Colors.GREEN_700)
            return
        # Descarga OK: aplica y cierra la app (no retorna; se reinicia sola).
        AutoUpdater().aplicar_y_salir(ruta)


def _pantalla_cargando(page: ft.Page, titulo: str, mensaje: str) -> None:
    """Splash centrado (spinner + título + mensaje) para que el usuario NUNCA vea
    la ventana en blanco: se usa al arrancar (carga de módulos pesados como OCR/
    Playwright) y durante la búsqueda/descarga de actualizaciones."""
    page.controls.clear()
    page.add(
        ft.Container(
            content=ft.Column(
                [
                    ft.ProgressRing(width=46, height=46, stroke_width=4),
                    ft.Text(titulo, size=20, weight=ft.FontWeight.BOLD,
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


def _pantalla_actualizando(page: ft.Page, mensaje: str) -> None:
    """Splash específico del auto-updater (busca/descarga/instala)."""
    _pantalla_cargando(page, "Actualizando Herramientas Tesorería", mensaje)


async def _buscar_actualizaciones(page: ft.Page) -> bool:
    """Solo en la app empaquetada: busca una versión más nueva en GitHub y, si la
    hay, la descarga y la instala mostrando una pantalla clara; la app se cierra y
    se reinicia sola al terminar. Devuelve True si va a actualizar. Cualquier fallo
    (sin red, sin PAT, etc.) se ignora para no impedir el arranque.

    La parte de red (buscar release + descargar el instalador) corre en un hilo
    (asyncio.to_thread) para NO congelar la interfaz: así el splash y su spinner
    siguen vivos mientras se conecta y descarga."""
    if not getattr(sys, "frozen", False):
        return False  # en desarrollo no se autoactualiza
    try:
        from core import entorno
        from core.auto_updater import AutoUpdater

        if not entorno.github_pat(requerido=False):
            return False  # sin PAT configurado: se omite el chequeo
        _pantalla_actualizando(page, "Buscando actualizaciones…")
        await asyncio.sleep(0.05)  # cede el control para que el splash SÍ se pinte
        up = AutoUpdater()

        # Al iniciar la descarga (callback desde el hilo), se agenda el cambio de
        # mensaje del splash en el loop de la interfaz (page.run_task es seguro
        # entre hilos).
        def al_iniciar_descarga(tag: str) -> None:
            page.run_task(_splash_descargando, page, tag)

        ruta = await asyncio.to_thread(up.buscar_y_descargar, al_iniciar_descarga)
        if ruta is None:
            # NO se limpia el splash: main() lo reemplaza por "Iniciando la
            # aplicación…" (evita un parpadeo en blanco antes de construir).
            return False
        up.aplicar_y_salir(ruta)  # en el hilo principal: sys.exit() cierra bien
        return True
    except Exception:  # noqa: BLE001 — el updater nunca debe tumbar el arranque
        return False


async def _splash_descargando(page: ft.Page, tag: str) -> None:
    _pantalla_actualizando(
        page,
        f"Descargando e instalando la versión {tag}.\n"
        "La aplicación se reiniciará automáticamente al terminar.",
    )


def _comprobar_update_sync() -> str | None:
    """Chequeo SÍNCRONO (pensado para correr en un hilo): devuelve el tag de la
    versión disponible o None. Solo en la app empaquetada con PAT; cualquier fallo
    -> None (no molesta al usuario ni bloquea el arranque)."""
    if not getattr(sys, "frozen", False):
        return None
    try:
        from core import entorno
        from core.auto_updater import AutoUpdater

        if not entorno.github_pat(requerido=False):
            return None
        return AutoUpdater().hay_actualizacion()
    except Exception:  # noqa: BLE001 — el chequeo nunca debe estorbar
        return None


async def _revisar_actualizacion_2do_plano(page: ft.Page, app: "AppTesoreria") -> None:
    """Tras cargar la app, revisa en segundo plano (en un hilo, sin frenar nada)
    si hay una versión nueva; si la hay, prende el botón de actualización para que
    el usuario la aplique cuando quiera (no se instala a la fuerza)."""
    tag = await asyncio.to_thread(_comprobar_update_sync)
    if tag:
        app.marcar_actualizacion_disponible(tag)


_CLAVE_VENTANA = "ventana"
_CLAVE_TEMA = "tema"


def _tema_oscuro_guardado() -> bool:
    """True si la última sesión quedó en modo oscuro (best-effort: si no hay
    preferencia o falla la lectura, arranca en claro)."""
    try:
        from core import preferencias
        return preferencias.cargar_valor(_CLAVE_TEMA) == "oscuro"
    except Exception:  # noqa: BLE001
        return False


def _guardar_tema_oscuro(oscuro: bool) -> None:
    """Recuerda la preferencia de tema (claro/oscuro) para el próximo arranque."""
    try:
        from core import preferencias
        preferencias.guardar_valor(_CLAVE_TEMA, "oscuro" if oscuro else "claro")
    except Exception:  # noqa: BLE001
        pass


def _restaurar_ventana(page: ft.Page) -> None:
    """Aplica el tamaño/posición/maximizado guardados de la última sesión. La
    primera vez (sin estado guardado) abre la ventana maximizada."""
    from core import preferencias

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
    from core import preferencias

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


async def main(page: ft.Page) -> None:
    page.title = "Herramienta Integral de Tesorería"
    # Localización en español (México): los textos internos de los controles nativos
    # de Material (p. ej. el diálogo del DatePicker: título, OK/Cancelar, meses y
    # días) los pone la localización de Flutter según el locale de la página. Sin
    # esto salen en inglés y no concuerdan con el resto de la interfaz.
    page.locale_configuration = ft.LocaleConfiguration(
        supported_locales=[ft.Locale("es", "MX"), ft.Locale("es", "ES"),
                           ft.Locale("en", "US")],
        current_locale=ft.Locale("es", "MX"),
    )
    # Icono de la ventana / barra de tareas (mismo que el del escritorio).
    # Ruta relativa al assets_dir (rutas.BUNDLE), igual que el logo del encabezado.
    page.window.icon = "Imagenes/icon.ico"
    # Padding inferior menor que el resto: el pie va al fondo y, con 18 abajo, el
    # espacio bajo la firma superaba al de encima (spacing de la columna) y el
    # texto quedaba por encima del centro. Se equilibra reduciéndolo.
    page.padding = ft.Padding.only(left=18, right=18, top=18, bottom=10)
    # Tema recordado de la última sesión (claro por defecto).
    page.theme_mode = (
        ft.ThemeMode.DARK if _tema_oscuro_guardado() else ft.ThemeMode.LIGHT
    )
    # Barras de scroll siempre visibles e interactivas (no solo al hacer hover),
    # para que el usuario pueda desplazarse en tablas anchas sin adivinar.
    _barra = ft.ScrollbarTheme(
        thumb_visibility=True, track_visibility=True, thickness=12, interactive=True,
    )
    page.theme = ft.Theme(scrollbar_theme=_barra)
    page.dark_theme = ft.Theme(scrollbar_theme=_barra)
    # Splash inmediato: es lo PRIMERO que se pinta, para que el usuario no vea la
    # ventana en blanco mientras se prepara la app y se busca actualización. El
    # 'await' cede el control para que el cliente ALCANCE a pintarlo antes del
    # trabajo pesado que viene (imports, red); si no, se quedaría en blanco.
    _pantalla_cargando(page, "Herramientas Tesorería", "Iniciando…")
    await asyncio.sleep(0.05)
    # Identidad en la barra de tareas: la ventana la crea un flet.exe genérico
    # (Flet abre el cliente en un proceso aparte); se etiqueta con el icono y el
    # comando de re-lanzamiento correctos para que anclarla funcione bien.
    _configurar_taskbar(page)
    # Restaura el estado de la ventana de la última sesión (o maximizada la 1ra
    # vez) y empieza a recordarlo. Best-effort: nunca debe impedir arrancar.
    try:
        _restaurar_ventana(page)
        _vigilar_ventana(page)
    except Exception:  # noqa: BLE001 — la persistencia de ventana no es crítica
        pass

    # Construcción de la app real. Ya NO se fuerza la actualización al iniciar:
    # la app arranca de inmediato en la versión instalada (arranque ágil) y la
    # revisión de nuevas versiones se hace DESPUÉS, en segundo plano; el usuario
    # decide cuándo aplicarla (botón de la barra superior). Se mantiene el splash
    # "Iniciando la aplicación…" mientras se importan módulos pesados (OCR/
    # Playwright) y se arman las pantallas; el 'await' cede el control para que el
    # splash SÍ se pinte antes de bloquear en la construcción. Si un import/
    # arranque falla, se muestra una pantalla de error clara con opción de
    # reintentar / buscar actualización (por si el fix ya está publicado).
    _pantalla_cargando(page, "Herramientas Tesorería", "Iniciando la aplicación…")
    await asyncio.sleep(0.05)
    try:
        app = _arrancar_app(page)
    except Exception as exc:  # noqa: BLE001 — se reporta en pantalla, no se propaga
        _pantalla_error_arranque(page, exc)
        return

    # Revisión de actualizaciones en SEGUNDO PLANO (no frena el arranque): si hay
    # una versión nueva, prende el botón de actualización para que el usuario la
    # aplique cuando termine sus pendientes.
    page.run_task(_revisar_actualizacion_2do_plano, page, app)


def _configurar_taskbar(page: ft.Page) -> None:
    """Etiqueta la ventana (creada por el flet.exe cliente) con la identidad de la
    app: AppUserModelID + comando/icono de re-lanzamiento. Así, al anclarla a la
    barra de tareas, Windows crea el pin apuntando a Tesoreria.exe con el icono
    correcto (y no al flet.exe genérico con argumentos de sesión). Solo aplica en
    la app empaquetada en Windows; best-effort, nunca impide arrancar."""
    if sys.platform != "win32" or not getattr(sys, "frozen", False):
        return
    try:
        from core import win_taskbar

        exe = os.path.join(rutas.INSTALL, "Tesoreria.exe")
        icono = os.path.join(rutas.INSTALL, "icon.ico")
        win_taskbar.configurar_identidad(
            titulo=page.title,
            relaunch_cmd=f'"{exe}"',
            icon_path=icono if os.path.exists(icono) else None,
            display="Herramientas Tesorería",
        )
    except Exception:  # noqa: BLE001 — la identidad de la barra no es crítica
        pass


def _arrancar_app(page: ft.Page) -> "AppTesoreria":
    """Importa (de forma perezosa) y arranca la app completa. Devuelve la instancia
    de AppTesoreria (para conectarle la revisión de actualizaciones). Se invoca
    dentro de un try/except desde main() para que un módulo roto no tumbe el
    proceso antes de que el updater haya podido correr."""
    from core import db, ocr

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

    return AppTesoreria(page)


def _pantalla_error_arranque(page: ft.Page, exc: Exception) -> None:
    """Pantalla clara cuando la app no puede iniciar (p. ej. un módulo roto).
    Evita el diálogo crudo del sistema y ofrece reintentar (que reintenta el
    chequeo de actualización, por si el fix ya está publicado)."""
    detalle = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))

    async def reintentar(_e=None) -> None:
        _pantalla_cargando(page, "Herramientas Tesorería", "Reintentando…")
        await asyncio.sleep(0.05)
        if await _buscar_actualizaciones(page):
            return
        try:
            _arrancar_app(page)
        except Exception as exc2:  # noqa: BLE001
            _pantalla_error_arranque(page, exc2)

    page.controls.clear()
    page.add(
        ft.Container(
            content=ft.Column(
                [
                    ft.Icon(ft.Icons.ERROR_OUTLINE, size=48, color=ft.Colors.ERROR),
                    ft.Text("No se pudo iniciar la aplicación",
                            size=20, weight=ft.FontWeight.BOLD,
                            text_align=ft.TextAlign.CENTER),
                    ft.Text(
                        "Ocurrió un error al cargar un componente. Si hay una "
                        "actualización disponible, al reintentar se aplicará sola; "
                        "si el problema persiste, reinstala desde el instalador.",
                        size=14, text_align=ft.TextAlign.CENTER,
                        color=ft.Colors.ON_SURFACE_VARIANT,
                    ),
                    ft.FilledButton(
                        "Reintentar / buscar actualización",
                        icon=ft.Icons.REFRESH, on_click=reintentar,
                    ),
                    ft.Divider(),
                    ft.Text("Detalle técnico:", size=12, weight=ft.FontWeight.BOLD),
                    ft.Container(
                        content=ft.Column(
                            [ft.Text(detalle, size=11, selectable=True,
                                     font_family="monospace")],
                            scroll=ft.ScrollMode.AUTO,
                        ),
                        height=200, width=640,
                        bgcolor=ft.Colors.SURFACE_CONTAINER_HIGHEST,
                        border_radius=8, padding=10,
                    ),
                ],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                alignment=ft.MainAxisAlignment.CENTER,
                spacing=14,
                scroll=ft.ScrollMode.AUTO,
            ),
            alignment=ft.Alignment(0, 0),
            padding=24,
            expand=True,
        )
    )
    page.update()


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
