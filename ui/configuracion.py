"""Configuración: credenciales de inicio de sesión del SIPP (para el RPA).

Se abre como diálogo desde el botón de la barra superior. Captura usuario y
contraseña, que siempre se guardan localmente con la contraseña cifrada (ver
core/credenciales.py). Otras pantallas (p. ej. el RPA de dispersión) leen estas
credenciales con el método credenciales().
"""

from __future__ import annotations

import asyncio

import flet as ft

from core import ajustes_api, credenciales, cuentas_bancarias, cuentas_dispersion
from ui.comun import GRIS, ROJO, VERDE, tarjeta

CENTRO = ft.Alignment(0, 0)
# Margen ("canalón") para que la barra de scroll vertical no tape el contenido
# (mismo criterio que la pantalla de Dispersión).
_GUTTER_SCROLL = 14
# A partir de este ancho de PANTALLA (px) los dos grupos van lado a lado (50/50);
# por debajo, se apilan verticalmente.
_BREAKPOINT = 1000
# Ancho del diálogo con 2 columnas / 1 columna (topado a ~0.92 del ancho de pantalla).
_ANCHO_2COL = 960
_ANCHO_1COL = 460
# La altura del modal se ajusta al contenido, entre el 40% y el 80% de la pantalla.
_ALTO_MIN_FRAC = 0.40
_ALTO_MAX_FRAC = 0.80


class SeccionConfiguracion:
    """Diálogo de configuración con las credenciales del SIPP."""

    def __init__(self, app):
        self.app = app
        self.page = app.page
        self._abierto = False        # el diálogo está mostrándose
        self._alto_contenido = 0.0   # alto natural del contenido (medido en vivo)
        self._construir()
        self._cargar_credenciales()
        self._cargar_ajustes_api()

    # --------------------------------------------------- helpers de layout
    @staticmethod
    def _apartado(titulo: str, ayuda: "str | None", *controles) -> ft.Control:
        """Un sub-apartado: título y —si tiene descripción— un ícono de ayuda cuyo
        tooltip la muestra al instante (en vez de un subtítulo que alarga el modal)."""
        encabezado = [ft.Text(titulo, size=14, weight=ft.FontWeight.BOLD)]
        if ayuda:
            encabezado.append(ft.Icon(
                ft.Icons.HELP_OUTLINE, size=18, color=GRIS,
                tooltip=ft.Tooltip(
                    message=ayuda, wait_duration=ft.Duration(milliseconds=0))))
        return ft.Column(
            [ft.Row(encabezado, spacing=6, tight=True,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER),
             *controles],
            spacing=8, tight=True)

    def _recalcular_layout(self) -> None:
        """Coloca los grupos lado a lado (50/50) si la PANTALLA es más ancha que
        _BREAKPOINT, o apilados si no; ajusta el ancho del diálogo y la altura."""
        w = self.page.width or 1200
        dos_col = w > _BREAKPOINT
        self._grupo_sistema.expand = dos_col      # en Row, expand=True -> 50/50
        self._grupo_catalogos.expand = dos_col
        if dos_col:
            self._holder.content = ft.Row(
                [self._grupo_sistema, self._grupo_catalogos], spacing=14,
                vertical_alignment=ft.CrossAxisAlignment.START)
            ancho = min(_ANCHO_2COL, int(w * 0.92))
        else:
            self._holder.content = ft.Column(
                [self._grupo_sistema, self._grupo_catalogos], spacing=14, tight=True,
                horizontal_alignment=ft.CrossAxisAlignment.STRETCH)
            ancho = min(_ANCHO_1COL, int(w * 0.92))
        self._cont_alto.width = ancho
        self._titulo_row.width = ancho
        self._ajustar_alto_valor()

    def _ajustar_alto(self, e) -> None:
        """on_size_change del contenido: guarda su alto natural y re-ajusta."""
        self._alto_contenido = float(getattr(e, "height", 0) or 0)
        self._ajustar_alto_valor()

    def _ajustar_alto_valor(self) -> None:
        """Altura del modal = alto del contenido, acotada al 40%–80% de la pantalla."""
        ph = self.page.height or 800
        alto = self._alto_contenido or (_ALTO_MIN_FRAC * ph)
        self._cont_alto.height = min(
            max(alto, _ALTO_MIN_FRAC * ph), _ALTO_MAX_FRAC * ph)
        self._safe_update()

    def _on_resize(self, _e=None) -> None:
        """Registrado en app.registrar_on_resize: reacomoda el modal si está abierto."""
        if not self._abierto:
            return
        self._recalcular_layout()
        self._safe_update()

    def _mostrar_procesando(self, visible: bool,
                            texto: str = "Procesando archivo…") -> None:
        """Muestra/oculta el overlay de espera sobre el contenido del modal."""
        self._overlay_texto.value = texto
        self._overlay.visible = visible
        self._safe_update()

    def _safe_update(self) -> None:
        # Actualiza SOLO el diálogo (no toda la página): así el overlay, la altura y
        # el reacomodo del modal no fuerzan un re-render del app completo (que puede
        # ser pesado si una pantalla de fondo tiene muchas tablas/filas y bloquea la UI).
        try:
            self.dialogo.update()
        except (RuntimeError, AssertionError, AttributeError):
            pass  # aún no montado; se reflejará al renderizar

    # ------------------------------------------------------------ UI
    def _construir(self) -> None:
        # --- Controles (expand: llenan el ancho de su columna, responsivos) ---
        self.tf_usuario = ft.TextField(
            label="Usuario", dense=True, content_padding=10, expand=True)
        self.tf_contrasena = ft.TextField(
            label="Contraseña", password=True, can_reveal_password=False,
            dense=True, content_padding=10, expand=True)
        self.tf_api_url = ft.TextField(
            label="URL base de la API", dense=True, content_padding=10,
            hint_text="https://api.quetzaltic.dev", expand=True)
        self.tf_api_token = ft.TextField(
            label="Token de la API", password=True, can_reveal_password=True,
            dense=True, content_padding=10,
            hint_text="Déjalo vacío para conservar el actual", expand=True)
        self.txt_estado_cuentas = ft.Text(size=12)
        self.txt_estado_cuentas_disp = ft.Text(size=12)
        self.txt_api_token_estado = ft.Text(size=12)
        self._actualizar_estado_cuentas()
        self._actualizar_estado_cuentas_dispersion()
        self._actualizar_estado_token()

        # --- Sub-apartados (la descripción va como tooltip de un ícono de ayuda) ---
        cred = self._apartado(
            "Credenciales SIPP", None, self.tf_usuario, self.tf_contrasena)
        api = self._apartado(
            "Configuración de API",
            "URL y token de los microservicios. El token se guarda cifrado en este "
            "equipo (DPAPI); nunca en claro ni en la instalación.",
            self.tf_api_url, self.tf_api_token,
            ft.Row(
                [self.txt_api_token_estado,
                 ft.TextButton("Quitar token", on_click=self._quitar_token)],
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                vertical_alignment=ft.CrossAxisAlignment.CENTER))
        calib = self._apartado(
            "Hoja de calibración de cheques",
            "Imprime una hoja con cuadrícula milimétrica y las esquinas identificadas "
            "(SI/SD/II/ID) para calibrar la posición del cheque. Sale a escala real (1:1).",
            ft.OutlinedButton(
                "Imprimir hoja de calibración", icon=ft.Icons.PRINT,
                on_click=self._imprimir_calibracion, expand=True))
        cuentas = self._apartado(
            "Catálogo de Cuentas",
            "Adjunta el Excel de cuentas bancarias; se guardará donde la app lo "
            "consulta y reemplazará al anterior si ya había uno.",
            ft.OutlinedButton(
                "Adjuntar Excel de cuentas", icon=ft.Icons.UPLOAD_FILE,
                on_click=self._adjuntar_cuentas, expand=True),
            self.txt_estado_cuentas)
        cuentas_disp = self._apartado(
            "Catálogo de Cuentas Dispersión",
            "Adjunta el Excel con las columnas 'id Empresa', 'Cuenta' y 'CLABE "
            "interbancaria' (opcional). Determina las cuentas y CLABEs que aparecen "
            "en los selectores de la pantalla de Dispersión (por empresa).",
            ft.OutlinedButton(
                "Adjuntar Excel de cuentas de dispersión", icon=ft.Icons.UPLOAD_FILE,
                on_click=self._adjuntar_cuentas_dispersion, expand=True),
            self.txt_estado_cuentas_disp)

        # --- Grupos (tarjeta: título + cuerpo) ---
        self._grupo_sistema = tarjeta("Sistema", ft.Column(
            [cred, ft.Divider(), api, ft.Divider(), calib], spacing=14, tight=True))
        self._grupo_catalogos = tarjeta("Catálogos", ft.Column(
            [cuentas, ft.Divider(), cuentas_disp], spacing=14, tight=True))

        # --- Holder responsivo (Row 50/50 o Column apilada) + medición de altura.
        # on_size_change mide el alto natural del contenido para acotar el del modal.
        self._holder = ft.Container(on_size_change=self._ajustar_alto)
        contenido_scroll = ft.Column(
            [ft.Container(self._holder,
                         padding=ft.Padding.only(right=_GUTTER_SCROLL))],
            scroll=ft.ScrollMode.AUTO, tight=True)
        # Contenedor con la ALTURA acotada; su ancho/alto se fijan en _recalcular_layout.
        self._cont_alto = ft.Container(contenido_scroll)

        # --- Overlay "Procesando…" (waitscreen al subir catálogos) ---
        self._overlay_texto = ft.Text("Procesando archivo…", size=14, color=GRIS)
        self._overlay = ft.Container(
            ft.Column(
                [ft.ProgressRing(), self._overlay_texto],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                alignment=ft.MainAxisAlignment.CENTER, spacing=14),
            alignment=CENTRO, visible=False, expand=True, border_radius=8,
            bgcolor=ft.Colors.with_opacity(0.75, ft.Colors.SURFACE))

        # Encabezado: título + botón "X". Su ancho se fija en _recalcular_layout.
        self._titulo_row = ft.Row(
            [ft.Text("Configuración", size=22, weight=ft.FontWeight.BOLD),
             ft.IconButton(icon=ft.Icons.CLOSE, tooltip="Cerrar",
                           on_click=self._cerrar)],
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            vertical_alignment=ft.CrossAxisAlignment.CENTER)

        self.dialogo = ft.AlertDialog(
            modal=True,
            title=self._titulo_row,
            content=ft.Stack([self._cont_alto, self._overlay]),
            actions=[
                ft.FilledButton("Aceptar", icon=ft.Icons.CHECK, on_click=self._guardar),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self._recalcular_layout()

    # -------------------------------------------------------- acciones
    def abrir(self, _e=None) -> None:
        self._abierto = True
        self._recalcular_layout()  # ya con page.width/height reales
        self.page.show_dialog(self.dialogo)

    def _cerrar(self, _e=None) -> None:
        self._abierto = False
        self.page.pop_dialog()

    # ---------------------------------------------- impresión de cheques
    def _imprimir_calibracion(self, _e=None) -> None:
        """Abre el diálogo de impresión para la hoja de calibración."""
        from core import impresion
        from ui.dialogo_impresion import DialogoImpresion

        DialogoImpresion(
            self.app,
            titulo="Imprimir hoja de calibración",
            mensaje=(
                "Se imprimirá una hoja de calibración para cheques.\n"
                "¿Realizar la impresión?"
            ),
            clave_pref="impresora_cheques",
            al_imprimir=impresion.imprimir_hoja_calibracion,
        ).abrir()

    def _guardar(self, _e=None) -> None:
        """Guarda credenciales (contraseña cifrada) y ajustes de la API (URL como
        preferencia; token cifrado con DPAPI solo si se capturó uno nuevo)."""
        usuario, contrasena = self.credenciales()
        credenciales.guardar(usuario, contrasena)
        ajustes_api.guardar_base_url(self.tf_api_url.value or "")
        token = (self.tf_api_token.value or "").strip()
        if token:  # vacío -> se conserva el token guardado (no se borra al guardar)
            ajustes_api.guardar_token(token)
        self._cerrar()
        self.app.avisar("Configuración guardada.", VERDE)

    # -------------------------------------------------- integración (API)
    def _cargar_ajustes_api(self) -> None:
        """Precarga la URL base guardada (el token NO se muestra por seguridad)."""
        self.tf_api_url.value = ajustes_api.base_url() or ""

    def _actualizar_estado_token(self) -> None:
        """Refleja si hay un token guardado localmente (sin mostrarlo)."""
        if ajustes_api.hay_token_local():
            self.txt_api_token_estado.value = "Token guardado ✓"
            self.txt_api_token_estado.color = VERDE
        else:
            self.txt_api_token_estado.value = "Sin token guardado."
            self.txt_api_token_estado.color = GRIS

    def _quitar_token(self, _e=None) -> None:
        """Elimina el token guardado localmente."""
        ajustes_api.borrar_token()
        self.tf_api_token.value = ""
        self._actualizar_estado_token()
        self.page.update()
        self.app.avisar("Token de la API eliminado.", VERDE)

    # ----------------------------------------------- catálogo de cuentas
    def _actualizar_estado_cuentas(self) -> None:
        """Refleja si ya hay un Excel de cuentas cargado."""
        if cuentas_bancarias.hay_excel():
            self.txt_estado_cuentas.value = "Archivo de cuentas cargado ✓"
            self.txt_estado_cuentas.color = VERDE
        else:
            self.txt_estado_cuentas.value = "Sin archivo de cuentas cargado."
            self.txt_estado_cuentas.color = GRIS

    async def _adjuntar_cuentas(self, _e=None) -> None:
        """Deja elegir el Excel de cuentas y lo instala (reemplazando el anterior).
        El procesamiento corre en un hilo con un overlay de espera para no congelar
        la UI ni dar la sensación de que 'no pasa nada'."""
        archivos = await self.app.picker.pick_files(
            dialog_title="Selecciona el Excel de cuentas bancarias",
            allowed_extensions=["xlsx", "xls"],
            allow_multiple=False,
        )
        if not archivos:
            return
        # La instalación (copia + lectura del Excel) corre en un hilo con el overlay
        # visible. La recarga de catálogos es rápida (relectura de Excels pequeños,
        # ~decenas de ms) y se hace después en el hilo de la UI.
        self._mostrar_procesando(True, "Procesando catálogo de cuentas…")
        error, empresas = None, 0
        try:
            empresas = await asyncio.to_thread(
                cuentas_bancarias.instalar_excel, archivos[0].path)
        except cuentas_bancarias.ExcelCuentasInvalido as exc:
            error = f"{exc} Se conservó el archivo anterior."  # rollback ya hecho
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            error = f"No se pudo guardar el archivo: {exc}"
        finally:
            self._mostrar_procesando(False)
        if error:
            self.app.avisar(error, ROJO)
            return
        self._actualizar_estado_cuentas()
        self._safe_update()
        self._recargar_catalogos()
        self.app.avisar(
            f"Excel de cuentas actualizado ({empresas} empresa(s)).", VERDE)

    # ------------------------------------------- cuentas de dispersión
    def _actualizar_estado_cuentas_dispersion(self) -> None:
        """Refleja si ya hay un Excel de cuentas de dispersión cargado."""
        if cuentas_dispersion.hay_excel():
            self.txt_estado_cuentas_disp.value = "Archivo de cuentas de dispersión cargado ✓"
            self.txt_estado_cuentas_disp.color = VERDE
        else:
            self.txt_estado_cuentas_disp.value = "Sin archivo de cuentas de dispersión cargado."
            self.txt_estado_cuentas_disp.color = GRIS

    async def _adjuntar_cuentas_dispersion(self, _e=None) -> None:
        """Deja elegir el Excel de cuentas de dispersión (id Empresa + Cuenta) y lo
        instala (reemplazando el anterior). Procesa en un hilo con overlay de espera."""
        archivos = await self.app.picker.pick_files(
            dialog_title="Selecciona el Excel de cuentas de dispersión",
            allowed_extensions=["xlsx", "xls"],
            allow_multiple=False,
        )
        if not archivos:
            return
        # Instalación en un hilo con overlay; recarga rápida después (hilo de la UI).
        self._mostrar_procesando(True, "Procesando catálogo de cuentas de dispersión…")
        error, empresas = None, 0
        try:
            empresas = await asyncio.to_thread(
                cuentas_dispersion.instalar_excel, archivos[0].path)
        except cuentas_dispersion.ExcelCuentasDispersionInvalido as exc:
            error = f"{exc} Se conservó el archivo anterior."
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            error = f"No se pudo guardar el archivo: {exc}"
        finally:
            self._mostrar_procesando(False)
        if error:
            self.app.avisar(error, ROJO)
            return
        self._actualizar_estado_cuentas_dispersion()
        self._safe_update()
        self._recargar_catalogos()
        self.app.avisar(
            f"Cuentas de dispersión actualizadas ({empresas} empresa(s)).", VERDE)

    def _recargar_catalogos(self) -> None:
        """Refresca en caliente las pantallas que consultan el catálogo de
        cuentas (para que el Excel nuevo se refleje sin reabrir la app)."""
        for pantalla in ("devoluciones", "dispersion_no_pemex"):
            try:
                getattr(self.app, pantalla).recargar_catalogo()
            except Exception:  # noqa: BLE001 — no debe romper el guardado
                pass

    # --------------------------------------------------- credenciales
    def _cargar_credenciales(self) -> None:
        """Precarga las credenciales guardadas, si las hay."""
        datos = credenciales.cargar()
        if datos is None:
            return
        usuario, contrasena = datos
        self.tf_usuario.value = usuario
        self.tf_contrasena.value = contrasena

    def credenciales(self) -> tuple[str, str]:
        """Devuelve (usuario, contraseña) tal como están capturados ahora."""
        return (self.tf_usuario.value or "").strip(), self.tf_contrasena.value or ""
