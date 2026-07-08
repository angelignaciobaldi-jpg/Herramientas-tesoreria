"""Configuración: credenciales de inicio de sesión del SIPP (para el RPA).

Se abre como diálogo desde el botón de la barra superior. Captura usuario y
contraseña, que siempre se guardan localmente con la contraseña cifrada (ver
core/credenciales.py). Otras pantallas (p. ej. el RPA de dispersión) leen estas
credenciales con el método credenciales().
"""

from __future__ import annotations

import flet as ft

from core import credenciales, cuentas_bancarias
from ui.comun import GRIS, ROJO, VERDE

# Ancho útil del modal: los inputs lo abarcan de margen a margen.
_ANCHO = 400


class SeccionConfiguracion:
    """Diálogo de configuración con las credenciales del SIPP."""

    def __init__(self, app):
        self.app = app
        self.page = app.page
        self._construir()
        self._cargar_credenciales()

    # ------------------------------------------------------------ UI
    def _construir(self) -> None:
        self.tf_usuario = ft.TextField(
            label="Usuario", width=_ANCHO, height=40, dense=True, content_padding=10,
        )
        self.tf_contrasena = ft.TextField(
            label="Contraseña", width=_ANCHO, password=True, height=40,
            can_reveal_password=False, dense=True, content_padding=10,
        )
        # Apartado "Credenciales SIPP" dentro de la configuración.
        credenciales_apartado = ft.Column(
            [
                ft.Text("Credenciales SIPP", size=15, weight=ft.FontWeight.BOLD),
                self.tf_usuario,
                self.tf_contrasena,
            ],
            spacing=12, tight=True,
        )

        # Apartado "Catálogo de cuentas": adjuntar el Excel de cuentas bancarias.
        self.txt_estado_cuentas = ft.Text(size=12)
        cuentas_apartado = ft.Column(
            [
                ft.Text("Catálogo de cuentas", size=15, weight=ft.FontWeight.BOLD),
                ft.Text(
                    "Adjunta el Excel de cuentas bancarias; se guardará donde la "
                    "app lo consulta y reemplazará al anterior si ya había uno.",
                    size=12, color=GRIS,
                ),
                ft.OutlinedButton(
                    "Adjuntar Excel de cuentas", icon=ft.Icons.UPLOAD_FILE,
                    on_click=self._adjuntar_cuentas, width=_ANCHO,
                ),
                self.txt_estado_cuentas,
            ],
            spacing=10, tight=True,
        )
        self._actualizar_estado_cuentas()

        # Apartado "Impresión de cheques": hoja de calibración para ubicar el
        # cheque sobre la hoja portadora (ver core/impresion.py).
        impresion_apartado = ft.Column(
            [
                ft.Text("Impresión de cheques", size=15, weight=ft.FontWeight.BOLD),
                ft.Text(
                    "Imprime una hoja con cuadrícula milimétrica y las esquinas "
                    "identificadas (SI/SD/II/ID) para calibrar la posición del "
                    "cheque. Sale a escala real (1:1).",
                    size=12, color=GRIS,
                ),
                ft.OutlinedButton(
                    "Imprimir hoja de calibración", icon=ft.Icons.PRINT,
                    on_click=self._imprimir_calibracion, width=_ANCHO,
                ),
            ],
            spacing=10, tight=True,
        )

        self.dialogo = ft.AlertDialog(
            modal=True,
            # Encabezado: título grande en negritas + botón "X" para cerrar.
            title=ft.Row(
                [
                    ft.Text("Configuración", size=25, weight=ft.FontWeight.BOLD),
                    ft.IconButton(
                        icon=ft.Icons.CLOSE, tooltip="Cerrar", on_click=self._cerrar,
                    ),
                ],
                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
                width=_ANCHO,
            ),
            content=ft.Column(
                [
                    credenciales_apartado, ft.Divider(),
                    cuentas_apartado, ft.Divider(),
                    impresion_apartado,
                ],
                spacing=18, tight=True, width=_ANCHO, scroll=ft.ScrollMode.AUTO,
            ),
            actions=[
                ft.FilledButton("Aceptar", icon=ft.Icons.CHECK, on_click=self._guardar),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )

    # -------------------------------------------------------- acciones
    def abrir(self, _e=None) -> None:
        self.page.show_dialog(self.dialogo)

    def _cerrar(self, _e=None) -> None:
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
        """Guarda las credenciales (la contraseña, cifrada)."""
        usuario, contrasena = self.credenciales()
        credenciales.guardar(usuario, contrasena)
        self._cerrar()
        self.app.avisar("Configuración guardada.", VERDE)

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
        """Deja elegir el Excel de cuentas y lo copia a la ubicación esperada por
        la app (reemplazando el anterior)."""
        archivos = await self.app.picker.pick_files(
            dialog_title="Selecciona el Excel de cuentas bancarias",
            allowed_extensions=["xlsx", "xls"],
            allow_multiple=False,
        )
        if not archivos:
            return
        try:
            empresas = cuentas_bancarias.instalar_excel(archivos[0].path)
        except cuentas_bancarias.ExcelCuentasInvalido as exc:
            # Formato inesperado: ya se restauró el Excel anterior (rollback).
            self.app.avisar(f"{exc} Se conservó el archivo anterior.", ROJO)
            return
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.app.avisar(f"No se pudo guardar el archivo: {exc}", ROJO)
            return
        self._actualizar_estado_cuentas()
        self.txt_estado_cuentas.update()
        self._recargar_catalogos()
        self.app.avisar(
            f"Excel de cuentas actualizado ({empresas} empresa(s)).", VERDE)

    def _recargar_catalogos(self) -> None:
        """Refresca en caliente las pantallas que consultan el catálogo de
        cuentas (para que el Excel nuevo se refleje sin reabrir la app)."""
        try:
            self.app.devoluciones.recargar_catalogo()
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
