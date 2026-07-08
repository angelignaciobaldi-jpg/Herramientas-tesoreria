"""Diálogo de impresión reutilizable: mensaje de confirmación + selector de
impresora, con la impresora elegida recordada en las preferencias.

Está pensado para cualquier acción que deba mandar algo a imprimir (por ahora la
hoja de calibración de cheques, más adelante los cheques mismos). La lógica de
impresión concreta se inyecta como callback `al_imprimir(nombre_impresora)`, así
que este diálogo no sabe *qué* se imprime, solo *dónde*.

Uso:
    DialogoImpresion(
        app,
        titulo="Imprimir hoja de calibración",
        mensaje="Se imprimirá ...\\n¿Realizar la impresión?",
        clave_pref="impresora_cheques",
        al_imprimir=impresion.imprimir_hoja_calibracion,
    ).abrir()
"""

from __future__ import annotations

import flet as ft

from core import impresion, preferencias
from ui.comun import ROJO, VERDE

_ANCHO = 400


class DialogoImpresion:
    """Confirmación + selección de impresora antes de imprimir."""

    def __init__(
        self,
        app,
        *,
        mensaje: str,
        al_imprimir,
        clave_pref: str = "impresora",
        titulo: str = "Imprimir",
        texto_boton: str = "Imprimir",
    ):
        self.app = app
        self.page = app.page
        self.mensaje = mensaje
        self.al_imprimir = al_imprimir
        self.clave_pref = clave_pref
        self.titulo = titulo
        self.texto_boton = texto_boton
        self.dialogo: ft.AlertDialog | None = None
        self.dd_impresora: ft.Dropdown | None = None

    # ------------------------------------------------------------ acciones
    def abrir(self, _e=None) -> None:
        impresoras = impresion.listar_impresoras()
        if not impresoras:
            self.app.avisar("No se encontraron impresoras instaladas.", ROJO)
            return

        # Preselección: la recordada si sigue disponible; si no, la del sistema.
        elegida = preferencias.cargar_valor(self.clave_pref)
        if elegida not in impresoras:
            elegida = impresion.impresora_predeterminada()
        if elegida not in impresoras:
            elegida = impresoras[0]

        self.dd_impresora = ft.Dropdown(
            label="Impresora", value=elegida, width=_ANCHO,
            options=[ft.dropdown.Option(n) for n in impresoras],
        )
        self.dialogo = ft.AlertDialog(
            modal=True,
            title=ft.Text(self.titulo, size=20, weight=ft.FontWeight.BOLD),
            content=ft.Column(
                [ft.Text(self.mensaje), self.dd_impresora],
                spacing=18, tight=True, width=_ANCHO,
            ),
            actions=[
                ft.TextButton("Cancelar", on_click=self._cerrar),
                ft.FilledButton(
                    self.texto_boton, icon=ft.Icons.PRINT, on_click=self._imprimir),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self.page.show_dialog(self.dialogo)

    def _cerrar(self, _e=None) -> None:
        self.page.pop_dialog()

    def _imprimir(self, _e=None) -> None:
        nombre = self.dd_impresora.value
        # Recordar la elección para la próxima vez (aunque falle la impresión).
        preferencias.guardar_valor(self.clave_pref, nombre)
        self._cerrar()
        try:
            self.al_imprimir(nombre)
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.app.avisar(f"No se pudo imprimir: {exc}", ROJO)
            return
        self.app.avisar(f"Enviado a imprimir en «{nombre}».", VERDE)
