"""Pantalla: Alta de beneficiarios.

Carga de estados de cuenta (OCR), tabla editable de beneficiarios, guardado en
la base y exportación (TXT de dispersión Bancomer o Excel de alta Banregio).
"""

from __future__ import annotations

import asyncio
import os
from datetime import date

import flet as ft

from core import (
    db, exportador, exportador_alta_bancomer, exportador_alta_banregio, ocr,
    reporte_cuentas,
)
from core.catalogo_bancos import banco_desde_clabe
from core.extractores import extraer_clabes, extraer_datos, nombre_desde_archivo, validar_clabe
from ui.comun import (
    EXTENSIONES, GRIS, NARANJA, ROJO, VERDE,
    W_ACCIONES, W_BANCO, W_CLABE, W_ESTADO, W_MONTO, W_NOMBRE,
    celda_centrada, encabezado_col, fmt_monto, parse_monto, solo_digitos,
    tarjeta, validar,
)

# Fondos de fila según la conciliación con el reporte de cuentas.
SIN_COINCIDENCIA = ft.Colors.with_opacity(0.16, ft.Colors.RED)      # no coincide
SUGERIDO_NOMBRE = ft.Colors.with_opacity(0.18, ft.Colors.AMBER)     # CLABE sugerida por nombre


class FilaBeneficiario:
    """Una fila editable de la tabla (un beneficiario, guardado o pendiente)."""

    def __init__(self, seccion: "SeccionAltaBeneficiarios", id_: int | None, clabe: str,
                 beneficiario: str, alias: str, email: str, origen: str = "",
                 monto: float | None = None, ruta_archivo: str | None = None):
        self.seccion = seccion
        self.id = id_          # None mientras no se haya guardado en la base
        self.origen = origen   # nombre del archivo de origen (informativo)
        self.ruta_archivo = ruta_archivo  # ruta completa para previsualizar
        # Estado de conciliación con el reporte de cuentas:
        #   None     -> no aplica (sin reporte importado)
        #   "ok"     -> la CLABE coincide con el reporte
        #   "nombre" -> CLABE sugerida porque el nombre coincide (fila ámbar)
        #   "sin"    -> no coincide ni por CLABE ni por nombre (fila roja)
        self.conciliacion: str | None = None

        self.tf_clabe = ft.TextField(
            value=clabe, dense=True, width=W_CLABE, max_length=18, text_size=12,
            content_padding=8, text_align=ft.TextAlign.CENTER, on_change=self._cambio_clabe,
        )
        self.tf_monto = ft.TextField(
            value=fmt_monto(monto), dense=True, width=W_MONTO, text_size=12,
            content_padding=8, text_align=ft.TextAlign.RIGHT, hint_text="0.00",
            on_change=self._cambio,
        )
        self.txt_banco = ft.Text(
            banco_desde_clabe(solo_digitos(clabe)) or "—", size=12,
            text_align=ft.TextAlign.CENTER,
        )
        self.tf_benef = ft.TextField(
            value=beneficiario, dense=True, width=W_NOMBRE, text_size=12,
            content_padding=8, on_change=self._cambio,
        )
        self.tf_alias = ft.TextField(
            value=alias, dense=True, width=W_NOMBRE, text_size=12,
            content_padding=8, on_change=self._cambio,
        )
        self.tf_email = ft.TextField(
            value=email, dense=True, width=W_NOMBRE, text_size=12,
            content_padding=8, on_change=self._cambio,
        )
        self.ico_estado = ft.Icon(ft.Icons.CIRCLE, size=16, color=GRIS)

        self.snapshot = self._valores()
        acciones = ft.Row(
            [
                ft.IconButton(
                    icon=ft.Icons.VISIBILITY_OUTLINED, tooltip="Ver archivo",
                    icon_color=ft.Colors.BLUE_700, on_click=lambda e: self.previsualizar(),
                ),
                ft.IconButton(
                    icon=ft.Icons.SAVE_OUTLINED, tooltip="Guardar", icon_color=VERDE,
                    on_click=lambda e: self.guardar(),
                ),
                ft.IconButton(
                    icon=ft.Icons.DELETE_OUTLINE, tooltip="Eliminar", icon_color=ROJO,
                    on_click=lambda e: self.seccion.eliminar_fila(self),
                ),
            ],
            spacing=0,
            alignment=ft.MainAxisAlignment.CENTER,
        )
        self.fila = ft.DataRow(
            selected=False,
            on_select_change=self._al_seleccionar,
            cells=[
                ft.DataCell(celda_centrada(self.ico_estado, W_ESTADO)),
                ft.DataCell(self.tf_clabe),
                ft.DataCell(self.tf_monto),
                ft.DataCell(celda_centrada(self.txt_banco, W_BANCO)),
                ft.DataCell(self.tf_benef),
                ft.DataCell(self.tf_alias),
                ft.DataCell(self.tf_email),
                ft.DataCell(celda_centrada(acciones, W_ACCIONES)),
            ],
        )
        self._actualizar_estado()

    def _al_seleccionar(self, e) -> None:
        self.fila.selected = str(e.data).lower() == "true"
        self.seccion.page.update()

    # ------------------------------------------------------------- helpers
    def _valores(self) -> tuple[str, str, str, str, str]:
        return (
            solo_digitos(self.tf_clabe.value),
            (self.tf_benef.value or "").strip(),
            (self.tf_alias.value or "").strip(),
            (self.tf_email.value or "").strip(),
            (self.tf_monto.value or "").strip(),
        )

    def _cambio_clabe(self, _e) -> None:
        self.txt_banco.value = banco_desde_clabe(solo_digitos(self.tf_clabe.value)) or "—"
        # Si hay un reporte importado, re-concilia esta fila con la nueva CLABE.
        self.seccion._conciliar_una(self)
        self._actualizar_estado()
        self.seccion.page.update()

    def _cambio(self, _e) -> None:
        self._actualizar_estado()
        self.seccion.page.update()

    def _actualizar_estado(self) -> None:
        clabe = solo_digitos(self.tf_clabe.value)
        if not validar_clabe(clabe):
            self.ico_estado.icon = ft.Icons.ERROR
            self.ico_estado.color = ROJO
            self.ico_estado.tooltip = "CLABE inválida"
        elif self.id is None:
            self.ico_estado.icon = ft.Icons.RADIO_BUTTON_UNCHECKED
            self.ico_estado.color = NARANJA
            self.ico_estado.tooltip = "Pendiente de guardar"
        elif self._valores() != self.snapshot:
            self.ico_estado.icon = ft.Icons.EDIT
            self.ico_estado.color = NARANJA
            self.ico_estado.tooltip = "Cambios sin guardar"
        else:
            self.ico_estado.icon = ft.Icons.CHECK_CIRCLE
            self.ico_estado.color = VERDE
            self.ico_estado.tooltip = "Guardado"

    @property
    def pendiente(self) -> bool:
        """True si la fila tiene cambios o altas sin persistir."""
        return self.id is None or self._valores() != self.snapshot

    # ------------------------------------------------------------ acciones
    def guardar(self, silencioso: bool = False) -> bool:
        clabe, beneficiario, alias, email, monto_txt = self._valores()
        error = validar(clabe, beneficiario, alias, email)
        if error:
            if not silencioso:
                self.seccion.avisar(error, ROJO)
            return False
        try:
            monto = parse_monto(monto_txt)
        except ValueError:
            if not silencioso:
                self.seccion.avisar("El monto debe ser un número válido (≥ 0).", ROJO)
            return False
        banco = banco_desde_clabe(clabe)
        try:
            if self.id is None:
                self.id = db.guardar(clabe, beneficiario, alias, email, banco, monto, self.ruta_archivo)
            else:
                db.actualizar(self.id, clabe, beneficiario, alias, email, banco, monto, self.ruta_archivo)
        except db.CLABEDuplicada:
            if not silencioso:
                self.seccion.avisar("Esa CLABE ya pertenece a otro beneficiario.", ROJO)
            return False
        self.tf_monto.value = fmt_monto(monto)  # normaliza la presentación
        self.snapshot = self._valores()
        self._actualizar_estado()
        self.seccion._refrescar_candado_export()
        self.seccion.page.update()
        if not silencioso:
            self.seccion.avisar("Beneficiario guardado.", VERDE)
        return True

    # ------------------------------------------------------- conciliación
    def aplicar_reporte(self, rep: reporte_cuentas.CuentaReporte,
                        incluir_clabe: bool = False) -> None:
        """Complementa el registro con los datos del reporte (fuente autorizada):
        toma el nombre del beneficiario y el correo. Con incluir_clabe también
        rellena la CLABE (cuando la coincidencia fue por nombre)."""
        if incluir_clabe and rep.clabe:
            self.tf_clabe.value = rep.clabe
            self.txt_banco.value = banco_desde_clabe(rep.clabe) or "—"
        if rep.beneficiario:
            self.tf_benef.value = rep.beneficiario
            self.tf_alias.value = rep.beneficiario
        if rep.correo:
            self.tf_email.value = rep.correo
        self._actualizar_estado()

    def marcar_conciliacion(self, estado: str | None) -> None:
        """Pinta la fila según la conciliación: rojo si no coincide; ámbar si la
        CLABE se sugirió por nombre (revisar); sin color si coincide por CLABE o
        si no hay reporte importado."""
        self.conciliacion = estado
        self.fila.color = (
            SIN_COINCIDENCIA if estado == "sin"
            else SUGERIDO_NOMBRE if estado == "nombre"
            else None
        )

    def previsualizar(self) -> None:
        """Abre el archivo original del registro en el visor predeterminado del
        sistema, para revisar el documento y corregir lo que haga falta."""
        ruta = self.ruta_archivo
        if not ruta or not os.path.exists(ruta):
            self.seccion.avisar("No se encontró el archivo original de este registro.", ROJO)
            return
        try:
            os.startfile(ruta)  # Windows: abre en el visor predeterminado
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.seccion.avisar(f"No se pudo abrir el archivo: {exc}", ROJO)


class SeccionAltaBeneficiarios:
    """Pantalla de alta de beneficiarios. Recibe el 'shell' de la app (que
    expone page, picker y avisar)."""

    def __init__(self, app):
        self.app = app
        self.page = app.page
        self.picker = app.picker
        self.filas: list[FilaBeneficiario] = []
        # Reporte de cuentas importado (para conciliar). Vacío = no se subió.
        self.catalogo_reporte: dict[str, reporte_cuentas.CuentaReporte] = {}
        self.nombre_reporte = ""
        self.contenido = self._construir()

    def avisar(self, mensaje: str, color: str | None = None) -> None:
        self.app.avisar(mensaje, color)

    # ===================================================== construcción UI
    def _construir(self) -> ft.Control:
        # --- Sección 1: carga de archivos (uno o varios) ---
        self.txt_estado = ft.Text("", color=GRIS, size=12)
        self.anillo = ft.ProgressRing(width=18, height=18, stroke_width=2, visible=False)
        self.btn_cargar = ft.FilledButton(
            content="Seleccionar estados de cuenta",
            icon=ft.Icons.UPLOAD_FILE,
            on_click=self._seleccionar,
        )
        # Importación del reporte de cuentas (Excel) para conciliar.
        self.txt_estado_rep = ft.Text("", color=GRIS, size=12)
        self.btn_reporte = ft.OutlinedButton(
            content="Importar reporte de cuentas (Excel)",
            icon=ft.Icons.TABLE_VIEW,
            on_click=self._importar_reporte,
        )
        self.btn_quitar_reporte = ft.IconButton(
            icon=ft.Icons.CLOSE, tooltip="Quitar reporte (cancela la conciliación)",
            icon_color=GRIS, visible=False, on_click=self._quitar_reporte,
        )
        seccion_carga = tarjeta(
            "1. Cargar estados de cuenta",
            ft.Column(
                [
                    ft.Row(
                        [
                            self.btn_cargar,
                            self.anillo,
                            ft.Text("Puedes seleccionar varios archivos a la vez.", color=GRIS, size=12, italic=True),
                        ],
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        spacing=12,
                    ),
                    self.txt_estado,
                    ft.Divider(height=1),
                    ft.Row(
                        [
                            self.btn_reporte,
                            self.btn_quitar_reporte,
                            ft.Text(
                                "Opcional: súbelo para conciliar las cuentas; se "
                                "completa el nombre/correo y se marcan en rojo las "
                                "que no aparezcan en el reporte.",
                                color=GRIS, size=12, italic=True, expand=True,
                            ),
                        ],
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        spacing=12,
                    ),
                    self.txt_estado_rep,
                ],
                spacing=8,
            ),
        )

        # --- Sección 2: tabla editable con todos los registros ---
        self.enc_benef = encabezado_col("Beneficiario", W_NOMBRE)
        self.enc_alias = encabezado_col("Alias", W_NOMBRE)
        self.enc_email = encabezado_col("Email de notificación", W_NOMBRE)
        self.tabla = ft.DataTable(
            columns=[
                ft.DataColumn(label=encabezado_col("Estado", W_ESTADO)),
                ft.DataColumn(label=encabezado_col("CLABE", W_CLABE)),
                ft.DataColumn(label=encabezado_col("Monto", W_MONTO), numeric=True),
                ft.DataColumn(label=encabezado_col("Banco", W_BANCO)),
                ft.DataColumn(label=self.enc_benef),
                ft.DataColumn(label=self.enc_alias),
                ft.DataColumn(label=self.enc_email),
                ft.DataColumn(label=encabezado_col("Acciones", W_ACCIONES)),
            ],
            rows=[],
            show_checkbox_column=True,
            column_spacing=14,
            heading_row_color=ft.Colors.SURFACE_CONTAINER_HIGHEST,
            heading_row_height=46,
            data_row_min_height=48,
            data_row_max_height=48,
            divider_thickness=1,
            border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT),
            border_radius=10,
            vertical_lines=ft.BorderSide(1, ft.Colors.with_opacity(0.4, ft.Colors.OUTLINE_VARIANT)),
        )
        self.txt_resumen = ft.Text("", color=GRIS, size=12)
        # Formato/banco de exportación: Bancomer -> TXT (dispersión) ;
        # Banregio -> Excel de alta de cuentas.
        self.dd_formato = ft.Dropdown(
            label="Exportar para", width=300, value="Bancomer",
            options=[
                ft.dropdown.Option(key="Bancomer", text="Bancomer · Dispersión (TXT)"),
                ft.dropdown.Option(key="BancomerAlta", text="Bancomer · Alta de cuentas (carpeta)"),
                ft.dropdown.Option(key="Banregio", text="Banregio · Alta (Excel)"),
            ],
            on_select=self._cambio_formato_export,
        )
        self.btn_export = ft.FilledButton(
            content="Exportar TXT (Bancomer)", icon=ft.Icons.DOWNLOAD,
            on_click=self._exportar,
        )
        barra = ft.Row(
            [
                ft.OutlinedButton(
                    content="Seleccionar todos", icon=ft.Icons.CHECKLIST,
                    on_click=self._seleccionar_todos,
                ),
                ft.OutlinedButton(
                    content="Asignar monto", icon=ft.Icons.ATTACH_MONEY,
                    on_click=self._asignar_monto_seleccionados,
                ),
                ft.FilledButton(
                    content="Guardar seleccionados", icon=ft.Icons.SAVE,
                    on_click=self._guardar_seleccionados,
                ),
                ft.OutlinedButton(
                    content="Guardar pendientes", icon=ft.Icons.SAVE_OUTLINED,
                    on_click=self._guardar_pendientes,
                ),
                self.dd_formato,
                self.btn_export,
                ft.OutlinedButton(
                    content="Eliminar seleccionados", icon=ft.Icons.DELETE_SWEEP_OUTLINED,
                    on_click=self._eliminar_seleccionados,
                    style=ft.ButtonStyle(color=ROJO),
                ),
                self.txt_resumen,
            ],
            spacing=10,
            wrap=True,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        seccion_tabla = tarjeta(
            "2. Revisión y edición de beneficiarios",
            ft.Column(
                [
                    barra,
                    self._leyenda(),
                    ft.Row([self.tabla], scroll=ft.ScrollMode.AUTO),
                ],
                spacing=12,
            ),
        )

        return ft.Column(
            [seccion_carga, seccion_tabla],
            spacing=14,
            scroll=ft.ScrollMode.AUTO,
            expand=True,
        )

    def _leyenda(self) -> ft.Control:
        """Leyenda de los íconos de la columna Estado (mismos íconos/colores
        que usa cada fila en _actualizar_estado)."""
        items = [
            (ft.Icons.CHECK_CIRCLE, VERDE, "Guardado"),
            (ft.Icons.RADIO_BUTTON_UNCHECKED, NARANJA, "Pendiente de guardar"),
            (ft.Icons.EDIT, NARANJA, "Cambios sin guardar"),
            (ft.Icons.ERROR, ROJO, "CLABE inválida (requiere atención)"),
        ]
        chips = [
            ft.Row(
                [ft.Icon(ico, color=color, size=16), ft.Text(txt, size=12, color=GRIS)],
                spacing=5,
                tight=True,
            )
            for ico, color, txt in items
        ]
        # Muestras de las filas coloreadas por la conciliación con el reporte.
        def swatch(color, borde, texto):
            return ft.Row(
                [
                    ft.Container(width=16, height=16, bgcolor=color,
                                 border=ft.Border.all(1, borde), border_radius=3),
                    ft.Text(texto, size=12, color=GRIS),
                ],
                spacing=5,
                tight=True,
            )

        chip_ambar = swatch(SUGERIDO_NOMBRE, NARANJA,
                            "Fila ámbar: CLABE sugerida por nombre (verifica)")
        chip_rojo = swatch(SIN_COINCIDENCIA, ROJO,
                          "Fila en rojo: sin coincidencia en el reporte")
        return ft.Row(
            [ft.Text("Leyenda:", size=12, weight=ft.FontWeight.BOLD, color=GRIS),
             *chips, chip_ambar, chip_rojo],
            spacing=18,
            wrap=True,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )

    # ========================================================= utilidades
    def _redibujar_tabla(self) -> None:
        self.tabla.rows = [f.fila for f in self.filas]
        self._ajustar_anchos()
        self._remarcar_conciliacion()
        self._actualizar_resumen()
        self._refrescar_candado_export()
        self.page.update()

    # ======================================================== conciliación
    async def _importar_reporte(self, _e) -> None:
        """Importa el Excel 'Reporte Cuentas Bancarias' y concilia la tabla."""
        archivos = await self.picker.pick_files(
            dialog_title="Selecciona el reporte de cuentas (Excel)",
            allowed_extensions=["xlsx", "xls"],
            allow_multiple=False,
        )
        if not archivos:
            return
        ruta = archivos[0].path
        try:
            catalogo = await asyncio.to_thread(reporte_cuentas.leer, ruta)
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.avisar(f"No se pudo leer el reporte: {exc}", ROJO)
            return
        if not catalogo:
            self.avisar("El reporte no contiene cuentas con CLABE válida.", NARANJA)
            return
        self.catalogo_reporte = catalogo
        self.nombre_reporte = os.path.basename(ruta)
        self.btn_quitar_reporte.visible = True
        self._conciliar()
        self._redibujar_tabla()
        self.avisar(
            f"Reporte importado: {len(catalogo)} cuenta(s). Conciliación aplicada.",
            VERDE,
        )

    def _quitar_reporte(self, _e) -> None:
        """Cancela la conciliación: olvida el reporte y limpia el marcado rojo."""
        self.catalogo_reporte = {}
        self.nombre_reporte = ""
        self.btn_quitar_reporte.visible = False
        self.txt_estado_rep.value = ""
        for f in self.filas:
            f.marcar_conciliacion(None)
        self.page.update()

    def _conciliar(self) -> None:
        """Concilia TODAS las filas con el reporte. Primero por CLABE; si no hay
        CLABE o no coincide, intenta por nombre del beneficiario y, si encuentra
        una coincidencia clara, sugiere su CLABE (fila ámbar para revisar). Las
        que no coinciden de ninguna forma se marcan en rojo. Solo actúa si hay un
        reporte importado (es decir, si el usuario subió ambos archivos)."""
        if not self.catalogo_reporte:
            return
        por_clabe = por_nombre = sin_match = 0
        for f in self.filas:
            clabe = solo_digitos(f.tf_clabe.value)
            rep = self.catalogo_reporte.get(clabe) if len(clabe) == 18 else None
            if rep:
                f.aplicar_reporte(rep)
                f.marcar_conciliacion("ok")
                por_clabe += 1
                continue
            rep_n = reporte_cuentas.buscar_por_nombre(self.catalogo_reporte, f.tf_benef.value)
            if rep_n:
                f.aplicar_reporte(rep_n, incluir_clabe=True)
                f.marcar_conciliacion("nombre")
                por_nombre += 1
            elif len(clabe) == 18:
                f.marcar_conciliacion("sin")
                sin_match += 1
            else:
                f.marcar_conciliacion(None)  # CLABE incompleta: el ícono ya avisa
        self.txt_estado_rep.value = (
            f"Reporte '{self.nombre_reporte}': {por_clabe} por CLABE, "
            f"{por_nombre} por nombre (CLABE sugerida, revisa en ámbar), "
            f"{sin_match} sin coincidencia (en rojo)."
        )

    def _remarcar_conciliacion(self) -> None:
        """Refresca SOLO el color de las filas (sin reescribir datos), para
        mantener el marcado tras redibujar la tabla. Conserva el aviso ámbar de
        las filas cuya CLABE se sugirió por nombre."""
        for f in self.filas:
            if not self.catalogo_reporte:
                f.marcar_conciliacion(None)
                continue
            if f.conciliacion == "nombre":
                f.marcar_conciliacion("nombre")  # preserva el aviso de revisión
                continue
            clabe = solo_digitos(f.tf_clabe.value)
            if len(clabe) != 18:
                f.marcar_conciliacion(None)
            elif clabe in self.catalogo_reporte:
                f.marcar_conciliacion("ok")
            else:
                f.marcar_conciliacion("sin")

    def _conciliar_una(self, fila: FilaBeneficiario) -> None:
        """Concilia una sola fila por CLABE (al editar su CLABE en vivo). No
        sugiere CLABE por nombre aquí para no pisar lo que el usuario escribe."""
        if not self.catalogo_reporte:
            fila.marcar_conciliacion(None)
            return
        clabe = solo_digitos(fila.tf_clabe.value)
        if len(clabe) != 18:
            fila.marcar_conciliacion(None)
            return
        rep = self.catalogo_reporte.get(clabe)
        if rep:
            fila.aplicar_reporte(rep)
            fila.marcar_conciliacion("ok")
        else:
            fila.marcar_conciliacion("sin")

    def _cambio_formato_export(self, _e=None) -> None:
        """Ajusta el botón de exportar según el formato/banco elegido."""
        if self.dd_formato.value == "Banregio":
            self.btn_export.content = "Exportar Excel (Banregio)"
            self.btn_export.icon = ft.Icons.TABLE_VIEW
        elif self.dd_formato.value == "BancomerAlta":
            self.btn_export.content = "Generar carpeta de alta (Bancomer)"
            self.btn_export.icon = ft.Icons.CREATE_NEW_FOLDER
        else:
            self.btn_export.content = "Exportar TXT (Bancomer)"
            self.btn_export.icon = ft.Icons.DOWNLOAD
        self._refrescar_candado_export()

    def _refrescar_candado_export(self) -> None:
        """Bloquea la exportación si falta un monto. El candado SOLO aplica al
        formato Bancomer de dispersión (que usa montos); ni el alta Banregio ni
        el alta de cuentas Bancomer usan montos."""
        if self.dd_formato.value != "Bancomer":
            self.btn_export.disabled = False
            self.btn_export.tooltip = (
                "Genera el archivo Excel de alta para Banregio"
                if self.dd_formato.value == "Banregio"
                else "Genera la carpeta con los TXT de alta de cuentas Bancomer"
            )
            self.page.update()
            return
        sin_monto = sum(1 for b in db.listar() if b.monto is None)
        self.btn_export.disabled = sin_monto > 0
        self.btn_export.tooltip = (
            f"Hay {sin_monto} registro(s) guardado(s) sin monto. Captura el monto "
            "y guarda para poder exportar."
            if sin_monto else "Genera el archivo TXT de dispersión"
        )
        self.page.update()

    def _ajustar_anchos(self) -> None:
        """Reparte el ancho disponible entre los campos de texto largos
        (beneficiario, alias, email) para que crezcan al agrandar la ventana."""
        ancho = self.page.width or 1180
        # Columnas de ancho fijo: estado, CLABE, monto, banco, acciones.
        fijo = W_ESTADO + W_CLABE + W_MONTO + W_BANCO + W_ACCIONES
        overhead = 170  # paddings de la tarjeta y espaciamiento entre columnas
        disponible = ancho - fijo - overhead
        w = max(W_NOMBRE, int(disponible / 3))
        for f in self.filas:
            f.tf_benef.width = w
            f.tf_alias.width = w
            f.tf_email.width = w
        # Los encabezados acompañan el ancho de sus columnas.
        self.enc_benef.width = w
        self.enc_alias.width = w
        self.enc_email.width = w

    def _on_resize(self, _e) -> None:
        self._ajustar_anchos()
        self.page.update()

    def _actualizar_resumen(self) -> None:
        total = len(self.filas)
        pendientes = sum(1 for f in self.filas if f.pendiente)
        self.txt_resumen.value = (
            f"{total} registro(s) · {pendientes} pendiente(s) de guardar"
            if total else "Sin registros todavía."
        )

    # ============================================================ datos DB
    def cargar_desde_db(self) -> None:
        """Carga los beneficiarios guardados. La llama el shell tras agregar la
        página (para que page.update() ya funcione)."""
        for b in db.listar():
            self.filas.append(
                FilaBeneficiario(self, b.id, b.clabe, b.beneficiario, b.alias,
                                 b.email or "", monto=b.monto, ruta_archivo=b.ruta_archivo)
            )
        self._redibujar_tabla()

    # ======================================================= carga archivos
    async def _seleccionar(self, _e) -> None:
        archivos = await self.picker.pick_files(
            dialog_title="Selecciona uno o varios estados de cuenta",
            allowed_extensions=EXTENSIONES,
            allow_multiple=True,
        )
        if not archivos:
            return

        self.btn_cargar.disabled = True
        self.anillo.visible = True
        identificados = 0
        errores: list[str] = []

        for i, archivo in enumerate(archivos, start=1):
            nombre = os.path.basename(archivo.path)
            self.txt_estado.value = f"Procesando {i}/{len(archivos)}: {nombre}…"
            self.page.update()
            try:
                texto, uso_ocr = await asyncio.to_thread(ocr.extraer_texto, archivo.path)
                datos = extraer_datos(texto)
                # Si la capa de texto no dio nada útil (p. ej. impresión de un
                # correo de Outlook con el estado como imagen), forzar OCR.
                if not datos.clabe and not datos.beneficiario and not uso_ocr:
                    texto, _ = await asyncio.to_thread(ocr.extraer_texto, archivo.path, True)
                    datos = extraer_datos(texto)
                # Último recurso para la CLABE: OCR en modo "texto disperso", que
                # recupera números en tablas/celdas que la segmentación normal
                # omite (p. ej. la tabla de una carta de asignación de cuenta).
                if not datos.clabe:
                    texto_sp = await asyncio.to_thread(ocr.texto_disperso, archivo.path)
                    clabes_sp = extraer_clabes(texto_sp)
                    if clabes_sp:
                        datos.clabe = clabes_sp[0]
                        datos.banco = banco_desde_clabe(datos.clabe)
                # El nombre del archivo (si parece nombre de persona) es la
                # fuente más confiable del beneficiario; tiene prioridad sobre
                # el OCR. Si no, se usa lo identificado en el documento.
                beneficiario = nombre_desde_archivo(nombre) or datos.beneficiario
                self.filas.append(
                    FilaBeneficiario(
                        self, None, datos.clabe, beneficiario, beneficiario,
                        datos.email, origen=nombre, ruta_archivo=archivo.path,
                    )
                )
                identificados += 1
                self._redibujar_tabla()
            except Exception as exc:  # noqa: BLE001 — se reporta al usuario
                errores.append(f"{nombre}: {exc}")

        self.btn_cargar.disabled = False
        self.anillo.visible = False
        resumen = f"{identificados} de {len(archivos)} archivo(s) identificado(s) y agregado(s) a la tabla."
        if errores:
            resumen += " Con error: " + "; ".join(errores)
        self.txt_estado.value = resumen
        # Si ya hay un reporte importado, concilia los nuevos registros.
        self._conciliar()
        self._actualizar_resumen()
        self.page.update()

    # =========================================================== acciones
    def _seleccionar_todos(self, _e) -> None:
        """Selecciona todas las filas; si ya estaban todas, las deselecciona."""
        if not self.filas:
            return
        nuevo = not all(f.fila.selected for f in self.filas)
        for f in self.filas:
            f.fila.selected = nuevo
        self.page.update()

    def _asignar_monto_seleccionados(self, _e) -> None:
        """Aplica un mismo monto a todas las filas seleccionadas, para no
        capturarlo uno por uno."""
        seleccionados = [f for f in self.filas if f.fila.selected]
        if not seleccionados:
            self.avisar("No hay registros seleccionados (marca las casillas).", GRIS)
            return

        tf = ft.TextField(
            label="Monto", hint_text="0.00", autofocus=True,
            text_align=ft.TextAlign.RIGHT, prefix_icon=ft.Icons.ATTACH_MONEY,
        )

        def aplicar(_ev):
            try:
                monto = parse_monto(tf.value)
            except ValueError:
                self.avisar("Monto inválido. Usa solo números (ej. 1500.00).", ROJO)
                return
            if monto is None:
                self.avisar("Captura un monto.", ROJO)
                return
            for fila in seleccionados:
                fila.tf_monto.value = fmt_monto(monto)
                fila._actualizar_estado()  # queda como cambio pendiente de guardar
            self.page.pop_dialog()
            self.page.update()
            self.avisar(
                f"Monto asignado a {len(seleccionados)} registro(s). "
                "Usa 'Guardar seleccionados' para aplicarlo.",
                VERDE,
            )

        self.page.show_dialog(
            ft.AlertDialog(
                modal=True,
                title=ft.Text(f"Asignar monto a {len(seleccionados)} seleccionado(s)"),
                content=ft.Container(content=tf, width=300),
                actions=[
                    ft.TextButton(content="Cancelar", on_click=lambda e: self.page.pop_dialog()),
                    ft.FilledButton(content="Aplicar", on_click=aplicar),
                ],
            )
        )

    def _guardar_varias(self, filas, etiqueta_vacia: str) -> None:
        if not filas:
            self.avisar(etiqueta_vacia, GRIS)
            return
        guardados = sum(1 for f in filas if f.guardar(silencioso=True))
        fallidos = len(filas) - guardados
        self._actualizar_resumen()
        self._refrescar_candado_export()
        self.page.update()
        if fallidos:
            self.avisar(
                f"{guardados} guardado(s). {fallidos} con datos inválidos o CLABE "
                "duplicada (revisa las filas en rojo/naranja).",
                NARANJA,
            )
        else:
            self.avisar(f"{guardados} beneficiario(s) guardado(s).", VERDE)

    def _guardar_seleccionados(self, _e) -> None:
        self._guardar_varias(
            [f for f in self.filas if f.fila.selected],
            "No hay registros seleccionados (marca las casillas o usa 'Seleccionar todos').",
        )

    def _guardar_pendientes(self, _e) -> None:
        self._guardar_varias(
            [f for f in self.filas if f.pendiente],
            "No hay registros pendientes de guardar.",
        )

    async def _exportar(self, _e) -> None:
        """Exporta los registros guardados en el formato del banco elegido:
        Bancomer -> TXT de dispersión ; Banregio -> Excel de alta de cuentas."""
        guardados = db.listar()
        if not guardados:
            self.avisar("No hay registros guardados para exportar.", ROJO)
            return

        if self.dd_formato.value == "Banregio":
            await self._exportar_alta_banregio(guardados)
        elif self.dd_formato.value == "BancomerAlta":
            await self._exportar_alta_bancomer(guardados)
        else:
            await self._exportar_dispersion_bancomer(guardados)

    async def _exportar_alta_bancomer(self, guardados) -> None:
        """Genera la carpeta 'CUENTAS PARA ALTA BANCOMER - dd-mm-aaaa' con los
        TXT de alta de cuentas para el portal Bancomer, separando los registros
        que tienen correo de los que no (estos requieren capturar el correo)."""
        validos = [b for b in guardados if validar_clabe(b.clabe)]
        if not validos:
            self.avisar("No hay registros con CLABE válida para exportar.", ROJO)
            return
        con_correo = [
            (b.clabe, b.beneficiario, b.email or "")
            for b in validos if (b.email or "").strip()
        ]
        sin_correo = [
            (b.clabe, b.beneficiario, "")
            for b in validos if not (b.email or "").strip()
        ]

        destino = await self.picker.get_directory_path(
            dialog_title="Elige dónde crear la carpeta de alta de cuentas Bancomer",
        )
        if not destino:
            return
        # El nombre lleva la fecha; en Windows una carpeta no admite '/', así que
        # se usa dd-mm-aaaa en lugar de dd/mm/aaaa.
        fecha = date.today().strftime("%d-%m-%Y")
        carpeta = os.path.join(destino, f"CUENTAS PARA ALTA BANCOMER - {fecha}")
        try:
            os.makedirs(carpeta, exist_ok=True)
            resumen: list[str] = []
            if con_correo:
                with open(os.path.join(carpeta, "Cuentas con correo.txt"),
                          "w", encoding="latin-1", newline="") as fh:
                    fh.write(exportador_alta_bancomer.generar_txt(con_correo))
                resumen.append(f"{len(con_correo)} con correo")
            if sin_correo:
                with open(os.path.join(carpeta, "Cuentas sin correo.txt"),
                          "w", encoding="latin-1", newline="") as fh:
                    fh.write(exportador_alta_bancomer.generar_txt(sin_correo))
                resumen.append(f"{len(sin_correo)} sin correo")
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.avisar(f"No se pudo generar la carpeta: {exc}", ROJO)
            return
        try:
            os.startfile(carpeta)  # abre la carpeta generada en el explorador
        except Exception:  # noqa: BLE001 — abrir es opcional
            pass
        self.avisar("Carpeta de alta generada (" + ", ".join(resumen) + ").", VERDE)

    async def _exportar_dispersion_bancomer(self, guardados) -> None:
        # Candado: no se exporta si algún registro guardado no tiene monto.
        sin_monto = sum(1 for b in guardados if b.monto is None)
        if sin_monto:
            self.avisar(
                f"No se puede exportar: {sin_monto} registro(s) guardado(s) sin monto. "
                "Captura el monto y guárdalo.",
                ROJO,
            )
            return
        registros = [
            (b.clabe, b.monto, b.beneficiario, b.alias)
            for b in guardados
            if validar_clabe(b.clabe)
        ]
        ruta = await self.picker.save_file(
            dialog_title="Guardar archivo de dispersión TXT (Bancomer)",
            file_name="dispersion.txt", allowed_extensions=["txt"],
        )
        if not ruta:
            return
        if not ruta.lower().endswith(".txt"):
            ruta += ".txt"
        try:
            contenido = exportador.generar_txt(registros)
            with open(ruta, "w", encoding="latin-1", newline="") as fh:
                fh.write(contenido)
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.avisar(f"No se pudo guardar el archivo: {exc}", ROJO)
            return
        self.avisar(f"TXT generado con {len(registros)} registro(s) guardado(s).", VERDE)

    async def _exportar_alta_banregio(self, guardados) -> None:
        registros = [
            (b.clabe, b.beneficiario, b.email or "")
            for b in guardados
            if validar_clabe(b.clabe)
        ]
        if not registros:
            self.avisar("No hay registros con CLABE válida para exportar.", ROJO)
            return
        ruta = await self.picker.save_file(
            dialog_title="Guardar archivo de alta (Banregio)",
            file_name="alta_banregio.xls", allowed_extensions=["xls"],
        )
        if not ruta:
            return
        if not ruta.lower().endswith(".xls"):
            ruta += ".xls"
        try:
            exportador_alta_banregio.generar(ruta, registros)
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.avisar(f"No se pudo guardar el archivo: {exc}", ROJO)
            return
        self.avisar(f"Excel de alta (Banregio) generado con {len(registros)} registro(s).", VERDE)

    def _eliminar_seleccionados(self, _e) -> None:
        """Elimina todas las filas seleccionadas (con confirmación)."""
        seleccionados = [f for f in self.filas if f.fila.selected]
        if not seleccionados:
            self.avisar("No hay registros seleccionados (marca las casillas).", GRIS)
            return

        def confirmar(_ev):
            for fila in seleccionados:
                if fila.id is not None:
                    db.eliminar(fila.id)
                self.filas.remove(fila)
            self.page.pop_dialog()
            self._redibujar_tabla()  # refresca resumen y candado de exportación
            self.avisar(f"{len(seleccionados)} registro(s) eliminado(s).", GRIS)

        self.page.show_dialog(
            ft.AlertDialog(
                modal=True,
                title=ft.Text("Confirmar eliminación"),
                content=ft.Text(
                    f"¿Eliminar {len(seleccionados)} registro(s) seleccionado(s)? "
                    "Esta acción no se puede deshacer."
                ),
                actions=[
                    ft.TextButton(content="Cancelar", on_click=lambda e: self.page.pop_dialog()),
                    ft.FilledButton(content="Eliminar", on_click=confirmar,
                                    color=ft.Colors.WHITE, bgcolor=ROJO),
                ],
            )
        )

    def eliminar_fila(self, fila: FilaBeneficiario) -> None:
        def confirmar(_e):
            if fila.id is not None:
                db.eliminar(fila.id)
            self.filas.remove(fila)
            self.page.pop_dialog()
            self._redibujar_tabla()
            self.avisar("Registro eliminado.", GRIS)

        # Si la fila aún no se guarda, se quita directo sin confirmar.
        if fila.id is None:
            self.filas.remove(fila)
            self._redibujar_tabla()
            return

        self.page.show_dialog(
            ft.AlertDialog(
                modal=True,
                title=ft.Text("Confirmar eliminación"),
                content=ft.Text("¿Eliminar este beneficiario de la base? Esta acción no se puede deshacer."),
                actions=[
                    ft.TextButton(content="Cancelar", on_click=lambda e: self.page.pop_dialog()),
                    ft.FilledButton(content="Eliminar", on_click=confirmar, color=ft.Colors.WHITE, bgcolor=ROJO),
                ],
            )
        )
