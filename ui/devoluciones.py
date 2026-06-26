"""Pantalla: Generar dispersión de devoluciones.

Captura movimientos (CLABE, monto, beneficiario, concepto, día) y genera el
archivo TXT del banco elegido (Banregio o Bancomer) o un reporte Excel.
"""

from __future__ import annotations

from datetime import date

import flet as ft

from core import cuentas_bancarias, exportador_devoluciones, reporte_excel
from ui.comun import (
    ROJO, VERDE, W_ACCIONES, W_CLABE, W_MONTO, W_NOMBRE,
    celda_centrada, encabezado_col, parse_monto, solo_digitos, tarjeta,
)


class FilaDevolucion:
    """Una fila editable de la tabla de devoluciones (un movimiento)."""

    def __init__(self, seccion: "SeccionDevoluciones"):
        self.seccion = seccion
        # Sin max_length (evita el contador "X/18"); el límite de 18 dígitos se
        # mantiene por código en _limitar_clabe.
        self.tf_clabe = ft.TextField(
            dense=True, width=W_CLABE, text_size=12,
            content_padding=8, text_align=ft.TextAlign.CENTER,
            on_change=self._limitar_clabe,
        )
        self.tf_monto = ft.TextField(
            dense=True, width=W_MONTO, text_size=12, content_padding=8,
            text_align=ft.TextAlign.RIGHT, hint_text="0.00",
        )
        self.tf_benef = ft.TextField(dense=True, width=W_NOMBRE, text_size=12, content_padding=8)
        self.tf_concepto = ft.TextField(dense=True, width=W_NOMBRE, text_size=12, content_padding=8)
        self.fila = ft.DataRow(
            cells=[
                ft.DataCell(self.tf_clabe),
                ft.DataCell(self.tf_monto),
                ft.DataCell(self.tf_benef),
                ft.DataCell(self.tf_concepto),
                ft.DataCell(celda_centrada(
                    ft.IconButton(
                        icon=ft.Icons.DELETE_OUTLINE, tooltip="Quitar", icon_color=ROJO,
                        on_click=lambda e: self.seccion.eliminar_fila(self),
                    ),
                    W_ACCIONES,
                )),
            ]
        )

    def _limitar_clabe(self, _e=None) -> None:
        """Limita la CLABE a 18 dígitos sin usar max_length (para no mostrar el
        contador que desalineaba la fila) y muestra una leyenda si no cumple la
        regla de 18 dígitos exactos."""
        limpio = solo_digitos(self.tf_clabe.value)[:18]
        if limpio != (self.tf_clabe.value or ""):
            self.tf_clabe.value = limpio
        # Sin error si está vacía (fila aún sin capturar) o si ya tiene 18.
        self.tf_clabe.error = (
            None if len(limpio) in (0, 18)
            else ft.Text("Debe tener 18 dígitos exactos.", color=ROJO, size=11)
        )
        self.seccion.page.update()

    def valores(self) -> tuple[str, str, str, str]:
        return (
            solo_digitos(self.tf_clabe.value),
            (self.tf_monto.value or "").strip(),
            (self.tf_benef.value or "").strip(),
            (self.tf_concepto.value or "").strip(),
        )


class SeccionDevoluciones:
    """Pestaña para capturar movimientos y generar el TXT de devoluciones,
    eligiendo el banco (Banregio o Bancomer)."""

    def __init__(self, app):
        self.app = app
        self.page = app.page
        self.filas: list[FilaDevolucion] = []
        self.catalogo = cuentas_bancarias.CatalogoCuentas()
        self.contenido = self._construir()
        self._agregar_fila()  # arranca con 2 movimientos (mínimo requerido)
        self._agregar_fila()
        # Sin page.update() aquí: la página aún no se ha construido.
        self.tabla.rows = [f.fila for f in self.filas]

    # ------------------------------------------------------------ UI
    def _construir(self) -> ft.Control:
        self.dd_banco = ft.Dropdown(
            label="Banco que dispersa", width=200, value="Banregio",
            options=[
                ft.dropdown.Option(key="Banregio", text="Banregio"),
                ft.dropdown.Option(key="Bancomer", text="Bancomer"),
            ],
            on_select=self._cambio_banco,
        )
        # Empresa que dispersa (de dónde sale el pago).
        self.dd_empresa = ft.Dropdown(
            label="Empresa que dispersa", width=420, enable_filter=True, editable=True,
            options=[ft.dropdown.Option(key=e, text=e) for e in self.catalogo.empresas()],
            on_select=self._actualizar_cuentas,
        )
        # Cuenta origen: se llena sola según empresa + banco (no se escribe).
        self._mapa_num_cuenta: dict[str, str] = {}  # clabe -> número de cuenta
        self.dd_origen = ft.Dropdown(
            label="Cuenta origen (CLABE)", width=300, options=[],
            on_select=self._mostrar_num_cuenta,
        )
        # Número de cuenta (informativo, no editable).
        self.tf_num_cuenta = ft.TextField(
            label="Número de cuenta", width=220, read_only=True,
        )
        # Config Banregio (solo fecha) / Bancomer (solo folio). Van en la misma
        # fila que la cuenta origen; se muestra uno u otro según el banco.
        # Sin max_length (evita el contador "8/8" que desalineaba la fila); el
        # límite de 8 dígitos se mantiene por código en _limitar_fecha.
        self.tf_fecha = ft.TextField(
            label="Fecha (DDMMAAAA)", width=170,
            value=date.today().strftime("%d%m%Y"),
            on_change=self._limitar_fecha,
        )
        self.tf_folio = ft.TextField(label="Folio", width=150, value="0023626H", visible=False)

        nota = ""
        if not self.catalogo.disponible():
            nota = ("⚠ No se pudo leer el Excel de cuentas bancarias. Si lo tienes "
                    "abierto en Excel, ciérralo y reabre la aplicación (después se "
                    "usará la última versión leída aunque esté abierto).")
        config = tarjeta(
            "1. Banco y datos del archivo",
            ft.Column(
                [
                    ft.Row([self.dd_banco, self.dd_empresa],
                           wrap=True, vertical_alignment=ft.CrossAxisAlignment.CENTER, spacing=16),
                    ft.Row([self.dd_origen, self.tf_num_cuenta, self.tf_fecha, self.tf_folio],
                           wrap=True, vertical_alignment=ft.CrossAxisAlignment.CENTER, spacing=16),
                    ft.Text(nota, color=ROJO, size=12, visible=bool(nota)),
                ],
                spacing=12,
            ),
        )

        self.tabla = ft.DataTable(
            columns=[
                ft.DataColumn(label=encabezado_col("CLABE", W_CLABE)),
                ft.DataColumn(label=encabezado_col("Monto", W_MONTO), numeric=True),
                ft.DataColumn(label=encabezado_col("Beneficiario", W_NOMBRE)),
                ft.DataColumn(label=encabezado_col("Concepto / Referencia", W_NOMBRE)),
                ft.DataColumn(label=encabezado_col("", W_ACCIONES)),
            ],
            rows=[],
            column_spacing=14,
            heading_row_color=ft.Colors.SURFACE_CONTAINER_HIGHEST,
            heading_row_height=46,
            data_row_min_height=48,
            # Permite que una fila crezca para mostrar la leyenda de la CLABE
            # cuando no tiene 18 dígitos (las filas válidas se quedan en 48).
            data_row_max_height=78,
            divider_thickness=1,
            border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT),
            border_radius=10,
        )
        tabla = tarjeta(
            "2. Movimientos",
            ft.Column(
                [
                    ft.Row([
                        ft.OutlinedButton(content="Agregar movimiento", icon=ft.Icons.ADD,
                                          on_click=lambda e: self._agregar_y_redibujar()),
                        ft.FilledButton(content="Generar TXT", icon=ft.Icons.DESCRIPTION_OUTLINED,
                                        on_click=self._generar),
                        ft.OutlinedButton(content="Generar Excel", icon=ft.Icons.TABLE_VIEW,
                                          on_click=self._generar_excel),
                    ], spacing=10, wrap=True),
                    ft.Row([self.tabla], scroll=ft.ScrollMode.AUTO),
                ],
                spacing=12,
            ),
        )
        return ft.Column([config, tabla], spacing=14, scroll=ft.ScrollMode.AUTO, expand=True)

    def _limitar_fecha(self, _e=None) -> None:
        """Mantiene la fecha en máximo 8 dígitos (DDMMAAAA) sin usar max_length
        (para no mostrar el contador que desalineaba la fila) y muestra una
        leyenda si no cumple la regla de 8 dígitos."""
        limpio = solo_digitos(self.tf_fecha.value)[:8]
        if limpio != (self.tf_fecha.value or ""):
            self.tf_fecha.value = limpio
        self.tf_fecha.error = (
            None if len(limpio) == 8
            else ft.Text("La fecha debe tener 8 dígitos (DDMMAAAA).", color=ROJO, size=11)
        )
        self.page.update()

    def _cambio_banco(self, _e) -> None:
        # La fecha (de devolución) se usa siempre: en el TXT de Banregio y como
        # Fecha de devolución del reporte Excel. El folio es solo de Bancomer.
        self.tf_folio.visible = self.dd_banco.value == "Bancomer"
        self._actualizar_cuentas()

    def _actualizar_cuentas(self, _e=None) -> None:
        """Llena la cuenta origen (CLABE) y el número de cuenta según la empresa
        y el banco seleccionados."""
        empresa = self.dd_empresa.value
        cuentas = self.catalogo.cuentas(empresa, self.dd_banco.value) if empresa else []
        self.dd_origen.options = [
            ft.dropdown.Option(key=clabe, text=f"{clabe}  ({divisa})")
            for clabe, divisa, _num in cuentas
        ]
        self._mapa_num_cuenta = {clabe: num for clabe, _divisa, num in cuentas}
        # Preselecciona la primera (las PESOS/MXP vienen primero).
        self.dd_origen.value = cuentas[0][0] if cuentas else None
        self._mostrar_num_cuenta()
        self.page.update()

    def _mostrar_num_cuenta(self, _e=None) -> None:
        self.tf_num_cuenta.value = self._mapa_num_cuenta.get(self.dd_origen.value, "")
        self.page.update()

    def _redibujar(self) -> None:
        self.tabla.rows = [f.fila for f in self.filas]
        self.page.update()

    def _agregar_fila(self) -> None:
        self.filas.append(FilaDevolucion(self))

    def _agregar_y_redibujar(self) -> None:
        self._agregar_fila()
        self._redibujar()

    def eliminar_fila(self, fila: FilaDevolucion) -> None:
        self.filas.remove(fila)
        self._redibujar()

    # ------------------------------------------------------- generación
    def _recolectar(self):
        """Valida y devuelve los movimientos como [(clabe, monto, benef,
        concepto), ...]; o None (tras avisar) si hay un dato inválido."""
        registros = []
        for fila in self.filas:
            clabe, monto_txt, benef, concepto = fila.valores()
            if not (clabe or monto_txt or benef or concepto):
                continue  # ignora filas totalmente vacías
            if len(clabe) != 18:
                self.app.avisar("Hay CLABE(s) que no tienen 18 dígitos.", ROJO)
                return None
            try:
                monto = parse_monto(monto_txt)
            except ValueError:
                self.app.avisar("Hay un monto inválido (usa solo números).", ROJO)
                return None
            if monto is None:
                self.app.avisar("Falta capturar algún monto.", ROJO)
                return None
            registros.append((clabe, monto, benef, concepto))
        # TODO: volver a exigir mínimo 2 movimientos cuando se active el candado.
        if len(registros) < 1:
            self.app.avisar("Captura al menos 1 movimiento válido.", ROJO)
            return None
        return registros

    def _contexto(self) -> dict:
        # La fecha (de devolución) alimenta tanto el TXT de Banregio como el
        # reporte Excel; se toma siempre del campo de fecha.
        return {
            "empresa": self.dd_empresa.value or "",
            "banco": self.dd_banco.value or "",
            "cuenta_origen": self.dd_origen.value or "",
            "num_cuenta": self.tf_num_cuenta.value or "",
            "fecha": self.tf_fecha.value or "",
        }

    async def _generar(self, _e) -> None:
        registros = self._recolectar()
        if registros is None:
            return
        if not self.dd_empresa.value:
            self.app.avisar("Elige la empresa que dispersa.", ROJO)
            return
        origen = solo_digitos(self.dd_origen.value)
        if len(origen) != 18:
            self.app.avisar(
                "No hay cuenta origen para esa empresa y banco. Elige otra empresa "
                "o revisa el Excel de cuentas.", ROJO)
            return

        movimientos = registros  # (clabe, monto, beneficiario, concepto)
        banco = self.dd_banco.value
        if banco == "Banregio":
            fecha = solo_digitos(self.tf_fecha.value)
            if len(fecha) != 8:
                self.app.avisar("La fecha debe tener 8 dígitos (DDMMAAAA).", ROJO)
                return
            contenido = exportador_devoluciones.generar_banregio(movimientos, fecha)
            nombre_def = "devolucion_banregio.txt"
        else:
            folio = (self.tf_folio.value or "").strip()
            contenido = exportador_devoluciones.generar_bancomer(movimientos, origen, folio)
            nombre_def = "devolucion_bancomer.txt"

        ruta = await self.app.picker.save_file(
            dialog_title=f"Guardar TXT de devoluciones ({banco})",
            file_name=nombre_def, allowed_extensions=["txt"],
        )
        if not ruta:
            return
        if not ruta.lower().endswith(".txt"):
            ruta += ".txt"
        try:
            with open(ruta, "w", encoding="latin-1", newline="") as fh:
                fh.write(contenido)
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.app.avisar(f"No se pudo guardar el archivo: {exc}", ROJO)
            return
        self.app.avisar(f"TXT de {banco} generado con {len(movimientos)} movimiento(s).", VERDE)

    async def _generar_excel(self, _e) -> None:
        registros = self._recolectar()
        if registros is None:
            return
        ruta = await self.app.picker.save_file(
            dialog_title="Guardar reporte Excel de devoluciones",
            file_name="reporte_devoluciones.xlsx", allowed_extensions=["xlsx"],
        )
        if not ruta:
            return
        if not ruta.lower().endswith(".xlsx"):
            ruta += ".xlsx"
        try:
            reporte_excel.generar(ruta, self._contexto(), registros)
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.app.avisar(f"No se pudo generar el Excel: {exc}", ROJO)
            return
        self.app.avisar(f"Reporte Excel generado con {len(registros)} movimiento(s).", VERDE)
