"""Pantalla: Generar dispersión de devoluciones.

Flujo, en una sola tabla:

  1. Se consultan las solicitudes AUTORIZADAS del SIPP por empresa(s). Sus datos
     llegan llenos pero son EDITABLES (empresa, cliente, CLABE, monto, concepto);
     cada fila puede restaurarse a los datos originales de la solicitud.
     En esa misma tabla se pueden agregar MOVIMIENTOS MANUALES (sin folio, porque
     no existen como solicitud en el sistema).
  2. Se seleccionan varias filas y se les asigna la cuenta ORIGEN de pago
     (empresa que dispersa + banco + CLABE origen).
  3. Se genera UN archivo TXT por cada empresa origen / banco.

El banco destino NO se captura: se deduce de los 3 primeros dígitos de la CLABE.
"""

from __future__ import annotations

import os
import re
from datetime import date

import flet as ft

from core import (
    cuentas_bancarias, exportador_devoluciones, reporte_excel,
    solicitudes_devolucion,
)
from core.catalogo_bancos import banco_desde_clabe
from ui.comun import (
    GRIS, NARANJA, ROJO, VERDE, W_BANCO, W_CLABE, W_MONTO, W_NOMBRE,
    celda_centrada, encabezado_col, fmt_monto, parse_monto, solo_digitos, tarjeta,
)

# Anchos de la tabla unificada de solicitudes/movimientos.
W_FOLIO = 120
W_EMPRESA = 215
W_CONCEPTO = 200
W_ORIGEN = 230
W_ACC = 110

# Fondo de la fila cuya cuenta origen de pago ya fue asignada.
ASIGNADA = ft.Colors.with_opacity(0.14, ft.Colors.GREEN)


def _sanear_archivo(nombre: str) -> str:
    """Deja un texto usable como nombre de archivo en Windows."""
    limpio = re.sub(r'[<>:"/\\|?*\x00-\x1f]', " ", nombre or "").strip()
    return re.sub(r"\s+", " ", limpio) or "SIN NOMBRE"


class FilaSolicitud:
    """Una fila de la tabla: o una solicitud traída del SIPP, o un movimiento
    capturado a mano (`manual=True`, sin folio).

    Campos EDITABLES: empresa solicitante, cliente, CLABE, monto y concepto.
    Campos automáticos: folio (solo SIPP), banco destino (de la CLABE) y la cuenta
    origen de pago (se asigna en el paso 2).
    """

    def __init__(self, seccion: "SeccionDevoluciones",
                 sol: solicitudes_devolucion.SolicitudDevolucion | None = None,
                 manual: bool = False):
        self.seccion = seccion
        self.manual = manual
        self.original = sol          # datos originales del SIPP (None si es manual)
        self.asignacion: dict | None = None  # cuenta origen de pago (paso 2)

        self.txt_folio = ft.Text(
            (sol.folio if sol else "— manual —"), size=12,
            italic=manual, color=(GRIS if manual else None),
        )
        self.dd_empresa = ft.Dropdown(
            dense=True, width=W_EMPRESA, text_size=12, content_padding=8,
            options=[ft.dropdown.Option(key=e, text=e)
                     for e in seccion.empresas_solicitantes()],
            on_select=self._cambio,
        )
        self.tf_cliente = ft.TextField(
            dense=True, width=W_NOMBRE, text_size=12, content_padding=8,
            on_change=self._cambio,
        )
        # Sin max_length (evita el contador "X/18" que desalinea la fila); el
        # límite de 18 dígitos se aplica por código en _cambio_clabe.
        self.tf_clabe = ft.TextField(
            dense=True, width=W_CLABE, text_size=12, content_padding=8,
            text_align=ft.TextAlign.CENTER, on_change=self._cambio_clabe,
        )
        self.txt_banco = ft.Text("—", size=12, text_align=ft.TextAlign.CENTER)
        self.tf_monto = ft.TextField(
            dense=True, width=W_MONTO, text_size=12, content_padding=8,
            text_align=ft.TextAlign.RIGHT, hint_text="0.00", on_change=self._cambio,
        )
        self.tf_concepto = ft.TextField(
            dense=True, width=W_CONCEPTO, text_size=12, content_padding=8,
            on_change=self._cambio,
        )
        self.txt_origen = ft.Text("— sin asignar —", size=12, color=GRIS)

        acciones = ft.Row(
            [
                ft.IconButton(
                    icon=ft.Icons.RESTORE, tooltip=(
                        "Limpiar el movimiento (empezar de cero)" if manual
                        else "Restaurar los datos originales de la solicitud"),
                    icon_color=ft.Colors.BLUE_700, on_click=lambda e: self.restaurar(),
                ),
                ft.IconButton(
                    icon=ft.Icons.DELETE_OUTLINE, tooltip="Quitar de la tabla",
                    icon_color=ROJO, on_click=lambda e: self.seccion.eliminar_fila(self),
                ),
            ],
            spacing=0, alignment=ft.MainAxisAlignment.CENTER, tight=True,
        )
        self.fila = ft.DataRow(
            selected=False,
            on_select_change=self._al_seleccionar,
            cells=[
                ft.DataCell(celda_centrada(self.txt_folio, W_FOLIO)),
                ft.DataCell(self.dd_empresa),
                ft.DataCell(self.tf_cliente),
                ft.DataCell(self.tf_clabe),
                ft.DataCell(celda_centrada(self.txt_banco, W_BANCO)),
                ft.DataCell(self.tf_monto),
                ft.DataCell(self.tf_concepto),
                ft.DataCell(celda_centrada(self.txt_origen, W_ORIGEN)),
                ft.DataCell(celda_centrada(acciones, W_ACC)),
            ],
        )
        self.restaurar(avisar=False)  # carga los datos originales (o deja vacío)

    # ---------------------------------------------------------- edición
    def restaurar(self, avisar: bool = True) -> None:
        """Devuelve los campos editables a los datos ORIGINALES de la solicitud
        (o los deja vacíos si es un movimiento manual). NO toca la cuenta origen
        de pago ya asignada (esa se quita con 'Quitar asignación')."""
        sol = self.original
        self.dd_empresa.value = sol.empresa if sol else None
        self.tf_cliente.value = sol.cliente if sol else ""
        self.tf_clabe.value = sol.clabe if sol else ""
        self.tf_monto.value = fmt_monto(sol.monto) if sol else ""
        self.tf_concepto.value = sol.concepto if sol else ""
        self.tf_clabe.error = None
        # El banco del SIPP es el de partida; si luego se edita la CLABE, se
        # recalcula a partir de ella (ver _cambio_clabe).
        self.txt_banco.value = (sol.banco if sol else "") or "—"
        if avisar:
            self.seccion._refrescar()

    def _cambio(self, _e=None) -> None:
        self.seccion._actualizar_contador()
        self.seccion.page.update()

    def _cambio_clabe(self, _e=None) -> None:
        """Limita la CLABE a 18 dígitos, avisa si no los tiene y deduce el banco
        destino de sus 3 primeros dígitos."""
        limpio = solo_digitos(self.tf_clabe.value)[:18]
        if limpio != (self.tf_clabe.value or ""):
            self.tf_clabe.value = limpio
        # Sin error si está vacía (fila aún sin capturar) o si ya tiene 18.
        self.tf_clabe.error = (
            None if len(limpio) in (0, 18)
            else ft.Text("Debe tener 18 dígitos exactos.", color=ROJO, size=11)
        )
        self.txt_banco.value = banco_desde_clabe(limpio) or "—"
        self._cambio()

    # ------------------------------------------------------- selección
    def _al_seleccionar(self, e) -> None:
        self.fila.selected = str(e.data).lower() == "true"
        self.seccion._actualizar_contador()
        self.seccion.page.update()

    @property
    def seleccionada(self) -> bool:
        return bool(self.fila.selected)

    def valores(self) -> tuple[str, str, str, str, str]:
        """(clabe, monto_texto, cliente, concepto, empresa_solicitante)."""
        return (
            solo_digitos(self.tf_clabe.value),
            (self.tf_monto.value or "").strip(),
            (self.tf_cliente.value or "").strip(),
            (self.tf_concepto.value or "").strip(),
            self.dd_empresa.value or "",
        )

    def vacia(self) -> bool:
        """True si es una fila manual sin nada capturado (se ignora al exportar)."""
        clabe, monto, cliente, concepto, empresa = self.valores()
        return not (clabe or monto or cliente or concepto or empresa)

    # ----------------------------------------------------- asignación
    def asignar(self, datos: dict) -> None:
        self.asignacion = datos
        self.txt_origen.value = (
            f"{datos['empresa_origen']} · {datos['banco']}\n{datos['cuenta_origen']}"
        )
        self.txt_origen.color = None
        self.fila.color = ASIGNADA

    def limpiar_asignacion(self) -> None:
        self.asignacion = None
        self.txt_origen.value = "— sin asignar —"
        self.txt_origen.color = GRIS
        self.fila.color = None


class SeccionDevoluciones:
    """Consulta de solicitudes (SIPP) + captura manual, asignación de cuenta
    origen de pago y generación de un TXT por empresa origen."""

    def __init__(self, app):
        self.app = app
        self.page = app.page
        self.filas: list[FilaSolicitud] = []
        self._empresas_val = {
            e: True for e in solicitudes_devolucion.empresas()  # todas por defecto
        }
        self.catalogo = cuentas_bancarias.CatalogoCuentas()
        self.contenido = self._construir()

    def empresas_solicitantes(self) -> list[str]:
        """Empresas del desplegable 'Empresa solicitante' (las del SIPP)."""
        return solicitudes_devolucion.empresas()

    def recargar_catalogo(self) -> None:
        """Relee el Excel de cuentas y refresca el combo de empresas que dispersan.
        Se llama tras adjuntar un Excel nuevo en Configuración."""
        self.catalogo = cuentas_bancarias.CatalogoCuentas()
        empresas = self.catalogo.empresas()
        self.dd_empresa.options = [ft.dropdown.Option(key=e, text=e) for e in empresas]
        self.txt_sin_catalogo.visible = not self.catalogo.disponible()
        if self.dd_empresa.value not in empresas:
            self.dd_empresa.value = None
            self.dd_origen.options = []
            self.dd_origen.value = None
            self.tf_num_cuenta.value = ""
        try:
            self.page.update()
        except (RuntimeError, AssertionError):
            pass  # aún no montada la página

    # ================================================================ UI
    def _construir_tabla(self) -> ft.Control:
        """Tarjeta 1: solicitudes del SIPP (editables) + movimientos manuales."""
        self.tf_empresas = ft.TextField(
            label="Empresas a consultar", width=280, read_only=True,
            value=self._resumen_empresas(), suffix_icon=ft.Icons.ARROW_DROP_DOWN,
            on_click=lambda e: self._abrir_selector_empresas(),
        )
        self.txt_contador = ft.Text("Sin registros.", size=12, color=GRIS,
                                    weight=ft.FontWeight.BOLD)
        self.tabla = ft.DataTable(
            columns=[
                ft.DataColumn(label=encabezado_col("Folio", W_FOLIO)),
                ft.DataColumn(label=encabezado_col("Empresa solicitante", W_EMPRESA)),
                ft.DataColumn(label=encabezado_col("Cliente / Beneficiario", W_NOMBRE)),
                ft.DataColumn(label=encabezado_col("CLABE Beneficiario", W_CLABE)),
                ft.DataColumn(label=encabezado_col("Banco destino", W_BANCO)),
                ft.DataColumn(label=encabezado_col("Monto", W_MONTO), numeric=True),
                ft.DataColumn(label=encabezado_col("Concepto", W_CONCEPTO)),
                ft.DataColumn(label=encabezado_col("Cuenta origen de pago", W_ORIGEN)),
                ft.DataColumn(label=encabezado_col("Acciones", W_ACC)),
            ],
            rows=[],
            show_checkbox_column=True,   # selección múltiple
            column_spacing=14,
            heading_row_color=ft.Colors.SURFACE_CONTAINER_HIGHEST,
            heading_row_height=46,
            data_row_min_height=54,
            # Deja crecer la fila para mostrar la leyenda de CLABE incompleta.
            data_row_max_height=86,
            divider_thickness=1,
            border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT),
            border_radius=10,
        )
        return tarjeta(
            "1. Solicitudes autorizadas (SIPP) y movimientos manuales",
            ft.Column(
                [
                    ft.Row(
                        [
                            self.tf_empresas,
                            ft.FilledButton(
                                content="Consultar solicitudes", icon=ft.Icons.SEARCH,
                                on_click=self._consultar,
                            ),
                            ft.OutlinedButton(
                                content="Agregar movimiento manual", icon=ft.Icons.ADD,
                                on_click=self._agregar_manual,
                            ),
                            ft.OutlinedButton(
                                content="Seleccionar todas", icon=ft.Icons.CHECKLIST,
                                on_click=self._seleccionar_todas,
                            ),
                            ft.OutlinedButton(
                                content="Quitar asignación", icon=ft.Icons.LAYERS_CLEAR,
                                on_click=self._quitar_asignacion,
                            ),
                            ft.OutlinedButton(
                                content="Generar Excel", icon=ft.Icons.TABLE_VIEW,
                                on_click=self._generar_excel,
                            ),
                            self.txt_contador,
                        ],
                        spacing=10, wrap=True,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    ),
                    ft.Text(
                        "Los datos llegan llenos del SIPP pero puedes editarlos (empresa, "
                        "cliente, CLABE, monto y concepto); el botón ↺ restaura el original. "
                        "Los movimientos manuales no llevan folio y su banco se deduce de la "
                        "CLABE. Datos de prueba mientras se conecta el SIPP.",
                        color=GRIS, size=12, italic=True,
                    ),
                    ft.Row([self.tabla], scroll=ft.ScrollMode.AUTO),
                ],
                spacing=12,
            ),
        )

    def _construir(self) -> ft.Control:
        # --- Paso 2: cuenta origen de pago (empresa que dispersa + banco) ---
        self.dd_banco = ft.Dropdown(
            label="Banco que dispersa", width=200, value="Banregio",
            options=[
                ft.dropdown.Option(key="Banregio", text="Banregio"),
                ft.dropdown.Option(key="Bancomer", text="Bancomer"),
            ],
            on_select=self._cambio_banco,
        )
        self.dd_empresa = ft.Dropdown(
            label="Empresa que dispersa", width=420, enable_filter=True, editable=True,
            options=[ft.dropdown.Option(key=e, text=e) for e in self.catalogo.empresas()],
            on_select=self._actualizar_cuentas,
        )
        self._mapa_num_cuenta: dict[str, str] = {}  # clabe -> número de cuenta
        self.dd_origen = ft.Dropdown(
            label="Cuenta origen (CLABE)", width=300, options=[],
            on_select=self._mostrar_num_cuenta,
        )
        self.tf_num_cuenta = ft.TextField(
            label="Número de cuenta", width=220, read_only=True,
        )
        self.tf_fecha = ft.TextField(
            label="Fecha (DDMMAAAA)", width=170,
            value=date.today().strftime("%d%m%Y"),
            on_change=self._limitar_fecha,
        )
        self.tf_folio = ft.TextField(label="Folio", width=150, value="0023626H", visible=False)
        self.txt_sin_catalogo = ft.Text(
            "⚠ No se pudo leer el Excel de cuentas bancarias. Adjúntalo en "
            "Configuración (⚙) o, si lo tienes abierto en Excel, ciérralo.",
            color=ROJO, size=12, visible=not self.catalogo.disponible(),
        )
        config = tarjeta(
            "2. Cuenta origen de pago (empresa que dispersa y banco)",
            ft.Column(
                [
                    ft.Row([self.dd_banco, self.dd_empresa],
                           wrap=True, vertical_alignment=ft.CrossAxisAlignment.CENTER, spacing=16),
                    ft.Row([self.dd_origen, self.tf_num_cuenta, self.tf_fecha, self.tf_folio],
                           wrap=True, vertical_alignment=ft.CrossAxisAlignment.CENTER, spacing=16),
                    ft.Row([
                        ft.FilledButton(
                            content="Asignar a filas seleccionadas",
                            icon=ft.Icons.ACCOUNT_BALANCE_WALLET,
                            on_click=self._asignar_a_seleccionadas,
                        ),
                    ], wrap=True, vertical_alignment=ft.CrossAxisAlignment.CENTER, spacing=16),
                    self.txt_sin_catalogo,
                ],
                spacing=12,
            ),
        )

        # --- Paso 3: generación (un TXT por empresa origen / banco) ---
        self.txt_grupos = ft.Text("Sin filas asignadas.", size=12, color=GRIS)
        generar = tarjeta(
            "3. Generar dispersión",
            ft.Column(
                [
                    ft.Row(
                        [
                            ft.FilledButton(
                                content="Generar TXT de dispersión",
                                icon=ft.Icons.CREATE_NEW_FOLDER,
                                on_click=self._generar_dispersion,
                            ),
                            self.txt_grupos,
                        ],
                        spacing=12, wrap=True,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    ),
                    ft.Text(
                        "Se genera un archivo TXT por cada empresa origen de pago (y banco). "
                        "No se genera si alguna CLABE no tiene 18 dígitos o falta el monto.",
                        color=GRIS, size=12, italic=True,
                    ),
                ],
                spacing=8,
            ),
        )
        return ft.Column(
            [self._construir_tabla(), config, generar],
            spacing=14, scroll=ft.ScrollMode.AUTO, expand=True,
        )

    # ------------------------------------------------ paso 2: cuenta origen
    def _limitar_fecha(self, _e=None) -> None:
        limpio = solo_digitos(self.tf_fecha.value)[:8]
        if limpio != (self.tf_fecha.value or ""):
            self.tf_fecha.value = limpio
        self.tf_fecha.error = (
            None if len(limpio) == 8
            else ft.Text("La fecha debe tener 8 dígitos (DDMMAAAA).", color=ROJO, size=11)
        )
        self.page.update()

    def _cambio_banco(self, _e) -> None:
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
        self.dd_origen.value = cuentas[0][0] if cuentas else None
        self._mostrar_num_cuenta()
        self.page.update()

    def _mostrar_num_cuenta(self, _e=None) -> None:
        self.tf_num_cuenta.value = self._mapa_num_cuenta.get(self.dd_origen.value, "")
        self.page.update()

    # -------------------------------------- paso 1: consulta y tabla
    def _resumen_empresas(self) -> str:
        sel = [e for e, v in self._empresas_val.items() if v]
        total = len(self._empresas_val)
        if not sel:
            return "Ninguna"
        if len(sel) == total:
            return f"Todas ({total})"
        if len(sel) == 1:
            return sel[0]
        return f"{len(sel)} empresas"

    def _marcar_empresa(self, empresa: str, valor: bool) -> None:
        self._empresas_val[empresa] = bool(valor)

    def _abrir_selector_empresas(self) -> None:
        """Despliega las empresas (casillas) para consultar una, varias o todas."""
        casillas = [
            ft.Checkbox(
                label=e, value=v,
                on_change=lambda ev, emp=e: self._marcar_empresa(emp, ev.control.value),
            )
            for e, v in self._empresas_val.items()
        ]

        def todas(_ev=None):
            for c in casillas:
                c.value = True
                self._marcar_empresa(c.label, True)
            self.page.update()

        def ninguna(_ev=None):
            for c in casillas:
                c.value = False
                self._marcar_empresa(c.label, False)
            self.page.update()

        def cerrar(_ev=None):
            self.tf_empresas.value = self._resumen_empresas()
            self.page.pop_dialog()
            self.page.update()

        self.page.show_dialog(
            ft.AlertDialog(
                modal=True,
                title=ft.Text("Empresas a consultar"),
                content=ft.Container(
                    content=ft.Column(casillas, tight=True, spacing=2,
                                      scroll=ft.ScrollMode.AUTO),
                    width=380,
                ),
                actions=[
                    ft.TextButton("Todas", on_click=todas),
                    ft.TextButton("Ninguna", on_click=ninguna),
                    ft.FilledButton("Listo", on_click=cerrar),
                ],
            )
        )

    def _consultar(self, _e=None) -> None:
        """Trae las solicitudes autorizadas de las empresas elegidas. Conserva los
        movimientos manuales ya capturados (solo reemplaza las del SIPP)."""
        empresas = [e for e, v in self._empresas_val.items() if v]
        if not empresas:
            self.app.avisar("Elige al menos una empresa para consultar.", ROJO)
            return
        try:
            resultados = solicitudes_devolucion.consultar(empresas)
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.app.avisar(f"No se pudo consultar: {exc}", ROJO)
            return
        manuales = [f for f in self.filas if f.manual]
        self.filas = [FilaSolicitud(self, s) for s in resultados] + manuales
        self._refrescar()
        if resultados:
            self.app.avisar(
                f"{len(resultados)} solicitud(es) autorizada(s) encontrada(s).", VERDE)
        else:
            self.app.avisar(
                "No hay solicitudes autorizadas para las empresas elegidas.", NARANJA)

    def _agregar_manual(self, _e=None) -> None:
        """Agrega una fila vacía para capturar un movimiento a mano (sin folio)."""
        self.filas.append(FilaSolicitud(self, sol=None, manual=True))
        self._refrescar()

    def eliminar_fila(self, fila: FilaSolicitud) -> None:
        if fila in self.filas:
            self.filas.remove(fila)
        self._refrescar()

    def _seleccionar_todas(self, _e=None) -> None:
        if not self.filas:
            return
        nuevo = not all(f.seleccionada for f in self.filas)
        for f in self.filas:
            f.fila.selected = nuevo
        self._actualizar_contador()
        self.page.update()

    def _refrescar(self) -> None:
        self.tabla.rows = [f.fila for f in self.filas]
        self._actualizar_contador()
        self.page.update()

    def _grupos_asignados(self) -> dict[tuple[str, str, str], list[FilaSolicitud]]:
        """Agrupa las filas ASIGNADAS por (banco, empresa origen, cuenta origen):
        cada grupo será un archivo TXT."""
        grupos: dict[tuple[str, str, str], list[FilaSolicitud]] = {}
        for f in self.filas:
            a = f.asignacion
            if not a or f.vacia():
                continue
            grupos.setdefault(
                (a["banco"], a["empresa_origen"], a["cuenta_origen"]), []).append(f)
        return grupos

    def _actualizar_contador(self) -> None:
        total = len(self.filas)
        if not total:
            self.txt_contador.value = "Sin registros."
            self.txt_grupos.value = "Sin filas asignadas."
            return
        manuales = sum(1 for f in self.filas if f.manual)
        seleccionadas = sum(1 for f in self.filas if f.seleccionada)
        asignadas = sum(1 for f in self.filas if f.asignacion)
        self.txt_contador.value = (
            f"{total} registro(s) ({manuales} manual(es)) · {seleccionadas} "
            f"seleccionado(s) · {asignadas} asignado(s)"
        )
        grupos = self._grupos_asignados()
        self.txt_grupos.value = (
            f"{len(grupos)} archivo(s) TXT a generar, con {asignadas} registro(s)."
            if grupos else "Sin filas asignadas."
        )

    # ------------------------------------------------------- asignación
    def _quitar_asignacion(self, _e=None) -> None:
        objetivo = [f for f in self.filas if f.seleccionada and f.asignacion]
        if not objetivo:
            self.app.avisar("No hay filas seleccionadas con asignación.", GRIS)
            return
        for f in objetivo:
            f.limpiar_asignacion()
        self._refrescar()
        self.app.avisar(f"Asignación quitada a {len(objetivo)} registro(s).", GRIS)

    def _asignar_a_seleccionadas(self, _e=None) -> None:
        """Asigna la cuenta origen de pago (empresa + banco + CLABE) a TODAS las
        filas seleccionadas. Se puede repetir con otra empresa origen para otro
        grupo: cada grupo generará su propio TXT."""
        seleccionadas = [f for f in self.filas if f.seleccionada]
        if not seleccionadas:
            self.app.avisar("Selecciona al menos una fila (marca las casillas).", ROJO)
            return
        empresa = self.dd_empresa.value
        if not empresa:
            self.app.avisar("Elige la empresa que dispersa (origen del pago).", ROJO)
            return
        origen = solo_digitos(self.dd_origen.value)
        if len(origen) != 18:
            self.app.avisar(
                "No hay cuenta origen para esa empresa y banco. Elige otra empresa "
                "o revisa el Excel de cuentas.", ROJO)
            return
        banco = self.dd_banco.value or ""
        fecha = solo_digitos(self.tf_fecha.value)
        folio = (self.tf_folio.value or "").strip()
        if banco == "Banregio" and len(fecha) != 8:
            self.app.avisar("La fecha debe tener 8 dígitos (DDMMAAAA).", ROJO)
            return
        if banco == "Bancomer" and not folio:
            self.app.avisar("Captura el folio (lo requiere el formato Bancomer).", ROJO)
            return
        datos = {
            "empresa_origen": empresa, "banco": banco, "cuenta_origen": origen,
            "num_cuenta": self.tf_num_cuenta.value or "",
            "fecha": fecha, "folio": folio,
        }
        for f in seleccionadas:
            f.asignar(dict(datos))
        self._refrescar()
        self.app.avisar(
            f"Cuenta origen de {empresa} ({banco}) asignada a "
            f"{len(seleccionadas)} registro(s).", VERDE)

    # ------------------------------------------------------- generación
    def _registros(self, filas: list[FilaSolicitud]) -> list[tuple] | None:
        """Valida las filas y las devuelve como [(clabe, monto, cliente, concepto)].
        Devuelve None (tras avisar) si algún dato impide generar el archivo."""
        registros = []
        for f in filas:
            if f.vacia():
                continue  # fila manual sin capturar: se ignora
            clabe, monto_txt, cliente, concepto, _empresa = f.valores()
            quien = cliente or f.txt_folio.value
            if len(clabe) != 18:
                self.app.avisar(
                    f"La CLABE de «{quien}» no tiene 18 dígitos. Corrígela para poder "
                    "generar el archivo.", ROJO)
                return None
            try:
                monto = parse_monto(monto_txt)
            except ValueError:
                self.app.avisar(f"El monto de «{quien}» no es un número válido.", ROJO)
                return None
            if monto is None:
                self.app.avisar(f"Falta capturar el monto de «{quien}».", ROJO)
                return None
            if not cliente:
                self.app.avisar("Falta el cliente/beneficiario en algún registro.", ROJO)
                return None
            registros.append((clabe, monto, cliente, concepto))
        return registros

    async def _generar_dispersion(self, _e=None) -> None:
        """Genera UN archivo TXT por cada (empresa origen + banco), con las filas
        que se le asignaron. Se guardan en una carpeta con la fecha del día."""
        grupos = self._grupos_asignados()
        if not grupos:
            self.app.avisar(
                "No hay filas con cuenta origen asignada. Selecciónalas y usa "
                "'Asignar a filas seleccionadas'.", ROJO)
            return
        # Se validan TODAS antes de escribir nada (no dejar archivos a medias).
        por_grupo: dict[tuple[str, str, str], list[tuple]] = {}
        for clave, filas in grupos.items():
            registros = self._registros(filas)
            if registros is None:
                return  # el aviso ya lo dio _registros
            if registros:
                por_grupo[clave] = registros
        if not por_grupo:
            self.app.avisar("Las filas asignadas no tienen datos que exportar.", ROJO)
            return

        destino = await self.app.picker.get_directory_path(
            dialog_title="Carpeta donde guardar los TXT de dispersión",
        )
        if not destino:
            return
        carpeta = os.path.join(
            destino, f"DISPERSION DEVOLUCIONES - {date.today().strftime('%d-%m-%Y')}")
        generados: list[str] = []
        try:
            os.makedirs(carpeta, exist_ok=True)
            for (banco, empresa, cuenta), registros in por_grupo.items():
                datos = grupos[(banco, empresa, cuenta)][0].asignacion
                if banco == "Banregio":
                    contenido = exportador_devoluciones.generar_banregio(
                        registros, datos["fecha"])
                else:  # Bancomer
                    contenido = exportador_devoluciones.generar_bancomer(
                        registros, cuenta, datos["folio"])
                nombre = _sanear_archivo(f"{empresa} - {banco}") + ".txt"
                with open(os.path.join(carpeta, nombre), "w",
                          encoding="latin-1", newline="") as fh:
                    fh.write(contenido)
                generados.append(f"{nombre} ({len(registros)} mov.)")
        except PermissionError:
            self.app.avisar(
                "No se pudo guardar: algún archivo de la carpeta está abierto. "
                "Ciérralo e intenta de nuevo.", ROJO)
            return
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.app.avisar(f"No se pudo generar la dispersión: {exc}", ROJO)
            return
        try:
            os.startfile(carpeta)  # abre la carpeta generada
        except Exception:  # noqa: BLE001 — abrirla es opcional
            pass
        self.app.avisar(
            f"{len(generados)} archivo(s) TXT generados: " + ", ".join(generados), VERDE)

    def _contexto(self) -> dict:
        return {
            "empresa": self.dd_empresa.value or "",
            "banco": self.dd_banco.value or "",
            "cuenta_origen": self.dd_origen.value or "",
            "num_cuenta": self.tf_num_cuenta.value or "",
            "fecha": self.tf_fecha.value or "",
        }

    async def _generar_excel(self, _e=None) -> None:
        """Reporte Excel con TODOS los registros de la tabla (SIPP + manuales)."""
        registros = self._registros(self.filas)
        if registros is None:
            return
        if not registros:
            self.app.avisar("No hay registros que exportar.", ROJO)
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
        except PermissionError:
            self.app.avisar(
                "No se pudo guardar: el archivo está abierto en Excel. Ciérralo e "
                "intenta de nuevo (o guarda con otro nombre).", ROJO)
            return
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.app.avisar(f"No se pudo generar el Excel: {exc}", ROJO)
            return
        self.app.avisar(
            f"Reporte Excel generado con {len(registros)} movimiento(s).", VERDE)
