"""Pantalla: Alta de beneficiarios.

Carga de estados de cuenta (OCR), tabla editable de beneficiarios, guardado en
la base y exportación (TXT de dispersión Bancomer o Excel de alta Banregio).
"""

from __future__ import annotations

import asyncio
import os
from datetime import date

import flet as ft
from playwright.async_api import TimeoutError as PlaywrightTimeoutError

from core import (
    db, exportador_alta_bancomer, exportador_alta_banregio, exportador_alta_spei,
    ocr, reporte_cuentas,
)
from core.catalogo_bancos import banco_desde_clabe
from core.extractores import extraer_clabes, extraer_datos, nombre_desde_archivo, validar_clabe
from core.rpa_sipp import (
    BucleRpa, ErrorSipp, SesionSipp, asegurar_navegador, necesita_navegador,
)
from ui.comun import (
    EXTENSIONES, GRIS, NARANJA, ROJO, VERDE,
    W_ACCIONES, W_BANCO, W_CLABE, W_ESTADO, W_MONTO, W_NOMBRE, W_TIPO,
    celda_centrada, encabezado_col, fmt_monto, parse_monto, solo_digitos,
    tarjeta, validar,
)

# Tipos de beneficiario que el RPA puede filtrar en el reporte del SIPP.
# (etiqueta visible, clave de coincidencia que usa el RPA para marcar la opción).
TIPOS_BENEFICIARIO = [
    ("Proveedores", "proveedor"),
    ("Acreedores Diversos", "acreedor"),
    ("Deudores Diversos", "deudor"),
]

class _CargaDetenida(Exception):
    """Señal interna: el usuario pulsó 'Detener' mientras se analizaba un archivo.
    El archivo en proceso se descarta y vuelve a la cola de pendientes."""


class _TokenCarga:
    """Testigo de cancelación de UNA corrida de carga.

    Cada corrida tiene el suyo. Es clave que no sea una bandera compartida: al
    detener, el OCR del archivo en curso queda huérfano en su hilo (no se puede
    matar un hilo en Python) y va consultando este testigo para salirse. Si el
    testigo se reiniciara al arrancar la siguiente corrida, ese OCR huérfano
    creería que ya no está cancelado y seguiría analizando el documento entero,
    compitiendo por la CPU y retrasando la reanudación."""

    __slots__ = ("cancelado",)

    def __init__(self) -> None:
        self.cancelado = False


# Fondos de fila según la conciliación con el reporte de cuentas.
SIN_COINCIDENCIA = ft.Colors.with_opacity(0.16, ft.Colors.RED)      # no coincide
SUGERIDO_NOMBRE = ft.Colors.with_opacity(0.18, ft.Colors.AMBER)     # CLABE sugerida por nombre
SOLO_REPORTE = ft.Colors.with_opacity(0.14, ft.Colors.BLUE)         # dato solo del reporte (sin estado de cuenta)


class FilaBeneficiario:
    """Una fila editable de la tabla (un beneficiario, guardado o pendiente)."""

    def __init__(self, seccion: "SeccionAltaBeneficiarios", id_: int | None, clabe: str,
                 beneficiario: str, alias: str, email: str, origen: str = "",
                 monto: float | None = None, ruta_archivo: str | None = None,
                 desde_reporte: bool = False):
        self.seccion = seccion
        self.id = id_          # None mientras no se haya guardado en la base
        self.origen = origen   # nombre del archivo de origen (informativo)
        self.ruta_archivo = ruta_archivo  # ruta completa para previsualizar
        # True si la fila se creó a partir del reporte (no hay estado de cuenta
        # descargado). Se pinta en azul para distinguirla y su 'Ver archivo' avisa.
        self.desde_reporte = desde_reporte
        # Estado de conciliación con el reporte de cuentas:
        #   None     -> no aplica (sin reporte importado)
        #   "ok"     -> la CLABE coincide con el reporte
        #   "nombre" -> CLABE sugerida porque el nombre coincide (fila ámbar)
        #   "sin"    -> no coincide ni por CLABE ni por nombre (fila roja)
        self.conciliacion: str | None = None

        # Sin max_length (evita el contador "X/18" que desalineaba la fila); el
        # límite de 18 dígitos se mantiene por código en _cambio_clabe.
        self.tf_clabe = ft.TextField(
            value=clabe, dense=True, width=W_CLABE, text_size=12,
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
        # Tipo de beneficiario: se llena al conciliar con el reporte ('—' si no).
        self.txt_tipo = ft.Text("—", size=12, text_align=ft.TextAlign.CENTER)
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
                ft.DataCell(celda_centrada(self.txt_tipo, W_TIPO)),
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
        # Limita la CLABE a 18 dígitos (sin max_length, para no mostrar contador).
        limpio = solo_digitos(self.tf_clabe.value)[:18]
        if limpio != (self.tf_clabe.value or ""):
            self.tf_clabe.value = limpio
        self.txt_banco.value = banco_desde_clabe(limpio) or "—"
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
        self.set_tipo(rep.tipo)
        self._actualizar_estado()

    def set_tipo(self, tipo: str) -> None:
        """Muestra el tipo de beneficiario tomado del reporte (o '—' si no aplica)."""
        self.txt_tipo.value = tipo or "—"

    def marcar_conciliacion(self, estado: str | None) -> None:
        """Pinta la fila según la conciliación: rojo si no coincide; ámbar si la
        CLABE se sugirió por nombre (revisar); azul si el dato viene solo del
        reporte (sin estado de cuenta); sin color si coincide por CLABE o si no
        hay reporte importado."""
        self.conciliacion = estado
        self.fila.color = (
            SIN_COINCIDENCIA if estado == "sin"
            else SUGERIDO_NOMBRE if estado == "nombre"
            else SOLO_REPORTE if self.desde_reporte
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

    # Empresa y sucursal del SIPP donde el RPA descarga los anexos.
    RPA_EMPRESA = "ASKE"
    RPA_SUCURSAL = "Corporativo"

    def __init__(self, app):
        self.app = app
        self.page = app.page
        self.picker = app.picker
        self.filas: list[FilaBeneficiario] = []
        # Texto de la barra de búsqueda del listado (filtra por beneficiario/CLABE).
        self.filtro_texto = ""
        # Control de la carga de estados de cuenta (para poder detener/reanudar):
        # _cargando indica que hay una carga en curso; _token_carga es el testigo
        # de cancelación de esa corrida; _rutas_pendientes guarda los que faltaron
        # al detener, para reanudar desde el último capturado.
        self._cargando = False
        self._token_carga: _TokenCarga | None = None
        self._rutas_pendientes: list[str] = []
        # Reporte de cuentas importado (para conciliar). Vacío = no se subió.
        self.catalogo_reporte: dict[str, reporte_cuentas.CuentaReporte] = {}
        self.nombre_reporte = ""
        # RPA del SIPP (descarga de anexos). Se crean al primer uso.
        self.sesion_rpa: SesionSipp | None = None
        self.bucle_rpa: BucleRpa | None = None
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
        # Cargar una carpeta completa: evita el diálogo de multiselección (que
        # falla al elegir muchos archivos con nombres largos) y es lo más cómodo
        # para la carpeta que deja el RPA.
        self.btn_cargar_carpeta = ft.OutlinedButton(
            content="Cargar carpeta completa",
            icon=ft.Icons.FOLDER_OPEN,
            on_click=self._seleccionar_carpeta,
        )
        # Detener la carga en curso (visible solo mientras se procesa) y reanudar
        # desde el último registro capturado (visible solo si quedaron pendientes).
        self.btn_detener_carga = ft.OutlinedButton(
            content="Detener", icon=ft.Icons.STOP_CIRCLE_OUTLINED, visible=False,
            on_click=self._detener_carga, style=ft.ButtonStyle(color=ROJO),
            tooltip="Detiene la carga de inmediato. Los archivos ya cargados se "
                    "conservan; el que se estaba analizando y los que faltan quedan "
                    "pendientes para reanudar.",
        )
        self.btn_reanudar_carga = ft.FilledButton(
            content="Reanudar carga", icon=ft.Icons.PLAY_ARROW, visible=False,
            on_click=self._reanudar_carga,
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
        # Registra en la tabla las cuentas del reporte que NO se descargaron como
        # estado de cuenta (usa el reporte como fuente para no perder registros).
        self.btn_completar_reporte = ft.OutlinedButton(
            content="Registrar faltantes del reporte",
            icon=ft.Icons.PLAYLIST_ADD, visible=False,
            on_click=self._completar_desde_reporte,
            tooltip="Agrega a la tabla las cuentas del reporte que no tienen estado "
                    "de cuenta descargado, con su nombre, correo y tipo.",
        )
        seccion_carga = tarjeta(
            "1. Cargar estados de cuenta",
            ft.Column(
                [
                    ft.Row(
                        [
                            self.btn_cargar,
                            self.btn_cargar_carpeta,
                            self.btn_detener_carga,
                            self.btn_reanudar_carga,
                            self.anillo,
                            ft.Text("Selecciona varios archivos, o carga una carpeta "
                                    "completa (recomendado para la carpeta del RPA).",
                                    color=GRIS, size=12, italic=True),
                        ],
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        spacing=12, wrap=True,
                    ),
                    self.txt_estado,
                    ft.Divider(height=1),
                    ft.Row(
                        [
                            self.btn_reporte,
                            self.btn_quitar_reporte,
                            self.btn_completar_reporte,
                            ft.Text(
                                "Opcional: súbelo para conciliar las cuentas; se "
                                "completa el nombre/correo y se marcan en rojo las "
                                "que no aparezcan en el reporte.",
                                color=GRIS, size=12, italic=True,
                            ),
                        ],
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                        spacing=12, wrap=True,
                    ),
                    self.txt_estado_rep,
                ],
                spacing=8,
            ),
        )

        seccion_rpa = self._construir_rpa()

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
                ft.DataColumn(label=encabezado_col("Tipo", W_TIPO)),
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
        # Formato de exportación: Bancomer -> carpeta con TXT de alta de cuentas ;
        # Banregio -> Excel de alta de cuentas.
        self.dd_formato = ft.Dropdown(
            label="Exportar para", width=300, value="BancomerAlta",
            options=[
                ft.dropdown.Option(key="BancomerAlta", text="Bancomer · Alta de cuentas (Carpeta TXT)"),
                ft.dropdown.Option(key="Banregio", text="Banregio · Alta (Excel)"),
            ],
            on_select=self._cambio_formato_export,
        )
        self.btn_export = ft.FilledButton(
            content="Generar carpeta de alta (Bancomer)", icon=ft.Icons.CREATE_NEW_FOLDER,
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
        # Búsqueda dentro del listado + contador de registros (mejora la consulta).
        self.tf_buscar = ft.TextField(
            hint_text="Buscar por beneficiario, alias o CLABE…",
            prefix_icon=ft.Icons.SEARCH, dense=True, expand=True,
            on_change=self._filtrar,
        )
        self.txt_contador = ft.Text("0 registros", size=12, color=GRIS,
                                    weight=ft.FontWeight.BOLD)
        barra_busqueda = ft.Row(
            [self.tf_buscar, self.txt_contador],
            spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        seccion_tabla = tarjeta(
            "2. Revisión y edición de beneficiarios",
            ft.Column(
                [
                    barra,
                    self._leyenda(),
                    barra_busqueda,
                    ft.Row([self.tabla], scroll=ft.ScrollMode.AUTO),
                ],
                spacing=12,
            ),
        )

        return ft.Column(
            [seccion_carga, seccion_rpa, seccion_tabla],
            spacing=14,
            scroll=ft.ScrollMode.AUTO,
            expand=True,
        )

    def _construir_rpa(self) -> ft.Control:
        """Tarjeta del RPA: descarga los anexos del SIPP por rango de fechas."""
        # Selectores de fecha tipo calendario (DatePicker). El campo es de solo
        # lectura y abre el calendario al hacer clic; el RPA lee su texto.
        self.dp_ini = ft.DatePicker(
            first_date=date(2020, 1, 1), last_date=date(2035, 12, 31),
            help_text="Fecha inicio de la consulta",
            on_change=lambda e: self._fecha_elegida(self.tf_rpa_ini, self.dp_ini),
        )
        self.dp_fin = ft.DatePicker(
            first_date=date(2020, 1, 1), last_date=date(2035, 12, 31),
            help_text="Fecha fin de la consulta",
            on_change=lambda e: self._fecha_elegida(self.tf_rpa_fin, self.dp_fin),
        )
        self.tf_rpa_ini = ft.TextField(
            label="Fecha inicio consulta", hint_text="DD/MM/AAAA", width=210,
            read_only=True, suffix_icon=ft.Icons.CALENDAR_MONTH,
            on_click=lambda e: self.page.show_dialog(self.dp_ini),
        )
        self.tf_rpa_fin = ft.TextField(
            label="Fecha fin consulta", hint_text="DD/MM/AAAA", width=210,
            read_only=True, suffix_icon=ft.Icons.CALENDAR_MONTH,
            on_click=lambda e: self.page.show_dialog(self.dp_fin),
        )
        self.btn_rpa = ft.FilledButton(
            content="Iniciar descarga", icon=ft.Icons.SMART_TOY_OUTLINED,
            on_click=self._iniciar_rpa_anexos,
        )
        self.anillo_rpa = ft.ProgressRing(width=18, height=18, stroke_width=2, visible=False)
        # Selector de tipos de beneficiario a descargar (uno, dos o los tres).
        # Arrancan los tres activos (comportamiento previo: descarga todo). El
        # estado vive en _tipos_val; se edita en un diálogo que abre el campo
        # 'tf_tipos', para que se vea como un desplegable (igual que las fechas).
        self._tipos_val = {clave: True for _etq, clave in TIPOS_BENEFICIARIO}
        self.tf_tipos = ft.TextField(
            label="Tipos de beneficiario", width=240, read_only=True,
            value=self._resumen_tipos(), suffix_icon=ft.Icons.ARROW_DROP_DOWN,
            on_click=lambda e: self._abrir_selector_tipos(),
        )
        # Barra de avance de la descarga de anexos (determinada: hechos/total).
        # Oculta hasta que empieza la descarga y se conoce el total de registros.
        self.pb_rpa = ft.ProgressBar(width=360, value=0, visible=False)
        self.txt_rpa = ft.Text("", color=GRIS, size=12)
        return tarjeta(
            "Descargar anexos del SIPP",
            ft.Column(
                [
                    ft.Row(
                        [self.tf_tipos, self.tf_rpa_ini, self.tf_rpa_fin,
                         self.btn_rpa, self.anillo_rpa],
                        spacing=12, wrap=True,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    ),
                    ft.Text(
                        f"Entra al SIPP con las credenciales de Configuración, empresa "
                        f"{self.RPA_EMPRESA} / sucursal {self.RPA_SUCURSAL}, filtra los "
                        "tipos de beneficiario elegidos por el rango de fechas y "
                        "descarga los anexos.",
                        color=GRIS, size=12, italic=True,
                    ),
                    self.pb_rpa,
                    self.txt_rpa,
                ],
                spacing=8,
            ),
        )

    def _fecha_elegida(self, campo: ft.TextField, dp: ft.DatePicker) -> None:
        """Vuelca la fecha elegida en el calendario al campo, como DD/MM/AAAA."""
        if dp.value:
            campo.value = dp.value.strftime("%d/%m/%Y")
            self.page.update()

    def _tipos_seleccionados(self) -> list[str]:
        """Claves de los tipos de beneficiario activos en el selector."""
        return [clave for _etq, clave in TIPOS_BENEFICIARIO if self._tipos_val.get(clave)]

    def _resumen_tipos(self) -> str:
        """Texto que muestra el campo desplegable según los tipos elegidos."""
        seleccion = [etq for etq, clave in TIPOS_BENEFICIARIO if self._tipos_val.get(clave)]
        if not seleccion:
            return "Ninguno"
        if len(seleccion) == len(TIPOS_BENEFICIARIO):
            return "Todos los tipos"
        if len(seleccion) == 1:
            return seleccion[0]
        return f"{len(seleccion)} tipos seleccionados"

    def _marcar_tipo(self, clave: str, valor: bool) -> None:
        self._tipos_val[clave] = bool(valor)

    def _abrir_selector_tipos(self) -> None:
        """Despliega las opciones de tipo de beneficiario (casillas) en un diálogo.
        Al cerrar, refleja la selección en el campo desplegable."""
        casillas = [
            ft.Checkbox(
                label=etq, value=self._tipos_val.get(clave),
                on_change=lambda e, c=clave: self._marcar_tipo(c, e.control.value),
            )
            for etq, clave in TIPOS_BENEFICIARIO
        ]

        def cerrar(_e=None):
            self.tf_tipos.value = self._resumen_tipos()
            self.page.pop_dialog()
            self.page.update()

        self.page.show_dialog(
            ft.AlertDialog(
                modal=True,
                title=ft.Text("Tipos de beneficiario a descargar"),
                content=ft.Container(
                    content=ft.Column(casillas, tight=True, spacing=2),
                    width=320,
                ),
                actions=[ft.FilledButton("Listo", on_click=cerrar)],
            )
        )

    async def _iniciar_rpa_anexos(self, _e) -> None:
        """Lanza el RPA: login en el SIPP, filtra los tipos de beneficiario
        elegidos por el rango de fechas y descarga los anexos en la carpeta que
        elija el usuario."""
        usuario, contrasena = self.app.config.credenciales()
        if not usuario or not contrasena:
            self.avisar("Captura usuario y contraseña en Configuración (ícono ⚙).", ROJO)
            return
        fi = solo_digitos(self.tf_rpa_ini.value)
        ff = solo_digitos(self.tf_rpa_fin.value)
        if len(fi) != 8 or len(ff) != 8:
            self.avisar("Captura fecha inicio y fecha fin (DD/MM/AAAA).", ROJO)
            return
        tipos = self._tipos_seleccionados()
        if not tipos:
            self.avisar("Selecciona al menos un tipo de beneficiario a descargar.", ROJO)
            return
        carpeta = await self.picker.get_directory_path(
            dialog_title="Carpeta donde guardar los anexos descargados",
        )
        if not carpeta:
            return

        await self._cerrar_sesion_rpa()  # cierra una corrida previa si la hubiera

        # Primera vez (app empaquetada): descarga Chromium mostrando aviso, igual
        # que en la pantalla de Dispersión (No Pemex), para que no parezca colgada.
        if necesita_navegador():
            if self.bucle_rpa is None:
                self.bucle_rpa = BucleRpa()
            self.btn_rpa.disabled = True
            self.pb_rpa.value = None  # indeterminada
            self.pb_rpa.visible = True
            self.txt_rpa.value = "Instalando componentes extra, espere un momento…"
            self.page.update()
            try:
                await asyncio.wrap_future(self.bucle_rpa.enviar(asegurar_navegador()))
            except Exception as exc:  # noqa: BLE001 — se reporta al usuario
                self._fin_rpa(f"RPA: no se pudo preparar el navegador: {exc}", ROJO)
                return
            self.pb_rpa.visible = False
            self.pb_rpa.value = 0

        self.btn_rpa.disabled = True
        self.anillo_rpa.visible = True
        self.txt_rpa.value = "RPA: iniciando sesión en el SIPP…"
        self.page.update()
        try:
            descargados, danados = await self._correr_rpa_anexos(
                usuario, contrasena, fi, ff, carpeta, tipos,
            )
        except ErrorSipp as exc:
            await self._cerrar_sesion_rpa()
            self._fin_rpa(f"RPA: {exc}", ROJO)
            return
        except PlaywrightTimeoutError:
            # Timeout de navegación de Playwright: casi siempre es la conexión o
            # que el portal del SIPP tardó/no respondió, no un fallo del RPA.
            await self._cerrar_sesion_rpa()
            self._fin_rpa(
                "RPA: la página del SIPP tardó demasiado en responder. Suele ser "
                "por una conexión a internet lenta o inestable (o el portal está "
                "caído/muy lento). Revisa tu conexión e inténtalo de nuevo.",
                ROJO,
            )
            return
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            await self._cerrar_sesion_rpa()
            self._fin_rpa(f"RPA: error inesperado: {exc}", ROJO)
            return
        if descargados == 0 and danados == 0:
            self._fin_rpa("RPA: no se encontraron anexos para el rango elegido.", NARANJA)
            return
        mensaje = f"RPA: {descargados} anexo(s) descargado(s) en la carpeta elegida."
        color = VERDE
        if danados:
            mensaje += (
                f" {danados} llegaron dañados y se apartaron en la subcarpeta "
                "'_no_validos' (revísalos manualmente)."
            )
            color = NARANJA
        self._fin_rpa(mensaje, color)

    async def _correr_rpa_anexos(
        self, usuario, contrasena, fi, ff, carpeta, tipos,
    ) -> tuple[int, int]:
        """Ejecuta todo el flujo del RPA en el bucle del hilo dedicado (Playwright
        requiere que todo corra en el mismo loop). Devuelve (descargados, dañados)."""
        if self.bucle_rpa is None:
            self.bucle_rpa = BucleRpa()
        self.sesion_rpa = SesionSipp(headless=False)
        sesion = self.sesion_rpa

        async def flujo() -> tuple[int, int]:
            await sesion.iniciar()
            await sesion.login(usuario, contrasena)
            await sesion.seleccionar_empresa_sucursal(self.RPA_EMPRESA, self.RPA_SUCURSAL)
            return await sesion.descargar_anexos_beneficiarios(
                fi, ff, carpeta, tipos=tipos, progreso=self._progreso_rpa,
            )

        return await asyncio.wrap_future(self.bucle_rpa.enviar(flujo()))

    def _progreso_rpa(self, hechos: int, total: int) -> None:
        """Callback que el RPA (en su hilo aparte) llama para reportar avance.
        Agenda la actualización de la barra en el hilo/loop de la interfaz con
        page.run_task (seguro entre hilos)."""
        self.page.run_task(self._aplicar_progreso_rpa, hechos, total)

    async def _aplicar_progreso_rpa(self, hechos: int, total: int) -> None:
        """Refleja el avance de la descarga en la barra de progreso del RPA."""
        if total <= 0:
            self.pb_rpa.visible = False
            self.txt_rpa.value = "RPA: no se encontraron anexos para el rango elegido."
        else:
            self.pb_rpa.visible = True
            self.pb_rpa.value = hechos / total
            self.txt_rpa.value = f"RPA: descargando anexos… {hechos} de {total}"
        self.page.update()

    async def _cerrar_sesion_rpa(self) -> None:
        """Cierra el navegador del RPA si quedó abierto (best-effort)."""
        if self.sesion_rpa is None:
            return
        sesion = self.sesion_rpa
        self.sesion_rpa = None
        try:
            if self.bucle_rpa is not None:
                await asyncio.wrap_future(self.bucle_rpa.enviar(sesion.cerrar()))
        except Exception:  # noqa: BLE001 — el cierre no debe propagar errores
            pass

    def _fin_rpa(self, mensaje: str, color: str) -> None:
        """Restablece los controles del RPA y muestra el resultado."""
        self.btn_rpa.disabled = False
        self.anillo_rpa.visible = False
        self.pb_rpa.visible = False
        self.pb_rpa.value = 0
        self.txt_rpa.value = mensaje
        self.page.update()
        self.avisar(mensaje, color)

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
        chip_azul = swatch(SOLO_REPORTE, ft.Colors.BLUE,
                          "Fila azul: dato del reporte (sin estado de cuenta)")
        return ft.Row(
            [ft.Text("Leyenda:", size=12, weight=ft.FontWeight.BOLD, color=GRIS),
             *chips, chip_ambar, chip_rojo, chip_azul],
            spacing=18,
            wrap=True,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )

    # ========================================================= utilidades
    def _redibujar_tabla(self) -> None:
        visibles = self._filas_visibles()
        self.tabla.rows = [f.fila for f in visibles]
        self._ajustar_anchos()
        self._remarcar_conciliacion()
        self._actualizar_resumen()
        self._actualizar_contador(len(visibles))
        self._actualizar_btn_completar()
        self._refrescar_candado_export()
        self.page.update()

    # ----------------------------------------------- búsqueda / contador
    def _filas_visibles(self) -> list[FilaBeneficiario]:
        """Filas que pasan el filtro de la barra de búsqueda (por beneficiario,
        alias o CLABE). Sin texto de búsqueda, devuelve todas."""
        texto = self.filtro_texto
        if not texto:
            return self.filas
        digitos = solo_digitos(texto)
        visibles = []
        for f in self.filas:
            benef = (f.tf_benef.value or "").lower()
            alias = (f.tf_alias.value or "").lower()
            clabe = solo_digitos(f.tf_clabe.value)
            if texto in benef or texto in alias or (digitos and digitos in clabe):
                visibles.append(f)
        return visibles

    def _filtrar(self, _e=None) -> None:
        """Aplica el texto de la barra de búsqueda y redibuja el listado."""
        self.filtro_texto = (self.tf_buscar.value or "").strip().lower()
        self._redibujar_tabla()

    def _actualizar_contador(self, mostrados: int | None = None) -> None:
        """Refresca el contador de registros del listado (total y, al filtrar,
        cuántos se están mostrando)."""
        total = len(self.filas)
        if mostrados is None:
            mostrados = len(self._filas_visibles())
        if total == 0:
            self.txt_contador.value = "0 registros"
        elif mostrados == total:
            self.txt_contador.value = f"{total} registro(s)"
        else:
            self.txt_contador.value = f"Mostrando {mostrados} de {total} registro(s)"

    def _actualizar_btn_completar(self) -> None:
        """Muestra el botón 'Registrar faltantes del reporte' con la cuenta de
        cuentas del reporte que aún no están en la tabla."""
        if not self.catalogo_reporte:
            self.btn_completar_reporte.visible = False
            return
        existentes = {solo_digitos(f.tf_clabe.value) for f in self.filas}
        faltan = sum(1 for clabe in self.catalogo_reporte if clabe not in existentes)
        self.btn_completar_reporte.visible = True
        self.btn_completar_reporte.disabled = faltan == 0
        self.btn_completar_reporte.content = (
            f"Registrar faltantes del reporte ({faltan})" if faltan
            else "Reporte completo (sin faltantes)"
        )

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
        self.btn_completar_reporte.visible = False
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
                f.set_tipo("")
                sin_match += 1
            else:
                f.marcar_conciliacion(None)  # CLABE incompleta: el ícono ya avisa
                f.set_tipo("")
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
                f.set_tipo("")
                continue
            if f.conciliacion == "nombre":
                f.marcar_conciliacion("nombre")  # preserva el aviso de revisión
                continue
            clabe = solo_digitos(f.tf_clabe.value)
            if len(clabe) != 18:
                f.marcar_conciliacion(None)
                f.set_tipo("")
            elif clabe in self.catalogo_reporte:
                f.marcar_conciliacion("ok")
                f.set_tipo(self.catalogo_reporte[clabe].tipo)
            else:
                f.marcar_conciliacion("sin")
                f.set_tipo("")

    def _conciliar_una(self, fila: FilaBeneficiario) -> None:
        """Concilia una sola fila por CLABE (al editar su CLABE en vivo). No
        sugiere CLABE por nombre aquí para no pisar lo que el usuario escribe."""
        if not self.catalogo_reporte:
            fila.marcar_conciliacion(None)
            fila.set_tipo("")
            return
        clabe = solo_digitos(fila.tf_clabe.value)
        if len(clabe) != 18:
            fila.marcar_conciliacion(None)
            fila.set_tipo("")
            return
        rep = self.catalogo_reporte.get(clabe)
        if rep:
            fila.aplicar_reporte(rep)
            fila.marcar_conciliacion("ok")
        else:
            fila.marcar_conciliacion("sin")
            fila.set_tipo("")

    def _completar_desde_reporte(self, _e=None) -> None:
        """Usa el reporte como fuente: agrega a la tabla las cuentas del reporte
        que NO tienen un estado de cuenta cargado (comparando por CLABE), con su
        nombre, correo y tipo. Así, si el RPA descargó menos anexos que registros
        hay en el reporte, esos faltantes quedan igualmente registrados."""
        if not self.catalogo_reporte:
            self.avisar("Primero importa el reporte de cuentas (Excel).", ROJO)
            return
        existentes = {solo_digitos(f.tf_clabe.value) for f in self.filas}
        faltantes = [
            rep for clabe, rep in self.catalogo_reporte.items()
            if clabe not in existentes
        ]
        if not faltantes:
            self.avisar(
                "No hay faltantes: todas las cuentas del reporte ya están en la tabla.",
                GRIS,
            )
            return
        for rep in faltantes:
            fila = FilaBeneficiario(
                self, None, rep.clabe, rep.beneficiario, rep.beneficiario,
                rep.correo, origen="(reporte)", desde_reporte=True,
            )
            fila.set_tipo(rep.tipo)
            fila.marcar_conciliacion("ok")  # coincide por CLABE; se pinta en azul
            self.filas.append(fila)
        self._redibujar_tabla()
        self.avisar(
            f"{len(faltantes)} registro(s) agregados desde el reporte (en azul, sin "
            "estado de cuenta). Revisa el monto y guárdalos.",
            VERDE,
        )

    def _cambio_formato_export(self, _e=None) -> None:
        """Ajusta el botón de exportar según el formato elegido."""
        if self.dd_formato.value == "Banregio":
            self.btn_export.content = "Exportar Excel (Banregio)"
            self.btn_export.icon = ft.Icons.TABLE_VIEW
        else:  # BancomerAlta
            self.btn_export.content = "Generar carpeta de alta (Bancomer)"
            self.btn_export.icon = ft.Icons.CREATE_NEW_FOLDER
        self._refrescar_candado_export()

    def _refrescar_candado_export(self) -> None:
        """Bloquea la exportación si falta un monto. El candado aplica al alta de
        cuentas Bancomer (el monto va dentro del TXT); el alta Banregio no usa
        montos."""
        if self.dd_formato.value == "Banregio":
            self.btn_export.disabled = False
            self.btn_export.tooltip = "Genera el archivo Excel de alta para Banregio"
            self.page.update()
            return
        sin_monto = sum(1 for b in db.listar() if b.monto is None)
        self.btn_export.disabled = sin_monto > 0
        self.btn_export.tooltip = (
            f"Hay {sin_monto} registro(s) guardado(s) sin monto. Captura el monto "
            "y guarda para poder exportar."
            if sin_monto else "Genera la carpeta con los TXT de alta de cuentas Bancomer"
        )
        self.page.update()

    def _ajustar_anchos(self) -> None:
        """Reparte el ancho disponible entre los campos de texto largos
        (beneficiario, alias, email) para que crezcan al agrandar la ventana."""
        ancho = self.page.width or 1180
        # Columnas de ancho fijo: estado, CLABE, monto, banco, tipo, acciones.
        fijo = W_ESTADO + W_CLABE + W_MONTO + W_BANCO + W_TIPO + W_ACCIONES
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
        """Carga por selección de archivos (uno o varios) desde el diálogo."""
        archivos = await self.picker.pick_files(
            dialog_title="Selecciona uno o varios estados de cuenta",
            allowed_extensions=EXTENSIONES,
            allow_multiple=True,
        )
        if not archivos:
            return
        await self._procesar_archivos([a.path for a in archivos])

    async def _seleccionar_carpeta(self, _e) -> None:
        """Carga TODOS los estados de cuenta (por extensión) de una carpeta. Es lo
        más cómodo para la carpeta que deja el RPA y evita el diálogo de
        multiselección, que falla al elegir muchos archivos con nombres largos.
        Solo toma el primer nivel, así que ignora la subcarpeta '_no_validos'."""
        carpeta = await self.picker.get_directory_path(
            dialog_title="Elige la carpeta con los estados de cuenta",
        )
        if not carpeta:
            return
        exts = tuple("." + e.lower() for e in EXTENSIONES)
        try:
            rutas = [
                os.path.join(carpeta, n) for n in sorted(os.listdir(carpeta))
                if n.lower().endswith(exts) and os.path.isfile(os.path.join(carpeta, n))
            ]
        except OSError as exc:
            self.avisar(f"No se pudo leer la carpeta: {exc}", ROJO)
            return
        if not rutas:
            self.avisar("La carpeta no tiene estados de cuenta (PDF o imagen).", NARANJA)
            return
        await self._procesar_archivos(rutas)

    def _detener_carga(self, _e=None) -> None:
        """Corta la carga de inmediato: el análisis del archivo en proceso se
        abandona (no se agrega a la tabla) y vuelve a la cola de pendientes junto
        con los que faltan. Los archivos ya cargados se conservan."""
        if not self._cargando or self._token_carga is None:
            return
        self._token_carga.cancelado = True
        self.btn_detener_carga.disabled = True
        self.btn_detener_carga.content = "Deteniendo…"
        self.page.update()

    async def _ocr_cancelable(self, token: "_TokenCarga", func, ruta, *extra):
        """Corre `func` (OCR, pesado) en un hilo pasándole el callback de
        cancelación de ESTA corrida.

        El OCR es cancelable de verdad: consulta `cancelado` entre páginas y, sobre
        todo, MATA el proceso tesseract si se detiene (ver core.ocr). Por eso el
        hilo termina casi enseguida al pulsar Detener; se ESPERA su fin (no se
        abandona, así no quedan hilos ni tesseract huérfanos ocupando CPU) y se
        traduce OCRCancelado en _CargaDetenida para dejar el archivo pendiente.

        Se usa el `token` de la corrida (no una bandera compartida) para que, si
        alguna vez quedara un hilo colgado, siga viendo SU cancelación y no estorbe
        a una corrida posterior."""
        try:
            return await asyncio.to_thread(
                func, ruta, *extra, cancelado=lambda: token.cancelado)
        except ocr.OCRCancelado:
            raise _CargaDetenida()

    async def _reanudar_carga(self, _e=None) -> None:
        """Continúa la carga con los archivos que quedaron pendientes al detener."""
        pendientes = self._rutas_pendientes
        if not pendientes:
            return
        self._rutas_pendientes = []
        self.btn_reanudar_carga.visible = False
        await self._procesar_archivos(pendientes)

    async def _ocr_archivo(self, token: "_TokenCarga", ruta: str) -> "FilaBeneficiario":
        """Analiza UN estado de cuenta (OCR + extracción) y devuelve su fila.
        Lanza _CargaDetenida si se detuvo la carga a media marcha (para que el
        archivo quede pendiente). Aísla el trabajo de un archivo para poder correr
        varios en paralelo desde _procesar_archivos."""
        nombre = os.path.basename(ruta)
        texto, uso_ocr = await self._ocr_cancelable(token, ocr.extraer_texto, ruta)
        datos = extraer_datos(texto)
        # Si la capa de texto no dio nada útil (p. ej. impresión de un correo de
        # Outlook con el estado como imagen), forzar OCR.
        if not datos.clabe and not datos.beneficiario and not uso_ocr:
            texto, _ = await self._ocr_cancelable(token, ocr.extraer_texto, ruta, True)
            datos = extraer_datos(texto)
        # Último recurso para la CLABE: OCR en modo "texto disperso", que recupera
        # números en tablas/celdas que la segmentación normal omite (p. ej. la
        # tabla de una carta de asignación de cuenta).
        if not datos.clabe:
            texto_sp = await self._ocr_cancelable(token, ocr.texto_disperso, ruta)
            clabes_sp = extraer_clabes(texto_sp)
            if clabes_sp:
                datos.clabe = clabes_sp[0]
                datos.banco = banco_desde_clabe(datos.clabe)
        # El nombre del archivo (si parece nombre de persona) es la fuente más
        # confiable del beneficiario; tiene prioridad sobre el OCR.
        beneficiario = nombre_desde_archivo(nombre) or datos.beneficiario
        return FilaBeneficiario(
            self, None, datos.clabe, beneficiario, beneficiario,
            datos.email, origen=nombre, ruta_archivo=ruta,
        )

    async def _procesar_archivos(self, rutas: list[str]) -> None:
        """Procesa una lista de rutas (OCR + extracción) y las agrega a la tabla.
        Compartido por la carga por archivos y por carpeta. Se puede detener con
        el botón Detener: el archivo en proceso termina y los que falten quedan en
        _rutas_pendientes para reanudar desde el último capturado."""
        if not rutas:
            return
        token = _TokenCarga()  # testigo de ESTA corrida (ver _TokenCarga)
        self._token_carga = token
        self._cargando = True
        self.btn_cargar.disabled = True
        self.btn_cargar_carpeta.disabled = True
        self.btn_reanudar_carga.visible = False
        self.btn_detener_carga.visible = True
        self.btn_detener_carga.disabled = False
        self.btn_detener_carga.content = "Detener"
        self.anillo.visible = True
        self.page.update()

        total = len(rutas)
        identificados = 0
        procesados = 0                       # completados (ok o error) en esta corrida
        errores: list[tuple[str, str]] = []  # (nombre_archivo, motivo)
        completado = [False] * total
        # OCR en PARALELO: se analizan varios archivos a la vez (cada uno lanza su
        # tesseract), aprovechando varios núcleos. Se deja uno libre para la
        # interfaz/el sistema y se topa a 4 para no agotar RAM.
        n_workers = max(1, min(4, (os.cpu_count() or 2) - 1))
        limite = asyncio.Semaphore(n_workers)

        async def procesar_uno(idx: int, ruta: str) -> None:
            nonlocal identificados, procesados
            async with limite:
                if token.cancelado:
                    return  # no se alcanzó a empezar: queda pendiente
                nombre = os.path.basename(ruta)
                try:
                    fila = await self._ocr_archivo(token, ruta)
                except _CargaDetenida:
                    return  # abortado a media marcha: queda pendiente
                except Exception as exc:  # noqa: BLE001 — se reporta al usuario
                    errores.append((nombre, self._motivo_error(exc)))
                else:
                    self.filas.append(fila)
                    identificados += 1
                    self._redibujar_tabla()
                completado[idx] = True
                procesados += 1
                self.txt_estado.value = f"Analizando estados de cuenta… {procesados}/{total}"
                self.page.update()

        await asyncio.gather(*(procesar_uno(i, r) for i, r in enumerate(rutas)))

        detenido = token.cancelado
        # Pendientes: los que no se completaron (no empezados + abortados a media).
        restantes = [r for k, r in enumerate(rutas) if not completado[k]]
        self._token_carga = None
        self._cargando = False
        self.btn_cargar.disabled = False
        self.btn_cargar_carpeta.disabled = False
        self.anillo.visible = False
        self.btn_detener_carga.visible = False

        if detenido and restantes:
            self._rutas_pendientes = restantes
            self.btn_reanudar_carga.visible = True
            self.btn_reanudar_carga.content = (
                f"Reanudar carga ({len(restantes)} pendientes)"
            )
        else:
            self._rutas_pendientes = []
            self.btn_reanudar_carga.visible = False

        if detenido and restantes:
            resumen = (f"Carga detenida: {identificados} cargado(s), "
                       f"{len(restantes)} pendiente(s) por procesar.")
        else:
            resumen = f"{identificados} de {procesados} archivo(s) cargado(s) en la tabla."
        if errores:
            resumen += " " + self._resumen_errores(errores)
        self.txt_estado.value = resumen
        self.txt_estado.color = NARANJA if (errores or (detenido and restantes)) else GRIS
        if errores:
            self.avisar(
                f"{len(errores)} archivo(s) no se pudieron leer (ver el detalle abajo).",
                NARANJA,
            )
        elif detenido and restantes:
            self.avisar(
                f"Carga detenida. Quedaron {len(restantes)} pendiente(s); usa "
                "'Reanudar carga' para continuar desde donde se quedó.",
                NARANJA,
            )
        # Si ya hay un reporte importado, concilia los nuevos registros.
        self._conciliar()
        self._actualizar_resumen()
        self.page.update()

    @staticmethod
    def _motivo_error(exc: Exception) -> str:
        """Traduce la excepción al cargar un archivo a un motivo breve y claro."""
        s = str(exc).lower()
        tipo = type(exc).__name__.lower()
        if "cannot identify image" in s or "unidentifiedimage" in tipo:
            return "imagen no válida o dañada/incompleta"
        if isinstance(exc, FileNotFoundError):
            return "no se encontró el archivo"
        if isinstance(exc, ValueError) and "extens" in s:
            return "tipo de archivo no admitido"
        if "tesseract" in s or "ocr" in tipo:
            return "se necesita OCR y no está disponible"
        if isinstance(exc, PermissionError):
            return "archivo en uso o sin permisos"
        return "no se pudo procesar"

    @staticmethod
    def _resumen_errores(errores: list[tuple[str, str]], max_nombres: int = 8) -> str:
        """Arma un texto conciso de los archivos con error, agrupados por motivo y
        mostrando solo los nombres (sin rutas ni volcados técnicos)."""
        por_motivo: dict[str, list[str]] = {}
        for nombre, motivo in errores:
            por_motivo.setdefault(motivo, []).append(nombre)
        partes = []
        for motivo, nombres in por_motivo.items():
            visibles = ", ".join(nombres[:max_nombres])
            if len(nombres) > max_nombres:
                visibles += f" y {len(nombres) - max_nombres} más"
            partes.append(f"{len(nombres)} con {motivo} ({visibles})")
        return "No se pudieron leer: " + " · ".join(partes) + "."

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
        """Exporta los registros guardados en el formato elegido: Bancomer ->
        carpeta con los TXT de alta de cuentas ; Banregio -> Excel de alta."""
        guardados = db.listar()
        if not guardados:
            self.avisar("No hay registros guardados para exportar.", ROJO)
            return

        if self.dd_formato.value == "Banregio":
            await self._exportar_alta_banregio(guardados)
        else:  # BancomerAlta
            await self._exportar_alta_bancomer(guardados)

    async def _exportar_alta_bancomer(self, guardados) -> None:
        """Genera la carpeta 'CUENTAS PARA ALTA BANCOMER - dd-mm-aaaa' con los
        TXT de alta de cuentas, separados por banco (Bancomer 012 vs otros) y,
        dentro de cada uno, por correo (con vs sin). El alta del portal Bancomer
        aplica solo a las cuentas 012; las que no tienen correo se separan para
        capturarlo antes de subirlas."""
        validos = [b for b in guardados if validar_clabe(b.clabe)]
        if not validos:
            self.avisar("No hay registros con CLABE válida para exportar.", ROJO)
            return
        # El monto va dentro del TXT: no se exporta si algún registro no lo tiene.
        sin_monto = sum(1 for b in validos if b.monto is None)
        if sin_monto:
            self.avisar(
                f"No se puede exportar: {sin_monto} registro(s) sin monto. "
                "Captura el monto y guárdalo.",
                ROJO,
            )
            return

        def es_bancomer(b) -> bool:
            return b.clabe[:3] == "012"

        def tiene_correo(b) -> bool:
            return bool((b.email or "").strip())

        # Bancomer (012) usa su formato; otros bancos usan el formato SPEI.
        gen_bancomer = exportador_alta_bancomer.generar_txt
        gen_spei = exportador_alta_spei.generar_txt
        # (nombre de archivo, filtro, etiqueta, generador del TXT)
        grupos = [
            ("Cuentas Bancomer con correo.txt",
             lambda b: es_bancomer(b) and tiene_correo(b), "Bancomer con correo", gen_bancomer),
            ("Cuentas Bancomer sin correo.txt",
             lambda b: es_bancomer(b) and not tiene_correo(b), "Bancomer sin correo", gen_bancomer),
            ("Cuentas otros bancos con correo.txt",
             lambda b: not es_bancomer(b) and tiene_correo(b), "otros bancos con correo", gen_spei),
            ("Cuentas otros bancos sin correo.txt",
             lambda b: not es_bancomer(b) and not tiene_correo(b), "otros bancos sin correo", gen_spei),
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
            for nombre_archivo, filtro, etiqueta, generador in grupos:
                sub = [
                    (b.clabe, b.monto, b.beneficiario, b.email or "")
                    for b in validos if filtro(b)
                ]
                if not sub:
                    continue
                with open(os.path.join(carpeta, nombre_archivo),
                          "w", encoding="latin-1", newline="") as fh:
                    fh.write(generador(sub))
                resumen.append(f"{len(sub)} {etiqueta}")
        except PermissionError:
            self.avisar(
                "No se pudo guardar: algún archivo de la carpeta está abierto. "
                "Ciérralo e intenta de nuevo.", ROJO)
            return
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.avisar(f"No se pudo generar la carpeta: {exc}", ROJO)
            return
        try:
            os.startfile(carpeta)  # abre la carpeta generada en el explorador
        except Exception:  # noqa: BLE001 — abrir es opcional
            pass
        self.avisar("Carpeta de alta generada (" + ", ".join(resumen) + ").", VERDE)

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
        except PermissionError:
            self.avisar(
                "No se pudo guardar: el archivo está abierto en Excel. Ciérralo e "
                "intenta de nuevo (o guarda con otro nombre).", ROJO)
            return
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
