"""Pantalla: Dispersión (No Pemex).

Ofrece los controles para operar el RPA del SIPP: una tarjeta de filtros (los
mismos del modal "Agregar Facturas/Solicitudes de Pago") y los botones para
iniciar/pausar/reanudar y detener la ejecución. Al iniciar, hace login en el
SIPP, selecciona empresa/sucursal de la sesión y, por cada combinación de
empresa × tipo de solicitud elegida, aplica los filtros, busca y descarga el
Excel del reporte.

Las credenciales de inicio de sesión se capturan en el menú "Configuración" de
la barra superior (ver ui/configuracion.py); aquí se leen desde ahí al arrancar.

El RPA corre en un bucle de asyncio en un hilo aparte (BucleRpa) para no
congelar la interfaz y para que Playwright pueda lanzar el navegador en Windows.
Los métodos _pausar_rpa / _reanudar_rpa quedan como puntos de conexión para
cuando exista el proceso de dispersión que se pueda pausar.
"""

from __future__ import annotations

import asyncio
import datetime
import os
import re
import unicodedata

import flet as ft
from playwright.async_api import TimeoutError as PlaywrightTimeoutError

from core import (
    conciliacion, cuentas_dispersion, exportador_devoluciones, preferencias,
    reporte_dispersion, reporte_dispersion_export, rutas, tipo_cambio,
)
from core.reporte_dispersion import FilaSolicitud
from core.rpa_sipp import (
    BucleRpa,
    ControlRpa,
    ErrorSipp,
    FiltrosSolicitudPago,
    RpaDetenido,
    SesionSipp,
    asegurar_navegador,
    necesita_navegador,
)
from ui.comun import (CENTRO, EMPRESAS, GRIS, ID_POR_EMPRESA, NARANJA,
                      NOMBRES_EMPRESAS, ROJO, ROJO_BOTON, VERDE, encabezado_col)

# Formato de fecha que pide el modal del SIPP.
_RE_FECHA = re.compile(r"^\d{2}/\d{2}/\d{4}$")

# Alineaciones para las celdas de la tabla.
_IZQ = ft.Alignment(-1, 0)
_DER = ft.Alignment(1, 0)

# Anchos de columna de la tabla de solicitudes (encabezado y celdas).
_W_CHK = 40
_W_FOLIO = 50
_W_TIPO = 40
_W_FOLIO_FAC = 80
_W_PROV = 300
_W_CTA = 280
_W_FECHA = 100
_W_TSOL = 150
_W_MONTO = 100
_W_MONEDA = 70
_W_PRODUCTO = 220

# Ancho aproximado de un carácter a size=12 (px). Sirve para decidir, sin medir
# el render, si el texto de una celda probablemente se recorta y por tanto amerita
# tooltip. Se toma bajo a propósito (conservador): ante la duda, se deja tooltip.
_PX_POR_CHAR = 6.0

# Anchos de cada columna, en orden, para estimar si la tabla desborda a lo ancho.
_ANCHOS_COLS = [
    _W_CHK, _W_FOLIO, _W_TIPO, _W_FOLIO_FAC, _W_PROV, _W_CTA, _W_FECHA, _W_FECHA,
    _W_TSOL, _W_MONTO, _W_MONTO, _W_MONTO, _W_MONEDA, _W_PRODUCTO,
]
# Separación entre columnas y margen lateral del DataTable. Se mantienen chicos
# para respetar los anchos definidos y aprovechar el espacio (pantallas chicas).
_COL_SPACING = 8
_MARGEN_H = 6
# Ancho total aproximado de la tabla (columnas + separación + márgenes).
_ANCHO_TABLA = (
    sum(_ANCHOS_COLS) + _COL_SPACING * (len(_ANCHOS_COLS) - 1) + _MARGEN_H * 2
)
# Alto de cada fila de datos de la tabla.
_ALTO_FILA = 44
# Alto fijo del panel de filtros (ya no es un acordeón colapsable). Se ajusta
# para que el contenido (combos, campos, botones y estado en dos líneas) quepa
# sin scroll interno en una ventana normal. Si el contenido no cabe (p. ej.
# ventana muy angosta), el panel hace scroll interno; la pantalla completa
# también tiene scroll para recorrer panel + tabla.
_ALTO_PANEL_FILTROS = 370

# Paginación de la tabla ("lazy load"): se renderiza un máximo APROXIMADO de
# filas por página, respetando los grupos por cuenta (no se parte un grupo, para
# que su fila TOTAL siga cuadrando). Es el enfoque viable en Flet, cuyo DataTable
# no virtualiza: construye TODAS las filas que se le asignen. Subir este número
# muestra más de una vez pero cuesta más rendimiento.
_FILAS_POR_PAGINA = 100

# Texto de ayuda del ícono de interrogante junto a "Solicitudes a pagar".
_AYUDA_SOLICITUDES = (
    "Se generarán las dispersiones en SIPP, el proceso de pago en el banco NO se verá intervenido."
)

# Texto de ayuda del ícono de interrogante junto a "Buscar solicitudes de pago".
_AYUDA_BUSCAR = (
    'Busca las solicitudes de pago pendientes en la sección "Dispersiones (no '
    'Pemex)" del "Dashboard Tesorería".\n\n'
    "Búsqueda por reportes:\n"
    "- Permite subir uno o más reportes de solicitudes de pago (no Pemex).\n"
    "- Se buscarán y agregarán las facturas que se encuentren en el reporte y en "
    "SIPP.\n"
    "- El proceso de búsqueda será automático."
)

# --- Formato de moneda ---------------------------------------------------
def _fmt_moneda(valor: float | None) -> str:
    """Formatea un monto como moneda con 2 decimales (p. ej. $1,234.50)."""
    return f"${(valor or 0):,.2f}"


# --- Colores de fila por tipo de solicitud -------------------------------
def _norm(texto: str) -> str:
    """Normaliza (minúsculas, sin acentos, espacios colapsados) para comparar
    tipos de solicitud sin depender de acentos/mayúsculas."""
    base = unicodedata.normalize("NFKD", str(texto or ""))
    base = "".join(c for c in base if not unicodedata.combining(c))
    return " ".join(base.lower().split())


# Tipo de solicitud -> color de la fila (hex).
_COLOR_TIPO = {
    _norm("Pago Facturas"): "#428bca",
    _norm("Pago Extraordinario"): "#f0ad4e",
    _norm("Pago Estadias"): "#5cb85c",
    _norm("Pago Extraordinario Facturas"): "#999999",
    _norm("Pago General de Fletes"): "#9f8aff",
}
# Rojo cuando Saldo Programado != Saldo Factura (tiene prioridad sobre el tipo).
_COLOR_DESCUADRE = "#d9534f"
# Opacidad del tinte de fila (para que el texto siga legible).
_OPACIDAD_FILA = 0.30

# Entradas de la leyenda (etiqueta, color).
_LEYENDA = [
    ("Pago Facturas", "#428bca"),
    ("Pago Extraordinario", "#f0ad4e"),
    ("Pago Estadías", "#5cb85c"),
    ("Pago Extraordinario Facturas", "#999999"),
    ("Pago General de Fletes", "#9f8aff"),
    ("Saldo prog. ≠ Saldo factura", _COLOR_DESCUADRE),
]


def _color_fila(f: FilaSolicitud) -> str | None:
    """Color de fondo de la fila: rojo si Saldo Programado y Saldo Factura no
    coinciden; si coinciden, el color del tipo de solicitud (o None si no aplica)."""
    if round((f.saldo_factura or 0) - (f.saldo_programado or 0), 2) != 0:
        base = _COLOR_DESCUADRE
    else:
        base = _COLOR_TIPO.get(_norm(f.tipo_solicitud))
    return ft.Colors.with_opacity(_OPACIDAD_FILA, base) if base else None


def _fmt_fecha(d: datetime.date) -> str:
    return d.strftime("%d/%m/%Y")


def _parse_fecha(texto) -> "datetime.date | None":
    """Parsea 'DD/MM/AAAA' a date; None si está vacío o no es una fecha válida.
    Los reportes traen la Fecha Vencimiento en ese formato de texto."""
    s = str(texto or "").strip()
    if not s:
        return None
    try:
        return datetime.datetime.strptime(s, "%d/%m/%Y").date()
    except ValueError:
        return None


def _clave_empresa_moneda(f: FilaSolicitud) -> str:
    """Clave de agrupación de una fila: empresa + tipo de moneda (p. ej.
    'Abastecedora - MXN'). Las dispersiones se separan por empresa Y por moneda;
    si la moneda viene vacía, la clave es solo la empresa. La moneda ya llega
    normalizada del lector (MN -> MXN, sin puntos)."""
    empresa = f.empresa or "(Sin empresa)"
    moneda = (f.moneda or "").strip()
    return f"{empresa} - {moneda}" if moneda else empresa


def _nombre_empresa_limpio(emp) -> str:
    """Nombre corto de la empresa SIN el sufijo de moneda (para el filtro del modal
    y la selección de sesión). Se toma de los movimientos (traen el nombre limpio);
    respaldo: la clave 'Empresa - Moneda' recortando el sufijo."""
    if emp.movimientos and emp.movimientos[0].empresa:
        return emp.movimientos[0].empresa
    return (emp.empresa or "").rsplit(" - ", 1)[0].strip()


def _rango_fechas_vencimiento(emp) -> tuple[str, str]:
    """(fecha_inicio, fecha_fin) en DD/MM/AAAA: la fecha de VENCIMIENTO más ANTIGUA y
    la más RECIENTE de los movimientos de la empresa (se usan como Fecha Inicio /
    Fecha Fin del filtro del modal). ('', '') si no hay fechas parseables."""
    fechas = [d for d in (_parse_fecha(m.fecha_vencimiento) for m in emp.movimientos)
              if d is not None]
    if not fechas:
        return "", ""
    return min(fechas).strftime("%d/%m/%Y"), max(fechas).strftime("%d/%m/%Y")


def _sanear_archivo(nombre: str) -> str:
    """Nombre de archivo válido en Windows (quita caracteres no permitidos)."""
    limpio = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(nombre or "")).strip().strip(".")
    return limpio or "archivo"


def _ruta_unica(ruta: str) -> str:
    """Si `ruta` existe, agrega ' (n)' antes de la extensión hasta hallar un nombre
    libre (evita sobrescribir)."""
    if not os.path.exists(ruta):
        return ruta
    base, ext = os.path.splitext(ruta)
    n = 2
    while os.path.exists(f"{base} ({n}){ext}"):
        n += 1
    return f"{base} ({n}){ext}"


def _pares_proveedor_cuenta(emp) -> list[tuple[str, str]]:
    """Pares (proveedor, cuenta_bancaria) DISTINTOS de los movimientos, en orden de
    aparición. Cada par corresponde a una fila de la tabla de pagos del SIPP donde
    se capturan concepto/referencia."""
    vistos: set[tuple[str, str]] = set()
    pares: list[tuple[str, str]] = []
    for m in emp.movimientos:
        clave = (m.proveedor or "", m.cuenta_bancaria or "")
        if clave in vistos:
            continue
        vistos.add(clave)
        pares.append(clave)
    return pares


def _fecha_valida(texto: str) -> bool:
    """True si `texto` es una fecha real con formato DD/MM/AAAA."""
    if not _RE_FECHA.match(texto):
        return False
    try:
        datetime.datetime.strptime(texto, "%d/%m/%Y")
        return True
    except ValueError:
        return False


def _label_requerido(texto: str) -> ft.Text:
    """Etiqueta de campo requerido: el texto seguido de un asterisco ROJO (para
    que resalte). Se usa como `label` (Control) de los inputs obligatorios."""
    return ft.Text(
        spans=[
            ft.TextSpan(texto + " "),
            ft.TextSpan("*", ft.TextStyle(color=ROJO, weight=ft.FontWeight.BOLD)),
        ],
        size=12,
    )


class _Multiseleccion:
    """Combo de multiselección con 'chips': se elige una opción del desplegable
    y se agrega como etiqueta (chip) con una 'x' para quitarla.

    El modal del SIPP usa selects de selección única; aquí permitimos elegir
    varias opciones y el RPA itera por cada una. `valores()` devuelve las
    seleccionadas (en el orden en que se agregaron)."""

    def __init__(self, etiqueta: "str | ft.Control", opciones: list[str], page):
        self.page = page
        self.etiqueta = etiqueta
        self._opciones = list(opciones)
        self._seleccion: list[str] = []

        # Combo: solo muestra las opciones aún no elegidas; al elegir, se agrega.
        # Sin ancho fijo: llena la columna donde se coloque (ResponsiveRow).
        self.dd = ft.Dropdown(
            label=etiqueta, enable_filter=True, editable=True,
            options=[ft.dropdown.Option(key=o, text=o) for o in self._opciones],
            on_select=self._agregar,
        )
        # Caja con los chips de lo seleccionado (se ajusta en varias líneas).
        self._chips = ft.Row(wrap=True, spacing=6, run_spacing=6)
        self._caja = ft.Container(
            content=self._chips,
            border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT),
            border_radius=8,
            padding=ft.Padding.symmetric(horizontal=10, vertical=8),
        )
        self._refrescar()
        # STRETCH: que el combo y la caja ocupen todo el ancho de la columna.
        self.control = ft.Column(
            [self.dd, self._caja], spacing=6,
            horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
        )

    def valores(self) -> list[str]:
        return list(self._seleccion)

    def establecer(self, valores: list[str]) -> None:
        """Fija la selección (solo con opciones válidas, sin duplicar) y refresca
        el combo y los chips. Se usa al cargar una selección guardada."""
        self._seleccion = []
        for v in valores:
            if v in self._opciones and v not in self._seleccion:
                self._seleccion.append(v)
        self._refrescar()
        try:
            self.control.update()
        except (RuntimeError, AssertionError):
            pass  # aún no montado en la página; se refleja al renderizar

    def _refrescar(self) -> None:
        """Reconstruye las opciones disponibles del combo y los chips."""
        self.dd.options = [
            ft.dropdown.Option(key=o, text=o)
            for o in self._opciones if o not in self._seleccion
        ]
        if self._seleccion:
            self._chips.controls = [self._chip(v) for v in self._seleccion]
        else:
            self._chips.controls = [
                ft.Text("Ninguna seleccionada", size=14, color=GRIS, margin=ft.Padding.symmetric(horizontal=10, vertical=9))
            ]

    def _chip(self, valor: str) -> ft.Chip:
        return ft.Chip(
            label=ft.Text(valor),
            on_delete=lambda _e, v=valor: self._quitar(v),
            delete_icon_tooltip="Quitar",
        )

    def _agregar(self, _e) -> None:
        valor = self.dd.value
        if valor and valor not in self._seleccion:
            self._seleccion.append(valor)
        self.dd.value = None  # deja el combo listo para elegir otra
        self._refrescar()
        self.control.update()  # dirigido a este combo, no a toda la página

    def _quitar(self, valor: str) -> None:
        if valor in self._seleccion:
            self._seleccion.remove(valor)
        self._refrescar()
        self.control.update()  # dirigido a este combo, no a toda la página


class _TablaSolicitudes:
    """Tabla de solicitudes de UNA empresa.

    Agrupa las filas por cuenta bancaria y agrega, tras cada grupo, una fila
    'TOTAL PROGRAMADO' con la suma del saldo programado del grupo. Colorea cada
    fila (verde si Saldo Factura == Saldo Programado, rojo si difieren) y ofrece
    un check por fila con 'seleccionar todas' en el encabezado. Al agregar nuevos
    reportes evita duplicar filas ya presentes (por su 'clave')."""

    def __init__(self, page, empresa: str = "", cuentas=None,
                 fecha_venc_default: "datetime.date | None" = None,
                 moneda: str = "", clabes=None):
        self.page = page
        # Empresa (nombre corto) de esta tabla y sus cuentas de origen ya
        # filtradas por 'Alias corto' (se resuelven una vez, al crear la tabla).
        self.empresa = empresa
        # Moneda del grupo (p. ej. 'USD' o 'MXN'). En tablas USD se ofrece marcar
        # proveedores para 'pagar en pesos' (se les genera un TXT aparte en pesos).
        self.moneda = (moneda or "").strip().upper()
        # Proveedores marcados para pagar en pesos (solo aplica en tablas USD).
        self._pagar_pesos: set[str] = set()
        self._cuentas = list(cuentas or [])
        # Pares (cuenta, clabe) de la empresa: se MUESTRA la cuenta y se OPERA con la
        # CLABE (cuenta origen del TXT en pesos). Solo CLABEs válidas.
        self._clabes = list(clabes or [])
        self.cuenta_elegida = None  # CuentaBancaria elegida en el selector
        self.filas: list[FilaSolicitud] = []
        self._claves: set[tuple] = set()
        # Filtro de Fecha Vencimiento POR EMPRESA (date | None). Solo se muestran
        # las filas cuyo vencimiento sea <= a esta fecha; None = mostrar todas.
        # Arranca con el valor del filtro principal (default al crear la tabla).
        self._fecha_venc_filtro = fecha_venc_default
        self.chk_todos = ft.Checkbox(value=False, on_change=self._marcar_todas)
        # Selección PERSISTENTE entre páginas: se guarda por 'clave' de fila, no
        # en el checkbox (que se reconstruye al paginar). Así no se pierde lo
        # elegido en otras páginas. seleccionadas() lee de aquí.
        self._sel: set[tuple] = set()
        # Checkboxes de la página ACTUAL (para 'seleccionar todas' sin rebuild).
        self._checks_filas: list[ft.Checkbox] = []
        # Paginación a nivel de grupos por cuenta: página actual y el reparto de
        # cuentas por página (se recalcula en _reconstruir).
        self._pagina = 0
        self._paginas: list[list[str]] = [[]]
        self.tabla = ft.DataTable(
            columns=[
                ft.DataColumn(label=ft.Container(self.chk_todos, width=_W_CHK, alignment=CENTRO)),
                ft.DataColumn(label=encabezado_col("Folio", _W_FOLIO), ),
                ft.DataColumn(label=encabezado_col("Folio Factura", _W_FOLIO_FAC)),
                ft.DataColumn(label=encabezado_col("Tipo Solicitud", _W_TSOL)),
                ft.DataColumn(label=encabezado_col("Proveedor", _W_PROV)),
                ft.DataColumn(label=encabezado_col("Cuenta Bancaria", _W_CTA)),
                ft.DataColumn(label=encabezado_col("Total Fact.", _W_MONTO),numeric=True),
                ft.DataColumn(label=encabezado_col("Saldo Fact.", _W_MONTO),numeric=True),
                ft.DataColumn(label=encabezado_col("Saldo Prog.", _W_MONTO),numeric=True),
                ft.DataColumn(label=encabezado_col("Tipo", _W_TIPO)),
                ft.DataColumn(label=encabezado_col("Fh. Fact.", _W_FECHA)),
                ft.DataColumn(label=encabezado_col("Fh. Ven.", _W_FECHA)),
                ft.DataColumn(label=encabezado_col("Moneda", _W_MONEDA)),
                ft.DataColumn(label=encabezado_col("Producto", _W_PRODUCTO)),
            ],
            rows=[],
            column_spacing=_COL_SPACING,
            horizontal_margin=_MARGEN_H,  # sin el margen por defecto (24) que empujaba el check
            heading_row_color=ft.Colors.SURFACE_CONTAINER_HIGHEST,
            heading_row_height=46,
            data_row_min_height=_ALTO_FILA,
            data_row_max_height=_ALTO_FILA,
            divider_thickness=1,
            border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT),
            border_radius=10,
            vertical_lines=ft.BorderSide(
                1, ft.Colors.with_opacity(0.4, ft.Colors.OUTLINE_VARIANT)),
        )
        # Paginador (solo visible con más de una página).
        self._btn_prev = ft.IconButton(
            icon=ft.Icons.CHEVRON_LEFT, tooltip="Página anterior",
            on_click=lambda _e: self._ir_a_pagina(self._pagina - 1),
        )
        self._btn_next = ft.IconButton(
            icon=ft.Icons.CHEVRON_RIGHT, tooltip="Página siguiente",
            on_click=lambda _e: self._ir_a_pagina(self._pagina + 1),
        )
        self._lbl_pagina = ft.Text("", size=12, color=GRIS)
        self._pager = ft.Row(
            [self._btn_prev, self._lbl_pagina, self._btn_next],
            spacing=6, tight=True, visible=False,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        # --- Selector de Fecha Vencimiento POR EMPRESA (mismo patrón que el filtro
        # principal). Cambiarlo re-filtra ESTA tabla; vaciarlo muestra todas.
        self.dp_venc = ft.DatePicker(
            value=fecha_venc_default,
            first_date=datetime.date(2020, 1, 1),
            last_date=datetime.date(2035, 12, 31),
            help_text="Mostrar vencimientos hasta esta fecha",
            on_change=self._cambio_fecha_venc,
        )
        self.btn_limpiar_venc = ft.IconButton(
            icon=ft.Icons.CLOSE, icon_size=16,
            visible=fecha_venc_default is not None,
            tooltip="Quitar el filtro de vencimiento (mostrar todas)",
            on_click=self._limpiar_fecha_venc,
            width=24, height=24, padding=0,
            style=ft.ButtonStyle(padding=0),
        )
        # Los 4 inputs de la fila SIN 'dense' ni 'height' explícitos: así usan la
        # altura estándar de Material, que es idéntica entre TextField y Dropdown
        # (con 'dense'/'height' el Dropdown no queda a la misma altura que los
        # TextField y los bordes inferiores no coinciden).
        self.tf_venc = ft.TextField(
            label="Fecha Vencimiento", hint_text="DD/MM/AAAA", read_only=True,
            width=200,
            value=_fmt_fecha(fecha_venc_default) if fecha_venc_default else "",
            suffix=ft.Row(
                [self.btn_limpiar_venc, ft.Icon(ft.Icons.CALENDAR_MONTH, size=18)],
                spacing=4, tight=True,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            on_click=lambda e: self.page.show_dialog(self.dp_venc),
        )
        # --- Selector de CUENTA de origen (requerido) filtrado a esta empresa,
        # más Concepto y Referencia de pago (opcionales). Van en la misma línea
        # que el filtro de vencimiento. Cada 'cuenta' es el valor 'Cuenta' del
        # catálogo de dispersión: es lo que se muestra y por lo que busca el RPA.
        self.dd_cuenta = ft.Dropdown(
            label=_label_requerido("Cuenta Bancaria Origen"),
            width=340, enable_filter=True, editable=True,
            options=[ft.dropdown.Option(key=c, text=c) for c in self._cuentas],
            on_select=self._elegir_cuenta,
        )
        # Selector de la CLABE interbancaria de origen para el pago en pesos (solo
        # se muestra en tablas USD, junto al check; mismo estilo que dd_cuenta).
        self.clabe_elegida = None
        self.dd_clabe_origen = ft.Dropdown(
            label="Cuenta Origen (pago en pesos)",
            width=340, enable_filter=True, editable=True,
            # Se muestra la cuenta (banco/empresa) pero el valor (key) es la CLABE.
            options=[ft.dropdown.Option(key=cl, text=cta) for cta, cl in self._clabes],
            on_select=self._elegir_clabe,
        )
        self.tf_concepto = ft.TextField(
            label="Concepto de Pago", width=200)
        self.tf_referencia = ft.TextField(
            label="Referencia de Pago", width=200)
        self._filtro_row = ft.Row(
            [self.tf_venc, self.dd_cuenta, self.tf_concepto, self.tf_referencia],
            spacing=8, wrap=True, vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        # La tabla solo necesita scroll HORIZONTAL: su alto es natural (las filas
        # de la PÁGINA actual) y el scroll VERTICAL lo hace la pantalla completa
        # (ver _construir). El Row queda acotado al ancho del viewport (la columna
        # contenedora estira sus hijos) y desborda a lo ancho con su barra.
        self._tabla_scroll = ft.Row([self.tabla], scroll=ft.ScrollMode.AUTO)
        # Mensaje que sustituye a las filas cuando el filtro de Fecha Vencimiento
        # oculta TODAS las solicitudes. El DataTable de Flet no soporta 'colspan',
        # así que se muestra a todo lo ancho, debajo del encabezado (equivale a una
        # fila que abarca toda la tabla). Visible solo cuando no hay filas visibles.
        self._msg_vacio_filtro = ft.Container(
            content=ft.Text(
                "No hay solicitudes cuya fecha de vencimiento sea menor o igual a "
                "la seleccionada",
                size=12, color=GRIS, italic=True,
                text_align=ft.TextAlign.CENTER),
            visible=False, alignment=CENTRO,
            padding=ft.Padding.symmetric(vertical=16, horizontal=12),
            border=ft.Border(top=ft.BorderSide(1, ft.Colors.OUTLINE_VARIANT)),
        )
        # Barra de "Pagar en pesos" (solo en tablas USD): un check por proveedor.
        # Se llena en _reconstruir_pesos según los proveedores presentes.
        self._pesos_holder = ft.Container(visible=False)
        self.control = ft.Column(
            [self._filtro_row, self._pesos_holder, self._pager,
             self._tabla_scroll, self._msg_vacio_filtro],
            spacing=8)

    # ------------------------------------------- cuenta / concepto / referencia
    def _elegir_cuenta(self, _e=None) -> None:
        self.cuenta_elegida = self.dd_cuenta.value or None

    def set_cuentas(self, cuentas) -> None:
        """Reemplaza las cuentas del selector (p. ej. tras recargar el catálogo).
        Conserva la selección actual si la cuenta elegida sigue existiendo."""
        self._cuentas = list(cuentas or [])
        self.dd_cuenta.options = [
            ft.dropdown.Option(key=c, text=c) for c in self._cuentas
        ]
        if self.dd_cuenta.value not in self._cuentas:
            self.dd_cuenta.value = None
            self.cuenta_elegida = None
        try:
            self.dd_cuenta.update()
        except (RuntimeError, AssertionError):
            pass

    def cuenta_seleccionada(self):
        """Cuenta elegida en el selector, o None si no se ha elegido."""
        return self.cuenta_elegida

    def _elegir_clabe(self, _e=None) -> None:
        self.clabe_elegida = self.dd_clabe_origen.value or None

    def set_clabes(self, clabes) -> None:
        """Reemplaza los pares (cuenta, clabe) del selector de origen para pago en
        pesos (p. ej. tras recargar el catálogo). Se muestra la cuenta y se opera con
        la CLABE. Conserva la selección (por CLABE) si sigue existiendo."""
        self._clabes = list(clabes or [])
        self.dd_clabe_origen.options = [
            ft.dropdown.Option(key=cl, text=cta) for cta, cl in self._clabes
        ]
        claves = {cl for _cta, cl in self._clabes}
        if self.dd_clabe_origen.value not in claves:
            self.dd_clabe_origen.value = None
            self.clabe_elegida = None
        try:
            self.dd_clabe_origen.update()
        except (RuntimeError, AssertionError):
            pass

    def clabe_origen_pesos(self) -> str:
        """CLABE de origen elegida para el TXT en pesos ('' si no se eligió). Es el
        VALOR (key) del selector: la CLABE, aunque se muestre la cuenta."""
        return (self.clabe_elegida or "").strip()

    def cuenta_pesos_texto(self) -> str:
        """Texto (cuenta con banco/empresa) de la CLABE elegida en el selector de
        pago en pesos ('' si no se eligió). Sirve para determinar el banco/formato."""
        cl = self.clabe_elegida
        for cuenta, clabe in self._clabes:
            if clabe == cl:
                return cuenta
        return ""

    def concepto(self) -> str:
        return (self.tf_concepto.value or "").strip()

    def referencia(self) -> str:
        return (self.tf_referencia.value or "").strip()

    def agregar(self, nuevas: list[FilaSolicitud]) -> int:
        """Agrega filas evitando duplicados (por clave). Devuelve cuántas se
        agregaron realmente."""
        agregadas = 0
        for f in nuevas:
            clave = f.clave()
            if clave in self._claves:
                continue
            self._claves.add(clave)
            self.filas.append(f)
            agregadas += 1
        self._reconstruir()
        return agregadas

    # ------------------------------------------------ pagar en pesos (USD)
    def es_usd(self) -> bool:
        return self.moneda == "USD"

    def proveedores_pagar_pesos(self) -> set[str]:
        """Proveedores (de esta tabla USD) marcados para pagar en pesos. Solo se
        consideran los que además tienen alguna solicitud SELECCIONADA."""
        if not self.es_usd():
            return set()
        provs_sel = {f.proveedor for f in self.seleccionadas()}
        return {p for p in self._pagar_pesos if p in provs_sel}

    def _proveedores(self) -> list[str]:
        """Proveedores distintos presentes en la tabla, en orden de aparición."""
        vistos: list[str] = []
        for f in self.filas:
            if f.proveedor and f.proveedor not in vistos:
                vistos.append(f.proveedor)
        return vistos

    def _reconstruir_pesos(self) -> None:
        """(Re)arma la barra de 'Pagar en pesos' con un check por proveedor. Solo
        visible en tablas USD; oculta si no hay proveedores."""
        provs = self._proveedores() if self.es_usd() else []
        if not provs:
            self._pesos_holder.visible = False
            self._pesos_holder.content = None
            return
        # Limpia marcas de proveedores que ya no están.
        self._pagar_pesos &= set(provs)

        def _toggle(e, prov):
            (self._pagar_pesos.add if e.control.value
             else self._pagar_pesos.discard)(prov)
            self._actualizar_estado_clabe()

        checks = [
            ft.Checkbox(
                label=prov, value=prov in self._pagar_pesos,
                on_change=lambda e, p=prov: _toggle(e, p))
            for prov in provs
        ]
        # El selector de origen va a la DERECHA de los checks. Sin 'expand' (inflaba
        # el contenedor); los checks se envuelven en su propia fila a la izquierda.
        self._pesos_holder.content = ft.Container(
            ft.Column(
                [
                    ft.Text(
                        "Pagar en pesos (por proveedor): se genera un TXT aparte "
                        "en pesos al finalizar.", size=12, color=GRIS),
                    ft.Row(
                        [ft.Row(checks, wrap=True, spacing=16, run_spacing=6),
                         self.dd_clabe_origen],
                        spacing=16,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER),
                ],
                spacing=6, tight=True),
            padding=ft.Padding.symmetric(horizontal=8, vertical=6),
            bgcolor=ft.Colors.SURFACE_CONTAINER_HIGHEST, border_radius=8)
        self._pesos_holder.visible = True
        # Estado inicial del selector: activo solo si hay algún proveedor marcado.
        self._actualizar_estado_clabe(refrescar=False)

    def _actualizar_estado_clabe(self, refrescar: bool = True) -> None:
        """Activa el selector de origen solo si hay algún proveedor marcado para
        pagar en pesos; si no, lo deshabilita y lo vacía."""
        activo = bool(self._pagar_pesos)
        self.dd_clabe_origen.disabled = not activo
        if not activo:
            self.dd_clabe_origen.value = None
            self.clabe_elegida = None
        if refrescar:
            try:
                self.dd_clabe_origen.update()
            except (RuntimeError, AssertionError):
                pass

    def quitar(self, claves: set) -> int:
        """Quita las filas cuya clave esté en `claves` (p. ej. las ya dispersadas),
        limpia su selección y reconstruye. Devuelve cuántas quitó."""
        antes = len(self.filas)
        self.filas = [f for f in self.filas if f.clave() not in claves]
        self._claves = {f.clave() for f in self.filas}
        self._sel = {c for c in self._sel if c in self._claves}
        self._reconstruir()
        return antes - len(self.filas)

    def seleccionadas(self) -> list[FilaSolicitud]:
        """Movimientos (filas) seleccionados y VISIBLES (las ocultas por el filtro
        de vencimiento no se dispersan). La selección vive en self._sel por
        'clave' y persiste aunque una fila se oculte."""
        return [f for f in self._filas_visibles() if f.clave() in self._sel]

    def _filas_visibles(self) -> list[FilaSolicitud]:
        """Filas que cumplen el filtro de Fecha Vencimiento (<= al filtro por
        empresa). Si el filtro es None, son todas. Las que no cumplen se ocultan
        (no se borran: siguen en self.filas)."""
        lim = self._fecha_venc_filtro
        if lim is None:
            return self.filas
        visibles = []
        for f in self.filas:
            d = _parse_fecha(f.fecha_vencimiento)
            if d is not None and d <= lim:
                visibles.append(f)
        return visibles

    def _reconstruir(self) -> None:
        """Reconstruye SOLO las filas de la página actual (lazy load). Agrupa por
        cuenta y pagina sobre las filas VISIBLES (según el filtro de vencimiento);
        las ocultas no se renderizan ni cuentan en los totales."""
        visibles = self._filas_visibles()
        # Agrupa las filas visibles por cuenta bancaria, en orden de aparición.
        grupos: dict[str, list[FilaSolicitud]] = {}
        orden: list[str] = []
        for f in visibles:
            if f.cuenta_bancaria not in grupos:
                grupos[f.cuenta_bancaria] = []
                orden.append(f.cuenta_bancaria)
            grupos[f.cuenta_bancaria].append(f)
        # Reparte los grupos en páginas (~_FILAS_POR_PAGINA filas) y acota la
        # página actual al rango válido.
        self._paginas = self._calcular_paginas(orden, grupos)
        self._pagina = max(0, min(self._pagina, len(self._paginas) - 1))

        self._checks_filas = []
        renglones: list[ft.DataRow] = []
        for cuenta in self._paginas[self._pagina]:
            grupo = grupos[cuenta]
            total = sum((f.saldo_programado or 0) for f in grupo)
            for f in grupo:
                renglones.append(self._fila_datos(f))
            renglones.append(self._fila_total(total))
        self.tabla.rows = renglones
        # Si hay filas pero el filtro de vencimiento las oculta TODAS, se muestra el
        # mensaje a todo lo ancho (la tabla queda solo con su encabezado).
        self._msg_vacio_filtro.visible = bool(self.filas) and not visibles
        # 'Seleccionar todas' refleja las filas VISIBLES seleccionadas.
        self.chk_todos.value = bool(visibles) and all(
            f.clave() in self._sel for f in visibles)
        # Barra de 'pagar en pesos' (solo tablas USD).
        self._reconstruir_pesos()
        self._actualizar_pager(len(visibles))

    def _calcular_paginas(
        self, orden: list[str], grupos: dict[str, list[FilaSolicitud]]
    ) -> list[list[str]]:
        """Reparte las cuentas en páginas acumulando hasta ~_FILAS_POR_PAGINA
        filas por página, SIN partir un grupo (una cuenta que sola supere el tope
        queda en su propia página). Cada fila TOTAL cuenta como una fila más."""
        paginas: list[list[str]] = []
        actual: list[str] = []
        filas_actual = 0
        for cuenta in orden:
            n = len(grupos[cuenta]) + 1  # +1 por la fila TOTAL del grupo
            if actual and filas_actual + n > _FILAS_POR_PAGINA:
                paginas.append(actual)
                actual, filas_actual = [], 0
            actual.append(cuenta)
            filas_actual += n
        paginas.append(actual)
        return paginas or [[]]

    def _actualizar_pager(self, n_visibles: int) -> None:
        """Muestra/oculta el paginador y actualiza etiqueta y botones. `n_visibles`
        es el total de filas visibles (tras el filtro de vencimiento)."""
        total = len(self._paginas)
        self._pager.visible = total > 1
        if total > 1:
            self._lbl_pagina.value = (
                f"Página {self._pagina + 1} de {total}  ·  "
                f"{n_visibles} movimientos"
            )
            self._btn_prev.disabled = self._pagina <= 0
            self._btn_next.disabled = self._pagina >= total - 1

    def _ir_a_pagina(self, indice: int) -> None:
        """Cambia de página y repinta (update dirigido a esta tabla)."""
        indice = max(0, min(indice, len(self._paginas) - 1))
        if indice == self._pagina:
            return
        self._pagina = indice
        self._reconstruir()
        self._repintar()

    # -------------------------------------------- filtro de vencimiento
    def _cambio_fecha_venc(self, _e=None) -> None:
        """Aplica la fecha elegida en el calendario como filtro de esta tabla:
        vuelve a mostrar solo las filas con vencimiento <= a ella."""
        d = self.dp_venc.value
        self._fecha_venc_filtro = (
            d.date() if isinstance(d, datetime.datetime) else d)
        self.tf_venc.value = (
            _fmt_fecha(self._fecha_venc_filtro) if self._fecha_venc_filtro else "")
        self.btn_limpiar_venc.visible = self._fecha_venc_filtro is not None
        self._pagina = 0
        self._reconstruir()
        self._repintar()

    def _limpiar_fecha_venc(self, _e=None) -> None:
        """Quita el filtro de vencimiento de esta tabla: muestra TODAS sus filas."""
        self._fecha_venc_filtro = None
        self.dp_venc.value = None
        self.tf_venc.value = ""
        self.btn_limpiar_venc.visible = False
        self._pagina = 0
        self._reconstruir()
        self._repintar()

    def _repintar(self) -> None:
        """Update dirigido a esta tabla (silencioso si aún no está montada)."""
        try:
            self.control.update()
        except (RuntimeError, AssertionError):
            pass  # aún no montada; se reflejará al renderizar

    def _celda(self, texto: str, ancho: int, alineacion=CENTRO) -> ft.DataCell:
        """Celda de una línea: si el texto no cabe se recorta con '…' y el valor
        completo queda en el tooltip. Centrada salvo que se pida a la derecha.

        El tooltip solo se agrega cuando el texto es lo bastante largo como para
        recortarse (estimado por ancho); así se evita crear un MouseRegion por
        cada celda corta (montos, folios, fechas), lo que aligera el render."""
        texto = str(texto or "")
        if alineacion is _DER:
            align_txt = ft.TextAlign.RIGHT
        elif alineacion is _IZQ:
            align_txt = ft.TextAlign.LEFT
        else:
            align_txt = ft.TextAlign.CENTER
        # Tooltip solo si probablemente se recorta (texto más ancho que la celda).
        tip = texto if len(texto) * _PX_POR_CHAR > ancho else None
        return ft.DataCell(
            ft.Container(
                ft.Text(
                    texto, size=12, text_align=align_txt, width=ancho,
                    max_lines=1, no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS,
                ),
                width=ancho, alignment=alineacion, tooltip=tip,
            )
        )

    def _fila_datos(self, f: FilaSolicitud) -> ft.DataRow:
        # value inicial desde la selección persistente; al marcar/desmarcar se
        # actualiza self._sel (así sobrevive al cambio de página).
        chk = ft.Checkbox(value=f.clave() in self._sel)

        def _al_check(_e, f=f, c=chk):
            (self._sel.add if c.value else self._sel.discard)(f.clave())

        chk.on_change = _al_check
        self._checks_filas.append(chk)

        # Todas las celdas son TEXTO PLANO (no editables): 'Cuenta Bancaria' y
        # 'Saldo Programado' se muestran como los demás campos. Saldo Programado
        # usa el mismo formato de moneda y alineación derecha que Total/Saldo
        # Fact., para que la columna luzca uniforme con los demás totales.
        return ft.DataRow(
            color=_color_fila(f),
            cells=[
                ft.DataCell(ft.Container(chk, width=_W_CHK, alignment=CENTRO)),
                self._celda(f.folio, _W_FOLIO),
                self._celda(f.folio_factura, _W_FOLIO_FAC),
                self._celda(f.tipo_solicitud, _W_TSOL),
                self._celda(f.proveedor, _W_PROV),
                self._celda(f.cuenta_bancaria, _W_CTA),
                self._celda(_fmt_moneda(f.total_factura), _W_MONTO, _DER),
                self._celda(_fmt_moneda(f.saldo_factura), _W_MONTO, _DER),
                self._celda(_fmt_moneda(f.saldo_programado), _W_MONTO, _DER),
                self._celda(f.tipo, _W_TIPO),
                self._celda(f.fecha_factura, _W_FECHA),
                self._celda(f.fecha_vencimiento, _W_FECHA),
                self._celda(f.moneda, _W_MONEDA),
                self._celda(f.producto, _W_PRODUCTO, _IZQ),
            ],
        )

    def _fila_total(self, total: float) -> ft.DataRow:
        def vacia(ancho):
            return ft.DataCell(ft.Container(width=ancho))

        etiqueta = ft.DataCell(
            ft.Container(
                ft.Text("TOTAL PROGRAMADO", size=12, weight=ft.FontWeight.BOLD,
                        text_align=ft.TextAlign.RIGHT),
                width=_W_MONTO, alignment=_DER,
            )
        )
        # El total del grupo ya no se edita en vivo (Saldo Prog. es texto): se
        # calcula una vez y se muestra con el mismo formato de moneda.
        total_text = ft.Text(
            _fmt_moneda(total), size=12, weight=ft.FontWeight.BOLD,
            text_align=ft.TextAlign.RIGHT, width=_W_MONTO,
        )
        valor = ft.DataCell(ft.Container(total_text, width=_W_MONTO, alignment=_DER))
        # IMPORTANTE: este orden debe coincidir con el de las columnas y el de
        # _fila_datos (Check, Folio, Folio Factura, Tipo Solicitud, Proveedor,
        # Cuenta Bancaria, Total Fact., Saldo Fact., Saldo Prog., Tipo, Fh. Fact.,
        # Fh. Ven., Moneda, Producto). Si no, celdas vacías anchas ensanchan
        # columnas equivocadas.
        return ft.DataRow(
            color=ft.Colors.SURFACE_CONTAINER_HIGHEST,
            cells=[
                vacia(_W_CHK), vacia(_W_FOLIO), vacia(_W_FOLIO_FAC),
                vacia(_W_TSOL),   # Tipo Solicitud
                vacia(_W_PROV),   # Proveedor
                vacia(_W_CTA),    # Cuenta Bancaria
                vacia(_W_MONTO),  # Total Fact.
                etiqueta,         # bajo 'Saldo Fact.'
                valor,            # bajo 'Saldo Prog.'
                vacia(_W_TIPO),   # Tipo
                vacia(_W_FECHA),  # Fh. Fact.
                vacia(_W_FECHA),  # Fh. Ven.
                vacia(_W_MONEDA),    # Moneda
                vacia(_W_PRODUCTO),  # Producto
            ],
        )

    def _marcar_todas(self, _e) -> None:
        # Selecciona/deselecciona todas las filas VISIBLES (de todas las páginas,
        # respetando el filtro de vencimiento) y refleja el cambio en los
        # checkboxes de la página actual. Las ocultas no se seleccionan.
        claves_visibles = {f.clave() for f in self._filas_visibles()}
        if self.chk_todos.value:
            self._sel |= claves_visibles
        else:
            self._sel -= claves_visibles
        for chk in self._checks_filas:
            chk.value = self.chk_todos.value
        try:
            self.tabla.update()  # dirigido a esta tabla, no a toda la página
        except (RuntimeError, AssertionError):
            pass


class SeccionDispersionNoPemex:
    """Pestaña para operar el RPA de dispersión (No Pemex)."""

    # Empresa y plaza/sucursal con que se inicia la SESIÓN del SIPP (no es el
    # filtro de búsqueda, que el usuario elige abajo).
    EMPRESA_SESION = "Abastecedora"
    SUCURSAL_SESION = "Corporativo"

    # PRUEBAS: si es True, la operación de dispersión NO cierra el navegador al
    # terminar/detener, para poder inspeccionar el estado en el SIPP. Volver a
    # False al finalizar las pruebas.
    MANTENER_NAVEGADOR_PRUEBAS = True

    # Empresas disponibles para la dispersión (No Pemex). La fuente única vive en
    # ui.comun (se comparte con otras pantallas); aquí se referencian.
    EMPRESAS = EMPRESAS
    NOMBRES_EMPRESAS = NOMBRES_EMPRESAS
    ID_POR_EMPRESA = ID_POR_EMPRESA

    # Opciones fijas del combo "Tipo de Solicitud" (las del modal del SIPP).
    TIPOS_SOLICITUD = [
        "Pago Facturas",
        "Pago Extraordinario",
        "Pago Estadias",
        "Pago Extraordinario Facturas",
        "Pago General de Fletes",
    ]

    def __init__(self, app):
        self.app = app
        self.page = app.page
        # Estado de la ejecución: "detenido" | "ejecutando" | "pausado".
        self.estado = "detenido"
        self.sesion: SesionSipp | None = None
        self.bucle: BucleRpa | None = None
        # Control cooperativo de pausa/detención del flujo del RPA y el Future de
        # la corrida en curso (para cancelar la operación en el acto al detener).
        self._ctrl: ControlRpa | None = None
        self._future_rpa = None
        # Rutas de los Excel de reporte descargados (para procesar después).
        self.rutas_reporte: list[str] = []
        # Cuántas combinaciones empresa × tipo se intentaron en la última corrida.
        self.combinaciones_intentadas = 0
        # Una tabla por empresa (acumula entre corridas del RPA, sin duplicar).
        self._tablas_por_empresa: dict[str, _TablaSolicitudes] = {}
        # Fechas (inicio, fin) usadas en la(s) búsqueda(s) por cada grupo
        # empresa+moneda. Se fusionan al agregar nuevas búsquedas (inicio = la más
        # antigua, fin = la más reciente) y se usan como filtro en la dispersión.
        self._fechas_por_grupo: dict[str, tuple] = {}
        # Empresa cuya tabla se muestra cuando hay tabs (más de una empresa).
        self._empresa_activa: str | None = None
        # Catálogo de 'Cuentas de dispersión' (se carga una vez; el selector por
        # empresa se filtra por el ID de la empresa). Si el usuario actualiza el
        # Excel en Configuración, se refresca con recargar_catalogo.
        self.catalogo_dispersion = cuentas_dispersion.CatalogoCuentasDispersion()
        # --- Operación "Generar Dispersión" (modal aparte del RPA de búsqueda de
        # arriba). Estado propio: detenido|ejecutando|pausado|completado|error.
        self._disp_estado_op = "detenido"
        self._disp_ctrl: ControlRpa | None = None   # control cooperativo pausa/detención
        self._disp_task = None                       # asyncio.Task (envoltura en el loop de la UI)
        self._disp_future = None                     # Future del flujo en el hilo del RPA (cancelable)
        self._disp_loop_ui = None                    # loop de la UI (para marshalar estatus desde el hilo RPA)
        self._conc_dispersion = None                 # payload conciliado a dispersar
        # Folios generados por cada dispersión guardada: [{folio, empresa,
        # cuenta_origen, monto}]. Alimenta la descarga de TXT y el resumen final.
        self._folios_dispersados: list[dict] = []
        # Resultado de la descarga de layouts (TXT) y su carpeta (para el resumen).
        self._disp_resultados_txt: list[dict] = []
        self._disp_carpeta_txt: str | None = None
        # 'Pagar en pesos': proveedores marcados por grupo y la CLABE de origen
        # elegida por grupo (se capturan al generar la dispersión), TXT en pesos
        # generados, tipo de cambio y posible error.
        self._pesos_por_grupo: dict[str, set] = {}
        self._clabe_pesos_por_grupo: dict[str, str] = {}
        self._cuenta_pesos_por_grupo: dict[str, str] = {}
        self._pesos_generados: list[dict] = []
        self._tipo_cambio: float | None = None
        self._pesos_error: str | None = None
        # Tipo de cambio de VISTA PREVIA (para mostrarlo en el modal 'Solicitudes a
        # dispersar' cuando hay proveedores USD marcados 'pagar en pesos'). Se
        # consulta al generar la dispersión; None si no aplica o no se pudo obtener.
        self._tc_preview: float | None = None
        self._tc_preview_error: str | None = None
        self.contenido = self._construir()
        self._construir_dialogo_dispersion()
        # Carga automática de los filtros guardados (Empresa / Tipo), si existen.
        self._cargar_preferencias_iniciales()

    # ------------------------------------------------------------ UI
    def _construir(self) -> ft.Control:
        """Un único panel 'Filtros de búsqueda' a todo el ancho, con los filtros
        en un grid de 12 columnas (estilo Bootstrap) y los botones al final."""
        # --- Filtros ---
        # Empresa / Tipo: multiselección (Empresa es requerida -> asterisco rojo).
        self.ms_empresa = _Multiseleccion(
            _label_requerido("Empresa"), self.NOMBRES_EMPRESAS, self.page)
        self.ms_tipo = _Multiseleccion(
            "Tipo de Solicitud", self.TIPOS_SOLICITUD, self.page)

        # Fechas: requeridas. Selectores tipo calendario (DatePicker), igual que en
        # Alta de beneficiarios: el campo es de solo lectura y abre el calendario
        # al hacer clic; el RPA lee su texto. Inicio = hace 20 días; Fin = hoy.
        hoy = datetime.date.today()
        inicio_defecto = hoy - datetime.timedelta(days=20)
        self.dp_fecha_ini = ft.DatePicker(
            value=inicio_defecto,
            first_date=datetime.date(2020, 1, 1),
            last_date=datetime.date(2035, 12, 31),
            help_text="Fecha Inicio",
            on_change=lambda e: self._fecha_elegida(
                self.tf_fecha_ini, self.dp_fecha_ini),
        )
        self.dp_fecha_fin = ft.DatePicker(
            value=hoy,
            first_date=datetime.date(2020, 1, 1),
            last_date=datetime.date(2035, 12, 31),
            help_text="Fecha Fin",
            on_change=lambda e: self._fecha_elegida(
                self.tf_fecha_fin, self.dp_fecha_fin),
        )
        # Fecha Vencimiento: filtro OPCIONAL (sin valor por defecto).
        self.dp_fecha_venc = ft.DatePicker(
            first_date=datetime.date(2020, 1, 1),
            last_date=datetime.date(2035, 12, 31),
            help_text="Fecha de vencimiento",
            on_change=lambda e: self._fecha_elegida(
                self.tf_fecha_venc, self.dp_fecha_venc),
        )
        self.tf_fecha_ini = ft.TextField(
            label=_label_requerido("Fecha Inicio"), value=_fmt_fecha(inicio_defecto),
            hint_text="DD/MM/AAAA", read_only=True,
            suffix_icon=ft.Icons.CALENDAR_MONTH,
            on_click=lambda e: self.page.show_dialog(self.dp_fecha_ini),
        )
        self.tf_fecha_fin = ft.TextField(
            label=_label_requerido("Fecha Fin"), value=_fmt_fecha(hoy),
            hint_text="DD/MM/AAAA", read_only=True,
            suffix_icon=ft.Icons.CALENDAR_MONTH,
            on_click=lambda e: self.page.show_dialog(self.dp_fecha_fin),
        )
        # Botón para limpiar la Fecha Vencimiento (es opcional): solo visible
        # cuando el campo tiene un valor, para poder buscar sin ese filtro. Se
        # constriñe su tamaño (width/height/padding) para que su área táctil por
        # defecto (~48px) no estire la altura del campo al aparecer.
        self.btn_limpiar_venc = ft.IconButton(
            icon=ft.Icons.CLOSE, icon_size=18, visible=False,
            tooltip="Quitar la fecha de vencimiento",
            on_click=self._limpiar_fecha_venc,
            width=20, height=20, padding=0,
            style=ft.ButtonStyle(padding=0),
        )
        self.tf_fecha_venc = ft.TextField(
            label="Fecha Vencimiento", hint_text="DD/MM/AAAA", read_only=True,
            suffix=ft.Row(
                [self.btn_limpiar_venc, ft.Icon(ft.Icons.CALENDAR_MONTH, size=18)],
                spacing=4, tight=True,
                vertical_alignment=ft.CrossAxisAlignment.CENTER,
            ),
            on_click=lambda e: self.page.show_dialog(self.dp_fecha_venc),
        )
        self.tf_folio = ft.TextField(label="Folio Solicitud")  # opcional

        # Proveedor y Cuenta Bancaria (Proveedor): ocultos por ahora (no se usan);
        # se conservan definidos para reactivarlos cuando definamos su origen.
        self.dd_proveedor = ft.Dropdown(
            label="Proveedor", disabled=True, hint_text="Pendiente", visible=False,
        )
        self.dd_cuenta = ft.Dropdown(
            label="Cuenta Bancaria (Proveedor)", disabled=True,
            hint_text="Pendiente", visible=False,
        )

        # --- Botones de ejecución ---
        # Botón que intercala Iniciar / Pausar / Reanudar según el estado.
        self.btn_iniciar = ft.FilledButton(
            content="Iniciar", icon=ft.Icons.PLAY_ARROW, on_click=self._iniciar_pausar,
        )
        # Detener solo se habilita mientras el RPA está en marcha o en pausa.
        self.btn_detener = ft.OutlinedButton(
            content="Detener", icon=ft.Icons.STOP, on_click=self._detener,
            disabled=True,
        )
        # Subir un Excel de solicitudes ya filtrado (sin correr el RPA).
        # OCULTO temporalmente: la carga manual por reportes se deshabilita por
        # ahora (para reactivarla, poner visible=True).
        self.btn_subir = ft.OutlinedButton(
            content="Búsqueda por reportes", icon=ft.Icons.UPLOAD_FILE,
            on_click=self._subir_reporte, visible=False,
        )
        # Estado del RPA: "Estado:" (label) + el valor coloreado según el estado.
        # Ambos en negrita cursiva.
        self.lbl_estado = ft.Text(
            "Estatus del Robot:", size=13, weight=ft.FontWeight.BOLD, italic=True, color=GRIS)
        self.txt_estado = ft.Text(
            "Detenido", size=13, weight=ft.FontWeight.BOLD, italic=True, color=NARANJA)
        # Aviso + barra (indeterminada) para la descarga del navegador la 1ra vez.
        self.txt_install = ft.Text(
            "Instalando componentes extra, espere un momento…",
            size=13, color=GRIS, visible=False,
        )
        self.barra_install = ft.ProgressBar(visible=False)

        # --- Grid (12 columnas) ---
        def col(control, ancho_md):
            control.col = {"sm": 12, "md": ancho_md}
            return control

        fila_combos = ft.ResponsiveRow(
            [
                col(self._combo_guardable(self.ms_empresa, "empresas"), 6),
                col(self._combo_guardable(self.ms_tipo, "tipos"), 6),
            ],
            spacing=16, run_spacing=16,
            vertical_alignment=ft.CrossAxisAlignment.START,
        )
        # Fechas + Folio a col-md-3 (4 campos × 3 = fila completa de 12). Proveedor
        # y Cuenta Bancaria quedan ocultos por ahora, así que no van en el grid.
        fila_campos = ft.ResponsiveRow(
            [
                col(self.tf_fecha_ini, 3),
                col(self.tf_fecha_fin, 3),
                col(self.tf_fecha_venc, 3),
                col(self.tf_folio, 3),
            ],
            spacing=16, run_spacing=16,
            vertical_alignment=ft.CrossAxisAlignment.START,
        )
        # Leyenda "* Campo requerido." con el asterisco en ROJO (para que llame la
        # atención); el resto en gris. Va en la misma línea que los botones.
        leyenda_requerido = ft.Text(
            spans=[
                ft.TextSpan("* ", ft.TextStyle(color=ROJO, weight=ft.FontWeight.BOLD)),
                ft.TextSpan("Campo requerido.", ft.TextStyle(color=GRIS)),
            ],
            size=11,
        )
        # Parte inferior en DOS líneas: arriba los botones de ejecución
        # (Iniciar/Pausar + Detener) con la leyenda de requerido a la derecha, y
        # abajo el estado del RPA.
        fila_botones = ft.Row(
            [
                ft.Row(
                    [self.btn_iniciar, self.btn_detener],
                    spacing=10, tight=True,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                ),
                leyenda_requerido,
            ],
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
            wrap=True,
        )
        fila_estado = ft.Row(
            [self.lbl_estado, self.txt_estado],
            spacing=6, tight=True,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )

        # Panel de filtros de ALTURA FIJA (ya no es un acordeón). El título es
        # estático; el cuerpo hace scroll interno si no cabe en la altura fija.
        cuerpo_filtros = ft.Column(
            [
                fila_combos,
                fila_campos,
                self.txt_install,
                self.barra_install,
                fila_botones,
                fila_estado,
            ],
            spacing=14, scroll=ft.ScrollMode.AUTO, expand=True,
        )
        # Ícono de ayuda junto al título: tooltip inmediato al pasar el mouse
        # (wait_duration=0), sin necesidad de hacer click.
        icono_ayuda_buscar = ft.Icon(
            ft.Icons.HELP_OUTLINE, size=18, color=GRIS,
            tooltip=ft.Tooltip(
                message=_AYUDA_BUSCAR,
                wait_duration=ft.Duration(milliseconds=0),
            ),
        )
        # Encabezado del panel: título (con ayuda) a la izquierda y el botón de
        # búsqueda por reportes a la derecha.
        encabezado_panel = ft.Row(
            [
                ft.Row(
                    [
                        ft.Text("Buscar solicitudes de pago",
                                weight=ft.FontWeight.BOLD, size=15),
                        icono_ayuda_buscar,
                    ],
                    spacing=4, tight=True,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                ),
                self.btn_subir,
            ],
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        panel = ft.Card(
            content=ft.Container(
                content=ft.Column(
                    [encabezado_panel, cuerpo_filtros],
                    spacing=12, expand=True,
                ),
                padding=16,
                height=_ALTO_PANEL_FILTROS,
            ),
        )
        panel_tabla = self._construir_tabla()
        # Scroll de TODA la pantalla: el Column externo (expand) llena el área de
        # la sección y hace scroll vertical para recorrer panel + tabla. El panel
        # tiene altura fija y la tabla toma su alto natural (solo scroll horizontal).
        return ft.Column(
            [panel, panel_tabla],
            spacing=14, expand=True, scroll=ft.ScrollMode.AUTO,
            horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
        )

    # ----------------------------------------------------- tabla solicitudes
    def _construir_tabla(self) -> ft.Card:
        """Panel con la(s) tabla(s) de solicitudes. Con una sola empresa muestra
        una tabla; con varias, una tira de tabs (una por empresa). Empieza vacío;
        se llena tras el RPA con volcar_reportes()."""
        self.txt_tabla_vacia = ft.Text(
            "Aún no hay solicitudes para mostrar. Ejecuta una búsqueda.",
            size=12, color=GRIS,
        )
        # Indicador de carga mientras se leen los Excel.
        self._cargando = ft.Row(
            [
                ft.ProgressRing(width=18, height=18, stroke_width=2),
                ft.Text("Leyendo información de los reportes…", size=13, color=GRIS),
            ],
            spacing=10, visible=False,
        )
        # Tira de tabs (una por empresa); visible solo con más de una empresa.
        self._tira_holder = ft.Container(visible=False)
        # Columna persistente (expand): TODOS los controles viven aquí y solo se
        # alterna su 'visible' (más fiable que intercambiar 'content', que en Flet
        # no re-renderiza bien al reusar instancias). Las tablas se van agregando.
        self._contenedor_tablas = ft.Column(
            [self._cargando, self.txt_tabla_vacia, self._tira_holder],
            spacing=12,
        )
        # Botón que concilia la selección y (a futuro) dispara la dispersión.
        self.btn_dispersar = ft.FilledButton(
            content="Generar Dispersión",
            icon=ft.Icons.ACCOUNT_BALANCE_WALLET,
            on_click=self._generar_dispersiones,
        )
        # Botón para exportar a Excel las solicitudes seleccionadas (una hoja por
        # empresa), con el mismo layout que el reporte que se lee del SIPP.
        self.btn_reporte = ft.OutlinedButton(
            content="Generar Reporte",
            icon=ft.Icons.DESCRIPTION_OUTLINED,
            on_click=self._generar_reporte,
            style=ft.ButtonStyle(color=VERDE),  # verde: evoca un Excel
        )
        # Botón para vaciar TODAS las tablas de dispersión (pide confirmación).
        # En rojo por ser una acción destructiva; oculto mientras no haya tablas.
        self.btn_limpiar_tablas = ft.OutlinedButton(
            content="Eliminar todo",
            icon=ft.Icons.DELETE_OUTLINE,
            on_click=self._confirmar_eliminar_todo,
            style=ft.ButtonStyle(color=ROJO_BOTON),
            visible=False,
        )
        # Ícono de ayuda junto al título: el tooltip aparece de inmediato al pasar
        # el mouse (wait_duration=0), sin necesidad de hacer click.
        icono_ayuda = ft.Icon(
            ft.Icons.HELP_OUTLINE, size=18, color=GRIS,
            tooltip=ft.Tooltip(
                message=_AYUDA_SOLICITUDES,
                wait_duration=ft.Duration(milliseconds=0),
            ),
        )
        encabezado_tabla = ft.Row(
            [
                ft.Row(
                    [
                        ft.Text("Solicitudes a pagar",
                                weight=ft.FontWeight.BOLD, size=15),
                        icono_ayuda,
                    ],
                    spacing=4, tight=True,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                ),
                ft.Row(
                    [self.btn_reporte, self.btn_limpiar_tablas, self.btn_dispersar],
                    spacing=10, tight=True, wrap=True,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                ),
            ],
            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        # Card de alto natural: crece con la tabla y el scroll vertical lo maneja
        # la pantalla completa. La tabla mantiene su propio scroll horizontal.
        cuerpo = ft.Column(
            [
                encabezado_tabla,
                self._leyenda(),
                self._contenedor_tablas,
            ],
            spacing=10,
        )
        return ft.Card(
            content=ft.Container(content=cuerpo, padding=16),
        )

    def _leyenda(self) -> ft.Control:
        """Leyenda de colores de fila: uno por tipo de solicitud y el rojo del
        descuadre (Saldo Programado != Saldo Factura)."""
        chips = []
        for etiqueta, color in _LEYENDA:
            chips.append(
                ft.Row(
                    [
                        ft.Container(
                            width=14, height=14,
                            bgcolor=ft.Colors.with_opacity(_OPACIDAD_FILA, color),
                            border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT),
                            border_radius=3,
                        ),
                        ft.Text(etiqueta, size=11, color=GRIS),
                    ],
                    spacing=5, tight=True,
                )
            )
        return ft.Row(
            chips,
            wrap=True, spacing=14, run_spacing=6,
        )

    def _mostrar_cargando(self, visible: bool) -> None:
        """Muestra el indicador de carga (ocultando el resto) mientras se leen
        los reportes."""
        self._cargando.visible = visible
        if visible:
            self.txt_tabla_vacia.visible = False
            self._tira_holder.visible = False
            for tabla in self._tablas_por_empresa.values():
                tabla.control.visible = False
        # Update dirigido (solo esta pantalla), no toda la página.
        self._contenedor_tablas.update()

    async def _subir_reporte(self, _e=None) -> None:
        """Deja subir uno o varios Excel de solicitudes ya filtrados y los vuelca
        en la tabla (la empresa se toma de la celda C3 de cada archivo)."""
        archivos = await self.app.picker.pick_files(
            dialog_title="Selecciona el/los Excel de solicitudes",
            allowed_extensions=["xlsx", "xls"], allow_multiple=True,
        )
        if not archivos:
            return
        self._mostrar_cargando(True)
        rutas = [a.path for a in archivos]
        try:
            filas = await asyncio.to_thread(reporte_dispersion.leer_varios, rutas)
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self._reconstruir_tablas()  # oculta el "cargando"
            self.app.avisar(f"No se pudo leer el reporte: {exc}", ROJO)
            return
        if not filas:
            self._reconstruir_tablas()
            self.app.avisar(
                "El archivo no tiene solicitudes reconocibles (formato inesperado).",
                NARANJA)
            return
        self.volcar_reportes(filas)
        self.app.avisar(f"{len(filas)} solicitud(es) cargada(s) del reporte.", VERDE)

    # ------------------------------------------- generar dispersiones
    def _seleccion_por_empresa(self) -> dict[str, list[FilaSolicitud]]:
        """Movimientos seleccionados por empresa (solo empresas con selección).
        Mantiene la separación por empresa (una tabla por empresa)."""
        resultado: dict[str, list[FilaSolicitud]] = {}
        for empresa, tabla in self._tablas_por_empresa.items():
            seleccionadas = tabla.seleccionadas()
            if seleccionadas:
                resultado[empresa] = seleccionadas
        return resultado

    def _texto_tipo_solicitud(self) -> str:
        """Valor para el filtro 'Tipo Solicitud' del reporte: 'Todos' si están
        todos (o ninguno) seleccionados; si no, los tipos elegidos unidos por ', '."""
        tipos = self.ms_tipo.valores()
        if not tipos or set(tipos) == set(self.TIPOS_SOLICITUD):
            return "Todos"
        return ", ".join(tipos)

    async def _generar_reporte(self, _e=None) -> None:
        """Exporta a Excel las solicitudes SELECCIONADAS (una hoja por empresa),
        con el mismo formato del reporte que se lee del SIPP. Los datos de cada
        celda del bloque de filtros (B3:G6) se toman del filtro principal."""
        seleccion = self._seleccion_por_empresa()
        if not seleccion:
            self.app.avisar("Selecciona al menos un movimiento en la tabla.", NARANJA)
            return
        venc = (self.tf_fecha_venc.value or "").strip()
        folio = (self.tf_folio.value or "").strip()
        filtros = {
            "fecha_inicio": (self.tf_fecha_ini.value or "").strip(),
            "fecha_fin": (self.tf_fecha_fin.value or "").strip(),
            "fecha_vencimiento": venc or "N/A",
            "folio": folio or "Todos",
            "tipo_solicitud": self._texto_tipo_solicitud(),
        }
        ruta = await self.app.picker.save_file(
            dialog_title="Guardar reporte de solicitudes seleccionadas",
            file_name="Reporte Solicitudes Seleccionadas.xlsx",
            allowed_extensions=["xlsx"],
        )
        if not ruta:
            return
        if not ruta.lower().endswith(".xlsx"):
            ruta += ".xlsx"
        try:
            reporte_dispersion_export.generar(ruta, seleccion, filtros)
        except PermissionError:
            self.app.avisar(
                "No se pudo guardar: el archivo está abierto en Excel. Ciérralo e "
                "intenta de nuevo (o guarda con otro nombre).", ROJO)
            return
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.app.avisar(f"No se pudo generar el reporte: {exc}", ROJO)
            return
        n = sum(len(v) for v in seleccion.values())
        # Aviso con botón "Abrir" para abrir el reporte recién generado sin tener
        # que buscarlo en el explorador. Duración amplia para dar tiempo al clic.
        self.app.avisar(
            f"Reporte generado con {n} solicitud(es) en {len(seleccion)} hoja(s).",
            VERDE, accion="Abrir",
            on_accion=lambda _e=None: self._abrir_archivo(ruta),
            duracion=ft.Duration(seconds=12))

    def _abrir_archivo(self, ruta: str) -> None:
        """Abre un archivo en el programa predeterminado del sistema (Windows)."""
        try:
            os.startfile(ruta)  # noqa: S606 — abre en el visor predeterminado
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.app.avisar(f"No se pudo abrir el archivo: {exc}", ROJO)

    async def _generar_dispersiones(self, _e=None) -> None:
        """Valida la selección y la cuenta, concilia el payload y abre el diálogo
        de 'Generar Dispersión' desde el cual se inicia el RPA. El diálogo es modal
        y no se puede cerrar mientras la operación esté en curso."""
        seleccion = self._seleccion_por_empresa()
        if not seleccion:
            self.app.avisar("Selecciona al menos un movimiento en la tabla.", NARANJA)
            return
        # La Cuenta de origen es REQUERIDA para dispersar: cada tabla con
        # movimientos seleccionados debe tener una cuenta elegida.
        sin_cuenta = [
            grupo for grupo in seleccion
            if self._tablas_por_empresa[grupo].cuenta_seleccionada() is None
        ]
        if sin_cuenta:
            self.app.avisar(
                "Falta elegir la Cuenta en: " + ", ".join(sin_cuenta) + ".", NARANJA)
            return
        # Datos de pago por empresa (cuenta elegida + concepto + referencia) para
        # adjuntarlos al payload de la dispersión (los usará el RPA: la 'cuenta' es
        # el valor 'Cuenta' por el que se busca la cuenta en SIPP).
        datos_pago: dict[str, dict] = {}
        for grupo in seleccion:
            tabla = self._tablas_por_empresa[grupo]
            datos_pago[grupo] = {
                "cuenta": tabla.cuenta_seleccionada() or "",
                "concepto_pago": tabla.concepto(),
                "referencia_pago": tabla.referencia(),
            }
        # Proveedores marcados 'pagar en pesos' por grupo USD (se capturan ahora,
        # porque después de dispersar las tablas se vacían). Solo cuentan los que
        # tienen alguna solicitud seleccionada.
        self._pesos_por_grupo = {
            grupo: self._tablas_por_empresa[grupo].proveedores_pagar_pesos()
            for grupo in seleccion
            if self._tablas_por_empresa[grupo].proveedores_pagar_pesos()
        }
        # La CLABE de origen del pago en pesos es REQUERIDA si el grupo tiene
        # proveedores marcados (es la cuenta origen del TXT en pesos).
        sin_clabe = [
            grupo for grupo in self._pesos_por_grupo
            if not self._tablas_por_empresa[grupo].clabe_origen_pesos()
        ]
        if sin_clabe:
            self.app.avisar(
                "Falta elegir la CLABE de origen (pago en pesos) en: "
                + ", ".join(sin_clabe) + ".", NARANJA)
            return
        self._clabe_pesos_por_grupo = {
            grupo: self._tablas_por_empresa[grupo].clabe_origen_pesos()
            for grupo in self._pesos_por_grupo
        }
        # Texto de la cuenta origen elegida por grupo (para saber el banco/formato
        # del layout en pesos: Banregio vs Bancomer).
        self._cuenta_pesos_por_grupo = {
            grupo: self._tablas_por_empresa[grupo].cuenta_pesos_texto()
            for grupo in self._pesos_por_grupo
        }
        # Tipo de cambio (DOF) para mostrarlo en el modal 'Solicitudes a dispersar'
        # cuando hay proveedores USD marcados 'pagar en pesos'. Se consulta ahora
        # (cacheado por sesión) en un hilo para no congelar la UI; si falla, se
        # guarda el error para avisarlo en el modal (no impide dispersar).
        self._tc_preview = None
        self._tc_preview_error = None
        if self._pesos_por_grupo:
            try:
                self._tc_preview = await asyncio.to_thread(tipo_cambio.tipo_cambio_usd)
            except Exception as exc:  # noqa: BLE001 — se reporta en el modal
                self._tc_preview_error = str(exc)
        # Conciliación: separa por empresa, valida requeridos y cuadre por cuenta.
        # Las 'válidas' son las que irían al RPA. El payload queda guardado para la
        # operación (y para el botón 'Ver datos' del diálogo).
        self._conc_dispersion = conciliacion.conciliar(seleccion, datos_pago)
        self._abrir_dialogo_dispersion()

    # ---------------------------------------- diálogo "Generar Dispersión"
    def _construir_dialogo_dispersion(self) -> None:
        """Construye (una vez) el diálogo modal de 'Generar Dispersión': texto
        guía, botones de operación del RPA (Iniciar/Pausar/Detener), estatus y
        botones de cierre. No se puede cerrar mientras la operación esté en curso
        (el botón Cerrar se deshabilita y es modal, sin cierre por fuera)."""
        mensaje = ft.Text(
            "Se generarán las dispersiones de pago de las solicitudes "
            "seleccionadas en SIPP.\n"
            "Antes de continuar, revise bien la información que tomará el robot "
            "para generar las dispersiones.\n"
            "Presione 'Iniciar' para comenzar la operación.",
            size=13,
        )
        nota = ft.Text(
            "NOTA: No se podrá cerrar esta ventana hasta que se complete o se "
            "detenga la operación.",
            size=12, italic=True, color=NARANJA,
        )
        # Botón que intercala Iniciar / Pausar / Reanudar según el estado.
        self.btn_disp_iniciar = ft.FilledButton(
            content="Iniciar", icon=ft.Icons.PLAY_ARROW,
            on_click=self._disp_iniciar_pausar,
        )
        # Detener: solo habilitado mientras la operación corre o está en pausa.
        self.btn_disp_detener = ft.OutlinedButton(
            content="Detener", icon=ft.Icons.STOP, on_click=self._disp_detener,
            style=ft.ButtonStyle(color=ROJO_BOTON), disabled=True,
        )
        self.lbl_disp_estado = ft.Text(
            "Estatus del Robot:", size=13, weight=ft.FontWeight.BOLD,
            italic=True, color=GRIS)
        self.txt_disp_estado = ft.Text(
            "Detenido", size=13, weight=ft.FontWeight.BOLD, italic=True,
            color=NARANJA)
        # Revisar los datos que tomará el robot (payload conciliado). Se bloquea
        # mientras la operación corre para no apilar diálogos sobre el modal.
        self.btn_disp_ver = ft.TextButton(
            "Ver datos", icon=ft.Icons.VISIBILITY_OUTLINED,
            on_click=self._disp_ver_datos,
        )
        # Cerrar: se deshabilita mientras la operación corre (bloqueo del modal).
        self.btn_disp_cerrar = ft.TextButton("Cerrar", on_click=self._disp_cerrar)
        self._dlg_dispersion = ft.AlertDialog(
            modal=True,
            title=ft.Text("Generar dispersiones", weight=ft.FontWeight.BOLD),
            content=ft.Container(
                content=ft.Column(
                    [
                        mensaje, nota, ft.Divider(),
                        ft.Row([self.btn_disp_iniciar, self.btn_disp_detener],
                               spacing=10, tight=True),
                        ft.Row(
                            [self.lbl_disp_estado,
                             ft.Container(self.txt_disp_estado, expand=True)],
                            spacing=6,
                            vertical_alignment=ft.CrossAxisAlignment.START),
                    ],
                    tight=True, spacing=14,
                ),
                width=520,
            ),
            actions=[self.btn_disp_ver, self.btn_disp_cerrar],
            actions_alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
        )

    def _abrir_dialogo_dispersion(self) -> None:
        """Abre el diálogo en estado 'detenido' (listo para iniciar)."""
        self._disp_estado_op = "detenido"
        self._disp_fijar_estado("Detenido", NARANJA)
        self._disp_refrescar_controles()
        self.page.show_dialog(self._dlg_dispersion)
        self._disp_update()

    def _disp_fijar_estado(self, texto: str, color: str) -> None:
        self.txt_disp_estado.value = texto
        self.txt_disp_estado.color = color

    def _disp_update(self) -> None:
        """Refresca la página (para reflejar el diálogo); silencioso si aún no está
        montado."""
        try:
            self.page.update()
        except (RuntimeError, AssertionError):
            pass

    def _disp_refrescar_controles(self) -> None:
        """Ajusta botones y el bloqueo del cierre según el estado de la operación."""
        e = self._disp_estado_op
        corriendo = e in ("ejecutando", "pausado")
        if e == "ejecutando":
            self.btn_disp_iniciar.content = "Pausar"
            self.btn_disp_iniciar.icon = ft.Icons.PAUSE
            self.btn_disp_iniciar.disabled = False
        elif e == "pausado":
            self.btn_disp_iniciar.content = "Reanudar"
            self.btn_disp_iniciar.icon = ft.Icons.PLAY_ARROW
            self.btn_disp_iniciar.disabled = False
        elif e == "completado":
            # Ya se dispersó: no se re-ejecuta (evita duplicar); solo cerrar.
            self.btn_disp_iniciar.content = "Iniciar"
            self.btn_disp_iniciar.icon = ft.Icons.PLAY_ARROW
            self.btn_disp_iniciar.disabled = True
        else:  # detenido / error -> se puede (re)iniciar
            self.btn_disp_iniciar.content = "Iniciar"
            self.btn_disp_iniciar.icon = ft.Icons.PLAY_ARROW
            self.btn_disp_iniciar.disabled = False
        self.btn_disp_detener.disabled = not corriendo
        # BLOQUEO: no se puede cerrar ni ver datos mientras la operación corre.
        self.btn_disp_cerrar.disabled = corriendo
        self.btn_disp_ver.disabled = corriendo
        self._disp_update()

    async def _disp_iniciar_pausar(self, _e=None) -> None:
        """Arranca, pausa o reanuda la operación según el estado actual."""
        estado = self._disp_estado_op
        if estado in ("detenido", "error"):
            # Sin credenciales no tiene caso abrir el navegador (fallaría el login).
            usuario, contrasena = self.app.config.credenciales()
            if not usuario or not contrasena:
                self.app.avisar(
                    "Captura usuario y contraseña en Configuración.", ROJO)
                return
            self._disp_estado_op = "ejecutando"
            self._disp_fijar_estado("Iniciando…", VERDE)
            self._disp_refrescar_controles()
            self._disp_task = asyncio.create_task(self._ejecutar_dispersion())
            try:
                await self._disp_task
            except (RpaDetenido, asyncio.CancelledError):
                self._disp_estado_op = "detenido"
                self._disp_fijar_estado("Operación detenida.", ROJO)
            except Exception as exc:  # noqa: BLE001 — se reporta al usuario
                self._disp_estado_op = "error"
                self._disp_fijar_estado(f"Error: {exc}", ROJO)
                self.app.avisar(f"Falló la generación de dispersiones: {exc}", ROJO)
            else:
                self._disp_estado_op = "completado"
                self._disp_fijar_estado("Operación completada.", VERDE)
            finally:
                self._disp_task = None
                self._disp_future = None
                self._disp_ctrl = None
            self._disp_refrescar_controles()
            # En CUALQUIER desenlace, quitar de la tabla las combinaciones que sí se
            # dispersaron (evita re-dispersarlas si la operación se detuvo/falló a
            # media marcha; solo barre lo efectivamente guardado).
            barridas = self._eliminar_dispersadas()
            if self._disp_estado_op == "completado":
                self._disp_cerrar()   # el estado 'completado' permite cerrarlo
                self._mostrar_resumen_dispersion()
            elif barridas:
                # Se detuvo/falló a media marcha pero algunas sí se dispersaron:
                # se avisa que esas ya se quitaron de la tabla.
                self.app.avisar(
                    f"{barridas} combinación(es) sí se dispersó/dispersaron y se "
                    "quitó(aron) de la tabla.", NARANJA)
        elif estado == "ejecutando":
            self._disp_estado_op = "pausado"
            if self._disp_ctrl is not None:
                self._disp_ctrl.pausar()
            self._disp_fijar_estado("En pausa.", ft.Colors.AMBER_700)
            self._disp_refrescar_controles()
        elif estado == "pausado":
            self._disp_estado_op = "ejecutando"
            if self._disp_ctrl is not None:
                self._disp_ctrl.reanudar()
            self._disp_fijar_estado("En ejecución…", VERDE)
            self._disp_refrescar_controles()

    async def _disp_detener(self, _e=None) -> None:
        """Solicita detener: señala la detención cooperativa y cancela la operación
        en curso. El estado final 'detenido' lo fija _disp_iniciar_pausar al
        desenrollarse."""
        if self._disp_estado_op not in ("ejecutando", "pausado"):
            return
        if self._disp_ctrl is not None:
            self._disp_ctrl.detener()      # aborta en el próximo punto de control
        if self._disp_future is not None:
            self._disp_future.cancel()     # interrumpe el flujo en el hilo del RPA
        if self._disp_task is not None:
            self._disp_task.cancel()       # desenrolla la espera en el loop de la UI
        self._disp_fijar_estado("Deteniendo…", NARANJA)
        self.btn_disp_detener.disabled = True
        self._disp_update()

    def _disp_cerrar(self, _e=None) -> None:
        """Cierra el diálogo. No hace nada si la operación está en curso (el botón
        además está deshabilitado en ese caso)."""
        if self._disp_estado_op in ("ejecutando", "pausado"):
            return
        self.page.pop_dialog()

    def _disp_ver_datos(self, _e=None) -> None:
        """Muestra el detalle de lo que tomará el robot (payload conciliado)."""
        if self._conc_dispersion is not None:
            self._mostrar_datos_dispersion(self._conc_dispersion)

    async def _ejecutar_dispersion(self) -> None:
        """Operación REAL del RPA de dispersión: abre el navegador, inicia sesión y,
        por cada empresa+moneda VÁLIDA del payload conciliado, registra la dispersión
        en SIPP (buscar → seleccionar → aceptar → cuenta de origen → pagos → Guardar)
        y acumula el folio generado en self._folios_dispersados.

        Corre en el hilo del RPA (BucleRpa) con ControlRpa para pausa/detención
        cooperativa (mismo patrón que _arrancar_rpa); el estatus se marshala al loop
        de la UI porque Flet no es thread-safe."""
        self._disp_loop_ui = asyncio.get_running_loop()
        self._folios_dispersados = []
        self._disp_resultados_txt = []
        self._disp_carpeta_txt = None
        self._pesos_generados = []
        self._tipo_cambio = None
        self._pesos_error = None
        # (self._pesos_por_grupo y self._clabe_pesos_por_grupo se fijan en
        # _generar_dispersiones y se conservan para el TXT en pesos.)
        validas = self._conc_dispersion.validas if self._conc_dispersion else []
        if not validas:
            return
        if self.bucle is None:
            self.bucle = BucleRpa()
        self._disp_ctrl = ControlRpa(self.bucle._loop)
        ctrl = self._disp_ctrl
        self.sesion = SesionSipp(headless=False)
        sesion = self.sesion
        usuario, contrasena = self.app.config.credenciales()
        total = len(validas)
        # Fechas del filtro por combinación empresa+moneda, resueltas en el hilo de
        # la UI (el flujo corre en otro hilo): las guardadas de la búsqueda o, si no
        # hay, el rango por vencimiento como respaldo.
        fechas_por_grupo = {
            emp.empresa: self._fechas_dispersion(emp) for emp in validas
        }

        async def flujo() -> None:
            await sesion.iniciar()
            await sesion.login(usuario, contrasena)
            # La empresa/sucursal de SESIÓN se fija UNA vez; cada iteración filtra el
            # modal por su empresa (obj. 3), no cambia la sesión.
            await sesion.seleccionar_empresa_sucursal(
                self.EMPRESA_SESION, self.SUCURSAL_SESION)
            for i, emp in enumerate(validas, start=1):
                await ctrl.punto_control()  # pausa/detención entre empresas
                empresa = _nombre_empresa_limpio(emp)
                self._disp_estado_seguro(f"Dispersando {i}/{total}: {emp.empresa}…", VERDE)
                # 1) Entrar a 'Registrar Dispersión (No Pemex)' y abrir el modal.
                await sesion.ir_a_registrar_dispersion_no_pemex()
                await sesion.abrir_modal_agregar_solicitudes()
                # 2) Filtros: empresa de la iteración + las fechas guardadas de la
                #    búsqueda de ese grupo; Tipo Solicitud y Folio siempre vacíos.
                fecha_ini, fecha_fin = fechas_por_grupo.get(emp.empresa, ("", ""))
                await sesion.fijar_filtros(FiltrosSolicitudPago(
                    empresa=empresa, fecha_inicio=fecha_ini, fecha_fin=fecha_fin))
                # 3) Buscar, marcar las solicitudes elegidas y aceptar (obj. 4).
                await sesion.buscar_solicitudes()
                await sesion.seleccionar_solicitudes_por_folio(
                    [(m.folio, m.folio_factura) for m in emp.movimientos])
                await sesion.aceptar_solicitudes_dispersion()
                await ctrl.punto_control()
                # 4) Cuenta de origen (respaldo manual si no hay match, obj. 5).
                #    Devuelve el texto de la cuenta elegida (banco + cuenta) para el
                #    nombre del TXT y el resumen.
                cuenta_origen = await self._elegir_cuenta_origen(sesion, ctrl, emp.cuenta)
                # 5) Concepto/referencia por proveedor+cuenta (obj. 6/7).
                await sesion.llenar_pagos_proveedores(
                    _pares_proveedor_cuenta(emp),
                    emp.concepto_pago, emp.referencia_pago)
                # 6) Guardar y capturar el folio nuevo generado (obj. 8).
                folio = await sesion.guardar_dispersion()
                monto = sum((m.saldo_programado or 0) for m in emp.movimientos)
                moneda = emp.movimientos[0].moneda if emp.movimientos else ""
                # Se registra en cuanto guardar_dispersion REGRESA (la dispersión ya
                # quedó guardada en SIPP). 'clave' = empresa+moneda: identifica la
                # combinación dispersada para el barrido posterior de la tabla.
                self._folios_dispersados.append({
                    "folio": folio, "empresa": empresa, "clave": emp.empresa,
                    "moneda": moneda, "cuenta_origen": cuenta_origen or "",
                    "monto": monto})
            # 7) Descargar los TXT (layouts) de las dispersiones generadas.
            await self._descargar_txts_dispersion(sesion)
            # 8) Generar los TXT en PESOS (proveedores marcados en tablas USD).
            await asyncio.to_thread(self._generar_txts_pesos)
            return None

        self._disp_future = self.bucle.enviar(flujo())
        try:
            await asyncio.wrap_future(self._disp_future)
        finally:
            self._disp_future = None
            # En pruebas se deja el navegador abierto para inspeccionar el SIPP.
            if not self.MANTENER_NAVEGADOR_PRUEBAS:
                await self._detener_rpa()   # cierra el navegador (best-effort)

    async def _elegir_cuenta_origen(self, sesion, ctrl, cuenta: str) -> str:
        """Elige la cuenta de ORIGEN por coincidencia parcial de texto y devuelve el
        texto de la cuenta elegida (banco + cuenta). Si no hay coincidencia (obj. 5,
        último recurso), PAUSA el robot y espera a que el usuario elija la cuenta en
        el navegador y presione 'Reanudar'; al reanudar relee la cuenta y, si sigue
        vacía, vuelve a pausar."""
        if await sesion.seleccionar_cuenta_origen(cuenta):
            return (await sesion.cuenta_origen_valor()).strip()
        while True:
            self._disp_pausar_para_cuenta(cuenta)
            ctrl.pausar()
            await ctrl.punto_control()  # bloquea hasta Reanudar (o aborta al Detener)
            val = (await sesion.cuenta_origen_valor()).strip()
            if val:
                self._disp_estado_seguro("En ejecución…", VERDE)
                return val

    def _fechas_dispersion(self, emp) -> tuple[str, str]:
        """(fecha_inicio, fecha_fin) DD/MM/AAAA para el filtro del modal de esa
        combinación empresa+moneda: las fechas guardadas de la(s) búsqueda(s). Si el
        grupo no tiene fechas guardadas (p. ej. reporte cargado a mano), cae al
        rango por fecha de vencimiento como respaldo."""
        fi, ff = self._fechas_por_grupo.get(emp.empresa, (None, None))
        if fi is not None and ff is not None:
            return fi.strftime("%d/%m/%Y"), ff.strftime("%d/%m/%Y")
        return _rango_fechas_vencimiento(emp)

    async def _descargar_txts_dispersion(self, sesion) -> None:
        """Descarga los layouts (TXT) de las dispersiones generadas a una carpeta del
        día y guarda el resultado para el resumen. Resistente: si algo falla, deja el
        resultado vacío y la operación se considera terminada igual (las dispersiones
        ya quedaron guardadas)."""
        folios = [d for d in self._folios_dispersados if d.get("folio") is not None]
        if not folios:
            return
        self._disp_estado_seguro("Descargando archivos de dispersión…", VERDE)
        carpeta = self._carpeta_txt_dispersion()
        items = [{"folio": d["folio"], "empresa": d["empresa"],
                  "nombre": self._nombre_txt(d)} for d in folios]

        # Contador en el estatus del robot: así se ve el avance y, si algún layout
        # no genera descarga, se descarta rápido (timeout corto) sin parecer colgado.
        def _progreso(hechos, total):
            self._disp_estado_seguro(
                f"Descargando archivos de dispersión… {hechos}/{total}", VERDE)

        try:
            resultados = await sesion.descargar_layouts_dispersion(
                items, carpeta, progreso=_progreso)
        except Exception:  # noqa: BLE001 — la descarga no debe tumbar la operación
            resultados = []
        self._disp_carpeta_txt = carpeta
        self._disp_resultados_txt = resultados

    def _generar_txts_pesos(self) -> None:
        """Genera, en la carpeta de descargas, un TXT en PESOS por cada grupo USD
        dispersado con proveedores marcados 'pagar en pesos'. El importe = (total
        dispersado por proveedor en USD) × tipo de cambio del DOF (cacheado por
        sesión). El FORMATO del layout depende del banco de la cuenta de origen
        elegida (Banregio o BBVA/Bancomer). No tumba la operación si algo falla."""
        conc = self._conc_dispersion
        if conc is None or not self._pesos_por_grupo:
            return
        por_clave_folio = {
            d.get("clave"): d for d in self._folios_dispersados if d.get("clave")}
        por_clave_emp = {emp.empresa: emp for emp in conc.validas}
        # Grupos USD efectivamente dispersados y con proveedores marcados.
        pendientes = [
            clave for clave, provs in self._pesos_por_grupo.items()
            if provs and clave in por_clave_folio and clave in por_clave_emp
        ]
        if not pendientes:
            return
        try:
            tc = tipo_cambio.tipo_cambio_usd()
        except Exception as exc:  # noqa: BLE001 — se reporta en el resumen
            self._pesos_error = str(exc)
            return
        self._tipo_cambio = tc
        carpeta = self._disp_carpeta_txt or self._carpeta_txt_dispersion()
        os.makedirs(carpeta, exist_ok=True)
        for clave in pendientes:
            provs = self._pesos_por_grupo[clave]
            emp = por_clave_emp[clave]
            folio_entry = por_clave_folio[clave]
            folio = folio_entry.get("folio")
            # Un registro por proveedor marcado: (cuenta_prov, monto_pesos, nombre,
            # concepto). El importe se convierte USD -> MXN con el tipo de cambio.
            registros, total_pesos = [], 0.0
            for prov in provs:
                movs = [m for m in emp.movimientos if m.proveedor == prov]
                usd = sum((m.saldo_programado or 0) for m in movs)
                pesos = round(usd * tc, 2)
                cuenta_prov = re.sub(r"\D", "", movs[0].cuenta_bancaria) if movs else ""
                registros.append((cuenta_prov, pesos, prov, emp.concepto_pago))
                total_pesos += pesos
            if not registros:
                continue
            # CLABE de origen elegida en el selector del grupo (cuenta origen del
            # TXT en pesos); respaldo: los dígitos del texto de la cuenta de origen.
            clabe_sel = self._clabe_pesos_por_grupo.get(clave, "")
            clabe_origen = re.sub(
                r"\D", "", clabe_sel or folio_entry.get("cuenta_origen") or "")
            # El formato del layout depende del banco de la cuenta origen elegida.
            cuenta_texto = self._cuenta_pesos_por_grupo.get(clave, "")
            if exportador_devoluciones.banco_formato(cuenta_texto) == "banregio":
                # Banregio: separado por comas; usa la fecha (DDMMAAAA) de hoy.
                hoy = datetime.date.today().strftime("%d%m%Y")
                contenido = exportador_devoluciones.generar_banregio(registros, hoy)
            else:  # BBVA / Bancomer (ancho fijo) — formato por defecto
                contenido = exportador_devoluciones.generar_bancomer(
                    registros, clabe_origen, str(folio or ""))
            nombre = _sanear_archivo(self._nombre_txt(folio_entry) + " Pesos") + ".txt"
            ruta = _ruta_unica(os.path.join(carpeta, nombre))
            try:
                with open(ruta, "w", encoding="latin-1", newline="") as fh:
                    fh.write(contenido)
            except Exception:  # noqa: BLE001 — un TXT que falle no aborta el resto
                continue
            self._pesos_generados.append({
                "empresa": emp.empresa, "archivo": ruta,
                "total_pesos": total_pesos, "proveedores": len(registros)})

    @staticmethod
    def _carpeta_txt_dispersion() -> str:
        """Carpeta destino de los TXT (generados y descargados): una subcarpeta por
        día 'DD-MM-AAAA' dentro de 'Dispersiones (No Pemex)', bajo la carpeta de
        descargas (se usa '-' porque '/' no es válido en rutas Windows)."""
        hoy = datetime.date.today().strftime("%d-%m-%Y")
        return os.path.join(
            rutas.DATOS, "descargas", "Dispersiones (No Pemex)", hoy)

    @staticmethod
    def _nombre_txt(d: dict) -> str:
        """Nombre base del TXT: 'Folio Empresa Banco Cuenta' (el 'Banco Cuenta' es el
        texto de la cuenta de origen elegida)."""
        partes = [str(d.get("folio") or ""), d.get("empresa") or "",
                  d.get("cuenta_origen") or ""]
        return " ".join(p.strip() for p in partes if p and p.strip())

    def _disp_en_ui(self, fn) -> None:
        """Ejecuta `fn` en el loop de la UI. El flujo del RPA corre en otro hilo, así
        que las actualizaciones de Flet se marshalan con call_soon_threadsafe."""
        loop = self._disp_loop_ui
        if loop is not None and not loop.is_closed():
            loop.call_soon_threadsafe(fn)
        else:
            fn()

    def _disp_estado_seguro(self, texto: str, color: str) -> None:
        """Fija el estatus del modal de forma segura desde el hilo del RPA."""
        def aplicar():
            self._disp_fijar_estado(texto, color)
            self._disp_update()
        self._disp_en_ui(aplicar)

    def _disp_pausar_para_cuenta(self, cuenta: str) -> None:
        """Pone el modal en 'pausado' con el aviso para que el usuario elija la
        cuenta de origen en el navegador y presione 'Reanudar' (obj. 5)."""
        def aplicar():
            self._disp_estado_op = "pausado"
            self._disp_fijar_estado(
                f"No se encontró la cuenta de origen «{cuenta}». Selecciónala en el "
                "navegador y presiona «Reanudar».", ft.Colors.AMBER_700)
            self._disp_refrescar_controles()
        self._disp_en_ui(aplicar)

    @staticmethod
    def _celda_resumen(texto, ancho, derecha=False, bold=False):
        return ft.DataCell(ft.Container(
            ft.Text(str(texto or ""), size=12,
                    weight=ft.FontWeight.BOLD if bold else None,
                    text_align=ft.TextAlign.RIGHT if derecha else ft.TextAlign.LEFT,
                    max_lines=1, no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS,
                    width=ancho),
            width=ancho, alignment=_DER if derecha else _IZQ))

    def _tabla_resumen_mxn(self, folios: list[dict]) -> ft.Control:
        """Tabla de dispersiones en MXN: Folio | Empresa | Total MXN, con total."""
        c = self._celda_resumen
        W_FOL, W_EMP, W_MON = 70, 250, 150
        filas, total = [], 0.0
        for d in folios:
            total += d.get("monto") or 0
            filas.append(ft.DataRow(cells=[
                c(d.get("folio") if d.get("folio") is not None else "—", W_FOL),
                c(d.get("empresa"), W_EMP),
                c(_fmt_moneda(d.get("monto")), W_MON, derecha=True)]))
        filas.append(ft.DataRow(cells=[
            c("", W_FOL), c("TOTAL", W_EMP, bold=True),
            c(_fmt_moneda(total), W_MON, derecha=True, bold=True)]))
        return ft.DataTable(
            columns=[
                ft.DataColumn(label=encabezado_col("Folio", W_FOL)),
                ft.DataColumn(label=encabezado_col("Empresa", W_EMP)),
                ft.DataColumn(label=encabezado_col("Total MXN", W_MON), numeric=True)],
            rows=filas, column_spacing=10, horizontal_margin=8, heading_row_height=34,
            data_row_min_height=30, data_row_max_height=30,
            border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT), border_radius=8)

    def _tabla_resumen_usd(
        self, folios: list[dict], pesos_por_clave: dict) -> ft.Control:
        """Tabla de dispersiones en USD: Folio | Empresa | Total USD | Total MXN.
        El Total MXN es el importe convertido de los proveedores pagados en pesos de
        esa dispersión; 'N/A' si esa dispersión no tuvo pago en pesos. Con totales."""
        c = self._celda_resumen
        W_FOL, W_EMP, W_USD, W_MXN = 60, 190, 115, 115
        filas, tot_usd, tot_mxn = [], 0.0, 0.0
        for d in folios:
            tot_usd += d.get("monto") or 0
            mxn = pesos_por_clave.get(d.get("clave"))
            if mxn is not None:
                tot_mxn += mxn
            filas.append(ft.DataRow(cells=[
                c(d.get("folio") if d.get("folio") is not None else "—", W_FOL),
                c(d.get("empresa"), W_EMP),
                c(_fmt_moneda(d.get("monto")), W_USD, derecha=True),
                c(_fmt_moneda(mxn) if mxn is not None else "N/A", W_MXN, derecha=True)]))
        filas.append(ft.DataRow(cells=[
            c("", W_FOL), c("TOTAL", W_EMP, bold=True),
            c(_fmt_moneda(tot_usd), W_USD, derecha=True, bold=True),
            c(_fmt_moneda(tot_mxn) if tot_mxn else "N/A", W_MXN, derecha=True, bold=True)]))
        return ft.DataTable(
            columns=[
                ft.DataColumn(label=encabezado_col("Folio", W_FOL)),
                ft.DataColumn(label=encabezado_col("Empresa", W_EMP)),
                ft.DataColumn(label=encabezado_col("Total USD", W_USD), numeric=True),
                ft.DataColumn(label=encabezado_col("Total MXN", W_MXN), numeric=True)],
            rows=filas, column_spacing=10, horizontal_margin=8, heading_row_height=34,
            data_row_min_height=30, data_row_max_height=30,
            border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT), border_radius=8)

    def _mostrar_resumen_dispersion(self) -> None:
        """Resumen final: dos tablas (MXN y USD) que solo se muestran si tienen
        movimientos, con su total por moneda; el título de USD incluye el tipo de
        cambio cuando hubo pagos en pesos, y la tabla USD trae Total USD y Total MXN.
        Abajo, cuántos TXT se descargaron/generaron y un botón para abrir la carpeta."""
        folios = self._folios_dispersados
        resultados = self._disp_resultados_txt or []
        carpeta = self._disp_carpeta_txt
        pesos_gen = self._pesos_generados or []
        descargados = sum(1 for r in resultados if r.get("ok"))
        no_descargados = len(folios) - descargados
        hay_carpeta = bool(
            carpeta and os.path.isdir(carpeta) and (descargados or pesos_gen))

        # Importe en pesos (convertido) por combinación empresa+moneda dispersada.
        pesos_por_clave = {
            p.get("empresa"): p.get("total_pesos") for p in pesos_gen}
        usd = [d for d in folios if (d.get("moneda") or "").upper() == "USD"]
        mxn = [d for d in folios if (d.get("moneda") or "").upper() != "USD"]
        tc = self._tipo_cambio

        cuerpo: list[ft.Control] = [
            ft.Text(f"Se generaron {len(folios)} dispersión(es).", size=13)]
        if mxn:
            cuerpo.append(ft.Text("Dispersiones generadas en MXN", size=13,
                                  weight=ft.FontWeight.BOLD))
            cuerpo.append(self._tabla_resumen_mxn(mxn))
        if usd:
            titulo = "Dispersiones generadas en USD"
            if tc:
                titulo += f" (T.C. {_fmt_moneda(tc)} MXN)"
            cuerpo.append(ft.Text(titulo, size=13, weight=ft.FontWeight.BOLD))
            cuerpo.append(self._tabla_resumen_usd(usd, pesos_por_clave))
        cuerpo.append(ft.Divider())
        cuerpo.append(ft.Text(
            f"Archivos TXT descargados: {descargados}"
            + (f"   ·   No descargados: {no_descargados}"
               if no_descargados > 0 else ""),
            size=13, weight=ft.FontWeight.BOLD,
            color=VERDE if no_descargados == 0 else NARANJA))
        if pesos_gen:
            cuerpo.append(ft.Text(
                f"Archivos TXT en pesos generados: {len(pesos_gen)}",
                size=12, color=VERDE))
        elif self._pesos_error:
            cuerpo.append(ft.Text(
                f"No se generaron los TXT en pesos: {self._pesos_error}",
                size=12, color=ROJO))

        acciones: list[ft.Control] = []
        if hay_carpeta:
            acciones.append(ft.FilledButton(
                "Abrir carpeta", icon=ft.Icons.FOLDER_OPEN,
                on_click=lambda _e: self._abrir_archivo(carpeta)))
        acciones.append(ft.TextButton(
            "Cerrar", on_click=lambda _e: self.page.pop_dialog()))

        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Text("Resumen de la dispersión", weight=ft.FontWeight.BOLD),
            content=ft.Container(
                content=ft.Column(cuerpo, tight=True, spacing=10,
                                  scroll=ft.ScrollMode.AUTO),
                width=560, height=500),
            actions=acciones,
            actions_alignment=ft.MainAxisAlignment.SPACE_BETWEEN if hay_carpeta
            else ft.MainAxisAlignment.END,
        )
        self.page.show_dialog(dlg)

    def _eliminar_dispersadas(self) -> int:
        """Quita de las tablas las solicitudes de las combinaciones (empresa+moneda)
        que SÍ se dispersaron —las registradas en self._folios_dispersados—, para
        que no puedan volver a mandarse a dispersar por error. Funciona también si la
        operación se detuvo/falló a media marcha: solo barre lo efectivamente
        dispersado (una dispersión por combinación). Devuelve cuántas combinaciones
        barrió. Si una tabla queda vacía, se retira."""
        conc = self._conc_dispersion
        if conc is None or not self._folios_dispersados:
            return 0
        # Claves (empresa+moneda) efectivamente dispersadas y el emp de cada una
        # (para saber qué movimientos quitar de su tabla).
        dispersadas = {d.get("clave") for d in self._folios_dispersados if d.get("clave")}
        por_clave = {emp.empresa: emp for emp in conc.validas}
        barridas = 0
        vacias: list[str] = []
        for clave in dispersadas:
            emp = por_clave.get(clave)
            tabla = self._tablas_por_empresa.get(clave)
            if emp is None or tabla is None:
                continue
            tabla.quitar({m.clave() for m in emp.movimientos})
            barridas += 1
            if not tabla.filas:
                vacias.append(clave)
        # Retira del árbol las tablas que quedaron sin filas (y sus fechas).
        for grupo in vacias:
            self._fechas_por_grupo.pop(grupo, None)
            tabla = self._tablas_por_empresa.pop(grupo, None)
            if tabla is not None:
                try:
                    self._contenedor_tablas.controls.remove(tabla.control)
                except ValueError:
                    pass
        if self._empresa_activa not in self._tablas_por_empresa:
            self._empresa_activa = None
        self._reconstruir_tablas()
        return barridas

    def _mostrar_datos_dispersion(self, conc: "conciliacion.Conciliacion") -> None:
        """Muestra, de forma amigable (tablas, sin JSON), los datos que tomará el
        robot, separados por empresa + tipo de moneda (la misma separación de la
        pantalla). Pensado para usuarios no técnicos."""
        empresas = conc.empresas if conc else []
        if not empresas:
            cuerpo: ft.Control = ft.Text("No hay datos que mostrar.", size=12, color=GRIS)
        else:
            secciones: list[ft.Control] = []
            # Banner con el tipo de cambio: solo si hay proveedores USD marcados
            # 'pagar en pesos' (es el TC con que se convertirán a MXN).
            banner_tc = self._banner_tipo_cambio()
            if banner_tc is not None:
                secciones.append(banner_tc)
            for i, e in enumerate(empresas):
                if i:
                    secciones.append(ft.Divider())
                secciones.append(self._seccion_datos_empresa(e))
            cuerpo = ft.Column(
                secciones, scroll=ft.ScrollMode.AUTO, tight=True, spacing=14)
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Text("Solicitudes a dispersar", weight=ft.FontWeight.BOLD),
            content=ft.Container(content=cuerpo, width=860, height=580),
            actions=[
                ft.TextButton("Cerrar", on_click=lambda _e: self.page.pop_dialog()),
            ],
        )
        self.page.show_dialog(dlg)

    def _banner_tipo_cambio(self) -> ft.Control | None:
        """Banner con el tipo de cambio (DOF) que se usará para convertir a MXN los
        pagos de proveedores USD marcados 'pagar en pesos'. None si no hay ningún
        proveedor marcado (no aplica). Si el TC no se pudo obtener, lo avisa."""
        if not self._pesos_por_grupo:
            return None
        if self._tc_preview is not None:
            icono, color = ft.Icons.CURRENCY_EXCHANGE, VERDE
            texto = (f"Tipo de cambio (DOF): $ {_fmt_moneda(self._tc_preview)} MXN "
                     "por USD. Se usará para convertir a pesos los proveedores "
                     "marcados 'pagar en pesos'.")
        else:
            icono, color = ft.Icons.WARNING_AMBER, NARANJA
            detalle = f" ({self._tc_preview_error})" if self._tc_preview_error else ""
            texto = ("No se pudo obtener el tipo de cambio del DOF" + detalle
                     + ". El TXT en pesos se generará con el valor vigente al "
                     "momento de dispersar.")
        return ft.Container(
            content=ft.Row(
                [ft.Icon(icono, color=color, size=18),
                 ft.Text(texto, size=12, color=color, weight=ft.FontWeight.W_500,
                         expand=True)],
                spacing=8, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            padding=ft.Padding.symmetric(horizontal=12, vertical=8),
            border=ft.Border.all(1, color), border_radius=8,
        )

    def _seccion_datos_empresa(self, e: "conciliacion.EmpresaDispersion") -> ft.Control:
        """Bloque de una empresa+moneda: título con estado, datos de pago (cuenta
        origen / concepto / referencia), errores (si hay) y la tabla compacta de
        los movimientos a dispersar."""
        color_estado = VERDE if e.valida else NARANJA
        chip = ft.Container(
            ft.Text("Lista para dispersar" if e.valida else "Con observaciones",
                    size=10, color=color_estado, weight=ft.FontWeight.BOLD),
            padding=ft.Padding.symmetric(horizontal=8, vertical=2),
            border=ft.Border.all(1, color_estado), border_radius=10,
        )
        encabezado = ft.Row(
            [ft.Text(e.empresa, weight=ft.FontWeight.BOLD, size=14), chip],
            spacing=10, wrap=True, vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        info = ft.Row(
            [
                self._dato_compacto(
                    "Cuenta origen", e.cuenta or "—"),
                self._dato_compacto("Concepto", e.concepto_pago or "—"),
                self._dato_compacto("Referencia", e.referencia_pago or "—"),
            ],
            wrap=True, spacing=20, run_spacing=4,
        )
        hijos: list[ft.Control] = [encabezado, info]
        if e.errores:
            hijos.append(ft.Column(
                [ft.Text(f"• {m}", size=11, color=ROJO) for m in e.errores],
                spacing=1, tight=True))
        hijos.append(self._tabla_datos_movimientos(e))
        return ft.Column(hijos, spacing=8, tight=True)

    @staticmethod
    def _dato_compacto(etiqueta: str, valor: str) -> ft.Control:
        """Par 'Etiqueta: valor' en una línea compacta."""
        return ft.Row(
            [
                ft.Text(f"{etiqueta}:", size=11, color=GRIS,
                        weight=ft.FontWeight.BOLD),
                ft.Text(valor, size=11),
            ],
            spacing=4, tight=True,
        )

    def _tabla_datos_movimientos(
            self, e: "conciliacion.EmpresaDispersion") -> ft.Control:
        """Tabla compacta de los movimientos a dispersar de una empresa: folio,
        folio factura, proveedor, cuenta bancaria destino y los importes (total y
        saldo de factura + saldo programado), con sus totales."""
        W_FOLIO, W_FOLIO_FAC, W_PROV, W_CTA, W_MONTO = 60, 90, 200, 200, 105

        def celda(texto, ancho, derecha=False, bold=False):
            return ft.DataCell(ft.Container(
                ft.Text(
                    str(texto or ""), size=11,
                    weight=ft.FontWeight.BOLD if bold else None,
                    text_align=ft.TextAlign.RIGHT if derecha else ft.TextAlign.LEFT,
                    max_lines=1, no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS,
                    width=ancho),
                width=ancho, tooltip=str(texto or "") or None,
                alignment=_DER if derecha else _IZQ))

        filas: list[ft.DataRow] = []
        tot_prog = 0.0
        for m in e.movimientos:
            tot_prog += m.saldo_programado or 0
            filas.append(ft.DataRow(cells=[
                celda(m.folio, W_FOLIO),
                celda(m.folio_factura, W_FOLIO_FAC),
                celda(m.proveedor, W_PROV),
                celda(m.cuenta_bancaria, W_CTA),
                celda(_fmt_moneda(m.total_factura), W_MONTO, derecha=True),
                celda(_fmt_moneda(m.saldo_factura), W_MONTO, derecha=True),
                celda(_fmt_moneda(m.saldo_programado), W_MONTO, derecha=True),
            ]))
        # Solo se totaliza el Saldo Programado (la etiqueta va en la columna previa).
        filas.append(ft.DataRow(cells=[
            celda("", W_FOLIO), celda("", W_FOLIO_FAC), celda("", W_PROV),
            celda("", W_CTA), celda("", W_MONTO),
            celda("TOTAL PROGRAMADO", W_MONTO, derecha=True, bold=True),
            celda(_fmt_moneda(tot_prog), W_MONTO, derecha=True, bold=True),
        ]))
        tabla = ft.DataTable(
            columns=[
                ft.DataColumn(label=encabezado_col("Folio", W_FOLIO)),
                ft.DataColumn(label=encabezado_col("Folio Factura", W_FOLIO_FAC)),
                ft.DataColumn(label=encabezado_col("Proveedor", W_PROV)),
                ft.DataColumn(label=encabezado_col("Cuenta Bancaria", W_CTA)),
                ft.DataColumn(label=encabezado_col("Total Fact.", W_MONTO),
                              numeric=True),
                ft.DataColumn(label=encabezado_col("Saldo Fact.", W_MONTO),
                              numeric=True),
                ft.DataColumn(label=encabezado_col("Saldo Prog.", W_MONTO),
                              numeric=True),
            ],
            rows=filas,
            column_spacing=14, horizontal_margin=6,
            heading_row_height=34, data_row_min_height=30, data_row_max_height=30,
            divider_thickness=1,
            border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT), border_radius=8,
        )
        return ft.Row([tabla], scroll=ft.ScrollMode.AUTO)

    def _cuentas_de_empresa(self, nombre_empresa: str) -> list[str]:
        """Cuentas de dispersión de una empresa: se emparejan por el ID de la
        empresa (EMPRESAS/ID_POR_EMPRESA). [] si el nombre no tiene id o no hay
        cuentas cargadas para ese id."""
        id_empresa = self.ID_POR_EMPRESA.get(nombre_empresa)
        return self.catalogo_dispersion.cuentas_por_id_empresa(id_empresa)

    def _clabes_de_empresa(self, nombre_empresa: str) -> list[tuple[str, str]]:
        """Pares (cuenta, clabe) de una empresa para el selector de CLABE de origen:
        se MUESTRA la cuenta (banco/empresa) y se OPERA con la CLABE. Solo CLABEs
        válidas. [] si no hay."""
        id_empresa = self.ID_POR_EMPRESA.get(nombre_empresa)
        return self.catalogo_dispersion.cuentas_clabe_por_id_empresa(id_empresa)

    def recargar_catalogo(self) -> None:
        """Refresca el catálogo de cuentas de dispersión en caliente (tras subir un
        Excel nuevo en Configuración) y actualiza los selectores de cada tabla ya
        creada (cuenta origen y CLABE de pago en pesos)."""
        self.catalogo_dispersion = cuentas_dispersion.CatalogoCuentasDispersion()
        for tabla in self._tablas_por_empresa.values():
            tabla.set_cuentas(self._cuentas_de_empresa(tabla.empresa))
            tabla.set_clabes(self._clabes_de_empresa(tabla.empresa))

    def volcar_reportes(
        self, filas: list[FilaSolicitud],
        fecha_ini: str | None = None, fecha_fin: str | None = None,
    ) -> None:
        """Agrupa las filas por empresa y las agrega a la tabla de cada una
        (creándola si no existe), sin duplicar. Luego reconstruye la vista.

        `fecha_ini`/`fecha_fin` (DD/MM/AAAA), si se proveen, son las fechas de la
        búsqueda que produjo estas filas: se guardan/fusionan por grupo empresa+moneda
        (inicio = la más antigua, fin = la más reciente) para usarlas como filtro en
        la dispersión.

        La Fecha Vencimiento del filtro principal se toma como valor por DEFECTO
        del filtro por empresa, pero SOLO al crear la tabla: las tablas ya
        existentes conservan su propio filtro (el filtro principal no muta la
        tabla, ver objetivo 3.4)."""
        fecha_venc_default = _parse_fecha(self.tf_fecha_venc.value)
        fi, ff = _parse_fecha(fecha_ini), _parse_fecha(fecha_fin)
        # Se separa por empresa Y por tipo de moneda: cada combinación distinta
        # (p. ej. 'Abastecedora - MXN' y 'Abastecedora - USD') va en su propia
        # tabla/pestaña, igual que la separación por empresa pero con el sufijo de
        # la moneda.
        por_grupo: dict[str, list[FilaSolicitud]] = {}
        for f in filas:
            por_grupo.setdefault(_clave_empresa_moneda(f), []).append(f)
        for grupo, fs in por_grupo.items():
            tabla = self._tablas_por_empresa.get(grupo)
            if tabla is None:
                # La empresa (nombre corto) sale de las filas; sus cuentas se
                # resuelven UNA vez por su ID (no en cada cambio de tab).
                empresa_corta = fs[0].empresa if fs else ""
                tabla = _TablaSolicitudes(
                    self.page, empresa=empresa_corta,
                    cuentas=self._cuentas_de_empresa(empresa_corta),
                    fecha_venc_default=fecha_venc_default,
                    moneda=fs[0].moneda if fs else "",
                    clabes=self._clabes_de_empresa(empresa_corta))
                self._tablas_por_empresa[grupo] = tabla
                # Se agrega UNA vez al árbol; luego solo se alterna su 'visible'.
                self._contenedor_tablas.controls.append(tabla.control)
            tabla.agregar(fs)
            # Guarda/fusiona las fechas de la búsqueda para este grupo.
            self._fusionar_fechas_grupo(grupo, fi, ff)
        self._reconstruir_tablas()

    def _fusionar_fechas_grupo(self, grupo: str, fi, ff) -> None:
        """Fusiona las fechas de búsqueda (date | None) del grupo empresa+moneda:
        conserva la fecha inicio más ANTIGUA y la fecha fin más RECIENTE vistas."""
        if fi is None and ff is None:
            return
        act_fi, act_ff = self._fechas_por_grupo.get(grupo, (None, None))
        nueva_fi = min([d for d in (act_fi, fi) if d is not None], default=None)
        nueva_ff = max([d for d in (act_ff, ff) if d is not None], default=None)
        self._fechas_por_grupo[grupo] = (nueva_fi, nueva_ff)

    def _reconstruir_tablas(self) -> None:
        """Ajusta qué se ve: placeholder, una tabla o tabs (según cuántas
        empresas), la tira de tabs y el aviso de scroll horizontal."""
        self._cargando.visible = False
        empresas = list(self._tablas_por_empresa.keys())
        hay = bool(empresas)
        self.txt_tabla_vacia.visible = not empresas
        # Empresa activa: si la actual ya no existe (o no había), toma la primera.
        if hay and self._empresa_activa not in empresas:
            self._empresa_activa = empresas[0]
        # Tira de tabs: se muestra siempre que haya al menos una empresa (aunque
        # sea una sola), para que el usuario vea el botón 'Empresa - Moneda'.
        if hay:
            self._tira_holder.content = ft.Row(
                [self._boton_tab(emp) for emp in empresas], wrap=True, spacing=8)
        self._tira_holder.visible = hay
        # Visibilidad de cada tabla: solo la activa (con una sola, es esa misma).
        for empresa, tabla in self._tablas_por_empresa.items():
            tabla.control.visible = empresa == self._empresa_activa
        # 'Eliminar todo' solo tiene sentido con al menos una tabla.
        self.btn_limpiar_tablas.visible = bool(empresas)
        # Update dirigido (solo esta pantalla): no recorre las otras pestañas.
        self._contenedor_tablas.update()

    def _confirmar_eliminar_todo(self, _e=None) -> None:
        """Pide confirmación antes de vaciar todas las tablas de dispersión."""
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Text("Eliminar información"),
            content=ft.Text(
                "Se eliminará toda la información en las tablas de dispersión.\n"
                "¿Desea continuar?"),
            actions=[
                ft.TextButton(
                    "Cancelar", on_click=lambda _e: self.page.pop_dialog()),
                ft.FilledButton(
                    "Eliminar", color=ft.Colors.WHITE, bgcolor=ROJO,
                    on_click=self._eliminar_todo),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self.page.show_dialog(dlg)

    def _eliminar_todo(self, _e=None) -> None:
        """Vacía TODAS las tablas: quita sus controles del árbol, olvida las
        empresas y vuelve al estado inicial (placeholder). Cierra el diálogo."""
        self.page.pop_dialog()
        for tabla in self._tablas_por_empresa.values():
            try:
                self._contenedor_tablas.controls.remove(tabla.control)
            except ValueError:
                pass
        self._tablas_por_empresa.clear()
        self._fechas_por_grupo.clear()
        self._empresa_activa = None
        self._reconstruir_tablas()
        self.app.avisar("Se eliminó la información de las tablas.", VERDE)

    def _boton_tab(self, empresa: str) -> ft.Control:
        """Botón de la tira de tabs; el de la empresa activa va resaltado."""
        activo = empresa == self._empresa_activa
        fabrica = ft.FilledButton if activo else ft.OutlinedButton
        return fabrica(content=empresa,
                       on_click=lambda _e, m=empresa: self._cambiar_tab(m))

    def _cambiar_tab(self, empresa: str) -> None:
        self._empresa_activa = empresa
        self._reconstruir_tablas()

    # ------------------------------------------ guardar/cargar filtros
    def _combo_guardable(self, ms: "_Multiseleccion", clave: str) -> ft.Control:
        """Envuelve un combo de multiselección con una barra (guardar/cargar) que
        recuerda su selección entre sesiones, de forma independiente por combo."""
        guardar = ft.IconButton(
            icon=ft.Icons.SAVE_OUTLINED, icon_size=18,
            tooltip="Guardar esta selección",
            on_click=lambda _e: self._guardar_seleccion(ms, clave),
        )
        cargar = ft.IconButton(
            icon=ft.Icons.RESTORE, icon_size=18,
            tooltip="Cargar la selección guardada",
            on_click=lambda _e: self._cargar_seleccion(ms, clave),
        )
        barra = ft.Row(
            [guardar, cargar], spacing=0, tight=True,
            alignment=ft.MainAxisAlignment.END,
        )
        return ft.Column([barra, ms.control], spacing=0)

    def _guardar_seleccion(self, ms: "_Multiseleccion", clave: str) -> None:
        preferencias.guardar_lista(clave, ms.valores())
        self.app.avisar("Selección guardada.", VERDE)

    def _cargar_seleccion(self, ms: "_Multiseleccion", clave: str) -> None:
        valores = preferencias.cargar_lista(clave)
        if not valores:
            self.app.avisar("No hay una selección guardada aún.", NARANJA)
            return
        ms.establecer(valores)
        self.app.avisar("Selección cargada.", VERDE)

    def _cargar_preferencias_iniciales(self) -> None:
        """Aplica automáticamente al arrancar las selecciones guardadas de Empresa
        y Tipo de Solicitud (si existen), para no cargarlas a mano cada vez. Sin
        aviso (es transparente); si no hay nada guardado, no hace nada."""
        for ms, clave in ((self.ms_empresa, "empresas"), (self.ms_tipo, "tipos")):
            valores = preferencias.cargar_lista(clave)
            if valores:
                ms.establecer(valores)

    def _fecha_elegida(self, campo: ft.TextField, dp: ft.DatePicker) -> None:
        """Vuelca la fecha elegida en el calendario al campo, como DD/MM/AAAA."""
        if dp.value:
            campo.value = dp.value.strftime("%d/%m/%Y")
            # La Fecha Vencimiento es opcional: al fijarla, muestra el botón de
            # limpiar para poder quitarla luego.
            if campo is self.tf_fecha_venc:
                self.btn_limpiar_venc.visible = True
            self.page.update()

    def _limpiar_fecha_venc(self, _e=None) -> None:
        """Vacía la Fecha Vencimiento (filtro opcional) y oculta el botón limpiar."""
        self.tf_fecha_venc.value = ""
        self.dp_fecha_venc.value = None
        self.btn_limpiar_venc.visible = False
        self.page.update()

    def _validar_filtros(self) -> str:
        """Devuelve un mensaje de error si los filtros no son válidos; '' si OK."""
        if not self.ms_empresa.valores():
            return "Selecciona al menos una empresa."
        # El tipo de solicitud es OPCIONAL: si no se elige ninguno, se busca sin
        # filtrar por tipo (todas las solicitudes).
        fi = (self.tf_fecha_ini.value or "").strip()
        ff = (self.tf_fecha_fin.value or "").strip()
        if not _fecha_valida(fi):
            return "La Fecha Inicio debe tener formato DD/MM/AAAA válido."
        if not _fecha_valida(ff):
            return "La Fecha Fin debe tener formato DD/MM/AAAA válido."
        # Inicio no puede ser posterior a Fin. (Fecha Vencimiento es opcional y no
        # se valida: puede ir vacía para no filtrar por ella.)
        d_ini = datetime.datetime.strptime(fi, "%d/%m/%Y")
        d_fin = datetime.datetime.strptime(ff, "%d/%m/%Y")
        if d_ini > d_fin:
            return "La Fecha Inicio no puede ser mayor que la Fecha Fin."
        return ""

    # ----------------------------------------------------- ejecución
    async def _iniciar_pausar(self, _e) -> None:
        """Único botón que arranca, pausa o reanuda según el estado actual."""
        if self.estado == "detenido":
            usuario, contrasena = self.app.config.credenciales()
            if not usuario or not contrasena:
                self.app.avisar(
                    "Captura usuario y contraseña en Configuración.", ROJO)
                return
            error = self._validar_filtros()
            if error:
                self.app.avisar(error, ROJO)
                return
            # Pasa a 'ejecutando' y HABILITA pausa/detención DESDE YA. Antes el
            # botón quedaba deshabilitado toda la corrida, por lo que pausa y
            # detención eran inalcanzables. Toman efecto en el siguiente punto de
            # control del flujo (o cancelan la operación en curso, al Detener).
            self.estado = "ejecutando"
            self.btn_iniciar.content = "Pausar"
            self.btn_iniciar.icon = ft.Icons.PAUSE
            self.btn_iniciar.disabled = False
            self.btn_detener.disabled = False
            self._estado("Iniciando sesión en el SIPP…", VERDE)
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
            except (RpaDetenido, asyncio.CancelledError):
                # Detención pedida por el usuario: cierre limpio (no es un error).
                self._mostrar_instalacion(False)
                await self._detener_rpa()
                self.estado = "detenido"
                self.btn_iniciar.disabled = False
                self._refrescar_controles()
                self.app.avisar("RPA detenido.", ROJO)
                return
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
            # La operación del RPA terminó: cierra el navegador y vuelve a
            # 'detenido' (los XLSX quedaron en disco para la lectura posterior).
            await self._detener_rpa()
            self.estado = "detenido"
            self.btn_iniciar.disabled = False
            n = len(self.rutas_reporte)
            if n == 0:
                # Ninguna búsqueda trajo resultados: se termina y se notifica.
                self.app.avisar(
                    "No se encontraron resultados en ninguna de las búsquedas. "
                    "Se terminó la operación del RPA.", NARANJA)
            else:
                # Lee los XLSX descargados (en un hilo, para no congelar la UI) y
                # vuelca sus filas en la tabla, agrupadas por empresa y sin
                # duplicar respecto a corridas anteriores. Muestra un indicador de
                # carga mientras se leen para dejar claro que no se colgó.
                self._mostrar_cargando(True)
                filas = await asyncio.to_thread(
                    reporte_dispersion.leer_varios, self.rutas_reporte)
                # Guarda por grupo las fechas de ESTA búsqueda (inicio/fin del
                # filtro), para usarlas como filtro en la dispersión.
                self.volcar_reportes(
                    filas,
                    fecha_ini=(self.tf_fecha_ini.value or "").strip(),
                    fecha_fin=(self.tf_fecha_fin.value or "").strip())
                sin_datos = self.combinaciones_intentadas - n
                extra = f" ({sin_datos} sin resultados)" if sin_datos > 0 else ""
                self.app.avisar(
                    f"RPA completado: {n} reporte(s) (Excel) descargado(s){extra}.",
                    VERDE)
        elif self.estado == "ejecutando":
            self.estado = "pausado"
            await self._pausar_rpa()
        else:  # pausado
            self.estado = "ejecutando"
            await self._reanudar_rpa()
        self._refrescar_controles()

    async def _detener(self, _e) -> None:
        """Solicita detener el RPA. La detención es cooperativa: se marca la señal
        y se cancela la operación en curso. El cierre del navegador y el aviso
        final los hace _iniciar_pausar cuando el flujo se desenrolla, para no
        cerrar el navegador desde dos lados a la vez."""
        if self.estado == "detenido":
            return
        if self._ctrl is not None:
            self._ctrl.detener()   # aborta en el próximo punto de control
        if self._future_rpa is not None:
            self._future_rpa.cancel()  # interrumpe la operación en curso
        # Feedback inmediato; el estado final 'detenido' lo fija _iniciar_pausar.
        self.btn_iniciar.disabled = True
        self.btn_detener.disabled = True
        self._estado("Deteniendo…", NARANJA)
        self.page.update()

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

    def _estado(self, texto: str, color: str) -> None:
        """Fija el texto del estado del RPA y su color (sin actualizar la página;
        el llamador decide cuándo refrescar)."""
        self.txt_estado.value = texto
        self.txt_estado.color = color

    def _refrescar_controles(self) -> None:
        """Ajusta etiquetas, íconos y disponibilidad de los botones al estado."""
        if self.estado == "ejecutando":
            self.btn_iniciar.content = "Pausar"
            self.btn_iniciar.icon = ft.Icons.PAUSE
            self.btn_detener.disabled = False
            self._estado("En ejecución", VERDE)
        elif self.estado == "pausado":
            self.btn_iniciar.content = "Reanudar"
            self.btn_iniciar.icon = ft.Icons.PLAY_ARROW
            self.btn_detener.disabled = False
            self._estado("En pausa", ft.Colors.AMBER_700)
        else:  # detenido
            self.btn_iniciar.content = "Iniciar"
            self.btn_iniciar.icon = ft.Icons.PLAY_ARROW
            self.btn_detener.disabled = True
            self._estado("Detenido", NARANJA)
        self.page.update()

    # ------------------------------------------------- hooks del RPA
    async def _correr(self, coro):
        """Corre una corrutina del RPA en el bucle del hilo dedicado y espera su
        resultado sin congelar la interfaz."""
        if self.bucle is None:
            self.bucle = BucleRpa()
        return await asyncio.wrap_future(self.bucle.enviar(coro))

    def _estado_seguro(self, texto: str, color: str) -> None:
        """Fija el estatus del RPA de búsqueda de forma segura desde el hilo del RPA
        (marshala la actualización de Flet al loop de la UI)."""
        def aplicar():
            self._estado(texto, color)
            try:
                self.page.update()
            except (RuntimeError, AssertionError):
                pass
        loop = getattr(self, "_loop_ui", None)
        if loop is not None and not loop.is_closed():
            loop.call_soon_threadsafe(aplicar)
        else:
            aplicar()

    async def _arrancar_rpa(self, usuario: str, contrasena: str) -> None:
        """Abre el navegador, inicia sesión y, por cada empresa × tipo elegido,
        aplica filtros, busca y descarga el Excel del reporte."""
        # Loop de la UI (para marshalar el estatus desde el hilo del RPA).
        self._loop_ui = asyncio.get_running_loop()
        self.sesion = SesionSipp(headless=False)
        sesion = self.sesion

        # Lee los filtros en el hilo de la UI antes de pasar al hilo del RPA.
        empresas = self.ms_empresa.valores()
        # Tipo opcional: si no se elige ninguno, se hace UNA búsqueda por empresa
        # sin filtrar por tipo (None -> fijar_filtros omite ese filtro).
        tipos = self.ms_tipo.valores() or [None]
        fecha_ini = (self.tf_fecha_ini.value or "").strip()
        fecha_fin = (self.tf_fecha_fin.value or "").strip()
        folio = (self.tf_folio.value or "").strip() or None

        # Controlador de pausa/detención, atado al bucle del RPA. Se crea antes de
        # lanzar el flujo; el flujo lo consulta en sus puntos de control.
        if self.bucle is None:
            self.bucle = BucleRpa()
        self._ctrl = ControlRpa(self.bucle._loop)
        ctrl = self._ctrl

        total = len(empresas) * len(tipos)

        async def flujo() -> list[str]:
            self._estado_seguro("Iniciando sesión en el SIPP…", VERDE)
            await sesion.iniciar()
            await sesion.login(usuario, contrasena)
            self._estado_seguro("Seleccionando empresa y sucursal…", VERDE)
            await sesion.seleccionar_empresa_sucursal(
                self.EMPRESA_SESION, self.SUCURSAL_SESION)
            self._estado_seguro("Abriendo el buscador de solicitudes…", VERDE)
            await sesion.ir_a_registrar_dispersion_no_pemex()
            await sesion.abrir_modal_agregar_solicitudes()
            rutas: list[str] = []
            # SIPP filtra por una sola empresa/tipo a la vez: se itera cada
            # combinación reutilizando el mismo modal. Si una búsqueda no trae
            # resultados, se continúa con la siguiente.
            k = 0
            for empresa in empresas:
                await ctrl.punto_control()  # pausa/detención entre empresas
                for tipo in tipos:
                    await ctrl.punto_control()  # pausa/detención entre búsquedas
                    k += 1
                    prefijo = empresa if tipo is None else f"{empresa} - {tipo}"
                    self._estado_seguro(f"Buscando {k}/{total}: {prefijo}…", VERDE)
                    await sesion.fijar_filtros(
                        FiltrosSolicitudPago(
                            empresa=empresa,
                            fecha_inicio=fecha_ini,
                            fecha_fin=fecha_fin,
                            folio_solicitud=folio,
                            tipo_solicitud=tipo,
                        )
                    )
                    # Se separan Buscar y Descargar para dar retroalimentación con el
                    # número de solicitudes encontradas antes de bajar el Excel.
                    encontradas = await sesion.buscar_solicitudes()
                    if encontradas:
                        self._estado_seguro(
                            f"Descargando {k}/{total}: {prefijo} "
                            f"({encontradas} sol.)…", VERDE)
                        ruta = await sesion.descargar_reporte_excel(prefijo=prefijo)
                        if ruta:
                            rutas.append(ruta)
                    else:
                        self._estado_seguro(
                            f"Sin resultados {k}/{total}: {prefijo}", VERDE)
            self._estado_seguro("Procesando reportes descargados…", VERDE)
            return rutas

        # Se guarda el Future para poder cancelar la operación en curso al detener
        # (la pausa/detención cooperativa solo actúa en los puntos de control).
        self._future_rpa = self.bucle.enviar(flujo())
        try:
            self.rutas_reporte = await asyncio.wrap_future(self._future_rpa)
        finally:
            self._future_rpa = None
        # Búsquedas que sí descargaron vs total intentadas (para el aviso).
        self.combinaciones_intentadas = len(empresas) * len(tipos)

    async def _pausar_rpa(self) -> None:
        """Pide pausar: el flujo se detendrá en su próximo punto de control (tras
        terminar la búsqueda/descarga en curso)."""
        if self._ctrl is not None:
            self._ctrl.pausar()

    async def _reanudar_rpa(self) -> None:
        """Reanuda el flujo pausado."""
        if self._ctrl is not None:
            self._ctrl.reanudar()

    async def _detener_rpa(self) -> None:
        """Cierra el navegador y libera la sesión (best-effort)."""
        self._ctrl = None
        self._future_rpa = None
        if self.sesion is None:
            return
        sesion = self.sesion
        self.sesion = None
        try:
            await self._correr(sesion.cerrar())
        except Exception:  # noqa: BLE001 — el cierre no debe propagar errores
            pass
