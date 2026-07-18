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
import calendar
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
from ui.tabla_responsiva import (Cabecera, ColumnaTabla, FilaDatos,
                                 SegmentoCabecera, TablaResponsiva)
from ui.tabla_responsiva import DER as _TDER
from ui.tabla_responsiva import IZQ as _TIZQ

# Formato de fecha que pide el modal del SIPP.
_RE_FECHA = re.compile(r"^\d{2}/\d{2}/\d{4}$")

# Alineaciones para las celdas de la tabla.
_IZQ = ft.Alignment(-1, 0)
_DER = ft.Alignment(1, 0)

# Margen ("canalón") reservado para que las barras de scroll (que Flet dibuja
# ENCIMA del contenido) no se solapen con las tablas y obstruyan la información.
_GUTTER_SCROLL = 14

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
    "Para dispersar:\n"
    "1) Marca las solicitudes a pagar (usa el check del encabezado de cada grupo "
    "para seleccionar todo un proveedor de una vez).\n"
    "2) Elige la 'Cuenta Bancaria Origen' (obligatoria).\n"
    "3) Si vas a pagar en pesos a un proveedor en USD, márcalo y elige la CLABE de "
    "origen del pago en pesos.\n\n"
    "Se generarán las dispersiones en SIPP; el proceso de pago en el banco NO se "
    "verá intervenido."
)

# Texto de ayuda del ícono de interrogante junto a "Buscar solicitudes de pago".
_AYUDA_BUSCAR = (
    'Busca las solicitudes de pago pendientes en la sección "Dispersiones (no '
    'Pemex)" del "Dashboard Tesorería".'
)

# --- Formato de moneda ---------------------------------------------------
def _fmt_moneda(valor: float | None) -> str:
    """Formatea un monto como moneda con 2 decimales (p. ej. $1,234.50)."""
    return f"${(valor or 0):,.2f}"


def _fmt_tc(valor: float | None) -> str:
    """Formatea el tipo de cambio con 4 decimales (p. ej. $17.5993), tal como lo
    publica el DOF, para que coincida con el valor usado en el cálculo."""
    return f"${(valor or 0):,.4f}"


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


# Columnas de la tabla de solicitudes, definidas por PORCENTAJE del ancho de la
# tarjeta (la clase TablaResponsiva las convierte a px según el tamaño de ventana).
# La col 0 es el check (sin extractor). NO se muestran Proveedor/Cuenta (van en la
# banda de grupo) ni Moneda (va en la pestaña empresa-moneda): es solo VISUAL,
# FilaSolicitud conserva todos sus campos. Suma de porcentajes ≈ 100 (llena la
# tarjeta sin scroll; si algún día se suben para pasar de 100, aparece scroll).
# Cada entrada: (etiqueta, pct, alineación, extractor|None).
_COLS_PCT = [
    ("", 3, CENTRO, None),                                    # check
    ("Folio", 4, CENTRO, lambda f: f.folio),
    ("Folio Factura", 7, CENTRO, lambda f: f.folio_factura),
    ("Tipo Solicitud", 13, CENTRO, lambda f: f.tipo_solicitud),
    ("Total Fact.", 9, _TDER, lambda f: _fmt_moneda(f.total_factura)),
    ("Saldo Fact.", 9, _TDER, lambda f: _fmt_moneda(f.saldo_factura)),
    ("Saldo Prog.", 9, _TDER, lambda f: _fmt_moneda(f.saldo_programado)),
    ("Tipo", 3, CENTRO, lambda f: f.tipo),
    ("Fh. Fact.", 9, CENTRO, lambda f: f.fecha_factura),
    ("Fh. Ven.", 9, CENTRO, lambda f: f.fecha_vencimiento),
    ("Producto", 24, _TIZQ, lambda f: f.producto),
]


def _fmt_fecha(d: datetime.date) -> str:
    return d.strftime("%d/%m/%Y")


def _fechas_defecto(hoy: "datetime.date | None" = None) -> tuple[datetime.date, datetime.date]:
    """Rango por defecto de los filtros de fecha, pensado para el cierre mensual:

    - Fecha inicio: día 1 del MES ANTERIOR.
    - Fecha fin: último día del MES EN CURSO; salvo que HOY ya sea el último día del
      mes, en cuyo caso se extiende al día 10 del MES SIGUIENTE (para alcanzar las
      solicitudes que caen a inicios del mes que entra).
    """
    hoy = hoy or datetime.date.today()
    # Día 1 del mes anterior.
    if hoy.month == 1:
        inicio = datetime.date(hoy.year - 1, 12, 1)
    else:
        inicio = datetime.date(hoy.year, hoy.month - 1, 1)
    # Último día del mes en curso.
    ultimo_dia = calendar.monthrange(hoy.year, hoy.month)[1]
    if hoy.day == ultimo_dia:
        # Hoy es fin de mes: extender al 10 del mes siguiente.
        if hoy.month == 12:
            fin = datetime.date(hoy.year + 1, 1, 10)
        else:
            fin = datetime.date(hoy.year, hoy.month + 1, 10)
    else:
        fin = datetime.date(hoy.year, hoy.month, ultimo_dia)
    return inicio, fin


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
    """Tabla de solicitudes de UNA empresa (render por "secciones").

    Agrupa las filas por cuenta bancaria y muestra cada grupo como una BANDA de
    cabecera (Proveedor + Cuenta + Total Programado + check que selecciona todo el
    grupo) seguida de sus filas de detalle COMPACTAS (sin Proveedor/Cuenta/Moneda:
    ya van en la banda y en la pestaña empresa-moneda; solo se ocultan, el objeto
    conserva sus campos). Colorea cada fila (verde si Saldo Factura == Saldo
    Programado, rojo si difieren) y ofrece un check por fila con 'seleccionar todas'
    en el encabezado. Al agregar nuevos reportes evita duplicar filas ya presentes
    (por su 'clave')."""

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
        # Pago en pesos POR PAR (proveedor, cuenta_bancaria) — cada grupo de la tabla
        # es un par. Solo aplica en tablas USD. Persisten entre re-renders de la barra.
        # Solo alimentan el TXT en pesos; no se escriben en el formulario de SIPP.
        self._pagar_pesos: set[tuple] = set()          # pares marcados
        self._concepto_prov: dict[tuple, str] = {}     # {(prov, cuenta): concepto}
        self._ref_prov: dict[tuple, str] = {}          # {(prov, cuenta): referencia}
        self._clabe_prov: dict[tuple, str] = {}        # {(prov, cuenta): clabe origen}
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
        # Render con TablaResponsiva: columnas por PORCENTAJE del ancho de la tarjeta
        # (se adaptan al tamaño de ventana; scroll horizontal si superan el 100%). El
        # check 'seleccionar todas' va en la columna 0 del encabezado. Las bandas de
        # grupo se pasan como filas-cabecera al reconstruir.
        columnas = [
            ColumnaTabla(
                etiqueta, pct, alineacion,
                encabezado_control=(self.chk_todos if i == 0 else None))
            for i, (etiqueta, pct, alineacion, _fn) in enumerate(_COLS_PCT)
        ]
        self._tabla = TablaResponsiva(
            self.page, columnas,
            ancho_inicial=(getattr(self.page, "width", None) or 1200) - 90)
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
        # La Cuenta Origen del pago en pesos ya NO es única por grupo: cada par
        # (proveedor, cuenta beneficiario) tiene su propio selector, creado en
        # _reconstruir_pesos (con opciones de self._clabes y valor en self._clabe_prov).
        self.tf_concepto = ft.TextField(
            label="Concepto de Pago", width=200)
        self.tf_referencia = ft.TextField(
            label="Referencia de Pago", width=200)
        self._filtro_row = ft.Row(
            [self.tf_venc, self.dd_cuenta, self.tf_concepto, self.tf_referencia],
            spacing=8, wrap=True, vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        # Mensaje que sustituye a las filas cuando el filtro de Fecha Vencimiento
        # oculta TODAS las solicitudes. Visible solo cuando no hay filas visibles.
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
        # STRETCH: estira la tabla a lo ancho de la tarjeta para que TablaResponsiva
        # mida el ancho REAL disponible (y las columnas se dimensionen por %).
        self.control = ft.Column(
            [self._filtro_row, self._pesos_holder, self._pager,
             self._tabla.control, self._msg_vacio_filtro],
            spacing=8, horizontal_alignment=ft.CrossAxisAlignment.STRETCH)

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

    def set_clabes(self, clabes) -> None:
        """Reemplaza los pares (cuenta, clabe) de ORIGEN (opciones de los selectores de
        pago en pesos, p. ej. tras recargar el catálogo). Descarta selecciones por par
        cuya CLABE ya no exista y reconstruye la barra."""
        self._clabes = list(clabes or [])
        claves = {cl for _cta, cl in self._clabes}
        self._clabe_prov = {
            par: cl for par, cl in self._clabe_prov.items() if cl in claves}
        self._reconstruir_pesos()
        self._repintar()

    def _cuenta_texto_de_clabe(self, clabe: str) -> str:
        """Texto (cuenta con banco/empresa) de una CLABE de origen ('' si no está).
        Sirve para determinar el banco/formato del TXT en pesos."""
        for cuenta, cl in self._clabes:
            if cl == clabe:
                return cuenta
        return ""

    def clabes_pesos(self) -> dict[tuple, str]:
        """CLABE de origen elegida por par (proveedor, cuenta) marcado 'pagar en
        pesos' (solo pares con clabe elegida)."""
        pares = self.pares_pagar_pesos()
        return {par: cl for par, cl in self._clabe_prov.items()
                if par in pares and cl}

    def cuentas_pesos_texto(self) -> dict[tuple, str]:
        """Texto de la cuenta origen elegida por par (para decidir banco/formato)."""
        return {par: self._cuenta_texto_de_clabe(cl)
                for par, cl in self.clabes_pesos().items()}

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

    def pares_pagar_pesos(self) -> set[tuple]:
        """Pares (proveedor, cuenta_bancaria) de esta tabla USD marcados para pagar en
        pesos. Solo se consideran los que además tienen alguna solicitud SELECCIONADA."""
        if not self.es_usd():
            return set()
        pares_sel = {(f.proveedor, f.cuenta_bancaria) for f in self.seleccionadas()}
        return {par for par in self._pagar_pesos if par in pares_sel}

    def conceptos_pesos(self) -> dict[tuple, str]:
        """Concepto de Pago por par marcado 'pagar en pesos' (solo con valor)."""
        pesos = self.pares_pagar_pesos()
        return {par: v for par, v in self._concepto_prov.items() if par in pesos and v}

    def referencias_pesos(self) -> dict[tuple, str]:
        """Referencia de Pago por par marcado 'pagar en pesos' (solo con valor)."""
        pesos = self.pares_pagar_pesos()
        return {par: v for par, v in self._ref_prov.items() if par in pesos and v}

    def _pares_prov_cuenta(self) -> list[tuple]:
        """Pares (proveedor, cuenta_bancaria) DISTINTOS presentes en la tabla, en orden
        de aparición (equivalen a los grupos de la tabla)."""
        vistos: list[tuple] = []
        for f in self.filas:
            par = (f.proveedor, f.cuenta_bancaria)
            if f.proveedor and par not in vistos:
                vistos.append(par)
        return vistos

    def _reconstruir_pesos(self) -> None:
        """(Re)arma la barra de 'Pagar en pesos': una fila por cada par (proveedor,
        cuenta beneficiario) de la tabla, con su check y —cuando está marcado— su
        propia Cuenta Origen, Concepto y Referencia (habilitados solo si está marcado;
        al desmarcar se limpian). Solo visible en tablas USD; oculta si no hay pares."""
        pares = self._pares_prov_cuenta() if self.es_usd() else []
        if not pares:
            self._pesos_holder.visible = False
            self._pesos_holder.content = None
            return
        # Limpia marcas y valores de pares que ya no están en la tabla.
        vigentes = set(pares)
        self._pagar_pesos &= vigentes
        self._concepto_prov = {
            k: v for k, v in self._concepto_prov.items() if k in vigentes}
        self._ref_prov = {k: v for k, v in self._ref_prov.items() if k in vigentes}
        self._clabe_prov = {
            k: v for k, v in self._clabe_prov.items() if k in vigentes}

        def _toggle(e, par):
            if e.control.value:
                self._pagar_pesos.add(par)
            else:
                # Al desmarcar el par se limpia la info de sus inputs (Cuenta,
                # Concepto, Referencia), como el selector al deshabilitarse.
                self._pagar_pesos.discard(par)
                self._concepto_prov.pop(par, None)
                self._ref_prov.pop(par, None)
                self._clabe_prov.pop(par, None)
            # Re-render de la barra: habilita/deshabilita los inputs del par.
            self._reconstruir_pesos()
            self._repintar()

        def _fila_par(par: tuple) -> ft.Control:
            prov, cuenta = par
            marcado = par in self._pagar_pesos
            etiqueta = f"{prov} · {cuenta}" if cuenta else str(prov)
            chk = ft.Checkbox(
                label=etiqueta, value=marcado,
                on_change=lambda e, p=par: _toggle(e, p))
            # Cuenta Origen / Concepto / Referencia SIEMPRE se muestran; habilitados
            # solo si el par está marcado. Sin 'dense' (altura estándar de Material).
            dd_origen = ft.Dropdown(
                label="Cuenta Origen (pago en pesos)", width=340,
                enable_filter=True, editable=True, disabled=not marcado,
                value=self._clabe_prov.get(par),
                options=[ft.dropdown.Option(key=cl, text=cta)
                         for cta, cl in self._clabes],
                on_select=lambda e, p=par: self._clabe_prov.__setitem__(
                    p, e.control.value or ""))
            tf_concepto = ft.TextField(
                label="Concepto de Pago", width=200, disabled=not marcado,
                value=self._concepto_prov.get(par, ""),
                on_change=lambda e, p=par: self._concepto_prov.__setitem__(
                    p, (e.control.value or "").strip()))
            tf_ref = ft.TextField(
                label="Referencia de Pago", width=200, disabled=not marcado,
                value=self._ref_prov.get(par, ""),
                on_change=lambda e, p=par: self._ref_prov.__setitem__(
                    p, (e.control.value or "").strip()))
            # Orden pedido: Cuenta · Concepto · Referencia (tras el check del par).
            return ft.Row(
                [chk, dd_origen, tf_concepto, tf_ref],
                spacing=12, wrap=True,
                vertical_alignment=ft.CrossAxisAlignment.CENTER)

        filas = ft.Column([_fila_par(p) for p in pares], spacing=10, tight=True)
        self._pesos_holder.content = ft.Container(
            ft.Column(
                [
                    ft.Text(
                        "Pagar en pesos (por proveedor y cuenta): se genera un TXT "
                        "aparte en pesos al finalizar. Cuenta Origen es requerida; "
                        "Concepto y Referencia son opcionales.",
                        size=12, color=GRIS),
                    filas,
                ],
                spacing=6, tight=True),
            padding=ft.Padding.symmetric(horizontal=8, vertical=6),
            bgcolor=ft.Colors.SURFACE_CONTAINER_HIGHEST, border_radius=8)
        self._pesos_holder.visible = True

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
        filas_tabla: list = []
        for cuenta in self._paginas[self._pagina]:
            grupo = grupos[cuenta]
            total = sum((f.saldo_programado or 0) for f in grupo)
            # Banda del grupo (proveedor+cuenta+total) con un check que selecciona/
            # deselecciona TODAS las solicitudes de ese grupo de una sola vez; debajo
            # sus filas de detalle.
            proveedor = grupo[0].proveedor if grupo else ""
            filas_tabla.append(self._banda_grupo(
                proveedor, cuenta, {f.clave() for f in grupo}, total))
            for f in grupo:
                filas_tabla.append(self._fila_detalle(f))
        # 'Seleccionar todas' refleja las filas VISIBLES seleccionadas (se fija ANTES
        # de pintar, porque el check vive en el encabezado que reconstruye la tabla).
        self.chk_todos.value = bool(visibles) and all(
            f.clave() in self._sel for f in visibles)
        self._tabla.set_contenido(filas_tabla)
        # Si hay filas pero el filtro de vencimiento las oculta TODAS, se muestra el
        # mensaje a todo lo ancho.
        self._msg_vacio_filtro.visible = bool(self.filas) and not visibles
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

    def _banda_grupo(self, proveedor: str, cuenta: str, claves: set,
                     total: float) -> Cabecera:
        """Fila-cabecera (banda) de un grupo (proveedor+cuenta) con su Total Programado
        y un check tri-estado que selecciona/deselecciona TODAS sus solicitudes a la vez
        (marcado=todas, vacío=ninguna, indeterminado=algunas). Devuelve una `Cabecera`:
        el 1er segmento (col del check) lleva el check; el 2º abarca el resto con
        'proveedor · cuenta ……… TOTAL PROG. $X'."""
        seleccionadas = claves & self._sel
        if not seleccionadas:
            estado = False
        elif seleccionadas == claves:
            estado = True
        else:
            estado = None  # indeterminado

        def _toggle(_e, claves=claves):
            if claves & self._sel == claves:
                self._sel -= claves
            else:
                self._sel |= claves
            self._reconstruir()
            self._repintar()  # repinta tabla + barra de pesos (dependen de _sel)

        chk = ft.Checkbox(value=estado, tristate=True, on_change=_toggle)
        # Izquierda: proveedor (fuerte) · cuenta (tenue), ocupa el hueco (expand).
        prov_txt = str(proveedor or "—")
        cta_txt = str(cuenta or "")
        izq = ft.Row(
            [
                ft.Text(prov_txt, size=13, weight=ft.FontWeight.BOLD,
                        max_lines=1, no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS,
                        tooltip=prov_txt if len(prov_txt) > 40 else None),
                ft.Text("·", size=13, color=GRIS),
                ft.Text(cta_txt, size=12, weight=ft.FontWeight.BOLD, color=GRIS,
                        max_lines=1, no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS,
                        tooltip=cta_txt if len(cta_txt) > 34 else None),
            ],
            spacing=8, tight=True, expand=True,
            vertical_alignment=ft.CrossAxisAlignment.CENTER)
        der = ft.Row(
            [
                ft.Text("TOTAL PROG.", size=11, weight=ft.FontWeight.BOLD, color=GRIS),
                ft.Text(_fmt_moneda(total), size=13, weight=ft.FontWeight.BOLD),
            ],
            spacing=6, tight=True, vertical_alignment=ft.CrossAxisAlignment.CENTER)
        # alineacion=None -> el Row llena el ancho del segmento y el total queda a la
        # derecha (izq expande). El check va centrado en su columna (col 0).
        info = ft.Row([izq, der], vertical_alignment=ft.CrossAxisAlignment.CENTER)
        return Cabecera([
            SegmentoCabecera(1, chk, alineacion=CENTRO),
            SegmentoCabecera(len(_COLS_PCT) - 1, info, alineacion=None,
                             padding=ft.Padding.only(right=10)),
        ])

    def _fila_detalle(self, f: FilaSolicitud) -> FilaDatos:
        """Fila de detalle (sin Proveedor/Cuenta/Moneda, que ya van en la banda / la
        pestaña). El check por fila persiste la selección en self._sel. Devuelve una
        `FilaDatos`: celda 0 = check (control); resto = textos de `_COLS_PCT`."""
        chk = ft.Checkbox(value=f.clave() in self._sel)

        def _al_check(_e, f=f, c=chk):
            (self._sel.add if c.value else self._sel.discard)(f.clave())

        chk.on_change = _al_check
        self._checks_filas.append(chk)

        celdas: list = [chk]
        for _etiqueta, _pct, _alin, fn in _COLS_PCT[1:]:
            celdas.append(fn(f))
        return FilaDatos(celdas, bgcolor=_color_fila(f))

    def _marcar_todas(self, _e) -> None:
        # Selecciona/deselecciona todas las filas VISIBLES (de todas las páginas,
        # respetando el filtro de vencimiento) y reconstruye para reflejar los checks
        # por fila y las bandas. Las ocultas no se seleccionan.
        claves_visibles = {f.clave() for f in self._filas_visibles()}
        if self.chk_todos.value:
            self._sel |= claves_visibles
        else:
            self._sel -= claves_visibles
        self._reconstruir()
        self._repintar()


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
        # 'Pagar en pesos': por grupo, los PARES (proveedor, cuenta beneficiario)
        # marcados y, por par, la clabe origen / texto de cuenta / concepto /
        # referencia (se capturan al generar la dispersión, porque las tablas se
        # vacían). Más: TXT en pesos generados, tipo de cambio y posible error.
        # Clave interna de cada dict interno: la tupla (proveedor, cuenta_bancaria).
        self._pesos_por_grupo: dict[str, set] = {}
        self._clabe_pesos_por_grupo: dict[str, dict[tuple, str]] = {}
        self._cuenta_pesos_por_grupo: dict[str, dict[tuple, str]] = {}
        self._concepto_prov_por_grupo: dict[str, dict[tuple, str]] = {}
        self._ref_prov_por_grupo: dict[str, dict[tuple, str]] = {}
        # Referencia leída del DOM por el RPA (respaldo), por par.
        self._ref_dom_por_grupo: dict[str, dict[tuple, str]] = {}
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
        # El campo es de solo lectura y abre el calendario al hacer clic; el RPA lee
        # su texto. Defaults pensados para el cierre mensual: Inicio = día 1 del mes
        # anterior; Fin = último día del mes en curso (o el 10 del siguiente si hoy
        # es fin de mes). Ver _fechas_defecto().
        inicio_defecto, fin_defecto = _fechas_defecto()
        self.dp_fecha_ini = ft.DatePicker(
            value=inicio_defecto,
            first_date=datetime.date(2020, 1, 1),
            last_date=datetime.date(2035, 12, 31),
            help_text="Fecha Inicio",
            on_change=lambda e: self._fecha_elegida(
                self.tf_fecha_ini, self.dp_fecha_ini),
        )
        self.dp_fecha_fin = ft.DatePicker(
            value=fin_defecto,
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
            label=_label_requerido("Fecha Fin"), value=_fmt_fecha(fin_defecto),
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
        # El contenido va dentro de un contenedor con padding derecho (gutter) para
        # que la barra de scroll vertical (que Flet dibuja encima) no se solape con
        # las tarjetas (mismo criterio que el modal de solicitudes a dispersar).
        contenido = ft.Column(
            [panel, panel_tabla],
            spacing=14, horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
        )
        return ft.Column(
            [ft.Container(contenido, padding=ft.Padding.only(right=_GUTTER_SCROLL))],
            expand=True, scroll=ft.ScrollMode.AUTO,
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
            spacing=12, horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
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
            spacing=10, horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
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

    def _todas_por_empresa(self) -> dict[str, list[FilaSolicitud]]:
        """TODAS las solicitudes encontradas por empresa (no solo las seleccionadas),
        para el reporte 'Todas las solicitudes'. Una tabla por empresa."""
        resultado: dict[str, list[FilaSolicitud]] = {}
        for empresa, tabla in self._tablas_por_empresa.items():
            if tabla.filas:
                resultado[empresa] = list(tabla.filas)
        return resultado

    def _texto_tipo_solicitud(self) -> str:
        """Valor para el filtro 'Tipo Solicitud' del reporte: 'Todos' si están
        todos (o ninguno) seleccionados; si no, los tipos elegidos unidos por ', '."""
        tipos = self.ms_tipo.valores()
        if not tipos or set(tipos) == set(self.TIPOS_SOLICITUD):
            return "Todos"
        return ", ".join(tipos)

    def _generar_reporte(self, _e=None) -> None:
        """Al presionar 'Generar Reporte' pregunta QUÉ reporte generar: todas las
        solicitudes encontradas o solo las seleccionadas (o cancelar). Cada opción
        exporta con el mismo formato del reporte del SIPP (_exportar_reporte)."""
        async def _todas(_e=None):
            self.page.pop_dialog()
            await self._exportar_reporte(
                self._todas_por_empresa(), "Todas las Solicitudes")

        async def _seleccionadas(_e=None):
            self.page.pop_dialog()
            await self._exportar_reporte(
                self._seleccion_por_empresa(), "Solicitudes Seleccionadas")

        dialogo = ft.AlertDialog(
            modal=True,
            title=ft.Text("Generar Reporte"),
            content=ft.Text("Seleccione qué tipo de reporte quiere generar."),
            actions=[
                ft.TextButton("Cancelar", on_click=lambda _e: self.page.pop_dialog()),
                ft.OutlinedButton("Todas las solicitudes", on_click=_todas),
                ft.FilledButton("Solicitudes seleccionadas", on_click=_seleccionadas),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self.page.show_dialog(dialogo)

    async def _exportar_reporte(
        self, seleccion: dict, etiqueta: str) -> None:
        """Exporta a Excel el `seleccion` dado (una hoja por empresa), con el mismo
        formato del reporte que se lee del SIPP. `etiqueta` distingue el tipo en el
        nombre de archivo y los avisos. Los datos del bloque de filtros (B3:G6) se
        toman del filtro principal."""
        if not seleccion:
            self.app.avisar(
                "No hay solicitudes para el reporte "
                f"'{etiqueta.lower()}'.", NARANJA)
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
            dialog_title=f"Guardar reporte de {etiqueta.lower()}",
            file_name=f"Reporte {etiqueta}.xlsx",
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
        """Abre un archivo o la carpeta de resultados en el sistema, trayéndolo al
        frente de la app (ver AppTesoreria.abrir_en_sistema)."""
        self.app.abrir_en_sistema(ruta)

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
        # Pares (proveedor, cuenta beneficiario) marcados 'pagar en pesos' por grupo
        # USD (se capturan ahora, porque después de dispersar las tablas se vacían).
        # Solo cuentan los pares con alguna solicitud seleccionada.
        self._pesos_por_grupo = {
            grupo: self._tablas_por_empresa[grupo].pares_pagar_pesos()
            for grupo in seleccion
            if self._tablas_por_empresa[grupo].pares_pagar_pesos()
        }
        # La Cuenta Origen del pago en pesos es REQUERIDA por CADA par marcado (es la
        # cuenta origen del TXT en pesos). Se reúnen los pares sin clabe para avisar.
        self._clabe_pesos_por_grupo = {
            grupo: self._tablas_por_empresa[grupo].clabes_pesos()
            for grupo in self._pesos_por_grupo
        }
        sin_clabe = [
            f"{prov} · {cuenta}"
            for grupo, pares in self._pesos_por_grupo.items()
            for (prov, cuenta) in pares
            if (prov, cuenta) not in self._clabe_pesos_por_grupo.get(grupo, {})
        ]
        if sin_clabe:
            self.app.avisar(
                "Falta elegir la Cuenta Origen (pago en pesos) en: "
                + ", ".join(sin_clabe) + ".", NARANJA)
            return
        # Texto de la cuenta origen elegida por par (para saber el banco/formato del
        # layout en pesos: Banregio vs Bancomer).
        self._cuenta_pesos_por_grupo = {
            grupo: self._tablas_por_empresa[grupo].cuentas_pesos_texto()
            for grupo in self._pesos_por_grupo
        }
        # Concepto/Referencia POR PAR (pago en pesos), capturados en la UI. Solo
        # alimentan el TXT en pesos. La referencia final se resuelve en el RPA con el
        # respaldo del DOM (ver _ejecutar_dispersion / _generar_txts_pesos).
        self._concepto_prov_por_grupo = {
            grupo: self._tablas_por_empresa[grupo].conceptos_pesos()
            for grupo in self._pesos_por_grupo
        }
        self._ref_prov_por_grupo = {
            grupo: self._tablas_por_empresa[grupo].referencias_pesos()
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
        self._ref_dom_por_grupo = {}  # {grupo: {proveedor: referencia_DOM}}
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
                # 5) Concepto/referencia por proveedor+cuenta (obj. 6/7). Devuelve la
                #    referencia que el portal traía precargada (DOM) por proveedor+
                #    cuenta: se guarda como respaldo del TXT en pesos (obj. 4.2).
                referencias_dom = await sesion.llenar_pagos_proveedores(
                    _pares_proveedor_cuenta(emp),
                    emp.concepto_pago, emp.referencia_pago)
                if emp.empresa in self._pesos_por_grupo:
                    # referencias_dom ya viene por par {(prov,cuenta): ref}; se guarda
                    # tal cual (solo las no vacías) como respaldo del TXT en pesos.
                    self._ref_dom_por_grupo[emp.empresa] = {
                        par: ref for par, ref in referencias_dom.items() if ref}
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
        """Genera, en la carpeta de descargas, los TXT en PESOS de los grupos USD
        dispersados con pares (proveedor, cuenta beneficiario) marcados 'pagar en
        pesos'. El importe = (saldo programado del par en USD) × tipo de cambio del
        DOF. Como cada par elige su propia Cuenta Origen, se agrupan los registros por
        cuenta origen y se genera UN ARCHIVO por cada una (su banco define el formato:
        Banregio o BBVA/Bancomer). No tumba la operación si algo falla."""
        conc = self._conc_dispersion
        if conc is None or not self._pesos_por_grupo:
            return
        por_clave_folio = {
            d.get("clave"): d for d in self._folios_dispersados if d.get("clave")}
        por_clave_emp = {emp.empresa: emp for emp in conc.validas}
        # Grupos USD efectivamente dispersados y con pares marcados.
        pendientes = [
            clave for clave, pares in self._pesos_por_grupo.items()
            if pares and clave in por_clave_folio and clave in por_clave_emp
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
            pares = self._pesos_por_grupo[clave]
            emp = por_clave_emp[clave]
            folio_entry = por_clave_folio[clave]
            folio = folio_entry.get("folio")
            # Concepto/Referencia por par (obj. 4). El layout tiene un solo campo de
            # texto ("concepto"): lleva el Concepto y, si hay, se le anexa la
            # Referencia (app o, en su defecto, la precargada en el DOM).
            conceptos = self._concepto_prov_por_grupo.get(clave, {})
            refs_app = self._ref_prov_por_grupo.get(clave, {})
            refs_dom = self._ref_dom_por_grupo.get(clave, {})
            clabes = self._clabe_pesos_por_grupo.get(clave, {})
            cuentas_txt = self._cuenta_pesos_por_grupo.get(clave, {})
            # Agrupa los registros por CUENTA ORIGEN (un archivo por origen).
            por_origen: dict[str, dict] = {}
            for par in pares:
                prov, cuenta = par
                clabe_origen = clabes.get(par, "")
                if not clabe_origen:  # sin cuenta origen (validado antes; por si acaso)
                    continue
                movs = [m for m in emp.movimientos
                        if m.proveedor == prov and m.cuenta_bancaria == cuenta]
                usd = sum((m.saldo_programado or 0) for m in movs)
                pesos = round(usd * tc, 2)
                cuenta_benef = re.sub(r"\D", "", cuenta or "")
                concepto = (conceptos.get(par) or emp.concepto_pago or "").strip()
                referencia = (refs_app.get(par) or refs_dom.get(par) or "").strip()
                texto = f"{concepto} {referencia}".strip() if referencia else concepto
                bucket = por_origen.setdefault(clabe_origen, {
                    "cuenta_texto": cuentas_txt.get(par, ""),
                    "registros": [], "total": 0.0})
                bucket["registros"].append((cuenta_benef, pesos, prov, texto))
                bucket["total"] += pesos
            # Un TXT por cada cuenta origen.
            for clabe_origen, bucket in por_origen.items():
                registros = bucket["registros"]
                if not registros:
                    continue
                clabe_dig = re.sub(
                    r"\D", "", clabe_origen or folio_entry.get("cuenta_origen") or "")
                cuenta_texto = bucket["cuenta_texto"]
                if exportador_devoluciones.banco_formato(cuenta_texto) == "banregio":
                    # Banregio: separado por comas; usa la fecha (DDMMAAAA) de hoy.
                    hoy = datetime.date.today().strftime("%d%m%Y")
                    contenido = exportador_devoluciones.generar_banregio(registros, hoy)
                else:  # BBVA / Bancomer (ancho fijo) — formato por defecto
                    contenido = exportador_devoluciones.generar_bancomer(
                        registros, clabe_dig, str(folio or ""))
                # Distintivo de la cuenta origen en el nombre (últimos 4 dígitos), para
                # no colisionar cuando hay varias cuentas origen en el mismo grupo.
                sufijo = f"Pesos {clabe_dig[-4:]}" if clabe_dig else "Pesos"
                nombre = _sanear_archivo(
                    self._nombre_txt(folio_entry) + " " + sufijo) + ".txt"
                ruta = _ruta_unica(os.path.join(carpeta, nombre))
                try:
                    with open(ruta, "w", encoding="latin-1", newline="") as fh:
                        fh.write(contenido)
                except Exception:  # noqa: BLE001 — un TXT que falle no aborta el resto
                    continue
                self._pesos_generados.append({
                    "empresa": emp.empresa, "archivo": ruta,
                    "total_pesos": bucket["total"], "proveedores": len(registros)})

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
                titulo += f" (T.C. {_fmt_tc(tc)} MXN)"
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
            # El contenido va dentro de un contenedor con padding derecho: así la
            # barra de scroll vertical (que Flet dibuja encima) queda en ese margen y
            # no se solapa con la última columna de las tablas.
            cuerpo = ft.Column(
                [ft.Container(
                    ft.Column(secciones, tight=True, spacing=14),
                    padding=ft.Padding.only(right=_GUTTER_SCROLL))],
                scroll=ft.ScrollMode.AUTO, tight=True)
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
            texto = (f"Tipo de cambio (DOF): {_fmt_tc(self._tc_preview)} MXN "
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
        saldo de factura + saldo programado), con sus totales. Si el grupo es USD y
        tiene proveedores marcados 'pagar en pesos', agrega una columna con la
        equivalencia en pesos (Saldo Programado × T.C.) para esos proveedores."""
        W_FOLIO, W_FOLIO_FAC, W_PROV, W_CTA, W_MONTO, W_PESOS = (
            60, 90, 200, 200, 105, 115)

        # Pares (proveedor, cuenta beneficiario) del grupo marcados 'pagar en pesos' y
        # el tipo de cambio. La columna solo aparece si hay marcados y hay T.C.
        pesos_set = self._pesos_por_grupo.get(e.empresa, set())
        tc = self._tc_preview
        mostrar_pesos = bool(pesos_set) and bool(tc)

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
        tot_pesos = 0.0
        for m in e.movimientos:
            tot_prog += m.saldo_programado or 0
            celdas = [
                celda(m.folio, W_FOLIO),
                celda(m.folio_factura, W_FOLIO_FAC),
                celda(m.proveedor, W_PROV),
                celda(m.cuenta_bancaria, W_CTA),
                celda(_fmt_moneda(m.total_factura), W_MONTO, derecha=True),
                celda(_fmt_moneda(m.saldo_factura), W_MONTO, derecha=True),
                celda(_fmt_moneda(m.saldo_programado), W_MONTO, derecha=True),
            ]
            if mostrar_pesos:
                if (m.proveedor, m.cuenta_bancaria) in pesos_set:
                    pesos = round((m.saldo_programado or 0) * tc, 2)
                    tot_pesos += pesos
                    celdas.append(celda(_fmt_moneda(pesos), W_PESOS, derecha=True))
                else:  # (proveedor, cuenta) USD que NO se paga en pesos
                    celdas.append(celda("—", W_PESOS, derecha=True))
            filas.append(ft.DataRow(cells=celdas))
        # Fila de totales (Saldo Programado y, si aplica, equivalencia en pesos).
        total_celdas = [
            celda("", W_FOLIO), celda("", W_FOLIO_FAC), celda("", W_PROV),
            celda("", W_CTA), celda("", W_MONTO),
            celda("TOTAL PROGRAMADO", W_MONTO, derecha=True, bold=True),
            celda(_fmt_moneda(tot_prog), W_MONTO, derecha=True, bold=True),
        ]
        if mostrar_pesos:
            total_celdas.append(
                celda(_fmt_moneda(tot_pesos), W_PESOS, derecha=True, bold=True))
        filas.append(ft.DataRow(cells=total_celdas))

        columnas = [
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
        ]
        if mostrar_pesos:
            columnas.append(ft.DataColumn(
                label=encabezado_col("Equiv. MXN", W_PESOS), numeric=True))

        tabla = ft.DataTable(
            columns=columnas,
            rows=filas,
            column_spacing=14, horizontal_margin=6,
            heading_row_height=34, data_row_min_height=30, data_row_max_height=30,
            divider_thickness=1,
            border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT), border_radius=8,
        )
        # Padding inferior: reserva sitio para la barra de scroll horizontal (que
        # Flet dibuja encima) y evita que tape la fila TOTAL PROGRAMADO.
        return ft.Row(
            [ft.Container(tabla, padding=ft.Padding.only(bottom=_GUTTER_SCROLL))],
            scroll=ft.ScrollMode.AUTO)

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
