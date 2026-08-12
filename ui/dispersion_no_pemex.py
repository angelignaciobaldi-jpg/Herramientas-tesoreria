"""Pantalla: Dispersión (No Pemex).

Ofrece los controles para operar el RPA del SIPP: una tarjeta de filtros (los
mismos del modal "Agregar Facturas/Solicitudes de Pago") y los botones para
iniciar/pausar/reanudar y detener la ejecución. Al iniciar, hace login en el
SIPP, selecciona empresa/sucursal de la sesión y, por cada combinación de
empresa × tipo de solicitud elegida, aplica los filtros, busca y descarga el
Excel del reporte.

Las credenciales de inicio de sesión se capturan en el menú "Configuración" de
la barra superior (ver ui/configuracion.py); aquí se leen desde ahí al arrancar.

La BÚSQUEDA de solicitudes consulta el microservicio (core/api.py) y vuelca las
filas mapeadas (reporte_dispersion.desde_api); ya no abre el navegador. El RPA
(Playwright, en un bucle de asyncio aparte con BucleRpa para no congelar la interfaz)
se conserva SOLO para el paso de DISPERSAR en el SIPP.
"""

from __future__ import annotations

import asyncio
import calendar
import datetime
import os
import re
import unicodedata

import flet as ft

from core import (
    ajustes_api, api, conciliacion, cuentas_dispersion, exportador_devoluciones,
    preferencias, reporte_dispersion, reporte_dispersion_export, rutas, tipo_cambio,
)
from core.reporte_dispersion import FilaSolicitud
from core.rpa_sipp import (
    BucleRpa,
    ControlRpa,
    FiltrosSolicitudPago,
    RpaDetenido,
    SesionSipp,
)
from ui.comun import (CENTRO, EMPRESAS, GRIS, ID_POR_EMPRESA, NARANJA,
                      NOMBRES_EMPRESAS, ROJO, ROJO_BOTON, VERDE)
from ui.tabla_responsiva import (Cabecera, ColumnaTabla, FilaDatos,
                                 SegmentoCabecera, TablaResponsiva)
from ui.tabla_responsiva import DER as _TDER
from ui.tabla_responsiva import IZQ as _TIZQ

# Formato de fecha que pide el modal del SIPP.
_RE_FECHA = re.compile(r"^\d{2}/\d{2}/\d{4}$")

# Alineaciones para las celdas de la tabla.
_IZQ = ft.Alignment(-1, 0)
_DER = ft.Alignment(1, 0)

# Fondo de las filas del resumen PENDIENTES de subir su comprobante (modo 'errores').
_AMARILLO_PENDIENTE = ft.Colors.YELLOW_100

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
    'Busca las solicitudes de pago pendientes.'
)

# --- Formato de moneda ---------------------------------------------------
def _fmt_moneda(valor: float | None) -> str:
    """Formatea un monto como moneda con 2 decimales (p. ej. $1,234.50)."""
    return f"${(valor or 0):,.2f}"


def _fmt_tc(valor: float | None) -> str:
    """Formatea el tipo de cambio con 4 decimales (p. ej. $17.5993), tal como lo
    publica el DOF, para que coincida con el valor usado en el cálculo."""
    return f"${(valor or 0):,.4f}"


def _fecha_tc_texto() -> str:
    """Fecha del DOF con la que se pide el tipo de cambio, en 'DD/MM/AAAA': el día
    hábil anterior (los lunes, el viernes pasado). Ver `tipo_cambio.fecha_referencia`.
    Solo para mostrar cuando aún no se consultó el DOF; cuando ya hay respuesta se
    muestra la fecha que devolvió el DOF, que es la que de verdad se usó."""
    return tipo_cambio.fecha_referencia().strftime("%d/%m/%Y")


def _ultimos_digitos(texto, n: int = 4) -> str:
    """Últimos `n` dígitos de un texto, ignorando cualquier no-dígito (espacios,
    guiones, y máscaras como '*0012'). '' si no hay dígitos. Se usa para casar
    cuentas/CLABEs por sus últimos dígitos (los comprobantes vienen enmascarados)."""
    d = re.sub(r"\D", "", str(texto or ""))
    return d[-n:] if d else ""


def _norm_nombre_doc(nombre: str) -> str:
    """Normaliza un nombre de archivo para comparar el 'documento_lectura' que devuelve
    el extractor contra los PDFs subidos: minúsculas, sin espacios en los extremos y sin
    la extensión '.pdf'. Tolera diferencias de caja/extensión entre lo subido y lo que
    reporta la API."""
    s = (nombre or "").strip().lower()
    return s[:-4] if s.endswith(".pdf") else s


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


def _fecha_ddmmaaaa_a_iso(texto: str) -> str:
    """'DD/MM/AAAA' -> 'YYYY-MM-DD' (formato que pide el endpoint). '' si viene
    vacío o no es una fecha válida (la UI ya trabaja en DD/MM/AAAA)."""
    s = (texto or "").strip()
    if not s:
        return ""
    try:
        return datetime.datetime.strptime(s, "%d/%m/%Y").strftime("%Y-%m-%d")
    except ValueError:
        return ""


# Catálogo de tipos de solicitud: NOMBRE (como se ve en la UI) -> id que espera el
# endpoint (`tipoSolicitud`). Es la fuente única: la lista del combo
# (SeccionDispersionNoPemex.TIPOS_SOLICITUD) se deriva de estas claves, así el mapeo
# nombre->id al buscar siempre cuadra.
_TIPO_SOLICITUD_ID: dict[str, int] = {
    "Pago Factura": 1,
    "Pago Extraordinario": 2,
    "Pago Estadias": 3,
    "Pago Extraordinario Facturas": 4,
    "Pago General de Fletes": 5,
}


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
        # Checkboxes montados de la página ACTUAL, para reflejar la selección SIN
        # reconstruir: (checkbox, clave de la fila) y (checkbox de banda, claves del
        # grupo). Ver _aplicar_seleccion_en_vivo.
        self._checks_filas: list[tuple] = []
        self._checks_bandas: list[tuple] = []
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
        self._checks_bandas = []
        filas_tabla: list = []
        for cuenta in self._paginas[self._pagina]:
            grupo = grupos[cuenta]
            # Con las notas de crédito ya descontadas, para que el total de la banda
            # sea el mismo del TXT (ver reporte_dispersion.total_a_pagar).
            total = reporte_dispersion.total_a_pagar(grupo)
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
            # Sin reconstruir: solo se mueven los checks ya montados y la barra de
            # pesos (ver _aplicar_seleccion_en_vivo).
            self._aplicar_seleccion_en_vivo()

        chk = ft.Checkbox(value=estado, tristate=True, on_change=_toggle)
        self._checks_bandas.append((chk, claves))
        # Izquierda: proveedor (fuerte) · cuenta (tenue), ocupa el hueco (expand).
        prov_txt = str(proveedor or "—")
        cta_txt = str(cuenta or "")
        # Los Text llevan `expand` (no `tight`) para que el Row les FIJE un ancho: sin
        # eso toman su ancho natural, el '…' nunca entra y el texto se desborda
        # encimándose con el total de la derecha. 3/2 da más espacio al proveedor.
        izq = ft.Row(
            [
                ft.Text(prov_txt, size=13, weight=ft.FontWeight.BOLD, expand=3,
                        max_lines=1, no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS,
                        tooltip=prov_txt),
                ft.Text("·", size=13, color=GRIS),
                ft.Text(cta_txt, size=12, weight=ft.FontWeight.BOLD, color=GRIS,
                        expand=2,
                        max_lines=1, no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS,
                        tooltip=cta_txt or None),
            ],
            spacing=8, expand=True,
            vertical_alignment=ft.CrossAxisAlignment.CENTER)
        der = ft.Row(
            [
                ft.Text("TOTAL PROG.", size=11, weight=ft.FontWeight.BOLD, color=GRIS),
                ft.Text(_fmt_moneda(total), size=13, weight=ft.FontWeight.BOLD,
                        no_wrap=True,
                        tooltip=f"TOTAL PROG. {_fmt_moneda(total)}"),
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
            # Marcar una fila cambia el tri-estado de su banda, el check de
            # 'seleccionar todas' y la barra de pesos: se reflejan EN VIVO, sin
            # reconstruir (ver _aplicar_seleccion_en_vivo).
            self._aplicar_seleccion_en_vivo(salvo=c)

        chk.on_change = _al_check
        self._checks_filas.append((chk, f.clave()))

        celdas: list = [chk]
        for _etiqueta, _pct, _alin, fn in _COLS_PCT[1:]:
            celdas.append(fn(f))
        return FilaDatos(celdas, bgcolor=_color_fila(f))

    def _aplicar_seleccion_en_vivo(self, salvo=None) -> None:
        """Refleja `self._sel` en los controles YA montados, SIN reconstruir la tabla.

        Al pintar, lo único que depende de la selección es: el check de cada fila, el
        tri-estado de cada banda, el check 'seleccionar todas' y la barra de 'pagar en
        pesos'. Ni los colores de fila ni los totales cambian con ella.

        Reconstruir para marcar casillas creaba de cero los ~24 controles de cada fila
        —más de 2000 en una página llena— y por eso 'seleccionar todas' tardaba
        segundos en un beneficiario con muchos movimientos. Mutar los checks montados
        es proporcional al número de casillas, no al de celdas.

        `salvo` es el checkbox que el propio usuario acaba de mover: no se le reescribe
        el valor para no pelear con el estado que Flet ya le puso.
        """
        for chk, clave in self._checks_filas:
            if chk is salvo:
                continue
            chk.value = clave in self._sel
        for chk, claves in self._checks_bandas:
            sel = claves & self._sel
            chk.value = True if sel == claves and claves else (None if sel else False)
        visibles = self._filas_visibles()
        self.chk_todos.value = bool(visibles) and all(
            f.clave() in self._sel for f in visibles)
        # La barra de 'pagar en pesos' solo muestra los pares SELECCIONADOS, así que
        # sí hay que rearmarla; es pequeña (un bloque por par marcado).
        self._reconstruir_pesos()
        self._repintar()

    def _marcar_todas(self, _e) -> None:
        # Selecciona/deselecciona todas las filas VISIBLES (de todas las páginas,
        # respetando el filtro de vencimiento). Las ocultas no se seleccionan. Los
        # checks se mueven en vivo, sin reconstruir la página (ver
        # _aplicar_seleccion_en_vivo).
        claves_visibles = {f.clave() for f in self._filas_visibles()}
        if self.chk_todos.value:
            self._sel |= claves_visibles
        else:
            self._sel -= claves_visibles
        self._aplicar_seleccion_en_vivo(salvo=self.chk_todos)


class SeccionDispersionNoPemex:
    """Pestaña para operar el RPA de dispersión (No Pemex)."""

    # Empresa y plaza/sucursal con que se inicia la SESIÓN del SIPP (no es el
    # filtro de búsqueda, que el usuario elige abajo).
    EMPRESA_SESION = "Abastecedora"
    SUCURSAL_SESION = "Corporativo"

    # PRUEBAS: si es True, la operación de dispersión NO cierra el navegador al
    # terminar/detener, para poder inspeccionar el estado en el SIPP. En operación
    # normal va en False: al terminar cada RPA se cierra el navegador y la app vuelve
    # al frente.
    MANTENER_NAVEGADOR_PRUEBAS = False

    # Empresas disponibles para la dispersión (No Pemex). La fuente única vive en
    # ui.comun (se comparte con otras pantallas); aquí se referencian.
    EMPRESAS = EMPRESAS
    NOMBRES_EMPRESAS = NOMBRES_EMPRESAS
    ID_POR_EMPRESA = ID_POR_EMPRESA

    # Opciones fijas del combo "Tipo de Solicitud" (las del modal del SIPP).
    # Nombres del combo 'Tipo de Solicitud': se derivan del catálogo (nombre->id),
    # para que la selección de la UI mapee directo al id que pide el endpoint.
    TIPOS_SOLICITUD = list(_TIPO_SOLICITUD_ID)

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
        # Guard anti doble-click de 'Generar Dispersión': preparar el payload puede
        # tardar (TC del DOF + conciliación de muchas solicitudes) y, sin bloqueo, un
        # segundo clic lanzaría el flujo dos veces (excepción). `_dlg_cargando_disp`
        # es la pantalla de carga que se muestra mientras se prepara.
        self._generando_dispersion = False
        self._dlg_cargando_disp: ft.AlertDialog | None = None
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
        # Cuentas origen cambiadas a mano desde el resumen, por folio. La app las
        # aplica a sus filas, pero SIPP se corrige aparte: esto deja el rastro de qué
        # falta reflejar allá y es el punto de enganche si algún día se automatiza
        # (la tabla del SIPP expone 'editarDispersionProveedoresNoPemex(item)').
        self._cuentas_origen_pendientes_sipp: dict = {}
        self._tipo_cambio: float | None = None
        self._tc_fecha: str | None = None       # fecha DOF del TC usado (DD/MM/AAAA)
        self._pesos_error: str | None = None
        # Tipo de cambio de VISTA PREVIA (para mostrarlo en el modal 'Solicitudes a
        # dispersar' cuando hay proveedores USD marcados 'pagar en pesos'). Se
        # consulta al generar la dispersión; None si no aplica o no se pudo obtener.
        # `_tc_preview_fecha` es la fecha de publicación del DOF de ese valor.
        self._tc_preview: float | None = None
        self._tc_preview_fecha: str | None = None
        self._tc_preview_error: str | None = None
        # Comprobantes de pago cargados por el usuario en el resumen, por folio
        # (folio -> ruta de archivo). Las reglas de asignación están por definir.
        self._comprobantes: dict[str, str] = {}
        # Comprobantes LEÍDOS por la API extractor (data.comprobantes de la respuesta);
        # se conservan para el paso de asignación a folios (por definir).
        self._comprobantes_leidos: list[dict] = []
        # --- RPA de SUBIDA de comprobantes (Proveedores No Pemex) --------------
        # Movimientos (_id_fila) cuyo comprobante ya se subió OK; los errores de la
        # última subida; y el 'modo' de render del resumen (normal|exito|errores).
        self._subidos: set[str] = set()
        self._subida_errores: list[str] = []
        self._resumen_modo: str = "normal"
        # Estado/controles del RPA de subida (propio, paralelo al de dispersión).
        self._sub_estado_op = "detenido"
        self._sub_ctrl: ControlRpa | None = None
        self._sub_task = None
        self._sub_future = None
        self._dlg_subida: ft.AlertDialog | None = None
        self._sub_al_cerrar = None
        # Diálogo del resumen y callback a ejecutar CUANDO Flutter confirme su cierre.
        # Necesario porque Flet solo desmonta el diálogo de más arriba: para cerrar el
        # resumen (que queda DEBAJO de la confirmación) hay que esperar a que la
        # confirmación se desmonte y encadenar por on_dismiss.
        self._dlg_resumen: ft.AlertDialog | None = None
        self._resumen_al_cerrar = None
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
        # Búsqueda por API (ya no abre el navegador): consulta el endpoint y llena
        # las tablas. Se deshabilita mientras consulta.
        self.btn_iniciar = ft.FilledButton(
            content="Buscar solicitudes", icon=ft.Icons.SEARCH,
            on_click=self._buscar_solicitudes,
        )
        # 'Detener' ya no aplica a la búsqueda (la API es rápida): se oculta. Se
        # conserva la referencia para no romper el layout ni otros usos.
        self.btn_detener = ft.OutlinedButton(
            content="Detener", icon=ft.Icons.STOP, disabled=True, visible=False,
        )
        # Subir un Excel de solicitudes ya filtrado (sin correr el RPA).
        # OCULTO temporalmente: la carga manual por reportes se deshabilita por
        # ahora (para reactivarla, poner visible=True).
        self.btn_subir = ft.OutlinedButton(
            content="Búsqueda por reportes", icon=ft.Icons.UPLOAD_FILE,
            on_click=self._subir_reporte, visible=False,
        )

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
        # Parte inferior: el botón de búsqueda con la leyenda de requerido a la
        # derecha.
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

        # Panel de filtros de ALTURA FIJA (ya no es un acordeón). El título es
        # estático; el cuerpo hace scroll interno si no cabe en la altura fija.
        cuerpo_filtros = ft.Column(
            [
                fila_combos,
                fila_campos,
                fila_botones,
            ],
            spacing=14,
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
        # El panel toma su ALTURA NATURAL (se ajusta al contenido): así no queda
        # espacio muerto bajo el botón de búsqueda. En ventanas angostas los combos
        # se apilan y el panel crece; el scroll de toda la pantalla lo recorre.
        panel = ft.Card(
            content=ft.Container(
                content=ft.Column(
                    [encabezado_panel, cuerpo_filtros],
                    spacing=12,
                ),
                padding=16,
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
                ft.Text("Buscando solicitudes…", size=13, color=GRIS),
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
        y no se puede cerrar mientras la operación esté en curso.

        Preparar el payload puede tardar (consulta del TC al DOF + conciliación de
        muchas solicitudes), así que se muestra una pantalla de carga y se bloquea
        el botón / la re-entrada para que un segundo clic no dispare el flujo dos
        veces (lo que provocaba una excepción)."""
        # Guard anti doble-click: si ya se está preparando, ignora el nuevo clic.
        if self._generando_dispersion:
            return
        self._generando_dispersion = True
        self.btn_dispersar.disabled = True
        mostro_cargando = False
        abrir = False
        try:
            seleccion = self._seleccion_por_empresa()
            if not seleccion:
                self.app.avisar(
                    "Selecciona al menos un movimiento en la tabla.", NARANJA)
                return
            # La Cuenta de origen es REQUERIDA para dispersar: cada tabla con
            # movimientos seleccionados debe tener una cuenta elegida.
            sin_cuenta = [
                grupo for grupo in seleccion
                if self._tablas_por_empresa[grupo].cuenta_seleccionada() is None
            ]
            if sin_cuenta:
                self.app.avisar(
                    "Falta elegir la Cuenta en: " + ", ".join(sin_cuenta) + ".",
                    NARANJA)
                return
            # Datos de pago por empresa (cuenta elegida + concepto + referencia) para
            # adjuntarlos al payload de la dispersión (los usará el RPA: la 'cuenta'
            # es el valor 'Cuenta' por el que se busca la cuenta en SIPP).
            datos_pago: dict[str, dict] = {}
            for grupo in seleccion:
                tabla = self._tablas_por_empresa[grupo]
                datos_pago[grupo] = {
                    "cuenta": tabla.cuenta_seleccionada() or "",
                    "concepto_pago": tabla.concepto(),
                    "referencia_pago": tabla.referencia(),
                }
            # Pares (proveedor, cuenta beneficiario) marcados 'pagar en pesos' por
            # grupo USD (se capturan ahora, porque después de dispersar las tablas se
            # vacían). Solo cuentan los pares con alguna solicitud seleccionada.
            self._pesos_por_grupo = {
                grupo: self._tablas_por_empresa[grupo].pares_pagar_pesos()
                for grupo in seleccion
                if self._tablas_por_empresa[grupo].pares_pagar_pesos()
            }
            # La Cuenta Origen del pago en pesos es REQUERIDA por CADA par marcado (es
            # la cuenta origen del TXT en pesos). Se reúnen los pares sin clabe.
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
            # Texto de la cuenta origen elegida por par (para saber el banco/formato
            # del layout en pesos: Banregio vs Bancomer).
            self._cuenta_pesos_por_grupo = {
                grupo: self._tablas_por_empresa[grupo].cuentas_pesos_texto()
                for grupo in self._pesos_por_grupo
            }
            # Concepto/Referencia POR PAR (pago en pesos), capturados en la UI. Solo
            # alimentan el TXT en pesos. La referencia final se resuelve en el RPA con
            # el respaldo del DOM (ver _ejecutar_dispersion / _generar_txts_pesos).
            self._concepto_prov_por_grupo = {
                grupo: self._tablas_por_empresa[grupo].conceptos_pesos()
                for grupo in self._pesos_por_grupo
            }
            self._ref_prov_por_grupo = {
                grupo: self._tablas_por_empresa[grupo].referencias_pesos()
                for grupo in self._pesos_por_grupo
            }
            # A partir de aquí viene lo lento (TC del DOF + conciliación): se muestra
            # la pantalla de carga y se cede el control para que alcance a pintarse
            # antes del trabajo pesado.
            self._abrir_cargando_dispersion()
            mostro_cargando = True
            await asyncio.sleep(0.02)
            # Tipo de cambio (DOF) para el modal 'Solicitudes a dispersar' cuando hay
            # proveedores USD marcados 'pagar en pesos'. En un hilo (no congela la UI);
            # si falla, se guarda el error para avisarlo en el modal (no impide
            # dispersar). Siempre con la fecha del día hábil anterior (los lunes,
            # el viernes pasado); se muestra la fecha que devuelve el DOF.
            self._tc_preview = None
            self._tc_preview_fecha = None
            self._tc_preview_error = None
            if self._pesos_por_grupo:
                try:
                    self._tc_preview, self._tc_preview_fecha = (
                        await asyncio.to_thread(
                            tipo_cambio.tipo_cambio_usd_detalle))
                except Exception as exc:  # noqa: BLE001 — se reporta en el modal
                    self._tc_preview_error = str(exc)
            # Conciliación (separa por empresa, valida requeridos y cuadre por cuenta)
            # en un hilo: con muchas solicitudes puede tardar y congelaría la UI. Las
            # 'válidas' son las que irían al RPA; el payload queda guardado para la
            # operación (y para el botón 'Ver datos' del diálogo).
            self._conc_dispersion = await asyncio.to_thread(
                conciliacion.conciliar, seleccion, datos_pago)
            abrir = True
        finally:
            if mostro_cargando:
                self._cerrar_cargando_dispersion()
            self.btn_dispersar.disabled = False
            self._generando_dispersion = False
        # El diálogo real se abre DESPUÉS de cerrar la pantalla de carga (evita
        # apilar diálogos y que el cierre de carga tape al diálogo de dispersión).
        if abrir:
            self._abrir_dialogo_dispersion()

    def _abrir_cargando_dispersion(self) -> None:
        """Muestra la pantalla de carga (modal) mientras se prepara la dispersión.
        Al ser modal, bloquea la interacción con el botón mientras carga."""
        self._dlg_cargando_disp = ft.AlertDialog(
            modal=True,
            content=ft.Column(
                [
                    ft.Text("Preparando la dispersión…", size=20,
                            weight=ft.FontWeight.BOLD,
                            text_align=ft.TextAlign.CENTER),
                    ft.ProgressRing(width=32, height=32, stroke_width=3),
                ],
                spacing=20, tight=True,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER),
        )
        self.page.show_dialog(self._dlg_cargando_disp)
        self._disp_update()

    def _cerrar_cargando_dispersion(self) -> None:
        """Cierra la pantalla de carga de la preparación de la dispersión."""
        if self._dlg_cargando_disp is None:
            return
        self._dlg_cargando_disp = None
        try:
            self.page.pop_dialog()
        except Exception:  # noqa: BLE001 — el cierre no debe propagar
            pass

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

    # =============================== RPA de SUBIDA de comprobantes (Proveedores)
    def _filas_pendientes_subida(self) -> list[dict]:
        """Filas del resumen que faltan por subir (no en _subidos) y que tienen un
        comprobante vinculado con archivo existente en disco (lo único subible)."""
        pendientes: list[dict] = []
        for fila in self._folios_dispersados or []:
            if self._id_fila(fila) in self._subidos:
                continue
            ruta = self._comprobantes.get(self._id_fila(fila))
            if ruta and os.path.exists(ruta):
                pendientes.append(fila)
        return pendientes

    def _folios_sin_capturar(self, pendientes: list[dict]) -> list[str]:
        """Combinaciones empresa · cuenta origen cuyo FOLIO no se pudo capturar al
        guardar en SIPP (guardar_dispersion devolvió None).

        Sin folio no hay forma segura de ubicar su fila: empresa y cuenta origen se
        repiten entre dispersiones, así que buscar por ellas podría marcar la
        equivocada. Esas se dejan fuera del marcado y se reportan para hacerlas a
        mano, en vez de arriesgar un pago confirmado por error."""
        sin_folio = []
        for fila in pendientes:
            if fila.get("folio") is not None:
                continue
            etq = f"{fila.get('empresa') or '?'} · {fila.get('cuenta_origen') or '?'}"
            if etq not in sin_folio:
                sin_folio.append(etq)
        return sin_folio

    def _dispersiones_a_marcar(self, pendientes: list[dict]) -> list[dict]:
        """Dispersiones (una entrada por FOLIO) que hay que marcar como pagadas antes
        de subir: las de las filas `pendientes`, deduplicadas por folio.

        El FOLIO es el identificador de la dispersión: lo devuelve el propio SIPP al
        guardarla y es la primera columna de su tabla. Agrupar por él es lo correcto
        —y más fiable que empresa + cuenta origen, que se repiten entre dispersiones—:
        varias filas del resumen (una por proveedor+cuenta destino) pertenecen a UNA
        sola dispersión, y el total que muestra el SIPP es la suma de todas ellas.

        Cada entrada trae lo necesario para ubicar y VERIFICAR su fila en la tabla de
        dispersiones del SIPP:
        - `cuenta`: la Cuenta Origen elegida en SIPP (`cuenta_origen`), que es la que
          muestra esa tabla. NUNCA la cuenta en pesos del TXT: esa no entró a SIPP.
        - `total`: la suma de TODAS las filas del folio en _folios_dispersados (no solo
          las pendientes), en la moneda de la dispersión —SIPP la registró en USD
          aunque el pago se haga en pesos—, porque la tabla muestra el total completo.
        """
        totales: dict = {}
        for fila in self._folios_dispersados or []:
            folio = fila.get("folio")
            if folio is not None:
                totales[folio] = totales.get(folio, 0.0) + (fila.get("monto") or 0)
        items, vistos = [], set()
        for fila in pendientes:
            folio = fila.get("folio")
            if folio is None or folio in vistos:
                continue
            vistos.add(folio)
            items.append({
                "folio": folio,
                "empresa": fila.get("empresa") or "",
                "cuenta": fila.get("cuenta_origen") or "",
                "fecha": fila.get("fecha") or "",
                "total": round(totales.get(folio, 0.0), 2),
            })
        return items

    def _registrar_errores_marcaje(self, marcajes: list[dict]) -> None:
        """Acumula en _subida_errores las dispersiones que NO quedaron marcadas como
        pagadas. 'marcada' y 'ya_pagada' son resultados buenos y no se reportan.

        Se avisa además que su subida va a fallar: la búsqueda de comprobantes filtra
        por estatus PAGADA, así que una dispersión sin marcar no devuelve filas."""
        for m in marcajes or []:
            if m.get("estado") in ("marcada", "ya_pagada"):
                continue
            detalle = m.get("detalle") or m.get("estado") or "motivo desconocido"
            self._subida_errores.append(
                f"Dispersión {m.get('folio')} ({m.get('empresa')}): no se pudo marcar "
                f"como pagada — {detalle}. Sus comprobantes no se encontrarán.")

    def _subida_todo_ok(self) -> bool:
        """True si TODOS los movimientos del resumen tienen su comprobante subido."""
        ids = {self._id_fila(f) for f in (self._folios_dispersados or [])}
        return bool(ids) and ids <= self._subidos

    def _construir_dialogo_subida(self) -> None:
        """Construye el diálogo modal del RPA de subida de comprobantes: texto guía,
        estatus del robot y botón Detener. Arranca solo (el usuario ya consintió)."""
        mensaje = ft.Text(
            "Se subirán a SIPP los comprobantes vinculados a cada movimiento "
            "(pestaña 'Proveedores (No Pemex)').\n"
            "La operación inició automáticamente.",
            size=13)
        nota = ft.Text(
            "NOTA: No cierre esta ventana hasta que la operación termine o se "
            "detenga.",
            size=12, italic=True, color=NARANJA)
        self.btn_sub_detener = ft.OutlinedButton(
            content="Detener", icon=ft.Icons.STOP, on_click=self._sub_detener,
            style=ft.ButtonStyle(color=ROJO_BOTON), disabled=True)
        self.lbl_sub_estado = ft.Text(
            "Estatus del Robot:", size=13, weight=ft.FontWeight.BOLD,
            italic=True, color=GRIS)
        self.txt_sub_estado = ft.Text(
            "Iniciando…", size=13, weight=ft.FontWeight.BOLD, italic=True,
            color=VERDE)
        self.btn_sub_cerrar = ft.TextButton("Cerrar", on_click=self._sub_cerrar)
        self._dlg_subida = ft.AlertDialog(
            modal=True,
            title=ft.Text("Subir comprobantes de pago", weight=ft.FontWeight.BOLD),
            content=ft.Container(
                content=ft.Column(
                    [
                        mensaje, nota, ft.Divider(),
                        ft.Row([self.btn_sub_detener], spacing=10, tight=True),
                        ft.Row(
                            [self.lbl_sub_estado,
                             ft.Container(self.txt_sub_estado, expand=True)],
                            spacing=6,
                            vertical_alignment=ft.CrossAxisAlignment.START),
                    ],
                    tight=True, spacing=14,
                ),
                width=520),
            actions=[self.btn_sub_cerrar],
            actions_alignment=ft.MainAxisAlignment.END,
            on_dismiss=self._sub_on_dismiss,
        )

    def _sub_on_dismiss(self, _e=None) -> None:
        """Se dispara cuando Flutter confirma el cierre del diálogo de subida. Ejecuta el
        callback encadenado (p. ej. mostrar el resumen) sin apilar diálogos."""
        cb = self._sub_al_cerrar
        self._sub_al_cerrar = None
        if cb is not None:
            cb()

    def _cerrar_subida_luego(self, despues=None) -> None:
        """Cierra el diálogo de subida y ejecuta `despues` CUANDO Flutter confirme su
        cierre (así el resumen se muestra recién cuando la subida ya se desmontó)."""
        dlg = self._dlg_subida
        if dlg is None or not getattr(dlg, "open", False):
            if despues is not None:
                despues()
            return
        self._sub_al_cerrar = despues
        dlg.open = False
        dlg.update()

    def _abrir_dialogo_subida(self) -> None:
        """Muestra el diálogo del RPA de subida y lo ARRANCA automáticamente."""
        self._construir_dialogo_subida()
        self._sub_estado_op = "ejecutando"
        self._sub_refrescar_controles()
        self.page.show_dialog(self._dlg_subida)
        self._disp_update()
        self.page.run_task(self._sub_iniciar)

    def _sub_fijar_estado(self, texto: str, color: str) -> None:
        self.txt_sub_estado.value = texto
        self.txt_sub_estado.color = color

    def _sub_estado_seguro(self, texto: str, color: str) -> None:
        """Fija el estatus del diálogo de subida de forma segura desde el hilo RPA."""
        def aplicar():
            self._sub_fijar_estado(texto, color)
            self._disp_update()
        self._disp_en_ui(aplicar)

    def _sub_refrescar_controles(self) -> None:
        corriendo = self._sub_estado_op == "ejecutando"
        self.btn_sub_detener.disabled = not corriendo
        self.btn_sub_cerrar.disabled = corriendo
        self._disp_update()

    def _sub_cerrar(self, _e=None) -> None:
        """Cierra el diálogo de subida (solo cuando no está corriendo)."""
        if self._sub_estado_op == "ejecutando":
            return
        try:
            self.page.pop_dialog()
        except Exception:  # noqa: BLE001 — el cierre no debe propagar
            pass

    async def _sub_detener(self, _e=None) -> None:
        """Solicita detener la subida: detención cooperativa + cancelación."""
        if self._sub_estado_op != "ejecutando":
            return
        if self._sub_ctrl is not None:
            self._sub_ctrl.detener()
        if self._sub_future is not None:
            self._sub_future.cancel()
        if self._sub_task is not None:
            self._sub_task.cancel()
        self._sub_fijar_estado("Deteniendo…", NARANJA)
        self.btn_sub_detener.disabled = True
        self._disp_update()

    async def _sub_iniciar(self) -> None:
        """Corre la subida y, al terminar (éxito/detenido/error), cierra el diálogo y
        re-muestra el resumen en modo 'exito' o 'errores'."""
        self._sub_estado_op = "ejecutando"
        self._sub_fijar_estado("Iniciando…", VERDE)
        self._sub_refrescar_controles()
        self._sub_task = asyncio.create_task(self._ejecutar_subida_comprobantes())
        try:
            await self._sub_task
        except (RpaDetenido, asyncio.CancelledError):
            self._sub_estado_op = "detenido"
            self._sub_fijar_estado("Operación detenida.", ROJO)
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self._sub_estado_op = "error"
            self._sub_fijar_estado(f"Error: {exc}", ROJO)
            self.app.avisar(f"Falló la subida de comprobantes: {exc}", ROJO)
        else:
            self._sub_estado_op = "completado"
            self._sub_fijar_estado("Operación completada.", VERDE)
        finally:
            self._sub_task = None
            self._sub_future = None
            self._sub_ctrl = None
        self._sub_refrescar_controles()
        # Cerrar el diálogo de subida y, cuando se desmonte, mostrar el resumen en el
        # modo correspondiente (encadenado para no apilar diálogos).
        self._resumen_modo = "exito" if self._subida_todo_ok() else "errores"
        self._cerrar_subida_luego(self._mostrar_resumen_dispersion)

    async def _ejecutar_subida_comprobantes(self) -> None:
        """Marca las dispersiones como pagadas y sube los comprobantes.

        login → MARCAR PAGOS (pestaña 'Dispersiones (No Pemex)': una pasada por cada
        dispersión pendiente, pulsando 'Confirmar Pago') → pestaña 'Proveedores (No
        Pemex)' → estatus 'Pagado' (una vez) → por fila: buscar por proveedor +
        folio(s) de documento y adjuntar su PDF.

        El marcaje va PRIMERO porque la subida filtra por estatus PAGADA: sin marcar,
        las filas de esa dispersión no aparecen en la búsqueda. Los errores no abortan
        (se acumulan en _subida_errores). Corre en el hilo del RPA (BucleRpa)."""
        self._disp_loop_ui = asyncio.get_running_loop()
        self._subida_errores = []
        pendientes = self._filas_pendientes_subida()
        if not pendientes:
            self._sub_estado_seguro(
                "No hay comprobantes vinculados por subir.", NARANJA)
            return
        if self.bucle is None:
            self.bucle = BucleRpa()
        self._sub_ctrl = ControlRpa(self.bucle._loop)
        ctrl = self._sub_ctrl
        self.sesion = SesionSipp(headless=False)
        sesion = self.sesion
        usuario, contrasena = self.app.config.credenciales()
        total = len(pendientes)

        a_marcar = self._dispersiones_a_marcar(pendientes)
        for etq in self._folios_sin_capturar(pendientes):
            self._subida_errores.append(
                f"{etq}: no se capturó el folio al guardar la dispersión, así que no "
                "se pudo marcar como pagada. Hay que marcarla a mano en el SIPP.")

        def _prog_marcaje(hechos: int, total_m: int) -> None:
            self._sub_estado_seguro(
                f"Marcando pagos… {hechos}/{total_m}", VERDE)

        async def flujo() -> None:
            await sesion.iniciar()
            await sesion.login(usuario, contrasena)
            await sesion.seleccionar_empresa_sucursal(
                self.EMPRESA_SESION, self.SUCURSAL_SESION)
            # Marcar como pagadas las dispersiones que se van a subir. La tabla es una
            # PESTAÑA del DashboardTesor, vecina de 'Proveedores (No Pemex)':
            # marcar_pagos_dispersion entra sola, sin pasar por el alta de dispersiones.
            if a_marcar:
                await ctrl.punto_control()
                marcajes = await sesion.marcar_pagos_dispersion(
                    a_marcar, progreso=_prog_marcaje,
                    punto_control=ctrl.punto_control)
                self._registrar_errores_marcaje(marcajes)
            await sesion.ir_a_tab_proveedores_no_pemex()
            # El estatus 'Pagado' se fija UNA sola vez, antes del primer 'Buscar'.
            await sesion.seleccionar_estatus_pagado()
            for i, fila in enumerate(pendientes, start=1):
                await ctrl.punto_control()
                idf = self._id_fila(fila)
                ruta = self._comprobantes.get(idf)
                prov = fila.get("proveedor") or ""
                self._sub_estado_seguro(
                    f"Subiendo {i}/{total}: {prov}…", VERDE)
                id_prov = fila.get("id_proveedor")
                # Empresa (best-effort): acota los resultados; si no se encuentra, NO
                # bloquea la búsqueda. Se fija por fila, antes de probar sus folios.
                await sesion.seleccionar_empresa_filtro(fila.get("empresa") or "")
                folios = [f for f in (fila.get("folios_documento") or []) if f]
                if not folios and fila.get("folio_documento"):
                    folios = [fila["folio_documento"]]
                subido = False
                try:
                    for folio in folios:
                        # Se busca un folio_documento por grupo; si no da resultado,
                        # se prueba el siguiente hasta agotarlos.
                        n = await sesion.buscar_pago_proveedor(
                            id_prov, prov, str(folio))
                        if n > 0:
                            ok = await sesion.adjuntar_comprobante_en_resultado(
                                ruta, importe=fila.get("monto"))
                            if ok:
                                subido = True
                                break
                    if subido:
                        self._subidos.add(idf)
                    else:
                        self._subida_errores.append(
                            f"{prov}: no se encontró/subió el comprobante.")
                except RpaDetenido:
                    raise
                except Exception as exc:  # noqa: BLE001 — un fallo no aborta el resto
                    self._subida_errores.append(f"{prov}: {exc}")

        self._sub_future = self.bucle.enviar(flujo())
        try:
            await asyncio.wrap_future(self._sub_future)
        finally:
            self._sub_future = None
            if not self.MANTENER_NAVEGADOR_PRUEBAS:
                await self._detener_rpa()   # cierre obligatorio del navegador

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
        self._tc_fecha = None
        self._pesos_error = None
        self._comprobantes = {}
        self._comprobantes_leidos = []
        # Nueva dispersión: se reinicia el estado de la subida de comprobantes.
        self._subidos = set()
        self._subida_errores = []
        self._resumen_modo = "normal"
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
                # Se registra en cuanto guardar_dispersion REGRESA (la dispersión ya
                # quedó guardada en SIPP). El folio se DESGLOSA en una fila por
                # (proveedor, cuenta beneficiaria): es el mismo desglose en que el
                # sistema separa la dispersión al marcarla como pagada (n movimientos,
                # uno por proveedor + cuenta beneficiario).
                self._folios_dispersados.extend(
                    self._filas_resumen_de_empresa(
                        emp, folio, empresa, cuenta_origen or "",
                        _fmt_fecha(datetime.date.today())))
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
        # El layout (TXT) es por DISPERSIÓN (folio), no por movimiento: como ahora
        # _folios_dispersados trae varias filas por folio (una por proveedor+cuenta),
        # se deduplica por folio para descargar cada layout una sola vez.
        vistos: set = set()
        folios = []
        for d in self._folios_dispersados:
            f = d.get("folio")
            if f is None or f in vistos:
                continue
            vistos.add(f)
            folios.append(d)
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
            # Día hábil anterior (los lunes, el viernes pasado).
            tc, tc_fecha = tipo_cambio.tipo_cambio_usd_detalle()
        except Exception as exc:  # noqa: BLE001 — se reporta en el resumen
            self._pesos_error = str(exc)
            return
        self._tipo_cambio = tc
        self._tc_fecha = tc_fecha
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
                # Con las notas de crédito descontadas (ver total_a_pagar).
                usd = reporte_dispersion.total_a_pagar(movs)
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
                nombre = self._nombre_txt_pesos(folio_entry, clabe_dig)
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

    def _nombre_txt_pesos(self, d: dict, clabe_dig: str) -> str:
        """Nombre completo (con .txt) del layout en PESOS de una dispersión USD
        pagada en MXN: 'Pesos NNNN Folio Empresa Banco Cuenta'.

        El distintivo va al PRINCIPIO, no al final: en la carpeta del día estos
        archivos conviven con los layouts en dólares que descarga el SIPP, y llevando
        'Pesos' adelante el usuario los distingue de un vistazo y quedan todos juntos
        al ordenar por nombre. `NNNN` son los últimos 4 dígitos de la cuenta origen en
        pesos, para no colisionar cuando el mismo grupo usa varias cuentas origen."""
        distintivo = f"Pesos {clabe_dig[-4:]}" if clabe_dig else "Pesos"
        return _sanear_archivo(
            f"{distintivo} {self._nombre_txt(d)}".strip()) + ".txt"

    @staticmethod
    def _filas_resumen_de_empresa(
        emp, folio, empresa: str, cuenta_origen: str, fecha: str = "",
    ) -> list[dict]:
        """Desglosa una dispersión (empresa+moneda) en filas de resumen, UNA por
        (proveedor, cuenta beneficiaria) —el mismo corte en que el sistema separa la
        dispersión al marcarla como pagada—. Cada fila agrega el Saldo Programado del
        par y conserva, OCULTOS (no se muestran en la tabla), los nu_FolioDocumento de
        sus solicitudes y la CLABE interbancaria del beneficiario, para validaciones y
        la vinculación de comprobantes.

        `fecha` (DD/MM/AAAA) es la del registro en SIPP: se captura al guardar, no al
        usarla, para que siga siendo correcta si la subida de comprobantes se retoma
        otro día. Es la que muestra la columna 'Fecha Pago' de la tabla de
        dispersiones y sirve para verificar la fila antes de marcarla como pagada."""
        moneda = emp.movimientos[0].moneda if emp.movimientos else ""
        grupos: dict[tuple, list] = {}
        orden: list[tuple] = []
        for m in emp.movimientos:
            par = (m.proveedor, m.cuenta_bancaria)
            if par not in grupos:
                grupos[par] = []
                orden.append(par)
            grupos[par].append(m)
        filas: list[dict] = []
        for par in orden:
            prov, cuenta = par
            movs = grupos[par]
            folios_doc = [m.folio_documento for m in movs if m.folio_documento]
            beneficiarios = [
                (m.clabe_interbancaria_proveedor or m.cuenta_bancaria)
                for m in movs
                if (m.clabe_interbancaria_proveedor or m.cuenta_bancaria)]
            # id del proveedor (de la API): se usa para el select "id - nombre" del
            # RPA de subida de comprobantes (tab Proveedores No Pemex).
            id_prov = next(
                (m.id_proveedor for m in movs if m.id_proveedor is not None), None)
            filas.append({
                "folio": folio, "empresa": empresa, "clave": emp.empresa,
                "moneda": moneda, "cuenta_origen": cuenta_origen, "fecha": fecha,
                "proveedor": prov, "cuenta_destino": cuenta, "par": par,
                # Lo que de verdad se paga, ya con las notas de crédito descontadas,
                # para que cuadre con el TXT, con el total que muestra el SIPP y con
                # el importe que se busca al subir el comprobante.
                "monto": reporte_dispersion.total_a_pagar(movs),
                # --- Ocultos (no se muestran; validaciones/vinculación) ------------
                "id_proveedor": id_prov,
                "folio_documento": folios_doc[0] if folios_doc else "",
                "folios_documento": folios_doc,
                "beneficiarios": beneficiarios,
            })
        return filas

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

    # Ancho del contenido del modal de resumen (para dimensionar las tablas).
    _RESUMEN_ANCHO = 880

    def _cols_resumen(self, categoria: str) -> list[ColumnaTabla]:
        """Columnas (porcentuales) del resumen según la categoría de moneda. Cada fila
        es un (proveedor, cuenta destino), el corte con que el sistema separa la
        dispersión al pagarla. La última columna es 'Acciones'; 'usd_pesos' añade
        'Cuenta pesos' y 'Total MXN'. El Proveedor va a la izquierda (nombres largos),
        los montos a la derecha y el resto centrado.

        Empresa y Cuenta origen NO son columnas: son iguales para toda la dispersión,
        así que viven en la banda del folio (ver `_banda_folio`). Eso deja el ancho
        para lo que sí cambia fila a fila.

        'Cuenta pesos' es la Cuenta Origen EN PESOS del par: es la que nombra el TXT
        generado ('… Pesos 7045') y por tanto la que permite vincular cada fila con
        su archivo. Va aparte de 'Cuenta destino', que es la cuenta del beneficiario.

        Los porcentajes suman 99 (no 100) a propósito: así el ancho total queda
        justo por debajo del disponible y no aparece un scroll horizontal por el
        redondeo px + los gaps entre columnas (mismo criterio que _COLS_PCT)."""
        pesos = categoria == "usd_pesos"
        total_lbl = "Total MXN" if categoria == "mxn" else "Total USD"
        cols = [
            ColumnaTabla("Proveedor", 22 if pesos else 36, _TIZQ),
            ColumnaTabla("Cuenta destino", 19 if pesos else 30, CENTRO),
        ]
        if pesos:
            cols.append(ColumnaTabla("Cuenta pesos", 19, CENTRO))
        cols.append(ColumnaTabla(total_lbl, 12 if pesos else 18, _TDER))
        if pesos:
            cols.append(ColumnaTabla("Total MXN", 12, _TDER))
        cols.append(ColumnaTabla("Acciones", 15, CENTRO))
        return cols

    def _acciones_resumen(self, d: dict, categoria: str) -> ft.Control:
        """Botones de acción de una fila del resumen, según `self._resumen_modo`:
        - 'exito': solo VER (si hay comprobante); sin edición.
        - 'errores': movimiento SUBIDO (en _subidos) → solo VER; PENDIENTE → Agregar/
          Editar para adjuntar y reintentar.
        - 'normal': sin comprobante → Agregar (+); con comprobante → VER + EDITAR.
        Agregar/Editar nunca coexisten.

        Las acciones del comprobante van como ICONOS directos (son las de uso
        frecuente); las demás —regenerar el TXT en pesos y eliminar el movimiento—
        van en un menú de tres puntos, que además evita apretar la columna. En modo
        'exito' la operación ya cerró: solo consulta, sin menú."""
        modo = self._resumen_modo
        idf = self._id_fila(d)
        tiene = bool(self._comprobantes.get(idf))
        subido = idf in self._subidos
        botones: list[ft.Control] = []

        def ver() -> ft.Control:
            return ft.IconButton(
                ft.Icons.VISIBILITY, icon_size=18, tooltip="Ver comprobante",
                on_click=lambda _e, dd=d: self._ver_comprobante(dd))

        def editar(icono, tooltip) -> ft.Control:
            return ft.IconButton(
                icono, icon_size=18, tooltip=tooltip,
                on_click=lambda _e, dd=d: self.page.run_task(
                    self._vincular_comprobante_individual, dd))

        def menu() -> ft.Control:
            items = []
            if categoria == "usd_pesos":
                items.append(ft.PopupMenuItem(
                    content=ft.Text("Regenerar TXT en pesos", size=13),
                    icon=ft.Icons.AUTORENEW,
                    on_click=lambda _e, dd=d: self._dialogo_regenerar_txt(dd)))
            items.append(ft.PopupMenuItem(
                content=ft.Text("Eliminar movimiento", size=13, color=ROJO),
                icon=ft.Icons.DELETE_OUTLINE,
                on_click=lambda _e, dd=d: self._confirmar_eliminar_fila(dd)))
            return ft.PopupMenuButton(
                icon=ft.Icons.MORE_VERT, icon_size=18, tooltip="Más acciones",
                items=items)

        if modo == "exito":
            # Operación terminada con éxito: solo consulta.
            if tiene:
                botones.append(ver())
            return ft.Row(botones, spacing=0, tight=True,
                          alignment=ft.MainAxisAlignment.CENTER)
        if modo == "errores" and subido:
            # Ya subido: ver el comprobante; el resto sigue en el menú.
            botones.append(ver())
            botones.append(menu())
            return ft.Row(botones, spacing=0, tight=True,
                          alignment=ft.MainAxisAlignment.CENTER)

        # Modo normal, o pendiente en modo errores: se puede adjuntar/editar.
        if tiene:
            botones.append(ver())
            botones.append(editar(ft.Icons.EDIT, "Modificar comprobante"))
        else:
            botones.append(editar(ft.Icons.ADD, "Agregar comprobante"))
        botones.append(menu())
        return ft.Row(botones, spacing=0, tight=True,
                      alignment=ft.MainAxisAlignment.CENTER)

    def _total_folio(self, folio) -> float:
        """Total de la dispersión COMPLETA (todas sus filas), en su moneda. Es el
        importe que muestra el SIPP para ese folio, aunque en esta tabla solo se vean
        algunas de sus filas."""
        return round(sum((d.get("monto") or 0) for d in self._folios_dispersados
                         if d.get("folio") == folio), 2)

    def _banda_folio(self, d: dict, n_cols: int) -> Cabecera:
        """Banda agrupadora de una dispersión: 'Folio N · Empresa · Cuenta origen' a la
        izquierda, con el botón de editar la cuenta origen, y el TOTAL FOLIO a la
        derecha.

        El total es el del folio COMPLETO, no el de las filas visibles: un mismo folio
        puede repartir sus filas entre las tablas 'USD' y 'USD pago en MXN' (ver
        `_rango_moneda_fila`, que clasifica por FILA). Por eso se etiqueta 'TOTAL
        FOLIO', distinto del 'TOTAL' de la tabla, que sí suma solo lo que se ve."""
        folio = d.get("folio")
        empresa = str(d.get("empresa") or "—")
        cuenta = str(d.get("cuenta_origen") or "—")
        editable = self._resumen_modo != "exito"
        # Los Text llevan `expand` (no `tight`) para que el Row les fije un ancho y el
        # '…' funcione; con el ancho natural se desbordarían sobre el total.
        izq_items = [
            ft.Text(f"Folio {folio}" if folio is not None else "Sin folio",
                    size=13, weight=ft.FontWeight.BOLD, no_wrap=True,
                    tooltip=f"Folio de la dispersión en SIPP: {folio}"),
            ft.Text("·", size=13, color=GRIS),
            ft.Text(empresa, size=12, weight=ft.FontWeight.BOLD, expand=2,
                    max_lines=1, no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS,
                    tooltip=empresa),
            ft.Text("·", size=13, color=GRIS),
            ft.Text(cuenta, size=12, color=GRIS, expand=3,
                    max_lines=1, no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS,
                    tooltip=f"Cuenta origen: {cuenta}"),
        ]
        if editable:
            izq_items.append(ft.IconButton(
                ft.Icons.ACCOUNT_BALANCE, icon_size=16,
                tooltip="Cambiar la cuenta origen de esta dispersión",
                on_click=lambda _e, dd=d: self._dialogo_cuenta_origen(dd)))
        total = self._total_folio(folio)
        der = ft.Row(
            [ft.Text("TOTAL FOLIO", size=11, weight=ft.FontWeight.BOLD, color=GRIS),
             ft.Text(_fmt_moneda(total), size=13, weight=ft.FontWeight.BOLD,
                     no_wrap=True,
                     tooltip=f"Total de la dispersión {folio}: {_fmt_moneda(total)}")],
            spacing=6, tight=True, vertical_alignment=ft.CrossAxisAlignment.CENTER)
        info = ft.Row(
            [ft.Row(izq_items, spacing=6, expand=True,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER), der],
            vertical_alignment=ft.CrossAxisAlignment.CENTER)
        return Cabecera(
            [SegmentoCabecera(n_cols, info, alineacion=None,
                              padding=ft.Padding.only(left=10, right=10))],
            alto=40)

    def _tabla_resumen(self, folios: list[dict], categoria: str,
                       ancho: float) -> ft.Control:
        """Tabla de resumen (TablaResponsiva, columnas porcentuales) para una
        categoría de moneda: una BANDA por dispersión (folio) seguida de sus filas
        —una por (proveedor, cuenta destino)—, y al cierre la fila TOTAL de la tabla.

        Las filas llegan agrupadas por folio en el orden en que se dispersaron."""
        cols = self._cols_resumen(categoria)
        n_cols = len(cols)
        tabla = TablaResponsiva(
            self.page, cols, ancho_inicial=ancho, alto_fila=46)

        def negrita(texto: str) -> ft.Control:
            return ft.Text(str(texto or ""), size=12, weight=ft.FontWeight.BOLD,
                           max_lines=1, no_wrap=True,
                           overflow=ft.TextOverflow.ELLIPSIS)

        # Agrupa por folio conservando el orden de aparición.
        grupos: dict = {}
        orden: list = []
        for d in folios:
            folio = d.get("folio")
            if folio not in grupos:
                grupos[folio] = []
                orden.append(folio)
            grupos[folio].append(d)

        filas: list = []
        tot_prin, tot_mxn = 0.0, 0.0
        for folio in orden:
            del_folio = grupos[folio]
            filas.append(self._banda_folio(del_folio[0], n_cols))
            for d in del_folio:
                monto = d.get("monto") or 0
                tot_prin += monto
                celdas: list = [
                    d.get("proveedor") or "—",
                    d.get("cuenta_destino") or "—",
                ]
                if categoria == "usd_pesos":
                    celdas.append(self._cuenta_pesos_mostrar(d) or "—")
                celdas.append(_fmt_moneda(monto))
                if categoria == "usd_pesos":
                    mxn = self._monto_mxn_fila(d)
                    if mxn is not None:
                        tot_mxn += mxn
                    celdas.append(_fmt_moneda(mxn) if mxn is not None else "N/A")
                celdas.append(self._acciones_resumen(d, categoria))
                # En modo 'errores', las filas NO subidas (pendientes) se resaltan en
                # amarillo para que el usuario adjunte/reintente su comprobante.
                pendiente = (self._resumen_modo == "errores"
                             and self._id_fila(d) not in self._subidos)
                filas.append(FilaDatos(
                    celdas, bgcolor=_AMARILLO_PENDIENTE if pendiente else None))
        # Fila TOTAL de la tabla (suma de lo VISIBLE aquí, a diferencia del TOTAL
        # FOLIO de cada banda). Las celdas vacías corresponden a Cuenta destino y
        # —en 'usd_pesos'— Cuenta pesos.
        total_celdas: list = [negrita("TOTAL"), ""]
        if categoria == "usd_pesos":
            total_celdas.append("")
        total_celdas.append(negrita(_fmt_moneda(tot_prin)))
        if categoria == "usd_pesos":
            total_celdas.append(negrita(_fmt_moneda(tot_mxn) if tot_mxn else "N/A"))
        total_celdas.append("")
        filas.append(FilaDatos(total_celdas))
        tabla.set_contenido(filas)
        return tabla.control

    def _cuenta_pesos_mostrar(self, d: dict) -> str:
        """Cuenta Origen EN PESOS elegida para el par de una fila 'USD pago en MXN':
        la cuenta de la que sale el pago en pesos.

        Es la que da nombre al TXT generado ('… Pesos 7045', los últimos 4 dígitos de
        su CLABE), así que es el dato que permite vincular cada fila del resumen con
        su archivo. Si no está el texto de la cuenta, cae a la CLABE. '' si el par no
        tiene cuenta en pesos (no debería pasar: es requerida antes de dispersar)."""
        clave, par = d.get("clave") or "", d.get("par")
        return (self._cuenta_pesos_por_grupo.get(clave, {}).get(par)
                or self._clabe_pesos_por_grupo.get(clave, {}).get(par)
                or "")

    def _monto_mxn_fila(self, d: dict) -> float | None:
        """Equivalente en MXN del Saldo Programado (USD) del par: USD × T.C. del DOF
        (redondeado a 2, como el TXT en pesos). None si aún no hay T.C."""
        tc = self._tipo_cambio
        if not tc:
            return None
        return round((d.get("monto") or 0) * tc, 2)

    def _mostrar_resumen_dispersion(self) -> None:
        """Resumen final: hasta tres tablas (MXN, USD y USD pago en MXN) —solo las
        que tengan movimientos— con su total y una columna de acciones; una leyenda
        consolidada de dispersiones/archivos; un botón para cargar comprobantes; y
        los botones de continuar/terminar la operación."""
        folios = self._folios_dispersados
        resultados = self._disp_resultados_txt or []
        carpeta = self._disp_carpeta_txt
        pesos_gen = self._pesos_generados or []
        descargados = sum(1 for r in resultados if r.get("ok"))
        # Las DISPERSIONES se cuentan por folio único (cada folio son varias filas, una
        # por proveedor+cuenta); los archivos (layouts) también son por folio.
        n_dispersiones = len(
            {d.get("folio") for d in folios if d.get("folio") is not None})
        no_descargados = n_dispersiones - descargados
        hay_carpeta = bool(
            carpeta and os.path.isdir(carpeta) and (descargados or pesos_gen))
        tc = self._tipo_cambio

        # Categorías por moneda, POR FILA (proveedor+cuenta): MXN (0), USD puro (1),
        # USD pago en MXN (2, solo el par marcado 'pagar en pesos').
        mxn, usd, usd_pesos = [], [], []
        for d in folios:
            rango = self._rango_moneda_fila(d)
            (mxn if rango == 0 else usd if rango == 1 else usd_pesos).append(d)

        ancho_tabla = self._RESUMEN_ANCHO - 2 * _GUTTER_SCROLL - 24

        # La leyenda consolidada ("Se generaron X…") va en el SUBTÍTULO fijo del modal
        # (siempre visible, aunque se haga scroll) — ver `titulo` más abajo. Aquí solo
        # quedan los avisos secundarios de problemas (pueden desplazarse con el cuerpo).
        cuerpo: list[ft.Control] = []
        if no_descargados > 0:
            cuerpo.append(ft.Text(
                f"Nota: {no_descargados} archivo(s) no se pudieron recuperar.",
                size=12, color=NARANJA))
        if self._pesos_error:
            cuerpo.append(ft.Text(
                f"No se generaron los TXT en pesos: {self._pesos_error}",
                size=12, color=ROJO))

        def seccion(titulo: str, folios_cat: list[dict], categoria: str) -> None:
            cuerpo.append(ft.Text(titulo, size=13, weight=ft.FontWeight.BOLD))
            cuerpo.append(self._tabla_resumen(
                folios_cat, categoria, ancho_tabla))

        if mxn:
            seccion("Dispersiones en MXN", mxn, "mxn")
        if usd:
            seccion("Dispersiones en USD", usd, "usd")
        if usd_pesos:
            titulo = "Dispersiones en USD pago en MXN"
            if tc:
                titulo += f" (T.C. {_fmt_tc(tc)} MXN)"
            seccion(titulo, usd_pesos, "usd_pesos")

        # Título del modal (fijo): 1ª línea con el nombre y el botón de 'Carga de
        # comprobantes de pago' (+ su ícono de ayuda) a la derecha; 2ª línea con la
        # leyenda consolidada como SUBTÍTULO, que queda siempre visible aunque se haga
        # scroll en el cuerpo. El botón abre el selector; la asignación folio↔comprobante
        # está por definir.
        # El botón de carga masiva de comprobantes (+ su ícono de ayuda) se OCULTA en
        # el modo 'exito' (ya se terminó todo; no hay nada que cargar).
        if self._resumen_modo == "exito":
            controles_carga: ft.Control = ft.Container()
        else:
            controles_carga = ft.Row(
                [
                    ft.OutlinedButton(
                        "Carga de comprobantes de pago",
                        icon=ft.Icons.UPLOAD_FILE,
                        on_click=self._cargar_comprobantes),
                    ft.Icon(
                        ft.Icons.HELP_OUTLINE, size=18, color=GRIS,
                        tooltip="Subir los comprobantes generados en el "
                                "banco tras realizar los pagos.\nEstos se "
                                "vincularán automáticamente a los movimientos "
                                "generados."),
                ],
                spacing=8, tight=True,
                vertical_alignment=ft.CrossAxisAlignment.CENTER)
        titulo = ft.Column(
            [
                ft.Row(
                    [
                        ft.Text("Resumen de la dispersión", weight=ft.FontWeight.BOLD),
                        controles_carga,
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER),
                self._subtitulo_resumen(n_dispersiones, descargados),
            ],
            tight=True, spacing=4)

        # Botones del modal según el modo. En 'exito' solo queda 'Cerrar' (cierra el
        # resumen y el navegador). En 'normal'/'errores': Abrir carpeta · Terminar ·
        # Continuar operación.
        if self._resumen_modo == "exito":
            acciones: list[ft.Control] = [
                ft.Container(),
                ft.FilledButton(
                    "Cerrar", icon=ft.Icons.CHECK_CIRCLE_OUTLINE,
                    on_click=lambda _e: self._dialogo_terminar_operacion()),
            ]
        else:
            izquierda: ft.Control = (
                ft.FilledButton(
                    "Abrir carpeta", icon=ft.Icons.FOLDER_OPEN,
                    on_click=lambda _e: self._abrir_archivo(carpeta))
                if hay_carpeta else ft.Container())
            grupo_derecha = ft.Row(
                [
                    ft.OutlinedButton(
                        "Terminar operación", icon=ft.Icons.STOP_CIRCLE_OUTLINED,
                        on_click=lambda _e: self._dialogo_terminar_operacion()),
                    ft.FilledButton(
                        "Continuar operación", icon=ft.Icons.PLAY_ARROW,
                        on_click=lambda _e: self._dialogo_continuar_operacion()),
                ],
                spacing=10, tight=True)
            acciones = [izquierda, grupo_derecha]

        # El contenido va con padding derecho para la barra de scroll vertical.
        contenido = ft.Container(
            content=ft.Column(
                [ft.Container(
                    ft.Column(cuerpo, tight=True, spacing=10),
                    padding=ft.Padding.only(right=_GUTTER_SCROLL))],
                tight=True, scroll=ft.ScrollMode.AUTO),
            width=self._RESUMEN_ANCHO, height=560)

        # Si ya hay un resumen ABIERTO, se ACTUALIZA EN SITIO (título/contenido/botones)
        # en vez de cerrar+reabrir: así no se apilan diálogos de resumen (cerrar+mostrar
        # dejaba el viejo montado debajo, porque Flet solo desmonta el de más arriba).
        dlg = self._dlg_resumen
        if dlg is not None and getattr(dlg, "open", False):
            dlg.title = titulo
            dlg.content = contenido
            dlg.actions = acciones
            dlg.update()
            return
        dlg = ft.AlertDialog(
            modal=True,
            title=titulo,
            content=contenido,
            actions=acciones,
            actions_alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
            on_dismiss=self._resumen_on_dismiss,
        )
        self._dlg_resumen = dlg
        self.page.show_dialog(dlg)

    def _resumen_on_dismiss(self, _e=None) -> None:
        """Se dispara cuando Flutter confirma el cierre del resumen. Ejecuta el
        callback pendiente (encadenado tras cerrar la confirmación), si lo hay."""
        cb = self._resumen_al_cerrar
        self._resumen_al_cerrar = None
        if self._dlg_resumen is not None and not self._dlg_resumen.open:
            self._dlg_resumen = None
        if cb is not None:
            cb()

    def _cerrar_resumen_luego(self, despues=None) -> None:
        """Cierra el resumen y ejecuta `despues` CUANDO Flutter confirme su cierre. Se
        usa desde el on_dismiss de la confirmación (ahí el resumen ya es el diálogo de
        más arriba y sí se puede desmontar)."""
        dlg = self._dlg_resumen
        if dlg is None or not getattr(dlg, "open", False):
            if despues is not None:
                despues()
            return
        self._resumen_al_cerrar = despues
        dlg.open = False
        dlg.update()

    def _subtitulo_resumen(self, n_dispersiones: int, descargados: int) -> ft.Control:
        """Subtítulo fijo del resumen según `self._resumen_modo`:
        - 'exito'  → "La dispersión se ha realizado con éxito" (VERDE).
        - 'errores'→ "Ocurrieron errores al subir los comprobantes" (NARANJA).
        - 'normal' → leyenda de dispersiones/archivos generados (GRIS)."""
        if self._resumen_modo == "exito":
            return ft.Text("La dispersión se ha realizado con éxito",
                           size=13, weight=ft.FontWeight.BOLD, color=VERDE)
        if self._resumen_modo == "errores":
            return ft.Text("Ocurrieron errores al subir los comprobantes",
                           size=13, weight=ft.FontWeight.BOLD, color=NARANJA)
        # Los archivos se acotan al número de dispersiones: si el usuario eliminó del
        # resumen la última fila de un folio, ese folio ya no se cuenta pero su
        # descarga sí sigue en _disp_resultados_txt (que no guarda el folio y no se
        # puede reconciliar). Sin el tope saldrían absurdos tipo '1 dispersión y 2
        # archivos'.
        return ft.Text(
            f"Se generaron {n_dispersiones} dispersiones y se recuperaron "
            f"{min(descargados, n_dispersiones)} archivos.",
            size=13, weight=ft.FontWeight.W_500, color=GRIS)

    def _mostrar_resumen_ejemplo(self, _e=None) -> None:
        """PRUEBAS: rellena el estado con datos de ejemplo y abre el modal de resumen,
        para revisar el diseño y dar retro sin ejecutar una dispersión real.

        El ejemplo cubre los tres casos que cambian el layout: un folio con DOS
        beneficiarios (para ver la banda agrupando), un SEGUNDO folio de la misma
        empresa y cuenta origen (para comprobar que no se mezclan) y una entrada 'USD
        pago en MXN' (columna Cuenta pesos + Total MXN + regenerar TXT)."""
        def fila(folio, empresa, clave, moneda, origen, par, monto, docs, id_prov):
            # `docs`: lista de nu_FolioDocumento del grupo (uno por MOVIMIENTO). El
            # grupo (proveedor+cuenta) muestra 1 fila agregada, pero conserva todos sus
            # folios de documento (para la búsqueda por folio en el filtro del SIPP) y
            # el id del proveedor (para el select 'id - nombre' del RPA de subida).
            return {
                "folio": folio, "empresa": empresa, "clave": clave, "moneda": moneda,
                "cuenta_origen": origen, "fecha": _fmt_fecha(datetime.date.today()),
                "proveedor": par[0], "cuenta_destino": par[1],
                "par": par, "monto": monto, "id_proveedor": id_prov,
                "folio_documento": docs[0] if docs else "",
                "folios_documento": list(docs),
                "beneficiarios": [par[1]]}

        # --- REAL: última dispersión (folio 286, ACP Combustibles, MXN) ----------
        # Un folio con DOS beneficiarios (TRION y VALERO): así se ve el desglose por
        # proveedor + cuenta. Los id_proveedor son PLACEHOLDER: para una prueba
        # end-to-end real deben ser los reales de la API.
        origen_acp = "BBVA BANCOMER 0104728025 ACP COMBUSTIBLES"
        clave_acp = "ACP Combustibles - MXN"
        par_trion = ("TRION CORPORATION FUEL AND GAS",
                     "BANORTE - 072580012454559986")
        par_valero = ("VALERO MARKETING AND SUPPLY DE MEXICO",
                      "BANCO 110 - 110180000776465174")

        # Segundo folio de la MISMA empresa y cuenta origen: comprueba que las bandas
        # no se fusionan y que cada una lleva su propio TOTAL FOLIO.
        par_otro = ("TRANSPORTES Y EQUIPOS DEL NOROESTE",
                    "BANREGIO - 058744000012345678")
        # Entrada USD pago en MXN (otra empresa/moneda), para ver esa tabla completa.
        origen_ps = "BANREGIO NAVOJOA DLLS"
        clave_ps = "Petro Smart - USD"
        par_wind = ("ALMACENADORA DE GAS WINDSTAR, S.A. DE C.V.",
                    "SANTANDER - 014164655049324435")

        self._folios_dispersados = [
            fila("286", "ACP Combustibles", clave_acp, "MXN", origen_acp,
                 par_trion, 1523170.19, ["146"], 1146),
            fila("286", "ACP Combustibles", clave_acp, "MXN", origen_acp,
                 par_valero, 2404645.38,
                 ["3657735499", "3657735500", "3657735451"], 3657),
            fila("287", "ACP Combustibles", clave_acp, "MXN", origen_acp,
                 par_otro, 84200.50, ["4471"], 4471),
            fila("762", "Petro Smart", clave_ps, "USD", origen_ps,
                 par_wind, 25122.73, ["B38268", "B38269"], 7620),
        ]
        # 3 dispersiones (folios 286, 287 y 762) con su layout recuperado.
        self._disp_resultados_txt = [{"ok": True}, {"ok": True}, {"ok": True}]
        self._disp_carpeta_txt = None
        # El par de Petro Smart va marcado 'pagar en pesos': así aparece la tabla
        # 'USD pago en MXN' con sus columnas extra.
        self._pesos_por_grupo = {clave_ps: {par_wind}}
        self._cuenta_pesos_por_grupo = {
            clave_ps: {par_wind: "BBVA BANCOMER 012744001230117045 PETRO SMART"}}
        self._clabe_pesos_por_grupo = {clave_ps: {par_wind: "012744001230117045"}}
        self._concepto_prov_por_grupo = {clave_ps: {par_wind: "PAGO FACT PSC NA"}}
        self._ref_prov_por_grupo = {clave_ps: {par_wind: "B38268"}}
        self._pesos_generados = []
        self._tipo_cambio = 17.2195
        self._tc_fecha = _fecha_tc_texto()
        self._pesos_error = None
        self._comprobantes = {}
        # Estado de subida en limpio (el ejemplo siempre abre en modo normal).
        self._subidos = set()
        self._subida_errores = []
        self._resumen_modo = "normal"
        self._mostrar_resumen_dispersion()

    # -------------------------------------- acciones del resumen
    @staticmethod
    def _id_fila(d: dict) -> str:
        """Clave estable de una fila del resumen: folio + proveedor + cuenta destino.
        Identifica el MOVIMIENTO al que se vincula un comprobante, ya que un folio se
        abre en varias filas (una por proveedor + cuenta beneficiaria) y no puede
        distinguirse solo por el folio."""
        prov, cuenta = d.get("par") or (d.get("proveedor"), d.get("cuenta_destino"))
        return f"{d.get('folio')}‖{prov}‖{cuenta}"

    def _ver_comprobante(self, d: dict) -> None:
        """Abre el comprobante vinculado a una fila del resumen (si hay)."""
        ruta = self._comprobantes.get(self._id_fila(d))
        if not ruta:
            self.app.avisar(
                "Este movimiento aún no tiene un comprobante vinculado.", NARANJA)
        elif os.path.exists(ruta):
            self._abrir_archivo(ruta)
        else:
            self.app.avisar(
                "El comprobante está vinculado, pero no se encontró el archivo de "
                "origen en disco.", NARANJA)

    # ------------------------------------ vinculación comprobante ↔ dispersión
    def _objetivo_vinculacion(self, fila: dict) -> dict:
        """Datos normalizados de UNA fila del resumen (par proveedor+cuenta) para casar
        un comprobante. ABSTRAÍDO para reutilizarlo también en la vinculación masiva:
          - `origenes`: cuenta(s) origen. En 'USD pago en MXN' es la cuenta origen EN
            PESOS elegida para ese par; si no, la cuenta origen de la dispersión.
          - `beneficiarios`: CLABE(s) interbancaria(s) del proveedor de la fila.
          - `total`: total a casar. En 'USD pago en MXN' es el total EN MXN del par
            (Saldo Programado USD × T.C.).
        """
        clave = fila.get("clave") or ""
        par = fila.get("par")
        es_pesos = bool(par) and par in self._pesos_por_grupo.get(clave, set())
        beneficiarios = set(fila.get("beneficiarios") or [])
        if not beneficiarios and fila.get("cuenta_destino"):
            beneficiarios = {fila["cuenta_destino"]}
        if es_pesos:
            # Cuenta origen EN PESOS del par (1.1) y total EN MXN (3.1).
            origenes = set()
            cta = self._cuenta_pesos_por_grupo.get(clave, {}).get(par)
            clabe = self._clabe_pesos_por_grupo.get(clave, {}).get(par)
            origenes.update(o for o in (cta, clabe) if o)
            total = round((fila.get("monto") or 0) * (self._tipo_cambio or 0), 2)
        else:
            origenes = {fila.get("cuenta_origen") or ""}
            total = fila.get("monto") or 0
        return {
            "origenes": {o for o in origenes if o},
            "beneficiarios": {b for b in beneficiarios if b},
            "total": float(total or 0),
            "es_pesos": es_pesos,
        }

    def _comprobante_coincide(self, comprobante: dict, folio_dict: dict) -> dict:
        """Evalúa las 3 reglas de vinculación de un comprobante (dict de la respuesta
        del extractor) contra una dispersión (folio). REUTILIZABLE para carga
        individual y masiva. Devuelve {'origen','beneficiario','total','coincide'}
        (bools); 'coincide' = las tres reglas se cumplen.

        Las cuentas del comprobante vienen enmascaradas (p. ej. '*0012'), así que la
        comparación es por los ÚLTIMOS dígitos: cuenta origen (regla 1) y CLABE
        interbancaria del beneficiario vs cuenta destino (regla 2). El total (regla 3)
        se compara con tolerancia de centavos."""
        obj = self._objetivo_vinculacion(folio_dict)
        origen_comp = _ultimos_digitos(comprobante.get("cuenta_origen"))
        origen_ok = bool(origen_comp) and any(
            _ultimos_digitos(o) == origen_comp for o in obj["origenes"])
        destino_comp = _ultimos_digitos(comprobante.get("cuenta_destino"))
        benef_ok = bool(destino_comp) and any(
            _ultimos_digitos(b) == destino_comp for b in obj["beneficiarios"])
        importe = comprobante.get("importe")
        total_ok = importe is not None and abs(
            float(importe) - obj["total"]) < 0.01
        return {
            "origen": origen_ok, "beneficiario": benef_ok, "total": total_ok,
            "coincide": origen_ok and benef_ok and total_ok,
        }

    async def _vincular_comprobante_individual(self, fila: dict) -> None:
        """Carga INDIVIDUAL: pide UN PDF, lo lee en el extractor, valida las reglas
        contra ESTE movimiento (par proveedor+cuenta) y lo vincula. Si no coincide del
        todo, pide confirmación antes de adjuntar (regla 4)."""
        archivos = await self.app.picker.pick_files(
            dialog_title="Selecciona el comprobante (PDF)",
            allowed_extensions=["pdf"], allow_multiple=False)
        if not archivos:
            return
        ruta = archivos[0].path
        if not ajustes_api.base_url_extractor():
            self.app.avisar(
                "Configura la URL de la API extractor en Configuración.", NARANJA)
            return
        # Lee el PDF en el extractor (con un spinner mientras responde).
        spinner = ft.AlertDialog(
            modal=True,
            content=ft.Column(
                [ft.ProgressRing(width=32, height=32, stroke_width=3),
                 ft.Text("Leyendo comprobante…", size=14)],
                spacing=16, tight=True,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER))
        self.page.show_dialog(spinner)
        self._disp_update()
        error = None
        resp = None
        try:
            resp = await asyncio.to_thread(api.leer_comprobantes_pagos, [ruta])
        except api.ErrorApi as exc:
            error = str(exc)
        finally:
            self.page.pop_dialog()  # spinner
        if error:
            self.app.avisar(f"No se pudo leer el comprobante: {error}", ROJO)
            return
        comprobantes = ((resp or {}).get("data") or {}).get("comprobantes") or []
        if not comprobantes:
            self.app.avisar("El comprobante no devolvió datos legibles.", NARANJA)
            return
        # Coincide si ALGÚN comprobante leído del PDF casa con el movimiento.
        coincide = any(
            self._comprobante_coincide(c, fila).get("coincide")
            for c in comprobantes)
        if coincide:
            self._adjuntar_comprobante(fila, ruta)
            return
        # Regla 4: no coincide del todo -> confirmar antes de adjuntar.
        def aceptar(_e=None) -> None:
            self.page.pop_dialog()  # confirmación
            self._adjuntar_comprobante(fila, ruta)

        self.page.show_dialog(ft.AlertDialog(
            modal=True,
            title=ft.Text("Comprobante no coincide", weight=ft.FontWeight.BOLD),
            content=ft.Text(
                "Se detectó que el comprobante no coincide totalmente con el "
                "movimiento.\n¿Adjuntar de todas formas?"),
            actions=[
                ft.TextButton("Cancelar", on_click=lambda _e: self.page.pop_dialog()),
                ft.FilledButton("Adjuntar", on_click=aceptar),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        ))

    def _adjuntar_comprobante(self, fila: dict, ruta: str) -> None:
        """Vincula el archivo `ruta` a ESE movimiento (fila del resumen) y refresca el
        resumen (los íconos de la columna 'Acciones' pasan a Ver/Editar)."""
        self._comprobantes[self._id_fila(fila)] = ruta
        # Re-render EN SITIO del resumen (no cerrar+reabrir: evita apilar diálogos).
        self._mostrar_resumen_dispersion()
        prov = fila.get("proveedor") or f"folio {fila.get('folio')}"
        self.app.avisar(f"Comprobante vinculado a {prov}.", VERDE)

    async def _cargar_comprobantes(self, _e=None) -> None:
        """Sube al extractor los PDFs de comprobantes elegidos por el usuario, en
        LOTES de 10, mostrando una barra de progreso ((lotes_procesados/total)*100).
        Acumula los comprobantes leídos en self._comprobantes_leidos (para la
        asignación a folios, por definir) y avisa un resumen al terminar."""
        # 1) Selección de PDFs (multi-archivo).
        archivos = await self.app.picker.pick_files(
            dialog_title="Selecciona los comprobantes de pago (PDF)",
            allowed_extensions=["pdf"], allow_multiple=True)
        if not archivos:
            return
        rutas = [a.path for a in archivos]
        # 2) La API extractor debe estar configurada (URL). Si no, se avisa.
        if not ajustes_api.base_url_extractor():
            self.app.avisar(
                "Configura la URL de la API extractor en Configuración.", NARANJA)
            return
        # 3) Lotes de 10 archivos.
        _TAM_LOTE = 10
        lotes = [rutas[i:i + _TAM_LOTE] for i in range(0, len(rutas), _TAM_LOTE)]
        total = len(lotes)

        # 4) Modal con barra de progreso. El ProgressRing es INDETERMINADO: gira de
        # forma continua en el cliente (Flutter) aunque un lote tarde, para que el
        # usuario vea siempre actividad y no crea que la app se congeló; la
        # ProgressBar (determinada) muestra el avance por lotes.
        anillo = ft.ProgressRing(width=34, height=34, stroke_width=4)
        barra = ft.ProgressBar(value=0, width=360, bar_height=14)
        texto = ft.Text(f"Procesando lote 0 de {total}… (0%)", size=13, color=GRIS)
        dlg = ft.AlertDialog(
            modal=True,
            content=ft.Column(
                [anillo,
                 ft.Text("Leyendo comprobantes…", size=20,
                         weight=ft.FontWeight.BOLD, text_align=ft.TextAlign.CENTER),
                 texto, barra],
                spacing=16, tight=True,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER))
        self.page.show_dialog(dlg)
        self._disp_update()

        # 5) Subida por lotes; tras cada await se refresca el progreso.
        comprobantes: list[dict] = []
        fallidos: list = []
        errores: list[str] = []
        try:
            for i, lote in enumerate(lotes, start=1):
                try:
                    resp = await asyncio.to_thread(
                        api.leer_comprobantes_pagos, lote)
                    data = (resp or {}).get("data") or {}
                    comprobantes.extend(data.get("comprobantes") or [])
                    fallidos.extend((resp or {}).get("failedResults") or [])
                except api.ErrorApi as exc:
                    errores.append(str(exc))
                barra.value = i / total
                texto.value = (
                    f"Procesando lote {i} de {total}… ({int(i / total * 100)}%)")
                self._disp_update()
        finally:
            self.page.pop_dialog()

        # 6) Guarda lo leído.
        self._comprobantes_leidos = comprobantes
        n_ok = len(comprobantes)
        n_fallidos = len(fallidos)

        # 7) Vinculación MASIVA: asigna cada comprobante a la dispersión (fila) que le
        # corresponde, con las MISMAS reglas del alta individual (origen + beneficiario
        # + total), pero SIN el diálogo de confirmación (ese es solo para carga
        # individual). Refresca el resumen si hubo vínculos nuevos.
        n_vinc, n_sin_disp, n_sin_arch = self._vincular_comprobantes_masivo(
            comprobantes, rutas)
        if n_vinc:
            self._refrescar_resumen_dispersion()

        # 8) Resumen del resultado (lectura + vinculación).
        detalle = []
        if n_vinc:
            detalle.append(f"{n_vinc} vinculado(s)"
                           + (f", {n_sin_arch} sin archivo" if n_sin_arch else ""))
        if n_sin_disp:
            detalle.append(f"{n_sin_disp} sin dispersión que coincida")
        vinc_txt = ("Vinculación: " + "; ".join(detalle) + "."
                    if detalle else "Ningún comprobante coincidió con una dispersión.")
        if errores:
            self.app.avisar(
                f"Se leyeron {n_ok} comprobante(s); {len(errores)} lote(s) "
                f"fallaron: {errores[0]}. {vinc_txt}", ROJO)
        elif n_fallidos:
            self.app.avisar(
                f"Se leyeron {n_ok} comprobante(s); {n_fallidos} ilegible(s). "
                f"{vinc_txt}", NARANJA)
        elif n_ok:
            color = VERDE if n_vinc and not n_sin_disp else NARANJA
            self.app.avisar(
                f"Se leyeron {n_ok} comprobante(s). {vinc_txt}", color)
        else:
            self.app.avisar("No se leyó ningún comprobante.", NARANJA)

    def _vincular_comprobantes_masivo(
        self, comprobantes: list[dict], rutas: list[str],
    ) -> tuple[int, int, int]:
        """Asigna cada comprobante leído a la fila del resumen (dispersión) que le
        corresponde, con las MISMAS reglas del alta individual (_comprobante_coincide),
        SIN el diálogo de confirmación. El archivo de origen se resuelve por el campo
        'documento_lectura' del comprobante contra los PDFs subidos (nombre exacto y
        normalizado). Cada fila se vincula a lo sumo con UN comprobante.

        Devuelve (vinculados, sin_dispersion, sin_archivo):
          - vinculados: comprobantes casados con una fila (incluye los sin_archivo).
          - sin_dispersion: no casó con ninguna fila.
          - sin_archivo: casó, pero no se pudo resolver el PDF de origen (se registra el
            vínculo igual; 'Ver' no podrá abrirlo hasta corregir el nombre en la API)."""
        filas = self._folios_dispersados or []
        ruta_exacta = {os.path.basename(r): r for r in rutas}
        ruta_norm = {_norm_nombre_doc(os.path.basename(r)): r for r in rutas}
        vinculados = sin_dispersion = sin_archivo = 0
        ocupadas = set(self._comprobantes)  # filas que ya tienen comprobante
        for c in comprobantes:
            fila = next(
                (f for f in filas
                 if self._id_fila(f) not in ocupadas
                 and self._comprobante_coincide(c, f).get("coincide")),
                None)
            if fila is None:
                sin_dispersion += 1
                continue
            nombre = (c.get("documento_lectura") or "").strip()
            ruta = ruta_exacta.get(nombre) or ruta_norm.get(_norm_nombre_doc(nombre))
            if ruta is None:
                sin_archivo += 1
                ruta = nombre or "(archivo no resuelto)"
            self._comprobantes[self._id_fila(fila)] = ruta
            ocupadas.add(self._id_fila(fila))
            vinculados += 1
        return vinculados, sin_dispersion, sin_archivo

    def _refrescar_resumen_dispersion(self) -> None:
        """Refresca EN SITIO el modal de resumen para reflejar los comprobantes recién
        vinculados (íconos de la columna 'Acciones'), sin cerrar+reabrir (no apila)."""
        self._mostrar_resumen_dispersion()

    def _confirmar_eliminar_fila(self, d: dict) -> None:
        """Pide confirmación antes de quitar un movimiento del resumen. Es destructivo
        (borra su comprobante vinculado) y devuelve sus solicitudes a la tabla."""
        prov = d.get("proveedor") or "—"
        cuenta = d.get("cuenta_destino") or "—"
        folio = d.get("folio")

        def eliminar(_e=None) -> None:
            self.page.pop_dialog()
            self._eliminar_fila_resumen(d)

        self.page.show_dialog(ft.AlertDialog(
            modal=True,
            title=ft.Text("Eliminar movimiento"),
            content=ft.Column(
                [ft.Text(f"Se quitará del resumen del folio {folio}:", size=13),
                 ft.Text(f"{prov} · {cuenta}", size=13,
                         weight=ft.FontWeight.BOLD),
                 ft.Text(
                     "Sus solicitudes volverán a la tabla para poder dispersarlas de "
                     "nuevo, y el total del folio bajará. En el SIPP la dispersión "
                     "sigue como está: hay que corregirla ahí por separado.",
                     size=12, color=NARANJA)],
                spacing=8, tight=True),
            actions=[
                ft.TextButton("Cancelar", on_click=lambda _e: self.page.pop_dialog()),
                ft.FilledButton("Eliminar", icon=ft.Icons.DELETE_OUTLINE,
                                on_click=eliminar),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        ))

    def _eliminar_fila_resumen(self, d: dict) -> None:
        """Quita un movimiento del resumen y libera sus solicitudes.

        Pasos: saca la fila de `_folios_dispersados`, limpia el estado que se indexa
        por ella (comprobante, subido) y por su PAR dentro del grupo (solo si ninguna
        otra fila lo usa), regenera el TXT en pesos si aplica —para que el archivo ya
        no la incluya— y devuelve sus solicitudes a la tabla de la empresa.
        """
        idf = self._id_fila(d)
        clave = d.get("clave") or ""
        par = d.get("par") or (d.get("proveedor"), d.get("cuenta_destino"))
        era_pesos = self._rango_moneda_fila(d) == 2

        self._folios_dispersados = [
            f for f in self._folios_dispersados if self._id_fila(f) != idf]
        self._comprobantes.pop(idf, None)
        self._subidos.discard(idf)

        # El par solo se limpia si NINGUNA otra fila del mismo grupo lo usa (un grupo
        # puede haberse dispersado en más de un folio).
        sigue_usado = any(
            (f.get("clave") or "") == clave
            and (f.get("par") or (f.get("proveedor"), f.get("cuenta_destino"))) == par
            for f in self._folios_dispersados)
        if not sigue_usado:
            pares = self._pesos_por_grupo.get(clave)
            if pares is not None:
                pares.discard(par)
                if not pares:
                    self._pesos_por_grupo.pop(clave, None)
            for dic in (self._clabe_pesos_por_grupo, self._cuenta_pesos_por_grupo,
                        self._concepto_prov_por_grupo, self._ref_prov_por_grupo,
                        self._ref_dom_por_grupo):
                por_par = dic.get(clave)
                if por_par is not None:
                    por_par.pop(par, None)
                    if not por_par:
                        dic.pop(clave, None)

        # El TXT en pesos agrupa varios pares por cuenta origen: hay que reescribirlo
        # sin este movimiento. Se le pasan TODOS los pares que quedan porque
        # _regenerar_txt_pesos reemplaza los dicts del grupo completos, no fusiona.
        aviso_txt = ""
        if era_pesos and not sigue_usado:
            restantes = sorted(self._pesos_por_grupo.get(clave) or set())
            if restantes:
                clabes = self._clabe_pesos_por_grupo.get(clave, {})
                conceptos = self._concepto_prov_por_grupo.get(clave, {})
                refs = self._ref_prov_por_grupo.get(clave, {})
                ok, msg = self._regenerar_txt_pesos(
                    d,
                    {p: clabes.get(p, "") for p in restantes},
                    {p: conceptos.get(p, "") for p in restantes},
                    {p: refs.get(p, "") for p in restantes})
                aviso_txt = f" {msg}" if not ok else " Se regeneró el TXT en pesos."
            else:
                # Sin pares en pesos: se borran los archivos del grupo.
                for p in [x for x in self._pesos_generados
                          if x.get("empresa") == clave]:
                    try:
                        os.remove(p.get("archivo") or "")
                    except OSError:
                        pass
                    self._pesos_generados.remove(p)
                aviso_txt = " Se eliminó el TXT en pesos del grupo."

        liberados = self._liberar_movimientos(clave, par)
        detalle = (f" Se liberaron {liberados} solicitud(es)." if liberados
                   else " No se pudieron liberar sus solicitudes (sin conciliación).")
        self.app.avisar("Movimiento eliminado del resumen." + detalle + aviso_txt,
                        VERDE if liberados else NARANJA)

        if not self._folios_dispersados:
            # Sin filas el modal quedaría vacío: se cierra en vez de dejar el hueco.
            self._cerrar_resumen_luego(None)
            return
        self._mostrar_resumen_dispersion()

    def _liberar_movimientos(self, clave: str, par: tuple) -> int:
        """Devuelve a la tabla de su empresa las solicitudes del `par` dentro del
        grupo `clave`, para poder volver a dispersarlas. Devuelve cuántas liberó.

        Los movimientos originales solo viven en la conciliación; el filtro por par es
        el mismo de `_filas_resumen_de_empresa` y `_regenerar_txt_pesos`. Se reusa
        `volcar_reportes`, que recrea la tabla del grupo si había quedado vacía y se
        retiró del árbol, y evita duplicados."""
        conc = self._conc_dispersion
        if conc is None:
            return 0
        emp = next((e for e in conc.validas if e.empresa == clave), None)
        if emp is None:
            return 0
        prov, cuenta = par
        movs = [m for m in emp.movimientos
                if m.proveedor == prov and m.cuenta_bancaria == cuenta]
        if not movs:
            return 0
        self.volcar_reportes(movs)
        return len(movs)

    def _dialogo_cuenta_origen(self, folio_dict: dict) -> None:
        """Modal para cambiar la Cuenta Origen de una dispersión (por FOLIO).

        Es la cuenta con la que la dispersión quedó registrada en SIPP. Cambiarla aquí
        actualiza el registro de la app; corregirla en SIPP es aparte (ver
        `_aplicar_cuenta_origen`). Las opciones son TODAS las cuentas de la empresa
        —`_cuentas_de_empresa`, igual que el selector del encabezado de tabla—, no las
        de `_clabes_de_empresa`, que están filtradas a las que sirven para el TXT en
        pesos y son otra cosa."""
        folio = folio_dict.get("folio")
        empresa = folio_dict.get("empresa") or ""   # nombre limpio, NUNCA 'clave'
        actual = folio_dict.get("cuenta_origen") or ""
        opciones = self._cuentas_de_empresa(empresa)
        if not opciones:
            self.app.avisar(
                f"No hay cuentas de dispersión cargadas para «{empresa}». Revisa el "
                "catálogo de cuentas en Configuración.", NARANJA)
            return
        dd = ft.Dropdown(
            label=_label_requerido("Cuenta Bancaria Origen"), width=420,
            enable_filter=True, editable=True,
            value=actual if actual in opciones else None,
            options=[ft.dropdown.Option(key=c, text=c) for c in opciones])

        def guardar(_e=None) -> None:
            nueva = (dd.value or "").strip()
            if not nueva:
                self.app.avisar("Elige una cuenta origen.", NARANJA)
                return
            if nueva == actual:
                self.page.pop_dialog()
                return
            n = self._aplicar_cuenta_origen(folio, nueva)
            self.page.pop_dialog()
            self.app.avisar(
                f"Cuenta origen del folio {folio} actualizada en {n} movimiento(s). "
                "Recuerda cambiarla también en el SIPP.", VERDE)
            self._mostrar_resumen_dispersion()

        self.page.show_dialog(ft.AlertDialog(
            modal=True,
            title=ft.Text(f"Cuenta origen — folio {folio}",
                          weight=ft.FontWeight.BOLD),
            content=ft.Container(
                content=ft.Column(
                    [ft.Text(f"Actual: {actual or '—'}", size=12, color=GRIS),
                     dd,
                     ft.Text(
                         "El cambio aplica a todos los movimientos de esta "
                         "dispersión. En el SIPP hay que actualizarla por separado; "
                         "mientras no coincidan, el robot no confirmará su pago.",
                         size=11, color=NARANJA)],
                    spacing=12, tight=True),
                width=460, height=170),
            actions=[
                ft.TextButton("Cancelar", on_click=lambda _e: self.page.pop_dialog()),
                ft.FilledButton("Guardar", icon=ft.Icons.SAVE, on_click=guardar),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        ))

    def _aplicar_cuenta_origen(self, folio, cuenta: str) -> int:
        """Fija `cuenta` como Cuenta Origen en TODAS las filas de `folio` y anota el
        cambio como pendiente de reflejar en SIPP. Devuelve cuántas filas cambió.

        Es el punto de enganche para automatizar la edición en el portal: cuando se
        implemente, basta con leer `_cuentas_origen_pendientes_sipp` y disparar el
        paso RPA sobre esos folios."""
        n = 0
        for d in self._folios_dispersados:
            if d.get("folio") == folio:
                d["cuenta_origen"] = cuenta
                n += 1
        if n:
            self._cuentas_origen_pendientes_sipp[folio] = cuenta
        return n

    def _dialogo_regenerar_txt(self, folio_dict: dict) -> None:
        """Modal para regenerar los TXT en pesos de una dispersión USD pago en MXN.
        Como la Cuenta Origen se elige POR PROVEEDOR, muestra una fila por cada par
        (proveedor · cuenta) con su propio selector de Cuenta Origen (mismas opciones
        que en 'pagar en pesos') + concepto y referencia editables. Al regenerar,
        reescribe los TXT (mismo nombre y carpeta que los originales, reemplazándolos:
        uno por cuenta origen, como en la generación original)."""
        clave = folio_dict.get("clave") or ""
        empresa = folio_dict.get("empresa") or ""
        pares = sorted(self._pesos_por_grupo.get(clave) or set())
        if not pares:
            self.app.avisar(
                "Esta dispersión no tiene proveedores marcados 'pagar en pesos'.",
                NARANJA)
            return
        opciones = self._clabes_de_empresa(empresa)
        clabes_prev = self._clabe_pesos_por_grupo.get(clave, {})
        concep_prev = self._concepto_prov_por_grupo.get(clave, {})
        refs_prev = self._ref_prov_por_grupo.get(clave, {})

        # Un bloque por par (proveedor · cuenta): selector de Cuenta Origen + concepto
        # + referencia. Se guardan las referencias de los controles por par.
        controles: dict[tuple, tuple] = {}
        bloques: list[ft.Control] = []
        for par in pares:
            prov, cuenta = par
            etiqueta = f"{prov} · {cuenta}" if cuenta else str(prov)
            dd = ft.Dropdown(
                label="Cuenta Origen (pago en pesos)", width=340,
                enable_filter=True, editable=True,
                value=clabes_prev.get(par) or None,
                options=[ft.dropdown.Option(key=cl, text=cta) for cta, cl in opciones])
            tfc = ft.TextField(
                label="Concepto de pago", width=200, value=concep_prev.get(par, ""))
            tfr = ft.TextField(
                label="Referencia bancaria", width=200, value=refs_prev.get(par, ""))
            controles[par] = (dd, tfc, tfr)
            bloques.append(ft.Column(
                [ft.Text(etiqueta, size=12, weight=ft.FontWeight.BOLD),
                 ft.Row([dd, tfc, tfr], spacing=12, wrap=True)],
                spacing=6, tight=True))

        def regenerar(_e=None) -> None:
            clabes_par, concep_par, refs_par = {}, {}, {}
            for par, (dd, tfc, tfr) in controles.items():
                clabes_par[par] = (dd.value or "").strip()
                concep_par[par] = (tfc.value or "").strip()
                refs_par[par] = (tfr.value or "").strip()
            if not any(clabes_par.values()):
                self.app.avisar(
                    "Elige la Cuenta Origen de al menos un proveedor.", NARANJA)
                return
            ok, msg = self._regenerar_txt_pesos(
                folio_dict, clabes_par, concep_par, refs_par)
            self.page.pop_dialog()
            self.app.avisar(msg, VERDE if ok else ROJO)

        folio = folio_dict.get("folio")
        alto = min(460, 70 + 116 * len(pares))
        self.page.show_dialog(ft.AlertDialog(
            modal=True,
            title=ft.Text(
                f"Regenerar TXT en pesos — folio {folio}",
                weight=ft.FontWeight.BOLD),
            content=ft.Container(
                content=ft.Column(bloques, tight=True, spacing=16,
                                  scroll=ft.ScrollMode.AUTO),
                width=800, height=alto),
            actions=[
                ft.TextButton("Cancelar", on_click=lambda _e: self.page.pop_dialog()),
                ft.FilledButton("Regenerar", icon=ft.Icons.AUTORENEW,
                                on_click=regenerar),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        ))

    def _regenerar_txt_pesos(
        self, folio_dict: dict, clabes_par: dict, conceptos_par: dict,
        refs_par: dict,
    ) -> tuple[bool, str]:
        """Regenera los TXT en pesos de una dispersión (USD pago en MXN) con la Cuenta
        Origen / concepto / referencia elegidos POR PAR (proveedor · cuenta). Agrupa
        por cuenta origen y genera UN archivo por cada una (su banco define el
        formato), igual que la generación original. Reemplaza los TXT previos de esta
        dispersión (borra los anteriores para no dejar sobrantes si cambió el origen).
        Devuelve (ok, mensaje)."""
        clave = folio_dict.get("clave") or ""
        empresa = folio_dict.get("empresa") or ""
        conc = self._conc_dispersion
        emp = next((e for e in conc.validas if e.empresa == clave), None) \
            if conc is not None else None
        if emp is None:
            return False, "No se encontró la información de la dispersión para regenerar."
        pares = self._pesos_por_grupo.get(clave) or set()
        if not pares:
            return False, "Esta dispersión no tiene proveedores marcados 'pagar en pesos'."
        # Tipo de cambio del día hábil anterior (los lunes, el viernes pasado);
        # reutiliza el ya obtenido si existe.
        tc = self._tipo_cambio
        if not tc:
            try:
                tc, tc_fecha = tipo_cambio.tipo_cambio_usd_detalle()
                self._tipo_cambio = tc
                self._tc_fecha = tc_fecha
            except Exception as exc:  # noqa: BLE001 — se reporta al usuario
                return False, f"No se pudo obtener el tipo de cambio: {exc}"
        # Texto de cada cuenta origen (define banco/formato del layout).
        texto_por_clabe = {cl: cta for cta, cl in self._clabes_de_empresa(empresa)}
        # Agrupa los registros por CUENTA ORIGEN (un archivo por origen), con el
        # concepto/referencia de cada par.
        por_origen: dict[str, dict] = {}
        for par in pares:
            clabe_origen = (clabes_par.get(par) or "").strip()
            if not clabe_origen:
                continue
            prov, cuenta = par
            movs = [m for m in emp.movimientos
                    if m.proveedor == prov and m.cuenta_bancaria == cuenta]
            # Con las notas de crédito descontadas (ver total_a_pagar).
            usd = reporte_dispersion.total_a_pagar(movs)
            pesos = round(usd * tc, 2)
            concepto = (conceptos_par.get(par) or emp.concepto_pago or "").strip()
            referencia = (refs_par.get(par) or "").strip()
            texto = f"{concepto} {referencia}".strip() if referencia else concepto
            bucket = por_origen.setdefault(clabe_origen, {
                "cuenta_texto": texto_por_clabe.get(clabe_origen, ""),
                "registros": [], "total": 0.0})
            bucket["registros"].append(
                (re.sub(r"\D", "", cuenta or ""), pesos, prov, texto))
            bucket["total"] += pesos
        if not por_origen:
            return False, "Ningún proveedor tiene Cuenta Origen para regenerar."
        # Borra los TXT anteriores de esta dispersión (para reemplazarlos y no dejar
        # sobrantes si cambió alguna cuenta origen).
        carpeta = self._disp_carpeta_txt or self._carpeta_txt_dispersion()
        os.makedirs(carpeta, exist_ok=True)
        for p in [x for x in self._pesos_generados if x.get("empresa") == clave]:
            try:
                os.remove(p.get("archivo"))
            except OSError:
                pass
            self._pesos_generados.remove(p)
        # Un TXT por cada cuenta origen (mismo nombre que en la generación original).
        folio = folio_dict.get("folio")
        generados = 0
        for clabe_origen, bucket in por_origen.items():
            registros = bucket["registros"]
            clabe_dig = re.sub(r"\D", "", clabe_origen or "")
            cuenta_texto = bucket["cuenta_texto"]
            if exportador_devoluciones.banco_formato(cuenta_texto) == "banregio":
                hoy = datetime.date.today().strftime("%d%m%Y")
                contenido = exportador_devoluciones.generar_banregio(registros, hoy)
            else:  # BBVA / Bancomer (ancho fijo) — formato por defecto
                contenido = exportador_devoluciones.generar_bancomer(
                    registros, clabe_dig, str(folio or ""))
            nombre = self._nombre_txt_pesos(folio_dict, clabe_dig)
            ruta = os.path.join(carpeta, nombre)  # sobrescribe (mismo nombre)
            try:
                with open(ruta, "w", encoding="latin-1", newline="") as fh:
                    fh.write(contenido)
            except Exception as exc:  # noqa: BLE001 — se reporta al usuario
                return False, f"No se pudo escribir el TXT: {exc}"
            self._pesos_generados.append({
                "empresa": clave, "archivo": ruta,
                "total_pesos": bucket["total"], "proveedores": len(registros)})
            generados += 1
        # Persiste los valores editados por par (por si se regenera de nuevo).
        self._clabe_pesos_por_grupo[clave] = {
            par: cl for par, cl in clabes_par.items() if cl}
        self._concepto_prov_por_grupo[clave] = dict(conceptos_par)
        self._ref_prov_por_grupo[clave] = dict(refs_par)
        self._cuenta_pesos_por_grupo[clave] = {
            par: texto_por_clabe.get(cl, "")
            for par, cl in clabes_par.items() if cl}
        return True, f"{generados} TXT en pesos regenerado(s)."

    def _dialogo_continuar_operacion(self) -> None:
        """Confirmación para continuar con la SUBIDA de comprobantes. Al aceptar: se
        cierra la confirmación, luego (encadenado por on_dismiss) el resumen y, cuando
        éste se desmonta, se abre el RPA de subida (que arranca solo)."""
        estado = {"aceptar": False}

        def al_cerrar_confirmacion(_e=None) -> None:
            if estado["aceptar"]:
                # El resumen ya es el diálogo de arriba: cerrarlo y, tras su cierre,
                # abrir el diálogo del RPA de subida.
                self._cerrar_resumen_luego(self._abrir_dialogo_subida)

        def aceptar(_e=None) -> None:
            estado["aceptar"] = True
            self.page.pop_dialog()  # confirmación -> dispara al_cerrar_confirmacion

        self.page.show_dialog(ft.AlertDialog(
            modal=True,
            title=ft.Text("Continuar operación", weight=ft.FontWeight.BOLD),
            content=ft.Text(
                "Se subirán los archivos vinculados a las dispersiones.\n"
                "¿Continuar con la operación?"),
            on_dismiss=al_cerrar_confirmacion,
            actions=[
                ft.TextButton("Cancelar", on_click=lambda _e: self.page.pop_dialog()),
                ft.FilledButton("Aceptar", on_click=aceptar),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        ))

    def _dialogo_terminar_operacion(self) -> None:
        """Confirmación para terminar la operación. Al aceptar: se cierra la
        confirmación, luego (encadenado por on_dismiss) el resumen y, cuando éste se
        desmonta, se cierra —de forma obligatoria— el navegador. El encadenamiento es
        necesario porque Flet solo desmonta el diálogo de más arriba."""
        estado = {"aceptar": False}

        def cerrar_navegador() -> None:
            self.page.run_task(self._detener_rpa)

        def al_cerrar_confirmacion(_e=None) -> None:
            if estado["aceptar"]:
                self._cerrar_resumen_luego(cerrar_navegador)

        def aceptar(_e=None) -> None:
            estado["aceptar"] = True
            self.page.pop_dialog()  # confirmación -> dispara al_cerrar_confirmacion

        self.page.show_dialog(ft.AlertDialog(
            modal=True,
            title=ft.Text("Terminar operación", weight=ft.FontWeight.BOLD),
            content=ft.Text(
                "Se dará por terminada la operación de dispersión.\n\n ¿Continuar?"),
            on_dismiss=al_cerrar_confirmacion,
            actions=[
                ft.TextButton("Cancelar", on_click=lambda _e: self.page.pop_dialog()),
                ft.FilledButton("Aceptar", on_click=aceptar),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        ))

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

    def _rango_moneda(self, clave: str, moneda: str) -> int:
        """Rango de orden/categoría por moneda de un grupo (empresa+moneda):
        0 = MXN (u otras), 1 = USD, 2 = USD pago en MXN (USD con ≥1 par marcado
        'pagar en pesos', según self._pesos_por_grupo). La moneda llega ya
        normalizada (MN→MXN)."""
        if (moneda or "").strip().upper() == "USD":
            return 2 if self._pesos_por_grupo.get(clave) else 1
        return 0

    def _rango_moneda_fila(self, d: dict) -> int:
        """Categoría por moneda de UNA fila del resumen (proveedor+cuenta): 0 = MXN,
        1 = USD, 2 = USD pago en MXN. A diferencia de _rango_moneda (por grupo), aquí
        'usd_pesos' aplica solo si ESE par está marcado 'pagar en pesos', de modo que un
        mismo grupo USD puede repartir sus filas entre USD y USD pago en MXN."""
        if (d.get("moneda") or "").strip().upper() != "USD":
            return 0
        par = d.get("par")
        if par and par in self._pesos_por_grupo.get(d.get("clave") or "", set()):
            return 2
        return 1

    def _mostrar_datos_dispersion(self, conc: "conciliacion.Conciliacion") -> None:
        """Muestra, de forma amigable (tablas, sin JSON), los datos que tomará el
        robot, separados por empresa + tipo de moneda (la misma separación de la
        pantalla). Pensado para usuarios no técnicos."""
        empresas = conc.empresas if conc else []
        if not empresas:
            cuerpo: ft.Control = ft.Text("No hay datos que mostrar.", size=12, color=GRIS)
        else:
            # Orden por moneda: MXN → USD → USD pago en MXN (sort estable: dentro de
            # cada categoría se respeta el orden original).
            empresas = sorted(empresas, key=lambda e: self._rango_moneda(
                e.empresa, e.movimientos[0].moneda if e.movimientos else ""))
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
            con_fecha = (f" con fecha del {self._tc_preview_fecha}"
                         if self._tc_preview_fecha else "")
            texto = (f"Tipo de cambio: {_fmt_tc(self._tc_preview)} MXN (tomado del "
                     f"Diario Oficial de la Federación{con_fecha})")
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
        datos = [
            self._dato_compacto("Cuenta origen", e.cuenta or "—"),
            self._dato_compacto("Concepto", e.concepto_pago or "—"),
            self._dato_compacto("Referencia", e.referencia_pago or "—"),
        ]
        # En entradas 'USD pago en MXN' (con pares marcados 'pagar en pesos') se
        # añade la(s) cuenta(s) origen elegidas para el pago en pesos.
        if e.empresa in self._pesos_por_grupo:
            cuentas_pesos = [
                c for c in dict.fromkeys(
                    self._cuenta_pesos_por_grupo.get(e.empresa, {}).values()) if c]
            datos.append(self._dato_compacto(
                "Cuenta de pago en pesos",
                " · ".join(cuentas_pesos) if cuentas_pesos else "—"))
        info = ft.Row(datos, wrap=True, spacing=20, run_spacing=4)
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
        """Tabla de los movimientos a dispersar de una empresa, AGRUPADOS por
        proveedor+cuenta: una banda-cabecera por grupo con 'proveedor · cuenta … TOTAL
        PROG. $X' (igual que la tabla principal) y, al final, una banda con el TOTAL
        GENERAL, para que el usuario vea el total por proveedor y el total general. Las
        columnas de detalle son Folio, Folio Factura y los importes (proveedor y cuenta
        ya van en la banda). Si el grupo es USD con proveedores 'pagar en pesos', añade
        la columna Equiv. MXN (Saldo Programado × T.C.)."""
        # Pares (proveedor, cuenta beneficiario) del grupo marcados 'pagar en pesos' y
        # el tipo de cambio. La columna Equiv. MXN solo aparece si hay marcados y T.C.
        pesos_set = self._pesos_por_grupo.get(e.empresa, set())
        tc = self._tc_preview
        mostrar_pesos = bool(pesos_set) and bool(tc)

        cols = [
            ColumnaTabla("Folio", 11 if mostrar_pesos else 13, CENTRO),
            ColumnaTabla("Folio Factura", 15 if mostrar_pesos else 18, CENTRO),
            ColumnaTabla("Total Fact.", 18 if mostrar_pesos else 23, _TDER),
            ColumnaTabla("Saldo Fact.", 18 if mostrar_pesos else 23, _TDER),
            ColumnaTabla("Saldo Prog.", 18 if mostrar_pesos else 22, _TDER),
        ]
        if mostrar_pesos:
            cols.append(ColumnaTabla("Equiv. MXN", 19, _TDER))
        n_cols = len(cols)

        def banda(prov, cuenta, total, total_pesos, general=False) -> Cabecera:
            """Banda (Cabecera a todo lo ancho) con 'proveedor · cuenta' a la izquierda
            y el total a la derecha. Si `general`, es el TOTAL GENERAL (sin proveedor,
            con un fondo distinto para diferenciarlo de las bandas de grupo)."""
            if general:
                izq: ft.Control = ft.Container(expand=True)
                etiqueta = "TOTAL GENERAL PROG."
            else:
                prov_txt, cta_txt = str(prov or "—"), str(cuenta or "")
                # Los Text llevan `expand` (no `tight`) para que el Row les FIJE un
                # ancho: sin eso toman su ancho natural, el '…' nunca entra y el texto
                # se desborda encimándose con el total de la derecha. El reparto 3/2
                # da más espacio al proveedor, que suele ser el más largo.
                izq = ft.Row(
                    [ft.Text(prov_txt, size=13, weight=ft.FontWeight.BOLD,
                             max_lines=1, no_wrap=True, expand=3,
                             overflow=ft.TextOverflow.ELLIPSIS,
                             tooltip=prov_txt),
                     ft.Text("·", size=13, color=GRIS),
                     ft.Text(cta_txt, size=12, weight=ft.FontWeight.BOLD, color=GRIS,
                             max_lines=1, no_wrap=True, expand=2,
                             overflow=ft.TextOverflow.ELLIPSIS,
                             tooltip=cta_txt or None)],
                    spacing=8, expand=True,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER)
                etiqueta = "TOTAL PROG."
            tam = 14 if general else 13
            der_items = [
                ft.Text(etiqueta, size=11, weight=ft.FontWeight.BOLD, color=GRIS),
                ft.Text(_fmt_moneda(total), size=tam, weight=ft.FontWeight.BOLD,
                        no_wrap=True, tooltip=f"{etiqueta} {_fmt_moneda(total)}"),
            ]
            if total_pesos is not None:
                der_items += [
                    ft.Text("· Equiv. MXN", size=11, weight=ft.FontWeight.BOLD,
                            color=GRIS),
                    ft.Text(_fmt_moneda(total_pesos), size=tam,
                            weight=ft.FontWeight.BOLD, no_wrap=True,
                            tooltip=f"Equivalente en MXN {_fmt_moneda(total_pesos)}"),
                ]
            der = ft.Row(der_items, spacing=6, tight=True,
                         vertical_alignment=ft.CrossAxisAlignment.CENTER)
            info = ft.Row([izq, der], vertical_alignment=ft.CrossAxisAlignment.CENTER)
            return Cabecera(
                [SegmentoCabecera(n_cols, info, alineacion=None,
                                  padding=ft.Padding.only(left=10, right=10))],
                bgcolor=ft.Colors.SURFACE_CONTAINER_HIGHEST if general else None,
                alto=40)

        # Agrupa los movimientos por (proveedor, cuenta) en orden de aparición.
        grupos: dict[tuple, list] = {}
        orden: list[tuple] = []
        for m in e.movimientos:
            k = (m.proveedor, m.cuenta_bancaria)
            if k not in grupos:
                grupos[k] = []
                orden.append(k)
            grupos[k].append(m)

        filas: list = []
        tot_prog = 0.0
        tot_pesos = 0.0
        for (prov, cuenta) in orden:
            movs = grupos[(prov, cuenta)]
            # Los TOTALES van sin notas de crédito (la factura ligada ya las trae
            # descontadas); las filas de detalle sí las muestran, en negativo.
            grupo_prog = reporte_dispersion.total_a_pagar(movs)
            es_pesos = mostrar_pesos and (prov, cuenta) in pesos_set
            grupo_pesos = round(grupo_prog * tc, 2) if es_pesos else None
            filas.append(banda(prov, cuenta, grupo_prog, grupo_pesos))
            tot_prog += grupo_prog
            for m in movs:
                celdas: list = [
                    m.folio, m.folio_factura,
                    _fmt_moneda(m.total_factura),
                    _fmt_moneda(m.saldo_factura),
                    _fmt_moneda(m.saldo_programado),
                ]
                if mostrar_pesos:
                    if es_pesos:
                        p = round((m.saldo_programado or 0) * tc, 2)
                        tot_pesos += p
                        celdas.append(_fmt_moneda(p))
                    else:  # (proveedor, cuenta) USD que NO se paga en pesos
                        celdas.append("—")
                filas.append(FilaDatos(celdas))
        # Banda de TOTAL GENERAL (debajo de los grupos).
        filas.append(banda(None, None, tot_prog,
                           tot_pesos if mostrar_pesos else None, general=True))

        tabla = TablaResponsiva(
            self.page, cols,
            ancho_inicial=860 - 2 * _GUTTER_SCROLL - 24, alto_fila=40)
        tabla.set_contenido(filas)
        return tabla.control

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

    def _reiniciar_tablas(self) -> None:
        """Vacía TODAS las tablas y su estado por grupo (empresa+moneda): quita sus
        controles del árbol y olvida empresas/fechas. NO reconstruye ni avisa; lo hace
        quien llama. Se usa para volver al estado inicial y como reinicio antes de una
        búsqueda nueva (las solicitudes no se acumulan entre búsquedas)."""
        for tabla in self._tablas_por_empresa.values():
            try:
                self._contenedor_tablas.controls.remove(tabla.control)
            except ValueError:
                pass
        self._tablas_por_empresa.clear()
        self._fechas_por_grupo.clear()
        self._empresa_activa = None

    def _eliminar_todo(self, _e=None) -> None:
        """Vacía TODAS las tablas: quita sus controles del árbol, olvida las
        empresas y vuelve al estado inicial (placeholder). Cierra el diálogo."""
        self.page.pop_dialog()
        self._reiniciar_tablas()
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

    # ----------------------------------------------------- búsqueda (API)
    async def _buscar_solicitudes(self, _e=None) -> None:
        """Consulta las solicitudes a dispersar en el endpoint (sin navegador ni
        login) y vuelca las filas en las tablas. Hace UNA llamada por empresa
        seleccionada (el endpoint recibe un solo id de empresa); si alguna falla, se
        avisa sin abortar las demás. El RPA queda solo para el paso de dispersar."""
        error = self._validar_filtros()
        if error:
            self.app.avisar(error, ROJO)
            return
        fecha_ini_ui = (self.tf_fecha_ini.value or "").strip()
        fecha_fin_ui = (self.tf_fecha_fin.value or "").strip()
        fecha_ini = _fecha_ddmmaaaa_a_iso(fecha_ini_ui) or None
        fecha_fin = _fecha_ddmmaaaa_a_iso(fecha_fin_ui) or None
        folio_txt = (self.tf_folio.value or "").strip()
        folio = int(folio_txt) if folio_txt.isdigit() else None
        # Tipo(s): si no se elige ninguno o se eligen TODOS -> sin filtro (una llamada
        # por empresa). Si se eligen algunos, una llamada por empresa × tipo (con su id).
        tipos_sel = self.ms_tipo.valores()
        if not tipos_sel or set(tipos_sel) == set(self.TIPOS_SOLICITUD):
            tipos_id = [None]
        else:
            tipos_id = [
                _TIPO_SOLICITUD_ID[t] for t in tipos_sel if t in _TIPO_SOLICITUD_ID
            ] or [None]

        self.btn_iniciar.disabled = True
        self._mostrar_cargando(True)
        self.page.update()

        filas: list[FilaSolicitud] = []
        errores: list[str] = []
        try:
            for nombre in self.ms_empresa.valores():
                id_empresa = self.ID_POR_EMPRESA.get(nombre)
                if id_empresa is None:
                    errores.append(f"{nombre} (sin id de empresa)")
                    continue
                for tipo_id in tipos_id:
                    try:
                        resp = await asyncio.to_thread(
                            api.dispersiones_no_pemex, id_empresa,
                            fecha_inicio=fecha_ini, fecha_fin=fecha_fin,
                            tipo_solicitud=tipo_id, folio_solicitud=folio)
                    except api.ErrorApi as exc:
                        errores.append(f"{nombre}: {exc}")
                        continue
                    filas.extend(reporte_dispersion.desde_api(resp))
        finally:
            self._mostrar_cargando(False)
            self.btn_iniciar.disabled = False

        n = len(filas)
        # Cada búsqueda parte de cero: se descartan las tablas de la búsqueda anterior
        # (las solicitudes NO se acumulan entre búsquedas). La carga por Excel sí
        # acumula a propósito, por eso el reinicio vive aquí y no en volcar_reportes.
        self._reiniciar_tablas()
        if filas:
            # Vuelca (agrupa por empresa+moneda, crea tablas, dedup por clave) y
            # guarda las fechas de la búsqueda por grupo para la dispersión.
            self.volcar_reportes(
                filas, fecha_ini=fecha_ini_ui, fecha_fin=fecha_fin_ui)
        else:
            self._reconstruir_tablas()  # oculta el "cargando" / muestra placeholder
        self.page.update()

        if errores:
            self.app.avisar(
                "Algunas consultas fallaron: " + "; ".join(errores[:4])
                + ("…" if len(errores) > 4 else ""), ROJO)
        elif n:
            self.app.avisar(f"{n} solicitud(es) encontrada(s).", VERDE)
        else:
            self.app.avisar(
                "No se encontraron solicitudes con esos filtros.", NARANJA)

    # ------------------------------------------------- hooks del RPA
    async def _correr(self, coro):
        """Corre una corrutina del RPA en el bucle del hilo dedicado y espera su
        resultado sin congelar la interfaz."""
        if self.bucle is None:
            self.bucle = BucleRpa()
        return await asyncio.wrap_future(self.bucle.enviar(coro))

    async def _detener_rpa(self) -> None:
        """Cierra el navegador, libera la sesión y trae la app al frente (best-effort)."""
        self._ctrl = None
        self._future_rpa = None
        sesion = self.sesion
        self.sesion = None
        if sesion is not None:
            try:
                await self._correr(sesion.cerrar())
            except Exception:  # noqa: BLE001 — el cierre no debe propagar errores
                pass
        # Tras cerrar el navegador del RPA, devolver el foco a la ventana de la app.
        self._enfocar_app()

    def _enfocar_app(self) -> None:
        """Trae la ventana de la app al primer plano (tras cerrar el navegador del
        RPA). Best-effort y NO-OP fuera de Windows."""
        try:
            from core import win_taskbar
            win_taskbar.traer_al_frente(self.page.title)
        except Exception:  # noqa: BLE001 — el foco no es crítico
            pass
