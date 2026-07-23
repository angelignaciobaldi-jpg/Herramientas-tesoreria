"""Tabla responsiva con columnas por PORCENTAJE del ancho del contenedor.

A diferencia de un layout de anchos fijos en píxeles, aquí cada columna declara un
PORCENTAJE (`pct`) del ancho disponible. La tabla mide su propio contenedor con el
evento `on_size_change` de Flet y convierte esos porcentajes a píxeles en cada cambio
de tamaño, así las columnas crecen o se encogen con la ventana. Si la suma de
porcentajes supera el 100%, el contenido excede el ancho disponible y aparece una
barra de scroll horizontal (el `Row(scroll=AUTO)` interior).

Es un componente GENÉRICO y reutilizable (no atado a ninguna pantalla). Soporta filas
"cabecera" OPCIONALES (bandas a todo lo ancho con colspan), intercaladas con las filas
de datos: sirven para agrupar (p. ej. proveedor + cuenta + total) o para cualquier
encabezado de sección. El contenido de cada celda puede ser texto (se trunca con '…' y
tooltip si no cabe) o cualquier `Control` de Flet (p. ej. un `Checkbox`).

Uso típico:

    cols = [
        ColumnaTabla("", 4, encabezado_control=chk_todos),   # columna del check
        ColumnaTabla("Folio", 6),
        ColumnaTabla("Importe", 12, alineacion=DER),
        ...
    ]
    tabla = TablaResponsiva(page, cols)
    tabla.set_contenido([
        Cabecera([SegmentoCabecera(1, chk_grupo), SegmentoCabecera(len(cols) - 1, info)]),
        FilaDatos([chk_fila, "F-001", "$1,234.00", ...], bgcolor=color),
        ...
    ])
    # tabla.control se agrega al árbol de la pantalla.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional

import flet as ft

from ui.comun import CENTRO

# Alineaciones (mismas que usan las pantallas): izquierda / derecha / centro.
IZQ = ft.Alignment(-1, 0)
DER = ft.Alignment(1, 0)

# Ancho aproximado de un carácter a size=12 (px): decide, sin medir el render, si el
# texto de una celda probablemente se recorta (y por tanto amerita tooltip).
_PX_POR_CHAR = 6.0
# Canalón para que la barra de scroll horizontal no tape la última fila.
_GUTTER_SCROLL = 14
_ALTO_FILA = 44
_ALTO_ENCABEZADO = 46
_COL_SPACING = 8
# Ancho supuesto antes del primer `on_size_change` (se corrige al medir de verdad).
_ANCHO_INICIAL_DEFECTO = 1100.0
# No repintar por cambios de ancho menores a esto (evita trabajo durante el arrastre).
_UMBRAL_REMEDIR = 2.0


# --------------------------------------------------------------- modelo de datos
@dataclass
class ColumnaTabla:
    """Una columna de la tabla. `pct` es su porcentaje del ancho disponible.

    `encabezado_control`: si se da, se dibuja ese control en la celda del encabezado
    (p. ej. el check 'seleccionar todas') en vez de la etiqueta de texto.
    `ancho_min_px`: piso opcional en píxeles (para que no se encoja de más)."""

    etiqueta: str
    pct: float
    alineacion: ft.Alignment = field(default_factory=lambda: CENTRO)
    encabezado_control: Optional[ft.Control] = None
    ancho_min_px: int = 0


@dataclass
class SegmentoCabecera:
    """Tramo de una fila-cabecera que abarca `cols` columnas consecutivas (colspan).
    Su ancho = suma de los px de esas columnas (más los gaps internos), así una banda
    puede alinear su primera celda con la columna del check y ocupar el resto."""

    cols: int
    contenido: ft.Control
    alineacion: ft.Alignment = field(default_factory=lambda: IZQ)
    padding: Optional[ft.Padding] = None


@dataclass
class Cabecera:
    """Fila-cabecera OPCIONAL a todo lo ancho (una o varias, intercaladas con datos)."""

    segmentos: list[SegmentoCabecera]
    bgcolor: Optional[str] = None  # None -> PRIMARY_CONTAINER al pintar
    alto: int = _ALTO_FILA


@dataclass
class FilaDatos:
    """Fila de datos: una celda por columna (str -> texto truncado; Control -> tal cual)."""

    celdas: list  # list[str | ft.Control], alineada a las columnas
    bgcolor: Optional[str] = None
    alto: int = _ALTO_FILA


# --------------------------------------------------------------- la tabla
class TablaResponsiva:
    """Tabla cuyas columnas se dimensionan por porcentaje del ancho del contenedor.

    Mide su propio ancho con `on_size_change` y recomputa los píxeles por columna en
    cada cambio de tamaño. `control` es el widget a insertar en el árbol; debe quedar
    en un contenedor que lo estire a lo ancho (columnas con `horizontal_alignment=
    STRETCH`) para que la medición refleje el ancho real disponible."""

    def __init__(self, page, columnas: list[ColumnaTabla], *,
                 con_encabezado: bool = True, spacing: int = _COL_SPACING,
                 alto_fila: int = _ALTO_FILA, alto_encabezado: int = _ALTO_ENCABEZADO,
                 ancho_inicial: "float | None" = None):
        self.page = page
        self.columnas = list(columnas)
        self.con_encabezado = con_encabezado
        self.spacing = spacing
        self.alto_fila = alto_fila
        self.alto_encabezado = alto_encabezado
        self._filas: list = []            # list[Cabecera | FilaDatos]
        self._px: list[int] = []          # ancho en px por columna (computado)
        self._ancho_total = 0             # ancho total de una fila (px)
        self._disponible = float(
            ancho_inicial or (getattr(page, "width", None) or _ANCHO_INICIAL_DEFECTO))

        # Encabezado fijo (si aplica) + cuerpo, dentro de un marco con borde. El
        # Row exterior aporta el scroll HORIZONTAL cuando el contenido excede el ancho.
        self._encabezado_holder = ft.Container(visible=con_encabezado)
        self._cuerpo = ft.Column(spacing=0, tight=True)
        hijos_marco = [self._encabezado_holder, self._cuerpo] if con_encabezado \
            else [self._cuerpo]
        self._marco = ft.Container(
            ft.Column(hijos_marco, spacing=0, tight=True),
            border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT), border_radius=10,
            clip_behavior=ft.ClipBehavior.ANTI_ALIAS)
        self._scroll = ft.Row(
            [ft.Container(self._marco, padding=ft.Padding.only(bottom=_GUTTER_SCROLL))],
            scroll=ft.ScrollMode.AUTO)
        # El contenedor medidor: STRETCH del padre le da el ancho real disponible, y
        # `on_size_change` nos avisa cada vez que cambia (resize, cambio de pestaña…).
        self._medidor = ft.Container(self._scroll, on_size_change=self._medir)
        self.control = ft.Column(
            [self._medidor], spacing=0, tight=True,
            horizontal_alignment=ft.CrossAxisAlignment.STRETCH)

        # Referencias para mutar anchos SIN recrear controles al redimensionar
        # (recrear re-parentearía checkboxes existentes y Flet puede rechazarlo).
        self._enc_refs: list = []   # (container, texto|None) por columna del encabezado
        self._enc_row = None        # Container del encabezado
        self._filas_refs: list = []  # una entrada por fila renderizada (datos/cabecera)
        self._recomputar_px()
        if con_encabezado:
            self._encabezado_holder.content = self._construir_encabezado()
        self._pintar_cuerpo()

    # ----------------------------------------------------------- API pública
    def set_contenido(self, filas: list) -> None:
        """Fija las filas (mezcla de `Cabecera` y `FilaDatos`, en orden) y repinta el
        cuerpo (crea controles nuevos con los px actuales)."""
        self._filas = list(filas)
        self._pintar_cuerpo()

    def set_columnas(self, columnas: list[ColumnaTabla]) -> None:
        """Reemplaza las columnas (p. ej. para cambiar porcentajes) y repinta todo."""
        self.columnas = list(columnas)
        self._encabezado_holder.visible = self.con_encabezado
        self._recomputar_px()
        if self.con_encabezado:
            self._encabezado_holder.content = self._construir_encabezado()
        self._pintar_cuerpo()

    def contenedor_fila(self, indice: int):
        """Container de la fila renderizada en la posición `indice` (según el orden
        pasado a set_contenido). Permite mutar su `bgcolor` en vivo —sin reconstruir
        la tabla— para no perder el foco de un campo que se está editando. Devuelve
        None si el índice no existe."""
        if 0 <= indice < len(self._filas_refs):
            return self._filas_refs[indice].get("cont")
        return None

    # ----------------------------------------------------------- medición / px
    def _medir(self, e) -> None:
        """Handler de `on_size_change`: si el ancho cambió lo suficiente, recomputa
        los px por columna y SOLO muta los anchos ya montados (columnas que crecen/
        encogen con la ventana), sin recrear controles."""
        w = float(getattr(e, "width", 0) or 0)
        if w <= 0 or abs(w - self._disponible) < _UMBRAL_REMEDIR:
            return
        self._disponible = w
        self._recomputar_px()
        self._aplicar_px()

    def _recomputar_px(self) -> None:
        """Convierte los porcentajes a píxeles sobre el ancho disponible actual
        (descontando los gaps entre columnas). Respeta `ancho_min_px`."""
        n = len(self.columnas)
        if n == 0:
            self._px, self._ancho_total = [], 0
            return
        contenido = max(0.0, self._disponible - self.spacing * (n - 1))
        px = []
        for c in self.columnas:
            p = round(c.pct / 100.0 * contenido)
            if c.ancho_min_px:
                p = max(p, c.ancho_min_px)
            px.append(max(1, int(p)))
        self._px = px
        self._ancho_total = sum(px) + self.spacing * (n - 1)

    # ----------------------------------------------------------- render
    @staticmethod
    def _borde_inferior(opaco: bool = False):
        color = ft.Colors.OUTLINE_VARIANT if opaco \
            else ft.Colors.with_opacity(0.5, ft.Colors.OUTLINE_VARIANT)
        return ft.Border(bottom=ft.BorderSide(1, color))

    def _mk_celda(self, contenido, ancho: int, alineacion=CENTRO, bold: bool = False):
        """Devuelve (container, texto|None). Si `contenido` es un Control, se coloca
        tal cual (texto=None); si es texto, se recorta con '…' y lleva tooltip solo si
        probablemente no cabe. Se devuelve el Text para poder mutar su ancho en resize."""
        if isinstance(contenido, ft.Control):
            return ft.Container(contenido, width=ancho, alignment=alineacion), None
        texto = str(contenido or "")
        if alineacion is DER:
            align_txt = ft.TextAlign.RIGHT
        elif alineacion is IZQ:
            align_txt = ft.TextAlign.LEFT
        else:
            align_txt = ft.TextAlign.CENTER
        tip = texto if len(texto) * _PX_POR_CHAR > ancho else None
        t = ft.Text(texto, size=12, text_align=align_txt, width=ancho,
                    weight=ft.FontWeight.BOLD if bold else None,
                    max_lines=1, no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS)
        return ft.Container(t, width=ancho, alignment=alineacion, tooltip=tip), t

    def _construir_encabezado(self) -> ft.Container:
        self._enc_refs = []
        celdas = []
        for i, c in enumerate(self.columnas):
            ancho = self._px[i]
            if c.encabezado_control is not None:
                cont = ft.Container(c.encabezado_control, width=ancho, alignment=CENTRO)
                self._enc_refs.append((cont, None))
            else:
                cont, t = self._mk_celda(c.etiqueta, ancho, CENTRO, bold=True)
                self._enc_refs.append((cont, t))
            celdas.append(cont)
        self._enc_row = ft.Container(
            ft.Row(celdas, spacing=self.spacing, tight=True,
                   vertical_alignment=ft.CrossAxisAlignment.CENTER),
            width=self._ancho_total, height=self.alto_encabezado,
            bgcolor=ft.Colors.SURFACE_CONTAINER_HIGHEST,
            border=self._borde_inferior())
        return self._enc_row

    def _construir_fila_datos(self, fila: FilaDatos) -> ft.Container:
        celdas = []
        refs = []
        for i, c in enumerate(self.columnas):
            contenido = fila.celdas[i] if i < len(fila.celdas) else ""
            cont, t = self._mk_celda(contenido, self._px[i], c.alineacion)
            celdas.append(cont)
            refs.append((cont, t))
        row = ft.Container(
            ft.Row(celdas, spacing=self.spacing, tight=True,
                   vertical_alignment=ft.CrossAxisAlignment.CENTER),
            width=self._ancho_total, height=fila.alto, bgcolor=fila.bgcolor,
            border=self._borde_inferior())
        self._filas_refs.append({"tipo": "datos", "cont": row, "celdas": refs})
        return row

    def _construir_cabecera(self, cab: Cabecera) -> ft.Container:
        """Fila-cabecera con segmentos (colspan). El ancho de cada segmento = suma de
        los px de sus columnas + los gaps internos, para cuadrar con las filas de datos."""
        segmentos = []
        refs = []
        col = 0
        for seg in cab.segmentos:
            k = max(1, seg.cols)
            ancho = sum(self._px[col:col + k]) + self.spacing * (k - 1)
            cont = ft.Container(seg.contenido, width=ancho, alignment=seg.alineacion,
                                padding=seg.padding)
            segmentos.append(cont)
            refs.append((cont, col, k))
            col += k
        row = ft.Container(
            ft.Row(segmentos, spacing=self.spacing, tight=True,
                   vertical_alignment=ft.CrossAxisAlignment.CENTER),
            width=self._ancho_total, height=cab.alto,
            bgcolor=cab.bgcolor if cab.bgcolor is not None
            else ft.Colors.PRIMARY_CONTAINER,
            border=self._borde_inferior(opaco=True))
        self._filas_refs.append({"tipo": "cab", "cont": row, "segs": refs})
        return row

    def _pintar_cuerpo(self) -> None:
        """(Re)construye el CUERPO con controles nuevos y los px actuales, y repinta.
        El encabezado NO se recrea aquí (se mantiene para no re-parentear su check)."""
        self._filas_refs = []
        cuerpo = []
        for fila in self._filas:
            if isinstance(fila, Cabecera):
                cuerpo.append(self._construir_cabecera(fila))
            else:
                cuerpo.append(self._construir_fila_datos(fila))
        self._cuerpo.controls = cuerpo
        self._safe_update()

    def _aplicar_px(self) -> None:
        """Muta EN SITIO los anchos de encabezado y filas ya montadas según los px
        actuales (sin recrear controles) y repinta. Se usa al redimensionar."""
        if self._enc_row is not None:
            self._enc_row.width = self._ancho_total
            for i, (cont, t) in enumerate(self._enc_refs):
                cont.width = self._px[i]
                if t is not None:
                    t.width = self._px[i]
        for rec in self._filas_refs:
            rec["cont"].width = self._ancho_total
            if rec["tipo"] == "datos":
                for i, (cont, t) in enumerate(rec["celdas"]):
                    cont.width = self._px[i]
                    if t is not None:
                        t.width = self._px[i]
            else:  # cabecera: cada segmento re-suma los px de sus columnas
                for cont, col, k in rec["segs"]:
                    cont.width = sum(self._px[col:col + k]) + self.spacing * (k - 1)
        self._safe_update()

    def _safe_update(self) -> None:
        try:
            self.control.update()
        except (RuntimeError, AssertionError):
            pass  # aún no montada; se reflejará al renderizar
