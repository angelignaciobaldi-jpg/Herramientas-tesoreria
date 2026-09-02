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

import asyncio
import os
import re
from collections import Counter
from datetime import date

import flet as ft

from core import (
    api, comprobantes as _comprobantes, cuentas_bancarias,
    exportador_devoluciones, lector_comprobantes, pdf_paginas,
    reporte_excel, solicitudes_devolucion,
)
from core.catalogo_bancos import banco_a_mostrar, banco_desde_clabe
from core.rpa_sipp import (
    BucleRpa, ControlRpa, ErrorSipp, RpaDetenido, SesionSipp,
    asegurar_navegador, necesita_navegador,
)
from ui.comun import (
    CENTRO, GRIS, ID_POR_EMPRESA, NARANJA, NOMBRES_EMPRESAS, ROJO, VERDE,
    fmt_monto, parse_monto, solo_digitos, tarjeta,
)
from ui.tabla_responsiva import DER, IZQ, ColumnaTabla, FilaDatos, TablaResponsiva

# Columnas de la tabla unificada de solicitudes/movimientos, por PORCENTAJE del
# ancho disponible (TablaResponsiva las convierte a px según el tamaño de ventana;
# si la suma supera el 100% —o los mínimos no caben— aparece scroll horizontal).
# Cada entrada: (etiqueta, pct, alineación, ancho_min_px). La col 0 es el check.
_COLS = [
    ("", 3, CENTRO, 44),
    ("Folio", 8, CENTRO, 90),
    ("Empresa solicitante", 12, CENTRO, 150),
    ("Cliente / Beneficiario", 14, IZQ, 150),
    ("CLABE Beneficiario", 12, CENTRO, 150),
    ("Banco destino", 8, CENTRO, 90),
    ("Monto", 8, DER, 90),
    ("Concepto", 10, IZQ, 130),
    ("Cuenta origen de pago", 11, CENTRO, 150),
    ("Comprobante", 8, CENTRO, 110),
    ("Acciones", 6, CENTRO, 90),
]

# Fondo de la fila cuya cuenta origen de pago ya fue asignada.
ASIGNADA = ft.Colors.with_opacity(0.14, ft.Colors.GREEN)
# Fondo de las filas de deudores diversos (empleados/otros), para distinguirlas
# de las devoluciones a clientes.
DEUDOR_DIVERSO = ft.Colors.with_opacity(0.16, ft.Colors.INDIGO)
# Fondo de las filas que se repiten (misma CLABE, monto, beneficiario y empresa
# que otra ya capturada): se pintan de rojo para que el usuario las detecte.
DUPLICADO = ft.Colors.with_opacity(0.22, ft.Colors.RED)

# Tipos de pago a deudor (etiqueta, valor que recibe la API): 1=Devoluciones, 0=Pagos.
TIPOS_PAGO_DEUDOR = [("Devoluciones pendientes", "1"), ("Pagos pendientes", "0")]


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
                 manual: bool = False, es_deudor: bool = False):
        self.seccion = seccion
        self.manual = manual
        # True si la fila es un deudor diverso (empleado/otro): se pinta distinto
        # y su folio es editable (a diferencia de los clientes, cuyo folio es fijo).
        self.es_deudor = es_deudor
        self.original = sol          # datos originales del SIPP (None si es manual)
        self.asignacion: dict | None = None  # cuenta origen de pago (paso 2)
        # Ruta del PDF del comprobante de pago vinculado a esta fila (paso 3).
        # None mientras no se haya cargado/casado ninguno.
        self.comprobante: str | None = None
        # True si el banco que reporta el SIPP contradice al de la CLABE.
        # Lo recalcula _pintar_banco() en cada cambio.
        self.banco_discrepa = False
        # Lectura del comprobante (cuentas, importe, fecha de aplicacion...):
        # de ella sale la Referencia que el RPA escribira en el SIPP.
        self.lectura_comprobante: dict | None = None
        # True si esta fila repite (CLABE+monto+beneficiario+empresa) a otra ya
        # capturada; lo recalcula la sección en cada cambio/consulta.
        self.es_duplicado = False
        # Color de fondo actual (str|None) y contenedor renderizado en la tabla
        # responsiva (para recolorear en vivo sin reconstruir; lo asigna la sección
        # tras pintar). None mientras no esté en la página visible.
        self.color: str | None = None
        self.contenedor = None
        # Casilla de selección de la fila (col 0). Reemplaza el check nativo del
        # DataTable: la selección vive aquí (no hay page.update por fila).
        self.chk_sel = ft.Checkbox(value=False, on_change=self._al_seleccionar)

        # Folio: cliente = texto fijo (folio del SIPP); deudor diverso = editable
        # (alfanumérico); manual = "— manual —" (no lleva folio). Sin anchos fijos:
        # las celdas llenan el ancho que les da la columna responsiva.
        if es_deudor:
            self.folio_ctrl = ft.TextField(
                value=(sol.folio if sol else ""), dense=True,
                text_size=12, content_padding=8, hint_text="folio",
                on_change=self._cambio,
            )
        else:
            self.folio_ctrl = ft.Text(
                (sol.folio if (sol and not manual) else "— manual —"), size=12,
                italic=manual, color=(GRIS if manual else None),
                text_align=ft.TextAlign.CENTER,
            )
        self.dd_empresa = ft.Dropdown(
            dense=True, text_size=12, content_padding=8,
            options=[ft.dropdown.Option(key=e, text=e)
                     for e in seccion.empresas_solicitantes()],
            on_select=self._cambio,
        )
        self.tf_cliente = ft.TextField(
            dense=True, text_size=12, content_padding=8,
            on_change=self._cambio,
        )
        # Sin max_length (evita el contador "X/18" que desalinea la fila); el
        # límite de 18 dígitos se aplica por código en _cambio_clabe.
        self.tf_clabe = ft.TextField(
            dense=True, text_size=12, content_padding=8,
            text_align=ft.TextAlign.CENTER, on_change=self._cambio_clabe,
        )
        self.txt_banco = ft.Text("—", size=12, text_align=ft.TextAlign.CENTER)
        # Banco tal como lo reportó el SIPP (se conserva aunque se edite la
        # CLABE, para poder seguir contrastando ambos datos).
        self.banco_reportado = ""
        self.tf_monto = ft.TextField(
            dense=True, text_size=12, content_padding=8,
            text_align=ft.TextAlign.RIGHT, hint_text="0.00", on_change=self._cambio,
        )
        self.tf_concepto = ft.TextField(
            dense=True, text_size=12, content_padding=8,
            on_change=self._cambio,
        )
        # Cuenta origen de pago (paso 2): dos líneas —empresa · banco (truncada) y
        # la CLABE completa—, con tooltip que muestra todo el detalle al pasar el
        # mouse (el nombre de la empresa suele ser largo y no cabe en la celda).
        self.txt_origen_emp = ft.Text(
            "— sin asignar —", size=11, color=GRIS, text_align=ft.TextAlign.CENTER,
            max_lines=1, no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS,
        )
        self.txt_origen_clabe = ft.Text(
            "", size=11, weight=ft.FontWeight.BOLD, text_align=ft.TextAlign.CENTER,
            max_lines=1, no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS,
        )
        self.txt_origen = ft.Column(
            [self.txt_origen_emp, self.txt_origen_clabe],
            spacing=0, tight=True,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        )

        # Celda de la columna 'Comprobante': un icono que resume el estado y
        # cuyo tooltip trae el detalle (archivo y referencia que se escribira).
        self.icono_comprobante = ft.Icon(
            ft.Icons.ATTACH_FILE, size=18, color=GRIS)
        self.txt_comprobante = ft.Text("— sin comprobante —", size=11,
                                       color=GRIS, no_wrap=True)
        self.celda_comprobante = ft.Row(
            [self.icono_comprobante, self.txt_comprobante],
            spacing=4, alignment=ft.MainAxisAlignment.CENTER, tight=True,
        )

        self.acciones = ft.Row(
            [
                ft.IconButton(
                    icon=ft.Icons.RESTORE, tooltip=(
                        "Limpiar el movimiento (empezar de cero)" if manual
                        else "Restaurar los datos originales de la solicitud"),
                    icon_color=ft.Colors.BLUE_700, icon_size=18,
                    width=38, height=38, style=ft.ButtonStyle(padding=0),
                    on_click=lambda e: self.restaurar(),
                ),
                ft.IconButton(
                    icon=ft.Icons.DELETE_OUTLINE, tooltip="Quitar de la tabla",
                    icon_color=ROJO, icon_size=18,
                    width=38, height=38, style=ft.ButtonStyle(padding=0),
                    on_click=lambda e: self.seccion.eliminar_fila(self),
                ),
            ],
            spacing=2, alignment=ft.MainAxisAlignment.CENTER, tight=True,
        )
        self.restaurar(avisar=False)  # carga los datos originales (o deja vacío)
        self.actualizar_color()  # color base (deudor índigo; duplicado rojo; etc.)

    def fila_datos(self) -> FilaDatos:
        """Fila para la tabla responsiva: una celda (control) por columna, en el
        orden de `_COLS`. El fondo lleva el color actual de la fila."""
        return FilaDatos(
            [self.chk_sel, self.folio_ctrl, self.dd_empresa, self.tf_cliente,
             self.tf_clabe, self.txt_banco, self.tf_monto, self.tf_concepto,
             self.txt_origen, self.celda_comprobante, self.acciones],
            bgcolor=self.color,
        )

    @property
    def folio_valor(self) -> str:
        """Texto del folio (fijo en clientes, editable en deudores)."""
        return (self.folio_ctrl.value or "").strip()

    # ------------------------------------------------------------- color
    def actualizar_color(self) -> None:
        """Fija el color de fondo según el estado, con esta prioridad: duplicado
        (rojo, para que resalte como advertencia) > cuenta origen asignada (verde)
        > deudor diverso (índigo) > sin color. Si la fila ya está renderizada, muta
        el fondo del contenedor en vivo (sin reconstruir, sin perder el foco)."""
        if self.es_duplicado:
            self.color = DUPLICADO
        elif self.asignacion:
            self.color = ASIGNADA
        elif self.es_deudor:
            self.color = DEUDOR_DIVERSO
        else:
            self.color = None
        if self.contenedor is not None:
            self.contenedor.bgcolor = self.color

    def clave_duplicado(self) -> tuple:
        """Firma que identifica un pago repetido: misma CLABE, mismo monto, mismo
        beneficiario y misma empresa (beneficiario/empresa sin distinguir may/mín)."""
        clabe, monto_txt, cliente, _concepto, empresa = self.valores()
        try:
            monto = parse_monto(monto_txt)
        except ValueError:
            monto = None
        return (clabe, monto, cliente.strip().upper(), (empresa or "").strip().upper())

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
        # En deudores el folio es editable: se restaura a su original (vacío).
        if self.es_deudor:
            self.folio_ctrl.value = sol.folio if sol else ""
        # El banco lo reporta el SIPP; _pintar_banco lo contrasta con la CLABE
        # y marca la celda si se contradicen.
        self._pintar_banco(sol.banco if sol else "")
        if avisar:
            self.seccion._refrescar()

    def _cambio(self, _e=None) -> None:
        # Al editar (monto, cliente, empresa…) puede aparecer o resolverse un
        # duplicado: se re-evalúa toda la tabla y se repintan las filas.
        self.seccion._revisar_duplicados()
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
        # Cambió la CLABE: se revalúa contra el banco que reportó el SIPP (la
        # discrepancia puede aparecer o resolverse al corregirla).
        self._pintar_banco()
        self._cambio()

    # ------------------------------------------------------ banco destino
    def _pintar_banco(self, reportado: str | None = None) -> None:
        """Muestra el banco destino, CORRIGIÉNDOLO desde la CLABE si hace falta.

        El SIPP guarda el banco como texto, capturado aparte de la CLABE, así que
        los dos datos pueden contradecirse (se han visto solicitudes con CLABE de
        BanCoppel y "BBVA BANCOMER" como banco).

        Cuando se contradicen se muestra el banco que dice la CLABE, no el del
        SIPP. Es el único que describe lo que va a pasar: el pago viaja por la
        CLABE y el TXT se arma con su prefijo (ver
        `core.exportador_devoluciones.linea_bancomer`, que elige entre el layout
        de mismo banco y el de SPEI justamente por ahí). Dejar a la vista el del
        SIPP hacía creer que el dinero iría a un banco al que no va.

        La corrección es SOLO visual y se señala en naranja: el registro sigue
        necesitando revisión, porque la contradicción también puede significar que
        la equivocada sea la CLABE —y entonces el pago se iría a otro lado—. El
        tooltip conserva las dos versiones para poder verificarlo.

        `reportado` es el banco que dio el SIPP; con None se usa el que ya está
        guardado (al reevaluar tras editar la CLABE).
        """
        clabe = solo_digitos(self.tf_clabe.value)
        if reportado is None:
            reportado = self.banco_reportado
        self.banco_reportado = reportado or ""
        segun_clabe = banco_desde_clabe(clabe)
        mostrado, self.banco_discrepa = banco_a_mostrar(
            clabe, self.banco_reportado)
        self.txt_banco.value = mostrado or "—"
        if self.banco_discrepa:
            self.txt_banco.color = NARANJA
            self.txt_banco.weight = ft.FontWeight.BOLD
            self.txt_banco.tooltip = (
                f"Corregido desde la CLABE.\n"
                f"La CLABE ({clabe[:3]}…) es de {segun_clabe}, pero el SIPP "
                f"reportó «{self.banco_reportado}».\n"
                "Se muestra el de la CLABE porque es el que determina a dónde "
                "llega el pago y con el que se arma el TXT.\n"
                "Aun así revisa la solicitud: si la equivocada fuera la CLABE, "
                "el dinero iría a otro banco.")
        else:
            self.txt_banco.color = None
            self.txt_banco.weight = None
            self.txt_banco.tooltip = None
    # ------------------------------------------------------- selección
    def _al_seleccionar(self, _e=None) -> None:
        # El checkbox ya refleja su estado en el cliente; solo se refresca el
        # contador (sin reconstruir la tabla).
        self.seccion._actualizar_contador()
        self.seccion.page.update()

    @property
    def seleccionada(self) -> bool:
        return bool(self.chk_sel.value)

    def set_seleccion(self, valor: bool) -> None:
        self.chk_sel.value = bool(valor)

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
    def _pintar_origen(self) -> None:
        """Refleja la cuenta origen asignada en las dos líneas de la celda
        (empresa · banco arriba, CLABE abajo) y arma el tooltip con todo el
        detalle (empresa, banco, CLABE y número de cuenta) para verlo completo."""
        a = self.asignacion
        if a:
            emp = f"{a['empresa_origen']} · {a['banco']}"
            self.txt_origen_emp.value = emp
            self.txt_origen_emp.color = None
            self.txt_origen_clabe.value = a["cuenta_origen"]
            detalle = f"{emp}\nCLABE: {a['cuenta_origen']}"
            if a.get("num_cuenta"):
                detalle += f"\nNo. de cuenta: {a['num_cuenta']}"
            self.txt_origen.tooltip = detalle
        else:
            self.txt_origen_emp.value = "— sin asignar —"
            self.txt_origen_emp.color = GRIS
            self.txt_origen_clabe.value = ""
            self.txt_origen.tooltip = None

    def asignar(self, datos: dict) -> None:
        self.asignacion = datos
        self._pintar_origen()
        self.actualizar_color()

    # --------------------------------------------------- comprobante
    def _pintar_comprobante(self) -> None:
        """Refleja en la celda si la fila ya tiene comprobante vinculado. El
        tooltip lleva el archivo y la referencia que se escribira en el SIPP, que
        es lo que conviene revisar antes de subir."""
        if not self.comprobante:
            self.icono_comprobante.name = ft.Icons.ATTACH_FILE
            self.icono_comprobante.color = GRIS
            self.txt_comprobante.value = "— sin comprobante —"
            self.txt_comprobante.color = GRIS
            self.celda_comprobante.tooltip = None
            return
        nombre = os.path.basename(self.comprobante)
        self.icono_comprobante.name = ft.Icons.CHECK_CIRCLE
        self.icono_comprobante.color = ft.Colors.GREEN_700
        self.txt_comprobante.value = nombre[:18] + ("…" if len(nombre) > 18 else "")
        self.txt_comprobante.color = None
        detalle = f"Comprobante: {nombre}"
        ref = lector_comprobantes.referencia_aaaammdd(self.lectura_comprobante or {})
        if ref:
            detalle += f"\nReferencia que se escribira: {ref}"
        else:
            detalle += ("\nSIN fecha de aplicacion: habra que capturar la "
                        "referencia a mano.")
        self.celda_comprobante.tooltip = detalle

    def vincular_comprobante(self, ruta: str, lectura: dict | None = None) -> None:
        self.comprobante = ruta
        self.lectura_comprobante = lectura
        self._pintar_comprobante()

    def quitar_comprobante(self) -> None:
        self.comprobante = None
        self.lectura_comprobante = None
        self._pintar_comprobante()

    def objetivo_comprobante(self) -> "_comprobantes.Objetivo | None":
        """Qué hay que casar de ESTA fila, o None si aun no se puede casar.

        Se necesitan CLABE destino, monto y una cuenta origen asignada; sin esta
        ultima ningun comprobante puede coincidir (la regla de cuenta origen
        fallaria siempre). Se aportan CLABE y numero de cuenta del origen: el
        comprobante trae el NUMERO de cuenta, y como la CLABE termina en digito
        verificador, comparar solo contra ella nunca casaria."""
        clabe, monto_txt, _cliente, _concepto, _empresa = self.valores()
        a = self.asignacion or {}
        origenes = {o for o in (a.get("cuenta_origen"), a.get("num_cuenta")) if o}
        if not (clabe and origenes):
            return None
        try:
            total = float((monto_txt or "").replace(",", ""))
        except ValueError:
            return None
        return _comprobantes.Objetivo(
            origenes=origenes, beneficiarios={clabe}, total=total)

    def limpiar_asignacion(self) -> None:
        self.asignacion = None
        self._pintar_origen()
        # Al quitar la asignación, la fila vuelve a su color de origen.
        self.actualizar_color()


class SeccionDevoluciones:
    """Consulta de solicitudes (SIPP) + captura manual, asignación de cuenta
    origen de pago y generación de un TXT por empresa origen."""

    # Empresa y sucursal con que se INICIA LA SESIÓN del SIPP. Es solo la puerta
    # de entrada al portal, no un filtro: el listado de solicitudes autorizadas
    # muestra las de todas las empresas, así que no hace falta —ni conviene—
    # deducirla de cada solicitud. Deducirla fallaba: la sucursal no viene en la
    # API y el nombre que se armaba ("<Empresa> Corporativo") no existe para
    # todas. Mismo criterio que la pantalla de Dispersión (No Pemex).
    EMPRESA_SESION = "Aske"
    SUCURSAL_SESION = "Corporativo"

    def __init__(self, app):
        # --- RPA de subida de comprobantes al SIPP ---
        # Hilo con su propio loop (el RPA es asincrono y Flet no es
        # thread-safe); se crea la primera vez que se usa.
        self.bucle_rpa = None
        self.sesion_rpa = None
        self._sub_ctrl = None
        self._sub_corriendo = False
        # id(fila) de las que ya se registraron en esta corrida, para no
        # reintentarlas si se vuelve a pulsar el boton.
        self._sub_registradas: set = set()
        self.app = app
        self.page = app.page
        self.filas: list[FilaSolicitud] = []
        # Empresas y sus IDs: fuente única en ui/comun.EMPRESAS. Por defecto se
        # consultan todas (todas marcadas). Selección independiente para la
        # consulta de clientes y la de deudores diversos.
        self._empresas_val = {e: True for e in NOMBRES_EMPRESAS}
        self._empresas_deudor_val = {e: True for e in NOMBRES_EMPRESAS}
        self.catalogo = cuentas_bancarias.CatalogoCuentas()
        self.contenido = self._construir()

    def empresas_solicitantes(self) -> list[str]:
        """Empresas del desplegable 'Empresa solicitante' (las de ui/comun)."""
        return list(NOMBRES_EMPRESAS)

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
        # Update DIRIGIDO a los controles afectados (no page.update() de página
        # completa): se recarga con el modal de Configuración encima, así que
        # re-renderizar toda la app sería innecesario y podría congelar la UI.
        for control in (self.dd_empresa, self.txt_sin_catalogo,
                        self.dd_origen, self.tf_num_cuenta):
            try:
                control.update()
            except (RuntimeError, AssertionError):
                pass  # aún no montado; se reflejará al renderizar

    # ================================================================ UI
    def _construir_consulta(self) -> ft.Control:
        """Sección 1: consulta de solicitudes. Un desplegable elige el TIPO de
        beneficiario (Clientes / Deudores diversos) y, según la elección, se
        muestran los campos ya mapeados de esa consulta."""
        # --- Bloque CLIENTES (empresas + rango de fechas) ---
        self.tf_empresas = ft.TextField(
            label="Empresas a consultar", width=280, read_only=True,
            value=self._resumen(self._empresas_val), suffix_icon=ft.Icons.ARROW_DROP_DOWN,
            on_click=lambda e: self._abrir_selector(self._empresas_val, self.tf_empresas),
        )
        self.btn_consultar = ft.FilledButton(
            content="Consultar devoluciones clientes", icon=ft.Icons.SEARCH,
            on_click=self._consultar,
        )
        # Rango de fechas de la consulta (opcional): el servicio filtra por
        # fechaInicio/fechaFin. Campos de solo lectura con calendario (como el RPA).
        self.dp_fi = ft.DatePicker(
            first_date=date(2020, 1, 1), last_date=date(2035, 12, 31),
            help_text="Fecha inicio de la consulta",
            on_change=lambda e: self._fecha_consulta(self.tf_fi, self.dp_fi),
        )
        self.dp_ff = ft.DatePicker(
            first_date=date(2020, 1, 1), last_date=date(2035, 12, 31),
            help_text="Fecha fin de la consulta",
            on_change=lambda e: self._fecha_consulta(self.tf_ff, self.dp_ff),
        )
        self.tf_fi = ft.TextField(
            label="Fecha inicio", hint_text="DD/MM/AAAA", width=160, read_only=True,
            dense=True, content_padding=10, suffix_icon=ft.Icons.CALENDAR_MONTH,
            on_click=lambda e: self.page.show_dialog(self.dp_fi),
        )
        self.tf_ff = ft.TextField(
            label="Fecha fin", hint_text="DD/MM/AAAA", width=160, read_only=True,
            dense=True, content_padding=10, suffix_icon=ft.Icons.CALENDAR_MONTH,
            on_click=lambda e: self.page.show_dialog(self.dp_ff),
        )
        self.btn_limpiar_fechas = ft.IconButton(
            icon=ft.Icons.CLEAR, tooltip="Limpiar fechas (consultar sin rango)",
            icon_color=GRIS, on_click=self._limpiar_fechas_consulta,
        )
        self.bloque_clientes = ft.Column(
            [
                ft.Row(
                    [self.tf_empresas, self.tf_fi, self.tf_ff,
                     self.btn_limpiar_fechas, self.btn_consultar],
                    spacing=10, wrap=True,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                ),
                ft.Text(
                    "Consulta las devoluciones autorizadas de clientes por una o "
                    "varias empresas y (opcional) por rango de fechas. Los datos "
                    "llegan llenos pero puedes editarlos; el botón ↺ restaura el "
                    "original.",
                    color=GRIS, size=12, italic=True,
                ),
            ],
            spacing=12,
        )

        # --- Bloque DEUDORES DIVERSOS (tipo de pago + empresas) ---
        self.dd_tipo_deudor = ft.Dropdown(
            label="Tipo de pago a deudor", width=230, value="1",
            options=[ft.dropdown.Option(key=v, text=t) for t, v in TIPOS_PAGO_DEUDOR],
        )
        # Selector de empresas MULTI-selección (igual que el de clientes), con su
        # propia selección independiente (_empresas_deudor_val).
        self.tf_empresas_deudor = ft.TextField(
            label="Empresas a consultar", width=280, read_only=True,
            value=self._resumen(self._empresas_deudor_val),
            suffix_icon=ft.Icons.ARROW_DROP_DOWN,
            on_click=lambda e: self._abrir_selector(
                self._empresas_deudor_val, self.tf_empresas_deudor),
        )
        self.btn_consultar_deudor = ft.FilledButton(
            content="Consultar pagos pendientes", icon=ft.Icons.SEARCH,
            on_click=self._consultar_deudores,
        )
        self.bloque_deudores = ft.Column(
            [
                ft.Row(
                    [self.dd_tipo_deudor, self.tf_empresas_deudor,
                     self.btn_consultar_deudor],
                    spacing=10, wrap=True,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER,
                ),
                ft.Text(
                    "Consulta devoluciones/pagos pendientes a colaboradores u otros "
                    "deudores de la empresa elegida. Se agregan a la misma tabla en "
                    "color índigo; su folio y concepto son editables.",
                    color=GRIS, size=12, italic=True,
                ),
            ],
            spacing=12,
            visible=False,  # arranca en Clientes
        )

        # --- Desplegable que elige el tipo de beneficiario a consultar ---
        self.dd_tipo_beneficiario = ft.Dropdown(
            label="Tipo de beneficiario", width=280, value="clientes",
            options=[
                ft.dropdown.Option(key="clientes", text="Clientes"),
                ft.dropdown.Option(key="deudores", text="Deudores diversos"),
            ],
            on_select=self._cambio_tipo_beneficiario,
        )
        return tarjeta(
            "1. Consultar solicitudes (elige el tipo de beneficiario)",
            ft.Column(
                [self.dd_tipo_beneficiario, self.bloque_clientes, self.bloque_deudores],
                spacing=12,
            ),
        )

    def _cambio_tipo_beneficiario(self, _e=None) -> None:
        """Muestra el bloque de campos correspondiente al tipo de beneficiario
        elegido (Clientes o Deudores diversos)."""
        es_clientes = self.dd_tipo_beneficiario.value != "deudores"
        self.bloque_clientes.visible = es_clientes
        self.bloque_deudores.visible = not es_clientes
        self.page.update()

    def _construir_registros(self) -> ft.Control:
        """Tarjeta de los registros a dispersar (clientes + deudores + manuales),
        con las acciones que aplican a TODA la tabla, sin importar el origen.
        El botón 'Generar dispersión' va arriba a la derecha; abajo, solo la tabla."""
        self.txt_contador = ft.Text("Sin registros.", size=12, color=GRIS,
                                    weight=ft.FontWeight.BOLD)
        # Estado de la generación (cuántos TXT se generarán). Vive aquí porque el
        # botón 'Generar dispersión' ahora está en esta tarjeta.
        self.txt_grupos = ft.Text("Sin filas asignadas.", size=12, color=GRIS)
        # Tabla RESPONSIVA: columnas por porcentaje del ancho (se adaptan a la
        # ventana; scroll horizontal si no caben). Mismo modelo que Dispersión (No
        # Pemex). El check de selección va en la columna 0.
        columnas = [
            ColumnaTabla(etq, pct, alin, ancho_min_px=minpx)
            for etq, pct, alin, minpx in _COLS
        ]
        self.tabla = TablaResponsiva(
            self.page, columnas, alto_fila=58,
            ancho_inicial=(getattr(self.page, "width", None) or 1200) - 90)

        def swatch(color, texto):
            return ft.Row(
                [ft.Container(width=14, height=14, bgcolor=color, border_radius=3),
                 ft.Text(texto, size=12, color=GRIS)],
                spacing=5, tight=True,
            )
        leyenda = ft.Row(
            [ft.Text("Leyenda:", size=12, weight=ft.FontWeight.BOLD, color=GRIS),
             swatch(DEUDOR_DIVERSO, "Deudor diverso (empleado/otro)"),
             swatch(ASIGNADA, "Con cuenta origen asignada"),
             swatch(DUPLICADO, "Registro duplicado (revisar)")],
            spacing=18, wrap=True, vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        acciones = ft.Row(
            [
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
                    content="Mover cuentas a alta de beneficiarios",
                    icon=ft.Icons.MOVE_UP,
                    on_click=self._mover_a_alta,
                    tooltip="Copia las filas seleccionadas al módulo de Alta de "
                            "beneficiarios (CLABE, beneficiario, alias y banco).",
                ),
                ft.OutlinedButton(
                    content="Cargar comprobantes", icon=ft.Icons.UPLOAD_FILE,
                    on_click=self._cargar_comprobantes,
                    tooltip="Lee los PDF de comprobantes de pago y los vincula "
                            "con su registro por cuenta origen, cuenta destino "
                            "e importe.",
                ),
                ft.OutlinedButton(
                    content="Subir comprobantes a SIPP",
                    icon=ft.Icons.CLOUD_UPLOAD,
                    on_click=self._subir_comprobantes,
                    tooltip="Registra en el SIPP las devoluciones que ya "
                            "tienen comprobante vinculado: adjunta el PDF, "
                            "escribe la referencia y guarda.",
                ),
                ft.OutlinedButton(
                    content="Generar Excel", icon=ft.Icons.TABLE_VIEW,
                    on_click=self._generar_excel,
                ),
                ft.OutlinedButton(
                    content="Eliminar todos", icon=ft.Icons.DELETE_SWEEP_OUTLINED,
                    on_click=self._eliminar_todos,
                    style=ft.ButtonStyle(color=ROJO),
                ),
                self.txt_contador,
            ],
            spacing=10, wrap=True,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        # 'Generar dispersión' a la MISMA altura que el título de la tarjeta (arriba
        # a la derecha), con el resumen de cuántos TXT se generarán al lado. Abajo
        # queda solo la tabla. Se arma la tarjeta a mano (el helper `tarjeta` pone el
        # título en su propia línea).
        self.btn_generar = ft.FilledButton(
            content="Generar dispersión", icon=ft.Icons.CREATE_NEW_FOLDER,
            on_click=self._generar_dispersion,
            tooltip="Genera un archivo TXT por cada empresa origen de pago (y banco). "
                    "No se genera si alguna CLABE no tiene 18 dígitos o falta el monto.",
        )
        titulo_row = ft.Row(
            [
                ft.Text("Registros a dispersar (clientes + deudores + manuales)",
                        weight=ft.FontWeight.BOLD, size=15, expand=True),
                self.txt_grupos,
                self.btn_generar,
            ],
            spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        return ft.Card(
            content=ft.Container(
                content=ft.Column(
                    [titulo_row, acciones, leyenda, self.tabla.control],
                    spacing=12, horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
                ),
                padding=16,
            )
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

        # Orden nuevo: 1) consulta por tipo de beneficiario, 2) cuenta origen de
        # pago (arriba de la tabla, para no subir/bajar), 3) registros con la tabla
        # abajo y 'Generar dispersión' arriba a la derecha de esa tarjeta.
        return ft.Column(
            [self._construir_consulta(), config, self._construir_registros()],
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
    @staticmethod
    def _resumen(val: dict) -> str:
        """Texto del campo desplegable según las empresas marcadas en `val`."""
        sel = [e for e, v in val.items() if v]
        total = len(val)
        if not sel:
            return "Ninguna"
        if len(sel) == total:
            return f"Todas ({total})"
        if len(sel) == 1:
            return sel[0]
        return f"{len(sel)} empresas"

    def _resumen_empresas(self) -> str:  # consulta de clientes (compat)
        return self._resumen(self._empresas_val)

    def _abrir_selector(self, val: dict, campo: ft.TextField) -> None:
        """Despliega las empresas (casillas) para elegir una, varias o todas; al
        cerrar, refleja la selección en `campo`. Reutilizable para clientes y para
        deudores (cada uno con su propio `val`)."""
        casillas = [
            ft.Checkbox(
                label=e, value=v,
                on_change=lambda ev, emp=e: val.__setitem__(emp, ev.control.value),
            )
            for e, v in val.items()
        ]

        def todas(_ev=None):
            for c in casillas:
                c.value = True
                val[c.label] = True
            self.page.update()

        def ninguna(_ev=None):
            for c in casillas:
                c.value = False
                val[c.label] = False
            self.page.update()

        def cerrar(_ev=None):
            campo.value = self._resumen(val)
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

    def _fecha_consulta(self, campo: ft.TextField, dp: ft.DatePicker) -> None:
        """Vuelca la fecha elegida en el calendario al campo, como DD/MM/AAAA."""
        if dp.value:
            campo.value = dp.value.strftime("%d/%m/%Y")
            self.page.update()

    def _limpiar_fechas_consulta(self, _e=None) -> None:
        """Borra el rango de fechas (para consultar sin filtro de fecha)."""
        self.dp_fi.value = None
        self.dp_ff.value = None
        self.tf_fi.value = ""
        self.tf_ff.value = ""
        self.page.update()

    def _consultar_api(self, nombres: list[str], fecha_inicio: str | None,
                       fecha_fin: str | None) -> list:
        """Consulta la API (una llamada por empresa, usando su id de ui/comun) y
        junta todas las solicitudes. Corre en un hilo (lo llama _consultar con
        asyncio.to_thread) porque hace red y no debe congelar la interfaz."""
        todas = []
        for nombre in nombres:
            id_empresa = ID_POR_EMPRESA.get(nombre)
            if id_empresa is None:
                continue
            todas.extend(solicitudes_devolucion.consultar(
                id_empresa, nombre, fecha_inicio, fecha_fin))
        return todas

    async def _consultar(self, _e=None) -> None:
        """Trae las solicitudes autorizadas de las empresas elegidas desde la API,
        opcionalmente filtradas por rango de fechas. Conserva los movimientos
        manuales ya capturados (solo reemplaza las del SIPP)."""
        nombres = [e for e, v in self._empresas_val.items() if v]
        if not nombres:
            self.app.avisar("Elige al menos una empresa para consultar.", ROJO)
            return
        # Rango de fechas (opcional). El servicio las espera como 'YYYY-MM-DD'.
        fi, ff = self.dp_fi.value, self.dp_ff.value
        if fi and ff and fi > ff:
            self.app.avisar(
                "La fecha inicio no puede ser posterior a la fecha fin.", ROJO)
            return
        fecha_inicio = fi.strftime("%Y-%m-%d") if fi else None
        fecha_fin = ff.strftime("%Y-%m-%d") if ff else None

        self.btn_consultar.disabled = True
        self.txt_contador.value = "Consultando solicitudes…"
        self.page.update()
        try:
            resultados = await asyncio.to_thread(
                self._consultar_api, nombres, fecha_inicio, fecha_fin)
        except api.ApiSinConexion as exc:
            self._fin_consulta()
            self.app.avisar(
                "No se pudo conectar con el servicio de solicitudes. Revisa tu "
                f"conexión e inténtalo de nuevo. ({exc})", ROJO)
            return
        except api.ErrorRespuestaApi as exc:
            self._fin_consulta()
            self.app.avisar(
                f"El servicio respondió con error (HTTP {exc.status}). "
                f"{self._detalle_error(exc)}", ROJO)
            return
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self._fin_consulta()
            self.app.avisar(f"No se pudo consultar: {exc}", ROJO)
            return
        self._fin_consulta()
        # ACUMULA: agrega las nuevas solicitudes de clientes a lo ya capturado
        # (no borra consultas previas, deudores ni manuales). Los repetidos se
        # marcan en rojo al refrescar para que el usuario los detecte.
        self.filas = self.filas + [FilaSolicitud(self, s) for s in resultados]
        self._refrescar()
        dup = self._num_duplicados()
        aviso_dup = f" · {dup} posible(s) duplicado(s) en rojo" if dup else ""
        if resultados:
            self.app.avisar(
                f"{len(resultados)} solicitud(es) de clientes agregada(s).{aviso_dup}",
                NARANJA if dup else VERDE)
        else:
            self.app.avisar(
                "No hay solicitudes de clientes para las empresas elegidas.", NARANJA)

    @staticmethod
    def _detalle_error(exc: "api.ErrorRespuestaApi") -> str:
        """Mensaje legible del error del servicio: usa el 'msg' del cuerpo si viene
        (p. ej. 'El query param X es requerido'); si es 401/403 sugiere el token."""
        cuerpo = exc.datos
        if isinstance(cuerpo, dict) and cuerpo.get("msg"):
            return str(cuerpo["msg"])
        if isinstance(cuerpo, str) and cuerpo.strip():
            return cuerpo.strip()[:200]
        if exc.status in (401, 403):
            return "Revisa el token de la API en Configuración (⚙)."
        return ""

    def _fin_consulta(self) -> None:
        self.btn_consultar.disabled = False
        self._actualizar_contador()
        self.page.update()

    def _consultar_deudores_api(self, nombres: list[str], tipo: int) -> list:
        """Consulta la API de deudores (una llamada por empresa, usando su id) y
        junta todo. Corre en un hilo (lo llama _consultar_deudores)."""
        todas = []
        for nombre in nombres:
            id_empresa = ID_POR_EMPRESA.get(nombre)
            if id_empresa is None:
                continue
            todas.extend(
                solicitudes_devolucion.consultar_deudores(id_empresa, nombre, tipo))
        return todas

    async def _consultar_deudores(self, _e=None) -> None:
        """Consulta las devoluciones/pagos pendientes a deudores diversos de las
        empresas elegidas (multi-selección) y las agrega a la tabla (en índigo).
        Reemplaza SOLO las de deudores; conserva las de clientes y las manuales."""
        nombres = [e for e, v in self._empresas_deudor_val.items() if v]
        if not nombres:
            self.app.avisar(
                "Elige al menos una empresa para consultar deudores.", ROJO)
            return
        try:
            tipo = int(self.dd_tipo_deudor.value)
        except (TypeError, ValueError):
            tipo = 1

        self.btn_consultar_deudor.disabled = True
        self.txt_contador.value = "Consultando pagos pendientes…"
        self.page.update()
        try:
            resultados = await asyncio.to_thread(
                self._consultar_deudores_api, nombres, tipo)
        except api.ApiSinConexion as exc:
            self._fin_consulta_deudores()
            self.app.avisar(
                "No se pudo conectar con el servicio de deudores. Revisa tu "
                f"conexión e inténtalo de nuevo. ({exc})", ROJO)
            return
        except api.ErrorRespuestaApi as exc:
            self._fin_consulta_deudores()
            self.app.avisar(
                f"El servicio respondió con error (HTTP {exc.status}). "
                f"{self._detalle_error(exc)}", ROJO)
            return
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self._fin_consulta_deudores()
            self.app.avisar(f"No se pudo consultar deudores: {exc}", ROJO)
            return
        self._fin_consulta_deudores()
        # ACUMULA: agrega los nuevos deudores a lo ya capturado (no borra
        # consultas previas, clientes ni manuales). Los repetidos se marcan en rojo.
        self.filas = self.filas + [
            FilaSolicitud(self, s, es_deudor=True) for s in resultados]
        self._refrescar()
        dup = self._num_duplicados()
        aviso_dup = f" · {dup} posible(s) duplicado(s) en rojo" if dup else ""
        if resultados:
            self.app.avisar(
                f"{len(resultados)} pago(s) a deudores agregado(s).{aviso_dup}",
                NARANJA if dup else VERDE)
        else:
            self.app.avisar(
                "No hay pagos pendientes a deudores para esa empresa.", NARANJA)

    def _fin_consulta_deudores(self) -> None:
        self.btn_consultar_deudor.disabled = False
        self._actualizar_contador()
        self.page.update()

    def _agregar_manual(self, _e=None) -> None:
        """Agrega una fila vacía para capturar un movimiento a mano (sin folio)."""
        self.filas.append(FilaSolicitud(self, sol=None, manual=True))
        self._refrescar()

    def eliminar_fila(self, fila: FilaSolicitud) -> None:
        if fila in self.filas:
            self.filas.remove(fila)
        self._refrescar()

    def _eliminar_todos(self, _e=None) -> None:
        """Vacía la tabla (solicitudes consultadas + movimientos manuales), con
        confirmación por lo que se pierde (ediciones y capturas manuales)."""
        if not self.filas:
            self.app.avisar("No hay registros que eliminar.", GRIS)
            return

        def confirmar(_ev=None):
            cuantos = len(self.filas)
            self.filas = []
            self.page.pop_dialog()
            self._refrescar()
            self.app.avisar(f"{cuantos} registro(s) eliminado(s) de la tabla.", GRIS)

        self.page.show_dialog(
            ft.AlertDialog(
                modal=True,
                title=ft.Text("Eliminar todos los registros"),
                content=ft.Text(
                    f"¿Quitar los {len(self.filas)} registro(s) de la tabla "
                    "(solicitudes y movimientos manuales)? Se perderán las ediciones "
                    "y capturas. Las solicitudes del SIPP puedes volver a consultarlas."
                ),
                actions=[
                    ft.TextButton("Cancelar", on_click=lambda e: self.page.pop_dialog()),
                    ft.FilledButton("Eliminar todos", on_click=confirmar,
                                    color=ft.Colors.WHITE, bgcolor=ROJO),
                ],
                actions_alignment=ft.MainAxisAlignment.END,
            )
        )

    def _seleccionar_todas(self, _e=None) -> None:
        if not self.filas:
            return
        nuevo = not all(f.seleccionada for f in self.filas)
        for f in self.filas:
            f.set_seleccion(nuevo)
        self._actualizar_contador()
        self.page.update()

    def _revisar_duplicados(self) -> int:
        """Marca (en rojo) las filas que se repiten: mismo CLABE, monto,
        beneficiario y empresa. Todas las filas de un grupo repetido quedan
        marcadas para que el usuario compare y decida. Devuelve cuántas filas
        quedaron marcadas como duplicadas."""
        grupos: dict[tuple, list[FilaSolicitud]] = {}
        for f in self.filas:
            if f.vacia():
                f.es_duplicado = False
                f.actualizar_color()
                continue
            grupos.setdefault(f.clave_duplicado(), []).append(f)
        marcadas = 0
        for fs in grupos.values():
            duplicado = len(fs) > 1
            for f in fs:
                f.es_duplicado = duplicado
                f.actualizar_color()
                if duplicado:
                    marcadas += 1
        return marcadas

    def _refrescar(self) -> None:
        self._revisar_duplicados()
        self.tabla.set_contenido([f.fila_datos() for f in self.filas])
        # Enlaza cada fila con su contenedor renderizado para poder recolorearla
        # en vivo (al editar) sin reconstruir la tabla.
        for i, f in enumerate(self.filas):
            f.contenedor = self.tabla.contenedor_fila(i)
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

    def _num_duplicados(self) -> int:
        """Cuántas filas están marcadas como duplicadas (para el contador/avisos)."""
        return sum(1 for f in self.filas if f.es_duplicado)

    def _actualizar_contador(self) -> None:
        total = len(self.filas)
        if not total:
            self.txt_contador.value = "Sin registros."
            self.txt_grupos.value = "Sin filas asignadas."
            return
        deudores = sum(1 for f in self.filas if f.es_deudor)
        manuales = sum(1 for f in self.filas if f.manual)
        seleccionadas = sum(1 for f in self.filas if f.seleccionada)
        asignadas = sum(1 for f in self.filas if f.asignacion)
        duplicados = self._num_duplicados()
        aviso_dup = f" · {duplicados} duplicado(s) en rojo" if duplicados else ""
        # El banco que reporta el SIPP puede contradecir a la CLABE; se cuenta
        # aparte porque no impide dispersar, pero hay que revisarlo antes.
        discrepan = sum(1 for f in self.filas if f.banco_discrepa)
        aviso_banco = (f" · {discrepan} con banco corregido desde la CLABE"
                       if discrepan else "")
        self.txt_contador.value = (
            f"{total} registro(s) ({deudores} deudor(es), {manuales} manual(es)) · "
            f"{seleccionadas} seleccionado(s) · {asignadas} asignado(s)"
            f"{aviso_dup}{aviso_banco}"
        )
        grupos = self._grupos_asignados()
        self.txt_grupos.value = (
            f"{len(grupos)} archivo(s) TXT a generar, con {asignadas} registro(s)."
            if grupos else "Sin filas asignadas."
        )

    # ------------------------------------------------------- asignación
    def _mover_a_alta(self, _e=None) -> None:
        """Copia las solicitudes SELECCIONADAS al módulo de Alta de beneficiarios
        (para dar de alta las cuentas). Toma los valores actuales (editados) de
        CLABE, beneficiario y monto; el alias sale igual al beneficiario y el banco
        se deduce de la CLABE en el otro módulo. No las quita de devoluciones."""
        seleccionadas = [f for f in self.filas if f.seleccionada]
        if not seleccionadas:
            self.app.avisar(
                "Selecciona al menos una fila (marca las casillas).", ROJO)
            return
        registros, omitidos = [], 0
        for f in seleccionadas:
            clabe, monto_txt, cliente, _concepto, _empresa = f.valores()
            if not clabe and not cliente:
                omitidos += 1
                continue
            try:
                monto = parse_monto(monto_txt)
            except ValueError:
                monto = None  # el monto es opcional/editable en el otro módulo
            registros.append({"clabe": clabe, "beneficiario": cliente, "monto": monto})
        if not registros:
            self.app.avisar(
                "Las filas seleccionadas no tienen CLABE ni beneficiario.", ROJO)
            return
        try:
            agregados = self.app.alta.importar_desde_devoluciones(registros)
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.app.avisar(f"No se pudieron mover las cuentas: {exc}", ROJO)
            return
        # Lleva al usuario a la pestaña de Alta para que vea lo importado (morado).
        try:
            self.app._seleccionar_nav(0)
        except Exception:  # noqa: BLE001 — la navegación no debe romper el flujo
            pass
        extra = f" ({omitidos} sin datos, omitida(s))" if omitidos else ""
        self.app.avisar(
            f"{agregados} cuenta(s) movida(s) a Alta de beneficiarios{extra}. "
            "Aparecen en morado.", VERDE)

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

    # --------------------------------------------------- comprobantes
    async def _cargar_comprobantes(self, _e=None) -> None:
        """Lee un lote de comprobantes (PDF) y los vincula con sus registros.

        Los PDF de varias páginas se separan ANTES de leerlos: el banco entrega
        un solo archivo con un comprobante por página, y al SIPP hay que subir la
        página que corresponde a cada solicitud, no el documento completo.

        La lectura es LOCAL (core.lector_comprobantes): no necesita red ni token,
        y es la única que trae la fecha de aplicación, con la que se arma la
        referencia que el RPA escribirá en el SIPP.
        """
        pendientes = [f for f in self.filas if not f.vacia()]
        if not pendientes:
            self.app.avisar(
                "No hay registros a los que vincular comprobantes.", ROJO)
            return
        # Sin cuenta origen asignada no hay forma de casar (la regla de cuenta
        # origen fallaría siempre): se avisa antes de hacer trabajo inútil.
        objetivos = [(id(f), f.objetivo_comprobante()) for f in pendientes]
        casables = [(k, o) for k, o in objetivos if o is not None]
        if not casables:
            self.app.avisar(
                "Ningún registro se puede casar todavía: falta asignarles la "
                "cuenta origen de pago (o les falta CLABE/monto).", ROJO)
            return

        archivos = await self.app.picker.pick_files(
            dialog_title="Selecciona los comprobantes de pago (PDF)",
            allowed_extensions=["pdf"], allow_multiple=True)
        if not archivos:
            return
        elegidos = [a.path for a in archivos if a.path]
        if not elegidos:
            return

        try:
            # En un hilo: separar y leer PDFs es síncrono y congelaría la interfaz.
            rutas, info_sep = await asyncio.to_thread(self._separar_pdfs, elegidos)
            lecturas, errores = await asyncio.to_thread(
                lector_comprobantes.leer_varios, rutas)
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.app.avisar(f"No se pudieron leer los comprobantes: {exc}", ROJO)
            return

        # Se respeta lo ya vinculado: ni se reasignan filas que ya tienen
        # comprobante ni se reparte de nuevo un archivo ya usado. Así la carga
        # se puede repetir en tandas sin deshacer lo anterior.
        por_fila = {id(f): f for f in pendientes}
        ocupados = {k for k, f in por_fila.items() if f.comprobante}
        usados = {f.comprobante for f in pendientes if f.comprobante}
        libres = [r for r in rutas if r not in usados]
        res = _comprobantes.vincular(casables, lecturas, libres, ocupados=ocupados)

        # La lectura se guarda junto al vínculo: de ahí sale la referencia.
        por_archivo, _ = _comprobantes.repartir_lecturas(lecturas, libres)
        for clave, ruta in res.asignados.items():
            leidas = por_archivo.get(ruta) or []
            por_fila[clave].vincular_comprobante(
                ruta, leidas[0] if leidas else None)
        if res.asignados:
            self._refrescar()
        self._avisar_vinculacion(res, len(libres), errores, info_sep,
                                 sin_casar=len(objetivos) - len(casables))

    @staticmethod
    def _separar_pdfs(rutas: list[str]) -> tuple[list[str], dict]:
        """Separa a disco los PDF de varias páginas (una página = un archivo).

        Devuelve `(rutas_finales, info)`. Un PDF que no se pueda separar se envía
        completo: es preferible leerlo entero que perderlo. Los de una sola página
        se usan tal cual, sin dejar copias sueltas en disco.
        """
        finales: list[str] = []
        info = {"separados": 0, "paginas": 0, "errores": []}
        for ruta in rutas:
            carpeta = os.path.join(os.path.dirname(ruta), "_paginas")
            try:
                os.makedirs(carpeta, exist_ok=True)
                paginas = pdf_paginas.separar_paginas(ruta, carpeta)
            except Exception as exc:  # noqa: BLE001 — se reporta en el resumen
                info["errores"].append(f"{os.path.basename(ruta)}: {exc}")
                finales.append(ruta)
                continue
            if len(paginas) > 1:
                info["separados"] += 1
                info["paginas"] += len(paginas)
                finales.extend(paginas)
            else:
                finales.append(ruta)
        return finales, info

    def _avisar_vinculacion(self, res, total_archivos: int, errores: list,
                            info_sep: dict, sin_casar: int) -> None:
        """Resume qué pasó: cuántos se vincularon y, sobre todo, qué quedó fuera
        y por qué. Lo que no casa se dice explícito para resolverlo a mano, en vez
        de descubrirlo hasta el momento de subirlo al SIPP."""
        partes = []
        if info_sep.get("separados"):
            partes.append(f"{info_sep['separados']} PDF separado(s) en "
                          f"{info_sep['paginas']} página(s)")
        partes.append(f"{len(res.asignados)} de {total_archivos} vinculado(s)")
        if res.sin_movimiento:
            partes.append(f"{len(res.sin_movimiento)} sin registro que coincida")
        if errores:
            partes.append(f"{len(errores)} ilegible(s)")
        if res.sin_archivo:
            partes.append(f"{res.sin_archivo} lectura(s) sin archivo")
        if sin_casar:
            partes.append(f"{sin_casar} registro(s) sin cuenta origen asignada")
        if not res.asignados:
            color = ROJO
        elif res.sin_movimiento or errores or sin_casar:
            color = NARANJA
        else:
            color = VERDE
        self.app.avisar("Comprobantes: " + "; ".join(partes) + ".", color)

    # ------------------------------------------- subida al SIPP (RPA)
    def _filas_a_subir(self) -> list:
        """Filas con comprobante vinculado, folio y que no se hayan registrado ya
        en esta corrida. Sin folio no se puede ubicar la solicitud en el SIPP."""
        return [
            f for f in self.filas
            if f.comprobante and f.folio_valor
            and id(f) not in self._sub_registradas
        ]

    async def _subir_comprobantes(self, _e=None) -> None:
        """Pide confirmación y lanza el RPA que registra las devoluciones.

        La confirmación es UNA sola, al inicio del lote: el RPA guarda cada
        devolución por su cuenta (así se acordó), y preguntar por registro haría
        inútil automatizarlo. Pero guardar escribe en el SIPP y no se deshace, así
        que conviene un punto de escape antes de empezar."""
        if self._sub_corriendo:
            self.app.avisar("Ya hay una subida en curso.", NARANJA)
            return
        filas = self._filas_a_subir()
        if not filas:
            sin_comprobante = sum(1 for f in self.filas
                                  if not f.vacia() and not f.comprobante)
            sin_folio = sum(1 for f in self.filas
                            if f.comprobante and not f.folio_valor)
            detalle = []
            if sin_comprobante:
                detalle.append(f"{sin_comprobante} sin comprobante vinculado")
            if sin_folio:
                detalle.append(f"{sin_folio} sin folio de solicitud")
            if self._sub_registradas:
                detalle.append(f"{len(self._sub_registradas)} ya registrada(s)")
            self.app.avisar(
                "No hay nada que subir" + (": " + "; ".join(detalle) if detalle
                                           else "."), NARANJA)
            return

        usuario, contrasena = self.app.config.credenciales()
        if not usuario or not contrasena:
            self.app.avisar(
                "Captura tu usuario y contraseña del SIPP en Configuración (⚙).",
                ROJO)
            return

        def continuar(_ev=None):
            self.page.pop_dialog()
            self.page.run_task(self._ejecutar_subida, filas)

        muestra = [
            f"• Solicitud {f.folio_valor} — {f.valores()[2] or '(sin nombre)'} — "
            f"ref. {lector_comprobantes.referencia_aaaammdd(f.lectura_comprobante or {}) or '?'}"
            for f in filas[:8]
        ]
        if len(filas) > 8:
            muestra.append(f"… y {len(filas) - 8} más.")
        self.page.show_dialog(ft.AlertDialog(
            modal=True,
            title=ft.Text("Subir comprobantes al SIPP"),
            content=ft.Container(
                content=ft.Column(
                    [ft.Text(
                        f"Se registrarán {len(filas)} devolución(es) en el SIPP: se "
                        "adjunta el comprobante, se escribe la referencia y se "
                        "guarda. Esto NO se puede deshacer desde la herramienta.",
                        size=13),
                     ft.Column([ft.Text(m, size=12, color=GRIS) for m in muestra],
                               tight=True, spacing=3,
                               scroll=ft.ScrollMode.AUTO if len(muestra) > 6 else None,
                               height=200 if len(muestra) > 6 else None)],
                    tight=True, spacing=10),
                width=620),
            actions=[
                ft.TextButton("Cancelar",
                              on_click=lambda e: self.page.pop_dialog()),
                ft.FilledButton("Subir y guardar", on_click=continuar),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        ))

    async def _ejecutar_subida(self, filas: list) -> None:
        """Corre el RPA sobre `filas`. Cada solicitud arranca desde el listado, así
        que un fallo en una no arrastra a las siguientes."""
        self._sub_corriendo = True
        resultados = {"guardada": 0, "ya_estaba": 0, "error": 0}
        errores: list[str] = []
        usuario, contrasena = self.app.config.credenciales()
        try:
            if necesita_navegador():
                self.app.avisar("Descargando el navegador del RPA…", VERDE)
            if self.bucle_rpa is None:
                self.bucle_rpa = BucleRpa()
            self._sub_ctrl = ControlRpa(self.bucle_rpa._loop)
            self.sesion_rpa = SesionSipp(headless=False)
            sesion, ctrl = self.sesion_rpa, self._sub_ctrl

            async def flujo() -> None:
                if necesita_navegador():
                    await asegurar_navegador()
                await sesion.iniciar()
                await sesion.login(usuario, contrasena)
                # Sesión FIJA: el listado de autorizadas no depende de la empresa
                # con la que se entró (ver EMPRESA_SESION).
                await sesion.seleccionar_empresa_sucursal(
                    self.EMPRESA_SESION, self.SUCURSAL_SESION)
                await sesion.ir_a_devoluciones()
                for i, fila in enumerate(filas, start=1):
                    await ctrl.punto_control()
                    folio = fila.folio_valor
                    lectura = fila.lectura_comprobante or {}
                    ref = lector_comprobantes.referencia_aaaammdd(lectura)
                    fecha = lector_comprobantes.fecha_aplicacion_ddmmaaaa(lectura)
                    if not ref:
                        errores.append(
                            f"Solicitud {folio}: el comprobante no trae fecha de "
                            "aplicación, así que no hay referencia que escribir.")
                        resultados["error"] += 1
                        continue
                    self._sub_avisar(
                        f"Subiendo {i}/{len(filas)}: solicitud {folio}…")
                    try:
                        estado = await sesion.registrar_devolucion(
                            folio, fila.comprobante, ref, fecha)
                    except RpaDetenido:
                        raise
                    except Exception as exc:  # noqa: BLE001 — no aborta el resto
                        errores.append(f"Solicitud {folio}: {exc}")
                        resultados["error"] += 1
                        continue
                    resultados[estado] = resultados.get(estado, 0) + 1
                    if estado in ("guardada", "ya_estaba"):
                        self._sub_registradas.add(id(fila))
                    if estado == "error":
                        errores.append(
                            f"Solicitud {folio}: se abrió pero no se pudo "
                            "guardar (hay diagnóstico en _diagnostico_rpa).")

            await asyncio.wrap_future(self.bucle_rpa.enviar(flujo()))
        except RpaDetenido:
            self.app.avisar("Subida detenida.", NARANJA)
        except ErrorSipp as exc:
            self.app.avisar(f"El SIPP no respondió como se esperaba: {exc}", ROJO)
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.app.avisar(f"Falló la subida: {exc}", ROJO)
        finally:
            self._sub_corriendo = False
            await self._cerrar_sesion_rpa()
            self._resumen_subida(resultados, errores)

    def _sub_avisar(self, texto: str) -> None:
        """Estatus desde el hilo del RPA (Flet no es thread-safe)."""
        try:
            self.app.avisar(texto, VERDE)
        except Exception:  # noqa: BLE001 — avisar nunca debe tumbar el flujo
            pass

    async def _cerrar_sesion_rpa(self) -> None:
        """Cierra el navegador del RPA. Best-effort: si falla, no debe tapar el
        resultado de la subida."""
        sesion, self.sesion_rpa = self.sesion_rpa, None
        self._sub_ctrl = None
        if sesion is None:
            return
        try:
            await asyncio.wrap_future(self.bucle_rpa.enviar(sesion.cerrar()))
        except Exception:  # noqa: BLE001 — el cierre no propaga errores
            pass

    def _resumen_subida(self, resultados: dict, errores: list) -> None:
        """Resume el lote. Se distingue lo GUARDADO de lo que YA ESTABA: esto
        segundo no es un fallo —el SIPP ya no la lista entre las autorizadas porque
        se registró antes— y confundirlo con un error haría dudar de un proceso que
        funcionó."""
        partes = []
        if resultados.get("guardada"):
            partes.append(f"{resultados['guardada']} registrada(s)")
        if resultados.get("ya_estaba"):
            partes.append(f"{resultados['ya_estaba']} ya estaban registradas")
        if resultados.get("error"):
            partes.append(f"{resultados['error']} con error")
        if not partes:
            return
        color = ROJO if resultados.get("error") else VERDE
        texto = "Subida al SIPP: " + "; ".join(partes) + "."
        if errores:
            texto += " " + errores[0]
            if len(errores) > 1:
                texto += f" (y {len(errores) - 1} más)"
        self.app.avisar(texto, color)
        self._refrescar()
    # ------------------------------------------------------- generación
    def _registros(self, filas: list[FilaSolicitud]) -> list[tuple] | None:
        """Valida las filas y las devuelve como [(clabe, monto, cliente, concepto)].
        Devuelve None (tras avisar) si algún dato impide generar el archivo."""
        registros = []
        for f in filas:
            if f.vacia():
                continue  # fila manual sin capturar: se ignora
            clabe, monto_txt, cliente, concepto, _empresa = f.valores()
            quien = cliente or f.folio_valor
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

    def _repetidos_a_dispersar(self) -> list[tuple[FilaSolicitud, ...]]:
        """Grupos de filas ASIGNADAS que se repiten entre sí (mismo CLABE, monto,
        beneficiario y empresa): son las que provocarían un pago doble en el TXT."""
        exportables = [f for f in self.filas if f.asignacion and not f.vacia()]
        grupos: dict[tuple, list[FilaSolicitud]] = {}
        for f in exportables:
            grupos.setdefault(f.clave_duplicado(), []).append(f)
        return [tuple(fs) for fs in grupos.values() if len(fs) > 1]

    async def _generar_dispersion(self, _e=None) -> None:
        """Genera UN archivo TXT por cada (empresa origen + banco), con las filas
        que se le asignaron. Si detecta pagos repetidos entre los registros a
        dispersar, pide confirmación antes de continuar."""
        grupos = self._grupos_asignados()
        if not grupos:
            self.app.avisar(
                "No hay filas con cuenta origen asignada. Selecciónalas y usa "
                "'Asignar a filas seleccionadas'.", ROJO)
            return
        # Aviso si dos o más filas a dispersar son el mismo pago (pago doble).
        repetidos = self._repetidos_a_dispersar()
        if repetidos:
            self._confirmar_dispersion_con_duplicados(repetidos)
            return
        await self._continuar_generacion(grupos)

    def _confirmar_dispersion_con_duplicados(
            self, repetidos: list[tuple[FilaSolicitud, ...]]) -> None:
        """Muestra la advertencia de pagos repetidos y deja que el usuario valide
        si continúa (los datos son idénticos: se pagaría dos veces)."""
        total = sum(len(g) for g in repetidos)
        detalle = []
        for g in repetidos[:8]:
            clabe, monto_txt, cliente, _c, empresa = g[0].valores()
            detalle.append(
                f"• {cliente or '(sin nombre)'} — {empresa or 's/empresa'} — "
                f"${monto_txt or '0'} — {len(g)} veces")
        if len(repetidos) > 8:
            detalle.append(f"… y {len(repetidos) - 8} caso(s) más.")

        async def continuar(_ev=None):
            self.page.pop_dialog()
            await self._continuar_generacion(self._grupos_asignados())

        # El cuerpo del diálogo se dimensiona al contenido (no a la ventana): con
        # pocos casos crece natural (sin scroll); solo si son muchos acota el alto
        # y hace scroll interno. Ancho amplio para que cada caso entre en una línea.
        muchos = len(detalle) > 6
        lista = ft.Column(
            [ft.Text(t, size=12, color=GRIS) for t in detalle],
            tight=True, spacing=4,
            scroll=ft.ScrollMode.AUTO if muchos else None,
            height=220 if muchos else None,
        )
        self.page.show_dialog(
            ft.AlertDialog(
                modal=True,
                title=ft.Text("⚠ Hay pagos repetidos"),
                content=ft.Container(
                    content=ft.Column(
                        [ft.Text(
                            f"{total} de los registros a dispersar parecen el mismo "
                            "pago (mismo CLABE, monto, beneficiario y empresa). Si "
                            "continúas, se pagarían dos o más veces:", size=13),
                         lista],
                        tight=True, spacing=10),
                    width=640,
                ),
                actions=[
                    ft.TextButton("Cancelar y revisar",
                                  on_click=lambda e: self.page.pop_dialog()),
                    ft.FilledButton("Dispersar de todos modos", on_click=continuar,
                                    color=ft.Colors.WHITE, bgcolor=ROJO),
                ],
                actions_alignment=ft.MainAxisAlignment.END,
            )
        )

    async def _continuar_generacion(self, grupos: dict) -> None:
        """Valida y escribe los TXT (un archivo por empresa origen + banco)."""
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
        # Vía el shell de la app: es lo que hace que el Explorador salga POR ENCIMA
        # de la app (ver AppTesoreria.abrir_en_sistema).
        self.app.abrir_en_sistema(carpeta)
        self.app.avisar(
            f"{len(generados)} archivo(s) TXT generados: " + ", ".join(generados), VERDE)

    def _filas_reporte(self) -> list[dict] | None:
        """Arma los registros del reporte Excel: valida las filas y le agrega a
        cada una SU PROPIA cuenta origen asignada (empresa/banco/cuenta/número/
        fecha). Devuelve None si algún dato impide exportar (el aviso lo da
        `_registros`)."""
        registros = self._registros(self.filas)
        if registros is None:
            return None
        # `_registros` conserva el orden y omite las filas vacías; se emparejan con
        # las mismas filas no vacías para tomar su asignación individual.
        filas_validas = [f for f in self.filas if not f.vacia()]
        datos = []
        for f, (clabe, monto, cliente, concepto) in zip(filas_validas, registros):
            a = f.asignacion or {}
            datos.append({
                "empresa": a.get("empresa_origen", ""),
                "banco": a.get("banco", ""),
                "cuenta_origen": a.get("cuenta_origen", ""),
                "num_cuenta": a.get("num_cuenta", ""),
                # Fecha de la asignación de la fila; si no está asignada, la del panel.
                "fecha": a.get("fecha", "") or self.tf_fecha.value or "",
                "clabe": clabe, "monto": monto,
                "beneficiario": cliente, "concepto": concepto,
            })
        return datos

    async def _generar_excel(self, _e=None) -> None:
        """Reporte Excel con TODOS los registros de la tabla (SIPP + manuales),
        cada uno con su propia cuenta origen asignada."""
        datos = self._filas_reporte()
        if datos is None:
            return
        if not datos:
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
            reporte_excel.generar(ruta, datos)
        except PermissionError:
            self.app.avisar(
                "No se pudo guardar: el archivo está abierto en Excel. Ciérralo e "
                "intenta de nuevo (o guarda con otro nombre).", ROJO)
            return
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.app.avisar(f"No se pudo generar el Excel: {exc}", ROJO)
            return
        self.app.avisar(
            f"Reporte Excel generado con {len(datos)} movimiento(s).", VERDE)
