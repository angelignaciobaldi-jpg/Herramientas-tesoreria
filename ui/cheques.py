"""Pantalla "Cheques": operaciones con cheques.

En esta primera vuelta hay una sola operación —la impresión de cheques— y llega
hasta el modal de confirmación: el usuario captura beneficiario, monto, fecha,
moneda y layout de banco; la app arma los datos de impresión (incluido el MONTO
EN LETRA, derivado del número) y los muestra para confirmar. La impresión real se
cableará después (ver core/impresion.py + ui/dialogo_impresion.py).

El formulario captura un cheque, pero la capa de datos (core/cheques.py) trabaja
con listas, de modo que soportar varios cheques a la vez sea un cambio menor.
"""

from __future__ import annotations

import datetime

import flet as ft

from core import cheques as cheques_core
from core import impresion, preferencias
from ui.comun import GRIS, NARANJA, ROJO, VERDE, parse_monto

# Clave de preferencias donde se recuerda la impresora elegida para cheques.
_CLAVE_IMPRESORA = "impresora_cheques"

# Importe máximo capturable: numero_letras solo convierte hasta 999,999,999 (más
# allá lanzaría error). Se limita la escritura para no llegar a esa pantalla.
_MONTO_MAX = 999_999_999

def _fmt_fecha(d: datetime.date) -> str:
    return d.strftime("%d/%m/%Y")


def _label_requerido(texto: str) -> ft.Text:
    """Etiqueta de campo requerido: el texto seguido de un asterisco ROJO."""
    return ft.Text(
        spans=[
            ft.TextSpan(texto + " "),
            ft.TextSpan("*", ft.TextStyle(color=ROJO, weight=ft.FontWeight.BOLD)),
        ],
        size=12,
    )


class SeccionCheques:
    """Sección de operaciones con cheques (impresión de cheques)."""

    def __init__(self, app):
        self.app = app
        self.page = app.page
        self.dd_impresora: ft.Dropdown | None = None
        self._monto_ok = ""   # último valor aceptado del importe (para revertir)
        self.contenido = self._construir()

    # ------------------------------------------------------------ UI
    def _construir(self) -> ft.Control:
        hoy = datetime.date.today()

        # Sin overrides de altura/padding: todos los campos usan la altura por
        # defecto de Material (así quedan parejos, igual que en Dispersión). El
        # grid se arma con ResponsiveRow + 'col' (sm=12 apila en pantallas chicas;
        # md marca el ancho tipo Bootstrap: col-md-6 -> 6, col-md-3 -> 3).
        self.tf_beneficiario = ft.TextField(label=_label_requerido("Beneficiario"))
        self.tf_monto = ft.TextField(
            label=_label_requerido("Monto"), prefix_icon=ft.Icons.ATTACH_MONEY,
            hint_text="0.00", on_change=self._limitar_monto)
        self.dd_moneda = ft.Dropdown(
            label=_label_requerido("Moneda"), value="MXN", expand=True,
            options=[
                ft.dropdown.Option(key="MXN", text="MXN (pesos)"),
                ft.dropdown.Option(key="USD", text="USD (dólares)"),
            ],
        )
        # Fecha: calendario de solo lectura (igual que en Dispersión). Default hoy.
        self.dp_fecha = ft.DatePicker(
            value=hoy,
            first_date=datetime.date(2020, 1, 1),
            last_date=datetime.date(2035, 12, 31),
            help_text="Fecha del cheque",
            on_change=self._fecha_elegida,
        )
        self.tf_fecha = ft.TextField(
            label=_label_requerido("Fecha"), value=_fmt_fecha(hoy), read_only=True,
            suffix_icon=ft.Icons.CALENDAR_MONTH,
            on_click=lambda _e: self.page.show_dialog(self.dp_fecha),
        )
        self.dd_banco = ft.Dropdown(
            label=_label_requerido("Banco"), value=cheques_core.LAYOUTS_BANCO[0],
            expand=True,
            options=[ft.dropdown.Option(b) for b in cheques_core.LAYOUTS_BANCO],
        )

        # Encabezado: título + ícono de ayuda con el detalle de para qué sirve.
        encabezado = ft.Row(
            [
                ft.Text("Impresión manual", weight=ft.FontWeight.BOLD, size=15),
                ft.Icon(
                    ft.Icons.HELP_OUTLINE, size=18, color=GRIS,
                    # Tooltip inmediato al pasar el mouse (wait_duration=0).
                    tooltip=ft.Tooltip(
                        message=(
                            "En caso de querer imprimir un ticket con información "
                            "que no se encuentre de momento en el sistema lo puede "
                            "hacer llenando estos campos."),
                        wait_duration=ft.Duration(milliseconds=0),
                    ),
                ),
            ],
            spacing=8, vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )

        def col(control, ancho_md):
            control.col = {"sm": 12, "md": ancho_md}
            return control

        # Fila superior completa (6+3+3 = 12); la inferior deja libres las últimas
        # 6 columnas.
        fila_sup = ft.ResponsiveRow(
            [col(self.tf_beneficiario, 6), col(self.tf_fecha, 3),
             col(self.dd_banco, 3)],
            spacing=16, run_spacing=16,
            vertical_alignment=ft.CrossAxisAlignment.START,
        )
        # El botón va en la misma fila que Monto/Moneda, ocupando las 6 columnas
        # libres y fijado a la derecha (alineación derecha dentro de su columna).
        boton_generar = ft.Container(
            content=ft.FilledButton(
                "Generar cheque", icon=ft.Icons.PRINT, on_click=self._generar),
            alignment=ft.Alignment(1, 0),
        )
        fila_inf = ft.ResponsiveRow(
            [col(self.tf_monto, 3), col(self.dd_moneda, 3),
             col(boton_generar, 6)],
            spacing=16, run_spacing=16,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )

        formulario = ft.Card(
            content=ft.Container(
                content=ft.Column(
                    [encabezado, fila_sup, fila_inf],
                    spacing=14, tight=True,
                ),
                padding=16,
            ),
        )
        # STRETCH hace que la tarjeta abarque todo el ancho válido de la pantalla.
        return ft.Column(
            [formulario], spacing=14, scroll=ft.ScrollMode.AUTO, expand=True,
            horizontal_alignment=ft.CrossAxisAlignment.STRETCH)

    def _fecha_elegida(self, _e=None) -> None:
        """Vuelca la fecha elegida en el calendario al campo, como DD/MM/AAAA."""
        if self.dp_fecha.value:
            self.tf_fecha.value = _fmt_fecha(self.dp_fecha.value)
            self.page.update()

    def _limitar_monto(self, _e=None) -> None:
        """Limita el importe: máximo 2 decimales y parte entera <= _MONTO_MAX (tope
        de numero_letras). Recorta los decimales de más y revierte al último valor
        aceptado si se pasa del tope. No estorba textos parciales (los valida
        _generar)."""
        texto = self.tf_monto.value or ""
        # Máximo 2 decimales: recorta lo que sobre tras el punto.
        if "." in texto:
            entero, _, dec = texto.partition(".")
            if len(dec) > 2:
                texto = f"{entero}.{dec[:2]}"
        # Tope del importe (parte entera).
        try:
            valor = parse_monto(texto)
        except ValueError:
            valor = None  # texto no numérico (parcial): se deja, _generar valida
        if valor is not None and int(valor) > _MONTO_MAX:
            aceptado = self._monto_ok       # se pasó del tope: revierte
        else:
            aceptado = texto
            self._monto_ok = texto          # último valor aceptado
        if aceptado != (self.tf_monto.value or ""):
            self.tf_monto.value = aceptado
            self.tf_monto.update()

    # -------------------------------------------------------- acciones
    def _generar(self, _e=None) -> None:
        """Valida el formulario, arma los datos de impresión y abre el modal de
        confirmación. (La impresión real se cableará en una vuelta posterior.)"""
        beneficiario = (self.tf_beneficiario.value or "").strip()
        if not beneficiario:
            self.app.avisar("Captura el nombre del beneficiario.", NARANJA)
            return
        try:
            monto = parse_monto(self.tf_monto.value)
        except ValueError:
            self.app.avisar("El monto no es un número válido.", ROJO)
            return
        if not monto or monto <= 0:
            self.app.avisar("Captura un monto mayor a cero.", NARANJA)
            return
        if int(round(monto * 100)) // 100 > _MONTO_MAX:  # red de seguridad
            self.app.avisar(
                f"El importe no puede exceder {_MONTO_MAX:,}.", NARANJA)
            return
        fecha = self.dp_fecha.value or datetime.date.today()
        moneda = self.dd_moneda.value or "MXN"
        banco = self.dd_banco.value or ""

        # Lista de un cheque (la capa de datos ya soporta varios).
        lote = [cheques_core.Cheque(
            beneficiario=beneficiario, monto=monto, fecha=fecha, moneda=moneda)]
        datos = cheques_core.preparar_cheques(lote, banco)
        self._mostrar_confirmacion(datos)

    def _mostrar_confirmacion(
            self, datos: list[cheques_core.DatosImpresionCheque]) -> None:
        """Modal con los datos de cada cheque (itera la lista, aunque hoy traiga
        uno) + el selector de impresora integrado. El botón Imprimir lanza la
        PRUEBA de calibración con la impresora elegida aquí mismo."""
        bloques = [self._bloque_cheque(d) for d in datos]

        # Selector de impresora dentro del propio modal (sin abrir otro diálogo).
        impresoras = impresion.listar_impresoras()
        if impresoras:
            self.dd_impresora = ft.Dropdown(
                label="Impresora", width=488,
                value=self._impresora_preseleccionada(impresoras),
                options=[ft.dropdown.Option(n) for n in impresoras],
            )
            seccion_impresora: ft.Control = self.dd_impresora
        else:
            self.dd_impresora = None
            seccion_impresora = ft.Text(
                "No se encontraron impresoras instaladas.", color=ROJO, size=12)

        # tight=True para que el contenido abrace al cuerpo y no deje espacio en
        # blanco de más; scroll solo si hubiera muchos cheques.
        cuerpo = ft.Column(
            [*bloques, ft.Divider(), seccion_impresora],
            spacing=12, tight=True,
            scroll=ft.ScrollMode.AUTO if len(bloques) > 3 else None)

        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Text("Información del cheque", weight=ft.FontWeight.BOLD),
            content=ft.Container(content=cuerpo, width=520),
            actions=[
                ft.TextButton("Cerrar", on_click=lambda _e: self.page.pop_dialog()),
                ft.FilledButton(
                    "Imprimir", icon=ft.Icons.PRINT,
                    disabled=not impresoras,
                    on_click=lambda _e: self._imprimir_prueba(datos)),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        self.page.show_dialog(dlg)

    def _impresora_preseleccionada(self, impresoras: list[str]) -> str | None:
        """Impresora a preseleccionar: la recordada, luego la predeterminada del
        sistema y, si no, la primera de la lista."""
        elegida = preferencias.cargar_valor(_CLAVE_IMPRESORA)
        if elegida not in impresoras:
            elegida = impresion.impresora_predeterminada()
        if elegida not in impresoras:
            elegida = impresoras[0] if impresoras else None
        return elegida

    def _imprimir_prueba(
            self, datos: list[cheques_core.DatosImpresionCheque]) -> None:
        """Imprime el cheque (texto negro, vertical) con la impresora elegida en el
        modal. En esta vuelta se imprime el primer cheque del lote. (Si en el motor
        se activa el modo calibración, sale además la cuadrícula y los recuadros.)"""
        if not self.dd_impresora or not self.dd_impresora.value:
            self.app.avisar("Selecciona una impresora.", NARANJA)
            return
        d = datos[0]
        coords = cheques_core.layout_banco(d.banco) or {}
        if not coords:
            self.app.avisar(
                f"El banco «{d.banco}» no tiene un layout de impresión definido.",
                NARANJA)
            return
        campos = {
            "beneficiario": d.beneficiario,
            "fecha": d.fecha_texto,
            # El número va SIN el signo '$' (el cheque ya trae su símbolo impreso).
            "monto_numero": d.monto_texto.replace("$", "").strip(),
            "monto_letras": d.monto_letras,
        }
        nombre = self.dd_impresora.value
        preferencias.guardar_valor(_CLAVE_IMPRESORA, nombre)  # recuerda la elegida
        self.page.pop_dialog()
        try:
            impresion.imprimir_prueba_cheque(nombre, campos, coords)
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.app.avisar(f"No se pudo imprimir: {exc}", ROJO)
            return
        self.app.avisar(f"Enviado a imprimir en «{nombre}».", VERDE)

    def _bloque_cheque(
            self, d: cheques_core.DatosImpresionCheque) -> ft.Control:
        """Tarjeta con los datos de un cheque: cada campo en dos líneas (etiqueta
        en negrita pequeña arriba, valor abajo), en cuadrícula de dos columnas."""
        return ft.Container(
            content=ft.Column(
                [
                    ft.Row(
                        [self._campo("Beneficiario:", d.beneficiario),
                         self._campo("Monto:", f"{d.monto_texto} {d.moneda}")],
                        spacing=16,
                        vertical_alignment=ft.CrossAxisAlignment.START),
                    ft.Row(
                        [self._campo("Fecha:", d.fecha_texto),
                         self._campo("Layout del banco:", d.banco)],
                        spacing=16,
                        vertical_alignment=ft.CrossAxisAlignment.START),
                    self._campo("Cantidad con letra:", d.monto_letras),
                ],
                spacing=14, tight=True,
            ),
            padding=16,
            bgcolor=ft.Colors.SURFACE_CONTAINER_HIGHEST,
            border_radius=10,
            border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT),
        )

    @staticmethod
    def _campo(etiqueta: str, valor: str) -> ft.Control:
        """Campo en dos líneas: etiqueta (negrita, pequeña) sobre el valor. Ocupa
        el ancho disponible para acomodarse en cuadrícula."""
        return ft.Column(
            [
                ft.Text(etiqueta, size=11, weight=ft.FontWeight.BOLD, color=GRIS),
                ft.Text(valor, size=14, selectable=True),
            ],
            spacing=2, tight=True, expand=True,
        )
