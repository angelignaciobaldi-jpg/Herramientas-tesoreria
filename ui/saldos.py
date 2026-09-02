"""Pantalla: Saldos.

Sustituye el trabajo manual de pegar, hoja por hoja, los reportes de cada portal
bancario en `FORMATO DE SALDOS TESORERIA.xlsx`. El usuario sube los archivos tal
como los descarga —en cualquier orden y con cualquier nombre— y la herramienta
reconstruye el libro completo: las mismas pestañas, las mismas fórmulas, con cada
cuenta puesta en su fila.

El flujo tiene tres pasos:

  1. **Cargar archivos** — se leen en un hilo (ver `_leer_en_hilo`) y se muestra
     el avance archivo por archivo.
  2. **Revisar la cobertura** — cuántos renglones del formato quedaron llenos.
  3. **Generar** — pide dónde guardarlo y lo abre.

## Una sola pila

Entran dos clases de archivo: los reportes de los portales bancarios (lo único
obligatorio) y los insumos de flujo —créditos, Pemex, MGC, tesoro, nómina— que
alimentan las columnas de crédito y el calendario de proyección. Se suben por el
MISMO botón, aparecen en la MISMA lista y la herramienta decide cuál es cuál por
el contenido, igual que ya hacía con el banco. Lo único que los distingue en
pantalla es el marbete de su fila.

Pedirle al usuario que separe las dos pilas —o mostrárselas separadas— sería
devolverle justo el trabajo manual que este módulo vino a quitar. Y se puede
hacer sin riesgo porque los dos detectores no se pisan: medido sobre los archivos
reales, ninguno de los 24 reportes bancarios se confunde con un insumo ni al
revés. Aun así el orden importa —primero banco, después insumo—, porque los
reportes bancarios son el caso de todos los días y los insumos la excepción.

## Sobre el diseño de la pantalla

La cifra que manda es **cuántos de los 209 renglones del formato quedaron
llenos**, y por eso es lo único grande de la pantalla. Con «cuántas cuentas
identifiqué» uno cree que todo salió bien; con «201 de 209» se ve el hueco. La
barra de cobertura cambia de color sola: verde arriba de 95 %, ámbar entre 80 y
95, rojo abajo — para que no haya que leer el número para saber si algo falta.

Las dos acciones viven en el mismo renglón y en extremos opuestos: cargar a la
izquierda (donde empieza la lectura) y generar a la derecha, en verde, como
último paso. El verde solo aplica cuando el botón está habilitado: un botón
llamativo que no hace nada es una mentira visual.

La lista de archivos crece libre, sin tope de altura. Recortarla producía dos
barras de desplazamiento anidadas —la de la pantalla y la suya— y eso es peor que
una lista larga: el usuario no sabe cuál está moviendo. Los archivos que fallaron
se ordenan al final, porque intercalados se pierden entre veintitantos renglones
verdes y son justo los que hay que atender.

Nota de rendimiento: aquí NO se renderiza una tabla de cuentas. Con 20 archivos y
más de 200 saldos, pintar una tabla completa costaría cientos de widgets y el
render se notaría; la lista de archivos son unas decenas de filas y el detalle
vive en el Excel. Por lo mismo las «etiquetas» de banco son `Container`, no
`Chip`: un Chip es un control interactivo y aquí solo hace falta pintar texto.
"""

from __future__ import annotations

import asyncio
import datetime
import os
import traceback

import flet as ft

from core import rutas
from core import saldos as motor
from core import (saldos_estado, saldos_export, saldos_insumos,
                  saldos_lectores, saldos_plantilla)
from ui.comun import CENTRO, GRIS, NARANJA, ROJO, ROJO_BOTON, VERDE, tarjeta

# Todo lo que puede entrar: los portales emiten estos formatos (ver
# core/saldos_lectores) y los insumos de flujo siempre son hojas de cálculo.
_EXTENSIONES = ["xlsx", "xls", "csv", "txt", "pdf"]

# Cómo se llama cada insumo en pantalla.
_NOMBRES_INSUMO = {
    "CREDITOS": "Créditos",
    "PEMEX": "Pemex",
    "MGC": "MGC",
    "TESORO": "Tesoro",
    "NOMINA": "Nómina",
}

# Instrucciones de la pantalla. Viven en el tooltip del ícono de ayuda que va
# junto al título, no en un párrafo del cuerpo: solo hacen falta la primera vez.
_AYUDA = """Carga los archivos que se usarán para realizar el reporte de saldos.

Los formatos aceptados son los siguientes:
    - Excel (xlsx,xls)
    - CSV
    - Texto (txt)
    - PDF

La relación de facturas de Pemex se toma directo de la información del SIPP."""

# Umbrales de la barra de cobertura. Debajo del 80 % casi seguro falta un archivo
# entero, no una cuenta suelta.
_COBERTURA_BUENA = 0.95
_COBERTURA_REGULAR = 0.80


def _estilo_verde() -> ft.ButtonStyle:
    """Botón verde que se apaga al deshabilitarse.

    Si el color se pusiera plano, Material lo conservaría también en el estado
    deshabilitado y el botón seguiría gritando «púlsame» cuando todavía no hay
    nada que generar."""
    return ft.ButtonStyle(
        bgcolor={
            ft.ControlState.DEFAULT: VERDE,
            ft.ControlState.DISABLED: ft.Colors.with_opacity(
                0.12, ft.Colors.ON_SURFACE),
        },
        color={
            ft.ControlState.DEFAULT: ft.Colors.WHITE,
            ft.ControlState.DISABLED: ft.Colors.with_opacity(
                0.38, ft.Colors.ON_SURFACE),
        },
    )


def _filas_de(datos) -> int:
    """Cuántas filas trae una sección, sea ledger o el bloque de créditos."""
    if isinstance(datos, list):
        return len(datos)
    if isinstance(datos, dict) and "rangos" in datos:
        return sum(len(r["celdas"]) for r in datos["rangos"])
    return 0


def _etiqueta(texto: str, color) -> ft.Container:
    """Marbete de texto sobre fondo tenue (banco, tipo de insumo)."""
    return ft.Container(
        content=ft.Text(texto, size=11, color=color,
                        weight=ft.FontWeight.W_500),
        bgcolor=ft.Colors.with_opacity(0.10, color),
        padding=ft.Padding.symmetric(vertical=2, horizontal=8),
        border_radius=20)


class SeccionSaldos:
    """Pantalla de generación del reporte diario de saldos bancarios."""

    def __init__(self, app):
        self.app = app
        self.page = app.page
        # Archivos bancarios: [{"ruta", "banco", "cuentas", "error", "lineas"}]
        self.archivos: list[dict] = []
        # Insumos de flujo: [{"ruta", "tipo", "filas", "error", "datos"}]
        self.insumos: list[dict] = []
        self.lineas: list[saldos_lectores.LineaSaldo] = []
        self.asignacion: motor.Asignacion | None = None
        # Marca si hay una lectura en curso. `_progreso` corre en el hilo de
        # lectura y marshala con `page.run_task`, así que su último aviso puede
        # llegar DESPUÉS de que `_ocupado(False)` limpió la barra y dejarla
        # diciendo «Leyendo archivos… 5 de 5» para siempre.
        self._leyendo = False
        self._dlg_espera: ft.AlertDialog | None = None
        self._pestana = 0
        # Lo que quedó de días anteriores. Los ledgers de flujo se capturan por
        # semana, no a diario, así que se conservan y solo se tocan cuando el
        # usuario suba algo nuevo o los borre.
        #
        # NO se leen aquí: el libro de insumos ronda las 32 000 filas y tarda
        # varios segundos en abrirse, y el constructor corre en el arranque de
        # la app —con la pantalla de «Iniciando aplicación» al frente—. Quien no
        # venga a saldos pagaría esa espera para nada. Se difiere a `al_entrar`,
        # que el shell dispara la primera vez que se abre esta pantalla.
        self.guardados: dict = {}
        self._estado_cargado = False
        self._cargando_estado = False
        self._estado_listo = asyncio.Event()
        # La plantilla se carga una vez y se cachea en el módulo. Si falta el
        # artefacto derivado, la pantalla lo dice en vez de reventar al generar.
        try:
            self.plantilla = saldos_plantilla.cargar()
            self.error_plantilla = ""
        except saldos_plantilla.ErrorPlantilla as exc:
            self.plantilla = None
            self.error_plantilla = str(exc)
        self.contenido = self._construir()

    # ------------------------------------------------ carga diferida
    def al_entrar(self) -> None:
        """Gancho del shell: se dispara al abrir esta pantalla.

        Aquí se paga —una sola vez, y ya con la app en pantalla— la lectura de
        los insumos guardados que antes se hacía en el arranque."""
        if self._estado_cargado or self._cargando_estado:
            return
        self.page.run_task(self._cargar_estado)

    async def _cargar_estado(self) -> None:
        """Lee los insumos persistidos en un hilo y repinta al terminar."""
        if self._estado_cargado or self._cargando_estado:
            return
        self._cargando_estado = True
        self.cargando_insumos.visible = True
        self._refrescar(self.cargando_insumos)
        try:
            self.guardados = await asyncio.to_thread(
                saldos_estado.cargar_insumos)
        except BaseException as exc:  # noqa: BLE001 — nunca debe tumbar la pantalla
            self._registrar_error(exc, "(lectura de insumos guardados)")
            self.guardados = {}
        finally:
            self._estado_cargado = True
            self._cargando_estado = False
            self._estado_listo.set()
        self.cargando_insumos.visible = False
        self._pintar()

    async def _asegurar_estado(self) -> None:
        """Espera a tener los insumos guardados antes de tocarlos.

        Sin esto, fusionar o exportar mientras la lectura sigue en curso
        trabajaría sobre un diccionario vacío: se guardaría encima y se
        perderían las secciones capturadas en días anteriores."""
        if self._estado_cargado:
            return
        if self._cargando_estado:
            await self._estado_listo.wait()
        else:
            await self._cargar_estado()

    # ------------------------------------------------------------------ UI
    def _construir(self) -> ft.Control:
        self.btn_cargar = ft.FilledTonalButton(
            content="Cargar archivos", icon=ft.Icons.UPLOAD_FILE,
            on_click=self._cargar)
        self.btn_limpiar = ft.TextButton(
            content="Quitar todo", icon=ft.Icons.DELETE_SWEEP_OUTLINED,
            visible=False, on_click=self._limpiar,
            style=ft.ButtonStyle(color=ROJO_BOTON))
        self.btn_generar = ft.FilledButton(
            content="Generar reporte", icon=ft.Icons.TABLE_VIEW,
            disabled=True, on_click=self._generar, style=_estilo_verde())
        self.btn_formato = ft.TextButton(
            content="Descargar formato de insumos",
            icon=ft.Icons.DOWNLOAD_OUTLINED, on_click=self._descargar_formato,
            tooltip="Un Excel con una pestaña por sección —créditos, Pemex, MGC, "
                    "tesoro, nómina e impuestos— con lo que ya está capturado. "
                    "Se llena y se vuelve a subir junto con los reportes.")
        self.txt_estado = ft.Text("", size=12, color=GRIS)
        self.anillo = ft.ProgressRing(width=16, height=16, stroke_width=2,
                                      visible=False)
        self.barra = ft.ProgressBar(value=0, visible=False, bar_height=4,
                                    border_radius=2)

        # Las dos acciones en extremos opuestos del mismo renglón: el grupo de la
        # izquierda se expande y empuja el botón verde al borde derecho.
        acciones = ft.Row(
            [
                ft.Row([self.btn_cargar, self.btn_formato, self.btn_limpiar,
                        self.anillo, self.txt_estado],
                       spacing=8, expand=True, wrap=True,
                       vertical_alignment=ft.CrossAxisAlignment.CENTER),
                self.btn_generar,
            ],
            spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER)

        self.hero = ft.Container(visible=False)
        self.avisos = ft.Column(spacing=6, tight=True, visible=False)

        # Las dos vistas del resultado. Solo una está visible a la vez; la barra
        # de pestañas de abajo decide cuál.
        self.bancos = ft.Column(spacing=0, tight=True, visible=False)
        self.tab_saldos = ft.Tab(label="Saldos identificados")
        self.tab_archivos = ft.Tab(label="Archivos cargados")
        self.tab_insumos = ft.Tab(label="Insumos de flujo")
        self.secciones = ft.Column(spacing=0, tight=True, visible=False)
        # Aviso mientras se leen los insumos guardados al entrar. Solo se ve la
        # primera vez y unos segundos, pero sin él la pestaña de insumos
        # aparecería vacía y parecería que se perdió lo capturado.
        self.cargando_insumos = ft.Row(
            [ft.ProgressRing(width=14, height=14, stroke_width=2),
             ft.Text("Recuperando los insumos guardados…", size=12,
                     color=GRIS)],
            spacing=8, visible=False)
        # Se usa la TabBar SOLA, sin TabBarView. El TabBarView del ejemplo de
        # Flet exige `expand=True`, y expandir dentro de una columna con scroll
        # obliga a fijarle una altura: volverían las dos barras de desplazamiento
        # anidadas. Con el cuerpo fuera del control, el contenido fluye con la
        # página y solo hay una.
        self.pestanas = ft.Tabs(
            length=3, visible=False, on_change=self._cambiar_pestana,
            content=ft.TabBar(tabs=[self.tab_saldos, self.tab_archivos,
                                    self.tab_insumos]))
        # Lista de archivos cargados. Es un Column simple, no una TablaResponsiva:
        # son pocas filas y no necesita columnas proporcionales ni scroll propio.
        self.lista = ft.Column(spacing=0, tight=True, visible=False)
        self.vacio = self._zona_vacia()

        panel = ft.Column(
            [
                acciones,
                self.barra,
                self.hero,
                self.avisos,
                self.pestanas,
                self.cargando_insumos,
                self.bancos,
                self.lista,
                self.secciones,
                self.vacio,
            ],
            spacing=12, horizontal_alignment=ft.CrossAxisAlignment.STRETCH)

        controles = [tarjeta("Reporte de saldos bancarios", panel,
                             ayuda=_AYUDA)]
        if self.error_plantilla:
            controles.insert(0, self._aviso_plantilla())
        return ft.Column(controles, spacing=14, scroll=ft.ScrollMode.AUTO,
                         expand=True)

    def _zona_vacia(self) -> ft.Control:
        """Estado inicial.

        A propósito NO se parece a una zona de arrastre. Flet 0.85 no expone
        ningún evento para soltar archivos del sistema operativo —`DragTarget` y
        `Draggable` solo mueven controles dentro de la app—, así que un marco
        punteado con una nube prometería algo que no funciona: el usuario
        arrastraría, no pasaría nada, y pensaría que la app está rota.

        Tampoco lleva botón propio: «Cargar archivos» está a unos pixeles, justo
        arriba, y repetir la misma acción a dos centímetros solo obliga a decidir
        cuál de los dos usar. Este bloque únicamente informa que todavía no hay
        nada; la acción vive en la barra, que es donde seguirá estando cuando ya
        haya archivos cargados.

        Si algún día se compila una extensión de Flutter para recibir archivos
        soltados (`desktop_drop`), aquí es donde iría."""
        return ft.Container(
            content=ft.Column(
                [ft.Icon(ft.Icons.FOLDER_OPEN_OUTLINED, size=32, color=GRIS),
                 ft.Text("Todavía no has cargado ningún archivo",
                         size=13, color=GRIS)],
                spacing=8, tight=True,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            alignment=CENTRO, padding=ft.Padding.symmetric(vertical=24),
            border_radius=12,
            bgcolor=ft.Colors.SURFACE_CONTAINER_HIGHEST)

    def _aviso_plantilla(self) -> ft.Control:
        """Sin la plantilla derivada no hay reporte posible. Se dice de entrada."""
        return ft.Container(
            content=ft.Row(
                [ft.Icon(ft.Icons.ERROR_OUTLINE, size=20, color=ROJO),
                 ft.Text("No se pudo cargar la plantilla del formato: "
                         f"{self.error_plantilla}",
                         size=12, color=ROJO, expand=True)],
                spacing=10, vertical_alignment=ft.CrossAxisAlignment.START),
            padding=12, border_radius=10,
            bgcolor=ft.Colors.with_opacity(0.08, ROJO))

    # ------------------------------------------------------- carga de archivos
    async def _cargar(self, _e=None) -> None:
        archivos = await self.app.picker.pick_files(
            dialog_title="Selecciona los reportes de saldos y los insumos",
            allowed_extensions=_EXTENSIONES, allow_multiple=True)
        if not archivos:
            return
        rutas = [a.path for a in archivos]
        # Un archivo que ya está cargado no se vuelve a leer: repetirlo solo
        # produciría cuentas duplicadas que después hay que descartar.
        ya = {a["ruta"] for a in self.archivos} | {x["ruta"] for x in self.insumos}
        nuevas = [r for r in rutas if r not in ya]
        repetidos = len(rutas) - len(nuevas)
        if not nuevas:
            self.app.avisar("Esos archivos ya estaban cargados.", NARANJA)
            return

        self._ocupado(True, f"Leyendo {len(nuevas)} archivo(s)…")
        try:
            leidos = await asyncio.to_thread(self._leer_en_hilo, nuevas)
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self._ocupado(False)
            self.app.avisar(f"No se pudieron leer los archivos: {exc}", ROJO)
            return

        bancarios = [x for x in leidos if x["clase"] != "insumo"]
        insumos = [x for x in leidos if x["clase"] == "insumo"]
        self.archivos.extend(bancarios)
        # Dos archivos del mismo insumo: gana el último, que es lo que espera
        # quien vuelve a subir uno para corregir el anterior.
        for nuevo in insumos:
            self.insumos = [x for x in self.insumos if x["tipo"] != nuevo["tipo"]]
            self.insumos.append(nuevo)
        if insumos:
            # Antes de fundir hay que tener lo de días anteriores: si la lectura
            # diferida siguiera en curso, se guardaría encima de un diccionario
            # vacío y se perderían las secciones ya capturadas.
            await self._asegurar_estado()
            await asyncio.to_thread(self._guardar_insumos, insumos)

        self._reidentificar()
        self._ocupado(False)
        self._pintar()
        self.app.avisar(*self._resumen_carga(bancarios, insumos, repetidos))

    @staticmethod
    def _resumen_carga(bancarios, insumos, repetidos) -> tuple:
        """Mensaje del snackbar: qué entró y de qué tipo.

        Se dice por separado cuántos reportes y cuántos insumos porque el usuario
        subió una sola pila y no tiene por qué saber en cuál cayó cada archivo:
        este aviso es su confirmación de que la herramienta los separó bien."""
        con_error = [x for x in bancarios + insumos if x["error"]]
        partes = []
        buenos = len([x for x in bancarios if not x["error"]])
        if buenos:
            partes.append(f"{buenos} reporte(s)")
        if insumos:
            partes.append(f"{len(insumos)} insumo(s)")
        mensaje = ("Se leyó " + " y ".join(partes) + "." if partes
                   else "No se pudo leer ningún archivo.")
        if repetidos:
            mensaje += f" {repetidos} ya estaban cargados y se omitieron."
        if con_error:
            mensaje += f" {len(con_error)} no se pudieron leer."
        return mensaje, (NARANJA if con_error or not partes else VERDE)

    def _leer_en_hilo(self, rutas: list[str]) -> list[dict]:
        """Lee cada archivo y decide qué es. Corre en un hilo (lo llama `_cargar`
        con `asyncio.to_thread`) porque abre archivos y parsea PDF: en el hilo de
        la interfaz congelaría la ventana varios segundos.

        No se usa `saldos_lectores.leer_varios` porque devuelve las líneas de todo
        el lote junto y aquí hace falta el detalle POR ARCHIVO —qué se detectó y
        cuánto trajo— que es lo que se pinta en la lista.

        Un archivo que falle no aborta el resto — se marca y se sigue, que es lo
        que hace falta cuando se cargan doce de golpe."""
        out: list[dict] = []
        total = len(rutas)
        for i, ruta in enumerate(rutas, 1):
            out.append(self._clasificar(ruta))
            self._progreso(i, total)
        return out

    def _clasificar(self, ruta: str) -> dict:
        """Lee un archivo como reporte bancario y, si no lo es, como insumo.

        El orden no es arbitrario: los reportes bancarios son el caso de todos los
        días y los insumos la excepción, así que ante la duda gana el banco. Hoy
        no hay duda —ninguno de los dos detectores reclama archivos del otro—,
        pero el orden deja el comportamiento definido si algún día la hubiera.

        Cuando no es ninguno de los dos se reporta el fallo del LECTOR BANCARIO,
        no el del insumo: para un archivo que el usuario creía un reporte de
        banco, «no se reconoce de qué banco es» dice mucho más que «no parece
        ninguno de los insumos de flujo»."""
        registro = {"ruta": ruta, "clase": "banco", "banco": "", "cuentas": 0,
                    "tipo": "", "filas": 0, "error": "", "lineas": [],
                    "datos": None, "secciones": []}
        try:
            lineas, banco = saldos_lectores.leer(ruta)
            registro["banco"] = banco
            registro["cuentas"] = len(lineas)
            registro["lineas"] = lineas
            return registro
        except saldos_lectores.ErrorLector as exc:
            fallo_banco = str(exc)
        except Exception as exc:  # noqa: BLE001 — un archivo raro no aborta
            fallo_banco = f"{type(exc).__name__}: {exc}"

        try:
            tipo, datos = saldos_insumos.leer(ruta)
        except Exception:  # noqa: BLE001 — tampoco es insumo
            registro["error"] = fallo_banco
            return registro

        if tipo == "COMBINADO":
            # Un libro con varias pestañas reconocibles: trae más de una sección.
            registro.update(clase="insumo", tipo=tipo, datos=datos,
                            filas=sum(_filas_de(d) for d in datos.values()),
                            secciones=sorted(datos))
            return registro
        registro.update(clase="insumo", tipo=tipo, datos=datos,
                        filas=_filas_de(datos), secciones=[tipo])
        return registro

    def _guardar_insumos(self, insumos: list) -> None:
        """Funde lo recién subido con lo guardado y lo persiste.

        Solo se tocan las secciones que vinieron en los archivos: subir la nómina
        no puede borrar las facturas de Pemex capturadas la semana pasada. Corre
        en un hilo porque reescribe un libro de unas 32 000 filas."""
        nuevos = {}
        for registro in insumos:
            datos = registro.get("datos")
            if datos is None:
                continue
            if registro["tipo"] == "COMBINADO":
                nuevos.update(datos)
            else:
                nuevos[registro["tipo"]] = datos
        self.guardados = saldos_estado.fusionar_insumos(self.guardados, nuevos)
        saldos_estado.guardar_insumos(self.guardados)

    def _progreso(self, hechos: int, total: int) -> None:
        """Avance desde el hilo de lectura. `page.run_task` marshala al hilo de la
        interfaz, que es el único que puede tocar los controles de Flet.

        Va protegido a propósito: esto se llama DENTRO del hilo que está leyendo los
        archivos, y si fallara —la ventana se cerró a media carga, por ejemplo—
        tumbaría la lectura completa. Perder la barra de progreso es un detalle;
        perder los reportes ya leídos, no."""
        try:
            self.page.run_task(self._aplicar_progreso, hechos, total)
        except Exception:  # noqa: BLE001 — informar el avance nunca es crítico
            pass

    async def _aplicar_progreso(self, hechos: int, total: int) -> None:
        if not self._leyendo:
            return   # llegó tarde: la lectura ya terminó
        self.barra.value = hechos / total if total else 0
        self.txt_estado.value = f"Leyendo archivos… {hechos} de {total}"
        self._refrescar(self.barra, self.txt_estado)

    def _limpiar(self, _e=None) -> None:
        # Solo se quita lo de ESTA sesión. Lo guardado se borra sección por
        # sección desde su pestaña: 'Quitar todo' es para volver a empezar la
        # carga del día, no para tirar la captura de la semana.
        self.archivos.clear()
        self.insumos.clear()
        self.lineas.clear()
        self.asignacion = None
        self._pintar()

    # ------------------------------------------------------- identificación
    def _reidentificar(self) -> None:
        """Recalcula la colocación con TODOS los archivos cargados.

        Se rehace completa en vez de ir sumando por archivo porque la detección de
        cuentas repetidas necesita ver el conjunto: la misma cuenta puede venir en
        dos descargas distintas (a BBVA se le bajan cinco archivos)."""
        self.lineas = [x for a in self.archivos for x in a["lineas"]]
        self.asignacion = (motor.identificar(self.lineas, self.plantilla)
                           if self.lineas and self.plantilla else None)

    # ---------------------------------------------------------------- pintado
    def _fila_archivo(self, a: dict) -> ft.Control:
        if a["error"]:
            icono = ft.Icon(ft.Icons.ERROR_OUTLINE, size=18, color=ROJO)
            detalle = [ft.Text(a["error"], size=11, color=ROJO, expand=True,
                               max_lines=2,
                               overflow=ft.TextOverflow.ELLIPSIS)]
        else:
            icono = ft.Icon(ft.Icons.CHECK_CIRCLE_OUTLINE, size=18, color=VERDE)
            detalle = [_etiqueta(a["banco"], ft.Colors.PRIMARY),
                       ft.Text(f"{a['cuentas']} cuenta(s)", size=12, color=GRIS,
                               width=92, text_align=ft.TextAlign.RIGHT)]
        return self._fila(os.path.basename(a["ruta"]), a["ruta"], icono, detalle)

    def _fila_insumo(self, x: dict) -> ft.Control:
        if x["error"]:
            icono = ft.Icon(ft.Icons.ERROR_OUTLINE, size=18, color=ROJO)
            detalle = [ft.Text(x["error"], size=11, color=ROJO, expand=True,
                               max_lines=2,
                               overflow=ft.TextOverflow.ELLIPSIS)]
        else:
            icono = ft.Icon(ft.Icons.CHECK_CIRCLE_OUTLINE, size=18, color=VERDE)
            detalle = [
                _etiqueta(_NOMBRES_INSUMO.get(x["tipo"], x["tipo"]),
                          ft.Colors.TERTIARY),
                ft.Text(f"{x['filas']:,} fila(s)", size=12, color=GRIS,
                        width=92, text_align=ft.TextAlign.RIGHT)]
        return self._fila(os.path.basename(x["ruta"]), x["ruta"], icono, detalle)

    @staticmethod
    def _fila(nombre: str, ruta: str, icono, detalle: list) -> ft.Control:
        return ft.Container(
            content=ft.Row(
                [icono,
                 ft.Text(nombre, size=12, expand=True, max_lines=1, no_wrap=True,
                         overflow=ft.TextOverflow.ELLIPSIS, tooltip=ruta),
                 *detalle],
                spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            padding=ft.Padding.symmetric(vertical=7, horizontal=6),
            border=ft.Border(bottom=ft.BorderSide(1, ft.Colors.OUTLINE_VARIANT)))

    # ------------------------------------------------------------------ hero
    def _panel_cobertura(self) -> ft.Control:
        """La cifra que manda: cuántos renglones del formato quedaron llenos.

        Va grande y con barra de color porque es lo que el usuario tiene que
        mirar antes de generar. Los totales por divisa van al lado porque es lo
        segundo que revisa tesorería."""
        res = self.asignacion
        razon = res.llenos / res.total_renglones if res.total_renglones else 0
        color = (VERDE if razon >= _COBERTURA_BUENA
                 else NARANJA if razon >= _COBERTURA_REGULAR else ROJO)

        izquierda = ft.Column(
            [
                ft.Row(
                    [ft.Text(str(res.llenos), size=34,
                             weight=ft.FontWeight.BOLD, color=color),
                     ft.Text(f"/ {res.total_renglones}", size=16, color=GRIS)],
                    spacing=6, tight=True,
                    vertical_alignment=ft.CrossAxisAlignment.END),
                ft.Text("renglones del formato", size=11, color=GRIS),
            ],
            spacing=0, tight=True)

        centro = ft.Column(
            [
                ft.ProgressBar(value=razon, color=color, bar_height=8,
                               border_radius=4,
                               bgcolor=ft.Colors.with_opacity(0.15, color)),
                ft.Row([ft.Text(f"{razon:.0%} de cobertura", size=11,
                                color=GRIS),
                        ft.Text(f"{res.pegadas} filas pegadas en las pestañas",
                                size=11, color=GRIS)],
                       alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            ],
            spacing=6, tight=True)

        totales = res.totales()
        derecha = ft.Column(
            [self._total(divisa, monto)
             for divisa, monto in sorted(totales.items())] or
            [ft.Text("Sin saldos", size=12, color=GRIS)],
            spacing=2, tight=True,
            horizontal_alignment=ft.CrossAxisAlignment.END)

        # Sin `wrap`: activarlo convierte el Row en un Wrap, donde `expand` deja
        # de significar «toma el ancho sobrante» y la barra se estira a toda la
        # altura de la tarjeta. Los tres bloques caben de sobra en un renglón.
        return ft.Container(
            content=ft.Row(
                [izquierda,
                 ft.Container(content=centro, expand=True,
                              padding=ft.Padding.symmetric(horizontal=20)),
                 derecha],
                vertical_alignment=ft.CrossAxisAlignment.CENTER, spacing=16),
            padding=16, border_radius=12,
            bgcolor=ft.Colors.SURFACE_CONTAINER_HIGHEST)

    @staticmethod
    def _total(divisa: str, monto: float) -> ft.Control:
        etiqueta = "MX" if divisa == "MXN" else "DLS" if divisa == "USD" else divisa
        return ft.Row(
            [ft.Text(etiqueta, size=11, color=GRIS, width=32,
                     text_align=ft.TextAlign.RIGHT),
             ft.Text(f"${monto:,.2f}", size=15, weight=ft.FontWeight.BOLD)],
            spacing=8, tight=True,
            vertical_alignment=ft.CrossAxisAlignment.CENTER)

    # --------------------------------------------------- saldos por banco
    def _grupos_por_banco(self) -> list:
        """Un desplegable por pestaña, con sus cuentas y su total.

        Sirve para revisar ANTES de generar: los totales por banco se cotejan
        contra el portal sin abrir el Excel, y un saldo raro se ve a tiempo.

        Van COLAPSADOS y sus cuentas se construyen al desplegar (`on_change`),
        no antes. Son 216 cuentas: pintarlas todas de golpe son más de mil
        widgets y el render se nota — la misma razón por la que esta pantalla
        nunca tuvo una tabla. Colapsado son 15 renglones, y en la práctica se
        abren uno o dos.

        El orden es alfabético, no por monto: así la lista se ve igual todos los
        días y se puede comparar con la de ayer de un vistazo."""
        por_hoja: dict = {}
        for colocada in self.asignacion.colocadas.values():
            por_hoja.setdefault(colocada.destino.hoja, []).append(colocada)

        grupos = []
        for hoja in sorted(por_hoja):
            cuentas = sorted(por_hoja[hoja],
                             key=lambda c: -abs(c.saldo or 0.0))
            grupos.append(self._grupo_banco(hoja, cuentas))
        return grupos

    def _grupo_banco(self, hoja: str, cuentas: list) -> ft.Control:
        # Hay pestañas del formato donde conviven dos bancos: 'BX+ SCO' lleva
        # Scotiabank y Ve por Más. Ahí el nombre del grupo NO dice de quién es
        # cada cuenta, así que cada renglón trae su propio banco. En las demás
        # sería ruido: repetiría el encabezado en cada línea.
        compartida = hoja in saldos_plantilla.HOJAS_COMPARTIDAS
        totales: dict = {}
        for c in cuentas:
            divisa = (c.linea.moneda or "MXN").upper()
            totales[divisa] = totales.get(divisa, 0.0) + (c.saldo or 0.0)
        resumen = "  ·  ".join(
            "{} ${:,.2f}".format("MX" if d == "MXN" else "DLS" if d == "USD"
                                 else d, m)
            for d, m in sorted(totales.items()))

        tile = ft.ExpansionTile(
            title=ft.Row(
                [ft.Text(hoja, size=13, weight=ft.FontWeight.W_500,
                         expand=True),
                 ft.Text(f"{len(cuentas)} cuenta(s)", size=11, color=GRIS),
                 ft.Text(resumen, size=12, weight=ft.FontWeight.BOLD)],
                spacing=16, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            dense=True, min_tile_height=40,
            tile_padding=ft.Padding.symmetric(horizontal=6),
            controls_padding=ft.Padding.only(left=34, right=6, bottom=8),
            controls=[])
        # Las filas se arman la primera vez que se abre y se conservan: volver a
        # plegar y desplegar no debe rehacer el trabajo.
        def al_cambiar(_e, tile=tile, cuentas=cuentas):
            if not tile.controls:
                tile.controls = [self._fila_cuenta(c, compartida)
                                 for c in cuentas]
                self._refrescar(tile)   # traga el update si aún no está montado
        tile.on_change = al_cambiar
        return tile

    @staticmethod
    def _fila_cuenta(colocada, con_banco: bool = False) -> ft.Control:
        """Una cuenta dentro del desplegable de su pestaña.

        `con_banco` solo va en las pestañas que mezclan instituciones, donde el
        título del grupo no alcanza para saber de quién es la cuenta."""
        destino = colocada.destino
        divisa = (colocada.linea.moneda or "MXN").upper()
        celdas = [
            ft.Text(str(destino.cuenta or colocada.linea.cuenta), size=11,
                    color=GRIS, width=130,
                    no_wrap=True, overflow=ft.TextOverflow.ELLIPSIS),
            ft.Text(destino.titular or colocada.linea.titular or "",
                    size=12, expand=True, max_lines=1, no_wrap=True,
                    overflow=ft.TextOverflow.ELLIPSIS),
        ]
        if con_banco:
            celdas.append(_etiqueta(colocada.linea.banco, ft.Colors.PRIMARY))
        celdas += [
            ft.Text("" if divisa == "MXN" else divisa, size=11, color=GRIS,
                    width=34, text_align=ft.TextAlign.RIGHT),
            ft.Text(f"{colocada.saldo:,.2f}", size=12, width=126,
                    text_align=ft.TextAlign.RIGHT),
        ]
        return ft.Container(
            content=ft.Row(celdas, spacing=10,
                           vertical_alignment=ft.CrossAxisAlignment.CENTER),
            padding=ft.Padding.symmetric(vertical=4))

    # ------------------------------------------------- insumos persistidos
    def _filas_secciones(self) -> list:
        """Una fila por sección, con lo que hay guardado y cómo borrarlo.

        Se listan LAS SEIS aunque estén vacías: así se ve de un vistazo qué falta
        capturar, no solo lo que ya está."""
        filas = []
        for seccion in saldos_insumos.SECCIONES:
            datos = self.guardados.get(seccion)
            n = _filas_de(datos)
            nombre = _NOMBRES_INSUMO.get(seccion, seccion.capitalize())
            if n:
                icono = ft.Icon(ft.Icons.CHECK_CIRCLE_OUTLINE, size=18,
                                color=VERDE)
                detalle = ft.Text(f"{n:,} fila(s)", size=12, color=GRIS,
                                  width=110, text_align=ft.TextAlign.RIGHT)
                borrar = ft.IconButton(
                    icon=ft.Icons.DELETE_OUTLINE, icon_size=18,
                    icon_color=ROJO_BOTON, tooltip=f"Vaciar {nombre}",
                    on_click=lambda _e, sec=seccion: self.page.run_task(
                        self._olvidar, sec))
            else:
                icono = ft.Icon(ft.Icons.REMOVE_CIRCLE_OUTLINE, size=18,
                                color=GRIS)
                detalle = ft.Text("sin capturar", size=12, color=GRIS,
                                  italic=True, width=110,
                                  text_align=ft.TextAlign.RIGHT)
                borrar = ft.Container(width=40)
            filas.append(ft.Container(
                content=ft.Row(
                    [icono, ft.Text(nombre, size=12, expand=True), detalle,
                     borrar],
                    spacing=10,
                    vertical_alignment=ft.CrossAxisAlignment.CENTER),
                padding=ft.Padding.symmetric(vertical=2, horizontal=6),
                border=ft.Border(
                    bottom=ft.BorderSide(1, ft.Colors.OUTLINE_VARIANT))))

        nota = ("Estas secciones se conservan entre corridas: solo cambian "
                "cuando subes un archivo que las traiga o las vacías aquí. "
                "Los impuestos aún no los lee el reporte; la pestaña existe "
                "para poder irlos capturando.")
        filas.append(ft.Container(
            content=ft.Text(nota, size=11, color=GRIS),
            padding=ft.Padding.only(top=10, left=6, right=6)))
        return filas

    async def _olvidar(self, seccion: str) -> None:
        # Se le pasa lo que ya está en memoria: releer el libro para quitarle una
        # sección costaría varios segundos con la interfaz congelada.
        await self._asegurar_estado()
        self.guardados = await asyncio.to_thread(
            saldos_estado.olvidar_insumos, seccion, self.guardados)
        self._pintar()
        self.app.avisar(
            "Se vació {}.".format(_NOMBRES_INSUMO.get(seccion, seccion)),
            NARANJA)

    async def _descargar_formato(self, _e=None) -> None:
        """Entrega el libro de insumos para llenarlo y volverlo a subir.

        Sale con lo que ya está capturado, no en blanco: quien solo va a
        actualizar la nómina no tiene que recapturar las 16 000 filas de MGC."""
        ruta = await self.app.picker.save_file(
            dialog_title="Guardar el formato de insumos de flujo",
            file_name="INSUMOS DE FLUJO.xlsx", allowed_extensions=["xlsx"])
        if not ruta:
            return
        if not ruta.lower().endswith(".xlsx"):
            ruta += ".xlsx"

        # El diálogo de espera va DESPUÉS de elegir el archivo, nunca antes: el
        # selector del sistema es otra ventana y quedaría detrás del modal.
        #
        # `_asegurar_estado` entra aquí dentro a propósito. Normalmente ya está
        # resuelto —lo dispara `al_entrar`—, pero si el usuario pulsa el botón
        # antes de que termine, la lectura de los insumos se lleva sus segundos y
        # es justo lo que este modal existe para cubrir.
        escritas = fallo = None
        self._abrir_espera(
            "Preparando el formato de insumos…",
            "Se está armando el libro con lo que ya tienes capturado. "
            "Puede tardar unos segundos.")
        try:
            await self._asegurar_estado()
            escritas = await asyncio.to_thread(
                saldos_insumos.escribir_plantilla, ruta, self.guardados)
        except BaseException as exc:  # noqa: BLE001 — se reporta abajo, ya cerrado
            fallo = exc
        finally:
            # Se cierra ANTES de avisar nada: `app.avisar` apila el snackbar en la
            # MISMA pila que el modal, así que avisar primero se llevaría el
            # snackbar por delante y dejaría el modal puesto para siempre.
            self._cerrar_espera()

        if isinstance(fallo, PermissionError):
            self.app.avisar(
                "No se pudo guardar: el archivo está abierto en Excel.", ROJO)
            return
        if fallo is not None:
            registro = self._registrar_error(fallo, ruta)
            mensaje = "No se pudo generar el formato: {}".format(fallo)
            if registro:
                self.app.avisar(mensaje, ROJO, accion="Ver detalle",
                                on_accion=lambda _e=None: self.app.abrir_en_sistema(
                                    registro),
                                duracion=ft.Duration(seconds=12))
            else:
                self.app.avisar(mensaje, ROJO)
            return

        con_datos = sum(1 for n in escritas.values() if n)
        self.app.avisar(
            "Formato de insumos generado ({} de {} secciones con datos).".format(
                con_datos, len(escritas)),
            VERDE, accion="Abrir",
            on_accion=lambda _e=None: self.app.abrir_en_sistema(ruta),
            duracion=ft.Duration(seconds=12))

    def _lista_avisos(self) -> list:
        """Lo que hay que saber antes de generar, de más grave a menos."""
        res = self.asignacion
        avisos = []

        faltan = res.bancos_faltantes()
        if faltan:
            avisos.append((
                ft.Icons.WARNING_AMBER, NARANJA,
                "No llegó ningún saldo de " + ", ".join(faltan)
                + ". Si te faltó subir ese reporte, sus renglones saldrán "
                  "vacíos."))
        elif res.vacios:
            avisos.append((
                ft.Icons.HELP_OUTLINE, NARANJA,
                f"{len(res.vacios)} renglón(es) quedaron vacíos porque ningún "
                f"archivo trajo esa cuenta."))

        if res.nuevas:
            monto = sum(x.linea.saldo for x in res.nuevas)
            avisos.append((
                ft.Icons.ADD_CIRCLE_OUTLINE, NARANJA,
                f"{len(res.nuevas)} cuenta(s) llegaron y no están en ninguna "
                f"pestaña del formato (${monto:,.2f}). Van en la hoja «Cuentas "
                f"nuevas»; para que entren al reporte hay que actualizar la "
                f"plantilla."))

        if res.duplicados:
            avisos.append((
                ft.Icons.CONTENT_COPY, NARANJA,
                f"{len(res.duplicados)} cuenta(s) venían repetidas en más de un "
                f"archivo; se contaron una sola vez."))

        return [self._aviso(icono, color, texto)
                for icono, color, texto in avisos]

    @staticmethod
    def _aviso(icono, color, texto: str) -> ft.Control:
        """Aviso con una franja del color del acento a la izquierda: se distingue
        del texto normal sin tener que teñir toda la letra."""
        return ft.Container(
            content=ft.Row(
                [ft.Icon(icono, size=17, color=color),
                 ft.Text(texto, size=12, expand=True)],
                spacing=10, vertical_alignment=ft.CrossAxisAlignment.START),
            padding=ft.Padding.symmetric(vertical=8, horizontal=12),
            border_radius=8,
            bgcolor=ft.Colors.with_opacity(0.07, color),
            border=ft.Border(left=ft.BorderSide(3, color)))

    def _pintar(self) -> None:
        # Reportes e insumos van en la MISMA lista: el usuario subió una sola
        # pila por un solo botón, y separarlos en pantalla volvería a sugerir
        # que son dos trámites distintos. El marbete de cada fila dice cuál es.
        #
        # Los que fallaron se van al final. Intercalados se pierden entre
        # veintitantos renglones verdes, y son justo los que hay que atender:
        # agrupados abajo se leen de un vistazo. Dentro de cada grupo se respeta
        # el orden de carga (`sorted` es estable), para que la lista no se
        # reacomode sola entre una carga y la siguiente.
        cargados = sorted(self.archivos + self.insumos,
                          key=lambda x: bool(x["error"]))
        hay = bool(cargados)
        self.lista.controls = [
            self._fila_insumo(x) if x["clase"] == "insumo"
            else self._fila_archivo(x) for x in cargados]
        self.vacio.visible = not hay
        self.btn_limpiar.visible = hay
        self.tab_archivos.label = f"Archivos cargados ({len(cargados)})"
        # Basta con un archivo bancario leído: los insumos nunca son requisito.
        self.btn_generar.disabled = not (self.asignacion
                                         and self.asignacion.colocadas)

        if self.asignacion is not None:
            self.hero.content = self._panel_cobertura()
            self.hero.visible = True
            self.avisos.controls = self._lista_avisos()
            self.avisos.visible = bool(self.avisos.controls)
            self.bancos.controls = self._grupos_por_banco()
            self.tab_saldos.label = "Saldos identificados ({})".format(
                self.asignacion.pegadas)
        else:
            self.hero.visible = False
            self.hero.content = None
            self.avisos.visible = False
            self.avisos.controls = []
            self.bancos.controls = []
            self.tab_saldos.label = "Saldos identificados (0)"

        # La barra solo tiene sentido con algo que mostrar. Al aparecer se abre
        # en «Saldos identificados», que es lo que hay que revisar antes de
        # generar; la lista de archivos es el respaldo de qué se subió.
        self.pestanas.visible = hay or bool(self.guardados)
        self._aplicar_pestana()

        self.secciones.controls = self._filas_secciones()
        vivos = sum(1 for x in self.guardados.values() if x)
        self.tab_insumos.label = "Insumos de flujo ({} de {})".format(
            vivos, len(saldos_insumos.SECCIONES))
        self._refrescar(self.lista, self.vacio, self.pestanas,
                        self.btn_limpiar, self.btn_generar, self.hero,
                        self.avisos, self.bancos, self.secciones,
                        self.cargando_insumos)

    def _cambiar_pestana(self, e) -> None:
        self._pestana = e.control.selected_index
        self._aplicar_pestana()
        self._refrescar(self.bancos, self.lista, self.secciones)

    def _aplicar_pestana(self) -> None:
        """Muestra solo la vista de la pestaña activa.

        Se alterna la visibilidad en vez de reconstruir: los desplegables de
        banco que el usuario ya abrió conservan sus filas al ir y volver."""
        hay = bool(self.archivos or self.insumos)
        self.bancos.visible = hay and self._pestana == 0
        self.lista.visible = hay and self._pestana == 1
        self.secciones.visible = self._pestana == 2

    def _ocupado(self, activo: bool, mensaje: str = "") -> None:
        self._leyendo = activo
        self.anillo.visible = activo
        self.barra.visible = activo
        self.barra.value = 0 if activo else None
        self.btn_cargar.disabled = activo
        self.txt_estado.value = mensaje
        self._refrescar(self.anillo, self.barra, self.btn_cargar,
                        self.txt_estado)

    @staticmethod
    def _refrescar(*controles) -> None:
        """Update DIRIGIDO a los controles tocados, no `page.update()`: la pantalla
        vive montada junto a las demás y refrescar la página entera hace trabajo de
        más (ver el mismo patrón en ui/devoluciones.py)."""
        for c in controles:
            try:
                c.update()
            except (RuntimeError, AssertionError):
                pass  # aún no montado; se reflejará al renderizar

    # ------------------------------------------------------------- generación
    def _abrir_espera(self, titulo: str = "Generando el reporte de saldos…",
                      detalle: str = "Se están pegando los saldos en cada "
                                     "pestaña del formato. Puede tardar unos "
                                     "segundos.") -> None:
        """Modal de espera mientras se arma un libro.

        Hace falta porque escribir estos libros tarda varios segundos —el reporte
        abre el libro base, pega 216 filas y vuelca casi 32 000 renglones de
        ledgers; el formato de insumos vuelca esos mismos ledgers— y hasta ahora
        no pasaba nada visible: el usuario no sabía si estaba trabajando o se
        había colgado, y volvía a picarle al botón.

        Es MODAL a propósito: además de informar, impide dispararlo dos veces.
        Mismo patrón que en dispersión (No Pemex)."""
        self._dlg_espera = ft.AlertDialog(
            modal=True,
            content=ft.Column(
                [ft.Text(titulo, size=18,
                         weight=ft.FontWeight.BOLD,
                         text_align=ft.TextAlign.CENTER),
                 ft.ProgressRing(width=32, height=32, stroke_width=3),
                 ft.Text(detalle, size=12, color=GRIS,
                         text_align=ft.TextAlign.CENTER)],
                spacing=18, tight=True,
                horizontal_alignment=ft.CrossAxisAlignment.CENTER),
        )
        self.page.show_dialog(self._dlg_espera)
        self.page.update()

    def _cerrar_espera(self) -> None:
        """Cierra el modal de espera. Tiene que funcionar SIEMPRE.

        `page.pop_dialog()` saca lo que esté encima de la pila, no un diálogo en
        concreto, así que además se marca el nuestro como cerrado: si algo más se
        apiló en medio, `open = False` lo baja igual."""
        dlg, self._dlg_espera = getattr(self, "_dlg_espera", None), None
        if dlg is None:
            return
        try:
            self.page.pop_dialog()
        except Exception:  # noqa: BLE001 — se reintenta por la otra vía
            pass
        try:
            dlg.open = False
            self.page.update()
        except Exception:  # noqa: BLE001 — nada más que hacer; ya no bloquea
            pass

    def _registrar_error(self, exc: BaseException, ruta: str) -> str:
        """Guarda el detalle del fallo en un archivo y devuelve su ruta.

        El snackbar solo cabe una línea y se va solo a los segundos; cuando esto
        falle en la máquina de tesorería, lo que hace falta para diagnosticarlo es
        el traceback completo, no «no se pudo generar». Se acumula (modo 'a') para
        no perder el caso anterior si vuelve a pasar."""
        destino = os.path.join(rutas.DATOS, "saldos_errores.log")
        try:
            with open(destino, "a", encoding="utf-8") as f:
                cabecera = datetime.datetime.now().strftime(
                    " %d/%m/%Y %H:%M:%S ")
                f.write("\n{:=^70}\n".format(cabecera))
                f.write("archivo destino: {}\n".format(ruta))
                f.write("reportes: {} | insumos: {}\n".format(
                    len(self.archivos), len(self.insumos)))
                f.write("".join(traceback.format_exception(
                    type(exc), exc, exc.__traceback__)))
        except Exception:  # noqa: BLE001 — no poder registrar no agrava el fallo
            return ""
        return destino

    async def _generar(self, _e=None) -> None:
        if not (self.asignacion and self.asignacion.colocadas):
            self.app.avisar("No hay saldos que reportar.", NARANJA)
            return
        await self._asegurar_estado()
        hoy = datetime.date.today().strftime("%d-%m-%Y")
        ruta = await self.app.picker.save_file(
            dialog_title="Guardar el reporte de saldos",
            file_name=f"SALDOS {hoy}.xlsx", allowed_extensions=["xlsx"])
        if not ruta:
            return
        if not ruta.lower().endswith(".xlsx"):
            ruta += ".xlsx"
        # Lo guardado manda como base y lo de esta sesión ya está fundido ahí,
        # así que se usa `self.guardados`: es lo mismo que el usuario descargaría.
        insumos = dict(self.guardados)
        anterior = saldos_estado.totales_dia_anterior()

        # NO se avisa nada mientras el modal está abierto. `app.avisar` muestra el
        # snackbar con `page.show_dialog`, o sea LA MISMA PILA que el modal: si se
        # avisa antes de cerrarlo, el `pop_dialog` de después se lleva el snackbar
        # y el modal se queda para siempre. Era justo el cuelgue que solo aparecía
        # cuando la generación fallaba.
        info = fallo = None
        self._abrir_espera()
        try:
            info = await asyncio.to_thread(
                saldos_export.generar, ruta, self.asignacion, insumos,
                None, anterior)
        except BaseException as exc:  # noqa: BLE001 — se reporta abajo, ya cerrado
            fallo = exc
        finally:
            self._cerrar_espera()

        if fallo is not None:
            if isinstance(fallo, PermissionError):
                self.app.avisar(
                    "No se pudo guardar: el archivo está abierto en Excel. "
                    "Ciérralo e intenta de nuevo (o guarda con otro nombre).",
                    ROJO)
                return
            if isinstance(fallo, (asyncio.CancelledError, KeyboardInterrupt)):
                raise fallo   # no es un fallo nuestro: que siga su curso
            registro = self._registrar_error(fallo, ruta)
            mensaje = f"No se pudo generar el reporte: {fallo}"
            if registro:
                mensaje += f"  ·  Detalle en {os.path.basename(registro)}"
            self.app.avisar(mensaje, ROJO, duracion=ft.Duration(seconds=20))
            return

        saldos_estado.guardar_totales(info.get("totales_cabecera") or {})

        detalle = f"{info['llenos']} de {info['renglones']} renglones"
        if info["nuevas"]:
            detalle += f"; {info['nuevas']} cuenta(s) nueva(s) aparte"
        self.app.avisar(
            f"Reporte de saldos generado: {detalle}. "
            + ("Incluye la comparativa contra el {}. ".format(anterior[0])
               if info.get("comparativas") else "")
            + "Los totales aparecen al abrirlo en Excel.", VERDE, accion="Abrir",
            on_accion=lambda _e=None: self.app.abrir_en_sistema(ruta),
            duracion=ft.Duration(seconds=12))
