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
from core import diagnostico, portapapeles
from core import (saldos_estado, saldos_export, saldos_insumos,
                  saldos_lectores, saldos_plantilla)
from ui.comun import CENTRO, GRIS, NARANJA, ROJO, ROJO_BOTON, VERDE, tarjeta

# Todo lo que puede entrar: los portales emiten estos formatos (ver
# core/saldos_lectores) y los insumos de flujo siempre son hojas de cálculo.
_EXTENSIONES = ["xlsx", "xls", "csv", "txt", "pdf"]

# Alto máximo de la vista previa del pegado. Da para una docena de renglones —lo
# que suele traer un portal— sin que el diálogo se estire hasta salirse de la
# pantalla cuando alguien pega una tabla de doscientas filas.
_ALTO_VISTA_PEGADO = 260
# Hasta cuántas pestañas sin llegar se consideran «va avanzando» (ámbar) en vez
# de «apenas empieza» (rojo). Con quince bancos, tres pendientes es la recta
# final de una carga normal.
_BANCOS_POCOS = 3
# Ancho máximo del contenido de la pantalla. Más allá, las filas se estiran y
# separan el titular de su importe hasta perder la lectura.
_ANCHO_CONTENIDO = 1180
# Ancho útil de la vista previa y de la caja del diálogo. La tabla se reparte
# dentro de ese ancho, así que nunca hay que desplazarse en horizontal.
_ANCHO_DIALOGO_PEGADO = 860
_ANCHO_VISTA_PEGADO = 810
# Un carácter de Consolas a 11 px. Sirve para repartir el ancho entre columnas
# según lo que de verdad ocupa cada una, en vez de darles a todas lo mismo.
_ANCHO_CARACTER = 6.6
# Separación entre columnas de la vista previa. Se descuenta del presupuesto de
# ancho al repartirlo (ver `_anchos_columnas`).
_ESPACIO_COLUMNA = 8

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
        aviso = None
        try:
            self.guardados = await asyncio.to_thread(
                saldos_estado.cargar_insumos)
        except saldos_estado.ErrorEstado as exc:
            # Se avisa, no se traga: si lo guardado se rompió, el usuario tiene
            # que enterarse ANTES de generar un reporte con los paneles en cero.
            self.guardados = {}
            aviso = str(exc)
        except BaseException as exc:  # noqa: BLE001 — nunca debe tumbar la pantalla
            self._registrar_error(exc, "(lectura de insumos guardados)")
            self.guardados = {}
        finally:
            # El indicador se apaga AQUÍ, no después de pintar: si el repintado
            # fallara, dejarlo prendido es lo peor que puede pasar —la pantalla
            # se queda «Recuperando…» para siempre y parece colgada—. Solo se
            # cambia la bandera; de pintarla se encarga `_pintar` más abajo.
            self._estado_cargado = True
            self._cargando_estado = False
            self._estado_listo.set()
            self.cargando_insumos.visible = False

        # NO se pinta si el usuario tiene abierto el navegador de archivos. Esta
        # lectura tarda sus segundos y es normal que termine justo mientras él
        # está eligiendo un archivo; tocar los controles en ese momento congela
        # la ventana entera. Se espera a que cierre y recién entonces se pinta.
        await self.app.esperar_sin_dialogo_archivos()
        self._pintar()
        if aviso:
            self.app.avisar("No se recuperaron los insumos: " + aviso, NARANJA,
                            duracion=ft.Duration(seconds=15))

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
        # Tres formas de traer lo mismo, porque el diálogo de multiselección es
        # justo el que da problemas con muchos archivos de nombre largo —y los
        # reportes de los portales son eso—. Con la carpeta se evita elegirlos
        # uno por uno; con Ctrl+V, el diálogo entero.
        self.btn_carpeta = ft.TextButton(
            content="Cargar carpeta", icon=ft.Icons.FOLDER_OPEN_OUTLINED,
            tooltip="Toma de una carpeta todo lo que la herramienta reconozca, "
                    "sin abrir el selector de archivos",
            on_click=self._cargar_carpeta)
        self.btn_limpiar = ft.TextButton(
            content="Quitar todo", icon=ft.Icons.DELETE_SWEEP_OUTLINED,
            visible=False, on_click=self._limpiar,
            style=ft.ButtonStyle(color=ROJO_BOTON))
        # Aparece y desaparece con «Quitar todo»: un separador solo, colgando
        # después del último botón, se lee como un fallo de dibujo.
        self.separador = ft.Container(width=1, height=22, visible=False,
                                      bgcolor=ft.Colors.OUTLINE_VARIANT)
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
        # A la izquierda SOLO lo de todos los días —traer archivos—; a la
        # derecha el resultado. Antes convivían ahí cuatro acciones con el mismo
        # peso: «Descargar formato de insumos», que es una tarea semanal, y
        # «Quitar todo», que es destructiva, entre las dos de cargar. La primera
        # se movió a la pestaña de Insumos, que es su contexto; la segunda queda
        # apartada del grupo por un separador.
        acciones = ft.Row(
            [
                ft.Row([self.btn_cargar, self.btn_carpeta, self.separador,
                        self.btn_limpiar, self.anillo, self.txt_estado],
                       spacing=8, expand=True, wrap=True,
                       vertical_alignment=ft.CrossAxisAlignment.CENTER),
                self.btn_generar,
            ],
            spacing=12, vertical_alignment=ft.CrossAxisAlignment.CENTER)

        self.hero = ft.Container(visible=False)
        # Las quince pestañas del formato, justo debajo de la cifra de cobertura:
        # la cifra dice CUÁNTO falta y el tablero dice QUÉ falta.
        self.tablero = ft.Container(visible=False)
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
                self.tablero,
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
        # Ancho máximo y centrado. A pantalla completa las filas se estiraban a
        # 1500 px y quedaban mil de vacío entre el titular y su importe: el ojo
        # pierde el renglón a medio camino. Acotado, la fila se lee de un golpe.
        return ft.Row(
            [ft.Container(
                content=ft.Column(controles, spacing=14,
                                  scroll=ft.ScrollMode.AUTO, expand=True),
                width=_ANCHO_CONTENIDO, expand=False)],
            alignment=ft.MainAxisAlignment.CENTER, expand=True)

    def _zona_vacia(self) -> ft.Control:
        """Estado inicial.

        A propósito NO se parece a una zona de arrastre. Flet 0.85 no expone
        ningún evento para soltar archivos del sistema operativo —`DragTarget` y
        `Draggable` solo mueven controles dentro de la app—, así que un marco
        punteado con una nube prometería algo que no funciona: el usuario
        arrastraría, no pasaría nada, y pensaría que la app está rota.

        Lo que sí se puede es PEGAR, y eso se dice aquí porque no hay forma de
        adivinarlo: copiar los archivos en el Explorador y soltarlos con Ctrl+V
        se acerca bastante a arrastrarlos, y de paso se salta el diálogo.

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
                         size=13, color=GRIS),
                 ft.Text("Cópialos en el Explorador y pégalos aquí con Ctrl+V, "
                         "o usa «Cargar carpeta»",
                         size=11, color=GRIS, italic=True)],
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
        # Rastro paso a paso. Hay un cuelgue que solo aparece en una máquina, no
        # deja traceback —no es una excepción, se queda bloqueada— y obliga a
        # matar el proceso. Con esto, el último renglón del log dice en qué punto
        # exacto se detuvo. Cuesta unos microsegundos por paso.
        diagnostico.registrar("saldos._cargar: pidiendo archivos")
        archivos = await self.app.picker.pick_files(
            dialog_title="Selecciona los reportes de saldos y los insumos",
            allowed_extensions=_EXTENSIONES, allow_multiple=True)
        diagnostico.registrar("saldos._cargar: seleccionados",
                              "{}".format(len(archivos or [])))
        if not archivos:
            return
        await self._ingerir([a.path for a in archivos])

    async def _cargar_carpeta(self, _e=None) -> None:
        """Toma de una carpeta todo lo que reconozca, sin elegir uno por uno.

        Es la vía cómoda para la carpeta de descargas del navegador, donde caen
        los reportes tal cual los deja cada portal. Y evita el diálogo de
        MULTISELECCIÓN, que en esta app ya dio problemas con muchos archivos de
        nombre largo (misma nota en ui/alta_beneficiarios) — y los reportes de
        saldos son justo eso.

        Solo el primer nivel: las subcarpetas suelen ser de otros días."""
        diagnostico.registrar("saldos._cargar_carpeta: pidiendo carpeta")
        carpeta = await self.app.picker.get_directory_path(
            dialog_title="Elige la carpeta con los reportes de saldos")
        diagnostico.registrar("saldos._cargar_carpeta: carpeta",
                              str(carpeta or "")[:300])
        if not carpeta:
            return
        exts = tuple("." + e.lower() for e in _EXTENSIONES)
        try:
            rutas = [os.path.join(carpeta, n) for n in sorted(os.listdir(carpeta))
                     if n.lower().endswith(exts)
                     and os.path.isfile(os.path.join(carpeta, n))]
        except OSError as exc:
            self.app.avisar(f"No se pudo leer la carpeta: {exc}", ROJO)
            return
        if not rutas:
            self.app.avisar(
                "Esa carpeta no tiene archivos que la herramienta pueda leer "
                "({}).".format(", ".join(_EXTENSIONES)), NARANJA)
            return
        await self._ingerir(rutas)

    def _on_teclado(self, e) -> None:
        """Ctrl+V pega los archivos que se hayan copiado en el Explorador.

        El evento llega a TODAS las pantallas —el slot de teclado es único y el
        shell lo reparte—, así que lo primero es comprobar que la nuestra sea la
        que está al frente."""
        if not getattr(self.contenido, "visible", False):
            return
        if not (e.ctrl and str(e.key).lower() in ("v", "insert")):
            return
        if self._leyendo:
            return   # ya hay una lectura en curso
        self.page.run_task(self._pegar)

    async def _pegar(self, _e=None) -> None:
        """Carga lo que haya en el portapapeles: archivos o una tabla copiada.

        Los archivos ganan: si hay ambos —pasa al copiar desde Excel— lo que el
        usuario quiso traer casi siempre es el archivo."""
        diagnostico.registrar("saldos._pegar: leyendo portapapeles")
        rutas = await asyncio.to_thread(portapapeles.archivos)
        diagnostico.registrar("saldos._pegar: rutas", "{}".format(len(rutas)))
        if rutas:
            await self._ingerir(rutas)
            return

        texto = await asyncio.to_thread(portapapeles.texto)
        diagnostico.registrar("saldos._pegar: texto",
                              "{} caracteres".format(len(texto or "")))
        if texto and texto.strip():
            self._confirmar_pegado(texto)
            return
        self.app.avisar(
            "No hay nada que pegar. Copia los archivos en el Explorador, o "
            "selecciona la tabla del portal del banco y cópiala con Ctrl+C.",
            NARANJA)

    def _confirmar_pegado(self, texto: str) -> None:
        """Muestra qué se pegó y de qué banco cree que es, antes de cargarlo.

        NO se carga en silencio a propósito. Un pegado no trae nombre de archivo
        ni extensión, así que la firma tiene que desambiguar sola y hay bancos
        que comparten encabezados (`alias`, `cuenta`, `divisa`). Enseñar la tabla
        y dejar corregir el banco cuesta un clic; meter los saldos de un banco en
        la pestaña de otro no se ve hasta que el reporte ya se firmó."""
        filas = saldos_lectores.filas_pegadas(texto)
        detectado = saldos_lectores.detectar_pegado(filas)
        bancos = saldos_lectores.bancos_pegables()

        # Sin banco no se puede cargar, y el botón apagado junto a un selector
        # vacío ya lo dice: sustituye al texto que antes explicaba que no se
        # había reconocido, sin gastar un renglón en decirlo.
        boton = ft.FilledButton("Cargar", disabled=detectado is None,
                                style=_estilo_verde())
        # En un Column `expand` estira a lo ALTO, así que para ocupar todo el
        # ancho el selector va dentro de un Row, donde `expand` es horizontal.
        selector = ft.Dropdown(
            label="Banco", value=detectado, expand=True,
            options=[ft.dropdown.Option(key=b, text=b) for b in bancos])

        def elegir(_ev=None):
            boton.disabled = not selector.value
            # Se refresca el DIÁLOGO, no el botón: los controles de `actions`
            # viven fuera del árbol del contenido y actualizarlos por separado no
            # llega a la pantalla —el botón se quedaba apagado aunque ya hubiera
            # banco elegido, y no había forma de cargar—.
            self._refrescar(dlg)

        # `on_select`, NO `on_change`: el Dropdown de Flet 0.85 solo expone el
        # primero (el segundo se puede asignar sin error y no dispara nunca, que
        # es lo que dejaba el botón apagado con el banco ya elegido). Mismo
        # evento que usan los selectores de dispersión.
        selector.on_select = elegir

        def cargar(_ev=None):
            self.page.pop_dialog()
            self.page.run_task(self._cargar_pegado, texto, selector.value)

        boton.on_click = cargar
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Text("Pegar datos"),
            content=ft.Column(
                [ft.Row([selector], tight=True), self._vista_previa(filas)],
                spacing=12, tight=True, width=_ANCHO_DIALOGO_PEGADO,
                horizontal_alignment=ft.CrossAxisAlignment.STRETCH),
            actions=[
                ft.TextButton("Cancelar",
                              on_click=lambda e: self.page.pop_dialog()),
                boton,
            ],
            actions_alignment=ft.MainAxisAlignment.END)
        self.page.show_dialog(dlg)

    @staticmethod
    def _vista_previa(filas: list) -> ft.Control:
        """Todo lo pegado, en una tabla con UNA sola barra de desplazamiento.

        Se muestran TODOS los renglones: quien pega viene a comprobar que están
        sus cuentas, y un «… y 12 más» deja fuera justo lo que quería revisar.

        Antes había dos barras y se estorbaban. La horizontal pertenecía a un
        renglón tan alto como la tabla entera, así que quedaba al fondo del
        CONTENIDO y no de la caja: para alcanzarla había que bajar hasta el final,
        y al llegar tapaba la última fila. Anidar scrolls al revés solo cambia
        cuál de las dos estorba.

        La solución es que sobre una. Las columnas se reparten el ancho según lo
        que de verdad ocupa cada una —«Hora» no necesita lo mismo que «Titular»—,
        así que la tabla cabe y solo queda la barra vertical, siempre en el mismo
        sitio. Lo que no cabe se recorta con puntos suspensivos y se puede leer
        completo en el tooltip de la celda."""
        if not filas:
            return ft.Container(
                content=ft.Text("(vacío)", size=11, color=GRIS),
                padding=10, border_radius=6,
                bgcolor=ft.Colors.with_opacity(0.05, ft.Colors.ON_SURFACE),
                height=_ALTO_VISTA_PEGADO)

        def texto(v):
            return str(v if v is not None else "").strip()

        anchos = SeccionSaldos._anchos_columnas(filas)

        def celda(v, ancho, cabecera=False):
            t = texto(v)
            return ft.Text(
                t, size=11, font_family="Consolas", width=ancho, no_wrap=True,
                overflow=ft.TextOverflow.ELLIPSIS,
                weight=ft.FontWeight.BOLD if cabecera else None,
                # El tooltip salva lo que el recorte esconde, sin gastar ancho.
                tooltip=t if len(t) * _ANCHO_CARACTER > ancho else None)

        renglones = [
            ft.Row([celda(v, a, i == 0) for v, a in zip(f, anchos)],
                   spacing=_ESPACIO_COLUMNA, tight=True)
            for i, f in enumerate(filas)]
        return ft.Container(
            content=ft.Column(renglones, scroll=ft.ScrollMode.AUTO, spacing=2,
                              tight=True),
            # El relleno derecho deja libre el carril de la barra vertical, para
            # que no se monte sobre la última columna.
            padding=ft.Padding.only(left=10, top=10, bottom=10, right=22),
            border_radius=6,
            bgcolor=ft.Colors.with_opacity(0.05, ft.Colors.ON_SURFACE),
            height=_ALTO_VISTA_PEGADO)

    @staticmethod
    def _anchos_columnas(filas: list) -> list:
        """Ancho en píxeles de cada columna, repartiendo `_ANCHO_VISTA_PEGADO`.

        Cada columna pide lo que ocupa su celda más larga. Si entre todas se
        pasan del ancho disponible, se recorta a las MÁS ANCHAS primero: son las
        de texto libre —titulares, mensajes de error— donde perder el final no
        impide reconocer la fila, mientras que cuentas e importes se leen enteros
        precisamente porque son cortos."""
        columnas = max((len(f) for f in filas), default=0)
        if not columnas:
            return []
        pedidos = []
        for c in range(columnas):
            largo = max((len(str(f[c]).strip()) if c < len(f) and f[c] is not None
                         else 0) for f in filas)
            pedidos.append(max(4, largo) * _ANCHO_CARACTER + 8)

        # Los separadores entre columnas también ocupan: si no se descuentan del
        # presupuesto, la tabla se pasa por justo esos pixeles y reaparece la
        # barra horizontal que este reparto existe para evitar.
        disponible = _ANCHO_VISTA_PEGADO - _ESPACIO_COLUMNA * (columnas - 1)
        # Se baja el techo hasta que la suma quepa. Es un reparto sencillo y
        # estable: no depende del orden de las columnas ni deja huecos.
        #
        # El piso es lo que le tocaría a cada columna a partes iguales: por
        # debajo de eso, encoger a las anchas ya no puede hacer caber la tabla y
        # el bucle solo estrecharía las angostas sin ganar nada. Con un piso fijo
        # —antes 40 px— una tabla de veinte columnas se pasaba igual.
        techo = max(pedidos)
        piso = max(28.0, disponible / columnas)
        while sum(min(p, techo) for p in pedidos) > disponible and techo > piso:
            techo -= 5
        return [min(p, techo) for p in pedidos]

    async def _cargar_pegado(self, texto: str, banco: str) -> None:
        """Interpreta la tabla pegada y la suma como un reporte más."""
        diagnostico.registrar("saldos._cargar_pegado", str(banco))
        self._ocupado(True, "Leyendo lo pegado…")
        try:
            lineas, nombre = await asyncio.to_thread(
                saldos_lectores.leer_pegado, texto, banco)
        except saldos_lectores.ErrorLector as exc:
            self._ocupado(False)
            self.app.avisar(str(exc), ROJO, duracion=ft.Duration(seconds=12))
            return
        except BaseException as exc:  # noqa: BLE001 — se reporta al usuario
            self._ocupado(False)
            self._registrar_error(exc, "(pegado)")
            self.app.avisar(f"No se pudo leer lo pegado: {exc}", ROJO)
            return

        # Se etiqueta con la hora para distinguir dos pegados del mismo banco
        # —el portal a veces entrega pesos y dólares en tablas aparte— y para
        # que el descarte de repetidos no los confunda con un archivo.
        etiqueta = "(pegado) {} · {}".format(
            nombre, datetime.datetime.now().strftime("%H:%M:%S"))
        self.archivos.append({
            "ruta": etiqueta, "clase": "banco", "banco": nombre,
            "cuentas": len(lineas), "tipo": "", "filas": 0, "error": "",
            "lineas": lineas, "datos": None, "secciones": []})
        self._reidentificar()
        self._ocupado(False)
        self._pintar()
        self.app.avisar(
            "Se leyeron {} cuenta(s) de {} desde lo pegado.".format(
                len(lineas), nombre),
            VERDE if lineas else NARANJA)

    async def _ingerir(self, rutas: list) -> None:
        """Lee y suma los archivos, vengan de donde vengan.

        Los tres caminos —elegirlos, tomar una carpeta o pegarlos— terminan aquí:
        así el descarte de repetidos, la lectura y el aviso final son idénticos y
        no hay tres versiones que mantener."""
        diagnostico.registrar("saldos._ingerir: rutas", " | ".join(rutas)[:400])
        # Un archivo que ya está cargado no se vuelve a leer: repetirlo solo
        # produciría cuentas duplicadas que después hay que descartar.  # noqa
        ya = {a["ruta"] for a in self.archivos} | {x["ruta"] for x in self.insumos}
        nuevas = [r for r in rutas if r not in ya]
        repetidos = len(rutas) - len(nuevas)
        if not nuevas:
            self.app.avisar("Esos archivos ya estaban cargados.", NARANJA)
            return

        self._ocupado(True, f"Leyendo {len(nuevas)} archivo(s)…")
        diagnostico.registrar("saldos._cargar: leyendo", "{}".format(len(nuevas)))
        try:
            leidos = await asyncio.to_thread(self._leer_en_hilo, nuevas)
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            diagnostico.registrar("saldos._cargar: falló la lectura", str(exc)[:200])
            self._ocupado(False)
            self.app.avisar(f"No se pudieron leer los archivos: {exc}", ROJO)
            return
        diagnostico.registrar("saldos._cargar: leídos", "{}".format(len(leidos)))

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
            diagnostico.registrar("saldos._cargar: esperando estado")
            await self._asegurar_estado()
            diagnostico.registrar("saldos._cargar: guardando insumos")
            await asyncio.to_thread(self._guardar_insumos, insumos)

        diagnostico.registrar("saldos._cargar: identificando")
        self._reidentificar()
        self._ocupado(False)
        diagnostico.registrar("saldos._cargar: pintando")
        self._pintar()
        self.app.avisar(*self._resumen_carga(bancarios, insumos, repetidos))
        diagnostico.registrar("saldos._cargar: listo")

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
            diagnostico.registrar("saldos: leyendo archivo",
                                  "{}/{} {}".format(i, total,
                                                    os.path.basename(ruta)))
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

        # Atajo por los NOMBRES de las pestañas, antes de tocar el contenido. El
        # libro de insumos son 2 MB y 32 000 filas, y el detector bancario lo
        # abría entero solo para mirarle las primeras quince y concluir que no
        # era de ningún banco: quince segundos tirados en cada carga. El índice
        # de hojas se lee sin abrir una sola fila.
        if saldos_insumos.es_libro_de_insumos(ruta):
            return self._como_insumo(registro, "")

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

        return self._como_insumo(registro, fallo_banco)

    @staticmethod
    def _como_insumo(registro: dict, fallo_banco: str) -> dict:
        """Intenta leerlo como insumo de flujo y deja el registro listo.

        `fallo_banco` es el motivo que dio el lector bancario, y es el que se
        reporta si tampoco es insumo: para un archivo que el usuario creía un
        reporte de banco, «no se reconoce de qué banco es» dice mucho más que
        «no parece ninguno de los insumos». Va vacío cuando ni se intentó."""
        ruta = registro["ruta"]
        try:
            tipo, datos = saldos_insumos.leer(ruta)
        except Exception as exc:  # noqa: BLE001 — tampoco es insumo
            registro["error"] = fallo_banco or str(exc)
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
    def _tablero_bancos(self) -> ft.Control:
        """Las quince pestañas del formato, con cuáles ya llegaron.

        Sustituye al aviso en prosa —«No llegó ningún saldo de BAJIO, BANCOMER,
        BANCOPPEL, …»—, que era trece nombres corridos a lo ancho de la pantalla.
        La pregunta de cada mañana es «¿cuál me falta?», y esa lista obligaba a
        compararla de memoria contra lo ya subido.

        Aquí están SIEMPRE las quince: verde la que está completa, ámbar la que
        llegó a medias y apagada la que no ha llegado. Es una lista de pendientes
        que se va tachando sola conforme se cargan archivos."""
        cobertura = self.asignacion.cobertura_por_hoja()
        return ft.Row([self._ficha_banco(h, *v) for h, v in cobertura.items()],
                      spacing=6, wrap=True, run_spacing=6)

    @staticmethod
    def _ficha_banco(hoja: str, llenos: int, total: int) -> ft.Control:
        """Una pestaña del formato y cómo va. Verde / ámbar / apagada."""
        if llenos >= total:
            color, icono = VERDE, ft.Icons.CHECK_CIRCLE
        elif llenos:
            color, icono = NARANJA, ft.Icons.INCOMPLETE_CIRCLE
        else:
            color, icono = GRIS, ft.Icons.CIRCLE_OUTLINED
        detalle = ("{}/{}".format(llenos, total) if 0 < llenos < total
                   else str(total))
        return ft.Container(
            content=ft.Row(
                [ft.Icon(icono, size=13, color=color),
                 ft.Text(hoja, size=11, color=color,
                         weight=ft.FontWeight.W_500 if llenos else None),
                 ft.Text(detalle, size=10, color=GRIS)],
                spacing=5, tight=True,
                vertical_alignment=ft.CrossAxisAlignment.CENTER),
            padding=ft.Padding.symmetric(vertical=4, horizontal=9),
            border_radius=20,
            bgcolor=ft.Colors.with_opacity(0.10 if llenos else 0.04, color),
            border=ft.Border.all(1, ft.Colors.with_opacity(
                0.35 if llenos else 0.15, color)))

    def _panel_cobertura(self) -> ft.Control:
        """La cifra que manda: cuántos renglones del formato quedaron llenos.

        Va grande y con barra de color porque es lo que el usuario tiene que
        mirar antes de generar. Los totales por divisa van al lado porque es lo
        segundo que revisa tesorería."""
        res = self.asignacion
        razon = res.llenos / res.total_renglones if res.total_renglones else 0
        # El color lo mandan los BANCOS que faltan, no el porcentaje de
        # renglones. Con el porcentaje, subir los quince reportes de uno en uno
        # dejaba el indicador en rojo durante todo el proceso normal: solo se
        # ponía verde al final, así que el 90% del tiempo gritaba alarma sin que
        # hubiera nada mal. Contado por bancos distingue «voy a la mitad» de
        # «terminé y hay un hueco», que es lo que de verdad importa.
        faltan = len(res.bancos_faltantes())
        color = VERDE if not faltan else NARANJA if faltan <= _BANCOS_POCOS else ROJO

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
                tile.controls = [self._encabezado_cuentas(compartida)]
                tile.controls += [self._fila_cuenta(c, compartida)
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
            # Los ceros van atenuados. En una pestaña típica la mayoría de las
            # cuentas están en cero, y con el mismo peso que los saldos vivos
            # obligan a leer las quince filas para encontrar las dos que hay que
            # cotejar contra el portal.
            ft.Text(f"{colocada.saldo:,.2f}", size=12, width=126,
                    text_align=ft.TextAlign.RIGHT,
                    color=GRIS if not colocada.saldo else None),
        ]
        return ft.Container(
            content=ft.Row(celdas, spacing=10,
                           vertical_alignment=ft.CrossAxisAlignment.CENTER),
            padding=ft.Padding.symmetric(vertical=4))

    @staticmethod
    def _encabezado_cuentas(con_banco: bool = False) -> ft.Control:
        """Los títulos de las columnas del desplegable.

        Sin ellos, `7680454 · ABASTECEDORA … · 131,151.64` no dice si el primer
        número es la cuenta, la sucursal o un folio."""
        def celda(texto, **kw):
            return ft.Text(texto, size=10, color=GRIS,
                           weight=ft.FontWeight.W_500, **kw)
        celdas = [celda("Cuenta", width=130),
                  celda("Titular", expand=True)]
        if con_banco:
            celdas.append(celda("Banco", width=90))
        celdas += [celda("", width=34),
                   celda("Saldo", width=126, text_align=ft.TextAlign.RIGHT)]
        return ft.Container(
            content=ft.Row(celdas, spacing=10),
            padding=ft.Padding.only(bottom=4),
            border=ft.Border(
                bottom=ft.BorderSide(1, ft.Colors.OUTLINE_VARIANT)))

    # ------------------------------------------------- insumos persistidos
    def _filas_secciones(self) -> list:
        """Una fila por sección, con lo que hay guardado y cómo borrarlo.

        Se listan LAS SEIS aunque estén vacías: así se ve de un vistazo qué falta
        capturar, no solo lo que ya está."""
        filas = []
        acciones = [self.btn_formato]
        # «Vaciar todo» solo aparece si hay algo que borrar: quien viene a
        # empezar de cero no tiene que ir borrando de uno en uno.
        if any(self.guardados.values()):
            acciones.append(ft.TextButton(
                content="Vaciar todo", icon=ft.Icons.DELETE_SWEEP_OUTLINED,
                style=ft.ButtonStyle(color=ROJO_BOTON),
                on_click=self._confirmar_vaciar_todo))
        # Las dos acciones de los insumos en un solo renglón y en extremos
        # opuestos: traer el formato a la izquierda, vaciar a la derecha. El
        # formato se descarga desde AQUÍ —donde se está pensando en los
        # insumos— y no desde la barra de acciones diarias, que es de todos los
        # días mientras que esto se hace una vez por semana.
        filas.insert(0, ft.Container(
            content=ft.Row(acciones,
                           alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            padding=ft.Padding.only(bottom=6)))
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
        """Vacía una sección. Va con espera porque REESCRIBE el libro entero.

        Quitarle una sección a los insumos obliga a volver a guardar las otras
        cinco —unas 32 000 filas—, y eso tarda varios segundos. Sin aviso, el
        botón parecía no hacer nada y el usuario volvía a pulsarlo."""
        nombre = _NOMBRES_INSUMO.get(seccion, seccion)
        fallo = None
        self._abrir_espera(
            "Vaciando {}…".format(nombre),
            "Se está reescribiendo el archivo de insumos con lo que queda. "
            "Puede tardar unos segundos.")
        try:
            # Se le pasa lo que ya está en memoria: releer el libro para quitarle
            # una sección costaría varios segundos más.
            await self._asegurar_estado()
            self.guardados = await asyncio.to_thread(
                saldos_estado.olvidar_insumos, seccion, self.guardados)
        except BaseException as exc:  # noqa: BLE001 — se reporta abajo, ya cerrado
            fallo = exc
        finally:
            # Antes de avisar nada: el snackbar comparte pila con el modal.
            self._cerrar_espera()
        if fallo is not None:
            self._registrar_error(fallo, saldos_estado.RUTA_INSUMOS)
            self.app.avisar(f"No se pudo vaciar {nombre}: {fallo}", ROJO)
            return
        self._pintar()
        self.app.avisar("Se vació {}.".format(nombre), NARANJA)

    def _confirmar_vaciar_todo(self, _e=None) -> None:
        """Pregunta antes de borrarlo todo. No hay deshacer.

        Recapturar MGC o TESORO significa volver a exportar y subir decenas de
        miles de filas, así que un clic de más no puede costar eso."""
        vivas = [_NOMBRES_INSUMO.get(k, k)
                 for k in saldos_insumos.SECCIONES if self.guardados.get(k)]
        if not vivas:
            self.app.avisar("No hay insumos guardados que vaciar.", GRIS)
            return

        def confirmar(_ev=None):
            self.page.pop_dialog()
            self.page.run_task(self._vaciar_todo)

        self.page.show_dialog(ft.AlertDialog(
            modal=True,
            title=ft.Text("Vaciar todos los insumos"),
            content=ft.Text(
                "Se borrará lo capturado en {}: {}. No se puede deshacer; para "
                "recuperarlo habría que volver a subir los archivos.".format(
                    "{} secciones".format(len(vivas)) if len(vivas) > 1
                    else "1 sección",
                    ", ".join(vivas))),
            actions=[
                ft.TextButton("Cancelar",
                              on_click=lambda e: self.page.pop_dialog()),
                ft.FilledButton("Vaciar todo", on_click=confirmar,
                                color=ft.Colors.WHITE, bgcolor=ROJO),
            ],
            actions_alignment=ft.MainAxisAlignment.END))

    async def _vaciar_todo(self) -> None:
        """Borra el archivo de insumos completo."""
        fallo = None
        self._abrir_espera(
            "Vaciando los insumos…",
            "Se está borrando lo capturado de todas las secciones.")
        try:
            self.guardados = await asyncio.to_thread(
                saldos_estado.olvidar_insumos)
            # Ya no hay nada que leer del disco, así que la carga diferida queda
            # resuelta: sin esto, el próximo `_asegurar_estado` volvería a leer
            # un archivo que acabamos de borrar.
            self._estado_cargado = True
            self._estado_listo.set()
        except BaseException as exc:  # noqa: BLE001 — se reporta abajo, ya cerrado
            fallo = exc
        finally:
            self._cerrar_espera()
        if fallo is not None:
            self._registrar_error(fallo, saldos_estado.RUTA_INSUMOS)
            self.app.avisar(f"No se pudieron vaciar los insumos: {fallo}", ROJO)
            return
        self._pintar()
        self.app.avisar("Se vaciaron todos los insumos guardados.", NARANJA)

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

        # Qué pestañas faltan ya lo dice el tablero de fichas, con los quince
        # bancos a la vista; repetirlo aquí en prosa era la peor versión de la
        # misma información. Solo se conserva el aviso de los renglones sueltos,
        # que el tablero NO cubre: una pestaña puede haber llegado y aun así
        # dejar cuentas vacías dentro.
        if res.vacios and not res.bancos_faltantes():
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
        self.separador.visible = hay
        self.tab_archivos.label = f"Archivos cargados ({len(cargados)})"
        # Basta con un archivo bancario leído: los insumos nunca son requisito.
        self.btn_generar.disabled = not (self.asignacion
                                         and self.asignacion.colocadas)

        if self.asignacion is not None:
            self.hero.content = self._panel_cobertura()
            self.hero.visible = True
            self.tablero.content = self._tablero_bancos()
            self.tablero.visible = True
            self.avisos.controls = self._lista_avisos()
            self.avisos.visible = bool(self.avisos.controls)
            self.bancos.controls = self._grupos_por_banco()
            self.tab_saldos.label = "Saldos identificados ({})".format(
                self.asignacion.pegadas)
        else:
            self.hero.visible = False
            self.hero.content = None
            self.tablero.visible = False
            self.tablero.content = None
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
        # Las tres pestañas cuentan ELEMENTOS. El «5 de 6» de antes mezclaba una
        # métrica de completitud con dos de conteo, y sobre todo se leía como si
        # faltara algo urgente cuando cinco es lo normal: impuestos no se usa.
        vivos = sum(1 for x in self.guardados.values() if x)
        self.tab_insumos.label = "Insumos de flujo ({})".format(vivos)
        self._refrescar(self.lista, self.vacio, self.pestanas,
                        self.btn_limpiar, self.separador, self.btn_generar,
                        self.hero,
                        self.avisos, self.bancos, self.secciones,
                        self.cargando_insumos, self.tablero)

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
        # Se bloquean las TRES entradas, no solo el botón: si no, pegar con
        # Ctrl+V a media lectura arrancaría otra en paralelo sobre los mismos
        # archivos. `_leyendo` es lo que consulta el atajo de teclado.
        self.btn_cargar.disabled = activo
        self.btn_carpeta.disabled = activo
        self.txt_estado.value = mensaje
        self._refrescar(self.anillo, self.barra, self.btn_cargar,
                        self.btn_carpeta, self.txt_estado)

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

    @staticmethod
    def _capturas_de_la_semana() -> dict:
        """Lo que tesorería tecleó a mano en el reporte, para volver a ponerlo.

        Los pagos y los importes de ACP e IMPUESTOS no salen de ningún archivo:
        se escriben en Excel sobre el reporte ya generado. Como cada corrida
        parte del libro base, salían en blanco y había que recapturarlos.

        Se leen del ÚLTIMO reporte generado, que es donde están los más nuevos, y
        se guardan aparte por si ese archivo se mueve o se borra. Lo que traiga
        el archivo gana sobre lo guardado: es lo que el usuario acaba de escribir.

        Todo es best-effort. Si algo falla se genera sin las capturas —que es
        exactamente lo que pasaba antes— en vez de no generar."""
        lunes = saldos_export._lunes_de(datetime.datetime.now())
        try:
            guardadas = saldos_estado.manuales_semana(lunes)
        except Exception:  # noqa: BLE001 — se sigue sin lo guardado
            guardadas = {}
        try:
            del_archivo = saldos_export.leer_manuales(
                saldos_estado.ultimo_reporte())
        except Exception:  # noqa: BLE001 — se sigue sin releer el anterior
            del_archivo = {}
        guardadas.update(del_archivo)
        diagnostico.registrar(
            "saldos: capturas de la semana",
            "{} celda(s), semana del {}".format(len(guardadas), lunes.date()))
        if guardadas:
            try:
                saldos_estado.guardar_manuales(lunes, guardadas)
            except Exception:  # noqa: BLE001 — guardar no es crítico
                pass
        return guardadas

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
        # El botón se apaga desde el PRIMER clic y no se vuelve a encender hasta
        # el final. El modal de espera no alcanza: entre pulsar y verlo aparecer
        # hay que pasar por el diálogo de guardado y por la lectura de lo
        # capturado, y ahí el botón seguía vivo. `disabled` no se pierde con los
        # repintados intermedios porque `_pintar` lo recalcula igual.
        self.btn_generar.disabled = True
        self._refrescar(self.btn_generar)
        try:
            await self._generar_ya()
        finally:
            self.btn_generar.disabled = not (self.asignacion
                                             and self.asignacion.colocadas)
            self._refrescar(self.btn_generar)

    async def _generar_ya(self) -> None:
        await self._asegurar_estado()
        hoy = datetime.date.today().strftime("%d-%m-%Y")
        ruta = await self.app.picker.save_file(
            dialog_title="Guardar el reporte de saldos",
            file_name=f"SALDOS {hoy}.xlsx", allowed_extensions=["xlsx"])
        if not ruta:
            return
        if not ruta.lower().endswith(".xlsx"):
            ruta += ".xlsx"
        # La espera se abre AQUÍ, en cuanto hay ruta, y no justo antes de
        # escribir: entre medias se relee lo capturado del reporte anterior, y
        # eso tarda lo suyo. Con el modal después, la ventana se quedaba quieta
        # sin explicación y el usuario volvía a pulsar «Generar».
        #
        # No puede abrirse antes del diálogo de guardado: ese es una ventana del
        # sistema y quedaría por detrás del modal.
        #
        # NO se avisa nada mientras el modal está abierto. `app.avisar` muestra el
        # snackbar con `page.show_dialog`, o sea LA MISMA PILA que el modal: si se
        # avisa antes de cerrarlo, el `pop_dialog` de después se lleva el snackbar
        # y el modal se queda para siempre. Era justo el cuelgue que solo aparecía
        # cuando la generación fallaba.
        info = fallo = None
        self._abrir_espera()
        try:
            # Lo guardado manda como base y lo de esta sesión ya está fundido
            # ahí, así que se usa `self.guardados`: es lo mismo que el usuario
            # descargaría.
            insumos = dict(self.guardados)
            anterior = saldos_estado.totales_dia_anterior()
            manuales = await asyncio.to_thread(self._capturas_de_la_semana)
            info = await asyncio.to_thread(
                saldos_export.generar, ruta, self.asignacion, insumos,
                None, anterior, manuales)
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
        # Se recuerda ESTE reporte: mañana es de donde salen los pagos y los
        # importes que se escriban hoy en él.
        saldos_estado.guardar_ultimo_reporte(ruta)

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
