"""Pantalla: Saldos.

Sustituye el trabajo manual de pegar, hoja por hoja, los reportes de cada portal
bancario en `FORMATO DE SALDOS TESORERIA.xlsx`. El usuario sube los archivos tal
como los descarga —en cualquier orden y con cualquier nombre— y la herramienta los
lee, identifica de qué empresa es cada cuenta y genera el Excel listo para
imprimir.

El flujo tiene tres pasos y un solo panel:

  1. **Cargar reportes** — se leen en un hilo (ver `_leer_en_hilo`) y se muestra el
     avance archivo por archivo.
  2. **Revisar** — la lista dice qué banco se detectó en cada archivo y cuántas
     cuentas trajo; abajo, cuántas se pudieron atribuir a una empresa.
  3. **Generar el reporte** — pide dónde guardarlo y lo abre.

Nota de rendimiento: aquí NO se renderiza una tabla de cuentas. Con 20 archivos y
más de 200 saldos, pintar una tabla completa costaría cientos de widgets y el
render se notaría; la lista de archivos son una docena de filas y el detalle vive
en el Excel. Es una decisión deliberada, no una simplificación.
"""

from __future__ import annotations

import asyncio
import datetime
import os

import flet as ft

from core import saldos as motor
from core import cuentas_dispersion, saldos_export, saldos_lectores
from ui.comun import CENTRO, EMPRESAS, GRIS, NARANJA, ROJO, VERDE, tarjeta

# Extensiones que emiten los portales bancarios (ver core/saldos_lectores).
_EXTENSIONES = ["xlsx", "xls", "csv", "txt", "pdf"]


class SeccionSaldos:
    """Pantalla de generación del reporte diario de saldos bancarios."""

    def __init__(self, app):
        self.app = app
        self.page = app.page
        # Archivos cargados: [{"ruta", "banco", "cuentas", "error"}]
        self.archivos: list[dict] = []
        # Líneas leídas de todos los archivos y el resultado de identificarlas.
        self.lineas: list[saldos_lectores.LineaSaldo] = []
        self.resultado: motor.Resultado | None = None
        self.catalogo = cuentas_dispersion.CatalogoCuentasDispersion()
        self.contenido = self._construir()

    # ------------------------------------------------------------------ UI
    def _construir(self) -> ft.Control:
        self.btn_cargar = ft.FilledButton(
            content="Cargar reportes", icon=ft.Icons.UPLOAD_FILE,
            on_click=self._cargar)
        self.btn_limpiar = ft.OutlinedButton(
            content="Quitar todo", icon=ft.Icons.DELETE_SWEEP_OUTLINED,
            visible=False, on_click=self._limpiar)
        self.btn_generar = ft.FilledButton(
            content="Generar reporte", icon=ft.Icons.TABLE_VIEW,
            disabled=True, on_click=self._generar)

        self.txt_ayuda = ft.Text(
            "Sube los reportes de saldos tal como los descargas de cada portal "
            "(.xlsx, .xls, .csv, .txt o .pdf). El banco se detecta por el "
            "contenido, así que el nombre del archivo da igual.",
            size=12, color=GRIS)
        self.txt_estado = ft.Text("", size=13, color=GRIS)
        self.anillo = ft.ProgressRing(width=18, height=18, stroke_width=2,
                                      visible=False)
        self.barra = ft.ProgressBar(width=360, value=0, visible=False)

        # Lista de archivos cargados. Es un Column simple, no una TablaResponsiva:
        # son pocas filas y no necesita columnas proporcionales ni scroll propio.
        self.lista = ft.Column(spacing=0, tight=True)
        self.vacio = ft.Container(
            content=ft.Text("Todavía no has cargado ningún reporte.",
                            size=12, color=GRIS, italic=True),
            alignment=CENTRO, padding=ft.Padding.symmetric(vertical=18))

        self.resumen = ft.Container(visible=False)

        panel = ft.Column(
            [
                self.txt_ayuda,
                ft.Row([self.btn_cargar, self.btn_limpiar, self.anillo,
                        self.txt_estado],
                       spacing=10, wrap=True,
                       vertical_alignment=ft.CrossAxisAlignment.CENTER),
                self.barra,
                ft.Divider(height=8),
                self.lista,
                self.vacio,
                self.resumen,
                ft.Row([self.btn_generar], spacing=10),
            ],
            spacing=12, horizontal_alignment=ft.CrossAxisAlignment.STRETCH)

        # Estado inicial coherente: sin archivos, la lista se oculta y se ve el
        # mensaje de vacío. Se fija aquí y no en el constructor de cada control
        # para tener una sola fuente de verdad (_pintar).
        self.lista.visible = False
        return ft.Column(
            [tarjeta("Reporte de saldos bancarios", panel)],
            spacing=14, scroll=ft.ScrollMode.AUTO, expand=True)

    # ------------------------------------------------------- carga de archivos
    async def _cargar(self, _e=None) -> None:
        archivos = await self.app.picker.pick_files(
            dialog_title="Selecciona los reportes de saldos de los bancos",
            allowed_extensions=_EXTENSIONES, allow_multiple=True)
        if not archivos:
            return
        rutas = [a.path for a in archivos]
        # Un archivo que ya está cargado no se vuelve a leer: repetirlo solo
        # produciría cuentas duplicadas que después hay que descartar.
        nuevas = [r for r in rutas
                  if r not in {a["ruta"] for a in self.archivos}]
        repetidos = len(rutas) - len(nuevas)
        if not nuevas:
            self.app.avisar("Esos reportes ya estaban cargados.", NARANJA)
            return

        self._ocupado(True, f"Leyendo {len(nuevas)} archivo(s)…")
        try:
            leidos = await asyncio.to_thread(self._leer_en_hilo, nuevas)
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self._ocupado(False)
            self.app.avisar(f"No se pudieron leer los reportes: {exc}", ROJO)
            return
        self.archivos.extend(leidos)
        self._reidentificar()
        self._ocupado(False)
        self._pintar()

        con_error = [a for a in leidos if a["error"]]
        mensaje = f"{len(leidos) - len(con_error)} reporte(s) leído(s)."
        if repetidos:
            mensaje += f" {repetidos} ya estaban cargados y se omitieron."
        if con_error:
            mensaje += f" {len(con_error)} no se pudieron leer."
        self.app.avisar(mensaje, NARANJA if con_error else VERDE)

    def _leer_en_hilo(self, rutas: list[str]) -> list[dict]:
        """Lee los reportes. Corre en un hilo (lo llama `_cargar` con
        `asyncio.to_thread`) porque abre archivos y parsea PDF: en el hilo de la
        interfaz congelaría la ventana varios segundos.

        Un archivo que falle no aborta el resto — se marca y se sigue, que es lo
        que hace falta cuando se cargan doce de golpe."""
        out: list[dict] = []
        total = len(rutas)
        for i, ruta in enumerate(rutas, 1):
            registro = {"ruta": ruta, "banco": "", "cuentas": 0, "error": "",
                        "lineas": []}
            try:
                lineas, banco = saldos_lectores.leer(ruta)
                registro["banco"] = banco
                registro["cuentas"] = len(lineas)
                registro["lineas"] = lineas
            except saldos_lectores.ErrorLector as exc:
                registro["error"] = str(exc)
            except Exception as exc:  # noqa: BLE001 — un archivo raro no aborta
                registro["error"] = f"{type(exc).__name__}: {exc}"
            out.append(registro)
            self._progreso(i, total)
        return out

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
        self.barra.value = hechos / total if total else 0
        self.txt_estado.value = f"Leyendo reportes… {hechos} de {total}"
        self._refrescar(self.barra, self.txt_estado)

    def _limpiar(self, _e=None) -> None:
        self.archivos.clear()
        self.lineas.clear()
        self.resultado = None
        self._pintar()

    # ------------------------------------------------------- identificación
    def _reidentificar(self) -> None:
        """Recalcula la atribución con TODOS los archivos cargados.

        Se rehace completa en vez de ir sumando por archivo porque la detección de
        cuentas repetidas necesita ver el conjunto: la misma cuenta puede venir en
        dos descargas distintas (a BBVA se le bajan cinco archivos)."""
        self.lineas = [x for a in self.archivos for x in a["lineas"]]
        self.resultado = (motor.identificar(self.lineas, self.catalogo)
                          if self.lineas else None)

    # ---------------------------------------------------------------- pintado
    def _fila_archivo(self, a: dict) -> ft.Control:
        if a["error"]:
            icono = ft.Icon(ft.Icons.ERROR_OUTLINE, size=18, color=ROJO)
            detalle = ft.Text(a["error"], size=11, color=ROJO, expand=True,
                              max_lines=2, overflow=ft.TextOverflow.ELLIPSIS)
        else:
            icono = ft.Icon(ft.Icons.CHECK_CIRCLE_OUTLINE, size=18, color=VERDE)
            detalle = ft.Text(f"{a['banco']} · {a['cuentas']} cuenta(s)",
                              size=12, color=GRIS, expand=True)
        return ft.Container(
            content=ft.Row(
                [icono,
                 ft.Text(os.path.basename(a["ruta"]), size=12, width=320,
                         max_lines=1, no_wrap=True,
                         overflow=ft.TextOverflow.ELLIPSIS,
                         tooltip=a["ruta"]),
                 detalle],
                spacing=10, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            padding=ft.Padding.symmetric(vertical=6, horizontal=4),
            border=ft.Border(bottom=ft.BorderSide(1, ft.Colors.OUTLINE_VARIANT)))

    def _panel_resumen(self) -> ft.Control:
        """Qué se identificó y qué no. Se muestra ANTES de generar para que el
        usuario decida si le falta subir algún reporte o completar el catálogo."""
        res = self.resultado
        sin = len(res.sin_identificar)
        dup = len(res.duplicados)
        monto_sin = sum(x.linea.saldo for x in res.sin_identificar)
        filas = [
            ft.Row([ft.Icon(ft.Icons.ACCOUNT_BALANCE_WALLET, size=18, color=VERDE),
                    ft.Text(f"{len(res.identificados)} cuenta(s) identificada(s) "
                            f"en {len(res.por_empresa())} empresa(s)", size=13)],
                   spacing=8, vertical_alignment=ft.CrossAxisAlignment.CENTER),
        ]
        if sin:
            filas.append(ft.Row(
                [ft.Icon(ft.Icons.HELP_OUTLINE, size=18, color=NARANJA),
                 ft.Text(f"{sin} sin identificar por ${monto_sin:,.2f} — van en la "
                         f"hoja «Sin identificar» del reporte. Complétalas en el "
                         f"catálogo de cuentas para que aparezcan.",
                         size=12, color=NARANJA, expand=True)],
                spacing=8, vertical_alignment=ft.CrossAxisAlignment.START))
        if dup:
            filas.append(ft.Row(
                [ft.Icon(ft.Icons.CONTENT_COPY, size=18, color=NARANJA),
                 ft.Text(f"{dup} cuenta(s) venían repetidas en más de un archivo; "
                         f"se contaron una sola vez.", size=12, color=NARANJA,
                         expand=True)],
                spacing=8, vertical_alignment=ft.CrossAxisAlignment.START))
        return ft.Container(
            content=ft.Column(filas, spacing=8, tight=True),
            padding=12, border_radius=8,
            bgcolor=ft.Colors.SURFACE_CONTAINER_HIGHEST)

    def _pintar(self) -> None:
        hay = bool(self.archivos)
        self.lista.controls = [self._fila_archivo(a) for a in self.archivos]
        self.lista.visible = hay
        self.vacio.visible = not hay
        self.btn_limpiar.visible = hay
        self.btn_generar.disabled = not (self.resultado
                                         and self.resultado.identificados)
        if self.resultado is not None:
            self.resumen.content = self._panel_resumen()
            self.resumen.visible = True
        else:
            self.resumen.visible = False
            self.resumen.content = None
        self._refrescar(self.lista, self.vacio, self.btn_limpiar,
                        self.btn_generar, self.resumen)

    def _ocupado(self, activo: bool, mensaje: str = "") -> None:
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
    async def _generar(self, _e=None) -> None:
        if not (self.resultado and self.resultado.identificados):
            self.app.avisar("No hay saldos identificados que reportar.", NARANJA)
            return
        hoy = datetime.date.today().strftime("%d-%m-%Y")
        ruta = await self.app.picker.save_file(
            dialog_title="Guardar el reporte de saldos",
            file_name=f"SALDOS {hoy}.xlsx", allowed_extensions=["xlsx"])
        if not ruta:
            return
        if not ruta.lower().endswith(".xlsx"):
            ruta += ".xlsx"
        nombres = {e["id"]: e["Empresa"] for e in EMPRESAS}
        try:
            info = await asyncio.to_thread(
                saldos_export.generar, ruta, self.resultado, nombres)
        except PermissionError:
            self.app.avisar(
                "No se pudo guardar: el archivo está abierto en Excel. Ciérralo e "
                "intenta de nuevo (o guarda con otro nombre).", ROJO)
            return
        except Exception as exc:  # noqa: BLE001 — se reporta al usuario
            self.app.avisar(f"No se pudo generar el reporte: {exc}", ROJO)
            return
        detalle = (f"{info['cuentas']} cuenta(s) de {info['empresas']} empresa(s)")
        if info["sin_identificar"]:
            detalle += f"; {info['sin_identificar']} en la hoja «Sin identificar»"
        self.app.avisar(
            f"Reporte de saldos generado: {detalle}.", VERDE, accion="Abrir",
            on_accion=lambda _e=None: self.app.abrir_en_sistema(ruta),
            duracion=ft.Duration(seconds=12))

    # --------------------------------------------------------------- catálogo
    def recargar_catalogo(self) -> None:
        """Relee el catálogo de cuentas tras adjuntar uno nuevo en Configuración y
        vuelve a identificar lo que ya estaba cargado: es justo el momento en que el
        usuario acaba de agregar las cuentas que faltaban."""
        self.catalogo = cuentas_dispersion.CatalogoCuentasDispersion()
        if self.archivos:
            self._reidentificar()
            self._pintar()
