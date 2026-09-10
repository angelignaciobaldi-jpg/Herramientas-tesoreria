"""Lo que el reporte de saldos recuerda de un día para otro.

Los saldos bancarios se bajan a diario, pero el resto del formato no: las
facturas de MGC, Pemex y tesoro se capturan por semana, los movimientos de
créditos cada tanto, y la nómina y los impuestos según toque. Pedirle a la
usuaria que vuelva a subir todo eso cada mañana sería trabajo inventado.

Aquí se guardan dos cosas, en la carpeta de datos de la app:

  saldos_insumos.xlsx   las seis secciones, una por pestaña. Es el MISMO archivo
                        que se ofrece para descargar y volver a subir: no hay una
                        copia interna distinta de la que ve el usuario.
  saldos_totales.json   los totales de la cabecera de cada corrida, por fecha.

## Por qué el histórico va por fecha

El reporte compara contra el DÍA HÁBIL ANTERIOR, y esa comparación no se puede
atar a «la corrida pasada»: regenerar el reporte un martes en la tarde dejaría la
comparativa contra el mismo martes, o sea en ceros. Guardando por fecha y
tomando siempre el registro más reciente ANTERIOR a hoy, regenerar no mueve nada
y un lunes después de puente toma el viernes sin que nadie configure nada.
"""

from __future__ import annotations

import datetime
import json
import os

from . import rutas

RUTA_INSUMOS = os.path.join(rutas.DATOS, "saldos_insumos.xlsx")
# Lo que tesorería teclea en el calendario de flujo del reporte (pagos e
# importes). Va aparte de los totales porque tiene otra vida: se conserva por
# SEMANA, no por día, y no se compara contra nada — solo se repone.
RUTA_SEMANA = os.path.join(rutas.DATOS, "saldos_semana.json")
RUTA_TOTALES = os.path.join(rutas.DATOS, "saldos_totales.json")

# Cuántas corridas se conservan. No hace falta un archivo histórico: lo único que
# se consulta es el día anterior, y unas semanas alcanzan de sobra para cubrir
# puentes y vacaciones.
_MAX_HISTORICO = 60


def _hoy() -> str:
    return datetime.date.today().isoformat()


# ---------------------------------------------------------------- totales

def _leer_json(ruta: str) -> dict:
    """Un diccionario guardado en JSON, o vacío si no está o está corrupto."""
    if not os.path.exists(ruta):
        return {}
    try:
        with open(ruta, encoding="utf-8") as f:
            datos = json.load(f)
        return datos if isinstance(datos, dict) else {}
    except Exception:  # noqa: BLE001 — un archivo corrupto no debe tumbar nada
        return {}


def _escribir_json(ruta: str, datos: dict) -> None:
    """Guarda un diccionario. Best-effort: no poder recordar no impide reportar."""
    try:
        with open(ruta, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=1, sort_keys=True)
    except Exception:  # noqa: BLE001 — no poder recordar no impide reportar hoy
        pass


def _leer_totales() -> dict:
    return _leer_json(RUTA_TOTALES)


def guardar_totales(totales: dict, fecha: datetime.datetime = None) -> None:
    """Registra los totales de la cabecera de esta corrida, bajo su fecha.

    Si ya había una corrida del mismo día se reemplaza: la buena es la última,
    porque puede haberse regenerado con más archivos."""
    fecha = fecha or datetime.datetime.now()
    historico = _leer_totales()
    historico[fecha.date().isoformat()] = {
        "hora": fecha.strftime("%H:%M:%S"),
        "totales": {k: float(v) for k, v in (totales or {}).items()},
    }
    # Se poda por fecha, no por orden de escritura.
    for sobra in sorted(historico)[:-_MAX_HISTORICO]:
        historico.pop(sobra, None)
    _escribir_json(RUTA_TOTALES, historico)


def totales_dia_anterior(fecha: datetime.datetime = None) -> tuple:
    """Los totales de la última corrida ANTERIOR a `fecha`.

    Devuelve `(fecha_iso, hora, {celda: valor})`, o `(None, "", {})` si es la
    primera vez. Estrictamente anterior: regenerar hoy no se compara consigo
    mismo."""
    fecha = fecha or datetime.datetime.now()
    tope = fecha.date().isoformat()
    historico = _leer_totales()
    previas = [d for d in sorted(historico) if d < tope]
    if not previas:
        return None, "", {}
    dia = previas[-1]
    registro = historico[dia] or {}
    return dia, registro.get("hora", ""), registro.get("totales", {})


# ------------------------------------------------- capturas de la semana

def _clave_semana(lunes) -> str:
    return lunes.strftime("%Y-%m-%d") if hasattr(lunes, "strftime") else str(lunes)


def manuales_semana(lunes) -> dict:
    """Lo capturado a mano en el reporte de ESTA semana, {celda: valor}.

    Se guarda por semana porque es lo que dura: el lunes se empieza de cero, y
    conservar lo de la semana pasada metería pagos viejos en el calendario
    nuevo."""
    datos = _leer_json(RUTA_SEMANA)
    return dict(datos.get(_clave_semana(lunes)) or {})


def guardar_manuales(lunes, celdas: dict) -> None:
    """Guarda lo capturado, y deja SOLO la semana en curso.

    No se acumula histórico: estas celdas no se comparan contra nada, solo se
    reponen. Quedarse con las anteriores sería basura que crece sola."""
    if not celdas:
        return
    _escribir_json(RUTA_SEMANA, {_clave_semana(lunes): dict(celdas)})


def ultimo_reporte() -> str:
    """Ruta del último reporte generado, para poder releer lo capturado en él."""
    from . import preferencias
    return str(preferencias.cargar_valor("saldos_ultimo_reporte", "") or "")


def guardar_ultimo_reporte(ruta: str) -> None:
    from . import preferencias
    try:
        preferencias.guardar_valor("saldos_ultimo_reporte", str(ruta or ""))
    except Exception:  # noqa: BLE001 — recordarlo no es crítico
        pass


def olvidar_totales() -> None:
    """Borra el histórico de comparativas."""
    try:
        os.remove(RUTA_TOTALES)
    except OSError:
        pass


# ---------------------------------------------------------------- insumos

def hay_insumos() -> bool:
    return os.path.exists(RUTA_INSUMOS)


class ErrorEstado(Exception):
    """No se pudo leer o escribir lo que la app tenía guardado."""


def cargar_insumos() -> dict:
    """Los insumos guardados, {sección: datos}. Vacío si no hay nada.

    Un archivo ILEGIBLE no se trata como «no hay nada»: se aparta con otro
    nombre y se avisa. Tragárselo dejaba al usuario viendo todas las secciones
    «sin capturar», sin forma de saber si nunca las subió o si se le rompió el
    archivo —y con el siguiente guardado encima, borrando la evidencia—."""
    if not hay_insumos():
        return {}
    from . import saldos_insumos
    try:
        tipo, datos = saldos_insumos.leer(RUTA_INSUMOS)
    except Exception as exc:  # noqa: BLE001 — se traduce a un error propio
        apartado = _apartar(RUTA_INSUMOS)
        raise ErrorEstado(
            "el archivo de insumos guardado no se pudo leer ({}). Se apartó "
            "como «{}» y se empieza de cero; vuelve a subir tus archivos.".format(
                exc, os.path.basename(apartado) if apartado else "?")) from exc
    return datos if tipo == "COMBINADO" else {tipo: datos}


def _apartar(ruta: str) -> str:
    """Renombra un archivo ilegible en vez de dejarlo estorbando. Devuelve la
    ruta nueva, o "" si tampoco se pudo mover."""
    destino = "{}.dañado{}".format(*os.path.splitext(ruta))
    try:
        os.replace(ruta, destino)
        return destino
    except OSError:
        return ""


def guardar_insumos(datos: dict) -> dict:
    """Reescribe el archivo de insumos con lo que se le pase.

    Es el MISMO archivo que se ofrece para descargar: no hay una copia interna
    distinta de la que ve el usuario, así que lo que descargue es exactamente lo
    que la app va a usar mañana.

    Se escribe a un TEMPORAL y se reemplaza al final, nunca encima del bueno.
    Escribir en sitio es lo que destruyó un archivo real: el usuario pulsó
    «vaciar» varias veces porque el botón no daba señales, y dos rewrites del
    mismo libro de 1.6 MB se entrelazaron y lo dejaron ilegible. Con `os.replace`
    —atómico dentro del mismo volumen— lo peor que puede pasar es que quede el
    archivo anterior intacto."""
    from . import saldos_insumos
    temporal = RUTA_INSUMOS + ".tmp"
    try:
        escritas = saldos_insumos.escribir_plantilla(temporal, datos)
        os.replace(temporal, RUTA_INSUMOS)
        return escritas
    except Exception:  # noqa: BLE001 — no poder recordar no impide reportar hoy
        try:
            os.remove(temporal)
        except OSError:
            pass
        return {}


def fusionar_insumos(guardados: dict, nuevos: dict) -> dict:
    """Combina lo que ya había con lo que se acaba de subir.

    Una sección solo se toca si viene en `nuevos`: subir el archivo de nómina no
    puede borrar las facturas de Pemex capturadas la semana pasada. Y una sección
    que llega VACÍA sí borra: es la forma de limpiarla desde el Excel."""
    salida = dict(guardados or {})
    for seccion, datos in (nuevos or {}).items():
        if _vacia(datos):
            salida.pop(seccion, None)
        else:
            salida[seccion] = datos
    return salida


def _vacia(datos) -> bool:
    if not datos:
        return True
    if isinstance(datos, list):
        return len(datos) == 0
    rangos = datos.get("rangos") if isinstance(datos, dict) else None
    if rangos is None:
        return False
    return not any(any(v is not None for v in fila)
                   for r in rangos for fila in r["celdas"])


def olvidar_insumos(seccion: str = None, actuales: dict = None) -> dict:
    """Borra una sección (o todas si no se indica). Devuelve lo que quedó.

    `actuales` evita releer el libro cuando quien llama ya lo tiene en memoria:
    releerlo cuesta varios segundos y aquí bloquearía la interfaz."""
    if seccion is None:
        try:
            os.remove(RUTA_INSUMOS)
        except OSError:
            pass
        return {}
    quedan = dict(actuales) if actuales is not None else cargar_insumos()
    quedan.pop(seccion, None)
    guardar_insumos(quedan)
    return quedan


def resumen_insumos() -> dict:
    """Qué secciones hay guardadas: {seccion: (filas, fecha de captura)}.

    Se lee del propio archivo para que lo que se muestra en pantalla sea lo que
    de verdad se va a usar, y no un contador aparte que se puede desincronizar."""
    if not hay_insumos():
        return {}
    import openpyxl
    try:
        libro = openpyxl.load_workbook(RUTA_INSUMOS, read_only=True)
    except Exception:  # noqa: BLE001 — archivo dañado: se trata como vacío
        return {}
    try:
        modificado = datetime.datetime.fromtimestamp(
            os.path.getmtime(RUTA_INSUMOS))
        return {
            hoja.title: (max(hoja.max_row - 1, 0), modificado)
            for hoja in libro.worksheets
        }
    finally:
        libro.close()
