"""Extracción de texto de estados de cuenta.

Estrategia híbrida (decidida con el usuario):
  - PDF con texto seleccionable  -> se lee el texto directo (rápido y exacto).
  - PDF escaneado / página sin texto -> se rasteriza la página y se aplica OCR.
  - Imagen (JPG/PNG/TIFF)         -> OCR directo.

Dependencias: PyMuPDF (fitz) para leer/rasterizar PDF, Pillow + pytesseract
para el OCR. El motor Tesseract es local (offline): los documentos NO salen
del equipo.
"""

from __future__ import annotations

import io
import os
import re
import shutil
import subprocess
import tempfile
import time

import fitz  # PyMuPDF
import pytesseract
from PIL import Image

from . import rutas

# En Windows, evita que el tesseract.exe que lanzamos abra una ventana de consola.
_CREATE_NO_WINDOW = 0x08000000 if os.name == "nt" else 0

# Caracteres "normales" esperables en un documento en español. Si la capa de
# texto de un PDF trae muy pocos de estos (codificación de fuente rota, p. ej.
# texto que sale como "ÕÖÔÉÕÁ@ÆÓÅç..."), se considera ilegible y se usa OCR.
_RE_CHARS_NORMALES = re.compile(
    r"[0-9A-Za-zÁÉÍÓÚÜÑáéíóúüñ \t\r\n.,;:%$()/\-#°ºª'\"¡!¿?@&+*]"
)
_UMBRAL_TEXTO_CONFIABLE = 0.70

# Espacios Unicode (no-rompible, etc.) que algunos PDF usan en lugar del espacio
# normal; se normalizan para que la extracción de datos funcione.
_ESPACIOS_RAROS = {"\xa0": " ", " ": " ", " ": " ", " ": " ", "​": ""}

# Los PDF locales son de confianza: se desactiva el límite anti-"bomba" de Pillow
# para permitir páginas escaneadas a muy alta resolución.
Image.MAX_IMAGE_PIXELS = None


def _normalizar(texto: str) -> str:
    for raro, normal in _ESPACIOS_RAROS.items():
        texto = texto.replace(raro, normal)
    return texto

# --- Localización del binario de Tesseract -------------------------------
# 1º el Tesseract EMPAQUETADO junto a la app ({app}\Tesseract-OCR): asi funciona
# en una maquina limpia sin instalarlo aparte. Luego, ubicaciones habituales de
# una instalacion del sistema (util en desarrollo). Fuera del PATH -> se apunta.
_RUTAS_TESSERACT = [
    os.path.join(rutas.INSTALL, "Tesseract-OCR", "tesseract.exe"),
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    os.path.join(os.environ.get("LOCALAPPDATA", ""), r"Programs\Tesseract-OCR\tesseract.exe"),
]
for _ruta in _RUTAS_TESSERACT:
    if _ruta and os.path.exists(_ruta):
        pytesseract.pytesseract.tesseract_cmd = _ruta
        break

# Carpeta de modelos de idioma (spa+eng+osd). Se prioriza la del proyecto y, si
# no, la del Tesseract empaquetado; asi el OCR usa español sin tocar el sistema.
for _tessdata in (
    os.path.join(rutas.BUNDLE, "tessdata"),
    os.path.join(rutas.INSTALL, "Tesseract-OCR", "tessdata"),
):
    if os.path.isdir(_tessdata):
        os.environ["TESSDATA_PREFIX"] = _tessdata
        break

# Umbral: si una página de PDF trae menos de estos caracteres, se considera
# "sin texto real" y se manda a OCR.
_MIN_CARACTERES_PAGINA = 20
_DPI_OCR = 300


class OCRNoDisponible(RuntimeError):
    """Se requiere OCR pero el binario de Tesseract no está disponible."""


class OCRCancelado(RuntimeError):
    """Se abortó el análisis porque quien llamó pidió cancelar (ver `cancelado`).

    El OCR de un documento puede durar bastante (rasteriza y pasa Tesseract por
    CADA página). Para poder detenerlo de verdad, las funciones públicas aceptan
    un `cancelado()` que se consulta entre páginas: al devolver True se corta al
    momento en vez de seguir procesando el resto del documento."""


# Caché de las comprobaciones de Tesseract. Además de ahorrar trabajo, evita que
# se relance tesseract por CADA OCR: esas comprobaciones, hechas con pytesseract,
# abrían una ventana de consola que parpadeaba en pantalla (una por análisis, y
# varias a la vez con el OCR en paralelo).
_tess_disponible: bool | None = None
_tess_idioma: str | None = None


def _run_tesseract_oculto(args: list[str], timeout: int = 20) -> "subprocess.CompletedProcess":
    """Ejecuta tesseract capturando la salida y SIN abrir ventana de consola
    (creationflags en Windows). Se usa para las comprobaciones (versión/idiomas)
    en lugar de pytesseract, cuyas llamadas abren una consola que parpadea."""
    return subprocess.run(
        args, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        creationflags=_CREATE_NO_WINDOW, timeout=timeout,
    )


def tesseract_disponible() -> bool:
    """True si el binario de Tesseract responde. Se comprueba UNA vez y se cachea:
    evita relanzar tesseract (y su parpadeo de consola) en cada OCR."""
    global _tess_disponible
    if _tess_disponible is None:
        try:
            r = _run_tesseract_oculto(
                [pytesseract.pytesseract.tesseract_cmd, "--version"])
            _tess_disponible = r.returncode == 0
        except Exception:  # noqa: BLE001 — binario ausente/ruta mala -> no disponible
            _tess_disponible = False
    return _tess_disponible


def idioma_ocr() -> str:
    """'spa+eng' si el modelo de español está disponible; si no, 'eng'. Se cachea.

    Se prioriza comprobar el archivo del modelo (spa.traineddata) en vez de
    preguntar a tesseract: además de ser más rápido, evita abrir consola y no
    falla con rutas con acentos (p. ej. la carpeta 'tesorería').
    """
    global _tess_idioma
    if _tess_idioma is not None:
        return _tess_idioma
    carpeta = os.environ.get("TESSDATA_PREFIX", "")
    if carpeta and os.path.exists(os.path.join(carpeta, "spa.traineddata")):
        _tess_idioma = "spa+eng"
        return _tess_idioma
    try:
        r = _run_tesseract_oculto(
            [pytesseract.pytesseract.tesseract_cmd, "--list-langs"])
        if b"spa" in (r.stdout or b"") or b"spa" in (r.stderr or b""):
            _tess_idioma = "spa+eng"
            return _tess_idioma
    except Exception:  # noqa: BLE001
        pass
    _tess_idioma = "eng"
    return _tess_idioma


def _ocr_imagen(img: Image.Image, psm: int | None = None, cancelado=None) -> str:
    """Aplica OCR a una imagen lanzando tesseract como subproceso PROPIO, para
    poder MATARLO si `cancelado()` pasa a True.

    Se hace así (en vez de pytesseract.image_to_string) porque el tesseract que
    lanza pytesseract no lo controlamos: al abandonar el hilo de una carga
    detenida, ese proceso seguiría ocupando la CPU hasta terminar la página,
    retrasando la reanudación. Aquí, al detener, matamos el proceso al instante.

    Devuelve el mismo texto que pytesseract (saltos de línea normalizados a '\\n').
    """
    if not tesseract_disponible():
        raise OCRNoDisponible(
            "Se necesita OCR pero no se encontró Tesseract. Instálalo o "
            "revisa la ruta en core/ocr.py."
        )
    if img.mode not in ("L", "RGB"):
        img = img.convert("RGB")
    # Entrada y salida en archivos temporales (igual que pytesseract): evita el
    # pipe de stdout, cuyo buffer podría bloquear a tesseract si no se lee mientras
    # sondeamos la cancelación.
    tmpdir = tempfile.mkdtemp()
    entrada = os.path.join(tmpdir, "in.png")
    base_salida = os.path.join(tmpdir, "out")
    img.save(entrada, format="PNG")
    proc = None
    try:
        cmd = [pytesseract.pytesseract.tesseract_cmd, entrada, base_salida,
               "-l", idioma_ocr()]
        if psm is not None:
            cmd += ["--psm", str(psm)]
        proc = subprocess.Popen(
            cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=_CREATE_NO_WINDOW,
        )
        # Se sondea cada 50 ms; si piden cancelar, se mata tesseract al momento.
        while proc.poll() is None:
            if cancelado is not None and cancelado():
                proc.kill()
                proc.wait()
                raise OCRCancelado()
            time.sleep(0.05)
        with open(base_salida + ".txt", "r", encoding="utf-8", errors="replace") as fh:
            texto = fh.read()
        # tesseract escribe LF en el .txt, pero se normaliza por si acaso, para dar
        # exactamente el mismo texto que daba pytesseract.
        return texto.replace("\r\n", "\n").replace("\r", "\n")
    finally:
        if proc is not None and proc.poll() is None:
            proc.kill()
        shutil.rmtree(tmpdir, ignore_errors=True)


def _texto_confiable(texto: str) -> bool:
    """True si el texto de la capa del PDF parece legible (no codificación rota)."""
    if len(texto.strip()) < _MIN_CARACTERES_PAGINA:
        return False
    normales = len(_RE_CHARS_NORMALES.findall(texto))
    return normales / len(texto) >= _UMBRAL_TEXTO_CONFIABLE


def _ocr_pagina(pagina: "fitz.Page", psm: int | None = None, cancelado=None) -> str:
    pix = pagina.get_pixmap(dpi=_DPI_OCR)
    # Rasterizar ya costó; antes de pagar el Tesseract de esta página, se mira si
    # entretanto pidieron cancelar. El propio _ocr_imagen también es cancelable
    # (mata tesseract si se detiene a media página).
    if cancelado is not None and cancelado():
        raise OCRCancelado()
    # Imagen directa desde los píxeles del pixmap (sin round-trip a PNG): evita dos
    # codificaciones/decodificaciones PNG por página (más rápido).
    modo = "RGBA" if pix.alpha else "RGB"
    img = Image.frombytes(modo, (pix.width, pix.height), pix.samples)
    return _ocr_imagen(img, psm=psm, cancelado=cancelado)


def _texto_pdf(ruta: str, forzar_ocr: bool = False, psm: int | None = None,
               cancelado=None) -> tuple[str, bool]:
    """Devuelve (texto, se_uso_ocr) para un PDF.

    Usa la capa de texto del PDF si es legible; si la página no tiene texto, su
    codificación está rota, o se fuerza, rasteriza y aplica OCR. Forzar OCR es
    útil cuando la capa de texto trae solo "ruido" (p. ej. la impresión de un
    correo de Outlook) y el estado de cuenta real está como imagen. Con psm se
    fuerza siempre OCR usando ese modo de segmentación de página.

    `cancelado` (opcional): se consulta ANTES de cada página; si devuelve True se
    lanza OCRCancelado y no se procesa el resto del documento.
    """
    partes: list[str] = []
    uso_ocr = False
    with fitz.open(ruta) as doc:
        for pagina in doc:
            if cancelado is not None and cancelado():
                raise OCRCancelado(ruta)
            texto = pagina.get_text("text") or ""
            if psm is None and not forzar_ocr and _texto_confiable(texto):
                partes.append(texto)
            else:
                partes.append(_ocr_pagina(pagina, psm=psm, cancelado=cancelado))
                uso_ocr = True
    return "\n".join(partes), uso_ocr


def extraer_texto(ruta: str, forzar_ocr: bool = False, psm: int | None = None,
                  cancelado=None) -> tuple[str, bool]:
    """Extrae el texto de un estado de cuenta.

    Args:
        ruta: ruta a un archivo .pdf, .png, .jpg, .jpeg, .tif o .tiff.
        forzar_ocr: si es True, ignora la capa de texto del PDF y aplica OCR a
            todas las páginas (útil cuando la capa de texto es solo ruido).
        psm: modo de segmentación de página de Tesseract. Si se indica, fuerza
            OCR con ese modo (p. ej. 11 = "texto disperso", útil para tablas).
        cancelado: callable sin argumentos consultado entre páginas; si devuelve
            True se aborta con OCRCancelado (para poder detener una carga sin
            esperar a que termine el documento entero).

    Returns:
        (texto, se_uso_ocr): el texto plano del documento y si hubo que
        recurrir al OCR (útil para advertir sobre menor confiabilidad).

    Raises:
        FileNotFoundError: si la ruta no existe.
        ValueError: si la extensión no es compatible.
        OCRNoDisponible: si se requería OCR pero Tesseract no está disponible.
        OCRCancelado: si `cancelado()` devolvió True durante el análisis.
    """
    if not os.path.exists(ruta):
        raise FileNotFoundError(ruta)

    ext = os.path.splitext(ruta)[1].lower()
    if ext == ".pdf":
        texto, uso_ocr = _texto_pdf(ruta, forzar_ocr=forzar_ocr, psm=psm,
                                    cancelado=cancelado)
        return _normalizar(texto), uso_ocr
    if ext in (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"):
        if cancelado is not None and cancelado():
            raise OCRCancelado(ruta)
        return _normalizar(_ocr_imagen(Image.open(ruta), psm=psm, cancelado=cancelado)), True
    raise ValueError(f"Extensión no soportada: {ext}")


def texto_disperso(ruta: str, cancelado=None) -> str:
    """OCR en modo 'texto disperso' (PSM 11). Recupera números/CLABE que la
    segmentación normal de página omite (p. ej. la celda de una tabla en una
    carta de asignación de cuenta). Pierde el orden de lectura, por eso se usa
    solo como último recurso cuando no se halló la CLABE de otra forma."""
    texto, _ = extraer_texto(ruta, forzar_ocr=True, psm=11, cancelado=cancelado)
    return texto
