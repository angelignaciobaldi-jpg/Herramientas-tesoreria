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

import fitz  # PyMuPDF
import pytesseract
from PIL import Image

from . import rutas

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


def tesseract_disponible() -> bool:
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


def idioma_ocr() -> str:
    """Devuelve 'spa+eng' si el idioma español está disponible; si no, 'eng'.

    Se comprueba la existencia del archivo del modelo en lugar de llamar a
    pytesseract.get_languages(), porque ese comando falla al decodificar rutas
    con acentos (p. ej. la carpeta 'tesorería').
    """
    carpeta = os.environ.get("TESSDATA_PREFIX", "")
    if carpeta and os.path.exists(os.path.join(carpeta, "spa.traineddata")):
        return "spa+eng"
    try:
        if "spa" in pytesseract.get_languages(config=""):
            return "spa+eng"
    except Exception:
        pass
    return "eng"


def _ocr_imagen(img: Image.Image, psm: int | None = None) -> str:
    if not tesseract_disponible():
        raise OCRNoDisponible(
            "Se necesita OCR pero no se encontró Tesseract. Instálalo o "
            "revisa la ruta en core/ocr.py."
        )
    if img.mode not in ("L", "RGB"):
        img = img.convert("RGB")
    config = f"--psm {psm}" if psm is not None else ""
    return pytesseract.image_to_string(img, lang=idioma_ocr(), config=config)


def _texto_confiable(texto: str) -> bool:
    """True si el texto de la capa del PDF parece legible (no codificación rota)."""
    if len(texto.strip()) < _MIN_CARACTERES_PAGINA:
        return False
    normales = len(_RE_CHARS_NORMALES.findall(texto))
    return normales / len(texto) >= _UMBRAL_TEXTO_CONFIABLE


def _ocr_pagina(pagina: "fitz.Page", psm: int | None = None) -> str:
    pix = pagina.get_pixmap(dpi=_DPI_OCR)
    return _ocr_imagen(Image.open(io.BytesIO(pix.tobytes("png"))), psm=psm)


def _texto_pdf(ruta: str, forzar_ocr: bool = False, psm: int | None = None) -> tuple[str, bool]:
    """Devuelve (texto, se_uso_ocr) para un PDF.

    Usa la capa de texto del PDF si es legible; si la página no tiene texto, su
    codificación está rota, o se fuerza, rasteriza y aplica OCR. Forzar OCR es
    útil cuando la capa de texto trae solo "ruido" (p. ej. la impresión de un
    correo de Outlook) y el estado de cuenta real está como imagen. Con psm se
    fuerza siempre OCR usando ese modo de segmentación de página.
    """
    partes: list[str] = []
    uso_ocr = False
    with fitz.open(ruta) as doc:
        for pagina in doc:
            texto = pagina.get_text("text") or ""
            if psm is None and not forzar_ocr and _texto_confiable(texto):
                partes.append(texto)
            else:
                partes.append(_ocr_pagina(pagina, psm=psm))
                uso_ocr = True
    return "\n".join(partes), uso_ocr


def extraer_texto(ruta: str, forzar_ocr: bool = False, psm: int | None = None) -> tuple[str, bool]:
    """Extrae el texto de un estado de cuenta.

    Args:
        ruta: ruta a un archivo .pdf, .png, .jpg, .jpeg, .tif o .tiff.
        forzar_ocr: si es True, ignora la capa de texto del PDF y aplica OCR a
            todas las páginas (útil cuando la capa de texto es solo ruido).
        psm: modo de segmentación de página de Tesseract. Si se indica, fuerza
            OCR con ese modo (p. ej. 11 = "texto disperso", útil para tablas).

    Returns:
        (texto, se_uso_ocr): el texto plano del documento y si hubo que
        recurrir al OCR (útil para advertir sobre menor confiabilidad).

    Raises:
        FileNotFoundError: si la ruta no existe.
        ValueError: si la extensión no es compatible.
        OCRNoDisponible: si se requería OCR pero Tesseract no está disponible.
    """
    if not os.path.exists(ruta):
        raise FileNotFoundError(ruta)

    ext = os.path.splitext(ruta)[1].lower()
    if ext == ".pdf":
        texto, uso_ocr = _texto_pdf(ruta, forzar_ocr=forzar_ocr, psm=psm)
        return _normalizar(texto), uso_ocr
    if ext in (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"):
        return _normalizar(_ocr_imagen(Image.open(ruta), psm=psm)), True
    raise ValueError(f"Extensión no soportada: {ext}")


def texto_disperso(ruta: str) -> str:
    """OCR en modo 'texto disperso' (PSM 11). Recupera números/CLABE que la
    segmentación normal de página omite (p. ej. la celda de una tabla en una
    carta de asignación de cuenta). Pierde el orden de lectura, por eso se usa
    solo como último recurso cuando no se halló la CLABE de otra forma."""
    texto, _ = extraer_texto(ruta, forzar_ocr=True, psm=11)
    return texto
