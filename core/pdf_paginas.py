"""Separación de un PDF en una página por archivo.

El banco a veces entrega UN PDF con muchos comprobantes, uno por página. El
extractor de comprobantes (`core.api.leer_comprobantes_pagos`) recibe RUTAS de
archivo y, en su respuesta, lo único que ata cada comprobante a su archivo es el
campo `documento_lectura` (el nombre del archivo). Con un PDF de N páginas ese
nombre es el mismo para las N lecturas, así que no habría forma de saber qué
página adjuntar a qué movimiento.

Por eso se separan las páginas A DISCO **antes** de llamar al extractor: cada
página llega con un nombre único, la API devuelve un comprobante por página y el
reparto contra las dispersiones funciona sin cambiar ni una regla de casado.

Materializar los archivos no es opcional: el RPA de subida adjunta rutas reales
en el navegador, así que las páginas tienen que existir en disco.

Dependencia: PyMuPDF (módulo `pymupdf`), ya usado en `core.ocr`.
"""

from __future__ import annotations

import os

import pymupdf


class ErrorPdf(Exception):
    """No se pudo abrir o separar el PDF (corrupto, protegido o sin páginas)."""


def _ruta_libre(ruta: str) -> str:
    """Si `ruta` existe, agrega ' (n)' antes de la extensión hasta hallar un nombre
    libre. Nunca sobrescribe: una corrida anterior pudo dejar páginas del mismo
    nombre en la carpeta del día, y pisarlas rompería vínculos ya hechos."""
    if not os.path.exists(ruta):
        return ruta
    base, ext = os.path.splitext(ruta)
    n = 2
    while os.path.exists(f"{base} ({n}){ext}"):
        n += 1
    return f"{base} ({n}){ext}"


def _abrir(ruta_pdf: str) -> "pymupdf.Document":
    """Abre el PDF o lanza ErrorPdf con un mensaje que nombra el archivo."""
    nombre = os.path.basename(ruta_pdf)
    try:
        doc = pymupdf.open(ruta_pdf)
    except Exception as exc:  # noqa: BLE001 — se traduce a un error propio
        raise ErrorPdf(f"«{nombre}»: no se pudo abrir el PDF ({exc}).") from exc
    if doc.needs_pass:
        doc.close()
        raise ErrorPdf(f"«{nombre}»: el PDF está protegido con contraseña.")
    if doc.page_count < 1:
        doc.close()
        raise ErrorPdf(f"«{nombre}»: el PDF no tiene páginas.")
    return doc


def contar_paginas(ruta_pdf: str) -> int:
    """Páginas de un PDF. Lanza ErrorPdf si no se puede leer."""
    doc = _abrir(ruta_pdf)
    try:
        return doc.page_count
    finally:
        doc.close()


# Ancho por defecto al que se rasteriza una página para verla en la interfaz.
ANCHO_VISTA_PREVIA = 640

# Tope de ancho al rasterizar. A 2100 px una carta pesa ~285 KB y tarda ~136 ms;
# más allá el PNG crece rápido sin que se lea mejor, y para leer el detalle fino
# está el visor del sistema.
ANCHO_MAXIMO = 2400


def rasterizar_pagina(
    ruta_pdf: str, pagina: int = 0, ancho_px: int = ANCHO_VISTA_PREVIA,
) -> tuple[bytes, int, int]:
    """Rasteriza una página a PNG. Devuelve `(png, ancho, alto)` en píxeles.

    Flet no sabe dibujar PDF, pero sí un PNG (y `ft.Image` acepta los bytes tal
    cual). Se devuelven también las dimensiones porque quien lo muestra necesita
    fijarle tamaño explícito a la imagen para que el contenedor pueda hacer scroll
    cuando el zoom la saca de la vista.

    Cuesta ~30 ms a 700 px y ~136 ms a 2100, así que conviene llamarlo BAJO DEMANDA
    —cuando el usuario elige la página o cambia el zoom— y no de golpe para un lote.

    Lanza ErrorPdf si el PDF no se puede abrir o la página no existe."""
    doc = _abrir(ruta_pdf)
    try:
        if not 0 <= pagina < doc.page_count:
            raise ErrorPdf(
                f"«{os.path.basename(ruta_pdf)}»: no tiene página {pagina + 1}.")
        pag = doc[pagina]
        ancho = max(80, min(int(ancho_px), ANCHO_MAXIMO))
        escala = ancho / pag.rect.width if pag.rect.width else 1
        try:
            pix = pag.get_pixmap(matrix=pymupdf.Matrix(escala, escala), alpha=False)
            return pix.tobytes("png"), pix.width, pix.height
        except Exception as exc:  # noqa: BLE001 — se traduce a un error propio
            raise ErrorPdf(
                f"«{os.path.basename(ruta_pdf)}»: no se pudo generar la vista "
                f"previa ({exc}).") from exc
    finally:
        doc.close()


def separar_paginas(ruta_pdf: str, carpeta_destino: str) -> list[str]:
    """Escribe una página por archivo en `carpeta_destino` y devuelve sus rutas, en
    el orden del documento.

    Si el PDF trae UNA sola página devuelve `[ruta_pdf]` sin copiar nada: el caso
    habitual no paga ni el disco ni el riesgo de una copia, y aguas arriba todo
    sigue funcionando exactamente igual que antes de que existiera este módulo.

    Los archivos se llaman '<nombre original> pN.pdf'. El nombre base viene de un
    archivo que ya existe en disco, así que no hace falta sanearlo; solo se evita
    pisar nombres ya ocupados (ver `_ruta_libre`).

    Lanza ErrorPdf si el PDF no se puede abrir, está protegido o no tiene páginas.
    """
    doc = _abrir(ruta_pdf)
    try:
        if doc.page_count == 1:
            return [ruta_pdf]
        os.makedirs(carpeta_destino, exist_ok=True)
        base = os.path.splitext(os.path.basename(ruta_pdf))[0]
        salidas: list[str] = []
        for i in range(doc.page_count):
            destino = _ruta_libre(
                os.path.join(carpeta_destino, f"{base} p{i + 1}.pdf"))
            nuevo = pymupdf.open()
            try:
                nuevo.insert_pdf(doc, from_page=i, to_page=i)
                nuevo.save(destino)
            except Exception as exc:  # noqa: BLE001 — se traduce a un error propio
                raise ErrorPdf(
                    f"«{os.path.basename(ruta_pdf)}»: no se pudo escribir la "
                    f"página {i + 1} ({exc}).") from exc
            finally:
                nuevo.close()
            salidas.append(destino)
        return salidas
    finally:
        doc.close()
