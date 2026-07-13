"""Lectura de los reportes XLSX de 'Dispersión (No Pemex)' del SIPP.

El botón "Generar XLS" del modal descarga un Excel con este layout:
  - Filas 0-6: título y filtros aplicados (metadatos).
  - Fila de encabezados: Folio, Tipo de factura, Folio Factura, Empresa,
    Proveedor, Fecha Factura, Fecha Vencimiento, Tipo Solicitud, Moneda,
    Producto, Total Factura, Saldo Factura, Saldo Programado, Cuenta Bancaria,
    Comentarios.
  - Filas de datos y, al final de cada grupo por cuenta bancaria, una fila
    'TOTAL PROGRAMADO' propia del Excel (que aquí se ignora: los totales se
    recalculan en la UI).

`leer` / `leer_varios` devuelven filas ya tipadas (FilaSolicitud), listas para
volcarse en la tabla de la pantalla de dispersión.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

import openpyxl


@dataclass
class FilaSolicitud:
    """Un renglón del reporte de solicitudes de pago (una factura/solicitud)."""

    empresa: str = ""
    folio: str = ""
    tipo: str = ""
    folio_factura: str = ""
    proveedor: str = ""
    cuenta_bancaria: str = ""
    fecha_factura: str = ""
    fecha_vencimiento: str = ""
    tipo_solicitud: str = ""
    total_factura: float | None = None
    saldo_factura: float | None = None
    saldo_programado: float | None = None
    moneda: str = ""
    producto: str = ""
    comentarios: str = ""

    def clave(self) -> tuple:
        """Identidad de la fila para evitar duplicados al recargar reportes."""
        return (
            self.empresa, self.folio, self.tipo, self.folio_factura,
            self.proveedor, self.cuenta_bancaria, self.fecha_factura,
            self.fecha_vencimiento, self.tipo_solicitud,
            self.total_factura, self.saldo_factura, self.saldo_programado,
            self.moneda, self.producto, self.comentarios,
        )


# Encabezado en el Excel -> campo de FilaSolicitud.
_COLUMNAS = {
    "Folio": "folio",
    "Tipo de factura": "tipo",
    "Folio Factura": "folio_factura",
    "Empresa": "empresa",
    "Proveedor": "proveedor",
    "Fecha Factura": "fecha_factura",
    "Fecha Vencimiento": "fecha_vencimiento",
    "Tipo Solicitud": "tipo_solicitud",
    "Total Factura": "total_factura",
    "Saldo Factura": "saldo_factura",
    "Saldo Programado": "saldo_programado",
    "Cuenta Bancaria": "cuenta_bancaria",
    "Moneda": "moneda",
    "Producto": "producto",
    "Comentarios": "comentarios",
}
_CAMPOS_NUMERICOS = {"total_factura", "saldo_factura", "saldo_programado"}


def _texto(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return str(valor).strip()


def _numero(valor) -> float | None:
    if valor is None or valor == "":
        return None
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def normalizar_moneda(texto) -> str:
    """Normaliza el tipo de moneda a sus siglas en MAYÚSCULAS:

      - quita puntos y espacios entre las siglas ('U.S.D.' -> 'USD',
        'M.N.' -> 'MN', 'M.X.N.' -> 'MXN');
      - trata 'MN' (Moneda Nacional) como 'MXN'.

    Devuelve '' si no hay valor. Se usa para separar las dispersiones por moneda
    (además de por empresa), así que un mismo tipo escrito de varias formas debe
    colapsar a la misma sigla.
    """
    siglas = re.sub(r"[.\s]+", "", str(texto or "")).upper()
    return "MXN" if siglas == "MN" else siglas


def leer(ruta: str) -> list[FilaSolicitud]:
    """Lee un XLSX de reporte y devuelve sus filas de datos (sin las filas de
    'TOTAL PROGRAMADO' ni las vacías). Si el layout no se reconoce, devuelve []."""
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    try:
        ws = wb.active
        filas = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()

    # Localiza la fila de encabezados (la que trae 'Folio' y 'Folio Factura').
    encabezados = None
    inicio = 0
    for i, fila in enumerate(filas):
        vals = [str(c).strip() if c is not None else "" for c in fila]
        if "Folio" in vals and "Folio Factura" in vals:
            encabezados = vals
            inicio = i + 1
            break
    if encabezados is None:
        return []

    # Posición de cada campo según el encabezado (robusto a reordenamientos).
    posicion = {
        _COLUMNAS[h]: j for j, h in enumerate(encabezados) if h in _COLUMNAS
    }

    def valor(fila, campo):
        j = posicion.get(campo)
        if j is None or j >= len(fila):
            return None
        return fila[j]

    # La empresa del reporte va en los metadatos (celda "Empresa:" / C3). Es una
    # sola por archivo (reporte filtrado); se usa para TODAS las filas.
    empresa_reporte = _empresa_reporte(filas, inicio)

    resultado: list[FilaSolicitud] = []
    for fila in filas[inicio:]:
        # Saltar filas vacías y las de 'TOTAL PROGRAMADO' (sin Folio de dato).
        if valor(fila, "folio") in (None, ""):
            continue
        datos = {}
        for campo in _COLUMNAS.values():
            crudo = valor(fila, campo)
            datos[campo] = (
                _numero(crudo) if campo in _CAMPOS_NUMERICOS else _texto(crudo)
            )
        if empresa_reporte:
            datos["empresa"] = empresa_reporte
        # Normaliza la moneda (MN -> MXN, quita puntos) para que la separación por
        # moneda agrupe correctamente aunque venga escrita de distintas formas.
        datos["moneda"] = normalizar_moneda(datos.get("moneda", ""))
        resultado.append(FilaSolicitud(**datos))
    return resultado


def _empresa_reporte(filas: list, inicio: int) -> str:
    """Empresa a la que pertenece el reporte: se toma del metadato 'Empresa:'
    (el valor a su derecha) que aparece antes de la tabla; si no, de la celda C3
    (fila 3, columna C). Devuelve '' si no se encuentra."""
    for fila in filas[:inicio]:
        for j, celda in enumerate(fila):
            if str(celda or "").strip().rstrip(":").lower() == "empresa":
                for k in range(j + 1, len(fila)):
                    if fila[k] not in (None, ""):
                        return _texto(fila[k])
    # Respaldo: C3 directo (fila índice 2, columna índice 2).
    if len(filas) > 2 and len(filas[2]) > 2 and filas[2][2] not in (None, ""):
        return _texto(filas[2][2])
    return ""


def leer_varios(rutas: list[str]) -> list[FilaSolicitud]:
    """Lee varios reportes y concatena sus filas (en el orden dado)."""
    todas: list[FilaSolicitud] = []
    for ruta in rutas:
        todas.extend(leer(ruta))
    return todas
