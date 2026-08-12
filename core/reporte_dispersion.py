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


# Clave de 'Tipo de factura' de las notas de crédito. Sus importes SIEMPRE van en
# negativo: una nota de crédito resta de lo que se le paga al proveedor, así que al
# sumarla con el resto de sus documentos el neto sale solo.
TIPO_NOTA_CREDITO = "NC"

# Los tres campos monetarios de una solicitud. Se declaran antes de la clase porque
# __post_init__ los recorre para aplicarles el signo de las notas de crédito.
_IMPORTES = ("total_factura", "saldo_factura", "saldo_programado")


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
    # --- Campos de la API que NO se muestran en la tabla ---------------------
    # Se conservan (crudos, tipados) para validaciones futuras del proceso de
    # dispersión. Solo se llenan cuando la fila viene de la API (desde_api); al
    # leer del Excel quedan con su valor por defecto. No entran en `clave()`
    # (la identidad la dan los campos visibles) ni se muestran en la UI.
    id_empresa: int | None = None
    id_solicitud_pago: int | None = None
    id_solicitud_pago_detalle: int | None = None
    folio_documento: str = ""
    clabe_interbancaria_proveedor: str = ""
    id_proveedor: int | None = None

    def __post_init__(self) -> None:
        """Deja en NEGATIVO los importes de las notas de crédito ('NC').

        Va aquí y no en la lectura para que la regla valga igual venga la fila del
        Excel o de la API, y para que no se pueda construir una FilaSolicitud 'NC'
        con importes positivos por descuido.

        Se usa -abs() en vez de invertir el signo: así el resultado no depende de
        cómo lo mande el origen (que a veces ya trae el negativo) y reconstruir la
        fila —p. ej. al recargar un reporte— no lo vuelve positivo.

        Es lo que hace que `total_a_pagar` descuente las NC con una simple suma.
        """
        if not es_nota_credito(self):
            return
        for campo in _IMPORTES:
            valor = getattr(self, campo)
            if valor is not None:
                setattr(self, campo, -abs(valor))

    def clave(self) -> tuple:
        """Identidad de la fila para evitar duplicados al recargar reportes."""
        return (
            self.empresa, self.folio, self.tipo, self.folio_factura,
            self.proveedor, self.cuenta_bancaria, self.fecha_factura,
            self.fecha_vencimiento, self.tipo_solicitud,
            self.total_factura, self.saldo_factura, self.saldo_programado,
            self.moneda, self.producto, self.comentarios,
        )


def es_nota_credito(fila) -> bool:
    """True si la fila es una nota de crédito ('NC')."""
    return (getattr(fila, "tipo", "") or "").strip().upper() == TIPO_NOTA_CREDITO


def total_a_pagar(movimientos) -> float:
    """Suma de Saldo Programado de lo que REALMENTE se dispersa.

    Las notas de crédito SÍ entran, y restan. El reporte trae la factura con su
    importe completo y la NC como un renglón aparte, así que hay que descontarla para
    llegar al total correcto; el SIPP ya muestra ese neto en su propia pantalla.
    La resta sale sola porque `FilaSolicitud.__post_init__` deja los importes de las
    NC en negativo.

    Es la única forma correcta de totalizar un grupo de movimientos para pagar: usarla
    en vez de sumar `saldo_programado` a mano, para que la regla viva en un solo sitio.
    """
    return round(sum((m.saldo_programado or 0) for m in movimientos), 2)


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
_CAMPOS_NUMERICOS = set(_IMPORTES)


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


def _entero(valor) -> int | None:
    """Convierte a int (para los ids de la API). Acepta enteros, cadenas y flotantes
    'enteros' (p. ej. 5.0). None si viene vacío o no se puede convertir."""
    if valor is None or valor == "":
        return None
    try:
        return int(valor)
    except (TypeError, ValueError):
        try:
            return int(float(valor))
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


# ------------------------------------------------------------------ API (JSON)
# Campo del JSON de la API (endpoint /api/dispersiones/no_pemex) -> campo de
# FilaSolicitud. Se separan por tipo de conversión (texto / número / fecha).
_API_TEXTO = {
    "empresa": "nb_Empresa",              # nombre corto (coincide con ui.comun.EMPRESAS)
    "folio": "id_SolicitudPago",          # folio de la SOLICITUD (compartido por sus docs)
    "tipo": "cl_TipoDocumento",           # 'Tipo de factura' (p. ej. 'NF')
    "folio_factura": "nu_FolioDocumento",
    "proveedor": "nb_Proveedor",
    "cuenta_bancaria": "de_CuentaBancariaProveedor",  # 'BANCO - CLABE' (trae los dígitos)
    "tipo_solicitud": "de_TipoSolicitudpago",
    "producto": "nb_Producto",
    "comentarios": "de_Comentarios",
    # Ocultos (no se muestran; para validaciones futuras).
    "folio_documento": "nu_FolioDocumento",
    "clabe_interbancaria_proveedor": "nu_ClabeInterbancariaProveedor",
}
_API_NUMERO = {
    "total_factura": "im_Total",
    "saldo_factura": "im_Saldo",
    "saldo_programado": "im_SaldoSolicitud",  # saldo de la solicitud = 'Saldo Programado'
}
_API_FECHA = {
    "fecha_factura": "fh_Factura",
    "fecha_vencimiento": "fh_VencimientoFactura",
}
# Enteros (ids) ocultos: no se muestran; se conservan para validaciones futuras.
# id_Empresa se trata aparte (solo se toma si la fila no trae ya un id_empresa).
_API_ENTERO = {
    "id_solicitud_pago": "id_SolicitudPago",
    "id_solicitud_pago_detalle": "nd_SolicitudPagoDetalle",
    "id_proveedor": "id_Proveedor",
}


def _fecha_api(valor) -> str:
    """ISO ('YYYY-MM-DDTHH:MM:SS' o 'YYYY-MM-DD') -> 'DD/MM/AAAA' (lo que espera la
    UI, p. ej. el filtro de vencimiento). '' si viene vacío; el original si no parsea."""
    s = str(valor or "").strip()
    if not s:
        return ""
    try:
        anio, mes, dia = s[:10].split("-")
        return f"{int(dia):02d}/{int(mes):02d}/{anio}"
    except (ValueError, TypeError):
        return s


def desde_api(respuesta) -> list[FilaSolicitud]:
    """Convierte la respuesta del endpoint /api/dispersiones/no_pemex en filas
    tipadas (FilaSolicitud), listas para volcar en la tabla.

    Acepta la respuesta completa (dict con clave 'data') o directamente la lista de
    registros. Ignora entradas que no sean objetos. La moneda se normaliza (MN->MXN)
    y las fechas ISO se pasan a 'DD/MM/AAAA', igual que al leer el Excel.
    """
    if isinstance(respuesta, dict):
        registros = respuesta.get("data") or []
    else:
        registros = respuesta or []
    filas: list[FilaSolicitud] = []
    for reg in registros:
        if not isinstance(reg, dict):
            continue
        datos: dict = {}
        for campo, clave in _API_TEXTO.items():
            datos[campo] = _texto(reg.get(clave))
        for campo, clave in _API_NUMERO.items():
            datos[campo] = _numero(reg.get(clave))
        for campo, clave in _API_FECHA.items():
            datos[campo] = _fecha_api(reg.get(clave))
        for campo, clave in _API_ENTERO.items():
            datos[campo] = _entero(reg.get(clave))
        datos["moneda"] = normalizar_moneda(reg.get("c_MonedaSAT"))
        # id_Empresa: solo se toma de la API si la fila no trae ya un id_empresa
        # (se conserva el previo si existiera).
        if not datos.get("id_empresa"):
            datos["id_empresa"] = _entero(reg.get("id_Empresa"))
        filas.append(FilaSolicitud(**datos))
    return filas
