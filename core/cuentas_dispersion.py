"""Catálogo de 'Cuentas de dispersión' (por id de empresa).

Lee un Excel SENCILLO para filtrar las cuentas en la pantalla de Dispersión
(No Pemex):

  - ``id Empresa``          -> FK al id de la empresa (ver EMPRESAS en la pantalla).
  - ``Cuenta``              -> la cuenta que se MUESTRA en el selector y por la que
                               el RPA busca la cuenta.
  - ``CLABE interbancaria`` -> la CLABE de origen que se usa como cuenta origen del
                               TXT en pesos (opcional; se muestra en su propio
                               selector).

Es el complemento por-empresa del catálogo bancario general (`cuentas_bancarias`).
Se consulta por id de empresa (`cuentas_por_id_empresa`, `clabes_por_id_empresa`).

Las columnas se localizan por ENCABEZADO (fila 1), no por posición. Si el Excel se
actualiza, basta reabrir la app (o readjuntarlo en Configuración) para reflejarlo.
"""

from __future__ import annotations

import json
import os
import re

try:
    import openpyxl
except ImportError:  # openpyxl es opcional; sin él, el catálogo queda vacío.
    openpyxl = None

from . import rutas
from .exportador_devoluciones import banco_formato
from .extractores import validar_clabe

# El Excel lo actualiza el usuario, así que va junto al .exe (no empaquetado).
RUTA_EXCEL = os.path.join(rutas.DATOS, "Cuentas dispersion", "CUENTAS DISPERSION.xlsx")
# Copia en caché (permite seguir trabajando aunque el Excel esté abierto/bloqueado).
_RUTA_CACHE = os.path.join(rutas.DATOS, "_cuentas_dispersion_cache.json")

# Encabezados aceptados (normalizados) para cada columna.
_HDR_ID = ("id empresa",)
_HDR_CUENTA = ("cuenta",)
# CLABE interbancaria (opcional): cuenta origen del TXT en pesos.
_HDR_CLABE = ("clabe interbancaria", "clabe")


class ExcelCuentasDispersionInvalido(ValueError):
    """El Excel adjuntado no tiene el formato esperado (id Empresa + Cuenta)."""


def hay_excel() -> bool:
    """True si ya hay un Excel de cuentas de dispersión colocado para consulta."""
    return os.path.exists(RUTA_EXCEL)


def _norm(valor) -> str:
    """Normaliza un encabezado: minúsculas, sin espacios/underscores repetidos."""
    return re.sub(r"[\s_]+", " ", str(valor or "").strip().casefold())


def _id_empresa(valor) -> int | None:
    """Convierte el id de empresa (int, float o texto) a int; None si no es válido."""
    if valor is None or valor == "":
        return None
    try:
        return int(float(valor))
    except (TypeError, ValueError):
        return None


def _clabe_valida(texto) -> bool:
    """True si `texto` contiene una CLABE de 18 dígitos con dígito de control
    correcto (ignora espacios/guiones u otros separadores)."""
    dig = re.sub(r"\D", "", str(texto or ""))
    return len(dig) == 18 and validar_clabe(dig)


def _texto_cuenta(valor) -> str:
    """Cuenta como texto, sin el apóstrofo con que Excel fuerza texto y sin espacios
    en los extremos. Un número entero de Excel se muestra sin el '.0'."""
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return str(valor).replace("'", "").strip()


def _leer_excel(ruta: str) -> dict[int, list[dict]]:
    """Lee el Excel. Devuelve {} si no se puede (no existe, bloqueado, formato
    inesperado). Columnas por ENCABEZADO: 'id Empresa', 'Cuenta' y, opcional,
    'CLABE interbancaria'. Cada empresa mapea a una lista de registros
    {'cuenta': str, 'clabe': str} (sin duplicados)."""
    catalogo: dict[int, list[dict]] = {}
    if openpyxl is None or not os.path.exists(ruta):
        return catalogo
    try:
        wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
    except Exception:
        return catalogo  # p. ej. PermissionError si está abierto en Excel
    try:
        ws = wb[wb.sheetnames[0]]
        filas = ws.iter_rows(values_only=True)
        try:
            encabezados = next(filas)
        except StopIteration:
            return catalogo  # hoja vacía
        idx: dict[str, int] = {}
        for i, celda in enumerate(encabezados or ()):
            clave = _norm(celda)
            if clave and clave not in idx:
                idx[clave] = i
        i_id = next((idx[h] for h in _HDR_ID if h in idx), None)
        i_cta = next((idx[h] for h in _HDR_CUENTA if h in idx), None)
        i_clabe = next((idx[h] for h in _HDR_CLABE if h in idx), None)  # opcional
        if i_id is None or i_cta is None:
            return catalogo  # no es el Excel esperado -> {}

        def col(fila, j):
            return fila[j] if j is not None and j < len(fila) else None

        for fila in filas:
            if not fila:
                continue
            id_emp = _id_empresa(col(fila, i_id))
            cuenta = _texto_cuenta(col(fila, i_cta))
            clabe = _texto_cuenta(col(fila, i_clabe)) if i_clabe is not None else ""
            if id_emp is None or not cuenta:
                continue
            registros = catalogo.setdefault(id_emp, [])
            if not any(r["cuenta"] == cuenta and r["clabe"] == clabe
                       for r in registros):  # sin duplicar por empresa
                registros.append({"cuenta": cuenta, "clabe": clabe})
    finally:
        wb.close()
    return catalogo


def _cargar(ruta: str) -> dict[int, list[dict]]:
    """Carga el catálogo; si el Excel se puede leer, actualiza el caché; si no,
    usa la última lectura guardada en caché."""
    catalogo = _leer_excel(ruta)
    if catalogo:
        try:  # el JSON exige claves str: se guardan como texto y se reconvierten
            with open(_RUTA_CACHE, "w", encoding="utf-8") as fh:
                json.dump({str(k): v for k, v in catalogo.items()}, fh,
                          ensure_ascii=False)
        except Exception:
            pass
        return catalogo
    if os.path.exists(_RUTA_CACHE):
        try:
            with open(_RUTA_CACHE, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            return {int(k): list(v) for k, v in data.items()}
        except Exception:
            pass
    return {}


def instalar_excel(ruta_origen: str) -> int:
    """Instala el Excel elegido en RUTA_EXCEL de forma TRANSACCIONAL y devuelve
    cuántas empresas quedaron con cuentas.

    Respalda el actual, copia el nuevo y lo LEE para validarlo; si no se reconoce
    (formato inesperado), hace ROLLBACK y lanza ExcelCuentasDispersionInvalido. Si
    es válido, invalida el caché."""
    import shutil

    os.makedirs(os.path.dirname(RUTA_EXCEL), exist_ok=True)
    respaldo = None
    if os.path.exists(RUTA_EXCEL):
        respaldo = RUTA_EXCEL + ".bak"
        shutil.copyfile(RUTA_EXCEL, respaldo)
    try:
        shutil.copyfile(ruta_origen, RUTA_EXCEL)
        catalogo = _leer_excel(RUTA_EXCEL)
        if not catalogo:
            raise ExcelCuentasDispersionInvalido(
                "El archivo no tiene el formato esperado (columnas 'id Empresa' y "
                "'Cuenta')."
            )
    except Exception:
        if respaldo is not None:
            shutil.copyfile(respaldo, RUTA_EXCEL)
        elif os.path.exists(RUTA_EXCEL):
            os.remove(RUTA_EXCEL)
        raise
    finally:
        if respaldo is not None and os.path.exists(respaldo):
            os.remove(respaldo)
    try:
        if os.path.exists(_RUTA_CACHE):
            os.remove(_RUTA_CACHE)
    except OSError:
        pass
    return len(catalogo)


class CatalogoCuentasDispersion:
    """Acceso al catálogo de cuentas de dispersión por id de empresa."""

    def __init__(self, ruta: str = RUTA_EXCEL):
        self.datos = _cargar(ruta)

    def disponible(self) -> bool:
        return bool(self.datos)

    def empresas(self) -> list[int]:
        return sorted(self.datos.keys())

    def _registros(self, id_empresa) -> list[dict]:
        if id_empresa is None:
            return []
        try:
            clave = int(id_empresa)
        except (TypeError, ValueError):
            return []
        return self.datos.get(clave, [])

    def cuentas_por_id_empresa(self, id_empresa) -> list[str]:
        """Cuentas ('Cuenta') de una empresa por su id. [] si no hay o el id es
        None. Ordenadas (alfabético) y sin duplicar, para el selector."""
        cuentas = {r.get("cuenta", "") for r in self._registros(id_empresa)}
        return sorted(c for c in cuentas if c)

    def clabes_por_id_empresa(self, id_empresa) -> list[str]:
        """CLABEs interbancarias VÁLIDAS de una empresa por su id (cuenta origen del
        TXT en pesos). Solo se incluyen las que son una CLABE de 18 dígitos con
        dígito de control correcto. [] si no hay o el id es None. Ordenadas y sin
        duplicar."""
        clabes = {r.get("clabe", "") for r in self._registros(id_empresa)}
        return sorted(c for c in clabes if _clabe_valida(c))

    def cuentas_clabe_por_id_empresa(self, id_empresa) -> list[tuple[str, str]]:
        """Pares (cuenta, clabe) de una empresa donde la CLABE es VÁLIDA y el banco
        de la cuenta TIENE formato de generación en la app (BANREGIO / BBVA /
        BANCOMER). La 'cuenta' es el texto a MOSTRAR (trae banco/empresa, evita
        confusiones) y la CLABE es el valor con el que opera el TXT en pesos. Sin
        duplicar por CLABE y ordenado por el texto de la cuenta."""
        vistos: set[str] = set()
        pares: list[tuple[str, str]] = []
        for r in self._registros(id_empresa):
            clabe = r.get("clabe", "")
            cuenta = r.get("cuenta", "")
            # Solo bancos con layout soportado (evita cuentas sin formato en la app).
            if not _clabe_valida(clabe) or banco_formato(cuenta) is None \
                    or clabe in vistos:
                continue
            vistos.add(clabe)
            pares.append((cuenta or clabe, clabe))  # fallback: la CLABE como texto
        return sorted(pares, key=lambda p: p[0].casefold())
