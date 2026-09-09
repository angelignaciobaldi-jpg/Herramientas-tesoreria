"""Lectores de los reportes de saldos que emite cada portal bancario.

Cada banco entrega lo suyo en un formato distinto —csv, xlsx, xls antiguo, txt de
ancho fijo o pdf— y con encabezados propios. Este módulo los normaliza todos a una
misma lista de `LineaSaldo`, para que el resto del sistema no tenga que saber de
qué banco viene cada archivo.

El banco se detecta **por el contenido**, no por el nombre del archivo: cada lector
declara la firma de su reporte (los encabezados que espera) y se prueban en orden.
Así un archivo renombrado sigue funcionando, que es lo habitual cuando el usuario
descarga varios el mismo día. El nombre solo se usa como pista de desempate.

Principio de diseño: **si un reporte no se reconoce, se avisa; nunca se devuelven
saldos en cero**. Un cero silencioso en un reporte de tesorería es peor que un
error: se firma como bueno.

Formatos cubiertos y de dónde sale el dato (verificado contra descargas reales):

    BANORTE     csv        CUENTA · CLABE · SALDO DISPONIBLE
    SANTANDER   csv        Cuenta · Disponible
    BANAMEX     csv        Sucursal+Cuenta · Saldo · Moneda
    BANREGIO    xlsx       Cuenta · Empresa · Disponible
    MULTIVA     xlsx       Cuenta · Divisa · Saldo
    BAJIO       xlsx       encabezado en la fila 7; Cuentas de vista · Saldo Disponible
    HSBC        xlsx       Número de cuenta · Actual disponible   (ver _filas_xlsx)
    BANCOMER    xls        Cuenta · Divisa · Disponible           (requiere xlrd)
    SCOTIABANK  txt        ancho fijo de 134 caracteres
    MONEX       pdf        Contrato · Clabe · Total en pesos
    SABADELL    pdf        sección 'Cuentas' · Saldo disponible (las Líneas de
                           Crédito NO son saldos)

    INBURSA     xlsx       CUENTA · SALDO DISPONIBLE (la divisa va en la portada)
"""

from __future__ import annotations

import csv
import os
import re
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field

_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

from .catalogo_bancos import banco_desde_clabe

try:
    import openpyxl
except ImportError:  # sin openpyxl no se pueden leer los reportes en xlsx
    openpyxl = None

try:
    import xlrd  # solo para el .xls antiguo de BANCOMER (BIFF/OLE2)
except ImportError:
    xlrd = None

try:
    import pymupdf
except ImportError:
    pymupdf = None


class ErrorLector(Exception):
    """No se pudo leer el reporte (formato no reconocido, archivo dañado…)."""


@dataclass
class LineaSaldo:
    """Un saldo tal como lo reporta el banco, ya normalizado.

    `cuenta` y `clabe` van en dígitos, sin recortar ceros a la izquierda: el casado
    posterior compara colas de dígitos y un cero perdido rompe la comparación.
    """

    banco: str            # nombre canónico del banco
    cuenta: str           # dígitos de la cuenta, tal como los reporta el portal
    clabe: str            # CLABE si el reporte la trae; "" si no
    titular: str          # nombre que aparece en el reporte del banco
    saldo: float
    moneda: str           # MXN | USD (u otra sigla, en mayúsculas)
    origen: str = ""      # archivo del que salió, para diagnóstico
    extra: dict = field(default_factory=dict)


# --------------------------------------------------------------- utilidades

def _norm(texto) -> str:
    """Minúsculas, sin acentos y con los espacios colapsados. Para comparar
    encabezados: los portales cambian tildes y espacios entre versiones."""
    plano = unicodedata.normalize("NFKD", str(texto or ""))
    plano = "".join(c for c in plano if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", plano).strip().lower()


def _digitos(valor) -> str:
    """Solo los dígitos de un valor. Quita el apóstrofo con que los portales
    fuerzan texto ('0502939411), guiones, espacios y máscaras."""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return re.sub(r"\D", "", str(valor or ""))


def _a_float(valor) -> float | None:
    """Monto a float. Devuelve None si no hay número (celda vacía, guion, texto).

    Acepta las DOS convenciones, porque conviven en la misma herramienta: los
    archivos que descargan los portales traen '$13,290,864.21' —punto decimal—,
    pero al COPIAR la tabla del portal de Inbursa lo que viaja es el texto ya
    formateado a la europea, '$483.582,81'. Leer eso con la regla del punto daba
    483.58: mil veces menos, y con toda la pinta de un saldo bueno. Un cero
    silencioso se nota; este no.

    La regla: cuando aparecen los dos separadores, el ÚLTIMO es el decimal. Con
    una sola coma, es decimal si le siguen uno o dos dígitos ('$0,00'); si le
    siguen tres es de millares ('1,234'). Con un solo punto se conserva el
    comportamiento de siempre —decimal—, que es como leen hoy todos los demás
    bancos y no hay motivo para moverlo."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    txt = str(valor).strip()
    negativo = txt.startswith("(") and txt.endswith(")")
    crudo = re.sub(r"[^\d.,\-]", "", txt)
    if not crudo:
        return None

    ult_coma, ult_punto = crudo.rfind(","), crudo.rfind(".")
    if ult_coma >= 0 and ult_punto >= 0:
        decimal = "," if ult_coma > ult_punto else "."
    elif ult_coma >= 0:
        decimales = len(crudo) - ult_coma - 1
        decimal = "," if decimales in (1, 2) else ""
    else:
        decimal = "."   # solo punto: como siempre

    if decimal == ",":
        limpio = crudo.replace(".", "").replace(",", ".")
    elif decimal == ".":
        limpio = crudo.replace(",", "")
    else:
        limpio = crudo.replace(",", "").replace(".", "")

    limpio = re.sub(r"[^\d.\-]", "", limpio)
    if not limpio or limpio in ("-", ".", "-."):
        return None
    try:
        n = float(limpio)
    except ValueError:
        return None
    return -abs(n) if negativo else n


def _moneda(valor, defecto: str = "MXN") -> str:
    """Sigla de moneda normalizada. Los portales escriben MXP, MN, MXN, PESOS…"""
    s = re.sub(r"[^A-Z]", "", str(valor or "").upper())
    if s in ("MXP", "MN", "MXN", "PESOS", "PESO", "MEXICANO"):
        return "MXN"
    if s in ("USD", "DLLS", "DLL", "DOLARES", "DOLAR"):
        return "USD"
    return s or defecto


def _texto_plano(ruta: str, limite: int = 4096) -> str:
    """Primeros caracteres de un archivo de texto, probando codificaciones. Se usa
    para detectar el banco por sus encabezados."""
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            with open(ruta, encoding=enc) as fh:
                return fh.read(limite)
        except (UnicodeDecodeError, LookupError):
            continue
        except OSError as exc:
            raise ErrorLector(f"No se pudo abrir «{os.path.basename(ruta)}»: {exc}")
    return ""


def _filas_csv(ruta: str) -> list[list[str]]:
    """Filas de un csv, con el delimitador deducido y tolerante a codificación."""
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            with open(ruta, encoding=enc, newline="") as fh:
                muestra = fh.read(4096)
                fh.seek(0)
                try:
                    dialecto = csv.Sniffer().sniff(muestra, delimiters=",;\t|")
                except csv.Error:
                    dialecto = csv.excel  # coma, el caso habitual
                return [f for f in csv.reader(fh, dialecto)
                        if any(x.strip() for x in f)]
        except UnicodeDecodeError:
            continue
    raise ErrorLector(f"No se pudo decodificar «{os.path.basename(ruta)}».")


def _filas_xlsx_crudo(ruta: str) -> list[list]:
    """Filas de un .xlsx leyendo el XML del zip con el parser de la stdlib.

    Existe porque los reportes de HSBC e INBURSA rompen openpyxl con
    `TypeError: expected <class 'openpyxl.styles.fills.Fill'>` — su hoja de estilos
    trae un relleno que openpyxl no acepta, y falla en TODAS las combinaciones de
    flags (read_only, data_only, rich_text). Los valores, en cambio, están intactos:
    basta con no mirar los estilos.

    Se usa ElementTree y NO expresiones regulares: el orden de los atributos de un
    elemento XML es arbitrario y de hecho cambia entre portales —HSBC escribe
    `<c r="A1" t="s">` e INBURSA `<c t="s" r="A1">`—, así que cualquier patrón que
    dé por sentado el orden lee bien un archivo y devuelve cero filas del otro.
    """
    try:
        with zipfile.ZipFile(ruta) as z:
            nombres = z.namelist()
            sst: list[str] = []
            if "xl/sharedStrings.xml" in nombres:
                raiz = ET.fromstring(z.read("xl/sharedStrings.xml"))
                for si in raiz.findall(f"{_NS}si"):
                    # Una cadena puede venir partida en varios <t> (texto con
                    # formato mezclado); se concatenan todos.
                    sst.append("".join(t.text or "" for t in si.iter(f"{_NS}t")))
            hojas = sorted(n for n in nombres
                           if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
            if not hojas:
                raise ErrorLector("El archivo no tiene hojas.")
            raiz = ET.fromstring(z.read(hojas[0]))
    except ErrorLector:
        raise
    except (zipfile.BadZipFile, ET.ParseError, KeyError, OSError) as exc:
        raise ErrorLector(
            f"No se pudo leer «{os.path.basename(ruta)}»: {exc}") from exc

    datos = raiz.find(f"{_NS}sheetData")
    if datos is None:
        return []
    filas: list[list] = []
    for fila in datos.findall(f"{_NS}row"):
        celdas: list = []
        for c in fila.findall(f"{_NS}c"):
            tipo = c.get("t")
            if tipo == "inlineStr":
                bloque = c.find(f"{_NS}is")
                valor = ("".join(t.text or "" for t in bloque.iter(f"{_NS}t"))
                         if bloque is not None else None)
            else:
                v = c.find(f"{_NS}v")
                valor = v.text if v is not None else None
                if tipo == "s" and valor is not None and valor.lstrip("-").isdigit():
                    i = int(valor)
                    valor = sst[i] if 0 <= i < len(sst) else ""
                elif tipo in (None, "n") and valor is not None:
                    valor = _numero_xml(valor)
            # Se respeta la COLUMNA declarada en la referencia: una celda vacía en
            # medio de la fila no debe correr las de la derecha.
            destino = _indice_columna(c.get("r") or "")
            if destino is None:
                celdas.append(valor)
                continue
            while len(celdas) < destino:
                celdas.append(None)
            celdas.append(valor)
        filas.append(celdas)
    return filas


def _indice_columna(ref: str) -> int | None:
    """Índice 0-based de la columna de una referencia tipo 'AB12'. None si no
    trae letras (algunas hojas omiten el atributo 'r')."""
    letras = re.match(r"([A-Z]+)", ref or "")
    if not letras:
        return None
    n = 0
    for ch in letras.group(1):
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _numero_xml(texto: str):
    """Valor numérico de una celda del XML: int si es entero, float si no.

    Se convierte para que este lector devuelva lo mismo que openpyxl; si dejara
    cadenas, un número de cuenta y un importe se comportarían distinto según por
    qué vía se leyó el archivo."""
    try:
        n = float(texto)
    except ValueError:
        return texto
    return int(n) if n.is_integer() else n


def _filas_xlsx(ruta: str) -> list[list]:
    """Filas de un .xlsx. Intenta openpyxl y cae al lector crudo si falla.

    No se usa `read_only=True`: con el reporte de BAJÍO devuelve solo la primera
    columna (su hoja declara mal las dimensiones), y ese modo no da ninguna ventaja
    con archivos de unas decenas de filas."""
    if openpyxl is not None:
        try:
            wb = openpyxl.load_workbook(ruta, data_only=True)
            try:
                ws = wb[wb.sheetnames[0]]
                return [list(f) for f in ws.iter_rows(values_only=True)]
            finally:
                wb.close()
        except Exception:  # noqa: BLE001 — se reintenta con el lector crudo
            pass
    return _filas_xlsx_crudo(ruta)


def _filas_xls(ruta: str) -> list[list]:
    """Filas de un .xls antiguo (BIFF/OLE2). Requiere xlrd."""
    if xlrd is None:
        raise ErrorLector(
            "Para leer los reportes .xls de BBVA hace falta la librería 'xlrd' "
            "(pip install xlrd).")
    try:
        wb = xlrd.open_workbook(ruta)
        ws = wb.sheet_by_index(0)
        return [[c.value for c in ws.row(r)] for r in range(ws.nrows)]
    except Exception as exc:  # noqa: BLE001 — se traduce a un error propio
        raise ErrorLector(
            f"No se pudo leer «{os.path.basename(ruta)}»: {exc}") from exc


def _texto_pdf(ruta: str) -> str:
    """Texto de un PDF. Los reportes bancarios traen capa de texto, así que no se
    invoca OCR: si algún día llega uno escaneado, `core.ocr.extraer_texto` es la
    vía, pero cuesta segundos por página y aquí no hace falta."""
    if pymupdf is None:
        raise ErrorLector("Falta PyMuPDF para leer los reportes en PDF.")
    try:
        with pymupdf.open(ruta) as doc:
            return "\n".join(p.get_text() for p in doc)
    except Exception as exc:  # noqa: BLE001 — se traduce a un error propio
        raise ErrorLector(
            f"No se pudo leer «{os.path.basename(ruta)}»: {exc}") from exc


def _indice_encabezado(fila: list, alias: dict[str, tuple]) -> dict[str, int]:
    """Mapea nombre lógico -> índice de columna, buscando por encabezado.

    `alias` es {nombre_logico: (variantes aceptadas, ya normalizadas)}. Se compara
    por igualdad y, si no, por 'empieza con', porque algunos portales le pegan
    unidades o notas al encabezado."""
    idx: dict[str, int] = {}
    normalizadas = [_norm(c) for c in fila]
    for logico, variantes in alias.items():
        for i, celda in enumerate(normalizadas):
            if celda in variantes or any(celda.startswith(v) for v in variantes):
                idx[logico] = i
                break
    return idx


def _buscar_encabezado(filas: list[list], alias: dict[str, tuple],
                       obligatorias: tuple, limite: int = 25):
    """Encuentra la fila de encabezados y devuelve `(nº de fila, índice)`.

    No se asume la fila 1: el reporte de BAJÍO trae seis filas de portada y sus
    encabezados empiezan en la 7. Se escanea hasta `limite` filas buscando la
    primera que contenga todas las columnas obligatorias."""
    for n, fila in enumerate(filas[:limite]):
        if not fila:
            continue
        idx = _indice_encabezado(fila, alias)
        if all(o in idx for o in obligatorias):
            return n, idx
    return None, {}


# ------------------------------------------------------------------ lectores
# Cada lector devuelve list[LineaSaldo] o lanza ErrorLector. La detección del
# banco va aparte, en _detectar(): así un lector se puede forzar a mano.

_ALIAS_BANORTE = {
    "cuenta": ("cuenta",),
    "titular": ("titular / personalizacion", "titular"),
    "moneda": ("moneda",),
    "clabe": ("clabe",),
    "saldo": ("saldo disponible",),
}


def leer_banorte(ruta: str, filas: list = None) -> list[LineaSaldo]:
    filas = filas if filas is not None else _filas_csv(ruta)
    n, idx = _buscar_encabezado(filas, _ALIAS_BANORTE, ("cuenta", "saldo"))
    if n is None:
        raise ErrorLector("No se encontraron los encabezados de BANORTE.")
    out = []
    for fila in filas[n + 1:]:
        cuenta = _digitos(fila[idx["cuenta"]]) if idx["cuenta"] < len(fila) else ""
        if not cuenta:
            continue
        clabe = _digitos(fila[idx["clabe"]]) if "clabe" in idx else ""
        saldo = _a_float(fila[idx["saldo"]])
        out.append(LineaSaldo(
            banco="Banorte", cuenta=cuenta, clabe=clabe,
            titular=str(fila[idx["titular"]] if "titular" in idx else "").strip(),
            saldo=saldo or 0.0,
            moneda=_moneda(fila[idx["moneda"]] if "moneda" in idx else "")))
    return out


_ALIAS_SANTANDER = {
    "cuenta": ("cuenta",),
    "titular": ("descripcion",),
    "saldo": ("disponible",),
}


def leer_santander(ruta: str, filas: list = None) -> list[LineaSaldo]:
    filas = filas if filas is not None else _filas_csv(ruta)
    n, idx = _buscar_encabezado(filas, _ALIAS_SANTANDER, ("cuenta", "saldo"))
    if n is None:
        raise ErrorLector("No se encontraron los encabezados de SANTANDER.")
    out = []
    for fila in filas[n + 1:]:
        cuenta = _digitos(fila[idx["cuenta"]]) if idx["cuenta"] < len(fila) else ""
        if not cuenta:
            continue
        out.append(LineaSaldo(
            banco="Santander", cuenta=cuenta, clabe="",
            titular=str(fila[idx["titular"]] if "titular" in idx else "").strip(),
            saldo=_a_float(fila[idx["saldo"]]) or 0.0,
            moneda="MXN"))  # el consolidado de cheques no trae columna de moneda
    return out


_ALIAS_BANAMEX = {
    "sucursal": ("sucursal",),
    "cuenta": ("cuenta",),
    "saldo": ("saldo",),
    "moneda": ("moneda",),
    "error": ("mensaje de error",),
}


def leer_banamex(ruta: str, filas: list = None) -> list[LineaSaldo]:
    filas = filas if filas is not None else _filas_csv(ruta)
    n, idx = _buscar_encabezado(filas, _ALIAS_BANAMEX, ("cuenta", "saldo"))
    if n is None:
        raise ErrorLector("No se encontraron los encabezados de BANAMEX.")
    out = []
    for fila in filas[n + 1:]:
        if idx["cuenta"] >= len(fila):
            continue
        cuenta = _digitos(fila[idx["cuenta"]])
        if not cuenta:
            continue
        # El número completo es sucursal + cuenta: así es como está en el catálogo
        # (p. ej. sucursal 394 + cuenta 7680454 -> 3947680454).
        sucursal = _digitos(fila[idx["sucursal"]]) if "sucursal" in idx else ""
        out.append(LineaSaldo(
            banco="Banamex", cuenta=sucursal + cuenta, clabe="", titular="",
            saldo=_a_float(fila[idx["saldo"]]) or 0.0,
            moneda=_moneda(fila[idx["moneda"]] if "moneda" in idx else ""),
            extra={"cuenta_corta": cuenta, "sucursal": sucursal,
                   "aviso": str(fila[idx["error"]]).strip()
                   if "error" in idx and idx["error"] < len(fila) else ""}))
    return out


_ALIAS_BANREGIO = {
    "alias": ("alias",),
    "cuenta": ("cuenta",),
    "titular": ("empresa",),
    "saldo": ("disponible",),
}


def leer_banregio(ruta: str, filas: list = None) -> list[LineaSaldo]:
    filas = filas if filas is not None else _filas_xlsx(ruta)
    n, idx = _buscar_encabezado(filas, _ALIAS_BANREGIO, ("cuenta", "saldo"))
    if n is None:
        raise ErrorLector("No se encontraron los encabezados de BANREGIO.")
    out = []
    for fila in filas[n + 1:]:
        if idx["cuenta"] >= len(fila):
            continue
        cuenta = _digitos(fila[idx["cuenta"]])
        if not cuenta:
            continue
        out.append(LineaSaldo(
            banco="Banregio", cuenta=cuenta, clabe="",
            titular=str(fila[idx["titular"]] or "").strip() if "titular" in idx else "",
            saldo=_a_float(fila[idx["saldo"]]) or 0.0,
            # El reporte no trae divisa; el catálogo distingue las cuentas en
            # dólares, así que la moneda se resuelve al casar, no aquí.
            moneda="MXN",
            extra={"alias": str(fila[idx["alias"]] or "").strip()
                   if "alias" in idx and idx["alias"] < len(fila) else ""}))
    return out


_ALIAS_MULTIVA = {
    "titular": ("empresa",),
    "cuenta": ("cuenta",),
    "alias": ("alias",),
    "moneda": ("divisa",),
    "saldo": ("saldo",),
}


def leer_multiva(ruta: str, filas: list = None) -> list[LineaSaldo]:
    filas = filas if filas is not None else _filas_xlsx(ruta)
    n, idx = _buscar_encabezado(filas, _ALIAS_MULTIVA, ("cuenta", "saldo"))
    if n is None:
        raise ErrorLector("No se encontraron los encabezados de MULTIVA.")
    out = []
    for fila in filas[n + 1:]:
        if idx["cuenta"] >= len(fila):
            continue
        cuenta = _digitos(fila[idx["cuenta"]])
        if not cuenta:
            continue
        out.append(LineaSaldo(
            banco="Multiva Banco", cuenta=cuenta, clabe="",
            titular=str(fila[idx["titular"]] or "").strip() if "titular" in idx else "",
            saldo=_a_float(fila[idx["saldo"]]) or 0.0,
            moneda=_moneda(fila[idx["moneda"]] if "moneda" in idx else "")))
    return out


_ALIAS_BAJIO = {
    "producto": ("tipo de producto",),
    "cuenta": ("cuentas de vista",),
    "titular": ("nombre del cliente",),
    "moneda": ("divisa",),
    "saldo": ("saldo disponible",),
}


def leer_bajio(ruta: str, filas: list = None) -> list[LineaSaldo]:
    filas = filas if filas is not None else _filas_xlsx(ruta)
    n, idx = _buscar_encabezado(filas, _ALIAS_BAJIO, ("cuenta", "saldo"))
    if n is None:
        raise ErrorLector("No se encontraron los encabezados de BAJÍO.")
    out = []
    for fila in filas[n + 1:]:
        if idx["cuenta"] >= len(fila):
            continue
        # El mismo reporte mezcla cuentas con líneas de crédito y tarjetas; solo las
        # cuentas son saldo disponible. Sumar una línea de crédito inflaría el
        # reporte con dinero que no existe.
        producto = _norm(fila[idx["producto"]]) if "producto" in idx else "cuenta"
        if producto and not producto.startswith("cuenta"):
            continue
        cuenta = _digitos(fila[idx["cuenta"]])
        if not cuenta:
            continue
        out.append(LineaSaldo(
            banco="Banco del Bajío", cuenta=cuenta, clabe="",
            titular=str(fila[idx["titular"]] or "").strip() if "titular" in idx else "",
            saldo=_a_float(fila[idx["saldo"]]) or 0.0,
            moneda=_moneda(fila[idx["moneda"]] if "moneda" in idx else "")))
    return out


_ALIAS_HSBC = {
    "moneda": ("moneda",),
    "cuenta": ("numero de cuenta",),
    "titular": ("nombre de cuenta",),
    "saldo": ("actual disponible",),
}


def leer_hsbc(ruta: str, filas: list = None) -> list[LineaSaldo]:
    filas = filas if filas is not None else _filas_xlsx(ruta)
    n, idx = _buscar_encabezado(filas, _ALIAS_HSBC, ("cuenta", "saldo"))
    if n is None:
        raise ErrorLector("No se encontraron los encabezados de HSBC.")
    out = []
    for fila in filas[n + 1:]:
        if idx["cuenta"] >= len(fila):
            continue
        cuenta = _digitos(fila[idx["cuenta"]])
        if not cuenta:
            continue
        out.append(LineaSaldo(
            banco="HSBC", cuenta=cuenta, clabe="",
            titular=str(fila[idx["titular"]] or "").strip() if "titular" in idx else "",
            saldo=_a_float(fila[idx["saldo"]]) or 0.0,
            moneda=_moneda(fila[idx["moneda"]] if "moneda" in idx else "")))
    return out


_ALIAS_INBURSA = {
    "cuenta": ("cuenta",),
    "titular": ("empresa",),
    "producto": ("producto",),
    "saldo": ("saldo disponible",),
}


def leer_inbursa(ruta: str, filas: list = None) -> list[LineaSaldo]:
    """Inbursa entrega un consolidado por divisa: la moneda no está en una columna
    sino en una línea de portada ('Divisa: PESOS'), y la última fila es el total
    (sin número de cuenta, así que se descarta sola)."""
    filas = filas if filas is not None else _filas_xlsx(ruta)
    n, idx = _buscar_encabezado(filas, _ALIAS_INBURSA, ("cuenta", "saldo"))
    if n is None:
        raise ErrorLector("No se encontraron los encabezados de INBURSA.")
    portada = _norm(" ".join(str(v) for f in filas[:n] for v in f if v))
    moneda = "USD" if ("dolar" in portada or "usd" in portada) else "MXN"
    out = []
    for fila in filas[n + 1:]:
        if idx["cuenta"] >= len(fila):
            continue
        cuenta = _digitos(fila[idx["cuenta"]])
        if not cuenta:
            continue
        out.append(LineaSaldo(
            banco="Inbursa", cuenta=cuenta, clabe="",
            titular=str(fila[idx["titular"]] or "").strip() if "titular" in idx else "",
            saldo=_a_float(fila[idx["saldo"]]) or 0.0,
            moneda=moneda,
            extra={"producto": str(fila[idx["producto"]] or "").strip()
                   if "producto" in idx and idx["producto"] < len(fila) else ""}))
    return out


_ALIAS_BANCOMER = {
    "cuenta": ("cuenta",),
    "alias": ("alias",),
    "moneda": ("divisa",),
    "saldo": ("disponible",),
}


def leer_bancomer(ruta: str, filas: list = None) -> list[LineaSaldo]:
    filas = filas if filas is not None else _filas_xls(ruta) if ruta.lower().endswith(".xls") else _filas_xlsx(ruta)
    n, idx = _buscar_encabezado(filas, _ALIAS_BANCOMER, ("cuenta", "saldo"))
    if n is None:
        raise ErrorLector("No se encontraron los encabezados de BBVA.")
    out = []
    for fila in filas[n + 1:]:
        if idx["cuenta"] >= len(fila):
            continue
        cuenta = _digitos(fila[idx["cuenta"]])
        # La última fila suele ser 'Totales'; sin dígitos de cuenta se descarta sola.
        if not cuenta:
            continue
        out.append(LineaSaldo(
            banco="BBVA México", cuenta=cuenta, clabe="",
            titular=str(fila[idx["alias"]] or "").strip() if "alias" in idx else "",
            saldo=_a_float(fila[idx["saldo"]]) or 0.0,
            moneda=_moneda(fila[idx["moneda"]] if "moneda" in idx else "")))
    return out


# SCOTIABANK exporta ancho fijo de 134 caracteres, sin encabezado:
#   0-2 producto (CHQ) · 3-5 moneda · 6-14 plaza y coma · 16-35 cuenta (20 dígitos)
#   36-85 nombre · 86-102 saldo · 103+ país y estatus
_SCOTIA = re.compile(
    r"^(?P<producto>[A-Z]{3})(?P<moneda>[A-Z]{3})(?P<plaza>[A-Z ]+),\s*"
    r"(?P<cuenta>\d{12,24})(?P<titular>.{1,60}?)\s*"
    r"(?P<saldo>\d{6,}\.\d{2})(?P<pais>[A-Za-z]+?)(?P<estatus>Activa|No existe.*)?\s*$")


# El comprobante «SEL - Scotia en Línea», que es un PDF y no el txt de ancho
# fijo. Se reconoce por su título, que ningún otro reporte lleva.
_ES_SEL = "comprobante de consulta de saldos"
# La cuenta va pegada a la plaza al final del token: 'MAZATLAN,-11700512613'.
_SEL_CUENTA = re.compile(r"[^\d](\d{8,})\s*$")
_SEL_MONEDA = re.compile(r"^[A-Z]{3}$")
# Cuántas líneas se admiten entre la cuenta y su moneda. El titular puede partirse
# en dos ('…DEL PACIFICO' / 'SA') y a veces no viene; más allá de eso es que el
# renglón no era una cuenta y hay que soltarlo en vez de tragarse media página.
_SEL_MAX_TITULAR = 5


def _leer_scotia_sel(texto: str) -> list[LineaSaldo]:
    """El comprobante de Scotia en Línea, cuyo PDF sale un dato por renglón:

        MEXICO / CHQ-MXN- / MAZATLAN,-11700512613 / ABASTECEDORA … / SA
        / MXN / 498,827.01 / ACTIVA

    Se ancla en la CUENTA —el número largo al final del token de la plaza— y de
    ahí se avanza hasta la moneda; lo que queda en medio es el titular, que puede
    venir partido en dos renglones o no venir (una de las cuentas del comprobante
    real no tiene nombre). El importe es el renglón siguiente a la moneda.

    No se lee por posición ni por columnas: el PDF no las conserva."""
    # Se trabaja sobre TOKENS, no sobre renglones: del PDF llega un dato por
    # línea, pero al copiar la MISMA tabla desde el portal llega un renglón por
    # cuenta con tabuladores. Partiendo también por tabulador, las dos formas
    # quedan iguales y el recorrido es el mismo.
    lineas = [t.strip() for l in texto.splitlines() for t in l.split("\t")
              if t.strip()]
    salida = []
    i = 0
    while i < len(lineas):
        m = _SEL_CUENTA.search(lineas[i])
        # 'Total por Producto: 1,442,932.90' también acaba en dígitos: se descarta
        # por el rótulo, no por la forma.
        if not m or _norm(lineas[i]).startswith(("total", "folio")):
            i += 1
            continue
        cuenta = m.group(1)
        titular, j = [], i + 1
        while (j < len(lineas) and j - i <= _SEL_MAX_TITULAR
               and not _SEL_MONEDA.match(lineas[j])):
            titular.append(lineas[j])
            j += 1
        if j >= len(lineas) or not _SEL_MONEDA.match(lineas[j]):
            i += 1
            continue
        saldo = _a_float(lineas[j + 1]) if j + 1 < len(lineas) else None
        if saldo is None:
            i += 1
            continue
        estatus = lineas[j + 2] if j + 2 < len(lineas) else ""
        salida.append(LineaSaldo(
            banco="Scotiabank", cuenta=cuenta, clabe="",
            titular=" ".join(titular).strip(), saldo=saldo,
            moneda=_moneda(lineas[j]),
            extra={"estatus": estatus.strip()}))
        i = j + 2
    if not salida:
        raise ErrorLector(
            "El comprobante de Scotia en Línea no trae ninguna cuenta legible.")
    return salida


def leer_scotiabank(ruta: str, texto: str = None) -> list[LineaSaldo]:
    """Scotiabank entrega DOS formatos distintos y aquí se reparten.

    El de siempre es un txt de ancho fijo; el nuevo es el comprobante «Scotia en
    Línea» en PDF, que no tiene columnas que respetar. Se distinguen por el
    título del comprobante, así que un archivo no puede confundirse con el
    otro."""
    if texto is None:
        texto = (_texto_pdf(ruta) if ruta.lower().endswith(".pdf")
                 else _texto_plano(ruta, limite=1_000_000))
    # Se prueban los DOS y gana el que saque cuentas. Antes mandaba solo el
    # título del comprobante, y al PEGAR desde el portal ese título no siempre
    # viaja: el pegado se iba al lector de ancho fijo y moría con «ninguna línea
    # tiene el formato de SCOTIABANK» aun habiéndose identificado el banco.
    if _ES_SEL in _norm(texto):
        try:
            return _leer_scotia_sel(texto)
        except ErrorLector:
            pass   # traía el título pero no cuentas: se intenta el ancho fijo
    out = []
    for linea in texto.splitlines():
        if not linea.strip():
            continue
        m = _SCOTIA.match(linea.rstrip())
        if not m:
            continue
        cuenta = m.group("cuenta").lstrip("0")
        if not cuenta:
            continue
        out.append(LineaSaldo(
            banco="Scotiabank", cuenta=cuenta, clabe="",
            titular=m.group("titular").strip(),
            saldo=_a_float(m.group("saldo")) or 0.0,
            moneda=_moneda(m.group("moneda")),
            extra={"estatus": (m.group("estatus") or "").strip()}))
    if out:
        return out
    # El de ancho fijo no reconoció nada: el último intento es el comprobante,
    # que es como llega cuando se PEGA la tabla del portal —ahí no viene el
    # título por el que se distinguen—. Si tampoco es, sube su error, que dice
    # qué se esperaba encontrar.
    return _leer_scotia_sel(texto)


_MONEX_CLIENTE = re.compile(r"Cliente\s*:\s*(.+)")
_MONEX_CONTRATO = re.compile(r"Contrato\s*:\s*(\d+)")
_MONEX_CLABE = re.compile(r"Clabe\s*:\s*(\d{18})")
_MONEX_TOTAL = re.compile(r"Total en pesos\s*\n\s*([\d,.]+)")
# El efectivo en pesos y el saldo neto en dólares, para los contratos que traen
# las dos divisas. El comprobante los presenta en bloques separados —«Saldo de
# efectivo» y «Saldos en divisas»— y el formato les reserva DOS renglones.
_MONEX_EFECTIVO = re.compile(r"Saldo de efectivo\s*\n\s*([\d,.]+)")
# El renglón del dólar trae nueve importes seguidos —MD, 24hrs, 48hrs, >48hrs,
# tránsito, bloqueado, neto, tipo de cambio y valuación en pesos— y el que quiere
# tesorería es el PRIMERO, «Saldo MD». Por eso la expresión se queda con el
# primero y no busca 'Saldo neto': hoy coinciden porque las columnas intermedias
# van en cero, pero en cuanto haya dinero en tránsito dejarían de hacerlo.
_MONEX_USD = re.compile(r"DOLAR\s*\n?\s*AMERICANO\s*\n?\s*([\d,.]+)")


def leer_monex(ruta: str, texto: str = None) -> list[LineaSaldo]:
    """Monex emite un PDF por contrato, no un consolidado.

    Un contrato con dólares da DOS líneas, porque el formato lo parte en dos
    renglones: uno de pesos y otro de divisas (SALDOS!H34 y H94 dicen «DLS»).
    Antes se emitía solo el 'Total en pesos' —que ya incluye la valuación de los
    dólares—, así que el renglón de divisas quedaba vacío y los dólares se
    contaban como pesos: de las seis cuentas de Monex solo se llenaban cuatro.

    Al renglón de pesos le toca el SALDO DE EFECTIVO y al de divisas el «Saldo
    MD» EN DÓLARES, no su valuación en pesos: el reporte lo rotula DLS y en pesos
    saldría diecisiete veces mayor. En los contratos de una sola divisa
    'Saldo de efectivo' y 'Total en pesos' coinciden (verificado sobre los cuatro
    comprobantes), así que ahí no cambia nada."""
    texto = texto if texto is not None else _texto_pdf(ruta)
    contrato = _MONEX_CONTRATO.search(texto)
    clabe = _MONEX_CLABE.search(texto)
    total = _MONEX_TOTAL.search(texto)
    if not (contrato or clabe):
        raise ErrorLector("El PDF no parece un reporte de saldos de MONEX.")
    if total is None:
        raise ErrorLector(
            f"No se encontró el 'Total en pesos' en "
            f"«{os.path.basename(ruta)}».")
    cliente = _MONEX_CLIENTE.search(texto)
    num = contrato.group(1) if contrato else ""
    nombre = cliente.group(1).strip() if cliente else ""

    usd = _MONEX_USD.search(texto)
    efectivo = _MONEX_EFECTIVO.search(texto)
    # Sin bloque de divisas, efectivo y total coinciden: se prefiere el total por
    # ser el que este lector ha usado siempre.
    pesos = _a_float((efectivo or total).group(1)) if usd else _a_float(
        total.group(1))

    salida = [LineaSaldo(
        banco="Banco Monex", cuenta=num,
        clabe=clabe.group(1) if clabe else "",
        titular=nombre, saldo=pesos or 0.0, moneda="MXN")]
    if usd:
        salida.append(LineaSaldo(
            banco="Banco Monex", cuenta=num, clabe="", titular=nombre,
            saldo=_a_float(usd.group(1)) or 0.0, moneda="USD",
            # La marca es lo que permite mandarlo a su renglón: comparte contrato
            # con el de pesos, así que por número no hay forma de separarlos.
            extra={"renglon_divisa": True}))
    return salida


def leer_sabadell(ruta: str) -> list[LineaSaldo]:
    """Sabadell entrega la 'Posición Global': cuentas primero, líneas de crédito
    después. Solo la primera sección son saldos — una línea de crédito es dinero
    disponible para pedir prestado, no dinero en la cuenta.

    Se toma **Saldo disponible**, confirmado con tesorería. Ojo si alguien compara
    con el formato viejo: su fórmula apuntaba a SABADELL!E9, que por los
    encabezados de esa hoja parecería ser 'Saldo por aplicar'. Es un espejismo del
    pegado manual —la columna quedó recorrida—, no el criterio real."""
    texto = _texto_pdf(ruta)
    if "Posición Global" not in texto and "Saldo disponible" not in texto:
        raise ErrorLector("El PDF no parece un reporte de SABADELL.")
    lineas = [x.strip() for x in texto.splitlines()]
    # La sección de cuentas termina donde empiezan las líneas de crédito.
    try:
        fin = next(i for i, x in enumerate(lineas) if "Líneas de Crédito" in x)
    except StopIteration:
        fin = len(lineas)
    out = []
    for i, x in enumerate(lineas[:fin]):
        # Una cuenta es una corrida de dígitos seguida de la moneda y dos importes.
        if not re.fullmatch(r"\d{6,}", x):
            continue
        if i + 3 >= fin:
            continue
        moneda, por_aplicar, disponible = lineas[i + 1], lineas[i + 2], lineas[i + 3]
        if not re.fullmatch(r"[A-Z]{3}", moneda):
            continue
        saldo = _a_float(disponible)
        if saldo is None:
            continue
        out.append(LineaSaldo(
            banco="Banco Sabadell", cuenta=x.lstrip("0") or x, clabe="",
            titular=lineas[i - 1].strip() if i else "",
            saldo=saldo, moneda=_moneda(moneda),
            extra={"saldo_por_aplicar": _a_float(por_aplicar)}))
    if not out:
        raise ErrorLector(
            f"No se encontraron cuentas en «{os.path.basename(ruta)}».")
    return out


# --------------------------------------------------------------- detección
# (nombre, extensiones, marcas que deben aparecer en el texto, lector)
# El orden importa: se prueba de la marca más específica a la más genérica.
_LECTORES = (
    ("BANORTE", (".csv",), ("titular / personalizacion", "saldo disponible"),
     leer_banorte),
    ("SANTANDER", (".csv",), ("descripcion", "disponible", "sbc"), leer_santander),
    ("BANAMEX", (".csv",), ("tipo de cuenta", "sucursal", "mensaje de error"),
     leer_banamex),
    ("HSBC", (".xlsx",), ("numero de cuenta", "actual disponible"), leer_hsbc),
    ("BAJIO", (".xlsx",), ("cuentas de vista", "saldo disponible"), leer_bajio),
    ("INBURSA", (".xlsx",), ("saldo consolidado", "salvo buen cobro"),
     leer_inbursa),
    ("BANREGIO", (".xlsx",), ("alias", "empresa", "disponible", "en transito"),
     leer_banregio),
    ("MULTIVA", (".xlsx",), ("empresa", "cuenta", "alias", "divisa", "saldo"),
     leer_multiva),
    ("BANCOMER", (".xls", ".xlsx"), ("cuenta", "alias", "divisa", "disponible"),
     leer_bancomer),
    ("SCOTIABANK", (".txt", ".pdf"), ("chq",), leer_scotiabank),
    ("MONEX", (".pdf",), ("contrato", "clabe"), leer_monex),
    ("SABADELL", (".pdf",), ("posicion global", "saldo disponible"), leer_sabadell),
)


def _huella(ruta: str) -> str:
    """Texto representativo del archivo, normalizado, para detectar el banco."""
    ext = os.path.splitext(ruta)[1].lower()
    try:
        if ext in (".csv", ".txt"):
            return _norm(_texto_plano(ruta))
        if ext == ".pdf":
            return _norm(_texto_pdf(ruta)[:4000])
        if ext == ".xls":
            filas = _filas_xls(ruta)[:12]
        else:
            filas = _filas_xlsx(ruta)[:15]
        return _norm(" ".join(str(v) for f in filas for v in f if v is not None))
    except ErrorLector:
        raise
    except Exception:  # noqa: BLE001 — un archivo ilegible se reporta al detectar
        return ""


def es_temporal(ruta: str) -> bool:
    """Archivo de bloqueo de Office (`~$Reporte.xlsx`), no un reporte.

    Word y Excel crean uno junto a cada libro abierto. Aparecen en el diálogo de
    archivos y se cuelan con un 'seleccionar todo', pero ni siquiera se pueden
    abrir: Windows los tiene bloqueados."""
    return os.path.basename(ruta).startswith("~$")


def detectar(ruta: str) -> str | None:
    """Nombre del banco cuyo reporte parece ser `ruta`, o None.

    Se decide por el CONTENIDO. El nombre del archivo solo desempata cuando dos
    firmas encajan (p. ej. BANREGIO y MULTIVA comparten 'empresa/cuenta/alias').

    Devuelve None también cuando el archivo no se puede ni abrir. La firma dice
    `str | None` y dejar escapar la excepción la rompía: quien solo quiere saber
    de qué banco es un archivo no espera tener que atrapar nada."""
    if es_temporal(ruta):
        return None
    ext = os.path.splitext(ruta)[1].lower()
    try:
        huella = _huella(ruta)
    except ErrorLector:
        return None
    if not huella:
        return None
    candidatos = [(n, m) for n, exts, m, _ in _LECTORES
                  if ext in exts and all(x in huella for x in m)]
    if not candidatos:
        return None
    if len(candidatos) > 1:
        nombre_archivo = _norm(os.path.basename(ruta))
        for nombre, _ in candidatos:
            if _norm(nombre) in nombre_archivo:
                return nombre
        # Sin pista en el nombre, gana la firma más específica (más marcas).
        candidatos.sort(key=lambda c: -len(c[1]))
    return candidatos[0][0]


_POR_NOMBRE = {n: f for n, _, _, f in _LECTORES}

# --------------------------------------------------------------- pegado
# Hay bancos cuyo reporte tesorería no descarga: selecciona la tabla en el portal
# y la pega. Lo que Windows deja en el portapapeles es la misma tabla en texto
# —renglones por saltos de línea, columnas por tabulador—, así que se convierte a
# filas y se le da a los MISMOS lectores. No hay lógica de banco duplicada: solo
# cambia de dónde salen las filas.
#
# Los lectores de PDF (Monex, Sabadell) quedan fuera a propósito: no leen filas
# sino un texto con una disposición concreta, y un pegado la pierde.
# Sabadell queda fuera del pegado: su lector busca secciones dentro del texto de
# un PDF con una disposición concreta que al copiar se pierde. Monex SÍ entra:
# su lector ya trabaja sobre texto con expresiones regulares, así que le da igual
# venir de un PDF o del portapapeles.
_SIN_PEGADO = ("SABADELL",)


# --------------------------------------------------- recetas de pegado
# Lo que se copia del portal NO tiene la forma del archivo que ese mismo portal
# descarga, y a veces no tiene forma de tabla siquiera. Bancoppel entrega una
# línea de «etiqueta: valor»; el consolidado de HSBC trae un encabezado que no
# corresponde a las columnas de sus renglones. Con la regla de los archivos, esos
# pegados o no se reconocen o —peor— los reclama otro banco: el de Intercam se
# detectaba como BANORTE y su saldo habría acabado en la pestaña equivocada.
#
# Por eso estas RECETAS van aparte y se prueban ANTES: cada una exige marcas muy
# suyas, así que solo reclaman lo que es suyo. Lo que ninguna reconozca sigue
# cayendo en la regla de siempre, que no se toca.

def _celda(fila: list, i):
    """El valor de una columna, o None si la fila se queda corta.

    Los pegados traen filas de largo desigual —el portal recorta las columnas
    vacías del final—, así que indexar a secas revienta."""
    if i is None or i >= len(fila):
        return None
    return fila[i]


def _celdas(filas: list) -> list[str]:
    """Todas las celdas del pegado, normalizadas y sin las vacías."""
    return [_norm(v) for f in filas for v in f if _norm(v)]


def _texto_de(filas: list) -> str:
    """El pegado como texto plano, para las recetas que no miran celdas."""
    return "\n".join(" ".join(str(v or "") for v in fila) for fila in filas)


def _marca_bancoppel(filas: list) -> bool:
    """Un rótulo «Cuenta» seguido de su número y un «Saldo» seguido de su importe.

    Se mira el TEXTO y no las celdas: el pegado llega con tabuladores o con
    espacios según de dónde se copie, y exigiendo celdas exactas se perdía la
    mitad de los casos. Que AMBOS rótulos vengan seguidos de un valor es lo que
    impide reclamar cualquier tabla con columnas «Cuenta» y «Saldo»: ahí a un
    rótulo le sigue otro rótulo, no un número."""
    texto = _texto_de(filas)
    return bool(_BCP_CUENTA.search(texto) and _BCP_SALDO.search(texto))


# Los rótulos de Bancoppel y su valor, en el TEXTO. Se busca así y no por celdas
# porque el pegado llega de formas distintas según de dónde se copie: a veces los
# rótulos y los valores caen en la misma fila separados por tabuladores, y a
# veces el portal los manda en dos renglones —rótulos arriba, valores abajo—. Con
# una expresión sobre el texto da igual cuál de las dos sea.
_BCP_CUENTA = re.compile(r"cuenta\s*:?\s*([\d][\d\s-]{6,})", re.I)
_BCP_CLABE = re.compile(r"clabe\s*:?\s*([\d][\d\s-]{10,})", re.I)
_BCP_SALDO = re.compile(r"saldo\s*:?\s*\$?\s*([\d][\d.,]*)", re.I)


def _leer_bancoppel_pegado(filas: list) -> list[LineaSaldo]:
    """Bancoppel copia una línea de etiquetas y valores, no una tabla.

        Cuenta:  22000004794  l  CLABE:  137180220000047940    Saldo:  $170,350.43

    Se lee por ETIQUETA y sobre el TEXTO, no por celdas: entre «Cuenta:» y
    «CLABE:» el portal mete una celda suelta con una 'l', así que contar columnas
    se rompe; y según de dónde se copie, los rótulos y sus valores caen en la
    misma fila o en dos renglones distintos. Buscando el valor que sigue a cada
    rótulo, da igual cuál de las dos formas llegue."""
    texto = "\n".join(" ".join(str(v or "") for v in fila) for fila in filas)
    cuenta = _BCP_CUENTA.search(texto)
    saldo = _BCP_SALDO.search(texto)
    if not cuenta or not saldo:
        raise ErrorLector(
            "El pegado de BANCOPPEL no trae «Cuenta:» y «Saldo:» con sus "
            "valores. Copia el renglón completo de la cuenta.")
    clabe = _BCP_CLABE.search(texto)
    monto = _a_float(saldo.group(1))
    if monto is None:
        raise ErrorLector("No se entendió el saldo del pegado de BANCOPPEL.")
    return [LineaSaldo(
        banco="BanCoppel", cuenta=_digitos(cuenta.group(1)),
        clabe=_digitos(clabe.group(1)) if clabe else "", titular="",
        saldo=monto, moneda="MXN")]


# Una cuenta enmascarada de Intercam: '***-***94-001-1'. Es su rasgo más
# distintivo y aparece igual en las dos vistas de su portal.
#
# El GUION es imprescindible en el patrón: Banamex también enmascara sus cuentas
# —'**8363'— y sin exigirlo esta receta le reclamaría sus pegados.
_INTERCAM_CUENTA = re.compile(r"\*{2,}[\d*]*-[\d*\-]*\d")


def _marca_intercam(filas: list) -> bool:
    """Reconoce las DOS vistas del portal, que no traen las mismas columnas.

    Una lista los saldos con «Cta Anterior · Cuenta · Moneda · Alias · Saldo
    Disponible»; la otra —la que se imprime a PDF— con «Cuenta · Disponible ·
    Sobregiro · Bloqueado…». Lo común a ambas es la cuenta enmascarada, que
    ningún otro banco usa, así que basta con exigirla junto a alguna de las dos
    cabeceras."""
    # La cuenta enmascarada CON guiones no la produce ningún otro portal, así
    # que por sí sola alcanza: pedir además una cabecera concreta dejaba fuera
    # la vista que se copia sin tabuladores, donde los títulos van en una misma
    # celda y no casan uno a uno.
    return bool(_INTERCAM_CUENTA.search(_texto_de(filas)))


def _leer_intercam_pegado(filas: list) -> list[LineaSaldo]:
    """Intercam: tabla con «Cta Anterior · Cuenta · Moneda · Alias · Saldo…».

    Su número viene ENMASCARADO (`***-***94-001-1`) y así está también en el
    formato, que guarda el mismo texto. Los dígitos que quedan a la vista bastan
    para casarlo, que es como se casa cualquier cuenta enmascarada.

    Ojo con el alias: el portal lo rotula «CUENTA ENLACE KAPITAL», y por eso
    tesorería llama Kapital a este pegado; la pestaña del formato es INTERCAM."""
    alias = {"cuenta": ("cuenta",), "moneda": ("moneda",),
             "titular": ("alias",), "saldo": ("saldo disponible",)}
    n, idx = _buscar_encabezado(filas, alias, ("cuenta", "saldo"))
    if n is None:
        # La otra vista del portal no tiene columna «Saldo Disponible» sino
        # «Disponible», y mete la cuenta y su alias en la MISMA celda. Ahí no hay
        # encabezado que casar: se ancla en la cuenta enmascarada y se toma el
        # primer importe que le sigue, que es el disponible.
        return _leer_intercam_por_mascara(filas)
    salida = []
    for fila in filas[n + 1:]:
        cuenta = _digitos(_celda(fila, idx.get("cuenta")))
        saldo = _a_float(_celda(fila, idx.get("saldo")))
        if not cuenta or saldo is None:
            continue
        salida.append(LineaSaldo(
            banco="Intercam Banco", cuenta=cuenta, clabe="",
            titular=str(_celda(fila, idx.get("titular")) or "").strip(),
            saldo=saldo,
            moneda=_moneda(_celda(fila, idx.get("moneda")))))
    return salida


def _leer_intercam_por_mascara(filas: list) -> list[LineaSaldo]:
    """Intercam sin encabezado reconocible: la cuenta enmascarada manda.

    Es la vista que se imprime a PDF, donde la celda de la cuenta trae también su
    alias («***-***94-001-1 CUENTA ENLACE KAPITAL PESOS») y el disponible es el
    primer importe del renglón."""
    salida = []
    for fila in filas:
        valores = [str(v or "").strip() for v in fila]
        crudo = " ".join(valores)
        m = _INTERCAM_CUENTA.search(crudo)
        if not m:
            continue
        importe = next((_a_float(v) for v in valores
                        if not _INTERCAM_CUENTA.search(v)
                        and _a_float(v) is not None), None)
        if importe is None:
            # Cuenta y disponible pueden venir en la misma celda.
            resto = crudo[m.end():]
            hallado = re.search(r"\$?\s*([\d][\d.,]*)", resto)
            importe = _a_float(hallado.group(1)) if hallado else None
        if importe is None:
            continue
        titular = crudo[m.end():].strip()
        salida.append(LineaSaldo(
            banco="Intercam Banco", cuenta=_digitos(m.group(0)), clabe="",
            titular=re.sub(r"[\d$.,]+", " ", titular).strip()[:60],
            saldo=importe, moneda="MXN"))
    if not salida:
        raise ErrorLector(
            "No se encontró la cuenta enmascarada de INTERCAM con su saldo "
            "disponible.")
    return salida


def _marca_hsbc_pegado(filas: list) -> bool:
    celdas = _celdas(filas)
    return ("actual disponible" in celdas
            and any(c.startswith("disponible en libros") for c in celdas))


def _leer_hsbc_pegado(filas: list) -> list[LineaSaldo]:
    """El consolidado de HSBC copiado: el encabezado NO manda.

    Sus cuatro títulos («Actual disponible», «Disponible en libros»…) no se
    corresponden con las columnas de los renglones, que llegan así:

        (vacío)   4056511132   OPERADORA DE REC HUM   76,089.13

    Así que se lee por FORMA, no por columna: la celda que es puro número largo
    es la cuenta, la última que parece importe es el saldo y lo de en medio el
    titular. La divisa sale de la línea de sección —«Mexico HBMI (MXN - …)»— que
    encabeza cada bloque, y los renglones de subtotal se descartan: repiten un
    importe que ya está contado."""
    divisa = "MXN"
    salida = []
    for fila in filas:
        valores = [str(v or "").strip() for v in fila]
        crudo = " ".join(valores)
        seccion = re.search(r"\(([A-Za-z]{3})\s*-", crudo)
        if seccion and not any(_ES_CUENTA_HSBC.fullmatch(v) for v in valores):
            divisa = _moneda(seccion.group(1))
            continue
        primera = next((v for v in valores if v.strip()), "")
        if _norm(primera).startswith(("subtotal", "total")):
            continue
        cuenta = next((v for v in valores if _ES_CUENTA_HSBC.fullmatch(v)), "")
        if not cuenta:
            continue
        importes = [v for v in valores
                    if v is not cuenta and _a_float(v) is not None
                    and not _ES_CUENTA_HSBC.fullmatch(v)]
        if not importes:
            continue
        titular = next((v for v in valores
                        if v.strip() and v is not cuenta and v not in importes
                        and _a_float(v) is None), "")
        salida.append(LineaSaldo(
            banco="HSBC", cuenta=_digitos(cuenta), clabe="",
            titular=titular.strip(), saldo=_a_float(importes[-1]) or 0.0,
            moneda=divisa))
    return salida


# Una cuenta de HSBC en el pegado: solo dígitos y de 8 en adelante. Sirve para
# distinguirla del importe, que siempre trae separadores o decimales.
_ES_CUENTA_HSBC = re.compile(r"\d{8,}")

# (nombre de la pestaña, reconoce, lee). El nombre es el de la HOJA del formato,
# que es lo que espera el resto del sistema.
_RECETAS_PEGADO = (
    ("INTERCAM", _marca_intercam, _leer_intercam_pegado),
    ("HSBC", _marca_hsbc_pegado, _leer_hsbc_pegado),
    ("BANCOPPEL", _marca_bancoppel, _leer_bancoppel_pegado),
)
_POR_RECETA = {n: f for n, _m, f in _RECETAS_PEGADO}


def _receta_de(filas: list) -> str | None:
    """Nombre de la receta de pegado que reconoce estas filas, si alguna."""
    for nombre, reconoce, _leer in _RECETAS_PEGADO:
        try:
            if reconoce(filas):
                return nombre
        except Exception:  # noqa: BLE001 — una receta rota no tumba las demás
            continue
    return None

# Lectores que aceptan las filas ya separadas. Se listan aparte de `_POR_NOMBRE`
# —que los tiene todos— porque las firmas no son iguales: Scotiabank recibe el
# texto crudo y los de PDF ni siquiera participan. Un despacho único obligaría a
# llamarlos a todos con los mismos argumentos, que es justo lo que no se puede.
# Columnas que declara cada lector, para reconocer un pegado por su ENCABEZADO.
# Es distinto de la firma de `_LECTORES`: aquella se calcula sobre el archivo
# entero y puede apoyarse en textos de portada —Inbursa se reconoce por «saldo
# consolidado», que vive en el título y NO viaja al copiar la tabla—. Un pegado
# es justo un encabezado y sus renglones, así que se le pregunta a las columnas.
_COLUMNAS_POR_BANCO = {}


_PEGABLES_POR_FILAS = {
    "BANORTE": leer_banorte,
    "SANTANDER": leer_santander,
    "BANAMEX": leer_banamex,
    "BANREGIO": leer_banregio,
    "MULTIVA": leer_multiva,
    "BAJIO": leer_bajio,
    "HSBC": leer_hsbc,
    "INBURSA": leer_inbursa,
    "BANCOMER": leer_bancomer,
}


def filas_pegadas(texto: str) -> list[list[str]]:
    """Convierte en filas el texto de una tabla copiada.

    El tabulador es lo que ponen Excel, los navegadores y el propio Windows al
    copiar una tabla, así que manda. Si no hay ninguno se prueba el punto y coma.

    La COMA no se usa nunca como separador, aunque un pegado suelto pueda venir
    como CSV: los importes la llevan dentro —`$170,350.43`— y partir por ella
    convertía ese saldo en 170. Un separador de más rompe la cifra en silencio;
    uno de menos solo deja el renglón en una columna, que las recetas manejan.

    Las líneas en blanco se descartan: al seleccionar con el mouse suelen colarse
    al principio y al final, y correrían el índice del encabezado."""
    lineas = [l for l in (texto or "").splitlines() if l.strip()]
    if not lineas:
        return []
    for sep in ("\t", ";"):
        if any(sep in l for l in lineas):
            return [l.split(sep) for l in lineas]
    # Una sola columna: sigue siendo válido para los lectores que trabajan sobre
    # texto (Scotiabank) y para que la detección pueda al menos intentarlo.
    return [[l] for l in lineas]


def detectar_pegado(filas: list) -> str | None:
    """Banco cuya firma casa con una tabla pegada, o None.

    Es `detectar` sin el filtro de extensión —un pegado no tiene archivo, así que
    tampoco hay nombre del que sacar pistas—. Cuando dos firmas casan gana la más
    específica, igual que ahí; el resto lo decide el usuario en pantalla, que por
    eso ve el banco detectado y puede corregirlo."""
    # Las recetas van PRIMERO. No es una preferencia estética: la regla de los
    # archivos llega a reclamar pegados que no son suyos —el de Intercam lo
    # tomaba por BANORTE— y ahí el saldo acaba en otra pestaña sin que nadie lo
    # note. Las recetas exigen marcas muy específicas, así que solo se llevan lo
    # que de verdad reconocen; el resto sigue cayendo en la regla de siempre.
    receta = _receta_de(filas or [])
    if receta:
        return receta

    huella = _norm(" ".join(str(v) for f in (filas or [])[:15] for v in f
                            if v is not None))
    if not huella:
        return None
    candidatos = [(n, m) for n, _exts, m, _f in _LECTORES
                  if n not in _SIN_PEGADO and all(x in huella for x in m)]
    if candidatos:
        candidatos.sort(key=lambda c: -len(c[1]))
        return candidatos[0][0]
    # La firma no alcanzó: se pregunta a las columnas del encabezado, que es lo
    # único que un pegado trae con seguridad.
    return _por_encabezado(filas)


_COLUMNAS_POR_BANCO.update({
    "BANORTE": _ALIAS_BANORTE, "SANTANDER": _ALIAS_SANTANDER,
    "BANAMEX": _ALIAS_BANAMEX, "BANREGIO": _ALIAS_BANREGIO,
    "MULTIVA": _ALIAS_MULTIVA, "BAJIO": _ALIAS_BAJIO, "HSBC": _ALIAS_HSBC,
    "INBURSA": _ALIAS_INBURSA, "BANCOMER": _ALIAS_BANCOMER,
})
# Todos los lectores de filas exigen lo mismo para poder trabajar.
_OBLIGATORIAS = ("cuenta", "saldo")


def _por_encabezado(filas: list) -> str | None:
    """Banco cuyo ENCABEZADO casa mejor con lo pegado.

    Gana el que reconozca MÁS columnas: casi todos exigen solo «cuenta» y
    «saldo», así que quedarse con el primero que cumple el mínimo elegiría mal.
    Con el pegado de Inbursa, por ejemplo, Banamex también cumple el mínimo pero
    reconoce dos columnas frente a las cuatro de Inbursa."""
    mejor, mejor_n = None, 0
    for nombre, alias in _COLUMNAS_POR_BANCO.items():
        n, idx = _buscar_encabezado(filas, alias, _OBLIGATORIAS)
        if n is not None and len(idx) > mejor_n:
            mejor, mejor_n = nombre, len(idx)
    return mejor


def bancos_pegables() -> list[str]:
    """Pestañas del formato cuyo saldo se puede pegar, para ofrecerlas en
    pantalla.

    Incluye las de receta propia (Bancoppel, Intercam) y las que leen texto
    (Scotiabank, Monex), no solo las que casan por encabezado. Faltaban justo
    esas y por eso no había forma de cargarlas a mano cuando la detección
    fallaba."""
    return sorted(set(_PEGABLES_POR_FILAS) | set(_POR_RECETA)
                  | {"SCOTIABANK", "MONEX"})


def leer_pegado(texto: str, banco: str = None) -> tuple[list[LineaSaldo], str]:
    """Lee una tabla copiada del portal. Devuelve `(líneas, banco)`.

    `banco` fuerza el lector cuando la detección no acierta, que es el caso que
    la pantalla ofrece resolver a mano."""
    filas = filas_pegadas(texto)
    if not filas:
        raise ErrorLector("No hay nada que pegar.")
    nombre = banco or detectar_pegado(filas)
    if nombre is None:
        raise ErrorLector(
            "No se reconoce de qué banco es lo que pegaste. Asegúrate de haber "
            "copiado también el renglón de encabezados, o elige el banco a mano.")
    if nombre in _SIN_PEGADO:
        raise ErrorLector(
            "{} solo se puede cargar como archivo: su reporte es un PDF y al "
            "pegarlo se pierde la disposición que el lector necesita.".format(
                nombre))
    # Si hay receta para ese banco, manda: es la que sabe leer lo COPIADO, que
    # no tiene la forma del archivo descargado.
    if nombre in _POR_RECETA:
        return _POR_RECETA[nombre](filas), nombre
    # Scotiabank y Monex interpretan el TEXTO tal cual —ancho fijo el uno,
    # expresiones regulares sobre el PDF el otro—; los demás trabajan sobre las
    # filas ya separadas.
    if nombre == "SCOTIABANK":
        return leer_scotiabank("", texto=texto), nombre
    if nombre == "MONEX":
        return leer_monex("", texto=texto), nombre
    lector = _PEGABLES_POR_FILAS.get(nombre)
    if lector is None:
        raise ErrorLector("No hay lector de pegado para {}.".format(nombre))
    return lector("", filas=filas), nombre


def leer(ruta: str, banco: str | None = None) -> tuple[list[LineaSaldo], str]:
    """Lee un reporte y devuelve `(líneas, banco detectado)`.

    `banco` fuerza el lector cuando la detección falla (p. ej. un portal que cambió
    sus encabezados). Lanza ErrorLector si no se reconoce o si el lector falla."""
    if not os.path.exists(ruta):
        raise ErrorLector(f"No se encontró el archivo «{ruta}».")
    if es_temporal(ruta):
        raise ErrorLector(
            f"«{os.path.basename(ruta)}» es un archivo temporal que Excel crea "
            "mientras un libro está abierto, no un reporte.")
    nombre = banco or detectar(ruta)
    if nombre is None:
        # `detectar` devuelve None tanto si no reconoce el formato como si no
        # pudo abrir el archivo. Se vuelve a intentar la huella para dar el
        # motivo REAL —«está abierto en Excel» es accionable, «no se reconoce»
        # manda al usuario a buscar el problema donde no está—.
        _huella(ruta)
        raise ErrorLector(
            f"No se reconoce de qué banco es «{os.path.basename(ruta)}». "
            "Puede ser un formato nuevo del portal.")
    lector = _POR_NOMBRE.get(nombre)
    if lector is None:
        raise ErrorLector(f"No hay lector para «{nombre}».")
    lineas = lector(ruta)
    origen = os.path.basename(ruta)
    for x in lineas:
        x.origen = origen
        # Si el reporte trae CLABE, el banco sale de ella: es el dato duro. El
        # nombre del lector es solo la firma del formato.
        if x.clabe:
            canonico = banco_desde_clabe(x.clabe)
            if canonico:
                x.banco = canonico
    return lineas, nombre


def leer_varios(rutas, progreso=None) -> tuple[list[LineaSaldo], list[str]]:
    """Lee varios reportes. Devuelve `(líneas, errores)`.

    Un archivo que falle NO tumba el lote: se reporta y se sigue con los demás, que
    es lo que hace falta cuando se cargan doce reportes de golpe. `progreso(hechos,
    total)` se llama tras cada archivo.
    """
    lineas: list[LineaSaldo] = []
    errores: list[str] = []
    total = len(rutas)
    for i, ruta in enumerate(rutas, 1):
        try:
            nuevas, _ = leer(ruta)
            lineas.extend(nuevas)
        except ErrorLector as exc:
            errores.append(str(exc))
        except Exception as exc:  # noqa: BLE001 — un archivo raro no aborta el lote
            errores.append(f"«{os.path.basename(ruta)}»: {exc}")
        if progreso is not None:
            progreso(i, total)
    return lineas, errores
