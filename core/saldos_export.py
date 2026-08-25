"""Generación del reporte de saldos en Excel, listo para imprimir.

Replica la hoja `SALDOS` del formato que tesorería llena a mano: bloques por
empresa con banco · cuenta · saldo, repartidos en DOS columnas de bloques para que
todo quepa en una hoja tamaño OFICIO vertical. La configuración de impresión va
puesta en el archivo (papel, área y márgenes tomados del formato real, y ajuste
automático a una página), así que el usuario abre y manda a imprimir sin tocar
nada.

Qué NO incluye, y por qué: el formato original tiene además columnas de créditos
(C-H) y una proyección semanal de flujo (K-O) que se alimentan de las hojas
`CRÉDITOS`, `MGC`, `PEMEX`, `TESORO`, `NOMINA` e `IMPUESTOS`. Esos datos **no salen
de los portales bancarios**, así que este módulo no los puede calcular y deja esas
columnas libres. Es una omisión deliberada, no un olvido.

La segunda hoja, **Sin identificar**, es tan importante como la primera: lista las
cuentas que el banco reportó pero que no se pudieron atribuir a una empresa. Sin
ella, esos saldos desaparecerían del reporte sin dejar rastro, que es justo el
fallo silencioso que este módulo viene a eliminar.
"""

from __future__ import annotations

import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

from .saldos import Resultado, SaldoCuenta

# --- Geometría, tomada del formato real ---------------------------------
# Los bloques se reparten en DOS columnas para aprovechar el ancho de la hoja, tal
# como está armado el formato original.
_HOJA = "SALDOS"
_HOJA_SIN = "Sin identificar"
_FILA_INICIO = 9          # debajo del título y la fecha
# Alto de referencia del formato original (su área de impresión es A1:U131). Se usa
# para repartir los bloques en dos columnas parecidas, NO como tope: si hay más
# cuentas, la hoja crece y Excel ajusta la escala (ver _configurar_impresion).
_FILA_REFERENCIA = 131
_ULTIMA_COLUMNA = "U"
_PAPEL_OFICIO = 5         # PAPERSIZE_LEGAL

# Geometría de cada región. No son simétricas: en la izquierda el nombre de la
# empresa va en la columna F (a media banda) y alineado a la izquierda, mientras
# que en la derecha va en la Q y centrado. La etiqueta 'Total:' también cambia de
# columna según la región y la divisa. Todo esto está calcado del formato real —
# parece arbitrario, pero es lo que le da su aspecto reconocible.
_IZQ = {
    "banco": "A", "cuenta": "B", "divisa": "H", "saldo": "I",
    "empresa": "F", "empresa_alin": "left",
    "total_mxn": "B", "total_otras": "F", "total_alin": "right",
    "divisa_negrita": False,
}
_DER = {
    "banco": "Q", "cuenta": "R", "divisa": "T", "saldo": "U",
    "empresa": "Q", "empresa_alin": "center",
    "total_mxn": "S", "total_otras": "S", "total_alin": "center",
    "divisa_negrita": True,
}

# Anchos del formato original. Las columnas que no se llenan (C-G, J-P, S) se
# conservan igual para no alterar el reparto de la página al imprimir.
_ANCHOS = {
    "A": 19.7, "B": 13.1, "C": 14.4, "D": 13.1, "E": 15.7, "F": 15.7,
    "G": 15.7, "H": 15.7, "I": 26.0, "J": 8.1, "K": 19.7, "L": 20.7,
    "M": 21.1, "N": 18.4, "O": 25.7, "P": 7.4, "Q": 32.7, "R": 13.4,
    "S": 21.0, "T": 15.1, "U": 25.4,
}
_ALTO_FILA = 24.0
# La hoja original se ve SIN cuadrícula y con zoom al 50 %: con las líneas puestas
# el aspecto cambia por completo, aunque el contenido sea idéntico. El color de
# pestaña también es del formato (tema 2, ligeramente oscurecido).
_ZOOM = 50
_TAB_TEMA, _TAB_TINTE = 2, -0.1

_FMT_SALDO = "#,##0_);[Red](#,##0)"
_F_EMPRESA = Font(name="Calibri", size=18, bold=True)
_F_BANCO = Font(name="Calibri", size=18)
_F_DATO = Font(name="Calibri", size=16)
_F_CUENTA = Font(name="Calibri", size=16, color="FF000000")
_F_DIVISA = Font(name="Calibri", size=16)
_F_DIVISA_B = Font(name="Calibri", size=16, bold=True)
_F_TOTAL = Font(name="Calibri", size=16, bold=True)
_F_TITULO = Font(name="Calibri", size=22, bold=True)
_F_NOTA = Font(name="Calibri", size=11, color="FF808080")

_DER_ALIN = Alignment(horizontal="right")
_CEN_ALIN = Alignment(horizontal="center")
_IZQ_ALIN = Alignment(horizontal="left")

# Cada saldo va encajonado arriba y abajo; el total cierra con línea media.
_BORDE_DATO = Border(top=Side(style="thin"), bottom=Side(style="thin"))
_BORDE_TOTAL = Border(top=Side(style="thin"), bottom=Side(style="medium"))


def _bloques(res: Resultado, nombres: dict) -> list[dict]:
    """Un bloque por empresa: sus cuentas separadas por moneda.

    Se separan las divisas porque sumar pesos con dólares da una cifra que no
    significa nada; el formato original hace lo mismo (marca 'DLS' y total aparte).
    """
    out: list[dict] = []
    for id_empresa, saldos in res.por_empresa().items():
        mxn = [s for s in saldos if s.moneda == "MXN"]
        otras: dict[str, list[SaldoCuenta]] = {}
        for s in saldos:
            if s.moneda != "MXN":
                otras.setdefault(s.moneda, []).append(s)
        out.append({
            "empresa": nombres.get(id_empresa, f"Empresa {id_empresa}"),
            "id": id_empresa,
            "mxn": mxn,
            "otras": otras,
            # +1 encabezado, +1 total por moneda, +1 línea en blanco al final
            "alto": (len(mxn) + (2 if mxn else 0)
                     + sum(len(v) + 2 for v in otras.values()) + 1),
        })
    out.sort(key=lambda b: -sum(s.saldo for s in b["mxn"]))
    return out


def _repartir(bloques: list[dict]) -> tuple[list, list]:
    """Reparte los bloques en dos columnas de alto parecido.

    Se acumula en la izquierda mientras no se pase de la mitad del alto total; el
    resto va a la derecha. NUNCA se descarta un bloque por falta de espacio: la
    hoja se alarga y la impresión se ajusta. El formato original sí recorta —su
    contenido llega a la fila 143 y su área de impresión termina en la 131—, y ese
    es justo el tipo de pérdida silenciosa que este módulo viene a evitar."""
    total = sum(b["alto"] for b in bloques)
    objetivo = total / 2
    izq, der, acumulado = [], [], 0
    for b in bloques:
        if acumulado + b["alto"] / 2 <= objetivo:
            izq.append(b)
            acumulado += b["alto"]
        else:
            der.append(b)
    return izq, der


def _escribir_bloque(ws, bloque: dict, fila: int, cols: dict) -> int:
    """Escribe un bloque de empresa y devuelve la fila siguiente."""
    alin_empresa = (_IZQ_ALIN if cols["empresa_alin"] == "left" else _CEN_ALIN)
    celda = ws[f"{cols['empresa']}{fila}"]
    celda.value = bloque["empresa"]
    celda.font = _F_EMPRESA
    celda.alignment = alin_empresa
    ws.row_dimensions[fila].height = _ALTO_FILA
    fila += 1

    def grupo(saldos, divisa):
        nonlocal fila
        for s in saldos:
            ws.row_dimensions[fila].height = _ALTO_FILA
            b = ws[f"{cols['banco']}{fila}"]
            b.value = s.banco
            b.font = _F_BANCO
            b.alignment = _IZQ_ALIN
            # La cuenta va como TEXTO: los últimos dígitos pueden empezar en cero
            # ('0454') y Excel se los comería si la tratara como número.
            c = ws[f"{cols['cuenta']}{fila}"]
            c.value = _cola_visible(s.cuenta)
            c.font = _F_CUENTA
            c.alignment = _CEN_ALIN
            c.number_format = "@"
            if divisa != "MXN":
                d = ws[f"{cols['divisa']}{fila}"]
                d.value = "DLS" if divisa == "USD" else divisa
                d.font = _F_DIVISA_B if cols["divisa_negrita"] else _F_DIVISA
                d.alignment = _CEN_ALIN
            v = ws[f"{cols['saldo']}{fila}"]
            v.value = s.saldo
            v.font = _F_DATO
            v.alignment = _DER_ALIN
            v.number_format = _FMT_SALDO
            v.border = _BORDE_DATO
            fila += 1
        # Fila de total del grupo. La etiqueta cambia de columna según la divisa:
        # en la región izquierda el total en pesos se rotula en B y el de dólares
        # en F, tal como está el formato original.
        ws.row_dimensions[fila].height = _ALTO_FILA
        col_etq = cols["total_mxn"] if divisa == "MXN" else cols["total_otras"]
        etq = ws[f"{col_etq}{fila}"]
        etq.value = "Total: " if divisa == "MXN" else f"Total {divisa}: "
        etq.font = _F_TOTAL
        etq.alignment = (_DER_ALIN if cols["total_alin"] == "right" else _CEN_ALIN)
        t = ws[f"{cols['saldo']}{fila}"]
        # Fórmula, no el número ya sumado: así el usuario puede ajustar una cifra a
        # mano y ver el total actualizarse, que es como usa hoy el formato.
        ini_f, fin_f = fila - len(saldos), fila - 1
        col_s = cols["saldo"]
        t.value = f"=SUM({col_s}{ini_f}:{col_s}{fin_f})" if saldos else 0
        t.font = _F_TOTAL
        t.alignment = _DER_ALIN
        t.number_format = _FMT_SALDO
        t.border = _BORDE_TOTAL
        fila += 1

    if bloque["mxn"]:
        grupo(bloque["mxn"], "MXN")
    for divisa, saldos in sorted(bloque["otras"].items()):
        grupo(saldos, divisa)
    return fila + 1  # línea en blanco entre bloques


def _cola_visible(cuenta: str, n: int = 4) -> str:
    """Los últimos `n` dígitos, que es como el formato identifica cada cuenta.

    Se muestra la cola y no el número completo porque así lo lee el usuario en la
    hoja impresa (columna 'Cta.'), y porque a esa escala el número entero no cabe.
    """
    d = "".join(c for c in str(cuenta or "") if c.isdigit())
    return d[-n:].rjust(min(n, len(d)), "0") if d else ""


def _encabezado(ws, res: Resultado, fecha: datetime.datetime) -> None:
    """Título, fecha, totales generales y el rótulo 'Cta.' sobre cada región.

    Reproduce la cabecera del formato en lo que este módulo puede calcular. Se
    omiten los bloques de 'Cobertura' y 'Amortización' (filas 3-7 del original):
    salen de la hoja CRÉDITOS, que no viene de los portales bancarios."""
    ws["M1"] = "SALDOS"
    ws["M1"].font = _F_TITULO
    ws["M1"].alignment = _CEN_ALIN
    ws.row_dimensions[1].height = 22.5

    ws["K3"] = "Fecha:"
    ws["K3"].font = _F_TOTAL
    ws["K3"].alignment = _DER_ALIN
    ws["L3"] = fecha.strftime("%d/%m/%Y")
    ws["L3"].font = _F_DATO
    ws["K4"] = "Hora:"
    ws["K4"].font = _F_TOTAL
    ws["K4"].alignment = _DER_ALIN
    ws["L4"] = fecha.strftime("%H:%M")
    ws["L4"].font = _F_DATO

    # Totales generales por divisa, como en el original (P3/Q3 = MX, P4/Q4 = DLS).
    # No se convierten entre sí: un total que mezcle pesos y dólares no significa
    # nada. El desglose 'Combustibles / Resto' del formato es una agrupación de
    # negocio que este módulo no conoce, así que no se inventa.
    totales: dict[str, float] = {}
    for s_ in res.identificados:
        totales[s_.moneda] = totales.get(s_.moneda, 0.0) + s_.saldo
    ws["P3"] = "MX:"
    ws["P3"].font = _F_TOTAL
    ws["P3"].alignment = _DER_ALIN
    ws["Q3"] = totales.get("MXN", 0.0)
    ws["Q3"].font = _F_TOTAL
    ws["Q3"].alignment = _DER_ALIN
    ws["Q3"].number_format = _FMT_SALDO
    ws["P4"] = "DLS:"
    ws["P4"].font = _F_TOTAL
    ws["P4"].alignment = _DER_ALIN
    ws["Q4"] = totales.get("USD", 0.0)
    ws["Q4"].font = _F_TOTAL
    ws["Q4"].alignment = _DER_ALIN
    ws["Q4"].number_format = _FMT_SALDO

    # Rótulo de la columna de cuentas, en cada región (fila 8 del original).
    for cols in (_IZQ, _DER):
        celda = ws[f"{cols['cuenta']}8"]
        celda.value = "Cta."
        celda.font = _F_TOTAL
        celda.alignment = _CEN_ALIN

    if res.sin_identificar:
        aviso = ws["A3"]
        aviso.value = (f"{len(res.sin_identificar)} cuenta(s) sin identificar — "
                       f"ver la hoja «{_HOJA_SIN}»")
        aviso.font = Font(name="Calibri", size=12, bold=True, color="FFC00000")


def _configurar_impresion(ws, ultima_fila: int) -> None:
    """Papel, ajuste, área y márgenes: una sola hoja oficio, siempre completa.

    En vez de fijar la escala al 27 % del formato original, se le pide a Excel que
    ajuste a UNA página (`fitToWidth`/`fitToHeight`). Con una escala fija, el día
    que se den de alta más cuentas el reporte se imprimiría recortado sin que nadie
    se entere; ajustando, sale entero aunque la letra encoja un poco. El área
    también sigue al contenido, nunca al revés."""
    # Sin cuadrícula: es lo que hace que la hoja se vea como el formato y no como
    # una tabla de cálculo cualquiera. Afecta a la pantalla, no a la impresión
    # (imprimir cuadrícula ya viene desactivado por omisión).
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = _ZOOM
    ws.sheet_properties.tabColor = Color(theme=_TAB_TEMA, tint=_TAB_TINTE)
    ws.page_setup.paperSize = _PAPEL_OFICIO
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    fin = max(_FILA_REFERENCIA, ultima_fila)
    ws.print_area = f"A1:{_ULTIMA_COLUMNA}{fin}"
    ws.page_margins = PageMargins(
        left=0.2362204724409449, right=0.2362204724409449,
        top=0.7480314960629921, bottom=0.7480314960629921)
    for col, ancho in _ANCHOS.items():
        ws.column_dimensions[col].width = ancho


_ENC_SIN = (
    ("Banco", 22), ("Cuenta", 24), ("Titular en el reporte", 44),
    ("Saldo", 18), ("Moneda", 10), ("Archivo de origen", 40), ("Motivo", 60),
)


def _hoja_sin_identificar(wb, res: Resultado) -> None:
    """Las cuentas que el banco reportó y no se pudieron atribuir."""
    ws = wb.create_sheet(_HOJA_SIN)
    ws["A1"] = ("Cuentas que reportó el banco y NO se pudieron atribuir a una "
                "empresa. Complétalas en el catálogo de cuentas.")
    ws["A1"].font = Font(name="Calibri", size=11, bold=True)
    for i, (etq, ancho) in enumerate(_ENC_SIN, 1):
        c = ws.cell(3, i, etq)
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
        c.fill = PatternFill("solid", fgColor="FF317FB1")
        c.alignment = _CEN_ALIN
        ws.column_dimensions[get_column_letter(i)].width = ancho
    fila = 4
    for item in list(res.sin_identificar) + list(res.duplicados):
        ln = item.linea
        motivo = item.motivo
        if item.candidatos:
            motivo += " — candidatos: " + "; ".join(item.candidatos[:4])
        valores = (ln.banco, ln.cuenta, ln.titular, ln.saldo, ln.moneda,
                   ln.origen, motivo)
        for i, v in enumerate(valores, 1):
            c = ws.cell(fila, i, v)
            c.font = Font(name="Calibri", size=10)
            if i == 2:      # la cuenta como texto, para no perder ceros
                c.number_format = "@"
            if i == 4:
                c.number_format = _FMT_SALDO
                c.alignment = _DER_ALIN
        fila += 1
    ws.freeze_panes = "A4"


def generar(ruta: str, res: Resultado, nombres: dict | None = None,
            fecha: datetime.datetime | None = None) -> dict:
    """Escribe el reporte en `ruta`. Devuelve un resumen de lo generado.

    `nombres` mapea id de empresa -> nombre a mostrar (lo provee la interfaz desde
    `ui.comun.EMPRESAS`, para que el núcleo no dependa de la capa de UI).

    El resumen trae `filas`: cuántas ocupó el reporte. Ningún bloque se descarta
    por espacio — la hoja crece y la impresión se ajusta a una página.
    """
    nombres = nombres or {}
    fecha = fecha or datetime.datetime.now()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = _HOJA

    bloques = _bloques(res, nombres)
    izq, der = _repartir(bloques)

    _encabezado(ws, res, fecha)
    ultima = _FILA_INICIO
    for lado, cols in ((izq, _IZQ), (der, _DER)):
        fila = _FILA_INICIO
        for bloque in lado:
            fila = _escribir_bloque(ws, bloque, fila, cols)
        ultima = max(ultima, fila)
    _configurar_impresion(ws, ultima)

    if res.sin_identificar or res.duplicados:
        _hoja_sin_identificar(wb, res)
    wb.save(ruta)
    return {
        "empresas": len(bloques),
        "cuentas": len(res.identificados),
        "sin_identificar": len(res.sin_identificar),
        "duplicados": len(res.duplicados),
        "filas": ultima,
    }
