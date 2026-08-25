"""Identificación de saldos: de qué empresa es cada cuenta que reportó el banco.

Es el corazón del módulo de Saldos y lo que sustituye al mecanismo del formato de
Excel, donde cada saldo se tomaba de una CELDA FIJA de la hoja pegada
(`=BANAMEX!E4`). Esa relación cuenta↔empresa vivía en la posición de la fila: si el
portal cambiaba el orden, el reporte salía mal sin avisar. Aquí se casa por el
NÚMERO DE CUENTA contra el catálogo, que es un dato estable.

Reglas, de más a menos confiable (se aplican en orden y la primera que resuelve
gana):

  1. **CLABE completa** — 18 dígitos con dígito de control. Inequívoca.
  2. **Número de cuenta completo** — contra el `numeroCuenta` del catálogo y contra
     la cuenta que va embebida en la CLABE (posiciones 7 a 17).
  3. **Últimos 4 dígitos + banco** — último recurso, y SOLO si resuelve a una
     única cuenta.

Sobre la regla 3 hay que ser explícito, porque es la que puede hacer daño: en el
catálogo real **41 de 341 colas de 4 dígitos son ambiguas**. La cola `0012` de
Banregio, por ejemplo, corresponde a cinco empresas distintas, porque todas sus
cuentas empiezan igual (`1149…`) y solo difieren en medio. Por eso una cola
ambigua NO se resuelve al azar: se reporta como ambigua y el usuario decide. Meter
un saldo en la empresa equivocada es peor que dejarlo fuera — el hueco se ve, el
error no.

Lo mismo aplica a los DUPLICADOS: si dos reportes traen la misma cuenta (pasa con
BBVA, que se descarga en varios archivos), no se suman. Se toma uno y se avisa,
porque sumar dos veces el mismo saldo infla el reporte en silencio.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field

from . import cuentas_dispersion
from .catalogo_bancos import CATALOGO_BANCOS
from .extractores import validar_clabe
from .saldos_lectores import LineaSaldo

# Cuántos dígitos finales se comparan en la regla 3. Cuatro es lo que enseñan los
# reportes y las máscaras bancarias; con menos, la ambigüedad se dispara.
_COLA = 4

# Nombre canónico de banco -> prefijo de CLABE. Se deriva del catálogo de Banxico
# para no mantener dos listas: los lectores ya nombran los bancos con ese mismo
# vocabulario (ver saldos_lectores).
_PREFIJO_POR_BANCO = {nombre.casefold(): codigo
                      for codigo, nombre in CATALOGO_BANCOS.items()}

# Cómo escribe el banco el CATÁLOGO dentro del texto de la cuenta ('BBVA BANCOMER
# 0100647012 ABASTECEDORA'). Hace falta aparte del mapa canónico porque ahí los
# nombres van a mano, con abreviaturas y erratas. Sirve para saber de qué banco es
# una cuenta del catálogo que no tiene CLABE — 49 de las 220 están en ese caso, y
# sin esto quedarían fuera de la regla de la cola.
_PREFIJO_POR_TEXTO = {
    "BANAMEX": "002", "BANAMEC": "002",
    "BANBAJIO": "030", "BAJIO": "030",
    "BANCOPPEL": "137", "BANORTE": "072", "BANREGIO": "058",
    "BANCOMER": "012", "BBVA": "012",
    "BX": "113", "VEPORMAS": "113",
    "HSBC": "021", "INBURSA": "036", "INTERCAM": "630",
    "MONEX": "112", "MULTIVA": "132", "SABADELL": "156",
    "SANTANDER": "014", "SANTANDEER": "014",
    "SCOTIABANK": "044", "SCOTIANBANK": "044",
    "AFIRME": "062",
}


def prefijo_desde_texto(texto: str) -> str | None:
    """Prefijo de CLABE del banco que se menciona en un texto libre.

    Se compara sobre el texto sin acentos ni signos, palabra por palabra y de más
    largo a más corto, para que 'BANBAJIO' gane sobre 'BAJIO' y 'BANCOMER' no se
    confunda con 'BANCOPPEL'."""
    plano = re.sub(r"[^A-Z0-9 ]", " ", str(texto or "").upper())
    plano = unicodedata.normalize("NFKD", plano)
    plano = "".join(c for c in plano if not unicodedata.combining(c))
    palabras = plano.split()
    for clave in sorted(_PREFIJO_POR_TEXTO, key=len, reverse=True):
        if any(p.startswith(clave) for p in palabras):
            return _PREFIJO_POR_TEXTO[clave]
    return None


@dataclass
class SaldoCuenta:
    """Un saldo ya atribuido a una empresa del catálogo."""

    id_empresa: int
    banco: str
    cuenta: str              # número tal como lo reportó el banco
    cuenta_catalogo: str     # texto 'Cuenta' del catálogo (lo que ve el usuario)
    saldo: float
    moneda: str
    regla: str               # clabe | numero | cola  (cómo se resolvió)
    linea: LineaSaldo = None


@dataclass
class SinIdentificar:
    """Un saldo que NO se pudo atribuir, con el porqué."""

    linea: LineaSaldo
    motivo: str
    candidatos: list = field(default_factory=list)


@dataclass
class Resultado:
    identificados: list[SaldoCuenta] = field(default_factory=list)
    sin_identificar: list[SinIdentificar] = field(default_factory=list)
    duplicados: list[SinIdentificar] = field(default_factory=list)

    @property
    def total(self) -> int:
        return (len(self.identificados) + len(self.sin_identificar)
                + len(self.duplicados))

    def por_empresa(self) -> dict[int, list[SaldoCuenta]]:
        """Saldos agrupados por id de empresa, en el orden en que se encontraron."""
        out: dict[int, list[SaldoCuenta]] = {}
        for s in self.identificados:
            out.setdefault(s.id_empresa, []).append(s)
        return out


def _digitos(valor) -> str:
    return re.sub(r"\D", "", str(valor or ""))


def _sin_ceros(digitos: str) -> str:
    """Número sin ceros a la izquierda. Los portales rellenan a un ancho fijo
    (BBVA manda '000000000121510312') y el catálogo no, así que hay que comparar
    en la misma forma."""
    return digitos.lstrip("0")


def _colas(digitos: str) -> set[str]:
    """Colas de `_COLA` dígitos con las que puede aparecer esta cuenta.

    De una CLABE se derivan DOS: la suya y la de la cuenta que lleva embebida. No
    son la misma, porque el dígito verificador va al final y las desplaza
    (012320001103245316 termina en '5316', pero su cuenta termina en '4531')."""
    if not digitos:
        return set()
    colas = {digitos[-_COLA:]}
    if len(digitos) == 18:
        colas.add(digitos[6:17][-_COLA:])
    return colas


def _compatibles(reporte: str, catalogo: str) -> bool:
    """True si dos números de cuenta pueden ser EL MISMO, uno enmascarado.

    Compartir los últimos cuatro dígitos no basta ni de lejos. En el catálogo real,
    la cuenta 16084470201 de Abastecedora y la 388379020201 de Merarid son ambas de
    Bajío y ambas terminan en '0201': casarlas por la cola mandó 7.4 millones de
    pesos a la empresa equivocada. Enmascarar un número lo RECORTA, así que el
    corto tiene que ser sufijo del largo — '7012' de '0100647012' sí, '0201' de dos
    cuentas largas distintas no."""
    a, b = _sin_ceros(reporte), _sin_ceros(catalogo)
    if not a or not b:
        return False
    return a.endswith(b) or b.endswith(a)


def prefijo_banco(nombre: str) -> str | None:
    """Prefijo de CLABE del banco a partir de su nombre canónico."""
    return _PREFIJO_POR_BANCO.get((nombre or "").casefold())


class _Indice:
    """Índice del catálogo por cada forma en que puede venir una cuenta.

    Guarda TODAS las coincidencias por clave, no solo la primera: es lo que permite
    detectar que una clave es ambigua en vez de resolverla al azar."""

    def __init__(self, catalogo):
        self.por_clabe: dict[str, list] = {}
        self.por_numero: dict[str, list] = {}
        self.por_cola: dict[tuple, list] = {}
        for id_empresa in catalogo.empresas():
            for reg in catalogo._registros(id_empresa):
                self._agregar(id_empresa, reg)

    def _agregar(self, id_empresa: int, reg: dict) -> None:
        entrada = (id_empresa, reg)
        clabe = _digitos(reg.get("clabe", ""))
        numero = _digitos(reg.get("numero", ""))
        if len(clabe) == 18 and validar_clabe(clabe):
            self.por_clabe.setdefault(clabe, []).append(entrada)
            # La cuenta embebida en la CLABE es otra forma válida del número.
            embebida = _sin_ceros(clabe[6:17])
            if embebida:
                self.por_numero.setdefault(embebida, []).append(entrada)
        if numero:
            corto = _sin_ceros(numero)
            if corto:
                self.por_numero.setdefault(corto, []).append(entrada)
        # El banco sale de la CLABE cuando la hay; si no, del texto de la cuenta.
        prefijo = (clabe[:3] if len(clabe) == 18
                   else prefijo_desde_texto(reg.get("cuenta", "")))
        for cola in _colas(clabe) | _colas(numero):
            self.por_cola.setdefault((prefijo, cola), []).append(entrada)

    @staticmethod
    def _unico(entradas):
        """La entrada si la clave resuelve a UNA sola cuenta; None si no o si es
        ambigua. Varias filas del catálogo que apunten a la misma cuenta de la
        misma empresa cuentan como una."""
        if not entradas:
            return None
        distintas = {(i, r.get("cuenta", "")) for i, r in entradas}
        return entradas[0] if len(distintas) == 1 else None


def _formas_cuenta(linea: LineaSaldo) -> list[str]:
    """Todas las formas en que el número de esta línea puede estar en el catálogo.

    Un mismo número se escribe distinto según el portal: Banamex reporta sucursal y
    cuenta en columnas separadas y aquí se concatenan (7004 + 965783), pero el
    catálogo puede guardar solo la corta. Que el lector declare sus variantes deja
    el caso resuelto por número —la regla fuerte— en vez de depender de la cola."""
    formas = [_digitos(linea.cuenta)]
    for clave in ("cuenta_corta", "cuenta_alterna"):
        alterna = _digitos((linea.extra or {}).get(clave, ""))
        if alterna and alterna not in formas:
            formas.append(alterna)
    return [f for f in formas if f]


def _casar(linea: LineaSaldo, idx: _Indice):
    """Aplica la cascada. Devuelve `(entrada, regla, candidatos)`.

    `entrada` es None si no casó; `candidatos` trae las opciones cuando la cola
    resultó ambigua, para poder explicárselo al usuario."""
    clabe = _digitos(linea.clabe)
    if len(clabe) == 18:
        entrada = idx._unico(idx.por_clabe.get(clabe, []))
        if entrada:
            return entrada, "clabe", []

    for forma in _formas_cuenta(linea):
        corta = _sin_ceros(forma)
        if not corta:
            continue
        entrada = idx._unico(idx.por_numero.get(corta, []))
        if entrada:
            return entrada, "numero", []

    # Último recurso: cola + banco. Solo vale si resuelve a una única cuenta.
    prefijo = prefijo_banco(linea.banco) or (clabe[:3] if len(clabe) == 18 else None)
    digitos_reporte = _digitos(linea.cuenta)
    candidatos: list = []
    for cola in _colas(digitos_reporte) | _colas(clabe):
        entradas = idx.por_cola.get((prefijo, cola), [])
        if not entradas:
            continue
        # Solo cuentan las que además son compatibles como número: la cola sola
        # empareja cuentas que no tienen nada que ver (ver _compatibles).
        viables = [
            (i, r) for i, r in entradas
            if _compatibles(digitos_reporte, _digitos(r.get("numero", "")))
            or _compatibles(digitos_reporte, _digitos(r.get("clabe", ""))[6:17])
        ]
        entrada = idx._unico(viables)
        if entrada:
            return entrada, "cola", []
        candidatos.extend(viables)
    return None, "", candidatos


def identificar(lineas, catalogo=None) -> Resultado:
    """Atribuye cada saldo leído a una empresa del catálogo.

    `catalogo` es un `CatalogoCuentasDispersion`; si no se pasa, se carga el
    instalado. Devuelve un `Resultado` con lo identificado, lo que no se pudo y los
    duplicados — los tres importan: un reporte de saldos con huecos silenciosos es
    peor que uno que los señala.
    """
    if catalogo is None:
        catalogo = cuentas_dispersion.CatalogoCuentasDispersion()
    idx = _Indice(catalogo)
    res = Resultado()
    vistas: dict[tuple, SaldoCuenta] = {}

    for linea in lineas:
        entrada, regla, candidatos = _casar(linea, idx)
        if entrada is None:
            if candidatos:
                nombres = sorted({f"id {i} · {r.get('cuenta', '')}"
                                  for i, r in candidatos})
                res.sin_identificar.append(SinIdentificar(
                    linea,
                    f"los últimos {_COLA} dígitos corresponden a "
                    f"{len(nombres)} cuentas distintas del catálogo",
                    nombres))
            else:
                res.sin_identificar.append(SinIdentificar(
                    linea, "la cuenta no está en el catálogo"))
            continue

        id_empresa, reg = entrada
        clave = (id_empresa, reg.get("cuenta", ""))
        # La moneda del catálogo manda sobre la del reporte: varios portales
        # (Banregio, Santander) no traen columna de divisa, y el catálogo sí sabe
        # cuáles cuentas son en dólares.
        moneda_cat = cuentas_dispersion.MONEDAS.get(reg.get("moneda"))
        moneda = ("USD" if reg.get("moneda") == 2
                  else "MXN" if reg.get("moneda") == 1
                  else (linea.moneda or "MXN"))
        saldo = SaldoCuenta(
            id_empresa=id_empresa, banco=linea.banco, cuenta=linea.cuenta,
            cuenta_catalogo=reg.get("cuenta", ""), saldo=linea.saldo,
            moneda=moneda, regla=regla, linea=linea)

        anterior = vistas.get(clave)
        if anterior is not None:
            # Misma cuenta en dos archivos. No se suma: se conserva la primera y se
            # reporta, porque duplicar un saldo infla el total sin que se note.
            res.duplicados.append(SinIdentificar(
                linea,
                f"la cuenta ya venía en «{anterior.linea.origen}» "
                f"(${anterior.saldo:,.2f}); no se suma"))
            continue
        vistas[clave] = saldo
        res.identificados.append(saldo)
        _ = moneda_cat  # nombre legible de la moneda, por si se quiere mostrar

    return res


def totales_por_empresa(res: Resultado) -> dict[int, dict[str, float]]:
    """Total por empresa y moneda: {id_empresa: {'MXN': x, 'USD': y}}.

    No se convierten divisas: mezclar pesos y dólares en un solo total daría una
    cifra que no significa nada."""
    out: dict[int, dict[str, float]] = {}
    for s in res.identificados:
        out.setdefault(s.id_empresa, {}).setdefault(s.moneda, 0.0)
        out[s.id_empresa][s.moneda] += s.saldo
    return out


def resumen(res: Resultado) -> str:
    """Una línea con el desenlace, para el aviso de la interfaz."""
    partes = [f"{len(res.identificados)} cuenta(s) identificada(s)"]
    if res.sin_identificar:
        partes.append(f"{len(res.sin_identificar)} sin identificar")
    if res.duplicados:
        partes.append(f"{len(res.duplicados)} duplicada(s)")
    return " · ".join(partes) + "."
