"""La plantilla del reporte de saldos: qué cuenta va en qué fila de qué pestaña.

El formato de tesorería es un libro donde cada descarga bancaria se pega en su
pestaña y la hoja SALDOS la lee con referencias a celda fija (`=HSBC!C2`). Ese
mecanismo funciona mientras el portal entregue las cuentas SIEMPRE en el mismo
orden — y no lo hace. Cuando cambia, la referencia sigue apuntando a la misma
fila, que ahora tiene otra cuenta, y el reporte sale mal sin avisar.

Este módulo es la pieza que quita ese supuesto. No cambia las fórmulas: cambia
quién decide qué fila ocupa cada cuenta. Si nosotros colocamos cada saldo en su
fila canónica, `=HSBC!C2` vuelve a ser correcto por construcción.

Los dos artefactos que carga los produce `scripts/derivar_plantilla_saldos.py` a
partir del formato real:

  saldos_base.xlsx  el libro vacío, con sus 24 pestañas y todas sus fórmulas
  saldos_mapa.json  el mapeo cuenta -> (pestaña, fila) y las columnas de cada
                    región

## Cómo se casa una cuenta

En orden, y la primera que resuelve gana:

  1. **Número completo** — normalizado a puros dígitos sin ceros a la izquierda,
     porque los portales rellenan distinto (`000000000110311944` y `110311944`
     son la misma cuenta de BBVA).
  2. **Cuenta embebida en la CLABE** — posiciones 6 a 17.
  3. **(pestaña, últimos 4)** — solo para los 7 renglones del formato que no
     traen número, y solo si resuelve único DENTRO de esa pestaña.

La regla 3 es el último recurso y está acotada a propósito: se aplica únicamente
sobre renglones que ya sabemos que no tienen número, nunca como desempate
general. Meter un saldo en la empresa equivocada es peor que dejar el renglón
vacío — el hueco se ve, el error no.
"""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, replace

from . import rutas

_DIR = os.path.join(rutas.BUNDLE, "core", "datos")
RUTA_BASE = os.path.join(_DIR, "saldos_base.xlsx")
RUTA_MAPA = os.path.join(_DIR, "saldos_mapa.json")

# Prefijo de CLABE (los 3 primeros dígitos, según Banxico) -> pestaña donde se
# pega ese banco. Scotiabank y Ve por Más comparten la hoja 'BX+ SCO'; es una
# rareza del formato, no un error.
HOJA_POR_PREFIJO = {
    "002": "BANAMEX",
    "012": "BANCOMER",
    "014": "SANTANDER",
    "021": "HSBC",
    "030": "BAJIO",
    "036": "INBURSA",
    "044": "BX+ SCO",
    "058": "BANREGIO",
    "062": "AFIRME",
    "072": "BANORTE",
    "112": "MONEX",
    "113": "BX+ SCO",
    "132": "MULTIVA",
    "137": "BANCOPPEL",
    "156": "SABADELL",
    "630": "INTERCAM",
}


# Pestañas donde conviven MÁS DE UN banco. Se derivan del mapa de arriba en vez
# de escribirse a mano: si mañana el formato junta otro par de bancos en una hoja,
# basta con apuntarlo en HOJA_POR_PREFIJO y esto se entera solo.
#
# Sirve para que quien muestre las cuentas de una de estas pestañas sepa que el
# nombre de la hoja NO identifica al banco y tiene que decirlo por cuenta.
HOJAS_COMPARTIDAS = frozenset(
    hoja for hoja in set(HOJA_POR_PREFIJO.values())
    if sum(1 for h in HOJA_POR_PREFIJO.values() if h == hoja) > 1)


class ErrorPlantilla(Exception):
    """La plantilla no se pudo cargar o es incoherente."""


def digitos(valor) -> str:
    return re.sub(r"\D", "", str(valor or ""))


def normalizar_cuenta(valor) -> str:
    """Número comparable: puros dígitos y sin ceros a la izquierda."""
    d = digitos(valor)
    return d.lstrip("0") or d


@dataclass(frozen=True)
class Renglon:
    """Un renglón de la hoja SALDOS y la celda de banco de la que se alimenta."""

    banda: str          # 'A-I' | 'K-O' | 'Q-U'
    fila: int           # fila en la hoja SALDOS
    bloque: str         # nombre del bloque de empresa
    banco: str          # como lo escribe el formato ('Banbajio', 'BBVA TDE')
    cta4: str           # los 4 dígitos que muestra el formato
    marca: str          # 'DLS' | 'TDE' | 'CLN' | 'MZT' | 'LM' | ''
    hoja: str           # pestaña de descarga
    hoja_fila: int      # fila canónica dentro de esa pestaña
    hoja_col: str       # columna de la que lee la fórmula de SALDOS
    cuenta: str = None  # número completo, cuando el formato lo tiene
    titular: str = None
    reparado: bool = False   # su referencia estaba desfasada y se corrigió

    @property
    def cuenta_norm(self) -> str:
        return normalizar_cuenta(self.cuenta) if self.cuenta else ""

    @property
    def cola(self) -> str:
        return digitos(self.cta4)[-4:]

    def __str__(self) -> str:
        return "{} · {} {} ({}!{}{})".format(
            self.bloque, self.banco, self.cta4,
            self.hoja, self.hoja_col, self.hoja_fila)


# Pestañas cuyo portal ENMASCARA el número de cuenta. Banamex nunca lo manda
# completo: su reporte trae la sucursal en una columna y la cuenta como «**3101»,
# o sea los últimos cuatro dígitos y nada más. Para estas hojas la terminación no
# es un último recurso dudoso, es el único identificador que da la fuente.
HOJAS_ENMASCARADAS = frozenset({"BANAMEX"})


@dataclass(frozen=True)
class Destino:
    """Una fila de una pestaña de descarga: dónde se pega el saldo de una cuenta.

    Hay MÁS destinos que renglones de SALDOS. Las pestañas del formato reproducen
    la descarga completa del portal, y la hoja SALDOS solo lee algunas de sus
    filas: Banamex trae 16 cuentas y SALDOS toma 7; Banorte trae 24 y toma 21. Las
    otras son cuentas reales de la empresa —tarjetas, cuentas en ceros, líneas de
    crédito— que el reporte no desglosa pero que sí van pegadas en su pestaña.

    Confundir ambos conjuntos es lo que hacía que 15 cuentas que SÍ están en el
    formato se reportaran como «cuentas nuevas» y su pestaña saliera a medias.
    """

    hoja: str
    fila: int
    cuenta: str = None
    titular: str = None
    moneda: str = None
    # Solo BANAMEX: su portal reporta sucursal y cuenta en columnas separadas y el
    # formato guarda las dos. Hace falta para casar la forma corta que manda el
    # portal (ver `por_sucursal` en Plantilla).
    sucursal: str = ""
    renglon: Renglon = None   # el renglón de SALDOS que lo lee, si alguno
    # Otras filas de la MISMA pestaña que llevan esta misma cuenta. Banregio, por
    # ejemplo, repite AEROSERVICIOS en su bloque de "cuentas nuevas" (C18 y C47) y
    # el formato pone el saldo en las dos. Se escriben todas para que la pestaña
    # quede igual que la que arma tesorería.
    gemelos: tuple = ()

    @property
    def cuenta_norm(self) -> str:
        return normalizar_cuenta(self.cuenta) if self.cuenta else ""

    @property
    def en_reporte(self) -> bool:
        """Si además de pegarse en su pestaña, aparece en la hoja SALDOS."""
        return self.renglon is not None

    def __str__(self) -> str:
        if self.renglon is not None:
            return str(self.renglon)
        return "{}!{} · {}".format(self.hoja, self.fila,
                                   self.titular or self.cuenta or "?")


class Plantilla:
    """Los 209 renglones del formato, con sus índices de casado."""

    def __init__(self, mapa: dict, ruta_base: str = RUTA_BASE):
        self.ruta_base = ruta_base
        self.hojas = mapa["hojas"]
        self.ledgers = mapa.get("ledgers", {})
        self.bandas = {b["nombre"]: b for b in mapa["bandas"]}
        self.desfases = mapa.get("desfases_corregidos", [])
        # Totales de la cabecera, ya aplanados a los renglones que suman, y a qué
        # celda de la fila del día anterior le toca cada uno.
        self.totales_cabecera = mapa.get("totales_cabecera", {})
        self.espejo_totales = mapa.get("espejo_totales", {})
        self.celda_fecha_anterior = mapa.get("celda_fecha_anterior", "L5")
        self.celda_hora_anterior = mapa.get("celda_hora_anterior", "L6")
        self.renglones = [Renglon(**r) for r in mapa["renglones"]]

        self.destinos: list = []
        self.por_cuenta: dict[str, Destino] = {}
        self.por_cola: dict[tuple, Destino] = {}
        self.por_sucursal: dict[tuple, Destino] = {}
        self.por_cola_enmascarada: dict[tuple, Destino] = {}
        self._colas_ambiguas: set = set()
        self._indexar()

    # -- construcción ------------------------------------------------------

    def _indexar(self) -> None:
        """Arma los destinos —TODAS las filas de las pestañas— y sus índices.

        Se recorre el inventario de pestañas, no la lista de renglones: el
        reporte lee 209 filas pero las pestañas tienen 225, y las 16 restantes
        también hay que pegarlas para que la pestaña quede como la que arma
        tesorería a mano.

        Una cuenta repetida en dos destinos significaría que el mismo saldo
        alimenta dos celdas: o el mapa está mal, o el formato duplica la cuenta.
        En cualquier caso hay que verlo, no tragárselo."""
        por_celda = {(r.hoja, r.hoja_fila): r for r in self.renglones}

        for nombre, info in self.hojas.items():
            for fila, entrada in info["filas"].items():
                fila = int(fila)
                estaticos = entrada.get("estaticos") or {}
                self.destinos.append(Destino(
                    hoja=nombre, fila=fila,
                    cuenta=entrada.get("cuenta"),
                    titular=entrada.get("titular") or entrada.get("alias"),
                    moneda=entrada.get("moneda"),
                    sucursal=digitos(estaticos.get("sucursal")),
                    renglon=por_celda.get((nombre, fila))))

        # Agrupa por cuenta. Que una cuenta caiga en varias filas es legítimo
        # —el formato la repite en más de un bloque— y el saldo va a todas. Lo
        # que NO puede pasar es que dos RENGLONES de SALDOS lean la misma cuenta:
        # eso sería el mismo saldo contado dos veces en el reporte, y hay que
        # verlo, no tragárselo.
        grupos: dict[str, list] = {}
        for d in self.destinos:
            if d.cuenta_norm:
                grupos.setdefault(d.cuenta_norm, []).append(d)

        conflictos = []
        for cuenta, filas in grupos.items():
            con_renglon = [d for d in filas if d.renglon is not None]
            if len(con_renglon) > 1:
                conflictos.append((cuenta, con_renglon))
                continue
            principal = con_renglon[0] if con_renglon else filas[0]
            otras = tuple(d for d in filas if d is not principal)
            if otras:
                principal = replace(principal, gemelos=otras)
            self.por_cuenta[cuenta] = principal
        if conflictos:
            detalle = "; ".join(
                "{} en {}".format(c, " y ".join(str(d) for d in ds))
                for c, ds in conflictos[:5])
            raise ErrorPlantilla(
                "{} cuenta(s) alimentan más de un renglón de SALDOS: {}".format(
                    len(conflictos), detalle))

        # Índice sucursal + últimos 4. Banamex identifica una cuenta por la
        # sucursal y la terminación, no por el número completo: en el reporte que
        # manda el portal, la cuenta 7713101 de la sucursal 237 puede venir como
        # «2373101». No es una cola suelta —eso sería ambiguo—, es sucursal Y
        # terminación juntas, que en esta pestaña identifican una sola fila.
        #
        # Se concatena la sucursal tal cual y no se supone que mida tres dígitos:
        # en el formato conviven sucursales de tres (237, 394, 114) y de cuatro
        # (7001, 7004, 7006), y partir por posición fija casaría mal la mitad.
        #
        # Va acotado a la pestaña y solo si resuelve único, igual que la cola:
        # meter un saldo en la empresa equivocada es peor que dejar el renglón
        # vacío, porque el hueco se ve y el error no.
        ambiguas: set = set()
        for d in self.destinos:
            if not d.sucursal or not d.cuenta_norm:
                continue
            clave = (d.hoja, d.sucursal + d.cuenta_norm[-4:])
            if clave in self.por_sucursal:
                ambiguas.add(clave)
                continue
            self.por_sucursal[clave] = d
        for clave in ambiguas:
            self.por_sucursal.pop(clave, None)

        # Cola dentro de una pestaña ENMASCARADA. Es la misma idea que la regla
        # general de la cola, pero aquí se ofrece también a las filas que SÍ
        # tienen número, porque en esas hojas el portal no manda otra cosa: si no
        # se usara la terminación, un reporte de Banamex sin sucursal no casaría
        # con nada. Sigue exigiéndose que sea única dentro de la pestaña; si dos
        # cuentas compartieran final, ninguna casa y ambas salen como nuevas —el
        # hueco se ve, el saldo en la empresa equivocada no—.
        ambiguas = set()
        for d in self.destinos:
            if d.hoja not in HOJAS_ENMASCARADAS or not d.cuenta_norm:
                continue
            clave = (d.hoja, d.cuenta_norm[-4:])
            if clave in self.por_cola_enmascarada:
                ambiguas.add(clave)
                continue
            self.por_cola_enmascarada[clave] = d
        for clave in ambiguas:
            self.por_cola_enmascarada.pop(clave, None)

        # La regla de la cola solo se ofrece a los destinos que no tienen número
        # —y solo los tiene el renglón de SALDOS que los lee—, y solo si es única
        # dentro de su pestaña.
        for d in self.destinos:
            if d.cuenta_norm or d.renglon is None or not d.renglon.cola:
                continue
            clave = (d.hoja, d.renglon.cola)
            if clave in self.por_cola:
                self._colas_ambiguas.add(clave)
                continue
            self.por_cola[clave] = d
        for clave in self._colas_ambiguas:
            self.por_cola.pop(clave, None)

    # -- consulta ----------------------------------------------------------

    def hoja_de_prefijo(self, prefijo: str) -> str:
        """Pestaña donde se pega el banco de ese prefijo de CLABE."""
        return HOJA_POR_PREFIJO.get(str(prefijo or ""))

    def buscar(self, *formas) -> Destino:
        """Fila que corresponde a alguna de esas formas del número."""
        for forma in formas:
            if not forma:
                continue
            d = self.por_cuenta.get(normalizar_cuenta(forma))
            if d is not None:
                return d
        return None

    def buscar_por_sucursal(self, hoja: str, *formas) -> Destino:
        """Fila que casa con «sucursal + últimos 4» dentro de esa pestaña.

        Es la forma en que Banamex nombra sus cuentas. Se consulta DESPUÉS del
        número completo, para que un número de verdad siempre gane."""
        if not hoja:
            return None
        for forma in formas:
            d = self.por_sucursal.get((hoja, digitos(forma)))
            if d is not None:
                return d
        return None

    def buscar_por_cola(self, hoja: str, *formas) -> Destino:
        """Último recurso: (pestaña, últimos 4) sobre las filas sin número.

        En las pestañas enmascaradas se consultan además las filas que sí tienen
        número, porque ahí la terminación es todo lo que manda el portal."""
        if not hoja:
            return None
        indices = [self.por_cola]
        if hoja in HOJAS_ENMASCARADAS:
            indices.append(self.por_cola_enmascarada)
        for forma in formas:
            cola = digitos(forma)[-4:]
            if len(cola) < 4:
                continue
            for indice in indices:
                d = indice.get((hoja, cola))
                if d is not None:
                    return d
        return None

    def colas_de(self, hoja: str) -> set:
        """Terminaciones de las cuentas que la plantilla tiene en esa pestaña."""
        return {d.cuenta_norm[-4:] for d in self.destinos
                if d.hoja == hoja and d.cuenta_norm}

    def columnas(self, hoja: str, fila: int) -> dict:
        """Roles -> letra de columna para escribir en esa fila.

        Cada pestaña tiene varias regiones y no todas usan las mismas columnas:
        en BBVA el primer bloque trae el disponible en F y los demás en E."""
        info = self.hojas.get(hoja)
        if not info:
            return {}
        entrada = info["filas"].get(str(fila))
        if entrada is None:
            return {}
        return info["regiones"][entrada["region"]]["cols"]

    def region(self, hoja: str, fila: int) -> dict:
        info = self.hojas.get(hoja)
        if not info:
            return {}
        entrada = info["filas"].get(str(fila))
        if entrada is None:
            return {}
        return info["regiones"][entrada["region"]]

    @property
    def vacios_posibles(self) -> list:
        """Destinos que solo se pueden casar por cola: los frágiles."""
        return [d for d in self.destinos
                if not d.cuenta_norm and d.renglon is not None]


_cache = None


def cargar(forzar: bool = False) -> Plantilla:
    """Carga la plantilla (una sola vez por proceso: son 114 KB de JSON)."""
    global _cache
    if _cache is not None and not forzar:
        return _cache
    if not os.path.exists(RUTA_MAPA):
        raise ErrorPlantilla(
            "falta {}. Genéralo con:\n"
            "  python scripts/derivar_plantilla_saldos.py <formato.xlsx>"
            .format(RUTA_MAPA))
    if not os.path.exists(RUTA_BASE):
        raise ErrorPlantilla("falta el libro base {}".format(RUTA_BASE))
    with open(RUTA_MAPA, encoding="utf-8") as f:
        mapa = json.load(f)
    _cache = Plantilla(mapa)
    return _cache
