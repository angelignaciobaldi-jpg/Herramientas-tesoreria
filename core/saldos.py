"""Colocación de saldos: en qué fila del formato va cada cuenta que reportó el banco.

Es el corazón del módulo de Saldos. Sustituye al mecanismo del formato de Excel,
donde cada saldo se tomaba de una CELDA FIJA de la hoja pegada (`=HSBC!C2`): la
relación cuenta↔empresa vivía en la POSICIÓN de la fila, así que bastaba con que
el portal cambiara el orden de la descarga para que el reporte saliera mal en
silencio. No es hipotético — el formato que usa tesorería hoy trae seis renglones
de HSBC leyendo la cuenta equivocada.

Aquí no se tiran las fórmulas: se les quita el supuesto. Cada línea que leen los
lectores se casa por NÚMERO DE CUENTA contra la plantilla, que dice en qué fila
canónica va. Colocada ahí, `=HSBC!C2` vuelve a ser correcta por construcción.

Reglas, de más a menos confiable (la primera que resuelve gana):

  1. **CLABE completa** — 18 dígitos con dígito de control. Inequívoca.
  2. **Número de cuenta** — en todas las formas en que el portal puede darlo
     (completo, corto, y el embebido en la CLABE, posiciones 6 a 17).
  3. **Sucursal + últimos 4, dentro de la pestaña** — solo Banamex, que es el
     único que nombra así sus cuentas: la 7713101 de la sucursal 237 llega en el
     reporte como «2373101». Son siete u ocho dígitos, no cuatro, y la pareja es
     única en la pestaña; por eso va antes que la regla de la cola.
  4. **Últimos 4 dígitos dentro de la pestaña** — último recurso, acotado a los
     siete renglones que el formato no numera (Monex por alias, BX+, y los que se
     capturan a mano en Santander y Banamex), y solo si resuelve único.

Sobre la regla 4 hay que ser explícito, porque es la que puede hacer daño:
compartir los últimos cuatro dígitos no significa nada por sí solo. La cuenta
`16084470201` de Abastecedora y la `388379020201` de Merarid son ambas de Bajío y
ambas terminan en `0201`. Por eso la regla no se ofrece como desempate general:
solo la ven renglones que ya sabemos que no tienen número, y la plantilla
descarta de antemano cualquier cola que no sea única en su pestaña. Meter un
saldo en la empresa equivocada es peor que dejar el renglón vacío — el hueco se
ve, el error no.

Lo mismo aplica a los DUPLICADOS: si dos archivos traen la misma cuenta (pasa con
BBVA, que se descarga en varios), no se suman. Se toma uno y se avisa, porque
sumar dos veces el mismo saldo infla el reporte en silencio.

Ojo con una distinción que es fácil de perder: las pestañas del formato traen la
descarga COMPLETA del portal y la hoja SALDOS solo desglosa parte de ella. Banamex
manda 16 cuentas y el reporte muestra 7; Banorte manda 24 y muestra 21. Las demás
son cuentas reales —tarjetas, cuentas en ceros— que sí van pegadas en su pestaña
aunque el reporte no las liste. Por eso el casado va contra TODAS las filas de las
pestañas (`plantilla.destinos`) y no contra los 209 renglones de SALDOS.

Lo que no casa con ninguna fila no se pierde ni se cuela: va a `nuevas`, y el
exportador la lista en su propia pestaña. Es la única forma de enterarse de que
se abrió una cuenta que el formato todavía no contempla.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field

from .catalogo_bancos import CATALOGO_BANCOS
from .extractores import validar_clabe
from .saldos_lectores import LineaSaldo
from .saldos_plantilla import (Destino, Plantilla, Renglon,
                               cargar as cargar_plantilla)

# Nombre canónico de banco -> prefijo de CLABE. Se deriva del catálogo de Banxico
# para no mantener dos listas: los lectores nombran los bancos con ese mismo
# vocabulario (ver saldos_lectores).
_PREFIJO_POR_BANCO = {nombre.casefold(): codigo
                      for codigo, nombre in CATALOGO_BANCOS.items()}

# Cómo aparece el banco dentro de un texto libre. Hace falta aparte del mapa
# canónico porque algunos lectores anteponen 'Banco ' ('Banco Sabadell' no está
# en el catálogo, 'Sabadell' sí) y porque el formato escribe los nombres a mano.
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


# Cuentas que NO deben entrar al reporte, por banco y número tal como lo da el
# portal. Son cuentas de operación o en desuso que el formato nunca contempló:
# sin esta lista caen en «cuentas nuevas» y ensucian la pestaña de excepciones
# cada día, invitando a darlas de alta en la plantilla por error.
#
# Se excluye AQUÍ, en el embudo de identificación, y no en cada lector: así vale
# para todas las vías de carga y queda un solo lugar que mantener. La línea sale
# del flujo por completo: no se coloca, no se avisa en pantalla y no aparece en
# ninguna pestaña del libro. Queda en `Asignacion.excluidas` únicamente para
# poder revisarla desde el código si algún día hay que auditar la lista.
#
# El número se compara normalizado (solo dígitos) contra las mismas formas que usa
# el casado —la completa, la corta del portal y la embebida en la CLABE—, así que
# da igual con qué máscara venga escrito.
CUENTAS_EXCLUIDAS = (
    ("Banamex", "2375605601410"),
    ("Banamex", "23727547902799"),
    ("Banamex", "31617886014"),
    ("Banamex", "8182756227"),
    ("Scotiabank", "11700560642"),
    ("Scotiabank", "25603184671"),
    ("Scotiabank", "25605313032"),
)


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


def prefijo_banco(nombre: str) -> str | None:
    """Prefijo de CLABE del banco, por nombre canónico y si no por texto libre."""
    directo = _PREFIJO_POR_BANCO.get((nombre or "").casefold())
    return directo if directo else prefijo_desde_texto(nombre)


def _digitos(valor) -> str:
    return re.sub(r"\D", "", str(valor or ""))


# ---------------------------------------------------------------------------
# Resultado
# ---------------------------------------------------------------------------

@dataclass
class Colocada:
    """Un saldo ya asignado a su fila del formato."""

    destino: Destino
    linea: LineaSaldo
    regla: str          # clabe | numero | cola

    @property
    def saldo(self) -> float:
        return self.linea.saldo

    @property
    def renglon(self) -> Renglon:
        """El renglón de SALDOS que lee esta fila, si alguno.

        Puede ser None: hay cuentas que van pegadas en su pestaña pero que el
        reporte no desglosa (tarjetas, cuentas en ceros)."""
        return self.destino.renglon


@dataclass
class Suelta:
    """Una línea que no se colocó, con el porqué."""

    linea: LineaSaldo
    motivo: str


@dataclass
class Asignacion:
    """Lo que hay que escribir en el libro, y lo que quedó fuera."""

    plantilla: Plantilla
    colocadas: dict = field(default_factory=dict)   # (hoja, fila) -> Colocada
    nuevas: list = field(default_factory=list)      # Suelta
    duplicados: list = field(default_factory=list)  # Suelta
    excluidas: list = field(default_factory=list)   # Suelta

    @property
    def vacios(self) -> list:
        """Renglones de la hoja SALDOS que ningún archivo llenó.

        Es el dato que el módulo anterior no daba y que hace visible el hueco: si
        no subieron el archivo de Banorte, aquí salen sus 21 renglones.

        Se mide sobre los renglones de SALDOS, no sobre todas las filas de las
        pestañas: la cobertura que le importa al usuario es la del reporte que va
        a imprimir."""
        return [r for r in self.plantilla.renglones
                if (r.hoja, r.hoja_fila) not in self.colocadas]

    @property
    def llenos(self) -> int:
        """Renglones de SALDOS que quedaron con saldo."""
        return self.total_renglones - len(self.vacios)

    @property
    def pegadas(self) -> int:
        """Filas escritas en las pestañas, incluidas las que SALDOS no desglosa."""
        return len(self.colocadas)

    @property
    def total_renglones(self) -> int:
        return len(self.plantilla.renglones)

    def totales(self) -> dict:
        """Suma por divisa de lo que aparece en el reporte.

        Cuenta solo las filas que la hoja SALDOS desglosa. Las demás se pegan en
        su pestaña pero el reporte no las suma, así que incluirlas aquí daría un
        total que no cuadra con el que la usuaria ve impreso."""
        acumulado = {}
        for c in self.colocadas.values():
            if c.renglon is None:
                continue
            divisa = (c.linea.moneda or "MXN").upper()
            acumulado[divisa] = acumulado.get(divisa, 0.0) + (c.saldo or 0.0)
        return acumulado

    def cobertura_por_hoja(self) -> dict:
        """{pestaña: (renglones llenos, renglones que tiene)} para las 15 hojas.

        `bancos_faltantes` solo dice cuáles están en cero, que sirve para avisar
        pero no para saber POR DÓNDE VA la carga: con esto la pantalla puede
        mostrar las quince y cuáles ya llegaron, que es la pregunta de cada
        mañana. Se listan TODAS, también las que aún no tienen nada."""
        por_hoja: dict = {}
        for r in self.plantilla.renglones:
            datos = por_hoja.setdefault(r.hoja, [0, 0])
            datos[1] += 1
            if (r.hoja, r.hoja_fila) in self.colocadas:
                datos[0] += 1
        return {h: tuple(v) for h, v in sorted(por_hoja.items())}

    def bancos_faltantes(self) -> list:
        """Pestañas cuyos renglones quedaron TODOS vacíos.

        Distinto de 'faltan cuentas': si una pestaña entera está vacía es que no
        subieron ese archivo, y conviene decirlo así."""
        por_hoja = {}
        for r in self.plantilla.renglones:
            por_hoja.setdefault(r.hoja, [0, 0])[1] += 1
        for hoja, _fila in self.colocadas:
            if hoja in por_hoja:
                por_hoja[hoja][0] += 1
        return sorted(h for h, (puestos, _) in por_hoja.items() if puestos == 0)

    def resumen(self) -> str:
        partes = ["{} de {} renglones".format(self.llenos, self.total_renglones)]
        if self.nuevas:
            partes.append("{} cuenta(s) nueva(s)".format(len(self.nuevas)))
        if self.duplicados:
            partes.append("{} duplicada(s)".format(len(self.duplicados)))
        return " · ".join(partes)


# ---------------------------------------------------------------------------
# Casado
# ---------------------------------------------------------------------------

def _formas_cuenta(linea: LineaSaldo) -> list[str]:
    """Todas las formas en que el número de esta línea puede estar en el formato.

    Un mismo número se escribe distinto según el portal: Banamex reporta sucursal
    y cuenta en columnas separadas y el lector las concatena (394 + 7680454),
    pero el formato guarda solo la corta. Que el lector declare sus variantes deja
    el caso resuelto por número —la regla fuerte— en vez de depender de la cola."""
    formas = [_digitos(linea.cuenta)]
    corta = _digitos((linea.extra or {}).get("cuenta_corta", ""))
    if corta and corta not in formas:
        formas.append(corta)
    clabe = _digitos(linea.clabe)
    if len(clabe) == 18:
        embebida = clabe[6:17]
        if embebida and embebida not in formas:
            formas.append(embebida)
    return [f for f in formas if f]


# Índice de las exclusiones: prefijo de CLABE del banco -> números normalizados.
# Se agrupa por PREFIJO y no por el nombre del banco porque los lectores lo
# escriben de varias formas ('Scotiabank', 'Scotianbank', 'Banco Sabadell'); el
# prefijo de Banxico es el único identificador estable que ya maneja el módulo.
_EXCLUIDAS_POR_PREFIJO: dict[str, set] = {}
for _banco, _cuenta in CUENTAS_EXCLUIDAS:
    _pref = prefijo_banco(_banco) or prefijo_desde_texto(_banco)
    if _pref is None:
        raise RuntimeError(
            "CUENTAS_EXCLUIDAS: no se reconoce el banco «{}»".format(_banco))
    _EXCLUIDAS_POR_PREFIJO.setdefault(_pref, set()).add(_digitos(_cuenta))
del _banco, _cuenta, _pref


def colas_excluidas(plantilla: Plantilla) -> dict:
    """{prefijo: {terminación}} de las excluidas que se pueden reconocer por cola.

    Hace falta porque Banamex ENMASCARA el número: su reporte manda «**6227», no
    `8182756227`. Con la lista escrita en números completos, esas cuentas dejarían
    de reconocerse y volverían a aparecer como cuentas nuevas cada día.

    La guarda es la que hace segura una comparación de cuatro dígitos: una
    terminación solo se acepta si NO es la de ninguna cuenta que la plantilla
    tenga en esa pestaña. Así, ante la duda, siempre gana la plantilla y jamás se
    tira un saldo bueno tomándolo por excluido."""
    fuera = {}
    for prefijo, numeros in _EXCLUIDAS_POR_PREFIJO.items():
        hoja = plantilla.hoja_de_prefijo(prefijo)
        if not hoja:
            continue
        propias = plantilla.colas_de(hoja)
        colas = {n[-4:] for n in numeros if len(n) >= 4} - propias
        if colas:
            fuera[prefijo] = colas
    return fuera


def _excluida(linea: LineaSaldo, colas: dict = None) -> bool:
    """Si esta línea está en la lista de cuentas que no van al reporte.

    Se exige que COINCIDAN banco y número. Comparar solo el número invitaría a un
    choque entre bancos distintos —los formatos cortos son de 7 u 8 dígitos— y
    dejaría fuera un saldo bueno sin que nadie se entere.

    `colas` añade el reconocimiento por terminación para los portales que
    enmascaran el número (ver `colas_excluidas`); sin él solo se compara la forma
    completa."""
    clabe = _digitos(linea.clabe)
    prefijo = clabe[:3] if len(clabe) == 18 else prefijo_banco(linea.banco)
    formas = _formas_cuenta(linea)
    numeros = _EXCLUIDAS_POR_PREFIJO.get(prefijo)
    if numeros and any(f in numeros for f in formas):
        return True
    porcolas = (colas or {}).get(prefijo)
    if not porcolas:
        return False
    return any(f[-4:] in porcolas for f in formas if len(f) >= 4)


def _casar(linea: LineaSaldo, plantilla: Plantilla):
    """Aplica la cascada. Devuelve `(destino, regla)`; `destino` es None si no casó."""
    clabe = _digitos(linea.clabe)
    if len(clabe) == 18 and validar_clabe(clabe):
        destino = plantilla.buscar(clabe)
        if destino is not None:
            return destino, "clabe"

    formas = _formas_cuenta(linea)
    destino = plantilla.buscar(*formas)
    if destino is not None:
        return destino, "numero"

    # Las dos últimas reglas van acotadas a la pestaña de ese banco. El prefijo
    # sale de la CLABE cuando la hay y del nombre canónico cuando no.
    prefijo = clabe[:3] if len(clabe) == 18 else prefijo_banco(linea.banco)
    hoja = plantilla.hoja_de_prefijo(prefijo)

    # Banamex nombra sus cuentas por sucursal + terminación, no por el número
    # completo: la 7713101 de la sucursal 237 (PETROPLAZAS) puede llegar como
    # «2373101». Va antes que la cola porque compara siete u ocho dígitos, no
    # cuatro: identifica la fila, no la insinúa.
    destino = plantilla.buscar_por_sucursal(hoja, *formas)
    if destino is not None:
        return destino, "sucursal"

    destino = plantilla.buscar_por_cola(hoja, *formas)
    if destino is not None:
        return destino, "cola"

    return None, ""


def identificar(lineas, plantilla: Plantilla = None) -> Asignacion:
    """Coloca cada línea leída en su renglón del formato.

    `lineas` son las `LineaSaldo` que devolvieron los lectores. El orden importa
    solo para los duplicados: gana la primera que llegó."""
    plantilla = plantilla or cargar_plantilla()
    res = Asignacion(plantilla=plantilla)
    colas_fuera = colas_excluidas(plantilla)

    for linea in lineas or ():
        # Antes de casar: una cuenta excluida no debe llegar siquiera a la
        # plantilla. Si mañana se le abre un renglón, la exclusión seguiría
        # ganando y el renglón saldría vacío sin explicación.
        if _excluida(linea, colas_fuera):
            res.excluidas.append(Suelta(
                linea=linea,
                motivo="cuenta excluida del reporte a propósito"))
            continue
        destino, regla = _casar(linea, plantilla)
        if destino is None:
            res.nuevas.append(Suelta(
                linea=linea,
                motivo="la cuenta no está en ninguna pestaña del formato"))
            continue
        clave = (destino.hoja, destino.fila)
        previa = res.colocadas.get(clave)
        if previa is not None:
            res.duplicados.append(Suelta(
                linea=linea,
                motivo="ya venía en {}; no se suma".format(
                    previa.linea.origen or "otro archivo")))
            continue
        res.colocadas[clave] = Colocada(destino=destino, linea=linea,
                                        regla=regla)

    return res
