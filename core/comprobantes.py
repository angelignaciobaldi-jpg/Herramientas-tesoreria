"""Casado de comprobantes de pago contra los movimientos que los originaron.

Extraído de `ui/dispersion_no_pemex.py` para poder reusarlo desde la pantalla de
devoluciones: la mecánica es la misma (leer PDFs, atribuir cada lectura a su
archivo y decidir a qué movimiento corresponde) y duplicarla garantizaría que las
dos copias se separen con el tiempo.

Lo que vive aquí es lo que NO depende de una pantalla:

  - Normalización de nombres de archivo, para resolver el `documento_lectura` que
    devuelve el extractor contra las rutas que se le enviaron.
  - Reparto de las lecturas por archivo.
  - Las 3 reglas de coincidencia (cuenta origen, cuenta destino, importe).

Lo que NO vive aquí es de dónde salen los datos del movimiento: cada pantalla
arma su `Objetivo` (qué cuentas y qué total hay que casar) y se lo pasa a
`evaluar_coincidencia`. Así la dispersión sigue resolviendo sus cuentas por el
catálogo y las devoluciones pueden resolver las suyas por su propia asignación,
sin que este módulo sepa de ninguna de las dos.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field

# Dígitos que se comparan de una cuenta. Los comprobantes suelen enmascararlas
# (p. ej. '*0012'), así que solo se pueden casar por el final.
DIGITOS_CUENTA = 4
# Tolerancia al comparar importes: son pesos con centavos, no enteros.
TOLERANCIA_IMPORTE = 0.01


# --- Normalización -------------------------------------------------------
def ultimos_digitos(texto, n: int = DIGITOS_CUENTA) -> str:
    """Últimos `n` dígitos de un texto, ignorando cualquier no-dígito (espacios,
    guiones, y máscaras como '*0012'). '' si no hay dígitos."""
    d = re.sub(r"\D", "", str(texto or ""))
    return d[-n:] if d else ""


def claves_cuenta(texto, n: int = DIGITOS_CUENTA) -> set:
    """Colas de `n` dígitos con las que una MISMA cuenta puede aparecer en un
    comprobante bancario. Conjunto vacío si el texto no trae dígitos.

    Devuelve VARIAS porque el mismo número se escribe de dos formas y hay que
    casar cualquiera contra cualquiera: la CLABE completa (18 dígitos) y la
    cuenta enmascarada por sus últimos dígitos ('****7012'), que es como la
    imprimen los comprobantes.

    No son intercambiables: la CLABE lleva el DÍGITO VERIFICADOR al final, así
    que sus últimos 4 están desplazados respecto a los de la cuenta
    (012320001103245316 -> '5316', pero la cuenta que lleva dentro termina en
    '4531'). Comparar solo la cola de la CLABE hacía fallar la regla del
    beneficiario en todos los comprobantes que traen la cuenta enmascarada. De
    una CLABE se derivan las dos: su propia cola y la de la cuenta embebida
    (posiciones 7 a 17).
    """
    d = re.sub(r"\D", "", str(texto or ""))
    if not d:
        return set()
    claves = {d[-n:]}
    if len(d) == 18:                 # CLABE: 3 banco + 3 plaza + 11 cuenta + 1 dv
        claves.add(d[6:17][-n:])
    return claves

def norm_nombre_doc(nombre: str) -> str:
    """Normaliza un nombre de archivo para comparar el 'documento_lectura' que
    devuelve el extractor contra los PDFs subidos: minúsculas, sin espacios en los
    extremos y sin la extensión '.pdf'. Tolera diferencias de caja/extensión entre
    lo subido y lo que reporta la API."""
    s = (nombre or "").strip().lower()
    return s[:-4] if s.endswith(".pdf") else s


def reducir_nombre_doc(nombre: str) -> str:
    """Versión aún más laxa de `norm_nombre_doc`: solo letras y dígitos. Es el
    último recurso para casar el 'documento_lectura' con el archivo enviado, por si
    la API cambia espacios por guiones/guiones bajos al reportarlo. Las páginas
    separadas terminan en 'p1', 'p2'…, que sobreviven a esta reducción y las
    mantienen distinguibles entre sí."""
    return re.sub(r"[^a-z0-9]", "", norm_nombre_doc(nombre))


# --- Resolución de 'documento_lectura' -> ruta del archivo ---------------
def indices_por_nombre(rutas_pdf: list[str]) -> tuple[dict, dict, dict]:
    """Índices para resolver el 'documento_lectura' que devuelve el extractor
    contra las rutas enviadas, de más a menos estricto: nombre exacto, normalizado
    (sin caja ni '.pdf') y reducido a solo letras y dígitos.

    Los niveles laxos DESCARTAN las claves ambiguas: si dos archivos colapsan al
    mismo nombre reducido, ninguno se resuelve por ahí y se cae al criterio
    siguiente, en vez de adivinar y adjuntarle a un movimiento el comprobante de
    otro."""
    exacto: dict[str, str] = {}
    norm: dict[str, str] = {}
    reducido: dict[str, str] = {}
    dup_norm: set[str] = set()
    dup_red: set[str] = set()
    for ruta in rutas_pdf:
        base = os.path.basename(ruta)
        exacto.setdefault(base, ruta)
        k = norm_nombre_doc(base)
        if k in norm and norm[k] != ruta:
            dup_norm.add(k)
        norm.setdefault(k, ruta)
        k2 = reducir_nombre_doc(base)
        if k2 in reducido and reducido[k2] != ruta:
            dup_red.add(k2)
        reducido.setdefault(k2, ruta)
    for k in dup_norm:
        norm.pop(k, None)
    for k in dup_red:
        reducido.pop(k, None)
    return exacto, norm, reducido


def resolver_ruta(nombre: str, indices: tuple[dict, dict, dict]) -> str | None:
    """Ruta del PDF del que salió una lectura, por su 'documento_lectura'."""
    exacto, norm, reducido = indices
    n = (nombre or "").strip()
    if not n:
        return None
    return (exacto.get(n) or norm.get(norm_nombre_doc(n))
            or reducido.get(reducir_nombre_doc(n)))


def repartir_lecturas(
    comprobantes: list[dict], rutas_pdf: list[str],
) -> tuple[dict[str, list[dict]], int]:
    """Agrupa las lecturas del extractor por la RUTA del archivo del que salieron.

    Devuelve `(por_archivo, sin_archivo)`; `por_archivo` trae una entrada por cada
    ruta enviada (lista vacía si no devolvió lectura) y `sin_archivo` cuenta las
    lecturas cuyo 'documento_lectura' no correspondió a ningún archivo enviado."""
    indices = indices_por_nombre(rutas_pdf)
    por_archivo: dict[str, list[dict]] = {r: [] for r in rutas_pdf}
    sin_archivo = 0
    for c in comprobantes:
        ruta = resolver_ruta(c.get("documento_lectura"), indices)
        if ruta is None:
            sin_archivo += 1
            continue
        por_archivo[ruta].append(c)
    return por_archivo, sin_archivo


# --- Reglas de coincidencia ---------------------------------------------
@dataclass
class Objetivo:
    """Lo que hay que casar de UN movimiento, ya resuelto por su pantalla.

    - `origenes`: identificadores con los que el comprobante puede referirse a la
      cuenta de retiro. Conviene incluir TODOS los que se conozcan (nombre, número
      de cuenta y CLABE): un comprobante puede traer el número de cuenta mientras
      la pantalla guarda la CLABE, y como la CLABE termina en dígito verificador,
      sus últimos dígitos NO coinciden con los del número de cuenta.
    - `beneficiarios`: CLABE(s)/cuenta(s) destino admisibles.
    - `total`: importe esperado, en la misma moneda que reporta el comprobante.
    """

    origenes: set[str] = field(default_factory=set)
    beneficiarios: set[str] = field(default_factory=set)
    total: float = 0.0


def evaluar_coincidencia(comprobante: dict, objetivo: Objetivo) -> dict:
    """Evalúa las 3 reglas de vinculación de un comprobante (dict con
    `cuenta_origen`, `cuenta_destino` e `importe`) contra un `Objetivo`.

    Devuelve {'origen','beneficiario','total','coincide'} (bools); 'coincide' = las
    tres reglas se cumplen.

    Las cuentas del comprobante vienen enmascaradas (p. ej. '*0012'), así que la
    comparación es por los ÚLTIMOS dígitos: cuenta origen (regla 1) y cuenta
    destino vs beneficiario (regla 2). El total (regla 3) se compara con tolerancia
    de centavos.

    Cada lado aporta VARIAS colas posibles (ver `claves_cuenta`) y basta con
    que se crucen: una CLABE y la cuenta que lleva dentro terminan distinto,
    y el comprobante puede traer cualquiera de las dos."""
    origen_comp = claves_cuenta(comprobante.get("cuenta_origen"))
    origen_ok = bool(origen_comp) and any(
        origen_comp & claves_cuenta(o) for o in objetivo.origenes)
    destino_comp = claves_cuenta(comprobante.get("cuenta_destino"))
    benef_ok = bool(destino_comp) and any(
        destino_comp & claves_cuenta(b) for b in objetivo.beneficiarios)
    importe = comprobante.get("importe")
    total_ok = importe is not None and abs(
        float(importe) - objetivo.total) < TOLERANCIA_IMPORTE
    return {
        "origen": origen_ok, "beneficiario": benef_ok, "total": total_ok,
        "coincide": origen_ok and benef_ok and total_ok,
    }


# --- Vinculación archivo <-> movimiento ----------------------------------
@dataclass
class ResultadoVinculacion:
    """Qué pasó al repartir un lote de comprobantes entre los movimientos.

    - `asignados`: {clave del movimiento -> ruta del PDF}.
    - `sin_movimiento`: archivos CON lectura que no casaron con ningún movimiento.
    - `sin_asignar`: archivos que quedaron libres (incluye los anteriores). Se
      ofrecen para asignar a mano.
    - `sin_archivo`: lecturas cuyo 'documento_lectura' no correspondió a ningún
      archivo enviado.
    """

    asignados: dict = field(default_factory=dict)
    sin_movimiento: list = field(default_factory=list)
    sin_asignar: list = field(default_factory=list)
    sin_archivo: int = 0


def vincular(
    objetivos: list[tuple],
    comprobantes: list[dict],
    rutas_pdf: list[str],
    ocupados: set | None = None,
) -> ResultadoVinculacion:
    """Asigna cada ARCHIVO al movimiento que le corresponde, uno a uno.

    - `objetivos`: [(clave, Objetivo)] en el orden en que se muestran; `clave`
      identifica al movimiento para quien llama (índice, folio, lo que sea).
    - `comprobantes`: las lecturas (del extractor o de `core.lector_comprobantes`).
    - `rutas_pdf`: los archivos que se enviaron a leer.
    - `ocupados`: claves que YA tienen comprobante y no deben reasignarse.

    Se itera por ARCHIVO y no por lectura: lo que se sube al SIPP es un archivo,
    así que una lectura cuyo 'documento_lectura' no se pueda resolver no debe
    producir vínculo. Cada movimiento recibe a lo sumo un archivo y cada archivo va
    a un solo movimiento: ante la duda es preferible dejarlo sin asignar —queda a
    la vista para resolverlo a mano— que adjudicarle a alguien el comprobante de
    otro.
    """
    por_archivo, sin_archivo = repartir_lecturas(comprobantes, rutas_pdf)
    res = ResultadoVinculacion(sin_archivo=sin_archivo)
    tomados = set(ocupados or ())
    usados: set[str] = set()
    for ruta in rutas_pdf:
        # Un archivo repetido en la lista no debe vincularse dos veces: se
        # subiria el MISMO comprobante a dos solicitudes distintas.
        if ruta in usados:
            continue
        lecturas = por_archivo.get(ruta) or []
        clave = next(
            (k for k, obj in objetivos
             if k not in tomados
             and any(evaluar_coincidencia(c, obj)["coincide"] for c in lecturas)),
            None,
        )
        if clave is None:
            res.sin_asignar.append(ruta)
            if lecturas:
                res.sin_movimiento.append(ruta)
            continue
        res.asignados[clave] = ruta
        tomados.add(clave)
        usados.add(ruta)
    return res
