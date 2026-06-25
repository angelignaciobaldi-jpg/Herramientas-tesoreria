"""Identificación de datos en el texto de un estado de cuenta.

Extrae: CLABE, nombre del beneficiario, email de notificación y alias.

La CLABE se valida con su dígito de control oficial (algoritmo módulo 10 con
pesos 3-7-1), lo que permite descartar números de 18 dígitos que no son CLABE
y elegir la correcta cuando aparecen varias secuencias numéricas.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .catalogo_bancos import banco_desde_clabe

# --- Expresiones regulares ----------------------------------------------
_RE_EMAIL = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
# Corridas de dígitos que pueden traer espacios o guiones como separadores.
_RE_CORRIDA_DIGITOS = re.compile(r"\d[\d \-]{16,}\d")
_RE_ETIQUETA_CLABE = re.compile(r"CLABE[^0-9]{0,20}((?:\d[ \-]?){18})", re.IGNORECASE)
# Confusiones típicas del OCR letra->dígito (para recuperar CLABE dañadas).
_MAPA_OCR_DIGITO = str.maketrans(
    {"O": "0", "o": "0", "Q": "0", "D": "0", "I": "1", "l": "1", "|": "1",
     "S": "5", "B": "8", "b": "6", "Z": "2", "G": "6"}
)

# Etiquetas típicas que preceden al nombre del titular/beneficiario.
_ETIQUETAS_NOMBRE = [
    "beneficiario",
    "nombre del beneficiario",
    "titular",
    "nombre del titular",
    "nombre del cliente",
    "razón social",
    "razon social",
    "a nombre de",
    "cliente",
    "nombre",
]
_ETIQUETAS_ALIAS = ["alias", "referencia", "concepto", "apodo"]
_ETIQUETAS_EMAIL = ["correo", "email", "e-mail", "notificaci", "correo electr"]

# Palabras que NO deben confundirse con un nombre de persona/empresa.
_RUIDO_NOMBRE = re.compile(
    r"estado de cuenta|periodo|per[ií]odo|saldo|cuenta|clabe|sucursal|rfc|"
    r"fecha|p[aá]gina|tarjeta|n[uú]mero",
    re.IGNORECASE,
)


@dataclass
class DatosExtraidos:
    clabe: str = ""
    beneficiario: str = ""
    alias: str = ""
    email: str = ""
    banco: str = ""
    clabe_valida: bool = False
    clabes_candidatas: list[str] = field(default_factory=list)


# --- Validación de CLABE -------------------------------------------------
def validar_clabe(clabe: str) -> bool:
    """Valida una CLABE de 18 dígitos mediante su dígito de control."""
    if len(clabe) != 18 or not clabe.isdigit():
        return False
    pesos = (3, 7, 1)
    suma = sum((int(d) * pesos[i % 3]) % 10 for i, d in enumerate(clabe[:17]))
    control = (10 - (suma % 10)) % 10
    return control == int(clabe[17])


def extraer_clabes(texto: str) -> list[str]:
    """Devuelve las CLABE válidas encontradas, en orden de confiabilidad:

    1. Etiquetadas: precedidas por la palabra "CLABE".
    2. Sólidas: corridas de 18 dígitos contiguos (sin espacios), que es como
       suele imprimirse la CLABE.
    3. Fusionadas: solo si no hubo etiquetadas ni sólidas, se deslizan ventanas
       de 18 dígitos sobre corridas con espacios/guiones. Se evita este paso
       cuando ya hay candidatas sólidas porque la fusión de varios números
       (p. ej. cuenta + CLABE) puede producir falsos positivos con dígito de
       control válido por casualidad.
    """
    etiquetadas: list[str] = []
    solidas: list[str] = []
    fusionadas: list[str] = []

    for m in _RE_ETIQUETA_CLABE.finditer(texto):
        digitos = re.sub(r"\D", "", m.group(1))
        if validar_clabe(digitos) and digitos not in etiquetadas:
            etiquetadas.append(digitos)

    for token in re.findall(r"\d{18,}", texto):
        rango = 1 if len(token) == 18 else len(token) - 17
        for i in range(rango):
            sub = token[i : i + 18]
            if validar_clabe(sub) and sub not in solidas:
                solidas.append(sub)

    if not etiquetadas and not solidas:
        for m in _RE_CORRIDA_DIGITOS.finditer(texto):
            digitos = re.sub(r"\D", "", m.group(0))
            for i in range(len(digitos) - 17):
                sub = digitos[i : i + 18]
                # Estas rutas fusionan números vecinos (cuenta + CLABE) y pueden
                # producir secuencias con dígito de control válido por azar; se
                # exige además que el código de banco (3 primeros dígitos) exista
                # para descartar esos falsos positivos.
                if validar_clabe(sub) and banco_desde_clabe(sub) and sub not in fusionadas:
                    fusionadas.append(sub)

    fuzzy: list[str] = []
    if not etiquetadas and not solidas and not fusionadas:
        # Último recurso: el OCR a veces confunde dígitos con letras parecidas
        # (0->O, 1->I/l, 5->S, 8->B...). Se corrigen y se valida con el dígito
        # de control y un código de banco existente, que descartan cualquier
        # corrección equivocada.
        for m in re.finditer(r"[0-9OoQDIl|SBZGb]{18,}", texto):
            seg = m.group(0)
            for i in range(len(seg) - 17):
                sub = seg[i : i + 18].translate(_MAPA_OCR_DIGITO)
                if sub.isdigit() and validar_clabe(sub) and banco_desde_clabe(sub) and sub not in fuzzy:
                    fuzzy.append(sub)

    ordenadas: list[str] = []
    for c in (*etiquetadas, *solidas, *fusionadas, *fuzzy):
        if c not in ordenadas:
            ordenadas.append(c)
    return ordenadas


# --- Nombre del beneficiario --------------------------------------------
def _limpiar_nombre(valor: str) -> str:
    valor = valor.strip(" :.-\t")
    # Corta en el primer separador fuerte si la línea trae más cosas.
    valor = re.split(r"\s{3,}|\t|\| ", valor)[0].strip()
    return valor


def _parece_nombre(valor: str) -> bool:
    """Heurística: un nombre de persona/empresa es mayormente alfabético.

    Rechaza números (p. ej. "Número de cliente: 0123456"), códigos y valores
    dominados por dígitos.
    """
    v = valor.strip()
    if len(v) < 4:
        return False
    if _RUIDO_NOMBRE.search(v):
        return False
    letras = sum(c.isalpha() for c in v)
    digitos = sum(c.isdigit() for c in v)
    if letras < 4:               # debe tener suficientes letras
        return False
    if digitos >= 4:             # 4+ dígitos sugiere un número/código, no un nombre
        return False
    if digitos > letras:         # más dígitos que letras: no es un nombre
        return False
    return True


def _candidato_nombre(valor: str) -> str:
    """Devuelve el nombre limpio en mayúsculas si parece válido; si no, ''."""
    candidato = _limpiar_nombre(valor)
    return candidato.upper() if _parece_nombre(candidato) else ""


def _prefijo_es_numero(prefijo: str) -> bool:
    """True si la etiqueta viene precedida por 'número de', 'no. de', etc.,
    indicando que el valor que sigue es un número de cliente, no un nombre."""
    return bool(re.search(r"(n[uú]mero|nro\.?|no\.?|clave)\s*(de\s+)?$", prefijo))


def _beneficiario_por_etiqueta(lineas: list[str]) -> str:
    # Paso 1: valor en la MISMA línea que la etiqueta. Se recorre por orden de
    # prioridad de etiqueta (beneficiario/titular antes que el genérico "nombre").
    for etiqueta in _ETIQUETAS_NOMBRE:
        for linea in lineas:
            bajo = linea.lower()
            pos = bajo.find(etiqueta)
            if pos == -1 or _prefijo_es_numero(bajo[:pos]):
                continue
            candidato = _candidato_nombre(linea[pos + len(etiqueta):])
            if candidato:
                return candidato

    # Paso 2: etiqueta sola en su línea y el nombre en la línea siguiente
    # (p. ej. "Beneficiario:" seguido de "COMERCIALIZADORA ..."). La línea
    # siguiente no debe ser, a su vez, otro campo etiquetado.
    for etiqueta in _ETIQUETAS_NOMBRE:
        for i, linea in enumerate(lineas):
            bajo = linea.lower()
            pos = bajo.find(etiqueta)
            if pos == -1 or _prefijo_es_numero(bajo[:pos]):
                continue
            if linea[pos + len(etiqueta):].strip(" :.-\t"):
                continue  # ya traía algo en la línea (lo evaluó el paso 1)
            if i + 1 < len(lineas) and ":" not in lineas[i + 1]:
                candidato = _candidato_nombre(lineas[i + 1])
                if candidato:
                    return candidato
    return ""


# Palabras que delatan que una línea en mayúsculas es texto bancario o de
# domicilio, no el nombre del titular.
_PALABRAS_NO_NOMBRE = {
    "BANCO", "BANCOS", "HSBC", "SANTANDER", "BANAMEX", "BBVA", "BANCOMER",
    "BANORTE", "SCOTIABANK", "INBURSA", "BANREGIO", "AZTECA", "BANCOPPEL",
    "AFIRME", "MIFEL", "INVEX", "MULTIVA", "CITIBANAMEX", "INSTITUCION",
    "INSTITUCIÓN", "BANCA", "MULTIPLE", "MÚLTIPLE", "GRUPO", "FINANCIERO",
    "MEXICO", "MÉXICO", "ESTADO", "CUENTA", "CUENTAS", "RESUMEN", "GENERAL",
    "PRODUCTO", "PRODUCTOS", "SERVICIO", "SERVICIOS", "MONEDA", "NACIONAL",
    "PERIODO", "PERÍODO", "SUCURSAL", "PLAZA", "SALDO", "NOMINA", "NÓMINA",
    "SUPER", "CLIENTE", "NUMERO", "NÚMERO", "CLABE", "INTERBANCARIA",
    "REGISTRO", "FEDERAL", "CONTRIBUYENTES", "RFC", "CURP", "TELEFONO",
    "TELÉFONO", "FECHA", "CORTE", "SCANNED", "CAMSCANNER", "FLEXIBLE",
    "SIMPLE", "MICUENTA", "TARJETA", "COMISIONES", "INTERESES", "DEPOSITOS",
    "DEPÓSITOS", "RETIROS", "GAT", "IVA", "DIAS", "DÍAS", "TRANSCURRIDOS",
    "CODIGO", "CÓDIGO", "CHEQUES", "INFORMATIVO", "APLICABLE", "CONTRATO",
    "MEDIOS", "ACCESO", "RETIRO", "ABONOS", "CARGOS", "DOMICILIO",
    # encabezados de formularios/anexos
    "FIRMA", "COTITULAR", "TITULAR", "APELLIDOS", "APELLIDO", "RAZON", "RAZÓN",
    "SOCIAL", "OPERACION", "OPERACIÓN", "DESCRIPCION", "DESCRIPCIÓN", "ANEXO",
    "LIBRETON", "LIBRETÓN", "BASICO", "BÁSICO", "DIGITAL", "INDIVIDUAL",
    "MANCOMUNADA", "REGIMEN", "RÉGIMEN", "NACIONALIDAD", "SEXO", "CIVIL",
    "CASADO", "CASADA", "SOLTERO", "SOLTERA", "IDENTIFICACION", "IDENTIFICACIÓN",
    "CORREO", "ELECTRONICO", "ELECTRÓNICO", "EMAIL", "NACIMIENTO", "MEX",
    "DATOS", "CLIENTE", "PRODUCTO", "DESIGNADO", "DESIGNADOS", "EXTERIOR",
    "INTERIOR", "NUM", "NÚM", "HOGAR", "TIPO", "LIMITE", "LÍMITE", "ILIMITADOS",
    "COMO", "APARECE", "COMPLETO", "COMPLETOS", "CONSTANCIA", "SITUACION",
    "SITUACIÓN", "FISCAL", "DOCUMENTOS", "DOCUMENTO", "CIFRAS", "EXPRESADAS",
    "PESOS", "MEXICANOS", "MEXICANA", "GENERICO", "GENÉRICO", "PARENTESCO",
    "PORCENTAJE", "DIVISA", "COMERCIAL", "CARATULA", "CARÁTULA", "ACTIVACION",
    "ACTIVACIÓN", "CHEQUERA", "DISPERSION", "DISPERSIÓN", "PATRON", "PATRÓN",
    "DETALLE", "MOVIMIENTOS", "MOVIMIENTO", "OPER", "LIQ", "BENEFICIARIO",
    "BENEFICIARIOS", "ORDENANTE", "PERSONALIDAD", "JURIDICA", "JURÍDICA",
    "TRANSFERENCIA", "TRAS", "CTAS", "CTA", "CONCEPTO", "REFERENCIA", "FAVOR",
    "RESUMEN", "DESCRIPCION", "DESCRIPCIÓN", "PERIODO", "PERÍODO",
    # tokens de domicilio
    "COL", "COLONIA", "FRACC", "FRACCIONAMIENTO", "CENTRO", "PRIV", "PRIVADA",
    "CALLE", "AVENIDA", "BLVD", "BOULEVARD", "AND", "ANDADOR", "CDA", "CERRADA",
    "PROL", "PROLONGACION", "MZ", "LT", "EJ", "EJIDO",
}

# Encabezados que anteceden al nombre en formularios/anexos con formato tabular.
_ENCABEZADOS_NOMBRE = (
    "nombre(s) y apellidos",
    "nombre y apellidos",
    "nombre(s)",
    "apellidos / razón social",
    "apellidos / razon social",
    "apellidos y nombre",
)
_RE_TOKEN_ALFA = re.compile(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]{2,}$")


def _nombre_desde_fila(linea: str) -> str:
    """Extrae el nombre de una fila tabular como '20/03/2026 D7447770 JUAN PEREZ':
    toma la corrida más larga de palabras puramente alfabéticas (descarta fechas,
    números de cliente y códigos)."""
    mejor: list[str] = []
    actual: list[str] = []
    for token in linea.split():
        if _RE_TOKEN_ALFA.match(token) and token.upper() not in _PALABRAS_NO_NOMBRE:
            actual.append(token)
            if len(actual) > len(mejor):
                mejor = list(actual)
        else:
            actual = []
    nombre = " ".join(mejor)
    return nombre.upper() if len(mejor) >= 2 and _parece_nombre(nombre) else ""


def _beneficiario_tabular(lineas: list[str]) -> str:
    """Busca un encabezado de nombre y extrae el nombre de las líneas SIGUIENTES
    (el valor va debajo del encabezado; la propia línea de encabezado trae texto
    descriptivo como 'completo como aparece en la identificación')."""
    for i, linea in enumerate(lineas):
        if any(h in linea.lower() for h in _ENCABEZADOS_NOMBRE):
            for j in range(i + 1, min(i + 4, len(lineas))):
                nombre = _nombre_desde_fila(lineas[j])
                if nombre:
                    return nombre
    return ""
# Una línea con domicilio típicamente trae dígitos o un token de calle.
_RE_DIRECCION = re.compile(r"\d|^(C|CALLE|AV|AVE|AVENIDA|BLVD|COL|FRACC|PRIV|AND|CDA|PROL)\b")
# Prefijos de domicilio: una línea que empieza así no es un nombre de persona.
_PREFIJOS_DIRECCION = {
    "C", "AV", "AVE", "AVENIDA", "CALLE", "BLVD", "BOULEVARD", "CTO", "CIRCUITO",
    "AND", "ANDADOR", "PRIV", "PRIVADA", "FRACC", "COL", "CDA", "CERRADA", "PROL",
    "RETORNO", "PASEO", "CALZADA", "CALZ", "PASAJE", "PERIFERICO", "PERIFÉRICO",
}


def _candidato_linea_nombre(linea: str) -> str:
    """Devuelve el nombre si la línea es (casi) toda un nombre en mayúsculas.

    Tolera basura corta del OCR al final (p. ej. 'GRETEL ... FIERRO el' o
    '... PEREZ 1,') recortando los tokens finales que no sean palabras en
    mayúsculas, pero exige que el núcleo sean 2-6 palabras de nombre limpias."""
    tokens = linea.split()
    # Recorta basura final: tokens que no son palabras en MAYÚSCULAS.
    while tokens and not (tokens[-1].strip(".,").isupper() and _RE_TOKEN_ALFA.match(tokens[-1].strip(".,"))):
        tokens.pop()
    if not (2 <= len(tokens) <= 6):
        return ""
    palabras = [t.strip(".,") for t in tokens]
    # El núcleo debe ser todo palabras en mayúsculas, sin ruido ni prefijo calle.
    if not all(p.isupper() and _RE_TOKEN_ALFA.match(p) for p in palabras):
        return ""
    if palabras[0] in _PREFIJOS_DIRECCION:
        return ""
    if any(p in _PALABRAS_NO_NOMBRE for p in palabras):
        return ""
    nombre = " ".join(palabras)
    return nombre if _parece_nombre(nombre) else ""


def _beneficiario_sin_etiqueta(lineas: list[str]) -> str:
    """Detecta el nombre del titular sin etiqueta: el nombre en mayúsculas que
    encabeza el bloque de domicilio (una línea de domicilio aparece debajo)."""
    candidatos: list[str] = []
    for i, linea in enumerate(lineas):
        nombre = _candidato_linea_nombre(linea)
        if not nombre:
            continue
        if any(_RE_DIRECCION.search(s) for s in lineas[i + 1 : i + 4]):
            return nombre  # señal fuerte: nombre encima del domicilio
        candidatos.append(nombre)
    return candidatos[0] if candidatos else ""


def nombre_desde_archivo(nombre_archivo: str) -> str:
    """Deriva el nombre del beneficiario desde el nombre del archivo cuando este
    parece un nombre de persona (convención común: 'NOMBRE APELLIDO APELLIDO.pdf').
    Devuelve '' si el nombre de archivo no parece un nombre (p. ej. 'scan001')."""
    base = re.sub(r"\.[A-Za-z0-9]+$", "", nombre_archivo)  # quita la extensión
    # Nombres tipo hash/identificador (p. ej. '0BE29A2C2BBA0A19...') no son
    # nombres de persona; sus letras hexadecimales producían basura ("BE BBA").
    solo_alnum = re.sub(r"[^A-Za-z0-9]", "", base)
    if len(solo_alnum) >= 16 and re.fullmatch(r"[0-9A-Fa-f]+", solo_alnum):
        return ""
    base = re.sub(r"[_\-]+", " ", base)
    base = re.sub(r"\(.*?\)|\d+", " ", base)               # quita "(1)", números
    palabras = [p for p in base.split() if _RE_TOKEN_ALFA.match(p)]
    if len(palabras) >= 2 and _parece_nombre(" ".join(palabras)):
        return " ".join(palabras).upper()
    return ""


def extraer_beneficiario(texto: str) -> str:
    # El nombre del titular siempre está al inicio del documento; restringir a
    # la cabecera evita falsos positivos en movimientos y texto legal.
    lineas = [ln.strip() for ln in texto.splitlines() if ln.strip()]
    # La heurística nombre-sobre-domicilio es precisa (exige una línea de
    # domicilio debajo), así que puede mirar más abajo; las pasadas por
    # encabezado/etiqueta son más propensas a falsos positivos y se limitan más.
    return (
        _beneficiario_tabular(lineas[:50])
        or _beneficiario_sin_etiqueta(lineas[:70])
        or _beneficiario_por_etiqueta(lineas[:50])
    )


# --- Email ---------------------------------------------------------------
def extraer_email(texto: str) -> str:
    # Prioriza correos que aparecen cerca de etiquetas de notificación.
    for linea in texto.splitlines():
        bajo = linea.lower()
        if any(et in bajo for et in _ETIQUETAS_EMAIL):
            m = _RE_EMAIL.search(linea)
            if m:
                return m.group(0).lower()
    m = _RE_EMAIL.search(texto)
    return m.group(0).lower() if m else ""


# --- Alias ---------------------------------------------------------------
def extraer_alias(texto: str) -> str:
    for linea in texto.splitlines():
        bajo = linea.lower()
        for etiqueta in _ETIQUETAS_ALIAS:
            if etiqueta in bajo:
                resto = linea[bajo.index(etiqueta) + len(etiqueta):]
                candidato = _limpiar_nombre(resto)
                if 2 <= len(candidato) <= 60:
                    return candidato
    return ""


# --- Orquestador ---------------------------------------------------------
def extraer_datos(texto: str) -> DatosExtraidos:
    """Identifica todos los campos a partir del texto del estado de cuenta."""
    clabes = extraer_clabes(texto)
    clabe = clabes[0] if clabes else ""
    beneficiario = extraer_beneficiario(texto)
    # Por decisión del usuario, el alias es el mismo nombre del beneficiario
    # (queda editable en la interfaz para cada registro).
    alias = beneficiario

    return DatosExtraidos(
        clabe=clabe,
        beneficiario=beneficiario,
        alias=alias,
        email=extraer_email(texto),
        banco=banco_desde_clabe(clabe),
        clabe_valida=bool(clabe) and validar_clabe(clabe),
        clabes_candidatas=clabes,
    )
