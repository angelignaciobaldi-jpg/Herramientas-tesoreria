"""Catálogo de bancos por código CLABE (3 primeros dígitos).

Los primeros 3 dígitos de una CLABE identifican a la institución según el
catálogo del Banco de México. Se incluyen las instituciones más frecuentes;
puede ampliarse según se necesite.
"""

CATALOGO_BANCOS = {
    "002": "Banamex",
    "006": "Banco Nacional de Comercio Exterior",
    "009": "Banobras",
    "012": "BBVA México",
    "014": "Santander",
    "019": "Banjército",
    "021": "HSBC",
    "030": "Banco del Bajío",
    "036": "Inbursa",
    "042": "Mifel",
    "044": "Scotiabank",
    "058": "Banregio",
    "059": "Invex",
    "060": "Bansí",
    "062": "Afirme",
    "072": "Banorte",
    "106": "Bank of America",
    "108": "MUFG",
    "110": "JP Morgan",
    "112": "Banco Monex",
    "113": "Ve por Más",
    "124": "Banco Citi México",
    "127": "Banco Azteca",
    "128": "Autofin",
    "129": "Barclays",
    "130": "Compartamos",
    "132": "Multiva Banco",
    "133": "Actinver",
    "137": "BanCoppel",
    "138": "ABC Capital",
    "140": "Consubanco",
    "143": "CIBanco",
    "147": "Bankaool",
    "148": "Banco PagaTodo",
    "150": "Inmobiliario Mexicano",
    "151": "Donde",
    "152": "Bancrea",
    "154": "Banco Covalto",
    "155": "ICBC",
    "156": "Sabadell",
    "157": "Shinhan",
    "158": "Mizuho",
    "159": "Bank of China",
    "160": "Banco S3",
    "166": "Banco del Bienestar",
    "167": "Hey Banco",
    "168": "Sociedad Hipotecaria Federal",
    "600": "Monexcb",
    "601": "GBM",
    "602": "Masari",
    "605": "Valué",
    "608": "Vector",
    "613": "Multiva Casa de Bolsa",
    "616": "Finamex",
    "617": "Valmex",
    "620": "Profuturo",
    "630": "Intercam Banco",
    "631": "CI Bolsa",
    "636": "HDI Seguros",
    "637": "Order",
    "638": "Akala",
    "640": "JP Morgan Casa de Bolsa",
    "642": "Reforma",
    "646": "STP",
    "647": "Telecomunicaciones de México",
    "648": "Evercore",
    "649": "Skandia",
    "652": "Asea",
    "653": "Kuspit",
    "655": "Sofiexpress",
    "656": "Unagra",
    "659": "Opciones Empresariales del Noroeste",
    "661": "Alternativos",
    "670": "Libertad",
    "674": "Axa",
    "677": "Caja Pop Mexicana",
    "679": "Fnd",
    "684": "Operadora de Recursos Reforma",
    "686": "Storm-Fin",
    "689": "Fompet",
    "703": "Tesored",
    "706": "Arcus",
    "710": "NVIO",
    "722": "Mercado Pago",
    "723": "Cuenca",
    "728": "SPIN by OXXO",
    "812": "BBVA Bancomer (Seguros)",
    "846": "STP (legado)",
}


def banco_desde_clabe(clabe: str) -> str:
    """Devuelve el nombre del banco para una CLABE, o cadena vacía si no aplica."""
    if not clabe or len(clabe) < 3:
        return ""
    return CATALOGO_BANCOS.get(clabe[:3], "")


# --- Validación CLABE <-> banco reportado --------------------------------
# El SIPP guarda el banco del beneficiario como TEXTO, capturado aparte de la
# CLABE, así que los dos datos pueden contradecirse: se han visto solicitudes con
# una CLABE de BanCoppel (137) y "BBVA BANCOMER" como banco. Comparar ambos avisa
# de un registro mal capturado ANTES de dispersarle dinero.
#
# Los nombres no coinciden literalmente entre fuentes ("BBVA México" aquí, "BBVA
# BANCOMER" en el SIPP), así que se comparan por palabras significativas.

# Palabras que no distinguen a una institución y solo agregan ruido.
_PALABRAS_VACIAS = {
    "BANCO", "BANCA", "MULTIPLE", "MULTIPLE", "INSTITUCION", "DE", "DEL", "LA",
    "EL", "Y", "SA", "SAB", "CV", "SNC", "IBM", "GRUPO", "FINANCIERO", "SOFOM",
    "ER", "ENR", "MEXICO", "MEXICANO", "NACIONAL", "S", "A", "C", "V",
}

# Nombres comerciales distintos que designan a la MISMA institución. La clave se
# normaliza igual que el resto; el valor es la palabra canónica con la que se
# compara.
_ALIAS = {
    "BANCOMER": "BBVA",
    "CITIBANAMEX": "BANAMEX",
    "CITIBANEX": "BANAMEX",
    "CITI": "BANAMEX",
    "SERFIN": "SANTANDER",
    "IXE": "BANORTE",
    "BANCOMEXT": "COMERCIOEXTERIOR",
}


def _palabras(nombre: str) -> set[str]:
    """Palabras significativas de un nombre de banco, ya normalizadas (mayúsculas,
    sin acentos ni signos) y con los alias comerciales resueltos."""
    import re as _re
    import unicodedata as _ud

    base = _ud.normalize("NFKD", str(nombre or ""))
    base = "".join(c for c in base if not _ud.combining(c)).upper()
    crudas = [p for p in _re.split(r"[^A-Z0-9]+", base) if p]
    palabras = set()
    for p in crudas:
        p = _ALIAS.get(p, p)
        if p not in _PALABRAS_VACIAS and len(p) >= 3:
            palabras.add(p)
    return palabras


def codigo_desde_clabe(clabe: str) -> str:
    """Los 3 dígitos de institución de una CLABE, o '' si no se pueden leer."""
    import re as _re

    digitos = _re.sub(r"\D", "", str(clabe or ""))
    return digitos[:3] if len(digitos) >= 3 else ""


def coincide_banco(clabe: str, banco_reportado: str) -> bool | None:
    """¿El banco reportado concuerda con el que dice la CLABE?

    - `True`  : concuerdan.
    - `False` : se contradicen (registro sospechoso: la CLABE dice un banco y el
                SIPP otro).
    - `None`  : NO se puede juzgar, y hay que tratarlo distinto de `False`. Pasa
                si falta alguno de los dos datos, si la CLABE es muy corta, si su
                prefijo no está en el catálogo, o si el nombre reportado no deja
                ninguna palabra significativa. Marcar esos casos como error
                llenaría la pantalla de falsas alarmas.
    """
    esperado = banco_desde_clabe(codigo_desde_clabe(clabe))
    if not esperado or not str(banco_reportado or "").strip():
        return None
    palabras_esperadas = _palabras(esperado)
    palabras_reportadas = _palabras(banco_reportado)
    if not palabras_esperadas or not palabras_reportadas:
        return None
    if palabras_esperadas & palabras_reportadas:
        return True
    # Una puede estar contenida en la otra sin ser la misma palabra:
    # 'SANTANDER' vs 'BANCOSANTANDER' al juntar todo.
    junta_e = "".join(sorted(palabras_esperadas))
    junta_r = "".join(sorted(palabras_reportadas))
    if junta_e in junta_r or junta_r in junta_e:
        return True
    return False


def banco_a_mostrar(clabe: str, banco_reportado: str) -> tuple[str, bool]:
    """Qué banco destino mostrar en pantalla, y si hubo que corregirlo.

    Devuelve `(nombre, corregido)`:
      - `nombre`: el banco que dice la CLABE siempre que su prefijo esté en el
        catálogo; si no, el que reportó el SIPP (para no dejar la celda vacía
        teniendo el dato).
      - `corregido`: True solo cuando ambos datos se contradicen, es decir cuando
        lo mostrado NO es lo que venía en la solicitud.

    Manda la CLABE porque es la que describe lo que va a pasar: el pago viaja por
    ella y el TXT se arma con su prefijo (`core.exportador_devoluciones` elige
    entre el layout de mismo banco y el de SPEI justamente por ahí). Mostrar el
    banco del SIPP hacía creer que el dinero iría a un banco al que no va.

    `corregido=True` NO significa que el registro ya esté bien: la contradicción
    también puede venir de una CLABE equivocada, y en ese caso el pago se iría a
    otro lado. Por eso quien llama debe señalarlo para que se revise.
    """
    segun_clabe = banco_desde_clabe(codigo_desde_clabe(clabe))
    reportado = str(banco_reportado or "").strip()
    corregido = coincide_banco(clabe, reportado) is False
    return (segun_clabe or reportado or ""), corregido
