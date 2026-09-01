"""Pruebas de `core.comprobantes` y `core.lector_comprobantes`.

Sin dependencias de test externas: se corre con `python scripts/prueba_comprobantes.py`
y sale con código 1 si algo falla, para poder colgarlo del pipeline igual que
`smoke_import.py`.

Los casos salen de comprobantes BBVA Net Cash REALES (un pago mismo banco y uno
interbancario del mismo lote), incluidas sus dos trampas: el nombre del archivo
trae el total del lote y no el del pago, y los dos comprobantes del lote comparten
importe, así que solo la cuenta destino los distingue.
"""

from __future__ import annotations

import os
import sys

_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)

from core import comprobantes as c  # noqa: E402
from core import lector_comprobantes as lc  # noqa: E402
from core import catalogo_bancos as cb  # noqa: E402

# Lecturas equivalentes a las de los dos comprobantes reales del lote GC MOTORS.
MISMO_BANCO = {
    "documento_lectura": "DEV GC MOTORS $17,468.00-1.pdf",
    "cuenta_origen": "000000000117421184",     # número de cuenta, con ceros
    "cuenta_destino": "012730028914386037",    # CLABE del beneficiario
    "importe": 3227.00,
}
INTERBANCARIO = {
    "documento_lectura": "DEV GC MOTORS $17,468.00-2.pdf",
    "cuenta_origen": "000000000117421184",
    "cuenta_destino": "137730104690721058",
    "importe": 3227.00,
}

_fallos: list[str] = []


def check(condicion: bool, descripcion: str) -> None:
    print(f"  {'OK  ' if condicion else 'FALLA'}  {descripcion}")
    if not condicion:
        _fallos.append(descripcion)


def prueba_ultimos_digitos() -> None:
    print("\nultimos_digitos")
    check(c.ultimos_digitos("012730028914386037") == "6037", "CLABE -> últimos 4")
    check(c.ultimos_digitos("*0012") == "0012", "cuenta enmascarada")
    check(c.ultimos_digitos("PETRO SMART HERMOSILLO BBVA") == "",
          "nombre sin dígitos -> cadena vacía")
    check(c.ultimos_digitos(None) == "", "None -> cadena vacía")
    check(c.ultimos_digitos("117421184", 6) == "421184", "n configurable")


def prueba_nombres() -> None:
    print("\nnormalización de nombres")
    check(c.norm_nombre_doc("Comprobante.PDF") == "comprobante", "quita caja y .pdf")
    check(c.reducir_nombre_doc("DEV GC MOTORS $17,468.00-1.pdf") == "devgcmotors17468001",
          "reduce a letras y dígitos")
    check(c.reducir_nombre_doc("pago p1.pdf") != c.reducir_nombre_doc("pago p2.pdf"),
          "las páginas separadas siguen siendo distinguibles")


def prueba_resolucion_rutas() -> None:
    print("\nresolución documento_lectura -> ruta")
    rutas = [r"C:\x\DEV GC MOTORS $17,468.00-1.pdf",
             r"C:\x\DEV GC MOTORS $17,468.00-2.pdf"]
    idx = c.indices_por_nombre(rutas)
    check(c.resolver_ruta("DEV GC MOTORS $17,468.00-1.pdf", idx) == rutas[0],
          "nombre exacto")
    check(c.resolver_ruta("dev gc motors $17,468.00-2", idx) == rutas[1],
          "sin caja ni extensión")
    # El nivel laxo existe precisamente para esto: si la API reporta el nombre
    # con guiones bajos en vez de espacios/signos, debe seguir resolviendo.
    check(c.resolver_ruta("DEV_GC_MOTORS_17468001.pdf", idx) == rutas[0],
          "tolera espacios y signos cambiados por guiones bajos")
    check(c.resolver_ruta("otro comprobante.pdf", idx) is None,
          "un nombre ajeno no inventa ruta")
    check(c.resolver_ruta("", idx) is None, "nombre vacío -> None")

    # Ambigüedad: dos archivos que colapsan al mismo nombre reducido no deben
    # resolverse por el nivel laxo (adjudicarían el comprobante equivocado).
    ambiguas = [r"C:\a\pago-1.pdf", r"C:\b\pago 1.pdf"]
    idx2 = c.indices_por_nombre(ambiguas)
    check(c.resolver_ruta("pago1", idx2) is None,
          "nombres ambiguos NO se resuelven por el nivel laxo")
    check(c.resolver_ruta("pago-1.pdf", idx2) == ambiguas[0],
          "pero el nombre exacto sigue resolviendo")


def prueba_reparto() -> None:
    print("\nreparto de lecturas por archivo")
    rutas = [r"C:\x\DEV GC MOTORS $17,468.00-1.pdf",
             r"C:\x\DEV GC MOTORS $17,468.00-2.pdf",
             r"C:\x\sin lectura.pdf"]
    por_archivo, sin_archivo = c.repartir_lecturas(
        [MISMO_BANCO, INTERBANCARIO, {"documento_lectura": "ajeno.pdf"}], rutas)
    check(len(por_archivo) == 3, "una entrada por cada ruta enviada")
    check(por_archivo[rutas[0]] == [MISMO_BANCO], "cada lectura a su archivo")
    check(por_archivo[rutas[2]] == [], "archivo sin lectura queda con lista vacía")
    check(sin_archivo == 1, "la lectura sin archivo se cuenta aparte")


def prueba_coincidencia() -> None:
    print("\nlas 3 reglas de coincidencia")
    # La pantalla aporta TODOS los identificadores de la cuenta origen: aquí el
    # número de cuenta y la CLABE. La CLABE termina en dígito verificador, así que
    # sus últimos 4 NO coinciden con los del número de cuenta.
    objetivo = c.Objetivo(
        origenes={"0117421184", "012180001174211843"},
        beneficiarios={"012730028914386037"},
        total=3227.00,
    )
    r = c.evaluar_coincidencia(MISMO_BANCO, objetivo)
    check(r["coincide"], "el comprobante correcto casa por las 3 reglas")
    check(r["origen"] and r["beneficiario"] and r["total"], "las 3 dan verdadero")

    # El interbancario tiene el MISMO importe y la MISMA cuenta origen: solo la
    # cuenta destino evita que se le adjudique al movimiento equivocado.
    r2 = c.evaluar_coincidencia(INTERBANCARIO, objetivo)
    check(not r2["coincide"], "el otro comprobante del lote NO casa")
    check(r2["origen"] and r2["total"] and not r2["beneficiario"],
          "falla solo por beneficiario (mismo importe y misma cuenta origen)")

    # Si la pantalla solo conociera la CLABE, la regla de origen fallaría: es la
    # razón por la que Objetivo.origenes debe llevar también el número de cuenta.
    solo_clabe = c.Objetivo(origenes={"012180001174211843"},
                            beneficiarios={"012730028914386037"}, total=3227.00)
    check(not c.evaluar_coincidencia(MISMO_BANCO, solo_clabe)["origen"],
          "con solo la CLABE, la cuenta origen no casa (dígito verificador)")

    print("\n  bordes")
    check(not c.evaluar_coincidencia(
        {"cuenta_origen": "", "cuenta_destino": "", "importe": 3227.00},
        objetivo)["coincide"], "cuentas vacías no casan")
    check(not c.evaluar_coincidencia(
        {**MISMO_BANCO, "importe": None}, objetivo)["total"],
        "importe None no casa")
    check(c.evaluar_coincidencia(
        {**MISMO_BANCO, "importe": 3227.004}, objetivo)["total"],
        "diferencia menor a un centavo sí casa")
    check(not c.evaluar_coincidencia(
        {**MISMO_BANCO, "importe": 3227.02}, objetivo)["total"],
        "diferencia de dos centavos no casa")
    check(not c.evaluar_coincidencia(
        MISMO_BANCO, c.Objetivo())["coincide"],
        "objetivo vacío no casa con nada")


# --- Lector local de comprobantes ---------------------------------------
# Los comprobantes reales estan en .gitignore (llevan datos de clientes), asi que
# las pruebas arman PDFs sinteticos con el MISMO texto que emite BBVA Net Cash,
# incluidos los acentos ('depósito'/'aplicación') que el lector debe tolerar.
_PLANTILLA_MISMO_BANCO = """BBVA Net Cash
  BBVA Net Cash - Pago Mismo Banco
Datos de la operación
Tipo de operación: Grupo Pago Mismo Banco
  Descripción: DEVOLUCIONES GC MOTORS
  Importe: 3,227.00
  Cuenta de retiro: 000000000117421184
  Cuenta de depósito: 012730028914386037
  Titular de la cuenta: EMPRESA QUE PAGA SA
  Titular de la cuenta: BENEFICIARIO UNO
  Fecha de creación: 04/08/2026
  Fecha de aplicación: 04/08/2026
  Motivo de pago: APOYO PLACAS
  Folio de firma: 9485458461
  Folio único: I333202608041750400010641840
  Estado: {estado}
"""

_PLANTILLA_INTERBANCARIO = """BBVA Net Cash
  BBVA Net Cash - Pago Interbancario
Datos de la operación
Tipo de operación: Grupo Pago Interbancario
  Importe: 1,500.50
  Cuenta de retiro: 000000000117421184
  Cuenta de depósito: 137730104690721058
  Banco beneficiario: BANCOPPEL
  Fecha de creación: 03/08/2026
  Fecha de aplicación: 04/08/2026
  Concepto de pago: APOYO PLACAS
  Referencia: 0023626
  Clave de rastreo: 002601002608040000587607
  Folio interbancario: 0000587607
  Estado: Operado
"""


def _pdf_con_texto(ruta: str, texto: str) -> str:
    """Escribe un PDF de una pagina con `texto` en su capa de texto."""
    import pymupdf
    doc = pymupdf.open()
    pagina = doc.new_page()
    y = 60
    for linea in texto.splitlines():
        pagina.insert_text((50, y), linea, fontsize=9)
        y += 14
    doc.save(ruta)
    doc.close()
    return ruta


def prueba_lector() -> None:
    import tempfile
    print("\nlector local de comprobantes")
    tmp = tempfile.mkdtemp(prefix="prueba_comprobantes_")

    mb = _pdf_con_texto(os.path.join(tmp, "mismo banco.pdf"),
                        _PLANTILLA_MISMO_BANCO.format(estado="Operado"))
    ib = _pdf_con_texto(os.path.join(tmp, "interbancario.pdf"),
                        _PLANTILLA_INTERBANCARIO)

    lecturas, errores = lc.leer_varios([mb, ib])
    check(len(lecturas) == 2 and not errores, "lee las dos variantes sin errores")

    a, b = lecturas
    check(a["cuenta_origen"] == "000000000117421184",
          "cuenta de retiro (numero de cuenta con ceros)")
    check(a["cuenta_destino"] == "012730028914386037", "cuenta de deposito (CLABE)")
    check(a["importe"] == 3227.00, "importe con separador de miles -> float")
    check(b["importe"] == 1500.50, "importe con centavos")
    check(a["documento_lectura"] == "mismo banco.pdf",
          "documento_lectura = nombre del archivo, como el extractor")

    # La forma debe coincidir con la del extractor para que el casado sea indistinto.
    for clave in ("documento_lectura", "cuenta_origen", "cuenta_destino", "importe"):
        check(clave in a, f"expone '{clave}' igual que el extractor")

    print("\n  diferencias entre variantes")
    check(a["referencia"] == "" and a["clave_rastreo"] == "",
          "mismo banco NO trae referencia ni clave de rastreo")
    check(b["referencia"] == "0023626" and b["clave_rastreo"].startswith("0026"),
          "interbancario SI las trae")
    check(a["concepto"] == "APOYO PLACAS" and b["concepto"] == "APOYO PLACAS",
          "concepto sale de 'Motivo de pago' o de 'Concepto de pago'")

    print("\n  referencia para el SIPP (AAAAMMDD)")
    check(lc.referencia_aaaammdd(a) == "20260804", "fecha de aplicacion -> AAAAMMDD")
    check(lc.referencia_aaaammdd(b) == "20260804",
          "usa la de APLICACION, no la de creacion (03/08 en el interbancario)")
    check(lc.referencia_aaaammdd({"fecha_creacion": "09/12/2026"}) == "20261209",
          "sin fecha de aplicacion cae a la de creacion")
    check(lc.referencia_aaaammdd({}) == "",
          "sin ninguna fecha devuelve vacio (no inventa una referencia)")

    print("\n  estado de la operacion")
    check(lc.esta_aplicado(a), "'Operado' cuenta como aplicado")
    pendiente = _pdf_con_texto(
        os.path.join(tmp, "pendiente.pdf"),
        _PLANTILLA_MISMO_BANCO.format(estado="En proceso"))
    (p,), _ = lc.leer_varios([pendiente])
    check(not lc.esta_aplicado(p), "'En proceso' NO cuenta como aplicado")

    print("\n  el casado funciona igual con la lectura local")
    objetivo = c.Objetivo(origenes={"0117421184"},
                          beneficiarios={"012730028914386037"}, total=3227.00)
    check(c.evaluar_coincidencia(a, objetivo)["coincide"],
          "una lectura local casa con evaluar_coincidencia")
    check(not c.evaluar_coincidencia(b, objetivo)["coincide"],
          "y el comprobante ajeno sigue sin casar")

    print("\n  bordes")
    roto = os.path.join(tmp, "roto.pdf")
    with open(roto, "wb") as fh:
        fh.write(b"esto no es un pdf")
    lec, err = lc.leer_varios([roto])
    check(not lec and len(err) == 1, "un PDF ilegible se reporta y no rompe el lote")
    vacio = _pdf_con_texto(os.path.join(tmp, "vacio.pdf"), "Hoja sin datos de pago")
    lec2, err2 = lc.leer_varios([vacio])
    check(not lec2 and len(err2) == 1, "un PDF sin datos de pago se reporta aparte")
    lec3, err3 = lc.leer_varios([mb, roto, ib])
    check(len(lec3) == 2 and len(err3) == 1,
          "el lote continua: 2 leidos y 1 con error")

    import shutil
    shutil.rmtree(tmp, ignore_errors=True)


def prueba_vinculacion() -> None:
    print("\nvinculacion archivo <-> movimiento")
    rutas = [r"C:\x\DEV GC MOTORS $17,468.00-1.pdf",
             r"C:\x\DEV GC MOTORS $17,468.00-2.pdf"]
    # Dos movimientos del MISMO lote: misma cuenta origen y MISMO importe. Solo la
    # cuenta destino los distingue, que es el caso real del lote GC MOTORS.
    objetivos = [
        ("mov-A", c.Objetivo(origenes={"0117421184"},
                             beneficiarios={"012730028914386037"}, total=3227.00)),
        ("mov-B", c.Objetivo(origenes={"0117421184"},
                             beneficiarios={"137730104690721058"}, total=3227.00)),
    ]
    res = c.vincular(objetivos, [MISMO_BANCO, INTERBANCARIO], rutas)
    check(res.asignados == {"mov-A": rutas[0], "mov-B": rutas[1]},
          "cada comprobante va a SU movimiento pese al importe identico")
    check(not res.sin_asignar and not res.sin_movimiento and not res.sin_archivo,
          "no queda nada suelto")

    print("\n  un archivo por movimiento, y viceversa")
    # Dos movimientos identicos compitiendo por UN solo archivo.
    gemelos = [("mov-A", objetivos[0][1]), ("mov-A2", objetivos[0][1])]
    res2 = c.vincular(gemelos, [MISMO_BANCO], [rutas[0]])
    check(len(res2.asignados) == 1 and "mov-A" in res2.asignados,
          "un archivo se asigna a un solo movimiento (el primero libre)")

    # El mismo archivo no se reparte dos veces aunque haya dos movimientos que casen.
    res3 = c.vincular(gemelos, [MISMO_BANCO, MISMO_BANCO], [rutas[0], rutas[0]])
    check(len(set(res3.asignados.values())) == len(res3.asignados),
          "ningun archivo se asigna a dos movimientos")

    print("\n  lo que NO casa queda a la vista")
    ajeno = {**MISMO_BANCO, "cuenta_destino": "999999999999999999"}
    res4 = c.vincular(objetivos, [ajeno], [rutas[0]])
    check(not res4.asignados, "un comprobante ajeno no se asigna a nadie")
    check(res4.sin_movimiento == [rutas[0]],
          "se reporta como 'sin movimiento que coincida'")
    check(res4.sin_asignar == [rutas[0]], "y queda libre para asignarlo a mano")

    sin_lectura = c.vincular(objetivos, [], rutas)
    check(not sin_lectura.asignados and sin_lectura.sin_asignar == rutas,
          "sin lecturas, todos los archivos quedan libres")
    check(not sin_lectura.sin_movimiento,
          "pero NO se reportan como 'sin movimiento' (no se pudieron leer)")

    print("\n  respeta lo ya vinculado")
    res5 = c.vincular(objetivos, [MISMO_BANCO, INTERBANCARIO], rutas,
                      ocupados={"mov-A"})
    check("mov-A" not in res5.asignados, "un movimiento ya ocupado no se reasigna")
    check(res5.asignados == {"mov-B": rutas[1]}, "los demas si se asignan")
    check(res5.sin_asignar == [rutas[0]],
          "el archivo que sobra queda libre, no se fuerza a otro movimiento")

    lectura_huerfana = {**MISMO_BANCO, "documento_lectura": "no enviado.pdf"}
    res6 = c.vincular(objetivos, [lectura_huerfana], rutas)
    check(res6.sin_archivo == 1, "una lectura sin archivo se cuenta aparte")
    check(not res6.asignados, "y no produce vinculo")


def prueba_banco_clabe() -> None:
    """El SIPP guarda el banco como texto aparte de la CLABE, asi que pueden
    contradecirse. El caso real: CLABE de BanCoppel con "BBVA BANCOMER"."""
    print("\nvalidacion CLABE vs banco reportado")
    check(cb.coincide_banco("137730104690721058", "BBVA BANCOMER") is False,
          "detecta el caso real (CLABE 137 BanCoppel vs BBVA)")
    check(cb.coincide_banco("137730104690721058", "BANCOPPEL") is True,
          "acepta el banco correcto")
    check(cb.coincide_banco("072744001830318174", "SANTANDER") is False,
          "detecta otra contradiccion (072 Banorte vs Santander)")

    print("\n  nombres que NO coinciden literalmente pero son el mismo banco")
    check(cb.coincide_banco("012730028914386037", "BBVA BANCOMER") is True,
          "BBVA Mexico (catalogo) == BBVA BANCOMER (SIPP)")
    check(cb.coincide_banco("002028700475457474", "CITIBANAMEX") is True,
          "Banamex == CITIBANAMEX")
    check(cb.coincide_banco("002028700475457474", "banamex") is True,
          "no distingue mayusculas")
    check(cb.coincide_banco("058000000000000000", "BANREGIO S.A. DE C.V.") is True,
          "ignora la razon social (S.A. de C.V.)")

    print("\n  lo que NO se puede juzgar devuelve None, no False")
    check(cb.coincide_banco("137730104690721058", "") is None,
          "sin banco reportado")
    check(cb.coincide_banco("", "BANORTE") is None, "sin CLABE")
    check(cb.coincide_banco("12", "BANORTE") is None, "CLABE demasiado corta")
    check(cb.coincide_banco("999730104690721058", "BANORTE") is None,
          "prefijo que no esta en el catalogo")
    check(cb.coincide_banco("137730104690721058", "S.A. de C.V.") is None,
          "nombre sin ninguna palabra significativa")

    print("\n  el catalogo sigue resolviendo bien")
    check(cb.banco_desde_clabe("137730104690721058") == "BanCoppel", "137 BanCoppel")
    check(cb.codigo_desde_clabe("137730104690721058") == "137", "codigo de 3 digitos")
    check(cb.codigo_desde_clabe("13") == "", "codigo vacio si no alcanza")


def main() -> int:
    for prueba in (prueba_ultimos_digitos, prueba_nombres, prueba_resolucion_rutas,
                   prueba_reparto, prueba_coincidencia, prueba_lector,
                   prueba_vinculacion, prueba_banco_clabe):
        prueba()
    print()
    if _fallos:
        print(f"PRUEBAS DE COMPROBANTES: {len(_fallos)} FALLA(S)")
        for f in _fallos:
            print(f"  - {f}")
        return 1
    print("PRUEBAS DE COMPROBANTES OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
