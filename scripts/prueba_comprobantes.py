"""Pruebas de `core.comprobantes` (casado de comprobantes contra movimientos).

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


def main() -> int:
    for prueba in (prueba_ultimos_digitos, prueba_nombres, prueba_resolucion_rutas,
                   prueba_reparto, prueba_coincidencia):
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
