"""Verifica que el build salga apuntando al SIPP PRODUCTIVO.

La release 0.6.15 se publicó apuntando a `test.sipp.petroil.dev`: un PR dejó
intercambiadas las dos líneas de `BASE_URL` en `core/rpa_sipp.py` y el AutoUpdater
repartió esa versión a TODOS los usuarios, que operaron contra el ambiente de
pruebas hasta que alguien lo notó a mano.

Ya no se cambia de ambiente editando el código (para eso está la variable
`QUETZALTIC_SIPP_BASE_URL`), así que este chequeo vigila las dos cosas que pueden
mandar un build a otro ambiente:

  1. Que el DEFAULT DE FÁBRICA (`SesionSipp.BASE_URL_PRODUCTIVO`) sea el productivo.
     Es el que reciben las instalaciones, donde nadie define la variable.
  2. Que al compilar NO haya un override activo. Si la máquina de CI tuviera la
     variable definida, todo el build (incluido cualquier chequeo posterior) estaría
     mirando otro portal, y este script daría un falso OK.

Corre ANTES de compilar: si algo no apunta a producción, el job falla y el
instalador ni siquiera se genera.

Sale con código 1 (falla el job de CI) si algo no apunta a producción; 0 si todo
está en orden.
"""

from __future__ import annotations

import os
import sys

# Igual que en smoke_import.py: al correr como `python scripts/verificar_produccion.py`
# sys.path[0] es 'scripts/', no la raíz, así que 'core' no se resolvería.
_RAIZ = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _RAIZ not in sys.path:
    sys.path.insert(0, _RAIZ)

# URL del SIPP productivo. Es la única con la que se permite publicar un Release.
PRODUCTIVO = "https://sipp.petroil.com.mx"


def main() -> int:
    from core import entorno
    from core.rpa_sipp import SesionSipp

    errores: list[str] = []

    # 1) El default de fábrica: lo que recibe cualquier instalación.
    de_fabrica = getattr(SesionSipp, "BASE_URL_PRODUCTIVO", None)
    if de_fabrica != PRODUCTIVO:
        errores.append(
            f"BASE_URL_PRODUCTIVO = {de_fabrica!r}, se esperaba {PRODUCTIVO!r}")

    # 2) Ningún override activo al compilar (dejaría el build mirando otro portal).
    override = entorno.sipp_base_url()
    if override:
        errores.append(
            f"{entorno.VAR_SIPP_BASE_URL} está definida ({override!r}) en el entorno "
            "de compilación: quítala, es solo para pruebas locales")

    # 3) Y, en consecuencia, la URL efectiva y las URL_* derivadas.
    if SesionSipp.BASE_URL != PRODUCTIVO:
        errores.append(
            f"BASE_URL = {SesionSipp.BASE_URL!r}, se esperaba {PRODUCTIVO!r}")
    for nombre in sorted(a for a in vars(SesionSipp) if a.startswith("URL_")):
        valor = getattr(SesionSipp, nombre)
        if isinstance(valor, str) and not valor.startswith(PRODUCTIVO):
            errores.append(f"{nombre} = {valor!r}, no apunta a {PRODUCTIVO!r}")

    if errores:
        print("VERIFICACION DE PRODUCCION FALLO - el build NO apunta al SIPP productivo:")
        for err in errores:
            print(f"  - {err}")
        print()
        print("En core/rpa_sipp.py, BASE_URL_PRODUCTIVO debe ser")
        print(f'    "{PRODUCTIVO}"')
        print(f"y {entorno.VAR_SIPP_BASE_URL} NO debe estar definida al compilar")
        print("(esa variable es solo para probar en local contra otro ambiente).")
        return 1

    print(f"VERIFICACION DE PRODUCCION OK - BASE_URL = {PRODUCTIVO} (sin overrides)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
