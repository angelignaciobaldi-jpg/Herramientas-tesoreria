"""Verifica que el build salga apuntando al SIPP PRODUCTIVO.

Para probar contra el ambiente de pruebas se intercambian dos líneas en
`core/rpa_sipp.py` (se comenta la de producción y se descomenta la de test).
Es un cambio de dos renglones, fácil de olvidar al abrir el PR: así salió la
release 0.6.15, que quedó apuntando a `test.sipp.petroil.dev` y mandó a TODOS
los usuarios al ambiente de pruebas sin que nadie lo notara hasta usarla.

El daño no se limita a la máquina de quien probó: al publicarse el Release, el
AutoUpdater lo reparte a todo el mundo, y revertirlo obliga a sacar otra versión.

Por eso el chequeo corre en el pipeline ANTES de compilar: si la URL no es la de
producción, el job falla y el instalador ni siquiera se genera.

Se valida el valor EFECTIVO (importando la clase), no el texto del archivo: así
da igual cuál de las dos líneas esté comentada, cuántas variantes de URL se
acumulen arriba, o si alguien deja una `URL_*` apuntando a otro lado a mano.

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
    from core.rpa_sipp import SesionSipp

    errores: list[str] = []

    if SesionSipp.BASE_URL != PRODUCTIVO:
        errores.append(
            f"BASE_URL = {SesionSipp.BASE_URL!r}, se esperaba {PRODUCTIVO!r}"
        )

    # Las URL_* se arman a partir de BASE_URL, pero se revisan una por una por si
    # alguna quedó escrita a mano apuntando a otro ambiente.
    for nombre in sorted(a for a in vars(SesionSipp) if a.startswith("URL_")):
        valor = getattr(SesionSipp, nombre)
        if isinstance(valor, str) and not valor.startswith(PRODUCTIVO):
            errores.append(f"{nombre} = {valor!r}, no apunta a {PRODUCTIVO!r}")

    if errores:
        print("VERIFICACION DE PRODUCCION FALLO - el build NO apunta al SIPP productivo:")
        for err in errores:
            print(f"  - {err}")
        print()
        print("Corrige core/rpa_sipp.py: deja activa la linea")
        print(f'    BASE_URL = "{PRODUCTIVO}"')
        print("y comenta la del ambiente de pruebas. Luego vuelve a publicar el Release.")
        return 1

    print(f"VERIFICACION DE PRODUCCION OK - BASE_URL = {PRODUCTIVO}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
