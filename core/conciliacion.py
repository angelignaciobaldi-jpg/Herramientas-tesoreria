"""Conciliación de movimientos para generar dispersiones (No Pemex).

Toma los movimientos SELECCIONADOS por el usuario en la tabla (agrupados por
empresa) y prepara/valida la información antes de mandarla al RPA:

  1. Valida que cada movimiento traiga los campos requeridos.
  2. Agrupa por cuenta bancaria y concilia, por cuenta, que el total del importe
     (Saldo Facturado) coincida con el total del Saldo Programado.
  3. Marca cada empresa como 'válida' (sin errores y con >= 1 movimiento) para
     que el RPA itere solo sobre esas.

No depende de la interfaz (Flet); opera sobre FilaSolicitud, por lo que es
fácilmente testeable. La salida (a_debug) es una estructura 'cruda' pensada para
depurar lo que se enviará al RPA.
"""

from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field

from core.extractores import validar_clabe
from core.reporte_dispersion import FilaSolicitud

# Campos que un movimiento seleccionado DEBE traer para poder procesarse.
# (campo interno, etiqueta visible)
CAMPOS_REQUERIDOS = [
    ("folio", "Folio"),
    ("folio_factura", "Folio Factura"),
    ("tipo_solicitud", "Tipo Solicitud"),
    ("proveedor", "Proveedor"),
    ("cuenta_bancaria", "Cuenta Bancaria"),
    ("total_factura", "Total Factura"),
    ("saldo_factura", "Saldo Facturado"),
    ("saldo_programado", "Saldo Programado"),
]

# Campo que representa el "importe" a conciliar contra el Saldo Programado (por
# cuenta bancaria). Se usa Saldo Facturado, el contraparte directo del Saldo
# Programado. Si el negocio define que el "importe" es el Total Factura, basta
# cambiar esto a "total_factura".
_CAMPO_IMPORTE = "saldo_factura"

# Tolerancia (centavos) para comparar montos.
_TOL = 0.005

# Rango de dígitos aceptable para un "No. de cuenta" (no CLABE): cubre cuentas
# de 10-11 dígitos, tarjetas de 15-16 y otros formatos. La CLABE (18 dígitos)
# se valida aparte con su dígito de control.
_MIN_DIGITOS_CUENTA = 10
_MAX_DIGITOS_CUENTA = 20


def _falta(f: FilaSolicitud, campo: str) -> bool:
    """True si el campo no trae información utilizable."""
    valor = getattr(f, campo, None)
    if isinstance(valor, str):
        return not valor.strip()
    return valor is None  # montos: None = falta


def cuenta_valida(texto: str) -> bool:
    """True si `texto` contiene una cuenta bancaria válida.

    La cuenta suele venir como ``"BANCO - CUENTA"``, pero también puede llegar
    solo el número sin el banco. Como los nombres de banco no llevan dígitos,
    basta con extraer todos los dígitos del texto y verificar que formen:

      * una CLABE de 18 dígitos con dígito de control correcto, o
      * un número de cuenta con una cantidad de dígitos plausible.
    """
    dig = re.sub(r"\D", "", str(texto or ""))
    if not dig:
        return False
    if len(dig) == 18:
        return validar_clabe(dig)
    return _MIN_DIGITOS_CUENTA <= len(dig) <= _MAX_DIGITOS_CUENTA


@dataclass
class GrupoCuenta:
    """Movimientos de una misma cuenta bancaria (dentro de una empresa)."""

    cuenta_bancaria: str
    movimientos: list[FilaSolicitud]
    total_importe: float       # suma de Saldo Facturado
    total_programado: float    # suma de Saldo Programado

    @property
    def cuadra(self) -> bool:
        return abs(self.total_importe - self.total_programado) <= _TOL


@dataclass
class EmpresaDispersion:
    """Movimientos seleccionados de una empresa, agrupados y validados."""

    empresa: str
    movimientos: list[FilaSolicitud]           # seleccionados de la empresa
    grupos: list[GrupoCuenta] = field(default_factory=list)
    errores: list[str] = field(default_factory=list)
    # Datos de pago elegidos por el usuario en la tabla (cuenta de origen +
    # concepto/referencia). Se guardan para usarlos más adelante en el RPA.
    cuenta: str = ""                # 'Cuenta': valor por el que busca el RPA
    concepto_pago: str = ""
    referencia_pago: str = ""

    @property
    def valida(self) -> bool:
        """Válida = tiene al menos un movimiento y no tiene errores."""
        return bool(self.movimientos) and not self.errores


@dataclass
class Conciliacion:
    """Resultado global: una entrada por empresa con selección."""

    empresas: list[EmpresaDispersion]

    @property
    def validas(self) -> list[EmpresaDispersion]:
        return [e for e in self.empresas if e.valida]

    @property
    def hay_errores(self) -> bool:
        return any(e.errores for e in self.empresas)


def _validar_requeridos(movimientos: list[FilaSolicitud]) -> list[str]:
    errores: list[str] = []
    for i, f in enumerate(movimientos, 1):
        ref = f.folio_factura or f.folio or f"#{i}"
        faltantes = [etq for campo, etq in CAMPOS_REQUERIDOS if _falta(f, campo)]
        if faltantes:
            errores.append(f"Movimiento {ref}: falta(n) {', '.join(faltantes)}.")
        # La cuenta puede venir presente pero sin una CLABE / No. de cuenta
        # reconocible: se reporta aparte de "falta".
        if not _falta(f, "cuenta_bancaria") and not cuenta_valida(f.cuenta_bancaria):
            errores.append(
                f"Movimiento {ref}: la cuenta bancaria '{f.cuenta_bancaria}' "
                f"no contiene una CLABE o número de cuenta válido."
            )
    return errores


def _agrupar_por_cuenta(movimientos: list[FilaSolicitud]) -> list[GrupoCuenta]:
    por_cuenta: dict[str, list[FilaSolicitud]] = {}
    orden: list[str] = []
    for f in movimientos:
        cuenta = f.cuenta_bancaria or "(Sin cuenta)"
        if cuenta not in por_cuenta:
            por_cuenta[cuenta] = []
            orden.append(cuenta)
        por_cuenta[cuenta].append(f)
    grupos: list[GrupoCuenta] = []
    for cuenta in orden:
        movs = por_cuenta[cuenta]
        total_importe = round(sum(getattr(m, _CAMPO_IMPORTE) or 0 for m in movs), 2)
        total_prog = round(sum(m.saldo_programado or 0 for m in movs), 2)
        grupos.append(GrupoCuenta(cuenta, movs, total_importe, total_prog))
    return grupos


def conciliar(
    seleccion_por_empresa: dict[str, list[FilaSolicitud]],
    datos_pago: dict[str, dict] | None = None,
) -> Conciliacion:
    """Concilia y valida la selección. `seleccion_por_empresa` mapea cada empresa
    (clave 'Empresa - Moneda') a sus movimientos SELECCIONADOS. `datos_pago`, si se
    provee, mapea la MISMA clave a los datos de pago elegidos en la tabla (cuenta,
    concepto_pago, referencia_pago), que se adjuntan a cada empresa. Devuelve un
    Conciliacion con, por empresa, los grupos por cuenta y los errores encontrados."""
    datos_pago = datos_pago or {}
    empresas: list[EmpresaDispersion] = []
    for empresa, movimientos in seleccion_por_empresa.items():
        if not movimientos:
            continue
        errores = _validar_requeridos(movimientos)
        grupos = _agrupar_por_cuenta(movimientos)
        for g in grupos:
            if not g.cuadra:
                errores.append(
                    f"Cuenta '{g.cuenta_bancaria}': el importe facturado "
                    f"(${g.total_importe:,.2f}) no coincide con el programado "
                    f"(${g.total_programado:,.2f})."
                )
        pago = datos_pago.get(empresa, {})
        empresas.append(EmpresaDispersion(
            empresa, movimientos, grupos, errores,
            cuenta=pago.get("cuenta", ""),
            concepto_pago=pago.get("concepto_pago", ""),
            referencia_pago=pago.get("referencia_pago", ""),
        ))
    return Conciliacion(empresas)


def a_debug(conc: Conciliacion) -> dict:
    """Estructura 'cruda' (JSON-serializable) de lo que se enviaría al RPA, para
    depuración. Se mostrará temporalmente en la UI."""
    return {
        "empresas_validas": [e.empresa for e in conc.validas],
        "hay_errores": conc.hay_errores,
        "empresas": [
            {
                "empresa": e.empresa,
                "valida": e.valida,
                "errores": e.errores,
                "total_movimientos": len(e.movimientos),
                "cuenta": e.cuenta,
                "concepto_pago": e.concepto_pago,
                "referencia_pago": e.referencia_pago,
                "grupos": [
                    {
                        "cuenta_bancaria": g.cuenta_bancaria,
                        "total_importe": g.total_importe,
                        "total_programado": g.total_programado,
                        "cuadra": g.cuadra,
                        "movimientos": [asdict(m) for m in g.movimientos],
                    }
                    for g in e.grupos
                ],
            }
            for e in conc.empresas
        ],
    }
