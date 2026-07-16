"""Constantes y utilidades compartidas por las pantallas de la interfaz.

Centralizar esto evita duplicación y permite que cada pantalla viva en su
propio archivo (para trabajar en colaboración sin pisarse).
"""

from __future__ import annotations

import re

import flet as ft

from core.extractores import validar_clabe

# --- Validaciones / formatos ---------------------------------------------
RE_EMAIL = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
EXTENSIONES = ["pdf", "png", "jpg", "jpeg", "tif", "tiff", "bmp"]

# --- Colores -------------------------------------------------------------
VERDE = ft.Colors.GREEN_700
ROJO = ft.Colors.RED_700
NARANJA = ft.Colors.ORANGE_800
GRIS = ft.Colors.ON_SURFACE_VARIANT
# Rojo para el FOREGROUND de botones/íconos de acción destructiva. ROJO (RED_700)
# es un rojo oscuro fijo: sobre el fondo oscuro del tema nocturno tiene poco
# contraste (texto casi ilegible). ERROR es un ROL de tema que Material adapta a
# un rojo legible en claro y en oscuro, así que se usa para esos acentos.
ROJO_BOTON = ft.Colors.ERROR

# --- Empresas ------------------------------------------------------------
# Catálogo COMPLETO de empresas del SIPP (id + nombre), fuente única para los
# combos de las pantallas (dispersión No Pemex, devoluciones, etc.). El 'id' es
# el identificador de la empresa en el sistema: sirve para emparejar con las
# cuentas bancarias y para consultar la API (p. ej. solicitudes de devolución).
# Réplica de core/empresas.md (respuesta del SIPP). Respeta la capitalización.
EMPRESAS = [
    {"id": 1, "Empresa": "Abastecedora"},
    {"id": 2, "Empresa": "ACP Combustibles"},
    {"id": 55, "Empresa": "ADMINISTRADORA DE PRESTACION SOCIAL GP"},
    {"id": 24, "Empresa": "AENE PRODUCE SA DE CV"},
    {"id": 37, "Empresa": "AENEKA"},
    {"id": 27, "Empresa": "AEROSERVICIOS AG"},
    {"id": 19, "Empresa": "AMADO SABAS GUZMAN REYNAUD"},
    {"id": 20, "Empresa": "AMBIENTAL TEK RESOURCES"},
    {"id": 6, "Empresa": "Arza"},
    {"id": 11, "Empresa": "Asamaz"},
    {"id": 28, "Empresa": "ASFALTOS"},
    {"id": 12, "Empresa": "Aske"},
    {"id": 4, "Empresa": "Atunes"},
    {"id": 3, "Empresa": "ATUNES Y SARDINAS DE MEXICO"},
    {"id": 40, "Empresa": "BLUE PROPANE"},
    {"id": 17, "Empresa": "CARROZAS DE EPOCA SA DE CV"},
    {"id": 59, "Empresa": "DERIVADOS Y SERVICIOS DE ENERGIA"},
    {"id": 31, "Empresa": "DURAMAS RENUEVALLANTAS"},
    {"id": 52, "Empresa": "ELEKTRON MOTORS AMERICA"},
    {"id": 44, "Empresa": "ELYON LOGISTICS"},
    {"id": 50, "Empresa": "ESTACION DE GAS SANTA MONICA"},
    {"id": 36, "Empresa": "FODEN"},
    {"id": 58, "Empresa": "FUNDACION SEREN"},
    {"id": 39, "Empresa": "Gas Natural Petroil"},
    {"id": 30, "Empresa": "GC MOTORS DE OCCIDENTE"},
    {"id": 38, "Empresa": "IMAA"},
    {"id": 41, "Empresa": "JORGE ALBERTO ELIAS RETES"},
    {"id": 60, "Empresa": "JORGE CASAL GONZALEZ"},
    {"id": 23, "Empresa": "LLANTERA GUZMAN DE GUAMUCHIL"},
    {"id": 16, "Empresa": "Maquinaria Equipos y Construcciones EMMG SA de CV"},
    {"id": 42, "Empresa": "MARCO ANTONIO SANCHEZ ACOSTA"},
    {"id": 53, "Empresa": "Mazaport"},
    {"id": 51, "Empresa": "MAZPARK LOGISTICO"},
    {"id": 43, "Empresa": "MERARID"},
    {"id": 5, "Empresa": "Mexcapital"},
    {"id": 34, "Empresa": "Observatorio Express"},
    {"id": 32, "Empresa": "OCEANICA"},
    {"id": 56, "Empresa": "OPERACIONES TEMATICAS MZT"},
    {"id": 7, "Empresa": "Operadora"},
    {"id": 18, "Empresa": "OPERADORA TURISTICA OBSERVATORIO 1873"},
    {"id": 8, "Empresa": "Petro Smart"},
    {"id": 25, "Empresa": "PETRO SMART COMBUSTIBLES DEL PACIFICO"},
    {"id": 13, "Empresa": "PETROIL BRAND"},
    {"id": 57, "Empresa": "PETROIL ENERGY HOLDING"},
    {"id": 14, "Empresa": "PETROIL HOLDING"},
    {"id": 15, "Empresa": "PETROIL MARINE"},
    {"id": 10, "Empresa": "Petroplazas"},
    {"id": 26, "Empresa": "PETROPLAZAS AEROPUERTO"},
    {"id": 9, "Empresa": "PETROPLAZAS ESTACIONES"},
    {"id": 54, "Empresa": "Quetzaltic"},
    {"id": 29, "Empresa": "SERVICIO EL DURANGUENO"},
    {"id": 47, "Empresa": "SERVICIO J Y J"},
    {"id": 46, "Empresa": "SERVICIOS EDUCATIVOS IMAA"},
    {"id": 45, "Empresa": "SERVICIOS COMPLEMENTARIOS EDUCATIVOS"},
    {"id": 22, "Empresa": "SUPER LLANTAS DEL PACIFICO SA DE CV"},
    {"id": 49, "Empresa": "TAO MOTORS SA DE CV"},
    {"id": 48, "Empresa": "TRASLADOS ROEH"},
    {"id": 21, "Empresa": "TURISMO Y DESARROLLO SHEKINAH"},
]
# Nombres de empresa (para los combos) e índice nombre -> id (para el match con
# las cuentas bancarias). 'EMPRESAS' es la fuente única; estos se derivan de él.
NOMBRES_EMPRESAS = [e["Empresa"] for e in EMPRESAS]
ID_POR_EMPRESA = {e["Empresa"]: e["id"] for e in EMPRESAS}


# --- Anchos de columna (compartidos entre encabezado y celdas) -----------
W_ESTADO = 64
W_CLABE = 200
W_MONTO = 120
W_BANCO = 140
W_TIPO = 150
W_ACCIONES = 150
W_NOMBRE = 200
CENTRO = ft.Alignment(0, 0)


# --- Helpers de UI -------------------------------------------------------
def celda_centrada(contenido: ft.Control, ancho: int) -> ft.Container:
    return ft.Container(content=contenido, width=ancho, alignment=CENTRO)


def encabezado_col(titulo: str, ancho: int) -> ft.Container:
    return ft.Container(
        content=ft.Text(titulo, weight=ft.FontWeight.BOLD, size=13,
                        text_align=ft.TextAlign.CENTER),
        width=ancho, alignment=CENTRO,
    )


def tarjeta(titulo: str, cuerpo: ft.Control) -> ft.Card:
    return ft.Card(
        content=ft.Container(
            content=ft.Column(
                [ft.Text(titulo, weight=ft.FontWeight.BOLD, size=15), cuerpo],
                spacing=10,
            ),
            padding=16,
        )
    )


# --- Helpers de datos ----------------------------------------------------
def solo_digitos(texto: str | None) -> str:
    return re.sub(r"\D", "", texto or "")


def validar(clabe: str, beneficiario: str, alias: str, email: str) -> str:
    if not validar_clabe(clabe):
        return "La CLABE debe tener 18 dígitos y un dígito de control válido."
    if not beneficiario:
        return "Falta el nombre del beneficiario."
    if not alias:
        return "Falta el alias de la cuenta."
    if email and not RE_EMAIL.match(email):
        return "El email de notificación no tiene un formato válido."
    return ""


def parse_monto(texto: str | None) -> float | None:
    """Convierte el texto del monto a número. Vacío -> None. Lanza ValueError
    si no es un número válido o es negativo."""
    s = (texto or "").strip().replace("$", "").replace(",", "").replace(" ", "")
    if not s:
        return None
    valor = float(s)
    if valor < 0:
        raise ValueError("El monto no puede ser negativo.")
    return valor


def fmt_monto(monto: float | None) -> str:
    return "" if monto is None else f"{monto:,.2f}"
