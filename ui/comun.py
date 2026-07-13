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

# --- Empresas (No Pemex) -------------------------------------------------
# Empresas disponibles (dispersión y demás pantallas). Array de objetos
# {id, Empresa}: el 'id' es el identificador de la empresa en la base y sirve
# para emparejar con las cuentas bancarias por id. Respeta la capitalización.
EMPRESAS = [
    {"id": 1, "Empresa": "Abastecedora"},
    {"id": 2, "Empresa": "ACP Combustibles"},
    {"id": 20, "Empresa": "AMBIENTAL TEK RESOURCES"},
    {"id": 28, "Empresa": "ASFALTOS"},
    {"id": 40, "Empresa": "BLUE PROPANE"},
    {"id": 52, "Empresa": "ELEKTRON MOTORS AMERICA"},
    {"id": 44, "Empresa": "ELYON LOGISTICS"},
    {"id": 39, "Empresa": "Gas Natural Petroil"},
    {"id": 8, "Empresa": "Petro Smart"},
    {"id": 25, "Empresa": "PETRO SMART COMBUSTIBLES DEL PACIFICO"},
    {"id": 15, "Empresa": "PETROIL MARINE"},
    {"id": 46, "Empresa": "SERVICIOS  EDUCATIVOS IMAA"},
    {"id": 45, "Empresa": "SERVICIOS COMPLEMENTARIOS EDUCATIVOS"},
    {"id": 48, "Empresa": "TRASLADOS ROEH"},
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
