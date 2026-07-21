"""Color de la barra de título nativa de Windows (DWM).

La app es Flet: la ventana visible la crea un cliente Flutter (flet.exe) APARTE,
con su barra de título NATIVA de Windows (título + minimizar / maximizar / cerrar).
Esa barra la pinta el sistema, no Flutter, así que para que combine con el tema
claro/oscuro de la app hay que pedírselo al gestor de ventanas (DWM).

Windows 11 (build 22000+) permite pintar la parte no-cliente con
DwmSetWindowAttribute:

  - DWMWA_USE_IMMERSIVE_DARK_MODE (20): modo oscuro de la barra (afecta el color
    por defecto de los botones/borde cuando no se fija uno explícito).
  - DWMWA_CAPTION_COLOR (35): color de FONDO de la barra de título.
  - DWMWA_TEXT_COLOR    (36): color del TEXTO del título.
  - DWMWA_BORDER_COLOR  (34): color del BORDE de la ventana.

El color es un COLORREF (0x00BBGGRR). Igual que en `win_taskbar`, la ventana se
crea de forma asíncrona, así que el HWND se sondea por título en un hilo demonio.

Todo el trabajo nativo (ctypes) queda encapsulado aquí y es un NO-OP fuera de
Windows (o en Windows 10, donde los atributos no existen: DWM devuelve error y se
ignora), para no romper imports en otras plataformas / en el smoke test.
"""

from __future__ import annotations

import sys

if sys.platform == "win32":
    import ctypes
    import threading
    import time
    from ctypes import wintypes

    _dwmapi = ctypes.windll.dwmapi
    _user32 = ctypes.windll.user32

    # Atributos de DwmSetWindowAttribute (dwmapi.h). Requieren Windows 11 22000+
    # (salvo IMMERSIVE_DARK_MODE, desde Win10 2004): en versiones previas la
    # llamada devuelve un HRESULT de error y simplemente no hace nada.
    _DWMWA_USE_IMMERSIVE_DARK_MODE = 20
    _DWMWA_BORDER_COLOR = 34
    _DWMWA_CAPTION_COLOR = 35
    _DWMWA_TEXT_COLOR = 36

    _dwmapi.DwmSetWindowAttribute.argtypes = [
        wintypes.HWND, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD]
    _dwmapi.DwmSetWindowAttribute.restype = ctypes.c_long  # HRESULT
    _user32.EnumWindows.argtypes = [ctypes.c_void_p, wintypes.LPARAM]
    _user32.IsWindowVisible.argtypes = [wintypes.HWND]
    _user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    _user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]

    def _colorref(hex_color: str) -> int:
        """'#RRGGBB' (o 'RRGGBB') -> COLORREF 0x00BBGGRR que espera DWM."""
        s = hex_color.lstrip("#")
        r, g, b = int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)
        return r | (g << 8) | (b << 16)

    def _set_attr_dword(hwnd, atributo: int, valor: int) -> None:
        dato = wintypes.DWORD(valor)
        _dwmapi.DwmSetWindowAttribute(
            hwnd, atributo, ctypes.byref(dato), ctypes.sizeof(dato))

    def _buscar_hwnd(titulo: str):
        """HWND de la ventana visible cuyo título es exactamente `titulo` (None si
        no la hay). Enumera las ventanas de nivel superior."""
        encontrado = []

        @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        def cb(hwnd, _lparam):
            if not _user32.IsWindowVisible(hwnd):
                return True
            n = _user32.GetWindowTextLengthW(hwnd)
            if n <= 0:
                return True
            buf = ctypes.create_unicode_buffer(n + 1)
            _user32.GetWindowTextW(hwnd, buf, n + 1)
            if buf.value == titulo:
                encontrado.append(hwnd)
                return False  # deja de enumerar
            return True

        _user32.EnumWindows(cb, 0)
        return encontrado[0] if encontrado else None

    def _aplicar(hwnd, fondo: str, texto: str | None, borde: str | None,
                 oscuro: bool) -> None:
        # El flag de modo oscuro primero (ajusta lo que no se pinte explícito) y
        # luego los colores concretos de la barra.
        _set_attr_dword(hwnd, _DWMWA_USE_IMMERSIVE_DARK_MODE, 1 if oscuro else 0)
        _set_attr_dword(hwnd, _DWMWA_CAPTION_COLOR, _colorref(fondo))
        if texto:
            _set_attr_dword(hwnd, _DWMWA_TEXT_COLOR, _colorref(texto))
        if borde:
            _set_attr_dword(hwnd, _DWMWA_BORDER_COLOR, _colorref(borde))

    def pintar_barra(
        titulo: str,
        fondo: str,
        *,
        texto: str | None = None,
        borde: str | None = None,
        oscuro: bool = False,
        timeout: float = 20.0,
    ) -> None:
        """Pinta la barra de título de la ventana `titulo` con el color `fondo`
        (y opcionalmente `texto`/`borde`), en un hilo demonio que sondea el HWND
        hasta `timeout` s (la ventana la crea flet.exe de forma asíncrona; al
        cambiar de tema ya existe y se aplica en el primer intento).

        Los colores son '#RRGGBB'. Nunca lanza (best-effort)."""

        def worker() -> None:
            fin = time.monotonic() + timeout
            while time.monotonic() < fin:
                hwnd = _buscar_hwnd(titulo)
                if hwnd:
                    try:
                        _aplicar(hwnd, fondo, texto, borde, oscuro)
                    except Exception:  # noqa: BLE001 — el color no es crítico
                        pass
                    return
                time.sleep(0.3)

        threading.Thread(target=worker, daemon=True).start()

else:  # fuera de Windows: no-op (la app solo se distribuye para Windows)

    def pintar_barra(*_args, **_kwargs) -> None:
        return
