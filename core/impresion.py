"""Impresión directa en Windows mediante GDI (ctypes), sin dependencias extra.

¿Por qué GDI y no un PDF + visor?  Para calibrar la posición de un cheque lo
único que importa es que salga a **escala 1:1 exacta** (1 mm dibujado = 1 mm en
papel).  Cualquier visor de PDF (Chrome, Edge, Acrobat) puede meter "ajustar a
página" y desfasar todo.  Dibujando directo al contexto de dispositivo (DC) de
la impresora, midiendo en milímetros físicos, no hay escalado posible: es 1:1
por construcción.  Además evita traer pywin32 (aquí solo hay pywin32_ctypes, que
no imprime), siguiendo el mismo patrón ctypes de core/win_taskbar.py.

El origen de coordenadas (0, 0) es la **esquina física superior izquierda del
papel** (no del área imprimible).  La franja no imprimible del borde (~4-5 mm en
láser) simplemente se recorta; por eso la hoja de calibración marca dónde empieza
de verdad lo imprimible.

Todo el trabajo nativo queda encapsulado aquí y es un NO-OP fuera de Windows
(la app solo se distribuye para Windows), para no romper imports en el smoke test.
"""

from __future__ import annotations

import sys


def _rgb(r: int, g: int, b: int) -> int:
    """COLORREF de GDI (0x00BBGGRR)."""
    return r | (g << 8) | (b << 16)


if sys.platform == "win32":
    import ctypes
    from ctypes import wintypes

    _gdi = ctypes.windll.gdi32
    _ws = ctypes.WinDLL("winspool.drv")

    # --- Índices de GetDeviceCaps -------------------------------------------
    _HORZRES, _VERTRES = 8, 10
    _LOGPIXELSX, _LOGPIXELSY = 88, 90
    _PHYSICALWIDTH, _PHYSICALHEIGHT = 110, 111
    _PHYSICALOFFSETX, _PHYSICALOFFSETY = 112, 113

    _PS_SOLID = 0
    _TRANSPARENT = 1
    _BLACK_PEN, _SYSTEM_FONT = 7, 13  # stock objects
    _DEFAULT_CHARSET = 1

    # Flags de EnumPrinters: impresoras locales + conexiones de red.
    _ENUM_LOCAL, _ENUM_CONNECTIONS = 0x2, 0x4

    class _DOCINFO(ctypes.Structure):
        _fields_ = [
            ("cbSize", ctypes.c_int),
            ("lpszDocName", wintypes.LPCWSTR),
            ("lpszOutput", wintypes.LPCWSTR),
            ("lpszDatatype", wintypes.LPCWSTR),
            ("fwType", wintypes.DWORD),
        ]

    class _SIZE(ctypes.Structure):
        _fields_ = [("cx", ctypes.c_long), ("cy", ctypes.c_long)]

    class _PRINTER_INFO_4W(ctypes.Structure):
        _fields_ = [
            ("pPrinterName", wintypes.LPWSTR),
            ("pServerName", wintypes.LPWSTR),
            ("Attributes", wintypes.DWORD),
        ]

    # Handles devueltos son punteros: SIN restype=c_void_p Python los trunca a
    # 32 bits en x64 y todo revienta.
    _gdi.CreateDCW.restype = ctypes.c_void_p
    _gdi.CreateDCW.argtypes = [
        wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPCWSTR, ctypes.c_void_p]
    _gdi.CreateFontW.restype = ctypes.c_void_p
    _gdi.CreateFontW.argtypes = [ctypes.c_int] * 13 + [wintypes.LPCWSTR]
    _gdi.CreatePen.restype = ctypes.c_void_p
    _gdi.CreatePen.argtypes = [ctypes.c_int, ctypes.c_int, wintypes.COLORREF]
    _gdi.SelectObject.restype = ctypes.c_void_p
    _gdi.SelectObject.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    _gdi.GetStockObject.restype = ctypes.c_void_p
    _gdi.GetStockObject.argtypes = [ctypes.c_int]
    _gdi.DeleteObject.argtypes = [ctypes.c_void_p]
    _gdi.DeleteDC.argtypes = [ctypes.c_void_p]
    _gdi.GetDeviceCaps.restype = ctypes.c_int
    _gdi.GetDeviceCaps.argtypes = [ctypes.c_void_p, ctypes.c_int]
    _gdi.StartDocW.restype = ctypes.c_int
    _gdi.StartDocW.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    for _f in ("StartPage", "EndPage", "EndDoc"):
        getattr(_gdi, _f).restype = ctypes.c_int
        getattr(_gdi, _f).argtypes = [ctypes.c_void_p]
    _gdi.MoveToEx.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int,
                              ctypes.c_void_p]
    _gdi.LineTo.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int]
    _gdi.TextOutW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int,
                              wintypes.LPCWSTR, ctypes.c_int]
    _gdi.SetTextColor.argtypes = [ctypes.c_void_p, wintypes.COLORREF]
    _gdi.SetBkMode.argtypes = [ctypes.c_void_p, ctypes.c_int]
    _gdi.GetTextExtentPoint32W.argtypes = [ctypes.c_void_p, wintypes.LPCWSTR,
                                           ctypes.c_int, ctypes.c_void_p]

    _ws.EnumPrintersW.restype = wintypes.BOOL
    _ws.EnumPrintersW.argtypes = [
        wintypes.DWORD, wintypes.LPWSTR, wintypes.DWORD, ctypes.c_void_p,
        wintypes.DWORD, ctypes.POINTER(wintypes.DWORD),
        ctypes.POINTER(wintypes.DWORD)]
    _ws.GetDefaultPrinterW.restype = wintypes.BOOL
    _ws.GetDefaultPrinterW.argtypes = [wintypes.LPWSTR,
                                       ctypes.POINTER(wintypes.DWORD)]

    # ------------------------------------------------------------ impresoras
    def listar_impresoras() -> list[str]:
        """Nombres de las impresoras instaladas (locales + de red)."""
        flags = _ENUM_LOCAL | _ENUM_CONNECTIONS
        needed = wintypes.DWORD(0)
        returned = wintypes.DWORD(0)
        # 1er intento sin buffer: nos dice cuántos bytes hacen falta.
        _ws.EnumPrintersW(flags, None, 4, None, 0,
                          ctypes.byref(needed), ctypes.byref(returned))
        if needed.value == 0:
            return []
        buf = (ctypes.c_byte * needed.value)()
        ok = _ws.EnumPrintersW(flags, None, 4, buf, needed.value,
                               ctypes.byref(needed), ctypes.byref(returned))
        if not ok:
            return []
        info = ctypes.cast(buf, ctypes.POINTER(_PRINTER_INFO_4W))
        nombres = [info[i].pPrinterName for i in range(returned.value)
                   if info[i].pPrinterName]
        # Sin duplicados, conservando el orden.
        return list(dict.fromkeys(nombres))

    def impresora_predeterminada() -> str | None:
        """Nombre de la impresora predeterminada del sistema, o None."""
        size = wintypes.DWORD(0)
        _ws.GetDefaultPrinterW(None, ctypes.byref(size))
        if size.value == 0:
            return None
        buf = ctypes.create_unicode_buffer(size.value)
        if not _ws.GetDefaultPrinterW(buf, ctypes.byref(size)):
            return None
        return buf.value or None

    # ------------------------------------------------------- lienzo (mm)
    class _Lienzo:
        """Superficie de dibujo sobre el DC de la impresora, en milímetros.

        El origen (0, 0) es la esquina física superior izquierda del papel.
        Reúne los objetos GDI creados para liberarlos todos al cerrar."""

        def __init__(self, hdc):
            self.hdc = hdc
            self.dpix = _gdi.GetDeviceCaps(hdc, _LOGPIXELSX)
            self.dpiy = _gdi.GetDeviceCaps(hdc, _LOGPIXELSY)
            self.offx = _gdi.GetDeviceCaps(hdc, _PHYSICALOFFSETX)
            self.offy = _gdi.GetDeviceCaps(hdc, _PHYSICALOFFSETY)
            self.horzres = _gdi.GetDeviceCaps(hdc, _HORZRES)
            self.vertres = _gdi.GetDeviceCaps(hdc, _VERTRES)
            self.phys_w = _gdi.GetDeviceCaps(hdc, _PHYSICALWIDTH)
            self.phys_h = _gdi.GetDeviceCaps(hdc, _PHYSICALHEIGHT)
            self._objetos: list = []

        # --- conversión mm -> pixeles de dispositivo (restando el offset) ---
        def _px_x(self, mm: float) -> int:
            return int(round(mm / 25.4 * self.dpix)) - self.offx

        def _px_y(self, mm: float) -> int:
            return int(round(mm / 25.4 * self.dpiy)) - self.offy

        # --- medidas del papel / área imprimible, en mm ---------------------
        @property
        def ancho_mm(self) -> float:
            return self.phys_w / self.dpix * 25.4

        @property
        def alto_mm(self) -> float:
            return self.phys_h / self.dpiy * 25.4

        @property
        def margen_izq_mm(self) -> float:
            return self.offx / self.dpix * 25.4

        @property
        def margen_sup_mm(self) -> float:
            return self.offy / self.dpiy * 25.4

        @property
        def margen_der_mm(self) -> float:
            return (self.phys_w - self.offx - self.horzres) / self.dpix * 25.4

        @property
        def margen_inf_mm(self) -> float:
            return (self.phys_h - self.offy - self.vertres) / self.dpiy * 25.4

        # --- creación de recursos GDI (se liberan en cerrar) ----------------
        def crear_pluma(self, ancho_px: int, color: int):
            pluma = _gdi.CreatePen(_PS_SOLID, ancho_px, color)
            self._objetos.append(pluma)
            return pluma

        def crear_fuente(self, puntos: float, negrita: bool = False):
            alto = -int(round(puntos / 72 * self.dpiy))  # negativo = alto de car.
            peso = 700 if negrita else 400
            fuente = _gdi.CreateFontW(
                alto, 0, 0, 0, peso, 0, 0, 0, _DEFAULT_CHARSET, 0, 0, 0, 0,
                "Arial")
            self._objetos.append(fuente)
            return fuente

        # --- primitivas de dibujo (coordenadas en mm) -----------------------
        def linea(self, pluma, x1: float, y1: float, x2: float, y2: float):
            _gdi.SelectObject(self.hdc, pluma)
            _gdi.MoveToEx(self.hdc, self._px_x(x1), self._px_y(y1), None)
            _gdi.LineTo(self.hdc, self._px_x(x2), self._px_y(y2))

        def texto(self, fuente, x: float, y: float, texto: str,
                  color: int = 0, alinear: str = "izq"):
            """Dibuja `texto` con la esquina/borde dado en (x, y) mm.

            alinear: 'izq' (x = borde izquierdo), 'der' (x = borde derecho),
            'centro' (x = centro horizontal)."""
            _gdi.SelectObject(self.hdc, fuente)
            _gdi.SetTextColor(self.hdc, color)
            px, py = self._px_x(x), self._px_y(y)
            if alinear != "izq":
                tam = _SIZE()
                _gdi.GetTextExtentPoint32W(self.hdc, texto, len(texto),
                                           ctypes.byref(tam))
                px -= tam.cx if alinear == "der" else tam.cx // 2
            _gdi.TextOutW(self.hdc, px, py, texto, len(texto))

        def cerrar(self):
            # No se puede borrar un objeto seleccionado: primero se reponen los
            # stock objects, luego se borran los creados.
            _gdi.SelectObject(self.hdc, _gdi.GetStockObject(_BLACK_PEN))
            _gdi.SelectObject(self.hdc, _gdi.GetStockObject(_SYSTEM_FONT))
            for obj in self._objetos:
                _gdi.DeleteObject(obj)
            self._objetos.clear()

    def imprimir(nombre: str, doc_nombre: str, dibujar, salida: str | None = None):
        """Abre `nombre`, arranca un documento de una página y llama a
        `dibujar(lienzo)` para pintarla.  `salida` (ruta de archivo) es solo
        para pruebas con impresoras de archivo (p. ej. "Microsoft Print to
        PDF"); en producción va en None → sale por la impresora física."""
        hdc = _gdi.CreateDCW("WINSPOOL", nombre, salida, None)
        if not hdc:
            raise RuntimeError(f"No se pudo abrir la impresora «{nombre}».")
        try:
            info = _DOCINFO()
            info.cbSize = ctypes.sizeof(_DOCINFO)
            info.lpszDocName = doc_nombre
            info.lpszOutput = salida
            if _gdi.StartDocW(hdc, ctypes.byref(info)) <= 0:
                raise RuntimeError("La impresora rechazó el trabajo (StartDoc).")
            try:
                if _gdi.StartPage(hdc) <= 0:
                    raise RuntimeError("No se pudo iniciar la página.")
                _gdi.SetBkMode(hdc, _TRANSPARENT)
                lienzo = _Lienzo(hdc)
                try:
                    dibujar(lienzo)
                finally:
                    lienzo.cerrar()
                _gdi.EndPage(hdc)
            finally:
                _gdi.EndDoc(hdc)
        finally:
            _gdi.DeleteDC(hdc)

    # ---------------------------------------------------- hoja de calibración
    def _dibujar_calibracion(l: "_Lienzo") -> None:
        w, h = l.ancho_mm, l.alto_mm
        gris = _rgb(200, 200, 200)      # rejilla fina (5 mm)
        gris_med = _rgb(140, 140, 140)  # rejilla media (10 mm)
        negro = _rgb(0, 0, 0)           # rejilla gruesa (50 mm) y textos

        p_fino = l.crear_pluma(0, gris)
        p_medio = l.crear_pluma(0, gris_med)
        p_grueso = l.crear_pluma(max(1, int(l.dpix / 150)), negro)
        f_num = l.crear_fuente(6)
        f_grande = l.crear_fuente(22, negrita=True)
        f_chica = l.crear_fuente(8)
        f_titulo = l.crear_fuente(12, negrita=True)
        f_nota = l.crear_fuente(8)

        def pluma_para(mm: float):
            if round(mm) % 50 == 0:
                return p_grueso
            return p_medio if round(mm) % 10 == 0 else p_fino

        # Rejilla cada 5 mm (se recorta sola al área imprimible).
        x = 0.0
        while x <= w + 0.01:
            l.linea(pluma_para(x), x, 0, x, h)
            x += 5
        y = 0.0
        while y <= h + 0.01:
            l.linea(pluma_para(y), 0, y, w, y)
            y += 5

        # Números (mm) cada 10 mm sobre el borde imprimible superior e izquierdo.
        top = l.margen_sup_mm + 0.5
        left = l.margen_izq_mm + 0.5
        # (se omiten los ~15 mm de cada extremo para no chocar con las siglas
        # de esquina.)
        x = 0.0
        while x <= w + 0.01:
            if round(x) % 10 == 0 and 15 <= x <= w - 15:
                l.texto(f_num, x + 0.4, top, str(int(round(x))), color=negro)
            x += 5
        y = 0.0
        while y <= h + 0.01:
            if round(y) % 10 == 0 and 15 <= y <= h - 15:
                l.texto(f_num, left, y + 0.4, str(int(round(y))), color=negro)
            y += 5

        # Esquinas imprimibles: sigla + coordenada (mm) medida desde (0,0) físico.
        pl, pt = l.margen_izq_mm, l.margen_sup_mm
        pr, pb = w - l.margen_der_mm, h - l.margen_inf_mm
        b = 12  # largo del corchete en L

        def esquina(cx, cy, dx, dy, sigla, alinear):
            # Corchete en L apuntando hacia adentro.
            l.linea(p_grueso, cx, cy, cx + dx * b, cy)
            l.linea(p_grueso, cx, cy, cx, cy + dy * b)
            ty = cy + 2 if dy > 0 else cy - 10
            tx = cx + dx * 2
            l.texto(f_grande, tx, ty, sigla, color=negro, alinear=alinear)
            cy_txt = ty + 9 if dy > 0 else ty + 9
            l.texto(f_chica, tx, cy_txt,
                    f"({round(cx)}, {round(cy)}) mm", color=negro, alinear=alinear)

        esquina(pl, pt, +1, +1, "SI", "izq")
        esquina(pr, pt, -1, +1, "SD", "der")
        esquina(pl, pb, +1, -1, "II", "izq")
        esquina(pr, pb, -1, -1, "ID", "der")

        # Título y nota al centro-arriba.
        l.texto(f_titulo, w / 2, pt + 4, "HOJA DE CALIBRACIÓN — CHEQUES",
                color=negro, alinear="centro")
        l.texto(f_nota, w / 2, pt + 10,
                "Origen (0,0) = esquina superior izquierda del papel · "
                "regla en milímetros", color=negro, alinear="centro")

        # Barra de referencia de 100 mm para verificar la escala con una regla.
        cy = h / 2
        x0 = w / 2 - 50
        l.linea(p_grueso, x0, cy, x0 + 100, cy)
        l.linea(p_grueso, x0, cy - 2, x0, cy + 2)
        l.linea(p_grueso, x0 + 100, cy - 2, x0 + 100, cy + 2)
        l.texto(f_nota, w / 2, cy + 2,
                "Esta barra debe medir 100 mm con regla (si no, hay escalado)",
                color=negro, alinear="centro")

    def imprimir_hoja_calibracion(nombre: str, salida: str | None = None) -> None:
        """Imprime la hoja de calibración de cheques en la impresora `nombre`."""
        imprimir(nombre, "Hoja de calibración de cheques", _dibujar_calibracion,
                 salida=salida)

else:  # fuera de Windows: stubs que no rompen imports (la app es solo Windows)

    def listar_impresoras() -> list[str]:
        return []

    def impresora_predeterminada() -> str | None:
        return None

    def imprimir(*_args, **_kwargs) -> None:
        raise RuntimeError("La impresión solo está disponible en Windows.")

    def imprimir_hoja_calibracion(*_args, **_kwargs) -> None:
        raise RuntimeError("La impresión solo está disponible en Windows.")
