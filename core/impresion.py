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


def _tope_x_layout(coords: dict) -> float:
    """Máximo borde-derecho lógico (x, en mm) entre los campos del layout,
    considerando el ancho y la alineación de cada recuadro. Al girar 90° CCW ese
    borde es el que queda MÁS ARRIBA en la página, así que sirve para anclar el
    'margen_superior' (el elemento más alto se coloca en esa Y)."""
    tope = 0.0
    for p in coords.values():
        if not isinstance(p, dict) or "x" not in p:
            continue  # llaves escalares como 'margen_superior' se ignoran
        x = float(p["x"])
        ancho = float(p.get("ancho_max") or 0)
        alinear = p.get("alineacion", "izq")
        if alinear == "der":
            borde = x
        elif alinear == "centro":
            borde = x + ancho / 2
        else:
            borde = x + ancho
        tope = max(tope, borde)
    return tope


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

    # Margen superior (mm) por defecto para la impresión VERTICAL si un layout no
    # define 'margen_superior': el elemento más alto del cheque queda con su borde
    # superior a esta distancia del borde de la página.
    _MARGEN_SUP_DEFECTO = 10

    # Modo CALIBRACIÓN: True imprime la cuadrícula milimétrica + los datos en ROJO
    # con sus recuadros (para calibrar layouts nuevos). False = impresión NORMAL
    # (solo el texto del cheque, en negro, sin rejilla ni recuadros). Cámbialo a
    # True temporalmente cuando agregues/afines un layout.
    _MODO_CALIBRACION = False

    # Tamaño de fuente (pt) de los datos del cheque en la impresión. El monto en
    # letra usa una fuente más COMPACTA (condensada y algo menor) porque es larga;
    # si aun así no cabe, se parte en 2 renglones reduciendo hasta _PTS_LETRAS_MIN.
    # Fuente condensada para todos los datos del cheque (si no existe, GDI
    # sustituye por la más parecida sin fallar).
    _FAMILIA = "Arial Narrow"
    # Un solo tamaño para todos los campos; si un campo largo no cabe se reduce
    # hasta _PTS_MIN al partirlo en 2 renglones.
    _PTS_LETRAS = 9
    _PTS_LETRAS_MIN = 6
    # Interlínea de un campo a 2 renglones (factor del alto del texto): apretada
    # para que queden lo más pegados posible sin solaparse.
    _FACTOR_INTERLINEA = 0.9

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
            # Rotación 90° CCW (impresión vertical): cuando `girar` es True, cada
            # punto lógico (x, y) se mapea a la página como (y, largo_mm - x) y el
            # texto se dibuja con la fuente rotada. Se conmuta por dibujo.
            self.girar = False
            self.largo_mm = 0.0

        # --- conversión mm -> pixeles de dispositivo (restando el offset) ---
        def _px_x(self, mm: float) -> int:
            return int(round(mm / 25.4 * self.dpix)) - self.offx

        def _px_y(self, mm: float) -> int:
            return int(round(mm / 25.4 * self.dpiy)) - self.offy

        def _map(self, x: float, y: float) -> tuple[float, float]:
            """Punto lógico (marco horizontal del layout) -> punto de página (mm).
            Si `girar`, rota 90° CCW: (x, y) -> (y, largo_mm - x)."""
            return (y, self.largo_mm - x) if self.girar else (x, y)

        def _px_pt(self, x: float, y: float) -> tuple[int, int]:
            """Punto lógico (mm) -> pixeles de dispositivo, aplicando la rotación."""
            mx, my = self._map(x, y)
            return (int(round(mx / 25.4 * self.dpix)) - self.offx,
                    int(round(my / 25.4 * self.dpiy)) - self.offy)

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

        def crear_fuente(self, puntos: float, negrita: bool = False,
                         familia: str = "Arial"):
            alto = -int(round(puntos / 72 * self.dpiy))  # negativo = alto de car.
            peso = 700 if negrita else 400
            # Vertical: escapement/orientation = 900 (90° CCW, en décimas de grado).
            esc = 900 if self.girar else 0
            # Si la familia no existe (p. ej. 'Arial Narrow' sin Office), GDI
            # sustituye por la más parecida; no falla.
            fuente = _gdi.CreateFontW(
                alto, 0, esc, esc, peso, 0, 0, 0, _DEFAULT_CHARSET, 0, 0, 0, 0,
                familia)
            self._objetos.append(fuente)
            return fuente

        # --- primitivas de dibujo (coordenadas en mm) -----------------------
        def linea(self, pluma, x1: float, y1: float, x2: float, y2: float):
            _gdi.SelectObject(self.hdc, pluma)
            p1x, p1y = self._px_pt(x1, y1)
            p2x, p2y = self._px_pt(x2, y2)
            _gdi.MoveToEx(self.hdc, p1x, p1y, None)
            _gdi.LineTo(self.hdc, p2x, p2y)

        def texto(self, fuente, x: float, y: float, texto: str,
                  color: int = 0, alinear: str = "izq"):
            """Dibuja `texto` con la esquina/borde dado en (x, y) mm.

            alinear: 'izq' (x = borde izquierdo), 'der' (x = borde derecho),
            'centro' (x = centro horizontal). La alineación se resuelve en el marco
            lógico y luego se mapea (respeta la rotación vertical)."""
            _gdi.SelectObject(self.hdc, fuente)
            _gdi.SetTextColor(self.hdc, color)
            if alinear != "izq":
                tam = _SIZE()
                _gdi.GetTextExtentPoint32W(self.hdc, texto, len(texto),
                                           ctypes.byref(tam))
                ancho_mm = tam.cx / self.dpix * 25.4
                x -= ancho_mm if alinear == "der" else ancho_mm / 2
            px, py = self._px_pt(x, y)
            _gdi.TextOutW(self.hdc, px, py, texto, len(texto))

        def medir_mm(self, fuente, texto: str) -> float:
            """Ancho (mm) que ocupa `texto` con `fuente` en este dispositivo."""
            _gdi.SelectObject(self.hdc, fuente)
            tam = _SIZE()
            _gdi.GetTextExtentPoint32W(self.hdc, texto, len(texto),
                                       ctypes.byref(tam))
            return tam.cx / self.dpix * 25.4

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

    # ------------------------------------------------ prueba de cheque (overlay)
    def _dibujar_linea(l: "_Lienzo", fuente, pos: dict, texto: str, color: int,
                       guia, alto_txt: float) -> None:
        """Dibuja un renglón en `pos` ({x, y, alineacion, ancho_max}). Si `guia` no
        es None y hay ancho_max, traza además su recuadro fino (guía de calibración);
        en impresión normal `guia` es None y solo se dibuja el texto."""
        x, y = float(pos["x"]), float(pos["y"])
        alinear = pos.get("alineacion", "izq")
        l.texto(fuente, x, y, texto, color=color, alinear=alinear)
        ancho = pos.get("ancho_max")
        if ancho and guia is not None:
            if alinear == "der":
                xa, xb = x - ancho, x
            elif alinear == "centro":
                xa, xb = x - ancho / 2, x + ancho / 2
            else:
                xa, xb = x, x + ancho
            ya, yb = y - 0.6, y + alto_txt + 0.6
            l.linea(guia, xa, ya, xb, ya)   # borde superior
            l.linea(guia, xb, ya, xb, yb)   # borde derecho
            l.linea(guia, xb, yb, xa, yb)   # borde inferior
            l.linea(guia, xa, yb, xa, ya)   # borde izquierdo

    def _dibujar_datos_cheque(l: "_Lienzo", campos: dict, coords: dict,
                              color: int, cajas: bool = True) -> None:
        """Dibuja los datos del cheque (`campos`: campo -> texto) en las posiciones
        de `coords` (campo -> {x, y, alineacion, ancho_max} en mm). Si `cajas`,
        traza además el recuadro fino de cada campo (guía de calibración); en
        impresión normal `cajas=False` y solo sale el texto."""
        fuente = l.crear_fuente(_PTS_LETRAS, familia=_FAMILIA)
        guia = l.crear_pluma(0, color) if cajas else None
        alto_normal = _PTS_LETRAS / 72 * 25.4
        for campo in ("fecha", "monto_numero"):
            pos = coords.get(campo)
            texto = campos.get(campo)
            if pos and texto:
                _dibujar_linea(l, fuente, pos, str(texto), color, guia, alto_normal)
        # Beneficiario: fuente condensada; si no cabe, 2 renglones anclados en
        # 'beneficiario_2'.
        _dibujar_campo_multilinea(
            l, coords, campos.get("beneficiario"), "beneficiario",
            "beneficiario_2", _FAMILIA, _PTS_LETRAS, _PTS_LETRAS_MIN, color, guia)
        # Monto en letra: fuente condensada; si no cabe, 2 renglones anclados en
        # 'monto_letras_2'.
        _dibujar_campo_multilinea(
            l, coords, campos.get("monto_letras"), "monto_letras",
            "monto_letras_2", _FAMILIA, _PTS_LETRAS, _PTS_LETRAS_MIN,
            color, guia)

    def _dibujar_campo_multilinea(l: "_Lienzo", coords: dict, texto, campo1: str,
                                  campo2: str, familia: str, pts_base: float,
                                  pts_min: float, color: int, guia) -> None:
        """Dibuja un campo que puede requerir 2 renglones:
          - Si cabe en UN renglón (fuente base `pts_base`/`familia`) -> se dibuja en
            `campo1`.
          - Si no cabe -> se parte en 2 renglones anclados en `campo2` (o `campo1`
            si no está), reduciendo la fuente hasta `pts_min` y con interlínea
            apretada para que queden pegados sin solaparse."""
        pos1 = coords.get(campo1)
        if not pos1 or not texto:
            return
        texto = str(texto)
        ancho1 = pos1.get("ancho_max")
        fuente_base = l.crear_fuente(pts_base, familia=familia)
        alto_base = pts_base / 72 * 25.4
        # 1 renglón: cabe tal cual en el ancho de `campo1`.
        if not ancho1 or l.medir_mm(fuente_base, texto) <= ancho1:
            _dibujar_linea(l, fuente_base, pos1, texto, color, guia, alto_base)
            return
        # 2 renglones: la posición HORIZONTAL (x, alineación, ancho_max) sale
        # SIEMPRE del campo principal (`campo1`); `campo2` solo aporta la Y del
        # bloque (para subirlo y que quepan las 2 líneas). Así, aunque `campo2`
        # tenga x/ancho distintos, el x y el ancho del campo principal se respetan.
        y_ancla = float((coords.get(campo2) or pos1).get("y", pos1["y"]))
        fuente, lineas, pts = _ajustar_dos_lineas(
            l, texto, ancho1, familia, pts_base, pts_min)
        alto_txt = pts / 72 * 25.4
        pos_l1 = dict(pos1)
        pos_l1["y"] = y_ancla
        _dibujar_linea(l, fuente, pos_l1, lineas[0], color, guia, alto_txt)
        if len(lineas) > 1:
            pos_l2 = dict(pos1)
            pos_l2["y"] = y_ancla + alto_txt * _FACTOR_INTERLINEA
            _dibujar_linea(l, fuente, pos_l2, lineas[1], color, guia, alto_txt)

    def _envolver_dos_lineas(l: "_Lienzo", fuente, texto: str,
                             ancho_max: float) -> list[str]:
        """Parte `texto` en 1 o 2 renglones (por palabras) para que el 1º quepa en
        `ancho_max`. Llena el 1er renglón lo más posible y el resto va al 2º."""
        if not ancho_max or l.medir_mm(fuente, texto) <= ancho_max:
            return [texto]
        palabras = texto.split()
        linea1: list[str] = []
        for w in palabras:
            prueba = " ".join(linea1 + [w])
            if linea1 and l.medir_mm(fuente, prueba) > ancho_max:
                break
            linea1.append(w)
        resto = palabras[len(linea1):]
        return [" ".join(linea1)] + ([" ".join(resto)] if resto else [])

    def _ajustar_dos_lineas(l: "_Lienzo", texto: str, ancho: float, familia: str,
                            pts_base: float, pts_min: float):
        """Elige la fuente (`familia`) y el partido en 2 renglones: parte de
        `pts_base` y reduce hasta que ambos renglones quepan en `ancho` (o llega a
        `pts_min`). Devuelve (fuente, [renglones], pts)."""
        pts = pts_base
        while pts >= pts_min:
            fuente = l.crear_fuente(pts, familia=familia)
            lineas = _envolver_dos_lineas(l, fuente, texto, ancho)
            cabe = not ancho or all(l.medir_mm(fuente, ln) <= ancho for ln in lineas)
            if len(lineas) <= 2 and cabe:
                return fuente, lineas, pts
            pts -= 0.5
        # Último recurso: la fuente mínima con lo que resulte (hasta 2 renglones).
        fuente = l.crear_fuente(pts_min, familia=familia)
        return fuente, _envolver_dos_lineas(l, fuente, texto, ancho), pts_min

    def imprimir_prueba_cheque(nombre: str, campos: dict, coords: dict,
                               salida: str | None = None) -> None:
        """Imprime el cheque (texto negro, rotado 90° CCW) en la impresora `nombre`.

        En modo NORMAL (`_MODO_CALIBRACION = False`) sale solo el texto del cheque,
        en negro, sin rejilla ni recuadros. Con `_MODO_CALIBRACION = True` imprime
        además la cuadrícula milimétrica y los datos en rojo con sus recuadros, para
        calibrar layouts nuevos. `campos` mapea campo -> texto; `coords` mapea
        campo -> {x, y, alineacion, ancho_max} (mm)."""
        # Rojo sobre la rejilla al calibrar; negro puro en impresión normal.
        color = _rgb(210, 30, 30) if _MODO_CALIBRACION else _rgb(0, 0, 0)

        def dibujar(l: "_Lienzo") -> None:
            # 'margen_superior' (mm): el elemento MÁS ALTO del cheque queda con su
            # borde superior en esa Y. Girado, page_y = largo_mm - x, y el borde más
            # alto corresponde al mayor x lógico (_tope_x_layout); por eso
            # largo_mm = margen + tope_x hace que ese borde caiga en 'margen'.
            margen = coords.get("margen_superior", _MARGEN_SUP_DEFECTO)
            l.largo_mm = margen + _tope_x_layout(coords)
            # Solo al calibrar: cuadrícula (sin rotar) de referencia de página.
            if _MODO_CALIBRACION:
                l.girar = False
                _dibujar_calibracion(l)
            # Datos del cheque, ROTADOS 90° CCW (impresión vertical). Los recuadros
            # de guía solo se dibujan en modo calibración (`cajas`).
            l.girar = True
            _dibujar_datos_cheque(l, campos, coords, color,
                                  cajas=_MODO_CALIBRACION)

        doc = "Cheque (calibración)" if _MODO_CALIBRACION else "Cheque"
        imprimir(nombre, doc, dibujar, salida=salida)

else:  # fuera de Windows: stubs que no rompen imports (la app es solo Windows)

    def listar_impresoras() -> list[str]:
        return []

    def impresora_predeterminada() -> str | None:
        return None

    def imprimir(*_args, **_kwargs) -> None:
        raise RuntimeError("La impresión solo está disponible en Windows.")

    def imprimir_hoja_calibracion(*_args, **_kwargs) -> None:
        raise RuntimeError("La impresión solo está disponible en Windows.")

    def imprimir_prueba_cheque(*_args, **_kwargs) -> None:
        raise RuntimeError("La impresión solo está disponible en Windows.")
