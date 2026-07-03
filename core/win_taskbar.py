"""Identidad de la app en la barra de tareas de Windows.

Una app de escritorio Flet arranca como DOS procesos: el .exe de Python
(Tesoreria.exe) lanza un cliente Flutter APARTE (flet.exe) que es el DUEÑO de la
ventana visible. Como esa ventana pertenece a un flet.exe genérico (icono default
de Flet, argumentos de sesión):

  * al "Anclar a la barra de tareas" la app corriendo, Windows ancla el flet.exe
    (icono equivocado) con argumentos de esa sesión, por lo que el pin ya no
    vuelve a abrir bien la app;
  * la agrupación y el icono de la barra no son los de la app.

La solución soportada por Windows es etiquetar la VENTANA con las propiedades del
shell:

  - System.AppUserModel.ID                     (identidad / agrupación)
  - System.AppUserModel.RelaunchCommand         (qué ejecutar al anclar/re-lanzar)
  - System.AppUserModel.RelaunchIconResource    (icono del pin)
  - System.AppUserModel.RelaunchDisplayNameResource (nombre visible del pin)

Con RelaunchCommand puesto, al anclar la ventana Windows crea el acceso con ESE
comando (Tesoreria.exe) y ESE icono, no con el flet.exe crudo. Arregla el
lanzamiento del pin y el icono a la vez.

Todo el trabajo nativo (ctypes) queda encapsulado aquí y es un NO-OP fuera de
Windows, para no romper imports en otras plataformas / en el smoke test.
"""

from __future__ import annotations

import sys

# AppUserModelID estable de la app. DEBE coincidir con el AppUserModelID de los
# accesos directos del instalador (instalador.iss -> [Icons]) para que la ventana
# y el acceso se agrupen como la misma app.
AUMID = "QuetzalticSolutions.HerramientasTesoreria"


if sys.platform == "win32":
    import ctypes
    import threading
    import time
    from ctypes import wintypes

    _shell32 = ctypes.windll.shell32
    _ole32 = ctypes.windll.ole32
    _user32 = ctypes.windll.user32

    _VT_LPWSTR = 31  # PROPVARIANT de cadena Unicode (pwszVal)

    class _GUID(ctypes.Structure):
        _fields_ = [
            ("Data1", wintypes.DWORD),
            ("Data2", wintypes.WORD),
            ("Data3", wintypes.WORD),
            ("Data4", ctypes.c_ubyte * 8),
        ]

    class _PROPERTYKEY(ctypes.Structure):
        _fields_ = [("fmtid", _GUID), ("pid", wintypes.DWORD)]

    class _PROPVARIANT(ctypes.Structure):
        # Tamaño real de PROPVARIANT en x64 (24 bytes): cabecera + unión. Solo se
        # usa el puntero (VT_LPWSTR) pero se reserva el tamaño completo para que
        # SetValue no lea fuera del struct.
        _fields_ = [
            ("vt", wintypes.USHORT),
            ("wReserved1", wintypes.USHORT),
            ("wReserved2", wintypes.USHORT),
            ("wReserved3", wintypes.USHORT),
            ("p", ctypes.c_void_p),
            ("p2", ctypes.c_void_p),
        ]

    def _guid(texto: str) -> _GUID:
        g = _GUID()
        _ole32.CLSIDFromString(ctypes.c_wchar_p(texto), ctypes.byref(g))
        return g

    # fmtid común de las propiedades AppUserModel (propkey.h) y sus pid.
    _FMTID_AUM = "{9F4C2855-9F79-4B39-A8D0-E1D42DE1D5F3}"
    _IID_IPropertyStore = _guid("{886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99}")

    def _pkey(pid: int) -> _PROPERTYKEY:
        return _PROPERTYKEY(_guid(_FMTID_AUM), pid)

    _PKEY_ID = _pkey(5)               # AppUserModel.ID
    _PKEY_RELAUNCH_CMD = _pkey(2)     # AppUserModel.RelaunchCommand
    _PKEY_RELAUNCH_ICON = _pkey(3)    # AppUserModel.RelaunchIconResource
    _PKEY_RELAUNCH_NAME = _pkey(4)    # AppUserModel.RelaunchDisplayNameResource

    _shell32.SetCurrentProcessExplicitAppUserModelID.argtypes = [wintypes.LPCWSTR]
    _shell32.SHGetPropertyStoreForWindow.argtypes = [
        wintypes.HWND, ctypes.POINTER(_GUID), ctypes.POINTER(ctypes.c_void_p)]
    _shell32.SHGetPropertyStoreForWindow.restype = ctypes.c_long  # HRESULT
    _ole32.CoTaskMemAlloc.argtypes = [ctypes.c_size_t]
    _ole32.CoTaskMemAlloc.restype = ctypes.c_void_p
    _ole32.PropVariantClear.argtypes = [ctypes.POINTER(_PROPVARIANT)]
    _user32.EnumWindows.argtypes = [ctypes.c_void_p, wintypes.LPARAM]
    _user32.IsWindowVisible.argtypes = [wintypes.HWND]
    _user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    _user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]

    def _metodo(interfaz, indice, restype, argtypes):
        """Devuelve una función invocable del método COM Nº `indice` de la vtable
        de `interfaz` (un puntero c_void_p a la interfaz)."""
        vtabla = ctypes.cast(interfaz, ctypes.POINTER(ctypes.c_void_p))[0]
        func = ctypes.cast(vtabla, ctypes.POINTER(ctypes.c_void_p))[indice]
        proto = ctypes.WINFUNCTYPE(restype, ctypes.c_void_p, *argtypes)
        return proto(func)

    def _set_prop(store, key: _PROPERTYKEY, valor: str) -> None:
        # PROPVARIANT VT_LPWSTR: la cadena se aloja con CoTaskMemAlloc porque
        # PropVariantClear la liberará con CoTaskMemFree. SetValue hace su propia
        # copia dentro del store, así que se limpia el PROPVARIANT local al salir.
        pv = _PROPVARIANT()
        origen = ctypes.create_unicode_buffer(valor)
        n_bytes = (len(valor) + 1) * ctypes.sizeof(ctypes.c_wchar)
        buf = _ole32.CoTaskMemAlloc(n_bytes)
        if not buf:
            return
        ctypes.memmove(buf, origen, n_bytes)
        pv.vt = _VT_LPWSTR
        pv.p = buf
        try:
            set_value = _metodo(  # IPropertyStore::SetValue (índice 6 en la vtable)
                store, 6, ctypes.c_long,
                [ctypes.POINTER(_PROPERTYKEY), ctypes.POINTER(_PROPVARIANT)])
            set_value(store, ctypes.byref(key), ctypes.byref(pv))
        finally:
            _ole32.PropVariantClear(ctypes.byref(pv))

    def _aplicar(hwnd, relaunch_cmd, icon_res, display) -> None:
        store = ctypes.c_void_p()
        hr = _shell32.SHGetPropertyStoreForWindow(
            hwnd, ctypes.byref(_IID_IPropertyStore), ctypes.byref(store))
        if hr != 0 or not store:
            return
        try:
            _set_prop(store, _PKEY_ID, AUMID)
            if relaunch_cmd:
                _set_prop(store, _PKEY_RELAUNCH_CMD, relaunch_cmd)
            if icon_res:
                _set_prop(store, _PKEY_RELAUNCH_ICON, icon_res)
            if display:
                _set_prop(store, _PKEY_RELAUNCH_NAME, display)
            commit = _metodo(store, 7, ctypes.c_long, [])  # IPropertyStore::Commit
            commit(store)
        finally:
            release = _metodo(store, 2, ctypes.c_ulong, [])  # IUnknown::Release
            release(store)

    def _buscar_hwnd(titulo: str):
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

    def configurar_identidad(
        titulo: str,
        relaunch_cmd: str | None = None,
        icon_path: str | None = None,
        display: str | None = None,
        timeout: float = 20.0,
    ) -> None:
        """Etiqueta la ventana (la que tenga `titulo`) con la identidad de la app.
        La ventana la crea flet.exe de forma asíncrona, así que se sondea en un
        hilo demonio hasta `timeout` segundos. Nunca lanza (best-effort)."""
        try:
            _shell32.SetCurrentProcessExplicitAppUserModelID(AUMID)
        except Exception:  # noqa: BLE001
            pass
        icon_res = f"{icon_path},0" if icon_path else None

        def worker() -> None:
            fin = time.monotonic() + timeout
            while time.monotonic() < fin:
                hwnd = _buscar_hwnd(titulo)
                if hwnd:
                    try:
                        _aplicar(hwnd, relaunch_cmd, icon_res, display)
                    except Exception:  # noqa: BLE001
                        pass
                    return
                time.sleep(0.4)

        threading.Thread(target=worker, daemon=True).start()

else:  # fuera de Windows: no-op (la app solo se distribuye para Windows)

    def configurar_identidad(*_args, **_kwargs) -> None:
        return
