# Herramienta Integral de Tesoreria

Aplicacion de escritorio (Flet) para el area de tesoreria:

- **Alta de beneficiarios:** lee estados de cuenta (PDF/imagen) con OCR,
  identifica CLABE, beneficiario, alias y email; tabla editable; exporta el TXT
  de dispersion (Bancomer) o el Excel de alta (Banregio).
- **Generar dispersion devoluciones:** captura movimientos y genera el TXT del
  banco elegido (Banregio/Bancomer) y un reporte Excel.

---

## Instalar en una maquina nueva (Instalador + actualizacion automatica)

La app se distribuye como **instalador de Windows** (`Instalador_Quetzaltic.exe`,
generado con Inno Setup) y **se actualiza sola** desde las *releases* de GitHub
mediante el AutoUpdater (`core/auto_updater.py`).

1. Ejecuta **`Instalador_Quetzaltic.exe`** (pide permisos de administrador).
2. Abre **"Herramientas de Tesoreria"** desde el menu Inicio o el acceso directo
   del escritorio.

> Al iniciar, la app revisa si hay una version mas nueva publicada en GitHub; si
> la hay, descarga el nuevo instalador y lo aplica en silencio. **El usuario no
> reconstruye ni reinstala nada a mano.**

El AutoUpdater necesita el PAT del repo privado en la variable de entorno
`QUETZALTIC_GITHUB_PAT` (o un archivo `.env` junto al `.exe`; ver `.env.example`).

> **Scripts obsoletos:** `iniciar.bat`, `instalar.bat` y `preparar_paquete.bat`
> pertenecian al modelo anterior (correr desde el codigo + `git pull`) y quedaron
> deprecados. Ya no se usan.

---

## Datos que viven en cada maquina (NO estan en el repo)

Por seguridad, el repositorio **solo contiene codigo**. Estos archivos viven en
la maquina (junto a la app instalada o en la carpeta del proyecto en desarrollo):

- `Cuentas bancarias/CUENTAS BANCARIAS .xlsx` - catalogo de cuentas por empresa.
- `tessdata/` - modelos de OCR (incluye espanol).
- `tesoreria.db` y `_cuentas_cache.json` - se crean/actualizan solos.
- Tesseract OCR instalado en `C:\Program Files\Tesseract-OCR\`.

Para **actualizar las cuentas**: edita el Excel en `Cuentas bancarias/`,
cierralo y reabre la app (el cache tolera tenerlo abierto).

---

## Para desarrollar (editar el codigo)

Trabajo por ramas (una por persona). En desarrollo se corre desde el codigo y se
maneja git a mano.

```
pip install -r requirements.txt
python app.py
```

### Estructura
```
app.py                   Shell: ventana, logo, tema, pestanas
core/                    Backend (OCR, BD, extractores, exportadores, rutas)
ui/
  comun.py               Constantes y utilidades compartidas
  alta_beneficiarios.py  Pantalla "Alta de beneficiarios"
  devoluciones.py        Pantalla "Generar dispersion devoluciones"
```

### Scripts (obsoletos)
`iniciar.bat`, `instalar.bat` y `preparar_paquete.bat` quedaron **deprecados**
con el cambio al modelo de instalador + autoactualizacion. Solo muestran un aviso.

### Generar el instalable (build + empaquetado)
1. **Compilar la app** (onedir). El script empaqueta con los flags correctos
   (incluye el driver de Playwright via `--collect-all`). **No** empaqueta
   Chromium: el navegador se descarga en la primera ejecucion del RPA (instalador
   liviano). Corre dentro del `.venv`:
   ```
   construir.bat
   ```
   Genera `dist\Tesoreria\Tesoreria.exe`. (El nombre `Tesoreria` debe coincidir
   con `instalador.iss`.)
2. **Empaquetar** compilando `instalador.iss` con Inno Setup (`iscc instalador.iss`).
   Genera `Output\Instalador_Quetzaltic.exe`.
3. **Publicar** una *release* en GitHub con `tag_name` = la nueva version (mayor a
   `core/version.py`) y sube ese `.exe` como asset con nombre exacto
   `Instalador_Quetzaltic.exe`. El AutoUpdater lo detecta y actualiza a los demas.

> Antes de compilar, sube `AppVersion` en `instalador.iss` y `__version__` en
> `core/version.py` (deben quedar iguales).

> **RPA (Playwright):** el `.exe` NO trae Chromium (para no inflar el
> instalador). La primera vez que se usa el RPA, `core/rpa_sipp.py` descarga
> Chromium a `%LOCALAPPDATA%\Quetzaltic Solutions\Herramientas de Tesoreria\ms-playwright`
> (requiere internet, igual que el propio RPA). Las siguientes veces ya está.
