# Herramienta Integral de Tesoreria

Aplicacion de escritorio (Flet) para el area de tesoreria:

- **Alta de beneficiarios:** lee estados de cuenta (PDF/imagen) con OCR,
  identifica CLABE, beneficiario, alias y email; tabla editable; exporta el TXT
  de dispersion (Bancomer) o el Excel de alta (Banregio).
- **Generar dispersion devoluciones:** captura movimientos y genera el TXT del
  banco elegido (Banregio/Bancomer) y un reporte Excel.
- **Dispersion (No Pemex):** RPA (Playwright) que entra al SIPP, busca las
  solicitudes de pago por empresa/tipo/fechas, descarga los reportes Excel y los
  vuelca en una tabla (agrupada por cuenta bancaria, con tabs por empresa).

Las credenciales del SIPP y el Excel de cuentas se capturan en el menu
**Configuracion (⚙)** de la barra superior.

---

## Instalar en una maquina nueva (Instalador + actualizacion automatica)

La app se distribuye como **instalador de Windows** (`Instalador_Quetzaltic.exe`,
generado con Inno Setup) y **se actualiza sola** desde las *releases* de GitHub
mediante el AutoUpdater (`core/auto_updater.py`).

1. Ejecuta **`Instalador_Quetzaltic.exe`**. Se instala **por usuario** en
   `%LOCALAPPDATA%\Programs\...`, asi que **NO pide permisos de administrador**.
2. Abre **"Herramientas Tesoreria"** desde el menu Inicio o el acceso directo del
   escritorio.

> Al iniciar, la app revisa si hay una version mas nueva publicada en GitHub; si
> la hay, muestra una pantalla de "Actualizando...", descarga el nuevo instalador,
> lo aplica **sin UAC** (al ser por usuario) y **reinicia la app sola**. El usuario
> no reconstruye ni reinstala nada a mano.

El AutoUpdater necesita el PAT del repo privado en la variable de entorno
`QUETZALTIC_GITHUB_PAT` (o un archivo `.env` junto al `.exe`; ver `.env.example`).

### Migracion de la version antigua (admin -> por usuario)

Las versiones anteriores se instalaban en `Archivos de Programa` y requerian
administrador. Windows trata esa instalacion "per-machine" como **distinta** de la
nueva "per-user", asi que **no se migran solas**: quien ya tenga la version admin
debe **reinstalar UNA vez** con el instalador nuevo (conviene desinstalar la vieja
primero). A partir de ahi, todas las actualizaciones son **sin UAC**.

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

Para **actualizar las cuentas**: abre el menu **Configuracion (⚙) -> Catalogo de
cuentas -> Adjuntar Excel de cuentas**. La app lo copia a su ubicacion, lo valida
(si el formato no es el esperado, hace *rollback* al anterior) y recarga el
catalogo al momento. Tambien puedes reemplazar el Excel a mano en
`Cuentas bancarias/` y reabrir la app.

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
  configuracion.py       Modal de Configuracion (credenciales SIPP, Excel cuentas)
  alta_beneficiarios.py  Pantalla "Alta de beneficiarios"
  devoluciones.py        Pantalla "Generar dispersion devoluciones"
  dispersion_no_pemex.py Pantalla "Dispersion (No Pemex)" (RPA + tabla)
```

### Scripts (obsoletos)
`iniciar.bat`, `instalar.bat` y `preparar_paquete.bat` quedaron **deprecados**
con el cambio al modelo de instalador + autoactualizacion. Solo muestran un aviso.

### Publicar una nueva version (automatico via GitHub Actions)

El pipeline `.github/workflows/compilar.yml` se dispara al **publicar un Release** y
hace TODO: sincroniza la version, compila la app, arma el instalador con Inno Setup
y sube `Instalador_Quetzaltic.exe` como asset. Los demas equipos lo detectan y se
actualizan solos.

1. Mergea tu codigo a `main`.
2. En GitHub: **Releases -> Draft a new release** -> crea el tag con la nueva
   version (`0.5.6` o `v0.5.6`) -> **Publish**.
3. Listo. El CI escribe ese tag en `core/version.py` y en `AppVersion`
   (`instalador.iss`) **antes de compilar**, asi que el instalador SIEMPRE reporta
   exactamente su tag (por eso no hay bucles de actualizacion).

> **El tag del Release ES la version.** Ya NO hay que subir `version.py` /
> `AppVersion` a mano. Reglas: el tag nuevo debe ser **mayor** que el anterior, y
> crea un Release **nuevo** (editar uno viejo NO dispara el build).

### Compilar el instalable localmente (opcional, para probar)

1. `construir.bat` -> genera `dist\Tesoreria\Tesoreria.exe` (onedir; incluye el
   driver de Playwright via `--collect-all`; **no** empaqueta Chromium).
2. `iscc instalador.iss` -> genera `Output\Instalador_Quetzaltic.exe`.

> **RPA (Playwright):** el `.exe` NO trae Chromium (para no inflar el
> instalador). La primera vez que se usa el RPA, `core/rpa_sipp.py` descarga
> Chromium a `%LOCALAPPDATA%\Quetzaltic Solutions\Herramientas de Tesoreria\ms-playwright`
> (requiere internet, igual que el propio RPA). Las siguientes veces ya está.
