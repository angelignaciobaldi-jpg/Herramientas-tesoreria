@echo off
rem ============================================================
rem  Construye el ejecutable (dist\Tesoreria\Tesoreria.exe), listo
rem  para meterlo en el instalador (instalador.iss con Inno Setup).
rem
rem  NO empaqueta Chromium: el navegador se descarga en la primera
rem  ejecucion del RPA (a %LOCALAPPDATA%\...). Asi el instalador es
rem  liviano. El driver de Playwright (node) si va incluido via
rem  --collect-all, para poder lanzar/descargar el navegador.
rem
rem  Ejecutar dentro del entorno virtual (.venv activado), desde
rem  la carpeta del proyecto.
rem ============================================================
setlocal
cd /d "%~dp0"

echo Empaquetando con flet pack ...
rem tessdata / Imagenes se incluyen solo si existen (tessdata es opcional:
rem si no esta, el OCR usa el Tesseract instalado en el sistema).
set DATAARGS=
if exist "Imagenes\" set DATAARGS=%DATAARGS% --add-data "Imagenes:Imagenes"
if exist "tessdata\" set DATAARGS=%DATAARGS% --add-data "tessdata:tessdata"

flet pack app.py -n "Tesoreria" -D ^
  --icon "Imagenes\icon.ico" ^
  %DATAARGS% ^
  --hidden-import openpyxl ^
  --hidden-import xlwt ^
  --pyinstaller-build-args="--collect-all=playwright" ^
  -y
if errorlevel 1 (
  echo *** Fallo el empaquetado. ***
  pause & exit /b 1
)

echo.
echo ============================================================
echo   Listo: dist\Tesoreria\Tesoreria.exe
echo   Siguiente: compila instalador.iss con Inno Setup (iscc).
echo ============================================================
pause
