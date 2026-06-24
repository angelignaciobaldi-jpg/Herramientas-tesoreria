@echo off
chcp 65001 >nul
title Preparar paquete para compartir
cd /d "%~dp0"

REM Resuelve la carpeta padre a una ruta absoluta (sin "..").
pushd "%~dp0.."
set "DESTINO=%CD%\Tesoreria-paquete"
popd

echo ============================================================
echo   Preparando paquete para compartir a usuarios nuevos
echo ------------------------------------------------------------
echo   Copia el proyecto (incluyendo .git para auto-actualizar y
echo   los datos locales) a una carpeta limpia, sin artefactos ni
echo   datos de prueba.
echo ============================================================
echo   Destino: %DESTINO%
echo.

robocopy "%~dp0." "%DESTINO%" /E /R:1 /W:1 /NFL /NDL /NJH ^
  /XD "dist" "build" "__pycache__" "CARATULAS" "Archivos TXT" ".vscode" ^
  /XF "tesoreria.db" "Tesoreria.spec" "ALTABANREGIO 1.xls" "Codigo macro excel"

echo.
echo ============================================================
echo   Listo. Comparte la carpeta:
echo     %DESTINO%
echo.
echo   (Incluye .git, tessdata y Cuentas bancarias. El usuario solo
echo    corre instalar.bat una vez y luego iniciar.bat.)
echo ============================================================
pause
