@echo off
title [OBSOLETO] preparar_paquete.bat
rem ============================================================
rem  OBSOLETO - Este script ya no se usa.
rem
rem  Antes armaba una carpeta del codigo fuente para compartir.
rem  Con el nuevo modelo, lo que se distribuye es el INSTALADOR
rem  (Instalador_Quetzaltic.exe), generado al compilar la app con
rem  flet pack/PyInstaller y empaquetarla con Inno Setup
rem  (instalador.iss). Ver la seccion de build en el README.
rem ============================================================
echo ============================================================
echo   Este script (preparar_paquete.bat) quedo OBSOLETO.
echo.
echo   Para generar lo distribuible:
echo     1) Compila la app:   ver README (flet pack ...)
echo     2) Empaqueta:        compila instalador.iss con Inno Setup
echo        -^> Output\Instalador_Quetzaltic.exe
echo ============================================================
pause
