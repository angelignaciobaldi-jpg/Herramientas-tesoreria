@echo off
title Herramienta Integral de Tesoreria
cd /d "%~dp0"
echo ============================================================
echo   Herramienta Integral de Tesoreria
echo ============================================================
echo.
echo [1/4] Buscando actualizaciones (rama main)...
git checkout main
if errorlevel 1 (
  echo   * No se pudo cambiar a la rama main; se abrira la version local.
) else (
  git pull origin main
  if errorlevel 1 echo   * No se pudo actualizar; se abrira la version local.
)
echo.
echo [2/4] Verificando que Python este instalado...
set "PYEXE="
set "PYWEXE="
where py.exe >nul 2>nul && py -c "" >nul 2>nul && (set "PYEXE=py" & set "PYWEXE=pyw")
if not defined PYEXE where python.exe >nul 2>nul && python -c "" >nul 2>nul && (set "PYEXE=python" & set "PYWEXE=pythonw")
if not defined PYEXE (
  echo.
  echo *** No se encontro Python en este equipo. ***
  echo.
  echo Ejecuta primero  instalar.bat  para instalar Python, Git y Tesseract.
  echo Si YA lo instalaste, cierra esta ventana y vuelve a abrir iniciar.bat
  echo ^(Windows a veces necesita una ventana nueva para reconocer Python^).
  echo.
  echo Si sigue fallando: Configuracion ^> Aplicaciones ^> Alias de ejecucion
  echo y desactiva los alias "python.exe" y "python3.exe" de Microsoft Store.
  echo.
  pause
  exit /b 1
)
echo   Python detectado: %PYEXE%
echo.
echo [3/4] Verificando dependencias...
%PYEXE% -m pip install -r requirements.txt --quiet --disable-pip-version-check
echo.
echo [4/4] Iniciando la aplicacion...
echo   (esta ventana se cierra sola si la app abre correctamente)
powershell -NoProfile -ExecutionPolicy Bypass -Command "$log = Join-Path $env:TEMP ('tesoreria_error_' + $PID + '.log'); $p = Start-Process %PYWEXE% -ArgumentList 'app.py' -WorkingDirectory (Get-Location).Path -PassThru -RedirectStandardError $log; Start-Sleep -Seconds 7; Get-ChildItem (Join-Path $env:TEMP 'tesoreria_error_*.log') -ErrorAction SilentlyContinue | Where-Object { $_.FullName -ne $log } | Remove-Item -Force -ErrorAction SilentlyContinue; if ($p.HasExited -and $p.ExitCode -ne 0) { Write-Host ''; Write-Host '*** La aplicacion no pudo iniciar. Detalle del error: ***' -ForegroundColor Red; if (Test-Path $log) { Get-Content $log }; exit 1 } else { exit 0 }"
if errorlevel 1 (
  echo.
  echo Copia el error de arriba si necesitas reportarlo.
  pause
)
