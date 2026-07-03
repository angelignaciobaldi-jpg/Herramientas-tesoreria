[Setup]
; AppId FIJO: identifica la app entre versiones. Es lo que permite que el
; instalador descargado por el AutoUpdater actualice EN SITIO (sobrescribe) en
; vez de instalar una copia paralela. No lo cambies entre versiones.
AppId={{7E9F2A14-3C5B-4D88-9E21-6B0F4A2C1D33}
AppName=Herramientas Tesoreria
; Mantener en sync con core/version.py (__version__).
AppVersion=0.5.6
AppPublisher=Quetzaltic Solutions
; Instalacion POR USUARIO (en %LOCALAPPDATA%\Programs), NO en Archivos de
; Programa. Al ser una carpeta escribible por el usuario, la actualizacion
; silenciosa la sobrescribe SIN pedir permisos de administrador (sin UAC).
; {autopf} con PrivilegesRequired=lowest resuelve a {localappdata}\Programs.
DefaultDirName={autopf}\Quetzaltic Solutions\Herramientas Tesoreria
DefaultGroupName=Quetzaltic Solutions
OutputDir=.\Output
; Debe coincidir con el asset que busca el AutoUpdater (NOMBRE_ASSET):
;   Instalador_Quetzaltic.exe
OutputBaseFilename=Instalador_Quetzaltic
Compression=lzma2/ultra64
SolidCompression=yes
; 'lowest' = no solicita elevacion (sin UAC). Requisito para actualizar sin admin.
PrivilegesRequired=lowest

[Files]
; Carpeta de salida de flet pack/PyInstaller (onedir). El nombre 'Tesoreria'
; debe coincidir con el -n del build (ver README). Si cambias el nombre del
; build, actualiza esta ruta y el nombre del .exe en [Icons] y [Run].
Source: ".\dist\Tesoreria\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; Motor Tesseract OCR empaquetado (lo prepara el CI en '.\tesseract_bundle').
; skipifsourcedoesntexist: si no esta (build local sin OCR empaquetado) NO rompe
; la compilacion del instalador; la app cae a un Tesseract del sistema si existe.
Source: ".\tesseract_bundle\*"; DestDir: "{app}\Tesseract-OCR"; Flags: ignoreversion recursesubdirs createallsubdirs skipifsourcedoesntexist
; El PAT del AutoUpdater (QUETZALTIC_GITHUB_PAT) NO se distribuye: se registra a
; mano en cada maquina (variable de entorno del sistema). Por eso el .env no se
; incluye aqui.

[Icons]
; IconFilename apunta al icono incluido en {app}\Imagenes (la carpeta Imagenes
; se empaqueta completa). El .exe ya lleva el icono embebido, pero declararlo
; aqui garantiza que los accesos directos lo usen explicitamente. {autodesktop} y
; {group} son por-usuario (coherentes con la instalacion no elevada).
Name: "{group}\Herramientas Tesoreria"; Filename: "{app}\Tesoreria.exe"; IconFilename: "{app}\Imagenes\icon.ico"
; {autodesktop} = escritorio del usuario (no el comun, que requeriria admin).
Name: "{autodesktop}\Herramientas Tesoreria"; Filename: "{app}\Tesoreria.exe"; IconFilename: "{app}\Imagenes\icon.ico"

[Run]
; Ejecuta la app al terminar la instalacion (no en modo silencioso/actualizacion).
Filename: "{app}\Tesoreria.exe"; Description: "{cm:LaunchProgram,Herramientas Tesoreria}"; Flags: nowait postinstall skipifsilent
