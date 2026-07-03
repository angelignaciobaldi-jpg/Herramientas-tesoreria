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
; Icono para los accesos directos, copiado a la RAIZ de {app}. Se toma del
; codigo fuente (no del build): PyInstaller (onedir) mete 'Imagenes' dentro de
; {app}\_internal, asi que un IconFilename a {app}\Imagenes\icon.ico no existiria
; y el acceso saldria con un cuadro blanco. Copiarlo aqui garantiza la ruta.
Source: ".\Imagenes\icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; IconFilename apunta a {app}\icon.ico (copiado a la raiz en [Files]). NO usar
; {app}\Imagenes\icon.ico: PyInstaller (onedir) coloca 'Imagenes' dentro de
; {app}\_internal, por lo que esa ruta no existe y el acceso sale sin icono.
; WorkingDir: {app} para que los accesos arranquen en la carpeta de la app (si
; el usuario ancla ESTE acceso a la barra de tareas, Windows hereda el "Iniciar
; en"). La app ademas fija el CWD por codigo para cubrir cualquier lanzador.
; AppUserModelID: DEBE coincidir con AUMID en core/win_taskbar.py para que el
; acceso y la ventana (creada por el flet.exe cliente) se agrupen como la MISMA
; app en la barra de tareas.
Name: "{group}\Herramientas Tesoreria"; Filename: "{app}\Tesoreria.exe"; WorkingDir: "{app}"; IconFilename: "{app}\icon.ico"; AppUserModelID: "QuetzalticSolutions.HerramientasTesoreria"
; {autodesktop} = escritorio del usuario (no el comun, que requeriria admin).
Name: "{autodesktop}\Herramientas Tesoreria"; Filename: "{app}\Tesoreria.exe"; WorkingDir: "{app}"; IconFilename: "{app}\icon.ico"; AppUserModelID: "QuetzalticSolutions.HerramientasTesoreria"

[Run]
; Ejecuta la app al terminar la instalacion (no en modo silencioso/actualizacion).
Filename: "{app}\Tesoreria.exe"; Description: "{cm:LaunchProgram,Herramientas Tesoreria}"; Flags: nowait postinstall skipifsilent
