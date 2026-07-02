[Setup]
; AppId FIJO: identifica la app entre versiones. Es lo que permite que el
; instalador descargado por el AutoUpdater actualice EN SITIO (sobrescribe) en
; vez de instalar una copia paralela. No lo cambies entre versiones.
AppId={{7E9F2A14-3C5B-4D88-9E21-6B0F4A2C1D33}
AppName=Herramientas Tesoreria
; Mantener en sync con core/version.py (__version__).
AppVersion=0.5.5
AppPublisher=Quetzaltic Solutions
DefaultDirName={commonpf}\Quetzaltic Solutions\Herramientas Tesoreria
DefaultGroupName=Quetzaltic Solutions
OutputDir=.\Output
; Debe coincidir con el asset que busca el AutoUpdater (NOMBRE_ASSET):
;   Instalador_Quetzaltic.exe
OutputBaseFilename=Instalador_Quetzaltic
Compression=lzma2/ultra64
SolidCompression=yes
; Escribir en Archivos de Programa y la actualizacion silenciosa requieren admin.
PrivilegesRequired=admin

[Files]
; Carpeta de salida de flet pack/PyInstaller (onedir). El nombre 'Tesoreria'
; debe coincidir con el -n del build (ver README). Si cambias el nombre del
; build, actualiza esta ruta y el nombre del .exe en [Icons] y [Run].
Source: ".\dist\Tesoreria\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; IconFilename apunta al icono incluido en {app}\Imagenes (la carpeta Imagenes
; se empaqueta completa). El .exe ya lleva el icono embebido, pero declararlo
; aqui garantiza que los accesos directos lo usen explicitamente.
Name: "{group}\Herramientas Tesoreria"; Filename: "{app}\Tesoreria.exe"; IconFilename: "{app}\Imagenes\icon.ico"
Name: "{commondesktop}\Herramientas Tesoreria"; Filename: "{app}\Tesoreria.exe"; IconFilename: "{app}\Imagenes\icon.ico"

[Run]
; Ejecuta la app al terminar la instalacion (no en modo silencioso/actualizacion).
Filename: "{app}\Tesoreria.exe"; Description: "{cm:LaunchProgram,Herramientas Tesoreria}"; Flags: nowait postinstall skipifsilent
