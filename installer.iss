#define MyAppName "Formulář pro předvyplnění"
#define MyAppVersion "1.4.0-dev"
#define MyAppPublisher "Krajan Ondřej"
#define MyAddinFile "FormularProPredvyplneni.xlam"

[Setup]
AppId={{D9DAE0E6-7898-45BC-AF4E-1E71712936DB}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={userappdata}\Microsoft\Excel\XLSTART
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x86compatible x64compatible
OutputDir=Output
OutputBaseFilename=FormularProPredvyplneni_Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#MyAppName}
UninstallFilesDir={localappdata}\Programs\FormularProPredvyplneni\Uninstall
CloseApplications=yes
RestartApplications=no
SetupLogging=yes

[Languages]
Name: "czech"; MessagesFile: "compiler:Languages\Czech.isl"

[Files]
Source: "Build\{#MyAddinFile}"; DestDir: "{userappdata}\Microsoft\Excel\XLSTART"; Flags: ignoreversion

[Icons]
Name: "{userprograms}\Formulář pro předvyplnění\Odinstalovat"; Filename: "{uninstallexe}"

[InstallDelete]
; Odstranění souborů starších vývojových verzí.
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\VyplnitNazvoslovi.xlam"
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\unins000.dat"
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\unins000.exe"
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\unins000.msg"

[UninstallDelete]
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\{#MyAddinFile}"
Type: dirifempty; Name: "{userprograms}\Formulář pro předvyplnění"

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    MsgBox('Doplněk byl nainstalován. Pokud je Excel spuštěný, zavřete jej a znovu otevřete.', mbInformation, MB_OK);
end;
