#define MyAppName "MR_Helper"
#define MyAppVersion "1.4.0-dev"
#define MyAppPublisher "Krajan Ondřej"
#define MyAddinFile "MR_Helper.xlam"
#define MyLogoFile "MR_Helper_logo.jpg"

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
OutputBaseFilename=MR_Helper_Setup
SetupIconFile=assets\MR_Helper.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#MyAppName}
UninstallFilesDir={localappdata}\Programs\MR_Helper\Uninstall
CloseApplications=yes
RestartApplications=no
SetupLogging=yes

[Languages]
Name: "czech"; MessagesFile: "compiler:Languages\Czech.isl"

[Files]
Source: "Build\{#MyAddinFile}"; DestDir: "{userappdata}\Microsoft\Excel\XLSTART"; Flags: ignoreversion
Source: "Build\MR_Helper_assets\{#MyLogoFile}"; DestDir: "{userappdata}\Microsoft\Excel\XLSTART\MR_Helper_assets"; Flags: ignoreversion

[Icons]
Name: "{userprograms}\MR_Helper\Odinstalovat"; Filename: "{uninstallexe}"

[InstallDelete]
; Odstranění souborů starších vývojových verzí.
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\VyplnitNazvoslovi.xlam"
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\FormularProPredvyplneni.xlam"
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\MR_Helper_logo.jpg"
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\MR_Helper_logo.png"
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\unins000.dat"
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\unins000.exe"
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\unins000.msg"

[UninstallDelete]
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\{#MyAddinFile}"
Type: files; Name: "{userappdata}\Microsoft\Excel\XLSTART\MR_Helper_assets\{#MyLogoFile}"
Type: dirifempty; Name: "{userappdata}\Microsoft\Excel\XLSTART\MR_Helper_assets"
Type: dirifempty; Name: "{userprograms}\MR_Helper"

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    MsgBox('Doplněk MR_Helper byl nainstalován. Pokud je Excel spuštěný, zavřete jej a znovu otevřete.', mbInformation, MB_OK);
end;
