#define MyAppName "Quadstate Renewals"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Quadstate Insurance"
#define MyAppExeName "Quadstate Renewals.exe"

[Setup]
AppId={{9BBB9199-FC4A-5468-8899-4A6D8B73BFE5}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\{#MyAppName}
DefaultGroupName={#MyAppName}
PrivilegesRequired=lowest
OutputDir=C:/Downloads
OutputBaseFilename=Quadstate Renewals_Setup_v1.0.0
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=C:/Users/wayde/Documents/Quadstate Renewals.ico

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "D:/Dev/Python_Projects/QuadstateRenewals\dist\Quadstate Renewals\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: files; Name: "{app}\app.log"
Type: files; Name: "{localappdata}\{#MyAppName}\app.log"
Type: dirifempty; Name: "{localappdata}\{#MyAppName}"