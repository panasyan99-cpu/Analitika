#define MyAppName "Аналитика"
#define MyAppVersion "1.1.4"
#define MyAppPublisher "Princess Jewelry"
#define MyAppExeName "Analitika.exe"

[Setup]
AppId={{47D9F713-13E1-4B7D-8E9A-90DC0A82E2D1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
VersionInfoVersion=1.1.4.0
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription=Princess Jewelry sales analytics
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=installer_output
OutputBaseFilename=Analitika_Setup_1.1.4
SetupIconFile=assets\analitika.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
RestartApplications=no

[Files]
Source: "dist\Analitika\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "update_config.json"; DestDir: "{app}"; Flags: onlyifdoesntexist

[Dirs]
Name: "{userappdata}\Analitika\Output"
Name: "{userappdata}\Analitika\logs"

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; GroupDescription: "Ярлыки:"; Flags: unchecked

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Запустить {#MyAppName}"; Flags: nowait postinstall skipifsilent
