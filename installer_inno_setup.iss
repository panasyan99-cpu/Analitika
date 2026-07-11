#ifndef MyAppVersion
  #define MyAppVersion "1.1.6"
#endif
#ifndef MyAppVersionQuad
  #define MyAppVersionQuad "1.1.6.0"
#endif

#define MyAppName "Аналитика"
#define MyAppPublisher "Princess Jewelry"
#define MyAppExeName "Analitika.exe"

[Setup]
; Never change AppId: Windows and future installers use it to find the existing installation.
AppId={{47D9F713-13E1-4B7D-8E9A-90DC0A82E2D1}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://github.com/panasyan99-cpu/Analitika
AppUpdatesURL=https://github.com/panasyan99-cpu/Analitika/releases
VersionInfoVersion={#MyAppVersionQuad}
VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription=Princess Jewelry sales analytics
DefaultDirName={localappdata}\Programs\Analitika
DefaultGroupName={#MyAppName}
OutputDir=installer_output
OutputBaseFilename=Analitika_Setup_{#MyAppVersion}
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
UsePreviousAppDir=yes
UsePreviousGroup=yes
UsePreviousTasks=yes
DisableProgramGroupPage=yes
SetupLogging=yes

[Files]
Source: "dist\Analitika\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "update_config.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "version.json"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
Name: "{userdocs}\Analitika\Output"
Name: "{userdocs}\Analitika\History"
Name: "{userdocs}\Analitika\Settings"
Name: "{userdocs}\Analitika\Logs"
Name: "{userdocs}\Analitika\Updates"
Name: "{userdocs}\Analitika\Reports"

[Tasks]
Name: "desktopicon"; Description: "Создать ярлык на рабочем столе"; GroupDescription: "Ярлыки:"; Flags: unchecked

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; WorkingDir: "{app}"; Tasks: desktopicon

[Run]
; Interactive first install: optional checkbox on the final page.
Filename: "{app}\{#MyAppExeName}"; Description: "Запустить {#MyAppName}"; Flags: nowait postinstall skipifsilent
; Silent updater: reopen the application after replacing files.
Filename: "{app}\{#MyAppExeName}"; Parameters: "/updated"; Flags: nowait skipifnotsilent; Check: WizardSilent
