#ifndef AppVersion
  #define AppVersion "12.0.126"
#endif
#define AppName "GridForge"
#define AppPublisher "GridForge Project"
#define AppExeName "GridForge.exe"

[Setup]
AppId={{4E772F66-801A-4E47-9E8D-03B3A5697091}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={localappdata}\Programs\GridForge
DefaultGroupName=GridForge
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=..\..\build\release
OutputBaseFilename=GridForge-Setup-{#AppVersion}-win64
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupIconFile=..\gridforge.ico
UninstallDisplayIcon={app}\{#AppExeName}
CloseApplications=yes
RestartApplications=no
ChangesAssociations=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked
Name: "vbashortcut"; Description: "Show the optional Microsoft Excel VBA bridge in the Start menu"; GroupDescription: "Optional Excel integration:"; Flags: unchecked

[Files]
Source: "..\..\build\portable\GridForge\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\..\build\portable\GridForgeExcelBridge.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\..\docs\DESKTOP_USER_GUIDE.md"; DestDir: "{app}\Documentation"; Flags: ignoreversion
Source: "..\..\LICENSE"; DestDir: "{app}\Documentation"; Flags: ignoreversion

[Icons]
Name: "{group}\GridForge"; Filename: "{app}\{#AppExeName}"
Name: "{group}\GridForge Macro Studio"; Filename: "{app}\{#AppExeName}"; Parameters: "--macro-workbench"
Name: "{group}\Uninstall GridForge"; Filename: "{uninstallexe}"
Name: "{autodesktop}\GridForge"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon
Name: "{group}\Excel VBA Bridge (trusted workbooks only)"; Filename: "{app}\GridForgeExcelBridge.exe"; Tasks: vbashortcut

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch GridForge"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Intentionally does not erase %LOCALAPPDATA%\GridForge\WebView2 workbook drafts.
; Users should remove personal workbook data themselves if desired.






