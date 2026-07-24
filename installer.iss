; Inno Setup script for GST Reconciliation Tool.
; Build the app first (PyInstaller --onedir, see HANDOFF/README), then compile
; this with ISCC.exe to produce a real Windows installer (Start Menu entry,
; uninstaller, Add/Remove Programs listing) around dist\GST-Reconciliation-Tool\.

#define MyAppName "GST Reconciliation Tool"
#define MyAppVersion "1.1.1"
#define MyAppPublisher "Sai Srinivasa Bharadwaj"
#define MyAppExeName "GST-Reconciliation-Tool.exe"

[Setup]
AppId={{B6C6E1B0-6B7B-4B1E-9C1B-5B1B7B7B7B7B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer_output
OutputBaseFilename=GST-Reconciliation-Tool-Setup
SetupIconFile=assets\icon.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "dist\GST-Reconciliation-Tool\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
