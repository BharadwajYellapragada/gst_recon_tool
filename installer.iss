; Inno Setup script for GST Reconciliation Tool.
; Build the app first (PyInstaller --onedir, see HANDOFF/README), then compile
; this with ISCC.exe to produce a real Windows installer (Start Menu entry,
; uninstaller, Add/Remove Programs listing) around dist\GST-Reconciliation-Tool\.

#define MyAppName "GST Reconciliation Tool"
#define MyAppVersion "1.1.3"
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
Name: "freshstart"; Description: "Start fresh — do NOT carry forward clients, uploads, or reports saved by a previous install on this computer"; GroupDescription: "Existing data on this computer:"; Flags: unchecked

[Files]
Source: "dist\GST-Reconciliation-Tool\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if (CurPageID = wpSelectTasks) and WizardIsTaskSelected('freshstart') then
  begin
    if MsgBox('You checked "Start fresh". Any clients, uploads, and reports already saved on ' +
              'this computer from a previous install will be archived (renamed aside, not ' +
              'permanently deleted) and will no longer appear in the app — you will set up a ' +
              'brand new PIN on first launch. Your activation key is unaffected either way.' + #13#10 + #13#10 +
              'Continue with a fresh start?',
              mbConfirmation, MB_YESNO) = IDNO then
      Result := False;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  DataDir, Stamp, SecPath, DbPath: String;
begin
  // Only security.json (password/machine-lock state) and gst_recon.db.enc (the
  // encrypted database) are archived here -- deliberately NOT license.json,
  // which holds the activation key and is keyed purely to the machine
  // fingerprint, independent of the password/database. A fresh start must not
  // force re-activation.
  if (CurStep = ssPostInstall) and WizardIsTaskSelected('freshstart') then
  begin
    DataDir := ExpandConstant('{localappdata}\GSTReconTool');
    Stamp := GetDateTimeString('yyyymmdd_hhnnss', #0, #0);
    SecPath := DataDir + '\security.json';
    DbPath := DataDir + '\gst_recon.db.enc';
    if FileExists(SecPath) then
      RenameFile(SecPath, SecPath + '.orphaned_' + Stamp);
    if FileExists(DbPath) then
      RenameFile(DbPath, DbPath + '.orphaned_' + Stamp);
  end;
end;
