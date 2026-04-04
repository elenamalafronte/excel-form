; Inno Setup script for ExcelForm
; Build this after running PyInstaller and confirming dist\ExcelForm exists.

#define AppName "ExcelForm"
#define AppVersion "1.0.0"
#define AppPublisher "ExcelForm"
#define AppExeName "ExcelForm.exe"
#define AppDirName "ExcelForm"

[Setup]
AppId=ExcelFormApp
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\{#AppDirName}
DefaultGroupName={#AppName}
DisableProgramGroupPage=yes
OutputDir=installer-output
OutputBaseFilename=ExcelFormSetup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=lowest

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "dist\ExcelForm\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch {#AppName}"; Flags: nowait postinstall skipifsilent
