; ============================================
; BLP Stock Report Generator - Inno Setup Script
; ============================================
; Pastikan sudah menjalankan build_exe.bat terlebih dahulu
; sehingga folder dist\BLPStockReport\ sudah ada.

#define MyAppName "BLP Stock Report Generator"
#define MyAppVersion "1.0"
#define MyAppPublisher "BLP"
#define MyAppExeName "BLPStockReport.exe"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\BLPStockReport
DefaultGroupName={#MyAppName}
OutputDir=Output
OutputBaseFilename=BLPStockReport_Setup
SetupIconFile=
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Salin semua file dari hasil PyInstaller build
Source: "dist\BLPStockReport\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
