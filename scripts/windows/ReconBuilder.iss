#ifndef AppVersion
  #define AppVersion "1.1.0"
#endif

[Setup]
AppId={{2C94D43A-F10C-4AF9-86B1-0989ED339E64}
AppName=Auto Recon Biller
AppVersion={#AppVersion}
AppPublisher=Auto Recon Biller
VersionInfoVersion={#AppVersion}
VersionInfoProductVersion={#AppVersion}
DefaultDirName={autopf}\Recon Builder
DefaultGroupName=Auto Recon Biller
UninstallDisplayIcon={app}\ReconBuilder.exe
OutputDir=..\..\dist
OutputBaseFilename=ReconBuilder-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
SetupIconFile=..\..\assets\recon-builder.ico
CloseApplications=yes
RestartApplications=yes
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin

[Files]
Source: "..\..\dist\.build\ReconBuilder\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Auto Recon Biller"; Filename: "{app}\ReconBuilder.exe"
Name: "{autodesktop}\Auto Recon Biller"; Filename: "{app}\ReconBuilder.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Buat shortcut di Desktop"; GroupDescription: "Shortcut tambahan:"

[Run]
Filename: "{app}\ReconBuilder.exe"; Description: "Jalankan Auto Recon Biller"; Flags: nowait postinstall skipifsilent
