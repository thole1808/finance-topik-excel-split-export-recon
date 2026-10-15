#ifndef AppVersion
  #define AppVersion "1.1.0"
#endif

[Setup]
AppId={{2C94D43A-F10C-4AF9-86B1-0989ED339E64}
AppName=Recon Builder
AppVersion={#AppVersion}
AppPublisher=Recon Builder
DefaultDirName={autopf}\Recon Builder
DefaultGroupName=Recon Builder
UninstallDisplayIcon={app}\ReconBuilder.exe
OutputDir=..\..\dist
OutputBaseFilename=ReconBuilder-Setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin

[Files]
Source: "..\..\dist\ReconBuilder\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Recon Builder"; Filename: "{app}\ReconBuilder.exe"
Name: "{autodesktop}\Recon Builder"; Filename: "{app}\ReconBuilder.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Buat shortcut di Desktop"; GroupDescription: "Shortcut tambahan:"

[Run]
Filename: "{app}\ReconBuilder.exe"; Description: "Jalankan Recon Builder"; Flags: nowait postinstall skipifsilent
