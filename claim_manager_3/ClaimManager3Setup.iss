#ifndef AppVersion
  #define AppVersion "3.0.1"
#endif
#ifndef SourcePortableDir
  #define SourcePortableDir "..\Releases\desktop_portable\Claim Manager 3 Portable"
#endif
#ifndef OutputDir
  #define OutputDir "..\Releases\desktop_installer"
#endif

[Setup]
AppId={{4E0A24B5-8C63-4D50-94F1-83E5816BC8FD}
AppName=Claim Manager 3
AppVersion={#AppVersion}
AppVerName=Claim Manager 3 {#AppVersion}
AppPublisher=Fernando Marin
DefaultDirName={autopf}\Claim Manager 3
DefaultGroupName=Claim Manager 3
DisableProgramGroupPage=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=admin
OutputDir={#OutputDir}
OutputBaseFilename=Claim-Manager-3-Setup-{#AppVersion}
UninstallDisplayIcon={app}\Claim Manager 3\Claim Manager 3.exe
CloseApplications=yes
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "{#SourcePortableDir}\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs ignoreversion

[Icons]
Name: "{autodesktop}\Claim Manager 3"; Filename: "{app}\Claim Manager 3\Claim Manager 3.exe"; WorkingDir: "{app}"
Name: "{group}\Claim Manager 3"; Filename: "{app}\Claim Manager 3\Claim Manager 3.exe"; WorkingDir: "{app}"
Name: "{group}\Claim Manager 3 Office"; Filename: "{app}\Claim Manager 3 Office\Claim Manager 3 Office.exe"; WorkingDir: "{app}"
Name: "{group}\Uninstall Claim Manager 3"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\Claim Manager 3\Claim Manager 3.exe"; Description: "Launch Claim Manager 3"; Flags: nowait postinstall skipifsilent
