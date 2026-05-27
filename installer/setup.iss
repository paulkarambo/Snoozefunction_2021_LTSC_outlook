; Inno Setup script for Outlook Snooze Add-in
; Creates a per-user installer that places the manifest.xml

#define MyAppName "Outlook Snooze"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "SnoozeFunction"
#define MyAppURL "https://paulkarambo.github.io/Snoozefunction_2021_LTSC_outlook/"

[Setup]
AppId={{3A8F2C1E-5D7B-4A6C-9E3F-2B1D8C4A6F0E}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL=https://github.com/paulkarambo/Snoozefunction_2021_LTSC_outlook/issues
DefaultDirName={localappdata}\Programs\OutlookSnooze
DefaultGroupName=Outlook Snooze
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=Output
OutputBaseFilename=OutlookSnooze-Setup-{#MyAppVersion}
Compression=lzma
SolidCompression=yes
UninstallDisplayName=Outlook Snooze Add-in
DisableDirPage=auto

[Languages]
Name: english; MessagesFile: compiler:Default.isl

[Files]
Source: "..\manifest.xml"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{userdesktop}\Outlook Snooze - Sideload Instructions"; Filename: "{app}\SIDELOAD.txt"

[Dirs]
Name: "{app}"

[Run]
Filename: "{cmd}"; Parameters: "/C start notepad.exe ""{app}\SIDELOAD.txt"""; Flags: postinstall nowait skipifsilent; Description: "Open sideloading instructions"

[UninstallDelete]
Type: files; Name: "{app}\manifest.xml"
Type: files; Name: "{app}\SIDELOAD.txt"
Type: dirifempty; Name: "{app}"

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  SideloadTxt: string;
begin
  if CurStep = ssPostInstall then
  begin
    SideloadTxt := ExpandConstant('{app}\SIDELOAD.txt');
    SaveStringToFile(SideloadTxt,
      'OUTLOOK SNOOZE ADD-IN - SIDELOAD INSTRUCTIONS' + #13#10 +
      '============================================' + #13#10 + #13#10 +
      'The add-in has been installed to:' + #13#10 +
      ExpandConstant('{app}') + #13#10 + #13#10 +
      'To install it in Outlook:' + #13#10 + #13#10 +
      '1. Open Outlook 2021' + #13#10 +
      '2. Click "Get Add-ins" on the Home ribbon' + #13#10 +
      '3. Click "My add-ins" (left sidebar)' + #13#10 +
      '4. Click "Add a custom add-in" at the bottom' + #13#10 +
      '5. Select "Add from file..."' + #13#10 +
      '6. Browse to: ' + ExpandConstant('{app}\manifest.xml') + #13#10 +
      '7. Click Open' + #13#10 + #13#10 +
      'The add-in will appear in the Outlook ribbon' + #13#10 +
      'when reading an email.' + #13#10 + #13#10 +
      'SUPPORT: https://github.com/paulkarambo/Snoozefunction_2021_LTSC_outlook/issues',
      False);
  end;
end;
