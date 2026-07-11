; ============================================================================
;  VistaType LP + Braille Macros -- Inno Setup installer
; ============================================================================
;  Replaces the manual "copy three files to three locations" procedure.
;  Per-user install (no admin needed); everything lands in the user's profile.
;
;  What it does automatically (each item is a support call it removes):
;    * Copies LPandBRL.dotm            -> %AppData%\Microsoft\Word\STARTUP
;                                         (creates STARTUP if it doesn't exist)
;    * Copies LargePrintTemplate.dotx  -> %AppData%\Microsoft\Templates
;    * Copies Word.officeUI            -> %AppData%\Microsoft\Office
;                                         (backs up the user's existing one first)
;    * Registers the STARTUP folder as a Word Trusted Location
;    * Deletes the obsolete "Large Print Templates" folder
;    * Refuses to run while Word or Outlook is open (with a clear message)
;    * Provides a clean uninstaller
;
;  Build:  compiled by ISCC.exe on the Windows box (see installer/README.md and
;          `make installer`). Source files are staged into ..\dist first.
;
;  STILL NEEDS A TEST PASS on a real machine -- the Trusted Location registry
;  and running-Office detection are correct in principle but unverified here.
; ============================================================================

#define AppVer      "2.2.3"
#define DotmName    "LPandBRL.dotm"
#define DotxName    "LargePrintTemplate.dotx"
#define RibbonName  "Word.officeUI"
; Directory holding the three shipping files (staged by `make installer`).
#ifndef SrcDir
  #define SrcDir "..\dist"
#endif

[Setup]
AppName=VistaType LP + Braille Macros
AppVersion={#AppVer}
AppPublisher=Jerry Whittaker
AppPublisherURL=mailto:jerry@thewhittakers.org
DefaultDirName={userappdata}\Microsoft\Word\STARTUP
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=..\dist
OutputBaseFilename=VistaType-LP-Setup-{#AppVer}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName=VistaType LP + Braille Macros

[Files]
; Macro add-in -> STARTUP (DefaultDirName). {app} == the STARTUP folder here.
Source: "{#SrcDir}\{#DotmName}";   DestDir: "{app}";                         Flags: ignoreversion
; Large-print styles template -> user Templates folder.
Source: "{#SrcDir}\{#DotxName}";   DestDir: "{userappdata}\Microsoft\Templates"; Flags: ignoreversion
; Ribbon/QAT -> Office folder. Existing file is backed up in CurStepChanged.
Source: "{#SrcDir}\{#RibbonName}"; DestDir: "{userappdata}\Microsoft\Office";     Flags: ignoreversion

[InstallDelete]
; Remove the obsolete template folder from older versions.
Type: filesandordirs; Name: "{userappdata}\Microsoft\Templates\Large Print Templates"

[UninstallDelete]
Type: files; Name: "{app}\{#DotmName}"
Type: files; Name: "{userappdata}\Microsoft\Templates\{#DotxName}"

[Code]
const
  WORD_CLASS    = 'OpusApp';           { Word main window class }
  OUTLOOK_CLASS = 'rctrl_renwnd32';    { Outlook main window class }

{ --- Abort if Word or Outlook is running (they lock the target files). --- }
function InitializeSetup(): Boolean;
begin
  Result := True;
  if (FindWindowByClassName(WORD_CLASS) <> 0) or
     (FindWindowByClassName(OUTLOOK_CLASS) <> 0) then
  begin
    MsgBox('Please close Microsoft Word and Microsoft Outlook, then run this '
      + 'installer again.', mbError, MB_OK);
    Result := False;
  end;
end;

{ --- Find the installed Word version key (16.0, 15.0, ...) under HKCU. --- }
function OfficeVersion(): String;
var
  Versions: TArrayOfString;
  I: Integer;
begin
  Result := '';
  Versions := ['16.0', '15.0', '14.0'];
  for I := 0 to GetArrayLength(Versions) - 1 do
    if RegKeyExists(HKCU, 'Software\Microsoft\Office\' + Versions[I] + '\Word') then
    begin
      Result := Versions[I];
      Exit;
    end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  Office, Ribbon, Backup, Ver, Key: String;
begin
  if CurStep = ssInstall then
  begin
    { Back up an existing Word.officeUI before we overwrite it, so the user's
      own ribbon/QAT can be restored. Only if no backup already exists. }
    Office := ExpandConstant('{userappdata}\Microsoft\Office');
    Ribbon := Office + '\{#RibbonName}';
    Backup := Office + '\Word.officeUI.vistatype-backup';
    if FileExists(Ribbon) and (not FileExists(Backup)) then
      FileCopy(Ribbon, Backup, False);
  end;

  if CurStep = ssPostInstall then
  begin
    { Register the STARTUP folder as a Word Trusted Location so macros run
      without the user lowering Trust Center macro settings. Note: this does
      NOT override enterprise Group Policy or an "add-ins must be signed"
      policy -- code-signing the .dotm is the fix for those. }
    Ver := OfficeVersion();
    if Ver <> '' then
    begin
      Key := 'Software\Microsoft\Office\' + Ver
           + '\Word\Security\Trusted Locations\VistaTypeStartup';
      RegWriteStringValue(HKCU, Key, 'Path', ExpandConstant('{app}\'));
      RegWriteStringValue(HKCU, Key, 'Description', 'VistaType LP STARTUP add-ins');
      RegWriteDWordValue(HKCU, Key, 'AllowSubFolders', 1);
    end;
  end;
end;
