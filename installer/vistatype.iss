; ============================================================================
;  VistaType LP + Braille Macros -- Inno Setup installer
; ============================================================================
;  Replaces the manual "copy three files to three locations" procedure.
;  Per-user install (no admin needed); everything lands in the user's profile.
;
;  The ribbon is now EMBEDDED in LPandBRL.dotm (customUI14.xml), so there is no
;  Word.officeUI file to install and the user's own ribbon/QAT is never touched.
;
;  What it does automatically (each item is a support call it removes):
;    * Copies LPandBRL.dotm            -> %AppData%\Microsoft\Word\STARTUP
;                                         (creates STARTUP if it doesn't exist)
;    * Copies LargePrintTemplate.dotx  -> %AppData%\Microsoft\Templates
;    * Merges VistaType's QAT icons INTO the user's own Word.officeUI
;      (Merge-Qat.ps1) -- pre-stocks the quick-access toolbar without touching
;      the user's ribbon or their own QAT items; uninstall removes only ours
;    * Registers the STARTUP folder as a Word Trusted Location (and allows
;      network trusted locations, for roaming/redirected %AppData% profiles)
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

; The Makefile passes /DAppVer= on the ISCC command line. Without this guard the #define
; below would override it, so bumping the Makefile alone would still build an installer
; named with the OLD number - and the Makefile would then fail looking for the new one.
; The literal here is the fallback for building this script by hand, and is kept in step
; with the Makefile by "make bump".
#ifndef AppVer
  #define AppVer      "3.0.15"
#endif
#define DotmName    "LPandBRL.dotm"
#define DotxName    "LargePrintTemplate.dotx"
; Directory holding the three shipping files (staged by `make installer`).
#ifndef SrcDir
  #define SrcDir "..\dist"
#endif

[Setup]
AppName=VistaType LP + Braille Macros
AppVersion={#AppVer}
AppPublisher=Jerry Whittaker
AppPublisherURL=mailto:jerry@thewhittakers.org
AppCopyright=Copyright (C) 2015-2026 Jerry Whittaker (GNU GPL v3.0)
; Show the GPLv3 during install. (GPL governs copying/modifying, not mere use, so this
; page is informational; switch to InfoBeforeFile if you'd rather not require "I accept".)
LicenseFile={#SrcDir}\LICENSE.txt
DefaultDirName={userappdata}\Microsoft\Word\STARTUP
; {app} is Word's STARTUP folder, which Word scans for templates -- keep the uninstaller
; (unins000.exe/.dat) OUT of it, in our own per-user folder instead.
UninstallFilesDir={userappdata}\VistaType LP
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
; Macro add-in (with embedded ribbon) -> STARTUP. {app} == the STARTUP folder here.
Source: "{#SrcDir}\{#DotmName}";   DestDir: "{app}";                             Flags: ignoreversion
; Large-print styles template -> user Templates folder.
Source: "{#SrcDir}\{#DotxName}";   DestDir: "{userappdata}\Microsoft\Templates"; Flags: ignoreversion
; QAT merge helpers -> a persistent per-user folder (needed again at uninstall).
Source: "scripts\Merge-Qat.ps1";   DestDir: "{userappdata}\VistaType LP"; Flags: ignoreversion
Source: "scripts\Remove-Qat.ps1";  DestDir: "{userappdata}\VistaType LP"; Flags: ignoreversion
Source: "qat-template.officeUI";   DestDir: "{userappdata}\VistaType LP"; Flags: ignoreversion
; Install a copy of the GPL so the user "receives a copy of the license" per the GPL.
Source: "{#SrcDir}\LICENSE.txt";   DestDir: "{userappdata}\VistaType LP"; Flags: ignoreversion

[Run]
; Install VistaType's standard QAT (preserving the user's own icons + ribbon customizations).
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -NoProfile -File ""{userappdata}\VistaType LP\Merge-Qat.ps1"" -Template ""{userappdata}\VistaType LP\qat-template.officeUI"""; \
  Flags: runhidden; StatusMsg: "Adding VistaType quick-access icons..."

[UninstallRun]
; Remove only VistaType's QAT icons, leaving the user's own ribbon/QAT intact.
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -NoProfile -File ""{userappdata}\VistaType LP\Remove-Qat.ps1"""; \
  Flags: runhidden; RunOnceId: "RemoveVistaTypeQat"

[InstallDelete]
; Remove the obsolete template folder from older versions.
Type: filesandordirs; Name: "{userappdata}\Microsoft\Templates\Large Print Templates"
; Remove a stale uninstaller left in the STARTUP folder by pre-fix installers (it now lives
; in {userappdata}\VistaType LP). {app} is the STARTUP folder here.
Type: files; Name: "{app}\unins000.exe"
Type: files; Name: "{app}\unins000.dat"

[UninstallDelete]
Type: files;          Name: "{app}\{#DotmName}"
Type: files;          Name: "{userappdata}\Microsoft\Templates\{#DotxName}"
Type: filesandordirs; Name: "{userappdata}\VistaType LP"

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
  Ver, Key: String;
begin
  if CurStep = ssPostInstall then
  begin
    { Register the STARTUP folder as a Word Trusted Location so macros run
      without the user lowering Trust Center macro settings. Note: this does
      NOT override enterprise Group Policy or an "add-ins must be signed"
      policy -- code-signing the .dotm is the fix for those. }
    Ver := OfficeVersion();
    if Ver <> '' then
    begin
      { Allow trusted locations that resolve to a network path. Many institutional
        users have roaming/redirected %AppData%, so the STARTUP folder is a network
        location; without this, Word ignores the trusted location below. }
      RegWriteDWordValue(HKCU, 'Software\Microsoft\Office\' + Ver
        + '\Word\Security\Trusted Locations', 'AllowNetworkLocations', 1);

      Key := 'Software\Microsoft\Office\' + Ver
           + '\Word\Security\Trusted Locations\VistaTypeStartup';
      RegWriteStringValue(HKCU, Key, 'Path', ExpandConstant('{app}\'));
      RegWriteStringValue(HKCU, Key, 'Description', 'VistaType LP STARTUP add-ins');
      RegWriteDWordValue(HKCU, Key, 'AllowSubFolders', 1);
    end;
  end;
end;
