; ============================================================================
;  VistaType LP + Braille Macros -- Inno Setup installer
; ============================================================================
;  Replaces the manual "copy three files to three locations" procedure.
;  Per-user install (no admin needed); everything lands in the user's profile.
;
;  The ribbon is EMBEDDED in LPandBRL.dotm (customUI14.xml), so it merges with the
;  user's ribbon rather than replacing it, and no Word.officeUI ships.
;
;  The Quick Access Toolbar is the user's CHOICE, made on the wizard (see [Tasks]).
;  Until 3.0.32 the VistaType toolbar was merged in unconditionally, which pushed the
;  user's own icons to the right, discarded their separators, and hid buttons they had
;  deliberately kept -- Undo among them, because suppressing Word's defaults is the only
;  way to make a curated toolbar look clean. That merge is gone. Now: keep theirs and
;  append ours, or install ours whole (theirs saved and restorable), or leave it alone.
;
;  What it does automatically (each item is a support call it removes):
;    * Copies LPandBRL.dotm            -> %AppData%\Microsoft\Word\STARTUP
;                                         (creates STARTUP if it doesn't exist)
;    * Copies LargePrintTemplate.dotx  -> %AppData%\Microsoft\Templates
;    * Sets up the Quick Access Toolbar the way the user asked (Merge-Qat.ps1), always
;      saving their original first; uninstall removes only what we put there
;    * Registers the STARTUP folder as a Word Trusted Location (and allows
;      network trusted locations, for roaming/redirected %AppData% profiles)
;    * Deletes the obsolete "Large Print Templates" folder
;    * Refuses to run while Word or Outlook is open (with a clear message)
;    * Provides a clean uninstaller
;
;  Build:  compiled by ISCC.exe on the Windows box (see installer/README.md and
;          `make installer`). Source files are staged into ..\dist first.
;
;  Verified on the build box 7/29/2026: the Trusted Location keys land correctly, and
;  setup refuses (exit code 1, nothing installed) both while Word is running and while
;  something else holds the .dotm -- with a control run proving it still installs when
;  neither is true. Untested: a machine whose %AppData% is redirected to a network share,
;  which is the case AllowNetworkLocations exists for.
; ============================================================================

; The Makefile passes /DAppVer= on the ISCC command line. Without this guard the #define
; below would override it, so bumping the Makefile alone would still build an installer
; named with the OLD number - and the Makefile would then fail looking for the new one.
; The literal here is the fallback for building this script by hand, and is kept in step
; with the Makefile by "make bump".
#ifndef AppVer
  #define AppVer      "3.0.36"
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
; Jerry, 7/27/2026. NOTE: AppName above is deliberately NOT changed to match. Inno
; derives the product identity from AppName when no AppId is set, so renaming it would
; make an upgrade look like a different product and leave a second entry behind in
; Programs & Features. Renaming only the output file has no such effect.
OutputBaseFilename=VistaType LP and Braille Macros Setup {#AppVer}
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
Source: "scripts\Merge-Qat.ps1";     DestDir: "{userappdata}\VistaType LP"; Flags: ignoreversion
Source: "scripts\Remove-Qat.ps1";    DestDir: "{userappdata}\VistaType LP"; Flags: ignoreversion
; Both toolbars ship whatever the user chose: the curated one, and the append-safe icons.
; They are tiny, and shipping both means the Switch Toolbar button in Word can move between
; them later without needing the installer again.
Source: "qat-template.officeUI";     DestDir: "{userappdata}\VistaType LP"; Flags: ignoreversion
Source: "qat-icons-only.officeUI";   DestDir: "{userappdata}\VistaType LP"; Flags: ignoreversion
; The two VistaType ribbon tabs, in the form Word treats as the user's own custom tabs so
; that Customize the Ribbon lists them. Generated from the ribbon by tools/lib/build_ribbon_tabs.py.
Source: "ribbon-tabs.officeUI";      DestDir: "{userappdata}\VistaType LP"; Flags: ignoreversion
; Install a copy of the GPL so the user "receives a copy of the license" per the GPL.
Source: "{#SrcDir}\LICENSE.txt";   DestDir: "{userappdata}\VistaType LP"; Flags: ignoreversion

[Tasks]
; Deliberately [Tasks] and not [Components]: components choose which FILES get installed,
; and all three choices ship identical files -- only a script argument differs. It would
; also force a "Select Components" page showing "0.0 MB" beside each option, which reads as
; broken. Tasks additionally get remembered: Inno stores the chosen tasks in the per-user
; uninstall key and preselects them next time, so an upgrade keeps the user's decision
; without asking again. (That memory is keyed off AppName -- see the note in [Setup].)
;
; Leaving the parent UNCHECKED means "do not touch my toolbar at all", which is only safe
; because the six VistaType macros are also reachable by keyboard shortcut and from the
; VistaType LP ribbon tab.
Name: "qat";          GroupDescription: "Quick Access Toolbar (the small row of icons at the top of the Word window):"; \
                      Description: "Set up my Quick Access Toolbar for VistaType"
Name: "qat\mine";     Description: "Keep my toolbar exactly as it is, and add VistaType's icons on the end"; \
                      Flags: exclusive
Name: "qat\vista";    Description: "Replace my toolbar with VistaType's standard one (mine is saved, and I can get it back)"; \
                      Flags: exclusive unchecked
; Only offered when we actually hold a saved copy, i.e. an earlier VistaType install
; rewrote their toolbar. checkedonce so a later upgrade stops pestering them about it.
Name: "qat\restore";  Description: "Put back the toolbar I had before VistaType was installed, then add VistaType's icons"; \
                      Flags: exclusive unchecked checkedonce; Check: HasQatBackup

; Separate top-level task, NOT a child of qat: those three are Flags: exclusive, and a
; non-exclusive fourth sibling would confuse both Inno and the reader. Someone can perfectly
; well want the tabs and not the toolbar, or the other way round.
;
; Unchecked leaves the tabs exactly as they were before 3.0.34 -- supplied by the add-in,
; and not listed in Customize the Ribbon. So nobody can end up with no tabs at all.
Name: "ribbontabs"; GroupDescription: "Ribbon tabs:"; \
                    Description: "Put the VistaType tabs on my ribbon, so I can hide, rename and reorder them"

[Run]
; At most one of the first three runs -- the qat tasks are mutually exclusive, and none runs
; if the user left that parent unchecked. Each also carries the ribbon-tab decision, computed
; by TabsArg below, so the toolbar and the tabs are set up in a single pass.
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -NoProfile -File ""{userappdata}\VistaType LP\Merge-Qat.ps1"" -Mode Mine -Tabs {code:TabsArg} -FullTemplate ""{userappdata}\VistaType LP\qat-template.officeUI"" -IconsTemplate ""{userappdata}\VistaType LP\qat-icons-only.officeUI"" -TabsTemplate ""{userappdata}\VistaType LP\ribbon-tabs.officeUI"""; \
  Tasks: qat\mine; Flags: runhidden; StatusMsg: "Setting up your Quick Access Toolbar and ribbon..."
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -NoProfile -File ""{userappdata}\VistaType LP\Merge-Qat.ps1"" -Mode Vista -Tabs {code:TabsArg} -FullTemplate ""{userappdata}\VistaType LP\qat-template.officeUI"" -IconsTemplate ""{userappdata}\VistaType LP\qat-icons-only.officeUI"" -TabsTemplate ""{userappdata}\VistaType LP\ribbon-tabs.officeUI"""; \
  Tasks: qat\vista; Flags: runhidden; StatusMsg: "Setting up your Quick Access Toolbar and ribbon..."
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -NoProfile -File ""{userappdata}\VistaType LP\Merge-Qat.ps1"" -Mode Restore -Tabs {code:TabsArg} -FullTemplate ""{userappdata}\VistaType LP\qat-template.officeUI"" -IconsTemplate ""{userappdata}\VistaType LP\qat-icons-only.officeUI"" -TabsTemplate ""{userappdata}\VistaType LP\ribbon-tabs.officeUI"""; \
  Tasks: qat\restore; Flags: runhidden; StatusMsg: "Putting back your original Quick Access Toolbar..."
; ...and this one covers "leave my toolbar alone, but I do want the tabs", where none of the
; three above fires.
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -NoProfile -File ""{userappdata}\VistaType LP\Merge-Qat.ps1"" -Mode None -Tabs Install -TabsTemplate ""{userappdata}\VistaType LP\ribbon-tabs.officeUI"""; \
  Tasks: ribbontabs; Check: ToolbarUntouched; Flags: runhidden; StatusMsg: "Putting the VistaType tabs on your ribbon..."

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

{ --- Do we hold a saved copy of the toolbar this user had before VistaType touched it?
      Merge-Qat.ps1 writes <file>.vtqatbak once and never overwrites it, so on a machine
      that has had VistaType for years this is still their true original. A ZERO-length
      file is the sentinel meaning "there was no toolbar file before us" - there is nothing
      to put back in that case, so it must not count.

      Gates the "put back the toolbar I had" task: offering it on a machine with nothing
      saved would promise a rescue we cannot perform. --- }
function NonEmptyFile(const Path: String): Boolean;
var
  Size: Integer;
begin
  Result := False;
  if FileExists(Path) then
    if FileSize(Path, Size) then
      Result := Size > 0;
end;

function HasQatBackup(): Boolean;
begin
  { Both locations: Word reads Roaming on most machines, Local when the profile roams. }
  Result := NonEmptyFile(ExpandConstant('{userappdata}\Microsoft\Office\Word.officeUI.vtqatbak'));
  if not Result then
    Result := NonEmptyFile(ExpandConstant('{localappdata}\Microsoft\Office\Word.officeUI.vtqatbak'));
end;

{ --- What to pass Merge-Qat for -Tabs. Called inline from the Parameters of the [Run]
      entries, so the toolbar and the ribbon are set up by one run of the script rather than
      two. "Skip" leaves the ribbon exactly as it was: tabs supplied by the add-in, and not
      listed in Customize the Ribbon.

      NOTE: braces are Pascal's comment delimiters, so an inline code reference must never be
      written out inside one -- the closing brace ends the comment early and the rest of the
      sentence is compiled. That is what broke the 3.0.34 build first time round. --- }
function TabsArg(Param: String): String;
begin
  if WizardIsTaskSelected('ribbontabs') then
    Result := 'Install'
  else
    Result := 'Skip';
end;

{ --- True when the user left the toolbar alone, so none of the three qat entries above will
      fire. That is the one case where the tabs still need installing on their own. --- }
function ToolbarUntouched(): Boolean;
begin
  Result := not WizardIsTaskSelected('qat');
end;

{ --- Is a process running? Checked by NAME, not by window: Word can be running with no
      window at all - left behind by a crashed session, or started by Outlook as its mail
      editor - and it still holds the STARTUP add-in locked. The old window-only test walked
      straight past that case. Falls back to False if WMI is unavailable, so a locked-down
      machine degrades to the window test below rather than blocking the install. --- }
function IsProcessRunning(const ExeName: String): Boolean;
var
  Locator, WMI, Objs: Variant;
begin
  Result := False;
  try
    { Inno's Pascal Script cannot call a method on a function result - the object has to
      land in a variable first, or the compiler reports "Unknown identifier". }
    Locator := CreateOleObject('WbemScripting.SWbemLocator');
    WMI     := Locator.ConnectServer('.', 'root\CIMV2');
    Objs    := WMI.ExecQuery('SELECT Name FROM Win32_Process WHERE Name = "' + ExeName + '"');
    Result  := (Objs.Count > 0);
  except
    Result := False;
  end;
end;

{ --- Can we actually replace this file? Renaming it is the definitive test: if the rename
      succeeds nothing has it open, and we put it straight back. This catches holders the
      process test cannot name - a backup agent, antivirus, or the search indexer. --- }
function FileIsInUse(const FileName: String): Boolean;
var
  TempName: String;
begin
  Result := False;
  if not FileExists(FileName) then
    Exit;
  TempName := FileName + '.vtlocktest';
  if RenameFile(FileName, TempName) then
    RenameFile(TempName, FileName)
  else
    Result := True;
end;

{ --- Is Word or Outlook running, by process OR by window? --- }
function OfficeIsRunning(): Boolean;
begin
  Result := IsProcessRunning('WINWORD.EXE') or IsProcessRunning('OUTLOOK.EXE')
         or (FindWindowByClassName(WORD_CLASS) <> 0)
         or (FindWindowByClassName(OUTLOOK_CLASS) <> 0);
end;

{ --- Shared by install AND uninstall, so the two can never drift apart. Both need to
      touch the same STARTUP .dotm, and Word holds it locked whenever it is running.
      Verb is "installed" or "removed" so each path can word its own message. --- }
function OfficeIsClear(const Verb: String): Boolean;
var
  Dotm: String;
begin
  Result := True;
  Dotm := ExpandConstant('{userappdata}\Microsoft\Word\STARTUP\{#DotmName}');

  { SuppressibleMsgBox, not MsgBox. /SUPPRESSMSGBOXES has no effect on a plain MsgBox, so a
    silent or unattended install (an IT department deploying this) would sit forever waiting
    for a click nobody can give, instead of failing cleanly. Found 7/29/2026 by running the
    guard for real: the installer hung rather than aborting. Suppressed, it returns Default
    and we go on to return False, which is exactly the behavior wanted. }
  if OfficeIsRunning() then
  begin
    SuppressibleMsgBox('Microsoft Word or Outlook is still running.' #13#10 #13#10
      + 'Close every Word and Outlook window and try again. If you have already closed '
      + 'them, Word may still be running in the background - sign out of Windows and back '
      + 'in, then run this again.' #13#10 #13#10
      + 'While Word is running the add-in cannot be ' + Verb + '.',
      mbError, MB_OK, IDOK);
    Result := False;
    Exit;
  end;

  if FileIsInUse(Dotm) then
  begin
    SuppressibleMsgBox('The VistaType add-in file is being used by another program, so it '
      + 'cannot be ' + Verb + ':' #13#10 #13#10 + Dotm + #13#10 #13#10
      + 'This is usually antivirus, a backup program, or Windows Search reading the file. '
      + 'Wait a moment and try again, or restart the computer.',
      mbError, MB_OK, IDOK);
    Result := False;
  end;
end;

{ --- Abort the install if anything is holding the files we need to replace. --- }
function InitializeSetup(): Boolean;
begin
  Result := OfficeIsClear('replaced');
end;

{ --- Same guard on the way out. Uninstalling with Word open leaves the .dotm behind,
      because Word has it locked - so the add-in appears to have been removed while the
      old code carries on loading at every Word start. Remove-Qat.ps1 also rewrites the
      user's Word.officeUI here, which is not safe while Word has it loaded.
      (Jerry, 7/27/2026.) --- }
function InitializeUninstall(): Boolean;
begin
  Result := OfficeIsClear('removed');
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
