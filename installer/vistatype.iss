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
;      network trusted locations, for roaming/redirected %AppData% profiles).
;      NOT redundant with Word's own STARTUP trusted location: verified 7/30/2026 that
;      with Word's entry deleted and macro security at the default, ours alone is what
;      lets the add-in run. Removed again at uninstall.
;    * Takes LPandBRL.dotm off Word's Disabled Items list, where Word's crash protection
;      puts it after a bad exit and where a reinstall alone cannot reach it
;    * Checks Word is actually INSTALLED, not merely not running -- a Word add-in on a
;      machine with no Word installs perfectly and does nothing
;    * Deletes the obsolete "Large Print Templates" folder
;    * Refuses to run while Word or Outlook is open (with a clear message)
;    * Provides a clean uninstaller
;
;  Build:  compiled by ISCC.exe on the Windows box (see installer/README.md and
;          `make installer`). Source files are staged into ..\dist first.
;
;  Verified on the build box 7/29-30/2026: the Trusted Location keys land correctly and are
;  removed again at uninstall; setup refuses (exit code 1, nothing installed) both while
;  Word is running and while something else holds the .dotm, with a control run proving it
;  still installs when neither is true. Untested: a machine whose %AppData% is redirected to
;  a network share, which is the case AllowNetworkLocations exists for. See installer/README.md.
; ============================================================================

; The Makefile passes /DAppVer= on the ISCC command line. Without this guard the #define
; below would override it, so bumping the Makefile alone would still build an installer
; named with the OLD number - and the Makefile would then fail looking for the new one.
; The literal here is the fallback for building this script by hand, and is kept in step
; with the Makefile by "make bump".
#ifndef AppVer
  #define AppVer      "3.0.509"
#endif
#define DotmName    "LPandBRL.dotm"
#define DotxName    "LargePrintTemplate.dotx"
; The output file's name, defined once because it is now used TWICE: as the name Inno writes,
; and as OriginalFileName inside the .exe's own version block. Two literals would drift, and a
; wrong OriginalFileName is worse than a blank one.
#define OutputBase  "VistaType LP and Braille Macros Setup " + AppVer
; Where a transcriber gets the newest Setup.exe, the guides, and the source. Defined once
; because it appears on the welcome page AND in Programs & Features; releases are published
; here, so this is the address to give anyone asking for an update.
#define RepoUrl     "https://github.com/jerrywhittaker/vistatype-lp"
; A #define FontFamily "VistaTypeLP Legible" stood here until 8/20/2026, and a bundled typeface
; is back from 8/22/2026 -- a different one, which is the point. VistaTypeLP Legible was dropped
; for coverage: no Greek, no IPA, not enough mathematics, and a character the face has not got is
; substituted SILENTLY at another size. VistaTypeLP Sans is Noto Sans given the same ruler
; rescaling, and it carries Greek complete, the phonetic characters, and a slashed zero.
;
; The two do NOT collide, and that was checked rather than assumed. RemoveLegacyLegibleFont below
; still retires the old face on any machine that has it, and it matches on two things: the four
; old filenames in the value's data, and the literal 'vistatypelp legible' in the value's name.
; Neither can reach VistaTypeLPSans-*.ttf or "VistaTypeLP Sans (TrueType)". So an install that
; ships the new face and retires the old one in the same run is safe in either order.
#define FontFamily  "VistaTypeLP Sans"
;
; The OFL text goes in a folder of its OWN, and deliberately NOT the one the legacy license used.
; RemoveLegacyLegibleFont DelTrees "{userappdata}\VistaType LP Fonts" outright once the old face
; is gone -- so putting this license there would have it deleted on the first machine that still
; had Legible on it, leaving a font on disk with no license beside it. That is condition 2 of the
; OFL, and the breach would be invisible. A separate folder costs nothing and cannot be caught by
; that DelTree.
;
; It needs its own folder rather than {userappdata}\VistaType LP for the same reason the legacy
; one did: the fonts install uninsneveruninstall and stay behind, so their license has to outlive
; an uninstall too, and the VistaType LP folder is removed with everything else in it.
#define FontLicenseDir "{userappdata}\VistaType LP Sans Fonts"
;
; The per-user Fonts folder for the REMOVAL of the dropped typeface, written out rather than
; using Inno's {autofonts}. A removal must be able to run on any Windows the add-in supports,
; and {autofonts} resolves to {userfonts}, which Inno documents as Windows 10 version 1803 and
; later. Naming the folder outright means that on a Windows too old to have one it simply does
; not exist, deleting from a folder that is not there is a no-op, and nothing can raise at
; ssPostInstall on an install that has already put LPandBRL.dotm in STARTUP.
;
; The INSTALL of the current face does use {autofonts}, with MinVersion: 10.0.17134 -- see the
; [Files] note. There the version guard and the add-in's own Sh_Is_Font_Installed check are ONE
; mitigation and must not be separated: below 1803 the font does not install, and the add-in
; grays the Typeface choice out and says so rather than offering a face Word cannot find.
#define UserFontsDir "{localappdata}\Microsoft\Windows\Fonts"
; Directory holding the shipping files -- add-in, template, GPL, four font faces and their
; three OFL texts (staged by `make installer`).
#ifndef SrcDir
  #define SrcDir "..\dist"
#endif

[Setup]
AppName=VistaType LP + Braille Macros
AppVersion={#AppVer}
AppPublisher=Jerry Whittaker
AppPublisherURL=mailto:jerry@vistatypelp.org
; These two become the "Support" and "Update" links on the entry in Programs & Features, so a
; transcriber who has lost the download can find it from their own machine.
AppSupportURL={#RepoUrl}
AppUpdatesURL={#RepoUrl}/releases/latest
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
OutputBaseFilename={#OutputBase}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName=VistaType LP + Braille Macros
; Inno 6 hides the welcome page by default. With the folder page and the program-group page
; also off, that left the GPL as the FIRST thing a transcriber saw -- a wall of legal text
; with an "I accept" button, and nothing anywhere saying what they were installing. Turned
; back on 8/8/2026 so the license is the second page and has some context in front of it.
DisableWelcomePage=no

; ----------------------------------------------------------------------------
;  Identity: the icon, the artwork, and the version block inside the .exe
;
;  All added 8/15/2026. Until then the installer set NO icon and NO version numbers, so the
;  built file was a generic Inno Setup stub whose properties read FileVersion 0.0.0.0 with the
;  file-version and original-filename strings blank. That is not cosmetic: "no version
;  information" sits beside "unsigned" as one of the two heavily-weighted factors when one of
;  Microsoft's own tools was hit by the same kind of machine-learning detection that has been
;  deleting this Setup.exe. See docs/Code-Signing.md. A file that looks like a product is less
;  likely to be judged as though it were not one -- and a transcriber who checks Properties
;  before running an installer, which is exactly what a careful one does, now sees who made it.
;
;  branding/ is GENERATED from the artwork in assets/branding/ by tools/lib/build_branding.py
;  (`make branding`, and `make installer` refuses to run if it is missing). Regenerate it;
;  never hand-edit it. The several sizes of each image are so a high-DPI screen gets a sharp
;  one: Inno picks the closest and does not have to stretch it.
; ----------------------------------------------------------------------------
;
;  Jerry, 8/15/2026: there is deliberately NO WizardImageFile. That is the tall panel down the
;  left of the welcome and finished pages, and it carried the mark and the wordmark stacked --
;  both marks on one screen, which looked wrong. The panel is hidden on both pages by
;  InitializeWizard in [Code], which lays the wordmark across the top instead and lets the text
;  run the full width beneath it. The mark still appears top-right on every page in between
;  (WizardSmallImageFile, below) and on the .exe itself.
SetupIconFile=branding\vistatype.ico
WizardSmallImageFile=branding\wizard-small-58.png,branding\wizard-small-77.png,branding\wizard-small-97.png,branding\wizard-small-116.png,branding\wizard-small-124.png,branding\wizard-small-143.png,branding\wizard-small-159.png

; The icon shown beside the entry in Programs & Features. It has to be a file that survives the
; install, so the .ico is installed alongside the scripts in the per-user folder ([Files] below).
UninstallDisplayIcon={userappdata}\VistaType LP\vistatype.ico

; VersionInfoVersion is the one that fills VS_FIXEDFILEINFO, which is where 0.0.0.0 was coming
; from. Inno wants up to four numbers; AppVer is three (3.0.186), which it accepts. The rest were
; being half-filled from AppName/AppPublisher/AppCopyright -- set them outright so nothing is
; left to a default, and so OriginalFileName stops being blank.
VersionInfoVersion={#AppVer}
VersionInfoProductVersion={#AppVer}
VersionInfoTextVersion={#AppVer}
VersionInfoProductName=VistaType LP + Braille Macros
VersionInfoDescription=VistaType LP + Braille Macros Setup
VersionInfoCompany=Jerry Whittaker
VersionInfoCopyright=Copyright (C) 2015-2026 Jerry Whittaker (GNU GPL v3.0)
VersionInfoOriginalFileName={#OutputBase}.exe

; ----------------------------------------------------------------------------
;  Wizard wording. Overrides Inno's stock text; see the Messages section of its
;  help for the full list of names.
;
;  Rules for editing these: they are ONE LINE each -- use %n for a line break and
;  %n%n for a blank line, and write a literal percent sign as two of them. Keep the
;  welcome text short: WizardStyle=modern gives this label a fixed height beside the
;  side image, and anything that overflows is simply not shown, with no warning at
;  compile time. Roughly what is here is the ceiling. Check it on the VM after editing.
; ----------------------------------------------------------------------------
[Messages]
WelcomeLabel1=VistaType LP and Braille Macros

WelcomeLabel2=Version {#AppVer}%n%nJerry Whittaker's tools for transcribers are VBA add-ins for Microsoft Word: large print for readers with low vision, and tools for formatting braille source files for the Duxbury Braille Translator.%n%nAlso installs the {#FontFamily} typeface, which carries its own separate license.%n%nPlease close Word and Outlook before continuing.%n%nLatest version and guides:%n{#RepoUrl}
; Roughly 12 rendered lines against a label that shows about 13-15. It is at the ceiling, so
; anything added here must have something else taken out, and it must be LOOKED AT on the VM.
; Overflow is silently clipped from the bottom - no compile warning - and the bottom is where
; "close Word" and the address are.

; The license page's own heading, which by default asks the user to read an "important"
; agreement before continuing. The GPL governs copying and modifying, not using, so the
; stock wording overstates what a transcriber has to decide here.
LicenseLabel3=VistaType LP is free software under the GNU General Public License v3, shown below in full. Use it for anything, copy it, pass it on. There is no warranty.

; The last thing a transcriber reads. From 8/20/2026 there is nothing left that needs a restart:
; the bundled typeface is gone, and the add-in, the ribbon tabs and the toolbar are all live the
; instant setup finishes. This page said the opposite for ten days, because a per-user font is
; only taken into use at sign-in.
;
; Do NOT put a restart request back here without a reason as concrete as that one was. Asking a
; transcriber to restart mid-book, for nothing she can see, is how an installer earns a
; reputation.
FinishedLabel=Setup has finished installing VistaType LP and Braille Macros.%n%nThe macros, the ribbon tabs and the toolbar are ready to use now. Start Word and look for the VistaType LP and Braille Macros tabs.

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
; The icon Programs & Features draws beside the entry (UninstallDisplayIcon above points here).
; It has to be a file that OUTLIVES the install, which is why it is installed rather than only
; compiled in. Goes with the folder at uninstall, like everything else here.
Source: "branding\vistatype.ico";    DestDir: "{userappdata}\VistaType LP"; Flags: ignoreversion
; The wordmark laid across the top of the welcome and finished pages. dontcopy: these are read
; by the wizard itself and never installed. One per display scaling, drawn at its natural size,
; because Windows' stretch is not smooth and it shows on lettering. BMP because Inno Setup's
; scripting can only load a bitmap. See InitializeWizard in [Code].
Source: "branding\welcome-wordmark-*.bmp";                                 Flags: dontcopy

; ---------------------------------------------------------------------------------------
;  The bundled typeface -- VistaTypeLP Sans, from 8/22/2026
; ---------------------------------------------------------------------------------------
;  Four faces, not one. Word does NOT fall back inside a font family: tested on the build box
;  with a two-face family whose Regular carried U+2264 and whose Bold did not, and the bold one
;  came out of CAMBRIA MATH -- another typeface at another size, silently. Headings are bold, so
;  one face would put the old fault straight back into every heading with a symbol in it.
;
;  MinVersion 10.0.17134 is {autofonts}: Inno documents the per-user Fonts folder as Windows 10
;  version 1803 and later. On anything older these four lines simply do not run, and the face is
;  then not installed -- which the add-in HANDLES rather than ignores: UserForm_Initialize calls
;  Sh_Is_Font_Installed, grays the Typeface choice out and says so on the button. The version
;  guard and that check are ONE mitigation and must not be separated.
;
;  uninsneveruninstall, and this time it means it. Every book set in the face embeds its own
;  copy (Lp_Attach_The_Template sets EmbedTrueTypeFonts with SaveSubsetFonts False), but the
;  books on a machine are not ours to reason about at uninstall time, and a font that leaves
;  takes every unembedded document with it.
Source: "{#SrcDir}\VistaTypeLPSans-Regular.ttf";    DestDir: "{autofonts}"; FontInstall: "{#FontFamily}";             Flags: uninsneveruninstall; MinVersion: 10.0.17134
Source: "{#SrcDir}\VistaTypeLPSans-Bold.ttf";       DestDir: "{autofonts}"; FontInstall: "{#FontFamily} Bold";        Flags: uninsneveruninstall; MinVersion: 10.0.17134
Source: "{#SrcDir}\VistaTypeLPSans-Italic.ttf";     DestDir: "{autofonts}"; FontInstall: "{#FontFamily} Italic";      Flags: uninsneveruninstall; MinVersion: 10.0.17134
Source: "{#SrcDir}\VistaTypeLPSans-BoldItalic.ttf"; DestDir: "{autofonts}"; FontInstall: "{#FontFamily} Bold Italic"; Flags: uninsneveruninstall; MinVersion: 10.0.17134

;  Three licenses, because the face is three fonts. It is Noto Sans with Noto Sans Math and Noto
;  Sans Symbols folded into it, each under its own SIL Open Font License, and condition 2 wants
;  each one's text on disk beside the font. They install uninsneveruninstall for the same reason
;  the faces do -- a license may not leave while the font it covers is still there.
;
;  NOT into {userappdata}\VistaType LP Fonts, which is the legacy folder RemoveLegacyLegibleFont
;  DelTrees outright. Putting them there would delete them on the first machine that still had
;  VistaTypeLP Legible on it, leaving four fonts on disk with no license beside them -- a breach
;  that would be completely invisible. See the #define near the top.
Source: "{#SrcDir}\OFL.txt";                 DestDir: "{#FontLicenseDir}"; Flags: uninsneveruninstall
Source: "{#SrcDir}\OFL-NotoSansMath.txt";    DestDir: "{#FontLicenseDir}"; Flags: uninsneveruninstall
Source: "{#SrcDir}\OFL-NotoSansSymbols.txt"; DestDir: "{#FontLicenseDir}"; Flags: uninsneveruninstall

; ---------------------------------------------------------------------------------------
;  The PREVIOUS bundled typeface -- REMOVED 8/20/2026, and still being removed
; ---------------------------------------------------------------------------------------
;  Four VistaTypeLPLegible-*.ttf files installed to {autofonts} here, with their SIL Open Font
;  License to {userappdata}\VistaType LP Fonts, from 3.0.101 (8/8/2026) to 3.0.196 (8/20/2026).
;
;  Jerry dropped the typeface: its character set is Latin. English, French, German, Spanish and
;  Italian set correctly in it; Latin proper, mathematics, the IPA and the Greek that runs
;  through medical transcription do not, and Word fills a character the face has not got out of
;  another typeface at another size without saying a word. In large print that is the one thing
;  that must never happen.
;
;  It is not enough to stop shipping it -- see RemoveLegacyLegibleFont in [Code], which takes it
;  back off a machine that already has it. This is a deliberate reversal of the
;  uninsneveruninstall rule that used to be written here, and it is safe for one specific
;  reason: Lp_Attach_The_Template EMBEDS the face in every book set in it, unsubsetted, so those
;  books carry their own copy and go on setting and printing correctly with the font gone.
;
;  If a bundled typeface is ever considered again, the test to apply first is character
;  COVERAGE, before metrics, before legibility: Greek, the IPA, and the mathematical operators.
;  A face that fails it fails silently, which is why this one shipped at all.
;
;  That test was applied to VistaTypeLP Sans on 8/22/2026 and it FAILED the third leg on the
;  first attempt -- Greek and the IPA complete, but 3 of 15 common math characters, no arrows,
;  and none of the operators. Google's Noto Sans comes from the latin-greek-cyrillic project and
;  is scoped to those; mathematics lives in Noto Sans Math, a separate face. That is why the
;  build folds Noto Sans Math and Noto Sans Symbols in, and why the required-character list in
;  tools/lib/build_vistatypelp_sans.py now carries the operators, the arrows and the letterlike
;  symbols: so the BUILD fails on a face that cannot set them, instead of this note being
;  something someone has to remember to act on. It nearly was not acted on this time.

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
; The wording deliberately names BOTH halves of the product. The eight icons on this toolbar
; are Lp_, Dx_ and Sh_ macros - it serves large print and braille alike - so calling it
; "VistaType" alone reads as if braille users need not bother. (Jerry, 7/30/2026.)
Name: "qat";          GroupDescription: "Quick Access Toolbar (the small row of icons at the top of the Word window):"; \
                      Description: "Set up my Quick Access Toolbar for VistaType LP and Braille Macros"
; Order and default set by Jerry, 7/30/2026, and his reasoning is the point:
;   "most people do not use a custom QAT and with a click-through they get what they need
;    to produce large print and braille."
; So the common case has nothing of the user's to protect, and clicking straight through
; Next should leave a transcriber with a working setup rather than a bare toolbar. The
; people who HAVE built their own toolbar are the ones who read the options - and for them
; the second choice keeps it untouched. This is not a return to the pre-3.0.33 behavior
; that caused the complaints: that one merged silently, hid buttons they had chosen to keep,
; and offered no way back. This asks, saves what they had, and can restore it.
; With `exclusive`, the entry WITHOUT `unchecked` is the one selected.
Name: "qat\vista";    Description: "Replace my toolbar with the standard VistaType LP and Braille Macros toolbar (mine is saved, and I can get it back)"; \
                      Flags: exclusive
Name: "qat\mine";     Description: "Keep my toolbar exactly as it is, and add the VistaType LP and Braille Macros icons on the end"; \
                      Flags: exclusive unchecked
; Only offered when we actually hold a saved copy, i.e. an earlier VistaType install
; rewrote their toolbar. checkedonce so a later upgrade stops pestering them about it.
Name: "qat\restore";  Description: "Put back the toolbar I had before VistaType LP and Braille Macros was installed, then add its icons"; \
                      Flags: exclusive unchecked checkedonce; Check: HasQatBackup

; There was a "Ribbon tabs:" task here until 8/9/2026. It is gone, and the tabs are now always
; written into the user's own Word.officeUI. Jerry's call, and his reasoning: leaving it
; UNCHECKED is what makes Word behave abnormally, because Word does not list add-in tabs in
; Customize the Ribbon at all - so a transcriber who declined ended up with two tabs they could
; not hide, rename or reorder like every other tab they have. Offering that as a choice invited
; people to pick the odd one out. Nobody can end up with no tabs either way: the copies embedded
; in the add-in stay as the fallback and getVisible switches them off only once the user has
; their own. See DEVELOPMENT.md, "Embedded ribbon".

[Run]
; At most one of the first three runs -- the qat tasks are mutually exclusive, and none runs
; if the user left that parent unchecked. Each also writes the ribbon tabs, so the toolbar and
; the tabs are set up in a single pass. The tabs are not optional (8/9/2026); the fourth entry
; below covers the case where the toolbar is left alone and none of these three fires.
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -NoProfile -File ""{userappdata}\VistaType LP\Merge-Qat.ps1"" -Mode Mine -Tabs Install -FullTemplate ""{userappdata}\VistaType LP\qat-template.officeUI"" -IconsTemplate ""{userappdata}\VistaType LP\qat-icons-only.officeUI"" -TabsTemplate ""{userappdata}\VistaType LP\ribbon-tabs.officeUI"""; \
  Tasks: qat\mine; Flags: runhidden; StatusMsg: "Setting up your Quick Access Toolbar and ribbon..."
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -NoProfile -File ""{userappdata}\VistaType LP\Merge-Qat.ps1"" -Mode Vista -Tabs Install -FullTemplate ""{userappdata}\VistaType LP\qat-template.officeUI"" -IconsTemplate ""{userappdata}\VistaType LP\qat-icons-only.officeUI"" -TabsTemplate ""{userappdata}\VistaType LP\ribbon-tabs.officeUI"""; \
  Tasks: qat\vista; Flags: runhidden; StatusMsg: "Setting up your Quick Access Toolbar and ribbon..."
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -NoProfile -File ""{userappdata}\VistaType LP\Merge-Qat.ps1"" -Mode Restore -Tabs Install -FullTemplate ""{userappdata}\VistaType LP\qat-template.officeUI"" -IconsTemplate ""{userappdata}\VistaType LP\qat-icons-only.officeUI"" -TabsTemplate ""{userappdata}\VistaType LP\ribbon-tabs.officeUI"""; \
  Tasks: qat\restore; Flags: runhidden; StatusMsg: "Putting back your original Quick Access Toolbar..."
; ...and this one covers "leave my toolbar alone", where none of the three above fires. The tabs
; still have to be written, so this run does them on their own.
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -NoProfile -File ""{userappdata}\VistaType LP\Merge-Qat.ps1"" -Mode None -Tabs Install -TabsTemplate ""{userappdata}\VistaType LP\ribbon-tabs.officeUI"""; \
  Check: ToolbarUntouched; Flags: runhidden; StatusMsg: "Putting the VistaType tabs on your ribbon..."

[UninstallRun]
; Remove only VistaType's QAT icons, leaving the user's own ribbon/QAT intact.
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -NoProfile -File ""{userappdata}\VistaType LP\Remove-Qat.ps1"""; \
  Flags: runhidden; RunOnceId: "RemoveVistaTypeQat"

[InstallDelete]
; The bundled typeface, dropped 8/20/2026 -- see the note in [Files] and RemoveLegacyLegibleFont
; in [Code]. The files go here, at the START of the install; the registry entries that make
; Windows load them go in that procedure at the end, because Inno has no [Registry] flag that
; can find a value by its DATA and these were written with Inno's own FontInstall naming, which
; this installer no longer performs and so cannot assume.
;
; A file Windows has already loaded for this session CANNOT be deleted, and that is expected:
; Windows opens a per-user font at sign-in and holds it for the whole session. Inno ignores an
; [InstallDelete] it cannot carry out, which is the behavior wanted here - the unregistering
; below is what actually retires the font, and the file goes on the next run or is left as an
; orphan nothing points at. Never make this fatal.
Type: files; Name: "{#UserFontsDir}\VistaTypeLPLegible-Regular.ttf"
Type: files; Name: "{#UserFontsDir}\VistaTypeLPLegible-Bold.ttf"
Type: files; Name: "{#UserFontsDir}\VistaTypeLPLegible-Italic.ttf"
Type: files; Name: "{#UserFontsDir}\VistaTypeLPLegible-BoldItalic.ttf"
; Remove the obsolete template folder from older versions.
Type: filesandordirs; Name: "{userappdata}\Microsoft\Templates\Large Print Templates"
; Remove a stale uninstaller left in the STARTUP folder by pre-fix installers (it now lives
; in {userappdata}\VistaType LP). {app} is the STARTUP folder here.
Type: files; Name: "{app}\unins000.exe"
Type: files; Name: "{app}\unins000.dat"

[UninstallDelete]
Type: files;          Name: "{app}\{#DotmName}"
Type: files;          Name: "{userappdata}\Microsoft\Templates\{#DotxName}"
; The bundled typeface used to be excluded here on purpose, so that an uninstall could not
; reflow the transcriber's finished books. That whole arrangement is gone: the typeface was
; dropped from the product on 8/20/2026 and is removed at INSTALL time instead, by
; RemoveLegacyLegibleFont in [Code]. Nothing here installs a font any more, so there is nothing
; for this section to protect.
; NOT listed, and never add it: {userappdata}\VistaType LP Settings, which holds the
; transcriber's own Word settings that VistaType puts back for her (the spelling and grammar
; ones braille switches off, and whatever the settings library grows to hold). It is a separate
; folder from "VistaType LP" for exactly this reason - that one is wiped whole by the line
; below, and a reinstall must not cost her her settings. Same reasoning as the font license
; folder above. 8/18/2026.
Type: filesandordirs; Name: "{userappdata}\VistaType LP"

[Code]
const
  WORD_CLASS    = 'OpusApp';           { Word main window class }
  OUTLOOK_CLASS = 'rctrl_renwnd32';    { Outlook main window class }
  { The wordmark's width in pixels at 100% display scaling. Everything else on the welcome and
    finished pages is positioned from it. }
  WORDMARK_WIDTH = 260;

{ --- The welcome and finished pages: the wordmark above full-width text -------------------

  Jerry, 8/15/2026. Inno draws a tall image panel down the left of these two pages, and it was
  carrying the mark and the wordmark stacked one above the other. Two marks on one screen looked
  wrong, and the name only needs saying once.

  So the panel is hidden on both pages and the wordmark goes across the top instead, with the
  heading and body text moved left and widened to the full page. The mark has not gone: it is
  still top-right on every page in between, and it is the icon on the .exe itself.

  There is no WizardImageFile in [Setup] at all, which is why hiding the panel leaves nothing
  behind it.

  The image is loaded here rather than by Inno because Inno only puts it in that one place.
  Six widths ship, one per display scaling from 100% to 250%, and the nearest one at or above
  what is wanted is drawn at its NATURAL size -- Stretch is left off deliberately, because
  Windows' stretch is not smooth and it shows badly on lettering. They are BMP because Inno
  Setup's scripting can only load a bitmap, and flattened onto white to match the page.
  Generated by tools/lib/build_branding.py; see assets/branding/README.md.                  --- }
function WordmarkFile(): String;
var
  Widths: array[0..5] of Integer;
  Want, I, Pick: Integer;
begin
  Widths[0] := 260; Widths[1] := 325; Widths[2] := 390;
  Widths[3] := 455; Widths[4] := 520; Widths[5] := 650;
  Want := ScaleX(WORDMARK_WIDTH);
  { Smallest width that is not smaller than we want, so it is never scaled UP; the largest is
    the fallback for a display scaled beyond anything shipped. }
  Pick := 5;
  for I := 0 to 5 do
    if Widths[I] >= Want then
    begin
      Pick := I;
      Break;
    end;
  Result := 'welcome-wordmark-' + IntToStr(Widths[Pick]) + '.bmp';
end;

procedure LayOutBrandedPage(Page: TNewNotebookPage; Panel: TBitmapImage;
                            Heading: TNewStaticText; Body: TNewStaticText;
                            const FileName: String);
var
  Mark: TBitmapImage;
  Margin: Integer;
begin
  Panel.Visible := False;
  Margin := ScaleX(24);

  Mark := TBitmapImage.Create(WizardForm);
  Mark.Parent := Page;
  Mark.Bitmap.LoadFromFile(ExpandConstant('{tmp}\') + FileName);
  Mark.Stretch := False;
  Mark.Left := Margin;
  Mark.Top := ScaleY(24);
  Mark.Width := Mark.Bitmap.Width;
  Mark.Height := Mark.Bitmap.Height;

  Heading.Left := Margin;
  Heading.Top := Mark.Top + Mark.Height + ScaleY(18);
  Heading.Width := Page.Width - Margin * 2;

  Body.Left := Margin;
  Body.Top := Heading.Top + Heading.Height + ScaleY(10);
  Body.Width := Page.Width - Margin * 2;
  Body.Height := Page.Height - Body.Top - ScaleY(16);
end;

procedure InitializeWizard();
var
  FileName: String;
begin
  FileName := WordmarkFile();
  ExtractTemporaryFile(FileName);
  LayOutBrandedPage(WizardForm.WelcomePage, WizardForm.WizardBitmapImage,
                    WizardForm.WelcomeLabel1, WizardForm.WelcomeLabel2, FileName);
  LayOutBrandedPage(WizardForm.FinishedPage, WizardForm.WizardBitmapImage2,
                    WizardForm.FinishedHeadingLabel, WizardForm.FinishedLabel, FileName);
end;

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

{ --- TabsArg lived here until 8/9/2026, deciding whether to pass Install or Skip for -Tabs.
      The tabs are no longer optional, so the [Run] entries pass Install outright and the
      function is gone. Merge-Qat.ps1 still ACCEPTS -Tabs Skip; nothing in the installer asks
      for it any more.

      NOTE for whoever writes the next one: braces are Pascal's comment delimiters, so an
      inline code reference must never be written out inside a comment -- the closing brace
      ends the comment early and the rest of the sentence is compiled. That broke the 3.0.34
      build first time round. --- }

{ --- True when the user left the toolbar alone, so none of the three qat entries above will
      fire. That is the one case where the tabs still need writing on their own run. --- }
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
function WordIsRunning(): Boolean;
begin
  Result := IsProcessRunning('WINWORD.EXE') or (FindWindowByClassName(WORD_CLASS) <> 0);
end;

function OutlookIsRunning(): Boolean;
begin
  Result := IsProcessRunning('OUTLOOK.EXE') or (FindWindowByClassName(OUTLOOK_CLASS) <> 0);
end;

function OfficeIsRunning(): Boolean;
begin
  Result := WordIsRunning() or OutlookIsRunning();
end;

{ --- NAME the one that is actually in the way. The message used to say "Word or Outlook",
      which sent Jerry hunting for an Outlook that was not installed on that machine while a
      windowless WINWORD.EXE sat there holding the file (8/10/2026). A transcriber cannot act
      on "one of these two". --- }
function WhatIsRunning(): String;
begin
  if WordIsRunning() and OutlookIsRunning() then
    Result := 'Microsoft Word and Microsoft Outlook are still running.'
  else if OutlookIsRunning() then
    Result := 'Microsoft Outlook is still running.'
  else
    Result := 'Microsoft Word is still running.';
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
    { A function result cannot sit next to #13#10 the way a literal can - Pascal needs the
      plus signs. That is what broke this compile first time round. }
    SuppressibleMsgBox(WhatIsRunning() + #13#10 + #13#10
      + 'Close every Word and Outlook window and try again.' #13#10 #13#10
      + 'If you have already closed them, one of them may still be running with no window '
      + 'on screen - left behind by a crash, or started in the background. Sign out of '
      + 'Windows and back in, then run this again. (Or end WINWORD.EXE in Task Manager, on '
      + 'the Details tab.)' #13#10 #13#10
      + 'While it is running the add-in cannot be ' + Verb + '.',
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
{ --- Registry helpers that see BOTH views.
      A 32-bit installer on 64-bit Windows has its plain HKLM reads redirected into
      WOW6432Node, and 64-bit Office writes InstallRoot only to the 64-bit view - so a
      plain HKLM check misses it entirely. --- }
function RegKeyExistsEitherView(const SubKey: String): Boolean;
begin
  Result := RegKeyExists(HKLM, SubKey);
  if (not Result) and IsWin64 then
    Result := RegKeyExists(HKLM64, SubKey);
end;

function RegStringEitherView(const SubKey, ValueName: String; var Value: String): Boolean;
begin
  Result := RegQueryStringValue(HKLM, SubKey, ValueName, Value);
  if (not Result) and IsWin64 then
    Result := RegQueryStringValue(HKLM64, SubKey, ValueName, Value);
end;

{ --- Is Word INSTALLED? (Not "has Word been run", and not "is Word running".)
      Uses signals written by Office SETUP, so it is true even on a profile where Word has
      never been opened. --- }
function IsWordInstalled(): Boolean;
var
  Path: String;
begin
  Result := False;
  if RegStringEitherView('SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe',
                         '', Path) then
    if (Path <> '') and FileExists(RemoveQuotes(Path)) then
    begin
      Result := True;
      Exit;
    end;
  { Fallbacks: the COM registration, then any Office InstallRoot. }
  if RegKeyExistsEitherView('SOFTWARE\Classes\Word.Application\CurVer') then Result := True;
  if (not Result) and RegKeyExistsEitherView('SOFTWARE\Microsoft\Office\16.0\Word\InstallRoot') then Result := True;
  if (not Result) and RegKeyExistsEitherView('SOFTWARE\Microsoft\Office\15.0\Word\InstallRoot') then Result := True;
  if (not Result) and RegKeyExistsEitherView('SOFTWARE\Microsoft\Office\14.0\Word\InstallRoot') then Result := True;
end;

{ --- Which Office version's settings do we write to?
      HKLM FIRST, because Office setup writes InstallRoot there. The old version read only
      HKCU, which Word does not create until a user opens it for the first time - so someone
      who installed Office, then VistaType, then opened Word got NO trusted location at all,
      silently, and could not run the macros they had just installed. --- }
function OfficeVersion(): String;
var
  Versions: TArrayOfString;
  I: Integer;
  CurVer: String;
begin
  Result := '';
  Versions := ['16.0', '15.0', '14.0'];

  for I := 0 to GetArrayLength(Versions) - 1 do
    if RegKeyExistsEitherView('SOFTWARE\Microsoft\Office\' + Versions[I] + '\Word\InstallRoot') then
    begin
      Result := Versions[I];
      Exit;
    end;

  for I := 0 to GetArrayLength(Versions) - 1 do
    if RegKeyExists(HKCU, 'Software\Microsoft\Office\' + Versions[I] + '\Word') then
    begin
      Result := Versions[I];
      Exit;
    end;

  { e.g. "Word.Application.16" -> "16.0" }
  if RegStringEitherView('SOFTWARE\Classes\Word.Application\CurVer', '', CurVer) then
    if Pos('.16', CurVer) > 0 then Result := '16.0'
    else if Pos('.15', CurVer) > 0 then Result := '15.0'
    else if Pos('.14', CurVer) > 0 then Result := '14.0';

  { Word is here but unusually registered. Writing 16.0's key is better than writing
    nothing: a key Word may read beats a trusted location that certainly does not exist. }
  if (Result = '') and IsWordInstalled() then
    Result := '16.0';
end;

function InitializeSetup(): Boolean;
begin
  Result := OfficeIsClear('replaced');
  if not Result then Exit;

  { A Word add-in on a machine with no Word installs perfectly and does nothing at all -
    the folders get created, every file lands, and the user is told it succeeded. Say so
    instead. Not a hard block: Office can be registered in ways this does not recognize
    (containerised or MSIX installs), and refusing a legitimate install is worse than a
    no-op. Default is No, so an unattended deployment onto a machine without Word stops
    rather than silently doing nothing. }
  if not IsWordInstalled() then
    Result := SuppressibleMsgBox(
        'Microsoft Word does not appear to be installed on this computer.' #13#10 #13#10
      + 'VistaType is an add-in for Word: on its own it does nothing. If Word is not '
      + 'installed, install it first and then run this again.' #13#10 #13#10
      + 'Continue anyway?',
      mbConfirmation, MB_YESNO, IDNO) = IDYES;
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


{ --- Take the dropped VistaTypeLP Legible typeface back off the machine. ---------------

  VistaType LP bundled this face from 3.0.101 (8/8/2026) to 3.0.196 (8/20/2026) and no longer
  does. Not shipping it is only half the job: it was installed with uninsneveruninstall, so on
  every machine that ever ran one of those builds it is still there, still registered, and still
  the face Word offers in every font list.

  Removing a font a user's own documents might depend on is normally the wrong thing, and this
  file said so at length. It is right here because Lp_Attach_The_Template embeds the face,
  unsubsetted, in every book set in it -- those books carry their own copy and go on setting and
  printing correctly with the font gone. That is the whole reason Jerry could drop it.

  Two halves, and the REGISTRY half is the one that matters. Windows opens a per-user font at
  sign-in and holds the file for the session, so the .ttf usually cannot be deleted during an
  install and [InstallDelete] quietly fails on it. Unregistering always works, and at the next
  sign-in Windows simply does not load the face. The leftover file is then an orphan nothing
  points at, and the next run of this installer deletes it.

  So: nothing here is fatal, and nothing here reports failure to the transcriber. A font that
  survives one more session is not worth a scary dialog on an install that otherwise worked. }
procedure RemoveLegacyLegibleFont();
var
  FontsKey, FontDir, Data, LowerData, LowerName: String;
  Names: TArrayOfString;
  Files: TArrayOfString;
  I, J, Unregistered, Remaining: Integer;
begin
  FontsKey := 'Software\Microsoft\Windows NT\CurrentVersion\Fonts';
  FontDir  := ExpandConstant('{#UserFontsDir}');

  SetArrayLength(Files, 4);
  Files[0] := 'VistaTypeLPLegible-Regular.ttf';
  Files[1] := 'VistaTypeLPLegible-Bold.ttf';
  Files[2] := 'VistaTypeLPLegible-Italic.ttf';
  Files[3] := 'VistaTypeLPLegible-BoldItalic.ttf';

  Unregistered := 0;

  { Matched on the value's DATA (the file it points at) as well as its name. The name Inno wrote
    was "VistaTypeLP Legible (TrueType)" and friends, but that is Inno's own convention for a
    FontInstall this installer no longer performs -- and a machine could have been touched by a
    hand-install too. The filename is the thing that is genuinely ours. }
  if RegGetValueNames(HKCU, FontsKey, Names) then
  begin
    for I := 0 to GetArrayLength(Names) - 1 do
    begin
      LowerName := Lowercase(Names[I]);
      Data := '';
      RegQueryStringValue(HKCU, FontsKey, Names[I], Data);
      LowerData := Lowercase(Data);

      for J := 0 to GetArrayLength(Files) - 1 do
      begin
        if (Pos(Lowercase(Files[J]), LowerData) > 0) or
           (Pos('vistatypelp legible', LowerName) > 0) then
        begin
          if RegDeleteValue(HKCU, FontsKey, Names[I]) then
          begin
            Log('Unregistered dropped typeface: ' + Names[I]);
            Unregistered := Unregistered + 1;
          end;
          Break;
        end;
      end;
    end;
  end;

  { Tried again here as well as in [InstallDelete], because unregistering can be what releases
    it on a machine where the font was added mid-session and never loaded at sign-in. }
  for J := 0 to GetArrayLength(Files) - 1 do
  begin
    if FileExists(FontDir + '\' + Files[J]) then
      if not DeleteFile(FontDir + '\' + Files[J]) then
        Log('Dropped typeface still in use, left for the next run: ' + Files[J]);
  end;

  { Counted by ASKING, not by adding up what this run managed to do. The difference matters, and
    it is the OFL condition below that makes it matter: a font hand-installed somewhere else --
    machine-wide under C:\Windows\Fonts, say -- leaves all four of these absent, and tallying
    successes would read that as "all removed" and delete the license text of a font that is
    still on the machine. Which is the exact breach the guard exists to prevent. }
  Remaining := 0;
  for J := 0 to GetArrayLength(Files) - 1 do
    if FileExists(FontDir + '\' + Files[J]) then
      Remaining := Remaining + 1;

  { The OFL text lived in its OWN folder so that it could outlive an uninstall alongside a font
    that never left. The font is leaving, so the license goes with it -- but ONLY once no font
    file of ours is left in that folder AND we have actually retired something. While a font file
    is on disk, its license has to be on disk beside it; that is condition 2 of the OFL, and it
    does not stop applying because we have decided to stop shipping the face.

    RegDeleteValue and DeleteFile between them are what proves this machine ever had our copy. On
    a clean install nothing is unregistered and nothing is deleted, so the folder is left alone --
    it will not be there anyway, and a DelTree of a folder we have no evidence about is not a
    thing to do on someone else's machine. }
  if (Remaining = 0) and (Unregistered > 0) then
    DelTree(ExpandConstant('{userappdata}\VistaType LP Fonts'), True, True, True)
  else if Remaining > 0 then
    Log('Dropped typeface not fully removed yet; leaving its license folder in place.');

  if (Unregistered > 0) or (Remaining > 0) then
    Log('VistaTypeLP Legible retired: ' + IntToStr(Unregistered) + ' registry entries removed, '
        + IntToStr(Remaining) + ' of 4 files still present. Takes effect at the next sign-in.');
end;


{ --- Take SUPERSEDED copies of VistaTypeLP Sans off the machine. -----------------------

  A stray VistaTypeLPSans-*.ttf under a name that is not one of the four we install is a
  loaded gun, and the reason is the family name inside it. Every one of them says
  "VistaTypeLP Sans". Right-click one and choose Install -- which is exactly what a person
  does with a stray .ttf found in their Fonts folder -- and Windows rewrites the
  "VistaTypeLP Sans (TrueType)" registration to point at it. From that moment the family
  Word offers is whatever that old file happened to hold.

  Measured on the build box, 8/24/2026, on the two that were sitting there: 3,094 characters
  against the shipping 6,450. Three of fourteen math operators. No arrows at all. And no
  italic anywhere in the family, so Word slants the roman by machine and a slanted rescaled
  face is the wrong size. That is the fault that got VistaTypeLP Legible dropped on
  8/20/2026, wearing the right name.

  WHERE THEY COME FROM. Not from this installer, as far as anything shows: it replaces all
  four faces on every run, and it did so on a box where Windows had them locked -- a plain
  copy over the top failed with a sharing violation and Setup replaced them anyway, with no
  leftover. The path that does produce one is a HAND install, which is what the .zip built by
  make font-installer tells a tester to do: right-click the .ttf and choose Install. Install
  a font whose filename is already in the per-user Fonts folder and Windows keeps the old one
  under a timestamped name. That mechanism was NOT reproduced here -- the shell Install verb
  does nothing over SSH with no desktop, and the file is only locked after a sign-in -- so
  read it as the likeliest source and not a proven one. The two files were real either way.

  ORDER IS THE ONE THING THAT CAN GO WRONG, and it is why this runs at ssPostInstall. By the
  time it does, the four FontInstall entries have already rewritten the four registrations to
  the canonical paths. Run it first instead, on a machine whose family currently points at a
  stray, and deleting that file would leave the typeface broken until the install caught up.

  Unregister first, delete second -- the same order as RemoveLegacyLegibleFont and for the
  same two reasons: unregistering can be what releases a file Windows is holding, and a
  registration left pointing at a file that is gone is a broken entry in every font list.

  Matched on the value's DATA rather than its name, again as above. The name is whatever the
  font's own full name said when someone hand-installed it, and that is not ours to predict.

  The four canonical filenames are a WHITELIST, and they are the guard. Anything this walk
  can see other than those four is superseded by definition.

  Per-user only. PrivilegesRequired is lowest, so HKLM and the machine-wide Fonts folder are
  out of reach; a copy hand-installed "for all users" is beyond this and is left alone.

  Nothing here is fatal and nothing is reported to the transcriber. A stray that survives one
  more session is not worth a dialog on an install that otherwise worked. }
procedure RemoveSupersededSansFaces();
var
  FontsKey, FontDir, Data, LowerData, LowerName: String;
  Names, Strays: TArrayOfString;
  FindRec: TFindRec;
  I, J, N, Removed, Unregistered: Integer;
  Canonical: Boolean;
begin
  FontDir  := ExpandConstant('{#UserFontsDir}');
  FontsKey := 'Software\Microsoft\Windows NT\CurrentVersion\Fonts';

  { Collected first and acted on second. Deleting inside a FindFirst/FindNext walk is how an
    entry gets skipped. }
  N := 0;
  SetArrayLength(Strays, 0);
  if FindFirst(FontDir + '\VistaTypeLPSans-*.ttf', FindRec) then
  begin
    try
      repeat
        if FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY = 0 then
        begin
          LowerName := Lowercase(FindRec.Name);
          Canonical := (LowerName = 'vistatypelpsans-regular.ttf')    or
                       (LowerName = 'vistatypelpsans-bold.ttf')       or
                       (LowerName = 'vistatypelpsans-italic.ttf')     or
                       (LowerName = 'vistatypelpsans-bolditalic.ttf');
          if not Canonical then
          begin
            SetArrayLength(Strays, N + 1);
            Strays[N] := FindRec.Name;
            N := N + 1;
          end;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;

  if N = 0 then
    Exit;

  Unregistered := 0;
  if RegGetValueNames(HKCU, FontsKey, Names) then
  begin
    for I := 0 to GetArrayLength(Names) - 1 do
    begin
      Data := '';
      RegQueryStringValue(HKCU, FontsKey, Names[I], Data);
      LowerData := Lowercase(Data);

      for J := 0 to N - 1 do
      begin
        { A stray's filename carries the decoration, so it can never match the data of a
          canonical registration -- the canonical path is the shorter string. }
        if Pos(Lowercase(Strays[J]), LowerData) > 0 then
        begin
          if RegDeleteValue(HKCU, FontsKey, Names[I]) then
          begin
            Log('Unregistered superseded VistaTypeLP Sans copy: ' + Names[I] + ' = ' + Data);
            Unregistered := Unregistered + 1;
          end;
          Break;
        end;
      end;
    end;
  end;

  Removed := 0;
  for J := 0 to N - 1 do
  begin
    if DeleteFile(FontDir + '\' + Strays[J]) then
      Removed := Removed + 1
    else
      Log('Superseded VistaTypeLP Sans copy still in use, left for the next run: ' + Strays[J]);
  end;

  Log('Superseded VistaTypeLP Sans copies: ' + IntToStr(N) + ' found, ' + IntToStr(Removed)
      + ' deleted, ' + IntToStr(Unregistered) + ' registry entries removed.');
end;


{ --- Take LPandBRL.dotm back off Word's Disabled Items list.
      When Word exits badly with the add-in loaded - a crash, or the machine reset under it -
      its next start can offer "Word is running into problems with the lpandbrl.dotm add-in.
      Do you want to disable it now?" A Yes puts the add-in on Resiliency\DisabledItems, and
      from then on it never loads: the VistaType tabs still show, because they live in the
      user's Word.officeUI, but every button in them points into the add-in and none can draw.
      Reinstalling did NOT cure it (9/26/2026, 3.0.509 on the build box) - the entry
      outlives the file - and each time Word then saved its ribbon settings with the add-in
      still off, it stripped every VistaType button out of Word.officeUI.

      Word must be closed for this, and setup refuses to start while it is running (see
      InitializeSetup), so Word cannot rewrite the list or the ribbon settings behind us.
      That also means the order against the [Run] entries that rewrite the tabs does not
      matter: Word reads both fresh at its next start.

      Each entry is a binary blob with the add-in's full path inside it as UTF-16. The null
      bytes are stripped and the rest lowercased, and ONLY an entry naming lpandbrl.dotm is
      deleted. Anything else on the list belongs to another add-in and is left alone. --- }
procedure EnableDisabledAddIn();
var
  Versions, Names: TArrayOfString;
  I, J: Integer;
  Key, Text: String;
  Blob: AnsiString;
begin
  Versions := ['16.0', '15.0', '14.0'];
  for I := 0 to GetArrayLength(Versions) - 1 do
  begin
    Key := 'Software\Microsoft\Office\' + Versions[I] + '\Word\Resiliency\DisabledItems';
    if RegGetValueNames(HKCU, Key, Names) then
      for J := 0 to GetArrayLength(Names) - 1 do
        if RegQueryBinaryValue(HKCU, Key, Names[J], Blob) then
        begin
          Text := Blob;
          StringChangeEx(Text, #0, '', True);
          if Pos('lpandbrl.dotm', Lowercase(Text)) > 0 then
          begin
            if RegDeleteValue(HKCU, Key, Names[J]) then
              Log('Took LPandBRL.dotm off Word ' + Versions[I] + '''s Disabled Items list (' + Names[J] + ').')
            else
              Log('Could not take LPandBRL.dotm off Word ' + Versions[I] + '''s Disabled Items list (' + Names[J] + ').');
          end;
        end;
  end;
end;


procedure CurStepChanged(CurStep: TSetupStep);
var
  Ver, Key, TL: String;
  Existing: Cardinal;
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
      TL := 'Software\Microsoft\Office\' + Ver + '\Word\Security\Trusted Locations';

      { Allow trusted locations that resolve to a network path. Many institutional
        users have roaming/redirected %AppData%, so the STARTUP folder is a network
        location; without this, Word ignores the trusted location below.

        This one is a machine-wide security setting rather than something of ours, so
        record whether WE are the ones turning it on. Uninstall then undoes it only if we
        did - never if the user or another add-in had already set it. }
      if not RegQueryDWordValue(HKCU, TL, 'AllowNetworkLocations', Existing) then
        Existing := 0;
      if Existing <> 1 then
        RegWriteStringValue(HKCU, 'Software\VistaType LP', 'WeSetAllowNetworkLocations', '1');
      RegWriteDWordValue(HKCU, TL, 'AllowNetworkLocations', 1);

      { Ours, named after us, and removed at uninstall. Word ships its own trusted location
        for this same folder, so on a default machine this is belt-and-braces - but it is
        NOT redundant: verified 7/30/2026 that with Word's own entry deleted and macro
        security at the default, this key alone is what lets the add-in run. }
      Key := TL + '\VistaTypeStartup';
      RegWriteStringValue(HKCU, Key, 'Path', ExpandConstant('{app}\'));
      RegWriteStringValue(HKCU, Key, 'Description', 'VistaType LP STARTUP add-ins');
      RegWriteDWordValue(HKCU, Key, 'AllowSubFolders', 1);

      { Remember where we wrote, so uninstall does not have to guess the version again. }
      RegWriteStringValue(HKCU, 'Software\VistaType LP', 'OfficeVersion', Ver);
    end;

    { An add-in Word has switched off stays off through any number of reinstalls. }
    EnableDisabledAddIn();

    { Last, and deliberately after everything that makes the add-in work: an upgrade must leave
      a working VistaType LP behind even if the font clean-up hits something unexpected.

      Both must also run AFTER the FontInstall entries in [Files] have rewritten the four
      registrations to the canonical paths, which is exactly what ssPostInstall guarantees.
      RemoveSupersededSansFaces says at length why that order is the thing that can go wrong. }
    RemoveLegacyLegibleFont();
    RemoveSupersededSansFaces();
  end;
end;

{ --- Take the Trust Center entries back out. The old uninstaller left both behind
      permanently: a trusted location named after us, and a security setting we had
      switched on. Neither is ours to keep once VistaType is gone. --- }
procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  Ver, TL, Flag: String;
begin
  if CurUninstallStep = usUninstall then
  begin
    if not RegQueryStringValue(HKCU, 'Software\VistaType LP', 'OfficeVersion', Ver) then
      Ver := OfficeVersion();
    if Ver <> '' then
    begin
      TL := 'Software\Microsoft\Office\' + Ver + '\Word\Security\Trusted Locations';
      RegDeleteKeyIncludingSubkeys(HKCU, TL + '\VistaTypeStartup');

      { Only if we were the ones who turned it on. }
      if RegQueryStringValue(HKCU, 'Software\VistaType LP', 'WeSetAllowNetworkLocations', Flag) then
        if Flag = '1' then
          RegDeleteValue(HKCU, TL, 'AllowNetworkLocations');
    end;
    RegDeleteKeyIncludingSubkeys(HKCU, 'Software\VistaType LP');

    { And again here, because the normal case leaves work behind. On the first upgrade the
      registry entries go but the four .ttf files usually cannot -- Windows holds a per-user font
      open for the whole session. If the transcriber's next move is to UNINSTALL rather than to
      install again, nothing would ever come back for them, and four orphaned files plus a license
      folder would outlive the product that put them there. }
    RemoveLegacyLegibleFont();

    { RemoveSupersededSansFaces is deliberately NOT called here. The four current faces stay
      at uninstall on purpose -- uninsneveruninstall, because a font that leaves takes every
      unembedded document with it -- so there is no superseded/current split to reconcile, and
      a sweep at this point would only be deciding what to delete from a Fonts folder we have
      just promised to leave alone. }
  end;
end;
