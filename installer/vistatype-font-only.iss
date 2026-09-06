; ============================================================================================
;  VistaTypeLP Sans - font only
;
;  A standalone installer that puts the typeface on a machine and nothing else. Built for a beta
;  tester who needs the font without the add-in (Jerry, 8/23/2026). It is not part of `make
;  installer` and is not shipped to transcribers; build it by hand when someone needs it:
;
;     make font-installer
;
;  NO ADMINISTRATIVE RIGHTS ARE NEEDED. PrivilegesRequired=lowest, so {autofonts} resolves to the
;  user's OWN Fonts folder, %LOCALAPPDATA%\Microsoft\Windows\Fonts, and the registration goes
;  under HKCU. Exactly what the main installer does. It follows that the font is installed for
;  THAT USER only - another account on the same PC will not see it.
;
;  ANTIVIRUS. This file is deliberately plain: no [Code] section, no registry work of its own, no
;  elevation, no network. It carries a full version block and an icon, because "no version
;  information" sits beside "unsigned" as one of the two heavily-weighted factors behind the
;  machine-learning deletions that ate every build of the main Setup.exe on 8/13/2026 - see
;  docs/Code-Signing.md. What it CANNOT be is signed, because the certificate is not in use yet.
;  So it can still be flagged, and there is no way to promise otherwise. The .zip built beside it
;  is the answer to that: a font installed by right-click has no executable in it at all.
;
;  A DIFFERENT AppId FROM THE MAIN PRODUCT, and it must stay different. Inno identifies a product
;  by AppId; share one and installing or removing either would be treated as an upgrade or a
;  removal of the other.
;
;  THE THREE LICENSES TRAVEL WITH THE FONT and are not optional. The face is Noto Sans with Noto
;  Sans Math and Noto Sans Symbols folded in, each under its own SIL Open Font License, and
;  condition 2 wants each one's text on disk beside the font. Same folder the main installer uses
;  - and deliberately NOT %AppData%\VistaType LP Fonts, which the main product's uninstaller
;  deletes whole (it is the legacy folder from the dropped typeface).
;
;  UNLIKE THE MAIN INSTALLER, uninstalling this one DOES take the font away. There it is
;  uninsneveruninstall, because a font that leaves takes every document that did not embed it
;  with it. Here the person has the font for testing and should be able to give it back.
; ============================================================================================

#define FontFamily   "VistaTypeLP Sans"
#define FontVer      "1.1"
#define OutputBase   "VistaTypeLP-Sans-Font-Setup-" + FontVer

; Where the four .ttf files and the three OFL texts are staged. `make installer` puts them in
; dist/; this is overridden on the command line the same way.
#ifndef SrcDir
  #define SrcDir "..\dist"
#endif

; The same folder the main installer uses for these, so a machine that has both does not end up
; with two copies of the same three licenses.
#define FontLicenseDir "{userappdata}\VistaType LP Sans Fonts"

[Setup]
AppId={{9E3B4C21-77A5-4F0E-9C2D-1B6A5F0E7D34}
AppName={#FontFamily} font
AppVersion={#FontVer}
AppPublisher=Jerry Whittaker
AppPublisherURL=mailto:jerry@vistatypelp.org
AppCopyright=VistaTypeLP Sans is built from Noto Sans and is licensed under the SIL Open Font License 1.1

; The font's own license, shown before it is installed. The OFL governs copying and modifying
; rather than mere use, so this page is informational - but a font is somebody else's work and
; the terms should be in front of the person taking it.
LicenseFile={#SrcDir}\OFL.txt

; Nothing is installed into {app}; it exists only to hold the uninstaller.
DefaultDirName={userappdata}\VistaTypeLP Sans Font
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
DisableWelcomePage=no

; Per-user font installation is Windows 10 version 1803 and later. Refuse outright on anything
; older rather than run and install nothing, which would look like success.
MinVersion=10.0.17134

OutputDir=..\dist
OutputBaseFilename={#OutputBase}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayName={#FontFamily} font

SetupIconFile=branding\vistatype.ico
WizardSmallImageFile=branding\wizard-small-58.png,branding\wizard-small-77.png,branding\wizard-small-97.png,branding\wizard-small-116.png,branding\wizard-small-124.png,branding\wizard-small-143.png,branding\wizard-small-159.png

; The version block. See the note at the top: a file that looks like a product is less likely to
; be judged as though it were not one, and anyone careful enough to check Properties before
; running an installer sees who made it.
VersionInfoVersion={#FontVer}
VersionInfoProductVersion={#FontVer}
VersionInfoTextVersion={#FontVer}
VersionInfoProductName={#FontFamily} font
VersionInfoDescription={#FontFamily} font installer
VersionInfoCompany=Jerry Whittaker
VersionInfoCopyright=SIL Open Font License 1.1
VersionInfoOriginalFileName={#OutputBase}.exe

[Messages]
WelcomeLabel1={#FontFamily}

WelcomeLabel2=This installs the {#FontFamily} typeface and nothing else.%n%nIt does not need administrator rights: the font is installed for you alone, and no other account on this computer will see it.%n%nClose Word before continuing, or the new font will not appear in it until you restart Word.%n%nThe typeface is built from Noto Sans and carries the SIL Open Font License, shown on the next page.

FinishedLabelNoIcons=The {#FontFamily} typeface is installed.%n%nIf Word was open while it installed, close it and open it again before looking for the font.

[Files]
; Four faces, not one. Word does NOT fall back inside a font family - a bold character the Bold
; face has not got comes out of some other typeface at another size, silently - so all four go on
; or the emphasized words in a book set in this face come out at the wrong size.
Source: "{#SrcDir}\VistaTypeLPSans-Regular.ttf";    DestDir: "{autofonts}"; FontInstall: "{#FontFamily}"
Source: "{#SrcDir}\VistaTypeLPSans-Bold.ttf";       DestDir: "{autofonts}"; FontInstall: "{#FontFamily} Bold"
Source: "{#SrcDir}\VistaTypeLPSans-Italic.ttf";     DestDir: "{autofonts}"; FontInstall: "{#FontFamily} Italic"
Source: "{#SrcDir}\VistaTypeLPSans-BoldItalic.ttf"; DestDir: "{autofonts}"; FontInstall: "{#FontFamily} Bold Italic"

; The licenses, beside the font. Not optional - see the note at the top.
Source: "{#SrcDir}\OFL.txt";                 DestDir: "{#FontLicenseDir}"
Source: "{#SrcDir}\OFL-NotoSansMath.txt";    DestDir: "{#FontLicenseDir}"
Source: "{#SrcDir}\OFL-NotoSansSymbols.txt"; DestDir: "{#FontLicenseDir}"

[UninstallDelete]
; The folder itself, once its three files have gone with the uninstall above.
Type: dirifempty; Name: "{#FontLicenseDir}"
