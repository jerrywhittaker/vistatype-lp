# VistaType LP installer

Replaces the manual "copy three files to three locations" procedure with a single
per-user `Setup.exe`. No admin rights required.

## What it automates (each was a recurring support call)

| Manual step in the old doc | Installer does it |
|---|---|
| Copy `LPandBRL.dotm` to `STARTUP` (and create STARTUP if missing) | ✔ automatic |
| Copy `LargePrintTemplate.dotx` to Templates | ✔ automatic |
| Copy `Word.officeUI` to the Office folder | ✔ automatic, **backs up the existing one** |
| "Macros won't run" → Trust Center fiddling | ✔ registers STARTUP as a **Trusted Location** |
| Delete obsolete "Large Print Templates" folder | ✔ automatic |
| "Close Word and Outlook first" | ✔ refuses to run if they're open |
| (none — was impossible by hand) | ✔ clean **Uninstall** entry |

## Build

From Linux: `make installer` — stages the three shipping files into `dist/`
(renaming the built `.dotm` to `LPandBRL.dotm`), pushes to the Windows box, runs
`ISCC.exe`, and copies `dist/VistaType-LP-Setup-<ver>.exe` back.

Prereq on the Windows box: [Inno Setup 6](https://jrsoftware.org/isdl.php) installed
(free). Adjust `WIN_ISCC` in `build.config` if it's not at the default path.

To compile by hand on Windows instead: open `installer\vistatype.iss` in the Inno
Setup IDE with the three files present in `..\dist`, and click Compile.

## Not yet validated

The Trusted-Location registry write and the Word/Outlook running-check are correct in
principle but haven't been run on a real machine. Test on a clean Windows profile
before shipping. Known limits the installer **cannot** fix (they need code-signing or
IT involvement): enterprise Group Policy that disables trusted locations, or a policy
requiring add-ins to be signed by a trusted publisher.
