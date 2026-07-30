# VistaType LP installer

Replaces the manual "copy three files to three locations" procedure with a single
per-user `Setup.exe`. No admin rights required.

> This file is developer-facing (how the installer is built and what it does). The
> **end-user install guide** is [`../docs/Installation-Guide.md`](../docs/Installation-Guide.md).

## What it automates (each was a recurring support call)

| Manual step in the old doc | Installer does it |
|---|---|
| Copy `LPandBRL.dotm` to `STARTUP` (and create STARTUP if missing) | ✔ automatic |
| Copy `LargePrintTemplate.dotx` to Templates | ✔ automatic |
| Copy `Word.officeUI`, export/import your QAT first | ✘ **gone** — ribbon is embedded in the `.dotm` and merges non-destructively |
| Pre-stock the QAT quick-access icons | ✔ **asks** (`[Tasks]`): append VistaType's icons to the user's toolbar (default), install VistaType's toolbar whole, restore their pre-VistaType one, or leave it alone. Their ribbon is untouched either way |
| "Macros won't run" → Trust Center fiddling | ✔ registers STARTUP as a **Trusted Location** + allows **network** trusted locations (roaming `%AppData%`) |
| Delete obsolete "Large Print Templates" folder | ✔ automatic |
| "Close Word and Outlook first" | ✔ refuses to run if they're open |
| Show the license / give the user a copy | ✔ displays **GPLv3** during setup and installs a `LICENSE.txt` copy |
| (none — was impossible by hand) | ✔ clean **Uninstall** entry (removes only our QAT icons) |

The installer ships **two** files (`.dotm` + `.dotx`), not three. The toolbar is set up by
`scripts/Merge-Qat.ps1` at install time (`-Mode Mine|Vista|Restore`, chosen by the user in
`[Tasks]`) and taken back off by `scripts/Remove-Qat.ps1` at uninstall.

Two toolbar files ship, both to `%AppData%\VistaType LP`:

| File | Maintained how |
|---|---|
| `qat-template.officeUI` | **By hand.** VistaType's curated toolbar. Its ordering, separators, and the `visible="false"` entries that suppress Word's own default buttons are all deliberate and cannot be derived from anything. |
| `qat-icons-only.officeUI` | **Generated** by `tools/lib/build_qat.py` from the hidden `tab_LP_and_BRL_QAT_Icons` ribbon tab. VistaType's six icons and nothing else — no `mso:` entries, no `visible="false"` — so appending it to someone's own toolbar cannot hide or move what they have. |
| `ribbon-tabs.officeUI` | **Generated** by `tools/lib/build_ribbon_tabs.py` from the two visible ribbon tabs. Written into the user's own `Word.officeUI` by `-Tabs Install`, so Word lists the tabs in *Customize the Ribbon* and they can be hidden, reordered and renamed — which is impossible for tabs defined inside the add-in. Buttons are emitted as references (`idQ="x1:btn_..."`), so clicks still go through `RibbonAction` and the labels and icons are not duplicated. |
| `ribbon-button-ids.txt` | **Append-only record**, maintained by the same script. Every id that has shipped. Renaming or deleting a button breaks every tab and toolbar already installed in the field, which reference buttons by id — the entry renders blank with no error. `make build` stops if a shipped id disappears. |

The ribbon-tab install has one rule worth stating plainly, because it is what makes an
upgrade safe: **the contents of our tabs are ours; the tab's place, name and visibility are
the user's.** A re-install replaces only the `vt_grp_*` groups inside a VistaType tab and
leaves the tab element — position, label, `visible` — untouched, along with any group the
user added to it themselves.

`make build` runs `build_qat.py`, which also **validates** the curated file against that
ribbon tab and fails the build if they disagree. A stale `idQ` reference renders as a blank
button on the user's machine with nothing to warn you. (`tools/lib/extract_qat.py` and the
old `qat-controls.xml` are obsolete and no longer part of the pipeline.)

## Build

From Linux: `make installer` — stages the shipping files into `dist/`
(`LPandBRL.dotm`, `LargePrintTemplate.dotx`, `LICENSE.txt`), pushes to the Windows box,
runs `ISCC.exe`, and copies `"VistaType LP and Braille Macros Setup <ver>.exe"` back to local `dist/` **and**
onto the build box's Desktop (ready to double-click for a test install on the VM).

Prereq on the Windows box: [Inno Setup 6](https://jrsoftware.org/isdl.php) installed
(free). Adjust `WIN_ISCC` in `build.config` if it's not at the default path.

To compile by hand on Windows instead: open `installer\vistatype.iss` in the Inno
Setup IDE with the three files present in `..\dist`, and click Compile.

## Validation status

**The toolbar scripts now run for real.** As of 3.0.33 they are exercised on the Windows
build box by a test harness (`Test-Qat.ps1`, kept with the session scratch rather than in
the repo) that drives `Merge-Qat.ps1` and `Remove-Qat.ps1` against sample `Word.officeUI`
files via their `-TargetPaths` parameter, so the box's own Word setup is never touched.
45 checks cover: all three install choices; a heavily customized toolbar (own icons, own
separators, reordered built-ins, plus a custom ribbon tab that must survive); a machine
with no `Word.officeUI` at all; running an install twice; and six uninstall situations,
including uninstalling after the user added icons of their own, and an install done by an
older version that left no manifest. A malformed `Word.officeUI` must not abort either
script, and is tested too.

That harness caught two real defects before shipping: an illegal `--` inside an XML
comment in the generated toolbar file, which made the whole "keep my toolbar" path fail
silently; and an uninstall that did not restore the user's original toolbar if they had
added an icon since installing.

As of 3.0.35 the suite is **75 checks**, adding the ribbon tabs: installing them beside a
user's own custom tab and a hidden built-in; `-Mode None -Tabs Install` (tabs without
touching the toolbar); a re-install after the user moved, renamed and unticked one of our
tabs; uninstall; uninstall when there is no toolbar section at all; and a VistaType tab the
user has added their own group to, which must survive.

**The install guards were exercised for real on 7/29/2026** and both work: setup exits with
code 1 and installs nothing while Word is running, and likewise while another process holds
`LPandBRL.dotm` — with a control run confirming it still installs when neither is true. The
Trusted Location keys land with the right path, description and `AllowSubFolders`, and
`AllowNetworkLocations=1` is set on the parent.

That test found a real defect: the guards used Inno's plain `MsgBox`, which
**`/SUPPRESSMSGBOXES` does not suppress** — only `SuppressibleMsgBox` does. A silent or
unattended install therefore *hung* waiting for a click instead of failing cleanly. Fixed;
if you add another guard, use `SuppressibleMsgBox` with an explicit default.

**The Trusted Location is not redundant — proven 7/30/2026.** Word ships its own trusted
location for the STARTUP folder, so the suspicion was that ours merely duplicated it. Jerry
tested it the hard way: macro security back at Word's default (`VBAWarnings=2`) *and* Word's
own STARTUP entry deleted. VistaType still loaded and ran with no security warning. The only
thing trusting that folder was our key. Keep it.

**Three defects that test found, all fixed in 3.0.37:**

1. **The installer never checked whether Word was installed at all** — only whether it was
   *running*. On a machine with no Word every file lands, every folder is created, and the
   user is told it succeeded. `IsWordInstalled()` now asks, and setup offers to stop. It is a
   Yes/No rather than a hard block, because Office can be registered in ways this will not
   recognise (containerised or MSIX), and refusing a legitimate install is worse than a
   no-op. The default is No, so an unattended deployment onto a machine without Word stops.

2. **`OfficeVersion()` read `HKCU`, which Word does not create until a user first opens it.**
   So: install Office, install VistaType, *then* open Word for the first time — and the whole
   Trusted Location block was silently skipped. No trusted location, no
   `AllowNetworkLocations`, macros possibly blocked, and nothing to explain why. It now reads
   `HKLM\...\Office\<ver>\Word\InstallRoot` first, which **Office setup** writes, and falls
   back through HKCU, the `Word.Application` COM registration, and finally `16.0` if Word is
   demonstrably present but oddly registered. Both HKLM views are checked: a 32-bit installer
   on 64-bit Windows has plain `HKLM` reads redirected to `WOW6432Node`, and 64-bit Office
   writes `InstallRoot` only to the 64-bit view.

3. **Uninstall left the Trust Center entries behind permanently.** It now removes
   `VistaTypeStartup` and the `HKCU\Software\VistaType LP` bookkeeping key.
   `AllowNetworkLocations` is a machine-wide security setting rather than something of ours,
   so it is undone **only if we were the ones who turned it on** — recorded at install time
   as `WeSetAllowNetworkLocations`. If the user or another add-in had already set it, it
   stays. (Machines upgraded from 3.0.36 or earlier have no such record, so theirs is left
   alone — the conservative direction.)

**Still unverified, and not testable here:** a machine with `%AppData%` redirected to a
network share. That is the case `AllowNetworkLocations` exists for, and an institutional
profile is the only way to see it.

### Windows elevates the uninstaller, and we cannot stop it

Uninstalling raises a UAC prompt ("...from an unknown publisher..."), reported 7/30/2026.
It is **not** ours to fix, and the file is not at fault — verified by parsing the PE:

```
unins000.exe  32-bit, RT_MANIFEST resource present
  <requestedExecutionLevel level="asInvoker" uiAccess="false"/>
```

No `AppCompatFlags\Layers` entry forces it either, and there is only one uninstaller
registered. What triggers it is Windows' **installer-detection heuristic**, which elevates
executables whose *filename* looks like an installer; Inno Setup names every uninstaller
`unins000.exe` and offers no way to rename it. `Setup.exe` escapes the same fate only
because `VistaType LP and Braille Macros Setup <ver>.exe` does not match the pattern.

For an administrator it is one extra click and the uninstall works. **The open risk is a
transcriber who is not a local administrator** on a school or agency machine: they would be
asked for credentials they do not have and could not uninstall at all. Untested — it needs a
standard (non-admin) local account, which is a 20-minute test on the build box and does not
need Word, since nothing has to be opened.

The real fix is **code-signing** the installer and uninstaller. A signed binary carries a
named publisher, is not treated as a legacy unsigned installer, and also removes the
SmartScreen warning transcribers currently get downloading the `.exe` from GitHub. That is a
purchase and a yearly renewal, so it is Jerry's decision rather than a task.

**Known gap:** `Vt_Put_Tabs_Back_On_Ribbon` (in `RibbonCallbacks.bas`) restores the add-in's
own tabs for someone who deleted the installed ones, but has no button yet — it can only be
run from Alt+F8. Re-running the installer is the documented route in the meantime.

Note: enabling `AllowNetworkLocations` lets Word honor trusted locations on network
paths for that user (needed for roaming/redirected `%AppData%`). It slightly broadens
trust Word-wide, which is the right call for this user base but worth being aware of.

Known limits the installer **cannot** fix (they need code-signing or IT involvement):
enterprise Group Policy that disables trusted locations, or a policy requiring add-ins
to be signed by a trusted publisher.
