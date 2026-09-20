# Code signing, and Windows Defender deleting the installer

Written 8/13/2026, after Windows Defender began deleting the installer.
Rewritten 8/15/2026 after a proper look at what was happening.

## What happened, in order

- **8/13–8/14/2026.** Windows Defender on the build box began deleting
  `VistaType LP and Braille Macros Setup <ver>.exe` within seconds of it landing, calling it
  **`Trojan:Win32/Bearfoos.B!ml`**.
- **8/14/2026.** The Quick Access Toolbar and ribbon setup was rewritten to take PowerShell out
  of the installer completely, proven to produce identical results, and Defender deleted the new
  build just the same. Reverted. See `installer/qat-code.iss`.
- **8/14/2026, 6:19 PM.** The last detection on that machine. Fourteen in all, every one the same
  name.
- **8/14/2026, 8:34 PM.** Defender's definitions updated on that machine.
- **8/15/2026.** Both the last known-good installer (3.0.135) and the current one (3.0.186) scan
  **clean** on the build box, same Defender, using `MpCmdRun`.
- **8/15/2026, the same afternoon.** That turned out to be the smaller half of the picture.
  **The deletions had hit both testers' machines too**, not only the build box — so this was
  never one machine's bad definitions. And the first VirusTotal run on 3.0.186 came back with
  **five of seventy engines flagging it, Microsoft among them**, as
  `Trojan:Win32/Wacatac.B!ml` — a different machine-learning bucket, same kind of verdict.
  The others were APEX, CrowdStrike (60% confidence), Skyhigh, and Rising, whose label
  (`Macro.Run.f`) is reacting to the macro-enabled template inside.

- **8/15/2026, later.** The two free fixes were made — the full version block and VistaType's own
  icon and wizard artwork — and 3.0.187 was built and measured against 3.0.186. **It did not
  help.** Six engines flagged it instead of five, Microsoft still calling it
  `Trojan:Win32/Wacatac.B!ml`. (CrowdStrike dropped off; Cylance and DeepInstinct appeared. The
  difference between five and six across two different files is noise; the point is that
  nothing moved.) The fixes stay — they were right on their own merits — but they are not the
  answer, and nobody should spend another day on that theory. See below.

**So it is not over, and it is not local.** A quiet on-machine scan and a flagging cloud can
coexist: `MpCmdRun` asked the copy that machine already knows, while VirusTotal ran Microsoft's
engine against a file it was meeting for the first time. The second is closer to what a
transcriber's machine does with a fresh download.

## Nothing in the installer changed

The obvious suspicion was that something added between 3.0.135 — the last build known to install
in the field — and 3.0.186 set this off. It did not. Between those two, `installer/vistatype.iss`
changed by **one line**, and that line is the version number:

```
-  #define AppVer      "3.0.135"
+  #define AppVer      "3.0.186"
```

Nothing else. Not a new file, not a new PowerShell call, not a new registry write, not a changed
setting. `installer/qat-code.iss` was added in that stretch but is **not compiled into anything** —
there is no `#include` of it. The Makefile changed only to bump the version and to deliver the
finished `.exe` into a folder Defender is told to leave alone.

So the installer VistaType ships is, in every structural respect, the installer that worked. What
changed was the verdict, not the file.

## What the detection actually is

`Trojan:Win32/Bearfoos.B!ml` is not a match against known malware. The `!ml` means a **model scored
the file** and the score crossed a line. Microsoft's own encyclopedia entry for it carries no
description of behavior, because there is nothing specific to describe — the name is a bucket.
Other people's brand-new unsigned installers land in the same bucket routinely, including
Microsoft's own tools.

**There is therefore no signature to avoid.** There is only a score, built from trust signals the
file either has or has not got. That is why removing the PowerShell changed nothing: it removed one
input to a model that was already condemning the file on several others.

The timing says the same thing. Defender deleted the file *before anything ran* — seconds after it
appeared on disk. So what the installer **does** was never what was being judged. What the file
**is** was.

## What the file is, and what scores against it

Ranked heaviest first, from a full inventory of the built `.exe` on 8/15/2026:

1. **It is unsigned.** No certificate at all — the compiled file has no signature block. This is
   the dominant factor and the only one worth real money.
2. **Five hidden `powershell.exe -ExecutionPolicy Bypass` launches** (`installer/vistatype.iss:266–285`).
3. **A macro-enabled `.dotm` dropped into Word's STARTUP folder**, where it loads with every
   document (`:80`, `:144`).
4. **A Word Trusted Location created for that same folder**, plus `AllowNetworkLocations` set to 1
   (`:602`, `:608–611`). Drop a macro template, then trust the folder you dropped it in — that is
   the textbook shape of Office macro persistence, and it is what Defender's Office rules exist to
   catch. VistaType has an honest reason for every step; a model cannot see reasons.
5. Two `.ps1` files written into `%AppData%` and then run hidden.
6. ~~Fonts installed and registered under the user's own Fonts key.~~ **Gone 8/20/2026** — the
   bundled typeface was dropped from the product. The installer now *deletes* four `.ttf` files
   and clears their entries from that key instead, which is a different shape but not obviously a
   friendlier one to a machine-learning model. Worth re-running `make scan` to see whether the
   score moved either way.
7. A WMI `Win32_Process` query, to find a running Word (`:365–367`).
8. A rename-and-rename-back probe on a file in Word's STARTUP folder, to test whether Word has it
   locked (`:377–389`). Innocent, and also what ransomware rules watch for.
9. ~~**The version block is half empty.**~~ **Fixed 8/15/2026.** `VersionInfoVersion` was not
   set, so the compiled `.exe` reported **FileVersion 0.0.0.0** with the FileVersion and
   OriginalFileName strings blank. All eight version fields are now set outright in
   `installer/vistatype.iss`.
10. ~~**No `SetupIconFile`.**~~ **Fixed 8/15/2026.** The installer wore Inno Setup's stock icon
    and was generic in its resources — the same shell malware families use when they abuse Inno
    Setup. It now carries VistaType's own icon and wizard artwork, generated from
    `assets/branding/` by `tools/lib/build_branding.py`.
11. **Zero history.** Every `make installer` bumps the version, which changes the file, which makes
    a file Microsoft's cloud has never seen anywhere in the world. On 8/15/2026 `dist/` held **108
    installers, every one a different file**, built over eleven days. Defender weights
    never-seen-before heavily, and a new unsigned installer appearing several times an hour on one
    machine is the condition under which these verdicts are likeliest.

### Items 9 and 10 were fixed on 8/15/2026, and measured, and did not help

They cost nothing and they were the only free moves on the board, so they were worth making. In
a Microsoft-run project that hit this same kind of detection, "no version information" was
listed alongside "unsigned" as one of the two high-weight factors.

**Measured, not assumed.** 3.0.186 (no icon, FileVersion 0.0.0.0) scored **5 detections of 70**
on VirusTotal. 3.0.187, identical but for the full version block and VistaType's own icon and
artwork, scored **6**. Microsoft flagged both, with the same name.

That is now the second theory tested to destruction — the first being that the PowerShell was
doing it. Both were reasonable, both were checked properly, and neither was the cause. What is
left on the list above is item 1.

**Keep the fixes anyway.** A transcriber who right-clicks a download and reads Properties before
running it — which is what a careful one does — now sees who made it and which version it is.
That was worth doing whatever Defender thinks.

**Never reach for a self-signed certificate.** Windows treats a self-signed file as *less*
trustworthy than an unsigned one — self-signing is itself a marker of malware. It would make this
worse, not better.

## How to test, instead of guessing

Two instruments. Use both; they answer different questions.

### VirusTotal — about seventy engines at once

`make scan` uploads the newest installer to [VirusTotal](https://www.virustotal.com), waits for the
result, and prints every engine that flagged it. See `DEVELOPMENT.md` for setting the key up. It
also stops with an error if the count is above the limit, or if **Microsoft** in particular flags
it, so it can sit in front of a release.

Two things to know:

- **Uploads are public.** VirusTotal keeps the file forever, shares it with every antivirus vendor,
  and lets its paying customers download it. For a GPL installer given away free that is harmless,
  and it is genuinely useful — it puts the sample in front of the vendors who need to stop flagging
  it. **Never upload a transcriber's document.**
- **VirusTotal's "Microsoft" engine is not the Defender on a real machine.** VirusTotal runs the
  engines without their cloud, and an `!ml` verdict is a cloud judgment. VirusTotal can come back
  clean while a real machine still deletes the file. It is one instrument, not the verdict.

### Defender itself, without losing the file

On the build box:

```
"C:\Program Files\Windows Defender\MpCmdRun.exe" -Scan -ScanType 3 ^
    -File "C:\path\to\Setup.exe" -DisableRemediation
```

`-DisableRemediation` means it reports what it finds and does not delete it. Same machine, same
definitions, an answer in seconds. This is what established on 8/15/2026 that 3.0.135 and 3.0.186
are both clean today.

`Get-MpThreat` lists what Defender has found, and `Get-MpComputerStatus` gives the definition
version and when it last updated — enough to line a detection up against a definition change.

**One caution about both.** A file that has been on the machine before, or that Defender has already
sent to Microsoft, is not a fair test any more. The honest test is a **brand-new build with a new
version number**, delivered somewhere Defender is not told to ignore.

### The bisect ladder, if it comes back

If it returns and the cause has to be found rather than argued about, build a ladder of test
installers and scan each one: an empty installer first, then one that only copies the files, then
one that adds the Trusted Location, then one that adds the PowerShell, then the real thing.
Whichever rung starts getting flagged is the answer. About an hour's work, and it ends the
theorizing.

## Reporting a false positive to Microsoft

Free, and the only lever that helps anyone other than the person who owns the machine.

Submit the `.exe` at **[microsoft.com/wdsi/filesubmission](https://www.microsoft.com/en-us/wdsi/filesubmission)**,
choosing **"Software developer"** and marking it a false positive. Give the detection name, and say
plainly what the installer does and why. Set the priority to **Medium** — Low submissions may never
be looked at, and High is reserved for emergencies.

Medium turns round in a few days. If Microsoft agrees, the correction rides out in the next
definition update to every machine, not only the one that complained.

The catch: the fix is tied to that exact file, and the next build is a different file. Which is
what makes signing the durable answer.

## Signing: what it does, and what it does not

**Does:** puts a verified identity on the file. Windows can then say who made it, and the model has
something to weigh that is not guesswork. In practice this is what stops false-positive deletions of
small installers.

**Does not:** make the "Windows protected your PC" message vanish on day one. That is **SmartScreen**,
which judges a file's download history, and history has to be earned. Early users may still see it,
with a **More info → Run anyway** underneath.

**Note, changed since this file was first written:** an **EV certificate no longer skips that
warning**. Microsoft removed that behavior in 2024, and EV-signed files now build history like any
other. Paying the EV premium to clear SmartScreen is no longer worth doing.

**Also does, and this matters separately:** the same certificate can sign the **VBA project inside
`LPandBRL.dotm`**. Some institutions set Word to run only digitally signed macros, and on those
machines the add-in cannot work at all — the trusted-location entry the installer writes does not
override that policy. Signing the VBA project is the only fix, and it is arguably the bigger win of
the two.

## The options, as of 8/15/2026

| | Stops the deletion | Signs the VBA project | Runs unattended | Cost |
|---|---|---|---|---|
| **Azure Artifact Signing** | Yes | **No** | **Yes** | ~$10/month |
| **Certum open source** (card) | Yes | Yes | No — needs a PIN | ~$50–90 once |
| **Standard (OV) certificate** | Yes | Yes | No — needs a token | $150–300/year |
| EV certificate | Yes | Yes | No — needs a token | $400+/year |
| Self-signed | **No — makes it worse** | — | — | free |

### Azure Artifact Signing — new, and it fits this build

Microsoft's own service, renamed in 2026 from Trusted Signing. About **$9.99 a month**, and
crucially **no card and no PIN**. It reached general availability in April 2026 and now accepts
**self-employed individuals in the USA**; the old three-years-of-history requirement is gone.

Why it suits this project in particular: signing runs from the command line with no hardware
plugged in, so it drops into `make installer` over SSH with nobody at the keyboard. The card-based
options cannot do that — the card has to be in the build box, and a PIN prompt with nobody watching
stalls the build.

Setting up needs an Azure account and an identity check (government ID and a photograph of you, put
through a third-party verifier). Allow a few days; some people report longer.

**The catch, and it is a real one: Azure Artifact Signing does not sign VBA projects.** That needs a
certificate sitting in Windows' own certificate store. People who have tried it report the command
claiming success while no signature actually lands on the project. So the two goals have come apart:
this option fixes the installer and does nothing for institutional macro policy.

### Certum open source — cheapest, and it does sign macros

About **$50–90** once, including a **Mini cryptoCertum 3.7 card** and a reader. Two things to weigh:

1. **The publisher shown is "Open Source Developer", not "Jerry Whittaker."** It stops the deletion
   without putting your name on your own work.
2. **It is revoked if the software is ever distributed commercially.** VistaType is GPLv3 and given
   away, so it qualifies today.

Since June 2023 the private key must live on hardware, so a certificate file can no longer be
installed on the build box. Note before ordering: **the card cannot be returned or refunded.**

Proving who you are is more involved than for an ordinary certificate — a full copy of an ID
document, photographs of you holding it, a utility bill, and the URL of the open-source project.
Documents go to Certum by email; expect days rather than minutes.

Validity: they sell 1, 2 and 3 year terms, but since **1 March 2026** no code-signing certificate may
be issued for longer than **460 days**. A 2- or 3-year purchase is a shorter certificate plus free
reissues along the way. A one-year purchase is one certificate, once, and is the simplest thing.

### Ruled out: SignPath Foundation

SignPath gives free certificates to open-source projects, and VistaType's license would qualify. It
only signs files built by a **public automated build system it can check against the source**, and
VistaType is compiled by Word on a private Windows box over SSH, which it cannot see. Not available
without rebuilding how the project is built.

## Where signing would fit in the build

Signing has to happen on Windows, so it becomes the last step of `make installer` on the build box —
after Inno Setup produces the `.exe` and before it is copied back. Inno Setup has built-in support
for it (a `SignTool` entry), so it is a small change to `installer/vistatype.iss` plus one line in
the Makefile.

Three practical points:

- **With a card, the card lives on the build box.** Signing anywhere else means moving it, and a PIN
  prompt with nobody watching stalls an unattended build. Azure Artifact Signing has neither problem.
- **Always timestamp.** A signature with a timestamp stays valid after the certificate expires; one
  without it stops being trusted the day the certificate lapses, and every installer ever shipped
  goes bad at once. This matters doubly for Azure Artifact Signing, whose certificates last three
  days by design.
- **Signing the VBA project is a separate operation, but it is not manual** — corrected 8/17/2026.
  Word's VBA editor (**Tools → Digital Signature**) is one way and would have to be redone by hand
  on every build. It is not the only way: Microsoft publishes libraries that teach `signtool` to
  sign the VBA project inside a `.dotm` from the command line, so it can go into the build like
  anything else. They are installed and tested on the build box — see *The card is bought, and the
  build box is ready* at the end of this file.

## What to do, in order

1. ~~Fill in the version information and give the installer its own icon.~~ **Done and measured
   8/15/2026. It did not help** — 5 detections before, 6 after, Microsoft flagging both.
2. **Report it to Microsoft** as a software developer, now rather than later. Two obscure
   engines' worth of noise would not be worth the trouble; Microsoft flagging it, on machines
   belonging to real testers, is. Free, and the only lever that reaches a transcriber who
   downloads from GitHub tomorrow.
3. **Run the bisect ladder.** Two theories have now been tested and neither was the cause, so
   stop theorizing and find out. An empty installer, then one that only copies the files, then
   one that adds the Trusted Location, then one that adds the PowerShell — `make scan` each.
   About an hour.
4. ~~**Decide on signing.**~~ **Decided and bought, 8/17/2026** — the Certum open source kit, card
   and reader, three or four days for delivery. Everything that can be prepared without the card is
   done; see *The card is bought, and the build box is ready* at the end of this file.

Worth noticing about the two remaining unsigned-file factors, 9 and 10 having now been ruled
out: **Rising's label is `Macro.Run.f` and Skyhigh's is `BehavesLike.Win32.ObfuscatedPoly`.**
Those two are reacting to the macro-enabled template inside the installer, not to the installer's
own shape.

## Where this stands, 8/15/2026

**Jerry's call, on reading the scan results:** the engines are looking *inside* the installer at
the macro template, so the certificate to buy is one that **can sign the `.dotm` as well as the
`.exe`**.

That rules **Azure Artifact Signing out**, cheap and unattended though it is — it cannot sign a
VBA project, and a certificate that fixes only half the problem is the wrong purchase. It leaves
the two card-based options from the table above:

- **Certum open source**, about $50–90 once. Publisher reads "Open Source Developer", not
  Jerry Whittaker. Revoked if the software is ever sold, which VistaType is not.
- **A standard OV certificate**, $150–300 a year, carrying Jerry's own name.

Both need the card in the build box and a PIN, so `make installer` stops being fully unattended.
That is a real cost and it is accepted knowingly — it buys the one thing Azure cannot.

**Deferred, not abandoned: the bisect ladder** (above). It is the thing that would *confirm*
whether the macro template is really what the engines are reacting to, rather than inferring it
from two engine labels. Worth an hour before any money is spent, and worth it afterwards too,
because it would say whether signing the `.dotm` alone is enough.

**Also still to do:** report the false positive to Microsoft as a software developer. Free, and
independent of everything above.

### Deferred until signing is settled — do not start these

Jerry, 8/15/2026: **everything below waits on the certificate.** They were worked out, not
forgotten, and they are recorded here so the reasoning does not have to be rebuilt.

**The repository is still private, and stays that way for now.** Only Jerry's account can reach
it — there are no collaborators. So nobody can download from GitHub at all, and the Download
button on vistatypelp.org leads to a "page not found" for every visitor. That is known and
accepted for the moment: pointing the public at an installer their antivirus deletes is worse
than pointing them at nothing. **3.0.135, the build that worked in the field, is flagged too**
(5 of 70 engines, Microsoft calling it `Program:Win32/Wacapew.C!ml`), so releasing it would not
have avoided this.

Note for whenever it does go public: the GPL says anyone given the installer is entitled to the
source for that exact version, and the About dialog and the installed `LICENSE.txt` both promise
it. The history was checked on 8/15/2026 and is clean — `build.config` was never committed, and
there are no keys or tokens anywhere in it. (It also checked for the Braille Institute's
Atkinson Hyperlegible license PDF, from the typeface bundled between 3.0.101 and 3.0.196. That
typeface was dropped on 8/20/2026 and `assets/fonts/` deleted; the `.gitignore` rule refusing
the PDF by name stays, since the file is still on Jerry's disk.)

**A release-candidate lane was designed and deferred.** The shape, so it need not be worked out
twice:

- A fourth lane, `X.Y-rcN`, for builds handed to the testers. `3.0.X` stays the private build
  counter; `X.Y` stays a real release; `X.Y.0.N` stays a hotfix.
- **Two version strings, not one.** `VersionInfoVersion` accepts *numbers only* (up to four,
  dot-separated), so `3.1-rc1` breaks the build if fed to it. Everything a person reads — the
  About box, Programs and Features, the wizard, the file name — takes free text.
- The numeric half should be **the build the candidate was cut from**: `3.1-rc1 (build 3.0.188)`,
  numerically `3.0.188.0`. Anything cleverer collides with the hotfix lane, where `3.1.0.1`
  already means "3.1, repair 1". It also means a tester's report points at one exact commit.
- **Published with GitHub's pre-release flag**, which is excluded from "the newest release" — so
  a candidate can never become the public download by accident, which is what the website
  follows.
- **Tagged on `dev`; `master` never moves.** That keeps "master is the last released version"
  true.
- **Never put "rc" in `AppName`.** With no `AppId`, Inno derives product identity from it, so a
  renamed one looks like a different product and leaves a second entry in Programs and Features.
  Say it in `AppVersion`, `UninstallDisplayName`, the welcome and finished pages, and the About
  dialogs instead.
- Build side: leave `APPVER` and `make bump` alone (its arithmetic would trip over a label), add
  an optional `RCVER`, and pass the installer two defines instead of one.

## The card is bought, and the build box is ready — 8/17/2026

Jerry ordered the **Certum open source signing kit** — the card and its reader — on 8/17/2026, with
three or four days to delivery.

Everything that does not need the card in hand is now done, and more importantly **proved**, so the
day it arrives is a short day rather than the full day of troubleshooting other people report.

### What the build box now has

`tools/windows/Setup-SigningTools.ps1` did it, and can do it again if the box is ever rebuilt:

```bash
# vistabuild is WIN_HOST from build.config; C:/vt-signing is where the script puts things anyway
scp tools/windows/Setup-SigningTools.ps1 vistabuild:'C:/vt-signing/Setup-SigningTools.ps1'
ssh vistabuild 'powershell -ExecutionPolicy Bypass -File C:/vt-signing/Setup-SigningTools.ps1'
ssh vistabuild 'powershell -ExecutionPolicy Bypass -File C:/vt-signing/Setup-SigningTools.ps1 -SkipSdk -Register'
```

That copy on the box is a copy, and copies drift. The one under `tools/` in this repository is the
real one; if the box's ever disagrees, delete it and send this one again.

**Before that folder is ever deleted or moved, run the script with `-Unregister`.** Windows points
at those libraries machine-wide, and a dangling pointer breaks signature checking on Word files for
every program on the box, silently. A `DO-NOT-DELETE.txt` saying so is written into the folder.

- **`signtool.exe`, from Windows SDK 10.0.28000.2526** — the **`SigningTools` feature only**, which
  is a small install; the whole SDK is several gigabytes and nothing else in it is wanted. Three
  copies arrive, at
  `C:\Program Files (x86)\Windows Kits\10\bin\10.0.28000.0\{x64,x86,arm64}\signtool.exe`.
- **The Office Subject Interface Packages, 16.0.19416.43425** — `msosip.dll` and `msosipx.dll`,
  which is what lets `signtool` sign the VBA project inside a `.dotm`. Both the 64-bit and the
  32-bit set, unpacked to `C:\vt-signing\officesips-x64` and `...-x86`, and registered in both
  halves of the registry. They also bring `offsign.bat` (the wrapper that does the whole signing
  dance) and `offclearsig.exe` (removes existing signatures).
  **These do not ship with Office** — checked on 8/17/2026, there is no copy anywhere under
  Program Files, which is exactly why Microsoft publishes them separately.
- The **Visual C++ 2015–2022 runtimes** their readme requires: both already present.
- Everything lives under **`C:\vt-signing`**, deliberately outside the build folder, because
  `src/`, `tools/` and `installer/` are all wiped on that box before each build.
- Both Microsoft downloads were **checked with `Get-AuthenticodeSignature` before anything was run
  or registered** — both "Valid", both Microsoft Corporation. That check is not decoration: this
  package was reported in 2022 as being signed by *Microsoft Testing Root Certificate Authority
  2010*, which nothing trusts. It is fixed in this release. The script refuses to register a
  library whose signature is anything other than valid.

### Signing was rehearsed end to end, without the card

A throwaway self-signed certificate was created on the box, used, and removed again — confirmed gone
from all three certificate stores afterwards. Five things came out of it, and the second would have
cost a day.

1. **The `.dotm` signs, in all three signature formats.** `offsign.bat` clears any old signature,
   then signs and verifies three times over — the legacy format, the agile one, and the 2020 one —
   and reported exit code 0 against the real 1.6 MB shipping `LPandBRL.dotm`. Word wants all three;
   only one can be written per pass, which is why the wrapper exists.
2. **Use the x64 `signtool`, not the x86 one — Microsoft's own instructions are wrong for this
   box.** Their guidance says to use x86. Here x86 fails every single time with
   `SignerSign() failed (0x800403f4)` and no further explanation, while x64 signs the same file
   without complaint. Word on this box is 64-bit and the signing library has to match it. Starting
   from the documented command would have looked exactly like a faulty card or a bad certificate.
3. **Injecting the ribbon afterwards does not break the macro signature.** Tested directly: sign the
   `.dotm`, run `inject_customui.py` over it, verify again — still valid. So the signature covers
   the VBA project alone and not the rest of the file, and the signing step can sit on either side
   of the injection. (The keyboard shortcuts never touch the `.dotm` at all — they go into the
   `.dotx`.)
4. **`offsign.bat` is fussy in two ways.** Its first argument is the *folder* holding `signtool.exe`
   and it **must end with a backslash**. And drive it from a `.cmd` file on the box: quoted Windows
   paths passed down through ssh → PowerShell → cmd get mangled three different ways.
5. **Verifying an unsigned `.dotm` answers "No signature found"** rather than "cannot be verified".
   That one line is how you tell the libraries are properly registered, and it works before any
   certificate exists.

The command, for the record — only the thumbprint and the file change:

```bat
cd /d C:\vt-signing\officesips-x64
offsign.bat "C:\Program Files (x86)\Windows Kits\10\bin\10.0.28000.0\x64\" ^
  "sign /sha1 <THUMBPRINT> /fd SHA256 /tr http://time.certum.pl /td SHA256" ^
  "verify /pa" "C:\Users\jerry\vistatype-build\dist\LPandBRL.dotm"
```

**Take that `10.0.28000.0` folder name from the script's own output, not from this page.** It
prints every `signtool.exe` it can find, with versions; the SDK moves on and this page will not.
(The two numbers looking almost alike — installer `10.0.28000.2526`, folder `10.0.28000.0` — are
both right and neither is a typo, so resist tidying one to match the other.)

### When the card arrives — in this order

1. **Reader driver first, then proCertum CardManager, then reboot.** Expect to hunt for the reader's
   own driver; Certum does not always supply it.
2. **CardManager → Common Profile → Initialize profile**, and set the **PIN and the PUK**. Write
   both down somewhere permanent before going further. Exhausting the PUK attempts turns the card
   into a coaster, and Certum does not refund cards.
3. **Prove who you are.** A full copy of an ID document, photographs of you holding it, a utility
   bill, and the project's URL — or the automated route (IDNow: photograph the ID, record some head
   movements), which came back in about two days for one developer.

   **Ask Certum about this one now, before the card arrives — it is free and it is the item most
   likely to stall for days.** Their open source certificate is issued on the strength of a
   *visible* open source project, and **this repository is private**: only Jerry's account can
   reach it, which was a deliberate decision while the installer was being deleted by antivirus
   (see *Deferred until signing is settled* above). vistatypelp.org is public and can be offered
   instead, but whether that satisfies them is their call, not a guess worth making. The answer
   may be that the repository has to be made public first, which is a decision with its own
   consequences and Jerry's alone to make.
4. **Certum SignService** — a third application, and *not* SimplySign or SmartSign. Their activation
   page uses it to generate the key pair **on the card** and send Certum the request.
5. **Download the issued certificate, import it to the card** in CardManager, and then into Windows'
   own Personal certificate store.
6. **The gotcha everybody hits.** Windows shows the certificate with *"No key provider information"*
   and *"Missing stored keyset"* — it does not know the key is on the card. The fix is a small
   `keyinfo.inf` naming the Subject Key Identifier, then:

   ```
   certutil -repairstore -user MY <THUMBPRINT> keyinfo.inf
   ```

7. **Prove the certificate works on something disposable first** — a copy of any small `.exe` — so
   that a PIN problem and a signing problem never get confused with each other.
8. **Then the two real ones**: the `.dotm` with the command above, and the `Setup.exe`.

### Where signing lands in the build

- **The `.dotm` gets signed on the build box, as part of `make installer` — not `make build`.** Day
  to day builds then stay unattended and need no card in the reader; only a build that is going to a
  person needs it. It has to happen **before** Inno Setup compiles, or the installer carries the
  unsigned copy.
- **The `Setup.exe` gets signed by Inno Setup itself** — a `SignTool` entry in `installer/vistatype.iss`
  plus `SignedUninstaller=yes`, so the uninstaller left on the transcriber's machine carries a
  signature too.
- **Always timestamp** (`/tr` and `/td`). Without it, every installer ever shipped stops being
  trusted the day the certificate lapses.
- **Expect four PIN prompts per installer at least** — three for the `.dotm`, because each pass is
  its own session, and one for the `.exe`. `SignedUninstaller=yes` signs the uninstaller as a
  further operation, so five or six would not be a fault. Look at CardManager's PIN-caching
  setting before accepting that as permanent.
- **Find out where the PIN prompt appears when the build is driven over SSH, before wiring anything
  in.** The card asks for its PIN with a window. A window has nowhere to go in an SSH session that
  has no desktop, and the likely result is `make installer` hanging with no error at all — which is
  exactly how the leftover `~$` lock-file bug behaved (`DEVELOPMENT.md`, "Gotchas baked into the
  tooling"), and it cost a long time to find. Sign one throwaway file over a plain `ssh` before
  believing any of this works unattended. If the prompt cannot be reached that way, signing has to
  be a step Jerry runs at the box's own screen.

**The exact place, in the `installer-build` target of the Makefile:** between the line that copies
`$(SHIPFILES)` up to the box's `dist/` and the line that runs `ISCC`. At that moment the box's
`dist/LPandBRL.dotm` is the finished file with the ribbon already in it, and Inno has not read it
yet. Nothing else has to move.

**A signed `.dotm` must never travel back to this machine. This is the trap in the whole plan.**

`make deploy` copies `dist/LPandBRL.dotm` to the repo root, and the root `.dotm` is the **base every
future build starts from** — `push-src` sends it to the box and `Import-Vba.ps1` copies it and
imports the modules into that copy. So a signed file reaching the root would seed every later build
with a signature that stops matching the moment the modules are re-imported.

Signing at the point named above avoids it completely: the signature is applied to the box's own
`dist/` copy, minutes before Inno reads it, and it exists nowhere afterwards except inside the
`Setup.exe`. This machine's `dist/` and the tracked root `.dotm` both stay unsigned, which is
correct. **Do not "improve" this by copying the signed file back down.** The consequence is that a
`make deploy` smoke-test runs an unsigned add-in while transcribers get a signed one; that
difference is the safe side of the trade, and the signed copy gets tested by installing the
`Setup.exe`, which is how it should be tested anyway.

#### Measured 9/20/2026 — what actually happens, and it is not what this page assumed

A real copy of the build base was signed on the box with a throwaway self-signed certificate,
pushed through `Import-Vba.ps1`, and the result examined. Three findings, and the third changes
the shape of the risk.

1. **The signature is not inside `vbaProject.bin`.** That part came back byte-identical — all 302
   OLE streams, every size the same. It is its own part in the Office package:
   `word/vbaProjectSignature.bin`, plus `...SignatureAgile.bin` and `...SignatureV3.bin` for the
   other two formats. Anything looking for it among the OLE streams, the way the old binary `.doc`
   format stored it, will report a signed file as clean. The first version of
   `tools/lib/check_dotm_unsigned.py` did exactly that and passed a genuinely signed file.

2. **`offsign.bat` stops at its first verify.** With an untrusted certificate it signs the legacy
   format, fails `verify /pa` on the trust chain, and exits 5 without writing the other two. The
   signing itself is fine — this is the trust chain, not the card and not the file. Adding a
   self-signed certificate to `CurrentUser\Root` over SSH is refused ("The request is not
   supported"), so a full three-pass rehearsal needs an interactive session or a trusted cert.

3. **Word does not leave a stale signature — it RE-SIGNS.** One signature part went in; three came
   out, over code in which every module had been removed and re-imported. `signtool verify /pa /v`
   on the result complained only that the root was untrusted: no hash mismatch, no invalid
   signature. Word found the certificate in the user's store and signed the rebuilt project with
   it.

   **So the danger of a signed base is not a broken signature. It is that the build box quietly
   acquires a signing step on every build** — including plain `make build` and `make try`, which
   this page requires to stay unattended and card-free. What Word does when the certificate is
   *not* reachable (card out of the reader, PIN not supplied) was not measured, and the likely
   answer is the PIN window with nowhere to go that is warned about above: `make build` hanging
   with no error.

`tools/lib/check_dotm_unsigned.py` now refuses to promote or build from a signed `.dotm`, and is
wired into `push-src` and `make deploy`. It matches on the package part name by prefix, so a
fourth signature format is caught rather than ignored.

### Still unknown until the card is here

- **Whether Word accepts it.** Set Word's macro settings on the test box to require a digital
  signature, install, and check the add-in still loads. `Document.VBASigned` answers the narrower
  question of whether a signature is present at all.
- **Whether it moves the antivirus verdict.** `make scan` before and after — ideally the same build
  signed and unsigned. That number is what this whole exercise is for, and the two engines reacting
  to the macro template (`Macro.Run.f`, `BehavesLike.Win32.ObfuscatedPoly`) are the ones to watch.
- **Whether the bisect ladder is still worth an hour.** It would say whether signing the `.dotm`
  alone is what did the work. Cheaper to answer once there is a signature to compare against.

## Sources

- [Trojan:Win32/Bearfoos.B!ml — Microsoft Security Intelligence](https://www.microsoft.com/en-us/wdsi/threats/malware-encyclopedia-description?Name=Trojan%3AWin32%2FBearfoos.B%21ml)
- [Code signing options for Windows app developers — Microsoft Learn](https://learn.microsoft.com/en-us/windows/apps/package-and-deploy/code-signing-options)
- [Trusted Signing opens to individual developers — Microsoft](https://techcommunity.microsoft.com/blog/microsoft-security-blog/trusted-signing-is-now-open-for-individual-developers-to-sign-up-in-public-previ/4273554)
- [Azure Trusted Signing and VBA macros — Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/2045602/azure-trusted-signing-how-can-i-digitally-sign-a-v)
- [Cloud protection and Microsoft Defender Antivirus — Microsoft Learn](https://learn.microsoft.com/en-us/defender-endpoint/cloud-protection-microsoft-antivirus-sample-submission)
- [mpcmdrun.exe command-line tool — Microsoft Learn](https://learn.microsoft.com/en-gb/defender-endpoint/command-line-arguments-microsoft-defender-antivirus)
- [VirusTotal — false positives](https://docs.virustotal.com/docs/false-positive)
- [SignPath Foundation — conditions for open source projects](https://signpath.org/terms.html)
- [Certum — shorter code signing certificate validity](https://www.certum.eu/en/news/shortening-code-signing-certificate-validity/)
- [Certum Shop — Open Source Code Signing](https://shop.certum.eu/open-source-code-signing.html)
- [Certum — activating a Code Signing certificate on a card (PDF)](https://files.certum.eu/documents/manual_en/CS-Open_Source_Code_Signing_Certificate_activation.pdf)
- [Certum — signing with signtool and jarsigner (PDF)](https://www.files.certum.eu/documents/manual_en/Code-Signing-signing-the-code-using-tools-like-Singtool-and-Jarsigner_v2.3.pdf)
- [One developer's start-to-finish account, including the "missing stored keyset" fix](https://piers.rocks/2025/10/30/certum-open-source-code-sign.html)
- [Office Subject Interface Packages — download](https://www.microsoft.com/en-us/download/details.aspx?id=56617)
- [Upgrade signed Office VBA macro projects to V3 signature — Microsoft](https://devblogs.microsoft.com/microsoft365dev/upgrade-signed-office-vba-macro-projects-to-v3-signature/)
- [Windows SDK downloads](https://learn.microsoft.com/windows/apps/windows-sdk/downloads)
