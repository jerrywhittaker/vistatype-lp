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
6. Fonts installed and registered under the user's own Fonts key.
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
- **Signing the VBA project is separate and manual** — Word's VBA editor, **Tools → Digital
  Signature**. It would have to be redone on every build, which means it belongs in
  `Import-Vba.ps1` rather than in anyone's hands.

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
4. **Decide on signing.** With both testers affected this has stopped being a build-box
   curiosity — it is between a transcriber and the software. A certificate is the only thing
   that helps someone downloading from GitHub, and the choice is not obvious, because the
   installer and the macros may need different ones.

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
