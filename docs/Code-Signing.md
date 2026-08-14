# Code signing — what it would involve

Written 8/13/2026, after Windows Defender began deleting the installer.

## The problem it would solve

Defender deletes `VistaType LP and Braille Macros Setup <ver>.exe` within seconds of it
landing, and calls it **`Trojan:Win32/Bearfoos.B!ml`**. The `!ml` on the end is the important
part: it means a machine-learning **guess**, not a match against anything known. Nothing in the
file resembles real malware. A model looked at the behavior and decided.

What it is reacting to is a combination that is genuinely unusual and genuinely ours:

- the installer is **unsigned**, so nobody's name is on it
- it drops a **macro-enabled Word template** into Word's STARTUP folder, so it loads with every
  document the transcriber opens
- it writes into `%AppData%` and installs fonts

That is also an accurate description of Office macro malware. A model has no way to tell the two
apart without either a signature or a known publisher.

**Taking PowerShell out did not help.** That was tried in full on 8/13/2026 — the toolbar and
ribbon setup was rewritten in the installer's own scripting language and proven to produce
identical results, and Defender deleted the new build just the same, with the same detection
name. See `installer/qat-code.iss`. That leaves signing.

## What signing does, and what it does not

**Does:** puts a verified identity on the file. Windows can then say who made it, and both
Defender and SmartScreen have something to trust that is not guesswork. In practice this is what
stops false-positive deletions of small installers.

**Does not:** make the warning vanish on day one. SmartScreen judges a *reputation*, which a
standard certificate has to earn through downloads over time. Early users may still see
"Windows protected your PC" with a **More info → Run anyway**. An EV certificate gets reputation
immediately, and costs several times as much.

**Also does, and this matters separately:** the same certificate can sign the **VBA project
inside `LPandBRL.dotm`**. Some institutions set Word to run only digitally signed macros, and on
those machines the add-in currently cannot work at all — the trusted-location entry the installer
writes does not override that policy. Signing the VBA project is the only fix. If you ever buy a
certificate, this is arguably the bigger win of the two.

## The two ways to buy one

Since June 2023 the private key must live on **hardware** — a USB token or a smart card. You
cannot install a certificate file on the build box any more.

| | **Certum Open Source** | **Standard (OV) certificate** |
|---|---|---|
| Cost | about **$50–90** including the card and reader | about **$200–400 a year** |
| Publisher name shown | **"Open Source Developer"** | **your name** |
| Who may use it | individuals, non-commercial open-source projects only | anyone |
| SmartScreen | reputation must be earned | reputation must be earned |
| Sold by | Certum, and resellers | Sectigo, DigiCert, SSL.com, resellers |

There is a third tier, **EV**, at roughly $400–700 a year, which is the only one that clears
SmartScreen immediately.

### The catch with the cheap one

Two things to weigh before choosing it:

1. **The publisher shown to the transcriber is "Open Source Developer", not "Jerry Whittaker".**
   The certificate carries that phrase in place of an organization name. It stops the deletion,
   but it does not put your name on your own work.
2. **It is revoked if the software is distributed commercially.** VistaType is GPLv3 and given
   away, so it qualifies today. If you ever sell it, or bundle it into something sold, the
   certificate goes.

## What buying the Certum one actually involves

**Validity.** They sell 1, 2 and 3 year terms, but since **1 March 2026** no code-signing
certificate may be issued for longer than **460 days** — about 15 months. A 2- or 3-year purchase
is a shorter certificate plus free reissues along the way. A one-year purchase is one
certificate, once, and is the simplest thing.

**What arrives.** A **Mini cryptoCertum 3.7 card** and a reader. Note before ordering: **the card
cannot be returned or refunded** if you send the set back.

**Proving who you are.** More involved than an ordinary certificate:

- a full copy of an ID document — passport, driver's license or ID card, both sides
- photographs of you holding that document
- a utility bill
- the URL of the open-source project proving you work on it — the GitHub repository does this

Documents go to Certum by email. Expect this to take days rather than minutes.

**Setting it up.** Install Certum's card software and drivers on the build box, generate the key
**on the card**, complete the validation, and the certificate is issued onto the card.

## Where signing would fit in the build

Signing has to happen on Windows, with the card plugged in, so it becomes the last step of
`make installer` on the build box — after Inno Setup produces the `.exe` and before it is copied
back here.

Three practical points:

- **The card lives on the build box.** Signing on any other machine means moving it.
- **A PIN is needed.** Depending on the card's settings this may be typed once per session or
  once per signature, which matters because the build is driven over SSH with nobody watching.
  Worth checking before relying on an unattended build.
- **Always timestamp.** A signature with a timestamp stays valid after the certificate expires;
  one without it stops being trusted the day the certificate lapses, and every installer you have
  ever shipped goes bad at once.

Inno Setup has built-in support for this — a `SignTool` entry — so it is a small change to
`installer/vistatype.iss` plus one line in the Makefile.

Signing the VBA project is separate and manual: it is done in Word's VBA editor,
**Tools → Digital Signature**, choosing the certificate. It would have to be redone on every
build, which means it belongs in `Import-Vba.ps1` rather than in anyone's hands.

## What I would suggest

**If the goal is to stop the deletions and you are content with "Open Source Developer" as the
publisher:** the Certum open-source certificate, one-year term. Cheapest thing that works,
roughly $50–90 once, and it also unlocks signing the VBA project for institutional users.

**If your name on the installer matters** — and for something transcribers install on their work
machines, there is a fair argument that it does — a standard certificate at $200–400 a year is
the honest choice.

**Before spending anything**, worth doing: put the current installer on another Windows machine
and see whether Defender deletes it there too. If it survives, this is peculiar to the build box's
definitions and not a general verdict, and the whole question can wait.

## Sources

- [Certum — Transition to shorter Code Signing certificate validity periods](https://www.certum.eu/en/news/shortening-code-signing-certificate-validity/)
- [Certum Shop — Open Source Code Signing](https://shop.certum.eu/open-source-code-signing.html)
- [Certum support — Code Signing required documents](https://support.certum.eu/en/code-signing-required-documents/)
- [Certum — Code Signing certificates](https://www.certum.eu/en/code-signing-certificates/)
