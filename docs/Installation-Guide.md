# VistaType LP + Braille Macros — Installation Guide

**Version 3.0**

VistaType LP adds tools to Microsoft Word for producing **large-print** documents and
**braille** source files (for the Duxbury Braille Translator). This guide covers
installing it with the new **one-click installer**.

> If you used an older version, you no longer copy files by hand or export your Quick
> Access Toolbar first. The installer does everything, it **never disturbs your existing
> ribbon**, and it **asks what to do about your Quick Access Toolbar** rather than
> deciding for you.

---

## Before you start

- **Windows** with **Microsoft Word** (Word 2013 or newer; Microsoft 365 is fine).
- **Close Microsoft Word and Microsoft Outlook.** The installer will stop and remind you
  if either is still running — they hold the files open.
- **No administrator rights needed.** It installs just for your Windows user account.

---

## Install in three steps

1. **Download** `VistaType LP and Braille Macros Setup <version>.exe`.
2. **Close Word and Outlook**, then **double-click** the installer.
   - If Windows shows a blue *“Windows protected your PC”* message, click **More info →
     Run anyway**. (This appears because the installer isn’t code-signed; it is safe.)
3. Click **Next / Install** through the wizard, then **Finish**. It takes under a minute.

That’s it — nothing to restart. Open Word and look for the two new ribbon tabs.

---

## What the installer set up

So you know exactly what changed on your computer:

| It installed / changed | Where | Notes |
|---|---|---|
| The macros (`LPandBRL.dotm`) | Word **STARTUP** folder | Loads automatically every time Word opens |
| The large-print template (`LargePrintTemplate.dotx`) | Your **Templates** folder | Holds the large-print styles |
| Two ribbon tabs: **VistaType LP** and **Braille Macros** | Word ribbon | **Added next to your existing tabs** — your ribbon is not replaced |
| Quick Access Toolbar icons | Your QAT | **You choose** during install — see below. The default installs VistaType’s standard toolbar; your own is saved first and you can have it back at any time |
| Trust settings so macros run | Word Trust Center | Marks the STARTUP folder as trusted (including network/roaming profiles) |
| Removed leftover files from old versions | Templates folder | Cleans up the obsolete “Large Print Templates” folder |
| Removed the **VistaTypeLP Legible** typeface | Your personal **Fonts** folder | If an earlier version had installed it — see below |

---

## Confirm it’s working

1. Open **Microsoft Word** (a blank document is fine).
2. Look along the ribbon for the **VistaType LP** and **Braille Macros** tabs.
3. Click **VistaType LP → Help → Version & Updates**. If a message appears, the macros
   are running correctly.

If the tabs are there and buttons respond, you’re ready to go.

---

## The VistaTypeLP Legible typeface has been withdrawn

For a short time VistaType LP installed a typeface of its own, **VistaTypeLP Legible**, and
offered it beside Tahoma when you attached the large-print template. **It is gone.** Large-print
books are set in Tahoma, as they were for years, and the attach dialog no longer asks.

**Why it was withdrawn.** The typeface covers the Latin alphabet and the accents that go with
it — English, French, German, Spanish and Italian all set correctly. It does not cover Greek,
the phonetic alphabet, or most mathematical symbols. Word does not tell you when a character is
missing: it silently borrows that one character from some other typeface, at some other size.
In a large-print book that is the one thing that must never happen, and it turns up exactly
where you would least want it — Latin, mathematics, phonetics, and the Greek that runs all
through medical material. Tahoma covers the lot.

**What this means for you:**

- **Nothing to do.** Install the new version over the old one in the usual way.
- **Books you already made in it are safe, and are left alone.** Word stored a complete copy of
  the typeface inside each of those documents when you attached the template, so they still open,
  set and print exactly as they did. If you attach the template to one of them again, VistaType
  LP recognizes it and **keeps the typeface it is already in** — it will not convert your book to
  Tahoma or move your page breaks.
- **New books are Tahoma.** No question is asked.
- **The typeface is removed from your computer.** The installer takes it back off, because
  leaving a face on the machine that VistaType LP no longer supports invites it being picked by
  hand for the very documents it cannot set. The removal finishes the next time you sign in to
  Windows — Windows will not release a typeface in the middle of a session — so if you still see
  it in Word's font list today, sign out and back in.
- **If you deliberately want to keep it**, install it yourself from the Braille Institute's own
  **Atkinson Hyperlegible**, which is free. Do that knowing what is written above: a point size
  set in Word will print about 8% smaller than it says, and the missing characters are still
  missing.

---

## Your Quick Access Toolbar — you decide

The Quick Access Toolbar is the small row of icons at the very top of the Word window.

Most people never customize it, so the installer’s default sets it up for transcription
work — click straight through and you get the icons you need for large print and braille.
But some people spend years building their own, so from version 3.0.33 the installer
**asks** rather than assuming. There is a page in the wizard with three choices:

| Choice | What happens |
|---|---|
| **Replace my toolbar with the standard VistaType LP and Braille Macros toolbar** *(default)* | You get VistaType’s full toolbar, laid out for transcription work. Your own toolbar is saved first, and you can have it back at any time. |
| **Keep my toolbar exactly as it is, and add the VistaType LP and Braille Macros icons on the end** | Nothing of yours is moved, hidden, or reordered. VistaType’s six icons are added after your own, with a divider before them. |
| **Put back the toolbar I had before VistaType LP and Braille Macros was installed** | Offered only if we are holding a saved copy — that is, an earlier VistaType version had rewritten your toolbar. One click undoes that. |

Untick the box at the top of that page and VistaType will not touch your toolbar at all.

**Your own ribbon tabs and groups are never touched, whichever you choose.** The
**VistaType LP** and **Braille Macros** tabs are added alongside them.

---

## Your ribbon tabs — yours to arrange

The installer always puts the two VistaType tabs on your ribbon as tabs of your own, so they
behave like any other tab in Word. Go to **File → Options → Customize the Ribbon** and you can:

- **Untick a tab to hide it.** Hiding **Braille Macros** while you work on large print is
  the obvious one — untick it, and it is gone until you want it back.
- **Move them.** Use the arrows to put them wherever suits you in the tab order.
- **Rename them.**

**Re-installing will not undo any of that.** The rule is simple: *the buttons inside a
VistaType tab are ours to keep up to date; where the tab sits, what it is called and
whether it is showing are yours.* An upgrade refreshes the buttons and leaves the rest
exactly as you set it.

Up to version 3.0.117 this was a choice on the installer, and declining it left the tabs
showing on the ribbon but missing from Customize the Ribbon — so they were the only tabs you
could not hide, move or rename. That is not a useful thing to be able to choose, so the choice
is gone and the tabs are always yours to arrange.

> **Earlier versions did not ask.** Up to 3.0.32 the installer merged its toolbar into
> yours, which pushed your icons to the right, removed dividers you had placed, and hid
> some of Word’s own buttons — Undo among them. If that happened to you, the third choice
> above puts it right. Your original was saved at the time and is still there.

### Restoring a VistaType icon you removed

Run the installer again and choose **Keep my toolbar exactly as it is, and add the
VistaType LP and Braille Macros icons on the end**. It puts back any of the six that are
missing and leaves everything else alone. Running it twice changes nothing the second time.

---

## Updating to a newer version

1. Close **Word** and **Outlook**.
2. Run the new `VistaType LP and Braille Macros Setup <version>.exe`.

It replaces the previous files and re-applies the settings. It is safe to run over an
existing installation — nothing is duplicated.

---

## Uninstalling

1. Open **Settings → Apps → Installed apps** (or **Control Panel → Programs**).
2. Find **VistaType LP + Braille Macros** and choose **Uninstall**.

Uninstalling removes the macros, the template, and **only the VistaType Quick Access
Toolbar icons** — your own ribbon and toolbar customizations are left in place, including
anything you added *after* installing VistaType.

If you had chosen VistaType’s standard toolbar, uninstalling puts back the toolbar you
had before — and keeps any icons you added on top of ours in the meantime.

---

## Troubleshooting — common issues

First, whatever the symptom: **fully close Word and reopen it.** Almost all changes take
effect only when Word next starts.

### “I don’t see the VistaType quick-access icons”

This usually means the icons are there but hidden — not that the install failed.

- **The Quick Access Toolbar may be turned off.** In Microsoft 365, it is hidden by
  default. Turn it on: click the small **ribbon-options arrow** at the far right end of
  the ribbon (or right-click the ribbon) → **Show Quick Access Toolbar**.
- **The icons are at the far end, grouped together.** We add them as one block —
  set off by a divider line — *after* your existing toolbar buttons, so look at the
  right-hand end of the toolbar. If your toolbar is full, click the **” (More controls)**
  chevron at its end to see the overflow.
- **Give the toolbar more room.** Right-click the toolbar → **Show Quick Access Toolbar
  Below the Ribbon.** Below the ribbon it spans the full width and hides fewer icons.
- **Still missing one?** Run the installer again and choose **Keep my toolbar exactly as
  it is, and add VistaType’s icons on the end.** It restores any of the six that are
  missing and leaves the rest of your toolbar alone.

### “The VistaType LP / Braille Macros tabs don’t appear”

- **Look for a Security Warning bar** just below the ribbon that says macros were
  disabled — click **Enable Content**. (After the installer’s trust setup this shouldn’t
  keep coming back.)
- **The add-in may have been disabled by Word** (this can happen after a Word crash):
  *File → Options → Add-ins*. At the bottom, set **Manage: Disabled Items → Go**, and if
  `LPandBRL.dotm` is listed, select it and click **Enable**. Then restart Word.
- **Confirm the add-in is loaded:** *File → Options → Add-ins*, set **Manage: Word
  Add-ins → Go**, and make sure `LPandBRL.dotm` is checked under Global Templates.

### “Uninstalling asks for an administrator”

Windows occasionally shows **“Do you want to allow this app from an unknown publisher to
make changes to your device?”** when you uninstall. Click **Yes** and it proceeds normally.

VistaType does **not** need administrator — it installs entirely inside your own user
folders, and its uninstaller explicitly asks Windows *not* to elevate it. Windows raises the
prompt of its own accord, usually after an earlier install on that computer was interrupted.

If you are **not** an administrator and Windows asks for a password you don’t have, contact
us — VistaType can be removed by hand.

### “The VistaType tabs are there, but the buttons do nothing”

The tabs are on your ribbon but the macros behind them aren’t loaded — almost always
because **Word has disabled the add-in**, which it does after a crash.

*File → Options → Add-ins*. At the bottom set **Manage: Disabled Items → Go**. If
`LPandBRL.dotm` is listed, select it, click **Enable**, and restart Word.

If it isn’t listed, check the add-in is still installed: same dialog, **Manage: Word
Add-ins → Go**, and make sure `LPandBRL.dotm` is ticked under Global Templates. If it has
gone altogether, run the installer again.

### “Where did the VistaTypeLP Legible typeface go?”

It was withdrawn — see *The VistaTypeLP Legible typeface has been withdrawn* above. The short
version: it has no Greek, no phonetic alphabet and almost no mathematics, and Word substitutes a
missing character silently and at the wrong size. Large-print books are Tahoma again.

Books you have already produced in it are untouched and keep working: each one carries its own
copy of the typeface inside the file.

If Word still lists the typeface after installing the new version, **sign out of Windows and
back in**. Windows does not release a personal typeface in the middle of a session.

### “I removed the VistaType tabs and want them back”

Run the installer again. It puts them back without disturbing anything else on your ribbon.

### “I see two VistaType tabs” (after upgrading)

An old version may have left its ribbon file behind. Delete the file
`%AppData%\Microsoft\Office\Word.officeUI` **only if** you did not customize your own
ribbon; otherwise contact us (below) so we don’t remove your customizations. Restart Word.

### “Macros are blocked / a security warning keeps appearing”

- **Check macro security:** *File → Options → Trust Center → Trust Center Settings →
  Macro Settings.* “Disable all macros with notification” is fine — the installer’s
  trusted-location setting lets the VistaType macros run regardless.
- **Managed or organization computers:** Your workplace’s IT policies (Group Policy) can
  block macros or require add-ins to be digitally signed, and these override the
  installer. Contact your IT department, or reach out for help below.

**Need help?** Email **jerry@thewhittakers.org**.

---

## Manual installation (only if you can’t run the installer)

Some locked-down computers block running an `.exe`. In that case an administrator can
place the two files by hand:

| File | Copy it to |
|---|---|
| `LPandBRL.dotm` | `%AppData%\Microsoft\Word\STARTUP` (create the `STARTUP` folder if missing) |
| `LargePrintTemplate.dotx` | `%AppData%\Microsoft\Templates` |

Tip: paste `%AppData%\Microsoft\Word\STARTUP` into the File Explorer address bar to open
that folder directly. Close Word first, then reopen it after copying.

> With a manual install, the ribbon tabs still appear (they’re built into
> `LPandBRL.dotm`), but the Quick Access Toolbar icons are **not** added automatically.
> Everything is still reachable: the tools are on the **VistaType LP** and **Braille
> Macros** tabs, and the keyboard shortcuts work (Ctrl+Alt+Shift+I for Document Settings,
> for instance). To put an icon on your toolbar yourself, right-click any VistaType ribbon
> button → *Add to Quick Access Toolbar*. You may also need to confirm the STARTUP folder
> is a trusted location in the Trust Center.
