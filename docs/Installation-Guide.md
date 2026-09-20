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

That’s it. Open Word and look for the two new ribbon tabs — the macros, the tabs and
the toolbar are ready the moment setup finishes.

The one thing that is not is the **VistaTypeLP Sans** typeface. Windows takes a personal
typeface into use only when you next sign in, so sign out of Windows and back in before
you expect to see it — see *The VistaTypeLP Sans typeface* below.

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
| The **VistaTypeLP Sans** typeface, in four faces | Your personal **Fonts** folder | Installed for your Windows account only, which is what avoids the administrator password. **Left in place if you uninstall** — see below |
| The typeface’s license | `%AppData%\VistaType LP Sans Fonts` | Three SIL Open Font License texts, one per typeface it is built from |
| Removed the **VistaTypeLP Legible** typeface | Your personal **Fonts** folder | If an earlier version had installed it — see below |

Two things VistaType LP writes **itself**, the first time it runs rather than during the
install, and **leaves behind when you uninstall** — so a reinstall does not cost you them:

| It writes | Where | What it holds |
|---|---|---|
| `VistaType.ini` | `%AppData%\VistaType LP Settings` | Your own Word settings, so VistaType LP can put them back after a book has changed them |
| `AutoCorrect-DEF.txt`, `-LP.txt`, `-BRL.txt` | the same folder | **Three separate AutoCorrect lists, one per kind of document.** Word keeps only one AutoCorrect list for the whole program, so until 9/19/2026 the fraction entries a braille or large-print book removes came off your own letters too, permanently. Now each kind of document gets its own list and yours is left alone. Document Settings shows which one is in use |
| `VistaType-Errors.log` | the same folder | The last 50 faults, newest first. If a macro fails, the message offers to open this folder |

Formatted AutoCorrect entries of your own are never touched — their content lives inside
`Normal.dotm` and cannot be read back — so those stay shared across all three lists.

---

## Confirm it’s working

1. Open **Microsoft Word** (a blank document is fine).
2. Look along the ribbon for the **VistaType LP** and **Braille Macros** tabs.
3. Click **VistaType LP → Help → Version & Updates**. If a message appears, the macros
   are running correctly.

If the tabs are there and buttons respond, you’re ready to go.

---

## The VistaTypeLP Sans typeface

VistaType LP installs a typeface of its own, **VistaTypeLP Sans**, and offers it beside
Tahoma under **Font Choice** when you attach the large-print template. For a new book
VistaTypeLP Sans is the one already selected; Tahoma is one click away.

**What it is.** Google’s **Noto Sans**, with three changes made for large print:

- **The zero is slashed.** It *is* the ordinary zero — there is no setting to find and none
  to forget.
- **The whole face is enlarged by an exact 25/24**, so a point size set in Word matches the
  VistaType ruler. Tahoma runs small against it.
- **The whole of Noto Sans Math and Noto Sans Symbols is folded into every one of the four
  faces.**

That last change is the one that matters. It gives the typeface **6,450 characters against
Tahoma’s 3,772** — Greek, the phonetic alphabet and mathematics complete — so Word is never
left quietly filling a missing character out of some other typeface at some other size. In a
large-print book that is the one thing that must never happen.

It also has **four drawn faces where Tahoma has two.** There is no italic Tahoma and there
never was: what you see when you press Ctrl+I in a Tahoma book is Word slanting the roman by
machine. In VistaTypeLP Sans the italic and the bold italic are real.

VistaTypeLP Sans is released under the **SIL Open Font License**, which is what allows
VistaType LP to modify it, install it, and store a complete copy of it inside every book set
in it. That embedded copy means a book still sets and prints correctly on a computer that has
never seen the typeface.

### When to pick Tahoma instead

Two reasons, and both are real:

1. **Hebrew, Arabic or Thai in the text.** Tahoma sets all three properly. VistaTypeLP Sans
   sets none of them, and needs Tahoma or a third typeface for those runs.
2. **A book that is already set in Tahoma.** VistaTypeLP Sans fits about **nine percent less
   text on a line**, so the same book at the same point size **repaginates**. A book already
   in a reader’s hands must not be reset.

You do not have to think about fill-in lines. Their underscores draw with small holes in
VistaTypeLP Sans, so VistaType LP types them in Tahoma by itself whatever the book is set in,
and puts them back that way after every attach.

### If the Typeface choice is grayed out

Windows takes a personal typeface into use only when you next sign in. If VistaTypeLP Sans is
grayed out and reads **NOT INSTALLED (or Word needs restarting)**, close Word and reopen it;
if it is still grayed, **sign out of Windows and back in**. Tahoma is selected meanwhile, and
everything else works normally.

The typeface needs **Windows 10 version 1803 or newer**. On anything older it is not installed
at all and the choice stays grayed — the rest of VistaType LP is unaffected.

### Uninstalling leaves the typeface behind

That is deliberate. A book VistaType LP set in the typeface carries its own copy inside the
file and is safe either way — but a document that was typed in it by hand, outside VistaType
LP, carries nothing, and taking the typeface off the machine would change how that document
sets. The license text stays with it, for the same reason.

---

## The earlier VistaTypeLP Legible typeface was withdrawn

For twelve days in August 2026 VistaType LP installed a **different** typeface,
**VistaTypeLP Legible**. It is gone, and VistaTypeLP Sans is not a renamed version of it.

**Why it was withdrawn.** It covered the Latin alphabet and the accents that go with it —
English, French, German, Spanish and Italian all set correctly. It did not cover Greek, the
phonetic alphabet, or most mathematical symbols, and Word does not tell you when a character
is missing: it silently borrows that one character from another typeface, at another size.
That turned up exactly where you would least want it — Latin, mathematics, phonetics, and the
Greek that runs all through medical material.

**What this means for you:**

- **Nothing to do.** Install the new version over the old one in the usual way.
- **Books you already made in it are safe, and are left alone.** Word stored a complete copy
  of the typeface inside each of those documents, so they still open, set and print exactly as
  they did. **If you attach the template to such a book again it will be reset into the
  typeface you choose in the dialog, and its page breaks will move.** VistaType LP used to
  shield those books automatically; that was removed on 9/4/2026 once it was confirmed the
  face never went beyond beta testers and no finished book was ever produced in it. If you
  do have one, do not re-attach the template to it.
- **The typeface is removed from your computer.** Leaving a face on the machine that VistaType
  LP no longer supports invites it being picked by hand for the very documents it cannot set.
  The removal finishes the next time you sign in to Windows — Windows will not release a
  typeface in the middle of a session — so if you still see it in Word’s font list today, sign
  out and back in.

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
| **Keep my toolbar exactly as it is, and add the VistaType LP and Braille Macros icons on the end** | Nothing of yours is moved, hidden, or reordered. VistaType’s eight icons are added after your own, with a divider before them. |
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
VistaType LP and Braille Macros icons on the end**. It puts back any of the eight that are
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

**The VistaTypeLP Sans typeface stays on the computer**, along with its license — see
*Uninstalling leaves the typeface behind* above.

**Your settings folder stays too.** `%AppData%\VistaType LP Settings` — your saved Word
settings, your three AutoCorrect lists and the error log — is deliberately left alone, so
reinstalling does not cost you any of it. Delete that folder by hand if you really want a
clean slate.

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
  it is, and add VistaType’s icons on the end.** It restores any of the eight that are
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

It was withdrawn, and **VistaTypeLP Sans** takes its place — see *The earlier VistaTypeLP
Legible typeface was withdrawn* above. The short version: Legible had no Greek, no phonetic
alphabet and almost no mathematics, and Word substitutes a missing character silently and at
the wrong size. VistaTypeLP Sans carries all three.

Books you have already produced in Legible are untouched and keep working: each one carries its
own copy of the typeface inside the file. **But attaching the template to one again will reset
it into Tahoma or VistaTypeLP Sans and move its page breaks** — the protection that used to
prevent that was removed on 9/4/2026, because the face only ever reached beta testers.

If Word still lists Legible after installing the new version, **sign out of Windows and back
in**. Windows does not release a personal typeface in the middle of a session.

### “VistaTypeLP Sans is grayed out when I attach the template”

Same cause, the other way round: Windows takes a personal typeface into use only at sign-in.
Close Word and reopen it; if it is still grayed, sign out of Windows and back in. See *If the
Typeface choice is grayed out* above.

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

**Need help?** Email **jerry@vistatypelp.org**.

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
