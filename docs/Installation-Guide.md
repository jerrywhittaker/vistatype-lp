# VistaType LP + Braille Macros — Installation Guide

**Version 3.0**

VistaType LP adds tools to Microsoft Word for producing **large-print** documents and
**braille** source files (for the Duxbury Braille Translator). This guide covers
installing it with the new **one-click installer**.

> If you used an older version, you no longer copy files by hand or export your Quick
> Access Toolbar first. The installer does everything, and it **does not disturb your
> existing ribbon or Quick Access Toolbar**.

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

That’s it. Open Word and look for the two new ribbon tabs.

---

## What the installer set up

So you know exactly what changed on your computer:

| It installed / changed | Where | Notes |
|---|---|---|
| The macros (`LPandBRL.dotm`) | Word **STARTUP** folder | Loads automatically every time Word opens |
| The large-print template (`LargePrintTemplate.dotx`) | Your **Templates** folder | Holds the large-print styles |
| Two ribbon tabs: **VistaType LP** and **Braille Macros** | Word ribbon | **Added next to your existing tabs** — your ribbon is not replaced |
| Quick Access Toolbar icons | Your QAT | **Added to** your toolbar — your own icons are kept |
| Trust settings so macros run | Word Trust Center | Marks the STARTUP folder as trusted (including network/roaming profiles) |
| Removed leftover files from old versions | Templates folder | Cleans up the obsolete “Large Print Templates” folder |

---

## Confirm it’s working

1. Open **Microsoft Word** (a blank document is fine).
2. Look along the ribbon for the **VistaType LP** and **Braille Macros** tabs.
3. Click **VistaType LP → Help → Version & Updates**. If a message appears, the macros
   are running correctly.

If the tabs are there and buttons respond, you’re ready to go.

---

## Your ribbon and Quick Access Toolbar are safe

Unlike older versions, this installer **merges** its buttons in rather than overwriting
your customizations:

- Your own ribbon tabs and groups are untouched.
- Your own Quick Access Toolbar icons stay exactly where they were.
- You do **not** need to export or back up your QAT before installing.

### Restoring a Quick Access Toolbar icon you removed

If you accidentally remove one of the VistaType icons from your Quick Access Toolbar:

1. Go to the **LP and BRL QAT Icons** tab.
2. **Right-click** the icon you want back.
3. Choose **Add to Quick Access Toolbar**.

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
Toolbar icons** — your own ribbon and toolbar customizations are left in place.

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
- **Still missing one?** Add it back from the **LP and BRL QAT Icons** tab: right-click
  the icon → **Add to Quick Access Toolbar**.

### “The VistaType LP / Braille Macros tabs don’t appear”

- **Look for a Security Warning bar** just below the ribbon that says macros were
  disabled — click **Enable Content**. (After the installer’s trust setup this shouldn’t
  keep coming back.)
- **The add-in may have been disabled by Word** (this can happen after a Word crash):
  *File → Options → Add-ins*. At the bottom, set **Manage: Disabled Items → Go**, and if
  `LPandBRL.dotm` is listed, select it and click **Enable**. Then restart Word.
- **Confirm the add-in is loaded:** *File → Options → Add-ins*, set **Manage: Word
  Add-ins → Go**, and make sure `LPandBRL.dotm` is checked under Global Templates.

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
> Add the ones you want from the **LP and BRL QAT Icons** tab (right-click → *Add to
> Quick Access Toolbar*). You may also need to confirm the STARTUP folder is a trusted
> location in the Trust Center.
