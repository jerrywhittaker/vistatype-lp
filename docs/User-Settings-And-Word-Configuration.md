# AutoCorrect: keeping the transcriber's own settings, and still configuring Word per document

A plan. Written 8/17/2026, revised the same day after Jerry's proposal of a saved-configuration
library.

**Status, 8/18/2026 — Piece 1 is built, and so is the spelling-and-grammar half of the fault
found while measuring.** The fifteen settings all three configurations wrote identically are
deleted from all three, so they are the transcriber's now in every kind of document. The five
spelling and grammar settings braille switches off are noted and put back, **and they survive
Word closing** — Piece 5's store is built, as `%AppData%\VistaType LP Settings\VistaType.ini`,
written with `System.PrivateProfileString` exactly as recommended below. Every remaining write in
the large print and braille configurations is guarded, which was not in this plan and came out of
the same measurement.

**Pieces 2 and 4 are NOT built.** The twenty settings that genuinely differ are still written to
fixed values on an ordinary document, and there is no saved-configuration library. Piece 3 is built
for the five spelling settings only; extending it to the twenty is Piece 2 and uses the same store.

Two things learned while building it, both worth keeping in mind for Piece 2:

- **The store had to be a file, not variables.** VBA's `End` statement resets every module-level
  variable, and this project runs `End` on ordinary paths — 34 times in `LPandBrlMacros`, plus
  dialog Cancel buttons and `Sh_Is_Doc_Open` when no document is open. Held in memory, one Cancel in
  a braille file lost the transcriber's values, and the next switch recorded VistaType's own
  switched-off values as her preference, permanently.
- **Order matters inside the configuration subs.** The settings a book takes away are now written
  *after* the line that records which configuration is in force. Otherwise a sub that raises part
  way through leaves them switched off while the record still says `DEF`, and the next save captures
  VistaType's values as hers by the same route.

The problem, reported by a beta tester running 3.0.135: she changes settings in Word's AutoCorrect
dialog, closes Word, reopens it — and her changes are gone, replaced by what VistaType considers the
default setup. Confirmed in the code. Opening or creating any document that is neither large print
nor braille runs `MS_Set_Word_Config_For_New_Install`, which writes nine AutoCorrect settings and
fourteen AutoFormat settings to fixed values. A blank document at the start of a session is enough
to do it.

**What must not be lost:** Word carries **one** set of AutoCorrect settings for the whole program,
not one per document, so a braille file and a large print book open together cannot both have their
typing behavior in force. Making the configuration follow the document (8/9/2026) is what fixed
that, and it stays.

Both are wanted and both can be had, because the settings braille and large print need to switch are
**not** the settings the transcriber wants to keep.

---

## The dividing line the whole plan rests on

Every setting in these three configurations belongs to one of two groups, and the difference is not
a matter of taste.

**Settings that belong to the reader.** Fractions, ordinals, symbols, Word's automatic bulleted
lists, numbered lists, tables, borders and heading styles, and capitalizing the first letter of a
sentence or a table cell. These decide what the DBT receives and what a low-vision reader sees on
the page. The braille configuration **adds** the AutoCorrect entries that turn `1/2` into `½`,
because that is what translates correctly; large print **deletes** them, because a large print book
should keep `1/2` as it was typed. Opposite requirements, and neither is the transcriber's to choose
— she is not the one they are for. An **ordinary** document now adds them as well (Jerry, 8/18/2026):
outside braille and large print they are simply Word's own behavior, and stripping them out of a
transcriber's own documents was never the add-in's business. **Twenty settings and eighteen
fraction entries are in this group, and they must keep switching per document.**

**Settings that belong to the transcriber.** "Replace text as you type", smart quotes, hyperlinks,
two initial capitals, days of the week, cAPS LOCK, tab-as-indent, and the rest. Nothing in braille
or large print depends on any of them. **Fifteen settings are in this group, and they should be hers
— in every document type, permanently.**

Measured by comparing the three `MS_Set_Word_Config_For_*` subs setting by setting. It is not a
judgment call: the fifteen are the ones all three configurations already write *identically*.

**Decision, 8/17/2026: no question is asked about braille or large print.** An earlier draft would
have offered to apply a saved configuration to braille and large print files too. Dropped, because
the only settings it could affect are the twenty in the first group — so "yes" would mean wrong
output, chosen at a moment when nobody can judge it. The fifteen that are genuinely hers already
apply everywhere, without asking. If a transcriber one day asks for a specific setting their own way
in braille, that is the moment to look again, with the example in hand.

---

## Part 1 — How it will work, written for the transcriber

### Word has one set of AutoCorrect settings, not one per document

Everything in **File → Options → Proofing → AutoCorrect Options** belongs to Word itself, not to the
document you happen to have open. Change smart quotes while writing a braille file and you have
changed them for every document you will ever open, including your letters and your grocery list.

That is a Word design decision, not a VistaType one, and it is the reason for everything below.

### Why VistaType changes some of them

Braille and large print need Word to type differently.

A braille file needs `1/2` to become `½`, because that is what the Duxbury translator turns into
correct braille. A large print book needs the opposite — `1/2` stays exactly as you typed it. And
neither can afford Word inventing bulleted lists, tables or borders as you type, because the styles
that control what the reader sees are chosen deliberately and Word's guesses undo them.

So when you move to a braille file, VistaType sets those behaviors for braille; when you move to a
large print book, it sets them for large print; and when you move to an ordinary document, it puts
them back. It follows whichever document you are actually looking at — click between a book and a
letter and it changes each time, without asking you and without moving anything on your screen.

### What is changing: your own choices are now yours

Until now, "puts them back" meant putting back **VistaType's** idea of a normal setup, which quietly
overwrote yours. From this version it means putting back **yours**.

**Fifteen settings are no longer touched at all.** They were being set the same way for braille,
large print and ordinary documents alike — so VistaType was never switching them, only overriding
you. Among them:

- **Replace text as you type** — the master AutoCorrect switch
- Correct TWo INitial CApitals
- Capitalize names of days
- Correct accidental use of the cAPS LOCK key
- Straight quotes with smart quotes
- Internet paths with hyperlinks
- Automatically use suggestions from the spelling checker
- Show AutoCorrect Options buttons
- Set left- and first-indent with tabs

Change any of those and VistaType will never change it back, in any document, in any session. They
are yours.

**The rest are remembered.** For the settings VistaType does have to switch — capitalizing the first
letter of a sentence or a table cell, and the AutoFormat As You Type items for fractions, ordinals,
symbols, bulleted and numbered lists, tables and borders — VistaType now notes how **you** have them
set while you are working in an ordinary document. When you move to a braille file or a large print
book it changes them, as before. When you come back to an ordinary document, it puts back what it
noted, not a factory default.

You do not have to do anything for this. There is nothing to run and nothing to remember.

### Saving your settings under a name

Some transcribers work to more than one house style, or want a way back after experimenting. So the
settings you have arranged can also be saved and named.

Three commands:

- **Save My Settings** — takes the AutoCorrect settings you have now, asks for a name and a short
  description ("Agency style, no smart quotes"), and keeps them.
- **Choose My Settings** — lists what you have saved, with the descriptions, and puts the one you
  pick into force. Also where you delete one you no longer want.
- **Reset Word's Settings** — puts everything back the way VistaType expects, for when Word is
  behaving oddly and you want a known starting point.

These are a convenience, not something you need. Your settings are already remembered without them.

**What a saved set covers:** the settings that are yours. Braille files and large print books keep
their own typing behavior whichever set you have loaded, because those settings decide what the
braille translator and the printed page receive, not what you prefer. The fifteen settings listed
above are yours everywhere already, so a saved set carries those into every document type too.

### One thing to expect the first time

VistaType has to learn your settings from somewhere, and the only honest place to start is what is in
force the first time it looks. On a machine where the old behavior has been running, that will be
VistaType's own defaults rather than your preferences.

**So the first time after updating, set the AutoCorrect options the way you want them once more.**
From then on they hold — across documents, across sessions, and across braille and large print work.
This is also a good moment to use **Save My Settings**, so you have a way back.

### If you ever want to see what is in force

**Doc Info** on either tab reports which configuration Word is currently set for — large print,
braille, or ordinary. That line describes the document in front of you, which is the point of it.

---

## Part 2 — The plan

### The shape of it

The insight that makes the everyday case free: **while an ordinary document is in force, the live
settings ARE the transcriber's own.** Nothing needs to be asked and no dialog needs to exist. Save
them on the way out, put them back on the way in.

Five pieces, in the order they should be built. The first three fix the reported fault and need no
new UserForm at all. The last two are Jerry's saved-configuration library, which sits on top and is
easier to build once the first three exist — by then the add-in already knows which settings are the
transcriber's, and a saved configuration is simply a named copy of them.

### Piece 1 — the fifteen that were never being switched

All three configurations write these identically, so removing them changes no switching behavior
whatsoever. There is nothing to switch between:

```
replaceText                                  True     <- "Replace text as you type"
CorrectInitialCaps                           True
CorrectDays                                  True
CorrectCapsLock                              True
AutoFormatReplacePlainTextEmphasis           False   <- see note below
ReplaceTextFromSpellingChecker               True
DisplayAutoCorrectOptions                    True
TabIndentKey                                 True    <- left the list on 8/18/2026, see below
AutoFormatAsYouTypeReplaceQuotes             True
AutoFormatAsYouTypeReplaceHyperlinks         True
AutoFormatAsYouTypeReplacePlainTextEmphasis  False
AutoFormatAsYouTypeApplyHeadings             False
AutoFormatAsYouTypeDefineStyles              False
AutoFormatReplaceQuotes                      True
AutoFormatReplaceHyperlinks                  True
```

Delete all fifteen, from all three subs. On its own this fixes the loudest half of the complaint.

**`TabIndentKey` left this list the same day.** Jerry: braille must have "Set left- and
first-indent with tabs and backspaces" switched OFF. All three configurations did write it
True, so the measurement was right — but identical today does not mean identical tomorrow. It
is now a switched setting, noted and put back through the store like the five spelling ones,
rather than forced off in her own documents for good.

**Two corrections made on 8/18/2026, when this was built.** `AutoFormatReplacePlainTextEmphasis`
joined this list that morning: the default configuration wrote it True and a clean install has it
False, so correcting the default made all three agree. `CorrectKeyboardSetting` left the list and
is now written by none of the three at all — it follows the keyboard to the language being typed,
which belongs to a multilingual transcriber and to nothing in braille or large print.

**And a trap this list nearly walked into.** Three of the fifteen — plain text emphasis, quotes and
hyperlinks in the *on-demand* AutoFormat group — were not passive preferences. They were the
arguments to Word's actual **AutoFormat command**, which `Dx_Fix_Common_File_Errors` ran over the
whole document as its first step. Deleting them would have meant braille file cleanup converting
`*word*` to bold and **deleting the asterisks** on any machine where that setting was on. The
AutoFormat call was removed instead (Jerry, 8/18/2026) — it also ate empty paragraphs, exactly as
the large print copy did before 8/13/2026 — and with it gone, all fifteen really are the
transcriber's.

Deleting rather than keeping them is right for a second reason: each is a write to a setting that
roams, and it was that burst of writes which triggered Office's "restart to apply your privacy
settings" notice on 7/18/2026.

Three **display** settings are also written identically — `FormattingShowNextLevel`,
`DisplayRulers`, `ShowAll` — but they are a different question, already governed by
`Sh_Config_Skip_Display`, and they stay exactly as they are.

### Piece 2 — save and restore the twenty that do differ

The complete list. The helpers below work over exactly it:

| Setting | Ordinary | Large print | Braille |
|---|---|---|---|
| `CorrectSentenceCaps` | True | False | False |
| `CorrectTableCells` | True | False | False |
| `AutoFormatAsYouTypeReplaceFractions` | True | False | True |
| `AutoFormatAsYouTypeReplaceOrdinals` | True | False | True |
| `AutoFormatAsYouTypeReplaceSymbols` | True | False | True |
| `AutoFormatAsYouTypeApplyBorders` | True | False | False |
| `AutoFormatAsYouTypeApplyBulletedLists` | True | False | False |
| `AutoFormatAsYouTypeApplyNumberedLists` | True | False | False |
| `AutoFormatAsYouTypeApplyTables` | True | False | False |
| `AutoFormatAsYouTypeFormatListItemBeginning` | True | False | False |
| `AutoFormatApplyHeadings` | True | False | False |
| `AutoFormatApplyLists` | True | False | False |
| `AutoFormatApplyBulletedLists` | True | False | False |
| `AutoFormatApplyOtherParas` | True | False | False |
| `AutoFormatReplaceFractions` | True | False | True |
| `AutoFormatReplaceOrdinals` | True | False | True |
| `AutoFormatReplaceSymbols` | True | False | True |
| `AutoFormatPreserveStyles` | True | True | False |
| `AutoFormatPlainTextWordMail` | True | False | False |

Two new shared helpers, `Sh_` because both sides use them:

- **`Sh_Save_User_Typing_Settings`** — reads each setting in that list and stores it. Called only
  when what is in force is the transcriber's own.
- **`Sh_Restore_User_Typing_Settings`** — writes them back, keeping the
  `If .X <> value Then .X = value` guard so an unchanged setting is not written.

In `MS_Set_Word_Config_For_New_Install` the twenty sit in three places: the `With Options` block of
AutoFormat-As-You-Type settings, the `With AutoCorrect` block (`CorrectSentenceCaps` and
`CorrectTableCells`), and the second `With Options` block of plain AutoFormat settings. Work from the
table and count what you remove — twenty out, one call in.

**Where the helpers hook in, and this matters.** Not in `Sh_Apply_Word_Config`. About a dozen places
run the three `MS_Set_Word_Config_*` subs **directly**, including both cleanup sequences, and
anything put only in the funnel would be bypassed by all of them. It is the same reasoning that
already made the three subs write `Sh_ConfiguredAs` themselves rather than trusting their callers.

So:

- At the **top of `MS_Set_Word_Config_For_Large_Print`** and **`..._For_Braille`**: if the ordinary
  configuration is what is currently in force, call `Sh_Save_User_Typing_Settings` **before**
  imposing anything. That is the last moment the transcriber's values still exist.
- In **`MS_Set_Word_Config_For_New_Install`**: replace the twenty fixed writes with a call to
  `Sh_Restore_User_Typing_Settings`. If nothing has ever been stored, it writes today's values and
  stores them — which is what makes the first run after updating behave sensibly.

### Piece 3 — remembering across sessions what was left in force

`Sh_ConfiguredAs` is empty when Word starts, but Word's settings are not — they are whatever the last
session left. Without fixing this, a session that ends in a braille file starts the next one with
braille settings in force while the add-in believes nothing is configured, and would then save
braille's values as the transcriber's.

So the store needs two things, both surviving Word closing:

1. **The transcriber's values** for the twenty settings.
2. **Which configuration was last imposed** — written by the three subs at the same point they
   already set `Sh_ConfiguredAs`.

At startup, seed `Sh_ConfiguredAs` from the stored marker instead of leaving it empty. Then the "is
what is in force the transcriber's own?" test is honest on the first document of a session, which is
exactly the case that is broken today.

### Piece 4 — the saved-configuration library

Jerry's proposal, kept as a convenience layer on top rather than as the mechanism. Three commands,
and a store that holds any number of named configurations.

**What a saved configuration holds:** the fifteen settings from Piece 1 and the twenty from Piece 2,
a name, and a description. Nothing else — and in particular **not** the eighteen fraction entries,
which stay managed per document with no opt-out.

**Loading one** writes the fifteen into Word directly (they are hers everywhere, so there is nothing
to negotiate) and writes the twenty into the same store `Sh_Restore_User_Typing_Settings` reads —
then applies them if an ordinary document is in force. A braille file open at the time keeps braille's
typing behavior, and picks up the loaded values when the transcriber next returns to an ordinary
document. No flag, and no suspension of the per-document configuration: with Pieces 1 to 3 in place
the automatic behavior is no longer fighting her, so there is nothing to switch off.

**Three commands, three dialogs' worth of work:**

| Command | What it does |
|---|---|
| Save My Settings | Reads what is in force, asks for a name and description, stores it. Offers to overwrite a name that already exists. |
| Choose My Settings | Lists names with descriptions, loads the chosen one. Delete and rename live here too — one dialog, not three. |
| Reset Word's Settings | Writes VistaType's own defaults, and says so plainly. The support tool for a Word that is behaving oddly. |

Worth building **Reset Word's Settings first**, before the other two: it is the smallest, it needs no
list and no store, and it is the one that earns its keep the next time a tester's Word misbehaves.

### Piece 5 — where the store lives

One store for everything above: the current values, the last-imposed marker, and the named
configurations.

**Recommended: a small ini file under `%AppData%\VistaType LP Settings\`,** read and written with
Word's own `System.PrivateProfileString` so no file handling has to be written.

Two reasons for that folder name rather than `%AppData%\VistaType LP`:

- **The uninstaller deletes `%AppData%\VistaType LP` wholesale**, and a reinstall must not lose the
  transcriber's settings. This is the same reasoning that already puts `OFL.txt` in
  `%AppData%\VistaType LP Fonts` — see `installer/vistatype.iss` and the note in `CLAUDE.md`.
- A file **can be copied to a new machine**; a registry entry effectively cannot. Transcribers do
  get new computers.

The installer will therefore need this folder left out of `[UninstallDelete]` — deliberately, and
with a comment saying why, or someone will tidy it back in.

**Neither the registry nor an ini file can list its own sections**, so the store needs an index entry
naming the saved configurations. Keep that index in one place and let one helper own it.

### Found while measuring this: five settings that are turned off and never given back

Not AutoCorrect, and not what she reported, but the same fault in a different place, and it should be
decided at the same time.

The braille configuration turns these off:

```
Options.CheckGrammarAsYouType = False
Options.IgnoreMixedDigits     = False
Options.ContextualSpeller     = False
Options.LabelSmartTags        = False
Options.IgnoreUppercase       = False
```

and large print turns off the last two. **The ordinary configuration sets none of them, ever.**

So the moment a transcriber opens one braille file, grammar checking as you type is off — in that
braille file, in her ordinary documents, in every document she opens for the rest of that session,
and in every session afterwards, because Word remembers it. Nothing in the add-in ever turns it back
on, and she has no reason to connect it to VistaType.

Switching them off is right for braille: the grammar checker and the contextual speller flag braille
formatting endlessly. The fault is only that they are never restored.

**Decision needed:** the consistent answer is to treat these exactly like the twenty — note what she
had, and put it back on return to an ordinary document. Five more entries in the same list, no new
mechanism. The alternative, having the ordinary configuration turn them all on, is another fixed
value overriding her choice and repeats the original mistake.

### What must not change

- The switching behavior itself — `Sh_HandleDocumentActivated` and its guards, including the
  screen-updating test that keeps all of this off the macros' backs.
- The large print and braille configurations. Their values are unchanged; only where the *ordinary*
  values come from changes.
- The display side, and `Sh_Config_Skip_Display` with it. Switching documents still leaves the
  transcriber's screen alone.
- The eighteen fraction AutoCorrect **entries**. Braille and ordinary documents add them, large print
  deletes them, and they keep switching per document. Worth a line in the guide that these particular entries are
  managed, in case a transcriber ever adds one by hand.

### What to test in Word, in order

1. **The reported fault.** Ordinary document → turn off "Capitalize first letter of sentences" and
   turn off "Replace text as you type" → close Word → reopen → open a blank document. Both still as
   she left them.
2. **Switching still works.** With that ordinary document open, open a braille file. Fractions
   should give `½`, capitalization as braille needs it. **Doc Info** says braille.
3. **Coming back.** Click to the ordinary document. Sentence capitals off again, "Replace text as you
   type" still off. `1/2` now becomes `½` here too, as of 8/18/2026 — see the fractions note below.
   Doc Info says ordinary.
4. **Large print in the mix.** Three documents open — ordinary, braille, large print — clicking round
   all three. Fractions on in braille and in the ordinary document, off in large print.
5. **The master switch is never touched.** With "Replace text as you type" off, work in braille and
   large print and come back. Still off, every time.
6. **A macro run does not disturb it.** Full File Cleanup and Selection Cleanup on both tabs — both
   call the configuration subs directly, which is why they are worth testing specifically — then
   check the ordinary document's settings again.
7. **Word restarted from a braille file.** End a session with a braille file open, reopen Word, open
   an ordinary document. Her settings, not braille's.
8. **The privacy notice does not come back.** Several new documents in a row on a machine where
   Office is unlicensed or disconnected — that is where the 7/18/2026 notice appeared.
9. **The library.** Save a configuration, change several settings, load it back. Then load it while a
   braille file is active and confirm braille's typing behavior is untouched until an ordinary
   document is activated. Then delete it.
10. **Reset.** Run Reset Word's Settings and confirm it says what it did, and that the following
    document switch still behaves.

### Traps

- **The first run after updating cannot know her preferences.** It stores what is in force, which on
  an existing install is the add-in's own defaults. She sets them once more and they hold. Say so in
  the release note; it is the one part of this that needs explaining to a person.
- **Never save while a macro is running.** The guard exists in `Sh_HandleDocumentActivated`, but the
  save sits inside the LP and braille subs, which macros call directly — so the save must test that
  the ordinary configuration is genuinely in force rather than assume its caller is a transcriber.
- **The temporary document.** A scratch document created by a cleanup macro has `Normal.dotm`
  attached and is therefore an "ordinary document" by every test the add-in has. Saving settings
  while one is on screen is harmless today, because a macro will not have changed AutoCorrect behind
  the transcriber's back — but it is exactly the kind of thing that becomes wrong later, so skip the
  save when screen updating is off, the same marker used everywhere else.
- **Keep the `If .X <> value Then .X = value` guards.** They are what stopped the privacy notice.
- **One list, not several.** The settings to save and restore must be the same list in both helpers
  and in the library. Write them over one shared list so they cannot drift, and remember that adding
  a setting to any of the three configurations later means adding it there too.
- **New UserForms cannot be hand-written from Linux.** The binary `.frx` has to be created in Word —
  `tools/windows/New-UserForm.ps1`, and never `make pull` to seed one. See `CLAUDE.md`.
- **New ribbon buttons are append-only.** Every `btn_*` id already shipped is referenced by toolbars
  installed in the field; `installer/ribbon-button-ids.txt` records them and `make build` stops if
  one disappears. Add ids, never rename or remove.
- **Three new commands need somewhere to live.** Deciding whether they go on both tabs, one tab, or
  behind a menu is Jerry's call, and it affects the ribbon and the curated toolbar together.
