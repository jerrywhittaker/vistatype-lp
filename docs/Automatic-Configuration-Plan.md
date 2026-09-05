# Automatic Word configuration: the plan

Agreed by Jerry and the beta tester, 8/20/2026, from *Automatic Configuration Table.docx*.
Written up here before any code was touched.

**Status, 9/5/2026.** Pieces 1, 3, 4, 5 and 7 are built. Piece 6 was settled 9/4/2026 and needed
no code. **Piece 2 was built, tested and REJECTED by Jerry on 9/4/2026** — see its section below,
which is now a record of why it must not be built again rather than a plan to build it.

**DECISION 3 WAS REVERSED BY JERRY ON 9/5/2026, and half of Piece 4 came out with it.** A setting
the transcriber changes inside a large print or braille book is **not** a preference and never
reaches the user's letters — a book's settings live and die with the book being on screen. The learning
comparison in `Sh_Restore_Transcriber_Settings` is gone and `[BookApplied]` is now a flag and
nothing else. Read Decision 3 and Piece 4 before touching the ledger; the mechanism that was
removed is written out there so it does not get rebuilt.

This supersedes the unbuilt half of
`User-Settings-And-Word-Configuration.md` — Pieces 2 and 4 of that plan are replaced by
Piece 4 here, which answers a question the earlier plan never asked. Piece 1, 3 and 5 of
that document are built and stay as they are.

## The rule, in one sentence

**Large print and braille are authoritative and self-restoring. Everything else is hers, and
VistaType never overwrites it.**

Open a book and Word is configured for that book, every time, no matter what she did last.
Open anything else and Word is exactly as she left it.

## The table, as agreed

| | Large Print | Braille | Neither |
|---|---|---|---|
| After attaching a template | LP | Braille | Default — Normal, or some other template |
| Can she reconfigure the current document | yes | yes | yes |
| Re-attaching a template | resets to LP | resets to Braille | no — keeps her changes |
| Opening the document again | resets to LP | resets to Braille | no — keeps her changes |
| A new document | LP | Braille | keeps her changes |
| Opening a document | LP if it is LP | Braille if it is braille | keeps her changes |

Type is decided by the attached template alone: `LargePrintTemplate.dotx` is large print,
`BANA Braille*.dot*` is braille, anything else is neither.

## The seven decisions behind it

Each of these was asked and answered on 8/20/2026. They are written down because several of
them reverse something the code does deliberately today, and the reversal is only safe if the
reason is on the record.

1. **Switching between open documents never reconfigures Word.** Clicking from a book to a
   letter leaves the book's configuration in force. This reverses the 8/9/2026 "the
   configuration follows the document" work.
2. **The display side is in scope, not only AutoCorrect and AutoFormat.** An ordinary document
   gets both rulers on and print view; it gets formatting marks, the Styles pane and the
   navigation pane switched **off**. Revised 8/20/2026, later the same day: the first answer
   forced formatting marks on and left the Styles pane alone, and neither survived the question
   of what a plain letter should actually look like when it opens.
3. ~~**Reopening a book re-applies its configuration**, and a setting the user changes *inside* a book
   still counts as the user's for the user's ordinary documents.~~ **THE SECOND HALF IS REVERSED — Jerry,
   9/5/2026.** Reopening a book re-applies its configuration; that part stands. But a setting
   changed inside a book is **not** the user's and never reaches the user's letters. His words:

   > "changes in the word configuration settings for a braille file or a large print file should
   > have no effect on the word configuration settings for a 'letter'. Changes to the word
   > configuration settings for a large print or braille file are limited to that file and only
   > while it is open... the settings are for here and now and never get changed for any other
   > file or file type. Changes to the word configuration settings for a letter essetially means
   > that the user has decided that this how word will behave for 'letters' from that point
   > onward until they are changed again."

   **A book's settings live and die with the book being on screen. Only what the user changes in a
   letter is remembered.** The original reasoning — that the user does far more in Word than large
   print and braille and does not stop being an ordinary user when the user opens a book — was
   answered on its own terms: what the user does inside a book is for that book, here and now.
   Built out on 9/5/2026; see Piece 4.
4. **Document type is read from the attached template**, not from the `Box Black` style. That
   style test predates knowing how to test for an attached template. It is kept for one other
   job — see Piece 1.
5. **An ordinary document still gets the 18 compact fractions.** Large print deletes them, so
   returning to a letter has to put them back or she has lost them for the session.
6. **The Styles pane's sort and its filter are left alone in ordinary documents.** Both are
   saved inside her file, so forcing them overwrites what that document was carrying. The pane's
   *visibility* is a separate matter and is not left alone — see Decision 2.
7. **There is no first-run configuring pass.** Thirty-two of the thirty-four default settings
   already match Word's factory values, so on a fresh machine the only thing that needs doing
   is adding the fractions.

## What is actually changing

The two book configurations barely move. Almost all of this is about the third case, which is
where she spends most of her time and which VistaType currently treats as a configuration to
be imposed rather than as her own Word.

### Piece 1 — type is read from the attached template

`Sh_Doc_Config_Type` currently asks `Lp_Is_The_Attached_Template_LP`, which tests for the
`Box Black` style. It becomes a test of the attached template's name. The braille half already
works this way (`InStr(attached, "BANA Braille") > 0`) and needs nothing.
`LargePrintTemplate.dotx` is the only large print template name the code has ever known.

**The style test is kept, for a different question.** `Sh_HandleDocumentOpened` uses it today
for two jobs at once: *is this large print*, and then separately *is its template obsolete* —
LP book, but attached to `Normal.dot` or to anything that is not `LargePrintTemplate.dotx`, so
she gets the warning with the font size and margins and is sent to the attach dialog. Move the
type question to the template name and that branch can never fire again: the outer test would
already have said "not large print", and a book on an old template would quietly get the
ordinary configuration with no warning that it needs re-attaching.

So the two questions get two tests, and the function is renamed so its name says which it
answers:

- **Which configuration does this document get?** — the attached template is
  `LargePrintTemplate.dotx`.
- **Was this document ever made as a large print book?** — `Box Black` is present. Used only
  to catch a book whose template is missing or obsolete.

Ten callers in `LPandBrlMacros` and one in a form have to be triaged one at a time: each is
asking one of those two questions, and which one is not always obvious from the call site.

### Piece 2 — REJECTED BY JERRY, 9/4/2026. Do not build it.

It was built and reverted the same day. **The rule is the opposite of what this section proposed:
the configuration follows the document type on the screen, and it changes when she clicks into a
document of a different kind.** Jerry, on being shown it working:

> "The word configurations are to change when a different type of document is click. the
> configuration for the 'letter' should not be forced on the braille. Configurations should match
> the document type on the screen."

**What it would have done.** `Sh_HandleDocumentActivated` would stop calling
`Sh_Apply_Word_Config`, so a configuration would follow OPENING or CREATING a document and
nothing else.

**Why he said no, and it is the sentence the plan itself had already written down as "the cost".**
With a braille file and a letter open together, whichever was opened last held Word — so working
in the braille file under the letter's settings was a normal Tuesday, not an edge case. The plan
called that "real, and the price of the rule". It is not a price he is willing to pay, and the
reason is the work rather than the principle: a braille file typed under a letter's settings gets
Word's capitalization and its fractions, which is wrong in the file that goes to Duxbury.

**Three things the attempt was worth, all of them kept:**

- **A settings-loss fault it exposed in `MS_Reset_Word_Configuration`**, which is fixed and stays
  fixed. Step 3 marked the book record spent on the strength of the DOCUMENT being ordinary; it
  now also requires that no book configuration is in force. Those two agree while switching
  reconfigures, so it is a second net — but on a machine where Word does not raise the window
  events they can disagree all session, and the cost of getting it wrong is her own settings,
  permanently. That is the 8/18/2026 fault.
- **Document Settings could print a blank line** where "Word is configured for…" belongs, because
  `MS_Word_Config` is wiped by any `End` statement. It goes through `Sh_Word_Config_Line` now.
- **Two review findings that were only true under Piece 2 and are now moot**: the Reset button's
  hover text going stale on a switch, and Styles Pane: Recommended half-working. Both were
  symptoms of the switch not configuring, and both went back to normal with the revert.

**If this is ever reconsidered**, read the whole of this section first, and know that two things
the original text asserts are false: `Sh_Apply_Word_Config`'s `DisplayToo` argument and
`Sh_Config_Skip_Display` **cannot** be retired with the switch. `MS_Reset_Word_Configuration`
passes `DisplayToo:=False` deliberately, and `Dx_Attach_BANA_Template_Run` raises the flag across
the attach so the braille display changes at its stage four and not before (Jerry, 8/29/2026).

### Piece 3 — the ordinary configuration stops configuring

`MS_Set_Word_Config_For_New_Install` loses its twenty fixed Options and AutoCorrect writes.
It keeps exactly three jobs:

- restore her settings from the ledger (Piece 4),
- `Sh_Add_Compact_Fractions`,
- the display block, revised as below.

Of its eight display writes, four stay, one flips, three go, and one is added:

- **Stay** — `DisplayRulers`, `DisplayVerticalRuler`, `View.Type = wdPrintView`, and the Styles
  pane hide.
- **Flips** — `View.ShowAll` becomes `False`. A plain letter opens without pilcrows.
- **Go** — `StyleSortMethod`, `FormattingShowFilter`, `FormattingShowNextLevel`. All three are
  saved inside her file, so forcing them overwrote what that document was carrying: the
  7/24/2026 complaint, in her letters rather than her books.
- **Added** — `ActiveWindow.DocumentMap = False`, closing the navigation pane.

It should be renamed. It is not a new-install configuration any more and has not been one since
Decision 7 — something like `MS_Set_Word_Config_For_Ordinary_Document`.

### Piece 4 — the ledger grows, and learns to tell her changes from ours

This is the large piece and the one with no equivalent in the earlier plan.

**Today** `Sh_Save_Transcriber_Settings` notes six settings, and only while the ordinary
configuration is genuinely in force — guarded on `Sh_ConfiguredAs`, precisely so VistaType
never records its own value as her choice. Decision 3 makes that guard wrong: a change she
makes inside a book has to reach her letters, and today it cannot.

**What it has to become.** Two sections in
`%AppData%\VistaType LP Settings\VistaType.ini`:

- `[TranscriberSettings]` — her preferences, as now, but covering everything a book writes:
  roughly 22 settings for large print and 25 for braille, plus the three application-wide
  display settings in Piece 5 and the fractions.
- `[BookApplied]` — for each of those settings, the value the book configuration most recently
  wrote.

**BUILT, THEN HALF OF IT REMOVED — Jerry's hard wall, 9/5/2026.** What follows is what was built
on 8/20/2026 and is kept as the record of a mechanism that must not come back:

> How the user's changes are told apart from VistaType's, with no heuristics. When a book configuration is
> applied, record into `[BookApplied]` the value it set for every tracked setting. When an
> ordinary document takes over, compare live against `[BookApplied]`, setting by setting: **same**
> means the user did not touch it and the stored preference is left alone; **different** means the user
> changed it while working, and that becomes the user's new preference in `[TranscriberSettings]`.

**That comparison is gone.** Under Decision 3 as revised, a change made inside a book is not a
preference at all, so there is nothing to learn and nothing to tell apart. The learning loop came
out of `Sh_Restore_Transcriber_Settings`, which now does one job: restore from
`[TranscriberSettings]`, then clear the book record.

**The flag is raised at the TOP of each book configuration, not the end — 9/5/2026.** Review found
a hole that had been open since 8/20/2026 and was not caused by this change: written last, a book
configuration that *raised part way through* never raised the flag at all. Thirty of that book's
values were already standing in Word with no record that a book had put them there, so the next
ordinary document read them as the user's own and stored them permanently — the 8/18/2026 fault by
a third route, and silent, because `Sh_Apply_Word_Config` swallows the error. The flag now goes up
before the first setting is written, immediately **below** the guarded
`Sh_Save_Transcriber_Settings` and never above it (that save declines while the flag stands, so
raising it first would discard the change the user had just made in a letter). A configuration that
dies now leaves the flag standing, and saving stays suspended until a letter clears it — which is
correct, because after a failure nobody knows whose values are in Word.

**Do NOT bump `VT_STORE_STAMP_NOW` for any of this.** It stamps `[TranscriberSettings]`, whose
shape did not change — `Sh_Tracked_Settings` is untouched at 42 names. Bumping would re-impose the
starting AutoCorrect list on every machine in the field, and worse, the stamp-mismatch path is
itself a known route to the 8/18 fault, so bumping for tidiness could cause the very thing this
work prevents.

**`[BookApplied]` survives as a FLAG and nothing else.** `Sh_Note_Book_Settings` writes
`Saved=1` and no longer writes the 42 per-setting values — the learning loop was their only
reader, checked name by name. The flag itself is load-bearing and must not be removed: it is what
`Sh_Save_Transcriber_Settings` tests before saving, and therefore what stops Word being quit
inside a braille file and `AutoExec` recording braille's values as the user's preferences the next
morning. That is the 8/18/2026 fault and the flag is the whole wall.

**What still catches a change the user makes in a letter**, since the restore no longer learns: both
book configurations call `Sh_Save_Transcriber_Settings` at their top, guarded by
`If Sh_ConfiguredAs = "DEF"`. The user's live values are written down as the user's own in the moment before a book
takes Word over. Verified 9/5/2026 when the learning loop was removed — without it there would be
a hole, and there is not.

Two traps in that, both of which will bite silently:

- **Record what the configuration *targets*, not what it *wrote*.** Every write in the three MS_
  subs is guarded with `If <> ` — a setting already at the target is not written at all. Record
  the target value regardless, or a setting that happened to match beforehand looks like one the
  book never touched.
- **Write `[BookApplied]` after the configuration completes, not as it goes.** A sub that raises
  half way through would otherwise leave a record claiming settings it never reached, and the
  next return to an ordinary document would read her own values as the book's.

### Piece 5 — the display side

Most of the display side already behaves the way the table asks, for free, because of where
Word keeps it:

- **Per window** — view type, both rulers, style area width, formatting marks. Each open
  document has its own; they cannot leak between documents.
- **Saved in the document file** — style sort, "Select styles to show", `FormattingShowNextLevel`,
  `FormattingShowUserStyleName`. They travel with the book to any machine.
- **Application-wide** — the Styles pane, `ShowStylePreviews`, `RestrictLinkedStyles`. Only
  these three behave like the typing side.

**Two of the three need nothing, because Decision 2 settles them.** The ordinary configuration
closes the Styles pane and closes the navigation pane, so there is no earlier state to restore —
VistaType simply states what a plain letter looks like. That is worth being explicit about: it
means a pane she opens herself in a letter is closed again the next time she opens one. Decided,
not overlooked.

**The other two do need the ledger.** Large print sets `ShowStylePreviews` and
`RestrictLinkedStyles` to True and braille sets `RestrictLinkedStyles` to True. Neither
configuration mentions them on the way back, and **nothing has ever put them back** — the same
shape as the five spelling settings found on 8/18. They join Piece 4's tracked set.

**One more, found while writing this up.** `Dx_Attach_BANA_Template` runs
`CommandBars("Navigation").Visible = False`, which is application-wide, unlike the
`ActiveWindow.DocumentMap` line beside it. Under Decision 2 that is no longer a leak — the
ordinary configuration closes the navigation pane anyway — so the two lines simply need to agree
about which one does the closing. `DocumentMap` is the modern Navigation Pane and is per window;
`CommandBars("Navigation")` is the older Document Map task pane.

**Done, and NOT the way this said.** `Sh_Set_Navigation_Pane` writes **both** halves deliberately:
`DocumentMap` alone does nothing during `DocumentBeforeClose`, which is measured and recorded at
that sub. Do not "finish" this by dropping the command-bar write.

Where each configuration lands after this:

| | Large Print | Braille | Ordinary |
|---|---|---|---|
| View type | Print | Draft | Print |
| Formatting marks | on | — | on |
| Horizontal ruler | on | on | on |
| Vertical ruler | on | off | on |
| Style area width | 24.5 | 64.5 | — |
| Styles pane | on | off | **off** |
| Navigation pane | off | off | **off** |
| Style sort | — | Recommended | **left alone** |
| Select styles to show | — | Recommended | **left alone** |
| Style previews | on | — | **restored from the ledger** |
| Restrict linked styles | on | on | **restored from the ledger** |

### Piece 6 — SETTLED 9/4/2026, and it cost no code

**Jerry chose option 1: a document a macro makes for the transcriber is configured by the macro
that made it.** Option 2 had died that morning anyway — `Lp_Copy_To_Temp_Doc` was removed when
Change Picture Color came off the scratch document, so there was no "for the transcriber" route
left in it to configure anything.

It was written as preparation for Piece 2, which was then rejected — but **it stands on its own**.
The pick-up when her cursor reaches such a document sets the TYPING side only (`DisplayToo`
False), so a document left to it never gets formatting marks, the rulers, Print view or the
Styles pane.

**The gap turned out to be closed already.** Checking site by site, every document a macro makes
for the transcriber is configured, so the decision is a rule to hold to rather than a change to
build:

| What makes it | Screen at that moment | What configures it |
|---|---|---|
| `Dx_Attach_BANA_Template_Run`, blank document when nothing is open | OFF, so `Sh_HandleDocumentNew` skips it | the macro itself, `MS_Set_Word_Config_For_Braille`, twice |
| `Lp_Attach_Lp_Template`, blank document when nothing is open | ON | `Sh_HandleDocumentNew` as an ordinary document, then `Lp_Attach_The_Template` applies the large print configuration |
| `Sh_Convert_XML_File_To_Word_Document`, the converted book | ON — screen updating goes off 22 lines later | `Sh_HandleDocumentNew`. **A dependency on line order, not a guarantee**, and marked as such at that `Documents.Add` |
| `Sh_Copy_Ref_Pg_Tags_To_Temp_File`, the `$pg` validation list | OFF | the macro itself — braille or large print, matching the book it was made from |
| the two export macros | OFF — and **that**, not the `Visible:=False`, is what makes the handler skip them; Word raises NewDocument for a hidden document too | nothing, and nothing is needed: they are saved and closed inside the macro and she never sees them |

**What was written down rather than built:** the rule now stands in `Sh_HandleDocumentNew` with
the site-by-site audit beside it, and each creation site says what configures its document and
what would break if the screen were turned off around it.

### Piece 7 — the Styles pane when a book closes: nothing to do

This was an open question until Decision 2. She has a book and a letter open, the pane is up,
and she closes the book: `Sh_HandleDocumentClosing` takes the pane down, and the recovery in
`Sh_HandleDocumentActivated` only puts it back if what is on screen is a large print document.
While the ordinary configuration was going to leave the pane alone, that pulled the pane out of
her letter and needed fixing.

Now that a letter is defined as having no Styles pane, the pane coming down when the book closes
is simply correct — and it is what the 8/20/2026 code already does. **No change.** Recorded here
so the next person to look at it does not re-open the question.

## Build order

1. **Piece 1** — independent of everything else, and the smallest.
2. **Piece 7** — independent, finishes the 8/20 work.
3. **Piece 4** — the ledger, built and tested while the old behavior is still in place. Nothing
   depends on it yet, so it can be measured on its own.
4. **Piece 3** — flip the ordinary configuration over to the ledger. Piece 4 must be right first.
5. **Piece 6** — the decision. Made 9/4/2026, option 1, and it needed no code. Then
6. ~~**Piece 2**~~ — **rejected by Jerry on 9/4/2026 after being built and shown to him.** See its
   section above. The plan ends at five pieces and a rule.

## What must not change

- **The four `AutoAdd` properties and the four AutoCorrect exception lists are never written**,
  by anything, in any configuration. They are her accumulated work and cannot be rebuilt.
- **Large print still deletes the 18 compact fractions.** A large print book keeps `1/2` as it
  was typed, and that is not a preference.
- **The store stays a file**, not module variables. VBA's `End` runs on ordinary paths here — 34
  times in `LPandBrlMacros` alone — and wipes module state.
- **`UpdateStylesOnOpen = False`** after an attach. `Sh_Close_And_Reopen` depends on it.
- **Nothing in an application event may raise.** An error inside the `VtEvents` sink can stop
  Word calling back for the rest of the session, which kills document-type detection silently.

## What to test, in order

Each of these has an expected answer that differs from what the add-in does today.

1. Open a letter on a machine that has never run VistaType. Nothing about Word changes except
   that the 18 fractions appear.
2. A letter opens with no pilcrows, no Styles pane, no navigation pane, both rulers, print view.
3. In a letter, change the Styles pane sort and "Select styles to show". Close it, open it again.
   Both are as she left them.
4. In an LP book, turn grammar checking on. Close the book, open it again — grammar is off, LP
   won. Open a letter — grammar is **whatever the user last set it to in a letter**, and that change
   inside the book did NOT follow the user out. **Rewritten 9/5/2026**; this test used to end "grammar
   is on, the change followed the user out", which was Decision 3 before Jerry reversed it. Watch a
   setting where the user's stored preference and the book's value actually differ — the five
   spelling/grammar settings are a poor choice on a machine where the user already has them off,
   because then the two values coincide and the test passes either way.
5. Book and letter open together. Click between them. **The configuration follows the document on
   screen and changes each time** — Jerry's rule, 9/4/2026. This test used to read "nothing about
   Word's configuration changes either way", which was Piece 2 and is rejected. Capitalization and
   the fractions are the two easiest things to watch.
6. Open a braille file, then open a letter from disk. Grammar, tab-indent and the fractions are
   all back.
7. Open a large print book made on an obsolete template. She still gets the warning and the
   attach dialog.
8. Book and letter open, Styles pane up, close the book. The pane comes down — correct now,
   because a letter has no Styles pane.
9. Close a book with a picture selected. No crash, and document-type detection still works
   afterwards.

## Traps

- **`Sh_ConfiguredAs` is the hinge of the whole ledger.** It must record what is *actually* in
  force, never what was intended — the 8/18 note on ordering inside the configuration subs
  applies unchanged.
- **Guarded writes hide from the ledger.** See Piece 4.
- **`Lp_Is_The_Attached_Template_LP` has eleven callers asking two different questions.** Renaming
  it is what makes the triage in Piece 1 possible; leaving the old name is how one of them gets
  missed.
- **`Sh_Config_Skip_Display` is NOT to be retired**, and neither is `Sh_Apply_Word_Config`'s
  `DisplayToo`. This trap used to read as instructions for removing them along with Piece 2. Both
  have live callers that have nothing to do with switching: `MS_Reset_Word_Configuration` passes
  `DisplayToo:=False`, and `Dx_Attach_BANA_Template_Run` raises the flag across the attach so the
  braille display changes at its stage four and not before (Jerry, 8/29/2026). The guard appears
  six times and one of them wraps `Application.ScreenRefresh` alone.
- ~~**The leak in Piece 2 will be reported as a bug.**~~ Moot: Piece 2 was rejected, so there is
  no leak. It was written into `docs/How-Word-Settings-Work.md` while the change existed and taken
  back out with it.
