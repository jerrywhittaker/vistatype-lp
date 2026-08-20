# Automatic Word configuration: the plan

Agreed by Jerry and the beta tester, 8/20/2026, from *Automatic Configuration Table.docx*.
Written up here before any code was touched.

**Status: nothing below is built.** This supersedes the unbuilt half of
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
3. **Reopening a book re-applies its configuration**, and a setting she changes *inside* a book
   still counts as hers for her ordinary documents. She does far more in Word than large print
   and braille, and she does not stop being an ordinary user when she opens a book.
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

### Piece 2 — switching stops reconfiguring

`Sh_HandleDocumentActivated` stops calling `Sh_Apply_Word_Config`. The `VtEvents` hooks stay —
the handler still has the Styles-pane recovery added on 8/20 — but it no longer configures
anything.

That retires `Sh_Apply_Word_Config`'s `DisplayToo` argument and `Sh_Config_Skip_Display` with
it. Their only caller passing `False` was the switch, which was the entire reason they existed
(8/9/2026: "switching sets the typing, not the screen"). Every configuration is now a full one,
which removes six `If Not Sh_Config_Skip_Display Then` guards from the three MS_ subs.

`Sh_ConfiguredAs` stays. It is how the ledger in Piece 4 knows which configuration is in force,
and that job survives.

**The cost, stated plainly.** With a book and a letter open together, typing in the letter
happens under the book's configuration — grammar checking off, tab-indent off, the fifteen
non-Word fractions gone. It corrects itself the moment she opens or creates anything, so this
is nothing like the 8/18 fault where one braille file turned grammar off for every session
after. But it is real and it is the price of the rule.

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

**How her changes are told apart from ours**, with no heuristics. When a book configuration is
applied, record into `[BookApplied]` the value it set for every tracked setting. When an
ordinary document takes over, compare live against `[BookApplied]`, setting by setting:

- **same** — the book's value still stands, she did not touch it, and her stored preference is
  left alone,
- **different** — she changed it while working, and that becomes her new preference in
  `[TranscriberSettings]`.

Then restore from `[TranscriberSettings]`. A setting no book ever writes is not in the ledger
at all and is trivially hers already.

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
`CommandBars("Navigation")` is the older Document Map task pane. Use `DocumentMap` and stop
writing the command bar.

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

### Piece 6 — the gap Piece 2 opens, which needs an answer first

On 8/18 `Sh_HandleDocumentNew` was made to skip itself while a macro is running, because the
scratch document `Lp_Copy_To_Temp_Doc` makes — from about 35 places — was reconfiguring Word in
the middle of a job. The note records the handoff: a document a macro makes *for* the
transcriber "is configured when her cursor reaches it instead." That handoff was to the activate
handler, and Piece 2 removes it. Those documents would never be configured at all.

Three ways out, and one has to be chosen before Piece 2 ships:

1. configure such a document explicitly when the macro that made it finishes,
2. have the "for the transcriber" route in `Lp_Copy_To_Temp_Doc` configure it itself,
3. keep a narrow activate-time configuration for a document that has *never* been configured,
   and only for that case.

Option 3 is the smallest change and the easiest to reason about, but it puts back a piece of
exactly what Piece 2 removes, so it needs Jerry's word rather than mine.

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
5. **Piece 6** — the decision, then
6. **Piece 2** — last, because it is the one that cannot be half-done.

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
   won. Open a letter — grammar is **on**, her change followed her out.
5. Book and letter open together. Click between them. Nothing about Word's configuration changes
   either way.
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
- **Retiring `Sh_Config_Skip_Display` touches all three MS_ subs.** The guard appears six times
  and one of them wraps `Application.ScreenRefresh` alone.
- **The leak in Piece 2 will be reported as a bug.** Write it into the guide before it ships.
