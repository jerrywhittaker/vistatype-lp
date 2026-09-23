# How VistaType LP talks to the transcriber

How a dialog looks and reads, how a box that stays open behaves, how progress is shown, and how
a failed macro is reported.
Moved out of `CLAUDE.md` on 20 September 2026. Same text, no changes.

These are requirements about how a message looks and reads. **They are not a ban on anything** —
Jerry's correction, 8/26/2026, after the earlier wording was quoted back at him as a reason his
own `MsgBox` suggestion could not be done.

---

### How VistaType LP says things — dialogs and messages

Jerry's rule, 8/23/2026. It governs every new dialog and every message:

- **The message is at least 10 point Tahoma.** That is the floor, not the target — a bigger
  message is fine, a smaller one is not. The people using this software work on large print
  and braille, and several of them have low vision themselves.
- **The OK button is captioned `Okay`, never `OK`.**
- **`Okay` has an accelerator of `O`; `Cancel` has `C`** — Alt+O and Alt+C press them.
- Every CommandButton still gets a `ControlTipText` of its caption plus " Button", and a form
  is still Tahoma 10 throughout (the 8/7/2026 rules; both stand).

**These are requirements about how a message looks and reads. They are not a ban on anything** —
Jerry's correction, 8/26/2026, after this file's earlier wording was quoted back at him as a
reason his own `MsgBox` suggestion could not be done. He set how a message should appear; he did
not forbid a control or a mechanism.

What follows from that is a fact about `MsgBox`, not a rule of his: VBA's `MsgBox` draws in
whichever font Windows uses for message boxes, its button says `OK` and cannot be changed, and it
has no accelerators — so a plain `MsgBox` cannot meet the three. A UserForm can, which is why
messages here are UserForms. Where a `MsgBox` is the right answer anyway, say so and let Jerry
decide; do not cite the rule as settling it. `Sh_Message_Form` is the one dialog every message goes through, reached by two
subs in `LPandBrlMacros` and never touched directly:

```vba
Sh_Say "text", "VistaType LP (nnn)"             ' tell her. One Okay button.
If Sh_Ask("text", "VistaType LP (nnn)") Then    ' ask. Okay and Cancel; True means Okay.
```

The second argument is the title bar, and that is where the dialog's own number lives. Those
numbers identify one dialog out of all of them when a transcriber says what she saw, so a new
message takes the next unused number and an existing one keeps the number it has always had.
(Written `nnn` above on purpose — a real number in an example is one more hit to wade through
when hunting for the next unused one.)

Paragraph breaks are written `vbCr`, as everywhere else in this project. `Sh_Message_Form`
turns them into `vbCrLf` on the way into the text box, because an MSForms text box is not a
`MsgBox` and the one other form in this project that fills one uses `vbCrLf` throughout.

What is lost against `MsgBox`: its information / warning / question **icon**, and the sound
the warning one made. A UserForm has neither.

**If the form refuses to appear, `Sh_Say` and `Sh_Ask` fall back to a `MsgBox`** (8/26/2026).
Until then the trap around `.Show` swallowed the failure and the message was simply **lost** —
intolerable in an error handler, which is where it came up. Smaller and saying `OK` beats never
appearing. In `Sh_Ask` it changes an answer too: a question that could not be shown used to come
back `False`, i.e. Cancel, and silence that quietly means "no" cannot be told apart from the user
having chosen "no".

**Not retrofitted.** The two newest toolbar buttons — Reset Word Configuration and Styles Pane:
Recommended — were converted on 8/23/2026, five messages. **The attach-a-large-print-template path
followed on 9/6/2026, sixteen messages** — the obsolete-template warning (123), the missing-template
stop (141), the cancelled-save question (306), the could-not-reopen notice (233), the
stabilized-and-saved notice (201), and the eleven page-size and margin checks on
`LP_Attach_An_Lp_Template_Form` (189–199). Every one kept the number it already had.
They convert a feature at a time. Nothing new uses `MsgBox`.

**What is left, recounted 9/20/2026: 67 in `LPandBrlMacros`, 28 across the forms, 18 in the
three smaller modules — 113 in all.** The total has not moved since 9/6/2026, but the split has:
two fewer in the forms and two more in the module. Worth knowing that the total staying still
does not mean nothing changed. **Count with comment lines excluded**, or the number comes out
around 40 too high: the changelog and the version blocks say the word `MsgBox` constantly, and so
do commented-out calls. `grep -ah MsgBox <files> | sed 's/^[[:space:]]*//' | grep -v "^'"` is the
count that means something. The figures this paragraph carried until 9/6/2026 (85 / 48 / 20) were
made the unfiltered way on 8/23/2026 and were never right.

**Three things a conversion cannot carry, and they decide the ORDER of the remaining work:**
- **The icon and the sound.** Around 45 of the calls pass `vbExclamation`, `vbInformation`,
  `vbCritical` or `vbQuestion`, and the warning ones make a noise. A UserForm draws neither and
  plays neither. Weigh that per message rather than sweeping.
- **Yes/No has to be reworded.** `Sh_Ask` offers Okay and Cancel. Sixteen calls are `vbYesNo`;
  four are already `vbOKCancel` and are free. **None is three-button**, which is why this is
  possible at all.
- **`Sh_Message_Form` fixes which key presses which button, and it cannot be told otherwise.**
  `New-UserForm.ps1` built it with `Okay.Default = True` and `Cancel.Cancel = True`, so **Enter is
  always Okay and Esc is always Cancel**. Six `MsgBox` calls pass `vbDefaultButton2` — they are
  saying the SECOND button is the safe one, which this form cannot express. Word such a question
  so that Okay is the safe answer, or leave it a `MsgBox`.
- **A `vbYesNo` gains an escape hatch the moment it moves here.** A Yes/No `MsgBox` disables its own
  X and ignores Esc: the user has to answer it. `Sh_Message_Form` takes both as Cancel. On dialog
  306 that means Esc now carries on **without saving**, silently, where before it did nothing.
  Before converting any other Yes/No, ask what Cancel costs if it is pressed by reflex.

**Twenty of the calls read an answer; the rest only tell the user something.** The tell-only ones
convert almost mechanically. Do not attempt all of them in one sweep: nothing here compiles VBA,
there are no automated tests, and a scripted edit across this module has already deleted code out
of both book configurations with every guard passing and a green build.

**Hover text cannot be made 10 point Tahoma, on a ribbon button or on a form.** Word draws a
ribbon/QAT screentip and supertip itself in the Office UI font; customUI has no font attribute
of any kind, and a `getSupertip` callback supplies the words and nothing else. `ControlTipText`
on a UserForm control is a Windows tooltip and is the same. The only thing that changes either
is the reader's own Windows text size (Settings → Accessibility → Text size), which changes it
everywhere. Where hover text needs to be readable at 10 point or more, the answer is to put the
words in the dialog, which VistaType LP does control.


### Boxes that stay open — F6, and the out-of-range message

Jerry's rule, 9/23/2026. A box that stays open while the transcriber works in the book — a
**modeless** UserForm, shown with `.Show vbModeless` — follows these rules. That covers Table and
TOC Tools' TOC box (354), Resize Pictures in a Selected Range (375), Type Fill-In Lines (357),
the $pg validation menus, and **every box made modeless from now on**, new or converted. The progress bar and the "please
wait" notice are not boxes the transcriber works in, and are not covered.

- **F6 and Shift+F6 move the keyboard between the book and the box**, both ways. In the box,
  every control's `KeyDown` sends the keyboard back to the book on either key. In the book, the
  key is **shared**: Word keeps one command per key, so it is bound once, to
  `Sh_Box_ToggleFocus`, while any such box is up. A box joins the list with `Sh_Box_Opened` and
  leaves it with `Sh_Box_Closed` (both in `ShNonModalMessage`); **the newest box up has F6**, and
  closing it hands F6 to the box opened before. Never bind F6 for one box on its own — that is
  how a live box was left without it (review, 9/23/2026).
- **When the box first appears, the book has the keyboard**, not the box. So does every press of a button that
  types into the book (357), so the transcriber can move to the next place and press again.
- **Cancel becomes Done.** A box that stays open has nothing to cancel; Done closes it, and the
  title bar's X goes through the same Done.
- **Where the box works on a held range, moving the cursor out of it is said** — "Your cursor is
  out of the selected range." — at every new spot outside, including a picture selected outside
  it. Silent while any box's job is running, since the jobs move the selection themselves, and
  **only the newest box up watches** (`Sh_Box_On_Top`), or two boxes would each complain about
  every click in the other's range. Join the list before the box first moves the selection, and
  leave it after the box last does.

A gray selection does **not** mean the box has the keyboard. Word draws the selection the same
gray whichever of the two has it (measured 9/23/2026). Press Down arrow to tell: the cursor moves
in the book, or the choice moves in the box.


### How VistaType LP shows that it is working

Jerry's rule, 9/6/2026: **one progress indicator.** His words — *"the current state of progress
indicators makes VistaType LP look schizophrenic... I like to have just a progress bar but i do
like the text which indicates what is happening."*

There were four of VistaType LP's own: two near-identical spinner boxes
(`Sh_NonModalMessageForm`, `Sh_Please_Wait_Form`), the bar (`Sh_Convert_Progress_Form`), and a bar
drawn out of pipe characters in Word's status line by the two Export Selection macros. **Every
feature is on the bar now and both spinner boxes have no callers left** — they are still in the
project, unused, because deleting a UserForm is a deliberate job where the code naming it has to
go first.

**The shape.** A fill that only ever goes forwards, a percentage, a line of text naming the step,
and a spinner. Reached through four helpers in `ShNonModalMessage`, never by naming the form:

```vba
Sh_Progress_Open "Fixing common file errors"   ' modeless; starts the spinner
Sh_Progress_Say 40, "Removing square bullets"  ' moves it and says what is happening
Sh_Progress_Hide / Sh_Progress_Show            ' step aside for a modal dialog of Word's own
Sh_Progress_Close                              ' every path out, including the error handler
```

- **Never touch `Sh_Convert_Progress_Form` directly.** Every helper is gated on the module flag
  `Sh_Progress_Up`, never on the form's `.Visible`, because reading any property of an unloaded
  UserForm instantiates it and runs its `Initialize` — and this form's `Initialize` reads
  `Application.Left`, which hangs a Word driven over SSH with no desktop. That gating is also what
  makes a `Say` with no bar open cost nothing, so a macro can carry these calls and still be run
  from somewhere that shows no indicator.
- **The spinner is for the steps where nothing can move.** `Repaginate`, `InsertFile` and the Save
  As dialog are single calls into Word, and VBA is single-threaded. The fill says how far along;
  the spinner says still alive. Before it existed the only answer was a message apologizing for a
  frozen bar.
- **`Sh_Progress_Span lo, hi` lends a slice to a sequence that counts its own work 0–100**, so one
  counted job can run inside another. The attach hands 3–35 to File Cleanup and 74–90 to Normalize
  Styles. **Give the bar back** (`Sh_Progress_Span 0, 100`) the moment the inner sequence returns,
  or every later `Say` lands inside the slice and the bar stops part way along.
- **Count with a step helper, one per sequence** — `Dx_Ffc_Step`, `Lp_Ffc_Step`, `Lp_Ns_Step`. The
  total is **one more than the number of passes**, because each step announces itself *before* its
  pass runs and a bar must never read 100% while work is still going on. Where stages differ
  enormously in cost — the attach, the DAISY converter — hand-pick the percentages instead of
  counting equal steps. Jerry's choice, 9/6/2026: rough stages that always advance, never a bar
  that sits frozen looking like a hang.
- **Close it on every path**, including the error handler, and **take a copy of `Err.Number` and
  `Err.Description` first** — `Sh_Progress_Close` runs `Err.Clear` inside itself, so reporting
  after it reports 0.
- **Word's own indicator is not ours and cannot be removed.** Word draws a message and a bar in
  the status line while it saves, and on OneDrive while it uploads. VistaType LP writes nothing to
  `Application.StatusBar` any more. **Do not hide our bar to avoid the overlap** — that was tried
  at 3.0.384 and reversed at 3.0.385: an export to a LOCAL disk gets no indicator from Word at
  all, so hiding ours takes the indicator away from the only case that has nothing else, and on
  OneDrive it turned one steady box into three events. Jerry: *"we need to see the bar if the
  export is to a local disk."*


### How VistaType LP reports a macro that failed

Added 8/26/2026, at Jerry's request. Before it, of 175 `On Error` statements in the project,
**four** ever put an error in front of the user; the rest either swallowed it or let it reach
the user as **Word's own "Run-time error" dialog**, which offers **Debug** — and the VBA project
is not locked for viewing, so Debug opens the source on their machine.

- **`RibbonAction` is the single catch point** for all 47 ribbon and toolbar buttons. It calls
  each macro **directly**, through the generated `Sh_Dispatch` — *not* `Application.Run`, which
  never lets the error back out (see `build_ribbon_dispatch.py` above). `Application.Run` remains
  only as the fallback for a name the table lacks.
- **`Sh_Report_Error`** restores the screen first (`ScreenUpdating`, `DisplayAlerts`, the progress
  box and the please-wait box), then appends one line to the log, then shows dialog **240**
  through `Sh_Ask` — where Okay opens the folder holding the log. Nothing in it may raise: it
  runs on a path that has already gone wrong, and a second error there is Word's dialog again.
- **The log is `%AppData%\VistaType LP Settings\VistaType-Errors.log`**, beside `VistaType.ini`
  in the folder the uninstaller deliberately leaves alone — a log thrown away by the reinstall
  someone suggested to cure the fault is worth nothing. `Sh_Store_Folder` is shared by both.
- **`Sh_Last_Activity`** carries the "Where:" line. VBA gives a handler a number and a description
  and nothing else — no line number (`Erl` needs numbered lines, which this project does not use)
  and no call stack — so the last `SetActivityMessage` string is the only position marker there is.
  The form's `SetActivityMessage` records it as well as showing it.
- **Not total coverage.** The nine keyboard shortcuts in `src/keymap` and every UserForm button
  call their macros directly and are still unguarded.
- **`VT_VERSION` says which build produced a fault.** A placeholder in `src/vba` forever
  (`"unstamped"`); `Import-Vba.ps1` replaces it inside the **built** `.dotm`, the same pass that
  stamps the two About captions. Never put a real number in the source: the version bumps on
  every `make try` and it would be committed by accident. A log line reading `unstamped` means
  the build did not stamp it.
- **The log is newest-first and capped at 50 lines** (Jerry, 8/26/2026): the line anyone needs is
  the one they just caused. `Sh_Log_Newest_First` rewrites the whole file — Append can only add
  at the bottom. **It is not cleared when the user presses Cancel**, and that was considered and
  rejected: Cancel means "do not open the folder", not "throw this away", and the user who
  cancels is exactly the one whose repeats are the only record you would ever get.
