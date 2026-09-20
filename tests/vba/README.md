# The VBA tests

Run them with `make vba-test`. Word must be **closed** on the build box.

They take about a minute, most of which is copying the source up. Nothing is installed, Word
is never visible, and the throwaway document is closed without being saved.

## What this actually runs

Not the built add-in. `tools/lib/build_vba_test_bundle.py` reads the `'@uses` lines at the top
of each test module, lifts those procedures out of `src/vba` **as they stand now**, follows
what they call, brings the module-level declarations they need, and writes one module. The
runner imports that plus these test modules into a new invisible document and calls
`VtRunAllTests`.

Two reasons it works that way:

- It tests **the source in front of you**, not the last build.
- Most of the helpers worth testing are `Private`. A test document that merely referenced the
  add-in could not call them at all.

Adding a test means writing the assertions and adding a `'@uses Name` line. If that name is
not in `src/vba` the bundle is refused, by name — because a missing procedure is not an error
here, it is a hang.

## What CANNOT be tested this way

Every one of these was measured, and each cost at least one run that ended in a Word to kill
and stale lock files to sweep.

| Not testable | What you see |
|---|---|
| Anything that reaches a `Document`, `Selection` or `Range` | Usually a hang, sometimes a wrong answer from an empty document |
| A macro that ends in `MsgBox` | The modal dialog is behind an invisible Word; the run hangs until it times out |
| A UserForm | `.Show` blocks even modeless, and so does `Load` — the centering code reads `Application.Left` |
| `Tables.Add`, `ContentControls.Add` | Hang outright |
| The event path — `AutoExec`, `App_DocumentOpen`, the configuration subs | A Word started over SSH does not load the STARTUP add-in at all, so nothing fires and it looks exactly like a real fault |
| A discontiguous (Find All) selection | A UI-only state; it cannot be built from VBA |

So this suite is for **pure helpers**: a string or a number in, a string or a number out. There
are roughly 36 of those in `LPandBrlMacros`. Everything else is still a job for real Word, and
`docs/Ribbon-Test-Checklist.md` and `DEVELOPMENT.md` cover that.

## When a run hangs

`vbatests/results.txt` on the build box is written a line at a time and closed after each one,
on purpose, so a run that is killed still says how far it got. Read it first.

The usual cause is that the bundled module did not compile. In an invisible Word a compile
error is a modal dialog nobody can answer, which is why it hangs rather than failing. Look at
whatever was about to be called when the trace stops.

Afterwards there will be a `WINWORD.EXE` left on the box. **Read the process list before
ending anything**: `SessionId` 0 with `/Automation -Embedding` is one of these runs and is
safe to end; any other session is somebody's own Word with unsaved work in it. Then clear the
hidden `~$*` files, or the next run meets Word's "File In Use" dialog and hangs the same way.
`Run-VbaTests.ps1` refuses to start while any Word is running and prints the list.

## The library

`lib/TestSuite.cls` and `lib/TestCase.cls` are [vba-test](https://github.com/VBA-tools/vba-test)
by Tim Hall, MIT — see `lib/LICENSE-vba-test.txt`. It is plain VBA with no Excel dependency, it
runs from an ordinary Sub, and it is the only maintained option: Rubberduck, the name everyone
reaches for, was archived in March 2026 and could only ever run tests from its own window
inside the VBA editor, which is unreachable from here.

**`TestCase.cls` is modified in one place**, marked in the file. It held `Context As Dictionary`,
early bound against Microsoft Scripting Runtime. Adding that reference to the test document
succeeds — it is listed on the project afterwards — and the early-bound type still does not
resolve. That is a compile error, so it hung, at `S.Test(...)`, with no message. Late binding
works and needs no reference at all.

`lib/VtFileReporter.cls` is ours. vba-test's own reporter writes to the Immediate Window, which
nothing can read back out of a Word started over SSH.
