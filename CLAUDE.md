# VistaType LP

A Microsoft Word add-in that helps transcribers produce **large-print** documents for
visually impaired readers, and **braille** source files for the Duxbury Braille Translator
(DBT). Jerry Whittaker (jerry@vistatypelp.org), 2015–2026.

`src/` is the source of truth. `src/vba/*.bas` and `src/forms/*.frm` are the real code —
~16,500 lines, mostly in `LPandBrlMacros`. **Never hand-edit `LPandBRL.dotm`**; it is a build
output. Only Word can compile VBA, so `make build` ships the source to a Windows box over SSH
and brings the built add-in back.

Everything is namespaced: `Lp_` large print, `Dx_` braille/DBT, `Sh_` shared, `DN_`
DAISY/NIMAS, `MS_` Word configuration. Grep by prefix to find a feature area.

**This file is process only.** Everything else is in `docs/` — see the table below.

## Talking to Jerry

The full rules are in `~/.claude/CLAUDE.md` and apply everywhere. In short:
**Speak Word and VBA to him freely** — subs, UserForms, `.frx`, ranges, styles, section breaks,
DBT tables. He wrote all ~16,500 lines himself over twenty years.

**Strip the process jargon**, which he never needed: branch, merge, rebase, HEAD, upstream,
checkout, diff, CI, artifact, refactor, regression. Say the effect, not the mechanism. Never
say "just". Lead with what happened or what to do, then why. Short paragraphs — he is reading
in a terminal between Word sessions.

**American spellings everywhere**, in replies and in the product alike. The one exempt file is
`LICENSE`, which quotes the GPL verbatim.

## How to work here

**Check the open issues at the start of a session** — `gh issue list`. Defects are GitHub
issues from 9/20/2026. Tell Jerry briefly what is open as things he could pick up, but do not
start on one unless he asks. **Record any new defect the same way**, not as a note buried in a
document, saying what was verified and what was not.

**Delegate the work, not just the looking.** Getting back to Jerry quickly matters more than
doing it yourself, and a transcript in the main session slows every later turn.
**The trigger, so it cannot be argued away: three or more tool calls, or output you will not
read in full. Either one means an agent.** Below that, do it here — a single `grep` is faster
than a round trip. On 20/9/2026 the one job that was properly delegated, the pre-public sweep,
caught a leak the main session had missed.

- **`scout`** (Sonnet/low) — where is X, what calls Y, how many Z, is this claim still true.
- **`build-runner`** (Sonnet/low) — runs a make target; knows the five silent build failures.
- **`vba-test`** (Sonnet/high) — runs a macro headlessly; reports the text, not the transcript.
- **`diagnose`** (Opus/high) — a fault Jerry reported. Reads the error register first,
  reproduces before theorising.
- **`secrets-check`** (Opus/high) — **before any push touching signing, and before the repo is
  ever made public.**
- **`issue-triage`** (Sonnet/medium) — one issue: real? reachable? what would a fix disturb?
- **`vba-review`, `braille-lp-review`, `ship-safety`, `release-check`** — for a change.

Send independent jobs in **one message** so they run at once, and never do a search yourself
*and* delegate it.

## Where things are written down

| Question | File |
|---|---|
| Has this fault been reported and fixed before? | `docs/Reported-Errors.md` — **read first, always** |
| What is this file / folder / tool for? | `docs/Repo-Layout.md` |
| How should a dialog, message or progress bar look? | `docs/UI-Conventions.md` |
| What do the domain words mean? | `docs/Domain-Concepts.md` |
| How do I build, and what are the tooling traps? | `DEVELOPMENT.md` |
| How does a release work? (Jerry's own guide) | `docs/Daily-Workflow-and-Releases.md` |
| Word configuration and the settings ledger | `docs/Word-Configuration-Rules.md`, then `docs/User-Settings-And-Word-Configuration.md`; the transcriber-facing version is `docs/How-Word-Settings-Work.md` |
| Installing, updating, troubleshooting (for transcribers) | `docs/Installation-Guide.md` |
| Code signing, and why antivirus eats the installer | `docs/Code-Signing.md` |
| The bundled typeface: coverage, and what it will not set | `docs/VistaTypeLP-Sans.md` |
| Moving a macro onto a range, and the traps | `docs/Temp-Doc-Conversion-Checklist.md` |
| What to press when testing the ribbon | `docs/Ribbon-Test-Checklist.md` |
| The undo cost of every macro (closed survey) | `docs/Undo-Audit.md` |
| Building the Windows box | `docs/Build-VM-Setup.md` |
| The About dialog's license wording | `docs/Software-Agreement.md` |

## Before you change anything

**Read `docs/Reported-Errors.md` first** — Jerry's rule, 8/26/2026: *"I don't want to spend
time and tokens on re-investigating errors that have already been fixed."* Keyed on version,
error number, macro and step. Add the row **and a test, wherever the fault is testable**, in
the same change as the fix — Jerry, 9/20/2026. A fix with no test is how a fault comes back.

**Nothing here compiles VBA, and a compile gate is impossible** — measured 8/26/2026: VBA
compiles lazily per procedure and `Debug > Compile` is unreachable by automation. The compile
is a **human step** before any release; say so whenever a change touches VBA. Give every helper
`ByVal` parameters — the fault that proved this was a Variant passed to a ByRef `String`.

## Building and testing

**`make build` runs `make test`** (the Python suite, half a second), so every `try` and every
`installer` does too; **`make installer` also runs `make vba-test`** (the VBA suite, on the
box). A failing test stops the build. What each covers: `tests/README.md`, `tests/vba/README.md`.
`make try` bumps, builds and drops the add-in into Word's STARTUP folder on the build box;
`make build` builds without deploying; `make installer` bumps and produces the `Setup.exe`.
Word must be closed on the box. Hand these to `build-runner`.
**`make try` cannot test** `installer/**`, `src/ribbon/**`, `src/keymap/**`, or `*.dotx`.
`tools/lib/check_try_scope.py` prints this after every try — **say it in the reply too**
(Jerry, 8/23/2026). He should never have to work out for himself that what he is about to test
cannot be tested that way.

Every sub carries an inline `Version:`/`Date:` comment block, and `LPandBrlMacros`'s header
carries a dated changelog. Bump both when changing behavior.

## Git and releases

**`dev` is all day-to-day work. `main` is the last released version.** Work only ever moves
`dev` → `main`, fast-forward only, one tag per release. (`main` was `master` until 9/20/2026.)

**Version numbers:** the third digit — `3.0.460` — is a private build counter, bumped freely,
never published. `X.Y` — `3.1` — is a real release. A fourth digit — `3.1.0.1` — is a hotfix in
a separate lane so it can never collide with `dev`'s counter. **Only Jerry starts a release, and
by name** — "let's release 3.1"; a clean build is the normal end of a day, not a cue. Checklist,
hotfix and documentation-only routes are in `docs/Daily-Workflow-and-Releases.md`.

**Documentation-only changes go straight to `main`**, so Jerry's bookmark is never stale.
Only if the change touches nothing outside `docs/**`, `CLAUDE.md`, `DEVELOPMENT.md` and
`README*` — verify with `git diff --name-only`, never by eye. Write it on `main`, then merge.

## Hard rules

- **Never push code until Jerry says the word "push."** Not `dev`, not `main`, not tags.
  Documentation-only changes are exempt — push those freely.
- **Never commit code on `main`** (check `git branch --show-current` first), never force-push,
  and never rewrite or delete a `v*` tag — the tags are the recovery points.
- **Never publish a release without its `.exe`.** GitHub's source zip cannot be installed by a
  transcriber. Verify with `gh release view <tag> --json assets`; it must not be `[]`.
- **Never hand-edit `LPandBRL.dotm` or `RibbonDispatch.bas`** — both are generated.
- **Never rename or remove a `btn_*` id** in `src/ribbon/customUI14.xml` — toolbars in the
  field reference them by id and a renamed one draws blank, silently. Retire a button by moving
  it to the hidden `tab_VT_Retired_Buttons`.
- **A `.frm` must stay CRLF**, or Word dumps the designer header into the form's code module —
  and it fails only on the transcriber's machine.
- **A signed `.dotm` must never reach the repo root** — it is the base every build starts
  from. `make deploy` and `push-src` refuse one.
- **Never merge this project with `~/projects/vistatypelp-org`, or suggest it, or edit the
  site from here.** On a release, hand Jerry the prompt from the daily-workflow guide.
- **Messages go through `Sh_Say` / `Sh_Ask`**, never a bare `MsgBox`: 10 point Tahoma or more,
  the button says `Okay`, Alt+O and Alt+C. One progress indicator, the bar.
- **Say what actually happened.** Show failing output, name skipped steps, and say whether a
  number was measured or guessed.

## Keeping this file honest

It is loaded in full at the start of every session, so it is capped at **150 lines**. Anything
longer belongs in `docs/` with a pointer added to the table above.
When you change the build — `Makefile`, `DEVELOPMENT.md`, `tools/`, `src/ribbon/`,
`installer/` — update `docs/Repo-Layout.md` in the **same** change. A `PostToolUse` hook
(`.claude/hooks/claude-md-sync.py`) reminds you.
