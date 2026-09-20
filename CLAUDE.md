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

**Speak Word and VBA to him freely** — subs, modules, UserForms, `.frx`, ranges, styles,
section breaks, DBT tables. He wrote all of this himself over twenty years.

**Strip the process jargon**, because he never needed it: branch, merge, rebase, HEAD,
upstream, checkout, staging, diff, CI, artifact, refactor, regression. Say the effect, not the
mechanism. Never say "just". Lead with what happened or what he should do, then why. Short
paragraphs. He is reading in a terminal between Word sessions.

**American spellings everywhere** — in replies, dialogs, installer text, code comments alike.
license, color, behavior, recognize, gray, dialog. Periods not full stops, parentheses not
brackets. The one exempt file is `LICENSE`, which quotes the GPL verbatim.

## Start of a session: check the open issues

**Defects are recorded as GitHub issues** (from 9/20/2026). At the start of a session run
`gh issue list` and tell Jerry briefly what is open, as things he could pick up. Do not start
on one unless he asks.

**Record new defects the same way.** Anything found while working on something else — a bug, a
dead macro, an untested path — becomes an issue rather than a note buried in a document. Say
what was verified and what was not.

## Where things are written down

| Question | File |
|---|---|
| Has this fault been reported and fixed before? | `docs/Reported-Errors.md` — **read first, always** |
| What is this file / folder / tool for? | `docs/Repo-Layout.md` |
| How should a dialog, message or progress bar look? | `docs/UI-Conventions.md` |
| What do the domain words mean? | `docs/Domain-Concepts.md` |
| How do I build, and what are the tooling traps? | `DEVELOPMENT.md` |
| How does a release work? (Jerry's own guide) | `docs/Daily-Workflow-and-Releases.md` |
| How does Word configuration and the settings ledger work? | `docs/Word-Configuration-Rules.md` |
| What does VistaType do to the transcriber's own settings? | `docs/User-Settings-And-Word-Configuration.md` |
| The transcriber-facing explanation of the same | `docs/How-Word-Settings-Work.md` |
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
time and tokens on re-investigating errors that have already been fixed."* It is keyed on
version, error number, macro and step. Add the row in the **same change** as the fix.

**Nothing here compiles VBA, and a compile gate is impossible.** Measured 8/26/2026: VBA
compiles lazily per procedure, a probe macro proves nothing, and `Debug > Compile` cannot be
reached by automation. So the compile is a **human step** before any release. Say so whenever
a change touches VBA. Give every helper `ByVal` parameters — the fault that proved this was a
Variant passed to a ByRef `String`.

**Read the mechanism you are about to introduce, not only the feature you were asked about.**
More than one fault here was walked past because it was already written down elsewhere.

## Building and testing

`make try` is the short loop: it bumps the build number, builds, and drops the add-in straight
into Word's STARTUP folder on the build box. `make build` builds without deploying.
`make installer` bumps and produces the `Setup.exe`. Word must be closed on the box.

**`make try` cannot test** `installer/**`, `src/ribbon/**`, `src/keymap/**`, or `*.dotx`.
`tools/lib/check_try_scope.py` prints this after every try — **say it in the reply too**
(Jerry, 8/23/2026). He should never have to work out for himself that what he is about to test
cannot be tested that way.

Every sub carries an inline `Version:`/`Date:` comment block, and `LPandBrlMacros`'s header
carries a dated changelog. Bump both when changing behavior.

## Git and releases

**`dev` is all day-to-day work. `main` is the last released version.** Work only ever moves
`dev` → `main`, fast-forward only, one tag per release. (`main` was called `master` until
9/20/2026.)

**Version numbers:** the third digit — `3.0.460` — is a private build counter, bumped freely,
never published. `X.Y` — `3.1` — is a real release. A fourth digit — `3.1.0.1` — is a hotfix
to a released version, kept in a separate lane so it can never collide with `dev`'s counter.

**Only Jerry starts a release, and he starts it by name** — "let's release 3.1". A clean build
is the normal end of a day's work, not a cue to release. The full checklist, the hotfix route
and the documentation-only route are in `docs/Daily-Workflow-and-Releases.md`.

**Documentation-only changes go straight to `main`**, so Jerry's bookmark is never stale.
Qualifies only if the change touches nothing outside `docs/**`, `CLAUDE.md`, `DEVELOPMENT.md`
and `README*` — verify with `git diff --name-only`, never by eye. Write it on `main`, then
merge `main` into `dev`.

## Hard rules

- **Never push code until Jerry says the word "push."** Not `dev`, not `main`, not tags.
  Documentation-only changes are exempt — push those freely.
- **Never commit code on `main`.** Check `git branch --show-current` first.
- **Never force-push, never rewrite or delete a `v*` tag.** The tags are the recovery points.
- **Never publish a release without its `.exe`.** GitHub's automatic source zip cannot be
  installed by a transcriber, so a release without the installer delivers nothing. Verify with
  `gh release view <tag> --json assets` and confirm it is not `[]`.
- **Never edit `LPandBRL.dotm` by hand**, and never hand-edit `src/vba/RibbonDispatch.bas` —
  it is generated.
- **Never rename or remove a `btn_*` id** in `src/ribbon/customUI14.xml`. Every toolbar already
  installed in the field references them by id, and a renamed one draws blank with no error.
  To retire a button, move it to the hidden `tab_VT_Retired_Buttons`.
- **A `.frm` must stay CRLF.** LF endings make Word dump the designer header into the form's
  code module, and it fails only on the transcriber's machine.
- **A signed `.dotm` must never reach the repo root** — it is the base every build starts from.
  `make deploy` and `push-src` refuse one.
- **Never merge this project with `~/projects/vistatypelp-org`, and never suggest it.** The
  website is separate, permanently. On a real release, remind Jerry and hand him the prompt
  from `docs/Daily-Workflow-and-Releases.md` to paste into a session there. Do not edit it
  from here.
- **Messages go through `Sh_Say` / `Sh_Ask`**, never a bare `MsgBox`: at least 10 point Tahoma,
  the button says `Okay`, Alt+O and Alt+C. One progress indicator, the bar. See
  `docs/UI-Conventions.md`.
- **Say what actually happened.** If a test failed, show the output. If a step was skipped, say
  so. If a number was guessed rather than measured, say which.

## Keeping this file honest

It is loaded in full at the start of every session, so it is capped at **150 lines**. Anything
longer belongs in `docs/` with a pointer added to the table above.

When you change the build — `Makefile`, `DEVELOPMENT.md`, `tools/`, `src/ribbon/`,
`installer/` — update `docs/Repo-Layout.md` in the **same** change. A `PostToolUse` hook
(`.claude/hooks/claude-md-sync.py`) reminds you.
