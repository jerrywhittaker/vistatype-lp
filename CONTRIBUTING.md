# Contributing to VistaType LP

Thanks for your interest in improving VistaType LP — Word tools for large-print and
braille document production. Contributions of all kinds are welcome: bug reports, fixes,
new formatting tools, documentation, and installer improvements.

Please read this first — the project has an **unusual build model** (the code is VBA that
only Microsoft Word can compile), so the workflow isn't the typical "clone and run."

- **License:** by contributing you agree your contribution is licensed under the project's
  [GNU GPL v3.0](LICENSE) (inbound = outbound). Don't submit code you can't license this way.
- **Architecture & workflow:** [`DEVELOPMENT.md`](DEVELOPMENT.md) is the detailed guide.
  [`CLAUDE.md`](CLAUDE.md) is the architecture overview. This file is the short version.

---

## The one thing to understand first

**The text under `src/` is the source of truth; the `.dotm` is a build output.** VBA can
only be compiled by Word itself, so building runs on a **Windows machine with Word**,
driven from Linux/macOS over SSH. You edit text, run `make`, and a remote (or local)
Word instance imports and compiles it.

That means there are **two ways to contribute**:

1. **You have Windows + Word.** Point `build.config` at it and you can build and test
   end-to-end yourself. Best for anything touching runtime behavior.
2. **You don't.** You can still edit and submit **source** changes (`src/`, `installer/`,
   docs) — they're plain, reviewable text. A maintainer will build and validate in Word
   before merging. You can *read* all the VBA without Word via `make read` (see below).

---

## Setup

1. **Fork** the repo and clone your fork.
2. **Linux/macOS tooling** (for reading and non-Word tasks):
   - `git`, `python3` with the `olefile` package (`pip install olefile`) — used by the
     read/ribbon/QAT helper scripts.
3. **To build** (needed only for the compile/test step):
   - A **Windows box with Microsoft Word**, reachable over SSH, with *“Trust access to
     the VBA project object model”* enabled in Word. Copy `build.config.example` to
     `build.config` and set `WIN_HOST` / `WIN_DIR`.
   - **Inno Setup 6** on that box if you're building the installer.
4. Run `make help` to see all targets.

Full setup detail is in [`DEVELOPMENT.md`](DEVELOPMENT.md).

---

## Making changes

| You want to change… | Edit… | Then… |
|---|---|---|
| VBA logic | `src/vba/*.bas`, `src/vba/*.cls` | `make build`, smoke-test in Word |
| A dialog / UserForm | in the Word **VBE**, then `make pull` | forms' binary layout can't be authored from text |
| The ribbon | `src/ribbon/customUI14.xml` | rebuilt into the `.dotm` by `make build` |
| The installer | `installer/` (`vistatype.iss`, `scripts/`) | `make installer` |
| Docs | `*.md`, `docs/` | — |

**Never hand-edit `Normal.dotm`** — it's the build shell, not source.

Handy targets:

```
make read     # dump readable VBA to reference/ (no Windows needed) — great for browsing
make build    # import src/ -> dist/ .dotm via Word, embed the ribbon
make pull     # export canonical VBA/forms from the .dotm back into src/ (after VBE edits)
make installer# build the Setup.exe
```

---

## Coding conventions

- **Naming / namespacing.** Entry points are prefixed by area — `Lp_` (large print),
  `Dx_` (braille/Duxbury), `Sh_` (shared), `DN_` (DAISY/NIMAS), `MS_` (Word config).
  Match the surrounding style; grep the prefix to find related code.
- **Ribbon buttons** call the single `RibbonAction` dispatcher and name their macro in the
  control's `tag` — keep new buttons parameterless and follow that pattern.
- **Versioning.** Each Sub carries an inline `' Version / Date` comment; the
  `LPandBrlMacros` header keeps a dated changelog, and the About dialogs show the version
  string. Update these when you change behavior.
- Keep `Option Explicit`; write code that reads like its neighbors.

---

## Testing your change

Because there are no unit tests, **exercise the change in real Word** before submitting.
After `make build`, open `dist/Normal.dotm` and confirm:

- the VBA project **compiles** (VBE → *Debug → Compile*);
- the **ribbon loads** — the *VistaType LP* / *Braille Macros* tabs appear and buttons run;
- the specific behavior you changed works on a real document.

See the checklist in [`DEVELOPMENT.md`](DEVELOPMENT.md#always-smoke-test-a-build). If you
can't build (no Word), say so in your PR so a maintainer runs this pass.

---

## Submitting

1. Branch from `master` (`git switch -c fix-something`).
2. Make focused commits with clear messages explaining **why**, not just what.
3. Push to your fork and open a **pull request** describing the change, how you tested it
   (or that it needs a maintainer's Word test), and any docs updated.
4. Be ready for review feedback — for anything touching document formatting, expect
   questions about edge cases (tables, images, page numbering, braille vs. large print).

## Reporting bugs & requesting features

Open a **GitHub issue** with: what you did, what you expected, what happened, and your
Word version. For large-print/braille output problems, a small sample document helps a lot.

Questions or sensitive reports can also go to **jerry@thewhittakers.org**.
