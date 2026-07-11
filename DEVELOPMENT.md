# Developing VistaType LP

The whole workflow runs from this Linux terminal. Word can't compile VBA on Linux,
so a remote Windows+Word box acts as an invisible build server driven over SSH. You
edit text, run `make`, and get a finished `.dotm` back — you never open Word by hand.

## Source of truth

`src/` is authoritative. **Never hand-edit `Normal.dotm`** — it is a build output.

```
src/vba/        *.bas standard modules, *.cls class/document modules   (canonical text)
src/forms/      *.frm UserForms + *.frx binary layout                  (canonical)
src/ribbon/     Word.officeUI ribbon XML                               (plain file; no Word needed)
Normal.dotm     the shell/base .dotm: project references + non-VBA parts (tracked)
LargePrintTemplate.dotx   the attached large-print template (styles/page setup; tracked)
dist/           build outputs (gitignored)
reference/      read-only aids (gitignored mirror + interim form-code dump)
```

## One-time setup

1. On the **Windows box**: enable OpenSSH Server, set up key login so `ssh WIN_HOST`
   works with no password. In Word: *Options → Trust Center → Macro Settings →
   ☑ Trust access to the VBA project object model*.
2. On **Linux**: `cp build.config.example build.config` and set `WIN_HOST` / `WIN_DIR`.
3. **Seed canonical source** (the current `src/` was bootstrapped by the Linux reader,
   which cannot produce valid `.frx`): run `make pull` once. This exports IDE-native
   `.bas/.cls/.frm/.frx` from `Normal.dotm` into `src/`. Review with `git diff`, commit.

## Everyday loop

```
edit src/vba/*.bas        # or src/forms, src/ribbon — on Linux, in git
make build                # Word imports src/ -> dist/Normal.dotm, copied back here
# smoke-test dist/Normal.dotm in Word (see below), then:
make deploy               # promote dist/ artifacts to repo root
git add -A && git commit
```

- `make read` — regenerate `reference/vba-src/` from `Normal.dotm` with the Linux-only
  decompressor. Handy for diffing what's actually compiled into the binary; needs no Windows.
- `make pull` — pull canonical VBA source back into `src/` after anyone edits in the VBE.

## Always smoke-test a build

`make build` produces a `.dotm` but does not prove it runs. Before `make deploy`:
open `dist/Normal.dotm` in Word once, confirm the VBA project compiles (VBE →
*Debug → Compile*) and the ribbon loads. This is the one manual Word step; everything
else is automated.

## Gotchas baked into the tooling

- **`ThisDocument`** is a Document module — it can't be Import-ed. `Import-Vba.ps1`
  clears its code module and refills it from `src/vba/ThisDocument.cls`.
- **Forms** must round-trip as `.frm` + `.frx` (the `.frx` holds images/binary layout).
  The Linux reader can't rebuild `.frx`; that's why `make pull` (Word export) is the
  canonical seeder for `src/forms/`.
- **Project references** (MSWORD.OLB, FM20.DLL/MSForms, scrrun.dll, stdole) and the
  attached-toolbar ribbon live in the shell `Normal.dotm`, not in `src/`. The build
  imports code *into a copy of that shell*, so those are preserved automatically.
- **Version bumps**: follow the in-file convention — update the per-sub `' Version`
  comment and add a dated line to the `LPandBrlMacros` header changelog and the About
  forms' version string.
