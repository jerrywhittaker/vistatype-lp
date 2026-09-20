# The test suite

Run it with `make test`, or `python3 -m pytest tests/ -q`. It needs pytest
(`sudo apt install python3-pytest`) and about half a second. It never opens Word and never
touches the build box.

## What it covers

The build's own checkers and file generators — the Python under `tools/lib/`. These are the
only thing standing between a VBA mistake and a transcriber's machine, and until 9/20/2026
nothing tested them.

| Test file | What it protects |
|---|---|
| `test_check_vba_structure.py` | A module-level `Const` beside its procedure, a duplicated procedure name, a `With` with no `End With`. All three ship as "Compile error in hidden module". |
| `test_check_style_guards.py` | An `ActiveDocument.Styles("X")` lookup that is not guarded — run-time error 5941 on a document that does not carry the style. |
| `test_check_form_calls.py` | A UserForm calling a macro that has been renamed or removed. |
| `test_check_frm_eol.py` | A `.frm` that has lost its CRLF endings, which makes Word dump the designer header into the form's code module. |
| `test_check_vba_line_length.py` | A VBA line over 1023 characters, which makes `make build` hang with no error. |
| `test_trim_frm_blanks.py` | The blank lines Word piles up on every form export — and that trimming them keeps CRLF. |
| `test_build_ribbon_tabs.py` | The generated ribbon tabs, and the rule that a `btn_*` id which has shipped can never disappear. |
| `test_build_qat.py` | The Quick Access Toolbar files, and that the generated one can never hide a button the transcriber put there. |
| `test_inject_customui.py` | Embedding the ribbon in the built `.dotm`, and that doing it twice replaces rather than duplicates. |
| `test_check_try_scope.py` | Saying when `make try` cannot test what was changed — including that a version bump on its own does not count. |

Each test file also ends with a canary that runs the real checker over the real source. If
one of those fails, `make build` is already refusing.

## Two rules for anything added here

**A test must never write into the repository.** Some of these scripts have side effects —
`build_ribbon_tabs.py` appends a newly added button id to `installer/ribbon-button-ids.txt`
even with `--check-only`. Where a canary runs against real files, it copies them to a
temporary folder first.

**Fixtures are CRLF.** `check_vba_structure.py` splits on CRLF, and a `.frm` with bare LF is
itself a defect. A fixture written with plain LF tests something the build never sees. Use
`write_crlf` from `conftest.py`.

## The VBA tests are separate

They need Word, so they run on the build box: `make vba-test`, and `tests/vba/README.md`. This
suite covers one piece of that machinery — `test_build_vba_test_bundle.py`, which checks that
the right VBA source is gathered for them. If it gathers the wrong thing the VBA tests do not
fail, they hang, so it is worth testing here where there is no Word to hang.
