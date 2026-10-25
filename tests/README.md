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
| `test_form_handlers_match_controls.py` | A UserForm event handler whose control is not on the form — renamed or removed — which never runs and raises nothing. Reads the control names out of each `.frx`. Needs `python3-olefile`; skipped without it. |
| `test_check_frm_eol.py` | A `.frm` that has lost its CRLF endings, which makes Word dump the designer header into the form's code module. |
| `test_check_vba_line_length.py` | A VBA line over 1023 characters, which makes `make build` hang with no error. |
| `test_trim_frm_blanks.py` | The blank lines Word piles up on every form export — and that trimming them keeps CRLF. |
| `test_build_ribbon_tabs.py` | The generated ribbon tabs, and the rule that a `btn_*` id which has shipped can never disappear. |
| `test_build_qat.py` | The Quick Access Toolbar files, and that the generated one can never hide a button the transcriber put there. |
| `test_inject_customui.py` | Embedding the ribbon in the built `.dotm`, and that doing it twice replaces rather than duplicates. |
| `test_check_try_scope.py` | Saying when `make try` cannot test what was changed — including that a version bump on its own does not count. |
| `test_toc_never_deletes_a_line.py` | Format the TOC deleting a line of the contents page — a `.Delete`, a `.Cut` or a replace that loses a paragraph mark in `Lp_TOC_CleanAndFormat_TOC` or any `Lp_TOC_` helper it calls. The bullet-legend pass cut one TOC from 202 paragraphs to 10 (9/21/2026). |
| `test_toc_blank_line_before_unnumbered.py` | Format the TOC leaving out the blank line before a contents line with no page number — the no-page-number branch of `Lp_TOC_CleanAndFormat_TOC` must call `Lp_TOC_Space_A_Heading` through `Lp_TOC_Line_Gets_Blank_Before`, and the page number must be looked for after the non-breaking spaces are gone. The text decision itself is `tests/vba/TestTocBlankLine.bas` (9/21/2026). |
| `test_toc_blue_page_numbers.py` | Format the TOC reading a wholly bold line whose page number is blue as a section heading, so it gets no tab and no TOC 1 — the heading test in `Lp_TOC_CleanAndFormat_TOC` must ask `Lp_TOC_Line_Is_Link_Entry`, the blue must come off (`Lp_TOC_Unblue`) only after the heading list is built, and it must become Automatic with no underline. The text-and-color decision itself is `tests/vba/TestTocLinkEntry.bas` (9/21/2026). |
| `test_toc_bold_entries.py` | Format the TOC reading every entry of a contents page that is bold throughout as a section heading, leaving its black-square bullets in, and giving no tab to an entry whose page number is followed by a manual line break — U+25A0 must be in the bullet class, the bullets must come off and the line breaks be split (`Lp_TOC_Split_Line_Breaks`, through `Lp_TOC_Break_Ends_Entry`) before the heading list, bold may mark a heading only through `Lp_TOC_Bold_Marks_Headings`, and the bold count must read each line through `Lp_TOC_Count_Text` so tabbed and leadered entries count; the split needs a bullet after the break. The two decisions themselves are `tests/vba/TestTocBoldEntries.bas` (9/28/2026). |
| `test_toc_short_title_second_tab.py` | Format the TOC leaving a short title's page number with no dot leader, because its tab stops at the TOC style's 0.75" hanging indent — the second-tab pass in `Lp_TOC_CleanAndFormat_TOC` must run after the TOC style, after every tab has been turned into a space (so a second run cannot stack them), bottom up in the hidden scratch document, and before the reference page passes. The decision itself is `tests/vba/TestTocShortTitle.bas` (9/21/2026). |
| `test_autocorrect_switch_checks_record.py` | The AutoCorrect switch trusting its record of which list is up when Word's list says otherwise, which saved the large print list over the user's ordinary one for good after Word ended without a proper quit (issue #21) — `Sh_AutoCorrect_Switch` may skip on the record only once `Sh_AC_Checked` says this session has checked it, must ask `Sh_AC_Identify` before its first save, must save into no file when that answer is blank, and must never seed a list from one; `Sh_AC_Checked` comes only from `Sh_AC_Remember`'s answer. The decision itself, and the real switch against Word's real list, are `tests/vba/TestAutoCorrectTables.bas` (9/29/2026). |
| `test_autotag_letter_hyphen_letter.py` | AutoTag Ref Pages tagging a paragraph made only of letters — "Introduction", "Contents" — as `$pg`. The "letter, hyphen, letter" pass in `Lp_AutoTag_Page_Numbers` had lost its `^045`; it and the braille pass in `Dx_AutoTag_Page_Numbers` must both read `^013([A-Za-z]{1,}^045)([A-Za-z]{1,})^013`. Fails on the source before the fix, passes after it (issue #35, 10/16/2026). |
| `test_pg_validation_warning_lines.py` | Validate Reference Page Tags (367) reading a warning line as a tag. From 10/16/2026 the list carries red lines starting with a pointer mark ahead of any entry that looks wrong, and a count at the top. Locate on one must take the entry below it (`Sh_PgVal_TagParaAtCursor`), Return must step over them, Locate must try the Range kept for that entry (`Sh_PgVal_StoredHit`) before searching for the text, the entry count must skip them, and the mark must be built only in `Sh_PgVal_Mark`. The checks themselves are `tests/vba/TestPgValCheck.bas` (10/16/2026). |
| `test_daisy_empty_pagenums.py` | Convert DAISY/NIMAS to Word tagging none of the page numbers of a DAISY book made from an EPUB, which writes each as an empty `<pagenum ... id="pagenum_0000163" />` (Jerry, 10/25/2026). The two text patterns must stay as they are, and `Sh_Tag_Empty_Pagenums` must run straight after them for DAISY and NIMAS alike. Its two patterns are lifted out of the source and run here with VBScript's `$` rules: the number from the end of the id, zeros dropped; `$pg?` with no id, no number or all zeros; a page number with text in it left alone. The helper itself is `tests/vba/TestEmptyPagenums.bas`. |
| `test_nimas_mathml.py` | The DAISY/NIMAS converter and its equations (issue #37, 10/18/2026). Box 325 ("Large Print or Braille") must answer `"LP"`/`"BRL"`, and the loop that deleted the pictures for "omit images" must stay gone. The book is read and looked at for MathML before any question, and every question comes before the bar, before a file is written and before the Word document exists. Only a large print book with MathML is asked "Equations (398)"; Cancel stops with a `Sh_Say` (399). The equations are parked as placeholders in a separate import file (the `.txt` and `.html` beside the book keep their MathML) and put back after the picture loop and the prodnote quote pass, before the font pass: Word equations through MSXML and `InsertXML`, falling back to the picture; pictures embedded, not linked. Braille gets Word equations (OMML) like large print, not asked (Jerry, 10/20/2026), and `Sh_MathML_Put_In` has only the `"OMML"` and `"PICTURE"` modes. No new `MsgBox`, every helper `ByVal`. The review of 10/18/2026: `Sh_One_Log_Line` (which turned every character from U+8000 up into a space) must read `AscW` as 0 to 65535; `MML2OMML.XSL` must be compiled once as an `XSLTemplate` before the walk, with `transformNode` as the fallback; `altimg` must be read from a quote-aware start tag and unescaped as XML then as a URL; Word's Find must be reset after the walk and on its error path; a picture must go in before its placeholder comes out. The MathML-as-text code and the braille cleanup park that protected it were removed on 10/20/2026; the last test checks that none of them is defined or called anywhere in `src/`. The string helpers themselves are `tests/vba/TestMathML.bas`. |
| `test_compress_linear_math_selection.py` | Braille Compress Linear Math stays inside the selection: no pass uses `wdFindContinue`, and the x-to-times-sign pass stops at the end of the selection (found by reading, 10/18/2026: every x in the book became a times sign). |
| `test_manual_tag_range_code.py` | Braille Manual Tag Ref Page tagging a page range without the DBT continuation code that braille AutoTag Ref Pages gives it — `16-18` must become `$pg16-18[[*lec*]][[*i*]]18`. `Dx_Manual_Tag_with_Dollar_pg` must close up a spaced hyphen with `Dx_Closed_Up_Pg_Range`, then type what `Dx_Pg_Range_Continuation` returns after the `$pg` tag and before `Dx_Set_DBT_Codes_Color_and_Style`; before it reads the line it must take off any old code at the line's end, reading the paragraph with `IncludeHiddenText` (`Dx_Pg_Continuation_Code_Start`), since the code is hidden and `Selection.Text` misses it, and the shared `Sh_Remove_DollarPG_For_Retag` must stay free of it; the helper must write the same code as AutoTag's passes and code the same five range shapes, read out of `Dx_AutoTag_Page_Numbers` itself. The helper's answers are `tests/vba/TestPageNumbers.bas` (10/17/2026). |

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
