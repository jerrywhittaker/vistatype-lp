"""Fix Common File Errors takes the manual page breaks out of a braille file.

Issue 3, 9/24/2026. The "Removing page breaks" step of Dx_Fix_Common_File_Errors_Run ran
Dx_Remove_Breaks - the section-break pass - a second time, so braille cleanup never removed a
manual page break (Ctrl+Enter) and every one went on into the file DBT reads. DBT makes its own
pages, so a page break from Word has no place there. The step now runs Dx_Remove_Page_Breaks:

  * pass 1, wildcards on: (^013)^m^013{1,} -> \1. A page break on a line of its own goes, and
    the blank lines after it with it; the paragraph mark before it stays.
  * pass 2, wildcards on: ^m(^013) -> \1. A page break at the end of a paragraph goes; that
    paragraph's own mark stays.
  * pass 3, wildcards off: ^m -> ^p. A page break left inside a paragraph becomes a paragraph
    mark, so "alpha<page break>beta" becomes two paragraphs rather than "alphabeta".

The first version of pass 1, ^m^013{1,} -> nothing, ate the paragraph's own mark when a
paragraph ended in a page break: "alpha<break><para>beta" came out "alphabeta". Measured in
Word on the build box, 9/24/2026, on a .docx of that shape.

Pass 1 must come after Dx_Remove_Breaks: with wildcards on, Word's ^m finds section breaks as
well as page breaks, and those are Dx_Remove_Breaks' to deal with.

The same change deleted Dx_Remove_Section_Breaks and Dx_Red_Border_Images, which nothing
called. Dx_Remove_Section_Breaks replaced every ^b with nothing - the pattern that corrupted
books until 3.0.326 - so it must not come back.

Not unit-testable in VBA: both passes run Selection.Find on the document, and anything that
reaches a Document, Selection or Range hangs the headless runner (tests/vba/README.md). So this
is a check on the source; what the passes do to real text needs a run in Word.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

RUN_SUB = "Dx_Fix_Common_File_Errors_Run"
PAGE_SUB = "Dx_Remove_Page_Breaks"
GONE = ("Dx_Remove_Section_Breaks", "Dx_Red_Border_Images")


def the_procs(repo_root):
    procs = procedures(module_text(repo_root))
    for name in (RUN_SUB, PAGE_SUB, "Dx_Remove_Breaks"):
        assert name in procs, f"{name} is missing from LPandBrlMacros.bas"
    return procs


def macro_run_after_step(lines, step):
    """The macro the first code line after `Dx_Ffc_Step stepNo, "<step>"` runs, or None."""
    start = next((i for i, c in enumerate(lines)
                  if re.search(r'Dx_Ffc_Step\s+stepNo\s*,\s*"' + re.escape(step) + '"', c)),
                 None)
    assert start is not None, f'Fix Common File Errors has no "{step}" step'
    for c in lines[start + 1:]:
        if not c.strip():
            continue
        m = re.search(r'Application\.Run\s+MacroName:="(\w+)"', c)
        return m.group(1) if m else c.strip().split()[0]
    return None


def replace_passes(lines):
    """Each find-and-replace in a procedure, in order, as (text, replacement, wildcards).

    A setting not written in a With block keeps the value the block before it left, the way
    Selection.Find does.
    """
    passes = []
    text = repl = wild = None
    for c in lines:
        m = re.match(r'\s*\.Text\s*=\s*"(.*)"\s*$', c)
        if m:
            text = m.group(1)
        m = re.match(r'\s*\.Replacement\.Text\s*=\s*"(.*)"\s*$', c)
        if m:
            repl = m.group(1)
        m = re.match(r'\s*\.MatchWildcards\s*=\s*(True|False)\s*$', c, re.IGNORECASE)
        if m:
            wild = m.group(1).lower() == "true"
        if re.search(r"\.Execute\s+Replace:=wdReplaceAll", c):
            passes.append((text, repl, wild))
    return passes


def test_the_page_break_step_runs_the_page_break_macro(repo_root):
    lines = the_procs(repo_root)[RUN_SUB]
    ran = macro_run_after_step(lines, "Removing page breaks")
    assert ran == PAGE_SUB, (
        f'The "Removing page breaks" step runs {ran}, not {PAGE_SUB}. Until issue 3 it ran '
        "Dx_Remove_Breaks a second time, and no page break was ever removed.")


def test_section_breaks_go_once_and_before_the_page_breaks(repo_root):
    lines = the_procs(repo_root)[RUN_SUB]
    assert macro_run_after_step(lines, "Removing breaks") == "Dx_Remove_Breaks"
    runs = [i for i, c in enumerate(lines) if re.search(r'"Dx_Remove_Breaks"', c)]
    assert len(runs) == 1, "Dx_Remove_Breaks should run once in Fix Common File Errors"
    page = next(i for i, c in enumerate(lines) if re.search(r'"' + PAGE_SUB + '"', c))
    assert runs[0] < page, (
        "Section breaks must be gone before the page-break pass: with wildcards on, ^m "
        "finds section breaks too.")


def test_the_three_passes(repo_root):
    passes = replace_passes(the_procs(repo_root)[PAGE_SUB])
    assert passes == [("(^013)^m^013{1,}", "\\1", True),
                      ("^m(^013)", "\\1", True),
                      ("^m", "^p", False)], (
        "Dx_Remove_Page_Breaks should take out a page break on a line of its own with the "
        "blank lines after it, then a break at the end of a paragraph, keeping every "
        "paragraph mark that has text before it, then turn any break left inside a paragraph "
        "into a paragraph mark. Found: " + repr(passes))


def test_no_pass_deletes_a_break_with_the_mark_before_it(repo_root):
    """The 1.0 pattern ^m^013{1,} -> nothing ran "alpha<break><para>beta" into "alphabeta"."""
    for text, repl, _ in replace_passes(the_procs(repo_root)[PAGE_SUB]):
        assert not (text.startswith("^m^013") and repl == ""), (
            "A pass deletes a page break together with the paragraph mark after it. When a "
            "paragraph ends in a page break, that mark is the paragraph's own, and the words "
            "either side run together.")


def test_the_uncalled_macros_stay_gone(repo_root):
    src = repo_root / "src"
    for path in sorted(p for p in src.rglob("*") if p.is_file()):
        text = path.read_bytes().decode("latin-1")
        for name in GONE:
            # A dated note in the changelog may name it; a procedure or a call may not.
            for line in text.splitlines():
                code = line.split("'", 1)[0] if path.suffix in (".bas", ".frm", ".cls") else line
                assert name not in code, (
                    f"{name} is back in {path.relative_to(repo_root)}. It was deleted for "
                    "issue 3; Dx_Remove_Section_Breaks replaced every section break with "
                    "nothing.")


def test_a_break_at_the_start_of_the_document_is_taken_first(repo_root):
    """Pass 1 needs a paragraph mark before the break. At the very start of the document there
    is none, so pass 2 took the break and left a blank first paragraph (measured 9/24/2026).
    The start is dealt with on a Range before the Find passes run."""
    lines = the_procs(repo_root)[PAGE_SUB]
    code = [c.split("'", 1)[0] for c in lines]
    start = next((i for i, c in enumerate(code)
                  if "ActiveDocument.Range(0, 1).Text = Chr(12)" in c), None)
    assert start is not None, (
        "Dx_Remove_Page_Breaks no longer takes a page break at the start of the document; "
        "it would leave a blank first paragraph.")
    first_find = next(i for i, c in enumerate(code) if "Selection.Find.Execute" in c)
    assert start < first_find, "The start of the document must be dealt with before the passes."
    assert any("ActiveDocument.Content.End - 1" in c for c in code), (
        "The loop must stop short of the document's last paragraph mark, which Word will not "
        "delete.")
