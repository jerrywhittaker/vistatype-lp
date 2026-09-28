"""Attach LP Template no longer stops with 5904 at "Adding a blank line after each table".

9/28/2026. Diagnosed and reproduced on Jerry's book on the build box: run-time error 5904 "Cannot
edit Range" at nextP.Range.InsertParagraphBefore in Lp_Add_Blank_After_Tables, and not caught.

Why it raised: when the one paragraph between two tables has a HIDDEN paragraph mark, Word joins
it to the next paragraph - the next table's first cell. doc.Range(tEnd, tEnd).Paragraphs(1) then
starts outside a table and ends inside the next one, Information(wdWithInTable) reads False, and
the insert raises. 10 of 187 tables in his book; 8 of those paragraphs empty, 2 a visible bullet
and two spaces with only the mark hidden.

Why it was not caught: Lp_Attach_The_Template ran File Cleanup through Application.Run, and Word
never passes an error back out of Application.Run (the 3.0.256/3.0.257 register row), so
AttachFailed never ran.

What this checks:

  1. Lp_Add_Blank_After_Tables unhides from the table's end through the next paragraph mark,
     with hidden text included, BEFORE the Paragraphs(1) lookup - and unconditionally, not only
     when Font.Hidden reads True (that missed the bullet paragraphs on the build box).
  2. Lp_Attach_The_Template calls Lp_Fix_Common_File_Errors directly, never through
     Application.Run.

Not unit-testable in VBA: it needs a Document with tables, and the headless runner hangs on that
(tests/vba/README.md). So this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

TABLES = "Lp_Add_Blank_After_Tables"
ATTACH = "Lp_Attach_The_Template"
CLEANUP = "Lp_Fix_Common_File_Errors"


def the_procs(repo_root):
    procs = procedures(module_text(repo_root))
    for name in (TABLES, ATTACH, CLEANUP):
        assert name in procs, f"{name} is missing from LPandBrlMacros.bas"
    return procs


def first_line(lines, pattern):
    return next((i for i, c in enumerate(lines) if re.search(pattern, c)), None)


def test_the_gap_is_unhidden_before_the_lookup(repo_root):
    lines = the_procs(repo_root)[TABLES]
    tend = first_line(lines, r"^\s*tEnd\s*=\s*doc\.Tables\(i\)\.Range\.End\s*$")
    start = first_line(lines, r"^\s*Set\s+(\w+)\s*=\s*doc\.Range\(tEnd,\s*tEnd\)\s*$")
    assert tend is not None and start is not None, "a range must be set at the table's end"
    r = re.match(r"^\s*Set\s+(\w+)", lines[start]).group(1)
    include = first_line(lines, rf"^\s*{r}\.TextRetrievalMode\.IncludeHiddenText\s*=\s*True\s*$")
    until = first_line(lines, rf"^\s*{r}\.MoveEndUntil\s+vbCr\s*$")
    past = first_line(lines, rf"^\s*{r}\.MoveEnd\s+wdCharacter,\s*1\s*$")
    unhide = first_line(lines, rf"^\s*{r}\.Font\.Hidden\s*=\s*False\s*$")
    lookup = first_line(lines, r"Set nextP = doc\.Range\(tEnd, tEnd\)\.Paragraphs\(1\)")
    assert None not in (include, until, past, unhide), \
        f"{TABLES} must unhide from the table's end through the next paragraph mark"
    assert lookup is not None, f"{TABLES} no longer looks up the paragraph after the table"
    assert tend < start < include < until < past < unhide < lookup, \
        "the unhide must come after the table's end is known and before the Paragraphs(1) lookup"


def test_the_unhide_is_unconditional(repo_root):
    lines = the_procs(repo_root)[TABLES]
    unhide = first_line(lines, r"\.Font\.Hidden\s*=\s*False")
    assert unhide is not None
    assert not re.match(r"^\s*If\b", lines[unhide]), "the unhide must not sit on an If line"
    # Nothing between the For and the unhide may open an If - the bullet paragraphs on the build
    # box did not read Font.Hidden = True, and a test on it left their marks hidden.
    loop = first_line(lines, r"^\s*For i = doc\.Tables\.count To 1 Step -1")
    assert loop is not None and loop < unhide
    assert not any(re.match(r"^\s*If\b", c) for c in lines[loop:unhide]), \
        "the unhide must run for every table, not only when Font.Hidden reads True"


def test_the_attach_calls_file_cleanup_directly(repo_root):
    procs = the_procs(repo_root)
    lines = procs[ATTACH]
    body = "\n".join(lines)
    assert not re.search(rf'Application\.Run\s+MacroName:="{CLEANUP}"', body), \
        f"{ATTACH} must not reach {CLEANUP} through Application.Run - errors never come back out"
    assert first_line(lines, rf"^\s*{CLEANUP}\s*$") is not None, f"{ATTACH} must call {CLEANUP} directly"
    assert re.match(rf"^\s*(Public\s+)?Sub\s+{CLEANUP}\(\)\s*$", procs[CLEANUP][0]), \
        f"{CLEANUP} must stay a no-argument Sub for the direct call to compile"
