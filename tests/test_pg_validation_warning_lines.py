"""The $pg validation list carries warning lines, and Locate and Return step over them.

From 10/16/2026 (Jerry) Validate Reference Page Tags (367) checks the tags as it builds the list,
and puts a red line starting with a pointer mark (U+25BA) ahead of any entry that looks wrong,
with a count at the very top. Those lines are not entries. If Locate read one as a tag it would
search the book for "217 twice" and find nothing; if Return landed on one the transcriber would
have to step past it by hand on every warning.

So Locate takes the entry BELOW a warning line, Return skips them, and Locate goes to the very tag
an entry came from - a Range kept in list order when the list was built - before falling back to
searching for the text. None of that is unit-testable in VBA: it works on the Selection and on
two open documents, and anything that reaches a Document hangs the headless runner
(tests/vba/README.md). The checks themselves are: tests/vba/TestPgValCheck.bas. So this is a check
on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures, strip_comment

NOTE = re.compile(r"\bSh_PgVal_Is_Note_Line\b")


def _shnm(repo_root):
    path = repo_root / "src" / "vba" / "ShNonModalMessage.bas"
    return procedures(path.read_bytes().decode("latin-1"))


def _code(procs, name):
    assert name in procs, f"{name} is missing"
    return [strip_comment(c) for c in procs[name]]


def _calls(lines, name):
    return any(re.search(rf"\b{name}\b", c) for c in lines)


def test_the_mark_is_built_with_chrw_in_one_place(repo_root):
    text = module_text(repo_root)
    procs = procedures(text)
    body = _code(procs, "Sh_PgVal_Mark")
    assert any(re.search(r"ChrW\(\s*9658\s*\)", c) for c in body), (
        "Sh_PgVal_Mark must return ChrW(9658) - the module is ANSI and cannot hold the mark itself")
    others = [n for n, lines in procs.items() if n != "Sh_PgVal_Mark"
              and any(re.search(r"ChrW\(\s*9658\s*\)", strip_comment(c)) for c in lines)]
    assert not others, f"the mark is built outside Sh_PgVal_Mark: {others}"


def test_the_list_builder_checks_and_writes_the_count(repo_root):
    procs = procedures(module_text(repo_root))
    body = _code(procs, "Sh_PgVal_Copy_Tags_Into")
    assert _calls(body, "Sh_PgVal_Check"), "the list is no longer checked as it is built"
    assert _calls(body, "Sh_PgVal_Summary_Text"), "the count at the top of the list is gone"
    assert any(re.search(r"\bsources\.Add\b", c, re.IGNORECASE) for c in body), (
        "the list builder no longer keeps where each entry came from, so Locate cannot go to it")


def test_locate_reads_the_entry_below_a_warning_line(repo_root):
    procs = _shnm(repo_root)
    text_of = _code(procs, "Sh_PgVal_TextOfCurrentTag")
    assert _calls(text_of, "Sh_PgVal_TagParaAtCursor"), (
        "Sh_PgVal_TextOfCurrentTag must read the entry through Sh_PgVal_TagParaAtCursor")
    assert not any(re.search(r"Selection\.Paragraphs\(1\)", c) for c in text_of), (
        "Sh_PgVal_TextOfCurrentTag reads the cursor's own line again - on a warning line that "
        "searches the book for the warning")
    at_cursor = _code(procs, "Sh_PgVal_TagParaAtCursor")
    assert any(NOTE.search(c) for c in at_cursor), (
        "Sh_PgVal_TagParaAtCursor must step over warning lines with Sh_PgVal_Is_Note_Line")


def test_locate_goes_to_the_kept_range_first(repo_root):
    procs = _shnm(repo_root)
    body = _code(procs, "Sh_PgVal_LocateInDocument")
    stored = [i for i, c in enumerate(body) if re.search(r"\bSh_PgVal_StoredHit\b", c)]
    search = [i for i, c in enumerate(body) if re.search(r"\bSh_PgVal_FindInSource\b", c)]
    assert stored and search, "Locate must try the kept Range, then search for the text"
    assert stored[0] <= search[0], "Locate searches for the text before trying the kept Range"
    count = _code(procs, "Sh_PgVal_EntryNumber")
    assert any(NOTE.search(c) for c in count), (
        "Sh_PgVal_EntryNumber must not count warning lines, or every entry after one is "
        "matched to the wrong tag")


def test_return_skips_warning_lines(repo_root):
    body = _code(_shnm(repo_root), "Sh_PgVal_ReturnToTempAndAdvance")
    assert any(NOTE.search(c) for c in body), (
        "Return must step over warning lines and the count with Sh_PgVal_Is_Note_Line")


def test_done_lets_go_of_the_kept_ranges(repo_root):
    body = _code(_shnm(repo_root), "Sh_PgVal_Done")
    assert any(re.search(r"Set\s+Sh_PgVal_Sources\s*=\s*Nothing", c) for c in body), (
        "Done must empty Sh_PgVal_Sources")
