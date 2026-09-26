"""Bookmarks a VistaType LP macro leaves in a book are taken out when the book opens.

9/26/2026, 3.0.510. Three macros mark a place in the book with a bookmark while they work and
delete it before they end:

  * Export Selection to New File - DxExportStart and DxExportEnd. It also SAVES the book with
    them in it and does not save again after removing them (docs/Reported-Errors.md, 8/30/2026).
  * Format $pg Tags, large print and braille - TempPgNoFormat. Until this change nothing caught
    an error between adding it and deleting it.

If Word crashes part-way, and the book was saved or AutoRecover kept it, the bookmark stays in
the book for good. Jerry asked for two things:

  1. Sh_HandleDocumentOpened calls Sh_Remove_Leftover_Bookmarks FIRST - before the obsolete
     template branch, which leaves by Exit Sub. It deletes those three names and no others, and
     a book that opened unchanged is marked unchanged again, so Word does not ask "Save
     changes?" about a book the user never touched.
  2. Both Format $pg Tags macros go to eom on an error, delete the bookmark there, and raise the
     error again so RibbonAction still reports it.

Not unit-testable in VBA: all of it reaches a Document, and the headless runner hangs on that
(tests/vba/README.md). So this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

HELPER = "Sh_Remove_Leftover_Bookmarks"
OPEN = "Sh_HandleDocumentOpened"
NAMES = ("DxExportStart", "DxExportEnd", "TempPgNoFormat")
PG_MACROS = ("Dx_Format_Tagged_Page_Numbers", "Lp_Format_Page_Numbers")


def the_procs(repo_root):
    procs = procedures(module_text(repo_root))
    for name in (HELPER, OPEN) + PG_MACROS:
        assert name in procs, f"{name} is missing from LPandBrlMacros.bas"
    return procs


def first_line(lines, pattern):
    return next((i for i, c in enumerate(lines) if re.search(pattern, c)), None)


def test_the_open_handler_cleans_up_before_anything_can_leave_early(repo_root):
    lines = the_procs(repo_root)[OPEN]
    call = first_line(lines, rf"^\s*{HELPER}\s+ActiveDocument\s*$")
    assert call is not None, f"{OPEN} does not call {HELPER} ActiveDocument"
    # Only the license skip may come first. Any other way out, or the configuration, must follow.
    for pattern in (r"Sh_Apply_Word_Config", r"^(?!.*Sh_Skip_Open_Handler).*\bExit Sub\b",
                    r"On Error GoTo eom"):
        later = first_line(lines, pattern)
        assert later is None or later > call, f"'{pattern}' comes before the cleanup in {OPEN}"


def test_the_helper_removes_exactly_the_three_names(repo_root):
    body = "\n".join(the_procs(repo_root)[HELPER])
    quoted = set(re.findall(r'"([^"]*)"', body))
    assert quoted == set(NAMES), f"{HELPER} names {sorted(quoted)}, expected {sorted(NAMES)}"
    assert "Bookmarks.Exists(" in body and ".Delete" in body


def test_the_helper_cannot_raise_and_keeps_an_unchanged_book_unchanged(repo_root):
    lines = the_procs(repo_root)[HELPER]
    header = lines[0]
    assert re.search(r"\(ByVal Doc As Document\)", header), "the Document must be passed ByVal"
    resume = first_line(lines, r"On Error Resume Next")
    first_touch = first_line(lines, r"\bDoc\.")
    assert resume is not None and resume < first_touch, "On Error Resume Next must come first"
    assert first_line(lines, r"On Error GoTo 0") is None
    captured = first_line(lines, r"wasSaved\s*=\s*Doc\.Saved")
    deleted = first_line(lines, r"\.Delete")
    assert captured is not None and captured < deleted, "Saved must be read before any delete"
    assert first_line(lines, r"If removedAny And wasSaved Then Doc\.Saved = True") is not None


def test_both_pg_macros_delete_the_bookmark_on_a_failure_and_report_it(repo_root):
    procs = the_procs(repo_root)
    for name in PG_MACROS:
        lines = procs[name]
        trap = first_line(lines, r"On Error GoTo eom")
        added = first_line(lines, r'Bookmarks\.Add Name:="TempPgNoFormat"')
        assert trap is not None and added is not None and trap < added, \
            f"{name}: the error trap must be set before the bookmark is added"
        label = first_line(lines, r"^eom:")
        assert label is not None, f"{name} has no eom label"
        assert re.search(r"^\s*Exit Sub\s*$", lines[label - 2] + lines[label - 1], re.M), \
            f"{name}: the normal path must leave by Exit Sub before eom"
        handler = "\n".join(lines[label:])
        assert 'Bookmarks("TempPgNoFormat").Delete' in handler, f"{name}: eom does not delete it"
        assert f'Err.Raise errNum, "{name}", errText' in handler, \
            f"{name}: eom must raise the error again so RibbonAction reports it"
        # Nothing between the trap and eom may switch it off again.
        between = lines[trap + 1:label]
        assert not any(re.search(r"On Error (Resume Next|GoTo 0)", c) for c in between), \
            f"{name}: an On Error between the trap and eom would switch the cleanup off"


def test_the_large_print_macro_puts_the_screen_back_on_a_failure(repo_root):
    lines = the_procs(repo_root)["Lp_Format_Page_Numbers"]
    handler = "\n".join(lines[first_line(lines, r"^eom:"):])
    assert "Application.ScreenUpdating = su_Prev" in handler
