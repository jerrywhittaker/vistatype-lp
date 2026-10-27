"""FormattedText is never assigned onto a range that may still hold text.

Issue #44, 10/27/2026. From Word 16.0.20527, `rng.FormattedText = other.FormattedText` onto a
range that HOLDS TEXT no longer replaces it: the new text goes in FRONT and the old text stays.
Measured on the build box - Format TOC turned Jerry's 29-paragraph "Problem TOC 1" into 44, the
formatted TOC with the whole original still under it. Assigning onto an EMPTY range inserts, on
every Word, so that is the only kind of assignment this allows.

A range is taken as empty when the procedure has done one of these to it since it was last Set:

  * `X.Delete`                      - the old text is gone (the measured cure)
  * `X.Collapse ...`                - it never held any
  * `X.End = X.start`               - the same, by hand
  * or it is written `doc.Range(n, n)`, the same position twice.

Anything else must be in EXEMPT below, with the reason it cannot hold text. Replacing text in the
book goes through Sh_Replace_With_Formatted_Text, which puts the new text in at a collapsed copy of
the range and then deletes the old text - so it passes the same rule.

This is a check on the source, and a textual one: it does not follow branches, so a Delete in one
branch also covers an assignment in another. It cannot run Word, and tests/vba/README.md lists
anything that reaches a Range as a hang for the VBA suite.
"""
import re

from test_toc_never_deletes_a_line import procedures

ASSIGN = re.compile(r"^\s*(?!Set\b)(.*?)\.FormattedText\s*=", re.IGNORECASE)
SAME_POSITION = re.compile(r"\.Range\(\s*(\w+)\s*,\s*\1\s*\)$")

# (module, procedure, target) -> why the target cannot hold text when it is assigned.
EXEMPT = {
    ("LPandBrlMacros", "Dx_Exercise_Levels_Hidden", "tempDoc.Content"):
        "a document made by Documents.Add a few lines above - only its own final mark",
    ("LPandBrlMacros", "Lp_Exercise_Levels_Hidden", "tempDoc.Content"):
        "a document made by Documents.Add a few lines above - only its own final mark",
    ("LPandBrlMacros", "Lp_Horz_To_Vert_Hidden", "tempDoc.Content"):
        "a document made by Documents.Add a few lines above - only its own final mark",
    ("LPandBrlMacros", "Lp_Convert_Table_To_Pseudo_Columns", "tRng"):
        "a cell of the table Tables.Add has just made, less its end-of-cell marker",
    ("LPandBrlMacros", "Lp_Table_Transpose_Table", "dstCellRng"):
        "a cell of the table Tables.Add has just made, less its end-of-cell marker",
}


def emptied_before(lines, index, target):
    """True when the target is shown empty between its last Set and the assignment at index."""
    if SAME_POSITION.search(target):
        return True
    t = re.escape(target)
    set_line = re.compile(r"^\s*Set\s+" + t + r"\s*=", re.IGNORECASE)
    emptied = re.compile(
        r"(?<![\w.])" + t + r"\.(Delete|Collapse)\b"
        r"|(?<![\w.])" + t + r"\.End\s*=\s*" + t + r"\.start\b",
        re.IGNORECASE)
    for code in reversed(lines[:index]):
        if emptied.search(code):
            return True
        if set_line.match(code):
            m = re.match(r"^\s*Set\s+" + t + r"\s*=\s*(.*)$", code, re.IGNORECASE)
            return bool(SAME_POSITION.search(m.group(1).strip()))
    return False


def assignments_over_text(modules):
    """Every FormattedText assignment not shown onto an empty range: (module, proc, target, line).

    modules is {module name: source text}.
    """
    found = []
    for module, text in sorted(modules.items()):
        for proc, lines in procedures(text).items():
            for i, code in enumerate(lines):
                m = ASSIGN.match(code)
                if not m:
                    continue
                target = m.group(1).strip()
                if (module, proc, target) in EXEMPT:
                    continue
                if not emptied_before(lines, i, target):
                    found.append((module, proc, target, code.strip()))
    return found


def source_modules(repo_root):
    mods = {}
    for folder in ("src/vba", "src/forms"):
        for path in sorted((repo_root / folder).iterdir()):
            if path.suffix.lower() in (".bas", ".frm", ".cls"):
                mods[path.stem] = path.read_bytes().decode("latin-1")
    return mods


# --- the checker itself, on made-up source --------------------------------------------------

FIXTURE = '''Sub Good_Ones()
    Set a = doc.Range(5, 5)
    a.FormattedText = src.FormattedText
    Set b = doc.Paragraphs(2).Range
    b.Collapse wdCollapseStart
    b.FormattedText = src.FormattedText
    Set c = Selection.Range
    c.Delete
    c.FormattedText = src.FormattedText
    Set d = tbl.Cell(1, 1).Range
    d.End = d.start
    d.FormattedText = src.FormattedText
    doc.Range(0, 0).FormattedText = src.FormattedText
    reader.Content.Delete
    reader.Content.FormattedText = src.FormattedText
End Sub
Sub Bad_One()
    Set work = Selection.Range
    ' work.Delete in a comment does not count
    work.FormattedText = body.FormattedText
End Sub
Sub Reset_By_Set()
    Set w = doc.Range(1, 1)
    w.Collapse wdCollapseEnd
    Set w = doc.Paragraphs(3).Range
    w.FormattedText = body.FormattedText
End Sub
Sub Other_Name_Deleted()
    homeRng2.Delete
    homeRng.FormattedText = body.FormattedText
End Sub
'''


def test_the_checker_passes_every_empty_landing():
    found = assignments_over_text({"Fixture": FIXTURE})
    assert [p for _, p, _, _ in found if p == "Good_Ones"] == []


def test_the_checker_catches_an_assignment_over_a_selection():
    found = assignments_over_text({"Fixture": FIXTURE})
    assert ("Fixture", "Bad_One", "work") in [(m, p, t) for m, p, t, _ in found]


def test_a_new_set_cancels_an_earlier_collapse():
    found = assignments_over_text({"Fixture": FIXTURE})
    assert "Reset_By_Set" in [p for _, p, _, _ in found]


def test_a_delete_of_a_different_range_does_not_count():
    found = assignments_over_text({"Fixture": FIXTURE})
    assert "Other_Name_Deleted" in [p for _, p, _, _ in found]


# --- the real source -------------------------------------------------------------------------

def test_no_formatted_text_assignment_over_text(repo_root):
    found = assignments_over_text(source_modules(repo_root))
    assert found == [], (
        "FormattedText assigned onto a range that may hold text - on Word 16.0.20527 that "
        "inserts in front and leaves the old text (issue #44). Use "
        "Sh_Replace_With_Formatted_Text, or Delete/Collapse first:\n"
        + "\n".join(f"  {m}.{p}: {line}" for m, p, _, line in found))


def test_every_exemption_still_exists(repo_root):
    """A stale exemption would quietly cover a NEW assignment of the same name."""
    present = set()
    for module, text in source_modules(repo_root).items():
        for proc, lines in procedures(text).items():
            for code in lines:
                m = ASSIGN.match(code)
                if m:
                    present.add((module, proc, m.group(1).strip()))
    assert set(EXEMPT) - present == set()


def test_the_book_is_replaced_through_the_helper(repo_root):
    """The four macros that bring a scratch document home over a selection use the helper."""
    procs = procedures(source_modules(repo_root)["LPandBrlMacros"])
    for name in ("Lp_TOC_CleanAndFormat_TOC", "Lp_Exercise_Levels_Hidden",
                 "Dx_Exercise_Levels_Hidden", "Lp_Horz_To_Vert_Hidden"):
        body = "\n".join(procs[name])
        assert re.search(r"^\s*Sh_Replace_With_Formatted_Text\s+\w+\s*,\s*\w+", body, re.M), name


def test_the_helper_inserts_first_then_deletes_the_old_text(repo_root):
    lines = procedures(source_modules(repo_root)["LPandBrlMacros"])[
        "Sh_Replace_With_Formatted_Text"]
    code = [c.strip() for c in lines if c.strip()]
    collapse = next(i for i, c in enumerate(code) if c.startswith("landing.Collapse"))
    assign = next(i for i, c in enumerate(code) if c.startswith("landing.FormattedText ="))
    delete = next(i for i, c in enumerate(code) if c.startswith("oldRng.Delete"))
    assert collapse < assign < delete
