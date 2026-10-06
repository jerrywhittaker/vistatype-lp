"""Setting table border weights leaves Word's own default border width alone.

Issue #2, 10/6/2026. Lp_Set_Table_Border_Weights wrote Options.DefaultBorderLineWidth - a
Word-wide setting, not the book's - and never put it back, so every border drawn afterwards in
any document, large print or not, came out at the large-print weight. The line did nothing the
sub needed: the loop sets .LineWidth on each edge of each table directly. The line is gone.

Not unit-testable in VBA: the sub walks ActiveDocument.Tables, and anything that reaches a
Document hangs the headless runner (tests/vba/README.md). So this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures, strip_comment

SUB = "Lp_Set_Table_Border_Weights"
ASSIGNMENT = re.compile(r"\bOptions\s*\.\s*DefaultBorderLineWidth\s*=", re.IGNORECASE)


def test_nothing_writes_the_word_wide_border_default(repo_root):
    for path in sorted((repo_root / "src" / "vba").glob("*.bas")):
        text = path.read_bytes().decode("latin-1")
        for n, line in enumerate(text.splitlines(), 1):
            assert not ASSIGNMENT.search(strip_comment(line)), (
                f"{path.relative_to(repo_root)}:{n} sets Options.DefaultBorderLineWidth. That is "
                "Word's own setting for every document, not the book's (issue #2).")


def test_the_sub_still_sets_each_edge(repo_root):
    procs = procedures(module_text(repo_root))
    assert SUB in procs, f"{SUB} is missing from LPandBrlMacros.bas"
    code = [strip_comment(line) for line in procs[SUB]]
    assert any(re.search(r"\.LineWidth\s*=\s*targetWeight\b", c) for c in code), (
        f"{SUB} no longer sets .LineWidth = targetWeight on the table edges - the borders "
        "would keep whatever weight they had.")
