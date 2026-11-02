"""Strong and Emphasis to plain bold and italic: search first, change after.

Jerry, 11/2/2026: "Replacing Strong with bold and Emphasis with italic" took 340 to 450 s in every
File Cleanup of Basic Mat1, a NIMAS math book with 1,449 such runs. Lp_Style_To_Direct searched,
changed a run, and searched again. Now one read-only walk stores the runs, and only then is each
stored run changed: style off, bold or italic back on. Compared on a copy of that
book, the whole body XML, identical to the old version.

No Replace All: a style applied by Replace All throws direct bold away (measured 9/7/2026 and
11/2/2026), and in a book with Word equations a Replace All merged two of them (3.0.637).

A check on the source: the sub works on the book, which hangs the headless runner.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures, strip_comment

SUB = "Lp_Style_To_Direct"


def _code(repo_root):
    procs = procedures(module_text(repo_root))
    assert SUB in procs, f"{SUB} is missing from LPandBrlMacros.bas"
    return [strip_comment(c) for c in procs[SUB]]


def test_the_walk_changes_nothing(repo_root):
    code = _code(repo_root)
    start = next(n for n, c in enumerate(code) if re.search(r"Do\s+While\s+\.Execute", c, re.I))
    end = next(n for n, c in enumerate(code[start:], start) if re.match(r"^\s*Loop\b", c, re.I))
    walk = "\n".join(code[start:end])
    assert not re.search(r"\.(Style|Bold|Italic)\s*=", walk), (
        "the Find walk must only store where the runs are - changing a run between searches is "
        "what cost 340 to 450 s on a math book")


def test_each_stored_run_style_off_then_weight_and_no_replace_all(repo_root):
    code = _code(repo_root)
    joined = "\n".join(code)
    assert "wdReplaceAll" not in joined, (
        "no Replace All: in a book with Word equations a Replace All merged two equations with the "
        "text between them into one (3.0.637), and nothing here takes the equations out first")
    loop = next(n for n, c in enumerate(code) if re.search(r"For\s+i\s*=\s*0\s+To\s+runCount\s*-\s*1", c))
    after = "\n".join(code[loop:])
    style = after.find(".Style =")
    bold = after.find(".Bold = True")
    assert 0 <= style < bold, "style off first - it clears direct formatting - then the weight on"


def test_the_walk_cannot_go_round_for_ever(repo_root):
    joined = "\n".join(_code(repo_root))
    assert re.search(r"If\s+r\.End\s*<=\s*runEnd\(runCount\s*-\s*1\)\s+Then\s+Exit\s+Do", joined), (
        "a hit that does not move on must end the walk, or Word hangs with the screen off")
