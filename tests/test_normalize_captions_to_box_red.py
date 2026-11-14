"""Normalize Styles turns Caption paragraphs into Box Red (Jerry, 11/14/2026).

The pass must come AFTER Lp_Italics_To_Dashed_Underline: Caption's italic comes from the style,
and only once it is direct dashed underline does it survive the change to Box Red. It must come
BEFORE "Setting box border weights", so the new Box Red paragraphs get their borders in the same
run. Caption is found by wdStyleCaption so a Word in another language finds it.

The progress bar total, LP_NS_STEPS, must stay one more than the number of passes.

Not unit-testable in VBA: the pass works on a live Document, which hangs the headless runner
(tests/vba/README.md). So this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

SUB = "Lp_Normalize_Styles"
PASS = "Lp_Captions_To_Box_Red"


def _procs(repo_root):
    return procedures(module_text(repo_root))


def _index(lines, pattern):
    rx = re.compile(pattern, re.IGNORECASE)
    for i, line in enumerate(lines):
        if rx.search(line):
            return i
    return -1


def test_caption_pass_sits_between_italics_and_box_borders(repo_root):
    code = _procs(repo_root)[SUB]
    italics = _index(code, r"\bLp_Italics_To_Dashed_Underline\b")
    caption = _index(code, r"^\s*(?:Call\s+)?" + PASS + r"\b")
    borders = _index(code, r'Lp_Ns_Step\s+stepNo\s*,\s*"Setting box border weights"')
    assert italics >= 0, "Normalize Styles no longer dashes italics"
    assert caption >= 0, f"Normalize Styles does not call {PASS} directly"
    assert borders >= 0, "the box border step is missing"
    assert italics < caption < borders, (
        f"{PASS} must run after the italics step and before the box border steps")


def test_caption_pass_finds_caption_by_constant_and_writes_box_red(repo_root):
    code = "\n".join(_procs(repo_root)[PASS])
    assert re.search(r"Styles\(\s*wdStyleCaption\s*\)", code), "find Caption by wdStyleCaption"
    assert re.search(r'Styles\(\s*"Box Red"\s*\)', code), "the replacement style must be Box Red"
    assert re.search(r"\.Replacement\.Style\s*=\s*boxStyle", code)
    assert re.search(r"Exit Sub", code), "a missing style must leave the book alone"


def test_progress_total_is_one_more_than_the_passes(repo_root):
    procs = _procs(repo_root)
    steps = sum(1 for line in procs[SUB] if re.search(r"^\s*Lp_Ns_Step\s+stepNo\b", line))
    total = re.search(r"Const\s+LP_NS_STEPS\s+As\s+Long\s*=\s*(\d+)", "\n".join(procs["Lp_Ns_Step"]))
    assert total, "LP_NS_STEPS is missing"
    assert int(total.group(1)) == steps + 1, (
        f"{steps} passes need LP_NS_STEPS = {steps + 1}, not {total.group(1)}")
