"""Normalize Styles turns Caption paragraphs into Box Red (Jerry, 11/14/2026).

The pass must come BEFORE Lp_Italics_To_Dashed_Underline, and must put back the italic the
caption showed as direct italic, so the italics step dashes it. Measured on vistabuild
11/14/2026: a plain style change (Replace All or Paragraph.Style) AFTER the italics step stripped
most of the dashes and made the caption's not-italic words italic; the same change BEFORE it,
with nothing carried over, left the caption plain. It must also come before "Setting box border
weights", so the new Box Red paragraphs get their borders in the same run. Caption is found by
wdStyleCaption so a Word in another language finds it.

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


def test_caption_pass_runs_before_italics_and_box_borders(repo_root):
    code = _procs(repo_root)[SUB]
    italics = _index(code, r"\bLp_Italics_To_Dashed_Underline\b")
    caption = _index(code, r"^\s*(?:Call\s+)?" + PASS + r"\b")
    borders = _index(code, r'Lp_Ns_Step\s+stepNo\s*,\s*"Setting box border weights"')
    assert italics >= 0, "Normalize Styles no longer dashes italics"
    assert caption >= 0, f"Normalize Styles does not call {PASS} directly"
    assert borders >= 0, "the box border step is missing"
    assert caption < italics < borders, (
        f"{PASS} must run before the italics step and before the box border steps")


def test_caption_pass_finds_caption_by_constant_and_writes_box_red(repo_root):
    code = "\n".join(_procs(repo_root)[PASS])
    assert re.search(r"Styles\(\s*wdStyleCaption\s*\)", code), "find Caption by wdStyleCaption"
    assert re.search(r'Styles\(\s*"Box Red"\s*\)', code), "the replacement style must be Box Red"
    assert re.search(r"Exit Sub", code), "a missing style must leave the book alone"
    assert not re.search(r"\bwdReplaceAll\b", code), (
        "a Replace All of the style loses the caption's italic (measured 11/14/2026)")


def test_each_caption_keeps_the_italic_it_showed(repo_root):
    code = "\n".join(_procs(repo_root)["Lp_Caption_To_Box"])
    style = re.search(r"^\s*r\.Style\s*=\s*boxStyle", code, re.MULTILINE)
    read = re.search(r"whole\s*=\s*r\.Font\.Italic", code)
    assert style and read, "read the italic, then change the style"
    assert read.start() < style.start(), "the italic is read before the style change"
    after = code[style.end():]
    assert re.search(r"r\.Font\.Italic\s*=", after), "the paragraph's italic is set after the change"
    assert re.search(r"\.Range\(.*\)\.Font\.Italic\s*=\s*True", after), (
        "the italic stretches are put back after the change")


def test_progress_total_is_one_more_than_the_passes(repo_root):
    procs = _procs(repo_root)
    steps = sum(1 for line in procs[SUB] if re.search(r"^\s*Lp_Ns_Step\s+stepNo\b", line))
    total = re.search(r"Const\s+LP_NS_STEPS\s+As\s+Long\s*=\s*(\d+)", "\n".join(procs["Lp_Ns_Step"]))
    assert total, "LP_NS_STEPS is missing"
    assert int(total.group(1)) == steps + 1, (
        f"{steps} passes need LP_NS_STEPS = {steps + 1}, not {total.group(1)}")
