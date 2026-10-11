"""Every pass of Braille Selection Clean Up gets the selection (issue #31).

Most of the passes Dx_Selected_Cleanup_Form runs end with Sh_Return_User_To_Start_Position and
Selection.Collapse. Until 10/11/2026 the form ran them back to back, so every pass after the
first ticked one found nothing selected and worked on the whole book. Found by reading; the form
now holds the selection as a range and selects it again before each pass and at the end.

Not unit-testable in VBA: a UserForm blocks the headless runner (tests/vba/README.md). So this
is a check on the source.
"""
import re

from test_toc_box_stays_open import form_code

FORM = "Dx_Selected_Cleanup_Form"


def _okay(repo_root):
    code = [line.split("'")[0] for line in form_code(repo_root, FORM)]
    start = next(i for i, l in enumerate(code) if re.search(r"Sub\s+Okay_Click\b", l, re.I))
    end = next(i for i in range(start, len(code)) if re.match(r"\s*End\s+Sub\b", code[i], re.I))
    return [l for l in code[start:end] if l.strip()]


def test_the_selection_is_held_as_a_range(repo_root):
    body = "\n".join(_okay(repo_root))
    assert re.search(r"Set\s+keep\s*=\s*Selection\.Range", body, re.I), (
        f"{FORM} no longer holds the selection as a range before the passes run.")


def test_every_pass_is_preceded_by_selecting_it_again(repo_root):
    lines = _okay(repo_root)
    runs = [i for i, l in enumerate(lines) if re.search(r"Application\.Run\b", l, re.I)]
    assert runs, f"{FORM} runs no passes?"
    for i in runs:
        before = "\n".join(lines[max(0, i - 2):i + 1])
        assert re.search(r"\bkeep\.Select\b", before, re.I), (
            f"{FORM} runs a pass without selecting the held range first ({lines[i].strip()}). "
            "The pass before it collapsed the selection, so this one would work on the whole book.")
