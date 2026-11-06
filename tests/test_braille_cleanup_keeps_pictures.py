"""Braille Full File Cleanup keeps the pictures and enlarges the tiny ones (Jerry, 11/5/2026).

Duxbury ignores a picture but makes a transcriber's note from its alt text, so the pictures must
reach the DBT source file. "Deleting images" (Dx_Delete_Images) is gone from the sequence; a
picture smaller than the text is enlarged to the Normal size by large print's
Lp_Enlarge_Tiny_Pictures, as on the LP attach. Remove Pictures in Selection Clean Up still deletes
on request.
"""
import re
from pathlib import Path

SRC = Path(__file__).resolve().parent.parent / "src" / "vba" / "LPandBrlMacros.bas"
FORM = Path(__file__).resolve().parent.parent / "src" / "forms" / "Dx_Selected_Cleanup_Form.frm"


def _cleanup_body():
    text = SRC.read_text(encoding="latin-1")
    start = text.index("Public Sub Dx_Fix_Common_File_Errors_Run(")
    end = text.index("End Sub '***** end of Dx_Fix_Common_File_Errors_Run Macro *****", start)
    return text[start:end]


def _code_only(body):
    return "\n".join(l for l in body.splitlines() if not l.lstrip().startswith("'"))


def test_full_file_cleanup_does_not_delete_pictures():
    code = _code_only(_cleanup_body())
    assert "Dx_Delete_Images" not in code
    assert '"Deleting images"' not in code


def test_full_file_cleanup_enlarges_tiny_pictures():
    code = _code_only(_cleanup_body())
    assert re.search(r'Dx_Ffc_Step stepNo, "Enlarging pictures smaller than the text"\s*\n\s*'
                     r'Lp_Enlarge_Tiny_Pictures ActiveDocument', code)


def test_enlarge_runs_before_the_last_content_step():
    code = _code_only(_cleanup_body())
    assert code.index("Lp_Enlarge_Tiny_Pictures") < code.index('"Coloring the $pg tags red"')


def test_step_total_still_matches_the_passes():
    # DX_FFC_STEPS is one more than the passes (see Dx_Ffc_Step); one step out, one in.
    code = _code_only(_cleanup_body())
    passes = len(re.findall(r"Dx_Ffc_Step stepNo,", code))
    total = int(re.search(r"Const DX_FFC_STEPS As Long = (\d+)",
                          SRC.read_text(encoding="latin-1")).group(1))
    assert total == passes + 1


def test_selection_cleanup_still_offers_remove_pictures():
    assert '"Dx_Delete_Images"' in FORM.read_text(encoding="latin-1")
