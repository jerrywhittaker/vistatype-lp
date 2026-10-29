"""DAISY/NIMAS conversion: the error handlers leave the handler state before On Error Resume Next.

Issue #47, Jerry, 10/27/2026, on 3.0.613: the log read `DN_Menu_Starter err 5825 "Object has
been deleted." step "Putting in equation 500 of 1,861."`, and the conversion stopped. Diagnosed
10/28/2026 on the build box: the converted book had been closed partway, so something in the
equation loop failed; ConvertFailed then ran `finalDoc.Activate` on the closed book. Inside an
error handler `On Error Resume Next` catches nothing (measured on the box), so that second 5825
escaped to RibbonAction and was logged under the button's name, and the first error was lost.

The fix: `On Error GoTo -1` straight after the handler takes its copy of the error, in
ConvertFailed and in Sh_MathML_Put_Picture's Failed; and a closed book gets box 403, which says
so, instead of an error report.
"""
from test_toc_never_deletes_a_line import module_text, procedures

CONVERTER = "Sh_Convert_XML_File_To_Word_Document"
PICTURE = "Sh_MathML_Put_Picture"


def _handler(lines, label):
    start = next(i for i, l in enumerate(lines) if l.strip() == label)
    return [l.strip() for l in lines[start:]]


def _leaves_state_first(body):
    resume = body.index("On Error Resume Next")
    assert "On Error GoTo -1" in body[:resume], "On Error GoTo -1 must come before On Error Resume Next"


def test_convert_failed_leaves_the_handler_state(repo_root):
    body = _handler(procedures(module_text(repo_root))[CONVERTER], "ConvertFailed:")
    # the copy of the error is taken first, then the state is left
    assert body.index("errText = Err.Description") < body.index("On Error GoTo -1")
    _leaves_state_first(body)


def test_put_picture_failed_leaves_the_handler_state(repo_root):
    _leaves_state_first(_handler(procedures(module_text(repo_root))[PICTURE], "Failed:"))


def test_a_closed_book_gets_box_403_not_an_error_report(repo_root):
    body = _handler(procedures(module_text(repo_root))[CONVERTER], "ConvertFailed:")
    text = "\n".join(body)
    assert "If Not finalDoc Is Nothing And Not bookGone Then" in text
    assert '"VistaType LP (403)"' in text
    assert text.index('"VistaType LP (403)"') < text.index("Sh_Report_Error")
