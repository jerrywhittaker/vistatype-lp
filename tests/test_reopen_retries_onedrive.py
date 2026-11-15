"""LP attach: the close-and-reopen tries again when the book Word hands back is not really there.

Issue #55, a tester, 11/11/2026, on 3.0.643: the error log read `Lp_Attach_The_Template err 5825
"Object has been deleted." step "Reopening the document so Word rebuilds the Quick Style
gallery"`. That user always saves to OneDrive. The book was already saved; the document object
Documents.Open returned was gone a moment later. THE CAUSE IS NOT REPRODUCED (VistaBuild has no
OneDrive): the theory is that the upload of the save was still running, or that FullName was an
https:// address, or both.

The fix, 11/15/2026: Sh_Close_And_Reopen tries up to REOPEN_TRIES = 5 times, REOPEN_WAIT_SECONDS
= 2 apart (both guesses), says so on the bar, looks for the book by name before opening it again,
and accepts only a book Sh_Doc_Still_There can read. After the last try it returns Nothing, and
the attach shows box 233. The finishing touches after the reopen run under Resume Next. The
error log ends with the document's folder, `path "..."`, so the next report shows whether Word
was working from a web address.

What this cannot test: anything to do with OneDrive itself.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

REOPEN = "Sh_Close_And_Reopen"
ATTACH = "Lp_Attach_The_Template"


def _body(repo_root, name):
    return "\n".join(procedures(module_text(repo_root))[name])


def test_the_two_numbers_are_named_and_are_5_and_2(repo_root):
    body = _body(repo_root, REOPEN)
    assert re.search(r"Const REOPEN_TRIES As Long = 5\b", body)
    assert re.search(r"Const REOPEN_WAIT_SECONDS As Single = 2\b", body)
    assert "For tryNo = 1 To REOPEN_TRIES" in body
    assert "Sh_PauseSeconds REOPEN_WAIT_SECONDS" in body


def test_the_bar_says_it_is_waiting_for_onedrive(repo_root):
    body = _body(repo_root, REOPEN)
    assert '"Waiting for OneDrive to finish with the document (try "' in body
    assert '" of " & CStr(REOPEN_TRIES) & ")"' in body


def test_each_try_looks_by_name_first_and_checks_the_book(repo_root):
    body = _body(repo_root, REOPEN)
    loop = body[body.index("For tryNo = 1 To REOPEN_TRIES"):body.index("Next tryNo")]
    assert loop.index("Sh_Find_Open_Doc(docPath)") < loop.index("Documents.Open(")
    assert "If Sh_Doc_Still_There(reopened) Then Exit For" in loop
    # the open is caught inside the loop, so a failed open gets the next try, not Bail
    opened = loop.index("Documents.Open(")
    assert loop.rindex("On Error Resume Next", 0, opened) > loop.index("Sh_Find_Open_Doc")


def test_a_failed_reopen_returns_nothing_so_the_attach_shows_box_233(repo_root):
    body = _body(repo_root, REOPEN)
    after = body[body.index("Next tryNo"):]
    assert "Set Sh_Close_And_Reopen = reopened" in after
    assert "If reopened Is Nothing Then Exit Function" in after
    # the only other document ever handed back is the one asked for (early exits) - never ActiveDocument
    assert "ActiveDocument" not in body
    attach = _body(repo_root, ATTACH)
    gone = attach[attach.index("Set doc = Sh_Close_And_Reopen(doc)"):]
    assert gone.index("If doc Is Nothing Then") < gone.index('"VistaType LP (233)"') < gone.index("doc.Activate")


def test_read_only_is_read_under_resume_next(repo_root):
    body = _body(repo_root, REOPEN)
    ro = body.index("isReadOnly = reopened.ReadOnly")
    assert body.rindex("On Error Resume Next", 0, ro) > body.index("Next tryNo")


def test_the_still_there_check_reads_the_name_and_a_window(repo_root):
    body = _body(repo_root, "Sh_Doc_Still_There")
    assert "On Error Resume Next" in body
    assert "d.FullName" in body and "d.Windows(1)" in body


def test_the_finishing_touches_cannot_raise_and_attachfailed_comes_back(repo_root):
    attach = _body(repo_root, ATTACH)
    tail = attach[attach.index("Set doc = Sh_Close_And_Reopen(doc)"):attach.index("AttachFailed:")]
    resume = tail.index("On Error Resume Next")
    assert resume < tail.index("doc.Activate") < tail.index("ActiveWindow.View.Type")
    assert tail.index("Lp_Turn_on_Styles_Pane") < tail.index("On Error GoTo AttachFailed")


def test_the_error_log_ends_with_the_documents_folder(repo_root):
    body = _body(repo_root, "Sh_Error_Context")
    assert 's = s & "  path """ & ActiveDocument.Path & """"' in body
    # last on the line, so nothing before it moves
    assert body.index("ActiveDocument.Path") > body.index("wdStyleNormal")
