"""Real Columns closes its undo record and puts the screen back on an error (issue #30).

Found by reading, 10/11/2026, reviewing the #28 fix: Lp_Convert_Table_To_Real_Columns started
its "Convert Table to Columns" undo record and turned screen updating off with no handler in
force (an earlier On Error Resume Next ended with On Error GoTo 0). An error after that point
would have shown Word's own dialog with the record still open and the screen frozen. Never seen
in Word.

It now ends as Lp_Convert_Table_To_Pseudo_Columns does since 3.0.557: su_Prev taken before
On Error GoTo ErrorHandler, and a handler that restores screen updating, closes the record
If Not objUndo Is Nothing, and says what went wrong - through Sh_Say, as every new message does.

Not unit-testable in VBA: the sub needs a real table in a document and shows a UserForm, which
hang the headless runner (tests/vba/README.md). So this is a check on the source.
"""
import re

from test_toc_never_deletes_a_line import module_text, procedures

SUB = "Lp_Convert_Table_To_Real_Columns"


def _code(repo_root):
    procs = procedures(module_text(repo_root))
    assert SUB in procs, f"{SUB} is missing from LPandBrlMacros.bas"
    return procs[SUB]


def _first(body, pattern):
    return next((n for n, c in enumerate(body) if re.search(pattern, c, re.IGNORECASE)), None)


def test_handler_in_force_from_su_prev_to_the_end(repo_root):
    body = _code(repo_root)
    take = _first(body, r"su_Prev\s*=\s*Application\.ScreenUpdating")
    handler = _first(body, r"On Error GoTo ErrorHandler")
    record = _first(body, r"StartCustomRecord")
    assert take is not None and handler is not None and record is not None, (
        f"{SUB} must take su_Prev, set On Error GoTo ErrorHandler and start its undo record")
    assert take < handler < record, (
        f"{SUB}: su_Prev must be taken before On Error GoTo ErrorHandler, and the handler set "
        "before the undo record starts")
    assert not any(re.search(r"On Error GoTo 0\b", c, re.IGNORECASE) for c in body), (
        f"{SUB} switches error handling off with On Error GoTo 0 - nothing would close the undo "
        "record or put the screen back")


def test_handler_closes_the_record_and_restores_the_screen(repo_root):
    body = _code(repo_root)
    label = _first(body, r"^\s*ErrorHandler:")
    assert label is not None, f"{SUB} has no ErrorHandler: label"
    assert re.match(r"^\s*Exit Sub\s*$", [c for c in body[:label] if c.strip()][-1], re.IGNORECASE), (
        f"{SUB} must Exit Sub before ErrorHandler:, or a good run falls into the handler")
    tail = "\n".join(body[label:])
    assert re.search(r"Application\.ScreenUpdating\s*=\s*su_Prev", tail, re.IGNORECASE), (
        f"{SUB}'s handler must put screen updating back")
    assert re.search(r"If\s+Not\s+objUndo\s+Is\s+Nothing\s+Then\s+objUndo\.EndCustomRecord", tail,
                     re.IGNORECASE), (
        f"{SUB}'s handler must close the undo record If Not objUndo Is Nothing")
    assert re.search(r'Sh_Say\s+"Error: "\s*&\s*failText\s*,\s*"VistaType LP \(396\)"', tail), (
        f"{SUB}'s handler must say what went wrong through Sh_Say (message 396)")
    assert not re.search(r"\bMsgBox\b", tail, re.IGNORECASE), (
        f"{SUB}'s handler is a new message - it goes through Sh_Say, not MsgBox")
