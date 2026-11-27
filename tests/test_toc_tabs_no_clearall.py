"""LP attach: the TOC and Print Pg Num tab stops are set without TabStops.ClearAll (issue #56).

Two crash dumps of the second-attach crash (11/25 and 11/27/2026) both die inside Word's
TabStops.ClearAll, at the attach's 91% step, and the six calls in
Lp_Set_TOC_and_Print_Page_Num_Tab_Stops were the only ClearAll calls the attach makes. From
11/27/2026 that sub hands each style to Lp_Set_One_Style_Tab, which leaves a style holding the
one right tab stop alone and otherwise clears its stops one at a time.

These are checks on the source.
"""
from test_toc_never_deletes_a_line import module_text, procedures

SUB = "Lp_Set_TOC_and_Print_Page_Num_Tab_Stops"
HELPER = "Lp_Set_One_Style_Tab"
ATTACH = "Lp_Attach_The_Template"


def _code(lines):
    return [l for l in lines if not l.lstrip().startswith("'")]


def test_no_clearall_in_the_attach_tab_stops(repo_root):
    procs = procedures(module_text(repo_root))
    for name in (SUB, HELPER):
        assert not any("ClearAll" in l for l in _code(procs[name])), name


def test_every_style_goes_through_the_helper(repo_root):
    lines = _code(procedures(module_text(repo_root))[SUB])
    for style, leader in (("TOC 1", "Dots"), ("TOC 2", "Dots"), ("TOC 3", "Dots"),
                          ("TOC 4", "Dots"), ("TOC 5", "Dots"), ("Print Pg Num", "Spaces")):
        assert any(f'{HELPER} "{style}", tabPos, wdTabLeader{leader}' in l for l in lines), style
    assert not any("TabStops" in l for l in lines)


def test_the_helper_skips_a_style_already_right_and_clears_one_at_a_time(repo_root):
    text = module_text(repo_root)
    lines = _code(procedures(text)[HELPER])
    body = "\n".join(lines)
    assert f"Private Sub {HELPER}(ByVal styleName As String, ByVal tabPos As Single, ByVal tabLeader As Long)" in text
    assert "If Not Sh_Style_Exists(ActiveDocument, styleName) Then Exit Sub" in body
    assert "If stops.Count = 1 Then" in body and "Then Exit Sub" in body.split("If stops.Count = 1 Then")[1]
    assert "For i = stops.Count To 1 Step -1" in body and "stops(i).Clear" in body
    assert "stops.Add Position:=tabPos, Alignment:=wdAlignTabRight, Leader:=tabLeader" in body


def test_the_attach_still_calls_it_directly(repo_root):
    lines = _code(procedures(module_text(repo_root))[ATTACH])
    assert any(l.strip() == SUB for l in lines)
