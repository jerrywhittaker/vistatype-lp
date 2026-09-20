"""check_style_guards.py -- the run-time error 5941 gate.

ActiveDocument.Styles("X") dies with 5941 on a document that does not carry the style.
Jerry hit it on 8/5/2026 running AutoTag Ref Pages over an already-tagged braille file.
"""
import check_style_guards as mod


def bas(tmp_path, text, name="Mod.bas"):
    path = tmp_path / name
    path.write_bytes(text.encode("utf-8"))
    return path


def test_an_unguarded_lookup_is_reported(tmp_path):
    path = bas(tmp_path, """\
Sub Lp_One()
    ActiveDocument.Styles("Print Pg Num").Font.Bold = True
End Sub
""")
    bad = mod.check(path)
    assert len(bad) == 1
    lineno, name, _src = bad[0]
    assert lineno == 2
    assert name == "Print Pg Num"


def test_normal_is_exempt_because_every_document_has_it(tmp_path):
    path = bas(tmp_path, """\
Sub Lp_One()
    ActiveDocument.Styles("Normal").Font.Bold = True
End Sub
""")
    assert mod.check(path) == []


def test_a_matching_guard_covers_the_lookup(tmp_path):
    path = bas(tmp_path, """\
Sub Lp_One()
    If Sh_Style_Exists(ActiveDocument, "Print Pg Num") Then
        ActiveDocument.Styles("Print Pg Num").Font.Bold = True
    End If
End Sub
""")
    assert mod.check(path) == []


def test_sh_style_in_use_counts_as_a_guard(tmp_path):
    path = bas(tmp_path, """\
Sub Lp_One()
    If Sh_Style_In_Use(ActiveDocument, "RefPageNemeth") Then
        ActiveDocument.Styles("RefPageNemeth").Delete
    End If
End Sub
""")
    assert mod.check(path) == []


def test_a_guard_for_a_different_style_does_not_cover_it(tmp_path):
    """The failure that looks guarded in a diff and is not."""
    path = bas(tmp_path, """\
Sub Lp_One()
    If Sh_Style_Exists(ActiveDocument, "Print Pg Num") Then
        ActiveDocument.Styles("RefPageNemeth").Font.Bold = True
    End If
End Sub
""")
    bad = mod.check(path)
    assert len(bad) == 1
    assert bad[0][1] == "RefPageNemeth"


def test_the_guard_stops_applying_after_end_if(tmp_path):
    path = bas(tmp_path, """\
Sub Lp_One()
    If Sh_Style_Exists(ActiveDocument, "Print Pg Num") Then
        ActiveDocument.Styles("Print Pg Num").Font.Bold = True
    End If
    ActiveDocument.Styles("Print Pg Num").Font.Italic = True
End Sub
""")
    bad = mod.check(path)
    assert len(bad) == 1
    assert bad[0][0] == 5


def test_the_older_on_error_set_form_is_accepted(tmp_path):
    """Set st = Styles(...) under On Error Resume Next, then test it for Nothing."""
    path = bas(tmp_path, """\
Sub Lp_One()
    Dim st As Style
    On Error Resume Next
    Set st = ActiveDocument.Styles("Print Pg Num")
    On Error GoTo 0
    If Not st Is Nothing Then st.Font.Bold = True
End Sub
""")
    assert mod.check(path) == []


def test_the_on_error_form_stops_counting_after_on_error_goto_zero(tmp_path):
    path = bas(tmp_path, """\
Sub Lp_One()
    On Error Resume Next
    Set st = ActiveDocument.Styles("Print Pg Num")
    On Error GoTo 0
    Set st = ActiveDocument.Styles("RefPageNemeth")
End Sub
""")
    bad = mod.check(path)
    assert len(bad) == 1
    assert bad[0][1] == "RefPageNemeth"


def test_a_commented_out_lookup_is_ignored(tmp_path):
    path = bas(tmp_path, """\
Sub Lp_One()
    ' ActiveDocument.Styles("Print Pg Num").Font.Bold = True
End Sub
""")
    assert mod.check(path) == []


def test_the_real_repository_passes_its_own_check(repo_root):
    bad = []
    for f in sorted((repo_root / "src" / "vba").glob("*.bas")) + \
             sorted((repo_root / "src" / "vba").glob("*.cls")):
        bad.extend(mod.check(f))
    assert bad == []
