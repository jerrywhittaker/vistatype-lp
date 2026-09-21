"""Format the TOC must never delete a line of the transcriber's contents page.

Jerry, 9/21/2026. From 3.0.410 to 3.0.474, Lp_TOC_CleanAndFormat_TOC deleted every line that
began with a bullet and had no page number at its end - a pass meant to drop a three-line
legend in one NIMAS book. On "TOC Table 2 Formatting.docx", where the page numbers sit behind a
non-breaking space and about 98 bulleted reading titles carry no page number at all, it cut 202
paragraphs to 10. Jerry's decision: the macro must never delete a line, and the pass was taken
out rather than narrowed.

This cannot be run as a VBA test: the pass works on a Range, and tests/vba/README.md lists
anything that reaches a Range as a hang. So it is a check on the source. It reads
Lp_TOC_CleanAndFormat_TOC and every Lp_TOC_ helper it calls, directly or through another one,
and fails if any of them can remove a paragraph:

  * a .Delete or .Cut on anything, or
  * a find-and-replace that takes away a paragraph mark - fewer ^p / ^013 in the replacement
    than in the text it finds.

What it deliberately does NOT catch: Lp_TOC_Join_Orphan_Page_Numbers writes a TAB over the
paragraph mark of the line above a page number standing on its own line. That joins two lines
into one entry and loses no text, and it is intended.
"""
import re

ROOT_SUB = "Lp_TOC_CleanAndFormat_TOC"

PROC_START = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function)\s+(\w+)",
    re.IGNORECASE)
PROC_END = re.compile(r"^\s*End\s+(Sub|Function)\b", re.IGNORECASE)
PARA_MARK = re.compile(r"\^p|\^013|\^13\b", re.IGNORECASE)


def strip_comment(line):
    """The code part of a VBA line - everything before an apostrophe that is not in a string."""
    in_string = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_string = not in_string
        elif ch == "'" and not in_string:
            return line[:i]
    return line


def procedures(text):
    """Every procedure in a module, as {name: [code lines with the comments taken off]}."""
    procs = {}
    name = None
    for raw in text.splitlines():
        code = strip_comment(raw)
        if name is None:
            m = PROC_START.match(code)
            if m:
                name = m.group(2)
                procs[name] = [code]
        else:
            procs[name].append(code)
            if PROC_END.match(code):
                name = None
    return procs


def reached_from(procs, root):
    """root, and every Lp_TOC_ procedure it calls, directly or through another one."""
    seen = set()
    todo = [root]
    while todo:
        name = todo.pop()
        if name in seen or name not in procs:
            continue
        seen.add(name)
        body = "\n".join(procs[name][1:])
        for callee in re.findall(r"\b(Lp_TOC_\w+)", body):
            if callee in procs and callee not in seen:
                todo.append(callee)
    return seen


def line_removers(procs, names):
    """Every place in the named procedures that can remove a paragraph, as (name, text)."""
    found = []
    for name in sorted(names):
        lines = procs[name]
        for code in lines:
            if re.search(r"\.(Delete|Cut)\b", code, re.IGNORECASE):
                found.append((name, code.strip()))
            # Lp_TOC_Replace <range>, <find>, <replace>, <wildcards>
            m = re.match(r'\s*Lp_TOC_Replace\s+\w+\s*,\s*"([^"]*)"\s*,\s*"([^"]*)"', code)
            if m and len(PARA_MARK.findall(m.group(1))) > len(PARA_MARK.findall(m.group(2))):
                found.append((name, code.strip()))
        # A With ... .Find block: compare its .Text with its .Replacement.Text.
        find_text = None
        for code in lines:
            m = re.match(r'\s*\.Text\s*=\s*"([^"]*)"', code)
            if m:
                find_text = m.group(1)
            m = re.match(r'\s*\.Replacement\.Text\s*=\s*"([^"]*)"', code)
            if m and find_text is not None:
                if len(PARA_MARK.findall(find_text)) > len(PARA_MARK.findall(m.group(1))):
                    found.append((name, f'find "{find_text}" -> "{m.group(1)}"'))
                find_text = None
    return found


def module_text(repo_root):
    path = repo_root / "src" / "vba" / "LPandBrlMacros.bas"
    return path.read_bytes().decode("latin-1")


# --- the checker itself, on made-up source --------------------------------------------------

FIXTURE = '''Sub Lp_TOC_CleanAndFormat_TOC()
    ' paraRange.Delete in a comment does not count
    Lp_TOC_Helper workRng
    Lp_TOC_Replace workRng, "^t", " ", False
End Sub
Private Sub Lp_TOC_Helper(ByVal r As Range)
    Lp_TOC_Inner r
End Sub
Private Sub Lp_TOC_Inner(ByVal r As Range)
    Dim s As String
    s = "it's fine"
End Sub
Private Sub Lp_TOC_Not_Called()
    Selection.Delete
End Sub
Sub Something_Else()
    ActiveDocument.Paragraphs(1).Range.Delete
End Sub
'''


def test_the_checker_follows_helpers_and_ignores_comments():
    procs = procedures(FIXTURE)
    names = reached_from(procs, ROOT_SUB)
    assert names == {ROOT_SUB, "Lp_TOC_Helper", "Lp_TOC_Inner"}
    assert line_removers(procs, names) == []


def test_the_checker_catches_a_delete_in_a_helper_two_calls_down():
    procs = procedures(FIXTURE.replace('s = "it\'s fine"', "r.Paragraphs(1).Range.Delete"))
    found = line_removers(procs, reached_from(procs, ROOT_SUB))
    assert [n for n, _ in found] == ["Lp_TOC_Inner"]


def test_the_checker_catches_a_replace_that_takes_away_a_paragraph_mark():
    procs = procedures(FIXTURE.replace('"^t", " "', '"^p^p", "^p"'))
    found = line_removers(procs, reached_from(procs, ROOT_SUB))
    assert [n for n, _ in found] == [ROOT_SUB]


def test_the_checker_catches_a_find_block_that_takes_away_a_paragraph_mark():
    block = ('    With r.Find\n        .Text = "^013^013"\n'
             '        .Replacement.Text = "^p"\n    End With\n')
    procs = procedures(FIXTURE.replace("    Dim s As String\n", block))
    found = line_removers(procs, reached_from(procs, ROOT_SUB))
    assert [n for n, _ in found] == ["Lp_TOC_Inner"]


def test_a_replace_that_keeps_the_paragraph_mark_is_allowed():
    block = ('    With r.Find\n        .Text = "^013^032{1,}"\n'
             '        .Replacement.Text = "^p"\n    End With\n')
    procs = procedures(FIXTURE.replace("    Dim s As String\n", block))
    assert line_removers(procs, reached_from(procs, ROOT_SUB)) == []


# --- the real source -------------------------------------------------------------------------

def test_format_the_toc_never_deletes_a_line(repo_root):
    procs = procedures(module_text(repo_root))
    assert ROOT_SUB in procs, "Lp_TOC_CleanAndFormat_TOC is missing from LPandBrlMacros.bas"
    names = reached_from(procs, ROOT_SUB)
    # The join is found, so the helpers really are being followed.
    assert "Lp_TOC_Join_Orphan_Page_Numbers" in names
    assert line_removers(procs, names) == []
