"""No dialog number is used by two different messages (issue #29).

The number in a message's title bar - "VistaType LP (nnn)", "Braille Macros (nnn)" - is how a
transcriber tells Jerry which message they saw, so a new message takes the next unused number
(docs/UI-Conventions.md). On 9/7/2026 Format the TOC's in-a-table refusal was given 321, which
the braille save prompt (Dx_Prompt_Save_Before_Braille) had had since 9/3/2026. It is 395 now.

A number may still appear on more than one line when it is the SAME dialog: a form's caption
and the window-title constant that finds it, one message said from two places in the same words,
or one message titled for either product with IIf. Those lines are the same once the product
name in the title is set aside. The one exception is listed below with its reason.
"""
import collections
import re

from test_toc_never_deletes_a_line import strip_comment

NUMBER = re.compile(r'"[^"]*\((\d{3})\)"')
PRODUCT_TITLE = re.compile(r'"(VistaType LP|Braille Macros) \(\d{3}\)"')
CAPTION_OR_TITLE_CONST = re.compile(r"^(Caption\s*=|((Private|Public)\s+)?Const\s)", re.IGNORECASE)

# Same dialog on two different lines, on purpose.
SAME_DIALOG = {
    "240": "Sh_Report_Error: the same report, asked when there is a log folder to open, said when not",
}


def message_lines(repo_root):
    """{number: {(file, code line with the product name set aside)}} for every titled message."""
    found = collections.defaultdict(set)
    src = repo_root / "src"
    for path in sorted([*src.rglob("*.bas"), *src.rglob("*.frm"), *src.rglob("*.cls")]):
        for raw in path.read_bytes().decode("latin-1").split("\r\n"):
            code = strip_comment(raw).strip()
            if CAPTION_OR_TITLE_CONST.match(code):
                continue
            for m in NUMBER.finditer(code):
                found[m.group(1)].add(PRODUCT_TITLE.sub('"TITLE"', code))
    return found


def test_no_number_names_two_different_messages(repo_root):
    found = message_lines(repo_root)
    clashes = {n: sorted(lines) for n, lines in found.items()
               if len(lines) > 1 and n not in SAME_DIALOG}
    assert not clashes, (
        "a dialog number is used by more than one message - give the newer one the next unused "
        f"number: {clashes}")


def test_the_two_321s_are_apart(repo_root):
    found = message_lines(repo_root)
    assert found["321"] == {'If Not Sh_Ask(body, "TITLE") Then Exit Sub'}
    assert len(found["395"]) == 1
    assert "use the table tools on a table." in next(iter(found["395"]))


def test_the_same_dialog_exceptions_are_still_needed(repo_root):
    """An exception that no longer matches anything would quietly excuse a future clash."""
    found = message_lines(repo_root)
    for number in SAME_DIALOG:
        assert len(found[number]) > 1, f"{number} is no longer used twice - take it out of SAME_DIALOG"
