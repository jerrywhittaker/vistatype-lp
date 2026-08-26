#!/usr/bin/env python3
"""Generate src/vba/RibbonDispatch.bas from the embedded ribbon.

WHY THIS EXISTS, measured on the build box 8/26/2026.

Every ribbon and toolbar button used to reach its macro through `Application.Run control.Tag`
in RibbonAction. That is one dispatcher for all 47 buttons, and it looked like the perfect
place to catch an error from any of them.

It is not. **Word does not pass an error back out of Application.Run.** It handles the error
itself and puts up its own "Run-time error" dialog -- the one with End and Debug -- and the
caller's On Error handler is never entered. Proved twice: a headless test hung on the modal
dialog, and then four real button presses on the build box wrote the marker BEFORE the call and
never the one after it or the one in the handler.

A DIRECT call does propagate normally. So the dispatcher calls the macro directly, through a
Select Case, and this writes that Select Case from the same customUI14.xml that defines the
buttons -- so the two can never drift. Hand-maintaining 47 branches would drift on the first
button anyone added, and the failure would be silent: a button that does nothing.

Run by `make build`, before the source is pushed to the Windows box.
"""
import re
import sys
from pathlib import Path

RIBBON = Path("src/ribbon/customUI14.xml")
OUT = Path("src/vba/RibbonDispatch.bas")

HEADER = '''Attribute VB_Name = "RibbonDispatch"
' VistaType LP - generated ribbon dispatch table.
'
' ############################################################################
' ##  GENERATED FILE - DO NOT EDIT.                                         ##
' ##  Written by tools/lib/build_ribbon_dispatch.py from                    ##
' ##  src/ribbon/customUI14.xml on every `make build`. Change the ribbon,   ##
' ##  not this file; anything typed here is lost on the next build.         ##
' ############################################################################
'
' WHY A SELECT CASE AND NOT Application.Run
'
' RibbonAction used to reach every macro with `Application.Run control.Tag`, and wrapped it in
' On Error GoTo so one handler could report a failure in any of the 47 buttons.
'
' That does not work. WORD DOES NOT PASS AN ERROR BACK OUT OF Application.Run. It takes the
' error itself and shows its own "Run-time error" dialog - the one with End and Debug, which
' also offers to open this source on the user's machine - and the caller's handler is never
' entered. Measured on the build box 8/26/2026: four real button presses wrote the marker set
' immediately BEFORE the call and never the one immediately after it, nor the one first thing
' in the handler.
'
' A DIRECT call propagates normally, which is what this table provides. Every macro named by a
' button's tag is a plain no-argument Sub, which is what makes the table possible at all; the
' generator checks that and refuses to write a table if it ever stops being true.
'
Option Explicit

' True when macroName was found and run. False means it is not on the ribbon, and the caller
' decides what to do about that - RibbonAction falls back to Application.Run, so a button can
' never do nothing at all just because this file is out of date.
'
Public Function Sh_Dispatch(ByVal macroName As String) As Boolean
    Sh_Dispatch = True

    Select Case macroName
'''

FOOTER = '''        Case Else
            Sh_Dispatch = False
    End Select
End Function   '*** end of Sh_Dispatch ***
'''


def die(msg):
    sys.stderr.write("ERROR: " + msg.rstrip() + "\n")
    sys.exit(1)


def main():
    if not RIBBON.is_file():
        die("%s not found." % RIBBON)

    tags = sorted(set(re.findall(r'tag="([^"]+)"', RIBBON.read_text(encoding="utf-8"))))
    if not tags:
        die("no tag= attributes in %s - refusing to write an empty dispatch table." % RIBBON)

    # Every one has to be a plain no-argument Sub, or the Select Case will not compile -- and
    # nothing in this pipeline compiles VBA, so it would ship and fail as "Compile error in
    # hidden module" on the user's machine.
    sources = "".join(p.read_text(encoding="cp1252", errors="replace").replace("\r\n", "\n")
                      for p in sorted(Path("src/vba").glob("*.bas")))
    bad = []
    for tag in tags:
        n = len(re.findall(r"^(?:Public )?Sub %s\(\)" % re.escape(tag), sources, re.M))
        if n != 1:
            bad.append("%s (%d plain no-argument definitions, expected 1)" % (tag, n))
    if bad:
        die("these ribbon tags are not callable directly:\n  " + "\n  ".join(bad) +
            "\nThe dispatch table cannot call them, so RibbonAction could not report their\n"
            "errors. Give each one a plain no-argument Sub, or take the button off the ribbon.")

    body = "".join('        Case "%s": %s\n' % (t, t) for t in tags)
    OUT.write_text(HEADER + body + FOOTER, encoding="cp1252", newline="\r\n")
    print("wrote %s (%d ribbon macros)" % (OUT, len(tags)))


if __name__ == "__main__":
    main()
