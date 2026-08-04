#!/usr/bin/env python3
"""Refuse to build when a UserForm calls a macro that does not exist.

Nothing in the build compiles VBA. A form that calls a sub which has been renamed or removed
builds perfectly and then fails on the transcriber's machine as:

    Compile error in hidden module: Lp_About_Title_And_Agreement

which names the form and says nothing about the missing name. That happened on 8/3/2026 when
Sh_About_License_Text was removed from LPandBrlMacros and the two About dialogs were left
calling it.

The check is deliberately narrow, because a general VBA parser is not worth writing:

  * only names carrying the project's own prefixes are checked -- Lp_ Dx_ Sh_ MS_ DN_ Vt_
  * only the CODE half of a .frm is read; the designer header above the Attribute lines is
    layout, not code
  * a name is satisfied by a Sub, a Function, a Property or a module-level declaration
    anywhere under src/vba, or by another form's name (forms are referenced as objects)

That leaves the one thing worth catching: a form naming a module procedure that is gone.
"""
import glob
import os
import re
import sys

PREFIXES = ('Lp_', 'Dx_', 'Sh_', 'MS_', 'DN_', 'Vt_')
TOKEN = re.compile(r'\b((?:Lp_|Dx_|Sh_|MS_|DN_|Vt_)[A-Za-z0-9_]+)')
DEFINE = re.compile(
    r'^\s*(?:Public\s+|Private\s+|Friend\s+)?'
    r'(?:Static\s+)?'
    r'(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z0-9_]+)',
    re.M)
DECLARE = re.compile(
    r'^\s*(?:Public|Private|Dim|Global)\s+(?:WithEvents\s+)?([A-Za-z0-9_]+)', re.M)


def read(path):
    with open(path, 'rb') as fh:
        return fh.read().decode('cp1252', errors='replace')


def code_of_form(text):
    """Everything below the designer header -- the Attribute block ends it."""
    last = 0
    for m in re.finditer(r'^Attribute VB_.*$', text, re.M):
        last = m.end()
    return text[last:]


def main():
    defined = set()

    for path in sorted(glob.glob('src/vba/*.bas') + glob.glob('src/vba/*.cls')):
        text = read(path)
        defined.update(DEFINE.findall(text))
        defined.update(DECLARE.findall(text))

    # a form may be referenced by name, as an object
    for path in glob.glob('src/forms/*.frm'):
        defined.add(os.path.basename(path)[:-4])

    # Two different failures, so they are reported differently.
    #   direct call   -> the project will not COMPILE: "Compile error in hidden module: <form>"
    #   Application.Run "name" -> compiles fine, fails at RUN TIME when the button is pressed
    compile_breaks = []
    runtime_breaks = []

    for path in sorted(glob.glob('src/forms/*.frm')):
        code = code_of_form(read(path))
        form_defined = set(DEFINE.findall(code)) | set(DECLARE.findall(code))
        seen = set()
        for line in code.split('\n'):
            stripped = line.strip()
            if stripped.startswith("'"):
                continue
            late_bound = 'Application.Run' in stripped
            for name in TOKEN.findall(line):
                if name in defined or name in form_defined or name in seen:
                    continue
                seen.add(name)
                entry = (os.path.basename(path), name, stripped[:90])
                (runtime_breaks if late_bound else compile_breaks).append(entry)

    def show(items):
        for form, name, line in items:
            print('  %-38s %s' % (form, name))
            print('  %-38s   %s' % ('', line))

    if runtime_breaks:
        print('WARNING: %d macro name(s) run by a form but defined nowhere under src/vba.'
              % len(runtime_breaks))
        print('These COMPILE. They fail when the user presses the button, as')
        print('"The macro cannot be found or has been disabled".')
        show(runtime_breaks)
        print()

    if compile_breaks:
        print('ERROR: %d name(s) called by a form but defined nowhere under src/vba:'
              % len(compile_breaks))
        show(compile_breaks)
        print()
        print('This builds fine and then fails on the user\'s machine as')
        print('"Compile error in hidden module: <form name>".')
        sys.exit(1)

    print('form calls OK (%d names defined under src/vba)' % len(defined))


if __name__ == '__main__':
    main()
