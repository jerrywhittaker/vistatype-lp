#!/usr/bin/env python3
"""Say when `make try` is not enough and a real Setup.exe is needed.

`make try` puts the freshly built add-in straight into Word's STARTUP folder on the build box.
That tests the VBA and the UserForms exactly as they would ship -- it is the same .dotm the
installer packages -- and it skips compiling and running the installer.

What it CANNOT test is anything the installer puts somewhere else. Those changes look like they
did nothing, which is the expensive way to find out: the code is right, the test says it is
wrong, and the next hour goes on the wrong question.

So: after a `make try`, this prints what in the working tree needs a real installer instead.

It reads the WORKING TREE (git status). That covers the ordinary loop, where a change is being
tested before it is committed. It will not see something already committed and never installed --
which is why the message says "at least", not "only".
"""
import re
import subprocess
import sys

# path pattern -> why a .dotm swap does not cover it
NEEDS_INSTALLER = [
    (re.compile(r'^installer/'),
     'the installer itself: the Quick Access Toolbar, the ribbon tabs it writes,\n'
     '       trusted locations, the license file, the artwork, uninstalling'),
    (re.compile(r'^src/ribbon/'),
     'the ribbon. The tabs ARE embedded in the .dotm, but on a machine the installer\n'
     '       has already set up, the copy in the user\'s own Word.officeUI is what shows and the\n'
     '       embedded one is hidden (VtTabVisible). So a tab change can look like it did nothing'),
    (re.compile(r'^src/keymap/'),
     'the keyboard shortcuts: they are injected into LargePrintTemplate.dotx, which\n'
     '       only the installer puts in place'),
    (re.compile(r'\.dotx$'),
     'the large print template: styles and page setup, installed to Templates'),
    (re.compile(r'^assets/fonts/'),
     'the bundled typeface and its licenses, installed and registered by the installer'),
]


def only_the_version_bump(path):
    """True when the only thing changed in this file is the version number.

    `make try` bumps, and the bump edits installer/vistatype.iss - so without this the scope
    check fires on EVERY run, naming the installer as a change that needs an installer. A warning
    that always fires is a warning nobody reads, which is worse than none.
    """
    out = subprocess.run(['git', 'diff', '-U0', '--', path],
                         capture_output=True, text=True).stdout
    changed = [ln for ln in out.splitlines()
               if (ln.startswith('+') or ln.startswith('-'))
               and not ln.startswith('+++') and not ln.startswith('---')]
    if not changed:
        return False
    return all('#define AppVer' in ln or 'APPVER' in ln for ln in changed)


def changed_paths():
    out = subprocess.run(['git', 'status', '--porcelain'],
                         capture_output=True, text=True).stdout
    paths = []
    for line in out.splitlines():
        if len(line) < 4:
            continue
        p = line[3:].strip()
        # "old -> new" for a rename; the new name is what matters
        if ' -> ' in p:
            p = p.split(' -> ', 1)[1]
        paths.append(p.strip('"'))
    return paths


def main():
    hits = []
    for path in changed_paths():
        for pattern, why in NEEDS_INSTALLER:
            if pattern.search(path):
                if only_the_version_bump(path):
                    break
                hits.append((path, why))
                break

    print()
    print('  What this build does NOT put on the machine: the toolbar, the ribbon tabs the')
    print('  installer writes, LargePrintTemplate.dotx, the keyboard shortcuts and the fonts.')

    if not hits:
        print('  Nothing changed in the working tree needs them, so this is a fair test.')
        return 0

    print()
    print('  BUT AT LEAST THESE CHANGES NEED A REAL INSTALLER - `make installer`:')
    seen = set()
    for path, why in hits:
        print('    %s' % path)
        if why not in seen:
            print('       -> %s' % why)
            seen.add(why)
    print()
    print('  Testing them from this .dotm will show you nothing, or the old behavior.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
