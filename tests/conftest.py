"""Shared fixtures for the VistaType LP tool tests.

The scripts under tools/lib are run as scripts, not imported as a package -- there is no
__init__.py and there never needs to be. So put that folder on the import path here and
import each module by its plain name.
"""
import pathlib
import subprocess
import sys

import pytest

ROOT = pathlib.Path(__file__).resolve().parents[1]
TOOLS = ROOT / "tools" / "lib"

if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))


@pytest.fixture
def repo_root():
    """The real repository, for the few tests that read a tracked file."""
    return ROOT


@pytest.fixture
def run_tool():
    """Run one of the tools/lib scripts in a given folder and return the finished process.

    Several of the checkers glob for src/vba/*.bas relative to the CURRENT folder rather
    than taking a path, so the only way to test them on a fixture is to give them a folder
    that looks like the repository. That is what `cwd` is for.
    """
    def _run(script, cwd, *args):
        return subprocess.run(
            [sys.executable, str(TOOLS / script), *args],
            cwd=str(cwd), capture_output=True, text=True)
    return _run


def write_crlf(path, text):
    """Write VBA source the way Word writes it.

    Not a detail: check_vba_structure splits on CRLF, and a .frm with bare LF makes Word
    dump the designer header into the form's code module. A fixture written with plain LF
    tests something the build never sees.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(text.replace("\n", "\r\n").encode("cp1252"))
    return path
