#!/usr/bin/env python3
"""Refuse to overwrite a form edit that has not been pulled into src/ yet.

Jerry edits UserForm LAYOUT in the VBA editor, against the add-in Word actually loads:
%APPDATA%\\Microsoft\\Word\\STARTUP\\LPandBRL.dotm on the build box. He has to -- a .frx is
binary and cannot be authored from Linux.

Those edits live in that ONE file until someone exports them back into src/forms/. And every
`make try` copies a freshly built .dotm straight over it. So an unpulled edit dies quietly: the
build succeeds, nothing is reported, and the dialog simply goes back to how it looked before.
Word says nothing either -- as far as it is concerned the file it loads is merely newer.

Nothing detected that. This does, by comparing what is on the box against what this repo last
built, and `make try` stops when they disagree.

BY CONTENT, not by date. The obvious test -- "is the STARTUP copy newer than the last build?" --
cries wolf every time a real Setup.exe is installed, because installing rewrites the file and its
date with no edit involved. The installer ships the same bytes `make try` copies, so a hash gives
the right answer in both cases.

TWO hashes are acceptable, not one, and the second was learned the hard way: `make installer`
rebuilds dist/ WITHOUT touching the STARTUP folder, so from that moment the two legitimately
differ and every later `make try` was stopped. `make try` therefore records what it actually put
there (dist/.startup-hash, --stamp), and either answer passes. A form edited in the VBA editor
matches neither.

Word LOCKS that file while it is open, so it cannot be hashed then. That is not a problem for
`make try`, which needs Word closed anyway -- this just says so before the build rather than
after it, which saves two minutes and a version number. With --warn-only a locked file is
reported and allowed, for callers that do not require Word to be closed.

Set ALLOW_STARTUP_OVERWRITE=1 to go ahead anyway (see the message for when that is right).
"""
import base64
import hashlib
import os
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path

# Matches DOTM in the Makefile. Nothing derives it -- both name the shipping add-in.
DOTM = "LPandBRL.dotm"
BUILT = Path("dist") / DOTM
# What `make try` last copied INTO the STARTUP folder. dist/ is gitignored, so this is per
# machine and never travels. See main() for why one hash is not enough.
STAMP = Path("dist") / ".startup-hash"


def die(msg):
    sys.stderr.write(msg.rstrip() + "\n")
    sys.exit(1)


def read_host():
    """WIN_HOST out of build.config, the same file the Makefile reads."""
    host = os.environ.get("WIN_HOST", "").strip()
    if host:
        return host
    cfg = Path("build.config")
    if cfg.exists():
        m = re.search(r"^\s*WIN_HOST\s*=\s*(\S+)", cfg.read_text(), re.MULTILINE)
        if m:
            return m.group(1).strip()
    die("ERROR: WIN_HOST not set. Copy build.config.example -> build.config and edit it.")


def on_the_box(host):
    """What is in Word's STARTUP folder: hash, size, date, and whether Word has it open.

    -EncodedCommand because the default shell over SSH here is cmd.exe, which eats the quoting
    a Get-FileHash line needs. Base64 of UTF-16LE is what PowerShell expects.
    """
    ps = (
        "$p = Join-Path $env:APPDATA 'Microsoft\\Word\\STARTUP\\%s';"
        "Write-Output ('WORD=' + @(Get-Process WINWORD -ErrorAction SilentlyContinue).Count);"
        "if (-not (Test-Path $p)) { Write-Output 'ABSENT=1'; exit 0 };"
        "$f = Get-Item $p;"
        "Write-Output ('SIZE=' + $f.Length);"
        "Write-Output ('WHEN=' + $f.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'));"
        "try { Write-Output ('HASH=' + "
        "(Get-FileHash -LiteralPath $p -Algorithm SHA256 -ErrorAction Stop).Hash) }"
        "catch { Write-Output 'HASH=LOCKED' }" % DOTM
    )
    enc = base64.b64encode(ps.encode("utf-16-le")).decode("ascii")
    r = subprocess.run(["ssh", host, "powershell -NoProfile -EncodedCommand " + enc],
                       capture_output=True, text=True)
    if r.returncode != 0:
        die("ERROR: could not reach the build box (%s) to check Word's STARTUP folder.\n%s"
            % (host, r.stderr.strip()))
    out = {}
    for line in r.stdout.splitlines():
        if "=" in line:
            k, _, v = line.strip().partition("=")
            out[k] = v
    if "WORD" not in out:
        die("ERROR: could not read %s on %s. Got:\n%s" % (DOTM, host, r.stdout.strip()))
    return out


def main():
    warn_only = "--warn-only" in sys.argv[1:]
    stamping = "--stamp" in sys.argv[1:]
    host = read_host()
    box = on_the_box(host)

    if stamping:
        # Called by `make try` straight after it copies the add-in into the STARTUP folder, to
        # record what it put there. Never fatal: a missed stamp costs a false stop later, not a
        # broken build.
        h = box.get("HASH", "")
        if h and h != "LOCKED":
            STAMP.parent.mkdir(parents=True, exist_ok=True)
            STAMP.write_text(h.lower() + "\n")
        return 0

    if box.get("ABSENT"):
        # Nothing there to lose: a box that has never had the add-in installed or tried.
        return 0

    if box.get("HASH") == "LOCKED":
        if warn_only:
            print("  NOTE: Word is open on %s, so the add-in in its STARTUP folder could not be\n"
                  "        checked for an unpulled form edit." % host)
            return 0
        die("\nSTOP: Word is open on %s, so %s in its STARTUP folder is locked.\n"
            "\n"
            "Close Word on the build box and run this again. `make try` needs it closed in any\n"
            "case - this says so now rather than after the build, which saves a version number.\n"
            % (host, DOTM))

    if not BUILT.is_file():
        print("  NOTE: no %s yet, so an unpulled form edit cannot be detected this once." % BUILT)
        return 0

    theirs = box.get("HASH", "").lower()

    # TWO acceptable answers, not one, or this cries wolf on an ordinary day.
    #
    #   dist/LPandBRL.dotm -- what this repo last built. Matches after `make try` copies it, and
    #                         after the user installs a Setup.exe, which lays down the same bytes.
    #   dist/.startup-hash -- what `make try` last actually PUT there. Needed because
    #                         `make installer` rebuilds dist/ WITHOUT touching the STARTUP folder,
    #                         so from then on the two legitimately differ and comparing only
    #                         against dist/ stopped every later `make try`. Seen 8/26/2026, on the
    #                         first installer built after this guard existed.
    #
    # A form edited in the VBA editor matches NEITHER, which is the whole point.
    if theirs and theirs == hashlib.sha256(BUILT.read_bytes()).hexdigest():
        return 0
    if STAMP.is_file() and theirs and theirs == STAMP.read_text().strip().lower():
        return 0

    if os.environ.get("ALLOW_STARTUP_OVERWRITE", "").strip():
        print("  ALLOW_STARTUP_OVERWRITE is set - overwriting Word's STARTUP copy on %s." % host)
        return 0

    built_when = datetime.fromtimestamp(BUILT.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
    die(
        "\nSTOP: the add-in in Word's STARTUP folder on %s is not the one this repo last built.\n"
        "\n"
        "    on the box    %11s bytes   changed %s\n"
        "    last built    %11s bytes   changed %s   (%s)\n"
        "\n"
        "The usual reason is a UserForm edited in the VBA editor. Those edits live ONLY in that\n"
        "file. Building would copy straight over them, the build would succeed, nothing would be\n"
        "reported, and the dialog would quietly go back to how it looked before.\n"
        "\n"
        "  - Edited a form? Ask Claude to pull it into src/forms/ before building.\n"
        "  - Installed an older Setup.exe on purpose, or there is nothing there to keep?\n"
        "    Run it again as:\n"
        "\n"
        "        ALLOW_STARTUP_OVERWRITE=1 make try\n"
        % (host, format(int(box.get("SIZE", 0)), ","), box.get("WHEN", "?"),
           format(BUILT.stat().st_size, ","), built_when, BUILT)
    )


if __name__ == "__main__":
    sys.exit(main())
