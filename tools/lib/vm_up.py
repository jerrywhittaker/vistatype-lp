#!/usr/bin/env python3
"""Start the Windows build VM, if it is not already up and answering.

WHY THIS EXISTS
---------------
Everything that touches Word runs on the build box, and if the VM is off, every one of those
targets fails at the SSH with a connection error that says nothing about the cause. Starting it
by hand first is a step nobody should have to remember.

WSL can run Windows programs directly, and the VM lives on the same Windows host as this WSL,
so `VBoxManage.exe startvm` is reachable from here.

THE ORDER MATTERS, and it is deliberately cautious:

  1. If SSH already answers, do nothing at all and say nothing much. That is the common case,
     and it means this never interferes on a machine where the box is reached some other way.
  2. Only then look for VBoxManage, and only start the VM if it is one this host knows about.
  3. If there is no VirtualBox here, say so plainly rather than failing with a stack trace -
     the box may be somebody else's machine entirely.

Headless by default: a build wants no window. `--gui` starts it with one, which is what you
want for anything that has to SEE Word - screenshots, or driving the ribbon by hand. A headless
VM is still fully booted and still reachable over SSH; it just has no window on the host.

Usage:  python3 tools/lib/vm_up.py --host vistabuild [--name VistaBuild] [--gui] [--wait 180]
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import time

DEFAULT_NAME = "VistaBuild"
CANDIDATES = (
    "/mnt/c/Program Files/Oracle/VirtualBox/VBoxManage.exe",
    "/mnt/c/Program Files (x86)/Oracle/VirtualBox/VBoxManage.exe",
)

VM_LINE = re.compile(r'^"(.*)"\s+\{[0-9a-fA-F-]+\}\s*$')


def parse_vm_list(text):
    """The machine names out of `VBoxManage list vms` / `list runningvms`.

    Each line is  "Name" {uuid}  -- and a name may contain spaces, so take what is inside the
    quotes rather than splitting on whitespace.
    """
    names = []
    for line in text.splitlines():
        m = VM_LINE.match(line.strip())
        if m:
            names.append(m.group(1))
    return names


def find_vboxmanage(explicit=None):
    for path in ([explicit] if explicit else []) + list(CANDIDATES):
        if path and os.path.exists(path):
            return path
    return shutil.which("VBoxManage.exe") or shutil.which("VBoxManage")


def ssh_ok(host, timeout=8):
    """True when the box answers. BatchMode so a missing key fails rather than prompting."""
    r = subprocess.run(
        ["ssh", "-o", "BatchMode=yes", "-o", "ConnectTimeout=%d" % timeout,
         "-o", "StrictHostKeyChecking=accept-new", host, "exit"],
        capture_output=True, text=True)
    return r.returncode == 0


def vbox(vboxmanage, *args):
    return subprocess.run([vboxmanage, *args], capture_output=True, text=True)


def wait_for_ssh(host, seconds, tick=5):
    """Poll until the box answers. Returns the seconds taken, or None if it never did."""
    started = time.time()
    while time.time() - started < seconds:
        if ssh_ok(host, timeout=tick):
            return int(time.time() - started)
        time.sleep(tick)
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--host", required=True, help="the SSH host name for the box")
    ap.add_argument("--name", default=os.environ.get("VM_NAME") or DEFAULT_NAME)
    ap.add_argument("--vboxmanage", default=os.environ.get("VBOXMANAGE"))
    ap.add_argument("--gui", action="store_true",
                    help="start it with a window, for anything that has to SEE Word")
    ap.add_argument("--wait", type=int, default=180,
                    help="seconds to wait for SSH after starting it")
    args = ap.parse_args()

    if ssh_ok(args.host):
        print("build box %s is up" % args.host)
        return 0

    vboxmanage = find_vboxmanage(args.vboxmanage)
    if not vboxmanage:
        print("%s is not answering, and VirtualBox is not on this machine, so it cannot be\n"
              "started from here. Start it yourself, or set VBOXMANAGE in build.config."
              % args.host, file=sys.stderr)
        return 1

    known = parse_vm_list(vbox(vboxmanage, "list", "vms").stdout)
    if args.name not in known:
        print("%s is not answering, and VirtualBox on this machine has no VM called %r.\n"
              "It knows: %s\nSet VM_NAME in build.config if it is one of those."
              % (args.host, args.name, ", ".join(known) or "(none)"), file=sys.stderr)
        return 1

    if args.name in parse_vm_list(vbox(vboxmanage, "list", "runningvms").stdout):
        print("%s is already running but not answering on SSH yet - waiting." % args.name)
    else:
        mode = "gui" if args.gui else "headless"
        print("%s is not answering. Starting the %s VM (%s)..." % (args.host, args.name, mode))
        r = vbox(vboxmanage, "startvm", args.name, "--type", mode)
        if r.returncode != 0:
            print("could not start it:\n%s" % (r.stderr.strip() or r.stdout.strip()),
                  file=sys.stderr)
            return 1

    took = wait_for_ssh(args.host, args.wait)
    if took is None:
        print("%s started but did not answer on SSH within %d seconds. It may still be\n"
              "booting, or waiting at the sign-in screen. Try again in a moment."
              % (args.name, args.wait), file=sys.stderr)
        return 1

    print("%s is up and answering after %d seconds." % (args.host, took))
    return 0


if __name__ == "__main__":
    sys.exit(main())
