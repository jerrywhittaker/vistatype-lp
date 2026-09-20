"""vm_up.py -- starting the build VM when a target needs it.

The VM lives on the same Windows host as this WSL, and WSL can run Windows programs, so
VBoxManage is reachable. The order in vm_up.py is deliberately cautious and these pin it:
answering SSH beats everything, and a machine with no VirtualBox must say so rather than fail.
"""
import vm_up as mod

LIST_VMS = '''"VistaBuild" {a4b4bb1c-93ab-499e-9439-290d008cbedd}
"Some Other Box" {11111111-2222-3333-4444-555555555555}
'''


def test_the_machine_names_are_read_out_of_the_listing():
    assert mod.parse_vm_list(LIST_VMS) == ["VistaBuild", "Some Other Box"]


def test_a_name_with_spaces_survives():
    """Split on whitespace and "Some Other Box" becomes three machines that do not exist."""
    assert "Some Other Box" in mod.parse_vm_list(LIST_VMS)


def test_nothing_running_reads_as_nothing():
    assert mod.parse_vm_list("") == []
    assert mod.parse_vm_list("\n\n") == []


def test_a_line_that_is_not_a_machine_is_ignored():
    noise = 'WARNING: something\n"VistaBuild" {a4b4bb1c-93ab-499e-9439-290d008cbedd}\n'
    assert mod.parse_vm_list(noise) == ["VistaBuild"]


def test_an_answering_box_is_left_alone(monkeypatch, capsys):
    """The common case. It must not go looking for VirtualBox at all, so that a machine
    reaching the box some other way is never interfered with."""
    monkeypatch.setattr(mod, "ssh_ok", lambda host, timeout=8: True)
    monkeypatch.setattr(mod, "find_vboxmanage",
                        lambda explicit=None: (_ for _ in ()).throw(
                            AssertionError("must not look for VirtualBox")))
    monkeypatch.setattr("sys.argv", ["vm_up.py", "--host", "vistabuild"])
    assert mod.main() == 0
    assert "is up" in capsys.readouterr().out


def test_no_virtualbox_here_says_so_instead_of_failing(monkeypatch, capsys):
    monkeypatch.setattr(mod, "ssh_ok", lambda host, timeout=8: False)
    monkeypatch.setattr(mod, "find_vboxmanage", lambda explicit=None: None)
    monkeypatch.setattr("sys.argv", ["vm_up.py", "--host", "vistabuild"])
    assert mod.main() == 1
    err = capsys.readouterr().err
    assert "VirtualBox is not on this machine" in err
    assert "Start it yourself" in err


def test_an_unknown_machine_name_lists_what_there_is(monkeypatch, capsys):
    monkeypatch.setattr(mod, "ssh_ok", lambda host, timeout=8: False)
    monkeypatch.setattr(mod, "find_vboxmanage", lambda explicit=None: "/fake/VBoxManage.exe")
    monkeypatch.setattr(mod, "vbox",
                        lambda vb, *a: type("R", (), {"stdout": LIST_VMS, "returncode": 0})())
    monkeypatch.setattr("sys.argv", ["vm_up.py", "--host", "vistabuild", "--name", "Typo"])
    assert mod.main() == 1
    err = capsys.readouterr().err
    assert "no VM called 'Typo'" in err
    assert "VistaBuild" in err


def test_a_stopped_machine_is_always_started_with_a_window(monkeypatch, capsys):
    """Jerry, 9/20/2026: he drives every setup by hand to test it, so a VM that came up
    headless is no use to him. There is no headless mode."""
    calls = []

    def fake_vbox(vb, *a):
        calls.append(a)
        out = LIST_VMS if a[:2] == ("list", "vms") else ""
        return type("R", (), {"stdout": out, "returncode": 0, "stderr": ""})()

    monkeypatch.setattr(mod, "ssh_ok", lambda host, timeout=8: False)
    monkeypatch.setattr(mod, "find_vboxmanage", lambda explicit=None: "/fake/VBoxManage.exe")
    monkeypatch.setattr(mod, "vbox", fake_vbox)
    monkeypatch.setattr(mod, "wait_for_ssh", lambda host, seconds, tick=5: 42)
    monkeypatch.setattr("sys.argv", ["vm_up.py", "--host", "vistabuild"])

    assert mod.main() == 0
    assert ("startvm", "VistaBuild", "--type", "gui") in calls
    out = capsys.readouterr().out
    assert "with a window" in out
    assert "after 42 seconds" in out


def test_there_is_no_way_to_ask_for_headless(monkeypatch):
    """Deleted rather than left as an option nobody should take. If it comes back, so does a
    VM Jerry cannot look at."""
    import io, contextlib
    monkeypatch.setattr("sys.argv", ["vm_up.py", "--host", "vistabuild", "--headless"])
    with contextlib.redirect_stderr(io.StringIO()):
        try:
            mod.main()
        except SystemExit as exc:
            assert exc.code == 2      # argparse refuses the unknown flag
        else:
            raise AssertionError("--headless was accepted")


def test_a_running_machine_that_is_not_answering_yet_is_not_started_again(monkeypatch, capsys):
    """Starting an already-running VM fails; it is mid-boot, so just wait for it."""
    calls = []

    def fake_vbox(vb, *a):
        calls.append(a)
        return type("R", (), {"stdout": LIST_VMS, "returncode": 0, "stderr": ""})()

    monkeypatch.setattr(mod, "ssh_ok", lambda host, timeout=8: False)
    monkeypatch.setattr(mod, "find_vboxmanage", lambda explicit=None: "/fake/VBoxManage.exe")
    monkeypatch.setattr(mod, "vbox", fake_vbox)
    monkeypatch.setattr(mod, "wait_for_ssh", lambda host, seconds, tick=5: 7)
    monkeypatch.setattr("sys.argv", ["vm_up.py", "--host", "vistabuild"])

    assert mod.main() == 0
    assert not any(a[0] == "startvm" for a in calls)
    assert "already running but not answering" in capsys.readouterr().out


def test_a_machine_that_never_answers_is_reported_not_assumed(monkeypatch, capsys):
    def fake_vbox(vb, *a):
        out = LIST_VMS if a[:2] == ("list", "vms") else ""
        return type("R", (), {"stdout": out, "returncode": 0, "stderr": ""})()

    monkeypatch.setattr(mod, "ssh_ok", lambda host, timeout=8: False)
    monkeypatch.setattr(mod, "find_vboxmanage", lambda explicit=None: "/fake/VBoxManage.exe")
    monkeypatch.setattr(mod, "vbox", fake_vbox)
    monkeypatch.setattr(mod, "wait_for_ssh", lambda host, seconds, tick=5: None)
    monkeypatch.setattr("sys.argv", ["vm_up.py", "--host", "vistabuild", "--wait", "5"])

    assert mod.main() == 1
    assert "did not answer on SSH" in capsys.readouterr().err
