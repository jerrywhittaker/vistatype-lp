# Build/test VM setup

The VistaType LP build runs Word on a Windows virtual machine, driven from Linux over SSH.
Word is the only step that cannot happen on this side: only its own VBA engine can regenerate
the compiled p-code inside a `.dotm`.

> **This VM moved from Hyper-V to VirtualBox on 20 September 2026**, done by Todd. The reason
> was the code-signing hardware: the Certum card reader is an **ACR40T ICC Reader**, a CCID
> smart card reader (USB class 0x0B), and under Hyper-V the only way into the guest was RDP
> smart card redirection — which is per-session and therefore invisible to an SSH session, so
> unattended signing was impossible. VirtualBox does real USB passthrough and the reader is
> visible to every guest session, Session 0 included.
>
> **The host-side creation steps in this file were written for Hyper-V and have been removed**
> — `Enable-WindowsOptionalFeature`, `New-VMSwitch`, `New-VM`, `Enable-VMTPM`, the
> `netsh portproxy` recipe and `Checkpoint-VM`. None of them applies now, and writing
> VirtualBox equivalents from memory would be guessing. **If you ever build a second box, write
> that section then, from what you actually did.** Everything below is inside the guest or on
> the Linux side, so it is hypervisor-independent and still correct.

## What the box must end up with

| | |
|---|---|
| Windows | 11, any edition — VirtualBox does not need Pro, which Hyper-V did |
| Office | Word, 64-bit. Verified 9/20/2026 as Office Pro Plus 2021 |
| Reachable as | `ssh vistabuild` — the alias in `~/.ssh/config` carries the address, user and key, so none of it is written down here |
| SSH session | gets a full admin token, so no elevation step is needed |
| Inno Setup 6 | `C:\Program Files (x86)\Inno Setup 6\ISCC.exe`, for `make installer` |
| Windows SDK signing tools | `C:\Program Files (x86)\Windows Kits\10\bin\<ver>\x64\signtool.exe` |
| Office signing libraries | `C:\vt-signing\officesips-x64` and `-x86`, registered machine-wide |
| `AccessVBOM` | `1`, under `HKCU\Software\Microsoft\Office\16.0\Word\Security` |

Sizing that has been enough: **6 GB RAM, 64 GB disk, 2 CPUs.**

## The smart card, and what is not a fault

Measured on 9/20/2026, from an ordinary SSH session (Session 0):

- **The card layer works.** `SCardEstablishContext`, `SCardListReaders` and `SCardConnect` all
  return `0x00000000`, and the ATR comes back carrying `Certum01`.
- **`certutil -scinfo` fails** with `0x80070005` on `SCardAccessStartedEvent`. That is a Session 0
  artifact, **not** a passthrough failure. Probe through `winscard.dll` instead.
- **Microsoft's two smart card providers are refused** from Session 0 — both
  `Microsoft Base Smart Card Crypto Provider` and `Microsoft Smart Card Key Storage Provider`
  return `0x80090010 NTE_PERM`. A software CSP enumerates fine in the same session, so it is
  those providers specifically. **`cryptoCertum3 CSP` works**, and is what the signing plan
  should name. See `docs/Code-Signing.md`.
- **The card shows as "Unknown Smart Card" in Device Manager.** Cosmetic.
- **The Windows host cannot see the reader while the VM is running.** USB passthrough is
  exclusive. Expected; do not try to fix it.
- **BitLocker was decrypted on C:** as part of the move, deliberately. Do not re-enable it.


## Configure the guest

Inside the running guest, elevated PowerShell:

```powershell
# --- OpenSSH Server ---
Add-WindowsCapability -Online -Name OpenSSH.Server~~~~0.0.1.0
Start-Service sshd
Set-Service -Name sshd -StartupType Automatic
Get-NetFirewallRule -Name *ssh*      # confirm "OpenSSH SSH Server (sshd)" is enabled
```

> **Do NOT change the default SSH shell to PowerShell.** The Makefile uses cmd syntax
> (e.g. `if not exist ... mkdir`) over SSH, and invokes PowerShell explicitly where it
> needs it. Leaving the default (cmd) shell is correct.

**SSH key authentication** — pick the file that matches your build user:

- *Standard (non-admin) user:* put the Linux box's public key in
  `C:\Users\builduser\.ssh\authorized_keys`.
- *Administrator user* (Windows OpenSSH quirk — the per-user file is ignored for admins):
  put it in `C:\ProgramData\ssh\administrators_authorized_keys`, then lock the ACLs:

  ```powershell
  icacls "C:\ProgramData\ssh\administrators_authorized_keys" `
    /inheritance:r /grant "Administrators:F" /grant "SYSTEM:F"
  ```

Then `Restart-Service sshd`.

**Word:**
- Install Word from your VM Word installer.
- Enable *File → Options → Trust Center → Trust Center Settings → Macro Settings →*
  **☑ Trust access to the VBA project object model** (required by the import/export build).
  This corresponds to `AccessVBOM=1` under
  `HKCU\Software\Microsoft\Office\16.0\Word\Security`; the build cannot read the VBA
  project without it.
- The shell `LPandBRL.dotm` you build from **must not have a locked VBA project** — a
  password-locked project reads as 0 components over automation and silently produces an
  empty build (see DEVELOPMENT.md "The VBA project must not be locked").

**Inno Setup:** install Inno Setup 6 (default path `C:\Program Files (x86)\Inno Setup 6`).

**Note the guest IP:** `ipconfig` → record the IPv4 address (e.g. `192.168.1.50`).

---

## Wire up the Linux side

On the **Linux box**, add an SSH alias so port/user are handled cleanly (works for both
bridged and NAT). Edit `~/.ssh/config`:

```
Host vistabuild
    HostName <guest-ip or host-routable-ip>
    User builduser
    Port 22          # or 2222 if you used the NAT port-forward
```

First connection (accept the host key), then point `build.config` at it:

```
# build.config
WIN_HOST = vistabuild
WIN_DIR  = C:/Users/builduser/vistatype-build
WIN_PWSH = powershell
WIN_ISCC = "C:/Program Files (x86)/Inno Setup 6/ISCC.exe"
```

Test the whole path:

```bash
ssh vistabuild          # should log in with no password
make pull               # first canonical export of the VBA/forms from the .dotm
make build              # import src/ -> dist/ .dotm, embed ribbon
```

---

## Snapshots

VirtualBox has its own snapshot mechanism; the `Checkpoint-VM` recipe that stood here was
Hyper-V's and has been removed.

The idea behind it is still worth keeping: **take a snapshot of the clean state before testing
an installer**, so the Quick Access Toolbar merge can be tried repeatedly from a known start.
Take one *after* importing a customized `Word.officeUI`, not before — testing the merge against
a toolbar the transcriber has already changed is the case that has actually caused trouble.

**The VM is disposable by design.** Nothing on it is a source of truth; everything it builds
comes from `src/` on the Linux side.


## Keeping the eval alive

The Windows 11 Enterprise evaluation runs 90 days. To extend a test rig:

```powershell
slmgr /rearm      # resets the timer; usable a few times, then reboot
```

Or just keep the `clean-word` checkpoint and rebuild from the ISO when it lapses — the VM
is disposable by design.
