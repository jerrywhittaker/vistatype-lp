# Build/test VM setup (Hyper-V on Windows 11 Pro)

A clean, snapshottable Windows + Word VM to run the build/test pipeline, so it never
touches your daily-driver Word. The Linux box that runs `make` drives this VM over SSH.

All commands here run **on the Windows machine** (host or guest as noted) in an
**elevated PowerShell**, unless it says "on the Linux box." They are not run through
Claude — you run them yourself.

Target: **Hyper-V** guest running **Windows 11 Enterprise (Evaluation)** with Word and
Inno Setup installed.

---

## 0. Prerequisites

- **Windows 11 Pro** host with virtualization enabled in UEFI/BIOS (usually on by default).
- **Windows 11 Enterprise Evaluation ISO** — free, 90 days, from the Microsoft Evaluation
  Center (search "Windows 11 Enterprise evaluation"). The eval expires; see
  [Keeping the eval alive](#keeping-the-eval-alive).
- **Your Word installer** (the separate installable Word you have for the VM).
- **Inno Setup 6** installer (jrsoftware.org) — for building the `Setup.exe`.
- The **public SSH key** of the Linux box that will drive the build
  (`cat ~/.ssh/id_ed25519.pub` on the Linux box; create one with `ssh-keygen -t ed25519`
  if needed).
- Free disk for a ~64 GB dynamic VHDX; plan ~6 GB RAM for the guest.

---

## 1. Enable Hyper-V (host)

```powershell
Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V -All
# reboot when prompted
```

If it errors that virtualization is off, enable Intel VT-x / AMD-V in UEFI first.

---

## 2. Create a virtual switch (host)

Pick ONE, based on how the Linux box reaches this machine (yours is *routable but not the
same LAN*):

**A. Bridged / External switch** — the VM gets its own IP on the host's network. Use this
if that network is the one routable from the Linux box.

```powershell
Get-NetAdapter          # note your physical NIC name, e.g. "Ethernet"
New-VMSwitch -Name "ExternalLAN" -NetAdapterName "Ethernet" -AllowManagementOS $true
```

**B. NAT + port-forward** — use this if only the *host's* IP is routable (single routable
address). The VM stays behind the host; you forward a port to it (set up in step 6 once
you know the guest IP). Use Hyper-V's built-in **Default Switch** for the VM in step 3.

---

## 3. Create the VM (host)

Windows 11 requires a Generation-2 VM with Secure Boot **and a virtual TPM**.

```powershell
$vm = "VistaBuild"
New-VM -Name $vm -Generation 2 -MemoryStartupBytes 6GB `
       -NewVHDPath "C:\Hyper-V\$vm\$vm.vhdx" -NewVHDSizeBytes 64GB `
       -SwitchName "ExternalLAN"        # or "Default Switch" for option 2B

Set-VM -Name $vm -ProcessorCount 2 `
       -DynamicMemory -MemoryMinimumBytes 4GB -MemoryMaximumBytes 8GB

# Virtual TPM (required by Windows 11)
Set-VMKeyProtector -VMName $vm -NewLocalKeyProtector
Enable-VMTPM -VMName $vm

# Attach the Windows 11 Enterprise Eval ISO and boot from it first
Add-VMDvdDrive -VMName $vm -Path "C:\ISOs\Windows11_Enterprise_Eval.iso"
Set-VMFirmware -VMName $vm -FirstBootDevice (Get-VMDvdDrive -VMName $vm)

Start-VM -Name $vm
vmconnect.exe localhost $vm
```

Install Windows 11 in the console window. Tips:
- Enterprise Eval lets you create a **local account** (no Microsoft account needed).
- Name the machine something obvious like `VISTABUILD`.
- Create the **build user** you'll SSH in as (e.g., `builduser`).

---

## 4. Configure the guest

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

**Inno Setup:** install Inno Setup 6 (default path `C:\Program Files (x86)\Inno Setup 6`).

**Note the guest IP:** `ipconfig` → record the IPv4 address (e.g. `192.168.1.50`).

---

## 5. (Option 2B only) Port-forward on the host

If you used the NAT/Default Switch, forward a host port to the guest's SSH so the Linux
box can reach it via the host's routable IP:

```powershell
netsh interface portproxy add v4tov4 `
  listenport=2222 listenaddress=0.0.0.0 connectport=22 connectaddress=<guest-ip>
New-NetFirewallRule -DisplayName "SSH to VistaBuild VM" -Direction Inbound `
  -Action Allow -Protocol TCP -LocalPort 2222
```

Now the VM's SSH is reachable at `<host-routable-ip>:2222`.

---

## 6. Wire up the Linux side

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

## 7. Snapshot the clean state

Once Word + SSH + Inno are set up and a build works, checkpoint the VM so installer tests
can revert to a pristine Word:

```powershell
Checkpoint-VM -Name VistaBuild -SnapshotName "clean-word"
```

For installer validation (Issue #1), also make a checkpoint *after* importing a customized
`Word.officeUI`, so you can test the QAT merge against a machine that already has one.
Revert anytime with `Restore-VMCheckpoint -VMName VistaBuild -Name clean-word -Confirm:$false`.

---

## Keeping the eval alive

The Windows 11 Enterprise evaluation runs 90 days. To extend a test rig:

```powershell
slmgr /rearm      # resets the timer; usable a few times, then reboot
```

Or just keep the `clean-word` checkpoint and rebuild from the ISO when it lapses — the VM
is disposable by design.
