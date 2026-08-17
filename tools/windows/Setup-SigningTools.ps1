<#
.SYNOPSIS
    Puts the code-signing tools on the Windows build box. Run once, over SSH.

.DESCRIPTION
    Two things have to be present before a Certum card can sign anything here:

      1. signtool.exe, from the Windows SDK. Only the "Signing Tools for Desktop Apps"
         feature is installed - the whole SDK is several gigabytes and nothing else in
         it is wanted. An x64, an x86 and an arm64 copy all arrive.

         USE THE x64 ONE for a .dotm. Microsoft's own instructions say x86; measured on
         this box 8/17/2026, x86 fails every time with "SignerSign() failed"
         (0x800403f4) and x64 succeeds on the same file. Word here is 64-bit and the
         signing library has to match it. Both sets are registered below anyway, so the
         choice stays open if Office ever changes bitness.

      2. The Office Subject Interface Packages - msosip.dll and msosipx.dll. These teach
         signtool how to sign the VBA project inside a .dotm. They do NOT ship with
         Office (checked on this box, 8/17/2026: not present anywhere under Program
         Files), so they come from Microsoft's download center separately.

    Everything lands under -InstallRoot, which is deliberately OUTSIDE the repo folder:
    src/, tools/ and installer/ are all wiped on this box before each build.

    Downloads are verified with Get-AuthenticodeSignature before anything is run or
    registered, and the script STOPS rather than registering a library whose signature
    does not chain to a trusted root. That matters here: the OfficeSips package was
    reported in 2022 as being signed by "Microsoft Testing Root Certificate Authority
    2010", which is not trusted by anything, and this whole exercise is about not
    shipping unverifiable code.

    Nothing here needs the Certum card. Run it before the card arrives.

.PARAMETER InstallRoot
    Where the tools go. Default C:\vt-signing.

.PARAMETER SkipSdk
    Do not install the Windows SDK signing tools.

.PARAMETER SkipOfficeSips
    Do not fetch the Office SIP libraries.

.PARAMETER Register
    Register msosip.dll / msosipx.dll after fetching them. Needs administrator rights
    (an SSH session on this box already has them). Left off by default so the download
    can be inspected first.

.PARAMETER Unregister
    Undo -Register. Do this BEFORE deleting or moving -InstallRoot: the registration
    points machine-wide at the DLLs where they sit, and a dangling entry breaks
    signature checking on Office files for every process on the box - which would first
    show up at the worst moment, testing the signed add-in in Word.

.PARAMETER Force
    Re-download files already present.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File Setup-SigningTools.ps1
    powershell -ExecutionPolicy Bypass -File Setup-SigningTools.ps1 -Register
    powershell -ExecutionPolicy Bypass -File Setup-SigningTools.ps1 -Unregister
#>
[CmdletBinding()]
param(
    [string] $InstallRoot = 'C:\vt-signing',
    [switch] $SkipSdk,
    [switch] $SkipOfficeSips,
    [switch] $Register,
    [switch] $Unregister,
    [switch] $Force
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# --- Pinned downloads -------------------------------------------------------------
# Windows SDK 10.0.28000.2526 (July 2026), the stable build current on 8/17/2026.
# Newer links live on https://learn.microsoft.com/windows/apps/windows-sdk/downloads
$SdkUrl        = 'https://go.microsoft.com/fwlink/?linkid=2372508'
$SdkVersion    = '10.0.28000.2526'

# Office Subject Interface Packages 16.0.19416.43425 (21 Oct 2025).
# Page: https://www.microsoft.com/en-us/download/details.aspx?id=56617
$SipBase       = 'https://download.microsoft.com/download/c53e473c-3060-4ee9-ac5c-0ddbbeced4e5/'
$SipX64        = 'OfficeSips_x64_16-0-19416-43425.exe'
$SipX86        = 'OfficeSips_x86_16-0-19416-43425.exe'

# Spaces would have to survive Start-Process argument arrays and two self-extractors.
# The default has none; refuse rather than fail obscurely halfway through.
if ($InstallRoot -match '\s') { throw "-InstallRoot must not contain spaces: '$InstallRoot'" }

$dl = Join-Path $InstallRoot 'downloads'
New-Item -ItemType Directory -Path $dl -Force | Out-Null

function Say($msg)  { Write-Host "  $msg" }
function Head($msg) { Write-Host ''; Write-Host "=== $msg ===" }

# Returns $true only for a signature that is Valid - i.e. chains to a root this
# machine trusts. Anything else is reported and treated as a refusal.
function Test-Signature([string] $Path) {
    $sig = Get-AuthenticodeSignature -FilePath $Path
    Say "signature: $($sig.Status)"
    if ($sig.SignerCertificate) { Say "signed by: $($sig.SignerCertificate.Subject)" }
    if ($sig.Status -ne 'Valid') {
        Say "REFUSED - signature is '$($sig.Status)', not 'Valid'."
        return $false
    }
    return $true
}

# Downloads to a .part file and renames only on success, so a dropped transfer can
# never leave a truncated file that a later run mistakes for a good one. A file already
# present is re-checked for a valid signature before it is trusted - a half-download
# fails that check, and the cure is to fetch it again rather than to report a Microsoft
# signing problem that is not happening.
function Get-File([string] $Url, [string] $Path) {
    $leaf = Split-Path $Path -Leaf
    if ((Test-Path $Path) -and -not $Force) {
        Say "already here: $leaf ($([math]::Round((Get-Item $Path).Length/1MB,1)) MB)"
        if (Test-Signature $Path) { return }
        Say 'the copy on disk is not valid - fetching it again'
        Remove-Item $Path -Force
    }
    Say "downloading $leaf ..."
    $part = "$Path.part"
    Remove-Item $part -Force -ErrorAction SilentlyContinue
    Invoke-WebRequest -Uri $Url -OutFile $part -UseBasicParsing -TimeoutSec 600
    Move-Item $part $Path -Force
    Say "got $([math]::Round((Get-Item $Path).Length/1MB,1)) MB"
}

# --- 0. Undo the registration (do this before deleting or moving InstallRoot) ------
if ($Unregister) {
    Head 'unregistering the SIP libraries'
    foreach ($pair in @(@{ Dir = 'officesips-x64'; RegSvr = "$env:WINDIR\System32\regsvr32.exe" },
                        @{ Dir = 'officesips-x86'; RegSvr = "$env:WINDIR\SysWOW64\regsvr32.exe" })) {
        $dir = Join-Path $InstallRoot $pair.Dir
        if (-not (Test-Path $dir)) { Say "skipping $($pair.Dir) - not present"; continue }
        Get-ChildItem $dir -Filter 'msosip*.dll' -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            Say "regsvr32 /u $($_.FullName)"
            $p = Start-Process -FilePath $pair.RegSvr -Wait -PassThru -ArgumentList @('/u', '/s', $_.FullName)
            if ($p.ExitCode -ne 0) { Say "  WARNING: exit code $($p.ExitCode)" }
        }
    }
    Say 'done - the folder can now be deleted safely'
    return
}

# --- 1. Windows SDK signing tools -------------------------------------------------
if (-not $SkipSdk) {
    Head "Windows SDK signing tools ($SdkVersion)"
    $setup = Join-Path $dl 'winsdksetup.exe'
    Get-File $SdkUrl $setup
    if (-not (Test-Signature $setup)) { throw 'winsdksetup.exe failed signature check' }

    Say 'installing the SigningTools feature only, quietly ...'
    $log = Join-Path $InstallRoot 'winsdk-install.log'
    $p = Start-Process -FilePath $setup -Wait -PassThru -ArgumentList @(
        '/features', 'OptionId.SigningTools',
        '/quiet', '/norestart',
        '/ceip', 'off',
        '/log', $log
    )
    Say "installer exit code: $($p.ExitCode)"
    if ($p.ExitCode -ne 0 -and $p.ExitCode -ne 3010) {
        Say "see $log"
        throw "Windows SDK install failed (exit $($p.ExitCode))"
    }
}

Head 'signtool.exe now present?'
$signtools = @()
foreach ($root in @('C:\Program Files (x86)\Windows Kits', 'C:\Program Files\Windows Kits')) {
    if (Test-Path $root) {
        $signtools += Get-ChildItem -Path $root -Filter 'signtool.exe' -Recurse -ErrorAction SilentlyContinue
    }
}
if ($signtools.Count -eq 0) {
    Say 'NONE FOUND'
} else {
    $signtools | ForEach-Object { Say "$($_.FullName)  v$($_.VersionInfo.FileVersion)" }
}

# --- 2. Office Subject Interface Packages -----------------------------------------
if (-not $SkipOfficeSips) {
    Head 'Office Subject Interface Packages (msosip / msosipx)'
    $ok = $true
    foreach ($pair in @(@{ Name = $SipX64; Dir = 'officesips-x64' },
                        @{ Name = $SipX86; Dir = 'officesips-x86' })) {
        $exe = Join-Path $dl $pair.Name
        Get-File ($SipBase + $pair.Name) $exe
        if (-not (Test-Signature $exe)) { $ok = $false; continue }

        $out = Join-Path $InstallRoot $pair.Dir
        New-Item -ItemType Directory -Path $out -Force | Out-Null
        Say "extracting to $out ..."
        # Microsoft self-extractors take /extract:<dir>; /quiet keeps it from prompting.
        $p = Start-Process -FilePath $exe -Wait -PassThru -ArgumentList @('/extract:' + $out, '/quiet')
        Say "extractor exit code: $($p.ExitCode)"
        if ($p.ExitCode -ne 0) { throw "$($pair.Name) failed to extract (exit $($p.ExitCode))" }
        if (-not (Get-ChildItem $out -Filter 'msosip*.dll' -Recurse -ErrorAction SilentlyContinue)) {
            throw "$($pair.Name) extracted but produced no msosip*.dll"
        }
        Get-ChildItem $out -Recurse -ErrorAction SilentlyContinue |
            ForEach-Object { Say "  $($_.FullName.Substring($out.Length + 1))" }

        $readme = Get-ChildItem $out -Filter 'readme*.txt' -Recurse -ErrorAction SilentlyContinue |
                    Select-Object -First 1
        if ($readme) {
            Head "readme from $($pair.Dir)"
            Get-Content $readme.FullName | Select-Object -First 60 | ForEach-Object { Write-Host "  $_" }
        }
    }
    if (-not $ok) {
        Head 'STOPPED'
        Say 'One of the OfficeSips downloads is not validly signed. Nothing has been registered.'
        Say 'This is the 2022 "Microsoft Testing Root CA 2010" problem - decide before proceeding.'
        throw 'OfficeSips download failed its signature check'
    }
}

# --- 3. Register the SIP libraries ------------------------------------------------
if ($Register) {
    Head 'registering the SIP libraries'
    foreach ($pair in @(@{ Dir = 'officesips-x64'; RegSvr = "$env:WINDIR\System32\regsvr32.exe" },
                        @{ Dir = 'officesips-x86'; RegSvr = "$env:WINDIR\SysWOW64\regsvr32.exe" })) {
        $dir = Join-Path $InstallRoot $pair.Dir
        if (-not (Test-Path $dir)) { Say "skipping $($pair.Dir) - not extracted"; continue }
        Get-ChildItem $dir -Filter 'msosip*.dll' -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
            if (-not (Test-Signature $_.FullName)) {
                Say "NOT registered: $($_.Name)"
                return
            }
            Say "regsvr32 $($_.FullName)"
            $p = Start-Process -FilePath $pair.RegSvr -Wait -PassThru -ArgumentList @('/s', $_.FullName)
            Say "  exit code: $($p.ExitCode)"
            if ($p.ExitCode -ne 0) { throw "regsvr32 failed on $($_.Name) (exit $($p.ExitCode)) - is this session elevated?" }
        }
    }
}

# --- 4. Report --------------------------------------------------------------------
Head 'SIP registrations that name msosip'
$found = $false
foreach ($k in @(
    'HKLM:\SOFTWARE\Microsoft\Cryptography\OID\EncodingType 0\CryptSIPDllCreateIndirectData',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Cryptography\OID\EncodingType 0\CryptSIPDllCreateIndirectData'
)) {
    Get-ChildItem -LiteralPath $k -ErrorAction SilentlyContinue | ForEach-Object {
        $dll = (Get-ItemProperty -LiteralPath $_.PSPath -ErrorAction SilentlyContinue).Dll
        if ($dll -like '*msosip*') {
            $found = $true
            $hive = if ($k -like '*WOW6432Node*') { '32-bit' } else { '64-bit' }
            Say "$hive  $($_.PSChildName) -> $dll"
        }
    }
}
if (-not $found) { Say 'none - the VBA project cannot be signed from the command line yet' }

Head 'smart card plumbing (the card is not needed yet)'
$svc = Get-Service SCardSvr -ErrorAction SilentlyContinue
Say "SCardSvr: $($svc.Status) / start=$($svc.StartType)  (Manual is correct - it starts when a reader appears)"
$readers = Get-PnpDevice -Class SmartCardReader -ErrorAction SilentlyContinue
if ($readers) { $readers | ForEach-Object { Say "reader: $($_.FriendlyName) [$($_.Status)]" } }
else { Say 'no reader attached yet' }

# A folder at the root of C: looks like scratch. It is not: HKLM points at the DLLs
# inside it, and deleting them leaves a dangling registration that breaks signature
# checking on Office files for every process on this box.
$marker = Join-Path $InstallRoot 'DO-NOT-DELETE.txt'
@"
VistaType LP code-signing tools.

Windows itself points at the DLLs in this folder, machine-wide, for signing and checking
signatures on Word files. DELETING OR MOVING THIS FOLDER BREAKS THAT for every program on
this computer, and nothing will say so until a signature silently fails to verify.

To remove it properly, run this first:

    powershell -ExecutionPolicy Bypass -File Setup-SigningTools.ps1 -Unregister

then delete the folder. The script lives in the VistaType LP project, under
tools/windows/. Background: docs/Code-Signing.md.
"@ | Set-Content -Path $marker -Encoding UTF8

Head 'done'
Say "tools are under $InstallRoot ($(Split-Path $marker -Leaf) written there)"
Say 'Sign a .dotm with the x64 signtool, never the x86 one - see the header of this file.'
Say 'Before ever deleting that folder, run this script with -Unregister.'
Say 'Next step needs the Certum card: docs/Code-Signing.md, last section.'
