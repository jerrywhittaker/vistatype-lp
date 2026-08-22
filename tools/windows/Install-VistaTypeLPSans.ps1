<#
.SYNOPSIS
    Install (or update) VistaTypeLP Sans for the current user. No admin required.

.DESCRIPTION
    Per-user font install: copies to %LOCALAPPDATA%\Microsoft\Windows\Fonts and registers
    under HKCU, then broadcasts WM_FONTCHANGE so running apps pick it up.

    Two things this handles that a naive copy does not:

      * The Windows Font Cache Service keeps a lock on an installed .ttf even after
        RemoveFontResourceW and even with Word closed. Overwriting in place therefore
        fails with "being used by another process". We install under a timestamped
        filename and repoint the registry value, then delete the old file if we can.

      * A running WINWORD is not necessarily ours. A process in session 0 with
        /Automation or -Embedding is a leftover test instance and is safe to stop.
        Anything in another session is the user's own Word with possibly unsaved work,
        and this script will refuse rather than touch it.

.PARAMETER SourceDir
    Directory holding VistaTypeLPSans-Regular.ttf and VistaTypeLPSans-Bold.ttf.

.PARAMETER Uninstall
    Remove VistaTypeLP Sans instead of installing it.

.EXAMPLE
    .\install-vistatypelp-sans.ps1 -SourceDir .\dist\fonts
    .\install-vistatypelp-sans.ps1 -Uninstall
#>
[CmdletBinding()]
param(
    [string] $SourceDir = ".\dist\fonts",
    [switch] $Uninstall
)

$ErrorActionPreference = 'Stop'

$FontDir = Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts'
$RegKey  = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts'

# registry value name -> source filename
$Faces = [ordered]@{
    'VistaTypeLP Sans (TrueType)'      = 'VistaTypeLPSans-Regular.ttf'
    'VistaTypeLP Sans Bold (TrueType)' = 'VistaTypeLPSans-Bold.ttf'
}

Add-Type -Name VtlpGdi -Namespace Vtlp -MemberDefinition @'
[DllImport("gdi32.dll", CharSet=CharSet.Unicode)] public static extern int  AddFontResourceW(string p);
[DllImport("gdi32.dll", CharSet=CharSet.Unicode)] public static extern bool RemoveFontResourceW(string p);
[DllImport("user32.dll")] public static extern int SendMessageTimeout(
    IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam, uint flags, uint timeout, out IntPtr result);
'@

function Assert-NoForeignWord {
    $procs = Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" |
             Select-Object ProcessId, SessionId, CommandLine
    if (-not $procs) { return }

    foreach ($p in $procs) {
        $isOurs = ($p.SessionId -eq 0) -and ($p.CommandLine -match '/Automation|-Embedding')
        if ($isOurs) {
            Write-Host "  stopping leftover automation Word (pid $($p.ProcessId))"
            Stop-Process -Id $p.ProcessId -Force
        }
        else {
            throw ("Word is running in session $($p.SessionId) (pid $($p.ProcessId)). " +
                   "That is a real user session, not an automation instance - close Word " +
                   "and re-run. Refusing to touch it.")
        }
    }
    Get-ChildItem (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\~$*') -ErrorAction SilentlyContinue |
        ForEach-Object { Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue }
}

function Publish-FontChange {
    $HWND_BROADCAST = [IntPtr] 0xffff
    $WM_FONTCHANGE  = 0x001D
    $res = [IntPtr]::Zero
    [void][Vtlp.VtlpGdi]::SendMessageTimeout(
        $HWND_BROADCAST, $WM_FONTCHANGE, [IntPtr]::Zero, [IntPtr]::Zero, 2, 1000, [ref] $res)
}

function Remove-Face([string] $ValueName) {
    $current = (Get-ItemProperty -Path $RegKey -Name $ValueName -ErrorAction SilentlyContinue).$ValueName
    if (-not $current) { return $null }
    [void][Vtlp.VtlpGdi]::RemoveFontResourceW($current)
    Remove-ItemProperty -Path $RegKey -Name $ValueName -ErrorAction SilentlyContinue
    return $current
}

Assert-NoForeignWord

if ($Uninstall) {
    Write-Host "Uninstalling $($Faces.Count) face(s)..."
    foreach ($name in $Faces.Keys) {
        $old = Remove-Face $name
        if (-not $old) { Write-Host "  not installed: $name"; continue }
        if (Test-Path $old) {
            try   { Remove-Item $old -Force -ErrorAction Stop; Write-Host "  removed $name" }
            catch { Write-Host "  unregistered $name (file held by font cache, clears on reboot)" }
        }
        else { Write-Host "  removed $name" }
    }
    Publish-FontChange
    Write-Host "Done."
    return
}

if (-not (Test-Path $FontDir)) { New-Item -ItemType Directory -Path $FontDir -Force | Out-Null }
if (-not (Test-Path $RegKey))  { New-Item -Path $RegKey -Force | Out-Null }

$stamp = Get-Date -Format 'yyyyMMddHHmm'
Write-Host "Installing from $SourceDir"

foreach ($name in $Faces.Keys) {
    $src = Join-Path $SourceDir $Faces[$name]
    if (-not (Test-Path $src)) { throw "missing source font: $src  (run build-vistatypelp-sans.py first)" }

    $old = Remove-Face $name

    # Timestamped target: the font cache may still hold the previous file.
    $target = Join-Path $FontDir ($Faces[$name] -replace '\.ttf$', "-$stamp.ttf")
    Copy-Item $src $target -Force
    New-ItemProperty -Path $RegKey -Name $name -Value $target -PropertyType String -Force | Out-Null

    $added = [Vtlp.VtlpGdi]::AddFontResourceW($target)
    if ($added -lt 1) { throw "AddFontResourceW reported 0 faces for $target" }
    Write-Host ("  {0,-34} -> {1}" -f $name, (Split-Path $target -Leaf))

    if ($old -and (Test-Path $old) -and ($old -ne $target)) {
        try   { Remove-Item $old -Force -ErrorAction Stop }
        catch { Write-Host "     (previous file still locked by font cache; clears on reboot)" }
    }
}

Publish-FontChange

# Verify Windows actually resolves the family - a copied file is not an installed font.
Add-Type -AssemblyName System.Drawing
$fam = (New-Object System.Drawing.Text.InstalledFontCollection).Families |
       Where-Object { $_.Name -eq 'VistaTypeLP Sans' }
if (-not $fam) { throw "installed, but Windows does not report the family - install FAILED" }

$em = $fam.GetEmHeight('Regular')
$ls = $fam.GetLineSpacing('Regular')
$styles = @('Regular','Bold') | Where-Object { $fam.IsStyleAvailable($_) }
Write-Host ""
Write-Host ("Verified: {0}  em={1}  lineSpacing={2}  ratio={3:N4}  styles={4}" -f
            $fam.Name, $em, $ls, ($ls / $em), ($styles -join ','))
Write-Host "Word must be restarted to see a newly installed font."
