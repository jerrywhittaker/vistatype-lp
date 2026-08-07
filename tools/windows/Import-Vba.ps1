<#
Import-Vba.ps1  --  Run ON the Windows build box (has Word installed).

Builds the shipping .dotm FROM the text source tree. Strategy: start from a shell
copy of the existing .dotm (so project references, the attached-toolbar ribbon, and
all non-VBA parts are preserved), strip out the old code components, and re-import
the text source. Word regenerates correct p-code on save.

    <SrcRoot>\vba\*.bas   -> standard modules      (removed & re-imported)
    <SrcRoot>\vba\*.cls   -> class OR document mod  (document mods code-replaced in place)
    <SrcRoot>\forms\*.frm -> UserForms (+ .frx)     (removed & re-imported)

Document modules (e.g. ThisDocument) CANNOT be Import-ed; their code module is
cleared and refilled from the .cls body instead.

PREREQUISITE: same "Trust access to the VBA project object model" as Export-Vba.ps1.

Usage (typically invoked over SSH by the Makefile):
    powershell -ExecutionPolicy Bypass -File Import-Vba.ps1 `
        -Shell "C:\build\vistatype\LPandBRL.dotm" `
        -SrcRoot "C:\build\vistatype\src" `
        -OutDotm "C:\build\vistatype\dist\LPandBRL.dotm"
#>
param(
    [Parameter(Mandatory=$true)][string]$Shell,     # existing .dotm used as the base
    [Parameter(Mandatory=$true)][string]$SrcRoot,
    [Parameter(Mandatory=$true)][string]$OutDotm,
    [string]$AppVer  = "",                          # stamp this version into both About dialogs
    [string]$VerDate = "",                          # date shown beside it, e.g. 7/26/2026
    [string]$ProjectName = "LPandBRL"               # VBA project name of the built add-in
)
$ErrorActionPreference = "Stop"

$CT_StdModule = 1; $CT_ClassModule = 2; $CT_MSForm = 3; $CT_Document = 100

# Repo code modules are UTF-8; Word imports .bas/.cls as system ANSI (Windows-1252). Convert
# UTF-8 -> ANSI before handing a module to Word so chars like the smart quotes / bullet /
# en-dash used in the find-and-replace macros survive. Export-Vba.ps1 does the reverse.
$Ansi = [System.Text.Encoding]::GetEncoding(0)          # system ANSI (cp1252 on this box)
$Utf8 = New-Object System.Text.UTF8Encoding($false)     # UTF-8, no BOM

# Work on a copy so the shell .dotm is never mutated in place.
New-Item -ItemType Directory -Force -Path (Split-Path $OutDotm) | Out-Null

# Clear stale Word "owner files" before going anywhere near Word.
#
# Word writes a hidden ~$<name> file beside any document it has open, holding the name of
# whoever has it, and deletes it on a clean close. A Word that was KILLED leaves one behind.
# The next build then opens or saves that same document, Word reads the leftover owner file,
# decides somebody else has it locked, and raises the modal "File In Use" dialog. Headless
# over SSH nobody can answer it, so the build hangs forever - no error, nothing in the event
# log, Word alive and "responding" at near-zero CPU. Worse, the usual recovery (kill the hung
# Word) leaves a FRESH owner file, so every following build hangs the same way.
#
# This cost a full day on 7/26/2026: a ~$andBRL.dotm left over from 7/22 blocked every build
# in between. Note they are HIDDEN - plain "del" skips them and reports "Could Not Find",
# which reads like they were already gone. Remove-Item -Force does remove them.
#
# The files hold nothing but a user name, and this box runs no interactive Word, so there is
# never anything worth keeping.
function Remove-WordOwnerFiles([string]$dir) {
    if (-not $dir -or -not (Test-Path $dir)) { return }
    Get-ChildItem -Path $dir -Force -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name.StartsWith('~$') } |
        ForEach-Object {
            Write-Host "clearing stale Word lock file $($_.Name)"
            Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
        }
}
if (@(Get-Process WINWORD -ErrorAction SilentlyContinue).Count -gt 0) {
    Write-Host "WARNING: WINWORD is already running on this box - almost certainly a leftover from a hung build. It may hold the .dotm open and stall this one."
}
Remove-WordOwnerFiles (Split-Path $OutDotm)
Remove-WordOwnerFiles (Split-Path $Shell)

Copy-Item -Path $Shell -Destination $OutDotm -Force

function Strip-ClsHeader([string]$path) {
    # Return the code body of a .cls/.bas with the VERSION/BEGIN..END block and
    # leading "Attribute VB_*" lines removed (used for Document-module code replace).
    $lines = [System.IO.File]::ReadAllText($path, $Utf8) -split "`r`n|`n"
    $out = New-Object System.Collections.Generic.List[string]
    $inBeginBlock = $false; $started = $false
    foreach ($ln in $lines) {
        if (-not $started) {
            if ($ln -match '^VERSION ')            { continue }
            if ($ln -match '^BEGIN')               { $inBeginBlock = $true; continue }
            if ($inBeginBlock) { if ($ln -match '^END') { $inBeginBlock = $false }; continue }
            if ($ln -match '^Attribute VB_')       { continue }
            $started = $true
        }
        $out.Add($ln)
    }
    return ($out -join "`r`n")
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
# Suppress AutoOpen/AutoNew/etc. so opening the shell doesn't run macros (AutoOpen pops
# MsgBox dialogs) and hang the headless build.
$word.WordBasic.DisableAutoMacros(1)
try {
    $doc  = $word.Documents.Open($OutDotm, $false, $false)
    $proj = $doc.VBProject
    if ($proj -eq $null) { throw "No VBProject. Enable 'Trust access to the VBA project object model'." }

    # Guard against a silent codeless build: if the shell's project enumerates 0 components
    # (most commonly because it is password-locked in the VBE), stop rather than ship an
    # empty .dotm. See DEVELOPMENT.md ("The VBA project must not be locked").
    if ($proj.VBComponents.Count -eq 0) {
        throw "VBProject '$($proj.Name)' has 0 components. A password-locked project also enumerates 0 - unlock it in the VBE (Tools, project Properties, Protection tab). See DEVELOPMENT.md."
    }

    # Names of document-type components (code-replaced, never removed/imported).
    $docComps = @{}
    foreach ($c in $proj.VBComponents) { if ($c.Type -eq $CT_Document) { $docComps[$c.Name] = $true } }

    # 1) Remove existing std/class/form components (leave Document modules alone).
    $toRemove = @()
    foreach ($c in $proj.VBComponents) {
        if ($c.Type -eq $CT_StdModule -or $c.Type -eq $CT_ClassModule -or $c.Type -eq $CT_MSForm) {
            $toRemove += $c
        }
    }
    foreach ($c in $toRemove) { Write-Host "remove $($c.Name)"; $proj.VBComponents.Remove($c) }

    # 2) Import forms (.frm pulls its .frx automatically).
    Get-ChildItem -Path (Join-Path $SrcRoot "forms") -Filter *.frm -ErrorAction SilentlyContinue | ForEach-Object {
        Write-Host "import form $($_.Name)"; $proj.VBComponents.Import($_.FullName) | Out-Null
    }

    # 3) Import/replace vba modules.
    Get-ChildItem -Path (Join-Path $SrcRoot "vba") -Include *.bas,*.cls -File -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $base = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
        if ($docComps.ContainsKey($base)) {
            # Document module: clear its code module and refill from the .cls body.
            Write-Host "code-replace document module $base"
            $cm = $proj.VBComponents.Item($base).CodeModule
            if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
            $body = Strip-ClsHeader $_.FullName
            if ($body.Trim().Length -gt 0) { $cm.AddFromString($body) }
        } else {
            Write-Host "import module $($_.Name)"
            # Import needs an ANSI file (Word reads .bas/.cls in the system codepage); write a
            # temp copy transcoded UTF-8 -> ANSI, keeping the extension so the type is inferred.
            $tmp = Join-Path ([System.IO.Path]::GetTempPath()) $_.Name
            [System.IO.File]::WriteAllText($tmp, [System.IO.File]::ReadAllText($_.FullName, $Utf8), $Ansi)
            try   { $proj.VBComponents.Import($tmp) | Out-Null }
            finally { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
        }
    }

    # Rename the VBA project away from 'Normal' so it does not collide with the user's own
    # Normal.dotm project when this add-in is loaded from the STARTUP folder. VBProject.Name
    # is writable (unlike the project password); no code references the project name.
    if ($ProjectName -and $proj.Name -ne $ProjectName) {
        Write-Host "rename VBA project '$($proj.Name)' -> '$ProjectName'"
        $proj.Name = $ProjectName
    }

    # Stamp the version into both About dialogs.
    #
    # The caption lives in the binary .frx, not in code, so it cannot be edited as text on
    # Linux. We are already inside Word with the project open, so set it here through the
    # form designer rather than starting a second Word session just for this.
    #
    # The two .frx files under src/forms are ALSO one of the four places the version has to
    # agree, so when the caption actually changes we export those forms back over the source
    # tree and the Makefile copies them home. Only when it changes: an unconditional export
    # rewrites the binary every build and churns the repo for nothing.
    if ($AppVer) {
        $stamp = "This computer is running Version $AppVer"
        if ($VerDate) { $stamp += "    $VerDate" }
        foreach ($aboutName in @("Lp_About_Title_And_Agreement", "Dx_About_Title_And_Agreement")) {
            $about = $proj.VBComponents | Where-Object { $_.Name -eq $aboutName }
            if (-not $about) { Write-Host "WARNING: About form $aboutName not found - version not stamped"; continue }
            # The two forms do NOT use the same control name: the LP one is "VersionLabel",
            # the Braille one is "Label5". Match on the caption instead, so this keeps working
            # if either is renamed and does not need a per-form lookup table.
            #
            # EVERY matching control, not just the first. This stopped at the first one until
            # 8/7/2026, and the Braille form carries TWO version captions - so one was stamped
            # and the other kept "This computer is running Version Version 2.1.1  2/23/2026",
            # from February. That stale label is what Jerry kept reading as "the About shows an
            # old version", on a clean install as well as an upgrade, through eight build
            # numbers while the fault was hunted in the installer instead.
            $stampedAny = $false
            $sawCaption = $false
            foreach ($ctl in $about.Designer.Controls) {
                $cap = ""
                try { $cap = [string]$ctl.Caption } catch { }
                if ($cap -notlike "This computer is running Version*") { continue }
                $sawCaption = $true
                if ($cap -eq $stamp) { continue }
                Write-Host "stamp version in $aboutName : '$cap' -> '$stamp'"
                $ctl.Caption = $stamp
                $stampedAny = $true
            }
            if (-not $sawCaption) { Write-Host "WARNING: $aboutName has no version caption - version not stamped"; continue }
            if (-not $stampedAny) { Write-Host "version already $AppVer in $aboutName"; continue }
            $about.Export((Join-Path $SrcRoot "forms\$aboutName.frm"))
            Write-Host "exported $aboutName back to src/forms (Makefile copies it home)"
        }
    }

    $doc.Save()
    $doc.Close($true)
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
Write-Host "Build complete -> $OutDotm"
