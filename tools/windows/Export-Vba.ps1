<#
Export-Vba.ps1  --  Run ON the Windows build box (has Word installed).

Opens the .dotm and exports every VBA component to text so the Linux repo has a
canonical, import-ready source tree:
    *.bas  standard modules      -> <SrcRoot>\vba\
    *.cls  class modules         -> <SrcRoot>\vba\
    *.cls  document modules      -> <SrcRoot>\vba\   (ThisDocument; import handled specially)
    *.frm  UserForms (+ .frx)    -> <SrcRoot>\forms\

This is the authoritative extractor. The Linux-side Python decompressor is for
reading only; forms in particular need this step to get valid .frm/.frx.

PREREQUISITE (one-time, on the Windows box):
    Word > File > Options > Trust Center > Trust Center Settings > Macro Settings
        [x] Trust access to the VBA project object model

Usage (typically invoked over SSH by the Makefile):
    powershell -ExecutionPolicy Bypass -File Export-Vba.ps1 `
        -Dotm "C:\build\vistatype\LPandBRL.dotm" -SrcRoot "C:\build\vistatype\src"
#>
param(
    [Parameter(Mandatory=$true)][string]$Dotm,
    [Parameter(Mandatory=$true)][string]$SrcRoot
)
$ErrorActionPreference = "Stop"

$vbaDir  = Join-Path $SrcRoot "vba"
$formDir = Join-Path $SrcRoot "forms"
New-Item -ItemType Directory -Force -Path $vbaDir, $formDir | Out-Null

# VBA component type constants (vbext_ComponentType)
$CT_StdModule = 1; $CT_ClassModule = 2; $CT_MSForm = 3; $CT_Document = 100

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0   # wdAlertsNone
# Suppress AutoOpen/AutoNew/etc. -- opening this .dotm as a document would otherwise
# run AutoOpen (which pops MsgBox dialogs) and hang the headless build.
$word.WordBasic.DisableAutoMacros(1)
try {
    $doc = $word.Documents.Open($Dotm, $false, $false)   # ConfirmConversions=false, ReadOnly=false
    $proj = $doc.VBProject
    if ($proj -eq $null) { throw "No VBProject. Enable 'Trust access to the VBA project object model'." }

    # A password-locked VBA project ('Lock project for viewing') enumerates 0 components
    # through the object model, with no API to unlock it -- that would make this export a
    # silent no-op. Fail loudly instead so the lock gets fixed in the VBE.
    $count = $proj.VBComponents.Count
    if ($count -eq 0) {
        throw "VBProject '$($proj.Name)' has 0 components. A password-locked project also enumerates 0 - unlock it in the VBE (Tools, project Properties, Protection tab). See DEVELOPMENT.md."
    }

    # Clear prior exports so a renamed/removed component can't leave a stale file behind.
    # Done only after the project is confirmed non-empty, so a failed open never wipes src.
    Remove-Item -Path (Join-Path $vbaDir  "*.bas"), (Join-Path $vbaDir  "*.cls"),
                      (Join-Path $formDir "*.frm"), (Join-Path $formDir "*.frx") -Force -ErrorAction SilentlyContinue

    foreach ($comp in $proj.VBComponents) {
        $name = $comp.Name
        switch ($comp.Type) {
            $CT_StdModule   { $comp.Export((Join-Path $vbaDir  "$name.bas")) }
            $CT_ClassModule { $comp.Export((Join-Path $vbaDir  "$name.cls")) }
            $CT_Document    { $comp.Export((Join-Path $vbaDir  "$name.cls")) }  # ThisDocument
            $CT_MSForm      { $comp.Export((Join-Path $formDir "$name.frm")) }  # writes .frm + .frx
            default         { Write-Warning "Skipping $name (type $($comp.Type))" }
        }
        Write-Host "exported $name"
    }

    # Word exports code modules in the system ANSI codepage (Windows-1252). Re-encode the
    # code modules to UTF-8 so the repo source is UTF-8-clean and git-friendly. Import-Vba.ps1
    # converts them back to ANSI before handing them to Word. Forms (.frm) are ASCII and the
    # .frx is binary, so both are left byte-for-byte as Word wrote them.
    $ansi = [System.Text.Encoding]::GetEncoding(0)          # system ANSI (cp1252 on this box)
    $utf8 = New-Object System.Text.UTF8Encoding($false)     # UTF-8, no BOM
    Get-ChildItem -Path $vbaDir -Include *.bas,*.cls -File -Recurse | ForEach-Object {
        $text = [System.IO.File]::ReadAllText($_.FullName, $ansi)
        [System.IO.File]::WriteAllText($_.FullName, $text, $utf8)
    }

    $doc.Close($false)   # SaveChanges=false; we only read
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
Write-Host "Export complete -> $SrcRoot"
