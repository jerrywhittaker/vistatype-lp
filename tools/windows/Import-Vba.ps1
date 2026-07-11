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
        -Shell "C:\build\vistatype\Normal.dotm" `
        -SrcRoot "C:\build\vistatype\src" `
        -OutDotm "C:\build\vistatype\dist\Normal.dotm"
#>
param(
    [Parameter(Mandatory=$true)][string]$Shell,     # existing .dotm used as the base
    [Parameter(Mandatory=$true)][string]$SrcRoot,
    [Parameter(Mandatory=$true)][string]$OutDotm
)
$ErrorActionPreference = "Stop"

$CT_StdModule = 1; $CT_ClassModule = 2; $CT_MSForm = 3; $CT_Document = 100

# Work on a copy so the shell .dotm is never mutated in place.
New-Item -ItemType Directory -Force -Path (Split-Path $OutDotm) | Out-Null
Copy-Item -Path $Shell -Destination $OutDotm -Force

function Strip-ClsHeader([string]$path) {
    # Return the code body of a .cls/.bas with the VERSION/BEGIN..END block and
    # leading "Attribute VB_*" lines removed (used for Document-module code replace).
    $lines = Get-Content -LiteralPath $path
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
try {
    $doc  = $word.Documents.Open($OutDotm, $false, $false)
    $proj = $doc.VBProject
    if ($proj -eq $null) { throw "No VBProject. Enable 'Trust access to the VBA project object model'." }

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
            $proj.VBComponents.Import($_.FullName) | Out-Null
        }
    }

    $doc.Save()
    $doc.Close($true)
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
Write-Host "Build complete -> $OutDotm"
