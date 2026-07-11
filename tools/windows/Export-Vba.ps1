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
        -Dotm "C:\build\vistatype\Normal.dotm" -SrcRoot "C:\build\vistatype\src"
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
try {
    $doc = $word.Documents.Open($Dotm, $false, $false)   # ConfirmConversions=false, ReadOnly=false
    $proj = $doc.VBProject
    if ($proj -eq $null) { throw "No VBProject. Enable 'Trust access to the VBA project object model'." }

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
    $doc.Close($false)   # SaveChanges=false; we only read
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
Write-Host "Export complete -> $SrcRoot"
