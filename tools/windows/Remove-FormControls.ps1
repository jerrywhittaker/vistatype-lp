<#
Remove-FormControls.ps1  --  Run ON the Windows build box (has Word installed).

Takes controls OFF a UserForm and writes the .frm/.frx pair back out for src/forms/.

Why this exists: a UserForm's layout lives in the BINARY .frx, so a control cannot be deleted
from the Linux side -- deleting its lines from the .frm leaves the control still on the form,
still drawn, and now with no code behind it. Word has to do it. This does it headlessly, in a
throwaway blank document, so nothing tracked is touched (same approach as New-UserForm.ps1).

The alternative is Jerry deleting the control in the VBA editor and exporting the form by hand,
which is how CustomOrientationFrame came off on 8/8/2026. That is still fine for a job he wants
to eyeball. This is for the ones done as part of a code change, where the .frm's code has
already stopped referring to the control and the two must land together.

REMOVE THE CODE FIRST. A control removed while a handler still names it does not fail here and
does not fail the build -- it fails on the transcriber's machine, at the moment the dialog is
opened, as run-time error 424 (this project's forms have no Option Explicit, so a name that is
no longer a control is simply an empty Variant). tools/lib/check_form_calls.py does not catch
it either; it checks macro names, not control names.

A container is removed with everything inside it. Name the children as well anyway - they are
reported before they go, and the list is then a record of exactly what was taken off.

PREREQUISITE (one-time, on the Windows box):
    Word > File > Options > Trust Center > Trust Center Settings > Macro Settings
        [x] Trust access to the VBA project object model

Usage:
    powershell -ExecutionPolicy Bypass -File Remove-FormControls.ps1 `
        -Frm "C:\build\vistatype\src\forms\My_Form.frm" `
        -Controls "Child1,Child2,TheirFrame"

    -Controls is ONE comma-separated string, not a PowerShell array, and it is split here. That
    is forced by -File: PowerShell hands every argument after -File to the script as a literal
    string and never parses it as an expression, so a real [string[]] parameter binds "a,b,c" as
    a single control named "a,b,c" -- which looked exactly like "that control is not on the
    form" and removed nothing. Every other script in this folder is invoked with -File over SSH
    by the Makefile, so this is the shape that has to work.

    The .frx beside the .frm is read and rewritten in place. Afterwards, on the Linux side:
        python3 tools/lib/check_frm_eol.py     (the .frm MUST stay CRLF)
        python3 tools/lib/trim_frm_blanks.py   (Word pads the code section on every export)
#>
param(
    [Parameter(Mandatory=$true)][string]$Frm,
    [Parameter(Mandatory=$true)][string]$Controls
)
$ErrorActionPreference = "Stop"

$CT_MSFORM = 3

$Frm = (Resolve-Path $Frm).Path
$frx = [IO.Path]::ChangeExtension($Frm, ".frx")
if (-not (Test-Path $frx)) { throw "No .frx beside $Frm -- a form is a PAIR, and the layout is in the .frx." }

$wanted = @($Controls -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
if ($wanted.Count -eq 0) { throw "-Controls named nothing." }

$name   = [IO.Path]::GetFileNameWithoutExtension($Frm)
$outDir = Split-Path $Frm -Parent

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
# Same reason as Import-Vba.ps1: an Auto macro popping a MsgBox hangs a headless run.
$word.WordBasic.DisableAutoMacros(1)

$doc = $null
try {
    $doc  = $word.Documents.Add()
    $proj = $doc.VBProject
    if ($proj -eq $null) { throw "No VBProject. Enable 'Trust access to the VBA project object model'." }

    # Import names the component from the .frm's own VB_Name attribute, not from the filename.
    $comp = $proj.VBComponents.Import($Frm)
    Write-Host "Imported $($comp.Name) from $Frm"

    $designer = $comp.Designer
    $removed  = 0

    foreach ($ctl in $wanted) {
        $found = $null
        # Indexing a control that is not there throws, and a container removed earlier in this
        # loop takes its children with it -- so look before reaching, and say so either way.
        foreach ($c in $designer.Controls) { if ($c.Name -eq $ctl) { $found = $c; break } }
        if ($found -eq $null) {
            Write-Host ("  {0,-24} not present (already gone, or removed with its container)" -f $ctl)
            continue
        }
        # Reported before removal because it is the only record of where the hole is: whoever
        # tidies the form up afterwards needs to know what size gap was left and where.
        Write-Host ("  {0,-24} removing -- Left {1}, Top {2}, Width {3}, Height {4}" -f `
                    $found.Name, $found.Left, $found.Top, $found.Width, $found.Height)
        $designer.Controls.Remove($ctl)
        $removed++
    }

    if ($removed -eq 0) { throw "Nothing was removed. The form is unchanged; not rewriting it." }

    # Export writes BOTH files, over the originals. Word regenerates the .frx from the designer.
    $comp.Export($Frm)
    Write-Host "Exported $Frm and $frx ($removed control(s) removed)."
}
finally {
    if ($doc -ne $null) { $doc.Close(0) }   # 0 = wdDoNotSaveChanges
    $word.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
