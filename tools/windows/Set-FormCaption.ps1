<#
.SYNOPSIS
    Change the Caption of ONE control on an existing UserForm, and rewrite the .frm/.frx pair.
    With -NewTip, change its hover text (ControlTipText) too, or instead.

.DESCRIPTION
    The counterpart to Add-FormControl.ps1 and Remove-FormControls.ps1, for the case where the
    control is already there and only its WORDS are wrong.

    A caption set in the designer lives in the BINARY .frx, so it cannot be edited from Linux -
    the same reason the two About dialogs' version captions are stamped over SSH by Import-Vba.ps1
    rather than written into src/. This script is that technique made general.

    Matches the control by its CURRENT CAPTION rather than by name, which is what Import-Vba.ps1
    settled on for the About forms: the two of them use different control names for the same
    label, and matching on the words means no per-form lookup table to keep in step. Pass -Name
    instead when the caption is not unique, or is empty.

    Refuses when it matches nothing, and refuses when it matches more than one control unless
    -All is given. A caption change that silently hits nothing looks exactly like a change that
    worked, right up until the transcriber reads the old words.

    -NewTip sets the ControlTipText of the same matched control. The hover text lives in the
    .frx as well, and a misspelling there is just as visible as one in the caption. Give
    -NewCaption, -NewTip, or both; a parameter left out leaves that property alone.

    Runs in a throwaway blank document, so nothing tracked is touched but the .frm/.frx pair.

    PREREQUISITE: "Trust access to the VBA project object model", same as Export-Vba.ps1.

.EXAMPLE
    powershell -File tools\windows\Set-FormCaption.ps1 `
        -Frm src\forms\Dx_Type_Dashes_Form.frm `
        -OldCaption "Select all text containing text fractions (not dates like 12/25)" `
        -NewCaption "Select the range containing fractions (not dates like 12/25/2026 or 12/26)"

.EXAMPLE
    powershell -File tools\windows\Set-FormCaption.ps1 -Frm src\forms\Some_Form.frm `
        -Name Label7 -NewCaption "New words"

.EXAMPLE
    powershell -File tools\windows\Set-FormCaption.ps1 -Frm src\forms\Lp_Table_Convert_Options_Form.frm `
        -Name ColumnOnlyTable -NewTip "Table has column heading only radio button"

.NOTES
    Afterwards, on the Linux side:
        python3 tools/lib/check_frm_eol.py     (the .frm MUST stay CRLF)
        python3 tools/lib/trim_frm_blanks.py   (Word pads the code section on every export)
#>
param(
    [Parameter(Mandatory=$true)][string]$Frm,
    [string]$OldCaption = "",
    [string]$Name = "",
    [string]$NewCaption,
    [string]$NewTip,
    [switch]$All
)
$ErrorActionPreference = "Stop"

if ($OldCaption -eq "" -and $Name -eq "") {
    throw "Give -OldCaption or -Name, so there is something to match on."
}
# Asked by name, not by value: an empty caption is a legitimate thing to set.
$setCaption = $PSBoundParameters.ContainsKey('NewCaption')
$setTip     = $PSBoundParameters.ContainsKey('NewTip')
if (-not $setCaption -and -not $setTip) {
    throw "Give -NewCaption, -NewTip, or both, so there is something to change."
}

$Frm = (Resolve-Path $Frm).Path
$frx = [IO.Path]::ChangeExtension($Frm, ".frx")
if (-not (Test-Path $frx)) { throw "No .frx beside $Frm -- a form is a PAIR, and the layout is in the .frx." }

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

    $hits = @()
    foreach ($c in $comp.Designer.Controls) {
        # Not every control HAS a Caption - a TextBox does not - so ask and move on.
        $cap = ""
        try { $cap = [string]$c.Caption } catch { continue }

        if ($Name -ne "") {
            if ($c.Name -eq $Name) { $hits += $c }
        } elseif ($cap -eq $OldCaption) {
            $hits += $c
        }
    }

    if ($hits.Count -eq 0) {
        # Say what IS there. Hunting a caption that differs by one space is otherwise guesswork.
        Write-Host "Captions on $($comp.Name):"
        foreach ($c in $comp.Designer.Controls) {
            $cap = ""
            try { $cap = [string]$c.Caption } catch { continue }
            Write-Host ("  {0,-28} [{1}]" -f $c.Name, $cap)
        }
        throw "Nothing matched. Nothing changed."
    }

    if ($hits.Count -gt 1 -and -not $All) {
        foreach ($c in $hits) { Write-Host ("  matched {0}" -f $c.Name) }
        throw "$($hits.Count) controls matched. Pass -All if every one of them should be changed."
    }

    foreach ($c in $hits) {
        if (-not $setCaption) {
            # Still say what the caption is: it is the words a transcriber sees on the control.
            Write-Host ("{0}.Caption is '{1}' (left alone)" -f $c.Name, [string]$c.Caption)
        }
        if ($setCaption) {
            Write-Host ("set {0}.Caption : '{1}' -> '{2}'" -f $c.Name, [string]$c.Caption, $NewCaption)
            $c.Caption = $NewCaption
        }
        if ($setTip) {
            Write-Host ("set {0}.ControlTipText : '{1}' -> '{2}'" -f $c.Name, [string]$c.ControlTipText, $NewTip)
            $c.ControlTipText = $NewTip
        }
    }

    # Export writes BOTH files, over the originals. Word regenerates the .frx from the designer.
    $comp.Export($Frm)
    Write-Host "Exported $Frm and $frx"
}
finally {
    if ($doc -ne $null) { $doc.Close($false) }
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
