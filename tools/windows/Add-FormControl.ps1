<#
Add-FormControl.ps1  --  Run ON the Windows build box (has Word installed).

Puts ONE control ONTO an existing UserForm and writes the .frm/.frx pair back out for
src/forms/. The counterpart to Remove-FormControls.ps1, and it exists for the same reason: a
UserForm's layout lives in the BINARY .frx, which cannot be written from the Linux side, so
Word has to do it. Done headlessly in a throwaway blank document, so nothing tracked is
touched (same approach as New-UserForm.ps1 and Remove-FormControls.ps1).

The alternative is Jerry adding the control in the VBA editor and exporting the form by hand,
which is still the right answer for anything he wants to lay out by eye. This is for the ones
that come with a code change, where the .frm's code already refers to the control and the two
have to land together.

ADD THE CONTROL FIRST, THEN THE CODE THAT NAMES IT - the opposite order from removal. Code
naming a control that does not exist reaches the transcriber as run-time error 424 when the
dialog opens (these forms have no Option Explicit, so an unknown name is an empty Variant),
and neither this pipeline nor tools/lib/check_form_calls.py will say a word about it.

Jerry's form conventions are applied to whatever is added: Tahoma 10, a ControlTipText on a
CommandButton, and the O/C accelerators on Okay and Cancel. See New-UserForm.ps1, which does
the same for a form built from scratch.

PREREQUISITE (one-time, on the Windows box):
    Word > File > Options > Trust Center > Trust Center Settings > Macro Settings
        [x] Trust access to the VBA project object model

Usage:
    powershell -ExecutionPolicy Bypass -File Add-FormControl.ps1 `
        -Frm "C:\build\vistatype\src\forms\My_Form.frm" `
        -Type Label -Name KeyHelp -Left 6 -Top 36 -Width 470 -Height 16 `
        -FormHeight 87

    Sizes are in POINTS, the unit the MSForms designer uses - the same as New-UserForm.ps1.
    -FormHeight / -FormWidth are the FORM's outside size and include its title bar and border
    (about 29 points vertically, 12 horizontally, measured on the build box 8/23/2026); leave
    them off to grow the form by hand in code instead.

    A Label's text is better set in UserForm_Initialize than passed here, so that the wording
    stays reviewable as text rather than disappearing into the binary .frx. -Caption is for
    the cases where there is no code to set it.

    Afterwards, on the Linux side:
        python3 tools/lib/check_frm_eol.py     (the .frm MUST stay CRLF)
        python3 tools/lib/trim_frm_blanks.py   (Word pads the code section on every export)
#>
param(
    [Parameter(Mandatory=$true)][string]$Frm,
    [Parameter(Mandatory=$true)][string]$Name,
    [ValidateSet("Label","CommandButton","TextBox","CheckBox","OptionButton","Frame")]
    [string]$Type = "Label",
    [Parameter(Mandatory=$true)][int]$Left,
    [Parameter(Mandatory=$true)][int]$Top,
    [Parameter(Mandatory=$true)][int]$Width,
    [Parameter(Mandatory=$true)][int]$Height,
    [string]$Caption = "",
    [int]$FormWidth = 0,
    [int]$FormHeight = 0
)
$ErrorActionPreference = "Stop"

$CT_MSFORM = 3

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

    $designer = $comp.Designer

    # Refuse rather than end up with two controls of nearly the same name: MSForms would accept
    # the Add and then rename the new one (KeyHelp1), which compiles, draws, and is wired to
    # nothing.
    foreach ($c in $designer.Controls) {
        if ($c.Name -eq $Name) { throw "$($comp.Name) already has a control named $Name. Nothing added." }
    }

    # InsideWidth/InsideHeight, NOT Width/Height: the Designer object does not expose the form's
    # outside size to PowerShell at all - reading it gives an empty string and setting it throws
    # "The property 'Height' cannot be found on this object". The form's own size lives on the
    # COMPONENT, through Properties.Item, and has to be handed over as a STRING. Both traps are
    # New-UserForm.ps1's, rediscovered here on 8/23/2026 because this script did not copy it.
    Write-Host ("  form before: inside {0} x {1}" -f $designer.InsideWidth, $designer.InsideHeight)

    $ctl = $designer.Controls.Add("Forms.$Type.1", $Name, $true)
    $ctl.Left   = $Left
    $ctl.Top    = $Top
    $ctl.Width  = $Width
    $ctl.Height = $Height
    if ($Caption -ne "") { $ctl.Caption = $Caption }

    Write-Host ("  {0,-24} added -- {1}, Left {2}, Top {3}, Width {4}, Height {5}" -f `
                $Name, $Type, $Left, $Top, $Width, $Height)

    if ($FormWidth  -gt 0) { $comp.Properties.Item('Width').Value  = "$FormWidth" }
    if ($FormHeight -gt 0) { $comp.Properties.Item('Height').Value = "$FormHeight" }
    if ($FormWidth -gt 0 -or $FormHeight -gt 0) {
        Write-Host ("  form after : inside {0} x {1}" -f $designer.InsideWidth, $designer.InsideHeight)
    }

    # --- Jerry's form conventions, applied to the CONTROL just added and to nothing else.
    #
    # In VBA rather than from PowerShell, and for the same reason New-UserForm.ps1 does it that
    # way: "$ctl.Font.Size = 10" from PowerShell reports success and reads back empty, and
    # $comp.Properties.Item('Font') throws "Invalid object use". From VBA the same properties
    # behave. The existing controls are deliberately left alone - a form built before the
    # conventions is not retrofitted by adding one label to it.
    $tidy = $proj.VBComponents.Add(1)
    $tidy.CodeModule.AddFromString(@"
Sub Zz_Apply_Form_Conventions()
    Dim c As Object, ctl As Object
    Set c = ThisDocument.VBProject.VBComponents("$($comp.Name)")
    On Error Resume Next
    Set ctl = c.Designer.Controls("$Name")

    ctl.Font.Name = "Tahoma"
    ctl.Font.Size = 10

    If TypeName(ctl) = "CommandButton" Then
        ctl.ControlTipText = ctl.Caption & " Button"
        If ctl.Caption = "Okay" Then ctl.Accelerator = "O"
        If ctl.Caption = "Cancel" Then ctl.Accelerator = "C"
    End If
End Sub
"@)
    $doc.Activate()
    $word.Run('Zz_Apply_Form_Conventions') | Out-Null
    $proj.VBComponents.Remove($tidy)
    Write-Host "  applied Tahoma 10 (and the button rules, if it is a button)"

    # Export writes BOTH files, over the originals. Word regenerates the .frx from the designer.
    $comp.Export($Frm)
    Write-Host "Exported $Frm and $frx"
    Write-Host "next: the code that names $Name, then check_frm_eol.py and trim_frm_blanks.py"
}
finally {
    if ($doc -ne $null) { $doc.Close(0) }   # 0 = wdDoNotSaveChanges
    $word.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
