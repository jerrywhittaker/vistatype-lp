<#
Move-FormControl.ps1  --  Run ON the Windows build box (has Word installed).

Moves and/or resizes controls ALREADY ON a UserForm and rewrites the .frm/.frx pair in
src/forms/. The fourth of the headless form tools, and it exists for the same reason as the
other three: a control's POSITION and SIZE live in the binary .frx, so they cannot be edited
from Linux, and the .frm alone says nothing about where anything sits.

The set is now:

    New-UserForm.ps1        builds a brand-new form
    Add-FormControl.ps1     puts one control ON an existing form
    Remove-FormControls.ps1 takes named controls OFF
    Set-FormCaption.ps1     changes what a control SAYS
    Move-FormControl.ps1    changes where a control IS and how big     <-- this one

Built in a throwaway blank document, so nothing tracked on the box is touched (same approach
as the other three).

WHY IT WAS WRITTEN, 9/6/2026: consolidating VistaType LP onto ONE progress indicator meant
putting a spinner beside the bar, and then moving the bar to make room for it on the other
side. Add-FormControl could add the spinner but nothing could shift the four controls that
were already there, and rewriting them at run time in UserForm_Initialize would have left the
designer showing a layout no transcriber ever sees.

IT REPORTS THE OLD BOUNDS BEFORE IT CHANGES THEM, always. That printout is the only record of
where a control used to be, and it is what lets a move be undone.

It REFUSES when a name matches nothing, and prints every control on the form with its bounds,
so a near-miss is obvious rather than silent.

    powershell -ExecutionPolicy Bypass -File Move-FormControl.ps1 `
        -Frm "C:\Users\jerry\vistatype-build\src\forms\My_Form.frm" `
        -Name "BarTrack" -Left 58 -Width 348

Every bound is OPTIONAL: pass only what should change. -Left 58 moves it and leaves its size
alone. Sizes are in POINTS, the unit the MSForms designer uses.

-Controls takes several at once, as ONE comma-separated string with each control's bounds after
a colon, in the order Left,Top,Width,Height. An empty slot leaves that bound alone:

    -Controls "BarTrack:58,,348,;BarFill:60,,,;PctLabel:58,,348,"

PowerShell's -File never parses an argument as an expression, so it must be one string and not
an array - the same trap as Remove-FormControls.ps1's -Controls.

TWO TRAPS SHARED WITH THE OTHER TOOLS, both rediscovered the hard way:
  - The Designer does not expose the FORM's own Width and Height to PowerShell. They live on
    the COMPONENT, through Properties.Item, and must be handed over as STRINGS. Use
    -FormWidth / -FormHeight.
  - A control's FONT can only be set from VBA, never from PowerShell. This script does not
    touch fonts at all; set a font in the form's own code.

Afterwards run tools/lib/check_frm_eol.py and tools/lib/trim_frm_blanks.py, as with the others.
#>

param(
    [Parameter(Mandatory=$true)][string]$Frm,
    [string]$Name = "",
    [string]$Controls = "",
    [Nullable[int]]$Left = $null,
    [Nullable[int]]$Top = $null,
    [Nullable[int]]$Width = $null,
    [Nullable[int]]$Height = $null,
    [int]$FormWidth = 0,
    [int]$FormHeight = 0
)

$ErrorActionPreference = "Stop"

if ($Name -eq "" -and $Controls -eq "") {
    throw "Give -Name (with any of -Left -Top -Width -Height) or -Controls. Nothing to do."
}
if ($Name -ne "" -and $Controls -ne "") {
    throw "Give -Name or -Controls, not both."
}

# --- work out what to change, as a list of {name, left, top, width, height} with $null meaning
# --- "leave this bound alone". One shape for both ways of asking, so the loop below is one loop.
$wanted = @()
if ($Name -ne "") {
    $wanted += [pscustomobject]@{ Name=$Name; Left=$Left; Top=$Top; Width=$Width; Height=$Height }
} else {
    foreach ($piece in $Controls.Split(";")) {
        if ($piece.Trim() -eq "") { continue }
        $bits = $piece.Split(":")
        if ($bits.Count -ne 2) { throw "Malformed -Controls entry '$piece'. Expected Name:L,T,W,H" }
        $nums = $bits[1].Split(",")
        if ($nums.Count -ne 4) { throw "Malformed bounds in '$piece'. Expected four slots, L,T,W,H" }
        function Slot($s) { if ($s.Trim() -eq "") { $null } else { [int]$s.Trim() } }
        $wanted += [pscustomobject]@{
            Name   = $bits[0].Trim()
            Left   = (Slot $nums[0]); Top    = (Slot $nums[1])
            Width  = (Slot $nums[2]); Height = (Slot $nums[3])
        }
    }
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

    $designer = $comp.Designer
    Write-Host ("  form before: inside {0} x {1}" -f $designer.InsideWidth, $designer.InsideHeight)

    # Refuse on a name that matches nothing, and SAY WHAT IS THERE. A silent no-op here would be
    # read as "the move did not take" and chased in the .frx, which cannot be read from Linux.
    $have = @{}
    foreach ($c in $designer.Controls) { $have[$c.Name] = $c }

    $missing = @($wanted | Where-Object { -not $have.ContainsKey($_.Name) } | ForEach-Object { $_.Name })
    if ($missing.Count -gt 0) {
        Write-Host "  controls on $($comp.Name):"
        foreach ($c in $designer.Controls) {
            Write-Host ("    {0,-20} L={1,6} T={2,6} W={3,6} H={4,6}" -f $c.Name, $c.Left, $c.Top, $c.Width, $c.Height)
        }
        throw "No control named $($missing -join ', ') on $($comp.Name). Nothing moved."
    }

    foreach ($w in $wanted) {
        $c = $have[$w.Name]
        # The old bounds, printed BEFORE anything changes. This line is the record of where the
        # control used to be, and the only way back if the new position is wrong.
        Write-Host ("  {0,-20} was  L={1,6} T={2,6} W={3,6} H={4,6}" -f $c.Name, $c.Left, $c.Top, $c.Width, $c.Height)

        if ($w.Left   -ne $null) { $c.Left   = $w.Left }
        if ($w.Top    -ne $null) { $c.Top    = $w.Top }
        if ($w.Width  -ne $null) { $c.Width  = $w.Width }
        if ($w.Height -ne $null) { $c.Height = $w.Height }

        Write-Host ("  {0,-20} now  L={1,6} T={2,6} W={3,6} H={4,6}" -f $c.Name, $c.Left, $c.Top, $c.Width, $c.Height)
    }

    # The form's own size lives on the COMPONENT and must be a STRING - see the header.
    if ($FormWidth  -gt 0) { $comp.Properties.Item('Width').Value  = "$FormWidth" }
    if ($FormHeight -gt 0) { $comp.Properties.Item('Height').Value = "$FormHeight" }
    if ($FormWidth -gt 0 -or $FormHeight -gt 0) {
        Write-Host ("  form after : inside {0} x {1}" -f $designer.InsideWidth, $designer.InsideHeight)
    }

    # Export writes BOTH files, over the originals. Word regenerates the .frx from the designer.
    $comp.Export($Frm)
    Write-Host "Exported $Frm and $frx"
    Write-Host "next: check_frm_eol.py and trim_frm_blanks.py"
}
finally {
    if ($doc -ne $null) { $doc.Close(0) }   # 0 = wdDoNotSaveChanges
    $word.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
