<#
New-UserForm.ps1  --  Run ON the Windows build box (has Word installed).

Creates a brand-new UserForm and exports it as a .frm/.frx pair for src/forms/.

Why this exists: a UserForm is a text .frm plus a BINARY .frx holding the control layout.
The .frx cannot be written by hand from the Linux side, and `make pull` must never be used
to seed one -- it wipes src/vba and src/forms and re-exports everything from the .dotm at
the repo root, which is usually older than src/. So Word itself builds the form here, in a
throwaway blank document, and nothing tracked is touched.

This writes the form's LAYOUT only. Add its code afterwards by appending to the exported
.frm as text (with CRLF), then run `python3 tools/lib/check_frm_eol.py`. Long message text
belongs in UserForm_Initialize rather than in a static label, so it stays reviewable as text
instead of disappearing into the binary .frx.

Sizes are in POINTS (72 to the inch), the unit the MSForms designer uses. The .frm header
records the form in twips (20 to the point); Word does that conversion on export.

PREREQUISITE (one-time, on the Windows box):
    Word > File > Options > Trust Center > Trust Center Settings > Macro Settings
        [x] Trust access to the VBA project object model

Usage:
    powershell -ExecutionPolicy Bypass -File New-UserForm.ps1 `
        -Name Sh_My_Form -Caption "My Dialog" -OutDir "C:\build\vistatype\formout"

    Forms are created in Tahoma 10, and every CommandButton gets a ControlTipText of its
    caption plus " Button" - both Jerry's rules from 8/7/2026. An Okay button is captioned
    "Okay", never "OK", and Okay and Cancel carry accelerators of O and C - Jerry's rule from
    8/23/2026. Keep all of it on anything added to the .frm by hand afterwards.

    ... -InfoDialog        also lays in the read-only note shape: a locked, border-less
                           text box named Info_Text filling the form, and an Okay button named
                           Close_Me. That is the shape Sh_Prodnote_Info_Form uses -- the text
                           wraps, can be selected and copied, and grows a scroll bar instead
                           of clipping when Windows runs at a large display scale.

    ... -MessageDialog     lays in the SHARED MESSAGE shape: the same read-only text box,
                           named Msg_Text, with an Okay button and a Cancel button beside it.
                           This is what Sh_Message_Form is built from, and it is the shape
                           every message in VistaType LP is meant to take from 8/23/2026 -
                           a MsgBox can be none of 10 point Tahoma, "Okay", or an accelerator.
                           The caller hides the Cancel button when it has nothing to cancel.
#>
param(
    [Parameter(Mandatory=$true)][string]$Name,
    [Parameter(Mandatory=$true)][string]$OutDir,
    [string]$Caption = "",
    # Integers, not doubles: the VBE Properties setter rejects a System.Double outright
    # ("Unable to cast object of type 'System.Double' to type 'System.String'").
    [int]$Width      = 492,
    [int]$Height     = 400,
    [switch]$InfoDialog,
    [switch]$MessageDialog
)
$ErrorActionPreference = "Stop"

# VBA component type constant (vbext_ComponentType)
$CT_MSFORM = 3

if ($Caption -eq "") { $Caption = $Name }

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$target = Join-Path $OutDir "$Name.frm"
if (Test-Path $target) {
    throw "$target already exists. Refusing to overwrite a form -- move it aside first."
}

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

    $comp = $proj.VBComponents.Add($CT_MSFORM)
    $comp.Name = $Name

    # NOTE: $comp.Properties('Caption') throws "Could not find member" from PowerShell. The
    # parenthesised default-member form works in VBA but is not exposed here -- it has to be
    # .Item(name).Value.
    # ... and the numeric ones have to be handed over as STRINGS. A literal (.Value = 492)
    # happens to work because PowerShell converts a constant on the way in; a variable of the
    # same value does not, and fails with "Unable to cast object of type 'System.Int32' to
    # type 'System.String'".
    $comp.Properties.Item('Caption').Value = $Caption
    $comp.Properties.Item('Width').Value   = "$Width"
    $comp.Properties.Item('Height').Value  = "$Height"

    $d = $comp.Designer
    if ($d -eq $null) { throw "Designer is null for $Name" }

    if ($InfoDialog) {
        $pad       = 12
        $btnW      = 72
        $btnH      = 24
        # What the form's border and title bar actually take, MEASURED on the build box 8/23/2026
        # from an exported .frm: a form 480 x 300 reports a client area of 468 x 270.75. These
        # read 36 and 46 until then, which left every control 24 points short of the right edge
        # and 16 points short of the bottom - a dialog that looks shoved into its top left corner.
        $insideW   = $Width  - 12
        $insideH   = $Height - 30
        $textH     = $insideH - $btnH - ($pad * 3)
        # Precomputed, not written inline in the array below: PowerShell's comma binds TIGHTER
        # than arithmetic, so @('Width', $insideW - $x) parses as ('Width', $insideW) - $x and
        # dies with "[System.Object[]] does not contain a method named 'op_Subtraction'".
        $textW     = $insideW - ($pad * 2)

        $tb = $d.Controls.Add('Forms.TextBox.1', 'Info_Text', $true)
        # NOTE: setting $tb.Font.Size here throws a null reference. Leave the font alone --
        # the default Tahoma is what the rest of the forms use anyway.
        foreach ($pair in @(
            @('Left', $pad), @('Top', $pad),
            @('Width', $textW), @('Height', $textH),
            @('MultiLine', $true), @('WordWrap', $true),
            @('ScrollBars', 2),                # fmScrollBarsVertical
            @('Locked', $true), @('TabStop', $false),
            @('SpecialEffect', 0),             # fmSpecialEffectFlat
            @('BorderStyle', 0),               # fmBorderStyleNone
            @('BackColor', -2147483633)        # &H8000000F, the system form face color
        )) { $tb.($pair[0]) = $pair[1] }

        $btn = $d.Controls.Add('Forms.CommandButton.1', 'Close_Me', $true)
        $btn.Left    = $insideW - $pad - $btnW
        $btn.Top     = $textH + ($pad * 2)
        $btn.Width   = $btnW
        $btn.Height  = $btnH
        $btn.Caption     = 'Okay'         # "Okay", never "OK" - Jerry, 8/23/2026
        $btn.Accelerator = 'O'            # Alt+O presses it
        $btn.Default = $true              # Enter closes it
        $btn.Cancel  = $true              # Esc closes it too
    }

    if ($MessageDialog) {
        # The shared message shape. Same read-only text box as -InfoDialog, plus the two
        # buttons every VistaType message is allowed to end in.
        #
        # SIZES HERE ARE A STARTING POINT, NOT THE FINAL LAYOUT: Sh_Message_Form re-heights
        # itself and moves its buttons in UserForm_Initialize, because one message is a line
        # long and the next is a screen. What has to be right in the .frx is what cannot be
        # set from code without ceremony - the names, the captions, the accelerators, the
        # fonts, and Default/Cancel.
        $pad     = 12
        $btnW    = 72
        $btnH    = 24
        $gap     = 8
        # Measured, not assumed - see the note in the -InfoDialog block above.
        $insideW = $Width  - 12
        $insideH = $Height - 30
        $textH   = $insideH - $btnH - ($pad * 3)
        # Precomputed for the same reason the -InfoDialog block precomputes: PowerShell's
        # comma binds tighter than arithmetic inside an array literal.
        $textW   = $insideW - ($pad * 2)
        $cxLeft  = $insideW - $pad - $btnW
        $okLeft  = $cxLeft - $gap - $btnW
        $btnTop  = $textH + ($pad * 2)

        $tb = $d.Controls.Add('Forms.TextBox.1', 'Msg_Text', $true)
        # $tb.Font.Size here throws a null reference - the VBA block at the bottom does fonts.
        foreach ($pair in @(
            @('Left', $pad), @('Top', $pad),
            @('Width', $textW), @('Height', $textH),
            @('MultiLine', $true), @('WordWrap', $true),
            @('ScrollBars', 2),                # fmScrollBarsVertical
            @('Locked', $true), @('TabStop', $false),
            @('SpecialEffect', 0),             # fmSpecialEffectFlat
            @('BorderStyle', 0),               # fmBorderStyleNone
            @('BackColor', -2147483633)        # &H8000000F, the system form face color
        )) { $tb.($pair[0]) = $pair[1] }

        $ok = $d.Controls.Add('Forms.CommandButton.1', 'Okay_Button', $true)
        $ok.Left        = $okLeft
        $ok.Top         = $btnTop
        $ok.Width       = $btnW
        $ok.Height      = $btnH
        $ok.Caption     = 'Okay'
        $ok.Accelerator = 'O'
        $ok.Default     = $true            # Enter presses it

        $cx = $d.Controls.Add('Forms.CommandButton.1', 'Cancel_Button', $true)
        $cx.Left        = $cxLeft
        $cx.Top         = $btnTop
        $cx.Width       = $btnW
        $cx.Height      = $btnH
        $cx.Caption     = 'Cancel'
        $cx.Accelerator = 'C'
        $cx.Cancel      = $true            # Esc presses it
    }

    # --- Jerry's form conventions: Tahoma 10 and a ControlTipText on every button (8/7/2026),
    # and an accelerator on Okay and Cancel (8/23/2026).
    #
    # Done LAST, so it catches every control however it was added, and done in VBA rather than
    # from PowerShell. Setting the font from PowerShell silently does nothing: "$d.Font.Size =
    # 10" reports success and reads back empty, and $comp.Properties.Item('Font') throws
    # "Invalid object use". From VBA the same properties behave.
    #
    # Each CONTROL is set as well as the form. Setting only the form is not enough: MSForms
    # snaps the form's own size to 9.75, while a control takes 10 exactly, and the controls are
    # what anyone actually sees.
    #
    # The tip is what a screen reader announces. Several of the transcribers using this are
    # visually impaired themselves, which is the whole point of it.
    $tidy = $proj.VBComponents.Add(1)
    $tidy.CodeModule.AddFromString(@"
Sub Zz_Apply_Form_Conventions()
    Dim c As Object, ctl As Object
    Set c = ThisDocument.VBProject.VBComponents("$Name")
    On Error Resume Next

    c.Designer.Font.Name = "Tahoma"
    c.Designer.Font.Size = 10

    For Each ctl In c.Designer.Controls
        ctl.Font.Name = "Tahoma"
        ctl.Font.Size = 10
        If TypeName(ctl) = "CommandButton" Then
            ctl.ControlTipText = ctl.Caption & " Button"
            ' The two accelerators Jerry fixed on 8/23/2026. Set here as well as where the
            ' button is created, so a button added to this script later cannot miss them.
            If ctl.Caption = "Okay" Then ctl.Accelerator = "O"
            If ctl.Caption = "Cancel" Then ctl.Accelerator = "C"
        End If
    Next ctl
End Sub
"@)
    $doc.Activate()
    $word.Run('Zz_Apply_Form_Conventions') | Out-Null
    $proj.VBComponents.Remove($tidy)
    Write-Host "applied Tahoma 10, button tips and Okay/Cancel accelerators"

    $comp.Export($target)
    Write-Host "created $Name -> $target (+ .frx)"
    Write-Host "next: append the form's code to the .frm as text (CRLF), then check_frm_eol.py"
}
finally {
    if ($doc -ne $null) { $doc.Close(0) }    # wdDoNotSaveChanges
    $word.Quit()
}
