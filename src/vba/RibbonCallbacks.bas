Attribute VB_Name = "RibbonCallbacks"
' VistaType LP - Ribbon dispatcher and tab visibility for the embedded customUI.
'
' The embedded ribbon (src/ribbon/customUI14.xml) routes every button through
' RibbonAction; the macro to run is carried in the control's Tag attribute. This
' lets the existing parameterless entry-point Subs (Lp_About, Dx_File_Fix_Sequence,
' ...) stay exactly as they are -- no signature changes needed.
'
' TWO PLACES THE TABS CAN COME FROM
' ---------------------------------
' Word does not list add-in tabs in "File > Options > Customize the Ribbon", so tabs defined
' here cannot be hidden, reordered or renamed by the user. From 3.0.34 the installer instead
' writes VistaType's tabs into the user's OWN Word.officeUI, where Word treats them as
' ordinary custom tabs and all three become possible.
'
' The embedded tabs stay in place as the fallback, switched off by VtTabVisible when the
' installer says the user has their own. That single switch covers three cases at once:
'   * upgrading      - the embedded copies go dark in the same run that installs the user's,
'                      so the tabs never appear twice
'   * installing by hand (no installer, so no registry value) - the embedded tabs show
'                      exactly as they always did; nothing is lost on that path
'   * declining the installer option - same as installing by hand
'
' Requires a reference to the Microsoft Office object library (IRibbonControl / IRibbonUI),
' which the project already has (MSO.DLL).
'
' See DEVELOPMENT.md -> "Embedded ribbon".
Option Explicit

' Word hands this over once, at load. It is the only way to make Word re-ask a getVisible
' callback later. It goes Nothing if the VBA project resets (an unhandled error, or
' Debug > Reset) and stays Nothing until Word restarts - a known customUI quirk with no cure,
' so every use of it is guarded.
Private gRibbon As IRibbonUI

Private Const VT_REG_KEY  As String = "HKEY_CURRENT_USER\Software\VistaType LP\"
Private Const VT_REG_TABS As String = "UserRibbonTabs"

Public Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

Public Sub RibbonAction(ByVal control As IRibbonControl)
    ' control.Tag holds the name of the macro to run (set in customUI14.xml).
    '
    ' Sh_Pos_Depth is the nesting counter for Sh_Save_User_Position /
    ' Sh_Return_User_To_Start_Position. A macro that stops on an error never runs its
    ' closing Sh_Return_User_To_Start_Position, so the counter would stay above zero and
    ' every later macro would think it was nested and stop returning the user to their
    ' place. Clearing it here means the damage lasts one button press, not the session.
    Sh_Pos_Depth = 0
    Sh_Pos_Saved = False

    If Len(control.Tag) > 0 Then
        Application.Run control.Tag
    End If
End Sub

' Called by Word for the two visible embedded tabs. Show them only when the user does NOT
' have their own copies on the ribbon.
'
' An error raised inside a ribbon callback fails silently, and Word may then stop calling
' back for the rest of the session - so this must never be allowed to raise. Defaulting to
' True on any doubt is the safe direction: a duplicated tab is untidy, no tab at all looks
' like a broken install.
Public Sub VtTabVisible(ByVal control As IRibbonControl, ByRef returnedVal)
    On Error Resume Next
    returnedVal = Not Vt_User_Owns_Ribbon_Tabs()
    If Err.Number <> 0 Then
        returnedVal = True
        Err.Clear
    End If
End Sub

' True when the installer has put VistaType's tabs into the user's own Word.officeUI.
' Written by installer/scripts/Merge-Qat.ps1; absent on a hand-copied install.
Public Function Vt_User_Owns_Ribbon_Tabs() As Boolean
    Dim s As String
    On Error Resume Next
    s = System.PrivateProfileString("", VT_REG_KEY, VT_REG_TABS)
    On Error GoTo 0
    Vt_User_Owns_Ribbon_Tabs = (Trim$(s) = "1")
End Function

' Ask Word to re-read the getVisible callbacks. No-ops when the ribbon handle was lost,
' which is the normal state after a VBA project reset until Word is restarted.
Public Sub VtRefreshTabs()
    On Error Resume Next
    If Not gRibbon Is Nothing Then gRibbon.Invalidate
End Sub

' Escape hatch for someone who deletes the VistaType tabs out of their own ribbon: clearing
' the flag brings the built-in ones straight back, with no reinstall and no Word restart.
'
' NOT YET REACHABLE FROM THE UI. It needs a button somewhere that still works when there is
' no VistaType tab left to click -- the Document Settings dialog is the obvious home, since
' that is on the Quick Access Toolbar and on Ctrl+Alt+Shift+I. Until then it can only be run
' from the macro list (Alt+F8).
Public Sub Vt_Put_Tabs_Back_On_Ribbon()
    On Error Resume Next
    System.PrivateProfileString("", VT_REG_KEY, VT_REG_TABS) = "0"
    VtRefreshTabs
End Sub
