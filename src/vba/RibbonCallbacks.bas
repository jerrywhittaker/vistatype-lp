Attribute VB_Name = "RibbonCallbacks"
' VistaType LP - Ribbon dispatcher for the embedded customUI (customUI14.xml).
'
' The embedded ribbon (src/ribbon/customUI14.xml) routes every button through
' RibbonAction; the macro to run is carried in the control's Tag attribute. This
' lets the existing parameterless entry-point Subs (Lp_About, Dx_File_Fix_Sequence,
' ...) stay exactly as they are -- no signature changes needed.
'
' Requires a reference to the Microsoft Office object library (IRibbonControl),
' which the project already has (MSO.DLL).
'
' See DEVELOPMENT.md -> "Embedded ribbon".
Option Explicit

Public Sub RibbonAction(ByVal control As IRibbonControl)
    ' control.Tag holds the name of the macro to run (set in customUI14.xml).
    If Len(control.Tag) > 0 Then
        Application.Run control.Tag
    End If
End Sub
