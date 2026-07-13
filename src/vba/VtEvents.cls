VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VtEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' VistaType LP - Word application-event sink.
'
' Hooking Word's application events lets document-type detection run for every opened /
' new / closing document even when the add-in is loaded from the STARTUP folder (where the
' classic AutoOpen/AutoNew/AutoClose macros do not fire). Instantiated and wired once by
' AutoExec in LPandBrlMacros (module-level gEvents keeps this sink alive for the session).
Option Explicit

Public WithEvents App As Word.Application
Attribute App.VB_VarHelpID = -1

Private Sub App_DocumentOpen(ByVal Doc As Document)
    Sh_HandleDocumentOpened
End Sub

Private Sub App_NewDocument(ByVal Doc As Document)
    Sh_HandleDocumentNew
End Sub

Private Sub App_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)
    Sh_HandleDocumentClosing
End Sub
