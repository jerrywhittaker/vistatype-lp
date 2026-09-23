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
    Sh_HandleDocumentNew Doc
End Sub

Private Sub App_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)
    Sh_HandleDocumentClosing Doc
End Sub

' Raised every time a document window becomes the active one - clicking another window,
' Ctrl+F6, View > Switch Windows, returning to Word from another program, and as part of
' opening or creating a document. Word's Options are per-APPLICATION, so this is where the
' three configurations (large print, braille, default) follow the document the transcriber
' is actually looking at. Sh_HandleDocumentActivated does the cheap tests first; see the
' comment above it in LPandBrlMacros. 8/9/2026.
Private Sub App_WindowActivate(ByVal Doc As Document, ByVal Wn As Window)
    Sh_HandleDocumentActivated
End Sub

' WindowActivate alone is NOT enough, and that is not a guess: instrumented on the build box
' 8/9/2026, Word raised it ZERO times - not for Document.Activate, not for Windows(n).Activate,
' not while a document was being created. WindowSelectionChange is the one that can be relied
' on, because moving to another document's window moves the selection with it. It also fires on
' every cursor movement, which is why Sh_HandleDocumentActivated's first test is a single string
' comparison. Both are hooked: WindowActivate is the RIGHT event and costs nothing when it does
' fire, and the handler is safe to call twice.
'
' The TOC box (354) and Resize Pictures in a Selected Range (375) watch the cursor through here
' too, from 9/23/2026 - see Lp_Tocb_Watch_Cursor and Lp_Rst_Watch_Cursor. Each returns on its first
' test while its box is down.
Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
    Sh_HandleDocumentActivated
    Lp_Tocb_Watch_Cursor Sel
    Lp_Rst_Watch_Cursor Sel
End Sub
