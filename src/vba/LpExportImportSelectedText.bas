Attribute VB_Name = "LpExportImportSelectedText"
Option Explicit

'=== API DECLARATIONS ===
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub Lp_Export_Selection_To_NewFile()
    '
    ' copies selected text to a new file - Exact Clone of Master DNA
    '
    ' Version 1.1 Date: 2/13/2026
    '
    Dim srcDoc As Document, destDoc As Document
    Dim exportPath As String, originalPath As String
    Dim masterTemplate As String
    Dim originalBackgroundSave As Boolean
    Dim userChoice As VbMsgBoxResult
    Dim selStart As Long, selEnd As Long
    
    ' 1. VALIDATION & MASTER SAVE
    Set srcDoc = ActiveDocument
    If srcDoc.ReadOnly Then MsgBox "Selections cannot be exported from this file because it is marked as Read-Only.", vbCritical: Exit Sub
    If Selection.Type = wdSelectionIP Then
        MsgBox "Select text to be exported first.", vbExclamation, "VistaType LP (224)"
        Exit Sub
    End If
    
    ' Ensure the master is saved so we have a valid path for cloning
    If srcDoc.Path = "" Then
        userChoice = MsgBox("This document must be saved before exporting." & vbCrLf & vbCrLf & _
                            "Click OK to save now.", vbOKCancel + vbInformation, "VistaType LP")
        If userChoice = vbOK Then
            If Dialogs(wdDialogFileSaveAs).Show = 0 Then Exit Sub
            Set srcDoc = ActiveDocument
        Else
            Exit Sub
        End If
    End If

    selStart = Selection.start
    selEnd = Selection.End
    originalPath = srcDoc.fullName
    masterTemplate = srcDoc.AttachedTemplate.fullName ' Capture the actual template
    srcDoc.Save

    ' 2. EXPORT FILE DIALOG
    Do
        With Application.FileDialog(msoFileDialogSaveAs)
            .Title = "Export Selection As..."
            .InitialFileName = srcDoc.Path & "\"
            If .Show = -1 Then
                exportPath = .SelectedItems(1)
                If LCase(exportPath) = LCase(originalPath) Then
                    MsgBox "You cannot use the current file as the target file for export!", vbExclamation, "VistaType LP (225)"
                Else
                    Exit Do
                End If
            Else
                Exit Sub
            End If
        End With
    Loop

    ' 3. PREP
    Application.ScreenUpdating = False
    originalBackgroundSave = Application.Options.BackgroundSave
    Application.Options.BackgroundSave = False
    
    ' 4. THE EXPORT (INHERITING MASTER CHARACTERISTICS)
    Lp_UpdateProgressBar "Inheriting Master DNA...", 30
    
    ' Copy selection using FormattedText to preserve every single detail
    Dim rangeToCopy As Range
    Set rangeToCopy = Selection.Range
    
    ' Create the new document based on the Master File itself
    ' This is the only way to avoid the "Normal.dotm" attachment
    On Error Resume Next
    Set destDoc = Documents.Add(Template:=originalPath, Visible:=False)
    
    ' FALLBACK: If OneDrive blocks Template-loading, create blank and force-attach
    If Err.Number <> 0 Then
        Err.Clear
        Set destDoc = Documents.Add(Visible:=False)
        destDoc.AttachedTemplate = masterTemplate
    End If
    On Error GoTo 0
    
    Lp_UpdateProgressBar "Customizing Export...", 60
    ' Wipe the body content of the clone (Headers/Styles remain)
    destDoc.Content.Delete
    
    ' Paste the selection
    destDoc.Range(0, 0).FormattedText = rangeToCopy.FormattedText
    
    ' Explicitly copy PageSetup again just to be 100% certain
    destDoc.PageSetup = srcDoc.PageSetup
    
    ' Standard cleanup
    With destDoc.Range.Find
        .Text = "^b": .Replacement.Text = "^p": .Execute Replace:=wdReplaceAll
    End With
    
    ' Save and close the background doc
    destDoc.SaveAs2 fileName:=exportPath, FileFormat:=wdFormatXMLDocument
    destDoc.Close SaveChanges:=wdDoNotSaveChanges
    
    ' 5. RESTORE UI & FOCUS
    Lp_UpdateProgressBar "Finalizing Master...", 90
    srcDoc.Activate
    srcDoc.Range(selStart, selEnd).Select
    
    Application.ScreenUpdating = True
    Lp_UpdateProgressBar "Complete!", 100
    DoEvents
    
    ' 6. FINAL CHOICE (STILL IN MASTER)
    userChoice = MsgBox("Export Successful." & vbCrLf & vbCrLf & _
                        "The exported file will appear in Word's recent file list." & vbCrLf & vbCrLf & _
                        "Delete the selected text from this file?", _
                        vbYesNo + vbQuestion, "VistaType LP (227)")

    If userChoice = vbYes Then
        ' SMART DELETE
        If Right(Selection.Text, 1) <> Chr(13) And Selection.End < srcDoc.Content.End - 1 Then
            Selection.MoveEnd Unit:=wdCharacter, count:=1
        End If
        Selection.Delete
        
        ' END OF DOCUMENT GUARD
        If Selection.End >= srcDoc.Content.End - 1 Then
            If srcDoc.Characters.count > 1 Then
                If srcDoc.Characters(srcDoc.Characters.count - 1).Text <> Chr(13) Then
                    srcDoc.Content.InsertAfter vbCr
                End If
            Else
                srcDoc.Content.InsertAfter vbCr
            End If
        End If
        Application.ScreenRefresh
    End If

CleanExit:
    Application.Options.BackgroundSave = originalBackgroundSave
    Application.StatusBar = False

End Sub '  *** end of Lp_Export_Selection_To_NewFile macro ***

'=== PROGRESS BAR HELPER ===
' This must exist in the same module for the main macro to find it!
'
' Version 1.1 Date: 2/13/2026
'
Private Sub Lp_UpdateProgressBar(msg As String, pct As Long)
    Dim bars As String: Dim barCount As Integer
    Dim endTime As Double
    Dim frm As Object ' Needed for the generic close loop
    
    ' 1. Construct the visual bar
    barCount = Int(pct / 5)
    bars = String(barCount, "|") & String(20 - barCount, " ")
    
    ' 2. Update the UI
    Application.StatusBar = "[" & bars & "] " & pct & "% - " & msg
    Call Sh_ShowNonModalMessage("VistaType LP is Working", _
         msg & " (" & pct & "%)" & vbCrLf & "Progress: [" & bars & "]")
    
    ' 3. Force the pause so the user sees the update
    DoEvents
    endTime = Timer + 0.3
    Do While Timer < endTime
        DoEvents
    Loop
    
    ' 4. EXIT LOGIC: Clean up when we reach 100%
    If pct >= 100 Then
        ' Brief pause so they see the 100% mark
        endTime = Timer + 0.8
        Do While Timer < endTime: DoEvents: Loop
        
        ' --- THE GENERIC CLOSE LOOP ---
        ' This closes ANY nonmodal message window currently on screen
        On Error Resume Next
        For Each frm In VBA.UserForms
            Unload frm
        Next frm
        On Error GoTo 0
        
        ' Clear the Status Bar
        Application.StatusBar = ""
    End If

End Sub '*** end of Lp_UpdateProgressBar macro ***

'=====================================================================

Sub Lp_Import_Exported_Selection_File()
    Dim fd As FileDialog
    Dim strFilePath As String
    Dim sourceDoc As Document
    Dim destDoc As Document
    Dim sourceRange As Range
    Dim targetRange As Range
    
    Set destDoc = ActiveDocument
    
    ' 1. Anchor the destination immediately
    ' We collapse the range to a point so we don't overwrite existing text
    Set targetRange = Selection.Range
    targetRange.Collapse Direction:=wdCollapseStart
    
    ' 2. Select the file
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select the Word file to import"
        .Filters.Clear
        .Filters.Add "Word Documents", "*.docx; *.doc; *.docm", 1
        If .Show = -1 Then
            strFilePath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With

    On Error GoTo ErrorHandler
    
    ' 3. Open Source (Hidden)
    Application.ScreenUpdating = False
    Set sourceDoc = Documents.Open(fileName:=strFilePath, Visible:=False)

    ' 4. Normalize
    sourceDoc.Activate
    ' Using a separate function check
    If Lp_Is_The_Attached_Template_LP = True Then
        Call Lp_Normalize_Styles
    End If
    
    ' 5. Define Source Content (Excluding the final paragraph mark)
    Set sourceRange = sourceDoc.Content
    Do While sourceRange.Characters.count > 1 And _
       (sourceRange.Characters.Last.Previous.Text = vbCr Or _
        sourceRange.Characters.Last.Previous.Text = Chr(13))
        sourceRange.End = sourceRange.End - 1
    Loop
    
    ' 6. The "Fix" - Direct Range Transfer
    ' This replaces Copy/Paste and avoids Error 4605
    targetRange.FormattedText = sourceRange.FormattedText
    
    ' 7. Cleanup
    sourceDoc.Close SaveChanges:=False
    Set sourceDoc = Nothing
    
    destDoc.Activate
    Application.ScreenUpdating = True
    
    ' Final UI Polish: Move cursor to the end of the newly inserted text
    targetRange.Select
    Selection.Collapse Direction:=wdCollapseEnd
    ActiveWindow.ScrollIntoView Selection.Range, True
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description & " (Code: " & Err.Number & ")", vbCritical
    If Not sourceDoc Is Nothing Then sourceDoc.Close SaveChanges:=False
End Sub

