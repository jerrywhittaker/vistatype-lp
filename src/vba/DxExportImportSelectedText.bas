Attribute VB_Name = "DxExportImportSelectedText"
Option Explicit

'=== API DECLARATIONS ===
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Sub Dx_ExportSelectionToNewFile()
    '
    ' Exports selection to a new document preserving formatting/settings
    ' Works with Local drives and OneDrive personal and OnDrive Enterprise
    '
    ' Version: 1.1 Date: 2/13/2026
    '
    On Error GoTo ErrHandler

    Dim srcDoc As Document, destDoc As Document
    Dim rngSel As Range
    Dim undoRec As UndoRecord
    Dim fd As FileDialog
    
    Dim exportPath As String
    Dim originalBackgroundSave As Boolean
    Dim userChoice As VbMsgBoxResult
    Dim startPos As Long, endPos As Long
    
    '===========================================================
    ' 1. VALIDATE, READ-ONLY CHECK & ANCHOR SELECTION
    '===========================================================
    Set srcDoc = ActiveDocument
    
    If srcDoc.ReadOnly Then
        MsgBox "This document is 'Read-Only'. " & vbCrLf & vbCrLf & _
               "Please use a copy of the document with the 'Read-Only' state turned off.", _
               vbCritical, "Braille Macros"
        Exit Sub
    End If
    
    If Selection.Type = wdSelectionIP Then
        MsgBox "Please select the text you wish to export.", vbExclamation, "Braille Macros"
        Exit Sub
    End If

    Set rngSel = Selection.Range
    
    ' Place point-bookmarks to "lock" the boundaries
    ' We use these to identify the "keep zone" in the cloned file
    srcDoc.Bookmarks.Add Name:="DxExportStart", Range:=srcDoc.Range(rngSel.start, rngSel.start)
    srcDoc.Bookmarks.Add Name:="DxExportEnd", Range:=srcDoc.Range(rngSel.End, rngSel.End)

    ' File must be saved so the bookmarks exist on the disk/cloud version
    If srcDoc.Path = "" Then
        MsgBox "File must be saved before exporting selection.", vbExclamation, "Braille Macros"
        Exit Sub
    End If
    
    Dx_UpdateProgressBar "Syncing source file...", 10
    srcDoc.Save ' CRITICAL: Ensures the new bookmarks are in the file used as the template

    '===========================================================
    ' 2. CHOOSE EXPORT LOCATION
    '===========================================================
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    fd.Title = "Export Selection As..."
    fd.InitialFileName = srcDoc.Path & "\"
    
    If fd.Show = -1 Then
        exportPath = fd.SelectedItems(1)
    Else
        GoTo CleanExit
    End If

    '===========================================================
    ' 3. SETUP & OPTIMIZATION
    '===========================================================
    Application.ScreenUpdating = False
    originalBackgroundSave = Application.Options.BackgroundSave
    ' Turn off background save to force code to wait until save is fully complete
    Application.Options.BackgroundSave = False
    
    '===========================================================
    ' 4. CLONE THE FILE (OneDrive Safe Method)
    '===========================================================
    Dx_UpdateProgressBar "Cloning document...", 30
    
    ' FSO fails on OneDrive URLs (https://...).
    ' Instead, we create a NEW document using the current file as its Template.
    ' This clones content, styles, and page setup perfectly.
    Set destDoc = Documents.Add(Template:=srcDoc.fullName, Visible:=False)
    
    ' Save the new clone immediately to the target path
    Dx_UpdateProgressBar "Establishing target file...", 40
    destDoc.SaveAs2 fileName:=exportPath, FileFormat:=srcDoc.SaveFormat, AddToRecentFiles:=True

    '===========================================================
    ' 5. FIXED BOOKMARK-BASED DELETION
    '===========================================================
    Dx_UpdateProgressBar "Isolating selection...", 50
    
    If destDoc.Bookmarks.Exists("DxExportEnd") And destDoc.Bookmarks.Exists("DxExportStart") Then
        
        ' Capture static positions
        startPos = destDoc.Bookmarks("DxExportStart").Range.start
        endPos = destDoc.Bookmarks("DxExportEnd").Range.End
        
        ' A. Delete Tail First (Bottom up prevents shifting indices)
        If endPos < destDoc.Range.End - 1 Then
            destDoc.Range(endPos, destDoc.Range.End - 1).Delete
        End If
        
        ' B. Delete Head
        If startPos > 0 Then
            destDoc.Range(0, startPos).Delete
        End If
        
        ' C. Cleanup Bookmarks in dest
        If destDoc.Bookmarks.Exists("DxExportStart") Then destDoc.Bookmarks("DxExportStart").Delete
        If destDoc.Bookmarks.Exists("DxExportEnd") Then destDoc.Bookmarks("DxExportEnd").Delete
    End If
    
    ' Ensure clean paragraph ending
    If Right(destDoc.Range.Text, 1) <> vbCr Then destDoc.Range.InsertAfter vbCr
    ' Note: We don't necessarily need a double return, just one is usually sufficient,
    ' but keeping your logic:
    destDoc.Range.InsertAfter vbCr

    '===========================================================
    ' 6. SAVE AND CLOSE
    '===========================================================
    Dx_UpdateProgressBar "Uploading to OneDrive (Please Wait)...", 80
    
    ' Force a screen refresh before the heavy network save
    DoEvents
    
    destDoc.Save
    destDoc.Close SaveChanges:=True
    Set destDoc = Nothing
    
    ' Cleanup bookmarks in the source document
    If srcDoc.Bookmarks.Exists("DxExportStart") Then srcDoc.Bookmarks("DxExportStart").Delete
    If srcDoc.Bookmarks.Exists("DxExportEnd") Then srcDoc.Bookmarks("DxExportEnd").Delete
    
    Dx_UpdateProgressBar "Finalizing...", 100
    
    '===========================================================
    ' 7. RETURN & UNDOABLE DELETE
    '===========================================================
    Application.StatusBar = "Ready"
    srcDoc.Activate
    Application.ScreenUpdating = True
    
    userChoice = MsgBox( _
        "Successful Export." & vbCrLf & vbCrLf & _
        "You are still in the parent document." & vbCrLf & vbCrLf & _
        "Delete this selection from this parent file?", _
        vbYesNo + vbQuestion + vbDefaultButton2, _
        "Braille Macros")
    
    If userChoice = vbYes Then
        Set undoRec = Application.UndoRecord
        undoRec.StartCustomRecord "Export Deletion"
            rngSel.Delete
        undoRec.EndCustomRecord
        MsgBox "Deleted selection can be restored with Ctrl+Z", , "Braille Macros"
    End If

CleanExit:
    On Error Resume Next
    ' Cleanup source bookmarks if we crashed before step 6
    If Not srcDoc Is Nothing Then
        If srcDoc.Bookmarks.Exists("DxExportStart") Then srcDoc.Bookmarks("DxExportStart").Delete
        If srcDoc.Bookmarks.Exists("DxExportEnd") Then srcDoc.Bookmarks("DxExportEnd").Delete
    End If
    
    Application.Options.BackgroundSave = originalBackgroundSave
    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    ' Close destDoc if it was left open during an error
    If Not destDoc Is Nothing Then destDoc.Close SaveChanges:=False
    
    Set destDoc = Nothing: Set srcDoc = Nothing: Set fd = Nothing
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Braille Macros"
    Resume CleanExit
End Sub '*** end of Dx_ExportSelectionToNewFile macro ***

'=== PROGRESS BAR HELPER ===
Private Sub Dx_UpdateProgressBar(msg As String, pct As Long)
    Dim bars As String: Dim barCount As Integer
    Dim endTime As Double
    Dim frm As Object ' Needed for the generic close loop
    
    ' 1. Construct the visual bar
    barCount = Int(pct / 5)
    bars = String(barCount, "|") & String(20 - barCount, " ")
    
    ' 2. Update the UI
    Application.StatusBar = "[" & bars & "] " & pct & "% - " & msg
    Call Sh_ShowNonModalMessage("Braille Macros are Working", _
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
End Sub ' end of Dx_UpdateProgressBar macro ***

'**************************************************************************************

Sub Dx_Import_Exported_Selection_File()
    '
    ' Version: 1.2  Date: 1/23/2026
    '
    Dim fd As FileDialog
    Dim strFilePath As String
    Dim sourceDoc As Document
    Dim destDoc As Document
    Dim sourceRange As Range
    
    Set destDoc = ActiveDocument
    
    ' 1. SELECT THE FILE
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

    ' 2. OPEN SOURCE FILE AND DEFINE CONTENT
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    
    Set sourceDoc = Documents.Open(fileName:=strFilePath, Visible:=False)
    
    ' Set the range to the whole document
    Set sourceRange = sourceDoc.Content
    
    ' 3. TRIM TRAILING PARAGRAPH MARKS
    ' This loops backwards from the end of the doc to find where
    ' the actual text/tables/images end, ensuring only ONE mark remains.
    Do While sourceRange.Characters.count > 1 And _
       (sourceRange.Characters.Last.Previous = vbCr Or _
        sourceRange.Characters.Last.Previous = Chr(13))
        sourceRange.End = sourceRange.End - 1
    Loop
    
    ' Copy the tightened range
    sourceRange.Copy
    sourceDoc.Close SaveChanges:=False
    
    ' 4. PASTE AND POSITION CURSOR
    destDoc.Activate
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    
    ' Collapse the selection to the end of the pasted content
    Selection.Collapse Direction:=wdCollapseEnd
    
    ' 5. REFRESH AND SCROLL TO CURSOR
    Application.ScreenUpdating = True
    DoEvents
    
    ActiveWindow.ScrollIntoView Selection.Range, True
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "An error occurred in Dx_Import_Exported_Selection_File: " & Err.Description, vbCritical

End Sub   '*** end of Dx_Import_Exported_Selection_File  ***





