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
    ' Version: 1.2 Date: 8/30/2026
    '   - the "file must be saved" test now runs BEFORE the two bookmarks are added.
    '     It leaves by Exit Sub, which skips CleanExit, so on an unsaved document it was
    '     leaving DxExportStart and DxExportEnd behind in the transcriber's file
    '   - BackgroundSave is only restored if it was actually captured. Cancelling the
    '     export dialog jumped to CleanExit BEFORE the capture, so it wrote an
    '     uninitialized False into Application.Options.BackgroundSave - switching
    '     background saving off for Word itself, application-wide and permanently
    '   - refuses to export onto the document being exported from (large print has always
    '     had this test; braille went ahead and failed inside SaveAs2)
    '   - writes .docx, and forces the name to match, as the large print macro does. It
    '     used srcDoc.SaveFormat, so exporting from a .docm produced a macro-enabled file
    '   - every dialog now carries its own number
    '   - the exported file ends with ONE paragraph mark. A second one was added
    '     unconditionally, giving every braille export a trailing blank line
    ' Version: 1.1 Date: 2/13/2026
    '
    On Error GoTo ErrHandler

    Dim srcDoc As Document, destDoc As Document
    Dim rngSel As Range
    Dim undoRec As UndoRecord
    Dim fd As FileDialog
    
    Dim exportPath As String
    Dim originalBackgroundSave As Boolean
    Dim bgSaveCaptured As Boolean
    Dim userChoice As VbMsgBoxResult
    Dim startPos As Long, endPos As Long
    Dim errNum As Long
    Dim errText As String
    
    '===========================================================
    ' 1. VALIDATE, READ-ONLY CHECK & ANCHOR SELECTION
    '===========================================================
    Set srcDoc = ActiveDocument
    
    If srcDoc.ReadOnly Then
        MsgBox "This document is 'Read-Only'. " & vbCrLf & vbCrLf & _
               "Please use a copy of the document with the 'Read-Only' state turned off.", _
               vbCritical, "Braille Macros (254)"
        Exit Sub
    End If
    
    If Selection.Type = wdSelectionIP Then
        MsgBox "Please select the text you wish to export.", vbExclamation, "Braille Macros (255)"
        Exit Sub
    End If

    ' The file must be saved so the bookmarks below exist in the disk/cloud copy that
    ' step 4 clones. TESTED BEFORE THE BOOKMARKS ARE ADDED, not after: this leaves by
    ' Exit Sub, which does not pass through CleanExit, so adding them first left two
    ' stray bookmarks in the transcriber's document every time it fired.
    If srcDoc.Path = "" Then
        MsgBox "File must be saved before exporting selection.", vbExclamation, "Braille Macros (256)"
        Exit Sub
    End If

    Set rngSel = Selection.Range
    
    ' Place point-bookmarks to "lock" the boundaries
    ' We use these to identify the "keep zone" in the cloned file
    srcDoc.Bookmarks.Add Name:="DxExportStart", Range:=srcDoc.Range(rngSel.start, rngSel.start)
    srcDoc.Bookmarks.Add Name:="DxExportEnd", Range:=srcDoc.Range(rngSel.End, rngSel.End)
    
    Dx_UpdateProgressBar "Syncing source file...", 10
    srcDoc.Save ' CRITICAL: Ensures the new bookmarks are in the file used as the template

    '===========================================================
    ' 2. CHOOSE EXPORT LOCATION
    '===========================================================
    Do
        Set fd = Application.FileDialog(msoFileDialogSaveAs)
        fd.Title = "Export Selection As..."
        fd.InitialFileName = srcDoc.Path & "\"

        If fd.Show = -1 Then
            ' Force .docx, because wdFormatXMLDocument is what SaveAs2 writes below.
            exportPath = Dx_Force_Docx_Extension(fd.SelectedItems(1))

            If LCase(exportPath) = LCase(srcDoc.fullName) Then
                ' Exporting onto the source would have SaveAs2 write over a document that
                ' is open, which fails somewhere inside Word and arrives as a raw number.
                MsgBox "You cannot use the current file as the target file for export!", _
                       vbExclamation, "Braille Macros (257)"
            Else
                Exit Do
            End If
        Else
            GoTo CleanExit
        End If
    Loop

    '===========================================================
    ' 3. SETUP & OPTIMIZATION
    '===========================================================
    Application.ScreenUpdating = False
    originalBackgroundSave = Application.Options.BackgroundSave
    bgSaveCaptured = True
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
    destDoc.SaveAs2 fileName:=exportPath, FileFormat:=wdFormatXMLDocument, AddToRecentFiles:=True

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
    
    ' Ensure clean paragraph ending - exactly ONE mark.
    ' A second, unconditional InsertAfter stood here, so every exported braille file ended
    ' with a blank paragraph, which is a blank line in the braille. Jerry, 8/30/2026.
    If Right(destDoc.Range.Text, 1) <> vbCr Then destDoc.Range.InsertAfter vbCr

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
        "Braille Macros (258)")
    
    If userChoice = vbYes Then
        Set undoRec = Application.UndoRecord
        undoRec.StartCustomRecord "Export Deletion"
            rngSel.Delete
        undoRec.EndCustomRecord
        MsgBox "Deleted selection can be restored with Ctrl+Z", , "Braille Macros (259)"
    End If

CleanExit:
    On Error Resume Next
    ' Cleanup source bookmarks if we crashed before step 6
    If Not srcDoc Is Nothing Then
        If srcDoc.Bookmarks.Exists("DxExportStart") Then srcDoc.Bookmarks("DxExportStart").Delete
        If srcDoc.Bookmarks.Exists("DxExportEnd") Then srcDoc.Bookmarks("DxExportEnd").Delete
    End If
    
    ' ONLY if it was actually read. Cancelling the export dialog jumps straight here from
    ' step 2, before the capture, and writing back an uninitialized Boolean switched
    ' background saving off for Word itself - permanently, and nobody asked for it.
    If bgSaveCaptured Then Application.Options.BackgroundSave = originalBackgroundSave

    Application.ScreenUpdating = True
    Application.StatusBar = ""
    
    ' Close destDoc if it was left open during an error
    If Not destDoc Is Nothing Then destDoc.Close SaveChanges:=False
    
    Set destDoc = Nothing: Set srcDoc = Nothing: Set fd = Nothing
    Exit Sub

ErrHandler:
    ' Number and wording first: VBA clears Err on ANY On Error statement.
    errNum = Err.Number
    errText = Err.Description

    Application.ScreenUpdating = True
    Sh_Say "Export failed in Dx_ExportSelectionToNewFile." & vbCr & _
           "Error " & errNum & ": " & errText, _
           "Braille Macros (260)"
    Resume CleanExit
End Sub '*** end of Dx_ExportSelectionToNewFile macro ***

'=== FORCE THE .docx EXTENSION ===
'
' Version 1.0 Date: 8/30/2026
'
' The twin of Lp_Force_Docx_Extension. Kept as a separate copy rather than shared because
' these two modules stand alone; if a third caller ever wants it, move it to LPandBrlMacros
' as Sh_Force_Docx_Extension and delete both.
'
Private Function Dx_Force_Docx_Extension(ByVal filePath As String) As String
    Dim lastDot As Long, lastSlash As Long

    If LCase(Right(filePath, 5)) = ".docx" Then
        Dx_Force_Docx_Extension = filePath
        Exit Function
    End If

    lastDot = InStrRev(filePath, ".")
    lastSlash = InStrRev(filePath, "\")
    If lastDot > lastSlash And lastDot > 0 Then filePath = Left(filePath, lastDot - 1)

    Dx_Force_Docx_Extension = filePath & ".docx"
End Function '*** end of Dx_Force_Docx_Extension ***

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
    ' Version: 1.5  Date: 8/30/2026
    '   - one failure message shared with the large print macro: it now names the macro
    '     AND carries the error number, and goes through Sh_Say rather than MsgBox
    '   - the source document is closed before the message, and under Resume Next
    ' Version: 1.4  Date: 8/30/2026
    '   - inserts by DIRECT RANGE TRANSFER instead of the clipboard, the way
    '     Lp_Import_Exported_Selection_File has done since its own rewrite. The
    '     transcriber's clipboard is left alone, the insertion point is fixed before
    '     the dialog opens so it cannot drift while another document is opened and
    '     closed, and Error 4605 cannot arise. Lp_Normalize_Styles is deliberately
    '     NOT carried over - it is large print's and has no braille equivalent
    '   - the error handler now closes the source document. It is opened
    '     Visible:=False, so a failure between opening it and reading it left a
    '     document open in the session with NO WINDOW, which the transcriber could
    '     neither see nor close
    ' Version: 1.3  Date: 8/30/2026
    '   - refuses to import the document into itself (see step 2a)
    ' Version: 1.2  Date: 1/23/2026
    '
    Dim fd As FileDialog
    Dim strFilePath As String
    Dim sourceDoc As Document
    Dim destDoc As Document
    Dim sourceRange As Range
    Dim targetRange As Range
    Dim errNum As Long
    Dim errText As String
    
    Set destDoc = ActiveDocument
    
    ' 1. ANCHOR THE DESTINATION IMMEDIATELY
    ' Collapsed to a point so existing text is never overwritten, and taken NOW so the
    ' landing place cannot move while the source document is opened and closed below.
    Set targetRange = Selection.Range
    targetRange.Collapse Direction:=wdCollapseStart
    
    ' 2. SELECT THE FILE
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

    ' 2a. A document cannot import itself.
    ' Documents.Open on a file that is already open hands back THAT SAME document,
    ' so sourceDoc and destDoc become one and the same. Step 6 then closes the
    ' document being written into, and every range still pointing at it dies -- which
    ' reaches the user as run-time error 5825, "Object has been deleted", after the
    ' text has already been inserted.
    If StrComp(strFilePath, destDoc.FullName, vbTextCompare) = 0 Then
        Sh_Say "This is the document you are importing into." & vbCr & _
               "A file cannot import itself. Choose a different file.", _
               "Braille Macros (248)"
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    
    ' 3. OPEN SOURCE (HIDDEN)
    Application.ScreenUpdating = False
    Set sourceDoc = Documents.Open(fileName:=strFilePath, Visible:=False)
    
    ' 4. DEFINE SOURCE CONTENT, TRIMMING THE TRAILING PARAGRAPH MARKS
    ' Loops back from the end to where the text, tables and images really stop, so that
    ' only ONE mark comes across.
    Set sourceRange = sourceDoc.Content
    Do While sourceRange.Characters.count > 1 And _
       (sourceRange.Characters.Last.Previous.Text = vbCr Or _
        sourceRange.Characters.Last.Previous.Text = Chr(13))
        sourceRange.End = sourceRange.End - 1
    Loop
    
    ' 5. DIRECT RANGE TRANSFER
    ' Replaces sourceRange.Copy + Selection.PasteAndFormat. Nothing goes near the
    ' clipboard, so whatever the transcriber had on it survives the import.
    targetRange.FormattedText = sourceRange.FormattedText
    
    ' 6. CLEANUP
    sourceDoc.Close SaveChanges:=False
    Set sourceDoc = Nothing
    
    destDoc.Activate
    Application.ScreenUpdating = True
    
    ' 7. PUT THE CURSOR AFTER THE INSERTED TEXT AND SCROLL TO IT
    targetRange.Select
    Selection.Collapse Direction:=wdCollapseEnd
    ActiveWindow.ScrollIntoView Selection.Range, True
    
    Exit Sub

ErrorHandler:
    ' Take the number and the wording NOW, before anything else. VBA clears the Err
    ' object on ANY On Error statement, so reading Err after the one below would report
    ' error 0 with no description -- and the number is one of the four things
    ' docs/Reported-Errors.md is keyed on.
    errNum = Err.Number
    errText = Err.Description

    Application.ScreenUpdating = True

    ' Close the hidden source document BEFORE saying anything, so it cannot be stranded
    ' by a message that fails to appear. Resume Next because a failure in the cleanup of
    ' a path that has already gone wrong would reach the user as Word's own run-time
    ' error dialog -- the one offering Debug, which opens this source on their machine.
    On Error Resume Next
    If Not sourceDoc Is Nothing Then sourceDoc.Close SaveChanges:=False
    On Error GoTo 0

    Sh_Say "Import failed in Dx_Import_Exported_Selection_File." & vbCr & _
           "Error " & errNum & ": " & errText, _
           "Braille Macros (250)"
End Sub   '*** end of Dx_Import_Exported_Selection_File  ***





