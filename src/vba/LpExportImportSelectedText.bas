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
    ' Version 1.2 Date: 8/30/2026
    '   - HAS AN ERROR HANDLER. It had none: CleanExit was never jumped to, so any
    '     failure left Application.Options.BackgroundSave switched OFF for Word itself,
    '     application-wide and surviving a restart, plus ScreenUpdating off and the
    '     hidden export document open
    '   - BackgroundSave is only restored if it was actually captured. Restoring an
    '     uninitialized Boolean would have WRITTEN False, causing the very fault the
    '     restore exists to prevent
    '   - the delete offered at the end is wrapped in an undo record, so Ctrl+Z brings
    '     the text back, and the dialog says so
    '   - the delete no longer absorbs a character it was not given: it only reaches
    '     past the selection when the next character IS a paragraph mark
    '   - the export name is forced to .docx, which is the format actually written
    '   - Application.StatusBar = "" rather than False. StatusBar is a String in Word
    '     (it is Excel where False restores the default), so False put the word "False"
    '     on the status bar
    ' Version 1.1 Date: 2/13/2026
    '
    Dim srcDoc As Document, destDoc As Document
    Dim exportPath As String, originalPath As String
    Dim masterTemplate As String
    Dim originalBackgroundSave As Boolean
    Dim bgSaveCaptured As Boolean
    Dim userChoice As VbMsgBoxResult
    Dim selStart As Long, selEnd As Long
    Dim undoRec As UndoRecord
    Dim errNum As Long
    Dim errText As String

    On Error GoTo ErrHandler

    ' 1. VALIDATION & MASTER SAVE
    Set srcDoc = ActiveDocument
    If srcDoc.ReadOnly Then MsgBox "Selections cannot be exported from this file because it is marked as Read-Only.", vbCritical, "VistaType LP (296)": Exit Sub
    If Selection.Type = wdSelectionIP Then
        MsgBox "Select text to be exported first.", vbExclamation, "VistaType LP (224)"
        Exit Sub
    End If
    
    ' Ensure the master is saved so we have a valid path for cloning
    If srcDoc.Path = "" Then
        userChoice = MsgBox("This document must be saved before exporting." & vbCrLf & vbCrLf & _
                            "Click OK to save now.", vbOKCancel + vbInformation, "VistaType LP (251)")
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

                ' Force .docx, because wdFormatXMLDocument is what SaveAs2 writes below.
                ' Without this, naming the file .docm produced a file whose contents and
                ' whose name disagreed, and Word complains when it is opened.
                exportPath = Lp_Force_Docx_Extension(exportPath)

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
    bgSaveCaptured = True
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
    ' Back onto the handler - On Error GoTo 0 here would have left the rest of the macro
    ' unguarded, which is what this version exists to stop.
    On Error GoTo ErrHandler

    ' If the fallback failed too, destDoc is Nothing and the next line would raise 91 with
    ' nothing to say for itself. Say what actually went wrong instead.
    If destDoc Is Nothing Then
        Err.Raise 5, , "The export document could not be created from " & originalPath
    End If

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
    Set destDoc = Nothing
    
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
        ' Wrapped in an undo record so the whole deletion comes back as ONE Ctrl+Z.
        ' Without it the tidy-up below could need several presses, and a transcriber who
        ' answered Yes by mistake had no obvious way back.
        Set undoRec = Application.UndoRecord
        undoRec.StartCustomRecord "VistaType LP Export Deletion"

        ' SMART DELETE
        ' Reach one character past the selection ONLY when that character is a paragraph
        ' mark, so a whole-paragraph export does not leave an empty paragraph behind.
        ' It used to extend by one character WITHOUT looking at it, so exporting part of a
        ' paragraph and answering Yes deleted one character of text nobody had selected.
        If Right(Selection.Text, 1) <> Chr(13) And Selection.End < srcDoc.Content.End - 1 Then
            If srcDoc.Range(Selection.End, Selection.End + 1).Text = Chr(13) Then
                Selection.MoveEnd Unit:=wdCharacter, count:=1
            End If
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

        undoRec.EndCustomRecord
        Application.ScreenRefresh

        Sh_Say "The text has been exported and removed from this document." & vbCr & _
               "Press Ctrl+Z to bring it back.", "VistaType LP (252)"
    End If

CleanExit:
    ' Nothing here may raise: this runs on a path that has already gone wrong as often as
    ' not, and a second error would reach the transcriber as Word's own run-time dialog.
    On Error Resume Next

    If Not destDoc Is Nothing Then destDoc.Close SaveChanges:=wdDoNotSaveChanges
    Set destDoc = Nothing

    ' ONLY if it was actually read. Writing back an uncaptured Boolean would switch
    ' background saving off for Word - the exact fault this restore is here to prevent.
    If bgSaveCaptured Then Application.Options.BackgroundSave = originalBackgroundSave

    Application.ScreenUpdating = True
    ' "" and not False: in Word StatusBar is a String, so False showed the word "False".
    Application.StatusBar = ""
    Exit Sub

ErrHandler:
    ' Number and wording first: VBA clears Err on ANY On Error statement, and the number
    ' is one of the four things docs/Reported-Errors.md is keyed on.
    errNum = Err.Number
    errText = Err.Description

    Application.ScreenUpdating = True
    Sh_Say "Export failed in Lp_Export_Selection_To_NewFile." & vbCr & _
           "Error " & errNum & ": " & errText, _
           "VistaType LP (253)"
    Resume CleanExit

End Sub '  *** end of Lp_Export_Selection_To_NewFile macro ***

'=== FORCE THE .docx EXTENSION ===
'
' Version 1.0 Date: 8/30/2026
'
' Both export macros write wdFormatXMLDocument, so the name must end .docx or the file's
' contents and its name disagree and Word objects when the transcriber opens it.
' The dot is only treated as an extension when it comes AFTER the last backslash -
' "C:\Braille Books 2026\chapter one" has a dot in the folder and none in the name.
'
Private Function Lp_Force_Docx_Extension(ByVal filePath As String) As String
    Dim lastDot As Long, lastSlash As Long

    If LCase(Right(filePath, 5)) = ".docx" Then
        Lp_Force_Docx_Extension = filePath
        Exit Function
    End If

    lastDot = InStrRev(filePath, ".")
    lastSlash = InStrRev(filePath, "\")
    If lastDot > lastSlash And lastDot > 0 Then filePath = Left(filePath, lastDot - 1)

    Lp_Force_Docx_Extension = filePath & ".docx"
End Function '*** end of Lp_Force_Docx_Extension ***

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
    '
    ' Imports another Word file into this one at the cursor.
    '
    ' Version: 1.2  Date: 8/30/2026
    '   - one failure message shared with the braille macro: it now names the macro AND
    '     carries the error number, and goes through Sh_Say rather than MsgBox
    '   - the source document is closed before the message, and under Resume Next
    ' Version: 1.1  Date: 8/30/2026
    '   - refuses to import the document into itself (see step 2a)
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

    ' 2a. A document cannot import itself.
    ' Documents.Open on a file that is already open hands back THAT SAME document,
    ' so sourceDoc and destDoc become one and the same. Step 7 then closes the
    ' document being written into, and every range still pointing at it dies -- which
    ' reaches the user as run-time error 5825, "Object has been deleted", after the
    ' text has already been inserted.
    If StrComp(strFilePath, destDoc.FullName, vbTextCompare) = 0 Then
        Sh_Say "This is the document you are importing into." & vbCr & _
               "A file cannot import itself. Choose a different file.", _
               "VistaType LP (247)"
        Exit Sub
    End If

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

    Sh_Say "Import failed in Lp_Import_Exported_Selection_File." & vbCr & _
           "Error " & errNum & ": " & errText, _
           "VistaType LP (249)"
End Sub

