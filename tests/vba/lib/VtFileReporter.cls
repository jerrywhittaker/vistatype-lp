VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VtFileReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' A vba-test reporter that writes to a FILE.
'
' vba-test ships ImmediateReporter, which writes to the Immediate Window. Nothing can read
' that back out of a Word started over SSH, so the results would be invisible - which is the
' whole point of running the tests from Linux. This writes the same information to a file the
' runner reads back.
'
' It also totals across suites rather than printing a summary per suite, so one run gives one
' verdict. The last line is RESULT OK, RESULT FAILED or RESULT ERROR, and the runner keys on it.
'
' Based on ImmediateReporter from https://github.com/vba-tools/vba-test (MIT, Tim Hall).
'
' Version: 1.0  Date: 9/20/2026
'
Option Explicit

Private WithEvents pSuite As TestSuite
Attribute pSuite.VB_VarHelpID = -1

Private pPath As String
Private pTotal As Long
Private pPassed As Long
Private pFailed As Long
Private pPending As Long
Private pSkipped As Long

Public Sub StartFile(ByVal resultPath As String)
    Dim f As Integer

    pPath = resultPath
    f = FreeFile
    Open pPath For Output As #f
    Close #f

    pTotal = 0
    pPassed = 0
    pFailed = 0
    pPending = 0
    pSkipped = 0
End Sub

Public Sub ListenTo(Suite As TestSuite)
    Set pSuite = Suite
    WriteLine "=== " & Suite.Description & " ==="

    ' The suites here are built complete and then handed over, so their tests have already
    ' finished. Report those before listening for any that follow.
    Dim Test As TestCase
    For Each Test In Suite.Tests
        WriteResult Test
    Next Test
End Sub

Public Sub Finish()
    WriteLine ""
    WriteLine "TOTAL " & pTotal
    WriteLine "PASSED " & pPassed
    WriteLine "FAILED " & pFailed
    WriteLine "PENDING " & pPending
    WriteLine "SKIPPED " & pSkipped
    If pTotal = 0 Then
        WriteLine "RESULT ERROR no tests ran"
    ElseIf pFailed > 0 Then
        WriteLine "RESULT FAILED"
    Else
        WriteLine "RESULT OK"
    End If
End Sub

Public Sub Trace(ByVal note As String)
    ' Every line is written straight through, so if the run hangs, the file still says how far
    ' it got. The first attempt at this suite buffered - Print # only flushes on Close - and a
    ' timed-out run left a nought-byte file and nothing to go on.
    WriteLine ".. " & note
End Sub

Public Sub Abort(ByVal reason As String)
    ' An unhandled VBA error puts up a modal dialog behind an invisible Word and hangs the
    ' whole run until it times out. The caller traps instead and says so here.
    WriteLine ""
    WriteLine "RESULT ERROR " & reason
End Sub

' ============================================= '

Private Sub WriteResult(Test As TestCase)
    pTotal = pTotal + 1

    Select Case Test.Result
        Case TestResultType.Pass
            pPassed = pPassed + 1
            WriteLine "PASS  " & Test.Description
        Case TestResultType.Fail
            pFailed = pFailed + 1
            WriteLine "FAIL  " & Test.Description
            Dim Failure As Variant
            For Each Failure In Test.Failures
                WriteLine "        " & Failure
            Next Failure
        Case TestResultType.Pending
            pPending = pPending + 1
            WriteLine "PEND  " & Test.Description
        Case TestResultType.Skipped
            pSkipped = pSkipped + 1
            WriteLine "SKIP  " & Test.Description
    End Select
End Sub

Private Sub WriteLine(ByVal s As String)
    Dim f As Integer

    If Len(pPath) = 0 Then Exit Sub

    f = FreeFile
    Open pPath For Append As #f
    Print #f, s
    Close #f
End Sub

Private Sub pSuite_Result(Test As TestCase)
    WriteResult Test
End Sub

Private Sub pSuite_Group(Suite As TestSuite)
    WriteLine "=== " & Suite.Description & " ==="
End Sub
