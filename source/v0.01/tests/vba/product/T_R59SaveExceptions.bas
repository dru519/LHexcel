Attribute VB_Name = "T_R59SaveExceptions"
Option Explicit
Private mBook As Workbook
Private mFirst As Worksheet
Private mSecond As Worksheet

Public Function NxR59ExceptionSetup(ByVal kind As String, ByVal outputRoot As String) As String
    Set mBook = Workbooks.Add(xlWBATWorksheet)
    Set mFirst = mBook.Worksheets(1)
    mFirst.Name = "First"
    mFirst.Range("B2").Value2 = "preserve source"
    mFirst.Range("C3").Formula = "=21*2"
    Set mSecond = mBook.Worksheets.Add(After:=mFirst)
    mSecond.Name = "Second"
    mFirst.Select
    ActiveWindow.Zoom = 75
    mFirst.Range("D9").Select
    mSecond.Select
    ActiveWindow.Zoom = 125
    mSecond.Range("F12").Select
    NxR59ExceptionSetup = outputRoot & Application.PathSeparator & kind & ".xlsx"
    If kind = "write_error" Then mBook.SaveAs NxR59ExceptionSetup, xlOpenXMLWorkbook
    mSecond.Range("B2").Value2 = "pending edit"
End Function

Public Function NxR59ExceptionRun(ByVal kind As String, ByVal expectedPath As String) As String
    Dim result As CNxResult
    On Error GoTo Failed
    Set result = NxRunMannerSave(mBook)
    If kind = "new_save" Then
        Require result.Outcome = NxSuccess, "save outcome: " & result.Recovery
        Require StrComp(mBook.FullName, expectedPath, vbTextCompare) = 0, "unexpected SaveAs destination"
        Require ActiveSheet.Name = "First" And ActiveCell.Address = "$A$1" And ActiveWindow.Zoom = 100, "first saved view"
        mSecond.Select
        Require ActiveCell.Address = "$A$1" And ActiveWindow.Zoom = 100, "second saved view"
        Require mBook.FileFormat = xlOpenXMLWorkbook, "xlsx format"
    Else
        If kind = "new_cancel" Then
            Require result.Outcome = NxCancelled, "cancel outcome"
            Require Len(mBook.Path) = 0, "cancel created a file"
        Else
            Require result.Outcome = NxEnvironmentError, "injected write error must report failure: " & CStr(result.Outcome)
            Require Len(result.Recovery) > 0, "actionable write failure message"
        End If
        Require ActiveSheet.Name = "Second" And ActiveCell.Address = "$F$12" And ActiveWindow.Zoom = 125, "restore second view"
        mFirst.Select
        Require ActiveCell.Address = "$D$9" And ActiveWindow.Zoom = 75, "restore first view"
        Require Not mBook.Saved, "unsaved edit retained"
    End If
    Require mFirst.Range("B2").Value2 = "preserve source" And mFirst.Range("C3").Formula = "=21*2", "source values/formulas"
    Require mSecond.Range("B2").Value2 = "pending edit", "pending edit retained"
    Require Application.EnableEvents And Application.ScreenUpdating, "application state restored"
    NxR59ExceptionRun = "PASS|" & kind
    GoTo CleanUp
Failed:
    NxR59ExceptionRun = "FAIL|" & kind & "|" & Err.Description
CleanUp:
    NxR59ExceptionDispose
End Function

Public Sub NxR59ExceptionDispose()
    On Error Resume Next
    mBook.Close False
    Set mSecond = Nothing
    Set mFirst = Nothing
    Set mBook = Nothing
End Sub

Private Sub Require(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 5910, "r59 save exception probe", message
End Sub
