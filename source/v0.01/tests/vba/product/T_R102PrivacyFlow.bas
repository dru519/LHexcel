Attribute VB_Name = "T_R102PrivacyFlow"
Option Explicit
Public CancelStage As String
Public ScanPass As Long

Public Function Names() As String
    Names = "menu_links|dummy_layout|cancel_prepare|cancel_execute|cancel_report|menu_default_status"
End Function

Public Sub Checkpoint(ByVal stage As String)
    If stage = "scan" Then
        ScanPass = ScanPass + 1
        If CancelStage = "cancel_prepare" And ScanPass = 1 Then Err.Raise 18
        If CancelStage = "cancel_execute" And ScanPass = 2 Then Err.Raise 18
    ElseIf stage = "report" And CancelStage = "cancel_report" Then
        Err.Raise 18
    End If
End Sub

Public Function RunCase(ByVal name As String) As String
    Dim book As Workbook, report As Workbook, sheet As Worksheet, beforeBooks As Long
    Dim priorCancel As XlEnableCancelKey, priorStatus As Variant, priorEvents As Boolean
    Dim detail As String, originalStatus As Variant
    On Error GoTo Failed
    priorCancel = Application.EnableCancelKey: priorStatus = Application.StatusBar
    originalStatus = priorStatus
    priorEvents = Application.EnableEvents
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set sheet = book.Worksheets(1): sheet.Name = "개인'정보"
    sheet.Range("A1").Value2 = "성명": sheet.Range("A2").Value2 = "홍길동"
    sheet.Range("B1").Value2 = "연락처": sheet.Range("B2").Value2 = "010-1234-5678"
    sheet.Range("C1").Value2 = "업무일": sheet.Range("C2").Value2 = "2024-06-17"
    sheet.Range("D1").Value2 = "업무명": sheet.Range("D2").Value2 = "보고서"
    If name = "dummy_layout" Then
        sheet.Cells.Clear
        sheet.Range("B2").Value2 = "asff"
        sheet.Range("C3").Value2 = "010-5159-3999"
        sheet.Range("D3").Value2 = "890614-*******"
        sheet.Range("B4").Value2 = "홍길동이"
        sheet.Range("G3").Value2 = "김치"
        sheet.Range("H3").Value2 = "보고서"
        Require NxPrivacyAuditKind("891332-*******", "") = "", "invalid masked date accepted"
        Require NxPrivacyAuditKind("890614-1******", "") = "주민등록번호", "partly masked identifier missed"
        Require NxPrivacyAuditKind("890614-******", "") = "", "invalid mask length accepted"
    End If
    sheet.Range("D2").Select: book.Saved = True
    beforeBooks = Workbooks.Count
    CancelStage = name: ScanPass = 0
    Application.StatusBar = "Privacy flow restoration probe"
    If name = "menu_default_status" Then Application.StatusBar = vbNullString
    priorStatus = Application.StatusBar
    NxDataRunSpecialFromRibbon NX_FEATURE_DATA_PRIVACY_SCAN
    If (name = "menu_links" Or name = "dummy_layout" Or name = "menu_default_status") And Not ActiveWorkbook Is book Then Set report = ActiveWorkbook
    Require Application.EnableCancelKey = priorCancel, "cancel mode not restored"
    Require Application.StatusBar = priorStatus, "status not restored: " & CStr(priorStatus) & " -> " & CStr(Application.StatusBar)
    Require Application.EnableEvents = priorEvents, "events not restored"
    Require book.Saved, "source dirty"
    If name = "dummy_layout" Then
        Require sheet.Range("B4").Value2 = "홍길동이", "dummy source changed"
    Else
        Require sheet.Range("A2").Value2 = "홍길동", "source changed"
    End If
    If name = "menu_links" Or name = "dummy_layout" Or name = "menu_default_status" Then
        Require Workbooks.Count = beforeBooks + 1, "report absent"
        Set report = ActiveWorkbook
        Require Not report Is book, "source shown instead of report"
        Require report.Sheets(1).Name = "개인정보", "report name"
        If name = "dummy_layout" Then
            Require report.Sheets(1).Hyperlinks.Count = 3, "dummy findings or links missing"
            Require report.Sheets(1).Range("A8").Value2 = "주민등록번호", "masked kind"
            Require report.Sheets(1).Range("B8").Value2 = "890614-*******", "masked original changed"
            Require report.Sheets(1).Range("A9").Value2 = "이름 후보", "context name missed"
            report.Sheets(1).Range("C9").Hyperlinks(1).Follow
        Else
            Require report.Sheets(1).Hyperlinks.Count = 2, "menu result links missing"
            report.Sheets(1).Range("C7").Hyperlinks(1).Follow
        End If
        Require ActiveWorkbook Is book, "wrong link workbook"
        Require ActiveSheet Is sheet, "wrong link sheet"
        If name = "dummy_layout" Then
            Require ActiveCell.Address = "$B$4", "wrong dummy link cell"
        Else
            Require ActiveCell.Address = "$A$2", "wrong link cell"
        End If
        report.Close False: Set report = Nothing
    Else
        Require Workbooks.Count = beforeBooks, "cancel left partial report"
        Require ActiveWorkbook Is book, "cancel lost source"
        Require Selection.Address = "$D$2", "cancel lost selection"
    End If
    CancelStage = "": book.Close False
    Application.StatusBar = originalStatus
    RunCase = "PASS|" & name
    Exit Function
Failed:
    detail = CStr(Err.Number) & ":" & Err.Description
    On Error Resume Next
    CancelStage = ""
    If Not report Is Nothing Then report.Close False
    If Not book Is Nothing Then book.Close False
    Application.EnableCancelKey = priorCancel: Application.StatusBar = originalStatus
    Application.EnableEvents = priorEvents
    RunCase = "FAIL|" & name & "|" & detail
End Function

Private Sub Require(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 102, "PrivacyFlow", message
End Sub
