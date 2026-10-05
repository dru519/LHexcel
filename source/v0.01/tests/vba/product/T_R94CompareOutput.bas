Attribute VB_Name = "T_R94CompareOutput"
Option Explicit
Private mPanel As FNxWorkbookCompare, mPanelBook As Workbook
Public Function Names() As String
    Names = "selected_pair|append_repeat|blank_tail|defaults|reopen_links|formula_links|range_blank|all_sheets|file_popup"
End Function
Public Function RunCase(ByVal name As String) As String
    Dim a As Workbook, b As Workbook, report As Workbook, destination As Workbook
    Dim result As CNxResult, detail As Worksheet, summary As Worksheet, panel As Object
    Dim explorer As CNxCompareResultsSession
    Dim pa As String, pb As String, root As String, originalCount As Long, i As Long
    Dim oldEvents As Boolean, oldAlerts As Boolean, message As String
    Dim sourcePath As String, sheetName As String, cellAddress As String, issue As String, refSheet As Worksheet
    On Error GoTo Failed
    oldEvents = Application.EnableEvents: oldAlerts = Application.DisplayAlerts
    Application.EnableEvents = False: Application.DisplayAlerts = False
    Set a = Workbooks.Add(xlWBATWorksheet): Set b = Workbooks.Add(xlWBATWorksheet)
    a.Worksheets(1).Name = "기준'자료": b.Worksheets(1).Name = "비교 자료"
    a.Worksheets(1).Range("A1").Value2 = "항목": b.Worksheets(1).Range("A1").Value2 = "항목"
    a.Worksheets(1).Range("B2").Value2 = 100: b.Worksheets(1).Range("B2").Value2 = 120
    a.Worksheets(1).Range("Q1").Value2 = "보존": b.Worksheets(1).Range("Q1").Value2 = "보존"
    If name = "defaults" Then
        Set panel = New FNxDataCompare: panel.BindFeature "NX-FILE-WORKBOOK-COMPARE"
        Require panel.Controls("cboOutput").ListIndex = 0 And InStr(panel.Controls("cboOutput").Value, "새 창") > 0, "range default"
        Unload panel: Set panel = New FNxWorkbookCompare: panel.BindFeature
        Require panel.Controls("cboOutput").ListIndex = 0 And InStr(panel.Controls("cboOutput").Value, "새 창") > 0, "file default"
        Unload panel: Set panel = New FNxAgeCalculator: Load panel
        Require panel.Controls("cboOutputMode").Value = "새 통합문서", "age default"
        Unload panel: Set panel = Nothing
        GoTo Passed
    End If
    If name = "range_blank" Then
        b.Worksheets(1).Range("B2").Value2 = 100
        Set result = NxDataCompareCreate("NX-FILE-WORKBOOK-COMPARE", a.Worksheets(1).Range("A1:Q5"), b.Worksheets(1).Range("A1:Q2"))
        NxDataCompareFinish True: Set report = ActiveWorkbook
        Require InStr(report.Worksheets(1).Range("A3").Value2, "차이 0셀") > 0, "blank extent ignored"
        GoTo Passed
    End If
    root = Environ$("LHEXCEL_PROFILE_ROOT")
    pa = root & "\기준 " & name & ".xlsx": pb = root & "\비교 " & name & ".xlsx"
    If name = "blank_tail" Then
        a.Worksheets(1).Range("XFD100000").Interior.Color = vbRed
        b.Worksheets(1).Range("C2").Interior.Color = vbBlue
    End If
    a.SaveAs pa, xlOpenXMLWorkbook: b.SaveAs pb, xlOpenXMLWorkbook
    If name = "formula_links" Then
        a.Worksheets(1).Range("C2").Formula = "=B2*2"
        a.Save
        sourcePath = a.FullName
        Set refSheet = NxFormulaReferenceReport(a.Worksheets(1).Range("C2"), True)
        Set report = refSheet.Parent
        Require CanonicalPath(refSheet.Range("A2").Hyperlinks(1).Address) = CanonicalPath(a.FullName), "persisted source path: actual=" & refSheet.Range("A2").Hyperlinks(1).Address & "; expected=" & a.FullName
        report.SaveAs root & "\formula-report.xlsx", xlOpenXMLWorkbook
        report.Close False: Set report = Nothing
        a.Close False: Set a = Nothing
        Set report = Workbooks.Open(root & "\formula-report.xlsx", UpdateLinks:=0)
        report.Worksheets(1).Activate
        report.Worksheets(1).Range("A2").Hyperlinks(1).Follow
        Set a = ActiveWorkbook
        Require CanonicalPath(a.FullName) = CanonicalPath(sourcePath) And ActiveCell.Address = "$C$2", "closed formula source"
        report.Worksheets(1).Activate
        report.Worksheets(1).Range("C2").Hyperlinks(1).Follow
        Require ActiveWorkbook Is a, "reference source file: actual=" & ActiveWorkbook.FullName & "/" & ActiveSheet.Name & "/" & ActiveCell.Address & "; expected=" & a.FullName & "; link=" & report.Worksheets(1).Range("C2").Hyperlinks(1).Address & "#" & report.Worksheets(1).Range("C2").Hyperlinks(1).SubAddress
        Require ActiveCell.Address = "$B$2", "reference source cell"
        GoTo Passed
    End If
    If name = "all_sheets" Then
        b.Worksheets(1).Name = a.Worksheets(1).Name
        a.Worksheets.Add.Name = "추가 시트"
        a.Worksheets("추가 시트").Range("A1").Value2 = "기준 전용"
        a.Save: b.Save
        Set result = NxWorkbookCompareReport(pa, pb, "", False)
        NxWorkbookCompareFinishReport True: Set report = ActiveWorkbook
        Require report.Worksheets.Count = 6, "two snapshot pairs"
        Require report.Worksheets("비교요약").Range("B4").Value2 = 2, "missing sheet retained"
        GoTo Passed
    End If
    If name = "file_popup" Then
        a.Activate
        Set panel = New FNxWorkbookCompare: panel.BindFeature
        panel.Controls("txtBasePath").Value = pa: panel.Controls("txtComparePath").Value = pb
        panel.NxProbeLoadSheets
        Require panel.Controls("cboBaseSheet").ListCount = 2, "base sheet choices"
        Require panel.Controls("cboCompareSheet").ListCount = 2, "comparison sheet choices"
        panel.Controls("cboBaseSheet").ListIndex = 1: panel.Controls("cboCompareSheet").ListIndex = 1
        panel.NxProbeExecute
        Set report = ActiveWorkbook
        Require Not report Is a, "default new workbook execution"
        Require report.Worksheets(1).Range("B2").Value2 = 100, "popup compare execution"
        Unload panel: Set panel = Nothing
        GoTo Passed
    End If
    Set result = NxWorkbookCompareReport(pa, pb, "", True, "기준'자료", "비교 자료")
    NxWorkbookCompareFinishReport True
    Set report = ActiveWorkbook
    Set summary = report.Worksheets("비교요약"): Set detail = report.Worksheets("셀차이")
    Require summary.Range("B4").Value2 = 1, "difference count"
    Require report.Worksheets(1).Range("B2").Value2 = 100, "first baseline"
    Require report.Worksheets(2).Range("B2").Value2 = 120, "second comparison"
    Require report.Worksheets(1).Range("Q1").Value2 = "보존", "snapshot column Q preserved"
    Require report.Worksheets(1).Range("B2").Interior.Color = RGB(255, 199, 206), "highlight"
    Require detail.Range("D2").Value2 = "100" And detail.Range("E2").Value2 = "120", "plain values"
    Require detail.Columns("F:O").Hidden, "compact detail"
    detail.Range("E2").Hyperlinks(1).Follow
    Require ActiveSheet Is report.Worksheets(2), "snapshot link sheet"
    Require ActiveCell.Address = "$B$2", "snapshot link cell"
    If name = "reopen_links" Then
        report.SaveAs root & "\compare-report.xlsx", xlOpenXMLWorkbook
        report.Close False: Set report = Nothing
        a.Close False: Set a = Nothing
        b.Close False: Set b = Nothing
        Set report = Workbooks.Open(root & "\compare-report.xlsx", UpdateLinks:=0)
        Set detail = report.Worksheets("셀차이"): Set summary = report.Worksheets("비교요약")
        detail.Activate
        detail.Range("E2").Hyperlinks(1).Follow
        Require ActiveSheet Is report.Worksheets(2), "reopened snapshot link: actual=" & ActiveWorkbook.Name & "/" & ActiveSheet.Name & "/" & ActiveCell.Address & "; expected=" & report.Worksheets(2).Name & "; link=" & detail.Range("E2").Hyperlinks(1).Address & "#" & detail.Range("E2").Hyperlinks(1).SubAddress
        Require NxWorkbookCompareNavigate(summary, summary.Range("F6").Hyperlinks(1), issue), "original file link: " & issue
        Set b = ActiveWorkbook
        Require b.FullName = pb And ActiveSheet.Name = "비교 자료", "different-name original target"
        Require b.ReadOnly, "original read only"
        b.Close False: Set b = Nothing
        summary.Activate
        Application.EnableEvents = True
        summary.Range("F6").Hyperlinks(1).Follow
        Require ActiveWorkbook.FullName = pb And ActiveSheet.Name = "비교 자료", "original hyperlink event"
        Set b = ActiveWorkbook
        Require b.ReadOnly, "hyperlink event read only"
        Application.EnableEvents = False
    End If
    If name = "append_repeat" Then
        Set destination = Workbooks.Add(xlWBATWorksheet)
        destination.Worksheets(1).Name = "비교요약"
        destination.Worksheets(1).Range("A1").Value2 = "기존 자료"
        For i = 1 To 2
            originalCount = destination.Worksheets.Count
            NxCompareAppendResultSheets report, destination
            Set report = Nothing
            Require destination.Worksheets.Count = originalCount + 4, "append count"
            Require destination.Worksheets(1).Range("A1").Value2 = "기존 자료", "existing preserved"
            Set detail = destination.Worksheets(destination.Worksheets.Count)
            detail.Range("E2").Hyperlinks(1).Follow
            Require ActiveSheet Is destination.Worksheets(originalCount + 2), "appended link sheet"
            Require ActiveCell.Address = "$B$2", "appended link cell"
            detail.Activate
            Set explorer = New CNxCompareResultsSession
            explorer.Configure destination
            Require InStr(explorer.Navigate(1, 1), "이동했습니다") > 0, "appended result explorer"
            Require ActiveSheet Is destination.Worksheets(originalCount + 2), "explorer correct report"
            explorer.ReleaseSession: Set explorer = Nothing
            If i = 1 Then
                Set result = NxWorkbookCompareReport(pa, pb, "", False, "기준'자료", "비교 자료")
                NxWorkbookCompareFinishReport True: Set report = ActiveWorkbook
            End If
        Next i
    End If
Passed:
    RunCase = "PASS|" & name
    GoTo Clean
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Number & "|" & Err.Description
Clean:
    On Error Resume Next
    If Not panel Is Nothing Then Unload panel
    If Not report Is Nothing Then report.Close False
    If Not destination Is Nothing Then destination.Close False
    If Not a Is Nothing Then a.Close False
    If Not b Is Nothing Then b.Close False
    Application.EnableEvents = oldEvents: Application.DisplayAlerts = oldAlerts
End Function
Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise 5, "CompareOutput", detail
End Sub

Private Function CanonicalPath(ByVal path As String) As String
    ' CMD pushd maps UNC shares; Excel hyperlinks persist the UNC spelling.
    Dim drive As Object, share As String
    If Mid$(path, 2, 1) = ":" Then
        Set drive = CreateObject("Scripting.FileSystemObject").GetDrive(Left$(path, 2))
        If drive.DriveType = 3 Then
            share = drive.ShareName
            If Len(share) > 0 Then path = share & Mid$(path, 3)
        End If
    End If
    CanonicalPath = LCase$(path)
End Function

Public Sub OpenPanel()
    Set mPanelBook = Workbooks.Add(xlWBATWorksheet)
    Set mPanel = New FNxWorkbookCompare
    mPanel.BindFeature
    mPanel.Controls("txtBasePath").Value = Environ$("LHEXCEL_PROFILE_ROOT") & "\기준 selected_pair.xlsx"
    mPanel.Controls("txtComparePath").Value = Environ$("LHEXCEL_PROFILE_ROOT") & "\비교 selected_pair.xlsx"
    mPanel.NxProbeLoadSheets
    mPanel.Controls("cboBaseSheet").ListIndex = 1
    mPanel.Controls("cboCompareSheet").ListIndex = 1
    mPanel.Show vbModeless
End Sub
Public Function Caption() As String
    Caption = mPanel.Caption
End Function
Public Sub ClosePanel()
    Unload mPanel: Set mPanel = Nothing
    mPanelBook.Close False: Set mPanelBook = Nothing
End Sub
