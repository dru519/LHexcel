Attribute VB_Name = "T_ProductExhaustive"
Option Explicit

Public Function T_ProductExhaustive_CreateFixture(ByVal fixturePath As String) As Excel.Workbook
    Dim wb As Excel.Workbook
    Dim ws As Excel.Worksheet
    If Len(fixturePath) = 0 Then Err.Raise vbObjectError + 1970, "T_ProductExhaustive", "fixture path is required"
    Set wb = Application.Workbooks.Add
    Do While wb.Worksheets.Count < 3
        wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)
    Loop
    Set ws = wb.Worksheets(1)
    ws.Name = "Data"
    wb.Worksheets(2).Name = "Second"
    wb.Worksheets(3).Name = "Last"
    ws.Range("A1:D1").Value = Array("Category", "Value", "Date", "Text")
    ws.Range("A2:D2").Value = Array("A", 10, DateSerial(1990, 1, 1), "alpha")
    ws.Range("A3:D3").Value = Array("B", 20, DateSerial(2000, 2, 2), "beta")
    ws.Range("A4:D4").Value = Array("A", 30, DateSerial(2010, 3, 3), "alpha")
    ws.Range("A5:D5").Value = Array("C", 40, DateSerial(2020, 4, 4), "gamma")
    ws.Range("E2").Formula = "=B2*2"
    ws.Range("E3").Formula = "=B3*2"
    ws.Range("E4").Formula = "=SUM(B2:B3)"
    ws.Range("B2:D4").Select
    Application.DisplayAlerts = False
    wb.SaveAs fixturePath, xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    Set T_ProductExhaustive_CreateFixture = wb
End Function

Public Function T_ProductExhaustive_Run(ByVal routeId As String, ByVal fixturePath As String) As String
    Dim wb As Excel.Workbook
    On Error GoTo Failed
    If Len(routeId) = 0 Then Err.Raise vbObjectError + 1971, "T_ProductExhaustive", "route id is required"
    Set wb = Application.Workbooks.Open(fixturePath, ReadOnly:=False)
    wb.Worksheets(1).Activate
    wb.Worksheets(1).Range("B2:D4").Select
    If Left$(routeId, 7) = "NX-CMD-" Then
        NxRouteCommand routeId
    Else
        NxRouteFeature routeId
    End If
    T_ProductExhaustive_Run = "PASS|" & routeId
CleanUp:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Exit Function
Failed:
    T_ProductExhaustive_Run = "FAIL|" & routeId & "|" & Err.Source & "|" & Err.Description
    Resume CleanUp
End Function
