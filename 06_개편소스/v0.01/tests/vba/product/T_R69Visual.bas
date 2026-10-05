Attribute VB_Name = "T_R69Visual"
Option Explicit

Public Function Names() As String
    Names = "visual"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim target As Range
    Set target = ActiveWorkbook.Worksheets(1).Range("A1:D6")
    target.Value2 = "표본"
    target.Font.Name = "맑은 고딕"
    target.Font.Size = 9
    target.Columns.ColumnWidth = 18
    target.RowHeight = 30
    target.Range("A1:B1").ClearContents
    target.Range("A1:B1").Merge
    target.Cells(1, 1).Value2 = "병합 머리글"
    target.Cells(2, 1).Value2 = "여러 줄의 긴 설명을 보여주는 미리보기 표본입니다."
    target.Cells(2, 1).WrapText = True
    target.Rows(2).RowHeight = 45
    target.Cells(3, 1).Value2 = "첫째 줄" & vbLf & "둘째 줄"
    target.Cells(3, 1).WrapText = True
    target.Cells(4, 2).Formula = "=10+20"
    target.Interior.Color = RGB(255, 230, 128)
    target.Borders.LineStyle = xlContinuous
    target.Select
    ActiveWorkbook.Save
    RunCase = "PASS|" & name
End Function
