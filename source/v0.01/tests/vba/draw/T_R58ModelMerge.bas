Attribute VB_Name = "T_R58ModelMerge"
Option Explicit

' Direct native entrypoint for an instrumented Product.xlam copy.
Public Sub RunAll()
    TestCenterAcrossPreservesFirstCellFormula
    TestCenterAcrossRejectsRelocation
    TestRightAlignmentPreservesSource
    TestRepeatStillCopiesByOptIn
End Sub

Private Sub TestCenterAcrossPreservesFirstCellFormula()
    Dim book As Workbook, target As Range, beforeAlignment As Variant, detail As String
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("A1:C1")
    target.Cells(1, 1).Formula = "=1+1"
    beforeAlignment = target.Cells(1, 1).HorizontalAlignment
    NxModelMergeApply target, "CENTER", False
    T_R58ModelMergeAssert target.Cells(1, 1).Formula = "=1+1", "Center-across moved the source formula"
    T_R58ModelMergeAssert IsEmpty(target.Cells(1, 2).Value2) And IsEmpty(target.Cells(1, 3).Value2), "Center-across populated an adjacent cell"
    T_R58ModelMergeAssert target.HorizontalAlignment = xlCenterAcrossSelection, "Center-across alignment was not applied"
    NxModelMergeRestore target
    T_R58ModelMergeAssert target.Cells(1, 1).HorizontalAlignment = beforeAlignment, "Restore did not recover source alignment"
CleanUp:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description
    Resume CleanUp
End Sub

Private Sub TestCenterAcrossRejectsRelocation()
    Dim book As Workbook, target As Range, failed As Boolean, detail As String
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("A1:C1")
    target.Cells(1, 2).Formula = "=2+2"
    On Error Resume Next
    NxModelMergeApply target, "CENTER", False
    failed = (Err.Number <> 0)
    Err.Clear
    On Error GoTo Failed
    T_R58ModelMergeAssert failed, "Center-across accepted a non-first source by relocating it"
    T_R58ModelMergeAssert target.Cells(1, 2).Formula = "=2+2", "Rejected center-across changed the source formula"
    T_R58ModelMergeAssert IsEmpty(target.Cells(1, 1).Value2) And IsEmpty(target.Cells(1, 3).Value2), "Rejected center-across changed an adjacent cell"
CleanUp:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description
    Resume CleanUp
End Sub

Private Sub TestRightAlignmentPreservesSource()
    Dim book As Workbook, target As Range, detail As String
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("A1:C1")
    target.Cells(1, 2).Value2 = "keep-right"
    NxModelMergeApply target, "RIGHT", False
    T_R58ModelMergeAssert CStr(target.Cells(1, 2).Value2) = "keep-right", "Right alignment relocated the source value"
    T_R58ModelMergeAssert IsEmpty(target.Cells(1, 1).Value2) And IsEmpty(target.Cells(1, 3).Value2), "Right alignment populated an adjacent cell"
    T_R58ModelMergeAssert target.HorizontalAlignment = xlRight, "Right alignment was not applied"
    NxModelMergeRestore target
CleanUp:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description
    Resume CleanUp
End Sub

Private Sub TestRepeatStillCopiesByOptIn()
    Dim book As Workbook, target As Range, detail As String
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set target = book.Worksheets(1).Range("A1:C1")
    target.Cells(1, 1).Formula = "=3+3"
    NxModelMergeApply target, "CENTER", True
    T_R58ModelMergeAssert target.Cells(1, 1).Formula = "=3+3" And target.Cells(1, 2).Formula = "=3+3" And target.Cells(1, 3).Formula = "=3+3", "Repeat-value option did not copy the formula"
    T_R58ModelMergeAssert target.Cells(1, 2).NumberFormat <> ";;;", "Repeat-value visible anchor is hidden"
    T_R58ModelMergeAssert target.Cells(1, 1).NumberFormat = ";;;" And target.Cells(1, 3).NumberFormat = ";;;", "Repeat-value duplicates are not hidden"
    NxModelMergeRestore target
    T_R58ModelMergeAssert target.Cells(1, 1).Formula = "=3+3", "Repeat-value restore changed the original formula"
    T_R58ModelMergeAssert IsEmpty(target.Cells(1, 2).Value2) And IsEmpty(target.Cells(1, 3).Value2), "Repeat-value restore did not clear copied cells"
CleanUp:
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    If Len(detail) > 0 Then NxRaiseContractError detail
    Exit Sub
Failed:
    detail = Err.Description
    Resume CleanUp
End Sub

Private Sub T_R58ModelMergeAssert(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then NxRaiseContractError message
End Sub
