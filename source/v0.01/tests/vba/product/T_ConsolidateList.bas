Attribute VB_Name = "T_ConsolidateList"
Option Explicit

Public Function Names() As String
    Names = "selection_order|over20_rows|over20_sheets"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim panel As FNxFileConsolidate, book As Workbook, checkBook As Workbook
    Dim paths(0 To 24) As String, index As Long, root As String, output As String
    Dim result As CNxResult, planned As Collection, mode As String
    On Error GoTo Failed
    If name = "selection_order" Then
        Set panel = New FNxFileConsolidate
        Require panel.lstInputFiles.MultiSelect = fmMultiSelectExtended, "Ctrl/Shift selection"
        For index = 0 To 4: panel.lstInputFiles.AddItem Chr$(65 + index): Next index
        panel.lstInputFiles.Selected(1) = True: panel.lstInputFiles.Selected(3) = True
        panel.cmdMoveUp.Value = True
        Require CStr(panel.lstInputFiles.List(0)) = "B" And CStr(panel.lstInputFiles.List(2)) = "D", "noncontiguous up"
        Require panel.lstInputFiles.Selected(0) And panel.lstInputFiles.Selected(2), "selection follows files"
        panel.cmdMoveDown.Value = True
        For index = 0 To 4: Require panel.lstInputFiles.List(index) = Chr$(65 + index), "down restores order": Next index
        panel.cmdRemoveFile.Value = True
        Require panel.lstInputFiles.ListCount = 3, "remove all selected"
        Require panel.lstInputFiles.List(0) = "A" And panel.lstInputFiles.List(1) = "C" And panel.lstInputFiles.List(2) = "E", "survivor order"
        panel.cmdSelectAll.Value = True
        For index = 0 To 2: Require panel.lstInputFiles.Selected(index), "select all": Next index
        panel.cmdMoveUp.Value = True: panel.cmdMoveDown.Value = True
        Require panel.lstInputFiles.List(0) = "A" And panel.lstInputFiles.List(2) = "E", "all selected boundary"
        panel.cmdRemoveFile.Value = True
        Require panel.lstInputFiles.ListCount = 0, "remove all"
        Unload panel
    Else
        root = Environ$("LHEXCEL_PROFILE_ROOT") & "\" & name
        MkDir root
        Set book = Workbooks.Add(xlWBATWorksheet)
        For index = 0 To 24
            paths(index) = root & "\input" & Format$(index, "00") & ".xlsx"
            book.Worksheets(1).Range("A1").Value2 = index
            book.SaveAs paths(index), xlOpenXMLWorkbook
        Next index
        book.Close False: Set book = Nothing
        Set planned = NxConsolidationPlan(paths)
        Require planned.Count = 25, "more than 20 inputs"
        output = root & "\output.xlsx"
        mode = IIf(name = "over20_rows", "ROWS", "SHEETS")
        Set result = NxRunConsolidation(paths, output, False, False, mode)
        Require result.Outcome = NxSuccess, result.Recovery
        Set checkBook = Workbooks.Open(output, 0, True)
        For index = 0 To 24
            If mode = "ROWS" Then
                Require CStr(checkBook.Worksheets(1).Cells(index + 1, 1).Value2) = CStr(index), "row order"
            Else
                Require CStr(checkBook.Worksheets(index + 1).Range("A1").Value2) = CStr(index), "sheet order"
            End If
            Require Len(Dir$(paths(index))) > 0, "source preserved"
        Next index
        checkBook.Close False: Set checkBook = Nothing
    End If
    RunCase = "PASS|" & name
    Exit Function
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Number & ":" & Err.Description
    On Error Resume Next
    If Not panel Is Nothing Then Unload panel
    If Not book Is Nothing Then book.Close False
    If Not checkBook Is Nothing Then checkBook.Close False
End Function

Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise vbObjectError + 104, , detail
End Sub
