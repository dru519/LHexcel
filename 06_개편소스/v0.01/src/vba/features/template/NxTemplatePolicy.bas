Attribute VB_Name = "NxTemplatePolicy"
Option Explicit

Public Function NxTemplateAnalyze(ByVal source As Range, ByVal retained As Collection, _
    ByVal includeHidden As Boolean, ByVal exclusionAccepted As Boolean) As CNxTemplateAnalysis

    If source Is Nothing Or retained Is Nothing Then NxRaiseContractError "Template source and retained ranges are required"
    If source.Areas.Count <> 1 Then NxRaiseContractError "Template source must be one rectangular range"
    ValidateRetainedAreas source, retained
    Dim result As New CNxTemplateAnalysis
    result.Configure source, includeHidden
    InspectCells source, retained, result
    InspectResources source.Worksheet, source, result
    result.Seal
    If result.HasBlockedResource Then NxRaiseContractError result.BlockReason
    If result.HasExcludedDisplayResource And Not exclusionAccepted Then NxRaiseContractError "Unsupported display resource exclusion requires confirmation"
    Set NxTemplateAnalyze = result
End Function

Public Function NxTemplateFormulaDisposition(ByVal formulaCell As Range, ByVal source As Range) As String
    Dim formulaText As String
    If formulaCell Is Nothing Or source Is Nothing Then NxRaiseContractError "Template formula policy requires cell and source"
    If formulaCell.Cells.CountLarge <> 1 Or Not formulaCell.HasFormula Then NxRaiseContractError "Template formula policy requires one formula cell"
    formulaText = UCase$(CStr(formulaCell.FormulaR1C1))
    If HasHardBlockedFormulaToken(formulaText) Or FormulaUsesDefinedName(formulaCell, formulaText) Then
        NxTemplateFormulaDisposition = "block:ExternalReference"
    ElseIf FormulaIsArrayOrSpill(formulaCell, formulaText) Then
        NxTemplateFormulaDisposition = "block:ArrayFormula"
    ElseIf HasFunction(formulaText, Array("INDIRECT", "OFFSET")) Then
        NxTemplateFormulaDisposition = "warn:DynamicFormula"
    ElseIf HasFunction(formulaText, Array("CELL", "INFO")) Then
        NxTemplateFormulaDisposition = "warn:EnvironmentFormula"
    ElseIf HasFunction(formulaText, Array("NOW", "TODAY", "RAND", "RANDBETWEEN")) Then
        NxTemplateFormulaDisposition = "warn:VolatileFormula"
    ElseIf FormulaPrecedentsLeaveSource(formulaCell, source) Then
        NxTemplateFormulaDisposition = "block:OutOfBoundsReference"
    Else
        NxTemplateFormulaDisposition = "allow"
    End If
End Function

Public Sub ValidateRetainedAreas(ByVal source As Range, ByVal retained As Collection)
    Dim area As Range, overlap As Range
    If source Is Nothing Or retained Is Nothing Then NxRaiseContractError "Template retained-area validation requires source and areas"
    For Each area In retained
        If Not area.Worksheet Is source.Worksheet Then NxRaiseContractError "Retained area must use the source worksheet"
        Set overlap = Application.Intersect(source, area)
        If overlap Is Nothing Or CDbl(overlap.CountLarge) <> CDbl(area.CountLarge) Then NxRaiseContractError "Retained area must stay inside the template source"
        Set overlap = Nothing
    Next area
End Sub

Private Sub InspectCells(ByVal source As Range, ByVal retained As Collection, ByVal result As CNxTemplateAnalysis)
    Dim cell As Range, disposition As String
    Dim completed As Long
    For Each cell In source.Cells
        If completed Mod 256 = 0 Then NxTemplateProgress "등록 내용 검사", completed, CLng(source.CountLarge)
        If cell.HasFormula Then
            result.CountFormula
            disposition = NxTemplateFormulaDisposition(cell, source)
            If Left$(disposition, 6) = "block:" Then
                result.MarkBlocked cell.Address(False, False, xlA1, False) & "|" & disposition
            Else
                AddFormulaRisks cell, result
            End If
        ElseIf Not IsEmpty(cell.Value2) Then
            If CellIsRetained(cell, retained) Then result.CountStoredConstant Else result.CountExcludedConstant
        End If
        completed = completed + 1
    Next cell
End Sub

Private Sub AddFormulaRisks(ByVal cell As Range, ByVal result As CNxTemplateAnalysis)
    Dim formulaText As String, name As Variant
    formulaText = UCase$(CStr(cell.FormulaR1C1))
    For Each name In Array("NOW", "TODAY", "RAND", "RANDBETWEEN")
        If ContainsFunction(formulaText, CStr(name)) Then result.AddRisk "VolatileFormula", cell.Address(False, False, xlA1, False), CStr(name)
    Next name
    For Each name In Array("INDIRECT", "OFFSET")
        If ContainsFunction(formulaText, CStr(name)) Then result.AddRisk "DynamicFormula", cell.Address(False, False, xlA1, False), CStr(name)
    Next name
    For Each name In Array("CELL", "INFO")
        If ContainsFunction(formulaText, CStr(name)) Then result.AddRisk "EnvironmentFormula", cell.Address(False, False, xlA1, False), CStr(name)
    Next name
End Sub

Private Function CellIsRetained(ByVal cell As Range, ByVal retained As Collection) As Boolean
    Dim area As Range, overlap As Range
    For Each area In retained
        Set overlap = Application.Intersect(cell, area)
        If Not overlap Is Nothing Then CellIsRetained = True: Exit Function
    Next area
End Function

Private Function HasHardBlockedFormulaToken(ByVal formulaText As String) As Boolean
    Dim token As Variant
    For Each token In Array("!", "HTTP://", "HTTPS://", "\\", "WEBSERVICE(", "RTD(", "|")
        If InStr(1, formulaText, CStr(token), vbBinaryCompare) > 0 Then HasHardBlockedFormulaToken = True: Exit Function
    Next token
End Function

Private Function HasFunction(ByVal formulaText As String, ByVal names As Variant) As Boolean
    Dim name As Variant
    For Each name In names
        If ContainsFunction(formulaText, CStr(name)) Then HasFunction = True: Exit Function
    Next name
End Function

Private Function ContainsFunction(ByVal formulaText As String, ByVal functionName As String) As Boolean
    ContainsFunction = (InStr(1, formulaText, functionName & "(", vbBinaryCompare) > 0)
End Function

Private Function FormulaUsesDefinedName(ByVal formulaCell As Range, ByVal formulaText As String) As Boolean
    Dim item As Name, shortName As String
    On Error GoTo Failed
    For Each item In formulaCell.Worksheet.Parent.Names
        shortName = UCase$(CStr(item.Name))
        If InStrRev(shortName, "!") > 0 Then shortName = Mid$(shortName, InStrRev(shortName, "!") + 1)
        If Len(shortName) > 0 And InStr(1, formulaText, shortName, vbBinaryCompare) > 0 Then FormulaUsesDefinedName = True: Exit Function
    Next item
    Exit Function
Failed:
    FormulaUsesDefinedName = True
    Err.Clear
End Function

Private Function FormulaIsArrayOrSpill(ByVal formulaCell As Range, ByVal formulaText As String) As Boolean
    On Error Resume Next
    FormulaIsArrayOrSpill = CBool(formulaCell.HasArray)
    On Error GoTo 0
    If InStr(1, formulaText, "#", vbBinaryCompare) > 0 Then FormulaIsArrayOrSpill = True
End Function

Private Function FormulaPrecedentsLeaveSource(ByVal formulaCell As Range, ByVal source As Range) As Boolean
    Dim precedents As Range, overlap As Range
    On Error Resume Next
    Set precedents = formulaCell.DirectPrecedents
    Err.Clear
    On Error GoTo 0
    If precedents Is Nothing Then Exit Function
    Set overlap = Application.Intersect(precedents, source)
    If overlap Is Nothing Then
        FormulaPrecedentsLeaveSource = True
    Else
        FormulaPrecedentsLeaveSource = (CDbl(overlap.CountLarge) <> CDbl(precedents.CountLarge))
    End If
End Function

Private Sub InspectResources(ByVal sheet As Worksheet, ByVal source As Range, ByVal result As CNxTemplateAnalysis)
    Dim shape As Shape, formatCondition As Object
    If sheet.Parent.HasVBProject Then result.MarkBlocked "VBA project is not allowed in a template package"
    If sheet.QueryTables.Count > 0 Or sheet.PivotTables.Count > 0 Then result.MarkBlocked "External data or pivot resources are not allowed"
    On Error Resume Next
    If sheet.Parent.Connections.Count > 0 Then result.MarkBlocked "Workbook connections are not allowed"
    On Error GoTo 0
    For Each shape In sheet.Shapes
        If ShapeTouchesSource(shape, source) Then
            Select Case shape.Type
                Case msoEmbeddedOLEObject, msoLinkedOLEObject, msoOLEControlObject, msoFormControl
                    result.MarkBlocked "Executable or linked sheet object is not allowed"
                Case Else
                    result.MarkExcludedDisplayResource "shape"
            End Select
        End If
    Next shape
    If source.Hyperlinks.Count > 0 Then result.MarkExcludedDisplayResource "hyperlink"
    For Each formatCondition In source.FormatConditions
        If formatCondition.Type <> xlCellValue And formatCondition.Type <> xlExpression Then result.MarkExcludedDisplayResource "conditional-format"
    Next formatCondition
End Sub

Private Function ShapeTouchesSource(ByVal shape As Shape, ByVal source As Range) As Boolean
    Dim shapeCells As Range, overlap As Range
    On Error GoTo Conservative
    Set shapeCells = source.Worksheet.Range(shape.TopLeftCell, shape.BottomRightCell)
    Set overlap = Application.Intersect(shapeCells, source)
    ShapeTouchesSource = Not overlap Is Nothing
    Exit Function
Conservative:
    ShapeTouchesSource = True
    Err.Clear
End Function
