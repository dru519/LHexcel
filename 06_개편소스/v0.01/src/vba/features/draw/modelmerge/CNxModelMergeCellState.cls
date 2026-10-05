VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxModelMergeCellState"
Option Explicit

Private mWorkbookFullName As String
Private mWorksheetCodeName As String
Private mRangeAddress As String
Private mHasFormula As Boolean
Private mUsesFormula2 As Boolean
Private mFormulaValue As Variant
Private mValue2 As Variant
Private mNumberFormat As Variant
Private mHorizontalAlignment As Variant
Private mShrinkToFit As Variant

Public Sub Capture(ByVal cell As Range)
    If cell Is Nothing Then NxRaiseContractError "보존할 셀이 필요합니다."
    If cell.Cells.CountLarge <> 1 Then NxRaiseContractError "셀 상태는 한 셀씩 보존해야 합니다."

    mWorkbookFullName = cell.Worksheet.Parent.FullName
    mWorksheetCodeName = cell.Worksheet.CodeName
    mRangeAddress = cell.Address(True, True, xlA1, False)
    mHasFormula = CBool(cell.HasFormula)
    If mHasFormula Then CaptureFormula cell Else mValue2 = cell.Value2
    mNumberFormat = cell.NumberFormat
    mHorizontalAlignment = cell.HorizontalAlignment
    mShrinkToFit = cell.ShrinkToFit
End Sub

Public Sub Restore()
    Dim cell As Range
    Set cell = ResolveCell()
    RestoreTo cell
End Sub

Public Sub RestoreTo(ByVal cell As Range)
    If cell Is Nothing Then NxRaiseContractError "복구할 셀을 찾을 수 없습니다."
    If mHasFormula Then
        If mUsesFormula2 Then
            On Error Resume Next
            cell.Formula2 = mFormulaValue
            If Err.Number = 0 Then
                On Error GoTo 0
            Else
                Err.Clear
                On Error GoTo 0
                cell.Formula = mFormulaValue
            End If
        Else
            cell.Formula = mFormulaValue
        End If
    Else
        cell.Value2 = mValue2
    End If
    cell.NumberFormat = mNumberFormat
    cell.HorizontalAlignment = mHorizontalAlignment
    cell.ShrinkToFit = mShrinkToFit
End Sub

Public Property Get WorkbookFullName() As String
    WorkbookFullName = mWorkbookFullName
End Property

Public Property Get WorksheetCodeName() As String
    WorksheetCodeName = mWorksheetCodeName
End Property

Public Property Get RangeAddress() As String
    RangeAddress = mRangeAddress
End Property

Private Sub CaptureFormula(ByVal cell As Range)
    On Error Resume Next
    Err.Clear
    mFormulaValue = cell.Formula2
    mUsesFormula2 = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    If Not mUsesFormula2 Then mFormulaValue = cell.Formula
End Sub

Private Function ResolveCell() As Range
    Dim book As Workbook
    Dim sheet As Worksheet
    Dim candidateBook As Workbook
    Dim candidateSheet As Worksheet

    For Each candidateBook In Application.Workbooks
        If StrComp(candidateBook.FullName, mWorkbookFullName, vbTextCompare) = 0 Then
            Set book = candidateBook
            Exit For
        End If
    Next candidateBook
    If book Is Nothing Then NxRaiseContractError "원본 통합문서를 찾을 수 없습니다."

    For Each candidateSheet In book.Worksheets
        If StrComp(candidateSheet.CodeName, mWorksheetCodeName, vbBinaryCompare) = 0 Then
            Set sheet = candidateSheet
            Exit For
        End If
    Next candidateSheet
    If sheet Is Nothing Then NxRaiseContractError "원본 워크시트를 찾을 수 없습니다."

    On Error Resume Next
    Set ResolveCell = sheet.Range(mRangeAddress)
    On Error GoTo 0
    If ResolveCell Is Nothing Then NxRaiseContractError "원본 셀 주소를 찾을 수 없습니다."
End Function
