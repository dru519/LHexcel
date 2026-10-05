VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxModelMergeSnapshot"
Option Explicit

Private mWorkbookFullName As String
Private mWorksheetCodeName As String
Private mRangeAddress As String
Private mStates As Collection

Public Sub Capture(ByVal target As Range)
    Dim cell As Range
    Dim state As CNxModelMergeCellState
    If target Is Nothing Then NxRaiseContractError "보존할 범위가 필요합니다."
    If target.Areas.Count <> 1 Then NxRaiseContractError "연속된 한 범위만 보존할 수 있습니다."

    mWorkbookFullName = target.Worksheet.Parent.FullName
    mWorksheetCodeName = target.Worksheet.CodeName
    mRangeAddress = target.Address(True, True, xlA1, False)
    Set mStates = New Collection
    For Each cell In target.Cells
        Set state = New CNxModelMergeCellState
        state.Capture cell
        mStates.Add state
    Next cell
End Sub

Public Sub Restore()
    Dim i As Long
    Dim state As CNxModelMergeCellState
    Dim target As Range
    Set target = ResolveRange()
    If mStates Is Nothing Then NxRaiseContractError "복구할 원본 상태가 없습니다."
    If target.Cells.CountLarge <> mStates.Count Then NxRaiseContractError "복구 대상의 크기가 원본과 다릅니다."

    For i = mStates.Count To 1 Step -1
        Set state = mStates(i)
        state.RestoreTo target.Cells(i)
    Next i
End Sub

Public Function ResolveRange() As Range
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
    Set ResolveRange = sheet.Range(mRangeAddress)
    On Error GoTo 0
    If ResolveRange Is Nothing Then NxRaiseContractError "원본 범위를 찾을 수 없습니다."
End Function

Public Property Get WorkbookFullName() As String
    WorkbookFullName = mWorkbookFullName
End Property

Public Property Get WorksheetCodeName() As String
    WorksheetCodeName = mWorksheetCodeName
End Property

Public Property Get RangeAddress() As String
    RangeAddress = mRangeAddress
End Property

Public Property Get IdentityKey() As String
    IdentityKey = mWorkbookFullName & "|" & mWorksheetCodeName & "|" & mRangeAddress
End Property
