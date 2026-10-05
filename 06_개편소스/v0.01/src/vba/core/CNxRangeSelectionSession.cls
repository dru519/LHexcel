VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxRangeSelectionSession"
Option Explicit

Private mWorkbookName As String
Private mWorkbookFullName As String
Private mWorksheetName As String
Private mWorksheetCodeName As String
Private mRangeAddress As String
Private mConfigured As Boolean

Public Sub UpdateFromRange(ByVal selected As Range)
    If selected Is Nothing Then NxRaiseContractError "선택 범위가 필요합니다."
    If selected.Areas.Count <> 1 Then NxRaiseContractError "연속된 한 범위만 선택하세요."
    mWorkbookName = selected.Worksheet.Parent.Name
    mWorkbookFullName = selected.Worksheet.Parent.FullName
    mWorksheetName = selected.Worksheet.Name
    mWorksheetCodeName = selected.Worksheet.CodeName
    mRangeAddress = selected.Address(True, True, xlA1, False)
    mConfigured = True
End Sub

Public Function ResolveRange() As Range
    Dim book As Workbook
    Dim sheet As Worksheet
    Dim candidateBook As Workbook
    Dim candidateSheet As Worksheet
    EnsureConfigured

    For Each candidateBook In Application.Workbooks
        If StrComp(candidateBook.FullName, mWorkbookFullName, vbTextCompare) = 0 And _
            StrComp(candidateBook.Name, mWorkbookName, vbTextCompare) = 0 Then
            Set book = candidateBook
            Exit For
        End If
    Next candidateBook
    If book Is Nothing Then NxRaiseContractError "선택했던 통합문서를 찾을 수 없습니다."
    If Application.ActiveWorkbook Is Nothing Then NxRaiseContractError "활성 통합문서가 없습니다."
    If Not Application.ActiveWorkbook Is book Then NxRaiseContractError "다른 통합문서가 활성화되어 실행을 중단했습니다."

    For Each candidateSheet In book.Worksheets
        If StrComp(candidateSheet.CodeName, mWorksheetCodeName, vbBinaryCompare) = 0 And _
            StrComp(candidateSheet.Name, mWorksheetName, vbBinaryCompare) = 0 Then
            Set sheet = candidateSheet
            Exit For
        End If
    Next candidateSheet
    If sheet Is Nothing Then NxRaiseContractError "선택했던 워크시트를 찾을 수 없습니다."
    If Not Application.ActiveSheet Is sheet Then NxRaiseContractError "다른 워크시트가 활성화되어 실행을 중단했습니다."

    On Error Resume Next
    Set ResolveRange = sheet.Range(mRangeAddress)
    On Error GoTo 0
    If ResolveRange Is Nothing Then NxRaiseContractError "선택 범위를 다시 확인하세요."
End Function

Public Property Get WorkbookIdentity() As String
    EnsureConfigured
    WorkbookIdentity = mWorkbookName & "|" & mWorkbookFullName
End Property

Public Property Get WorksheetIdentity() As String
    EnsureConfigured
    WorksheetIdentity = mWorksheetCodeName & "|" & mWorksheetName
End Property

Public Property Get RangeAddress() As String
    EnsureConfigured
    RangeAddress = mRangeAddress
End Property

Public Property Get DisplayAddress() As String
    EnsureConfigured
    DisplayAddress = "'" & Replace$(mWorksheetName, "'", "''") & "'!" & mRangeAddress
End Property

Public Property Get IsConfigured() As Boolean
    IsConfigured = mConfigured
End Property

Private Sub EnsureConfigured()
    If Not mConfigured Then NxRaiseContractError "범위 선택 세션이 준비되지 않았습니다."
End Sub
