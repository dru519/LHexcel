Attribute VB_Name = "NxBatchRenameController"
Option Explicit

Public Function NxBatchRenamePreview(ByVal pathsText As String, ByVal findText As String, _
    ByVal replaceText As String, ByVal prefixText As String, ByVal suffixText As String, _
    ByRef HasDuplicateTarget As Boolean, ByRef HasExistingCollision As Boolean) As Collection

    Set NxBatchRenamePreview = NxBatchRenameBuildPreview(pathsText, findText, replaceText, prefixText, suffixText, _
        HasDuplicateTarget, HasExistingCollision)
End Function

Public Function NxSheetBatchRenamePreview(ByVal findText As String, ByVal replaceText As String, _
    ByVal prefixText As String, ByVal suffixText As String, ByRef HasDuplicateTarget As Boolean, _
    ByRef HasExistingCollision As Boolean) As Collection

    If ActiveWorkbook Is Nothing Then NxRaiseContractError "열린 통합문서가 없습니다."
    Set NxSheetBatchRenamePreview = NxSheetBatchRenameBuildPreview(ActiveWorkbook, findText, replaceText, _
        prefixText, suffixText, HasDuplicateTarget, HasExistingCollision)
End Function

Public Function NxBatchRenameRootTarget(ByVal preview As Collection) As String
    Dim row As Variant, separator As Long
    If preview Is Nothing Then NxRaiseContractError "이름 변경 미리보기가 필요합니다."
    For Each row In preview
        If CStr(row(2)) = "준비" Then
            separator = InStrRev(CStr(row(0)), Application.PathSeparator)
            If separator = 0 Then NxRaiseContractError "파일 경로를 확인하세요."
            NxBatchRenameRootTarget = Left$(CStr(row(0)), separator - 1)
            Exit Function
        End If
    Next row
    NxRaiseContractError "변경할 파일 이름이 없습니다."
End Function
