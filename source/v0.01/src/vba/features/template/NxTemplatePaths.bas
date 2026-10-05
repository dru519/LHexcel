Attribute VB_Name = "NxTemplatePaths"
Option Explicit

Public Function NxTemplateStoreRoot() As String
    If Len(ThisWorkbook.Path) = 0 Then NxRaiseContractError "내엑셀 설치 폴더를 확인할 수 없습니다."
    NxTemplateStoreRoot = ThisWorkbook.Path & "\내엑셀 템플릿"
End Function

Public Sub NxTemplateEnsureStore(Optional ByVal rootOverride As String = vbNullString)
    Dim root As String
    root = EffectiveRoot(rootOverride)
    EnsureFolder root
    EnsureFolder root & "\Packages"
    EnsureFolder root & "\Staging"
    EnsureFolder root & "\Quarantine"
End Sub

Public Function NxTemplatePackagesRoot(Optional ByVal rootOverride As String = vbNullString) As String
    NxTemplatePackagesRoot = EffectiveRoot(rootOverride) & "\Packages"
End Function

Public Function NxTemplateStagingRoot(Optional ByVal rootOverride As String = vbNullString) As String
    NxTemplateStagingRoot = EffectiveRoot(rootOverride) & "\Staging"
End Function

Public Function NxTemplateTemplateFolderPath(ByVal templateId As String, Optional ByVal rootOverride As String = vbNullString) As String
    If Not NxTemplateIdIsValid(templateId) Then NxRaiseContractError "Template ID is outside the closed ASCII contract"
    NxTemplateTemplateFolderPath = NxTemplatePackagesRoot(rootOverride) & "\" & templateId
End Function

Public Function NxTemplatePackagePath(ByVal templateId As String, Optional ByVal rootOverride As String = vbNullString) As String
    NxTemplatePackagePath = NxTemplateTemplateFolderPath(templateId, rootOverride) & "\template.xlsx"
End Function

Public Function NxTemplateMetadataPath(ByVal templateId As String, Optional ByVal rootOverride As String = vbNullString) As String
    NxTemplateMetadataPath = NxTemplateTemplateFolderPath(templateId, rootOverride) & "\metadata.ini"
End Function

Public Function NxTemplateStagingPath(ByVal operationId As String, Optional ByVal rootOverride As String = vbNullString) As String
    If Not NxTemplateIdIsValid(operationId) Then NxRaiseContractError "Template operation ID is invalid"
    NxTemplateStagingPath = NxTemplateStagingRoot(rootOverride) & "\" & operationId
End Function

Public Function NxTemplateQuarantinePath(ByVal quarantineId As String, Optional ByVal rootOverride As String = vbNullString) As String
    If Not NxTemplateIdIsValid(quarantineId) Then NxRaiseContractError "Quarantine ID is outside the closed ASCII contract"
    NxTemplateQuarantinePath = EffectiveRoot(rootOverride) & "\Quarantine\" & quarantineId
End Function

Public Function NxTemplateIdIsValid(ByVal value As String) As Boolean
    Dim index As Long, character As String
    If Len(value) < 8 Or Len(value) > 64 Then Exit Function
    If Left$(value, 1) = "-" Or Right$(value, 1) = "-" Then Exit Function
    For index = 1 To Len(value)
        character = Mid$(value, index, 1)
        If InStr(1, "abcdefghijklmnopqrstuvwxyz0123456789-", character, vbBinaryCompare) = 0 Then Exit Function
    Next index
    NxTemplateIdIsValid = True
End Function

Private Function EffectiveRoot(ByVal rootOverride As String) As String
    If Len(rootOverride) = 0 Then
        EffectiveRoot = NxTemplateStoreRoot()
    Else
        If InStr(rootOverride, vbNullChar) > 0 Or Right$(rootOverride, 1) = "\" Then NxRaiseContractError "Template root override is invalid"
        EffectiveRoot = rootOverride
    End If
End Function

Private Sub EnsureFolder(ByVal path As String)
    Dim fileSystem As Object, parent As String
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem.FolderExists(path) Then Exit Sub
    parent = fileSystem.GetParentFolderName(path)
    If Len(parent) > 0 And Not fileSystem.FolderExists(parent) Then EnsureFolder parent
    fileSystem.CreateFolder path
End Sub
