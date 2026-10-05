Attribute VB_Name = "NxPdfPolicy"
Option Explicit

Public Sub NxPdfValidateFeature(ByVal featureId As String)
    Select Case featureId
        Case "NX-FILE-PDF-CURRENT-SHEET", "NX-FILE-PDF-EACH-SHEET", _
             "NX-FILE-PDF-SELECTED-COMBINED", "NX-FILE-PDF-ALL-COMBINED", _
             "NX-FILE-PDF-SETTINGS"
        Case Else: NxRaiseContractError "지원하지 않는 PDF 작업입니다."
    End Select
End Sub

Public Function NxPdfDefaultFolder() As String
    If Not ActiveWorkbook Is Nothing Then
        If Len(ActiveWorkbook.Path) > 0 Then NxPdfDefaultFolder = ActiveWorkbook.Path: Exit Function
    End If
    NxPdfDefaultFolder = Environ$("USERPROFILE") & Application.PathSeparator & "Documents"
End Function

Public Function NxPdfDefaultBaseName(Optional ByVal featureId As String = vbNullString) As String
    Dim candidate As String, dotAt As Long
    If featureId = NX_FEATURE_FILE_PDF_CURRENT_SHEET And TypeName(ActiveSheet) = "Worksheet" Then
        NxPdfDefaultBaseName = NxPdfSafeName(ActiveSheet.Name)
        Exit Function
    ElseIf ActiveWorkbook Is Nothing Then
        candidate = "내엑셀_PDF"
    Else
        candidate = ActiveWorkbook.Name
        dotAt = InStrRev(candidate, ".")
        If dotAt > 1 Then candidate = Left$(candidate, dotAt - 1)
    End If
    NxPdfDefaultBaseName = NxPdfSafeName(candidate)
End Function

Public Function NxPdfQualityValue(ByVal qualityText As String) As Long
    If InStr(1, qualityText, "작은", vbTextCompare) > 0 Then _
        NxPdfQualityValue = xlQualityMinimum Else NxPdfQualityValue = xlQualityStandard
End Function

Public Function NxPdfOutputPath(ByVal outputFolder As String, ByVal baseName As String, _
    Optional ByVal sheetName As String = vbNullString) As String

    Dim fileName As String
    If Len(Dir$(outputFolder, vbDirectory)) = 0 Then NxRaiseContractError "PDF 저장 폴더를 찾을 수 없습니다."
    fileName = NxPdfSafeName(baseName)
    If Len(sheetName) > 0 Then fileName = fileName & "_" & NxPdfSafeName(sheetName)
    NxPdfOutputPath = outputFolder & Application.PathSeparator & fileName & ".pdf"
End Function

Public Sub NxPdfRequireVacantOutput(ByVal outputPath As String)
    If LCase$(Right$(outputPath, 4)) <> ".pdf" Then NxRaiseContractError "PDF 출력 경로를 확인하세요."
    If Len(Dir$(outputPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0 Then _
        NxRaiseContractError "같은 이름의 PDF가 이미 있습니다. 다른 이름을 사용하세요."
End Sub

Public Function NxPdfSettingsPath() As String
    Dim root As String, settingsFolder As String
    root = NxLHexcelProfileRoot()
    If Len(Dir$(root, vbDirectory)) = 0 Then MkDir root
    settingsFolder = root & Application.PathSeparator & "Settings"
    If Len(Dir$(settingsFolder, vbDirectory)) = 0 Then MkDir settingsFolder
    NxPdfSettingsPath = settingsFolder & Application.PathSeparator & "pdf-v1.cfg"
End Function

Public Sub NxPdfSaveSettings(ByVal outputFolder As String, ByVal baseName As String, ByVal qualityText As String)
    Dim finalPath As String, temporaryPath As String, backupPath As String, handle As Integer
    Dim hadPrior As Boolean, errorNumber As Long, errorDescription As String
    If Len(Dir$(outputFolder, vbDirectory)) = 0 Then NxRaiseContractError "PDF 기본 저장 폴더를 찾을 수 없습니다."
    If Len(NxPdfSafeName(baseName)) = 0 Then NxRaiseContractError "PDF 기본 파일명을 입력하세요."
    Call NxPdfQualityValue(qualityText)
    finalPath = NxPdfSettingsPath()
    temporaryPath = finalPath & "." & Replace$(NxCreateRunUuid(), "-", vbNullString) & ".tmp"
    backupPath = finalPath & ".bak"
    On Error GoTo Failed
    handle = FreeFile
    Open temporaryPath For Output Access Write As #handle
    Print #handle, "folder=" & outputFolder
    Print #handle, "base=" & baseName
    Print #handle, "quality=" & qualityText
    Close #handle
    handle = 0
    If Len(Dir$(backupPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0 Then Kill backupPath
    hadPrior = (Len(Dir$(finalPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0)
    If hadPrior Then Name finalPath As backupPath
    Name temporaryPath As finalPath
    If hadPrior And Len(Dir$(backupPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0 Then Kill backupPath
    Exit Sub
Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear
    On Error Resume Next
    If handle <> 0 Then Close #handle
    If Len(Dir$(temporaryPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0 Then Kill temporaryPath
    If hadPrior And Len(Dir$(backupPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0 And _
        Len(Dir$(finalPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) = 0 Then Name backupPath As finalPath
    On Error GoTo 0
    Err.Raise errorNumber, "LHexcel.File.PdfSettings", errorDescription
End Sub

Private Function NxPdfSafeName(ByVal value As String) As String
    Dim result As String, forbidden As Variant, token As Variant
    result = Trim$(value)
    forbidden = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    For Each token In forbidden
        result = Replace$(result, CStr(token), "_")
    Next token
    Do While Right$(result, 1) = "." Or Right$(result, 1) = " "
        result = Left$(result, Len(result) - 1)
        If Len(result) = 0 Then Exit Do
    Loop
    If Len(result) = 0 Then NxRaiseContractError "PDF 파일명을 입력하세요."
    If Len(result) > 120 Then result = Left$(result, 120)
    NxPdfSafeName = result
End Function
