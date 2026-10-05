Attribute VB_Name = "NxPng"
Option Explicit

Public Function NxPngExportChart(ByVal chartObject As Object, ByVal outputPath As String, Optional ByVal replaceToken As String = vbNullString) As Boolean
    If chartObject Is Nothing Then NxRaiseContractError "A single chart is required"
    NxPngValidateOutput outputPath, replaceToken
    If TypeName(chartObject) = "Chart" Then
        If Not chartObject.Export(Filename:=outputPath, FilterName:=NxImageExportFilter(outputPath)) Then NxRaiseContractError "차트 그림 저장에 실패했습니다."
    ElseIf Not chartObject.Chart.Export(Filename:=outputPath, FilterName:=NxImageExportFilter(outputPath)) Then
        NxRaiseContractError "Chart PNG export failed"
    End If
    NxPngExportChart = True
End Function

Public Sub NxPngValidateOutput(ByVal outputPath As String, ByVal replaceToken As String)
    Dim fso As Object, folderPath As String
    Call NxImageExportFilter(outputPath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = fso.GetParentFolderName(outputPath)
    If Len(folderPath) = 0 Then NxRaiseContractError "그림을 저장할 폴더를 선택하세요."
    If Not fso.FolderExists(folderPath) Then NxRaiseContractError "그림 저장 폴더를 찾을 수 없습니다."
    If fso.FolderExists(outputPath) Then NxRaiseContractError "파일 경로에 같은 이름의 폴더가 있습니다."
    If fso.FileExists(outputPath) And replaceToken <> "REPLACE" Then NxRaiseContractError "같은 이름의 그림 파일이 있습니다. 다른 이름으로 저장하세요."
End Sub

Public Function NxImageExportFilter(ByVal outputPath As String) As String
    Select Case LCase$(Mid$(outputPath, InStrRev(outputPath, ".") + 1))
        Case "png": NxImageExportFilter = "PNG"
        Case "jpg", "jpeg": NxImageExportFilter = "JPG"
        Case Else: NxRaiseContractError "그림 파일 형식은 PNG 또는 JPG를 선택하세요."
    End Select
End Function
