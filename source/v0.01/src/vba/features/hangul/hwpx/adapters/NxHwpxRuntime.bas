Attribute VB_Name = "NxHwpxRuntime"
Option Explicit

Private Const SCRIPT_RELATIVE As String = "tools\lhexcel_hwpx_table_export.ps1"
Private Const TEMPLATE_RELATIVE As String = "templates\표.hwpx"

Public Function NxHwpxEnsureEmbeddedResources() As Variant
    Dim root As String, scriptPath As String, templatePath As String
    root = NxHwpxEmbeddedRoot()
    scriptPath = NxHwpxJoinPath(root, SCRIPT_RELATIVE)
    templatePath = NxHwpxJoinPath(root, TEMPLATE_RELATIVE)
    EnsureEmbeddedFile scriptPath, NxHwpxEmbeddedScriptBase64(), NxHwpxEmbeddedScriptSize(), NxHwpxEmbeddedScriptSha256()
    EnsureEmbeddedFile templatePath, NxHwpxEmbeddedTemplateBase64(), NxHwpxEmbeddedTemplateSize(), NxHwpxEmbeddedTemplateSha256()
    NxHwpxEnsureEmbeddedResources = Array(scriptPath, templatePath)
End Function

Public Function NxHwpxEnsureEmbeddedTemplate() As String
    Dim templatePath As String
    templatePath = NxHwpxJoinPath(NxHwpxEmbeddedRoot(), TEMPLATE_RELATIVE)
    EnsureEmbeddedFile templatePath, NxHwpxEmbeddedTemplateBase64(), NxHwpxEmbeddedTemplateSize(), NxHwpxEmbeddedTemplateSha256()
    NxHwpxEnsureEmbeddedTemplate = templatePath
End Function

Public Function NxHwpxEmbeddedRoot() As String
    Dim basePath As String
    basePath = Environ$("TEMP")
    If Len(basePath) = 0 Then basePath = Environ$("TMP")
    If Len(basePath) = 0 Then NxRaiseContractError "HWPX 내장 리소스를 복원할 임시 경로를 찾을 수 없습니다."
    NxHwpxEmbeddedRoot = NxHwpxJoinPath(basePath, "LHexcel\embedded_support\" & NxHwpxEmbeddedResourceVersion())
End Function

Public Function NxHwpxEmbeddedResourceNames() As String
    NxHwpxEmbeddedResourceNames = SCRIPT_RELATIVE & "|" & TEMPLATE_RELATIVE
End Function

Public Function NxHwpxJoinPath(ByVal root As String, ByVal relativePath As String) As String
    If Len(root) = 0 Or Len(relativePath) = 0 Then NxRaiseContractError "HWPX 런타임 경로가 비어 있습니다."
    If Right$(root, 1) = "\" Or Right$(root, 1) = "/" Then
        NxHwpxJoinPath = root & relativePath
    Else
        NxHwpxJoinPath = root & "\" & relativePath
    End If
End Function

Public Function NxHwpxRuntimeOutputRoot() As String
    ' NxLHexcelProfileRoot validates the optional LHEXCEL_PROFILE_ROOT override.
    ' Native verification therefore stays isolated while production still
    ' defaults to LOCALAPPDATA\LHexcel.
    NxHwpxRuntimeOutputRoot = NxLHexcelProfileRoot() & "\Temp\Hwpx"
End Function

Private Sub EnsureEmbeddedFile(ByVal targetPath As String, ByVal base64Text As String, ByVal expectedSize As Long, ByVal expectedSha256 As String)
    If EmbeddedFileMatches(targetPath, expectedSize, expectedSha256) Then Exit Sub
    EnsureFolder ParentFolder(targetPath)
    On Error Resume Next
    If Len(Dir$(targetPath, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) > 0 Then Kill targetPath
    On Error GoTo 0
    WriteBase64File targetPath, base64Text
    If Not EmbeddedFileMatches(targetPath, expectedSize, expectedSha256) Then
        On Error Resume Next
        Kill targetPath
        On Error GoTo 0
        NxRaiseContractError "복원된 HWPX 내장 리소스의 무결성을 확인할 수 없습니다: " & targetPath
    End If
End Sub

Private Function EmbeddedFileMatches(ByVal path As String, ByVal expectedSize As Long, ByVal expectedSha256 As String) As Boolean
    On Error GoTo NotMatched
    If Len(Dir$(path, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)) = 0 Then Exit Function
    If FileLen(path) <> expectedSize Then Exit Function
    EmbeddedFileMatches = (StrComp(FileSha256(path), expectedSha256, vbBinaryCompare) = 0)
    Exit Function
NotMatched:
    EmbeddedFileMatches = False
End Function

Private Function FileSha256(ByVal path As String) As String
    ' Reuse the in-process Windows CNG implementation; no helper process.
    FileSha256 = NxTemplateFileSha256(path)
End Function

Private Function IsHex64(ByVal value As String) As Boolean
    Dim index As Long, code As Long
    If Len(value) <> 64 Then Exit Function
    For index = 1 To Len(value)
        code = AscW(Mid$(value, index, 1))
        If Not ((code >= 48 And code <= 57) Or (code >= 65 And code <= 70) Or (code >= 97 And code <= 102)) Then Exit Function
    Next index
    IsHex64 = True
End Function

Private Function QuoteCommand(ByVal value As String) As String
    QuoteCommand = Chr$(34) & Replace$(value, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

Private Function EncodePowerShellCommand(ByVal commandText As String) As String
    Dim stream As Object, dom As Object, node As Object, bytes As Variant
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "unicode"
    stream.Open
    stream.WriteText commandText
    stream.Position = 0
    stream.Type = 1
    bytes = stream.Read
    stream.Close
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = dom.createElement("b64")
    node.DataType = "bin.base64"
    node.NodeTypedValue = bytes
    EncodePowerShellCommand = Replace$(Replace$(node.Text, vbCr, vbNullString), vbLf, vbNullString)
End Function

Private Sub WriteBase64File(ByVal targetPath As String, ByVal base64Text As String)
    Dim dom As Object, node As Object, stream As Object
    Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = dom.createElement("b64")
    node.DataType = "bin.base64"
    node.Text = base64Text
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write node.NodeTypedValue
    stream.SaveToFile targetPath, 2
    stream.Close
End Sub

Private Sub EnsureFolder(ByVal folderPath As String)
    Dim fileSystem As Object, parent As String
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    If fileSystem.FolderExists(folderPath) Then Exit Sub
    parent = fileSystem.GetParentFolderName(folderPath)
    If Len(parent) > 0 And Not fileSystem.FolderExists(parent) Then EnsureFolder parent
    fileSystem.CreateFolder folderPath
End Sub

Private Function ParentFolder(ByVal path As String) As String
    ParentFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(path)
End Function

Private Function ElapsedSeconds(ByVal started As Single) As Double
    If Timer >= started Then ElapsedSeconds = Timer - started Else ElapsedSeconds = (86400# - started) + Timer
End Function
