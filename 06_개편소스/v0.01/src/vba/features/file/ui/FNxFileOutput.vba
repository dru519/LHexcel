Option Explicit

Private mFeatureId As String
Private mSelectedTarget As Object
Private mRangeSession As CNxRangeSelectionSession
Private mPreviewReady As Boolean
Private mAllowClose As Boolean

Private Sub UserForm_Initialize()
    cmdExecute.Default = True
    cmdCancel.Cancel = True
End Sub

Public Sub BindFeature(ByVal featureId As String)
    Select Case featureId
        Case "NX-FILE-RANGE-PNG"
            lblTitle.Caption = "선택 범위를 PNG 또는 JPG 그림 파일로 저장합니다."
        Case "NX-FILE-CHART-PNG"
            lblTitle.Caption = "선택한 차트를 PNG 또는 JPG 그림 파일로 저장합니다."
        Case Else
            NxRaiseContractError "지원하지 않는 저장 기능입니다."
    End Select
    mFeatureId = featureId
    txtRange.Visible = (featureId = "NX-FILE-RANGE-PNG")
    cmdPickRange.Visible = (featureId = "NX-FILE-RANGE-PNG")
    If featureId = "NX-FILE-RANGE-PNG" And TypeName(Application.Selection) = "Range" Then
        Set mRangeSession = NxRangeSessionFromSelection()
        txtRange.Value = mRangeSession.DisplayAddress
    End If
    InvalidatePreview
End Sub

Private Sub txtRange_Change()
    InvalidatePreview
End Sub

Private Sub cmdPickRange_Click()
    On Error GoTo Failed
    If mFeatureId <> "NX-FILE-RANGE-PNG" Then NxRaiseContractError "이 기능은 셀 범위를 사용하지 않습니다."
    If mRangeSession Is Nothing Then Set mRangeSession = New CNxRangeSelectionSession
    If NxPickRange(Me, mRangeSession) Then txtRange.Value = mRangeSession.DisplayAddress
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdBrowse_Click()
    Dim selectedPath As Variant
    Dim filter As String
    Dim defaultName As String
    filter = "PNG 파일 (*.png), *.png, JPG 파일 (*.jpg), *.jpg"
    defaultName = "내엑셀_이미지.png"
    selectedPath = Application.GetSaveAsFilename(InitialFileName:=defaultName, FileFilter:=filter, _
        FilterIndex:=1, Title:="그림 저장 폴더·이름·파일 형식 선택")
    If VarType(selectedPath) <> vbBoolean Then txtOutputPath.Value = CStr(selectedPath)
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo Failed
    RequireOutputPath
    Set mSelectedTarget = Nothing
    Select Case mFeatureId
        Case "NX-FILE-RANGE-PNG"
            If mRangeSession Is Nothing Then NxRaiseContractError "저장할 셀 범위를 선택하세요."
            NxRangeSessionUpdateFromText mRangeSession, CStr(txtRange.Value)
            Set mSelectedTarget = mRangeSession.ResolveRange()
            txtPreview.Value = "선택 범위 " & mSelectedTarget.Address(False, False) & "를 " & NxImageExportFilter(CStr(txtOutputPath.Value)) & "로 저장합니다."
        Case "NX-FILE-CHART-PNG"
            Set mSelectedTarget = ResolveSelectedChart()
            txtPreview.Value = "선택한 차트를 " & NxImageExportFilter(CStr(txtOutputPath.Value)) & "로 저장합니다."
    End Select
    txtPreview.Value = txtPreview.Value & vbCrLf & "기존 파일은 자동으로 덮어쓰지 않습니다."
    mPreviewReady = True
    cmdExecute.Enabled = True
    lblStatus.Caption = "내용을 확인한 뒤 저장하세요."
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub cmdExecute_Click()
    Dim result As CNxResult
    On Error GoTo Failed
    RequireOutputPath
    If mFeatureId = "NX-FILE-RANGE-PNG" Then
        NxRangeSessionUpdateFromText mRangeSession, CStr(txtRange.Value)
    Else
        Set mSelectedTarget = ResolveSelectedChart()
    End If
    If mFeatureId = "NX-FILE-RANGE-PNG" Then Set mSelectedTarget = mRangeSession.ResolveRange()
    Set result = NxFileRun(mFeatureId, CStr(txtOutputPath.Value), Empty, False, mSelectedTarget, False)
    If result Is Nothing Then Exit Sub
    If result.Outcome = NxCancelled Then Exit Sub
    If result.Outcome <> NxSuccess Then NxRaiseContractError "파일을 저장하지 못했습니다."
    RequestCancel
    Exit Sub
Failed:
    ShowError
End Sub

Private Sub RequireOutputPath()
    If Len(Trim$(CStr(txtOutputPath.Value))) = 0 Then NxRaiseContractError "저장 위치를 선택하세요."
    NxPngValidateOutput CStr(txtOutputPath.Value), vbNullString
End Sub

Private Function ResolveSelectedChart() As Object
    Dim selected As Object
    Dim candidate As Object
    Dim chart As Object
    Dim owner As Object
    On Error Resume Next
    Set selected = Application.Selection
    Select Case TypeName(selected)
        Case "Chart": Set chart = selected
        Case "ChartObject": Set ResolveSelectedChart = selected: On Error GoTo 0: Exit Function
        Case "Shape": Set chart = selected.Chart
        Case "ShapeRange"
            If selected.Count = 1 Then Set candidate = selected.Item(1): Set chart = candidate.Chart
        Case "ChartArea", "PlotArea", "Legend", "Series": Set chart = selected.Parent
    End Select
    On Error GoTo 0
    If chart Is Nothing Then NxRaiseContractError "차트 하나를 먼저 선택하세요."
    On Error Resume Next
    Set owner = chart.Parent
    On Error GoTo 0
    If Not owner Is Nothing Then
        If TypeName(owner) = "ChartObject" Then Set ResolveSelectedChart = owner: Exit Function
    End If
    Set ResolveSelectedChart = chart
End Function

Private Sub txtOutputPath_Change(): InvalidatePreview: End Sub
Private Sub cmdCancel_Click(): RequestCancel: End Sub

Private Sub InvalidatePreview()
    mPreviewReady = False
    Set mSelectedTarget = Nothing
    cmdExecute.Enabled = True
    lblStatus.Caption = "옵션을 확인하고 실행하세요."
End Sub

Private Sub RequestCancel()
    mAllowClose = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    Cancel = True
    RequestCancel
End Sub

Private Sub ShowError()
    Dim detail As String
    detail = Err.Description
    Err.Clear
    detail = NxUserErrorText(detail)
    MsgBox detail, vbExclamation + vbOKOnly, "내엑셀 - 저장"
End Sub

Private Sub txtPreview_Change()
    NxNativePreviewTextChanged Me, txtPreview
End Sub
