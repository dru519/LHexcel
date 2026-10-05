Option Explicit

Private mRegistry As CNxFeatureRegistry
Private mFeatures As Collection
Private mHealthy As Boolean
Private mDiagnostic As String
Private mAllowClose As Boolean

Public Sub BindRegistry(ByVal registry As CNxFeatureRegistry)
    On Error GoTo Failed
    If registry Is Nothing Then NxRaiseContractError "Product registry is required"
    ValidateProductRegistry registry
    Set mRegistry = registry
    mHealthy = True
    mDiagnostic = vbNullString
    LoadFeatures vbNullString
    SetHealthyState True
    Exit Sub
Failed:
    SetFailure "NX-RECOVERY-REGISTRY"
End Sub

Private Sub UserForm_Initialize()
    Set mFeatures = New Collection
    mHealthy = False
    SetFailure "NX-RECOVERY-NOT-BOUND"
End Sub

Private Sub UserForm_Activate()
    txtSearch.SetFocus
End Sub

Private Sub txtSearch_Change()
    If mHealthy Then LoadFeatures txtSearch.Text
End Sub

Private Sub lstFeatures_Click()
    If mHealthy Then cmdExecute.Enabled = (lstFeatures.ListIndex >= 0)
End Sub

Private Sub cmdExecute_Click()
    Dim featureId As String
    On Error GoTo Failed
    If Not mHealthy Then Exit Sub
    If lstFeatures.ListIndex < 0 Then Exit Sub
    featureId = CStr(lstFeatures.List(lstFeatures.ListIndex, 0))
    ' The shared ribbon target is the only normal execution boundary.
    NxRibbonExecuteTag "nx1|feature|" & featureId
    Exit Sub
Failed:
    SetFailure "NX-RECOVERY-ROUTER"
End Sub

Private Sub cmdCopyDiagnostic_Click()
    Dim clipboard As Object
    If Len(mDiagnostic) = 0 Then Exit Sub
    On Error GoTo Done
    Set clipboard = CreateObject("Forms.DataObject")
    clipboard.SetText mDiagnostic
    clipboard.PutInClipboard
Done:
End Sub

Private Sub cmdClose_Click()
    RequestCancel
End Sub

Public Sub RequestCancel()
    mAllowClose = True
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If mAllowClose Then Exit Sub
    Cancel = True
    RequestCancel
End Sub

Private Sub LoadFeatures(ByVal query As String)
    Dim categoryId As Variant
    Dim definition As CNxFeatureDefinition
    Dim categoryFeatures As Collection
    Dim needle As String
    Dim labelText As String
    Set mFeatures = New Collection
    lstFeatures.Clear
    needle = LCase$(Trim$(query))
    For Each categoryId In Array("NX-CAT-AI", "NX-CAT-TEMPLATE", "NX-CAT-DATA", "NX-CAT-DRAW", "NX-CAT-FILE", "NX-CAT-SYMBOLS", "NX-CAT-CALCULATOR", "NX-CAT-UTIL")
        Set categoryFeatures = mRegistry.FeaturesForCategory(CStr(categoryId))
        For Each definition In categoryFeatures
            labelText = NxGeneratedFeatureLabel(definition.FeatureId)
            If needle = vbNullString Or InStr(1, LCase$(definition.FeatureId & " " & labelText & " " & definition.LabelKey), needle, vbTextCompare) > 0 Then
                mFeatures.Add definition
                lstFeatures.AddItem definition.FeatureId
                lstFeatures.List(lstFeatures.ListCount - 1, 1) = labelText
            End If
        Next definition
    Next categoryId
    lblStatus.Caption = CStr(mFeatures.Count) & "개 기능을 찾았습니다."
    cmdExecute.Enabled = (lstFeatures.ListIndex >= 0)
End Sub

Private Sub SetFailure(ByVal code As String)
    mHealthy = False
    mDiagnostic = "지원 코드: " & code
    lblStatus.Caption = "기능 목록을 사용할 수 없습니다. 진단을 복사하거나 닫아 주세요."
    SetHealthyState False
End Sub

Private Sub SetHealthyState(ByVal healthy As Boolean)
    txtSearch.Enabled = healthy
    lstFeatures.Enabled = healthy
    cmdExecute.Enabled = healthy And (lstFeatures.ListIndex >= 0)
    cmdCopyDiagnostic.Enabled = Not healthy
End Sub
