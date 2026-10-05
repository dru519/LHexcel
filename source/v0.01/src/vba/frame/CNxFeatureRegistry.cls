VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFeatureRegistry"
Option Explicit

Private mCategories As Collection
Private mFeatures As Collection
Private mSealed As Boolean

Private Sub Class_Initialize()
    Set mCategories = New Collection
    Set mFeatures = New Collection
End Sub

Public Sub RegisterCategory(ByVal definition As CNxCategoryDefinition)
    If mSealed Then NxRaiseContractError "Sealed registry cannot change"
    If definition Is Nothing Then NxRaiseContractError "Registry requires a category definition"
    If Not definition.IsSealed Then NxRaiseContractError "Registry requires a sealed category definition"
    If ContainsKey(mCategories, definition.Id) Then NxRaiseContractError "Duplicate category ID"
    mCategories.Add definition, definition.Id
End Sub

Public Sub RegisterFeature(ByVal definition As CNxFeatureDefinition)
    If mSealed Then NxRaiseContractError "Sealed registry cannot change"
    If definition Is Nothing Then NxRaiseContractError "Registry requires a feature definition"
    If Not definition.IsSealed Then NxRaiseContractError "Registry requires a sealed feature definition"
    If ContainsKey(mFeatures, definition.FeatureId) Then NxRaiseContractError "Duplicate feature ID"
    If Not ContainsKey(mCategories, definition.CategoryId) Then NxRaiseContractError "Unknown feature category"
    mFeatures.Add definition, definition.FeatureId
End Sub

Public Function CategoryById(ByVal categoryId As String) As CNxCategoryDefinition
    If Not ContainsKey(mCategories, categoryId) Then NxRaiseContractError "Unknown category ID"
    Set CategoryById = mCategories.Item(categoryId)
End Function

Public Function FeatureById(ByVal featureId As String) As CNxFeatureDefinition
    If Not ContainsKey(mFeatures, featureId) Then NxRaiseContractError "Unknown feature ID"
    Set FeatureById = mFeatures.Item(featureId)
End Function

Public Function FeaturesForCategory(ByVal categoryId As String) As Collection
    Dim definition As CNxFeatureDefinition
    Dim ordered() As CNxFeatureDefinition
    Dim result As New Collection
    Dim count As Long
    Dim index As Long
    Dim inner As Long
    Dim temporary As CNxFeatureDefinition

    CategoryById categoryId
    For Each definition In mFeatures
        If definition.CategoryId = categoryId Then count = count + 1
    Next definition
    If count = 0 Then Set FeaturesForCategory = result: Exit Function

    ReDim ordered(1 To count)
    index = 0
    For Each definition In mFeatures
        If definition.CategoryId = categoryId Then
            index = index + 1
            Set ordered(index) = definition
        End If
    Next definition

    For index = 1 To count - 1
        For inner = index + 1 To count
            If ComesAfter(ordered(index), ordered(inner)) Then
                Set temporary = ordered(index)
                Set ordered(index) = ordered(inner)
                Set ordered(inner) = temporary
            End If
        Next inner
    Next index

    For index = 1 To count
        result.Add ordered(index)
    Next index
    Set FeaturesForCategory = result
End Function

Public Property Get CategoryCount() As Long: CategoryCount = mCategories.count: End Property
Public Property Get FeatureCount() As Long: FeatureCount = mFeatures.count: End Property
Public Property Get IsSealed() As Boolean: IsSealed = mSealed: End Property

Public Sub Seal()
    mSealed = True
End Sub

Private Function ComesAfter(ByVal leftValue As CNxFeatureDefinition, ByVal rightValue As CNxFeatureDefinition) As Boolean
    If leftValue.SortOrder <> rightValue.SortOrder Then
        ComesAfter = (leftValue.SortOrder > rightValue.SortOrder)
    Else
        ComesAfter = (StrComp(leftValue.FeatureId, rightValue.FeatureId, vbBinaryCompare) > 0)
    End If
End Function

Private Function ContainsKey(ByVal items As Collection, ByVal key As String) As Boolean
    Dim candidate As Object
    On Error Resume Next
    Set candidate = items.Item(key)
    ContainsKey = (Err.Number = 0 And Not candidate Is Nothing)
    Err.Clear
    On Error GoTo 0
End Function
