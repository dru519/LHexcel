Attribute VB_Name = "NxProductRegistry"
Option Explicit
Private Const NX_EXPECTED_CATEGORY_COUNT As Long = 8
Private mProductRegistry As CNxFeatureRegistry

' The registry is static product metadata. Build and validate it once per session;
' selection, workbook and availability state are never cached here.
Public Function NxCreateProductRegistry() As CNxFeatureRegistry
    Dim registry As CNxFeatureRegistry
    If Not mProductRegistry Is Nothing Then
        Set NxCreateProductRegistry = mProductRegistry
        Exit Function
    End If
    Set registry = NxCreateGeneratedProductRegistry()
    ValidateProductRegistry registry
    registry.Seal
    Set mProductRegistry = registry
    Set NxCreateProductRegistry = registry
End Function

Public Sub NxResetProductRegistry()
    Set mProductRegistry = Nothing
End Sub
Public Function CreateProductRegistry() As CNxFeatureRegistry
    Set CreateProductRegistry = NxCreateProductRegistry()
End Function
Public Sub ValidateProductRegistry(ByVal registry As CNxFeatureRegistry)
    Dim featureId As Variant, releaseIds As Variant
    If registry Is Nothing Then NxRaiseContractError "Product registry is required"
    If registry.CategoryCount <> NX_EXPECTED_CATEGORY_COUNT Then NxRaiseContractError "Product registry category count is invalid"
    releaseIds = NxGeneratedReleaseFeatureIds()
    If registry.FeatureCount <> NxProductSurfaceFeatureCount() Or _
       UBound(releaseIds) - LBound(releaseIds) + 1 <> NxProductSurfaceFeatureCount() Then _
        NxRaiseContractError "Product registry feature count is invalid"
    RequireCategory registry, "NX-CAT-AI"
    RequireCategory registry, "NX-CAT-TEMPLATE"
    RequireCategory registry, "NX-CAT-DATA"
    RequireCategory registry, "NX-CAT-UTIL"
    RequireCategory registry, "NX-CAT-DRAW"
    RequireCategory registry, "NX-CAT-FILE"
    RequireCategory registry, "NX-CAT-SYMBOLS"
    RequireCategory registry, "NX-CAT-CALCULATOR"
    ' Every contract release ID must resolve; unknown IDs raise from FeatureById.
    For Each featureId In releaseIds
        RequireFeature registry, CStr(featureId)
    Next featureId
End Sub

Private Sub RequireCategory(ByVal registry As CNxFeatureRegistry, ByVal categoryId As String)
    Dim definition As CNxCategoryDefinition
    Set definition = registry.CategoryById(categoryId)
End Sub

Private Sub RequireFeature(ByVal registry As CNxFeatureRegistry, ByVal featureId As String)
    Dim definition As CNxFeatureDefinition
    Set definition = registry.FeatureById(featureId)
End Sub
