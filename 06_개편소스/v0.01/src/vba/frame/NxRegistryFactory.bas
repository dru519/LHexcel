Attribute VB_Name = "NxRegistryFactory"
Option Explicit

Public Function CreateBaseRegistry() As CNxFeatureRegistry
    Dim registry As New CNxFeatureRegistry
    AddCategory registry, "NX-CAT-AI", "NX-ENTRY-AI", "AI 모드", 10
    AddCategory registry, "NX-CAT-TEMPLATE", "NX-ENTRY-TEMPLATE", "내엑셀 템플릿", 20
    AddCategory registry, "NX-CAT-DATA", "NX-ENTRY-ALL", "데이터/추가기능", 30
    AddCategory registry, "NX-CAT-DRAW", "NX-ENTRY-ALL", "표/그리기", 50
    AddCategory registry, "NX-CAT-FILE", "NX-ENTRY-ALL", "파일관리", 60
    AddCategory registry, "NX-CAT-SYMBOLS", "NX-ENTRY-ALL", "기호표", 70
    AddCategory registry, "NX-CAT-CALCULATOR", "NX-ENTRY-ALL", "계산기", 80
    AddCategory registry, "NX-CAT-UTIL", "NX-ENTRY-ALL", "추가기능", 90
    Set CreateBaseRegistry = registry
End Function

Private Sub AddCategory( _
    ByVal registry As CNxFeatureRegistry, _
    ByVal categoryId As String, _
    ByVal parentId As String, _
    ByVal label As String, _
    ByVal sortOrder As Long)

    Dim definition As New CNxCategoryDefinition
    definition.Configure categoryId, parentId, label, sortOrder
    definition.Seal
    registry.RegisterCategory definition
End Sub
