VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxCommandRegistry"
Option Explicit

Private mById As Object
Private mInOrder As Collection

Private Sub Class_Initialize()
    Set mById = CreateObject("Scripting.Dictionary")
    mById.CompareMode = vbBinaryCompare
    Set mInOrder = New Collection
End Sub

Public Sub Add(ByVal commandId As String, ByVal labelKo As String, ByVal categoryId As String, _
    ByVal groupId As String, ByVal handler As String, ByVal argumentKey As String, _
    ByVal mutatesDocument As Boolean, ByVal shortcutEligible As Boolean, _
    Optional ByVal recoveryPolicy As String = "none")
    Dim definition As New CNxCommandDefinition
    If mById.Exists(commandId) Then NxRaiseContractError "Duplicate command id"
    definition.Configure commandId, labelKo, categoryId, groupId, handler, argumentKey, mutatesDocument, shortcutEligible, recoveryPolicy
    definition.Seal
    mById.Add commandId, definition
    mInOrder.Add definition
End Sub

Public Function CommandById(ByVal commandId As String) As CNxCommandDefinition
    If mById.Exists(commandId) Then Set CommandById = mById(commandId)
End Function

Public Function AllCommands() As Collection
    Dim result As New Collection
    Dim definition As CNxCommandDefinition
    For Each definition In mInOrder
        result.Add definition
    Next definition
    Set AllCommands = result
End Function

Public Function CommandsForCategory(ByVal categoryId As String) As Collection
    Dim result As New Collection
    Dim definition As CNxCommandDefinition
    For Each definition In mInOrder
        If definition.CategoryId = categoryId Then result.Add definition
    Next definition
    Set CommandsForCategory = result
End Function

Public Property Get Count() As Long
    Count = mInOrder.Count
End Property
