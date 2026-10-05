Attribute VB_Name = "NxWorkflowRouter"
Option Explicit

Public Sub NxRouteWorkbench(ByVal workbenchId As String)
    ' This is intentionally a closed, non-executing launch surface.
    Select Case workbenchId
        Case "NX-WB-AI": NxAiOpenInput
        Case "NX-WB-DATA": NxRaiseContractError "데이터 목록에서 실행할 기능을 선택하세요."
        Case "NX-WB-DOC": NxDocumentOpen
        Case "NX-WB-DRAW": NxRaiseContractError "스타일·표그리기 목록에서 실행할 기능을 선택하세요."
        Case "NX-WB-FILE": NxRaiseContractError "파일관리 목록에서 실행할 기능을 선택하세요."
        Case Else: NxRaiseContractError "Unknown workbench"
    End Select
End Sub

Public Sub NxDocumentOpen()
    Dim form As New FNxDocument
    form.Show vbModeless
End Sub

Public Sub NxOpenStart()
    NxProductSettingsOpenMenu
End Sub
