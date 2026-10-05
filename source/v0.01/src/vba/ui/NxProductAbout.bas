Attribute VB_Name = "NxProductAbout"
Option Explicit

Public Sub NxProductShowAbout()
    Dim aboutForm As New FNxAbout
    aboutForm.BindProductInfo NxProductVersionText(), NxDistributionValidUntilText()
    aboutForm.Show vbModal
End Sub

Public Function NxProductVersionText() As String
    NxProductVersionText = "내엑셀 v0.01_r105"
End Function

Public Function NxProductCountsText() As String
    NxProductCountsText = "대분류 12개 · 기능 " & CStr(NxGeneratedNavigationItems().Count) & "개"
End Function

Public Function NxProductCommandCountText() As String
    NxProductCommandCountText = "전체기능 기준 · 중복 및 삭제 항목 제외"
End Function

Public Function NxProductBuildIdText() As String
    NxProductBuildIdText = "r105 · 시트비교 링크 및 파일통합 목차"
End Function
