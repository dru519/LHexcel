Option Explicit
Private todayValue As Date
Private noticeCount As Long
Private profilePath As String
Public Function PolicyTestDate() As Date
    PolicyTestDate = todayValue
End Function
Public Sub PolicyTestNotice(ByVal kind As String)
    noticeCount = noticeCount + 1
End Sub
Public Function NxLHexcelProfileRoot() As String
    NxLHexcelProfileRoot = profilePath
End Function
Private Sub AssertPolicy(ByVal condition As Boolean, ByVal label As String)
    If Not condition Then Err.Raise vbObjectError + 160, , label
End Sub
Public Function PolicyProbe(ByVal root As String) As String
    Dim handle As Integer, stored As String, cfg As String, previousNotices As Long
    On Error GoTo Failed
    profilePath = root
    todayValue = DateSerial(2027, 6, 29)
    AssertPolicy LHEDistributionPolicyCanOpen(), "expiry_before"
    todayValue = DateSerial(2027, 6, 30)
    AssertPolicy Not LHEDistributionPolicyCanOpen(), "expiry_at"
    todayValue = DateSerial(2027, 7, 1)
    AssertPolicy Not LHEDistributionPolicyCanOpen(), "expiry_after"
    noticeCount = 0
    todayValue = DateSerial(2026, 6, 16)
    LHEDistributionPolicyShowBirthdayIfNeeded
    AssertPolicy noticeCount = 0, "non_birthday"
    todayValue = DateSerial(2026, 6, 17)
    LHEDistributionPolicyShowBirthdayIfNeeded
    LHEDistributionPolicyShowBirthdayIfNeeded
    AssertPolicy noticeCount = 1, "same_year_dedup"
    todayValue = DateSerial(2027, 6, 17)
    LHEDistributionPolicyShowBirthdayIfNeeded
    AssertPolicy noticeCount = 2, "next_year_notice"
    cfg = root & "\Settings\distribution-notice-v1.cfg"
    handle = FreeFile
    Open cfg For Input Access Read Lock Read Write As #handle
    todayValue = DateSerial(2028, 6, 17)
    LHEDistributionPolicyShowBirthdayIfNeeded
    Line Input #handle, stored
    Close #handle
    AssertPolicy stored = "2027", "locked_write_preserves_previous_year"
    profilePath = root & "-legacy"
    MkDir profilePath
    cfg = profilePath & "\Settingsdistribution-notice-v1.cfg"
    handle = FreeFile
    Open cfg For Output As #handle
    Print #handle, "2026"
    Close #handle
    handle = 0
    todayValue = DateSerial(2026, 6, 17)
    previousNotices = noticeCount
    LHEDistributionPolicyShowBirthdayIfNeeded
    AssertPolicy noticeCount = previousNotices, "legacy_path_dedup"
    PolicyProbe = "PASS|8|date input and notice sink are test substitutions"
    Exit Function
Failed:
    PolicyProbe = "FAIL|" & Err.Description
    On Error Resume Next
    If handle > 0 Then Close #handle
End Function
Public Function PreviewProbe(ByVal productName As String) As String
    Dim cell As Range, preset As Variant, outputMode As Variant, result As String, cases As Long
    Dim beforeFormula As Variant, beforeFormat As Variant, originalSheets As Long
    On Error GoTo Failed
    Set cell = ThisWorkbook.Worksheets(1).Range("A1")
    cell.Value2 = " 1,200 "
    beforeFormula = cell.Formula
    beforeFormat = cell.NumberFormat
    originalSheets = ThisWorkbook.Worksheets.Count
    For Each preset In Array("기본 정리", "숫자 변환", "텍스트 변환", "날짜 정리", "표 구조 정리")
        For Each outputMode In Array("새 시트 만들기", "원본 변경")
            result = Application.Run("'" & productName & "'!NxDataNormalizationPreview", cell, CStr(preset), "yyyy-mm-dd", CStr(outputMode))
            AssertPolicy InStr(result, "A1") > 0 And InStr(result, CStr(outputMode)) > 0, "normalization_preview"
            cases = cases + 1
        Next outputMode
    Next preset
    For Each preset In Array("NX-DATA-UNIQUE-COUNT", "NX-DATA-DUPLICATE-LIST")
        result = Application.Run("'" & productName & "'!NxDataPreviewText", CStr(preset), cell, False, 1, True, False, True)
        AssertPolicy InStr(result, "고유 그룹") > 0, "data_preview"
        cases = cases + 1
    Next preset
    AssertPolicy cell.Formula = beforeFormula And cell.NumberFormat = beforeFormat, "preview_preserves_source"
    AssertPolicy ThisWorkbook.Worksheets.Count = originalSheets, "preview_no_output_sheet"
    PreviewProbe = "PASS|" & CStr(cases) & "|function checks; popup visual NOT_RUN"
    Exit Function
Failed:
    PreviewProbe = "FAIL|" & Err.Description
End Function
