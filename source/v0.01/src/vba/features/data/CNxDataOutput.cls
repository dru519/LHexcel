VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxDataOutput"
Option Explicit

Private mSource As Range, mOutput As Range, mColumns As Range
Private mSheet As Worksheet, mBook As Workbook
Private mMode As String, mStarted As Boolean, mWasSaved As Boolean

Public Sub Configure(ByVal source As Range, ByVal mode As String)
    If Not mSource Is Nothing Then NxRaiseContractError "결과 위치는 한 번만 구성합니다."
    Set mSource = source
    mMode = NxDataNormalizeOutputMode(mode)
    mWasSaved = source.Worksheet.Parent.Saved
    If mMode <> "새 통합문서" Then
        If source.Worksheet.Parent.ReadOnly Or source.Worksheet.Parent.ProtectStructure Then NxRaiseContractError "원본 통합문서의 읽기 전용·구조 보호를 확인하세요."
    End If
    If mMode = "오른쪽 새 열 삽입" Then
        If source.Worksheet.ProtectContents Then NxRaiseContractError "보호된 시트에는 결과 열을 삽입할 수 없습니다."
        If source.Column + source.Columns.Count * 2 - 1 > source.Worksheet.Columns.Count Then NxRaiseContractError "오른쪽에 새 열을 삽입할 공간이 없습니다."
    End If
End Sub

Public Function CreateOutput(ByVal featureId As String, ByVal plan As CNxPlan) As Range
    Dim effect As CNxSideEffect, expected As NxSideEffectKind, declared As Boolean
    If mSource Is Nothing Or plan Is Nothing Then NxRaiseContractError "결과 위치와 실행 계획이 필요합니다."
    If mStarted Or Not plan.IsSealed Or plan.FeatureId <> featureId Then NxRaiseContractError "결과 생성 계획을 확인하세요."
    expected = NxWorksheetCreate
    If mMode = "오른쪽 새 열 삽입" Then expected = NxColumnInsert
    For Each effect In plan.SideEffects
        If effect.Kind = expected And effect.MaximumCells >= mSource.CountLarge Then declared = True
    Next effect
    If Not declared Then NxRaiseContractError "승인되지 않은 결과 위치입니다."
    mStarted = True
    Select Case mMode
        Case "오른쪽 새 열 삽입"
            Set mColumns = mSource.Worksheet.Cells(1, mSource.Column + mSource.Columns.Count).Resize(1, mSource.Columns.Count).EntireColumn
            On Error GoTo InsertFailed
            mColumns.Insert Shift:=xlToRight
            On Error GoTo 0
            Set mColumns = mSource.Worksheet.Cells(1, mSource.Column + mSource.Columns.Count).Resize(1, mSource.Columns.Count).EntireColumn
            Set mSheet = mSource.Worksheet
            Set mOutput = mSheet.Cells(mSource.Row, mColumns.Column).Resize(mSource.Rows.Count, mSource.Columns.Count)
        Case "새 시트 만들기"
            Set mSheet = mSource.Worksheet.Parent.Worksheets.Add(After:=mSource.Worksheet.Parent.Worksheets(mSource.Worksheet.Parent.Worksheets.Count))
            mSheet.Name = NxDataNextNormalizationSheetName(mSheet.Parent, "변환결과")
            Set mOutput = mSheet.Range("A1").Resize(mSource.Rows.Count, mSource.Columns.Count)
        Case "새 통합문서"
            Set mBook = Application.Workbooks.Add(xlWBATWorksheet)
            Set mSheet = mBook.Worksheets(1)
            mSheet.Name = "변환결과"
            Set mOutput = mSheet.Range("A1").Resize(mSource.Rows.Count, mSource.Columns.Count)
        Case Else
            NxRaiseContractError "원본 변경은 셀 복구 저널을 사용해야 합니다."
    End Select
    If Not mBook Is Nothing Then NxDrawStyleGeneratedTable mBook, mOutput, 0
    Set CreateOutput = mOutput
    Exit Function
InsertFailed:
    Dim failure As Long, detail As String
    failure = Err.Number: detail = Err.Description
    Set mColumns = Nothing
    Err.Raise failure, "CNxDataOutput", detail
End Function

Public Property Get ResultSheet() As Worksheet
    Set ResultSheet = mSheet
End Property

Public Sub ActivateResult()
    If Not mOutput Is Nothing Then mOutput.Worksheet.Activate: mOutput.Select
End Sub

Public Sub Rollback()
    Dim alerts As Boolean, failure As Long, detail As String
    alerts = Application.DisplayAlerts
    On Error GoTo Failed
    Application.DisplayAlerts = False
    If Not mColumns Is Nothing Then
        mColumns.Delete Shift:=xlToLeft
    ElseIf Not mBook Is Nothing Then
        mBook.Close SaveChanges:=False
    ElseIf Not mSheet Is Nothing Then
        If mSheet Is mSource.Worksheet Then NxRaiseContractError "원본 시트는 삭제할 수 없습니다."
        mSheet.Delete
    End If
    Set mColumns = Nothing: Set mBook = Nothing: Set mSheet = Nothing: Set mOutput = Nothing
    If mStarted And mWasSaved Then mSource.Worksheet.Parent.Saved = True
    Application.DisplayAlerts = alerts
    Exit Sub
Failed:
    failure = Err.Number: detail = Err.Description: Err.Clear
    Application.DisplayAlerts = alerts
    Err.Raise failure, "CNxDataOutput.Rollback", "결과 위치 복구 실패: " & detail
End Sub
