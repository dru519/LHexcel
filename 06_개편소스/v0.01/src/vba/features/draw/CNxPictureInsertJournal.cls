VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxPictureInsertJournal"
Option Explicit

Private mWorksheet As Worksheet, mWorksheetIdentity As String
Private mCreated As Collection, mCreatedNames As Collection

Private Sub Class_Initialize()
    Set mCreated = New Collection
    Set mCreatedNames = New Collection
End Sub

Public Sub Bind(ByVal target As Range)
    If target Is Nothing Then NxRaiseContractError "그림 삽입 journal 대상이 필요합니다."
    Set mWorksheet = target.Worksheet: mWorksheetIdentity = CStr(target.Parent.Parent.FullName) & "|" & CStr(target.Worksheet.CodeName)
End Sub

Public Sub Capture(ByVal shape As Shape)
    If mWorksheet Is Nothing Or shape Is Nothing Then NxRaiseContractError "삽입 도형 journal이 구성되지 않았습니다."
    If Not shape.Parent Is mWorksheet Then NxRaiseContractError "journal은 다른 시트의 도형을 기록할 수 없습니다."
    mCreated.Add shape
    mCreatedNames.Add CStr(shape.Name)
End Sub

Public Sub Rollback()
    Dim index As Long, createdShape As Shape
    If mWorksheet Is Nothing Then Exit Sub
    If CStr(mWorksheet.Parent.FullName) & "|" & CStr(mWorksheet.CodeName) <> mWorksheetIdentity Then Exit Sub
    On Error Resume Next
    For index = mCreated.Count To 1 Step -1
        Set createdShape = Nothing
        Set createdShape = mCreated(index)
        If Not createdShape Is Nothing Then
            If createdShape.Parent Is mWorksheet Then createdShape.Delete
        End If
    Next index
    On Error GoTo 0
End Sub

Public Property Get CreatedNames() As Collection
    Dim answer As New Collection, item As Variant
    For Each item In mCreatedNames: answer.Add CStr(item): Next item
    Set CreatedNames = answer
End Property

Public Property Get WorksheetIdentity() As String: WorksheetIdentity = mWorksheetIdentity: End Property
