VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxTemplateTransaction"
Option Explicit

Private mRoot As String
Private mTemplateId As String
Private mCreatedFiles As Collection
Private mCreatedFolders As Collection
Private mCreatedSheets As Collection
Private mConfigured As Boolean
Private mCommitted As Boolean

Private Sub Class_Initialize()
    Set mCreatedFiles = New Collection
    Set mCreatedFolders = New Collection
    Set mCreatedSheets = New Collection
End Sub

Public Sub TrackCreatedFolder(ByVal path As String)
    EnsureActive
    If LCase$(Left$(path, Len(mRoot) + 1)) <> LCase$(mRoot & "\") Then NxRaiseContractError "Template transaction folder escaped its root"
    mCreatedFolders.Add path
End Sub

Public Sub Configure(ByVal root As String, ByVal templateId As String)
    If mConfigured Then NxRaiseContractError "Template transaction can be configured only once"
    If Len(root) = 0 Or Not NxTemplateIdIsValid(templateId) Then NxRaiseContractError "Template transaction identity is invalid"
    mRoot = root
    mTemplateId = templateId
    mConfigured = True
End Sub

Public Sub TrackCreatedFile(ByVal path As String)
    EnsureActive
    If LCase$(Left$(path, Len(mRoot) + 1)) <> LCase$(mRoot & "\") Then NxRaiseContractError "Template transaction file escaped its root"
    mCreatedFiles.Add path
End Sub

Public Sub TrackCreatedSheet(ByVal sheet As Worksheet)
    EnsureActive
    If sheet Is Nothing Then NxRaiseContractError "Template transaction sheet tracking is invalid"
    mCreatedSheets.Add sheet
End Sub

Public Sub Commit()
    EnsureActive
    mCommitted = True
End Sub

Public Sub Rollback()
    Dim index As Long, alerts As Boolean, fileSystem As Object
    If Not mConfigured Or mCommitted Then Exit Sub
    On Error Resume Next
    alerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For index = mCreatedSheets.Count To 1 Step -1
        mCreatedSheets.Item(index).Delete
    Next index
    Application.DisplayAlerts = alerts
    For index = mCreatedFiles.Count To 1 Step -1
        If Len(Dir$(CStr(mCreatedFiles.Item(index)))) > 0 Then Kill CStr(mCreatedFiles.Item(index))
    Next index
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    For index = mCreatedFolders.Count To 1 Step -1
        If fileSystem.FolderExists(CStr(mCreatedFolders.Item(index))) Then fileSystem.DeleteFolder CStr(mCreatedFolders.Item(index)), True
    Next index
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Rollback
End Sub

Private Sub EnsureActive()
    If Not mConfigured Or mCommitted Then NxRaiseContractError "Template transaction is not active"
End Sub
