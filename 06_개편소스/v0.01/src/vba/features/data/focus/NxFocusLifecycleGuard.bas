Attribute VB_Name = "NxFocusLifecycleGuard"
Option Explicit

Public Function NxFocusCleanupWorkbook(ByVal workbook As Workbook) As Boolean
    Dim sheet As Worksheet
    Dim wasSaved As Boolean
    Dim clean As Boolean
    On Error GoTo Failed
    If workbook Is Nothing Then NxFocusCleanupWorkbook = True: Exit Function
    If workbook Is ThisWorkbook Then NxFocusCleanupWorkbook = True: Exit Function
    wasSaved = workbook.Saved
    clean = True
    For Each sheet In workbook.Worksheets
        If Not NxFocusDeleteManagedRules(sheet) Then clean = False
    Next sheet
    If wasSaved And clean Then workbook.Saved = True
    NxFocusCleanupWorkbook = clean
    Exit Function
Failed:
    NxFocusCleanupWorkbook = False
End Function
Public Function NxFocusCleanupAllOpenWorkbooks() As Boolean
    Dim workbook As Workbook
    Dim clean As Boolean
    clean = True
    On Error GoTo Failed
    For Each workbook In Application.Workbooks
        If Not workbook Is ThisWorkbook Then
            If Not NxFocusCleanupWorkbook(workbook) Then clean = False
        End If
    Next workbook
    NxFocusCleanupAllOpenWorkbooks = clean
    Exit Function
Failed:
    NxFocusCleanupAllOpenWorkbooks = False
End Function
