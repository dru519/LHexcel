Attribute VB_Name = "T_R74Folder"
Option Explicit

Public Function Names() As String
    Names = "unsaved_folder_rejected"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim wb As Workbook, code As Long
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error Resume Next
    NxProductOpenSavedFolder
    code = Err.Number
    Err.Clear
    On Error GoTo 0
    wb.Close False
    If code <> 0 Then
        RunCase = "PASS|" & name
    Else
        RunCase = "FAIL|" & name & "|unsaved workbook accepted"
    End If
End Function
