Attribute VB_Name = "T_R69PopupLayout"
Option Explicit
Private mPanel As Object

Public Function Names() As String
    Names = "emphasis|presets|date|age|money|mask|compare|copy_result"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim target As Range, control As Object
    On Error GoTo Failed
    ClosePanel
    Set target = ActiveWorkbook.Worksheets(1).Range("B2:C3")
    target.Value2 = 1234.56789: target.Select
    Select Case name
        Case "emphasis"
            Set mPanel = New FNxNumberFormat: mPanel.BindTarget target, False, True
        Case "presets"
            Set mPanel = New FNxViewPresets: mPanel.BindSettings
        Case "percent", "decimal"
            Set mPanel = New FNxNumberFormat
            mPanel.BindTarget target, name = "percent"
        Case "resize"
            Set mPanel = New FNxResize: mPanel.BindTarget target
        Case "age"
            Set mPanel = New FNxAgeCalculator: mPanel.BindFeature NX_FEATURE_DATA_AGE
        Case "money", "mask"
            Set mPanel = New FNxDataSpecial
            mPanel.BindFeature IIf(name = "money", NX_FEATURE_DATA_KOREAN_MONEY, NX_FEATURE_DATA_PRIVACY_MASK)
        Case "compare"
            Set mPanel = New FNxWorkbookCompare: mPanel.BindFeature
        Case "copy_result"
            Set mPanel = New FNxCopySaveResult: mPanel.BindPath "C:\Documents\업무용 폴더\선택범위_사본.xlsx"
        Case "date"
            NxOpenDateConversion
            For Each control In VBA.UserForms
                If TypeName(control) = "FNxDataNormalize" Then Set mPanel = control
            Next control
    End Select
    If mPanel Is Nothing Then Err.Raise 5, , "Popup missing"
    If Not mPanel.Visible Then mPanel.Show vbModeless
    mPanel.Repaint
    DoEvents
    For Each control In mPanel.Controls
        If control.Visible Then
            If control.Left < 0 Or control.Top < 0 Or control.Left + control.Width > mPanel.InsideWidth + 0.5 Or control.Top + control.Height > mPanel.InsideHeight + 0.5 Then Err.Raise 5, , "Clipped control: " & control.Name
        End If
    Next control
    RunCase = "PASS|" & name
    Exit Function
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Description
End Function

Public Function Caption() As String
    Caption = mPanel.Caption
End Function

Public Sub ClosePanel()
    If Not mPanel Is Nothing Then Unload mPanel
    Set mPanel = Nothing
End Sub
