Attribute VB_Name = "T_R96Ribbon"
Option Explicit

Public Function Names() As String
    Names = "scoped_ids|live_availability|standalone_menu|preview_stale"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim control As New T_R96Control, homeXml As Variant, productXml As Variant
    Dim result As Variant, book As Workbook, panel As FNxDataNormalize
    Dim failure As Long, detail As String
    On Error GoTo Failed
    If name = "scoped_ids" Then
        control.Id = "NX-HOME-ALL"
        NxRibbonGetAllFunctionsContent control, homeXml
        control.Id = "NX-PROD-ALL"
        NxRibbonGetAllFunctionsContent control, productXml
        Require InStr(homeXml, "id=""dyn_NX_HOME_ALL_") > 0, "home namespace"
        Require InStr(productXml, "id=""dyn_NX_PROD_ALL_") > 0, "product namespace"
        Require InStr(homeXml, " enabled=") = 0, "fixed enabled attribute"
        Require InStr(homeXml, "getEnabled=""NxRibbonGetEnabled""") > 0, "availability callback"
        Require InStr(homeXml, "image=""NxBrand_save_16""") > 0, "brand image in home menu"
        Require InStr(productXml, "image=""NxBrand_calculator_16""") > 0, "brand image in product menu"
    ElseIf name = "live_availability" Then
        Set book = Workbooks.Add(xlWBATWorksheet)
        book.Worksheets(1).Range("B2").Select
        control.Tag = "nx1|feature|NX-UTIL-CALCULATOR"
        control.Id = "dyn_NX_HOME_ALL_calculator"
        NxRibbonGetEnabled control, result
        Require CBool(result), "home enabled"
        control.Id = "dyn_NX_PROD_ALL_calculator"
        NxRibbonGetEnabled control, result
        Require CBool(result), "product enabled"
        control.Tag = "invalid"
        NxRibbonGetEnabled control, result
        Require Not CBool(result), "invalid route rejected"
    ElseIf name = "standalone_menu" Then
        Require NxDistributionIsStandaloneXlam(), "development profile"
        homeXml = NxRibbonManagementMenuXml()
        Require InStr(homeXml, "NX-MGMT-INSTALL""") > 0, "standalone install"
        Require InStr(homeXml, "NX-MGMT-REMOVE""") > 0, "standalone remove"
    ElseIf name = "preview_stale" Then
        Set panel = New FNxDataNormalize
        panel.Controls("txtPreview").Value = "old result"
        panel.Controls("cboDateFormat").ListIndex = 1
        Require InStr(panel.Controls("txtPreview").Value, "설정이 변경") > 0, "stale preview cleared"
        Require panel.Controls("cmdExecute").Enabled, "direct execution retained"
        Require panel.Controls("txtPreview").BackColor = vbWhite, "white preview"
    Else
        Err.Raise 5, , "Unknown r96 case"
    End If
Clean:
    On Error Resume Next
    If Not panel Is Nothing Then Unload panel
    If Not book Is Nothing Then book.Close False
    On Error GoTo 0
    If failure <> 0 Then Err.Raise failure, "T_R96Ribbon", detail
    RunCase = "PASS|" & name
    Exit Function
Failed:
    failure = Err.Number: detail = Err.Description
    Resume Clean
End Function

Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise 5, "T_R96Ribbon", detail
End Sub
