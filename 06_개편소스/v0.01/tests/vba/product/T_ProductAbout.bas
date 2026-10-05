Attribute VB_Name = "T_ProductAbout"
Option Explicit

Public Function Names() As String
    Names = "menu_and_form"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim panel As FNxAbout, xml As Object, items As Object, item As Object
    On Error GoTo Failed
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Require xml.LoadXML(NxRibbonManagementMenuXml()), "management XML"
    Set items = xml.DocumentElement.ChildNodes
    Set item = items.Item(items.Length - 1)
    Require item.getAttribute("label") = "내엑셀 정보", "last menu item"
    Require item.getAttribute("tag") = "nx1|entry|NX-MGMT-FILE-INFO", "about route"
    Set panel = New FNxAbout
    panel.BindProductInfo NxProductVersionText(), "2027-12-31"
    Require panel.Caption = "내엑셀 정보", "dialog title"
    Require panel.lblVersion.Caption = NxProductVersionText(), "current version"
    Require panel.lblExpiry.Caption = "사용기한: 2027-12-31", "expiry binding"
    Require panel.lblExpiry.Top + panel.lblExpiry.Height <= panel.cmdClose.Top, "expiry and button separated"
    Require panel.cmdClose.Top + panel.cmdClose.Height <= panel.InsideHeight, "close button fits"
    Require NxDistributionValidUntilText() = "제한 없음 (개발용)", "development expiry"
    Unload panel
    RunCase = "PASS|" & name
    Exit Function
Failed:
    RunCase = "FAIL|" & name & "|" & Err.Number & ":" & Err.Description
    On Error Resume Next
    If Not panel Is Nothing Then Unload panel
End Function

Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise vbObjectError + 105, , detail
End Sub
