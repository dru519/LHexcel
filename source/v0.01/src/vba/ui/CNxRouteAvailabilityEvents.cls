VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxRouteAvailabilityEvents"
Option Explicit

Private WithEvents ExcelApp As Application

Public Sub Configure(ByVal applicationObject As Application)
    Set ExcelApp = applicationObject
End Sub

Public Sub Disconnect()
    Set ExcelApp = Nothing
End Sub

Private Sub ExcelApp_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    NxRouteAvailabilityNotifyChanged
End Sub

Private Sub ExcelApp_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
    NxWorkbookCompareFollowHyperlink Sh, Target
End Sub

Private Sub ExcelApp_SheetActivate(ByVal Sh As Object)
    NxRouteAvailabilityNotifyChanged
End Sub

Private Sub ExcelApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    NxRouteAvailabilityNotifyChanged
End Sub

Private Sub ExcelApp_WorkbookActivate(ByVal Wb As Workbook)
    NxRouteAvailabilityNotifyChanged
End Sub

Private Sub ExcelApp_WorkbookOpen(ByVal Wb As Workbook)
    NxRouteAvailabilityNotifyChanged
End Sub

Private Sub ExcelApp_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    NxViewPresetForgetWorkbook Wb
    NxRouteAvailabilityNotifyChanged
End Sub

Private Sub ExcelApp_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    NxRouteAvailabilityNotifyChanged
End Sub
