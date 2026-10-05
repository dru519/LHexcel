VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxDocumentNavigatorItem"
Option Explicit

' Session-only references. A sheet points to its book entry, never the reverse.
Public TokenId As Long
Public BookRef As Excel.Workbook
Public SheetRef As Object
Public ParentBook As CNxDocumentNavigatorItem
Public RecentWindow As Excel.Window
Public DisplayName As String
Public StateText As String
Public Available As Boolean
