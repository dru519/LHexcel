Attribute VB_Name = "NxTemplatePrint"
Option Explicit

Public Function NxTemplatePrintFingerprint(ByVal sheet As Worksheet) As String
    Dim key As Variant, text As String, value As String, setup As Object
    Set setup = sheet.PageSetup
    For Each key In Array("PaperSize", "Orientation", "LeftMargin", "RightMargin", "TopMargin", "BottomMargin", _
        "HeaderMargin", "FooterMargin", "CenterHorizontally", "CenterVertically", "Order", "BlackAndWhite", _
        "PrintGridlines", "PrintHeadings", "Zoom", "FitToPagesWide", "FitToPagesTall", "PrintArea", "PrintTitleRows", "PrintTitleColumns")
        value = CStr(CallByName(setup, CStr(key), VbGet))
        text = text & CStr(Len(value)) & ":" & value & "|"
    Next key
    NxTemplatePrintFingerprint = NxTemplateTextSha256(text)
End Function

Public Sub NxTemplateCopyPrintSettings(ByVal source As Range, ByVal target As Worksheet)
    Dim original As PageSetup, result As PageSetup
    Set original = source.Worksheet.PageSetup
    Set result = target.PageSetup
    NxTemplateProgress "인쇄 설정 복원", 0, 1
    With result
        .PaperSize = original.PaperSize
        .Orientation = original.Orientation
        .LeftMargin = original.LeftMargin
        .RightMargin = original.RightMargin
        .TopMargin = original.TopMargin
        .BottomMargin = original.BottomMargin
        .HeaderMargin = original.HeaderMargin
        .FooterMargin = original.FooterMargin
        .CenterHorizontally = original.CenterHorizontally
        .CenterVertically = original.CenterVertically
        .Order = original.Order
        .BlackAndWhite = original.BlackAndWhite
        .PrintGridlines = original.PrintGridlines
        .PrintHeadings = original.PrintHeadings
        .Zoom = False
        .FitToPagesWide = original.FitToPagesWide
        .FitToPagesTall = original.FitToPagesTall
        .Zoom = original.Zoom
        .PrintArea = MappedPrintRange(original.PrintArea, source, target, "area")
        .PrintTitleRows = MappedPrintRange(original.PrintTitleRows, source, target, "rows")
        .PrintTitleColumns = MappedPrintRange(original.PrintTitleColumns, source, target, "columns")
    End With
End Sub

Private Function MappedPrintRange(ByVal address As String, ByVal source As Range, ByVal target As Worksheet, ByVal kind As String) As String
    Dim selected As Range, overlap As Range, area As Range, mapped As Range, combined As Range
    If Len(address) = 0 Then Exit Function
    Set selected = source.Worksheet.Range(address)
    Set overlap = Application.Intersect(selected, source)
    If overlap Is Nothing Then Exit Function
    For Each area In overlap.Areas
        Set mapped = target.Cells(area.Row - source.Row + 1, area.Column - source.Column + 1).Resize(area.Rows.Count, area.Columns.Count)
        If kind = "rows" Then Set mapped = mapped.EntireRow
        If kind = "columns" Then Set mapped = mapped.EntireColumn
        If combined Is Nothing Then Set combined = mapped Else Set combined = Application.Union(combined, mapped)
    Next area
    MappedPrintRange = combined.Address(True, True, xlA1, False)
End Function

Public Function NxTemplatePrintNamesOnly(ByVal book As Workbook) As Boolean
    Dim item As Name, localName As String, referenced As Range
    On Error GoTo Rejected
    For Each item In book.Names
        localName = Mid$(item.Name, InStrRev(item.Name, "!") + 1)
        If localName <> "Print_Area" And localName <> "Print_Titles" Then Exit Function
        If InStr(item.Name, "!") = 0 Then Exit Function
        Set referenced = Nothing
        Set referenced = item.RefersToRange
        If referenced Is Nothing Then Exit Function
        If Not referenced.Worksheet.Parent Is book Then Exit Function
        If localName = "Print_Area" Then
            If Len(referenced.Worksheet.PageSetup.PrintArea) = 0 Then Exit Function
        Else
            If Len(referenced.Worksheet.PageSetup.PrintTitleRows) = 0 And Len(referenced.Worksheet.PageSetup.PrintTitleColumns) = 0 Then Exit Function
        End If
    Next item
    NxTemplatePrintNamesOnly = True
Rejected:
End Function
