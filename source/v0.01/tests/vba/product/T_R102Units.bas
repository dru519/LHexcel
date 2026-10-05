Attribute VB_Name = "T_R102Units"
Option Explicit

Public Function Names() As String
    Names = "catalog|search|glyphs|insert|recent|form"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim rows As Collection, emptyRecent As New Collection, recent As CNxSymbolRecent
    Dim book As Workbook, target As Range, prepared As CNxSymbolCellWrite
    Dim symbol As CNxSymbolRecord, chars As String, index As Long, detail As String
    Dim form As FNxSymbols
    On Error GoTo Failed
    chars = "㎟㎠㎡㎢㎣㎤㎥㎦㎜㎝㎞㎎㎏㎖㎘㏊²³"
    Select Case name
        Case "catalog"
            Set rows = NxSymbolSearch("practical", "", "", emptyRecent)
            Require rows.Count = 190, "catalog count"
            Set rows = NxSymbolSearch("practical", "unit_currency", "", emptyRecent)
            Require rows.Count = 34, "unit count"
            For index = 1 To Len(chars)
                Set symbol = NxSymbolRecordForCode(NxSymbolFirstCodePoint(Mid$(chars, index, 1)))
                Require symbol.CategoryId = "unit_currency", "unit category"
            Next index
        Case "search"
            Set rows = NxSymbolSearch("practical", "unit_currency", "면적", emptyRecent)
            Require rows.Count = 6, "area search"
            Set rows = NxSymbolSearch("practical", "unit_currency", "m2", emptyRecent)
            Require rows.Count = 4, "ascii alias search"
            Set rows = NxSymbolSearch("practical", "unit_currency", "㎡", emptyRecent)
            Require rows.Count = 1 And rows(1).CodePoint = &H33A1&, "symbol search"
            Set rows = NxSymbolSearch("unicode", "cjk_compatibility", "U+33A5", emptyRecent)
            Require rows.Count = 1 And rows(1).Character = "㎥", "unicode block search"
        Case "glyphs"
            For index = 1 To Len(chars)
                Require NxSymbolFontHasGlyph("맑은 고딕", Mid$(chars, index, 1)), "missing glyph " & CStr(index)
            Next index
        Case "insert"
            Set book = Workbooks.Add(xlWBATWorksheet)
            Set target = book.Worksheets(1).Range("B2:C3")
            target.Select
            Set prepared = NxSymbolPrepareActiveWrite("㎡", NxSymbolActiveSelectionIdentity())
            prepared.Commit
            Require Application.WorksheetFunction.CountIf(target, "㎡") = 4, "range insertion"
            target.Cells(1, 1).Select
            Set prepared = NxSymbolPrepareActiveWrite("㎥", NxSymbolActiveSelectionIdentity())
            prepared.Commit
            Require target.Cells(1, 1).Value2 = "㎡㎥", "text append"
            book.Close False: Set book = Nothing
        Case "recent"
            Set recent = New CNxSymbolRecent
            recent.Touch NxSymbolRecordForCode(&H33A1&)
            recent.Touch NxSymbolRecordForCode(&H33A5&)
            recent.Touch NxSymbolRecordForCode(&H33A1&)
            Require recent.Count = 2 And recent.Item(1).Character = "㎡", "recent unit history"
        Case "form"
            Set recent = New CNxSymbolRecent
            Set form = New FNxSymbols
            form.BindSession recent
            form.Controls("lstCategories").ListIndex = 3
            Require form.Controls("cmdGlyph19").Caption = "㎡", "area glyph absent from form"
            Require form.RuntimeActivateGlyph(19), "area glyph disabled"
            Require form.RuntimeBufferValue = "㎡", "area glyph buffer"
            Unload form: Set form = Nothing
        Case Else
            Err.Raise vbObjectError + 102, , "unknown unit case"
    End Select
    RunCase = "PASS|" & name
    Exit Function
Failed:
    detail = CStr(Err.Number) & ":" & Err.Description
    On Error Resume Next
    If Not form Is Nothing Then Unload form
    If Not book Is Nothing Then book.Close False
    RunCase = "FAIL|" & name & "|" & detail
End Function

Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise vbObjectError + 102, "UnitSymbols", detail
End Sub
