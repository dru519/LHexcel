Attribute VB_Name = "T_Symbols"
Option Explicit

Public Function TestNames() As Collection
    Dim names As New Collection
    names.Add "TestCodePointBmpRoundTrip"
    names.Add "TestSupplementaryRoundTrip"
    names.Add "TestRejectSurrogate"
    names.Add "TestRejectControl"
    names.Add "TestRejectNoncharacter"
    names.Add "TestDropLastScalar"
    names.Add "TestPracticalCatalogCategoriesExact"
    names.Add "TestCatalogCodepointsUnique"
    names.Add "TestSearchByCharacter"
    names.Add "TestSearchByKoreanName"
    names.Add "TestSearchByAlias"
    names.Add "TestSearchByCategory"
    names.Add "TestUnitSearch"
    names.Add "TestSearchByUPlus"
    names.Add "TestSearchByBareHex"
    names.Add "TestDirectValidCodeCreatesRecord"
    names.Add "TestRecentLruKeepsSixteen"
    names.Add "TestRecentMovesDuplicateToFront"
    names.Add "TestBufferAddAddsRecent"
    names.Add "TestEmptyCellWrite"
    names.Add "TestTextCellAppend"
    names.Add "TestContiguousSelectionWrite"
    names.Add "TestUnsafeCellKindsRejected"
    names.Add "TestTargetDriftRejected"
    names.Add "TestFontFallbackAndFormContract"
    Set TestNames = names
End Function

Public Sub RunCase(ByVal name As String)
    Select Case name
        Case "TestCodePointBmpRoundTrip": TestCodePointBmpRoundTrip
        Case "TestSupplementaryRoundTrip": TestSupplementaryRoundTrip
        Case "TestRejectSurrogate": TestRejectSurrogate
        Case "TestRejectControl": TestRejectControl
        Case "TestRejectNoncharacter": TestRejectNoncharacter
        Case "TestDropLastScalar": TestDropLastScalar
        Case "TestPracticalCatalogCategoriesExact": TestPracticalCatalogCategoriesExact
        Case "TestCatalogCodepointsUnique": TestCatalogCodepointsUnique
        Case "TestSearchByCharacter": TestSearchByCharacter
        Case "TestSearchByKoreanName": TestSearchByKoreanName
        Case "TestSearchByAlias": TestSearchByAlias
        Case "TestSearchByCategory": TestSearchByCategory
        Case "TestUnitSearch": TestUnitSearch
        Case "TestSearchByUPlus": TestSearchByUPlus
        Case "TestSearchByBareHex": TestSearchByBareHex
        Case "TestDirectValidCodeCreatesRecord": TestDirectValidCodeCreatesRecord
        Case "TestRecentLruKeepsSixteen": TestRecentLruKeepsSixteen
        Case "TestRecentMovesDuplicateToFront": TestRecentMovesDuplicateToFront
        Case "TestBufferAddAddsRecent": TestBufferAddAddsRecent
        Case "TestEmptyCellWrite": TestEmptyCellWrite
        Case "TestTextCellAppend": TestTextCellAppend
        Case "TestContiguousSelectionWrite": TestContiguousSelectionWrite
        Case "TestUnsafeCellKindsRejected": TestUnsafeCellKindsRejected
        Case "TestTargetDriftRejected": TestTargetDriftRejected
        Case "TestFontFallbackAndFormContract": TestFontFallbackAndFormContract
        Case Else: NxRaiseContractError "Unknown Symbols test"
    End Select
End Sub

Private Function EmptyRecent() As Collection
    Set EmptyRecent = New Collection
End Function

Private Function ActiveTestBook() As Workbook
    Dim book As Workbook
    Set book = Application.ActiveWorkbook
    If book Is Nothing Or book Is ThisWorkbook Then NxRaiseContractError "Symbols test data workbook is unavailable"
    Set ActiveTestBook = book
End Function

Private Sub TestCodePointBmpRoundTrip()
    Dim value As String
    value = NxSymbolFromCodePoint(&H2713&)
    NxTestHarness.AssertTrue NxSymbolFirstCodePoint(value) = &H2713& And NxSymbolFormatCode(&H2713&) = "U+2713", "BMP Unicode round trip changed"
End Sub

Private Sub TestSupplementaryRoundTrip()
    Dim value As String
    value = NxSymbolFromCodePoint(&H1F600)
    NxTestHarness.AssertTrue Len(value) = 2 And NxSymbolFirstCodePoint(value) = &H1F600 And NxSymbolFormatCode(&H1F600) = "U+1F600", "Supplementary Unicode round trip changed"
End Sub

Private Sub TestRejectSurrogate()
    NxTestHarness.AssertTrue Not NxSymbolIsAllowedCodePoint(&HD800&) And Not NxSymbolIsAllowedCodePoint(&HDFFF&), "Surrogate code point was accepted"
End Sub

Private Sub TestRejectControl()
    NxTestHarness.AssertTrue Not NxSymbolIsAllowedCodePoint(0) And Not NxSymbolIsAllowedCodePoint(&H1F&) And _
        Not NxSymbolIsAllowedCodePoint(&H7F&) And Not NxSymbolIsAllowedCodePoint(&H9F&), "Control code point was accepted"
End Sub

Private Sub TestRejectNoncharacter()
    NxTestHarness.AssertTrue Not NxSymbolIsAllowedCodePoint(&HFDD0&) And Not NxSymbolIsAllowedCodePoint(&HFDEF&) And _
        Not NxSymbolIsAllowedCodePoint(&HFFFE&) And Not NxSymbolIsAllowedCodePoint(&HFFFF&) And _
        Not NxSymbolIsAllowedCodePoint(&H1FFFE) And Not NxSymbolIsAllowedCodePoint(&H10FFFF), "Unicode noncharacter was accepted"
End Sub

Private Sub TestDropLastScalar()
    Dim value As String
    value = "A" & NxSymbolFromCodePoint(&H1F600)
    NxTestHarness.AssertTrue NxSymbolDropLastScalar(value) = "A" And NxSymbolDropLastScalar("A") = vbNullString, "Unicode scalar deletion changed"
End Sub

Private Sub TestPracticalCatalogCategoriesExact()
    Dim rows As Collection
    Set rows = NxSymbolPracticalCategories()
    NxTestHarness.AssertTrue rows.Count = 9 And CStr(rows(1)) = "punctuation|문장부호|10" And _
        CStr(rows(9)) = "other|기타|90", "Practical symbol categories changed"
End Sub

Private Sub TestCatalogCodepointsUnique()
    Dim records As Collection, seen As New Collection, record As CNxSymbolRecord, duplicate As Boolean
    Set records = NxSymbolSearch("practical", vbNullString, vbNullString, EmptyRecent())
    For Each record In records
        On Error Resume Next
        seen.Add True, "C" & CStr(record.CodePoint)
        If Err.Number <> 0 Then duplicate = True
        Err.Clear
        On Error GoTo 0
    Next record
    NxTestHarness.AssertTrue records.Count = 190 And seen.Count = 190 And Not duplicate, "Practical symbol code points are not an exact unique 190"
End Sub

Private Sub TestSearchByCharacter()
    Dim records As Collection
    Set records = NxSymbolSearch("practical", vbNullString, NxSymbolFromCodePoint(&H203B&), EmptyRecent())
    NxTestHarness.AssertTrue records.Count = 1 And records(1).CodePoint = &H203B&, "Character search changed"
End Sub

Private Sub TestSearchByKoreanName()
    Dim records As Collection
    Set records = NxSymbolSearch("practical", vbNullString, "가운뎃점", EmptyRecent())
    NxTestHarness.AssertTrue records.Count = 1 And records(1).CodePoint = &HB7&, "Korean name search changed"
End Sub

Private Sub TestSearchByAlias()
    Dim records As Collection
    Set records = NxSymbolSearch("practical", vbNullString, "당구장", EmptyRecent())
    NxTestHarness.AssertTrue records.Count = 1 And records(1).CodePoint = &H203B&, "Korean alias search changed"
End Sub

Private Sub TestSearchByCategory()
    Dim records As Collection, record As CNxSymbolRecord, valid As Boolean
    Set records = NxSymbolSearch("practical", "arrows", vbNullString, EmptyRecent())
    valid = records.Count = 16
    For Each record In records
        If record.CategoryId <> "arrows" Then valid = False
    Next record
    NxTestHarness.AssertTrue valid, "Category search changed"
End Sub

Private Sub TestSearchByUPlus()
    Dim records As Collection
    Set records = NxSymbolSearch("practical", vbNullString, "U+2192", EmptyRecent())
    NxTestHarness.AssertTrue records.Count = 1 And records(1).CodePoint = &H2192&, "U+ search changed"
End Sub

Private Sub TestUnitSearch()
    Dim records As Collection
    Set records = NxSymbolSearch("practical", "unit_currency", "제곱미터", EmptyRecent())
    NxTestHarness.AssertTrue records.Count >= 1, "Square metre name missing"
    Set records = NxSymbolSearch("practical", "unit_currency", "m2", EmptyRecent())
    NxTestHarness.AssertTrue records.Count = 4, "Area aliases missing"
    Set records = NxSymbolSearch("practical", "unit_currency", NxSymbolFromCodePoint(&H33A1&), EmptyRecent())
    NxTestHarness.AssertTrue records.Count = 1 And records(1).CodePoint = &H33A1&, "Square metre character missing"
    Set records = NxSymbolSearch("practical", "unit_currency", "면적", EmptyRecent())
    NxTestHarness.AssertTrue records.Count = 6, "Area keyword missing"
End Sub

Private Sub TestSearchByBareHex()
    Dim records As Collection
    Set records = NxSymbolSearch("practical", vbNullString, "2192", EmptyRecent())
    NxTestHarness.AssertTrue records.Count = 1 And records(1).CodePoint = &H2192&, "Bare hexadecimal search changed"
End Sub

Private Sub TestDirectValidCodeCreatesRecord()
    Dim records As Collection, record As CNxSymbolRecord
    Set records = NxSymbolSearch("practical", vbNullString, "1F600", EmptyRecent())
    Set record = records(1)
    NxTestHarness.AssertTrue record.CodePoint = &H1F600 And record.CodeText = "U+1F600" And record.BlockId = "direct", "Direct valid code record changed"
End Sub

Private Sub TestRecentLruKeepsSixteen()
    Dim recent As New CNxSymbolRecent, index As Long
    For index = 0 To 16
        recent.Touch NxSymbolRecordForCode(&H2190& + index)
    Next index
    NxTestHarness.AssertTrue recent.Count = 16 And recent.Item(1).CodePoint = &H21A0& And recent.Item(16).CodePoint = &H2191&, "Recent LRU limit changed"
End Sub

Private Sub TestRecentMovesDuplicateToFront()
    Dim recent As New CNxSymbolRecent
    recent.Touch NxSymbolRecordForCode(&H2190&)
    recent.Touch NxSymbolRecordForCode(&H2191&)
    recent.Touch NxSymbolRecordForCode(&H2190&)
    NxTestHarness.AssertTrue recent.Count = 2 And recent.Item(1).CodePoint = &H2190& And recent.Item(2).CodePoint = &H2191&, "Recent duplicate ordering changed"
End Sub

Private Sub TestBufferAddAddsRecent()
    Dim recent As New CNxSymbolRecent, form As FNxSymbols, detail As String
    On Error GoTo Failed
    Set form = New FNxSymbols
    form.BindSession recent
    form.HandleGlyphActivate 1
    NxTestHarness.AssertTrue recent.Count = 1 And Len(CStr(form.Controls("txtBuffer").Value)) > 0, "Buffer add did not update session recent symbols"
    Unload form
    Exit Sub
Failed:
    detail = Err.Description
    On Error Resume Next: If Not form Is Nothing Then Unload form: On Error GoTo 0
    NxRaiseContractError detail
End Sub

Private Sub TestEmptyCellWrite()
    Dim book As Workbook, target As Range, prepared As CNxSymbolCellWrite, identity As String, detail As String
    On Error GoTo Failed
    Set book = ActiveTestBook(): Set target = book.Worksheets(1).Range("B2")
    target.Clear: target.NumberFormat = "General": target.Select
    identity = NxSymbolActiveCellIdentity()
    Set prepared = NxSymbolPrepareActiveWrite(NxSymbolFromCodePoint(&H2713&), identity)
    prepared.Commit
    NxTestHarness.AssertTrue CStr(target.Value2) = NxSymbolFromCodePoint(&H2713&), "Empty cell symbol write changed"
    target.Clear: book.Worksheets(1).Range("A1").Select
    Exit Sub
Failed:
    detail = Err.Description: On Error Resume Next: If Not target Is Nothing Then target.Clear: target.NumberFormat = "General": If Not book Is Nothing Then book.Worksheets(1).Range("A1").Select: On Error GoTo 0: NxRaiseContractError detail
End Sub

Private Sub TestTextCellAppend()
    Dim book As Workbook, target As Range, prepared As CNxSymbolCellWrite, identity As String, detail As String
    On Error GoTo Failed
    Set book = ActiveTestBook(): Set target = book.Worksheets(1).Range("B2")
    target.Clear: target.NumberFormat = "@": target.Value2 = "앞": target.Select
    identity = NxSymbolActiveCellIdentity()
    Set prepared = NxSymbolPrepareActiveWrite(NxSymbolFromCodePoint(&H2192&), identity)
    prepared.Commit
    NxTestHarness.AssertTrue CStr(target.Value2) = "앞" & NxSymbolFromCodePoint(&H2192&), "Text cell append changed"
    target.Clear: target.NumberFormat = "General": book.Worksheets(1).Range("A1").Select
    Exit Sub
Failed:
    detail = Err.Description: On Error Resume Next: If Not target Is Nothing Then target.Clear: target.NumberFormat = "General": If Not book Is Nothing Then book.Worksheets(1).Range("A1").Select: On Error GoTo 0: NxRaiseContractError detail
End Sub

Private Sub TestContiguousSelectionWrite()
    Dim book As Workbook, target As Range, prepared As CNxSymbolCellWrite, identity As String, detail As String
    On Error GoTo Failed
    Set book = ActiveTestBook(): Set target = book.Worksheets(1).Range("B2:C3")
    target.Clear: target.NumberFormat = "General": target.Select
    identity = NxSymbolActiveSelectionIdentity()
    Set prepared = NxSymbolPrepareActiveWrite(NxSymbolFromCodePoint(&H2460&), identity)
    prepared.Commit
    NxTestHarness.AssertTrue prepared.TargetCount = 4 And Application.WorksheetFunction.CountIf(target, NxSymbolFromCodePoint(&H2460&)) = 4, _
        "Contiguous selection symbol write changed"
    target.Clear: book.Worksheets(1).Range("A1").Select
    Exit Sub
Failed:
    detail = Err.Description: On Error Resume Next: If Not target Is Nothing Then target.Clear: If Not book Is Nothing Then book.Worksheets(1).Range("A1").Select: On Error GoTo 0: NxRaiseContractError detail
End Sub

Private Sub TestUnsafeCellKindsRejected()
    Dim book As Workbook, target As Range, rejected As Long, identity As String, detail As String
    On Error GoTo Failed
    Set book = ActiveTestBook(): Set target = book.Worksheets(1).Range("B2"): target.Select

    target.Formula = "=1+1": identity = NxSymbolActiveCellIdentity(): On Error Resume Next: Call NxSymbolPrepareActiveWrite("x", identity): If Err.Number <> 0 Then rejected = rejected + 1: Err.Clear: On Error GoTo Failed
    target.Clear: target.Value2 = 123#: identity = NxSymbolActiveCellIdentity(): On Error Resume Next: Call NxSymbolPrepareActiveWrite("x", identity): If Err.Number <> 0 Then rejected = rejected + 1: Err.Clear: On Error GoTo Failed
    target.Clear: target.Value2 = True: identity = NxSymbolActiveCellIdentity(): On Error Resume Next: Call NxSymbolPrepareActiveWrite("x", identity): If Err.Number <> 0 Then rejected = rejected + 1: Err.Clear: On Error GoTo Failed
    target.Clear: target.Value = DateSerial(2026, 8, 9): target.NumberFormat = "yyyy-mm-dd": identity = NxSymbolActiveCellIdentity(): On Error Resume Next: Call NxSymbolPrepareActiveWrite("x", identity): If Err.Number <> 0 Then rejected = rejected + 1: Err.Clear: On Error GoTo Failed
    target.Clear: target.NumberFormat = "General": target.Value = CVErr(xlErrNA): identity = NxSymbolActiveCellIdentity(): On Error Resume Next: Call NxSymbolPrepareActiveWrite("x", identity): If Err.Number <> 0 Then rejected = rejected + 1: Err.Clear: On Error GoTo Failed

    NxTestHarness.AssertTrue rejected = 5, "Unsafe cell kinds were not all rejected"
    target.Clear: target.NumberFormat = "General": book.Worksheets(1).Range("A1").Select
    Exit Sub
Failed:
    detail = Err.Description: On Error Resume Next: If Not target Is Nothing Then target.Clear: target.NumberFormat = "General": If Not book Is Nothing Then book.Worksheets(1).Range("A1").Select: On Error GoTo 0: NxRaiseContractError detail
End Sub

Private Sub TestTargetDriftRejected()
    Dim book As Workbook, prepared As CNxSymbolCellWrite, failed As Boolean, identity As String, detail As String
    On Error GoTo Failed
    Set book = ActiveTestBook()
    book.Worksheets(1).Range("B2:C2").Clear: book.Worksheets(1).Range("B2").Select
    identity = NxSymbolActiveCellIdentity()
    Set prepared = NxSymbolPrepareActiveWrite("x", identity)
    book.Worksheets(1).Range("C2").Select
    On Error Resume Next: prepared.Commit: failed = Err.Number <> 0: Err.Clear: On Error GoTo Failed
    NxTestHarness.AssertTrue failed And IsEmpty(book.Worksheets(1).Range("B2").Value2), "Symbol target drift was not rejected"
    book.Worksheets(1).Range("B2:C2").Clear: book.Worksheets(1).Range("A1").Select
    Exit Sub
Failed:
    detail = Err.Description: On Error Resume Next: If Not book Is Nothing Then book.Worksheets(1).Range("B2:C2").Clear: book.Worksheets(1).Range("A1").Select: On Error GoTo 0: NxRaiseContractError detail
End Sub

Private Sub TestFontFallbackAndFormContract()
    Dim canDisplay As Boolean, form As FNxSymbols, registry As CNxFeatureRegistry, definition As CNxFeatureDefinition, detail As String
    On Error GoTo Failed
    NxTestHarness.AssertTrue NxSymbolResolveDisplayFont("NX-MISSING-FONT-7F2A", &H2713&, canDisplay) = "Segoe UI Symbol", "Symbol font fallback changed"
    Set registry = CreateBaseRegistry(): NxSymbolsRegisterFeatures registry
    Set definition = registry.FeatureById("NX-UTIL-SYMBOLS")
    NxTestHarness.AssertTrue definition.CategoryId = "NX-CAT-SYMBOLS" And definition.LaunchSurface = "direct", "Symbols registry contract changed"
    Set form = New FNxSymbols
    NxTestHarness.AssertTrue form.RuntimeGlyphCount = 50 And form.GlyphControlName(50) = "cmdGlyph50" And _
        form.Controls("cmdInsertCell").Default And form.Controls("cmdClose").Cancel, "Symbols form contract changed"
    Unload form
    Exit Sub
Failed:
    detail = Err.Description: On Error Resume Next: If Not form Is Nothing Then Unload form: On Error GoTo 0: NxRaiseContractError detail
End Sub
