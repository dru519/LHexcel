Attribute VB_Name = "NxSymbolCatalog"
Option Explicit

Private Const NX_SYMBOL_SEARCH_LIMIT As Long = 480
Private mPractical As Collection
Private mCategories As Collection
Private mBlocks As Collection

Public Function NxSymbolPracticalCategories() As Collection
    EnsureSymbolCatalog
    Set NxSymbolPracticalCategories = CopyStringRows(mCategories)
End Function

Public Function NxSymbolUnicodeBlocks() As Collection
    EnsureSymbolCatalog
    Set NxSymbolUnicodeBlocks = CopyStringRows(mBlocks)
End Function

Public Function NxSymbolSearch(ByVal tabId As String, ByVal categoryId As String, ByVal query As String, ByVal recentRecords As Collection) As Collection
    Dim result As New Collection, seen As New Collection, record As CNxSymbolRecord
    Dim directCode As Long, hasDirect As Boolean, blockRow As Variant, parts As Variant, codePoint As Long
    EnsureSymbolCatalog
    tabId = LCase$(Trim$(tabId))
    If tabId <> "practical" And tabId <> "unicode" And tabId <> "recent" Then NxRaiseContractError "기호표 탭이 잘못되었습니다."
    hasDirect = TryParseDirectCode(query, directCode)
    If hasDirect Then AddUniqueSymbol result, seen, NxSymbolRecordForCode(directCode)
    If tabId = "practical" Then
        For Each record In mPractical
            If Len(categoryId) = 0 Or record.CategoryId = categoryId Then
                If SymbolMatches(record, query) Then AddUniqueSymbol result, seen, record
            End If
            If result.Count >= NX_SYMBOL_SEARCH_LIMIT Then Exit For
        Next record
    ElseIf tabId = "unicode" Then
        For Each blockRow In mBlocks
            parts = Split(CStr(blockRow), "|")
            If Len(categoryId) = 0 Or CStr(parts(0)) = categoryId Then
                For codePoint = CLng("&H" & CStr(parts(2))) To CLng("&H" & CStr(parts(3)))
                    If NxSymbolIsAllowedCodePoint(codePoint) And NxSymbolIsDisplayableCodePoint(codePoint) Then
                        Set record = NxSymbolRecordForCode(codePoint)
                        If SymbolMatches(record, query) Then AddUniqueSymbol result, seen, record
                        If result.Count >= NX_SYMBOL_SEARCH_LIMIT Then Exit For
                    End If
                Next codePoint
            End If
            If result.Count >= NX_SYMBOL_SEARCH_LIMIT Then Exit For
        Next blockRow
    ElseIf Not recentRecords Is Nothing Then
        For Each record In recentRecords
            If SymbolMatches(record, query) Then AddUniqueSymbol result, seen, record
            If result.Count >= NX_SYMBOL_SEARCH_LIMIT Then Exit For
        Next record
    End If
    Set NxSymbolSearch = result
End Function

Public Function NxSymbolRecordForCode(ByVal codePoint As Long) As CNxSymbolRecord
    Dim record As CNxSymbolRecord, blockId As String
    EnsureSymbolCatalog
    If Not NxSymbolIsAllowedCodePoint(codePoint) Then NxRaiseContractError "사용할 수 없는 유니코드 코드입니다."
    For Each record In mPractical
        If record.CodePoint = codePoint Then Set NxSymbolRecordForCode = record: Exit Function
    Next record
    blockId = BlockIdForCode(codePoint)
    Set record = New CNxSymbolRecord
    record.Configure codePoint, "유니코드 " & NxSymbolFormatCode(codePoint), "직접 입력;코드 검색", "direct", blockId, codePoint + 1
    Set NxSymbolRecordForCode = record
End Function

Public Function NxSymbolNormalizeQuery(ByVal query As String) As String
    query = LCase$(Trim$(query))
    query = Replace$(query, " ", vbNullString)
    query = Replace$(query, vbTab, vbNullString)
    query = Replace$(query, vbCr, vbNullString)
    query = Replace$(query, vbLf, vbNullString)
    If Left$(query, 2) = "u+" Then query = Mid$(query, 3)
    NxSymbolNormalizeQuery = query
End Function

Private Sub EnsureSymbolCatalog()
    Dim row As Variant, parts As Variant, record As CNxSymbolRecord
    If Not mPractical Is Nothing Then Exit Sub
    Set mCategories = New Collection
    For Each row In NxGeneratedSymbolCategoryRows()
        parts = Split(CStr(row), "|")
        If UBound(parts) <> 2 Then NxRaiseContractError "기호 범주 생성물이 잘못되었습니다."
        mCategories.Add CStr(parts(0)) & "|" & NxSymbolDecodeUtf16Hex(CStr(parts(2))) & "|" & CStr(parts(1))
    Next row
    Set mBlocks = New Collection
    For Each row In NxGeneratedSymbolBlockRows()
        parts = Split(CStr(row), "|")
        If UBound(parts) <> 4 Then NxRaiseContractError "유니코드 블록 생성물이 잘못되었습니다."
        mBlocks.Add CStr(parts(0)) & "|" & NxSymbolDecodeUtf16Hex(CStr(parts(4))) & "|" & CStr(parts(2)) & "|" & CStr(parts(3)) & "|" & CStr(parts(1))
    Next row
    Set mPractical = New Collection
    For Each row In NxGeneratedSymbolRows()
        parts = Split(CStr(row), "|")
        If UBound(parts) <> 5 Then NxRaiseContractError "기호 카탈로그 생성물이 잘못되었습니다."
        Set record = New CNxSymbolRecord
        record.Configure CLng("&H" & CStr(parts(3))), NxSymbolDecodeUtf16Hex(CStr(parts(4))), _
            NxSymbolDecodeUtf16Hex(CStr(parts(5))), CStr(parts(0)), CStr(parts(1)), CLng(parts(2))
        mPractical.Add record
    Next row
    If mCategories.Count <> 9 Or mBlocks.Count <> 14 Or mPractical.Count <> 190 Then NxRaiseContractError "기호 카탈로그 개수가 잘못되었습니다."
End Sub

Private Function CopyStringRows(ByVal source As Collection) As Collection
    Dim result As New Collection, item As Variant
    For Each item In source
        result.Add CStr(item)
    Next item
    Set CopyStringRows = result
End Function

Private Function SymbolMatches(ByVal record As CNxSymbolRecord, ByVal query As String) As Boolean
    Dim normalized As String, haystack As String
    normalized = NxSymbolNormalizeQuery(query)
    If Len(normalized) = 0 Then SymbolMatches = True: Exit Function
    haystack = NxSymbolNormalizeQuery(record.SearchText)
    SymbolMatches = InStr(1, haystack, normalized, vbBinaryCompare) > 0
End Function

Private Function TryParseDirectCode(ByVal query As String, ByRef codePoint As Long) As Boolean
    Dim normalized As String, index As Long, character As String
    normalized = NxSymbolNormalizeQuery(query)
    If Len(normalized) = 0 Or Len(normalized) > 6 Then Exit Function
    For index = 1 To Len(normalized)
        character = Mid$(normalized, index, 1)
        If InStr(1, "0123456789abcdef", character, vbBinaryCompare) = 0 Then Exit Function
    Next index
    On Error GoTo InvalidCode
    codePoint = CLng("&H" & normalized)
    On Error GoTo 0
    If Not NxSymbolIsAllowedCodePoint(codePoint) Then Exit Function
    TryParseDirectCode = True
    Exit Function
InvalidCode:
    Err.Clear
End Function

Private Sub AddUniqueSymbol(ByVal result As Collection, ByVal seen As Collection, ByVal record As CNxSymbolRecord)
    Dim key As String
    If record Is Nothing Or result.Count >= NX_SYMBOL_SEARCH_LIMIT Then Exit Sub
    key = "C" & CStr(record.CodePoint)
    On Error Resume Next
    seen.Add True, key
    If Err.Number = 0 Then result.Add record
    Err.Clear
    On Error GoTo 0
End Sub

Private Function BlockIdForCode(ByVal codePoint As Long) As String
    Dim row As Variant, parts As Variant
    For Each row In mBlocks
        parts = Split(CStr(row), "|")
        If codePoint >= CLng("&H" & CStr(parts(2))) And codePoint <= CLng("&H" & CStr(parts(3))) Then
            BlockIdForCode = CStr(parts(0))
            Exit Function
        End If
    Next row
    BlockIdForCode = "direct"
End Function
