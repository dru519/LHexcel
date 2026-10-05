VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxLimitsPolicy"
Option Explicit

' Exact generated feature-contract.json and limits.json projection.  No prefixes are authorization.
Public Function ExactOwnerFor(ByVal featureId As String) As String
    Select Case featureId
        Case "NX-TEST-CORE": ExactOwnerFor = "P01"
        Case "NX-AI-SUMMARY", "NX-AI-CLEAN", "NX-AI-FORMULA", "NX-AI-WRITE", "NX-AI-IMAGE": ExactOwnerFor = "ai"
        Case "NX-TPL-REGISTER-SHEET", "NX-TPL-LIST", "NX-TPL-LOAD", "NX-TPL-RENAME", "NX-TPL-DELETE": ExactOwnerFor = "template"
        Case "NX-DATA-NORMALIZE", "NX-DATA-UNIQUE-COUNT", "NX-DATA-DUPLICATE-LIST", "NX-DATA-AGE", "NX-DATA-KOREAN-MONEY", "NX-DATA-PRIVACY-MASK", "NX-DATA-PRIVACY-SCAN", "NX-DATA-FOCUS-CELL", "NX-DATA-COPY-VISIBLE", "NX-DATA-PASTE-VISIBLE-VALUES": ExactOwnerFor = "data"
        Case "NX-HANGUL-TABLE-SEND", "NX-HANGUL-PICTURE-SEND": ExactOwnerFor = "hangul"
        Case "NX-DRAW-TITLE-TABLE", "NX-DRAW-BUSINESS-TABLE", "NX-DRAW-ROLE-STYLE", "NX-DRAW-CLEAR-INNER", "NX-DRAW-FIT-PICTURE", "NX-DRAW-INSERT-PICTURE": ExactOwnerFor = "draw"
        Case "NX-FILE-FOLDER-CREATE", "NX-FILE-CONSOLIDATE", "NX-FILE-MANNER-SAVE", _
             "NX-FILE-RANGE-PNG", "NX-FILE-CHART-PNG", "NX-FILE-SHEET-COPY-SAVE", _
             "NX-FILE-RANGE-COPY-SAVE", "NX-FILE-WORKBOOK-COMPARE", "NX-FILE-SHEET-COMPARE", "NX-FILE-FILE-COMPARE", "NX-FILE-BATCH-RENAME", _
             "NX-FILE-SHEET-BATCH-RENAME", "NX-FILE-PDF-CURRENT-SHEET", _
             "NX-FILE-PDF-EACH-SHEET", "NX-FILE-PDF-SELECTED-COMBINED", _
             "NX-FILE-PDF-ALL-COMBINED", "NX-FILE-PDF-SETTINGS": ExactOwnerFor = "file"
        Case "NX-UTIL-CALCULATOR": ExactOwnerFor = "calculator"
        Case "NX-UTIL-SYMBOLS": ExactOwnerFor = "symbols"
        Case "NX-UTIL-NAVIGATOR", "NX-UTIL-DOCUMENT-NAVIGATOR": ExactOwnerFor = "navigator"
        Case Else: NxRaiseContractError "Unknown feature owner"
    End Select
End Function

Public Function MaximumCellsFor(ByVal featureId As String) As Long
    Select Case featureId
        Case "NX-TEST-CORE": MaximumCellsFor = 200000
        Case "NX-AI-SUMMARY", "NX-AI-CLEAN", "NX-AI-FORMULA", "NX-AI-WRITE", "NX-AI-IMAGE": MaximumCellsFor = 50000
        Case "NX-HANGUL-TABLE-SEND", "NX-HANGUL-PICTURE-SEND": MaximumCellsFor = 16384
        Case "NX-TPL-REGISTER-SHEET", "NX-TPL-LIST", "NX-TPL-LOAD", "NX-TPL-RENAME", "NX-TPL-DELETE", "NX-DRAW-TITLE-TABLE", "NX-DRAW-BUSINESS-TABLE", "NX-DRAW-ROLE-STYLE", "NX-DRAW-CLEAR-INNER", "NX-DRAW-FIT-PICTURE", "NX-DRAW-INSERT-PICTURE": MaximumCellsFor = 100000
        Case "NX-DATA-NORMALIZE", "NX-DATA-UNIQUE-COUNT", "NX-DATA-DUPLICATE-LIST", "NX-DATA-AGE", "NX-DATA-KOREAN-MONEY", "NX-DATA-PRIVACY-MASK", "NX-DATA-PRIVACY-SCAN", "NX-DATA-FOCUS-CELL", "NX-DATA-COPY-VISIBLE", "NX-DATA-PASTE-VISIBLE-VALUES": MaximumCellsFor = 200000
        Case "NX-FILE-FOLDER-CREATE", "NX-FILE-CONSOLIDATE", "NX-FILE-MANNER-SAVE", _
             "NX-FILE-SHEET-COPY-SAVE", "NX-FILE-RANGE-COPY-SAVE", _
             "NX-FILE-BATCH-RENAME", "NX-FILE-SHEET-BATCH-RENAME", _
             "NX-FILE-PDF-CURRENT-SHEET", "NX-FILE-PDF-EACH-SHEET", _
             "NX-FILE-PDF-SELECTED-COMBINED", "NX-FILE-PDF-ALL-COMBINED", _
             "NX-FILE-PDF-SETTINGS": MaximumCellsFor = 500000
        Case "NX-FILE-WORKBOOK-COMPARE", "NX-FILE-SHEET-COMPARE", "NX-FILE-FILE-COMPARE": MaximumCellsFor = 1000000
        Case "NX-FILE-RANGE-PNG", "NX-FILE-CHART-PNG": MaximumCellsFor = 10000
        Case "NX-UTIL-CALCULATOR", "NX-UTIL-SYMBOLS", "NX-UTIL-NAVIGATOR", "NX-UTIL-DOCUMENT-NAVIGATOR": MaximumCellsFor = 10000
        Case Else: NxRaiseContractError "Unknown feature limit"
    End Select
End Function

Public Function MaximumAreasFor(ByVal featureId As String) As Long
    If ExactOwnerFor(featureId) = "ai" Then MaximumAreasFor = 20 Else MaximumAreasFor = 0
End Function

Public Function MaximumPromptCharsFor(ByVal featureId As String) As Long
    If ExactOwnerFor(featureId) = "ai" Then MaximumPromptCharsFor = 50000 Else MaximumPromptCharsFor = 0
End Function

Public Function MaximumPicturesFor(ByVal featureId As String) As Long
    If featureId = "NX-DRAW-INSERT-PICTURE" Then
        MaximumPicturesFor = 500
    ElseIf ExactOwnerFor(featureId) = "draw" Or featureId = "NX-AI-IMAGE" Then
        MaximumPicturesFor = 1
    Else
        MaximumPicturesFor = 0
    End If
End Function

Public Function MaximumDirectoryItemsFor(ByVal featureId As String) As Long
    If featureId = "NX-FILE-FOLDER-CREATE" Then MaximumDirectoryItemsFor = 500 Else MaximumDirectoryItemsFor = 0
End Function

Public Function MaximumInputFilesFor(ByVal featureId As String) As Long
    If featureId = "NX-FILE-CONSOLIDATE" Then MaximumInputFilesFor = 0: Exit Function
    If ExactOwnerFor(featureId) = "file" Then MaximumInputFilesFor = 20 Else MaximumInputFilesFor = 0
End Function

Public Function MaximumSheetsFor(ByVal featureId As String) As Long
    If ExactOwnerFor(featureId) = "file" Then MaximumSheetsFor = 100 Else MaximumSheetsFor = 0
End Function

Public Function MaximumBytesFor(ByVal featureId As String) As Double
    If ExactOwnerFor(featureId) = "file" Then MaximumBytesFor = 104857600# Else MaximumBytesFor = 0#
End Function

Public Function MaximumPngDimensionFor(ByVal featureId As String) As Long
    Select Case featureId
        Case "NX-FILE-RANGE-PNG", "NX-FILE-CHART-PNG": MaximumPngDimensionFor = 16384
        Case Else: MaximumPngDimensionFor = 0
    End Select
End Function

Public Function MaximumPngCellsFor(ByVal featureId As String) As Long
    Select Case featureId
        Case "NX-FILE-RANGE-PNG", "NX-FILE-CHART-PNG": MaximumPngCellsFor = 10000
        Case Else: MaximumPngCellsFor = 0
    End Select
End Function

Public Function PowerShellTimeoutSecondsFor(ByVal featureId As String) As Long
    PowerShellTimeoutSecondsFor = 0
End Function

Public Function CancelCheckIntervalRows() As Long
    CancelCheckIntervalRows = 1024
End Function

Public Sub ValidateEffectFor(ByVal featureId As String, ByVal effect As CNxSideEffect)
    If effect Is Nothing Then NxRaiseContractError "Policy requires an exact side effect"
    If effect.Owner <> ExactOwnerFor(featureId) Then NxRaiseContractError "Policy rejected effect owner"
    If effect.LimitDecision <> MaximumCellsFor(featureId) Then NxRaiseContractError "Policy rejected effect cell limit"
    If effect.Kind = NxCellMutation Then
        If effect.MaximumCells <= 0 Or effect.MaximumCells > MaximumCellsFor(featureId) Then NxRaiseContractError "Policy rejected effect maximum cells"
        If Len(effect.Target) = 0 Then NxRaiseContractError "Policy rejected empty cell target"
    ElseIf effect.Kind = NxWorksheetCreate Or effect.Kind = NxColumnInsert Then
        If effect.MaximumCells <= 0 Or effect.MaximumCells > MaximumCellsFor(featureId) Or Len(effect.Target) = 0 Then NxRaiseContractError "Policy rejected worksheet creation"
    ElseIf effect.Kind = NxBinaryPackageWrite Then
        If effect.MaximumCells <> 0 Or Len(effect.Target) = 0 Then NxRaiseContractError "Policy rejected binary package write"
    ElseIf Len(effect.Target) = 0 Or effect.MaximumCells <> 0 Then
        NxRaiseContractError "Policy rejected non-cell effect target or limit"
    End If
    If effect.Kind = NxColumnInsert Then
        Select Case featureId
            Case "NX-DATA-NORMALIZE", "NX-DATA-AGE", "NX-DATA-KOREAN-MONEY", "NX-DATA-PRIVACY-MASK"
            Case Else: NxRaiseContractError "Column insertion is restricted to data conversion"
        End Select
    End If
    If effect.Kind = NxShapeInsert Then
        If effect.DeclaredItems <= 0 Or effect.DeclaredItems > MaximumPicturesFor(featureId) Then NxRaiseContractError "Policy rejected shape insertion"
    End If
    If effect.Kind = NxDirectoryBatchCreate Then
        If effect.DeclaredItems <= 0 Or effect.DeclaredItems > MaximumDirectoryItemsFor(featureId) Then NxRaiseContractError "Policy rejected directory batch"
        If effect.DirectoryItems.Count <> effect.DeclaredItems Then NxRaiseContractError "Policy rejected directory batch item count"
    End If
    If MaximumPngCellsFor(featureId) > 0 And effect.MaximumCells > MaximumPngCellsFor(featureId) Then NxRaiseContractError "Policy rejected PNG max_cells"
End Sub

Public Function SourceDigest() As String
    ' Generated exact projection of feature-contract.json and limits.json; no prefix inference/static digest.
    Dim index As Long, featureId As Variant, canonical As String
    Dim ids As Variant
    ' BEGIN GENERATED RELEASE FEATURE IDS (generate_product_surface.py)
    ids = Array( _
        "NX-AI-SUMMARY", "NX-AI-CLEAN", "NX-AI-FORMULA", "NX-AI-WRITE", "NX-AI-IMAGE", _
        "NX-TPL-REGISTER-SHEET", "NX-TPL-LIST", "NX-TPL-LOAD", "NX-TPL-RENAME", "NX-TPL-DELETE", _
        "NX-DATA-NORMALIZE", "NX-DATA-UNIQUE-COUNT", "NX-DATA-DUPLICATE-LIST", "NX-DATA-AGE", "NX-DATA-KOREAN-MONEY", _
        "NX-DATA-PRIVACY-MASK", "NX-DATA-PRIVACY-SCAN", "NX-DATA-FOCUS-CELL", "NX-DATA-COPY-VISIBLE", "NX-DATA-PASTE-VISIBLE-VALUES", _
        "NX-HANGUL-TABLE-SEND", "NX-HANGUL-PICTURE-SEND", "NX-DRAW-TITLE-TABLE", "NX-DRAW-BUSINESS-TABLE", "NX-DRAW-ROLE-STYLE", _
        "NX-DRAW-CLEAR-INNER", "NX-DRAW-FIT-PICTURE", "NX-DRAW-INSERT-PICTURE", "NX-FILE-FOLDER-CREATE", "NX-FILE-CONSOLIDATE", _
        "NX-FILE-MANNER-SAVE", "NX-FILE-RANGE-PNG", "NX-FILE-CHART-PNG", "NX-FILE-SHEET-COPY-SAVE", "NX-FILE-RANGE-COPY-SAVE", _
        "NX-FILE-WORKBOOK-COMPARE", "NX-FILE-BATCH-RENAME", "NX-FILE-SHEET-BATCH-RENAME", "NX-FILE-PDF-CURRENT-SHEET", "NX-FILE-PDF-EACH-SHEET", _
        "NX-FILE-PDF-SELECTED-COMBINED", "NX-FILE-PDF-ALL-COMBINED", "NX-FILE-PDF-SETTINGS", "NX-UTIL-CALCULATOR", "NX-UTIL-SYMBOLS", _
        "NX-UTIL-NAVIGATOR", "NX-UTIL-DOCUMENT-NAVIGATOR", "NX-FILE-SHEET-COMPARE", "NX-FILE-FILE-COMPARE")
    ' END GENERATED RELEASE FEATURE IDS
    For Each featureId In ids
        canonical = canonical & CStr(featureId) & "|" & ExactOwnerFor(CStr(featureId)) & "|" & MaximumCellsFor(CStr(featureId)) & "|" & MaximumAreasFor(CStr(featureId)) & "|" & MaximumPromptCharsFor(CStr(featureId)) & "|" & MaximumPicturesFor(CStr(featureId)) & "|" & MaximumDirectoryItemsFor(CStr(featureId)) & "|" & MaximumInputFilesFor(CStr(featureId)) & "|" & MaximumSheetsFor(CStr(featureId)) & "|" & MaximumBytesFor(CStr(featureId)) & "|" & MaximumPngCellsFor(CStr(featureId)) & "|" & MaximumPngDimensionFor(CStr(featureId)) & "|" & PowerShellTimeoutSecondsFor(CStr(featureId)) & vbLf
    Next featureId
    SourceDigest = "NX-LIMITS-PROJECTION-1|" & NxInvariantUnsigned(CDbl(Len(canonical))) & "|" & canonical
End Function
