VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxExecutionContext"
Option Explicit

Private mWorkbookIdentity As String
Private mWorkbookKind As String
Private mSheetCodeName As String
Private mSheetOrdinal As Long
Private mSelectionType As String
Private mAreaCount As Long
Private mActualCellCount As Double
Private mReadOnly As Boolean
Private mProtectedFlags As String
Private mFingerprint As String
Private mSourceStateDigest As String
Private mCaptured As Boolean
Private mWorkbook As Workbook
Private mSheet As Worksheet
Private mSelection As Range
Private mChartIdentity As String
Private mSelectionAuthority As Collection
Private mSelectionAuthorityBound As Boolean

Friend Sub Capture(ByVal factoryToken As String, Optional ByVal captureSourceState As Boolean = True, Optional ByVal selectionLimit As Long = 200000)
    Dim area As Range
    If factoryToken <> NX_CONTEXT_FACTORY_TOKEN Then NxRaiseContractError "Context capture is factory-only"
    If mCaptured Then NxRaiseContractError "Context can be captured only once"
    Set mWorkbook = Application.ActiveWorkbook
    Set mSheet = Application.ActiveSheet
    If mWorkbook Is Nothing Or mSheet Is Nothing Then NxRaiseContractError "Active Excel identity is required"
    If Len(mWorkbook.Path) = 0 Then
        mWorkbookKind = "unsaved"
        mWorkbookIdentity = "unsaved:" & NxContextFactory.UnsavedWorkbookSessionToken(mWorkbook)
    Else
        mWorkbookKind = "saved"
        mWorkbookIdentity = mWorkbook.FullName
    End If
    mSheetCodeName = mSheet.CodeName
    mSheetOrdinal = FindSheetOrdinal(mWorkbook, mSheet)
    mReadOnly = mWorkbook.ReadOnly
    mProtectedFlags = ProtectedFlagsWire(mSheet)
    If TypeName(Application.Selection) <> "Range" Then
        If Not CaptureChartSelection() Then NxRaiseContractError "Range or chart selection is required"
        mSelectionType = "Chart"
        mAreaCount = 0
        mActualCellCount = 0
        mFingerprint = SerializeIdentity
        mSourceStateDigest = "NXCHARTSOURCE2|" & mChartIdentity
        mCaptured = True
        Exit Sub
    End If
    Set mSelection = Application.Selection
    If mSelection Is Nothing Then NxRaiseContractError "Range selection is required"

    mSelectionType = "Range"
    mAreaCount = mSelection.Areas.Count
    For Each area In mSelection.Areas
        If captureSourceState And area.Cells.CountLarge > selectionLimit Then NxRaiseContractError "Selection exceeds cell journal cap"
        mActualCellCount = mActualCellCount + CDbl(area.Cells.CountLarge)
    Next area
    mFingerprint = SerializeIdentity
    If captureSourceState Then mSourceStateDigest = CaptureSourceSnapshot
    mCaptured = True
End Sub

Private Function CaptureChartSelection() As Boolean
    Dim candidate As Object, chart As Object, chartName As String, selectionType As String
    On Error Resume Next
    Set candidate = Application.Selection
    If candidate Is Nothing Then Exit Function
    selectionType = TypeName(candidate)
    Select Case selectionType
        Case "Chart": Set chart = candidate
        Case "ChartObject", "Shape": Set chart = candidate.Chart
        Case "ShapeRange"
            If candidate.Count = 1 Then Set chart = candidate.Item(1).Chart
        Case "ChartArea", "PlotArea", "Legend", "Series": Set chart = candidate.Parent
    End Select
    If Not chart Is Nothing Then chartName = chart.Name
    If Err.Number <> 0 Or Len(Trim$(chartName)) = 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    mChartIdentity = mSheet.CodeName & "!" & chartName
    CaptureChartSelection = True
End Function

Private Function CurrentChartIdentity() As String
    Dim candidate As Object, chart As Object, selectionType As String
    On Error Resume Next
    Set candidate = Application.Selection
    If candidate Is Nothing Then Exit Function
    selectionType = TypeName(candidate)
    Select Case selectionType
        Case "Chart": Set chart = candidate
        Case "ChartObject", "Shape": Set chart = candidate.Chart
        Case "ShapeRange"
            If candidate.Count = 1 Then Set chart = candidate.Item(1).Chart
        Case "ChartArea", "PlotArea", "Legend", "Series": Set chart = candidate.Parent
    End Select
    If Not chart Is Nothing Then CurrentChartIdentity = mSheet.CodeName & "!" & chart.Name
    Err.Clear
    On Error GoTo 0
End Function

Public Property Get WorkbookIdentity() As String: WorkbookIdentity = mWorkbookIdentity: End Property
Public Property Get SheetCodeName() As String: SheetCodeName = mSheetCodeName: End Property
Public Property Get SheetOrdinal() As Long: SheetOrdinal = mSheetOrdinal: End Property
Public Property Get SelectionType() As String: SelectionType = mSelectionType: End Property
Public Property Get SelectionAreaCount() As Long: SelectionAreaCount = mAreaCount: End Property
Public Property Get ActualCellCount() As Double: ActualCellCount = mActualCellCount: End Property
Public Property Get IsReadOnly() As Boolean: IsReadOnly = mReadOnly: End Property
Public Property Get ProtectedFlags() As String: ProtectedFlags = mProtectedFlags: End Property
Public Property Get Fingerprint() As String: Fingerprint = mFingerprint: End Property

Friend Sub BindSelectionAuthority(ByVal authority As Collection)
    If authority Is Nothing Then NxRaiseContractError "Selection authority is required"
    If mSelectionAuthorityBound Then NxRaiseContractError "Selection authority can be bound only once"
    Set mSelectionAuthority = authority
    mSelectionAuthorityBound = True
End Sub

Friend Function InternalSelectionRange(ByVal authority As Collection) As Range
    If authority Is Nothing Or mSelectionAuthority Is Nothing Then NxRaiseContractError "Selection authority is required"
    If Not (authority Is mSelectionAuthority) Then NxRaiseContractError "Selection authority identity was rejected"
    Set InternalSelectionRange = mSelection
End Function

Friend Function InternalRelativeRange(ByVal authority As Collection, ByVal rowOffset As Long, _
    ByVal columnOffset As Long, ByVal rowCount As Long, ByVal columnCount As Long) As Range

    Dim baseRange As Range
    If authority Is Nothing Or mSelectionAuthority Is Nothing Then NxRaiseContractError "Selection authority is required"
    If Not (authority Is mSelectionAuthority) Then NxRaiseContractError "Selection authority identity was rejected"
    If mSelection.Areas.Count <> 1 Then NxRaiseContractError "Relative AI operations require one contiguous selection"
    If rowOffset < 0 Or columnOffset < 0 Or rowCount <= 0 Or columnCount <= 0 Then NxRaiseContractError "Relative target dimensions are invalid"
    Set baseRange = mSelection.Areas(1)
    If CDbl(rowOffset) + CDbl(rowCount) > baseRange.Rows.CountLarge Or _
        CDbl(columnOffset) + CDbl(columnCount) > baseRange.Columns.CountLarge Then
        NxRaiseContractError "Operation target escapes the captured selection"
    End If
    Set InternalRelativeRange = baseRange.Cells(1, 1).Offset(rowOffset, columnOffset).Resize(rowCount, columnCount)
End Function

Public Function SelectionTargetIdentity() As String
    SelectionTargetIdentity = "W=" & EncodeField(mWorkbookIdentity) & "S=" & EncodeField(mSheetCodeName) & "O=" & NxInvariantUnsigned(CDbl(mSheetOrdinal)) & "F=" & EncodeField(mFingerprint)
End Function

Public Function SourceStateDigest() As String
    SourceStateDigest = mSourceStateDigest
End Function

Public Function SourceStateMatchesCurrent() As Boolean
    If mSelectionType = "Chart" Then
        SourceStateMatchesCurrent = (mSourceStateDigest = "NXCHARTSOURCE2|" & CurrentChartIdentity())
    Else
        SourceStateMatchesCurrent = (mSourceStateDigest = CaptureSourceSnapshot)
    End If
End Function

Private Function CaptureSourceSnapshot() As String
    If mSelectionType = "Chart" Then
        CaptureSourceSnapshot = "NXCHARTSOURCE2|" & mChartIdentity
        Exit Function
    End If
    Dim xmlSnapshot As String
    If TryCaptureXmlSnapshot(xmlSnapshot) Then
        CaptureSourceSnapshot = xmlSnapshot
        Exit Function
    End If
    Dim cell As Range, payload As String, formulaMode As String, formulaValue As Variant
    Dim records() As String, index As Long
    Dim formats As Object, formatRecords As Collection, formatText As String, formatId As String
    Dim formatArray() As String, formatIndex As Long
    Set formats = CreateObject("Scripting.Dictionary")
    formats.CompareMode = vbBinaryCompare
    Set formatRecords = New Collection
    ReDim records(0 To CLng(mActualCellCount) - 1)
    For Each cell In mSelection.Cells
        payload = vbNullString
        payload = payload & SnapshotField("AD", cell.Address(False, False, xlA1, True))
        If cell.HasFormula Then
            formulaValue = ReadFormula2OrFallback(cell, formulaMode)
            payload = payload & SnapshotField("FM", formulaMode)
            payload = payload & SnapshotField("FV", SourceStateValue(formulaValue))
        Else
            payload = payload & SnapshotField("FM", "Value2")
            payload = payload & SnapshotField("FV", SourceStateValue(cell.Value2))
        End If
        formatText = SnapshotField("NF", SourceStateValue(cell.NumberFormat))
        formatText = formatText & SnapshotFontState(cell)
        formatText = formatText & SnapshotInteriorState(cell)
        formatText = formatText & SnapshotAlignmentState(cell)
        formatText = formatText & SnapshotProtectionState(cell)
        formatText = formatText & SnapshotBorderState(cell)
        ' Store repeated formatting once, without sampling or dropping any field.
        ' Cap the dictionary so highly heterogeneous sheets do not double memory.
        If formats.Exists(formatText) Then
            payload = payload & SnapshotField("FI", formats(formatText))
        ElseIf formats.Count < 4096 Then
            formatId = CStr(formats.Count)
            formats.Add formatText, formatId
            formatRecords.Add SnapshotField("FD", formatText)
            payload = payload & SnapshotField("FI", formatId)
        Else
            payload = payload & SnapshotField("FF", formatText)
        End If
        If cell.MergeCells Then payload = payload & SnapshotField("MG", cell.MergeArea.Address) Else payload = payload & SnapshotField("MG", vbNullString)
        records(index) = payload
        index = index + 1
    Next cell
    ReDim formatArray(0 To formatRecords.Count)
    For formatIndex = 1 To formatRecords.Count
        formatArray(formatIndex - 1) = formatRecords(formatIndex)
    Next formatIndex
    CaptureSourceSnapshot = "NXSOURCE4|" & SnapshotField("FD", Join(formatArray, vbNullString)) & SnapshotField("SS", Join(records, vbNullString))
End Function

Private Function TryCaptureXmlSnapshot(ByRef snapshot As String) As Boolean
    Dim area As Range, document As Object, styles As Object, table As Object
    Dim cachedValues As Object, cachedValue As Object, records() As String, index As Long
    On Error GoTo Unsupported
    ReDim records(0 To mSelection.Areas.Count - 1)
    For Each area In mSelection.Areas
        Set document = CreateObject("MSXML2.DOMDocument.6.0")
        document.async = False
        document.validateOnParse = False
        document.resolveExternals = False
        document.setProperty "ProhibitDTD", True
        If Not document.LoadXML(CStr(area.Value(xlRangeValueXMLSpreadsheet))) Then GoTo Unsupported
        document.setProperty "SelectionNamespaces", "xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet'"
        Set styles = document.selectSingleNode("/ss:Workbook/ss:Styles")
        Set table = document.selectSingleNode("/ss:Workbook/ss:Worksheet/ss:Table")
        If styles Is Nothing Or table Is Nothing Then GoTo Unsupported
        NormalizeDefaultXmlStyles styles, table
        ' A formula's cached result may change during normal recalculation.
        ' The formula itself remains in the table and is part of the approval.
        Set cachedValues = table.selectNodes("ss:Row/ss:Cell[@ss:Formula]/ss:Data")
        For Each cachedValue In cachedValues
            cachedValue.parentNode.removeChild cachedValue
        Next cachedValue
        records(index) = SnapshotField("AD", area.Address(False, False, xlA1, True)) & _
            SnapshotField("ST", styles.XML) & SnapshotField("TB", table.XML)
        index = index + 1
    Next area
    snapshot = "NXSOURCE5|" & Join(records, vbNullString)
    TryCaptureXmlSnapshot = True
    Exit Function
Unsupported:
    ' Unsupported Excel XML ranges use the existing complete property reader.
    Err.Clear
    snapshot = vbNullString
End Function

Private Sub NormalizeDefaultXmlStyles(ByVal styles As Object, ByVal table As Object)
    Dim defaultStyle As Object, style As Object, part As Object, defaultPart As Object
    Dim reference As Object, styleId As String, redundant As Boolean
    Set defaultStyle = styles.selectSingleNode("ss:Style[@ss:ID='Default']")
    If defaultStyle Is Nothing Then Exit Sub
    For Each style In styles.selectNodes("ss:Style[@ss:ID!='Default' and not(@ss:Parent)]")
        redundant = True
        For Each part In style.childNodes
            Set defaultPart = defaultStyle.selectSingleNode("ss:" & part.baseName)
            If defaultPart Is Nothing Then
                redundant = False
                Exit For
            End If
            If part.XML <> defaultPart.XML Then
                redundant = False
                Exit For
            End If
        Next part
        If redundant Then
            styleId = style.getAttribute("ss:ID")
            For Each reference In table.selectNodes(".//*[@ss:StyleID='" & styleId & "']")
                reference.removeAttribute "ss:StyleID"
            Next reference
            styles.removeChild style
        End If
    Next style
End Sub

Private Function ReadFormula2OrFallback(ByVal cell As Range, ByRef formulaMode As String) As Variant
    ReadFormula2OrFallback = NxReadFormula2OrFallback(cell, formulaMode)
End Function

Private Function SnapshotFontState(ByVal cell As Range) As String
    Dim fontState As Font
    Set fontState = cell.Font
    SnapshotFontState = SnapshotField("FN", SourceStateValue(fontState.Name))
    SnapshotFontState = SnapshotFontState & SnapshotField("FS", SourceStateValue(fontState.Size))
    SnapshotFontState = SnapshotFontState & SnapshotField("FB", SourceStateValue(fontState.Bold))
    SnapshotFontState = SnapshotFontState & SnapshotField("FI", SourceStateValue(fontState.Italic))
    SnapshotFontState = SnapshotFontState & SnapshotField("FU", SourceStateValue(fontState.Underline))
    SnapshotFontState = SnapshotFontState & SnapshotField("FX", SourceStateValue(fontState.Strikethrough))
    SnapshotFontState = SnapshotFontState & SnapshotField("FC", SourceStateValue(fontState.Color))
    SnapshotFontState = SnapshotFontState & SnapshotField("FJ", SourceStateValue(fontState.ColorIndex))
    SnapshotFontState = SnapshotFontState & SnapshotField("FT", SourceStateValue(fontState.TintAndShade))
End Function

Private Function SnapshotInteriorState(ByVal cell As Range) As String
    Dim fillState As Interior
    Set fillState = cell.Interior
    SnapshotInteriorState = SnapshotField("IP", SourceStateValue(fillState.Pattern))
    SnapshotInteriorState = SnapshotInteriorState & SnapshotField("PP", SourceStateValue(fillState.PatternColor))
    SnapshotInteriorState = SnapshotInteriorState & SnapshotField("PJ", SourceStateValue(fillState.PatternColorIndex))
    SnapshotInteriorState = SnapshotInteriorState & SnapshotField("IC", SourceStateValue(fillState.Color))
    SnapshotInteriorState = SnapshotInteriorState & SnapshotField("IJ", SourceStateValue(fillState.ColorIndex))
    SnapshotInteriorState = SnapshotInteriorState & SnapshotField("IT", SourceStateValue(fillState.TintAndShade))
End Function

Private Function SnapshotAlignmentState(ByVal cell As Range) As String
    SnapshotAlignmentState = SnapshotField("HA", SourceStateValue(cell.HorizontalAlignment))
    SnapshotAlignmentState = SnapshotAlignmentState & SnapshotField("VA", SourceStateValue(cell.VerticalAlignment))
    SnapshotAlignmentState = SnapshotAlignmentState & SnapshotField("OR", SourceStateValue(cell.Orientation))
    SnapshotAlignmentState = SnapshotAlignmentState & SnapshotField("AI", SourceStateValue(cell.AddIndent))
    SnapshotAlignmentState = SnapshotAlignmentState & SnapshotField("IL", SourceStateValue(cell.IndentLevel))
    SnapshotAlignmentState = SnapshotAlignmentState & SnapshotField("SF", SourceStateValue(cell.ShrinkToFit))
    SnapshotAlignmentState = SnapshotAlignmentState & SnapshotField("RO", SourceStateValue(cell.ReadingOrder))
    SnapshotAlignmentState = SnapshotAlignmentState & SnapshotField("WT", SourceStateValue(cell.WrapText))
End Function

Private Function SnapshotProtectionState(ByVal cell As Range) As String
    SnapshotProtectionState = SnapshotField("LK", SourceStateValue(cell.Locked))
    SnapshotProtectionState = SnapshotProtectionState & SnapshotField("FH", SourceStateValue(cell.FormulaHidden))
End Function

Private Function SnapshotBorderState(ByVal cell As Range) As String
    Dim edge As Variant, borderState As Border
    For Each edge In Array(xlDiagonalDown, xlDiagonalUp, xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
        Set borderState = cell.Borders(edge)
        SnapshotBorderState = SnapshotBorderState & SnapshotField("BL", SourceStateValue(borderState.LineStyle))
        SnapshotBorderState = SnapshotBorderState & SnapshotField("BW", SourceStateValue(borderState.Weight))
        SnapshotBorderState = SnapshotBorderState & SnapshotField("BC", SourceStateValue(borderState.Color))
        SnapshotBorderState = SnapshotBorderState & SnapshotField("BJ", SourceStateValue(borderState.ColorIndex))
        SnapshotBorderState = SnapshotBorderState & SnapshotField("BT", SourceStateValue(borderState.TintAndShade))
    Next edge
End Function

Private Function SourceStateValue(ByVal value As Variant) As String
    If IsEmpty(value) Then
        SourceStateValue = SnapshotField("EM", vbNullString)
    ElseIf IsNull(value) Then
        SourceStateValue = SnapshotField("NL", vbNullString)
    ElseIf IsError(value) Then
        SourceStateValue = SnapshotField("ER", LTrim$(Str$(CLng(value))))
    ElseIf VarType(value) = vbBoolean Then
        If CBool(value) Then SourceStateValue = SnapshotField("BO", "1") Else SourceStateValue = SnapshotField("BO", "0")
    ElseIf VarType(value) = vbString Then
        SourceStateValue = SnapshotField("ST", value & vbNullString)
    ElseIf IsNumeric(value) Then
        SourceStateValue = SnapshotField("NU", LTrim$(Str$(CDbl(value))))
    Else
        NxRaiseContractError "Source snapshot does not support this Variant type"
    End If
End Function

Private Function SnapshotField(ByVal tag As String, ByVal value As String) As String
    ' Length-prefixed UTF-16 text preserves exact values without repeated hex expansion.
    SnapshotField = tag & CStr(Len(value)) & ":" & value
End Function

Public Function FingerprintMatchesCurrent() As Boolean
    FingerprintMatchesCurrent = (NxContextFactory.CaptureCurrentFingerprint() = mFingerprint)
End Function

Public Function HasExactTarget(ByVal targetIdentity As String) As Boolean
    HasExactTarget = (targetIdentity = SelectionTargetIdentity)
End Function

Public Function MatchesRuntimeIdentity(ByVal other As CNxExecutionContext) As Boolean
    Dim otherWorkbook As Workbook
    Dim otherSheet As Worksheet
    Dim otherSelection As Range
    If other Is Nothing Then Exit Function
    Set otherWorkbook = other.RuntimeWorkbook
    Set otherSheet = other.RuntimeSheet
    If mSelectionType = "Chart" Or other.SelectionType = "Chart" Then
        MatchesRuntimeIdentity = (mWorkbook Is otherWorkbook) And (mSheet Is otherSheet) And _
            (mSelectionType = other.SelectionType) And (mChartIdentity = other.RuntimeChartIdentity)
        Exit Function
    End If
    Set otherSelection = other.InternalSelectionRange(mSelectionAuthority)
    MatchesRuntimeIdentity = (mWorkbook Is otherWorkbook) And (mSheet Is otherSheet) And SameRange(mSelection, otherSelection)
End Function

Friend Property Get RuntimeWorkbook() As Workbook
    Set RuntimeWorkbook = mWorkbook
End Property

Friend Property Get RuntimeSheet() As Worksheet
    Set RuntimeSheet = mSheet
End Property

Friend Property Get RuntimeChartIdentity() As String
    RuntimeChartIdentity = mChartIdentity
End Property

Private Function SerializeIdentity() As String
    Dim area As Range
    SerializeIdentity = "NXCTXFP1"
    SerializeIdentity = SerializeIdentity & FieldWire("SV", NxInvariantUnsigned(CDbl(NX_SCHEMA_VERSION)))
    SerializeIdentity = SerializeIdentity & FieldWire("PV", NX_VERSION)
    SerializeIdentity = SerializeIdentity & FieldWire("WK", mWorkbookKind)
    SerializeIdentity = SerializeIdentity & FieldWire("WI", mWorkbookIdentity)
    If mReadOnly Then
        SerializeIdentity = SerializeIdentity & FieldWire("RO", "1")
    Else
        SerializeIdentity = SerializeIdentity & FieldWire("RO", "0")
    End If
    SerializeIdentity = SerializeIdentity & FieldWire("SC", mSheetCodeName)
    SerializeIdentity = SerializeIdentity & FieldWire("SO", NxInvariantUnsigned(CDbl(mSheetOrdinal)))
    SerializeIdentity = SerializeIdentity & FieldWire("PR", mProtectedFlags)
    SerializeIdentity = SerializeIdentity & FieldWire("ST", mSelectionType)
    SerializeIdentity = SerializeIdentity & FieldWire("AC", NxInvariantUnsigned(CDbl(mAreaCount)))
    If mSelectionType = "Chart" Then
        SerializeIdentity = SerializeIdentity & FieldWire("CH", mChartIdentity)
    Else
        For Each area In mSelection.Areas
            SerializeIdentity = SerializeIdentity & FieldWire("AR", AreaWire(area))
        Next area
    End If
End Function

Private Function FieldWire(ByVal tag As String, ByVal value As String) As String
    Dim index As Long
    Dim unitValue As Long
    Dim payload As String
    Dim valueLength As Long
    If Len(tag) <> 2 Then NxRaiseContractError "Fingerprint tag must be two ASCII characters"
    valueLength = Len(value)
    If valueLength > 0 Then payload = String$(valueLength * 4, "0")
    For index = 1 To valueLength
        unitValue = AscW(Mid$(value, index, 1))
        If unitValue < 0 Then unitValue = unitValue + 65536
        Mid$(payload, ((index - 1) * 4) + 1, 4) = Right$("0000" & Hex$(unitValue), 4)
    Next index
    FieldWire = tag & Right$("00000000" & Hex$(valueLength), 8) & ":" & payload
End Function

Private Function AreaWire(ByVal area As Range) As String
    AreaWire = NxInvariantUnsigned(CDbl(area.Row)) & "," & NxInvariantUnsigned(CDbl(area.Column)) & "," & NxInvariantUnsigned(area.Rows.CountLarge) & "," & NxInvariantUnsigned(area.Columns.CountLarge)
End Function

Private Function FindSheetOrdinal(ByVal workbook As Workbook, ByVal sheet As Worksheet) As Long
    Dim index As Long
    For index = 1 To workbook.Worksheets.Count
        If workbook.Worksheets(index) Is sheet Then FindSheetOrdinal = index: Exit Function
    Next index
    NxRaiseContractError "Worksheet is not owned by workbook"
End Function

Private Function ProtectedFlagsWire(ByVal sheet As Worksheet) As String
    ProtectedFlagsWire = "C" & BooleanWire(sheet.ProtectContents) & "D" & BooleanWire(sheet.ProtectDrawingObjects) & "S" & BooleanWire(sheet.ProtectScenarios)
End Function

Private Function BooleanWire(ByVal value As Boolean) As String
    If value Then BooleanWire = "1" Else BooleanWire = "0"
End Function

Private Function EncodeField(ByVal value As String) As String
    EncodeField = NxInvariantUnsigned(CDbl(Len(value))) & ":" & value
End Function

Private Function SameRange(ByVal leftRange As Range, ByVal rightRange As Range) As Boolean
    On Error GoTo Different
    If leftRange Is Nothing Or rightRange Is Nothing Then Exit Function
    If Not (leftRange.Worksheet Is rightRange.Worksheet) Then Exit Function
    SameRange = (leftRange.Address(False, False, xlA1, True) = rightRange.Address(False, False, xlA1, True))
    Exit Function
Different:
    Err.Clear
End Function
