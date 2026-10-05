Attribute VB_Name = "NxSymbolFont"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function CreateFontW Lib "gdi32" (ByVal height As Long, ByVal width As Long, ByVal escapement As Long, ByVal orientation As Long, ByVal weight As Long, ByVal italic As Long, ByVal underline As Long, ByVal strikeOut As Long, ByVal charSet As Long, ByVal outputPrecision As Long, ByVal clipPrecision As Long, ByVal quality As Long, ByVal pitchAndFamily As Long, ByVal faceName As LongPtr) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hdc As LongPtr, ByVal objectHandle As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal objectHandle As LongPtr) As Long
Private Declare PtrSafe Function GetGlyphIndicesW Lib "gdi32" (ByVal hdc As LongPtr, ByVal textPointer As LongPtr, ByVal textLength As Long, ByRef glyphIndices As Integer, ByVal flags As Long) As Long
Private Declare PtrSafe Function GetTextFaceW Lib "gdi32" (ByVal hdc As LongPtr, ByVal bufferLength As Long, ByVal faceBuffer As LongPtr) As Long
#Else
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateFontW Lib "gdi32" (ByVal height As Long, ByVal width As Long, ByVal escapement As Long, ByVal orientation As Long, ByVal weight As Long, ByVal italic As Long, ByVal underline As Long, ByVal strikeOut As Long, ByVal charSet As Long, ByVal outputPrecision As Long, ByVal clipPrecision As Long, ByVal quality As Long, ByVal pitchAndFamily As Long, ByVal faceName As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal objectHandle As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal objectHandle As Long) As Long
Private Declare Function GetGlyphIndicesW Lib "gdi32" (ByVal hdc As Long, ByVal textPointer As Long, ByVal textLength As Long, ByRef glyphIndices As Integer, ByVal flags As Long) As Long
Private Declare Function GetTextFaceW Lib "gdi32" (ByVal hdc As Long, ByVal bufferLength As Long, ByVal faceBuffer As Long) As Long
#End If

Private Const NX_GGI_MARK_NONEXISTING_GLYPHS As Long = 1
Private Const NX_DEFAULT_CHARSET As Long = 1
Private Const NX_SYMBOL_FALLBACK_FONT As String = "Segoe UI Symbol"

Public Function NxSymbolResolveDisplayFont(ByVal requestedFont As String, ByVal codePoint As Long, ByRef canDisplay As Boolean) As String
    Dim character As String
    character = NxSymbolFromCodePoint(codePoint)
    If Len(Trim$(requestedFont)) = 0 Then requestedFont = ActiveCellFontName()
    If Len(requestedFont) > 0 Then
        If NxSymbolFontHasGlyph(requestedFont, character) Then
            canDisplay = True
            NxSymbolResolveDisplayFont = requestedFont
            Exit Function
        End If
    End If
    NxSymbolResolveDisplayFont = NX_SYMBOL_FALLBACK_FONT
    canDisplay = NxSymbolFontHasGlyph(NX_SYMBOL_FALLBACK_FONT, character)
End Function

Public Function NxSymbolFontHasGlyph(ByVal fontName As String, ByVal value As String) As Boolean
#If VBA7 Then
    Dim hdc As LongPtr, fontHandle As LongPtr, previousFont As LongPtr
#Else
    Dim hdc As Long, fontHandle As Long, previousFont As Long
#End If
    Dim glyphs() As Integer, actualFace As String, faceLength As Long, glyphResult As Long
    Dim index As Long, success As Boolean
    If Len(Trim$(fontName)) = 0 Or Len(value) = 0 Then Exit Function
    On Error GoTo CleanUp
    hdc = GetDC(0)
    If hdc = 0 Then GoTo CleanUp
    fontHandle = CreateFontW(0, 0, 0, 0, 400, 0, 0, 0, NX_DEFAULT_CHARSET, 0, 0, 0, 0, StrPtr(fontName))
    If fontHandle = 0 Then GoTo CleanUp
    previousFont = SelectObject(hdc, fontHandle)
    If previousFont = 0 Then GoTo CleanUp
    actualFace = String$(128, vbNullChar)
    faceLength = GetTextFaceW(hdc, 128, StrPtr(actualFace))
    If faceLength <= 0 Then GoTo CleanUp
    actualFace = Left$(actualFace, InStr(1, actualFace, vbNullChar, vbBinaryCompare) - 1)
    If StrComp(actualFace, fontName, vbTextCompare) <> 0 Then GoTo CleanUp
    ReDim glyphs(0 To Len(value) - 1)
    glyphResult = GetGlyphIndicesW(hdc, StrPtr(value), Len(value), glyphs(0), NX_GGI_MARK_NONEXISTING_GLYPHS)
    If glyphResult = -1 Then GoTo CleanUp
    For index = LBound(glyphs) To UBound(glyphs)
        If glyphs(index) = -1 Then GoTo CleanUp
    Next index
    success = True
CleanUp:
    On Error Resume Next
    If previousFont <> 0 And hdc <> 0 Then SelectObject hdc, previousFont
    If fontHandle <> 0 Then DeleteObject fontHandle
    If hdc <> 0 Then ReleaseDC 0, hdc
    On Error GoTo 0
    NxSymbolFontHasGlyph = success
End Function

Private Function ActiveCellFontName() As String
    On Error GoTo Failed
    If TypeName(Application.ActiveCell) = "Range" Then ActiveCellFontName = CStr(Application.ActiveCell.Font.Name)
    Exit Function
Failed:
    Err.Clear
End Function
