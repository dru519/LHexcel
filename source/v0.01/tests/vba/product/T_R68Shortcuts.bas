Attribute VB_Name = "T_R68Shortcuts"
Option Explicit

' Pure parser/capture/isolated draft tests. Never bind OnKey or save a profile.
Public Function Names() As String
    Names = "parser_aliases|parser_invalid|capture|reserved|punctuation|function_keys|draft_conflict|direct_input"
End Function

Public Function RunCase(ByVal name As String) As String
    Dim modifierName As String, keyName As String, reason As String, chord As Variant
    Dim draft As Object, view As FNxShortcutManager
    Const firstRoute As String = "feature:NX-DATA-NORMALIZE"
    Const secondRoute As String = "feature:NX-DATA-UNIQUE-COUNT"
    On Error GoTo Failed
    Select Case name
        Case "parser_aliases"
            RequireChord " shift + CONTROL + alt + q ", "Ctrl+Alt+Shift", "Q"
            RequireChord "alt+ctrl+q", "Ctrl+Alt", "Q"
            RequireChord "Ctrl+Alt+[;]", "Ctrl+Alt", "Semicolon"
            RequireChord "Control+Alt+WonBackslash", "Ctrl+Alt", "Backslash"
            RequireChord "Ctrl+Alt+₩", "Ctrl+Alt", "Backslash"
            RequireChord "Ctrl+Alt+leftsquarebracket", "Ctrl+Alt", "LeftBracket"
            RequireChord "Ctrl+Alt+Period", "Ctrl+Alt", "Period"
        Case "parser_invalid"
            For Each chord In Array("", "Ctrl", "Ctrl+Alt", "Ctrl+", "Ctrl+Control+Q", _
                "Alt+Alt+Q", "Shift+Shift+Q", "Ctrl+Alt+A+B", "Win+Q", "AltGr+Q", _
                "Ctrl++Alt+Q", "Ctrl+Alt+F0", "Ctrl+Alt+F16", "Ctrl+Alt+한")
                modifierName = "stale"
                keyName = "stale"
                Require Not NxShortcutParseCombination(CStr(chord), modifierName, keyName, reason), "invalid accepted: " & CStr(chord)
                Require Len(reason) > 0 And Len(modifierName) = 0 And Len(keyName) = 0, "invalid parser leaked capture"
            Next chord
        Case "capture"
            Require NxShortcutCaptureDisplay(vbKeyQ, 6) = "Ctrl+Alt+Q", "Ctrl+Alt capture changed"
            Require NxShortcutCaptureDisplay(vbKeyQ, 7) = "Ctrl+Alt+Shift+Q", "modifier order changed"
            Require NxShortcutCaptureDisplay(vbKeyQ, 3) = "Ctrl+Shift+Q", "Ctrl+Shift capture changed"
            Require NxShortcutCaptureDisplay(186, 6) = "Ctrl+Alt+Semicolon", "OEM semicolon missing"
            Require NxShortcutCaptureDisplay(220, 6) = "Ctrl+Alt+Backslash", "OEM won/backslash missing"
            Require NxShortcutCaptureDisplay(187, 7) = "Ctrl+Alt+Shift+Equals", "shifted OEM key changed"
            Require NxShortcutCaptureDisplay(vbKeyNumpad2, 6) = "Ctrl+Alt+2", "legacy numpad alias changed"
            Require NxShortcutCaptureDisplay(124, 1) = "Shift+F13", "Shift function key capture missing"
            Require NxShortcutCaptureDisplay(125, 4) = "Alt+F14", "Alt function key capture missing"
            Require Not NxShortcutCaptureKey(vbKeyControl, 2, modifierName, keyName), "modifier-only capture accepted"
            Require Len(modifierName) = 0 And Len(keyName) = 0, "unsupported capture leaked keys"
        Case "reserved"
            For Each chord In Array("Ctrl+C", "Ctrl+V", "Ctrl+Q", "Alt+F4", "Alt+F11", _
                "Alt+Shift+H", "Ctrl+Alt+Delete", "Ctrl+Shift+Escape", "A", "Shift+A", _
                "Ctrl+Shift+A", "Ctrl+Shift+C", "Ctrl+Shift+F", "Ctrl+Shift+G", "Ctrl+Shift+L", _
                "Ctrl+Shift+M", "Ctrl+Shift+O", "Ctrl+Shift+P", "Ctrl+Shift+U", "Ctrl+Shift+V", _
                "Ctrl+Shift+8", "Ctrl+*", "Ctrl+Alt+5", "Ctrl+Alt+=", "Ctrl+Alt+-", _
                "Ctrl+Alt+V", "Ctrl+Alt+P", "Ctrl+Alt+F5", "Ctrl+Alt+F9", _
                "Ctrl+Alt+Shift+C", "Ctrl+Alt+Shift+M", "Ctrl+Alt+Shift+P", _
                "Ctrl+Alt+Shift+F2", "Ctrl+Alt+Shift+F9")
                RequireReserved CStr(chord)
            Next chord
        Case "punctuation"
            RequireChord "Ctrl+Alt+?", "Ctrl+Alt+Shift", "Slash"
            RequireChord "Ctrl+Alt++", "Ctrl+Alt+Shift", "Equals"
            RequireChord "Ctrl+Alt+[+]", "Ctrl+Alt+Shift", "Equals"
            RequireChord "Ctrl+Alt+|", "Ctrl+Alt+Shift", "Backslash"
            Require NxShortcutOnKeySpec("Ctrl+Alt", "Semicolon") = "^%;", "semicolon OnKey"
            Require NxShortcutOnKeySpec("Alt+Control", "[;]") = "^%;", "OnKey alias differs"
            Require NxShortcutOnKeySpec("Ctrl+Alt", "?") = "^%+/", "shifted slash OnKey"
            Require NxShortcutOnKeySpec("Ctrl+Alt", "[") = "^%{[}", "bracket escaping"
            Require NxShortcutOnKeySpec("Ctrl+Alt", "|") = "^%+\", "pipe alias escaped profile separator"
            RequireReserved "Ctrl+Alt+PlusSign"
            RequireReserved "Ctrl+Alt+Shift+Equals"
            RequireReserved "Ctrl+Shift+:"
        Case "function_keys"
            For Each chord In Array("F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "F11", "F12", _
                "Ctrl+Shift+F1", "Ctrl+Shift+F2", "Ctrl+Shift+F3", "Ctrl+Shift+F4", "Ctrl+Shift+F5", _
                "Ctrl+Shift+F6", "Ctrl+Shift+F10", "Ctrl+Shift+F11", "Ctrl+Shift+F12")
                RequireReserved CStr(chord)
            Next chord
            RequireChord "F13", "None", "F13"
            Require NxShortcutOnKeySpec("None", "F13") = "{F13}", "plain F13 OnKey"
            Require NxShortcutOnKeySpec("Shift", "F13") = "+{F13}", "Shift F13 OnKey"
            Require NxShortcutOnKeySpec("Alt", "F14") = "%{F14}", "Alt F14 OnKey"
            Require NxShortcutOnKeySpec("Ctrl", "F15") = "^{F15}", "Ctrl F15 OnKey"
            Require NxShortcutOnKeySpec("Alt+Shift", "F15") = "%+{F15}", "Alt Shift F15 OnKey"
        Case "draft_conflict"
            Set draft = CreateObject("Scripting.Dictionary")
            draft.CompareMode = vbBinaryCompare
            NxShortcutDraftAssign draft, firstRoute, "Alt+Control", "[;]"
            Require NxShortcutDraftDisplay(draft, firstRoute) = "Ctrl+Alt+Semicolon", "draft not canonical"
            Require Len(NxShortcutDraftConflictReason(draft, "ctrl+alt", "Semicolon", firstRoute)) = 0, "self conflict"
            Require Len(NxShortcutDraftConflictReason(draft, "ctrl+alt", ";", secondRoute)) > 0, "alias duplicate accepted"
            Require Len(NxShortcutDraftConflictReason(draft, "Ctrl+Alt", "?", secondRoute)) = 0, "safe punctuation rejected"
            NxShortcutDraftAssign draft, secondRoute, "Ctrl+Alt", "?"
            Require NxShortcutDraftDisplay(draft, secondRoute) = "Ctrl+Alt+Shift+Slash", "shift alias not canonical"
            Require Len(NxShortcutDraftConflictReason(draft, "Ctrl+Alt+Shift", "Slash", firstRoute)) > 0, "shift alias collision accepted"
            Require Len(NxShortcutDraftConflictReason(draft, "Ctrl+Shift", "L", firstRoute)) > 0, "reserved draft accepted"
            Require draft.Count = 2, "conflict preview mutated draft"
            draft.Item(firstRoute) = Array("Alt+Control", "[;]")
            Require Len(NxShortcutDraftConflictReason(draft, "Ctrl+Alt", "Semicolon", secondRoute)) > 0, "noncanonical draft duplicate accepted"
        Case "direct_input"
            Set view = New FNxShortcutManager
            Require Not view.Controls("txtShortcut").Locked, "direct input still locked"
            If view.Controls("lstCommands").ListCount > 0 Then view.Controls("lstCommands").ListIndex = 0
            view.Controls("txtShortcut").Value = "Alt+Control+Semicolon"
            Require InStr(1, CStr(view.Controls("lblPreview").Caption), "Ctrl+Alt+Semicolon", vbBinaryCompare) > 0, "text change not parsed"
            view.Controls("txtShortcut").Value = "Ctrl+Shift+V"
            Require Not view.Controls("cmdAssign").Enabled, "native shortcut editable preview accepted"
            view.Controls("txtShortcut").Value = "Ctrl+Control+Q"
            Require Not view.Controls("cmdAssign").Enabled, "malformed direct input kept previous capture"
            Unload view
            Set view = Nothing
        Case Else
            Err.Raise vbObjectError + 968, "T_R68Shortcuts", "unknown case"
    End Select
    RunCase = "PASS|" & name
    Exit Function
Failed:
    RunCase = "FAIL|" & name & "|" & CStr(Err.Number) & "|" & Err.Description
    On Error Resume Next
    If Not view Is Nothing Then Unload view
    On Error GoTo 0
End Function

Private Sub RequireChord(ByVal text As String, ByVal expectedModifier As String, ByVal expectedKey As String)
    Dim modifierName As String, keyName As String, reason As String
    Require NxShortcutParseCombination(text, modifierName, keyName, reason), "parse failed: " & text & " / " & reason
    Require modifierName = expectedModifier And keyName = expectedKey, "wrong canonical chord: " & text & " -> " & modifierName & "+" & keyName
End Sub

Private Sub RequireReserved(ByVal text As String)
    Dim modifierName As String, keyName As String, reason As String
    Require NxShortcutParseCombination(text, modifierName, keyName, reason), "reserved syntax not recognized: " & text
    Require NxShortcutIsReservedCombination(modifierName, keyName), "native chord accepted: " & text
    Require OnKeyRejected(modifierName, keyName), "native chord reached OnKey spec: " & text
End Sub

Private Function OnKeyRejected(ByVal modifierName As String, ByVal keyName As String) As Boolean
    Dim result As String
    On Error GoTo Rejected
    result = NxShortcutOnKeySpec(modifierName, keyName)
    Exit Function
Rejected:
    OnKeyRejected = (InStr(1, Err.Description, "예약 단축키", vbBinaryCompare) > 0)
    Err.Clear
End Function

Private Sub Require(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise vbObjectError + 969, "T_R68Shortcuts", detail
End Sub
