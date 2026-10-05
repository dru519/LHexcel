Attribute VB_Name = "T_R62RenameWorkbench"
Option Explicit

#Const NX_R58_TEST_BUILD = False

Public Function T_R62RenameRules() As String
    Dim rules As New Collection, rule As CNxRenameRule, failed As Boolean
    Set rule = New CNxRenameRule: rule.Kind = "PREFIX": rule.Text = "pre_": rules.Add rule
    Set rule = New CNxRenameRule: rule.Kind = "REPLACE": rule.Text = "pre": rule.Replacement = "done": rules.Add rule
    Check NxRenameApplyRules("alpha.TXT", rules, 1, True) = "done_alpha.TXT", "ordered replace"
    NxRenameMoveItem rules, 2, -1
    Check NxRenameApplyRules("alpha.TXT", rules, 1, True) = "pre_alpha.TXT", "rule order changes output"
    Set rules = New Collection
    Set rule = New CNxRenameRule: rule.Kind = "REPLACE": rule.Text = "a": rule.Replacement = "x": rule.FirstOnly = True: rule.CaseSensitive = True: rules.Add rule
    Check NxRenameApplyRules("AaA.txt", rules, 1, True) = "AxA.txt", "replace case and first"
    Set rules = New Collection
    Set rule = New CNxRenameRule: rule.Kind = "NUMBER": rule.Text = "_": rule.Position = "FRONT": rule.Start = 10: rule.StepValue = 2: rule.Digits = 3: rules.Add rule
    Check NxRenameApplyRules("alpha.TXT", rules, 3, True) = "014_alpha.TXT", "number options"
    rule.Position = "BACK"
    Check NxRenameApplyRules("alpha.TXT", rules, 2, True) = "alpha_012.TXT", "number position"
    rule.Start = 2147483647: rule.StepValue = 1
    On Error Resume Next
    NxRenameApplyRules "alpha", rules, 2, False
    failed = (Err.Number <> 0): Err.Clear
    On Error GoTo 0
    Check failed, "number overflow blocked"
    Set rules = New Collection
    Set rule = New CNxRenameRule: rule.Kind = "DELETE": rule.Start = 3: rule.Position = "FRONT": rules.Add rule
    Check NxRenameApplyRules("alpha.txt", rules, 1, True) = "ha.txt", "delete front"
    rule.Position = "BACK": rule.Start = 2
    Check NxRenameApplyRules("alpha.txt", rules, 1, True) = "alp.txt", "delete back"
    Set rules = New Collection
    Set rule = New CNxRenameRule: rule.Kind = "FULL": rule.Text = "report": rules.Add rule
    Set rule = New CNxRenameRule: rule.Kind = "CASE": rule.Position = "UPPER": rules.Add rule
    Set rule = New CNxRenameRule: rule.Kind = "EXTENSION": rule.Text = ".pdf": rules.Add rule
    Check NxRenameApplyRules("alpha.txt", rules, 1, True) = "REPORT.pdf", "full case extension"
    rule.Text = vbNullString: rule.AllowEmptyExtension = True
    Check NxRenameApplyRules("alpha.txt", rules, 1, True) = "REPORT", "explicit extension removal"
    failed = False
    On Error Resume Next
    NxRenameApplyRules "alpha", rules, 1, False
    failed = (Err.Number <> 0): Err.Clear
    On Error GoTo 0
    Check failed, "extension cannot reach sheet engine"
    T_R62RenameRules = "PASS|order,replace,number,overflow,delete,full,case,extension"
End Function

Public Function T_R62RenameFiles(ByVal rootPath As String) As String
    Dim folder As String, a As String, b As String, target As String, held As String, fso As Object
    Dim paths As New Collection, rules As New Collection, overrides As Object, preview As Collection
    Dim hasDuplicate As Boolean, hasCollision As Boolean, failed As Boolean, count As Long, rule As CNxRenameRule
    Dim collected As Collection, child As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = NewFixtureFolder(rootPath)
    a = folder & "\a.txt": b = folder & "\b.txt"
    WriteText a, "first": WriteText b, "second"
    child = folder & "\child": fso.CreateFolder child: WriteText child & "\nested.txt", "nested"
    Set collected = New Collection
    count = NxRenameAddFolder(collected, folder, False)
    Check count = 2, "nonrecursive folder input"
    count = NxRenameAddFolder(collected, folder, True)
    Check count = 1 And collected.Count = 3, "recursive input deduplicates existing files"
    paths.Add a: paths.Add b
    Set overrides = CreateObject("Scripting.Dictionary"): overrides.CompareMode = vbTextCompare
    overrides.Add a, "b.txt": overrides.Add b, "a.txt"
    Set preview = NxRenameFileRulesPreview(paths, rules, overrides, hasDuplicate, hasCollision)
    Check Not hasDuplicate And Not hasCollision, "cycle preview"
    count = NxBatchRenameCommit(preview)
    Check count = 2 And ReadText(a) = "second" And ReadText(b) = "first", "two-stage cycle"
    count = NxBatchRenameUndoLast(True)
    Check count = 2 And ReadText(a) = "first" And ReadText(b) = "second", "cycle undo"
    Set rule = New CNxRenameRule: rule.Kind = "PREFIX": rule.Text = "rule_": rules.Add rule
    NxRenameMoveItem paths, 2, -1
    Set preview = NxRenameFileRulesPreview(paths, rules, overrides, hasDuplicate, hasCollision)
    Check CStr(preview(1)(1)) = a, "manual override survives item order and rule edit"
    overrides.Item(a) = "same.txt": overrides.Item(b) = "same.txt"
    Set preview = NxRenameFileRulesPreview(paths, rules, overrides, hasDuplicate, hasCollision)
    Check hasDuplicate, "duplicate target preview"
    failed = False
    On Error Resume Next
    NxBatchRenameCommit preview
    failed = (Err.Number <> 0): Err.Clear
    On Error GoTo 0
    Check failed And ReadText(a) = "first" And ReadText(b) = "second", "invalid preview does not mutate"
    Set paths = New Collection: paths.Add a
    overrides.Item(a) = "CON.txt"
    Set preview = NxRenameFileRulesPreview(paths, rules, overrides, hasDuplicate, hasCollision)
    Check Left$(CStr(preview(1)(2)), 3) = "오류:", "reserved filename blocked"
    overrides.Item(a) = "..\escaped.txt"
    Set preview = NxRenameFileRulesPreview(paths, rules, overrides, hasDuplicate, hasCollision)
    Check Left$(CStr(preview(1)(2)), 3) = "오류:", "directory escape blocked"
    target = folder & "\unrelated.txt": WriteText target, "unrelated"
    overrides.Item(a) = "unrelated.txt"
    Set preview = NxRenameFileRulesPreview(paths, rules, overrides, hasDuplicate, hasCollision)
    Check hasCollision, "existing collision preview"
    overrides.Item(a) = "renamed.txt"
    Set preview = NxRenameFileRulesPreview(paths, rules, overrides, hasDuplicate, hasCollision)
    count = NxBatchRenameCommit(preview)
    target = folder & "\renamed.txt": held = folder & "\held-original.txt"
    Name target As held
    WriteText target, "replacement owned by another process"
    failed = False
    On Error Resume Next
    NxBatchRenameUndoLast True
    failed = (Err.Number <> 0): Err.Clear
    On Error GoTo 0
    Check failed And Not fso.FileExists(a), "undo rejects replacement identity before any move"
    Check ReadText(target) = "replacement owned by another process" And ReadText(held) = "first", "unrelated file preserved"
    T_R62RenameFiles = "PASS|folder,cycle,undo,override,duplicate,collision,reserved,escape,identity|fixture=" & folder
End Function

Public Function T_R62RenameSheets() As String
    Dim book As Workbook, first As Worksheet, second As Worksheet, order As New Collection, rules As New Collection
    Dim overrides As Object, preview As Collection, hasDuplicate As Boolean, hasCollision As Boolean, count As Long
    Dim errorNumber As Long, detail As String
    On Error GoTo Failed
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set first = book.Worksheets(1): first.Name = "Alpha"
    Set second = book.Worksheets.Add(After:=first): second.Name = "Beta"
    order.Add "Beta": order.Add "Alpha"
    Set overrides = CreateObject("Scripting.Dictionary")
    overrides.Add "Alpha", "Beta": overrides.Add "Beta", "Alpha"
    Set preview = NxRenameSheetRulesPreview(book, order, rules, overrides, hasDuplicate, hasCollision)
    Check CStr(preview(1)(0)) = "Beta", "preview item order"
    Check book.Worksheets(1) Is first, "item order does not move worksheets"
    count = NxSheetBatchRenameCommit(book, preview)
    Check count = 2 And first.Name = "Beta" And second.Name = "Alpha", "sheet cycle object binding"
    count = NxBatchRenameUndoLast(False, book)
    Check first.Name = "Alpha" And second.Name = "Beta", "sheet undo"
    book.Close SaveChanges:=False
    T_R62RenameSheets = "PASS|item-order,object-binding,cycle,undo"
    Exit Function
Failed:
    errorNumber = Err.Number: detail = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise errorNumber, "T_R62RenameSheets", detail
End Function

Public Function T_R62RenameRecovery(ByVal rootPath As String) As String
#If NX_R58_TEST_BUILD Then
    Dim folder As String, a As String, b As String, paths As New Collection, rules As New Collection
    Dim overrides As Object, preview As Collection, stage As Variant, duplicate As Boolean, collision As Boolean, failed As Boolean
    Dim book As Workbook, first As Worksheet, second As Worksheet, order As Collection
    Dim errorNumber As Long, detail As String
    On Error GoTo Failed
    folder = NewFixtureFolder(rootPath): a = folder & "\a.txt": b = folder & "\b.txt"
    WriteText a, "first": WriteText b, "second"
    paths.Add a: paths.Add b
    Set overrides = CreateObject("Scripting.Dictionary"): overrides.Add a, "b.txt": overrides.Add b, "a.txt"
    For Each stage In Array("original-to-temporary", "temporary-to-target")
        Set preview = NxRenameFileRulesPreview(paths, rules, overrides, duplicate, collision)
        NxBatchRenameTestInjectMoveFailure CStr(stage), 2
        failed = False
        On Error Resume Next
        NxBatchRenameCommit preview
        failed = (Err.Number <> 0): Err.Clear
        On Error GoTo 0
        NxBatchRenameTestClearMoveFailure
        Check failed And ReadText(a) = "first" And ReadText(b) = "second", "injected recovery " & CStr(stage)
    Next stage
    Set book = Workbooks.Add(xlWBATWorksheet)
    Set first = book.Worksheets(1): first.Name = "Alpha"
    Set second = book.Worksheets.Add(After:=first): second.Name = "Beta"
    Set order = New Collection: order.Add "Beta": order.Add "Alpha"
    Set overrides = CreateObject("Scripting.Dictionary"): overrides.Add "Alpha", "Beta": overrides.Add "Beta", "Alpha"
    For Each stage In Array("sheet-original-to-temporary", "sheet-temporary-to-target")
        Set preview = NxRenameSheetRulesPreview(book, order, rules, overrides, duplicate, collision)
        NxBatchRenameTestInjectMoveFailure CStr(stage), 2
        failed = False
        On Error Resume Next
        NxSheetBatchRenameCommit book, preview
        failed = (Err.Number <> 0): Err.Clear
        On Error GoTo Failed
        NxBatchRenameTestClearMoveFailure
        Check failed And first.Name = "Alpha" And second.Name = "Beta", "injected recovery " & CStr(stage)
    Next stage
    book.Close SaveChanges:=False
    T_R62RenameRecovery = "PASS|both-file-and-sheet-stages|fixture=" & folder
    Exit Function
Failed:
    errorNumber = Err.Number: detail = Err.Description
    On Error Resume Next
    NxBatchRenameTestClearMoveFailure
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise errorNumber, "T_R62RenameRecovery", detail
#Else
    T_R62RenameRecovery = "NOT_RUN|requires NX_R58_TEST_BUILD=True in engine and fixture"
#End If
End Function

Private Function NewFixtureFolder(ByVal rootPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(rootPath) Then Err.Raise vbObjectError + 6260, , "Fixture root must exist"
    NewFixtureFolder = fso.BuildPath(rootPath, "r62-rename-" & Replace$(NxCreateRunUuid(), "-", vbNullString))
    If fso.FolderExists(NewFixtureFolder) Then Err.Raise vbObjectError + 6261, , "Fixture path collision"
    fso.CreateFolder NewFixtureFolder
End Function

Private Sub WriteText(ByVal path As String, ByVal value As String)
    Dim fso As Object, stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.CreateTextFile(path, False, True)
    stream.Write value
    stream.Close
End Sub

Private Function ReadText(ByVal path As String) As String
    Dim fso As Object, stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.OpenTextFile(path, 1, False, -1)
    ReadText = stream.ReadAll
    stream.Close
End Function

Private Sub Check(ByVal condition As Boolean, ByVal detail As String)
    If Not condition Then Err.Raise vbObjectError + 6262, "T_R62RenameWorkbench", detail
End Sub
