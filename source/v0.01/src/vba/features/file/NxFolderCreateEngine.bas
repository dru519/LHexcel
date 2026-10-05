Attribute VB_Name = "NxFolderCreateEngine"
Option Explicit

Public Function NxFolderCreateItemsFromRange(ByVal source As Range, Optional ByVal prefix As String = vbNullString, Optional ByVal suffix As String = vbNullString) As Collection
    Dim items As New Collection, seen As Object, rowIndex As Long, columnIndex As Long
    Dim cell As Range, part As String, relativePath As String, blankSeen As Boolean
    Dim item As CNxFolderCreateItem
    If source Is Nothing Then NxRaiseContractError "폴더 이름이 있는 엑셀 범위를 선택하세요."
    If source.Areas.Count <> 1 Then NxRaiseContractError "폴더 목록은 연속된 한 범위를 선택하세요."
    If source.Rows.Count > NX_FOLDER_CREATE_LIMIT Or source.Cells.CountLarge > 5000 Then NxRaiseContractError "폴더 목록은 최대 500행, 5,000셀까지 선택할 수 있습니다."
    Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
    For rowIndex = 1 To source.Rows.Count
        relativePath = vbNullString: blankSeen = False
        For columnIndex = 1 To source.Columns.Count
            Set cell = source.Cells(rowIndex, columnIndex)
            If IsError(cell.Value2) Then NxRaiseContractError "폴더 이름에 오류 값이 있습니다: " & cell.Address(False, False)
            part = Trim$(CStr(cell.Value2))
            If Len(part) = 0 Then
                blankSeen = True
            Else
                If blankSeen Then NxRaiseContractError "상위 폴더가 비어 있습니다. 왼쪽 열부터 입력하세요: " & cell.Address(False, False)
                part = prefix & part & suffix
                If Not NxFolderNameIsValid(part) Then NxRaiseContractError "사용할 수 없는 폴더 이름입니다: " & cell.Address(False, False) & " · " & part
                If Len(relativePath) > 0 Then relativePath = relativePath & "\"
                relativePath = relativePath & part
            End If
        Next columnIndex
        If Len(relativePath) > 0 Then
            If Not seen.Exists(relativePath) Then
                Set item = New CNxFolderCreateItem: item.Configure relativePath
                items.Add item: seen.Add relativePath, True
            End If
        End If
    Next rowIndex
    If items.Count = 0 Then NxRaiseContractError "선택한 범위에 폴더 이름이 없습니다."
    Set NxFolderCreateItemsFromRange = items
End Function

Public Function NxFolderCreateItemsText(ByVal items As Collection) As String
    Dim item As CNxFolderCreateItem
    For Each item In items
        NxFolderCreateItemsText = NxFolderCreateItemsText & item.RelativePath & vbCrLf
    Next item
End Function

' Read-only planning. Core owns NxDirectoryBatchCreate mutation and rollback.
Public Function NxFolderCreatePlan(ByVal baseFolder As String, ByVal items As Collection, ByRef directoryPaths As Collection, ByRef detail As String) As Boolean
    Dim item As CNxFolderCreateItem, finalPath As String, fso As Object, current As String, part As Variant, parts As Variant, seen As Object, requested As Object
    If items Is Nothing Then detail = "폴더 입력은 1~500개여야 합니다.": Exit Function
    If items.Count < 1 Or items.Count > NX_FOLDER_CREATE_LIMIT Then detail = "폴더 입력은 1~500개여야 합니다.": Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject"): If Not fso.FolderExists(baseFolder) Then detail = "승인된 기준 폴더가 존재하지 않습니다.": Exit Function
    Set directoryPaths = New Collection: Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = 1: Set requested = CreateObject("Scripting.Dictionary"): requested.CompareMode = 1
    For Each item In items
        If item.SelectedForCreate Then
            If Not NxFolderPathValidate(baseFolder, item.RelativePath, finalPath) Then detail = "폴더 경로 안전 규칙에서 제외됨: " & item.RelativePath: Exit Function
            If requested.Exists(finalPath) Then detail = "중복 폴더 요청: " & item.RelativePath: Exit Function
            requested.Add finalPath, True: parts = Split(Replace$(item.RelativePath, "/", "\"), "\"): current = baseFolder
            For Each part In parts
                current = current & "\" & CStr(part)
                If Not fso.FolderExists(current) Then If Not seen.Exists(LCase$(current)) Then seen.Add LCase$(current), True: directoryPaths.Add current
            Next part
        End If
    Next item
    If directoryPaths.Count = 0 Then detail = "요청한 폴더가 모두 이미 존재합니다.": Exit Function
    If directoryPaths.Count > NX_FOLDER_CREATE_LIMIT Then detail = "생성 예정 폴더 경로는 최대 500개입니다.": Exit Function
    detail = "생성 예정=" & CStr(directoryPaths.Count): NxFolderCreatePlan = True
End Function
