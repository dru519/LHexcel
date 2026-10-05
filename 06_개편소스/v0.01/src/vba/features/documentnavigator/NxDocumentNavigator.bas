Attribute VB_Name = "NxDocumentNavigator"
Option Explicit

Private mNavigator As FNxDocumentNavigator
Private mNextToken As Long
Private mMoving As Boolean
Private mDllItems As Collection

Public Sub NxDocumentNavigatorOpen()
    If NxHostShowDocumentNavigator() Then Exit Sub
    If mNavigator Is Nothing Then Set mNavigator = New FNxDocumentNavigator
    If mNavigator.Visible Then
        mNavigator.RefreshNavigator
        mNavigator.FocusNavigator
    Else
        mNavigator.Show vbModeless
        ' Show can return before Activate is delivered to a modeless form.
        mNavigator.RefreshNavigator
        mNavigator.FocusNavigator
    End If
End Sub

Public Sub NxDocumentNavigatorShutdown()
    NxHostHideDocumentNavigator
    Set mDllItems = Nothing
    If mNavigator Is Nothing Then Exit Sub
    Unload mNavigator
    Set mNavigator = Nothing
End Sub

' Only opaque tokens and display data cross the DLL boundary, never Range/Workbook objects.
Public Function NxDocumentNavigatorDllSnapshot() As Variant
    Dim item As CNxDocumentNavigatorItem, count As Long, index As Long
    Dim rows() As Variant, include As Boolean
    Set mDllItems = NxDocumentNavigatorSnapshot(mDllItems)
    For Each item In mDllItems
        If DllDisplayItem(item) Then count = count + 1
    Next item
    If count = 0 Then NxDocumentNavigatorDllSnapshot = Array(): Exit Function
    If count > 10000 Then NxRaiseContractError "탐색 목록이 너무 큽니다."
    ReDim rows(0 To count - 1, 0 To 4)
    For Each item In mDllItems
        If DllDisplayItem(item) Then
            rows(index, 0) = CStr(item.TokenId)
            rows(index, 1) = vbNullString
            If Not item.ParentBook Is Nothing Then rows(index, 1) = CStr(item.ParentBook.TokenId)
            rows(index, 2) = item.DisplayName
            rows(index, 3) = item.Available
            If item.SheetRef Is Nothing Then
                rows(index, 4) = (item.BookRef Is Application.ActiveWorkbook)
            Else
                rows(index, 4) = (item.SheetRef Is Application.ActiveSheet)
            End If
            index = index + 1
        End If
    Next item
    NxDocumentNavigatorDllSnapshot = rows
End Function

Private Function DllDisplayItem(ByVal item As CNxDocumentNavigatorItem) As Boolean
    If item.SheetRef Is Nothing Then
        DllDisplayItem = True
    ElseIf TypeOf item.SheetRef Is Excel.Worksheet Then
        DllDisplayItem = (item.SheetRef.Visible <> xlSheetVeryHidden)
    End If
End Function

Public Function NxDocumentNavigatorDllMove(ByVal token As String) As Boolean
    Dim item As CNxDocumentNavigatorItem, reason As String
    If mDllItems Is Nothing Or mMoving Then Exit Function
    For Each item In mDllItems
        If CStr(item.TokenId) = token Then
            If Not DllDisplayItem(item) Then Exit Function
            NxDocumentNavigatorDllMove = NxDocumentNavigatorMove(item, reason)
            Exit Function
        End If
    Next item
End Function

Public Sub NxDocumentNavigatorRelease(ByVal navigator As FNxDocumentNavigator)
    If mNavigator Is navigator Then Set mNavigator = Nothing
End Sub

Public Function NxDocumentNavigatorSnapshot(ByVal previous As Collection) As Collection
    Dim result As New Collection
    Dim book As Excel.Workbook, sheet As Object
    Dim bookItem As CNxDocumentNavigatorItem, sheetItem As CNxDocumentNavigatorItem
    For Each book In Application.Workbooks
        If IsVisibleDocument(book) Then
            Set bookItem = ReuseItem(previous, book, Nothing)
            bookItem.DisplayName = book.Name
            bookItem.StateText = "문서"
            bookItem.Available = True
            Set bookItem.RecentWindow = ChooseWindow(bookItem)
            result.Add bookItem
            For Each sheet In book.Sheets
                Set sheetItem = ReuseItem(previous, book, sheet)
                Set sheetItem.ParentBook = bookItem
                sheetItem.DisplayName = sheet.Name
                sheetItem.Available = (sheet.Visible = xlSheetVisible) And IsSupportedSheet(sheet)
                If TypeOf sheet Is Excel.Worksheet Then
                    sheetItem.StateText = "시트"
                ElseIf TypeOf sheet Is Excel.Chart Then
                    sheetItem.StateText = "차트"
                Else
                    sheetItem.StateText = "지원하지 않는 시트"
                End If
                Select Case sheet.Visible
                    Case xlSheetHidden: sheetItem.StateText = sheetItem.StateText & " · 숨김"
                    Case xlSheetVeryHidden: sheetItem.StateText = sheetItem.StateText & " · 매우 숨김"
                End Select
                result.Add sheetItem
            Next sheet
        End If
    Next book
    Set NxDocumentNavigatorSnapshot = result
End Function

Private Function ReuseItem(ByVal previous As Collection, ByVal book As Excel.Workbook, ByVal sheet As Object) As CNxDocumentNavigatorItem
    Dim item As CNxDocumentNavigatorItem
    If Not previous Is Nothing Then
        For Each item In previous
            If book Is item.BookRef Then
                If sheet Is item.SheetRef Then
                    Set ReuseItem = item
                    Exit Function
                End If
            End If
        Next item
    End If
    Set item = New CNxDocumentNavigatorItem
    mNextToken = mNextToken + 1
    item.TokenId = mNextToken
    Set item.BookRef = book
    Set item.SheetRef = sheet
    Set ReuseItem = item
End Function

Private Function IsVisibleDocument(ByVal book As Excel.Workbook) As Boolean
    Dim window As Excel.Window
    If book.IsAddin Then Exit Function
    For Each window In book.Windows
        If window.Visible Then IsVisibleDocument = True: Exit Function
    Next window
End Function

Private Function IsSupportedSheet(ByVal sheet As Object) As Boolean
    IsSupportedSheet = (TypeOf sheet Is Excel.Worksheet) Or (TypeOf sheet Is Excel.Chart)
End Function

Public Function NxDocumentNavigatorCanMove(ByVal item As CNxDocumentNavigatorItem) As Boolean
    Dim book As Excel.Workbook, sheet As Object
    If item Is Nothing Then Exit Function
    On Error GoTo Unavailable
    For Each book In Application.Workbooks
        If book Is item.BookRef Then
            If Not IsVisibleDocument(book) Then Exit Function
            If item.SheetRef Is Nothing Then
                NxDocumentNavigatorCanMove = True
                Exit Function
            End If
            For Each sheet In book.Sheets
                If sheet Is item.SheetRef Then
                    NxDocumentNavigatorCanMove = (sheet.Visible = xlSheetVisible) And IsSupportedSheet(sheet)
                    Exit Function
                End If
            Next sheet
            Exit Function
        End If
    Next book
Unavailable:
End Function

Private Function ChooseWindow(ByVal bookItem As CNxDocumentNavigatorItem) As Excel.Window
    Dim window As Excel.Window, current As Excel.Window, firstVisible As Excel.Window
    Dim recent As Excel.Window
    If Application.ActiveWorkbook Is bookItem.BookRef Then Set current = Application.ActiveWindow
    For Each window In bookItem.BookRef.Windows
        If window.Visible Then
            If SameWindowHandle(window, current) Then Set ChooseWindow = window: Exit Function
            If firstVisible Is Nothing Then Set firstVisible = window
            If SameWindowHandle(window, bookItem.RecentWindow) Then Set recent = window
        End If
    Next window
    If Not recent Is Nothing Then Set ChooseWindow = recent Else Set ChooseWindow = firstVisible
End Function

Private Function SameWindowHandle(ByVal leftWindow As Excel.Window, ByVal rightWindow As Excel.Window) As Boolean
    ' Excel can return different COM wrappers for the same native window.
    ' Callers first scope candidates to a live, object-identified workbook.
    If leftWindow Is Nothing Or rightWindow Is Nothing Then Exit Function
    On Error GoTo Unavailable
    SameWindowHandle = (leftWindow.Hwnd <> 0 And leftWindow.Hwnd = rightWindow.Hwnd)
Unavailable:
End Function

Private Function IsActiveTargetWindow(ByVal bookItem As CNxDocumentNavigatorItem, ByVal targetWindow As Excel.Window) As Boolean
    If Not Application.ActiveWorkbook Is bookItem.BookRef Then Exit Function
    IsActiveTargetWindow = SameWindowHandle(Application.ActiveWindow, targetWindow)
End Function

Public Function NxDocumentNavigatorMove(ByVal item As CNxDocumentNavigatorItem, ByRef reason As String) As Boolean
    Dim bookItem As CNxDocumentNavigatorItem
    Dim targetWindow As Excel.Window
    Dim currentSheet As Object
    reason = vbNullString
    If mMoving Then
        reason = "문서 이동을 처리 중입니다. 잠시 후 다시 시도하세요."
        Exit Function
    End If
    mMoving = True
    On Error GoTo Failed
    If Not NxDocumentNavigatorCanMove(item) Then
        reason = "대상이 닫혔거나 숨겨져 이동할 수 없습니다. 목록을 새로고침하세요."
        GoTo Finished
    End If
    If item.ParentBook Is Nothing Then Set bookItem = item Else Set bookItem = item.ParentBook
    Set targetWindow = ChooseWindow(bookItem)
    If targetWindow Is Nothing Then
        reason = "이동할 수 있는 문서 창이 없습니다."
        GoTo Finished
    End If
    If Not IsActiveTargetWindow(bookItem, targetWindow) Then targetWindow.Activate
    ' Activation can run user handlers: check identity/visibility again before sheet access.
    If Not IsActiveTargetWindow(bookItem, targetWindow) Then GoTo ChangedDuringMove
    If Not NxDocumentNavigatorCanMove(item) Then
        reason = "문서 전환 중 대상이 변경되었습니다. 목록을 새로고침하세요."
        GoTo Finished
    End If
    If Not item.SheetRef Is Nothing Then
        Set currentSheet = Application.ActiveSheet
        If Not item.SheetRef Is currentSheet Then item.SheetRef.Activate
    End If
    If Not IsActiveTargetWindow(bookItem, targetWindow) Then GoTo ChangedDuringMove
    If Not item.SheetRef Is Nothing Then
        Set currentSheet = Application.ActiveSheet
        If Not currentSheet Is item.SheetRef Then GoTo ChangedDuringMove
    End If
    Set bookItem.RecentWindow = targetWindow
    NxDocumentNavigatorMove = True
Finished:
    mMoving = False
    Exit Function
ChangedDuringMove:
    reason = "문서 이벤트가 이동 대상을 바꿨습니다. 현재 창을 확인하세요."
    GoTo Finished
Failed:
    reason = "이동하지 못했습니다: " & Err.Description
    GoTo Finished
End Function
