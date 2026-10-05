Attribute VB_Name = "NxProductSettings"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function NxOpenShell Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hwnd As LongPtr, ByVal operation As LongPtr, ByVal file As LongPtr, ByVal parameters As LongPtr, ByVal directory As LongPtr, ByVal show As Long) As LongPtr
#Else
Private Declare Function NxOpenShell Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hwnd As Long, ByVal operation As Long, ByVal file As Long, ByVal parameters As Long, ByVal directory As Long, ByVal show As Long) As Long
#End If

Public Sub NxProductSettingsOpenMenu()
    ' The management menu is rendered by NxRibbonManagementMenuXml.
End Sub

Public Sub NxProductInstall()
    Dim addIn As Object
    NxDistributionEnsureStandaloneXlam
    NxDistributionEnsureExecutable
    If ThisWorkbook Is Nothing Then NxRaiseContractError "Product workbook is required"
    If Len(ThisWorkbook.Path) = 0 Then NxRaiseContractError "Product workbook must be saved before installation"
    Set addIn = NxFindProductAddIn()
    If addIn Is Nothing Then Set addIn = Application.AddIns.Add(ThisWorkbook.FullName, False)
    addIn.Installed = True
End Sub

Public Sub NxProductRemove()
    Dim addIn As Object
    NxDistributionEnsureStandaloneXlam
    If ThisWorkbook Is Nothing Then NxRaiseContractError "Product workbook is required"
    Set addIn = NxFindProductAddIn()
    If addIn Is Nothing Then NxRaiseContractError "Installed product add-in was not found"
    addIn.Installed = False
End Sub

Public Sub NxProductOpenSavedFolder()
    If ActiveWorkbook Is Nothing Then NxRaiseContractError "열려 있는 엑셀 파일이 없습니다."
    If ActiveWorkbook Is ThisWorkbook Then NxRaiseContractError "폴더를 열 엑셀 파일을 먼저 선택해 주세요."
    If Len(ActiveWorkbook.Path) = 0 Then NxRaiseContractError "아직 저장되지 않은 엑셀 파일입니다. 먼저 저장해 주세요."
    NxProductOpenFolder ActiveWorkbook.Path
End Sub

Public Sub NxProductOpenInstallFolder()
    If Len(ThisWorkbook.Path) = 0 Then NxRaiseContractError "내엑셀 설치 폴더를 찾을 수 없습니다."
    NxProductOpenFolder ThisWorkbook.Path
End Sub

Private Sub NxProductOpenFolder(ByVal folderPath As String)
    Dim fso As Object
    Dim executable As String, parameters As String, verb As String
#If VBA7 Then
    Dim result As LongPtr
#Else
    Dim result As Long
#End If
    If LCase$(Left$(folderPath, 8)) = "https://" Then
        ThisWorkbook.FollowHyperlink Address:=folderPath
        Exit Sub
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then NxRaiseContractError "폴더를 찾을 수 없습니다. 경로나 연결 상태를 확인해 주세요: " & folderPath
    If InStr(folderPath, Chr$(34)) > 0 Then NxRaiseContractError "폴더 경로가 올바르지 않습니다."
    executable = Environ$("WINDIR") & "\explorer.exe"
    parameters = "/n,/e," & Chr$(34) & folderPath & Chr$(34)
    verb = "open"
    result = NxOpenShell(Application.hwnd, StrPtr(verb), StrPtr(executable), StrPtr(parameters), 0, 1)
    If result <= 32 Then NxRaiseContractError "폴더 창을 열지 못했습니다. 오류 코드: " & CStr(result)
End Sub

Public Sub NxProductOpenQatSettings()
    Application.CommandBars.ExecuteMso "QuickAccessToolbarMoreCommands"
End Sub

Public Sub NxProductOpenExcelAddInsSettings()
    Application.CommandBars.ExecuteMso "AddInManager"
End Sub

Public Sub NxProductShowFileInfo()
    NxProductShowAbout
End Sub

Public Sub NxProductOpenShortcutManager()
    NxShortcutsShowManager
End Sub

Private Function NxFindProductAddIn() As Object
    Dim candidate As Object
    For Each candidate In Application.AddIns
        If StrComp(candidate.FullName, ThisWorkbook.FullName, vbTextCompare) = 0 Then
            Set NxFindProductAddIn = candidate
            Exit Function
        End If
    Next candidate
End Function

Public Sub NxInstallProduct(): NxProductInstall: End Sub
Public Sub NxRemoveProduct(): NxProductRemove: End Sub
Public Sub NxOpenSavedFolder(): NxProductOpenSavedFolder: End Sub
Public Sub NxOpenQatSettings(): NxProductOpenQatSettings: End Sub
Public Sub NxOpenExcelAddInsSettings(): NxProductOpenExcelAddInsSettings: End Sub
Public Sub NxShowFileInfo(): NxProductShowFileInfo: End Sub
Public Sub NxOpenShortcutManager(): NxProductOpenShortcutManager: End Sub
