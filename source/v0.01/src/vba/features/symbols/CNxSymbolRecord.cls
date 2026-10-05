VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxSymbolRecord"
Option Explicit

Private mCodePoint As Long
Private mCharacter As String
Private mCodeText As String
Private mNameKo As String
Private mAliasesKo As String
Private mCategoryId As String
Private mBlockId As String
Private mSearchText As String
Private mSortOrder As Long
Private mSealed As Boolean

Public Sub Configure(ByVal codePoint As Long, ByVal nameKo As String, ByVal aliasesKo As String, _
                     ByVal categoryId As String, ByVal blockId As String, ByVal sortOrder As Long)
    If mSealed Then NxRaiseContractError "기호 레코드는 한 번만 구성할 수 있습니다."
    If Not NxSymbolIsAllowedCodePoint(codePoint) Then NxRaiseContractError "기호 코드가 잘못되었습니다."
    If Len(Trim$(nameKo)) = 0 Or Len(Trim$(categoryId)) = 0 Or Len(Trim$(blockId)) = 0 Then NxRaiseContractError "기호 메타데이터가 비어 있습니다."
    If sortOrder <= 0 Then NxRaiseContractError "기호 순서가 잘못되었습니다."
    mCodePoint = codePoint
    mCharacter = NxSymbolFromCodePoint(codePoint)
    mCodeText = NxSymbolFormatCode(codePoint)
    mNameKo = nameKo
    mAliasesKo = aliasesKo
    mCategoryId = categoryId
    mBlockId = blockId
    mSortOrder = sortOrder
    mSearchText = LCase$(mCharacter & " " & mCodeText & " " & Hex$(codePoint) & " " & nameKo & " " & aliasesKo & " " & categoryId & " " & blockId)
    mSealed = True
End Sub

Public Property Get CodePoint() As Long: RequireSealed: CodePoint = mCodePoint: End Property
Public Property Get Character() As String: RequireSealed: Character = mCharacter: End Property
Public Property Get CodeText() As String: RequireSealed: CodeText = mCodeText: End Property
Public Property Get NameKo() As String: RequireSealed: NameKo = mNameKo: End Property
Public Property Get AliasesKo() As String: RequireSealed: AliasesKo = mAliasesKo: End Property
Public Property Get CategoryId() As String: RequireSealed: CategoryId = mCategoryId: End Property
Public Property Get BlockId() As String: RequireSealed: BlockId = mBlockId: End Property
Public Property Get SearchText() As String: RequireSealed: SearchText = mSearchText: End Property
Public Property Get SortOrder() As Long: RequireSealed: SortOrder = mSortOrder: End Property

Private Sub RequireSealed()
    If Not mSealed Then NxRaiseContractError "기호 레코드가 구성되지 않았습니다."
End Sub
