VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxRoleStylePreview"
Option Explicit

Private mWorkbookIdentity As String, mWorksheetIdentity As String, mTargetAddress As String, mTargetDigest As String
Private mRoleCode As String, mDisplayMode As String, mAxisCode As String, mRelativeIndexes As String, mSealed As Boolean

Public Sub Capture(ByVal target As Range, ByVal roleCode As String, ByVal displayMode As String, ByVal axisCode As String, ByVal relativeIndexes As String, ByVal targetDigest As String)
    mWorkbookIdentity = target.Worksheet.Parent.Name: mWorksheetIdentity = target.Worksheet.CodeName
    mTargetAddress = target.Address(External:=True): mTargetDigest = targetDigest
    mRoleCode = roleCode: mDisplayMode = displayMode: mAxisCode = axisCode: mRelativeIndexes = relativeIndexes
    mSealed = False
End Sub

Public Sub Seal()
    If Len(mTargetAddress) = 0 Or Len(mTargetDigest) = 0 Then NxRaiseContractError "role-style preview cannot be sealed before capture"
    mSealed = True
End Sub

Public Property Get IsSealed() As Boolean: IsSealed = mSealed: End Property
Public Property Get WorkbookIdentity() As String: WorkbookIdentity = mWorkbookIdentity: End Property
Public Property Get WorksheetIdentity() As String: WorksheetIdentity = mWorksheetIdentity: End Property
Public Property Get TargetAddress() As String: TargetAddress = mTargetAddress: End Property
Public Property Get TargetDigest() As String: TargetDigest = mTargetDigest: End Property
Public Property Get RoleCode() As String: RoleCode = mRoleCode: End Property
Public Property Get DisplayMode() As String: DisplayMode = mDisplayMode: End Property
Public Property Get AxisCode() As String: AxisCode = mAxisCode: End Property
Public Property Get RelativeIndexes() As String: RelativeIndexes = mRelativeIndexes: End Property
