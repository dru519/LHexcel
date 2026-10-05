Attribute VB_Name = "NxFocusInvalidateStub"
Option Explicit

Private mInvalidationCount As Long

Public Sub NxRibbonInvalidateFocus()
    mInvalidationCount = mInvalidationCount + 1
End Sub

Public Sub NxFocusInvalidateStubReset()
    mInvalidationCount = 0
End Sub

Public Function NxFocusInvalidateStubCount() As Long
    NxFocusInvalidateStubCount = mInvalidationCount
End Function
