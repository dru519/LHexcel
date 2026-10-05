VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNxFileExecutionCancelAdapter"
Option Explicit

Implements INxExecutionCancelAdapter

Private Sub INxExecutionCancelAdapter_RequestCancel()
    NxFileOperationRequestCancel
End Sub
