VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INxFeatureCommand"
Option Explicit

' Plans are approved by CNxApproval; commands never receive reusable string authorization.

Public Property Get FeatureId() As String
End Property

Public Function BuildPlan(ByVal context As CNxExecutionContext) As CNxPlan
End Function

Public Function Execute(ByVal context As CNxExecutionContext, ByVal approvedPlan As CNxPlan) As CNxResult
End Function
