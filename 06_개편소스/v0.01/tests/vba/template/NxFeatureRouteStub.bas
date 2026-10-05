Attribute VB_Name = "NxFeatureRouteStub"
Option Explicit

Private mLastFeatureId As String

Public Sub NxRouteFeature(ByVal featureId As String)
    If Len(featureId) = 0 Then NxRaiseContractError "Feature route stub requires a feature ID"
    mLastFeatureId = featureId
End Sub

Public Sub NxFeatureRouteStubReset()
    mLastFeatureId = vbNullString
End Sub

Public Function NxFeatureRouteStubLastFeatureId() As String
    NxFeatureRouteStubLastFeatureId = mLastFeatureId
End Function
