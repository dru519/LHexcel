Attribute VB_Name = "NxSymbolsWheel"
Option Explicit

Public Function NxSymbolsWheelAttach(ByVal host As FNxSymbols) As Boolean
    NxSymbolsWheelAttach = NxFormWheelAttach(host)
End Function

Public Sub NxSymbolsWheelDetach(ByVal host As FNxSymbols)
    NxFormWheelDetach host
End Sub

Public Function NxSymbolsWheelIsAttached() As Boolean
    NxSymbolsWheelIsAttached = NxFormWheelIsAttached()
End Function
