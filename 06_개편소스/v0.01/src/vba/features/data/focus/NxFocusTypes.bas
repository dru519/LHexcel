Attribute VB_Name = "NxFocusTypes"
Option Explicit

Public Enum NxFocusState
    NxFocusOff = 0
    NxFocusEnabling = 1
    NxFocusOn = 2
    NxFocusDisabling = 3
    NxFocusDegraded = 4
End Enum

Public Const NX_FOCUS_SHAPE_CROSS As String = "criss-cross"
Public Const NX_FOCUS_SHAPE_HORIZONTAL As String = "horizontal"
Public Const NX_FOCUS_SHAPE_VERTICAL As String = "vertical"
Public Const NX_FOCUS_STYLE_WIDE As String = "wide-stripe"
Public Const NX_FOCUS_MAX_SELECTION_CELLS As Double = 100000#
Public Const NX_FOCUS_APPLIES_ROWS As Long = 4096
Public Const NX_FOCUS_APPLIES_COLUMNS As Long = 256
