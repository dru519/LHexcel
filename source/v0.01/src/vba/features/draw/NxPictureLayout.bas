Attribute VB_Name = "NxPictureLayout"
Option Explicit

Public Sub NxDrawFitPicture(ByVal pictureShape As Shape, ByVal target As Range, Optional ByVal pictureMargin As Double = 4#, Optional ByVal moveAndSize As Boolean = True)
    Dim availableWidth As Double, availableHeight As Double, scaleFactor As Double
    Dim originalWidth As Double, originalHeight As Double
    Dim oldLeft As Double, oldTop As Double, oldWidth As Double, oldHeight As Double, oldPlacement As XlPlacement
    Dim oldLockAspectRatio As MsoTriState
    Dim failureNumber As Long, failureSource As String, failureDescription As String
    If pictureShape Is Nothing Or target Is Nothing Then NxRaiseContractError "그림과 대상 셀 범위가 필요합니다."
    If target.Areas.Count <> 1 Or pictureMargin < 0 Then NxRaiseContractError "그림 대상은 연속 범위와 유효한 여백이어야 합니다."
    If target.Worksheet.ProtectDrawingObjects Or target.Worksheet.ProtectContents Then NxRaiseContractError "보호된 시트에는 그림을 배치할 수 없습니다."
    availableWidth = target.Width - pictureMargin * 2#: availableHeight = target.Height - pictureMargin * 2#
    If availableWidth <= 0 Or availableHeight <= 0 Or pictureShape.Width <= 0 Or pictureShape.Height <= 0 Then NxRaiseContractError "그림 또는 대상 영역 크기가 0입니다."
    oldLeft = pictureShape.Left: oldTop = pictureShape.Top: oldWidth = pictureShape.Width: oldHeight = pictureShape.Height: oldPlacement = pictureShape.Placement: oldLockAspectRatio = pictureShape.LockAspectRatio
    originalWidth = pictureShape.Width: originalHeight = pictureShape.Height
    On Error GoTo Failed
    scaleFactor = WorksheetFunction.Min(availableWidth / originalWidth, availableHeight / originalHeight)
    pictureShape.LockAspectRatio = msoTrue: pictureShape.Width = originalWidth * scaleFactor
    pictureShape.Left = target.Left + (target.Width - pictureShape.Width) / 2#: pictureShape.Top = target.Top + (target.Height - pictureShape.Height) / 2#
    If moveAndSize Then pictureShape.Placement = xlMoveAndSize
    Exit Sub
Failed:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    On Error Resume Next
    pictureShape.LockAspectRatio = msoFalse: pictureShape.Left = oldLeft: pictureShape.Top = oldTop: pictureShape.Width = oldWidth: pictureShape.Height = oldHeight: pictureShape.Placement = oldPlacement: pictureShape.LockAspectRatio = oldLockAspectRatio
    On Error GoTo 0
    Err.Raise failureNumber, failureSource, failureDescription
End Sub
