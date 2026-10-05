Attribute VB_Name = "NxNativePreview"
Option Explicit
Private mService As Object

Private Function PreviewService() As Object
    If mService Is Nothing Then
        Set mService = NxHostCreatePicturePreview()
    End If
    Set PreviewService = mService
End Function

Public Sub NxNativePreviewTextChanged(ByVal view As Object, ByVal text As Object)
    Dim service As Object, image As Object, key As String, width As Single
    On Error GoTo Fallback
    Set service = PreviewService()
    If service Is Nothing Then Exit Sub
    key = "nxNative_" & text.Name
    On Error Resume Next
    Set image = view.Controls(key)
    On Error GoTo Fallback
    If image Is Nothing Then
        Set image = view.Controls.Add("Forms.Image.1", key, True)
        width = text.Width
        image.Left = text.Left: image.Top = text.Top
        image.Width = width * 0.62: image.Height = text.Height
        image.Tag = CStr(width)
        image.PictureSizeMode = 3: image.BorderStyle = 0
    End If
    Set image.Picture = service.RenderSummary(CStr(text.Value), CLng(image.Width), CLng(image.Height))
    text.Left = image.Left + image.Width + 6
    text.Width = CSng(image.Tag) - image.Width - 6
    text.BackColor = vbWhite
    text.ScrollBars = 2
    image.Visible = True
    Exit Sub
Fallback:
    Err.Clear
    If Not image Is Nothing Then
        text.Left = image.Left: text.Width = CSng(image.Tag)
        image.Visible = False
    End If
End Sub

Public Sub NxNativePreviewLabels(ByVal view As Object, ByVal prefix As String, ByVal originLeft As Single, _
        ByVal top As Single, ByVal width As Single, ByVal height As Single)
    Dim service As Object, image As Object, item As Object, cells() As Variant, count As Long, index As Long
    On Error GoTo Fallback
    Set service = PreviewService()
    If service Is Nothing Then Exit Sub
    For Each item In view.Controls
        If Left$(item.Name, Len(prefix)) = prefix And item.Visible Then count = count + 1
    Next item
    If count = 0 Then Exit Sub
    ReDim cells(1 To count, 1 To 10)
    For Each item In view.Controls
        If Left$(item.Name, Len(prefix)) = prefix And item.Visible Then
            index = index + 1
            cells(index, 1) = item.Left - originLeft: cells(index, 2) = item.Top - top
            cells(index, 3) = item.Width: cells(index, 4) = item.Height
            cells(index, 5) = item.Caption: cells(index, 6) = item.BackColor
            cells(index, 7) = item.ForeColor: cells(index, 8) = item.Font.Size
            cells(index, 9) = item.Font.Bold: cells(index, 10) = item.TextAlign
        End If
    Next item
    On Error Resume Next
    Set image = view.Controls("nxNativeGrid")
    On Error GoTo Fallback
    If image Is Nothing Then Set image = view.Controls.Add("Forms.Image.1", "nxNativeGrid", True)
    image.Left = originLeft: image.Top = top: image.Width = width: image.Height = height
    image.PictureSizeMode = 3: image.BorderStyle = 0
    Set image.Picture = service.RenderCells(cells, CLng(width), CLng(height))
    image.Visible = True: image.ZOrder 0
    Exit Sub
Fallback:
    Err.Clear
    If Not image Is Nothing Then image.Visible = False
End Sub

Public Sub NxNativePreviewHide(ByVal view As Object)
    On Error Resume Next
    view.Controls("nxNativeGrid").Visible = False
End Sub
