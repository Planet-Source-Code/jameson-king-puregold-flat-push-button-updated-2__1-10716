Attribute VB_Name = "MouseOverFunction"
'Greg Siemon gsiemon@home.net
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Function MouseOver(Control As Object, X As Single, Y As Single, _
        Optional NewScaleLeft As Variant, Optional NewScaleWidth As Variant, _
        Optional NewScaleTop As Variant, Optional NewScaleHeight As Variant) As Boolean

Static MouseOverSW As Boolean, Ret As Long
Dim ControlScaleLeft As Single, ControlScaleWidth As Single
Dim ControlScaleTop As Single, ControlScaleHeight As Single

If TypeOf Control Is PictureBox Or _
    TypeOf Control Is Form Or _
    TypeOf Control Is UserDocument Or _
    TypeOf Control Is MDIForm Then
                        
    ControlScaleLeft = Control.ScaleLeft
    ControlScaleWidth = Control.ScaleWidth
    ControlScaleTop = Control.ScaleTop
    ControlScaleHeight = Control.ScaleHeight
    
    If Not IsMissing(NewScaleLeft) Then ControlScaleLeft = NewScaleLeft
    If Not IsMissing(NewScaleWidth) Then ControlScaleWidth = NewScaleWidth
    If Not IsMissing(NewScaleTop) Then ControlScaleTop = NewScaleTop
    If Not IsMissing(NewScaleHeight) Then ControlScaleHeight = NewScaleHeight

    If Sgn(ControlScaleWidth) = Sgn(ControlScaleLeft - X) Or _
        Sgn(ControlScaleWidth) <> Sgn(ControlScaleWidth + ControlScaleLeft - X) Or _
        Sgn(ControlScaleHeight) = Sgn(ControlScaleTop - Y) Or _
        Sgn(ControlScaleHeight) <> Sgn(ControlScaleHeight + ControlScaleTop - Y) Then
        MouseOverSW = False
        Ret = ReleaseCapture()
        Else
        Ret = SetCapture(Control.hwnd)
        MouseOverSW = True
        End If
    End If
MouseOver = MouseOverSW
End Function


