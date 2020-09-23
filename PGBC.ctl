VERSION 5.00
Begin VB.UserControl PGBC 
   AutoRedraw      =   -1  'True
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   ScaleHeight     =   540
   ScaleWidth      =   1200
   ToolboxBitmap   =   "PGBC.ctx":0000
End
Attribute VB_Name = "PGBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Event Declarations:
Event MouseEnter()
Event MouseExit()
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Hide() 'MappingInfo=UserControl,UserControl,-1,Hide
Attribute Hide.VB_Description = "Occurs when the control's Visible property changes to False."
Event Show() 'MappingInfo=UserControl,UserControl,-1,Show
Attribute Show.VB_Description = "Occurs when the control's Visible property changes to True."
Public Enum Style
 Flat = 1
 CPlus = 2
 DarkFlat = 3
 WhiteOutLine = 4
 ThickFlat = 5
End Enum
Public Enum Align
    aRight = 1
    aLeft = 2
    aCenter = 3
End Enum
'Default Property Values:
Const m_def_BorderStyleC = 1
'Const m_def_borderstyle = 1
Const m_def_Hover = False
'Const m_def_BorderStyleThick = False
'Const m_def_BorderStyleThick = False
Const m_def_UpT = True
Const m_def_FocusRect = False
'Const m_def_Hover = 1
Const m_def_Caption = "Caption"
Const m_def_AlignText = aCenter
'Property Variables:
Dim Custom As Boolean
Dim Custom2 As Boolean
Dim m_AlignText As Align
Dim m_BorderStyleC As Style
'Dim m_BorderStyle As Variant
Dim m_Hover As Boolean
Dim m_UpT As Boolean
'Dim m_UpT As Boolean
Dim m_FocusRect As Boolean
'Dim m_Hover As Integer
Dim m_Caption As String
'Propertie
Dim FocusT As Boolean
'Dim upT As Boolean
Dim mH As Boolean
 
 
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    RealPaint
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    UserControl_Paint
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    RealPaint
    PropertyChanged "Font"
End Property


Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub



Private Sub Timer1_Timer()

End Sub

Private Sub Hv_Timer()

End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_EnterFocus()
FocusT = True
RealPaint
End Sub

Private Sub UserControl_ExitFocus()
FocusT = False
m_UpT = True
RealPaint
End Sub

Private Sub UserControl_GotFocus()
FocusT = True
RealPaint
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
m_UpT = False
RealPaint
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
'RealPaint "Up"
m_UpT = True
RealPaint
End Sub

Private Sub UserControl_LostFocus()
FocusT = False
m_UpT = True
RealPaint
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'RealPaint "Down"
m_UpT = False
RealPaint
RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
' Detecting if mouse moved!
RaiseEvent MouseMove(Button, Shift, x, y)
If m_Hover <> True Then GoTo 5
Dim Rt As RECT
Dim X1 As Single
Dim Y1 As Single
GetWindowRect UserControl.hwnd, Rt
X1 = Rt.Left
Y1 = Rt.Top
If MoveM <> True Then
        SetCapture (UserControl.hwnd)
        RaiseEvent MouseEnter
        mH = True
        RealPaint
    Else:
        RaiseEvent MouseExit
        mH = False
        RealPaint
        ReleaseCapture
End If
5
End Sub

Function MoveM()
' This is pretty self explanatory
Dim Pt As POINTAPI
Dim Rt As RECT
GetCursorPos Pt
GetWindowRect UserControl.hwnd, Rt
 If Pt.x >= Rt.Left And Pt.x <= Rt.Right And Pt.y >= Rt.Top And Pt.y <= Rt.Bottom Then
 SetCapture (UserControl.hwnd)
 MoveM = False
 Else:
 ReleaseCapture
 MoveM = True
 End If
End Function

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'RealPaint "Up"
m_UpT = True
RealPaint
RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,AutoRedraw
Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = UserControl.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    UserControl.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,CurrentX
Public Property Get CurrentX() As Single
Attribute CurrentX.VB_Description = "Returns/sets the horizontal coordinates for next print or draw method."
    CurrentX = UserControl.CurrentX
End Property

Public Property Let CurrentX(ByVal New_CurrentX As Single)
    UserControl.CurrentX() = New_CurrentX
    PropertyChanged "CurrentX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,CurrentY
Public Property Get CurrentY() As Single
Attribute CurrentY.VB_Description = "Returns/sets the vertical coordinates for next print or draw method."
    CurrentY = UserControl.CurrentY
End Property

Public Property Let CurrentY(ByVal New_CurrentY As Single)
    UserControl.CurrentY() = New_CurrentY
    PropertyChanged "CurrentY"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = UserControl.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    UserControl.FontBold() = New_FontBold
    RealPaint
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    UserControl.FontItalic() = New_FontItalic
    RealPaint
    PropertyChanged "FontItalic"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,FontName
'Public Property Get FontName() As String
'    FontName = UserControl.FontName
'End Property
'
'Public Property Let FontName(ByVal New_FontName As String)
'    UserControl.FontName() = New_FontName
'    PropertyChanged "FontName"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    UserControl.FontSize() = New_FontSize
    RealPaint
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = UserControl.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    UserControl.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    UserControl.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hdc = UserControl.hdc
End Property

Private Sub UserControl_Hide()
    RaiseEvent Hide
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = UserControl.Image
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
    UserControl.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleLeft
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_Description = "Returns/sets the horizontal coordinates for the left edges of an object."
    ScaleLeft = UserControl.ScaleLeft
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
    UserControl.ScaleLeft() = New_ScaleLeft
    PropertyChanged "ScaleLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleMode
Public Property Get ScaleMode() As Integer
Attribute ScaleMode.VB_Description = "Returns/sets a value indicating measurement units for object coordinates when using graphics methods or positioning controls."
    ScaleMode = UserControl.ScaleMode
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
    UserControl.ScaleMode() = New_ScaleMode
    PropertyChanged "ScaleMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleTop
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_Description = "Returns/sets the vertical coordinates for the top edges of an object."
    ScaleTop = UserControl.ScaleTop
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
    UserControl.ScaleTop() = New_ScaleTop
    PropertyChanged "ScaleTop"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
    ScaleWidth = UserControl.ScaleWidth
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
    UserControl.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleX
Public Function ScaleX(ByVal Width As Single, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
Attribute ScaleX.VB_Description = "Converts the value for the width of a Form, PictureBox, or Printer from one unit of measure to another."
    ScaleX = UserControl.ScaleX(Width, FromScale, ToScale)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleY
Public Function ScaleY(ByVal Height As Single, Optional ByVal FromScale As Variant, Optional ByVal ToScale As Variant) As Single
Attribute ScaleY.VB_Description = "Converts the value for the height of a Form, PictureBox, or Printer from one unit of measure to another."
    ScaleY = UserControl.ScaleY(Height, FromScale, ToScale)
End Function


Private Sub UserControl_Paint()
RealPaint
End Sub
Private Sub UserControl_TrackMouse()

End Sub
Private Sub RealPaint()
Set moPaintEffects = New PaintEffects
Dim clrMask As OLE_COLOR
Dim clrFace As OLE_COLOR
Dim clrDark As OLE_COLOR
Dim clrLight As OLE_COLOR
Dim Rt As RECT
' Rt is Rect
' FocusT is the option for drawing the Focus Rectangle
' BSO = Button Down State
' BSO2 = For bttons with two seperate API's for sides this is the second Down Setting
' BRI = Button Up State
' BRI2 = For buttons with two API's for button Up State this is the second

Select Case m_BorderStyleC
Case Is = 1
        Custom = False
            If m_Hover <> True Then
            ' hover option off
                BSO = BDR_SUNKENOUTER
                BRI = BDR_RAISEDINNER
            Else:
            ' hover option on
                    If UpT = False Then
                    ' button is  pushed
                        BSO = BDR_SUNKENOUTER
                    Else:
                    ' button not pushed
                        If mH = True Then
                            BRI = BDR_RAISEDINNER
                        Else:
                            BRI = &H7
                        End If
                    End If
            End If
Case Is = 2
        Custom = False
            If m_Hover <> True Then
                BSO = &HA
                BRI = &H5
            Else:
                    If UpT = False Then
                        BSO = &HA
                    Else:
                        If mH = True Then
                            BRI = &H5
                        Else:
                            BRI = &H7
                        End If
                    End If
            End If
Case Is = 3
        Custom = True
            If m_Hover <> True Then
                BSO = BDR_SUNKENOUTER
                BSO2 = &H8
                
                BRI = &H1
                BRI2 = BDR_RAISEDINNER
            Else:
                    If UpT = False Then
                        BSO = BDR_SUNKENOUTER
                        BSO2 = &H8
                    Else:
                        If mH = True Then
                            BRI = &H1
                            BRI2 = BDR_RAISEDINNER
                        Else:
                            BRI = &H7
                            BRI2 = &H7
                        End If
                    End If
            End If
Case Is = 4
        Custom = True
            If m_Hover <> True Then
                BSO = BDR_RAISEDINNER
                BSO2 = BDR_SUNKENOUTER
                
                BRI = BDR_SUNKENOUTER
                BRI2 = BDR_RAISEDINNER
            Else:
                    If UpT = False Then
                        BSO = BDR_RAISEDINNER
                        BSO2 = BDR_SUNKENOUTER
                    Else:
                        If mH = True Then
                            BRI = BDR_SUNKENOUTER
                            BRI2 = BDR_RAISEDINNER
                        Else:
                            BRI = &H7
                            BRI2 = &H7
                        End If
                    End If
            End If
Case Is = 5
        Custom2 = True
        Thick = 1
            If m_Hover <> True Then
            ' hover option off
                BSO = BDR_SUNKENOUTER
                BRI = BDR_RAISEDINNER
            Else:
            ' hover option on
                    If UpT = False Then
                    ' button is  pushed
                        BSO = BDR_SUNKENOUTER
                    Else:
                    ' button not pushed
                        If mH = True Then
                            BRI = BDR_RAISEDINNER
                        Else:
                            BRI = &H7
                        End If
                    End If
            End If
End Select

'//// Draw Border //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
If UpT = False Then
         If Custom <> True Then
            If Custom2 <> True Then
                        UserControl.Cls
                        With Rt
                            .Left = 0
                            .Top = 0
                            .Right = UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels)
                            .Bottom = UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels)
                        End With
                        UserControl.Cls
                        Q = DrawEdge(UserControl.hdc, Rt, BSO, BF_RECT)
            Else:
            '  Custom Thick
                UserControl.Cls
                For x = 0 To Thick
                With Rt
                    .Top = x
                    .Left = x
                    .Right = UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels)
                    .Bottom = UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels)
                End With
                DrawEdge UserControl.hdc, Rt, BSO, BF_TOPLEFT
                Next x
                For y = Val(UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels) - Thick) To UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels)
                With Rt
                    .Top = 0
                    .Left = 0
                    .Right = y
                    .Bottom = UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels)
                End With
                DrawEdge UserControl.hdc, Rt, BSO, BF_RIGHT
                Next y
                For Z = Val(UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels) - Thick) To UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels)
                With Rt
                    .Top = 0
                    .Top = 0
                    .Right = UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels)
                    .Bottom = Z
                End With
                DrawEdge UserControl.hdc, Rt, BSO, BF_BOTTOM
                Next Z
                ' End custom Thick
            End If
         Else:
         'Thin borders
         UserControl.Cls
         With Rt
             .Left = 0
             .Top = 0
             .Right = UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels)
             .Bottom = UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels)
         End With
            DrawEdge UserControl.hdc, Rt, BSO, BF_BOTTOMRIGHT
            DrawEdge UserControl.hdc, Rt, BSO2, BF_TOPLEFT
            ' End thin borders
         End If
Else:
            If Custom <> True Then
                If Custom2 <> True Then
                            UserControl.Cls
                            With Rt
                                .Left = 0
                                .Top = 0
                                .Right = UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels)
                                .Bottom = UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels)
                            End With
                            UserControl.Cls
                            Q = DrawEdge(UserControl.hdc, Rt, BRI, BF_RECT)
                            Debug.Print Q
                            Dim R As RECT
                Else:
            '  Custom Thick
                UserControl.Cls
                For x = 0 To Thick
                With Rt
                    .Top = x
                    .Left = x
                    .Right = UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels)
                    .Bottom = UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels)
                End With
                DrawEdge UserControl.hdc, Rt, BRI, BF_TOPLEFT
                Next x
                For y = Val(UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels) - Thick) To UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels)
                With Rt
                    .Top = 0
                    .Left = 0
                    .Right = y
                    .Bottom = UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels)
                End With
                DrawEdge UserControl.hdc, Rt, BRI, BF_RIGHT
                Next y
                For Z = Val(UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels) - Thick) To UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels)
                With Rt
                    .Top = 0
                    .Top = 0
                    .Right = UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels)
                    .Bottom = Z
                End With
                DrawEdge UserControl.hdc, Rt, BRI, BF_BOTTOM
                Next Z
                ' End custom Thick
                End If
            Else:
            ' Thin borders
            UserControl.Cls
            With Rt
                .Left = 0
                .Top = 0
                .Right = UserControl.ScaleX(UserControl.ScaleWidth, UserControl.ScaleMode, vbPixels)
                .Bottom = UserControl.ScaleY(UserControl.ScaleHeight, UserControl.ScaleMode, vbPixels)
            End With
                DrawEdge UserControl.hdc, Rt, BRI, BF_BOTTOMRIGHT
                DrawEdge UserControl.hdc, Rt, BRI2, BF_TOPLEFT
            'End Thin Bordrrs
            End If
End If


Select Case m_AlignText
Case Is = 3
    AlgnT = DT_WORDBREAK Or DT_CENTER Or DT_CENTERCENTER Or DT_NOCLIP
Case Is = 2
    AlgnT = DT_WORDBREAK Or DT_LEFT Or DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_VCENTER
Case Is = 1
    AlgnT = DT_WORDBREAK Or DT_RIGHT Or DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_VCENTER
End Select

'//// Draw Text & Draw Disabled Look ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
If Enabled = False Then
    mStr = m_Caption
    UserControl.ScaleMode = vbPixels
    SetRect R, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    Q = DrawTextEx(UserControl.hdc, mStr, Len(mStr), R, AlgnT, ByVal 0&)
    moPaintEffects.PaintDisabledDC UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, UserControl.hdc, 0, 0, , , , hPal 'clrMask
Else:
    If UpT = True Then
            mStr = m_Caption
            UserControl.ScaleMode = vbPixels
            SetRect R, 0, 0, Val(UserControl.ScaleWidth), UserControl.ScaleHeight
            Q = DrawTextEx(UserControl.hdc, mStr, Len(mStr), R, AlgnT, ByVal 0&)
    Else:
            mStr = m_Caption
            UserControl.ScaleMode = vbPixels
            SetRect R, 0, 0, Val(UserControl.ScaleWidth), UserControl.ScaleHeight
            R.Left = Val(R.Left) + 2
            R.Top = Val(R.Top) + 2
            Q = DrawTextEx(UserControl.hdc, mStr, Len(mStr), R, AlgnT, ByVal 0&)
        End If
End If
'//// Draw Focus Rect /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
If FocusT = True Then
    If m_FocusRect = True Then
        SetRect Rt, 4, 4, Val(UserControl.ScaleWidth) - 4, Val(UserControl.ScaleHeight) - 4
        DrawFocusRect UserControl.hdc, Rt
    Else:
        '
    End If
Else:
End If
End Sub


Private Sub UserControl_Resize()
    UserControl_Paint
End Sub

Private Sub UserControl_Show()
    RaiseEvent Show
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,TextHeight
Public Function TextHeight(ByVal Str As String) As Single
Attribute TextHeight.VB_Description = "Returns the height of a text string as it would be printed in the current font."
    TextHeight = UserControl.TextHeight(Str)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,TextWidth
Public Function TextWidth(ByVal Str As String) As Single
Attribute TextWidth.VB_Description = "Returns the width of a text string as it would be printed in the current font."
    TextWidth = UserControl.TextWidth(Str)
End Function

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_Caption = m_def_Caption
    m_FocusRect = m_def_FocusRect
    m_UpT = m_def_UpT
    m_Hover = m_def_Hover
    m_BorderStyleC = m_def_BorderStyleC
End Sub

'Load property values from storage

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    UserControl.CurrentX = PropBag.ReadProperty("CurrentX", 0)
    UserControl.CurrentY = PropBag.ReadProperty("CurrentY", 0)
    UserControl.FontBold = PropBag.ReadProperty("FontBold", 0)
    UserControl.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    UserControl.FontName = PropBag.ReadProperty("FontName", "")
    UserControl.FontSize = PropBag.ReadProperty("FontSize", 0)
    UserControl.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    UserControl.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 3600)
    UserControl.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
    UserControl.ScaleMode = PropBag.ReadProperty("ScaleMode", 1)
    UserControl.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 4800)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    UserControl.FontTransparent = PropBag.ReadProperty("FontTransparent", True)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_FocusRect = PropBag.ReadProperty("FocusRect", m_def_FocusRect)
    m_UpT = PropBag.ReadProperty("UpT", m_def_UpT)
    m_Hover = PropBag.ReadProperty("Hover", m_def_Hover)
    m_BorderStyleC = PropBag.ReadProperty("BorderStyleC", m_def_BorderStyleC)
    m_AlignText = PropBag.ReadProperty("AlignText", m_def_AlignText)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("AutoRedraw", UserControl.AutoRedraw, False)
    Call PropBag.WriteProperty("CurrentX", UserControl.CurrentX, 0)
    Call PropBag.WriteProperty("CurrentY", UserControl.CurrentY, 0)
    Call PropBag.WriteProperty("FontBold", UserControl.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", UserControl.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", UserControl.FontName, "")
    Call PropBag.WriteProperty("FontSize", UserControl.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", UserControl.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", UserControl.FontUnderline, 0)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 3600)
    Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
    Call PropBag.WriteProperty("ScaleMode", UserControl.ScaleMode, 1)
    Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 4800)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("FontTransparent", UserControl.FontTransparent, True)

    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("FocusRect", m_FocusRect, m_def_FocusRect)
    Call PropBag.WriteProperty("UpT", m_UpT, m_def_UpT)
    Call PropBag.WriteProperty("Hover", m_Hover, m_def_Hover)
    Call PropBag.WriteProperty("BorderStyleC", m_BorderStyleC, m_def_BorderStyleC)
    Call PropBag.WriteProperty("AlignText", m_AlignText, m_def_AlignText)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Sets or Gets the Buttons Caption"
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    RealPaint
    PropertyChanged "Caption"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,FontTransparent
Public Property Get FontTransparent() As Boolean
Attribute FontTransparent.VB_Description = "Returns/sets a value that determines whether background text/graphics on a Form, Printer or PictureBox are displayed."
    FontTransparent = UserControl.FontTransparent
End Property

Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
    UserControl.FontTransparent() = New_FontTransparent
    RealPaint
    PropertyChanged "FontTransparent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

Public Property Get FocusRect() As Boolean
Attribute FocusRect.VB_Description = "Sets the control to paint a focus Rectangle"
    FocusRect = m_FocusRect
End Property

Public Property Let FocusRect(ByVal New_FocusRect As Boolean)
    m_FocusRect = New_FocusRect
    PropertyChanged "FocusRect"
End Property

Public Property Get UpT() As Boolean
Attribute UpT.VB_Description = "Sets the Buttons Atributes (Up/Down)"
Attribute UpT.VB_MemberFlags = "400"
    UpT = m_UpT
End Property

Public Property Let UpT(ByVal New_UpT As Boolean)
    If Ambient.UserMode = False Then Err.Raise 382
    m_UpT = New_UpT
    PropertyChanged "UpT"
End Property


Public Property Get Hover() As Boolean
Attribute Hover.VB_Description = "Sets the button to be Flat untill the mouse is over it."
    Hover = m_Hover
End Property

Public Property Let Hover(ByVal New_Hover As Boolean)
    m_Hover = New_Hover
    PropertyChanged "Hover"
End Property

Public Property Get BorderStyleC() As Style
Attribute BorderStyleC.VB_Description = "Sets Border Style\r\n"
    BorderStyleC = m_BorderStyleC
End Property

Public Property Let BorderStyleC(ByVal New_BorderStyleC As Style)
    m_BorderStyleC = New_BorderStyleC
    RealPaint
    PropertyChanged "BorderStyleC"
End Property
Public Property Let AlignText(ByVal New_AlignText As Align)
m_AlignText = New_AlignText
RealPaint
PropertyChanged "AlignText"
End Property
Public Property Get AlignText() As Align
AlignText = m_AlignText
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Cls
Public Sub Cls()
Attribute Cls.VB_Description = "Clears graphics and text generated at run time from a Form, Image, or PictureBox."
    UserControl.Cls
End Sub

