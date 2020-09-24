VERSION 5.00
Begin VB.UserControl HzxYOption 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1335
   ClipControls    =   0   'False
   FillColor       =   &H8000000F&
   FillStyle       =   0  'Solid
   ForeColor       =   &H8000000F&
   ScaleHeight     =   22
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   89
   ToolboxBitmap   =   "HzxYOption.ctx":0000
End
Attribute VB_Name = "HzxYOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum OptionState
    FalseNormal = 0
    TrueNormal = 1
    FalseDisabled = 2
    TrueDisabled = 3
    FalseOver = 4
    TrueOver = 5
    FalseDown = 6
    TrueDown = 7
End Enum

Private m_Value As Boolean
Private m_Caption As String
Private m_BackColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_State As OptionState
Private optImage(7) As StdPicture
Private CorX_Pic As Long
Private CorY_Pic As Long
Private CorX_Cap As Long
Private CorY_Cap As Long
Private CaptionHeight As Long
Private lngFormat As Long
Private CaptionRect As RECT

Private Const m_def_State = FalseNormal

Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub UserControl_Initialize()
    Ini
    UserControl.ScaleMode = vbPixels
    UserControl.PaletteMode = vbPaletteModeContainer
End Sub

Private Sub UserControl_InitProperties()
    m_Caption = Ambient.DisplayName
    Enabled = True
    m_Value = False
    Set UserControl.Font = Ambient.Font
    m_BackColor = Parent.BackColor
    m_ForeColor = Parent.ForeColor
    UserControl.BackColor = m_BackColor
    UserControl.ForeColor = m_ForeColor
    m_State = m_def_State
    Dim loop1 As Integer
    For loop1 = LBound(optImage) To UBound(optImage)
        Set optImage(loop1) = Nothing
    Next
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Value = .ReadProperty("Value", False)
        m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        Enabled = .ReadProperty("Enabled", True)
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        m_BackColor = .ReadProperty("BackColor", Parent.BackColor)
        m_ForeColor = .ReadProperty("ForeColor", Parent.ForeColor)
        UserControl.BackColor = m_BackColor
        UserControl.ForeColor = m_ForeColor
        Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", 0)
        Set optImage(0) = .ReadProperty("Pic_FalseNormal", Nothing)
        Set optImage(1) = .ReadProperty("Pic_TrueNormal", Nothing)
        Set optImage(2) = .ReadProperty("Pic_FalseDisabled", Nothing)
        Set optImage(3) = .ReadProperty("Pic_TrueDisabled", Nothing)
        Set optImage(4) = .ReadProperty("Pic_FalseOver", Nothing)
        Set optImage(5) = .ReadProperty("Pic_TrueOver", Nothing)
        Set optImage(6) = .ReadProperty("Pic_FalseDown", Nothing)
        Set optImage(7) = .ReadProperty("Pic_TrueDown", Nothing)
    End With
    m_State = IIf(m_Value, TrueNormal, FalseNormal)
    If Enabled = False Then m_State = m_State + FalseDisabled
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
        Call .WriteProperty("Value", m_Value, False)
        Call .WriteProperty("Caption", m_Caption, Ambient.DisplayName)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("BackColor", m_BackColor, Parent.BackColor)
        Call .WriteProperty("ForeColor", m_ForeColor, Parent.ForeColor)
        Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, 0)
        Call .WriteProperty("Pic_FalseNormal", optImage(0), Nothing)
        Call .WriteProperty("Pic_TrueNormal", optImage(1), Nothing)
        Call .WriteProperty("Pic_FalseDisabled", optImage(2), Nothing)
        Call .WriteProperty("Pic_TrueDisabled", optImage(3), Nothing)
        Call .WriteProperty("Pic_FalseOver", optImage(4), Nothing)
        Call .WriteProperty("Pic_TrueOver", optImage(5), Nothing)
        Call .WriteProperty("Pic_FalseDown", optImage(6), Nothing)
        Call .WriteProperty("Pic_TrueDown", optImage(7), Nothing)
    End With
End Sub
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = m_BackColor
    DrawPicture m_State
    DrawCaption
End Property
Public Property Get Caption() As String
    Caption = m_Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    Refresh
End Property
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    If New_Enabled <> UserControl.Enabled Then
        UserControl.Enabled() = New_Enabled
        m_State = IIf(New_Enabled, m_State Mod 2, 2 + (m_State Mod 2))
        DrawCaption
        DrawPicture m_State
    End If
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    DrawCaption
End Property
Public Property Get Pic_TrueNormal() As StdPicture
    Set Pic_TrueNormal = optImage(1)
End Property
Public Property Set Pic_TrueNormal(ByVal newPic As StdPicture)
    Set optImage(1) = newPic
    PropertyChanged "Pic_TrueNormal"
    If m_State = TrueNormal Then DrawPicture m_State
End Property
Public Property Get Pic_TrueDisabled() As StdPicture
    Set Pic_TrueDisabled = optImage(3)
End Property
Public Property Set Pic_TrueDisabled(ByVal newPic As StdPicture)
    Set optImage(3) = newPic
    PropertyChanged "Pic_TrueDisabled"
    If m_State = TrueDisabled Then DrawPicture m_State
End Property
Public Property Get Pic_TrueDown() As StdPicture
    Set Pic_TrueDown = optImage(7)
End Property
Public Property Set Pic_TrueDown(ByVal newPic As StdPicture)
    Set optImage(7) = newPic
    PropertyChanged "Pic_TrueDown"
    If m_State = TrueDown Then DrawPicture m_State
End Property
Public Property Get Pic_TrueOver() As StdPicture
    Set Pic_TrueOver = optImage(5)
End Property
Public Property Set Pic_TrueOver(ByVal newPic As StdPicture)
    Set optImage(5) = newPic
    PropertyChanged "Pic_TrueOver"
    If m_State = TrueOver Then DrawPicture m_State
End Property
Public Property Get Pic_FalseNormal() As StdPicture
    Set Pic_FalseNormal = optImage(0)
End Property
Public Property Set Pic_FalseNormal(ByVal newPic As StdPicture)
    Set optImage(0) = newPic
    PropertyChanged "Pic_FalseNormal"
    If m_State = FalseNormal Then DrawPicture m_State
End Property
Public Property Get Pic_FalseDisabled() As StdPicture
    Set Pic_FalseDisabled = optImage(2)
End Property
Public Property Set Pic_FalseDisabled(ByVal newPic As StdPicture)
    Set optImage(2) = newPic
    PropertyChanged "Pic_FalseDisabled"
    If m_State = FalseDisabled Then DrawPicture m_State
End Property
Public Property Get Pic_FalseDown() As StdPicture
    Set Pic_FalseDown = optImage(6)
End Property
Public Property Set Pic_FalseDown(ByVal newPic As StdPicture)
    Set optImage(6) = newPic
    PropertyChanged "Pic_FalseDown"
    If m_State = FalseDown Then DrawPicture m_State
End Property
Public Property Get Pic_FalseOver() As StdPicture
    Set Pic_FalseOver = optImage(4)
End Property
Public Property Set Pic_FalseOver(ByVal newPic As StdPicture)
    Set optImage(4) = newPic
    PropertyChanged "Pic_FalseOver"
    If m_State = FalseOver Then DrawPicture m_State
End Property
Public Property Get Value() As Boolean
    Value = m_Value
End Property
Public Property Let Value(ByVal vNewValue As Boolean)
    If m_Value <> vNewValue Then
        m_Value = vNewValue
        PropertyChanged "Value"
        If m_Value Then
            m_State = 2 * Int(m_State / 2) + 1
        Else
            m_State = 2 * Int(m_State / 2)
        End If
        DrawPicture m_State
        If m_Value Then ContainerCheck
    End If
End Property
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    UserControl.Cls
    Refresh
    PropertyChanged "Font"
End Property
Public Property Get FontUnderline() As Boolean
    FontUnderline = UserControl.FontUnderline
End Property
Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    UserControl.FontUnderline() = New_FontUnderline
    UserControl.Cls
    Refresh
    PropertyChanged "Font"
End Property
Public Property Get FontStrikethru() As Boolean
    FontStrikethru = UserControl.FontStrikethru
End Property
Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    UserControl.FontStrikethru() = New_FontStrikethru
    UserControl.Cls
    Refresh
    PropertyChanged "Font"
End Property
Public Property Get FontSize() As Single
    FontSize = UserControl.FontSize
End Property
Public Property Let FontSize(ByVal New_FontSize As Single)
    UserControl.FontSize() = New_FontSize
    UserControl.Cls
    Refresh
    PropertyChanged "Font"
End Property
Public Property Get FontName() As String
    FontName = UserControl.FontName
End Property
Public Property Let FontName(ByVal New_FontName As String)
    UserControl.FontName() = New_FontName
    UserControl.Cls
    Refresh
    PropertyChanged "Font"
End Property
Public Property Get FontItalic() As Boolean
    FontItalic = UserControl.FontItalic
End Property
Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    UserControl.FontItalic() = New_FontItalic
    UserControl.Cls
    Refresh
    PropertyChanged "Font"
End Property
Public Property Get FontBold() As Boolean
    FontBold = UserControl.FontBold
End Property
Public Property Let FontBold(ByVal New_FontBold As Boolean)
    UserControl.FontBold() = New_FontBold
    UserControl.Cls
    Refresh
    PropertyChanged "Font"
End Property
Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As StdPicture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_Click()
    If Not Value Then Value = True
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    m_State = IIf(Value, TrueDown, FalseDown)
    DrawPicture m_State
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enabled = True Then
        If Value Then
            m_State = TrueDown
        Else
            m_State = FalseDown
        End If
    End If
    DrawPicture m_State
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
    
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
    SetCapture hwnd
    If PointInControl(X, UserControl.ScaleWidth, Y, UserControl.ScaleHeight) Then
        If m_State < FalseOver Then
            If Button = vbLeftButton Then
                If Value Then
                    m_State = TrueDown
                Else
                    m_State = FalseDown
                End If
            Else
                If Value Then
                    m_State = TrueOver
                Else
                    m_State = FalseOver
                End If
            End If
            DrawPicture m_State
        End If
    Else
        If Value Then
            m_State = TrueNormal
        Else
            m_State = FalseNormal
        End If
        DrawPicture m_State
        RaiseEvent MouseOut
        ReleaseCapture
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        m_State = m_State - 2
        DrawPicture m_State
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    Refresh
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

Private Sub UserControl_Terminate()
    Dim loop1 As Integer
    For loop1 = LBound(optImage) To UBound(optImage)
        Set optImage(loop1) = Nothing
    Next
End Sub

Public Sub Refresh()
    UserControl.Cls
    CalPosition
    DrawCaption
    DrawPicture m_State
End Sub

Private Sub CalPosition()
        
    Dim TmpRect As RECT
    Dim TextSize As Size
    
    UserControl.ScaleMode = vbPixels
    CorX_Pic = 0

    With UserControl
        GetTextExtentPoint32 .hdc, m_Caption, LenB(StrConv(m_Caption, vbFromUnicode)), TextSize
        .Width = (TextSize.cx + 17) * 15
        Call SetRect(TmpRect, 17, 0, .ScaleWidth, .ScaleHeight)
    End With
    lngFormat = DT_WORDBREAK Or DT_LEFT
    CaptionHeight = DrawText(UserControl.hdc, m_Caption, -1, TmpRect, lngFormat Or DT_CALCRECT)
    If CaptionHeight > 1 Then
        With UserControl
            .Height = IIf(CaptionHeight >= 13, CaptionHeight * 15, 195)
            Call SetRect(CaptionRect, 17, Int((.ScaleHeight - CaptionHeight) / 2), .ScaleWidth, Int((.ScaleHeight + CaptionHeight) / 2))
            CorY_Pic = .ScaleHeight \ 2 - 6
        End With
    End If
End Sub

Private Sub DrawCaption()
    Dim TmpRGBColor1 As Long
    
    If UserControl.Enabled Then
        If Trim(m_Caption) > 0 Then
            TmpRGBColor1 = BreakApart(m_ForeColor)
            SetTextColor UserControl.hdc, TmpRGBColor1
            DrawText UserControl.hdc, m_Caption, -1, CaptionRect, lngFormat
        End If
    Else
        If Trim(m_Caption) > 0 Then
            TmpRGBColor1 = BreakApart(&H80000011)
            SetTextColor UserControl.hdc, TmpRGBColor1
            DrawText UserControl.hdc, m_Caption, -1, CaptionRect, lngFormat
        End If
    End If
End Sub

Private Sub DrawPicture(CurState As OptionState)
    
    Dim str As String
    Dim tempPic As StdPicture
    Dim OriW As Long, OriH As Long
    
    Select Case CurState
    Case FalseNormal
        str = "optFalseNormal"
    Case TrueNormal
        str = "optTrueNormal"
    Case FalseDisabled
        str = "optFalseDisabled"
    Case TrueDisabled
        str = "optTrueDisabled"
    Case FalseOver
        str = "optFalseOver"
    Case TrueOver
        str = "optTrueOver"
    Case FalseDown
        str = "optFalseDown"
    Case TrueDown
        str = "optTrueDown"
    End Select

    Select Case CurState
    Case FalseNormal, TrueNormal
        If Not optImage(CurState) Is Nothing Then
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, BreakApart(m_BackColor)
            GetOriWH optImage(CurState), OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, optImage(CurState), OriW, OriH
        Else
            Set tempPic = LoadResPicture(str, vbResIcon)
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, BreakApart(m_BackColor)
            GetOriWH tempPic, OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, tempPic, OriW, OriH
            Set tempPic = Nothing
        End If
    Case FalseDisabled, TrueDisabled
        If Not optImage(CurState) Is Nothing Then
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, BreakApart(m_BackColor)
            GetOriWH optImage(CurState), OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, optImage(CurState), OriW, OriH
        ElseIf Not optImage(CurState Mod 2) Is Nothing Then
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, BreakApart(m_BackColor)
            GetOriWH optImage(CurState Mod 2), OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, optImage(CurState Mod 2), OriW, OriH, Icon_Grey
        Else
            Set tempPic = LoadResPicture(str, vbResIcon)
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, BreakApart(m_BackColor)
            GetOriWH tempPic, OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, tempPic, OriW, OriH
            Set tempPic = Nothing
        End If
    Case FalseOver, TrueOver, FalseDown, TrueDown
        If Not optImage(CurState) Is Nothing Then
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, BreakApart(m_BackColor)
            GetOriWH optImage(CurState), OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, optImage(CurState), OriW, OriH
        ElseIf Not optImage(CurState Mod 2) Is Nothing Then
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, BreakApart(m_BackColor)
            GetOriWH optImage(CurState Mod 2), OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, optImage(CurState Mod 2), OriW, OriH
        Else
            Set tempPic = LoadResPicture(str, vbResIcon)
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, BreakApart(m_BackColor)
            GetOriWH tempPic, OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 13, 13, tempPic, OriW, OriH
            Set tempPic = Nothing
        End If
    End Select
End Sub

Private Sub GetOriWH(ByVal srcPic As StdPicture, OriW As Long, OriH As Long)
    
    OriW = UserControl.ScaleX(srcPic.Width, vbHimetric, vbPixels)
    OriH = UserControl.ScaleY(srcPic.Height, vbHimetric, vbPixels)

End Sub

Private Sub ContainerCheck()
    Dim Control As Object
    For Each Control In UserControl.Parent.Controls
        If TypeOf Control Is HzxYOption Then
            If Control.Container.hwnd = UserControl.ContainerHwnd Then
                If Control.hdc <> UserControl.hdc Then
                    If Control.Value = True Then Control.Value = False
                End If
            End If
        End If
    Next
End Sub
