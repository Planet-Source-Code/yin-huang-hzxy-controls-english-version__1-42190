VERSION 5.00
Begin VB.UserControl HzxYCheckBox 
   BackColor       =   &H00D8E9EC&
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1335
   FillStyle       =   0  'Solid
   ScaleHeight     =   330
   ScaleWidth      =   1335
   ToolboxBitmap   =   "HzxYCheckBox.ctx":0000
End
Attribute VB_Name = "HzxYCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum CheckStates
    UncheckedNormal = 0
    CheckedNormal = 1
    MixedNormal = 2
    UncheckedDisabled = 3
    CheckedDisabled = 4
    MixedDisabled = 5
    UncheckedOver = 6
    CheckedOver = 7
    MixedOver = 8
    UncheckedDown = 9
    CheckedDown = 10
    MixedDown = 11
End Enum

Public Enum CheckValues
    Unchecked = 0
    Checked = 1
    Mixed = 2
End Enum

Private m_Value As CheckValues
Private m_Caption As String
Private m_Font As Font
Private m_BackColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_State As CheckStates
Private chkImage(11) As StdPicture
Private CorX_Pic As Long
Private CorY_Pic As Long
Private CorX_Cap As Long
Private CorY_Cap As Long
Private CaptionHeight As Long
Private lngFormat As Long
Private CaptionRect As RECT

Private Const m_def_State = MixedNormal

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
    m_Value = Mixed
    Set UserControl.Font = Ambient.Font
    m_BackColor = Parent.BackColor
    m_ForeColor = Parent.ForeColor
    UserControl.BackColor = m_BackColor
    UserControl.ForeColor = m_ForeColor
    m_State = m_def_State
    Set chkImage(0) = LoadResPicture("chkUncheckedNormal", vbResBitmap)
    Set chkImage(1) = LoadResPicture("chkCheckedNormal", vbResBitmap)
    Set chkImage(2) = LoadResPicture("chkMixedNormal", vbResBitmap)
    Set chkImage(3) = LoadResPicture("chkUncheckedDisabled", vbResBitmap)
    Set chkImage(4) = LoadResPicture("chkCheckedDisabled", vbResBitmap)
    Set chkImage(5) = LoadResPicture("chkMixedDisabled", vbResBitmap)
    Set chkImage(6) = LoadResPicture("chkUncheckedOver", vbResBitmap)
    Set chkImage(7) = LoadResPicture("chkCheckedOver", vbResBitmap)
    Set chkImage(8) = LoadResPicture("chkMixedOver", vbResBitmap)
    Set chkImage(9) = LoadResPicture("chkUncheckedDown", vbResBitmap)
    Set chkImage(10) = LoadResPicture("chkCheckedDown", vbResBitmap)
    Set chkImage(11) = LoadResPicture("chkMixedDown", vbResBitmap)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Value = .ReadProperty("Value", Mixed)
        m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        Enabled = .ReadProperty("Enabled", True)
        Set UserControl.Font = .ReadProperty("Font", Ambient.Font)
        m_BackColor = .ReadProperty("BackColor", Parent.BackColor)
        m_ForeColor = .ReadProperty("ForeColor", Parent.ForeColor)
        UserControl.BackColor = m_BackColor
        UserControl.ForeColor = m_ForeColor
        Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", 0)
        Set chkImage(0) = .ReadProperty("Pic_UncheckedNormal", Nothing)
        Set chkImage(1) = .ReadProperty("Pic_CheckedNormal", Nothing)
        Set chkImage(2) = .ReadProperty("Pic_MixedNormal", Nothing)
        Set chkImage(3) = .ReadProperty("Pic_UncheckedDisabled", Nothing)
        Set chkImage(4) = .ReadProperty("Pic_CheckedDisabled", Nothing)
        Set chkImage(5) = .ReadProperty("Pic_MixedDisabled", Nothing)
        Set chkImage(6) = .ReadProperty("Pic_UncheckedOver", Nothing)
        Set chkImage(7) = .ReadProperty("Pic_CheckedOver", Nothing)
        Set chkImage(8) = .ReadProperty("Pic_MixedOver", Nothing)
        Set chkImage(9) = .ReadProperty("Pic_UncheckedDown", Nothing)
        Set chkImage(10) = .ReadProperty("Pic_CheckedDown", Nothing)
        Set chkImage(11) = .ReadProperty("Pic_MixedDown", Nothing)
    End With
    m_State = IIf(Enabled, m_Value, m_Value + 3)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
        Call .WriteProperty("Value", m_Value, Mixed)
        Call .WriteProperty("Caption", m_Caption, Ambient.DisplayName)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("BackColor", m_BackColor, Parent.BackColor)
        Call .WriteProperty("ForeColor", m_ForeColor, Parent.ForeColor)
        Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, 0)
        Call .WriteProperty("Pic_UncheckedNormal", chkImage(0), Nothing)
        Call .WriteProperty("Pic_CheckedNormal", chkImage(1), Nothing)
        Call .WriteProperty("Pic_MixedNormal", chkImage(2), Nothing)
        Call .WriteProperty("Pic_UncheckedDisabled", chkImage(3), Nothing)
        Call .WriteProperty("Pic_CheckedDisabled", chkImage(4), Nothing)
        Call .WriteProperty("Pic_MixedDisabled", chkImage(5), Nothing)
        Call .WriteProperty("Pic_UncheckedOver", chkImage(6), Nothing)
        Call .WriteProperty("Pic_CheckedOver", chkImage(7), Nothing)
        Call .WriteProperty("Pic_MixedOver", chkImage(8), Nothing)
        Call .WriteProperty("Pic_UncheckedDown", chkImage(9), Nothing)
        Call .WriteProperty("Pic_CheckedDown", chkImage(10), Nothing)
        Call .WriteProperty("Pic_MixedDown", chkImage(11), Nothing)
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
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
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
        m_State = IIf(New_Enabled, m_State Mod 3, 3 + (m_State Mod 3))
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
Public Property Get Pic_CheckedNormal() As StdPicture
    Set Pic_CheckedNormal = chkImage(1)
End Property
Public Property Set Pic_CheckedNormal(ByVal newPic As StdPicture)
    Set chkImage(1) = newPic
    PropertyChanged "Pic_CheckedNormal"
    If m_State = CheckedNormal Then DrawPicture m_State
End Property
Public Property Get Pic_CheckedDisabled() As StdPicture
    Set Pic_CheckedDisabled = chkImage(4)
End Property
Public Property Set Pic_CheckedDisabled(ByVal newPic As StdPicture)
    Set chkImage(4) = newPic
    PropertyChanged "Pic_CheckedDisabled"
    If m_State = CheckedDisabled Then DrawPicture m_State
End Property
Public Property Get Pic_CheckedDown() As StdPicture
    Set Pic_CheckedDown = chkImage(10)
End Property
Public Property Set Pic_CheckedDown(ByVal newPic As StdPicture)
    Set chkImage(10) = newPic
    PropertyChanged "Pic_CheckedDown"
    If m_State = CheckedDown Then DrawPicture m_State
End Property
Public Property Get Pic_CheckedOver() As StdPicture
    Set Pic_CheckedOver = chkImage(7)
End Property
Public Property Set Pic_CheckedOver(ByVal newPic As StdPicture)
    Set chkImage(7) = newPic
    PropertyChanged "Pic_CheckedOver"
    If m_State = CheckedOver Then DrawPicture m_State
End Property
Public Property Get Pic_MixedNormal() As StdPicture
    Set Pic_MixedNormal = chkImage(2)
End Property
Public Property Set Pic_MixedNormal(ByVal newPic As StdPicture)
    Set chkImage(2) = newPic
    PropertyChanged "Pic_MixedNormal"
    If m_State = MixedNormal Then DrawPicture m_State
End Property
Public Property Get Pic_MixedDisabled() As StdPicture
    Set Pic_MixedDisabled = chkImage(5)
End Property
Public Property Set Pic_MixedDisabled(ByVal newPic As StdPicture)
    Set chkImage(5) = newPic
    PropertyChanged "Pic_MixedDisabled"
    If m_State = MixedDisabled Then DrawPicture m_State
End Property
Public Property Get Pic_MixedDown() As StdPicture
    Set Pic_MixedDown = chkImage(11)
End Property
Public Property Set Pic_MixedDown(ByVal newPic As StdPicture)
    Set chkImage(11) = newPic
    PropertyChanged "Pic_MixedDown"
    If m_State = MixedDown Then DrawPicture m_State
End Property
Public Property Get Pic_MixedOver() As StdPicture
    Set Pic_MixedOver = chkImage(8)
End Property
Public Property Set Pic_MixedOver(ByVal newPic As StdPicture)
    Set chkImage(8) = newPic
    PropertyChanged "Pic_MixedOver"
    If m_State = MixedOver Then DrawPicture m_State
End Property
Public Property Get Pic_UncheckedNormal() As StdPicture
    Set Pic_UncheckedNormal = chkImage(0)
End Property
Public Property Set Pic_UncheckedNormal(ByVal newPic As StdPicture)
    Set chkImage(0) = newPic
    PropertyChanged "Pic_UncheckedNormal"
    If m_State = UncheckedNormal Then DrawPicture m_State
End Property
Public Property Get Pic_UncheckedDisabled() As StdPicture
    Set Pic_UncheckedDisabled = chkImage(3)
End Property
Public Property Set Pic_UncheckedDisabled(ByVal newPic As StdPicture)
    Set chkImage(3) = newPic
    PropertyChanged "Pic_UncheckedDisabled"
    If m_State = UncheckedDisabled Then DrawPicture m_State
End Property
Public Property Get Pic_UncheckedDown() As StdPicture
    Set Pic_UncheckedDown = chkImage(9)
End Property
Public Property Set Pic_UncheckedDown(ByVal newPic As StdPicture)
    Set chkImage(9) = newPic
    PropertyChanged "Pic_UncheckedDown"
    If m_State = UncheckedDown Then DrawPicture m_State
End Property
Public Property Get Pic_UncheckedOver() As StdPicture
    Set Pic_UncheckedOver = chkImage(6)
End Property
Public Property Set Pic_UncheckedOver(ByVal newPic As StdPicture)
    Set chkImage(6) = newPic
    PropertyChanged "Pic_UncheckedOver"
    If m_State = UncheckedOver Then DrawPicture m_State
End Property
Public Property Get Value() As CheckValues
    Value = m_Value
End Property
Public Property Let Value(ByVal vNewValue As CheckValues)
    If m_Value <> vNewValue Then
        m_Value = vNewValue
        PropertyChanged "Value"
        If m_Value = Unchecked Then
            m_State = 3 * Int(m_State / 3)
        ElseIf m_Value = Checked Then
            m_State = 3 * Int(m_State / 3) + 1
        Else
            m_State = 3 * Int(m_State / 3) + 2
        End If
        DrawPicture m_State
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
    m_Value = IIf(m_Value = Checked, Unchecked, Checked)
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    m_State = IIf(m_Value = Checked, CheckedDown, UncheckedDown)
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
        If m_Value = Checked Then
            m_State = CheckedDown
        ElseIf m_Value = Unchecked Then
            m_State = UncheckedDown
        Else
            m_State = MixedDown
        End If
    End If
    DrawPicture m_State
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
    
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
    SetCapture hwnd
    If PointInControl(X, UserControl.ScaleWidth, Y, UserControl.ScaleHeight) Then
        If m_State < UncheckedOver Then
            If Button = vbLeftButton Then
                If m_Value = Checked Then
                    m_State = CheckedDown
                ElseIf m_Value = Unchecked Then
                    m_State = UncheckedDown
                Else
                    m_State = MixedDown
                End If
            Else
                If m_Value = Checked Then
                    m_State = CheckedOver
                ElseIf m_Value = Unchecked Then
                    m_State = UncheckedOver
                Else
                    m_State = MixedOver
                End If
            End If
            DrawPicture m_State
        End If
    Else
        If m_Value = Checked Then
            m_State = CheckedNormal
        ElseIf m_Value = Unchecked Then
            m_State = UncheckedNormal
        Else
            m_State = MixedNormal
        End If
        DrawPicture m_State
        RaiseEvent MouseOut
        ReleaseCapture
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        m_State = IIf(m_Value = Checked, UncheckedOver, CheckedOver)
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
    For loop1 = LBound(chkImage) To UBound(chkImage)
        Set chkImage(loop1) = Nothing
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
        .Width = (TextSize.cx + 21) * 15
        Call SetRect(TmpRect, 21, 0, .ScaleWidth, .ScaleHeight)
    End With
    lngFormat = DT_WORDBREAK Or DT_LEFT
    CaptionHeight = DrawText(UserControl.hdc, m_Caption, -1, TmpRect, lngFormat Or DT_CALCRECT)
    If CaptionHeight > 1 Then
        With UserControl
            .Height = IIf(CaptionHeight >= 16, CaptionHeight * 15, 240)
            Call SetRect(CaptionRect, 21, Int((.ScaleHeight - CaptionHeight) / 2), .ScaleWidth, Int((.ScaleHeight + CaptionHeight) / 2))
            CorY_Pic = .ScaleHeight \ 2 - 8
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

Private Sub DrawPicture(CurState As CheckStates)
    
    Dim OriW As Long, OriH As Long
    
    DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, 16, 16, BreakApart(m_BackColor)
    Select Case CurState
    Case UncheckedNormal, CheckedNormal, MixedNormal
        If Not chkImage(CurState) Is Nothing Then
            GetOriWH chkImage(CurState), OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 16, 16, chkImage(CurState), OriW, OriH
        End If
    Case UncheckedDisabled, CheckedDisabled, MixedDisabled
        If Not chkImage(CurState) Is Nothing Then
            GetOriWH chkImage(CurState), OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 16, 16, chkImage(CurState), OriW, OriH
        ElseIf Not chkImage(CurState Mod 3) Is Nothing Then
            GetOriWH chkImage(CurState Mod 3), OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 16, 16, chkImage(CurState Mod 3), OriW, OriH, Icon_Grey
        End If
    Case UncheckedOver, CheckedOver, MixedOver, UncheckedDown, CheckedDown, MixedDown
        If Not chkImage(CurState) Is Nothing Then
            GetOriWH chkImage(CurState), OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 16, 16, chkImage(CurState), OriW, OriH
        ElseIf Not chkImage(CurState Mod 3) Is Nothing Then
            GetOriWH chkImage(CurState Mod 3), OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 16, 16, chkImage(CurState Mod 3), OriW, OriH
        End If
    End Select
End Sub

Private Sub GetOriWH(ByVal srcPic As StdPicture, OriW As Long, OriH As Long)
    
    OriW = UserControl.ScaleX(srcPic.Width, vbHimetric, vbPixels)
    OriH = UserControl.ScaleY(srcPic.Height, vbHimetric, vbPixels)

End Sub
