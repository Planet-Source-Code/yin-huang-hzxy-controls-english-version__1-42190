VERSION 5.00
Begin VB.UserControl HzxYTabLabel 
   Appearance      =   0  'Flat
   BackColor       =   &H00D8E9EC&
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   FillStyle       =   2  'Horizontal Line
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   100
   ToolboxBitmap   =   "HzxYTabLabel.ctx":0000
End
Attribute VB_Name = "HzxYTabLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Caption As String
Private m_Font As Font
Private m_ForeColor As OLE_COLOR
Private CorX_Pic As Long
Private CorY_Pic As Long
Private CorX_Cap As Long
Private CorY_Cap As Long
Private CaptionHeight As Long
Private lngFormat As Long
Private CaptionRect As RECT
Private m_Image As StdPicture
Private m_IsActive As Boolean
Private CurState As LabelState

Event Click()
Event MouseOut()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub UserControl_Initialize()
    Ini
    UserControl.ScaleMode = vbPixels
    UserControl.PaletteMode = vbPaletteModeContainer
End Sub

Private Sub UserControl_InitProperties()
    UserControl.ScaleMode = vbPixels
    m_IsActive = False
    m_Caption = "[No Caption]"
    Enabled = Parent.Enabled
    Set UserControl.Font = Parent.Font
    m_ForeColor = Parent.ForeColor
    UserControl.ForeColor = m_ForeColor
    Set m_Image = Nothing
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_IsActive = .ReadProperty("IsActive", False)
        m_Caption = .ReadProperty("Caption", "[No Caption]")
        Enabled = .ReadProperty("Enabled", Parent.Enabled)
        Set UserControl.Font = .ReadProperty("Font", Parent.Font)
        m_ForeColor = .ReadProperty("ForeColor", Parent.ForeColor)
        UserControl.ForeColor = m_ForeColor
        Set m_Image = .ReadProperty("Image", Nothing)
    End With
    CurState = IIf(m_IsActive, lblActive, lblNormal)
End Sub

Private Sub UserControl_Terminate()
    Set m_Image = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("IsActive", m_IsActive, False)
        Call .WriteProperty("Caption", m_Caption, "[No Caption]")
        Call .WriteProperty("Enabled", UserControl.Enabled, Parent.Enabled)
        Call .WriteProperty("Font", UserControl.Font, Parent.Font)
        Call .WriteProperty("ForeColor", m_ForeColor, Parent.ForeColor)
        Call .WriteProperty("Image", m_Image, Nothing)
    End With
End Sub
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
        PropertyChanged "Enabled"
        If UserControl.Enabled Then
            CurState = IIf(m_IsActive, lblActive, lblNormal)
        Else
            CurState = lblDisabeld
        End If
        Refresh
    End If
End Property
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Refresh
End Property
Public Property Get FontUnderline() As Boolean
    FontUnderline = UserControl.FontUnderline
End Property
Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    UserControl.FontUnderline() = New_FontUnderline
    PropertyChanged "Font"
    Refresh
End Property
Public Property Get FontStrikethru() As Boolean
    FontStrikethru = UserControl.FontStrikethru
End Property
Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    UserControl.FontStrikethru() = New_FontStrikethru
    PropertyChanged "Font"
    Refresh
End Property
Public Property Get FontSize() As Single
    FontSize = UserControl.FontSize
End Property
Public Property Let FontSize(ByVal New_FontSize As Single)
    UserControl.FontSize() = New_FontSize
    PropertyChanged "Font"
    Refresh
End Property
Public Property Get FontName() As String
    FontName = UserControl.FontName
End Property
Public Property Let FontName(ByVal New_FontName As String)
    UserControl.FontName() = New_FontName
    PropertyChanged "Font"
    Refresh
End Property
Public Property Get FontItalic() As Boolean
    FontItalic = UserControl.FontItalic
End Property
Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    UserControl.FontItalic() = New_FontItalic
    PropertyChanged "Font"
    Refresh
End Property
Public Property Get FontBold() As Boolean
    FontBold = UserControl.FontBold
End Property
Public Property Let FontBold(ByVal New_FontBold As Boolean)
    UserControl.FontBold() = New_FontBold
    PropertyChanged "Font"
    Refresh
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    Refresh
End Property
Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
Public Property Get Image() As StdPicture
    Set Image = m_Image
End Property
Public Property Set Image(NewImage As StdPicture)
    Set m_Image = NewImage
    PropertyChanged "Image"
    Refresh
End Property
Public Property Get IsActive() As Boolean
    IsActive = m_IsActive
End Property
Public Property Let IsActive(NewIsActive As Boolean)
    If m_IsActive <> NewIsActive Then
        m_IsActive = NewIsActive
        PropertyChanged "IsActive"
        If m_IsActive Then
            CurState = lblActive
            ContainerCheck
        ElseIf UserControl.Enabled Then
            CurState = lblNormal
        Else
            CurState = lblDisabeld
        End If
        Refresh
    End If
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
    If Not m_IsActive Then
        m_IsActive = True
        CurState = lblActive
        Refresh
        ContainerCheck
    End If
    RaiseEvent Click
End Sub
   
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    SetCapture hwnd
    If PointInControl(X, UserControl.ScaleWidth, Y, UserControl.ScaleHeight) Then
        If Not m_IsActive And CurState <> lblOver Then
            CurState = lblOver
            Refresh
        End If
    Else
        If Not m_IsActive And CurState <> lblNormal Then
            CurState = lblNormal
            Refresh
            RaiseEvent MouseOut
        End If
        ReleaseCapture
    End If
End Sub

Public Sub UserControl_Paint()
    Refresh
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

Public Sub Refresh()
    CalPosition
    DrawBackColor
    DrawCaption
    DrawPicture
    DrawBorder
    RoundCorners
End Sub

Public Sub CalPosition()
        
    Dim TmpRect As RECT
    Dim TextSize As Size
    Dim BaseWidth As Long
    
    UserControl.ScaleMode = vbPixels
    
    UserControl.Height = 510
    
    CorX_Pic = 5
    CorY_Pic = 6
    
    BaseWidth = IIf(m_Image Is Nothing, 5, 33)

    With UserControl
        GetTextExtentPoint32 .hdc, m_Caption, LenB(StrConv(m_Caption, vbFromUnicode)), TextSize
        If m_IsActive Then
            If .Width <> (TextSize.cx + BaseWidth + 6) * 15 Then
                .Width = (TextSize.cx + BaseWidth + 6) * 15
                .ScaleWidth = .Width \ 15
            End If
        Else
            If .Width <> (TextSize.cx + BaseWidth + 5) * 15 Then
                .Width = (TextSize.cx + BaseWidth + 5) * 15
                .ScaleWidth = .Width \ 15
            End If
        End If
        Call SetRect(TmpRect, BaseWidth, 6, .ScaleWidth - 5, .ScaleHeight - 4)
    End With
    lngFormat = DT_WORDBREAK Or DT_CENTER
    CaptionHeight = DrawText(UserControl.hdc, m_Caption, -1, TmpRect, lngFormat Or DT_CALCRECT)
    If CaptionHeight > 1 Then
        With UserControl
            Call SetRect(CaptionRect, BaseWidth, 18 - CaptionHeight \ 2, .ScaleWidth - 5, 18 - CaptionHeight \ 2 + CaptionHeight)
        End With
    End If
End Sub

Private Sub DrawBackColor()
    With UserControl
        If .Enabled Then
            If m_IsActive Then
                DrawRectangle .hdc, 0, 0, .ScaleWidth, .ScaleHeight, &HFEFCFC
            Else
                DrawRectangle .hdc, 0, 0, .ScaleWidth, .ScaleHeight, &HF0F0F0
            End If
        Else
            DrawRectangle .hdc, 0, 0, .ScaleWidth, .ScaleHeight, ShiftColor(&HFFFFFF, &H18, True)
        End If
    End With
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

Private Sub DrawPicture()
    
    Dim OriW As Long, OriH As Long
    
    If Not m_Image Is Nothing Then
        If UserControl.Enabled Then
            GetOriWH m_Image, OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 24, 24, m_Image, OriW, OriH
        Else
            GetOriWH m_Image, OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, 24, 24, m_Image, OriW, OriH, Icon_Grey
        End If
    End If
End Sub

Private Sub GetOriWH(ByVal srcPic As StdPicture, OriW As Long, OriH As Long)
    
    OriW = UserControl.ScaleX(srcPic.Width, vbHimetric, vbPixels)
    OriH = UserControl.ScaleY(srcPic.Height, vbHimetric, vbPixels)

End Sub

Private Sub DrawBorder()
    
    Dim oldPen As Long, hPen As Long, Color As Long
    Dim Coor_Left As Long, Coor_Right As Long
    Dim Coor_Top As Long, Coor_Bottom As Long
    
    If m_IsActive Then
        Coor_Left = 0
        Coor_Right = UserControl.ScaleWidth
        Coor_Top = 0
    Else
        Coor_Left = 1
        Coor_Right = UserControl.ScaleWidth
        Coor_Top = 2
    End If
    Coor_Bottom = UserControl.ScaleHeight
        
    Select Case CurState
    Case lblActive, lblOver
        DrawLine UserControl.hdc, Coor_Left, Coor_Top, Coor_Right, Coor_Top, &H2C8BE6
        DrawLine UserControl.hdc, Coor_Left, Coor_Top + 1, Coor_Right, Coor_Top + 1, &H3CC8FF
        DrawLine UserControl.hdc, Coor_Left, Coor_Top + 2, Coor_Right, Coor_Top + 2, &H3CC8FF
        
        If CurState = lblActive Then
            Dim TmpRect As RECT
            With UserControl
                Call SetRect(TmpRect, 4, 4, .ScaleWidth - 3, .ScaleHeight - 2)
                Call DrawFocusRect(.hdc, TmpRect)
            End With
        End If
    Case lblNormal
        Color = &HF0F0F0
        DrawLine UserControl.hdc, Coor_Left, Coor_Top, Coor_Right, Coor_Top, Color
        DrawLine UserControl.hdc, Coor_Left, Coor_Top + 1, Coor_Right, Coor_Top + 1, Color
        DrawLine UserControl.hdc, Coor_Left, Coor_Top + 2, Coor_Right, Coor_Top + 2, Color
    End Select
    
    If UserControl.Enabled Then
        Color = &HA09C98
    Else
        Color = ShiftColor(&HFFFFFF, -&H3C, True)
    End If
    With UserControl
        hPen = CreatePen(PS_SOLID, 1, Color)
        oldPen = SelectObject(.hdc, hPen)
        Arc .hdc, Coor_Left, Coor_Top, Coor_Left + 8, Coor_Top + 8, 4, 0, 0, 4
        Arc .hdc, Coor_Right - 8, Coor_Top, Coor_Right, Coor_Top + 8, Coor_Right, 4, Coor_Right - 4, 0
        SelectObject .hdc, oldPen
        DeleteObject hPen
    End With
    DrawLine UserControl.hdc, Coor_Left, Coor_Top, Coor_Left, Coor_Bottom, Color
    DrawLine UserControl.hdc, Coor_Left, Coor_Top, Coor_Right, Coor_Top, Color
    DrawLine UserControl.hdc, Coor_Right - 1, Coor_Top, Coor_Right - 1, Coor_Bottom, Color
    If m_IsActive Then
        DrawLine UserControl.hdc, Coor_Left + 1, Coor_Bottom + 1, Coor_Right - 1, Coor_Bottom + 1, &HFEFCFC
    Else
        DrawLine UserControl.hdc, Coor_Left, Coor_Bottom + 1, Coor_Right, Coor_Bottom + 1, Color
    End If
End Sub

Private Sub RoundCorners()
    Dim TempRect As Long, TempRect1 As Long, TempRect2 As Long
    Dim He As Long, Wi As Long
    
    He = UserControl.ScaleHeight
    Wi = UserControl.ScaleWidth
    
    TempRect = CreateRectRgn(0, 0, 0, 0)
    If m_IsActive Then
        TempRect1 = CreateRoundRectRgn(0, -1, Wi + 1, He + 10, 8, 8)
        TempRect2 = CreateRectRgn(0, 0, Wi + 1, He + 1)
    Else
        TempRect1 = CreateRoundRectRgn(1, 1, Wi + 1, He + 10, 8, 8)
        TempRect2 = CreateRectRgn(1, 2, Wi + 1, He + 1)
    End If
    CombineRgn TempRect, TempRect2, TempRect1, 1
    SetWindowRgn UserControl.hwnd, TempRect, True
    DeleteObject TempRect1
    DeleteObject TempRect2
    DeleteObject TempRect

End Sub

Private Sub ContainerCheck()
    Dim Control As Object
    For Each Control In UserControl.ParentControls
        If TypeOf Control Is HzxYTabLabel Then
            If Control.hdc <> UserControl.hdc Then
                If Control.IsActive = True Then Control.IsActive = False
            End If
        End If
    Next
End Sub
