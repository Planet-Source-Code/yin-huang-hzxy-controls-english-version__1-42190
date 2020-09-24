VERSION 5.00
Begin VB.UserControl HzxYFrame 
   Appearance      =   0  'Flat
   BackColor       =   &H00D8E9EC&
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1110
   ControlContainer=   -1  'True
   ScaleHeight     =   21
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   74
   ToolboxBitmap   =   "HzxYFrame.ctx":0000
End
Attribute VB_Name = "HzxYFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum fraBorderStyles
    fraNone = 0
    fraFixed_Single = 1
End Enum

Private m_ForeColor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_BorderStyle As fraBorderStyles
Private m_BorderColor As OLE_COLOR
Private m_Caption As String
Private m_Image As StdPicture
Private m_ImageWidth As Long
Private m_ImageHeight As Long
Private CorX_Pic As Long
Private CorY_Pic As Long
Private CorXLeft_Cap As Long
Private CorXRight_Cap As Long
Private CorY_Cap As Long
Private CorY_TopLine As Long
Private CaptionHeight As Long
Private lngFormat As Long
Private CaptionRect As RECT
Private m_ControlContainedControls As Boolean

Private Const m_def_ForeColor = &HD54600
Private Const m_def_BorderColor = &HA09C98
Private Const m_def_ImageWidth = 16
Private Const m_def_ImageHeight = 16
Private Const m_def_BaseLeft = 6
'Events
Event Click()
Event DblClick()
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
    Set UserControl.Font = Parent.Font
    m_BackColor = Parent.BackColor
    m_ForeColor = m_def_ForeColor
    UserControl.BackColor = m_BackColor
    UserControl.ForeColor = m_ForeColor
    m_ControlContainedControls = True
    m_BorderStyle = fraFixed_Single
    m_BorderColor = m_def_BorderColor
    Set m_Image = Nothing
    m_ImageWidth = m_def_ImageWidth
    m_ImageHeight = m_def_ImageHeight
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Caption = .ReadProperty("Caption", Ambient.DisplayName)
        Enabled = .ReadProperty("Enabled", True)
        Set UserControl.Font = .ReadProperty("Font", Parent.Font)
        m_BackColor = .ReadProperty("BackColor", Parent.BackColor)
        m_ForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
        UserControl.BackColor = m_BackColor
        UserControl.ForeColor = m_ForeColor
        m_ControlContainedControls = .ReadProperty("ControlContainedControls", True)
        m_BorderStyle = .ReadProperty("BorderStyle", fraFixed_Single)
        m_BorderColor = .ReadProperty("BorderColor", m_def_BorderColor)
        Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", 0)
        Set m_Image = .ReadProperty("Image", Nothing)
        m_ImageWidth = .ReadProperty("ImageWidth", m_def_ImageWidth)
        m_ImageHeight = .ReadProperty("ImageHeight", m_def_ImageHeight)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim loop1 As Integer
    With PropBag
        Call .WriteProperty("Caption", m_Caption, Ambient.DisplayName)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
        Call .WriteProperty("BackColor", m_BackColor, Parent.BackColor)
        Call .WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
        Call .WriteProperty("ControlContainedControls", m_ControlContainedControls, True)
        Call .WriteProperty("BorderStyle", m_BorderStyle, fraFixed_Single)
        Call .WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
        Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, 0)
        Call .WriteProperty("Image", m_Image, Nothing)
        Call .WriteProperty("ImageWidth", m_ImageWidth, m_def_ImageWidth)
        Call .WriteProperty("ImageHeight", m_ImageHeight, m_def_ImageHeight)
    End With
End Sub
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    If m_BackColor <> New_BackColor Then
        m_BackColor = New_BackColor
        PropertyChanged "BackColor"
        Refresh
    End If
End Property
Public Property Get BorderStyle() As fraBorderStyles
    BorderStyle = m_BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As fraBorderStyles)
    If m_BorderStyle <> New_BorderStyle Then
        m_BorderStyle = New_BorderStyle
        PropertyChanged "BorderStyle"
        Refresh
    End If
End Property
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property
Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    If m_BorderColor <> New_BorderColor Then
        m_BorderColor = New_BorderColor
        PropertyChanged "BorderColor"
        DrawBorder
    End If
End Property
Public Property Get ControlContainedControls() As Boolean
    ControlContainedControls = m_ControlContainedControls
End Property
Public Property Let ControlContainedControls(ByVal New_ControlContainedControls As Boolean)
    m_ControlContainedControls = New_ControlContainedControls
    PropertyChanged "ControlContainedControls"
End Property
Public Property Get Caption() As String
    Caption = m_Caption
End Property
Public Property Let Caption(NewCaption As String)
    m_Caption = NewCaption
    PropertyChanged "Caption"
'    DrawRectangle UserControl.hdc, CorXLeft_Cap, 0, CorXRight_Cap, CorY_TopLine + 1, BreakApart(m_BackColor)
    Refresh
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    If New_Enabled <> UserControl.Enabled Then
        UserControl.Enabled() = New_Enabled
        PropertyChanged "Enabled"
        DrawCaption
        DrawPicture
        DrawBorder
        Dim Control As Object
        If m_ControlContainedControls Then
            For Each Control In UserControl.ContainedControls
                Control.Enabled = New_Enabled
            Next
        Else
            For Each Control In UserControl.ContainedControls
                Control.Refresh
            Next
        End If
    End If
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    If m_ForeColor <> New_ForeColor Then
        m_ForeColor = New_ForeColor
        PropertyChanged "ForeColor"
        DrawCaption
    End If
End Property
Public Property Get Font() As Font
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Font.VB_UserMemId = -512
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
Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
Public Property Get Image() As StdPicture
    Set Image = m_Image
End Property
Public Property Set Image(ByVal NewImage As StdPicture)
    Set m_Image = NewImage
    PropertyChanged "Image"
    Refresh
End Property
Public Property Get ImageHeight() As Long
    ImageHeight = m_ImageHeight
End Property
Public Property Let ImageHeight(ByVal NewImageHeight As Long)
    If m_ImageHeight <> NewImageHeight Then
        m_ImageHeight = NewImageHeight
        PropertyChanged "ImageHeight"
        If Not m_Image Is Nothing Then Refresh
    End If
End Property
Public Property Get ImageWidth() As Long
    ImageWidth = m_ImageWidth
End Property
Public Property Let ImageWidth(ByVal NewImageWidth As Long)
    If m_ImageWidth <> NewImageWidth Then
        m_ImageWidth = NewImageWidth
        PropertyChanged "ImageWidth"
        If Not m_Image Is Nothing Then Refresh
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
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Paint()
    Refresh
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

Public Sub Refresh()
    DrawRectangle UserControl.hdc, m_def_BaseLeft + 1, 0, CorXRight_Cap, CorY_TopLine + 1, BreakApart(m_BackColor)
    CalPosition
    DrawBlock
    If Trim(m_Caption) <> "" Then DrawCaption
    If Not m_Image Is Nothing Then DrawPicture
    If m_BorderStyle = fraFixed_Single Then DrawBorder
    RoundCorners
End Sub

Public Sub CalPosition()
        
    Dim TmpRect As RECT
    Dim TextSize As Size
    Dim BaseLeft As Long
    
    UserControl.ScaleMode = vbPixels
    BaseLeft = m_def_BaseLeft
    
    CorY_TopLine = 0
    CorX_Pic = 0
    CorY_Pic = 0
    CorXLeft_Cap = 0
    CorXRight_Cap = 0
    CorY_Cap = 0
    
    If Not m_Image Is Nothing Then
        CorY_TopLine = m_ImageHeight \ 2
        CorX_Pic = BaseLeft + 2
        BaseLeft = CorX_Pic + m_ImageWidth
        CorXRight_Cap = BaseLeft + 2
    End If
    
    If Trim(m_Caption) <> "" Then
        CorXLeft_Cap = BaseLeft + 2
        GetTextExtentPoint32 UserControl.hdc, m_Caption, LenB(StrConv(m_Caption, vbFromUnicode)), TextSize
        CorXRight_Cap = CorXLeft_Cap + TextSize.cx + 2
        Call SetRect(TmpRect, CorXLeft_Cap + 1, 0, CorXRight_Cap - 1, UserControl.ScaleHeight)
        lngFormat = DT_WORDBREAK Or DT_LEFT
        CaptionHeight = DrawText(UserControl.hdc, m_Caption, -1, TmpRect, lngFormat Or DT_CALCRECT)
        If CaptionHeight > 1 Then
            If CaptionHeight \ 2 >= CorY_TopLine Then
                CorY_TopLine = CaptionHeight \ 2
                CorY_Pic = (CaptionHeight - m_ImageHeight) \ 2
                Call SetRect(CaptionRect, CorXLeft_Cap + 1, 0, CorXRight_Cap - 1, CaptionHeight)
            Else
                CorY_Cap = CorY_TopLine - CaptionHeight \ 2
                Call SetRect(CaptionRect, CorXLeft_Cap + 1, CorY_Cap, CorXRight_Cap - 1, CorY_Cap + CaptionHeight)
            End If
        End If
    End If

End Sub

Private Sub DrawBlock()
    Dim Wi As Long, He As Long
    
    With UserControl
        Wi = .ScaleWidth
        He = .ScaleHeight
        DrawRectangle .hdc, 0, CorY_TopLine, Wi, He, BreakApart(m_BackColor)
        DrawRectangle .hdc, m_def_BaseLeft + 1, 0, CorXRight_Cap, CorY_TopLine + 1, BreakApart(m_BackColor)
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
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, m_ImageWidth, m_ImageHeight, BreakApart(m_BackColor)
            GetOriWH m_Image, OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, m_ImageWidth, m_ImageHeight, m_Image, OriW, OriH
        Else
            DrawRectangle UserControl.hdc, CorX_Pic, CorY_Pic, m_ImageWidth, m_ImageHeight, BreakApart(m_BackColor)
            GetOriWH m_Image, OriW, OriH
            TransBlt UserControl.hdc, CorX_Pic, CorY_Pic, m_ImageWidth, m_ImageHeight, m_Image, OriW, OriH, Icon_Grey
        End If
    End If
End Sub

Private Sub GetOriWH(ByVal srcPic As StdPicture, OriW As Long, OriH As Long)
    
    OriW = UserControl.ScaleX(srcPic.Width, vbHimetric, vbPixels)
    OriH = UserControl.ScaleY(srcPic.Height, vbHimetric, vbPixels)

End Sub
Private Sub DrawBorder()

    Dim Color As Long
    Dim loop1 As Integer
    Dim Wi As Long, He As Long
    Dim TabLeftPos As Long
    Dim oldPen As Long, hPen As Long

    If m_BorderStyle = fraNone Then Exit Sub
    
    With UserControl
        Wi = .ScaleWidth
        He = .ScaleHeight
    End With
    
    Color = IIf(UserControl.Enabled, m_BorderColor, ShiftColor(&HFFFFFF, -&H3C, True))
    
    DrawLine UserControl.hdc, 0, CorY_TopLine, 0, He - 1, Color
    DrawLine UserControl.hdc, Wi - 1, CorY_TopLine, Wi - 1, He - 1, Color
    DrawLine UserControl.hdc, 0, He - 1, Wi - 1, He - 1, Color
    
    If m_Image Is Nothing And Trim(m_Caption) = "" Then
        DrawLine UserControl.hdc, 0, CorY_TopLine, Wi - 1, CorY_TopLine, Color
    ElseIf m_Image Is Nothing Then
        DrawLine UserControl.hdc, 0, CorY_TopLine, CorXLeft_Cap - 2, CorY_TopLine, Color
        DrawLine UserControl.hdc, CorXRight_Cap, CorY_TopLine, Wi - 1, CorY_TopLine, Color
    Else
        DrawLine UserControl.hdc, 0, CorY_TopLine, m_def_BaseLeft, CorY_TopLine, Color
        DrawLine UserControl.hdc, CorXRight_Cap, CorY_TopLine, Wi - 1, CorY_TopLine, Color
    End If
    
    With UserControl
        hPen = CreatePen(PS_SOLID, 1, Color)
        oldPen = SelectObject(.hdc, hPen)
        Arc .hdc, 0, CorY_TopLine, 8, CorY_TopLine + 8, 4, CorY_TopLine, 0, CorY_TopLine + 4
        Arc .hdc, Wi - 8, CorY_TopLine, Wi, CorY_TopLine + 8, Wi, CorY_TopLine + 4, Wi - 4, CorY_TopLine
        Arc .hdc, 0, He - 8, 8, He, 0, He - 4, 4, He
        Arc .hdc, Wi - 8, He - 8, Wi, He, Wi - 4, He, Wi, He - 4
        SelectObject .hdc, oldPen
        DeleteObject hPen
    End With

End Sub

Private Sub RoundCorners()
    Dim TempRect As Long, TempRect1 As Long, TempRect2 As Long, TempRect3 As Long
    Dim He As Long, Wi As Long
    Dim loop1 As Integer
    Dim re As Long
    Dim TabLeftPos As Long
    
    He = UserControl.ScaleHeight
    Wi = UserControl.ScaleWidth
    
    TempRect = CreateRectRgn(0, 0, Wi, He)
    TempRect1 = CreateRoundRectRgn(0, CorY_TopLine - 1, Wi + 1, He + 1, 8, 8)
    TempRect2 = CreateRectRgn(0, CorY_TopLine, Wi + 1, He + 1)
    CombineRgn TempRect, TempRect2, TempRect1, RGN_AND
    DeleteObject TempRect2
    DeleteObject TempRect1
        
    If CorXRight_Cap > 0 Then
        TempRect1 = CreateRectRgn(m_def_BaseLeft, 0, CorXRight_Cap, CorY_TopLine + 1)
        CombineRgn TempRect, TempRect, TempRect1, RGN_OR
        DeleteObject TempRect1
    End If
    
    SetWindowRgn UserControl.hwnd, TempRect, True
    DeleteObject TempRect
End Sub
