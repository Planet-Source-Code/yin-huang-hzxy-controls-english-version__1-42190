VERSION 5.00
Begin VB.UserControl HzxYProgressBar 
   Appearance      =   0  'Flat
   BackColor       =   &H00D8E9EC&
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   ControlContainer=   -1  'True
   ScaleHeight     =   23
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   129
   ToolboxBitmap   =   "HzxYProgressBar.ctx":0000
End
Attribute VB_Name = "HzxYProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum prgBarStyles
    ProgressBar = 0
    SearchBar = 1
End Enum

Public Enum prgBorderStyles
    prgNone = 0
    prgFixed_Single = 1
End Enum

Public Enum prgColorSets
    Custum = 0
    XP_Default = 1
    XP_Blue = 2
    XP_DarkBlue = 3
    XP_Gold = 4
    XP_Green = 5
    XP_Grey = 6
    XP_Orange = 7
    XP_Red = 8
End Enum

Public Enum prgFillDirections
    prgLeft = 0
    prgRight = 1
    prgUp = 2
    prgDown = 3
End Enum

Private m_BackColor As OLE_COLOR
Private m_BarColorSet As prgColorSets
Private m_BarBorderStyle As prgBorderStyles
Private m_BarFillDirection As prgFillDirections
Private m_BarImage As StdPicture
Private m_BarImageHeight As Long
Private m_BarImageWidth As Long
Private m_BarSpaceBetweenImages As Long
Private m_BarStyle As prgBarStyles
Private m_BorderColor As OLE_COLOR

Private m_Mini, m_Maxi, m_Value As Long, m_LastValue As Long
Private m_ForceRedraw As Boolean
Private Wi As Long, He As Long
Private StepImage As Integer
Private BeginPos As Integer, EndPos As Integer, TotalSize As Integer
Private Target_X As Single, Target_Y As Single, Target_Width As Single, Target_Height As Single
Private Source_X As Single, Source_Y As Single, Source_Width As Single, Source_Height As Single

Private Const m_def_BackColor = &HFFFCFF
Private Const m_def_BarColorSet = prgColorSets.XP_Default
Private Const m_def_BarFillDirection = prgRight
Private Const m_def_BarSpaceBetweenImages = 2
Private Const m_def_BarStyle = prgBarStyles.ProgressBar
Private Const m_def_BorderColor = &HA09C98
Private Const m_def_Max = 100
Private Const m_def_Min = 0

Event Resize()

Private Sub UserControl_Initialize()
    Ini
    With UserControl
        .ScaleMode = vbPixels
        .PaletteMode = vbPaletteModeContainer
    End With
End Sub

Private Sub UserControl_InitProperties()
    m_Maxi = m_def_Max
    m_Mini = m_def_Min
    m_Value = m_def_Min
    m_BarStyle = m_def_BarStyle
    m_BarColorSet = m_def_BarColorSet
    m_BarBorderStyle = prgFixed_Single
    m_BorderColor = m_def_BorderColor
    m_BarFillDirection = m_def_BarFillDirection
    m_BackColor = m_def_BackColor
    LoadImage
    GetSize
    Enabled = True
    m_ForceRedraw = True
    m_LastValue = m_Value
    m_BarSpaceBetweenImages = m_def_BarSpaceBetweenImages
End Sub

Private Sub LoadImage()
    Dim str As String
    Select Case m_BarFillDirection
    Case prgLeft, prgRight
        str = "0"
    Case prgUp, prgDown
        str = "1"
    End Select
    Set m_BarImage = LoadResPicture("prg" & m_BarColorSet & str, vbResBitmap)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Mini = .ReadProperty("Min", m_def_Min)
        m_Maxi = .ReadProperty("Max", m_def_Max)
        m_Value = .ReadProperty("Value", m_def_Min)
        m_BarStyle = .ReadProperty("BarStyle", m_def_BarStyle)
        m_BarColorSet = .ReadProperty("BarColorSet", m_def_BarColorSet)
        Set m_BarImage = .ReadProperty("Bar_Pic", Nothing)
        m_BarBorderStyle = .ReadProperty("BarBorderStyle", prgFixed_Single)
        m_BorderColor = .ReadProperty("BorderColor", m_def_BorderColor)
        m_BarFillDirection = .ReadProperty("BarFillDirection", m_def_BarFillDirection)
        m_BackColor = .ReadProperty("BackColor", m_def_BackColor)
        Enabled = .ReadProperty("Enabled", True)
        m_BarSpaceBetweenImages = .ReadProperty("BarSpaceBetweenImages", m_def_BarSpaceBetweenImages)
    End With
    m_ForceRedraw = True
    m_LastValue = m_Value
    GetSize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("Min", m_Mini, m_def_Min)
        Call .WriteProperty("Max", m_Maxi, m_def_Max)
        Call .WriteProperty("Value", m_Value, m_def_Min)
        Call .WriteProperty("BarStyle", m_BarStyle, m_def_BarStyle)
        Call .WriteProperty("BarColorSet", m_BarColorSet, m_def_BarColorSet)
        Call .WriteProperty("Bar_Pic", m_BarImage, Nothing)
        Call .WriteProperty("BarBorderStyle", m_BarBorderStyle, prgFixed_Single)
        Call .WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
        Call .WriteProperty("BarFillDirection", m_BarFillDirection, m_def_BarFillDirection)
        Call .WriteProperty("BackColor", m_BackColor, m_def_BackColor)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("BarSpaceBetweenImages", m_BarSpaceBetweenImages, m_def_BarSpaceBetweenImages)
    End With
End Sub
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    If m_BackColor <> New_BackColor Then
        m_BackColor = New_BackColor
        PropertyChanged "BackColor"
        m_ForceRedraw = True
        Refresh
    End If
End Property
Public Property Get BarBorderStyle() As prgBorderStyles
    BarBorderStyle = m_BarBorderStyle
End Property
Public Property Let BarBorderStyle(ByVal New_BarBorderStyle As prgBorderStyles)
    If m_BarBorderStyle <> New_BarBorderStyle Then
        m_BarBorderStyle = New_BarBorderStyle
        PropertyChanged "BarBorderStyle"
        m_ForceRedraw = True
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
        m_ForceRedraw = True
        Refresh
    End If
End Property
Public Property Get BarColorSet() As prgColorSets
    BarColorSet = m_BarColorSet
End Property
Public Property Let BarColorSet(ByVal New_BarColorSet As prgColorSets)
    m_BarColorSet = New_BarColorSet
    PropertyChanged "BarColorSet"
    If m_BarColorSet <> Custum Then
        LoadImage
        GetSize
        m_ForceRedraw = True
        Refresh
    End If
End Property
Public Property Get BarFillDirection() As prgFillDirections
    BarFillDirection = m_BarFillDirection
End Property
Public Property Let BarFillDirection(ByVal New_BarFillDirection As prgFillDirections)
    m_BarFillDirection = New_BarFillDirection
    PropertyChanged "BarFillDirection"
    m_ForceRedraw = True
    Refresh
End Property
Public Property Get Bar_Pic() As StdPicture
    Set Bar_Pic = m_BarImage
End Property
Public Property Set Bar_Pic(ByVal newPic As StdPicture)
    Set m_BarImage = newPic
    PropertyChanged "Bar_Pic"
    GetSize
    m_ForceRedraw = True
    Refresh
End Property
Public Property Get BarSpaceBetweenImages() As Long
    BarSpaceBetweenImages = m_BarSpaceBetweenImages
End Property
Public Property Let BarSpaceBetweenImages(ByVal New_BarSpaceBetweenImages As Long)
    m_BarSpaceBetweenImages = New_BarSpaceBetweenImages
    PropertyChanged "BarSpaceBetweenImages"
    m_ForceRedraw = True
    Refresh
End Property
Public Property Get BarStyle() As prgBarStyles
    BarStyle = m_BarStyle
End Property
Public Property Let BarStyle(ByVal New_BarStyle As prgBarStyles)
    If m_BarStyle <> New_BarStyle Then
        m_BarStyle = New_BarStyle
        PropertyChanged "BarStyle"
        m_ForceRedraw = True
        Refresh
    End If
End Property
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    If New_Enabled <> UserControl.Enabled Then
        UserControl.Enabled() = New_Enabled
        PropertyChanged "Enabled"
        m_ForceRedraw = True
        Refresh
    End If
End Property
Public Property Get Min() As Long
    Min = m_Mini
End Property
Public Property Let Min(ByVal New_Mini As Long)
    If New_Mini = m_Maxi Then
        MsgBox "Minimum can NOT be bigger than Maximum!", vbCritical, "Error"
        Exit Property
    ElseIf New_Mini = m_Maxi Then
        MsgBox "Minimum can NOT equal to Maximum!", vbCritical, "Error"
        Exit Property
    Else
        m_Mini = New_Mini
        PropertyChanged "Min"
        m_ForceRedraw = True
        Refresh
    End If
End Property
Public Property Get Max() As Long
    Max = m_Maxi
End Property
Public Property Let Max(ByVal New_Maxi As Long)
    If New_Maxi < m_Mini Then
        MsgBox "Maximum can NOT be smaller than Minimum!", vbCritical, "Error"
        Exit Property
    ElseIf New_Maxi = m_Mini Then
        MsgBox "Maximum can NOT equal to Minimum!", vbCritical, "Error"
        Exit Property
    Else
        m_Maxi = New_Maxi
        m_ForceRedraw = True
        PropertyChanged "Max"
        Refresh
    End If
End Property
Public Property Get Value() As Long
    Value = m_Value
End Property
Public Property Let Value(ByVal New_Value As Long)
    If UserControl.Ambient.UserMode = False And (New_Value > m_Maxi Or New_Value < m_Mini) Then
        If New_Value > m_Maxi Then
            MsgBox "Value can NOT be bigger than Maximum!", vbCritical, "Error"
            Exit Property
        ElseIf New_Value < m_Mini Then
            MsgBox "Value can NOT be smaller than Minimum!", vbCritical, "Error"
            Exit Property
        End If
    Else
        m_Value = New_Value
        PropertyChanged "Value"
        Refresh
    End If
End Property

Private Sub UserControl_Paint()
    m_ForceRedraw = True
    Refresh
End Sub

Private Sub UserControl_Resize()
    m_ForceRedraw = True
    Refresh
End Sub

Public Sub Refresh()
    UserControl.ScaleMode = vbPixels
    If m_ForceRedraw Then
        DrawBack
        GetSize
    End If
    Select Case m_BarStyle
    Case ProgressBar
       DrawProgressValue
    Case SearchBar
        DrawSearchValue
    End Select
    m_LastValue = m_Value
    m_ForceRedraw = False
End Sub

Private Sub UserControl_Terminate()
    Set m_BarImage = Nothing
End Sub

Private Sub DrawBack()
    With UserControl
        Wi = .ScaleWidth
        He = .ScaleHeight
        DrawRectangle .hdc, 0, 0, Wi, He, BreakApart(m_BackColor)
    End With
    
    If m_BarBorderStyle = prgFixed_Single Then DrawBorder
    RoundCorners
End Sub

Private Sub DrawBorder()
    Dim Color As Long
    Dim oldPen As Long, hPen As Long
    
    Color = IIf(UserControl.Enabled, m_BorderColor, ShiftColor(&HFFFFFF, -&H3C, True))
    
    With UserControl
        DrawRectangle .hdc, 0, 0, Wi, He, Color, True
    
        hPen = CreatePen(PS_SOLID, 1, Color)
        oldPen = SelectObject(.hdc, hPen)
        Arc .hdc, 0, 0, 8, 8, 4, 0, 0, 4
        Arc .hdc, Wi - 8, 0, Wi, 8, Wi, 4, Wi - 4, 0
        Arc .hdc, 0, He - 8, 8, He, 0, He - 4, 4, He
        Arc .hdc, Wi - 8, He - 8, Wi, He, Wi - 4, He, Wi, He - 4
        SelectObject UserControl.hdc, oldPen
        DeleteObject hPen
    End With
End Sub

Private Sub RoundCorners()
    Dim TempRect As Long, TempRect1 As Long, TempRect2 As Long
    
    TempRect = CreateRectRgn(0, 0, Wi, He)
    If m_BarBorderStyle = prgFixed_Single Then
        TempRect1 = CreateRoundRectRgn(0, -1, Wi + 1, He + 1, 8, 8)
        TempRect2 = CreateRectRgn(0, 0, Wi + 1, He + 1)
        CombineRgn TempRect, TempRect2, TempRect1, RGN_AND
        SetWindowRgn UserControl.hwnd, TempRect, True
        DeleteObject TempRect1
        DeleteObject TempRect2
    Else
        SetWindowRgn UserControl.hwnd, TempRect, True
    End If
    DeleteObject TempRect
End Sub

Private Sub GetSize()
    
    If m_BarImage Is Nothing Then Exit Sub
    
    Dim Temp As Integer
    
    Select Case m_BarFillDirection
    Case prgLeft, prgRight
        With UserControl
            m_BarImageWidth = .ScaleX(m_BarImage.Width, vbHimetric, vbPixels)
            m_BarImageHeight = .ScaleHeight - 6
        End With
        StepImage = m_BarImageWidth + m_BarSpaceBetweenImages
        Select Case m_BarFillDirection
        Case prgRight
            BeginPos = 3
            EndPos = UserControl.ScaleWidth - 4
            Temp = EndPos - BeginPos + 1
        Case prgLeft
            BeginPos = UserControl.ScaleWidth - 4
            EndPos = 3
            Temp = BeginPos - EndPos + 1
        End Select
        
        Select Case m_BarStyle
        Case ProgressBar
            TotalSize = IIf((Temp Mod StepImage) >= m_BarImageWidth, m_BarImageWidth, Temp Mod StepImage)
            TotalSize = (Temp \ StepImage) * m_BarImageWidth + TotalSize
        Case SearchBar
            TotalSize = Temp
        End Select
    
        Target_Y = 3
        Target_Height = m_BarImageHeight
        Source_Y = 0
        Source_Height = UserControl.ScaleY(m_BarImage.Height, vbHimetric, vbPixels)
    Case prgUp, prgDown
        
        With UserControl
            m_BarImageWidth = .ScaleWidth - 6
            m_BarImageHeight = .ScaleY(m_BarImage.Height, vbHimetric, vbPixels)
        End With
        StepImage = m_BarImageHeight + m_BarSpaceBetweenImages
        
        Select Case m_BarFillDirection
        Case prgDown
            BeginPos = 3
            EndPos = UserControl.ScaleHeight - 3
            Temp = EndPos - BeginPos + 1
        Case prgUp
            BeginPos = UserControl.ScaleHeight - 3
            EndPos = 3
            Temp = BeginPos - EndPos + 1
        End Select
        
        Select Case m_BarStyle
        Case ProgressBar
            TotalSize = IIf((Temp Mod StepImage) >= m_BarImageHeight, m_BarImageHeight, Temp Mod StepImage)
            TotalSize = (Temp \ StepImage) * m_BarImageHeight + TotalSize
        Case SearchBar
            TotalSize = Temp
        End Select
    
        Target_X = 3
        Target_Width = m_BarImageWidth
        Source_X = 0
        Source_Width = UserControl.ScaleX(m_BarImage.Width, vbHimetric, vbPixels)
    End Select

End Sub

Private Sub DrawProgressValue()

    Dim loop1 As Long
    Dim CurrentMaxValue As Long
    Dim ScaledLastValue As Long
    Dim ScaledValue As Long
    Dim ImageSize As Long
    
    If m_ForceRedraw Then m_LastValue = m_Mini
    If m_Value > m_LastValue Then
        If m_Value > m_Mini Then
    
            If m_LastValue < m_Mini Then m_LastValue = m_Mini
            ScaledLastValue = (m_LastValue - m_Mini) * TotalSize / (m_Maxi - m_Mini)
            ScaledValue = (m_Value - m_Mini) * TotalSize / (m_Maxi - m_Mini)
            ImageSize = StepImage - m_BarSpaceBetweenImages
            
            For loop1 = (ScaledLastValue \ ImageSize) To (ScaledValue \ ImageSize)
                CurrentMaxValue = (loop1 + 1) * ImageSize * (m_Maxi - m_Mini) \ TotalSize
                If CurrentMaxValue <= m_Value Then
                    Select Case m_BarFillDirection
                    Case prgLeft
                        Target_X = BeginPos - (loop1 + 1) * StepImage + m_BarSpaceBetweenImages
                        Target_Width = m_BarImageWidth
                        If Target_X < EndPos Then
                            Target_Width = BeginPos - loop1 * StepImage - EndPos
                            Target_X = EndPos
                        End If
                        If Target_Width < 0 Then Target_Width = 0
                        Source_Width = Target_Width
                        Source_X = m_BarImageWidth - Source_Width
                    Case prgRight
                        Target_X = loop1 * StepImage + BeginPos
                        Target_Width = m_BarImageWidth
                        If Target_X + Target_Width > EndPos Then Target_Width = EndPos - Target_X
                        If Target_Width < 0 Then Target_Width = 0
                        Source_X = 0
                        Source_Width = Target_Width
                    Case prgUp
                        Target_Y = BeginPos - (loop1 + 1) * StepImage + m_BarSpaceBetweenImages
                        Target_Height = m_BarImageHeight
                        If Target_Y < EndPos Then
                            Target_Height = Target_Height - EndPos + Target_Y
                            Target_Y = EndPos
                        End If
                        If Target_Height < 0 Then Target_Height = 0
                        Source_Height = Target_Height
                        Source_Y = m_BarImageHeight - Target_Height
                    Case prgDown
                        Target_Y = loop1 * StepImage + BeginPos
                        Target_Height = m_BarImageHeight
                        If Target_Y + Target_Height > EndPos Then Target_Height = EndPos - Target_Y
                        If Target_Height < 0 Then Target_Height = 0
                        Source_Y = 0
                        Source_Height = Target_Height
                    End Select
                    If Abs(Target_Width * Target_Height) > 0 Then
                        UserControl.PaintPicture m_BarImage, _
                                                 Target_X, Target_Y, Target_Width, Target_Height, _
                                                 Source_X, Source_Y, Source_Width, Source_Height
                    End If
                Else
                    Select Case m_BarFillDirection
                    Case prgLeft
                        Target_Width = ScaledValue Mod ImageSize
                        Target_X = BeginPos - loop1 * StepImage - Target_Width
                        If Target_X < EndPos Then
                            Target_Width = BeginPos - loop1 * StepImage - EndPos
                            Target_X = EndPos
                        End If
                        If Target_Width < 0 Then Target_Width = 0
                        If Target_Width > ImageSize Then Target_Width = ImageSize
                        Source_Width = Target_Width
                        Source_X = ImageSize - Source_Width
                    Case prgRight
                        Target_X = loop1 * StepImage + BeginPos
                        Target_Width = ScaledValue Mod ImageSize
                        If Target_X + Target_Width > EndPos Then Target_Width = EndPos - Target_X
                        If Target_Width < 0 Then Target_Width = 0
                        If Target_Width > ImageSize Then Target_Width = ImageSize
                        Source_X = 0
                        Source_Width = Target_Width
                    Case prgUp
                        Target_Height = ScaledValue Mod ImageSize
                        Target_Y = BeginPos - loop1 * StepImage - Target_Height
                        If Target_Y < EndPos Then
                            Target_Height = BeginPos - loop1 * StepImage - EndPos
                            Target_Y = EndPos
                        End If
                        If Target_Height < 0 Then Target_Height = 0
                        If Target_Height > ImageSize Then Target_Height = ImageSize
                        Source_Height = Target_Height
                        Source_Y = ImageSize - Source_Height
                    Case prgDown
                        Target_Y = loop1 * StepImage + BeginPos
                        Target_Height = ScaledValue Mod ImageSize
                        If Target_Y + Target_Height > EndPos Then Target_Height = EndPos - Target_Y
                        If Target_Height < 0 Then Target_Height = 0
                        If Target_Height > ImageSize Then Target_Height = ImageSize
                        Source_Y = 0
                        Source_Height = Target_Height
                    End Select
                    If Abs(Target_Width * Target_Height) > 0 Then
                        UserControl.PaintPicture m_BarImage, _
                                                 Target_X, Target_Y, Target_Width, Target_Height, _
                                                 Source_X, Source_Y, Source_Width, Source_Height
                    End If
                    Exit For
                End If
            Next
        End If
    ElseIf m_Value < m_LastValue Then
        If m_Value < m_Maxi Then
                
                Dim Blank_BeginX As Single, Blank_BeginY As Single
                Dim Blank_Width As Single, Blank_Height As Single
                
                ScaledValue = (m_Value - m_Mini) * TotalSize / (m_Maxi - m_Mini)
                ImageSize = StepImage - m_BarSpaceBetweenImages
                
                Select Case m_BarFillDirection
                Case prgLeft
                    Blank_BeginX = EndPos
                    Blank_BeginY = Target_Y
                    Blank_Width = BeginPos - (ScaledValue \ ImageSize) * StepImage - (ScaledValue Mod ImageSize) - EndPos
                    Blank_Height = Target_Height
                Case prgRight
                    Blank_BeginX = (ScaledValue \ ImageSize) * StepImage + (ScaledValue Mod ImageSize) + BeginPos
                    Blank_BeginY = Target_Y
                    Blank_Width = EndPos - Blank_BeginX
                    Blank_Height = Target_Height
                Case prgUp
                    Blank_BeginX = Target_X
                    Blank_BeginY = EndPos
                    Blank_Width = Target_Width
                    Blank_Height = BeginPos - (ScaledValue \ ImageSize) * StepImage - (ScaledValue Mod ImageSize) - EndPos
                Case prgDown
                    Blank_BeginX = Target_X
                    Blank_BeginY = (ScaledValue \ ImageSize) * StepImage + (ScaledValue Mod ImageSize) + BeginPos
                    Blank_Width = Target_Width
                    Blank_Height = EndPos - Blank_BeginY
                End Select
                
                DrawRectangle UserControl.hdc, Blank_BeginX, Blank_BeginY, Blank_Width, Blank_Height, BreakApart(m_BackColor)
        End If
    
    End If
    
End Sub

Private Sub DrawSearchValue()

    If m_Value <= m_Mini Or m_Value > m_Maxi Then Exit Sub
    Dim loop1 As Long
    Dim ImageSize As Long
    Dim ScaledValue As Long
    Dim Blank_BeginX As Single, Blank_BeginY As Single
    Dim Blank_Width As Single, Blank_Height As Single
                
    Select Case m_BarFillDirection
    Case prgLeft
        Blank_BeginX = EndPos - 1
        Blank_BeginY = Target_Y
        Blank_Width = TotalSize + 2
        Blank_Height = Target_Height
    Case prgRight
        Blank_BeginX = BeginPos - 1
        Blank_BeginY = Target_Y
        Blank_Width = TotalSize + 2
        Blank_Height = Target_Height
    Case prgUp
        Blank_BeginX = Target_X
        Blank_BeginY = EndPos - 1
        Blank_Width = Target_Width
        Blank_Height = TotalSize + 2
    Case prgDown
        Blank_BeginX = Target_X
        Blank_BeginY = BeginPos - 1
        Blank_Width = Target_Width
        Blank_Height = TotalSize + 2
    End Select
    DrawRectangle UserControl.hdc, Blank_BeginX, Blank_BeginY, Blank_Width, Blank_Height, BreakApart(m_BackColor)
    
    ImageSize = StepImage - m_BarSpaceBetweenImages
    ScaledValue = (m_Value - m_Mini) * (TotalSize - 3 * StepImage + m_BarSpaceBetweenImages) / (m_Maxi - m_Mini)
    
    Select Case m_BarFillDirection
    Case prgLeft, prgRight
        Target_Width = m_BarImageWidth
        Source_X = 0
        Source_Width = m_BarImageWidth
    Case prgUp, prgDown
        Target_Height = m_BarImageHeight
        Source_Y = 0
        Source_Height = m_BarImageHeight
    End Select
    
    For loop1 = 0 To 2
        Select Case m_BarFillDirection
        Case prgLeft
            Target_X = BeginPos - ScaledValue - (loop1 + 1) * StepImage + m_BarSpaceBetweenImages + 1
        Case prgRight
            Target_X = ScaledValue + loop1 * StepImage + BeginPos - 1
        Case prgUp
            Target_Y = BeginPos - ScaledValue - (loop1 + 1) * StepImage + m_BarSpaceBetweenImages + 1
        Case prgDown
            Target_Y = ScaledValue + loop1 * StepImage + BeginPos - 1
        End Select
        If Abs(Target_Width * Target_Height) > 0 Then
            UserControl.PaintPicture m_BarImage, _
                                     Target_X, Target_Y, Target_Width, Target_Height, _
                                     Source_X, Source_Y, Source_Width, Source_Height
        End If
    Next

End Sub
