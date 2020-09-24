VERSION 5.00
Begin VB.UserControl HzxYTabStrip 
   Appearance      =   0  'Flat
   BackColor       =   &H00D8E9EC&
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3675
   ControlContainer=   -1  'True
   FillStyle       =   0  'Solid
   ScaleHeight     =   125
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   245
   ToolboxBitmap   =   "HzxYTabStrip.ctx":0000
   Begin HzxYControlsEnglish.HzxYTabLabel TabLabel 
      Height          =   510
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1230
      _ExtentX        =   265
      _ExtentY        =   900
   End
End
Attribute VB_Name = "HzxYTabStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_ForeColor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_CurrentTab As Integer 'Currently selected tab
Private m_ControlChildControl As Boolean

'Events
Event TabClick(NewTabIndex As Integer, OldTabIndex As Integer)

Private Sub UserControl_Initialize()
    UserControl.ScaleMode = vbPixels
    UserControl.PaletteMode = vbPaletteModeContainer
End Sub

Private Sub UserControl_InitProperties()
    Ini
    Enabled = True
    Set UserControl.Font = Parent.Font
    m_ForeColor = Parent.ForeColor
    m_BackColor = &HFEFCFC
    UserControl.BackColor = m_BackColor
    UserControl.ForeColor = m_ForeColor
    m_ControlChildControl = True
    Set TabLabel(0).Font = UserControl.Font
    CountTabs = 3
    CurrentTab = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim loop1 As Integer
    With PropBag
        Enabled = .ReadProperty("Enabled", True)
        Set UserControl.Font = .ReadProperty("Font", Parent.Font)
        m_BackColor = .ReadProperty("BackColor", &HFEFCFC)
        m_ForeColor = .ReadProperty("ForeColor", Parent.ForeColor)
        UserControl.BackColor = m_BackColor
        UserControl.ForeColor = m_ForeColor
        m_ControlChildControl = .ReadProperty("ControlChildControl", True)
        Set MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", 0)
        
        CountTabs = .ReadProperty("CountTabs", 3)
        CurrentTab = .ReadProperty("CurrentTab", 0)
        For loop1 = 0 To TabLabel.Count - 1
            TabLabel(loop1).Caption = .ReadProperty("CurrentTab_Caption" & loop1, "[No Caption]")
            Set TabLabel(loop1).Image = .ReadProperty("CurrentTab_Image" & loop1, Nothing)
            Set TabLabel(loop1).Font = UserControl.Font
            Set TabLabel(loop1).MouseIcon = UserControl.MouseIcon
            TabLabel(loop1).MousePointer = UserControl.MousePointer
            If loop1 = CurrentTab Then TabLabel(loop1).IsActive = True
        Next
    End With
    RedrawTabs
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim loop1 As Integer
    With PropBag
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
        Call .WriteProperty("BackColor", m_BackColor, &HFEFCFC)
        Call .WriteProperty("ForeColor", m_ForeColor, Parent.ForeColor)
        Call .WriteProperty("ControlChildControl", m_ControlChildControl, True)
        Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, 0)

        .WriteProperty "CountTabs", TabLabel.Count, 3
        .WriteProperty "CurrentTab", m_CurrentTab, 0
        For loop1 = 0 To TabLabel.Count - 1
            .WriteProperty "CurrentTab_Caption" & loop1, TabLabel(loop1).Caption, "[No Caption]"
            .WriteProperty "CurrentTab_Image" & loop1, TabLabel(loop1).Image, Nothing
        Next
    End With
End Sub
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property
Public Property Get ControlChildControl() As Boolean
    ControlChildControl = m_ControlChildControl
End Property
Public Property Let ControlChildControl(ByVal New_ControlChildControl As Boolean)
    m_ControlChildControl = New_ControlChildControl
    PropertyChanged "ControlChildControl"
End Property
Public Property Get CountTabs() As Integer
    CountTabs = TabLabel.Count
End Property
Public Property Let CountTabs(NewValue As Integer)
    Dim loop1 As Integer
    If NewValue < 1 Then
        MsgBox "You must have at least 1 tab"
    ElseIf NewValue < TabLabel.Count Then
        For loop1 = TabLabel.Count - 1 To NewValue Step -1
            Unload TabLabel(loop1)
        Next
        If m_CurrentTab > NewValue - 1 Then
            CurrentTab = NewValue - 1
        End If
    ElseIf NewValue > TabLabel.Count Then
        For loop1 = TabLabel.Count To NewValue - 1
            Load TabLabel(loop1)
            TabLabel(loop1).Caption = "[No Caption]"
            TabLabel(loop1).IsActive = False
            TabLabel(loop1).ForeColor = m_ForeColor
            Set TabLabel(loop1).Font = UserControl.Font
            Set TabLabel(loop1).Image = Nothing
            TabLabel(loop1).Visible = True
        Next
        RedrawTabs
        Refresh
    End If
    PropertyChanged "CountTabs"
End Property
Public Property Get CurrentTab() As Integer
    CurrentTab = m_CurrentTab
End Property
Public Property Let CurrentTab(NewValue As Integer)
    If NewValue < 0 Or NewValue > TabLabel.Count - 1 Then
        MsgBox "CurrentTab " & NewValue & " does not exist"
    Else
        m_CurrentTab = NewValue
        PropertyChanged "CurrentTab"
        TabLabel(NewValue).IsActive = True
    End If
End Property
Public Property Get CurrentTab_Caption() As String
    CurrentTab_Caption = TabLabel(m_CurrentTab).Caption
End Property
Public Property Let CurrentTab_Caption(NewValue As String)
    TabLabel(m_CurrentTab).Caption = NewValue
    PropertyChanged "CurrentTab_Caption"
    RedrawTabs
    Refresh
End Property
Public Property Get CurrentTab_Image() As StdPicture
    Set CurrentTab_Image = TabLabel(m_CurrentTab).Image
End Property
Public Property Set CurrentTab_Image(NewValue As StdPicture)
    Set TabLabel(m_CurrentTab).Image = NewValue
    PropertyChanged "CurrentTab_Image"
    RedrawTabs
    Refresh
End Property
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    If New_Enabled <> UserControl.Enabled Then
        UserControl.Enabled() = New_Enabled
        PropertyChanged "Enabled"
        DrawOtherLines
        DrawTopLine
        Dim Control As Object
        For Each Control In UserControl
            Control.Enabled = New_Enabled
        Next
        If m_ControlChildControl Then
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
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    Dim loop1 As Integer
    For loop1 = 0 To TabLabel.Count - 1
        TabLabel(loop1).ForeColor = m_ForeColor
    Next
End Property
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Dim loop1 As Integer
    For loop1 = 0 To TabLabel.Count - 1
        Set TabLabel(loop1).Font = UserControl.Font
        RedrawTabs
        Refresh
    Next
End Property
Public Property Get FontUnderline() As Boolean
    FontUnderline = UserControl.FontUnderline
End Property
Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    UserControl.FontUnderline() = New_FontUnderline
    PropertyChanged "Font"
    Dim loop1 As Integer
    For loop1 = 0 To TabLabel.Count - 1
        Set TabLabel(loop1).Font = UserControl.Font
        RedrawTabs
        Refresh
    Next
End Property
Public Property Get FontStrikethru() As Boolean
    FontStrikethru = UserControl.FontStrikethru
End Property
Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    UserControl.FontStrikethru() = New_FontStrikethru
    PropertyChanged "Font"
    Dim loop1 As Integer
    For loop1 = 0 To TabLabel.Count - 1
        Set TabLabel(loop1).Font = UserControl.Font
        RedrawTabs
        Refresh
    Next
End Property
Public Property Get FontSize() As Single
    FontSize = UserControl.FontSize
End Property
Public Property Let FontSize(ByVal New_FontSize As Single)
    UserControl.FontSize() = New_FontSize
    PropertyChanged "Font"
    Dim loop1 As Integer
    For loop1 = 0 To TabLabel.Count - 1
        Set TabLabel(loop1).Font = UserControl.Font
        RedrawTabs
        Refresh
    Next
End Property
Public Property Get FontName() As String
    FontName = UserControl.FontName
End Property
Public Property Let FontName(ByVal New_FontName As String)
    UserControl.FontName() = New_FontName
    PropertyChanged "Font"
    Dim loop1 As Integer
    For loop1 = 0 To TabLabel.Count - 1
        Set TabLabel(loop1).Font = UserControl.Font
        RedrawTabs
        Refresh
    Next
End Property
Public Property Get FontItalic() As Boolean
    FontItalic = UserControl.FontItalic
End Property
Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    UserControl.FontItalic() = New_FontItalic
    PropertyChanged "Font"
    Dim loop1 As Integer
    For loop1 = 0 To TabLabel.Count - 1
        Set TabLabel(loop1).Font = UserControl.Font
        RedrawTabs
        Refresh
    Next
End Property
Public Property Get FontBold() As Boolean
    FontBold = UserControl.FontBold
End Property
Public Property Let FontBold(ByVal New_FontBold As Boolean)
    UserControl.FontBold() = New_FontBold
    PropertyChanged "Font"
    Dim loop1 As Integer
    For loop1 = 0 To TabLabel.Count - 1
        Set TabLabel(loop1).Font = UserControl.Font
        RedrawTabs
        Refresh
    Next
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
    Dim loop1 As Integer
    For loop1 = 0 To TabLabel.Count - 1
        Set TabLabel(loop1).MouseIcon = UserControl.MouseIcon
    Next
End Property
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
    Dim loop1 As Integer
    For loop1 = 0 To TabLabel.Count - 1
        TabLabel(loop1).MousePointer = UserControl.MousePointer
    Next
End Property

Private Sub TabLabel_Click(Index As Integer)
    If Index <> m_CurrentTab Then
        TabLabel(Index).IsActive = True
        RaiseEvent TabClick(Index, m_CurrentTab)
        m_CurrentTab = Index
    End If
End Sub

Private Sub UserControl_Paint()
    Refresh
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

Private Sub Refresh()
    DrawBlock
    DrawOtherLines
    DrawTopLine
End Sub

Private Sub DrawBlock()

    Dim Color As Long
    Dim loop1 As Integer
    Dim Wi As Long, He As Long
    Dim TabLeftPos As Long

    With UserControl
        Wi = .ScaleWidth
        He = .ScaleHeight
        DrawRectangle .hdc, 0, 35, Wi, He, &HFEFCFC
    End With
    
    Color = IIf(UserControl.Enabled, &HA09C98, ShiftColor(&HFFFFFF, -&H3C, True))
    If TabLeftPos < Wi Then DrawLine UserControl.hdc, TabLeftPos, 34, Wi, 34, Color
    DrawLine UserControl.hdc, 0, 34, 0, He - 1, Color
    DrawLine UserControl.hdc, Wi - 1, 34, Wi - 1, He - 1, Color
    DrawLine UserControl.hdc, 0, He - 1, Wi - 1, He - 1, Color
    
    RoundCorners

End Sub

Private Sub DrawOtherLines()

    Dim Color As Long
    Dim loop1 As Integer
    Dim Wi As Long, He As Long
    Dim TabLeftPos As Long

    With UserControl
        Wi = .ScaleWidth
        He = .ScaleHeight
    End With
    
    Color = IIf(UserControl.Enabled, &HA09C98, ShiftColor(&HFFFFFF, -&H3C, True))
    If TabLeftPos < Wi Then DrawLine UserControl.hdc, TabLeftPos, 34, Wi, 34, Color
    DrawLine UserControl.hdc, 0, 34, 0, He - 1, Color
    DrawLine UserControl.hdc, Wi - 1, 34, Wi - 1, He - 1, Color
    DrawLine UserControl.hdc, 0, He - 1, Wi - 1, He - 1, Color

End Sub

Private Sub DrawTopLine()

    Dim Color As Long
    Dim loop1 As Integer
    Dim Wi As Long, He As Long
    Dim TabLeftPos As Long

    With UserControl
        Wi = .ScaleWidth
        He = .ScaleHeight
    End With
    
    TabLeftPos = 0
    For loop1 = 0 To TabLabel.Count - 1
        If loop1 <> m_CurrentTab Then
            Color = IIf(UserControl.Enabled, &HA09C98, ShiftColor(&HFFFFFF, -&H3C, True))
            DrawLine UserControl.hdc, TabLeftPos, 34, TabLeftPos + TabLabel(loop1).Width, 34, Color
        Else
            Color = IIf(UserControl.Enabled, &HFEFCFC, ShiftColor(&HFFFFFF, &H18, True))
            DrawLine UserControl.hdc, TabLeftPos + 1, 34, TabLeftPos + TabLabel(loop1).Width - 1, 34, Color
        End If
        TabLeftPos = TabLeftPos + IIf(loop1 = m_CurrentTab, TabLabel(loop1).Width - 1, TabLabel(loop1).Width)
    Next

End Sub

Private Function RedrawTabs()
    Dim TabLeftPos As Long
    Dim loop1 As Integer
    TabLeftPos = 0
    For loop1 = 0 To TabLabel.Count - 1
        TabLabel(loop1).CalPosition
        TabLabel(loop1).Move TabLeftPos, 0
        TabLeftPos = TabLeftPos + IIf(loop1 = m_CurrentTab, TabLabel(loop1).Width - 1, TabLabel(loop1).Width)
    Next
End Function

Private Sub RoundCorners()
    Dim TempRect As Long, TempRect0 As Long, TempRect1 As Long, TempRect2 As Long, TempRect3 As Long
    Dim He As Long, Wi As Long
    Dim loop1 As Integer
    Dim re As Long
    Dim TabLeftPos As Long
    
    He = UserControl.ScaleHeight
    Wi = UserControl.ScaleWidth
    
    TempRect0 = CreateRectRgn(0, 0, 0, 0)
    TempRect1 = CreateRectRgn(0, 0, Wi, 34)
    
    TabLeftPos = 0
    For loop1 = 0 To TabLabel.Count - 1
        TempRect = CreateRectRgn(0, 0, 0, 0)
        TempRect2 = CreateRectRgn(0, 0, 0, 0)
        GetWindowRgn TabLabel(loop1).hwnd, TempRect2
        OffsetRgn TempRect2, TabLeftPos, 0
        CombineRgn TempRect, TempRect1, TempRect2, RGN_AND
        CombineRgn TempRect0, TempRect, TempRect0, RGN_OR
        DeleteObject TempRect
        DeleteObject TempRect2
        TabLeftPos = TabLeftPos + IIf(loop1 = m_CurrentTab, TabLabel(loop1).Width - 1, TabLabel(loop1).Width)
    Next
    TempRect2 = CreateRectRgn(0, 34, Wi, He)
    TempRect = CreateRectRgn(0, 0, Wi, He)
    CombineRgn TempRect, TempRect0, TempRect2, RGN_OR
    SetWindowRgn UserControl.hwnd, TempRect, True
    DeleteObject TempRect0
    DeleteObject TempRect1
    DeleteObject TempRect2
    DeleteObject TempRect

End Sub
