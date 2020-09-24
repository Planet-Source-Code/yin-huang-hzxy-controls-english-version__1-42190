VERSION 5.00
Begin VB.UserControl HzxYTopButtons 
   CanGetFocus     =   0   'False
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   315
   FillStyle       =   0  'Solid
   MaskColor       =   &H00000000&
   ScaleHeight     =   4.546
   ScaleMode       =   0  'User
   ScaleWidth      =   3.973
   ToolboxBitmap   =   "HzxYTopButtons.ctx":0000
End
Attribute VB_Name = "HzxYTopButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TopButtonTypes
    TopCloseButton = 0
    TopMaxButton = 1
    TopMinButton = 2
    TopRestoreButton = 3
    TopHelp = 4
End Enum

Public Enum TopColorSets
    TopBlue = 0
    TopGreen = 1
    TopSilver = 2
End Enum

Public Enum TopStates
    TopNormal = 0
    TopOver = 1
    TopDown = 2
    TopDisabled = 3
End Enum

Private m_ButtonType As TopButtonTypes
Private TopImage(3) As StdPicture
Private m_State As TopStates
Private m_ColorSet As TopColorSets

Private Const m_def_ButtonType = TopButtonTypes.TopCloseButton
Private Const m_def_ColorSet = TopColorSets.TopBlue

Event Click()
Attribute Click.VB_UserMemId = -600
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607

Private Sub UserControl_Initialize()
    Ini
    With UserControl
        .ScaleMode = vbPixels
        .PaletteMode = vbPaletteModeContainer
    End With
End Sub

Private Sub UserControl_InitProperties()
    m_ButtonType = m_def_ButtonType
    m_ColorSet = m_def_ColorSet
    Enabled = True
    m_State = TopNormal
    LoadImage m_ButtonType
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_ButtonType = .ReadProperty("ButtonType", m_def_ButtonType)
        m_ColorSet = .ReadProperty("m_ColorSet", m_def_ColorSet)
        Enabled = .ReadProperty("Enabled", True)
        Set TopImage(TopNormal) = .ReadProperty("Pic_Normal", Nothing)
        Set TopImage(TopOver) = .ReadProperty("Pic_Over", Nothing)
        Set TopImage(TopDown) = .ReadProperty("Pic_Down", Nothing)
        Set TopImage(TopDisabled) = .ReadProperty("Pic_Disabled", Nothing)
    End With
    m_State = TopNormal
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "ButtonType", m_ButtonType, m_def_ButtonType
        .WriteProperty "m_ColorSet", m_ColorSet, m_def_ColorSet
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "Pic_Normal", TopImage(TopNormal), Nothing
        .WriteProperty "Pic_Over", TopImage(TopOver), Nothing
        .WriteProperty "Pic_Down", TopImage(TopDown), Nothing
        .WriteProperty "Pic_Disabled", TopImage(TopDisabled), Nothing
    End With
End Sub

Private Sub LoadImage(CurButtonType As TopButtonTypes)
    Select Case CurButtonType
    Case TopCloseButton
        Set TopImage(TopNormal) = LoadResPicture("TopQuitNormal" & m_ColorSet, vbResBitmap)
        Set TopImage(TopOver) = LoadResPicture("TopQuitOver" & m_ColorSet, vbResBitmap)
        Set TopImage(TopDown) = LoadResPicture("TopQuitDown" & m_ColorSet, vbResBitmap)
        Set TopImage(TopDisabled) = LoadResPicture("TopQuitDisabled" & m_ColorSet, vbResBitmap)
    Case TopMaxButton
        Set TopImage(TopNormal) = LoadResPicture("TopMaxNormal" & m_ColorSet, vbResBitmap)
        Set TopImage(TopOver) = LoadResPicture("TopMaxOver" & m_ColorSet, vbResBitmap)
        Set TopImage(TopDown) = LoadResPicture("TopMaxDown" & m_ColorSet, vbResBitmap)
        Set TopImage(TopDisabled) = LoadResPicture("TopMaxDisabled" & m_ColorSet, vbResBitmap)
    Case TopMinButton
        Set TopImage(TopNormal) = LoadResPicture("TopMinNormal" & m_ColorSet, vbResBitmap)
        Set TopImage(TopOver) = LoadResPicture("TopMinOver" & m_ColorSet, vbResBitmap)
        Set TopImage(TopDown) = LoadResPicture("TopMinDown" & m_ColorSet, vbResBitmap)
        Set TopImage(TopDisabled) = LoadResPicture("TopMinDisabled" & m_ColorSet, vbResBitmap)
    Case TopRestoreButton
        Set TopImage(TopNormal) = LoadResPicture("TopRestoreNormal" & m_ColorSet, vbResBitmap)
        Set TopImage(TopOver) = LoadResPicture("TopRestoreOver" & m_ColorSet, vbResBitmap)
        Set TopImage(TopDown) = LoadResPicture("TopRestoreDown" & m_ColorSet, vbResBitmap)
        Set TopImage(TopDisabled) = LoadResPicture("TopRestoreDisabled" & m_ColorSet, vbResBitmap)
    Case TopHelp
        Set TopImage(TopNormal) = LoadResPicture("TopHelpNormal" & m_ColorSet, vbResBitmap)
        Set TopImage(TopOver) = LoadResPicture("TopHelpOver" & m_ColorSet, vbResBitmap)
        Set TopImage(TopDown) = LoadResPicture("TopHelpDown" & m_ColorSet, vbResBitmap)
        Set TopImage(TopDisabled) = LoadResPicture("TopHelpDisabled" & m_ColorSet, vbResBitmap)
    End Select
End Sub
Public Property Get ButtonType() As TopButtonTypes
    ButtonType = m_ButtonType
End Property
Public Property Let ButtonType(ByVal New_ButtonType As TopButtonTypes)
    If New_ButtonType <> m_ButtonType Then
        m_ButtonType = New_ButtonType
        PropertyChanged "ButtonType"
        LoadImage m_ButtonType
        Refresh
    End If
End Property
Public Property Get ColorSet() As TopColorSets
    ColorSet = m_ColorSet
End Property
Public Property Let ColorSet(ByVal New_ColorSet As TopColorSets)
    If m_ColorSet <> New_ColorSet Then
        m_ColorSet = New_ColorSet
        PropertyChanged "ColorSet"
        LoadImage m_ButtonType
        Refresh
    End If
End Property
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    If New_Enabled <> UserControl.Enabled Then
        UserControl.Enabled() = New_Enabled
        m_State = IIf(New_Enabled, TopNormal, TopDisabled)
        Refresh
    End If
End Property
Public Property Get Pic_Normal() As StdPicture
    Set Pic_Normal = TopImage(TopNormal)
End Property
Public Property Set Pic_Normal(ByVal newPic As StdPicture)
    Set TopImage(TopNormal) = newPic
    PropertyChanged "Pic_Normal"
    Refresh
End Property
Public Property Get Pic_Disabled() As StdPicture
    Set Pic_Disabled = TopImage(TopDisabled)
End Property
Public Property Set Pic_Disabled(ByVal newPic As StdPicture)
    Set TopImage(TopDisabled) = newPic
    PropertyChanged "Pic_Disabled"
    Refresh
End Property
Public Property Get Pic_Down() As StdPicture
    Set Pic_Down = TopImage(TopDown)
End Property
Public Property Set Pic_Down(ByVal newPic As StdPicture)
    Set TopImage(TopDown) = newPic
    PropertyChanged "Pic_Down"
    Refresh
End Property
Public Property Get Pic_Over() As StdPicture
    Set Pic_Over = TopImage(TopOver)
End Property
Public Property Set Pic_Over(ByVal newPic As StdPicture)
    Set TopImage(TopOver) = newPic
    PropertyChanged "Pic_Over"
    Refresh
End Property

Private Sub UserControl_Click()
    If m_ButtonType = TopMaxButton Then
        m_ButtonType = TopRestoreButton
        LoadImage m_ButtonType
        Refresh
    ElseIf m_ButtonType = TopRestoreButton Then
        m_ButtonType = TopMaxButton
        LoadImage m_ButtonType
        Refresh
    End If
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    m_State = TopDown
    Refresh
    Exit Sub
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_State = TopDown
    Refresh
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    SetCapture hwnd
    If PointInControl(X, UserControl.ScaleWidth, Y, UserControl.ScaleHeight) Then
        If m_State < TopOver Then
            If Button = vbLeftButton Then
                m_State = TopDown
            Else
                m_State = TopOver
            End If
            Refresh
        End If
    Else
        m_State = TopNormal
        Refresh
        RaiseEvent MouseOut
        ReleaseCapture
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        m_State = TopOver
        Refresh
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Sub UserControl_Paint()
    Refresh
End Sub

Public Sub UserControl_Resize()
    Refresh
End Sub

Public Sub Refresh()
    With UserControl
        .Height = 315
        .Width = 315
        .Cls
        If Not TopImage(m_State) Is Nothing Then .PaintPicture TopImage(m_State), 0, 0
    End With
End Sub
Private Sub UserControl_Terminate()
    Dim loop1 As Integer
    For loop1 = LBound(TopImage) To UBound(TopImage)
        Set TopImage(loop1) = Nothing
    Next
End Sub
