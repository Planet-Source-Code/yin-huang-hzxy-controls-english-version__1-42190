Attribute VB_Name = "modBas"
Option Explicit

Public Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Public Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Public Declare Function GetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hPal As Long, ByRef lpcolorref As Long)
Public Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Public Declare Function TrackMouseEvent Lib "comctl32.dll" Alias "_TrackMouseEvent" (ByRef lpEventTrack As TrackMouseEvent) As Long

Public Type RGBTRIPLE
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
End Type

Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBTRIPLE
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type Size
    cx As Long
    cy As Long
End Type

Public Type TrackMouseEvent
    cbSize As Long
    dwFlags As Long
    hwnd As Long
    dwHoverTime As Long
End Type

Enum CP
    PS_SOLID = 0
    PS_DASH = 1
    PS_DOT = 2
    PS_DASHDOT = 3
    PS_DASHDOTDOT = 4
    PS_NULL = 5
    PS_INSIDEFRAME = 6
End Enum

Enum CPR
    ALTERNATE = 1
    BDR_SUNKENINNER = &H8
    BDR_RAISEDOUTER = &H1
    BDR_RAISEDINNER = &H4
    BDR_SUNKENOUTER = &H2
    BF_LEFT = &H1
    BF_RIGHT = &H4
    BF_TOP = &H2
    BF_BOTTOM = &H8
    BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
End Enum

Enum DrawSt
    DST_COMPLEX = &H0
    DST_TEXT = &H1
    DST_PREFIXTEXT = &H2
    DST_ICON = &H3
    DST_BITMAP = &H4
    DSS_NORMAL = &H0
    DSS_UNION = &H10
    DSS_DISABLED = &H20
    DSS_MONO = &H80
    DSS_RIGHT = &H8000
End Enum

Enum DT
    DT_BOTTOM = &H8
    DT_CALCRECT = &H400
    DT_CENTER = &H1
    DT_CHARSTREAM = 4
    DT_DISPFILE = 6
    DT_EXPANDTABS = &H40
    DT_EXTERNALLEADING = &H200
    DT_INTERNAL = &H1000
    DT_LEFT = &H0
    DT_METAFILE = 5
    DT_NOCLIP = &H100
    DT_NOPREFIX = &H800
    DT_PLOTTER = 0
    DT_RASCAMERA = 3
    DT_RASDISPLAY = 1
    DT_RASPRINTER = 2
    DT_RIGHT = &H2
    DT_SINGLELINE = &H20
    DT_TABSTOP = &H80
    DT_TOP = &H0
    DT_VCENTER = &H4
    DT_WORDBREAK = &H10
    DT_WORD_ELLIPSIS = &H40000
    DT_END_ELLIPSIS = 32768
    DT_PATH_ELLIPSIS = &H4000
    DT_EDITCONTROL = &H2000
    '===================
    DT_INCENTER = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
End Enum

Enum Fnt
    FW_NORMAL = 400
    DEFAULT_CHARSET = 1
    OUT_DEFAULT_PRECIS = 0
    CLIP_DEFAULT_PRECIS = 0
    PROOF_QUALITY = 2
    DEFAULT_PITCH = 0
    LOGPIXELSY = 90
    COLOR_WINDOW = 5
End Enum

Enum IconDrawMe
    DI_MASK = &H1
    DI_IMAGE = &H2
    DI_NORMAL = DI_MASK Or DI_IMAGE
End Enum

Enum IconStates
    Icon_Normal = 0
    Icon_Grey = 1
    Icon_Disabled = 2
End Enum

Enum InterfaceColors
    icMistyRose = &HE1E4FF
    icSlateGray = &H908070
    icDodgerBlue = &HFF901E
    icDeepSkyBlue = &HFFBF00
    icSpringGreen = &H7FFF00
    icForestGreen = &H228B22
    icGoldenrod = &H20A5DA
    icFirebrick = &H2222B2
End Enum

Public Enum LabelState
    lblNormal = 0
    lblOver = 1
    lblActive = 2
    lblDisabeld = 3
End Enum

Enum OperaRGN
    RGN_AND = 1
    RGN_OR = 2
    RGN_XOR = 3
    RGN_DIFF = 4
    RGN_COPY = 5
    RGN_MAX = RGN_COPY
    RGN_MIN = RGN_AND
End Enum

Enum PhotoEffects
    vbSrcCopy = &HCC0020
    vbSrcAnd = &H8800C6
    vbSrcInvert = &H660046
    vbSrcErase = &H440328
    vbSrcPaint = &HEE0086
End Enum

Enum picScaleMe
    vbUser = 0
    vbTwips = 1
    vbPoints = 2
    vbPixels = 3
    vbCharacters = 4
    vbInches = 5
    vbMillimeters = 6
    vbCentimeters = 7
    vbHimetric = 8
    vbContainerPosition = 9
    vbContainerSize = 10
End Enum

Enum SND
    SND_SYNC = &H0
    SND_ASYNC = &H1
    SND_NODEFAULT = &H2
    SND_MEMORY = &H4&
    SND_LOOP = &H8
    SND_NOSTOP = &H10
    SND_NOWAIT = &H2000
    SND_FILENAME = &H20000
    SND_RESOURCE = &H40004
End Enum

Enum WM_Message
    TME_LEAVE = &H2
    WM_TIMER = &H113
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONDBLCLK = &H203
    WM_MYMSG = &H232
    WM_MOUSELEAVE = &H2A3
End Enum

Public Sub Ini()

End Sub

Public Sub TransBlt(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal srcPic As StdPicture, OriW As Long, OriH As Long, Optional ByVal IconState As IconStates = Icon_Normal, Optional ByVal ShadowColor As Long = -1)
    
    If DstW = 0 Or DstH = 0 Then Exit Sub
    
    Dim SrcDC As Long, SrcRect As RECT, SrcBmp As Long, SrcObj As Long
    Dim TmpDC As Long, TmpBmp As Long, TmpObj As Long
    Dim Data1() As RGBTRIPLE, Data2() As RGBTRIPLE
    Dim Info As BITMAPINFO, BrushRGB As RGBTRIPLE, gCol As Long
    Dim ToBeChange As Boolean
    Dim loopx As Long, loopy As Long
    Dim i As Long, iTop As Long, iLeft As Long
    Dim DisabledRGB As RGBTRIPLE, HighLightRGB As RGBTRIPLE, ShadowRGB As RGBTRIPLE
    Dim HaveChanged As Boolean

    Select Case IconState
    Case Icon_Normal
        Select Case srcPic.Type
        Case vbPicTypeBitmap
            SrcDC = CreateCompatibleDC(DstDC)
            SrcBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
            SrcObj = SelectObject(SrcDC, srcPic)
            
            StretchBlt DstDC, DstX, DstY, DstW, DstH, SrcDC, 0, 0, OriW, OriH, vbSrcCopy
            
'            DeleteObject SelectObject(SrcDC, SrcObj)
            DeleteObject SrcBmp
            DeleteDC SrcDC
        Case vbPicTypeIcon
            DrawIconEx DstDC, DstX, DstY, srcPic.Handle, DstW, DstH, 0, 0, DI_NORMAL
        End Select
    
    Case Icon_Disabled
        
        Const cShadow = &H808080
        Const cHighLight = &HFFFFFF
        
        Select Case srcPic.Type
        Case vbPicTypeBitmap
            DrawRectangle DstDC, DstX, DstY, DstW, DstH, cShadow
            Dim TmpRect As RECT
            TmpRect.Left = DstX
            TmpRect.Right = DstX + DstW
            TmpRect.Top = DstY
            TmpRect.Bottom = DstY + DstH
            DrawEdge DstDC, TmpRect, BDR_SUNKENINNER, BF_RECT
        Case vbPicTypeIcon
            SrcDC = CreateCompatibleDC(DstDC)
            SrcBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
            SrcObj = SelectObject(SrcDC, SrcBmp)
            BitBlt SrcDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
        
            TmpDC = CreateCompatibleDC(SrcDC)
            TmpBmp = CreateCompatibleBitmap(SrcDC, DstW, DstH)
            TmpObj = SelectObject(TmpDC, TmpBmp)
            BitBlt SrcDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
            DrawIconEx TmpDC, 0, 0, srcPic.Handle, DstW, DstH, 0, 0, DI_NORMAL
            
            ReDim Data1(DstW * DstH * 3 - 1)
            ReDim Data2(UBound(Data1))
            With Info.bmiHeader
                .biSize = Len(Info.bmiHeader)
                .biWidth = DstW
                .biHeight = DstH
                .biPlanes = 1
                .biBitCount = 24
            End With
    
            GetDIBits SrcDC, SrcBmp, 0, DstH, Data1(0), Info, 0
            GetDIBits TmpDC, TmpBmp, 0, DstH, Data2(0), Info, 0
            
            With DisabledRGB
                .rgbBlue = (cShadow \ &H10000) Mod &H100
                .rgbGreen = (cShadow \ &H100) Mod &H100
                .rgbRed = cShadow And &HFF
            End With
            
            With HighLightRGB
                .rgbBlue = (cHighLight \ &H10000) Mod &H100
                .rgbGreen = (cHighLight \ &H100) Mod &H100
                .rgbRed = cHighLight And &HFF
            End With
    
            For loopy = 0 To DstH - 1
                For loopx = DstW - 1 To 0 Step -1
                    i = loopy * DstW + loopx
                    If Data2(i).rgbRed = Data1(i).rgbRed And Data2(i).rgbGreen = Data1(i).rgbGreen And Data2(i).rgbBlue = Data1(i).rgbBlue Then '±³¾°É«
                        HaveChanged = False
                        If loopy < DstH - 1 Then
                            iTop = (loopy + 1) * DstW + loopx
                            If Data2(iTop).rgbRed <> Data1(iTop).rgbRed Or Data2(iTop).rgbGreen <> Data1(iTop).rgbGreen Or Data2(iTop).rgbBlue <> Data1(iTop).rgbBlue Then
                                HaveChanged = True
                                Data2(i) = HighLightRGB
                            End If
                        End If
                        If loopx > 0 And (Not HaveChanged) Then
                            iLeft = i - 1
                            If Data2(iLeft).rgbRed <> Data1(iLeft).rgbRed Or Data2(iLeft).rgbGreen <> Data1(iLeft).rgbGreen Or Data2(iLeft).rgbBlue <> Data1(iLeft).rgbBlue Then
                                Data2(i) = HighLightRGB
                            End If
                        End If
                    Else
                        Data2(i) = DisabledRGB
                    End If
                Next
            Next

            SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data2(0), Info, 0

            Erase Data1, Data2
            DeleteObject SelectObject(TmpDC, TmpObj)
            DeleteObject TmpBmp
            DeleteDC TmpDC
            DeleteObject SelectObject(SrcDC, SrcObj)
            DeleteObject SrcBmp
            DeleteDC SrcDC
        
        End Select
        
    Case Icon_Grey
        
        If ShadowColor <> -1 Then
            With ShadowRGB
                .rgbBlue = (cShadow \ &H10000) Mod &H100
                .rgbGreen = (cShadow \ &H100) Mod &H100
                .rgbRed = cShadow And &HFF
            End With
        End If
        
        Select Case srcPic.Type
        Case vbPicTypeBitmap
            SrcDC = CreateCompatibleDC(DstDC)
            SrcObj = SelectObject(SrcDC, srcPic)
            
            TmpDC = CreateCompatibleDC(SrcDC)
            TmpBmp = CreateCompatibleBitmap(SrcDC, DstW, DstH)
            TmpObj = SelectObject(TmpDC, TmpBmp)
            StretchBlt TmpDC, 0, 0, DstW, DstH, SrcDC, 0, 0, OriW, OriH, vbSrcCopy
        
            ReDim Data2(DstW * DstH * 3 - 1)
            With Info.bmiHeader
                .biSize = Len(Info.bmiHeader)
                .biWidth = DstW
                .biHeight = DstH
                .biPlanes = 1
                .biBitCount = 24
            End With
            
            GetDIBits TmpDC, TmpBmp, 0, DstH, Data2(0), Info, 0
        
            For loopy = 0 To DstH - 1
                For loopx = DstW - 1 To 0 Step -1
                    i = loopy * DstW + loopx
                    If ShadowColor <> -1 Then
                        Data2(i) = ShadowRGB
                    Else
                        With Data2(i)
                            gCol = CLng(.rgbRed * 0.3) + .rgbGreen * 0.59 + .rgbBlue * 0.11
                            .rgbRed = gCol
                            .rgbGreen = gCol
                            .rgbBlue = gCol
                        End With
                    End If
                Next
            Next
        
            SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data2(0), Info, 0
        
            Erase Data2
            DeleteObject SelectObject(TmpDC, TmpObj)
            DeleteObject TmpBmp
            DeleteDC TmpDC
'            DeleteObject SelectObject(SrcDC, SrcObj)
            DeleteDC SrcDC
        Case vbPicTypeIcon
            SrcDC = CreateCompatibleDC(DstDC)
            SrcBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
            SrcObj = SelectObject(SrcDC, SrcBmp)
            BitBlt SrcDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
        
            TmpDC = CreateCompatibleDC(SrcDC)
            TmpBmp = CreateCompatibleBitmap(SrcDC, DstW, DstH)
            TmpObj = SelectObject(TmpDC, TmpBmp)
            BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
            DrawIconEx TmpDC, 0, 0, srcPic.Handle, DstW, DstH, 0, 0, DI_NORMAL
            
            ReDim Data1(DstW * DstH * 3 - 1)
            ReDim Data2(UBound(Data1))
            With Info.bmiHeader
                .biSize = Len(Info.bmiHeader)
                .biWidth = DstW
                .biHeight = DstH
                .biPlanes = 1
                .biBitCount = 24
            End With
    
            GetDIBits SrcDC, SrcBmp, 0, DstH, Data1(0), Info, 0
            GetDIBits TmpDC, TmpBmp, 0, DstH, Data2(0), Info, 0
            
            For loopy = 0 To DstH - 1
                For loopx = DstW - 1 To 0 Step -1
                    i = loopy * DstW + loopx
                    If Data2(i).rgbRed <> Data1(i).rgbRed Or Data2(i).rgbGreen <> Data1(i).rgbGreen Or Data2(i).rgbBlue <> Data1(i).rgbBlue Then
                        If ShadowColor <> -1 Then
                            Data2(i) = ShadowRGB
                        Else
                            With Data2(i)
                                gCol = CLng(.rgbRed * 0.3) + .rgbGreen * 0.59 + .rgbBlue * 0.11
                                .rgbRed = gCol
                                .rgbGreen = gCol
                                .rgbBlue = gCol
                            End With
                        End If
                    End If
                Next
            Next
        
            SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data2(0), Info, 0
        
            Erase Data1, Data2
            DeleteObject SelectObject(TmpDC, TmpObj)
            DeleteObject TmpBmp
            DeleteDC TmpDC
            DeleteObject SelectObject(SrcDC, SrcObj)
            DeleteObject SrcBmp
            DeleteDC SrcDC
        End Select
    
    End Select

End Sub

Public Function PointInControl(X As Single, XMax As Single, Y As Single, YMax As Single) As Boolean
    If X >= 0 And X <= XMax And Y >= 0 And Y <= YMax Then
        PointInControl = True
    End If
End Function

Public Sub DrawLine(DstDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
    Dim TmpPoint As POINTAPI
    Dim oldPen As Long, hPen As Long

    hPen = CreatePen(PS_SOLID, 1, Color)
    oldPen = SelectObject(DstDC, hPen)
    
    MoveToEx DstDC, X1, Y1, TmpPoint
    LineTo DstDC, X2, Y2
    
    SelectObject DstDC, oldPen
    DeleteObject hPen

End Sub

Public Sub DrawRectangle(DstDC As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Color As Long, Optional OnlyBorder As Boolean = False)

    Dim bRECT As RECT
    Dim hBrush As Long

    bRECT.Left = X
    bRECT.Top = Y
    bRECT.Right = X + Width
    bRECT.Bottom = Y + Height

    hBrush = CreateSolidBrush(Color)

    If OnlyBorder Then
        FrameRect DstDC, bRECT, hBrush
    Else
        FillRect DstDC, bRECT, hBrush
    End If

    DeleteObject hBrush
End Sub

Public Function BreakApart(ByVal Color As Long) As Long
    Dim R As Integer, G As Integer, B As Integer
    R = getRedVal(Color)
    G = getGreenVal(Color)
    B = getBlueVal(Color)
    BreakApart = RGB(R, G, B)
End Function

Public Function getBlueVal(ByVal RGBCol As Long) As Integer
    RGBCol = Sys2RGB(RGBCol)
    If RGBCol < 0 Then RGBCol = 0
    getBlueVal = (RGBCol And &HFF0000) / &H10000
End Function

Public Function getGreenVal(ByVal RGBCol As Long) As Integer
    RGBCol = Sys2RGB(RGBCol)
    If RGBCol < 0 Then RGBCol = 0
    getGreenVal = ((RGBCol And &H100FF00) / &H100)
End Function

Public Function getRedVal(ByVal RGBCol As Long) As Integer
    RGBCol = Sys2RGB(RGBCol)
    If RGBCol < 0 Then RGBCol = 0
    getRedVal = RGBCol And &HFF
End Function

Public Function Sys2RGB(RGBCol As Long) As Long
    If RGBCol < 0 Then
        OleTranslateColor RGBCol, 0&, Sys2RGB
    Else
        Sys2RGB = RGBCol
    End If
End Function

Public Function ShiftColor(ByVal Color As Long, ByVal Value As Long, Optional isXP As Boolean = False) As Long
    Dim Red As Long, Blue As Long, Green As Long
    
    If Not isXP Then 'for XP button i use a work-aroud that works fine
        Value = Value \ 2 'this is just a tricky way to do it and will result in weird colors for WinXP and KDE2
        Blue = ((Color \ &H10000) Mod &H100) + Value
    Else
        Blue = ((Color \ &H10000) Mod &H100)
        Blue = Blue + ((Blue * Value) \ &HC0)
    End If
    Green = ((Color \ &H100) Mod &H100) + Value
    Red = (Color And &HFF) + Value
    
    If Value > 0 Then
        If Red > 255 Then Red = 255
        If Green > 255 Then Green = 255
        If Blue > 255 Then Blue = 255
    ElseIf Value < 0 Then
        If Red < 0 Then Red = 0
        If Green < 0 Then Green = 0
        If Blue < 0 Then Blue = 0
    End If
    
    ShiftColor = Red + 256& * Green + 65536 * Blue
End Function

Public Function SmoothColor(ByVal Color As Long) As Long
    Dim RColor As Integer, GColor As Integer, BColor As Integer
    RColor = getRedVal(Color)
    GColor = getGreenVal(Color)
    BColor = getBlueVal(Color)
    RColor = RColor + 76 - Int((RColor + 32) / 64) * 19
    GColor = GColor + 76 - Int((GColor + 32) / 64) * 19
    BColor = BColor + 76 - Int((BColor + 32) / 64) * 19
    SmoothColor = BreakApart(RGB(RColor, GColor, BColor))
End Function

Public Function PlayASound(SoundFile As String) As Boolean
    If Trim(SoundFile) <> "" Then
        Dim bArr() As Byte
        bArr = LoadResData(SoundFile, "SOUND")
        sndPlaySound bArr(0), SND_MEMORY + SND_ASYNC + SND_NOSTOP + SND_NOWAIT + SND_NODEFAULT
    End If
'    PlayASound = PlaySound(SoundFile, vbNull, SND_FILENAME _
'    + SND_ASYNC + SND_NOSTOP + SND_NOWAIT + SND_NODEFAULT)
End Function
