VERSION 5.00
Object = "*\A..\HZXYCO~1\HzxYControlsEnglish.vbp"
Begin VB.Form Form1 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   650
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin HzxYControlsEnglish.HzxYTabStrip HzxYTabStrip1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8705
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CountTabs       =   2
      CurrentTab      =   1
      CurrentTab_Caption0=   "HzxYTabStrip"
      CurrentTab_Image0=   "Form1.frx":0000
      CurrentTab_Caption1=   "Other Controls"
      Begin HzxYControlsEnglish.HzxYFrame fra 
         Height          =   3975
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7011
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16710908
         BorderStyle     =   0
         Begin HzxYControlsEnglish.HzxYFrame HzxYFrame9 
            Height          =   1215
            Left            =   4680
            TabIndex        =   28
            Top             =   120
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   2143
            Caption         =   "HzxYTopButtons"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16710908
            Begin HzxYControlsEnglish.HzxYXPButton HzxYXPButton1 
               Height          =   375
               Left            =   1560
               TabIndex        =   30
               Top             =   240
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               Caption         =   "Next TopButtonType"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonCaptionPercent=   100
               PictureAreaBackSmoothColor=   13882323
            End
            Begin HzxYControlsEnglish.HzxYCheckBox HzxYCheckBox9 
               Height          =   240
               Left            =   240
               TabIndex        =   29
               Top             =   840
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   423
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               Caption         =   "Enable"
               BackColor       =   16710908
               Pic_UncheckedNormal=   "Form1.frx":0324
               Pic_CheckedNormal=   "Form1.frx":0676
               Pic_MixedNormal =   "Form1.frx":09C8
               Pic_UncheckedDisabled=   "Form1.frx":0D1A
               Pic_CheckedDisabled=   "Form1.frx":106C
               Pic_MixedDisabled=   "Form1.frx":13BE
               Pic_UncheckedOver=   "Form1.frx":1710
               Pic_CheckedOver =   "Form1.frx":1A62
               Pic_MixedOver   =   "Form1.frx":1DB4
               Pic_UncheckedDown=   "Form1.frx":2106
               Pic_CheckedDown =   "Form1.frx":2458
               Pic_MixedDown   =   "Form1.frx":27AA
            End
            Begin HzxYControlsEnglish.HzxYTopButtons HzxYTopButtons1 
               Height          =   315
               Left            =   240
               Top             =   360
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   556
               ButtonType      =   1
               Pic_Normal      =   "Form1.frx":2AFC
               Pic_Over        =   "Form1.frx":308E
               Pic_Down        =   "Form1.frx":3620
               Pic_Disabled    =   "Form1.frx":3BB2
            End
            Begin HzxYControlsEnglish.HzxYXPButton HzxYXPButton2 
               Height          =   375
               Left            =   1560
               TabIndex        =   31
               Top             =   720
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   661
               Caption         =   "Next ColorSet"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonCaptionPercent=   100
               PictureAreaBackSmoothColor=   13882323
            End
         End
         Begin HzxYControlsEnglish.HzxYFrame HzxYFrame5 
            Height          =   2295
            Left            =   4200
            TabIndex        =   15
            Top             =   1440
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   4048
            Caption         =   "HzxYFrame"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16710908
            Begin HzxYControlsEnglish.HzxYCheckBox HzxYCheckBox8 
               Height          =   240
               Left            =   1320
               TabIndex        =   21
               Top             =   1560
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   423
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               Caption         =   "Caption"
               BackColor       =   16710908
               Pic_UncheckedNormal=   "Form1.frx":4144
               Pic_CheckedNormal=   "Form1.frx":4496
               Pic_MixedNormal =   "Form1.frx":47E8
               Pic_UncheckedDisabled=   "Form1.frx":4B3A
               Pic_CheckedDisabled=   "Form1.frx":4E8C
               Pic_MixedDisabled=   "Form1.frx":51DE
               Pic_UncheckedOver=   "Form1.frx":5530
               Pic_CheckedOver =   "Form1.frx":5882
               Pic_MixedOver   =   "Form1.frx":5BD4
               Pic_UncheckedDown=   "Form1.frx":5F26
               Pic_CheckedDown =   "Form1.frx":6278
               Pic_MixedDown   =   "Form1.frx":65CA
            End
            Begin HzxYControlsEnglish.HzxYCheckBox HzxYCheckBox7 
               Height          =   240
               Left            =   240
               TabIndex        =   20
               Top             =   1560
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   423
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Image"
               BackColor       =   16710908
               Pic_UncheckedNormal=   "Form1.frx":691C
               Pic_CheckedNormal=   "Form1.frx":6C6E
               Pic_MixedNormal =   "Form1.frx":6FC0
               Pic_UncheckedDisabled=   "Form1.frx":7312
               Pic_CheckedDisabled=   "Form1.frx":7664
               Pic_MixedDisabled=   "Form1.frx":79B6
               Pic_UncheckedOver=   "Form1.frx":7D08
               Pic_CheckedOver =   "Form1.frx":805A
               Pic_MixedOver   =   "Form1.frx":83AC
               Pic_UncheckedDown=   "Form1.frx":86FE
               Pic_CheckedDown =   "Form1.frx":8A50
               Pic_MixedDown   =   "Form1.frx":8DA2
            End
            Begin HzxYControlsEnglish.HzxYCheckBox HzxYCheckBox6 
               Height          =   240
               Left            =   1320
               TabIndex        =   19
               Top             =   1920
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   423
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               Caption         =   "Control Contained Controls"
               BackColor       =   16710908
               Pic_UncheckedNormal=   "Form1.frx":90F4
               Pic_CheckedNormal=   "Form1.frx":9446
               Pic_MixedNormal =   "Form1.frx":9798
               Pic_UncheckedDisabled=   "Form1.frx":9AEA
               Pic_CheckedDisabled=   "Form1.frx":9E3C
               Pic_MixedDisabled=   "Form1.frx":A18E
               Pic_UncheckedOver=   "Form1.frx":A4E0
               Pic_CheckedOver =   "Form1.frx":A832
               Pic_MixedOver   =   "Form1.frx":AB84
               Pic_UncheckedDown=   "Form1.frx":AED6
               Pic_CheckedDown =   "Form1.frx":B228
               Pic_MixedDown   =   "Form1.frx":B57A
            End
            Begin HzxYControlsEnglish.HzxYCheckBox HzxYCheckBox5 
               Height          =   240
               Left            =   2640
               TabIndex        =   18
               Top             =   1560
               Width           =   810
               _ExtentX        =   1429
               _ExtentY        =   423
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               Caption         =   "Border"
               BackColor       =   16710908
               Pic_UncheckedNormal=   "Form1.frx":B8CC
               Pic_CheckedNormal=   "Form1.frx":BC1E
               Pic_MixedNormal =   "Form1.frx":BF70
               Pic_UncheckedDisabled=   "Form1.frx":C2C2
               Pic_CheckedDisabled=   "Form1.frx":C614
               Pic_MixedDisabled=   "Form1.frx":C966
               Pic_UncheckedOver=   "Form1.frx":CCB8
               Pic_CheckedOver =   "Form1.frx":D00A
               Pic_MixedOver   =   "Form1.frx":D35C
               Pic_UncheckedDown=   "Form1.frx":D6AE
               Pic_CheckedDown =   "Form1.frx":DA00
               Pic_MixedDown   =   "Form1.frx":DD52
            End
            Begin HzxYControlsEnglish.HzxYFrame HzxYFrame6 
               Height          =   1215
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   2143
               Caption         =   "fra"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16710908
               Begin HzxYControlsEnglish.HzxYFrame HzxYFrame7 
                  Height          =   615
                  Left            =   120
                  TabIndex        =   22
                  Top             =   480
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   1085
                  Caption         =   "fra's Image Width"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Times New Roman"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   16710908
                  Begin HzxYControlsEnglish.HzxYOption FraW 
                     Height          =   225
                     Index           =   1
                     Left            =   1080
                     TabIndex        =   23
                     Top             =   240
                     Width           =   435
                     _ExtentX        =   767
                     _ExtentY        =   397
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Times New Roman"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Caption         =   "24"
                     BackColor       =   16710908
                  End
                  Begin HzxYControlsEnglish.HzxYOption FraW 
                     Height          =   225
                     Index           =   0
                     Left            =   240
                     TabIndex        =   24
                     Top             =   240
                     Width           =   435
                     _ExtentX        =   767
                     _ExtentY        =   397
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Times New Roman"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Value           =   -1  'True
                     Caption         =   "16"
                     BackColor       =   16710908
                  End
               End
               Begin HzxYControlsEnglish.HzxYFrame HzxYFrame8 
                  Height          =   615
                  Left            =   2040
                  TabIndex        =   25
                  Top             =   480
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   1085
                  Caption         =   "fra's Image Height"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Times New Roman"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   16710908
                  Begin HzxYControlsEnglish.HzxYOption FraH 
                     Height          =   225
                     Index           =   1
                     Left            =   1080
                     TabIndex        =   26
                     Top             =   240
                     Width           =   435
                     _ExtentX        =   767
                     _ExtentY        =   397
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Times New Roman"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Caption         =   "24"
                     BackColor       =   16710908
                  End
                  Begin HzxYControlsEnglish.HzxYOption FraH 
                     Height          =   225
                     Index           =   0
                     Left            =   240
                     TabIndex        =   27
                     Top             =   240
                     Width           =   435
                     _ExtentX        =   767
                     _ExtentY        =   397
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Times New Roman"
                        Size            =   9
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Value           =   -1  'True
                     Caption         =   "16"
                     BackColor       =   16710908
                  End
               End
            End
            Begin HzxYControlsEnglish.HzxYCheckBox HzxYCheckBox4 
               Height          =   240
               Left            =   240
               TabIndex        =   16
               Top             =   1920
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   423
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Value           =   1
               Caption         =   "Enable"
               BackColor       =   16710908
               Pic_UncheckedNormal=   "Form1.frx":E0A4
               Pic_CheckedNormal=   "Form1.frx":E3F6
               Pic_MixedNormal =   "Form1.frx":E748
               Pic_UncheckedDisabled=   "Form1.frx":EA9A
               Pic_CheckedDisabled=   "Form1.frx":EDEC
               Pic_MixedDisabled=   "Form1.frx":F13E
               Pic_UncheckedOver=   "Form1.frx":F490
               Pic_CheckedOver =   "Form1.frx":F7E2
               Pic_MixedOver   =   "Form1.frx":FB34
               Pic_UncheckedDown=   "Form1.frx":FE86
               Pic_CheckedDown =   "Form1.frx":101D8
               Pic_MixedDown   =   "Form1.frx":1052A
            End
         End
         Begin HzxYControlsEnglish.HzxYFrame HzxYFrame2 
            Height          =   2295
            Left            =   360
            TabIndex        =   2
            Top             =   1440
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   4048
            Caption         =   "HzxYOption"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16710908
            Begin HzxYControlsEnglish.HzxYFrame HzxYFrame4 
               Height          =   735
               Left            =   240
               TabIndex        =   3
               Top             =   1320
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   1296
               Caption         =   "Enable/Disable Op(0)"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16710908
               Begin HzxYControlsEnglish.HzxYOption En 
                  Height          =   225
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   4
                  Top             =   360
                  Width           =   795
                  _ExtentX        =   1402
                  _ExtentY        =   397
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Times New Roman"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Caption         =   "Disable"
                  BackColor       =   16710908
               End
               Begin HzxYControlsEnglish.HzxYOption En 
                  Height          =   225
                  Index           =   0
                  Left            =   240
                  TabIndex        =   5
                  Top             =   360
                  Width           =   735
                  _ExtentX        =   1296
                  _ExtentY        =   397
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Times New Roman"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Value           =   -1  'True
                  Caption         =   "Enable"
                  BackColor       =   16710908
               End
            End
            Begin HzxYControlsEnglish.HzxYFrame HzxYFrame3 
               Height          =   855
               Left            =   1560
               TabIndex        =   6
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   1508
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16710908
               Begin HzxYControlsEnglish.HzxYOption Op 
                  Height          =   225
                  Index           =   2
                  Left            =   240
                  TabIndex        =   7
                  Top             =   120
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   397
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Times New Roman"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   16710908
               End
               Begin HzxYControlsEnglish.HzxYOption Op 
                  Height          =   225
                  Index           =   3
                  Left            =   240
                  TabIndex        =   8
                  Top             =   480
                  Width           =   705
                  _ExtentX        =   1244
                  _ExtentY        =   397
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Times New Roman"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackColor       =   16710908
               End
            End
            Begin HzxYControlsEnglish.HzxYOption Op 
               Height          =   225
               Index           =   0
               Left            =   360
               TabIndex        =   9
               Top             =   360
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   397
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16710908
            End
            Begin HzxYControlsEnglish.HzxYOption Op 
               Height          =   225
               Index           =   1
               Left            =   360
               TabIndex        =   10
               Top             =   720
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   397
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16710908
            End
         End
         Begin HzxYControlsEnglish.HzxYFrame HzxYFrame1 
            Height          =   1215
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   2143
            Caption         =   "CheckBox"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16710908
            Begin HzxYControlsEnglish.HzxYCheckBox HzxYCheckBox3 
               Height          =   240
               Left            =   360
               TabIndex        =   12
               Top             =   840
               Width           =   2085
               _ExtentX        =   3678
               _ExtentY        =   423
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Caption         =   "Enable HzxYCheckBox1"
               BackColor       =   16710908
               Pic_UncheckedNormal=   "Form1.frx":1087C
               Pic_CheckedNormal=   "Form1.frx":10BCE
               Pic_MixedNormal =   "Form1.frx":10F20
               Pic_UncheckedDisabled=   "Form1.frx":11272
               Pic_CheckedDisabled=   "Form1.frx":115C4
               Pic_MixedDisabled=   "Form1.frx":11916
               Pic_UncheckedOver=   "Form1.frx":11C68
               Pic_CheckedOver =   "Form1.frx":11FBA
               Pic_MixedOver   =   "Form1.frx":1230C
               Pic_UncheckedDown=   "Form1.frx":1265E
               Pic_CheckedDown =   "Form1.frx":129B0
               Pic_MixedDown   =   "Form1.frx":12D02
            End
            Begin HzxYControlsEnglish.HzxYCheckBox HzxYCheckBox2 
               Height          =   240
               Left            =   2400
               TabIndex        =   13
               Top             =   360
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   423
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   16710908
               Pic_UncheckedNormal=   "Form1.frx":13054
               Pic_CheckedNormal=   "Form1.frx":133A6
               Pic_MixedNormal =   "Form1.frx":136F8
               Pic_UncheckedDisabled=   "Form1.frx":13A4A
               Pic_CheckedDisabled=   "Form1.frx":13D9C
               Pic_MixedDisabled=   "Form1.frx":140EE
               Pic_UncheckedOver=   "Form1.frx":14440
               Pic_CheckedOver =   "Form1.frx":14792
               Pic_MixedOver   =   "Form1.frx":14AE4
               Pic_UncheckedDown=   "Form1.frx":14E36
               Pic_CheckedDown =   "Form1.frx":15188
               Pic_MixedDown   =   "Form1.frx":154DA
            End
            Begin HzxYControlsEnglish.HzxYCheckBox HzxYCheckBox1 
               Height          =   240
               Left            =   360
               TabIndex        =   14
               Top             =   360
               Width           =   1560
               _ExtentX        =   2752
               _ExtentY        =   423
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               BackColor       =   16710908
               Pic_UncheckedNormal=   "Form1.frx":1582C
               Pic_CheckedNormal=   "Form1.frx":15B7E
               Pic_MixedNormal =   "Form1.frx":15ED0
               Pic_UncheckedDisabled=   "Form1.frx":16222
               Pic_CheckedDisabled=   "Form1.frx":16574
               Pic_MixedDisabled=   "Form1.frx":168C6
               Pic_UncheckedOver=   "Form1.frx":16C18
               Pic_CheckedOver =   "Form1.frx":16F6A
               Pic_MixedOver   =   "Form1.frx":172BC
               Pic_UncheckedDown=   "Form1.frx":1760E
               Pic_CheckedDown =   "Form1.frx":17960
               Pic_MixedDown   =   "Form1.frx":17CB2
            End
         End
      End
      Begin HzxYControlsEnglish.HzxYFrame fra 
         Height          =   3975
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   7011
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16710908
         BorderStyle     =   0
         Begin VB.TextBox Text1 
            BackColor       =   &H00FEFCFC&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3015
            Left            =   720
            MultiLine       =   -1  'True
            TabIndex        =   33
            Text            =   "Form1.frx":18004
            Top             =   720
            Width           =   7215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "How to Use"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00D54600&
            Height          =   495
            Left            =   3120
            TabIndex        =   34
            Top             =   120
            Width           =   2025
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub En_Click(Index As Integer)
    Op(0).Enabled = IIf(Index = 0, True, False)
End Sub

Private Sub FraH_Click(Index As Integer)
    HzxYFrame6.ImageHeight = IIf(Index = 0, 16, 24)
End Sub

Private Sub FraW_Click(Index As Integer)
    HzxYFrame6.ImageWidth = IIf(Index = 0, 16, 24)
End Sub

Private Sub HzxYCheckBox3_Click()
    HzxYCheckBox1.Enabled = (HzxYCheckBox3.Value = Checked)
End Sub

Private Sub HzxYCheckBox4_Click()
    HzxYFrame6.Enabled = (HzxYCheckBox4.Value = Checked)
End Sub

Private Sub HzxYCheckBox5_Click()
    HzxYFrame6.BorderStyle = IIf(HzxYCheckBox5.Value = Checked, fraFixed_Single, fraNone)
End Sub

Private Sub HzxYCheckBox6_Click()
    HzxYFrame6.ControlContainedControls = (HzxYCheckBox6.Value = Checked)
End Sub

Private Sub HzxYCheckBox7_Click()
    Set HzxYFrame6.Image = IIf(HzxYCheckBox7.Value = Checked, LoadPicture("Run.ico"), Nothing)
End Sub

Private Sub HzxYCheckBox8_Click()
    HzxYFrame6.Caption = IIf(HzxYCheckBox8.Value = Checked, "fra", "")
End Sub

Private Sub HzxYCheckBox9_Click()
    HzxYTopButtons1.Enabled = (HzxYCheckBox9.Value = Checked)
End Sub

Private Sub HzxYTabStrip1_TabClick(NewTabIndex As Integer, OldTabIndex As Integer)
    fra(NewTabIndex).Visible = True
    fra(OldTabIndex).Visible = False
End Sub

Private Sub HzxYXPButton1_Click()
    HzxYTopButtons1.ButtonType = (HzxYTopButtons1.ButtonType + 1) Mod 5
End Sub

Private Sub HzxYXPButton2_Click()
    HzxYTopButtons1.ColorSet = (HzxYTopButtons1.ColorSet + 1) Mod 3
End Sub
