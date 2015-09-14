VERSION 5.00
Begin VB.Form frmPocketDMX 
   Appearance      =   0  '2D
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "PocketDMX"
   ClientHeight    =   9120
   ClientLeft      =   1635
   ClientTop       =   1545
   ClientWidth     =   10440
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmPocketDMX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmPocketDMX.frx":0AE2
   ScaleHeight     =   9120
   ScaleWidth      =   10440
   Begin VB.Timer tmrTimer 
      Interval        =   100
      Left            =   9000
      Top             =   2550
   End
   Begin VB.CheckBox chkAutoSpeed 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      Caption         =   "auto"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5175
      TabIndex        =   15
      Top             =   8400
      Width           =   690
   End
   Begin PocketDMX.ValueBar vbrTilt 
      Height          =   3915
      Left            =   300
      TabIndex        =   7
      Top             =   4350
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   6906
      Orientation     =   1
      Max             =   255
      Value           =   0
   End
   Begin VB.Frame fraPTBox 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3930
      Left            =   600
      TabIndex        =   2
      Top             =   4350
      Width           =   3930
      Begin VB.PictureBox picPTBox 
         Appearance      =   0  '2D
         BackColor       =   &H00000000&
         FillColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3870
         Left            =   30
         MousePointer    =   2  'Kreuz
         ScaleHeight     =   3840
         ScaleWidth      =   3840
         TabIndex        =   3
         Top             =   30
         Width           =   3870
         Begin VB.Image imgCross 
            Height          =   405
            Left            =   885
            MousePointer    =   4  'Symbol
            Picture         =   "frmPocketDMX.frx":0DEC
            Top             =   1035
            Width           =   405
         End
      End
   End
   Begin VB.Frame fraColor 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   300
      TabIndex        =   1
      Top             =   2150
      Width           =   7665
      Begin VB.Shape shpBlack 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Ausgefüllt
         Height          =   540
         Index           =   2
         Left            =   4605
         Top             =   1005
         Width           =   33150
      End
      Begin VB.Line linSep 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   2
         X1              =   4650
         X2              =   0
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   36
         Left            =   3585
         Picture         =   "frmPocketDMX.frx":0E57
         Tag             =   "2-191"
         Top             =   1050
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   35
         Left            =   3075
         Picture         =   "frmPocketDMX.frx":165B
         Tag             =   "2-170"
         Top             =   1050
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   34
         Left            =   2565
         Picture         =   "frmPocketDMX.frx":1E5F
         Tag             =   "2-149"
         Top             =   1050
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   33
         Left            =   2055
         Picture         =   "frmPocketDMX.frx":2663
         Tag             =   "2-128"
         Top             =   1050
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   29
         Left            =   15
         Picture         =   "frmPocketDMX.frx":2E67
         Tag             =   "2-192"
         Top             =   1050
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   30
         Left            =   525
         Picture         =   "frmPocketDMX.frx":366B
         Tag             =   "2-222"
         Top             =   1050
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   31
         Left            =   1035
         Picture         =   "frmPocketDMX.frx":3E6F
         Tag             =   "2-243"
         Top             =   1050
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   32
         Left            =   1545
         Picture         =   "frmPocketDMX.frx":4673
         Tag             =   "2-253"
         Top             =   1050
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   37
         Left            =   4095
         Picture         =   "frmPocketDMX.frx":4E77
         Tag             =   "2-255"
         Top             =   1050
         Width           =   495
      End
      Begin VB.Shape shpColor 
         BorderColor     =   &H00C0C0C0&
         Height          =   510
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Shape shpBlack 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Ausgefüllt
         Height          =   495
         Index           =   1
         Left            =   7155
         Top             =   510
         Width           =   525
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   14
         Left            =   7155
         Picture         =   "frmPocketDMX.frx":567B
         Tag             =   "2-112"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   28
         Left            =   6645
         Picture         =   "frmPocketDMX.frx":61FF
         Tag             =   "2-108"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   13
         Left            =   6645
         Picture         =   "frmPocketDMX.frx":6D83
         Tag             =   "2-104"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   27
         Left            =   6135
         Picture         =   "frmPocketDMX.frx":7907
         Tag             =   "2-100"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   12
         Left            =   6135
         Picture         =   "frmPocketDMX.frx":848B
         Tag             =   "2-96"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   26
         Left            =   5625
         Picture         =   "frmPocketDMX.frx":900F
         Tag             =   "2-92"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   0
         Left            =   15
         Picture         =   "frmPocketDMX.frx":9B93
         Tag             =   "2-0"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   1
         Left            =   525
         Picture         =   "frmPocketDMX.frx":A717
         Tag             =   "2-8"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   2
         Left            =   1035
         Picture         =   "frmPocketDMX.frx":B29B
         Tag             =   "2-16"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   3
         Left            =   1545
         Picture         =   "frmPocketDMX.frx":BE1F
         Tag             =   "2-24"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   4
         Left            =   2055
         Picture         =   "frmPocketDMX.frx":C9A3
         Tag             =   "2-32"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   5
         Left            =   2565
         Picture         =   "frmPocketDMX.frx":D527
         Tag             =   "2-40"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   6
         Left            =   3075
         Picture         =   "frmPocketDMX.frx":E0AB
         Tag             =   "2-48"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   15
         Left            =   15
         Picture         =   "frmPocketDMX.frx":EC2F
         Tag             =   "2-4"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   16
         Left            =   525
         Picture         =   "frmPocketDMX.frx":F7B3
         Tag             =   "2-12"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   17
         Left            =   1035
         Picture         =   "frmPocketDMX.frx":10337
         Tag             =   "2-20"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   18
         Left            =   1545
         Picture         =   "frmPocketDMX.frx":10EBB
         Tag             =   "2-28"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   19
         Left            =   2055
         Picture         =   "frmPocketDMX.frx":11A3F
         Tag             =   "2-36"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   20
         Left            =   2565
         Picture         =   "frmPocketDMX.frx":125C3
         Tag             =   "2-44"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   21
         Left            =   3075
         Picture         =   "frmPocketDMX.frx":13147
         Tag             =   "2-52"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   7
         Left            =   3585
         Picture         =   "frmPocketDMX.frx":13CCB
         Tag             =   "2-56"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   8
         Left            =   4095
         Picture         =   "frmPocketDMX.frx":1484F
         Tag             =   "2-64"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   9
         Left            =   4605
         Picture         =   "frmPocketDMX.frx":153D3
         Tag             =   "2-72"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   10
         Left            =   5115
         Picture         =   "frmPocketDMX.frx":15F57
         Tag             =   "2-80"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   25
         Left            =   5115
         Picture         =   "frmPocketDMX.frx":16ADB
         Tag             =   "2-84"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   24
         Left            =   4605
         Picture         =   "frmPocketDMX.frx":1765F
         Tag             =   "2-76"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   23
         Left            =   4095
         Picture         =   "frmPocketDMX.frx":181E3
         Tag             =   "2-68"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   22
         Left            =   3585
         Picture         =   "frmPocketDMX.frx":18D67
         Tag             =   "2-60"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgColor 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   11
         Left            =   5625
         Picture         =   "frmPocketDMX.frx":198EB
         Tag             =   "2-88"
         Top             =   15
         Width           =   495
      End
   End
   Begin VB.Frame fraGobo 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   300
      TabIndex        =   0
      Top             =   570
      Width           =   6225
      Begin VB.Shape shpGobo 
         BorderColor     =   &H00808080&
         Height          =   510
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Shape shpBlack 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H00000000&
         FillStyle       =   0  'Ausgefüllt
         Height          =   510
         Index           =   0
         Left            =   5700
         Top             =   510
         Width           =   525
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   22
         Left            =   5715
         Picture         =   "frmPocketDMX.frx":1A46F
         Tag             =   "3-255"
         Top             =   15
         Width           =   495
      End
      Begin VB.Line linSep 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   1
         X1              =   5685
         X2              =   5685
         Y1              =   1010
         Y2              =   15
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   21
         Left            =   3630
         Picture         =   "frmPocketDMX.frx":1AC73
         Tag             =   "3-128"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   20
         Left            =   4140
         Picture         =   "frmPocketDMX.frx":1B477
         Tag             =   "3-149"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   19
         Left            =   4650
         Picture         =   "frmPocketDMX.frx":1BC7B
         Tag             =   "3-170"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   18
         Left            =   5160
         Picture         =   "frmPocketDMX.frx":1C47F
         Tag             =   "3-191"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   17
         Left            =   5160
         Picture         =   "frmPocketDMX.frx":1CC83
         Tag             =   "3-253"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   16
         Left            =   4650
         Picture         =   "frmPocketDMX.frx":1D487
         Tag             =   "3-243"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   15
         Left            =   4140
         Picture         =   "frmPocketDMX.frx":1DC8B
         Tag             =   "3-222"
         Top             =   15
         Width           =   495
      End
      Begin VB.Line linSep 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   0
         X1              =   3600
         X2              =   3600
         Y1              =   995
         Y2              =   15
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   14
         Left            =   3630
         Picture         =   "frmPocketDMX.frx":1E48F
         Tag             =   "3-192"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   13
         Left            =   3075
         Picture         =   "frmPocketDMX.frx":1EC93
         Tag             =   "3-104"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   12
         Left            =   2565
         Picture         =   "frmPocketDMX.frx":1F497
         Tag             =   "3-96"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   11
         Left            =   2055
         Picture         =   "frmPocketDMX.frx":1FC9B
         Tag             =   "3-88"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   10
         Left            =   1545
         Picture         =   "frmPocketDMX.frx":2049F
         Tag             =   "3-80"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   9
         Left            =   1035
         Picture         =   "frmPocketDMX.frx":20CA3
         Tag             =   "3-72"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   8
         Left            =   525
         Picture         =   "frmPocketDMX.frx":214A7
         Tag             =   "3-64"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   7
         Left            =   15
         Picture         =   "frmPocketDMX.frx":21CAB
         Tag             =   "3-56"
         Top             =   510
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   6
         Left            =   3075
         Picture         =   "frmPocketDMX.frx":224AF
         Tag             =   "3-48"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   5
         Left            =   2565
         Picture         =   "frmPocketDMX.frx":22CB3
         Tag             =   "3-40"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   4
         Left            =   2055
         Picture         =   "frmPocketDMX.frx":234B7
         Tag             =   "3-32"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   3
         Left            =   1545
         Picture         =   "frmPocketDMX.frx":23CB9
         Tag             =   "3-24"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   2
         Left            =   1035
         Picture         =   "frmPocketDMX.frx":244BD
         Tag             =   "3-16"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   1
         Left            =   525
         Picture         =   "frmPocketDMX.frx":25041
         Tag             =   "3-8"
         Top             =   15
         Width           =   495
      End
      Begin VB.Image imgGobo 
         Appearance      =   0  '2D
         BorderStyle     =   1  'Fest Einfach
         Height          =   480
         Index           =   0
         Left            =   15
         Picture         =   "frmPocketDMX.frx":25755
         Tag             =   "3-0"
         Top             =   15
         Width           =   495
      End
   End
   Begin PocketDMX.ValueBar vbrPan 
      Height          =   240
      Left            =   600
      TabIndex        =   8
      Top             =   8340
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   423
      Orientation     =   0
      Max             =   255
      Value           =   0
   End
   Begin PocketDMX.ValueBar vbrSpeed 
      Height          =   3915
      Left            =   5175
      TabIndex        =   9
      Top             =   4350
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   6906
      Orientation     =   1
      Max             =   238
      Value           =   0
   End
   Begin PocketDMX.ValueBar vbrShutter 
      Height          =   3915
      Left            =   6450
      TabIndex        =   11
      Top             =   4350
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   6906
      Orientation     =   1
      Max             =   223
      Value           =   0
   End
   Begin PocketDMX.ValueBar vbrLaser 
      Height          =   3915
      Left            =   7800
      TabIndex        =   13
      Top             =   4350
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   6906
      Orientation     =   1
      Max             =   111
      Value           =   0
   End
   Begin VB.Label lblFunc 
      BackColor       =   &H00000000&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   9150
      TabIndex        =   16
      Top             =   4050
      Width           =   840
   End
   Begin VB.Image imgReset 
      Height          =   525
      Left            =   9150
      Picture         =   "frmPocketDMX.frx":25F59
      ToolTipText     =   "Press 5 sec to reset"
      Top             =   4350
      Width           =   525
   End
   Begin VB.Image imgControl 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   480
      Index           =   5
      Left            =   5475
      Picture         =   "frmPocketDMX.frx":26E5F
      Tag             =   "3-255"
      ToolTipText     =   "Relative Speed"
      Top             =   7800
      Width           =   480
   End
   Begin VB.Image imgControl 
      Height          =   450
      Index           =   4
      Left            =   6750
      Picture         =   "frmPocketDMX.frx":27969
      ToolTipText     =   "Shutter Closed"
      Top             =   7800
      Width           =   525
   End
   Begin VB.Image imgControl 
      Height          =   465
      Index           =   1
      Left            =   6750
      Picture         =   "frmPocketDMX.frx":28653
      ToolTipText     =   "Shutter Open"
      Top             =   4350
      Width           =   525
   End
   Begin VB.Image imgControl 
      Height          =   450
      Index           =   3
      Left            =   8100
      Picture         =   "frmPocketDMX.frx":293A9
      ToolTipText     =   "Laser Off"
      Top             =   7800
      Width           =   525
   End
   Begin VB.Image imgControl 
      Height          =   450
      Index           =   2
      Left            =   8100
      Picture         =   "frmPocketDMX.frx":2A093
      ToolTipText     =   "Laser On"
      Top             =   4350
      Width           =   525
   End
   Begin VB.Label lblFunc 
      BackColor       =   &H00000000&
      Caption         =   "Laser"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   7800
      TabIndex        =   14
      Top             =   4050
      Width           =   840
   End
   Begin VB.Label lblFunc 
      BackColor       =   &H00000000&
      Caption         =   "Shutter"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   6450
      TabIndex        =   12
      Top             =   4050
      Width           =   840
   End
   Begin VB.Image imgControl 
      Appearance      =   0  '2D
      BorderStyle     =   1  'Fest Einfach
      Height          =   480
      Index           =   0
      Left            =   5400
      Picture         =   "frmPocketDMX.frx":2AD7D
      Tag             =   "3-255"
      ToolTipText     =   "Pan/Tilt Music Control"
      Top             =   4350
      Width           =   495
   End
   Begin VB.Label lblFunc 
      BackColor       =   &H00000000&
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   5175
      TabIndex        =   10
      Top             =   4050
      Width           =   840
   End
   Begin VB.Label lblFunc 
      BackColor       =   &H00000000&
      Caption         =   "Pan / Tilt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   300
      TabIndex        =   6
      Top             =   4050
      Width           =   840
   End
   Begin VB.Label lblFunc 
      BackColor       =   &H00000000&
      Caption         =   "Color"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   300
      TabIndex        =   5
      Top             =   1875
      Width           =   465
   End
   Begin VB.Label lblFunc 
      BackColor       =   &H00000000&
      Caption         =   "Gobo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   300
      TabIndex        =   4
      Top             =   300
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   8550
      Picture         =   "frmPocketDMX.frx":2B581
      Top             =   375
      Width           =   1500
   End
End
Attribute VB_Name = "frmPocketDMX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const iStartChn As Integer = 1

Dim iPortAddr As Integer
Dim cDMX(1 To 256) As Byte

Private pbResendDMX As Boolean
Private pbMoveCross As Boolean
Private pbAutoSpeed As Boolean
Private plLastSendTicks As Long

Private pbVbrPanSleep As Boolean
Private pbVbrTiltSleep As Boolean
Private pbVbrSpeedSleep As Boolean
Private pbVbrLaserSleep As Boolean
Private pbVbrShutterSleep As Boolean


Private Sub sendDMX()
  Static sbWorking As Boolean
  Dim n As Long, t As Long, a As Long
  
  If Not sbWorking Then
    sbWorking = True
  
    Do
      pbResendDMX = False 'reset state
      
      t = GetTickCount()
      
      Out iPortAddr + 2, 1       'put pin 14 of the printerport high, pin 1 low
      Out iPortAddr + 2, 1
      Out iPortAddr + 2, 1
    '  Sleep 5
      Out iPortAddr + 2, 3       'put pin 14 low, pin 1 low
      Out iPortAddr + 2, 3
      Out iPortAddr + 2, 3
      
      For n = 1 To 256
        Out iPortAddr, cDMX(n)    'put the channel data on the printerport
        Out iPortAddr, cDMX(n)
        Out iPortAddr, cDMX(n)
        '
        Out iPortAddr + 2, 2           'put pin 14 low, pin 1 high
        Out iPortAddr + 2, 2
        Out iPortAddr + 2, 2
        '
        Out iPortAddr + 2, 3           'put pin 14 low, pin 1 low
        Out iPortAddr + 2, 3
        Out iPortAddr + 2, 3
      Next n                                              'repeat 65 times (64 channels + 1 dummy)
    Loop While pbResendDMX
  
    sbWorking = False
  Else  ' caught working
    pbResendDMX = True
  End If
  
  plLastSendTicks = GetTickCount()
  
  'Debug.Print GetTickCount() - t
End Sub

Private Sub chkAutoSpeed_Click()
  pbAutoSpeed = (chkAutoSpeed.Value = vbChecked)
End Sub

'Private Sub Command1_Click()
'Dim t As Long
't = GetTickCount()
'sendDMX
'Label1.Caption = GetTickCount() - t
'
'End Sub

Private Sub Form_Load()
  iPortAddr = &H378
  
  'cDMX(iStartChn + chnShutter) = 255
  'cDMX(iStartChn + chnSpeed) = 0
  sendDMX
  
End Sub

Private Sub imgColor_Click(Index As Integer)
  Dim strDMX As String, iOffset As Integer, cData As Byte
  strDMX = imgColor(Index).Tag
  
  iOffset = Val(FirstArg(strDMX))
  cData = Val(ArgFrom(strDMX))
  
  cDMX(iStartChn + iOffset) = cData
  sendDMX
End Sub

Private Sub imgColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  shpColor.Left = imgColor(Index).Left - 15
  shpColor.Top = imgColor(Index).Top - 15
  shpColor.Visible = True
End Sub

Private Sub imgColor_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  shpColor.Visible = False
End Sub

Private Sub imgCross_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  pbMoveCross = True
  'imgCross_MouseMove Button, Shift, X, Y
End Sub

Private Sub imgCross_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Static ssDX As Single, ssDY As Single, slLastTicks As Long
  Dim p As POINTAPI
  Dim sX As Single, sY As Single, sDX As Single, sDY As Single, sDD As Single
  Dim cPan As Byte, cTilt As Byte
  
  'sDX = Abs(ssDX - X) / Screen.TwipsPerPixelX
  'ssDX = X
  'sDY = Abs(ssDY - Y) / Screen.TwipsPerPixelY
  'ssDY = Y
  
  'Debug.Print GetTickCount() - slLastTicks
  slLastTicks = GetTickCount()
    
  If pbMoveCross Then
    
    GetCursorPos p
    ScreenToClient picPTBox.hWnd, p
    sX = p.X * Screen.TwipsPerPixelX
    sY = p.Y * Screen.TwipsPerPixelY
    If sX <= 0 Then sX = 0
    If sY <= 0 Then sY = 0
    If sX >= picPTBox.ScaleWidth Then sX = picPTBox.ScaleWidth
    If sY >= picPTBox.ScaleWidth Then sY = picPTBox.ScaleHeight
    imgCross.Left = sX - 200
    imgCross.Top = sY - 200
    
    cPan = CByte((sX / picPTBox.ScaleWidth) * 255)
    cTilt = 255 - CByte((sY / picPTBox.ScaleHeight) * 255)
    
    cDMX(iStartChn + chnPan) = 255 - cPan
    cDMX(iStartChn + chnTilt) = cTilt
    
    
      sDD = Sqr((ssDX - sX) * (ssDX - sX) + (ssDY - sY) * (ssDY - sY)) / Screen.TwipsPerPixelX / 365 * 255
  ssDX = sX
  ssDY = sY
    If pbAutoSpeed Then
      pbVbrSpeedSleep = True
        vbrSpeed.Value = sDD / 255 * 238
        cDMX(iStartChn + chnSpeed) = sDD / 255 * 238
      pbVbrSpeedSleep = False
    End If
    
    sendDMX
    
    pbVbrPanSleep = True
      vbrPan.Value = cPan
    pbVbrPanSleep = False
    
    pbVbrTiltSleep = True
      vbrTilt.Value = cTilt
    pbVbrTiltSleep = False
    

    

    
    
  End If
  
End Sub

Private Sub imgCross_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  pbMoveCross = False
End Sub

Private Sub imgGobo_Click(Index As Integer)
  Dim strDMX As String, iOffset As Integer, cData As Byte
  
  On Error GoTo catch
  
  strDMX = imgGobo(Index).Tag
  
  iOffset = CInt(FirstArg(strDMX))
  cData = CByte(ArgFrom(strDMX))
  
  cDMX(iStartChn + iOffset) = cData
  sendDMX
  
catch:
End Sub

Private Sub imgGobo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  shpGobo.Left = imgGobo(Index).Left - 15
  shpGobo.Top = imgGobo(Index).Top - 15
  shpGobo.Visible = True
End Sub

Private Sub imgGobo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  shpGobo.Visible = False
End Sub

Private Sub imgControl_Click(Index As Integer)
  Select Case Index
    Case 0  ' speed/music
      cDMX(iStartChn + chnSpeed) = 255
      chkAutoSpeed.Value = vbUnchecked
    Case 1  ' shutter on
      cDMX(iStartChn + chnShutter) = 255
    Case 2  ' laser on
      cDMX(iStartChn + chnLaser) = 128
    Case 3  ' laser off
      pbVbrLaserSleep = True
        vbrLaser.Value = 0
      pbVbrLaserSleep = False
      cDMX(iStartChn + chnLaser) = 0
    Case 4  ' shutter close
      cDMX(iStartChn + chnShutter) = 0
    Case 5  ' speed/rel
      cDMX(iStartChn + chnSpeed) = 0
      chkAutoSpeed.Value = vbUnchecked
  End Select
  sendDMX
End Sub

Private Sub imgReset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  cDMX(iStartChn + chnLaser) = 240
  sendDMX
End Sub

Private Sub imgReset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  cDMX(iStartChn + chnLaser) = 0
  sendDMX
End Sub

Private Sub Label1_Click()
  Shell "explorer http://developer.brainkiller.org/pocketdmx/"
End Sub

Private Sub picPTBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  pbMoveCross = True
  imgCross_MouseMove Button, Shift, 200, 200
End Sub

Private Sub picPTBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgCross_MouseMove Button, Shift, X, Y
End Sub

Private Sub picPTBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  pbMoveCross = False
End Sub



Private Sub tmrTimer_Timer()
  If GetTickCount() - plLastSendTicks > 100 Then sendDMX
End Sub

Private Sub vbrLaser_Change()
  cDMX(iStartChn + chnLaser) = 16 + vbrLaser.Value
  sendDMX
End Sub

Private Sub vbrPan_Change()
  If Not pbVbrPanSleep Then
    imgCross.Left = picPTBox.ScaleWidth * (vbrPan.Value / 255) - 200
    cDMX(iStartChn + chnPan) = 255 - vbrPan.Value
    sendDMX
  End If
End Sub

Private Sub vbrShutter_Change()
  cDMX(iStartChn + chnShutter) = 16 + vbrShutter.Value
  sendDMX
End Sub

Private Sub vbrSpeed_Change()
  If pbVbrSpeedSleep Then Exit Sub
  chkAutoSpeed.Value = vbUnchecked
  cDMX(iStartChn + chnSpeed) = 16 + vbrSpeed.Value
  sendDMX
End Sub

Private Sub vbrTilt_Change()
  If Not pbVbrTiltSleep Then
    imgCross.Top = picPTBox.ScaleHeight * ((255 - vbrTilt.Value) / 255) - 200
    cDMX(iStartChn + chnTilt) = vbrTilt.Value
    sendDMX
  End If
End Sub
