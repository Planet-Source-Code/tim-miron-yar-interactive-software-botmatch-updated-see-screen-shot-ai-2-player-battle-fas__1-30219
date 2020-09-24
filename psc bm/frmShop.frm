VERSION 5.00
Begin VB.Form frmShop 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "BotMatch - Shop"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmShop.frx":0000
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrLoad 
      Interval        =   100
      Left            =   360
      Top             =   2535
   End
   Begin VB.PictureBox picPics 
      HasDC           =   0   'False
      Height          =   975
      Left            =   3480
      ScaleHeight     =   915
      ScaleWidth      =   1650
      TabIndex        =   129
      Top             =   -975
      Width           =   1710
      Begin VB.PictureBox WeapDrk 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   450
         Left            =   270
         Picture         =   "frmShop.frx":8A12
         ScaleHeight     =   450
         ScaleWidth      =   465
         TabIndex        =   147
         Top             =   375
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox DrkAmmo 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   420
         Left            =   435
         Picture         =   "frmShop.frx":8E2E
         ScaleHeight     =   420
         ScaleWidth      =   480
         TabIndex        =   146
         Top             =   270
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox drkShield 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   435
         Left            =   555
         Picture         =   "frmShop.frx":9215
         ScaleHeight     =   435
         ScaleWidth      =   420
         TabIndex        =   145
         Top             =   450
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.PictureBox drkTools 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   435
         Left            =   435
         Picture         =   "frmShop.frx":95D2
         ScaleHeight     =   435
         ScaleWidth      =   465
         TabIndex        =   144
         Top             =   510
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox WeapLit 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   450
         Left            =   345
         Picture         =   "frmShop.frx":9A64
         ScaleHeight     =   450
         ScaleWidth      =   465
         TabIndex        =   143
         Top             =   660
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.PictureBox LitAmmo 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   480
         Left            =   780
         Picture         =   "frmShop.frx":9F40
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   142
         Top             =   585
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox LitShield 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   480
         Left            =   120
         Picture         =   "frmShop.frx":A4A6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   141
         Top             =   90
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox LitTools 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   480
         Left            =   465
         Picture         =   "frmShop.frx":A99D
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   140
         Top             =   540
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox LitBuy 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   390
         Left            =   525
         Picture         =   "frmShop.frx":AF87
         ScaleHeight     =   390
         ScaleWidth      =   435
         TabIndex        =   139
         Top             =   225
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox LitSell 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   375
         Left            =   255
         Picture         =   "frmShop.frx":B48B
         ScaleHeight     =   375
         ScaleWidth      =   525
         TabIndex        =   138
         Top             =   570
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.PictureBox DrkBuy 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   390
         Left            =   105
         Picture         =   "frmShop.frx":B9D0
         ScaleHeight     =   390
         ScaleWidth      =   435
         TabIndex        =   137
         Top             =   435
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox DrkSell 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   375
         Left            =   450
         Picture         =   "frmShop.frx":BDCA
         ScaleHeight     =   375
         ScaleWidth      =   525
         TabIndex        =   136
         Top             =   585
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.PictureBox SBDwnButUP 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   240
         Left            =   1290
         Picture         =   "frmShop.frx":C1F9
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   135
         Top             =   255
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox SBDwnButP 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   240
         Left            =   1035
         Picture         =   "frmShop.frx":C53B
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   134
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox SBUpButUP 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   240
         Left            =   1290
         Picture         =   "frmShop.frx":C87D
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   133
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox SBUpButP 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   240
         Left            =   1035
         Picture         =   "frmShop.frx":CBBF
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   132
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox upAr 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   240
         Left            =   1530
         Picture         =   "frmShop.frx":CF01
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   131
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox DownAr 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   240
         Left            =   1545
         Picture         =   "frmShop.frx":D243
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   130
         Top             =   240
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImgBlueC 
         Height          =   960
         Left            =   0
         Picture         =   "frmShop.frx":D585
         Top             =   315
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image imgRedC 
         Height          =   960
         Left            =   255
         Picture         =   "frmShop.frx":E1EE
         Top             =   225
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Image LitDisk 
         Height          =   435
         Left            =   1125
         Picture         =   "frmShop.frx":EE4B
         Top             =   675
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Image DrkDisk 
         Height          =   435
         Left            =   1125
         Picture         =   "frmShop.frx":F22A
         Top             =   540
         Visible         =   0   'False
         Width           =   435
      End
   End
   Begin VB.PictureBox Text 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1560
      Left            =   540
      ScaleHeight     =   1560
      ScaleWidth      =   3720
      TabIndex        =   125
      Top             =   5100
      Width           =   3720
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   1545
         Left            =   45
         TabIndex        =   126
         Top             =   0
         Width           =   3660
      End
   End
   Begin VB.Timer tmrStat 
      Interval        =   1
      Left            =   5115
      Top             =   6825
   End
   Begin VB.PictureBox ButDown 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   240
      Left            =   7980
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   124
      Top             =   4740
      Width           =   240
   End
   Begin VB.PictureBox ButUp 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   240
      Left            =   7980
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   123
      Top             =   4380
      Width           =   240
   End
   Begin VB.PictureBox ShieldBar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   285
      Left            =   6030
      ScaleHeight     =   285
      ScaleWidth      =   2985
      TabIndex        =   122
      Top             =   6030
      Width           =   2985
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   99
         X1              =   2970
         X2              =   2970
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   98
         X1              =   2940
         X2              =   2940
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   97
         X1              =   2910
         X2              =   2910
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   96
         X1              =   2880
         X2              =   2880
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   95
         X1              =   2850
         X2              =   2850
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   94
         X1              =   2820
         X2              =   2820
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   93
         X1              =   2790
         X2              =   2790
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   92
         X1              =   2760
         X2              =   2760
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   91
         X1              =   2730
         X2              =   2730
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   90
         X1              =   2700
         X2              =   2700
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   89
         X1              =   2670
         X2              =   2670
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   88
         X1              =   2640
         X2              =   2640
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   87
         X1              =   2610
         X2              =   2610
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   86
         X1              =   2580
         X2              =   2580
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   85
         X1              =   2550
         X2              =   2550
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   84
         X1              =   2520
         X2              =   2520
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   83
         X1              =   2490
         X2              =   2490
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   82
         X1              =   2460
         X2              =   2460
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   81
         X1              =   2430
         X2              =   2430
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   80
         X1              =   2400
         X2              =   2400
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   79
         X1              =   2370
         X2              =   2370
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   78
         X1              =   2340
         X2              =   2340
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   77
         X1              =   2310
         X2              =   2310
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   76
         X1              =   2280
         X2              =   2280
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   75
         X1              =   2250
         X2              =   2250
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   74
         X1              =   2220
         X2              =   2220
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   73
         X1              =   2190
         X2              =   2190
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   72
         X1              =   2160
         X2              =   2160
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   71
         X1              =   2130
         X2              =   2130
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   70
         X1              =   2100
         X2              =   2100
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   69
         X1              =   2070
         X2              =   2070
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   68
         X1              =   2040
         X2              =   2040
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   67
         X1              =   2010
         X2              =   2010
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   66
         X1              =   1980
         X2              =   1980
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   65
         X1              =   1950
         X2              =   1950
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   64
         X1              =   1920
         X2              =   1920
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   63
         X1              =   1890
         X2              =   1890
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   62
         X1              =   1860
         X2              =   1860
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   61
         X1              =   1830
         X2              =   1830
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   60
         X1              =   1800
         X2              =   1800
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   59
         X1              =   1770
         X2              =   1770
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   58
         X1              =   1740
         X2              =   1740
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   57
         X1              =   1710
         X2              =   1710
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   56
         X1              =   1680
         X2              =   1680
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   55
         X1              =   1650
         X2              =   1650
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   54
         X1              =   1620
         X2              =   1620
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   53
         X1              =   1590
         X2              =   1590
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   52
         X1              =   1560
         X2              =   1560
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   51
         X1              =   1530
         X2              =   1530
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   50
         X1              =   1500
         X2              =   1500
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   49
         X1              =   1470
         X2              =   1470
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   48
         X1              =   1440
         X2              =   1440
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   47
         X1              =   1410
         X2              =   1410
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   46
         X1              =   1380
         X2              =   1380
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   45
         X1              =   1350
         X2              =   1350
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   44
         X1              =   1320
         X2              =   1320
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   43
         X1              =   1290
         X2              =   1290
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   42
         X1              =   1260
         X2              =   1260
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   41
         X1              =   1230
         X2              =   1230
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   40
         X1              =   1200
         X2              =   1200
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   39
         X1              =   1170
         X2              =   1170
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   38
         X1              =   1140
         X2              =   1140
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   37
         X1              =   1110
         X2              =   1110
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   36
         X1              =   1080
         X2              =   1080
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   35
         X1              =   1050
         X2              =   1050
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   34
         X1              =   1020
         X2              =   1020
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   33
         X1              =   990
         X2              =   990
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   32
         X1              =   960
         X2              =   960
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   31
         X1              =   930
         X2              =   930
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   30
         X1              =   900
         X2              =   900
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   29
         X1              =   870
         X2              =   870
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   28
         X1              =   840
         X2              =   840
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   27
         X1              =   810
         X2              =   810
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   26
         X1              =   780
         X2              =   780
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   25
         X1              =   750
         X2              =   750
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   24
         X1              =   720
         X2              =   720
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   23
         X1              =   690
         X2              =   690
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   22
         X1              =   660
         X2              =   660
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   21
         X1              =   630
         X2              =   630
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   20
         X1              =   600
         X2              =   600
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   19
         X1              =   570
         X2              =   570
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   18
         X1              =   540
         X2              =   540
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   17
         X1              =   510
         X2              =   510
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   16
         X1              =   480
         X2              =   480
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   15
         X1              =   450
         X2              =   450
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   14
         X1              =   420
         X2              =   420
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   13
         X1              =   390
         X2              =   390
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   12
         X1              =   360
         X2              =   360
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   11
         X1              =   330
         X2              =   330
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   10
         X1              =   300
         X2              =   300
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   9
         X1              =   270
         X2              =   270
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   8
         X1              =   240
         X2              =   240
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   7
         X1              =   210
         X2              =   210
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   6
         X1              =   180
         X2              =   180
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   5
         X1              =   150
         X2              =   150
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   4
         X1              =   120
         X2              =   120
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   3
         X1              =   90
         X2              =   90
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   2
         X1              =   60
         X2              =   60
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   1
         X1              =   30
         X2              =   30
         Y1              =   255
         Y2              =   0
      End
      Begin VB.Line LineShields 
         BorderColor     =   &H00C0C000&
         Index           =   0
         X1              =   0
         X2              =   0
         Y1              =   255
         Y2              =   0
      End
   End
   Begin VB.PictureBox ButDisk 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   8790
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   121
      Top             =   2430
      Width           =   435
   End
   Begin VB.PictureBox HealthBar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   6030
      ScaleHeight     =   300
      ScaleWidth      =   3000
      TabIndex        =   20
      Top             =   5415
      Width           =   3000
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   99
         Left            =   2970
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   120
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   98
         Left            =   2940
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   119
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   97
         Left            =   2910
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   118
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   96
         Left            =   2880
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   117
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   95
         Left            =   2850
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   116
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   94
         Left            =   2820
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   115
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   93
         Left            =   2790
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   114
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   92
         Left            =   2760
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   113
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   91
         Left            =   2730
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   112
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   90
         Left            =   2700
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   111
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   89
         Left            =   2670
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   110
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   88
         Left            =   2640
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   109
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   87
         Left            =   2610
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   108
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   86
         Left            =   2580
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   107
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   85
         Left            =   2550
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   106
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   84
         Left            =   2520
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   105
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   83
         Left            =   2490
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   104
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   82
         Left            =   2460
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   103
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   81
         Left            =   2430
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   102
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   80
         Left            =   2400
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   101
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   79
         Left            =   2370
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   100
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   78
         Left            =   2340
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   99
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   77
         Left            =   2310
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   98
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   76
         Left            =   2280
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   97
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   75
         Left            =   2250
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   96
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   74
         Left            =   2220
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   95
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   73
         Left            =   2190
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   94
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   72
         Left            =   2160
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   93
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   71
         Left            =   2130
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   92
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   70
         Left            =   2100
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   91
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   69
         Left            =   2070
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   90
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   68
         Left            =   2040
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   89
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   67
         Left            =   2010
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   88
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   66
         Left            =   1980
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   87
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   65
         Left            =   1950
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   86
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   64
         Left            =   1920
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   85
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   63
         Left            =   1890
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   84
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   62
         Left            =   1860
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   83
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   61
         Left            =   1830
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   82
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   60
         Left            =   1800
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   81
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   59
         Left            =   1770
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   80
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   58
         Left            =   1740
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   79
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   57
         Left            =   1710
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   78
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   56
         Left            =   1680
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   77
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   55
         Left            =   1650
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   76
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   54
         Left            =   1620
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   75
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   53
         Left            =   1590
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   74
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   52
         Left            =   1560
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   73
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   51
         Left            =   1530
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   72
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   50
         Left            =   1500
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   71
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   49
         Left            =   1470
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   70
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   48
         Left            =   1440
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   69
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   47
         Left            =   1410
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   68
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   46
         Left            =   1380
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   67
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   45
         Left            =   1350
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   66
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   44
         Left            =   1320
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   65
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   43
         Left            =   1290
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   64
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   42
         Left            =   1260
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   63
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   41
         Left            =   1230
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   62
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   40
         Left            =   1200
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   61
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   39
         Left            =   1170
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   60
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   38
         Left            =   1140
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   59
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   37
         Left            =   1110
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   58
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   36
         Left            =   1080
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   57
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   35
         Left            =   1050
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   56
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   34
         Left            =   1020
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   55
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   33
         Left            =   990
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   54
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   32
         Left            =   960
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   53
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   31
         Left            =   930
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   52
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   30
         Left            =   900
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   51
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   29
         Left            =   870
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   50
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   28
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   49
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   27
         Left            =   810
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   48
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   26
         Left            =   780
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   47
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   25
         Left            =   750
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   46
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   24
         Left            =   720
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   45
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   23
         Left            =   690
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   44
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   22
         Left            =   660
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   43
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   21
         Left            =   630
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   42
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   20
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   41
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   19
         Left            =   570
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   40
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   18
         Left            =   540
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   39
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   17
         Left            =   510
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   38
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   16
         Left            =   480
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   37
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   15
         Left            =   450
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   36
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   14
         Left            =   420
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   35
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   13
         Left            =   390
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   34
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   12
         Left            =   360
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   33
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   11
         Left            =   330
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   32
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   10
         Left            =   300
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   31
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   9
         Left            =   270
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   30
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   8
         Left            =   240
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   29
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   7
         Left            =   210
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   28
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   6
         Left            =   180
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   27
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   5
         Left            =   150
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   26
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   4
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   25
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   3
         Left            =   90
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   24
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   2
         Left            =   60
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   23
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   1
         Left            =   30
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   22
         Top             =   15
         Width           =   15
      End
      Begin VB.PictureBox HltPnt 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         HasDC           =   0   'False
         Height          =   255
         Index           =   0
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   21
         Top             =   15
         Width           =   15
      End
   End
   Begin VB.PictureBox picWeap 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   7425
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   18
      Top             =   4440
      Width           =   480
   End
   Begin VB.PictureBox SellBut 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   375
      Left            =   5880
      ScaleHeight     =   375
      ScaleWidth      =   525
      TabIndex        =   17
      Top             =   1365
      Width           =   525
   End
   Begin VB.PictureBox BuyBut 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   390
      Left            =   5235
      ScaleHeight     =   390
      ScaleWidth      =   435
      TabIndex        =   16
      Top             =   1350
      Width           =   435
   End
   Begin VB.PictureBox UtilBut 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   480
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   3120
      Width           =   480
   End
   Begin VB.PictureBox ShieldBut 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   480
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   2415
      Width           =   480
   End
   Begin VB.PictureBox AmmoBut 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   480
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   1710
      Width           =   480
   End
   Begin VB.PictureBox WeapBut 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   495
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   1005
      Width           =   480
   End
   Begin VB.Shape shpCatBorder 
      BorderColor     =   &H0080FF80&
      Height          =   540
      Left            =   465
      Top             =   975
      Width           =   540
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H0080FF80&
      Height          =   285
      Left            =   8115
      Top             =   435
      Width           =   1035
   End
   Begin VB.Label butBudget 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "Budget "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00002000&
      Height          =   225
      Left            =   8145
      TabIndex        =   155
      Top             =   975
      Width           =   975
   End
   Begin VB.Label butQuant 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "Quantity "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00002000&
      Height          =   225
      Left            =   8145
      TabIndex        =   154
      Top             =   720
      Width           =   975
   End
   Begin VB.Label ButSummary 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "Summary "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00002000&
      Height          =   225
      Left            =   8145
      TabIndex        =   153
      Top             =   465
      Width           =   975
   End
   Begin VB.Label lblOther 
      BackStyle       =   0  'Transparent
      Caption         =   "--- Not Available ---"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   735
      Left            =   5280
      TabIndex        =   152
      Top             =   465
      Width           =   2715
   End
   Begin VB.Shape shpDisp 
      BorderColor     =   &H00004000&
      FillColor       =   &H00002000&
      FillStyle       =   0  'Solid
      Height          =   825
      Index           =   4
      Left            =   7995
      Top             =   420
      Width           =   1155
   End
   Begin VB.Shape shpDisp 
      BorderColor     =   &H00004000&
      Height          =   795
      Index           =   3
      Left            =   5220
      Top             =   435
      Width           =   2790
   End
   Begin VB.Shape shpLstBorder 
      BorderColor     =   &H0000C000&
      Height          =   915
      Index           =   15
      Left            =   5160
      Top             =   375
      Width           =   4035
   End
   Begin VB.Shape shpLstBorder 
      BorderColor     =   &H00008000&
      Height          =   885
      Index           =   14
      Left            =   5175
      Top             =   390
      Width           =   4005
   End
   Begin VB.Shape shpLstBorder 
      BorderColor     =   &H00008000&
      Height          =   945
      Index           =   13
      Left            =   5145
      Top             =   360
      Width           =   4065
   End
   Begin VB.Shape shpLstBorder 
      BorderColor     =   &H00004000&
      Height          =   975
      Index           =   10
      Left            =   5130
      Top             =   345
      Width           =   4095
   End
   Begin VB.Label lblTotalPrice 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   7620
      TabIndex        =   151
      Top             =   1410
      Width           =   1485
   End
   Begin VB.Label lblDead 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Price -"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   210
      Left            =   6615
      TabIndex        =   150
      Top             =   1455
      Width           =   975
   End
   Begin VB.Label lblHull 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hull Strength"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   210
      Left            =   7995
      TabIndex        =   149
      Top             =   2955
      Width           =   1065
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Item Selected"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   525
      TabIndex        =   148
      Top             =   4740
      Width           =   1920
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H00008000&
      Index           =   13
      X1              =   254
      X2              =   168
      Y1              =   35
      Y2              =   35
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H00008000&
      Index           =   12
      X1              =   48
      X2              =   48
      Y1              =   243
      Y2              =   270
   End
   Begin VB.Shape shpDisp 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   660
      Shape           =   3  'Circle
      Top             =   465
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   1005
      Picture         =   "frmShop.frx":F5FE
      Top             =   390
      Width           =   1470
   End
   Begin VB.Label lblShield 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   210
      Left            =   5415
      TabIndex        =   128
      Top             =   5820
      Width           =   45
   End
   Begin VB.Label lblHealth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   210
      Left            =   5415
      TabIndex        =   127
      Top             =   5205
      Width           =   45
   End
   Begin VB.Shape shpLstBorder 
      BorderColor     =   &H00004000&
      Height          =   1680
      Index           =   12
      Left            =   480
      Top             =   5040
      Width           =   3840
   End
   Begin VB.Shape shpLstBorder 
      BorderColor     =   &H00008000&
      Height          =   1650
      Index           =   11
      Left            =   495
      Top             =   5055
      Width           =   3810
   End
   Begin VB.Shape shpLstBorder 
      BorderColor     =   &H0000C000&
      Height          =   1620
      Index           =   9
      Left            =   510
      Top             =   5070
      Width           =   3780
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00008000&
      FillColor       =   &H00002000&
      FillStyle       =   6  'Cross
      Height          =   1395
      Left            =   7845
      Shape           =   4  'Rounded Rectangle
      Top             =   2910
      Width           =   1365
   End
   Begin VB.Label lblStat 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   270
      Left            =   5265
      TabIndex        =   19
      Top             =   6510
      Width           =   3435
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H00008000&
      Index           =   8
      X1              =   346
      X2              =   346
      Y1              =   430
      Y2              =   342
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H00008000&
      Index           =   7
      X1              =   607
      X2              =   345
      Y1              =   342
      Y2              =   342
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H00008000&
      Index           =   6
      X1              =   607
      X2              =   607
      Y1              =   343
      Y2              =   437
   End
   Begin VB.Shape shpDisp 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   9045
      Shape           =   3  'Circle
      Top             =   6555
      Width           =   135
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H00008000&
      Index           =   5
      X1              =   602
      X2              =   582
      Y1              =   440
      Y2              =   440
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H0000C000&
      Index           =   4
      X1              =   539
      X2              =   539
      Y1              =   315
      Y2              =   306
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H00008000&
      Index           =   3
      X1              =   346
      X2              =   346
      Y1              =   263
      Y2              =   180
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H00008000&
      Index           =   2
      X1              =   394
      X2              =   346
      Y1              =   263
      Y2              =   263
   End
   Begin VB.Shape shpDisp 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   5130
      Shape           =   3  'Circle
      Top             =   2580
      Width           =   135
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H00008000&
      Index           =   1
      X1              =   580
      X2              =   528
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H00008000&
      Index           =   0
      X1              =   360
      X2              =   345
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Shape shpLstBorder 
      BorderColor     =   &H0000C000&
      Height          =   540
      Index           =   8
      Left            =   7395
      Top             =   4410
      Width           =   540
   End
   Begin VB.Shape shpLstBorder 
      BorderColor     =   &H00008000&
      Height          =   510
      Index           =   7
      Left            =   7410
      Top             =   4425
      Width           =   510
   End
   Begin VB.Shape shpLstBorder 
      BorderColor     =   &H00008000&
      Height          =   570
      Index           =   6
      Left            =   7380
      Top             =   4395
      Width           =   570
   End
   Begin VB.Shape shpLstBorder 
      BorderColor     =   &H00004000&
      Height          =   600
      Index           =   5
      Left            =   7365
      Top             =   4380
      Width           =   600
   End
   Begin VB.Image imgTitle 
      Height          =   240
      Left            =   5460
      Picture         =   "frmShop.frx":FA89
      Top             =   2535
      Width           =   2445
   End
   Begin VB.Image ImgBot 
      Height          =   960
      Left            =   6390
      Top             =   3465
      Width           =   960
   End
   Begin VB.Label LstItem 
      BackColor       =   &H00002000&
      Caption         =   "   Item Not Available"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   11
      Left            =   1275
      TabIndex        =   15
      Top             =   3825
      Width           =   3075
   End
   Begin VB.Label LstItem 
      BackColor       =   &H00002000&
      Caption         =   "   Item Not Available"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   10
      Left            =   1275
      TabIndex        =   14
      Top             =   3555
      Width           =   3075
   End
   Begin VB.Label LstItem 
      BackColor       =   &H00002000&
      Caption         =   "   Item Not Available"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   9
      Left            =   1275
      TabIndex        =   13
      Top             =   3285
      Width           =   3075
   End
   Begin VB.Label LstItem 
      BackColor       =   &H00002000&
      Caption         =   "   Item Not Available"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   8
      Left            =   1275
      TabIndex        =   12
      Top             =   3015
      Width           =   3075
   End
   Begin VB.Label LstItem 
      BackColor       =   &H00002000&
      Caption         =   "   Item Not Available"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   7
      Left            =   1275
      TabIndex        =   11
      Top             =   2745
      Width           =   3075
   End
   Begin VB.Label LstItem 
      BackColor       =   &H00002000&
      Caption         =   "   Item Not Available"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   6
      Left            =   1275
      TabIndex        =   10
      Top             =   2475
      Width           =   3075
   End
   Begin VB.Label LstItem 
      BackColor       =   &H00002000&
      Caption         =   "   Item Not Available"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   5
      Left            =   1275
      TabIndex        =   9
      Top             =   2205
      Width           =   3075
   End
   Begin VB.Label LstItem 
      BackColor       =   &H00002000&
      Caption         =   "   Item Not Available"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   4
      Left            =   1275
      TabIndex        =   8
      Top             =   1935
      Width           =   3075
   End
   Begin VB.Label LstItem 
      BackColor       =   &H00002000&
      Caption         =   "   Item Not Available"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   3
      Left            =   1275
      TabIndex        =   7
      Top             =   1665
      Width           =   3075
   End
   Begin VB.Label LstItem 
      BackColor       =   &H00002000&
      Caption         =   "   Item Not Available"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   2
      Left            =   1275
      TabIndex        =   6
      Top             =   1395
      Width           =   3075
   End
   Begin VB.Label LstItem 
      BackColor       =   &H00002000&
      Caption         =   "   Item Not Available"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   1275
      TabIndex        =   5
      Top             =   1125
      Width           =   3075
   End
   Begin VB.Shape shpLstBorder 
      BorderColor     =   &H00004000&
      Height          =   3405
      Index           =   4
      Left            =   1185
      Top             =   765
      Width           =   3255
   End
   Begin VB.Shape shpLstBorder 
      BorderColor     =   &H00008000&
      Height          =   3375
      Index           =   3
      Left            =   1200
      Top             =   780
      Width           =   3225
   End
   Begin VB.Shape shpLstBorder 
      BorderColor     =   &H00004000&
      Height          =   3285
      Index           =   1
      Left            =   1245
      Top             =   825
      Width           =   3135
   End
   Begin VB.Shape shpLstBorder 
      BorderColor     =   &H00008000&
      Height          =   3315
      Index           =   0
      Left            =   1230
      Top             =   810
      Width           =   3165
   End
   Begin VB.Label LstItem 
      BackColor       =   &H00002000&
      Caption         =   "   Item Not Available"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   1275
      TabIndex        =   4
      Top             =   855
      Width           =   3075
   End
   Begin VB.Shape shpLstBorder 
      BorderColor     =   &H0000C000&
      Height          =   3345
      Index           =   2
      Left            =   1215
      Top             =   795
      Width           =   3195
   End
   Begin VB.Shape shpOuterShield 
      BorderColor     =   &H00008000&
      Height          =   1920
      Left            =   5910
      Shape           =   3  'Circle
      Top             =   2970
      Width           =   1920
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00004000&
      FillStyle       =   0  'Solid
      Height          =   330
      Left            =   5160
      Top             =   6465
      Width           =   3570
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00002000&
      FillStyle       =   6  'Cross
      Height          =   1455
      Left            =   5190
      Top             =   5145
      Width           =   3945
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00008000&
      FillColor       =   &H00002000&
      FillStyle       =   6  'Cross
      Height          =   2115
      Left            =   330
      Shape           =   4  'Rounded Rectangle
      Top             =   4695
      Width           =   4140
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H00008000&
      Index           =   9
      X1              =   62
      X2              =   47
      Y1              =   35
      Y2              =   35
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H00008000&
      Index           =   10
      X1              =   48
      X2              =   48
      Y1              =   64
      Y2              =   34
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H00008000&
      Index           =   11
      X1              =   79
      X2              =   47
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H00008000&
      Index           =   14
      X1              =   254
      X2              =   254
      Y1              =   51
      Y2              =   35
   End
End
Attribute VB_Name = "frmShop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CurFocus As Byte
Private CurListFocus As Byte
  Private SEffectIndex As Byte
  Private SEffectFlag As Boolean
  
  Private Type AnimStatus
    CurPos As Integer
    StringLength As Integer
    Expression As String
  End Type
  
  Private TmrIndex As Byte
  Private DontNeedAnim As Boolean
  Private TmpStr As String
Private AnimStat As AnimStatus

Private DispSel As Byte 'Selected Panel Display

'is there an item thats selected
Private ActiveSelItem As Boolean
Private CurCate As Byte   'current category

Private Sub AmmoBut_Click()
SetToCategory 2
End Sub

Private Sub AmmoBut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CurFocus = 2 Then Exit Sub
ResCateButs
    CurFocus = 2
        AmmoBut.Picture = LitAmmo.Picture
SetUpStatus "Switch to Ammunitions Catelog"
End Sub

Private Sub butBudget_Click()
DispSel = 3
  shpBorder.Top = butBudget.Top - 2
End Sub

Private Sub butBudget_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CurFocus = 12 Then Exit Sub
ResCateButs
    CurFocus = 12
    With butBudget
    .ForeColor = 0
    .BackColor = &HFF00&
    End With
End Sub

Private Sub ButDisk_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CurFocus = 7 Then Exit Sub
ResCateButs
    CurFocus = 7
        ButDisk.Picture = LitDisk.Picture
    SetUpStatus "Save Bot..."
End Sub
Private Sub ButDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButDown.Picture = SBDwnButP.Picture
End Sub
Private Sub ButDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CurFocus = 9 Then Exit Sub
ResCateButs
    CurFocus = 9
        ButDown.Picture = SBDwnButUP.Picture
    SetUpStatus "Toggle Shield Type"
End Sub
Private Sub ButDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButDown.Picture = SBDwnButUP.Picture
End Sub

Private Sub butQuant_Click()
DispSel = 2
shpBorder.Top = butQuant.Top - 2
End Sub

Private Sub butQuant_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CurFocus = 11 Then Exit Sub
ResCateButs
    CurFocus = 11
    With butQuant
     .ForeColor = 0
     .BackColor = &HFF00&
    End With
End Sub

Private Sub ButSummary_Click()
   DispSel = 1
       ResCateButs
  shpBorder.Top = ButSummary.Top - 2
End Sub

Private Sub ButSummary_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CurFocus = 10 Then Exit Sub
ResCateButs
    CurFocus = 10
    With ButSummary
     .ForeColor = 0
     .BackColor = &HFF00&
    End With
End Sub

Private Sub ButUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButUp.Picture = SBUpButP.Picture
End Sub
Private Sub ButUp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CurFocus = 8 Then Exit Sub
ResCateButs
    CurFocus = 8
        ButUp.Picture = SBUpButUP.Picture
    SetUpStatus "Toggle Shield Type"
End Sub
Private Sub ButUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ButUp.Picture = SBUpButUP.Picture
End Sub
Private Sub BuyBut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CurFocus = 5 Then Exit Sub
ResCateButs
    CurFocus = 5
        BuyBut.Picture = LitBuy.Picture
    SetUpStatus "Buy the Selected Item"
End Sub

Private Sub Form_Click()
SetHealthBars 121, 200, 33, 500
End Sub

Private Sub lblDesc_Change()
 If lblDesc.Caption = "-NA-" Then _
 lblDesc.Caption = "----- Not Available -----"
End Sub

Private Sub LstItem_Click(Index As Integer)
Dim TmpStr As String
 lblDesc.Caption = RetrieveItemData(CurCate, Index, 2)
 TmpStr = RetrieveItemData(CurCate, Index, 1)
 
 '-Set title Caption-'
 If TmpStr = "-NA-" Then
 lblTitle.Caption = "No Item Selected"
   Else
   lblTitle.Caption = TmpStr
 End If
 
 TmpStr = RetrieveItemData(CurCate, Index, 5)
 If TmpStr = "-NA-" Then
  lblOther.Caption = "--- Not Available ---"
    Else
     lblOther.Caption = TmpStr
 End If
End Sub

Private Sub SellBut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CurFocus = 6 Then Exit Sub
ResCateButs
    CurFocus = 6
        SellBut.Picture = LitSell.Picture
    SetUpStatus "Sell the Selected Item"
End Sub
Private Sub Form_Load()
'Set default category button pictures...
UtilBut.Picture = drkTools.Picture
WeapBut.Picture = WeapDrk.Picture
ShieldBut.Picture = drkShield.Picture
AmmoBut.Picture = DrkAmmo.Picture
   BuyBut.Picture = DrkBuy.Picture
   SellBut.Picture = DrkSell.Picture
   
   ButDisk.Picture = DrkDisk.Picture
    Call GradiBar
   ButUp.Picture = upAr.Picture
   ButDown.Picture = DownAr.Picture
   
   'Test only
    ImgBot.Picture = ImgBlueC.Picture
End Sub

Public Sub ResCateButs()
If CurFocus = 0 Then Exit Sub
'0 = none
'1 = weaps
'2 = ammo
'3 = shields
'4 = utils
   
   DontNeedAnim = False
  lblStat.Caption = "_"
Select Case CurFocus
 Case 1
  WeapBut.Picture = WeapDrk.Picture
     CurFocus = 0
 Case 2
  AmmoBut.Picture = DrkAmmo.Picture
     CurFocus = 0
 Case 3
  ShieldBut.Picture = drkShield.Picture
     CurFocus = 0
 Case 4
  UtilBut.Picture = drkTools.Picture
     CurFocus = 0
 Case 5
   BuyBut.Picture = DrkBuy.Picture
   CurFocus = 0
 Case 6
   SellBut.Picture = DrkSell.Picture
   CurFocus = 0
 Case 7
   ButDisk.Picture = DrkDisk.Picture
   CurFocus = 0
 Case 8
   ButUp.Picture = upAr.Picture
   CurFocus = 0
 Case 9
   ButDown.Picture = DownAr.Picture
   CurFocus = 0
 Case 10
  With ButSummary
   .BackColor = &H8000&
   .ForeColor = &H2000&
   End With
  CurFocus = 0

Case 11
   With butQuant
   .BackColor = &H8000&
   .ForeColor = &H2000&
   End With
  CurFocus = 0
Case 12
   With butBudget
   .BackColor = &H8000&
   .ForeColor = &H2000&
   End With
  CurFocus = 0
End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ResCateButs
Call ResALLListColors
End Sub

Private Sub LstItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     With LstItem(Index)
        .BackColor = &HFF00&
        .ForeColor = &H2000&
     End With
End Sub

Sub LstItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If (CurListFocus - 1) = Index Then Exit Sub
   ResListColors
   CurListFocus = Index + 1
     With LstItem(CurListFocus - 1)
        .BackColor = &H4000&
        .ForeColor = &HC0FFC0
     End With
End Sub

Private Sub LstItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next 'avoid array error

 With LstItem(Index)
    .BackColor = &H8000&
    .ForeColor = 0
 End With
End Sub

Private Sub ShieldBut_Click()
SetToCategory 3
End Sub

Private Sub ShieldBut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'hover color for list
If CurFocus = 3 Then Exit Sub
ResCateButs
    CurFocus = 3
        ShieldBut.Picture = LitShield.Picture
SetUpStatus "Switch to Shields/Defence Catelog"
End Sub

Private Sub tmrLoad_Timer()
tmrLoad.Enabled = False
OpenFile App.Path & "\items.ini"
SetToCategory 1
End Sub

Private Sub tmrStat_Timer()
If TmrIndex = 5 Then
        If TmpStr = "_" Then
            TmpStr = ""
        Else
            TmpStr = "_"
        End If
        TmrIndex = 0
     Else
     TmrIndex = TmrIndex + 1
End If
    
If DontNeedAnim = False Then
    lblStat.Caption = TmpStr
  Else
With lblStat
.Caption = _
VBA.Left(AnimStat.Expression, AnimStat.CurPos) _
& TmpStr
.Refresh
End With
If AnimStat.CurPos = AnimStat.StringLength Then
   'DontNeedAnim = True
   'tmrStat.Enabled = False
   Exit Sub
End If
 AnimStat.CurPos = AnimStat.CurPos + 1
End If
End Sub

Private Sub UtilBut_Click()
SetToCategory 4
End Sub

Private Sub UtilBut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CurFocus = 4 Then Exit Sub
ResCateButs
    CurFocus = 4
        UtilBut.Picture = LitTools.Picture
    SetUpStatus "Switch to Upgrades && Others Catelog"
End Sub
Private Sub WeapBut_Click()
SetToCategory 1
End Sub

Private Sub WeapBut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If CurFocus = 1 Then Exit Sub
ResCateButs
    CurFocus = 1
        WeapBut.Picture = WeapLit.Picture
    SetUpStatus "Switch to Weapons Catelog"
End Sub

Public Sub ResListColors()
On Error Resume Next
 With LstItem(CurListFocus - 1)
  .BackColor = &H2000&
  .ForeColor = &HC000&
 End With
End Sub
Public Sub ResALLListColors()
If CurListFocus = 0 Then Exit Sub
Dim X As Byte
On Error Resume Next

For X = 0 To 11
    With LstItem(X)
      .BackColor = &H2000&
      .ForeColor = &HC000&
    End With
Next X
CurListFocus = 0
End Sub

Public Sub GradiBar()
'Draw Gradients on the health/Shield
'strength graph bars
Dim C As Long
Dim X As Long
Dim a As Long
C = 16582
a = 4210888

For X = 0 To 99
  HltPnt(X).BackColor = C
  LineShields(X).BorderColor = a
  C = C - 2
  a = a - 2
Next X
End Sub

Public Sub SetUpStatus(Expression As String)
If DontNeedAnim = True Then Exit Sub
DontNeedAnim = True 'done animating statusbar
AnimStat.Expression = Expression
AnimStat.StringLength = Len(Expression)
AnimStat.CurPos = 1
'tmrStat.Enabled = True
End Sub

Public Sub SetHealthBars(CurHealth As Integer _
, MaxHealth As Integer, CurShield As Integer _
, MaxShield As Integer)
'Procedure sets up health bar, and shield bar...
On Error Resume Next

Dim HealthPercent As Byte
Dim ShieldPercent As Byte
Dim X As Long

HealthPercent = Round((CurHealth / MaxHealth) * 100, 0)
ShieldPercent = Round((CurShield / MaxShield) * 100, 0)

lblHealth.Caption = "Hull Integrety   -  " & HealthPercent & "%"
lblShield.Caption = "Shield Power   -  " & ShieldPercent & "%"

For X = 0 To HealthPercent - 1
 HltPnt(X).Visible = True
Next X
For X = HealthPercent To 99
 HltPnt(X).Visible = False
Next X

For X = 0 To ShieldPercent - 1
 LineShields(X).Visible = True
Next X
For X = ShieldPercent To 99
 LineShields(X).Visible = False
Next X

End Sub

Public Sub SetToCategory(Section As Byte)
Dim C As Integer 'rep counter
Dim TmpStr1 As String 'temp string

CurCate = Section 'set current category variable

Select Case Section
 Case 1
   shpCatBorder.Top = WeapBut.Top - 2
 Case 2
   shpCatBorder.Top = AmmoBut.Top - 2
 Case 3
   shpCatBorder.Top = ShieldBut.Top - 2
 Case 4
   shpCatBorder.Top = UtilBut.Top - 2
End Select
      For C = 0 To 11
        TmpStr1 = RetrieveItemData(Section, C, 1)
        If TmpStr1 = "-NA-" Then
          LstItem(C).Caption = "----- No Item Available -----"
          Else
          LstItem(C).Caption = "   " & TmpStr1
        End If
         LstItem(C).Refresh
      Next C
End Sub
