VERSION 5.00
Begin VB.Form frmRes 
   BackColor       =   &H00404040&
   ClientHeight    =   8010
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   HasDC           =   0   'False
   Icon            =   "frmResources.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   534
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   592
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picWeapon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   9390
      Picture         =   "frmResources.frx":000C
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   33
      Top             =   1140
      Width           =   480
   End
   Begin VB.PictureBox SCDbuffer2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2490
      ScaleHeight     =   240
      ScaleWidth      =   480
      TabIndex        =   32
      Top             =   4605
      Width           =   480
   End
   Begin VB.PictureBox SCDbuffer1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2475
      ScaleHeight     =   240
      ScaleWidth      =   480
      TabIndex        =   31
      Top             =   4320
      Width           =   480
   End
   Begin VB.PictureBox picLazI 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   15
      Picture         =   "frmResources.frx":0448
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   29
      Top             =   3975
      Width           =   480
   End
   Begin VB.PictureBox picFPS 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   240
      Left            =   915
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   27
      Top             =   690
      Width           =   900
   End
   Begin VB.PictureBox picGo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   63.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   15
      ScaleHeight     =   96
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   26
      Top             =   2505
      Width           =   2400
   End
   Begin VB.PictureBox picPlasmaSpark 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   2460
      Picture         =   "frmResources.frx":108A
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   25
      Top             =   2520
      Width           =   960
   End
   Begin VB.PictureBox picSpark 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1545
      Picture         =   "frmResources.frx":16DA
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   24
      Top             =   60
      Width           =   240
   End
   Begin VB.PictureBox picPF4 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   5580
      Picture         =   "frmResources.frx":1A1C
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   138
      TabIndex        =   22
      Top             =   2805
      Width           =   2070
   End
   Begin VB.PictureBox picP2Frags 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CAC Lasko Even Weight"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   390
      Left            =   2505
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   134
      TabIndex        =   21
      Top             =   3900
      Width           =   2010
   End
   Begin VB.PictureBox picP1Frags 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "CAC Lasko Even Weight"
         Size            =   18
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   390
      Left            =   2430
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   134
      TabIndex        =   20
      Top             =   3495
      Width           =   2010
   End
   Begin VB.PictureBox picPAK 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   -15
      Picture         =   "frmResources.frx":1E54
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   19
      Top             =   510
      Width           =   2175
   End
   Begin VB.PictureBox picCS 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Left            =   5535
      Picture         =   "frmResources.frx":2299
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   143
      TabIndex        =   18
      Top             =   2535
      Width           =   2145
   End
   Begin VB.PictureBox picP2Health 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   420
      Left            =   0
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   17
      Top             =   705
      Width           =   915
   End
   Begin VB.PictureBox picP1Health 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   420
      Left            =   15
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   16
      Top             =   1125
      Width           =   915
   End
   Begin VB.PictureBox picP2Name 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   1335
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   15
      Top             =   975
      Width           =   1485
   End
   Begin VB.PictureBox picP1Name 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   3045
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   99
      TabIndex        =   14
      Top             =   960
      Width           =   1485
   End
   Begin VB.PictureBox PlasmaB 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   1
      Left            =   3540
      Picture         =   "frmResources.frx":3A7B
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   13
      Top             =   2505
      Width           =   960
   End
   Begin VB.PictureBox PlasmaB 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   960
      Index           =   0
      Left            =   4530
      Picture         =   "frmResources.frx":6ABD
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   12
      Top             =   2505
      Width           =   960
   End
   Begin VB.PictureBox cgBullet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   60
      Index           =   1
      Left            =   2895
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   11
      Top             =   1095
      Width           =   15
   End
   Begin VB.PictureBox cgBullet 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   60
      Index           =   0
      Left            =   4905
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   10
      Top             =   480
      Width           =   15
   End
   Begin VB.PictureBox picWeapon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   9885
      Picture         =   "frmResources.frx":9AFF
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   9
      Top             =   645
      Width           =   480
   End
   Begin VB.PictureBox picWeapon 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   9885
      Picture         =   "frmResources.frx":9F14
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   120
      Width           =   480
   End
   Begin VB.PictureBox picWeapon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   9900
      Picture         =   "frmResources.frx":9FFB
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   7
      Top             =   1140
      Width           =   480
   End
   Begin VB.PictureBox RPanel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   540
      Picture         =   "frmResources.frx":A3F1
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   6
      Top             =   4950
      Width           =   1920
   End
   Begin VB.PictureBox BPanel 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   510
      Picture         =   "frmResources.frx":AA86
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   5
      Top             =   3945
      Width           =   1920
   End
   Begin VB.PictureBox ybolt 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   405
      Index           =   1
      Left            =   1500
      Picture         =   "frmResources.frx":B11B
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   2
      TabIndex        =   4
      Top             =   60
      Width           =   30
   End
   Begin VB.PictureBox ybolt 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   405
      Index           =   0
      Left            =   1470
      Picture         =   "frmResources.frx":B235
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   2
      TabIndex        =   3
      Top             =   75
      Width           =   30
   End
   Begin VB.PictureBox CMask 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   975
      Picture         =   "frmResources.frx":B34F
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   2
      Top             =   45
      Width           =   480
   End
   Begin VB.PictureBox CRed 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   510
      Picture         =   "frmResources.frx":BF91
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   30
      Width           =   480
   End
   Begin VB.PictureBox CBlue 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   15
      Picture         =   "frmResources.frx":CBD3
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   15
      Width           =   480
   End
   Begin VB.PictureBox picRetroF1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   30
      Picture         =   "frmResources.frx":D815
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   28
      Top             =   1545
      Width           =   7695
   End
   Begin VB.PictureBox picGameHelp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   4245
      Left            =   4320
      Picture         =   "frmResources.frx":123E4
      ScaleHeight     =   283
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   303
      TabIndex        =   23
      Top             =   3765
      Width           =   4545
   End
   Begin VB.PictureBox picShields 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   1800
      Picture         =   "frmResources.frx":1397C
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   30
      Top             =   -45
      Width           =   3840
   End
End
Attribute VB_Name = "frmRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright 2001 - yar interactive

'this form holds all the game resources
'(images, buffers, ect.)
