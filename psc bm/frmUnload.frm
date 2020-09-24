VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1005
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   3705
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   3750
      Top             =   180
   End
   Begin VB.PictureBox picLoader 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   5
      Left            =   2717
      ScaleHeight     =   330
      ScaleWidth      =   285
      TabIndex        =   6
      Top             =   495
      Width           =   285
   End
   Begin VB.PictureBox picLoader 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   4
      Left            =   2312
      ScaleHeight     =   330
      ScaleWidth      =   285
      TabIndex        =   5
      Top             =   495
      Width           =   285
   End
   Begin VB.PictureBox picLoader 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   3
      Left            =   1907
      ScaleHeight     =   330
      ScaleWidth      =   285
      TabIndex        =   4
      Top             =   495
      Width           =   285
   End
   Begin VB.PictureBox picLoader 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   2
      Left            =   1502
      ScaleHeight     =   330
      ScaleWidth      =   285
      TabIndex        =   3
      Top             =   495
      Width           =   285
   End
   Begin VB.PictureBox picLoader 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   1
      Left            =   1097
      ScaleHeight     =   330
      ScaleWidth      =   285
      TabIndex        =   2
      Top             =   495
      Width           =   285
   End
   Begin VB.PictureBox picLoader 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   0
      Left            =   703
      ScaleHeight     =   330
      ScaleWidth      =   285
      TabIndex        =   1
      Top             =   495
      Width           =   285
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C000&
      Height          =   1005
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   3705
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      Height          =   420
      Left            =   570
      Shape           =   4  'Rounded Rectangle
      Top             =   450
      Width           =   2565
   End
   Begin VB.Label lblUnload 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Game... Please Wait"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   465
      TabIndex        =   0
      Top             =   135
      Width           =   2790
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
Timer1.Enabled = False
picLoader(0).BackColor = &HFF00&
 Load frmSetup
    Sleep 400
picLoader(1).BackColor = &HFF00&
 Load frmMsg
 Unload frmGameType
    Sleep 400
picLoader(2).BackColor = &HFF00&
 Load frmRes
    Sleep 250
picLoader(3).BackColor = &HFF00&
 Load frmMain
    Sleep 400
picLoader(4).BackColor = &HFF00&
 frmSetup.picAdOps.Picture = frmMsg.Picture
 Load frmHelp
    Sleep 400
picLoader(5).BackColor = &HFF00&
frmLoad.Refresh
Sleep 400
frmSetup.Show
frmSetup.Timer1.Enabled = True
frmLoad.Hide
Unload frmLoad
End Sub
