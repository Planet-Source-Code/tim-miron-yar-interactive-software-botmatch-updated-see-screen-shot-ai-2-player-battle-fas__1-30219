VERSION 5.00
Begin VB.Form frmGameType 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "BotMatch - Select Game Options"
   ClientHeight    =   5250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8145
   HasDC           =   0   'False
   LinkTopic       =   "BotMatch - Game Type"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGameType.frx":0000
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   543
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox BMletrot 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   19200
      Left            =   3285
      Picture         =   "frmGameType.frx":28DE
      ScaleHeight     =   1280
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   21
      Top             =   5100
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.PictureBox picAnim2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   5835
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   20
      Top             =   525
      Width           =   1920
   End
   Begin VB.Timer tmrAnim 
      Interval        =   60
      Left            =   7050
      Top             =   4905
   End
   Begin VB.PictureBox rotF 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   5400
      Picture         =   "frmGameType.frx":3805
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   7680
   End
   Begin VB.PictureBox picAnim 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2445
      Left            =   990
      Picture         =   "frmGameType.frx":6818
      ScaleHeight     =   2445
      ScaleWidth      =   2460
      TabIndex        =   17
      Top             =   2355
      Width           =   2460
   End
   Begin VB.Label lblMadeBy 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Tim Miron"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   5130
      TabIndex        =   23
      Top             =   135
      Width           =   1785
   End
   Begin VB.Image imgIcon 
      Height          =   1920
      Index           =   4
      Left            =   -885
      Picture         =   "frmGameType.frx":1A38A
      Top             =   1485
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image imgIcon 
      Height          =   1920
      Index           =   5
      Left            =   -960
      Picture         =   "frmGameType.frx":1C1AC
      Top             =   2550
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image imgIcon 
      Height          =   1920
      Index           =   3
      Left            =   -795
      Picture         =   "frmGameType.frx":1E238
      Top             =   2760
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image imgIcon 
      Height          =   1920
      Index           =   2
      Left            =   -1320
      Picture         =   "frmGameType.frx":20715
      Top             =   2790
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image imgIcon 
      Height          =   1920
      Index           =   1
      Left            =   -1305
      Picture         =   "frmGameType.frx":22F94
      Top             =   3075
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Image imgIcon 
      Height          =   1920
      Index           =   0
      Left            =   -1380
      Picture         =   "frmGameType.frx":24ADF
      Top             =   2565
      Visible         =   0   'False
      Width           =   1920
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BotMatch V 1.0  - (c) Copyright 2001  yar interactive software"
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
      Height          =   405
      Left            =   15
      TabIndex        =   22
      Top             =   4830
      Width           =   2655
   End
   Begin VB.Shape shpDisp 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   3600
      Shape           =   3  'Circle
      Top             =   4080
      Width           =   135
   End
   Begin VB.Shape shpDisp 
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   6645
      Shape           =   3  'Circle
      Top             =   2850
      Width           =   135
   End
   Begin VB.Label cmdEsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   420
      Index           =   1
      Left            =   7080
      TabIndex        =   19
      Top             =   30
      Width           =   465
   End
   Begin VB.Shape shpEsc 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      FillColor       =   &H00002000&
      FillStyle       =   0  'Solid
      Height          =   390
      Index           =   1
      Left            =   7110
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   435
   End
   Begin VB.Label cmdEsc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   420
      Index           =   0
      Left            =   7590
      TabIndex        =   18
      Top             =   30
      Width           =   465
   End
   Begin VB.Shape shpEsc 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      FillColor       =   &H00002000&
      FillStyle       =   0  'Solid
      Height          =   390
      Index           =   0
      Left            =   7605
      Shape           =   4  'Rounded Rectangle
      Top             =   75
      Width           =   435
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H00008000&
      Index           =   1
      X1              =   447
      X2              =   447
      Y1              =   197
      Y2              =   216
   End
   Begin VB.Label lblOptVal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P1 vs. Computer"
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
      Index           =   3
      Left            =   3225
      TabIndex        =   15
      Top             =   1125
      Width           =   1545
   End
   Begin VB.Label lblOptVal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Easy"
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
      Index           =   4
      Left            =   3225
      TabIndex        =   14
      Top             =   1500
      Width           =   435
   End
   Begin VB.Label lblOptVal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Low"
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
      Index           =   5
      Left            =   3225
      TabIndex        =   13
      Top             =   1875
      Width           =   405
   End
   Begin VB.Label lblOpt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Graphics Level:"
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
      Index           =   5
      Left            =   75
      TabIndex        =   12
      Top             =   1875
      Width           =   2535
   End
   Begin VB.Label lblOpt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Skill Level:"
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
      Index           =   4
      Left            =   75
      TabIndex        =   11
      Top             =   1500
      Width           =   2145
   End
   Begin VB.Label lblOpt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Game Type:"
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
      Index           =   3
      Left            =   90
      TabIndex        =   10
      Top             =   1125
      Width           =   2220
   End
   Begin VB.Label lblOpt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Load Saved Bot"
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
      Index           =   2
      Left            =   105
      TabIndex        =   9
      Top             =   750
      Width           =   2490
   End
   Begin VB.Label lblOpt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.yarinteractive.com"
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
      Index           =   1
      Left            =   75
      TabIndex        =   8
      Top             =   375
      Width           =   3750
   End
   Begin VB.Label lblOpt 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Exit Game"
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
      Index           =   0
      Left            =   75
      TabIndex        =   7
      Top             =   15
      Width           =   2070
   End
   Begin VB.Line linOpt 
      BorderColor     =   &H0000FFFF&
      Index           =   5
      Visible         =   0   'False
      X1              =   68
      X2              =   87
      Y1              =   130
      Y2              =   130
   End
   Begin VB.Line linOpt 
      BorderColor     =   &H0000FFFF&
      Index           =   4
      Visible         =   0   'False
      X1              =   68
      X2              =   87
      Y1              =   105
      Y2              =   105
   End
   Begin VB.Line linOpt 
      BorderColor     =   &H0000FFFF&
      Index           =   3
      Visible         =   0   'False
      X1              =   68
      X2              =   87
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Line linOpt 
      BorderColor     =   &H0000FFFF&
      Index           =   2
      Visible         =   0   'False
      X1              =   68
      X2              =   87
      Y1              =   55
      Y2              =   55
   End
   Begin VB.Line linOpt 
      BorderColor     =   &H0000FFFF&
      Index           =   1
      Visible         =   0   'False
      X1              =   68
      X2              =   87
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Line linOpt 
      BorderColor     =   &H0000FFFF&
      Index           =   0
      Visible         =   0   'False
      X1              =   68
      X2              =   87
      Y1              =   6
      Y2              =   6
   End
   Begin VB.Shape shpOpt 
      BorderColor     =   &H0000FFFF&
      Height          =   180
      Index           =   5
      Left            =   75
      Top             =   1860
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Shape shpOpt 
      BorderColor     =   &H0000FFFF&
      Height          =   180
      Index           =   4
      Left            =   75
      Top             =   1485
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Shape shpOpt 
      BorderColor     =   &H0000FFFF&
      Height          =   180
      Index           =   3
      Left            =   75
      Top             =   1110
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Shape shpOpt 
      BorderColor     =   &H0000FFFF&
      Height          =   180
      Index           =   2
      Left            =   75
      Top             =   735
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Shape shpOpt 
      BorderColor     =   &H0000FFFF&
      Height          =   180
      Index           =   1
      Left            =   75
      Top             =   360
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Shape shpOpt 
      BorderColor     =   &H0000FFFF&
      Height          =   180
      Index           =   0
      Left            =   75
      Top             =   0
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblBut 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   630
      Index           =   6
      Left            =   5820
      TabIndex        =   6
      Top             =   3855
      Width           =   1770
   End
   Begin VB.Label lblBut 
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   5
      Left            =   5880
      TabIndex        =   5
      Top             =   4365
      Width           =   1665
   End
   Begin VB.Label lblBut 
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   4
      Left            =   6045
      TabIndex        =   4
      Top             =   4575
      Width           =   1335
   End
   Begin VB.Label lblBut 
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   3
      Left            =   6255
      TabIndex        =   3
      Top             =   4785
      Width           =   900
   End
   Begin VB.Label lblBut 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   2
      Left            =   5910
      TabIndex        =   2
      Top             =   3720
      Width           =   1605
   End
   Begin VB.Label lblBut 
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   1
      Left            =   6045
      TabIndex        =   1
      Top             =   3510
      Width           =   1335
   End
   Begin VB.Label lblBut 
      BackStyle       =   0  'Transparent
      Height          =   210
      Index           =   0
      Left            =   6270
      TabIndex        =   0
      Top             =   3300
      Width           =   870
   End
   Begin VB.Shape shpOkBut 
      BorderColor     =   &H00008000&
      BorderWidth     =   4
      FillColor       =   &H00002000&
      FillStyle       =   0  'Solid
      Height          =   1755
      Left            =   5835
      Shape           =   3  'Circle
      Top             =   3270
      Width           =   1755
   End
   Begin VB.Line linDisp 
      BorderColor     =   &H00008000&
      Index           =   0
      X1              =   391
      X2              =   246
      Y1              =   276
      Y2              =   276
   End
End
Attribute VB_Name = "frmGameType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private GTAnimFrame As Byte 'animation frame for rotating sphere
Private RotAnimFrame As Byte 'animation fram for rotating letters
Private DoRotBM As Boolean 'do animation of rotating "BM"

Private p1team As Byte 'player's team...

Private Sub cmdEsc_Click(Index As Integer)
 'exit (X) and minimize (_) buttons
 Select Case Index
  Case 0
       MsgStat = 5 'prompt for exit...
       ShowMsg "You are about to exit the program", "BotMatch - Exit Game?"
  Case 1
  
  'minimize the window...
   frmGameType.WindowState = 1
   End Select
End Sub

Private Sub cmdEsc_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'exit (X) and minimize (_) buttons Hover effect...
If cmdEsc(Index).ForeColor = &HC000& Then
   resGTButs
    shpEsc(Index).BorderColor = &HFF00&
    cmdEsc(Index).ForeColor = &HC0FFC0
End If
End Sub

Private Sub Form_Load()
DoRotBM = True 'start rotating "BM" animation
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
resGTButs 'reset graphics, animations ect.
End Sub

Private Sub lblBut_Click(Index As Integer)
'Accept button...

'if they have a 2 player game selected,
'make options on 2-player form match
frmGameType.Hide
Load frmLoad
frmLoad.Show
frmLoad.Timer1.Enabled = True

'check graphics settings
Select Case HGraphics
 Case False
 frmSetup.OPLowG.Value = True
 Case True
 frmSetup.opHighG.Value = True
End Select
End Sub

Private Sub lblBut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If shpOkBut.BorderColor = &H8000& Then
 resGTButs 'reset Game-Type button colors
 shpOkBut.BorderColor = &HFF00&
 lblBut(6).ForeColor = &HC0FFC0
End If
End Sub

Public Sub resGTButs()
Dim C As Byte
If shpOkBut.BorderColor = &H8000& And _
lblBut(6).ForeColor = &HC000& And _
shpOpt(0).Visible = False And shpOpt(1).Visible = False _
And shpOpt(2).Visible = False _
And shpOpt(3).Visible = False _
And shpOpt(4).Visible = False _
And shpOpt(5).Visible = False _
And tmrAnim.Enabled = True _
And lblOptVal(3).ForeColor = &HE0E0E0 _
And lblOptVal(4).ForeColor = &HE0E0E0 _
And lblOptVal(5).ForeColor = &HE0E0E0 _
And cmdEsc(0).ForeColor = &HC000& _
And cmdEsc(1).ForeColor = &HC000& _
And shpEsc(0).BorderColor = &HC000& _
And shpEsc(1).BorderColor = &HC000& _
And DoRotBM = True Then Exit Sub

For C = 0 To 5
 shpOpt(C).Visible = False
 lblOpt(C).ForeColor = &HC000&
 linOpt(C).Visible = False
Next C

cmdEsc(0).ForeColor = &HC000&
cmdEsc(1).ForeColor = &HC000&
shpEsc(0).BorderColor = &HC000&
shpEsc(1).BorderColor = &HC000&

lblOptVal(3).ForeColor = &HE0E0E0
lblOptVal(4).ForeColor = &HE0E0E0
lblOptVal(5).ForeColor = &HE0E0E0

shpOkBut.BorderColor = &H8000&
lblBut(6).ForeColor = &HC000&

DoRotBM = True
picAnim2.Picture = Nothing
End Sub


Private Sub lblOpt_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0
'exit prompt
       MsgStat = 5 'prompt for exit...
       ShowMsg "You are about to exit the program", "BotMatch - Exit Game?"
Case 1
  Shell "Start http://www.yarinteractive.com"
Case 2
  MsgBox "Sorry, this function isn't available in this version.", vbInformation, "Not Available..."
Case 3
    Call ToggleGameType
Case 4
    Call ToggleSkill
Case 5
    Call SetGraphicsLvl
End Select
End Sub

Private Sub lblOpt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If lblOpt(Index).ForeColor = &HC000& Then
   resGTButs
   lblOpt(Index).ForeColor = &H80FFFF
   shpOpt(Index).Visible = True
   linOpt(Index).Visible = True
   
   lblOptVal(Index).ForeColor = &HFFFF&
   DoRotBM = False
   picAnim2.Picture = imgIcon(Index).Picture
End If
End Sub

Private Sub picColor_Click()
If p1team = 0 Then
p1team = 1
picColor.BackColor = vbBlue
Else
p1team = 0
picColor.BackColor = vbRed
End If
End Sub


Private Sub tmrAnim_Timer()
picAnim.Cls
picAnim2.Cls

BitBlt picAnim.hdc, 0, 0, 128, 128, rotF.hdc, GTAnimFrame * 128, 0, vbSrcPaint
If DoRotBM = True Then
    BitBlt picAnim2.hdc, 0, 0, 128, 128, _
    BMletrot.hdc, 0, RotAnimFrame * 128, vbSrcCopy
End If
    RotAnimFrame = RotAnimFrame + 1
    GTAnimFrame = GTAnimFrame + 1

        If RotAnimFrame = 10 Then RotAnimFrame = 0
        If GTAnimFrame = 4 Then GTAnimFrame = 0
End Sub

Public Sub ToggleGameType()
'toggle type of game...
If MPGameType = False Then
    MPGameType = True
    lblOptVal(3).Caption = "P1 vs. P2"
    Else
    MPGameType = False
    lblOptVal(3).Caption = "P1 vs. Computer"
End If
End Sub

Public Sub ToggleSkill()
'toggle game skill level
If HighSkill = True Then
   HighSkill = False 'low skill level
   DoHG = True
   lblOptVal(4).Caption = "Easy"

        Else         'high skill level
        HighSkill = True
        DoHG = False
        lblOptVal(4).Caption = "Hard"
    End If
End Sub

Public Sub SetGraphicsLvl()
'toggle graphics level...
If HGraphics = False Then
          HGraphics = True
            ShowFPS = False
            ShowHUD = True
            ShowRetro = True
            HGTE = True
            
            
        GridOn = True
        
    With frmSetup
        .chkGrid.Value = 1
        .chkAutoH.Value = 0
        .chkFPS.Value = 0
        .chkHUD.Value = 1
        .chkRetro.Value = 1
        .chkHGT.Value = 1
        .opHighG.Value = True
    End With
    
    lblOptVal(5).Caption = "High"
        
        ElseIf HGraphics = True Then
        HGraphics = False
    lblOptVal(5).Caption = "Low"
                With frmMain
                    .Refresh
                    .Cls
                    .Picture = Nothing
                End With
            GridOn = False
        
        With frmSetup
             .chkGrid.Value = 0
             .chkAutoH.Value = 0
             .chkFPS.Value = 0
             .chkHUD.Value = 1
             .chkRetro.Value = 0
             .chkHGT.Value = 0
             .OPLowG.Value = True
        End With
End If
End Sub
