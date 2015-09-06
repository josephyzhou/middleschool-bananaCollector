VERSION 5.00
Begin VB.Form BHunter 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Banana Collector"
   ClientHeight    =   4272
   ClientLeft      =   48
   ClientTop       =   624
   ClientWidth     =   6204
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4272
   ScaleWidth      =   6204
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2280
      Top             =   3360
   End
   Begin VB.Timer Randomizer 
      Interval        =   10
      Left            =   3720
      Top             =   3360
   End
   Begin VB.Timer CheckEat 
      Left            =   4200
      Top             =   1920
   End
   Begin VB.Timer ScoreMove 
      Left            =   4200
      Top             =   2400
   End
   Begin VB.Timer ManMove 
      Left            =   4200
      Top             =   3360
   End
   Begin VB.Timer BombMove 
      Left            =   4200
      Top             =   2880
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pop up Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   252
      Left            =   4920
      MouseIcon       =   "BananaHunter.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   2040
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Hotkey List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   252
      Left            =   4920
      MouseIcon       =   "BananaHunter.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Label lblMyname 
      BackColor       =   &H00FF8080&
      Caption         =   "           Banana Collector                           Yufei Zhou                      yufeizhou88@hotmail.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   1080
      TabIndex        =   2
      Top             =   1560
      Width           =   2988
   End
   Begin VB.Image Smoke 
      Height          =   720
      Left            =   2760
      Picture         =   "BananaHunter.frx":0884
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Image Time 
      Height          =   384
      Left            =   1560
      Picture         =   "BananaHunter.frx":0CC6
      Top             =   2880
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Food 
      Height          =   384
      Left            =   600
      Picture         =   "BananaHunter.frx":1108
      Top             =   2880
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Heart 
      Height          =   384
      Left            =   1080
      Picture         =   "BananaHunter.frx":154A
      Top             =   2880
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Man_Happy 
      Height          =   384
      Left            =   2160
      Picture         =   "BananaHunter.frx":198C
      Top             =   2640
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Label ShowHiScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   4800
      TabIndex        =   15
      Top             =   1320
      Width           =   1092
   End
   Begin VB.Label lblHiScore 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HI-SCORE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   252
      Left            =   4956
      TabIndex        =   14
      Top             =   1080
      Width           =   1068
   End
   Begin VB.Image Man_Sad 
      Height          =   384
      Left            =   2160
      Picture         =   "BananaHunter.frx":1DCE
      Top             =   2640
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image icoMan 
      Height          =   240
      Left            =   5640
      Picture         =   "BananaHunter.frx":20D8
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   240
   End
   Begin VB.Image icoBomb 
      Height          =   240
      Left            =   5640
      Picture         =   "BananaHunter.frx":251A
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   240
   End
   Begin VB.Image icoBanana 
      Height          =   240
      Left            =   5640
      Picture         =   "BananaHunter.frx":295C
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   240
   End
   Begin VB.Label ShowBanana 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5040
      TabIndex        =   13
      Top             =   3240
      Width           =   612
   End
   Begin VB.Label ShowBomb 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5040
      TabIndex        =   12
      Top             =   3840
      Width           =   612
   End
   Begin VB.Label lblBanana 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BANANAs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   252
      Left            =   4800
      TabIndex        =   11
      Top             =   3000
      Width           =   1212
   End
   Begin VB.Label lblBomb 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BOMBs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   372
      Left            =   4800
      TabIndex        =   10
      Top             =   3600
      Width           =   1212
   End
   Begin VB.Label ShowLives 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   5040
      TabIndex        =   8
      Top             =   2760
      Width           =   612
   End
   Begin VB.Label lblLives 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LIVEs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   252
      Left            =   4800
      TabIndex        =   7
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Label ShowScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4800
      TabIndex        =   6
      Top             =   840
      Width           =   1092
   End
   Begin VB.Label LblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SCORE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   4800
      TabIndex        =   5
      Top             =   600
      Width           =   1212
   End
   Begin VB.Label ShowLevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4800
      TabIndex        =   4
      Top             =   360
      Width           =   1092
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LEVEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1212
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   23
      Left            =   4080
      Picture         =   "BananaHunter.frx":2D9E
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   22
      Left            =   3600
      Picture         =   "BananaHunter.frx":31E0
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   21
      Left            =   3120
      Picture         =   "BananaHunter.frx":3622
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   20
      Left            =   2640
      Picture         =   "BananaHunter.frx":3A64
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   19
      Left            =   2160
      Picture         =   "BananaHunter.frx":3EA6
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   18
      Left            =   1680
      Picture         =   "BananaHunter.frx":42E8
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   17
      Left            =   1200
      Picture         =   "BananaHunter.frx":472A
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   16
      Left            =   720
      Picture         =   "BananaHunter.frx":4B6C
      Top             =   1560
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   15
      Left            =   4080
      Picture         =   "BananaHunter.frx":4FAE
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   14
      Left            =   3600
      Picture         =   "BananaHunter.frx":53F0
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   13
      Left            =   3120
      Picture         =   "BananaHunter.frx":5832
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   12
      Left            =   2640
      Picture         =   "BananaHunter.frx":5C74
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   11
      Left            =   2160
      Picture         =   "BananaHunter.frx":60B6
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   10
      Left            =   1680
      Picture         =   "BananaHunter.frx":64F8
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   9
      Left            =   1200
      Picture         =   "BananaHunter.frx":693A
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   8
      Left            =   720
      Picture         =   "BananaHunter.frx":6D7C
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   7
      Left            =   4080
      Picture         =   "BananaHunter.frx":71BE
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   6
      Left            =   3600
      Picture         =   "BananaHunter.frx":7600
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   5
      Left            =   3120
      Picture         =   "BananaHunter.frx":7A42
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   4
      Left            =   2640
      Picture         =   "BananaHunter.frx":7E84
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   3
      Left            =   2160
      Picture         =   "BananaHunter.frx":82C6
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   2
      Left            =   1680
      Picture         =   "BananaHunter.frx":8708
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   1
      Left            =   1200
      Picture         =   "BananaHunter.frx":8B4A
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Bomb 
      Height          =   384
      Index           =   0
      Left            =   720
      Picture         =   "BananaHunter.frx":8F8C
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Man 
      Height          =   384
      Left            =   120
      Picture         =   "BananaHunter.frx":93CE
      Top             =   360
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Banana 
      Height          =   384
      Index           =   7
      Left            =   4080
      Picture         =   "BananaHunter.frx":9810
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Banana 
      Height          =   384
      Index           =   6
      Left            =   3600
      Picture         =   "BananaHunter.frx":9C52
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Banana 
      Height          =   384
      Index           =   5
      Left            =   3120
      Picture         =   "BananaHunter.frx":A094
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Banana 
      Height          =   384
      Index           =   4
      Left            =   2640
      Picture         =   "BananaHunter.frx":A4D6
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Banana 
      Height          =   384
      Index           =   3
      Left            =   2160
      Picture         =   "BananaHunter.frx":A918
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Banana 
      Height          =   384
      Index           =   2
      Left            =   1680
      Picture         =   "BananaHunter.frx":AD5A
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Banana 
      Height          =   384
      Index           =   1
      Left            =   1200
      Picture         =   "BananaHunter.frx":B19C
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image Banana 
      Height          =   384
      Index           =   0
      Left            =   720
      Picture         =   "BananaHunter.frx":B5DE
      Top             =   120
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Label lblGuide 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3960
      Width           =   4695
   End
   Begin VB.Label lblProtection 
      BackColor       =   &H008080FF&
      Caption         =   "PROTECTION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   50
      TabIndex        =   9
      Top             =   50
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label lblField 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu mnGame 
      Caption         =   "&Game"
      Begin VB.Menu mnNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnPause 
         Caption         =   "&Pause"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnQuit 
         Caption         =   "&Quit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnSpeed 
      Caption         =   "&Speed"
      Begin VB.Menu mnFast 
         Caption         =   "&Fast"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnMedium 
         Caption         =   "&Medium"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnSlow 
         Caption         =   "&Slow"
         Shortcut        =   {F7}
      End
   End
   Begin VB.Menu mnAbout 
      Caption         =   "&About"
      Begin VB.Menu mnHow 
         Caption         =   "&How to Play"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnAbout2 
         Caption         =   "&About Banana Hunter"
      End
   End
End
Attribute VB_Name = "BHunter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ANS As Variant

Dim Lives As Integer        'Labels value
Dim Level As Integer
Dim Score As Integer
Dim Banana_Left As Integer
Dim Bomb_Left As Integer
Dim HiScore As Integer

Dim Banana_Num As Integer   'Banana/bomb Counter
Dim Bomb_Num As Integer

Dim Man_Prot As Boolean     'Man Values
Dim Man_Dir_S As Boolean
Dim Man_Dir_W As Boolean
Dim Man_Dir_E As Boolean
Dim Man_Dir_N As Boolean

Dim Man_Speed As Integer

Dim Bomb_Speed As Integer    'Bomb Values
Dim Bomb_Dir(23) As Integer

Dim Bonus(2) As Integer

Dim X As Integer

Private Sub Form_Load()
lblMyname.Visible = True
Randomizer.Interval = 10    'Game random number
lblGuide.Caption = "Welcome to Banana Collector."
End Sub

Private Sub Label1_Click()
MsgBox ("     ---===Hot key list===---     " & vbNewLine & "     -------------------------------     " & vbNewLine & "     F1 - How to play     " & vbNewLine & "     F2 - New Game     " & vbNewLine & "     F3 - Pause     " & vbNewLine & "     F4 - Quit     " & vbNewLine & "     F5 - Fast     " & vbNewLine & "     F6 - Medium     " & vbNewLine & "     F7 - Slow     ")

End Sub



Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGuide.Caption = "Click here to get the complete Hotkeys List."
End Sub

Private Sub Label2_Click()
frmMenu.Show
frmMenu.Left = BHunter.Left + 1000
frmMenu.Top = BHunter.Top + 1800

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblGuide.Caption = "Click here to show the pop up menu"

End Sub

Private Sub mnAbout2_Click() 'Menu About
Call TimeStop
lblMyname.Visible = False
frmAbout.Show
End Sub

Private Sub mnHow_Click()   '(F1) Menu How to PLay
Call TimeStop
lblMyname.Visible = False
frmHow.Show
BHunter.Hide
End Sub

Private Sub mnFast_Click()  'Fast Speed
Man_Speed = 80
Bomb_Speed = 40
mnFast.Checked = True
mnSlow.Checked = False
mnMedium.Checked = False
lblGuide.Caption = "Game Speed: Fast "

End Sub

Private Sub mnMedium_Click() 'Medium Speed
Man_Speed = 50
Bomb_Speed = 25
mnMedium.Checked = True
mnSlow.Checked = False
mnFast.Checked = False
lblGuide.Caption = "Game Speed: Midium "

End Sub

Private Sub mnSlow_Click()  'Slow Speed
Man_Speed = 20
Bomb_Speed = 10
mnSlow.Checked = True
mnFast.Checked = False
mnMedium.Checked = False
lblGuide.Caption = "Game Speed: Slow "

End Sub

Private Sub mnPause_Click()    '(F3) Pauses the Game...
Call TimeStop
lblGuide.Caption = "Game Paused..."
ANS = MsgBox("Click OK to Continue", vbOKOnly, "PAUSED")
Call TimeStart
lblGuide.Caption = "Game Continues..."
End Sub

Private Sub mnQuit_Click()  '(F4)Menu Quit
Call TimeStop
End
End Sub

Private Sub mnNew_Click()   '(F2) New Game
lblGuide.Caption = "Starting a new Game, current high score: " & Val(ShowHiScore.Caption)
lblMyname.Visible = False
Man_Sad.Visible = False
Man_Happy.Visible = False
Randomizer.Interval = 0

mnMedium.Checked = True

TimeStop
ANS = MsgBox("Click OK to Start Game", vbOKOnly, "New Game")

Man_Speed = 50
Bomb_Speed = 25

Level = 0       'Resets all values
Lives = 5
Score = 0

Call New_Level

End Sub

Private Sub New_Level()

Man_Dir_N = False
Man_Dir_S = False
Man_Dir_E = False
Man_Dir_W = False

Man_Prot = True
Level = Level + 1               '+ Level

Banana_Num = 8                  'Sets # of banana
Bomb_Num = 2 + (Level / 2)      'Sets # of bombs

If Bomb_Num >= 24 Then
    Bomb_Num = 24
End If

ManMove.Interval = 30   'Sets Timers
BombMove.Interval = 30
CheckEat.Interval = 15
ScoreMove.Interval = 100

Banana_Left = Banana_Num    'Set display stats
Bomb_Left = Bomb_Num

Call ShowAll
TimeStart

End Sub

Private Sub Form_KeyUp(Key As Integer, Shift As Integer) 'Player Movements...
Select Case Key
    Case Is = 40
        Man_Dir_S = False
    Case Is = 37
        Man_Dir_W = False
    Case Is = 39
        Man_Dir_E = False
    Case Is = 38
        Man_Dir_N = False
End Select
End Sub
Private Sub Form_KeyDown(Key As Integer, Shift As Integer)
Man_Prot = False

Select Case Key
    Case Is = 40
        Man_Dir_S = True
    Case Is = 37
        Man_Dir_W = True
    Case Is = 39
        Man_Dir_E = True
    Case Is = 38
        Man_Dir_N = True
End Select
End Sub

Private Sub ManMove_Timer() 'Man Timer

If Man_Prot = True Then     'The Protection Label
lblProtection.Visible = True
Else: lblProtection.Visible = False
End If

'Man Moving .....................................
    If Man_Dir_S = True Then
        Man.Top = Man.Top + Man_Speed
    End If

    If Man_Dir_W = True Then
        Man.Left = Man.Left - Man_Speed
    End If
    
    If Man_Dir_E = True Then
        Man.Left = Man.Left + Man_Speed
    End If

    If Man_Dir_N = True Then
        Man.Top = Man.Top - Man_Speed
    End If
    

'Man Hits wall..........................
    If Man.Top <= 0 Then
        Man.Top = 1
        Man_Dir_N = False
    End If

    If Man.Left <= 0 Then
        Man.Left = 1
        Man_Dir_W = False
    End If
    
    If Man.Top >= 3360 Then
        Man.Top = 3359
        Man_Dir_S = False
    End If
    
    If Man.Left >= 4200 Then
        Man.Left = 4199
        Man_Dir_S = False
    End If

End Sub

Private Sub CheckEat_Timer()    'Checks for Man over Banana & ETC

    'Man eats Banana
For X = 0 To 7
    If Man_Prot = False And Banana.Item(X).Visible = True And Man.Top >= (Banana.Item(X).Top - 360) And Man.Top <= (Banana.Item(X).Top + 360) And Man.Left >= (Banana.Item(X).Left - 360) And Man.Left <= (Banana.Item(X).Left + 360) Then
        Banana.Item(X).Visible = False
        Score = Score + 100
        Banana_Left = Banana_Left - 1
    End If
Next X

'Bonus Items (HEARTS)
If Man_Prot = False And Heart.Visible = True And Man.Top >= (Heart.Top - 360) And Man.Top <= (Heart.Top + 360) And Man.Left >= (Heart.Left - 360) And Man.Left <= (Heart.Left + 360) Then
    Heart.Visible = False
    Lives = Lives + 1
End If
'Bonus Items (TIME)
If Man_Prot = False And Time.Visible = True And Man.Top >= (Time.Top - 360) And Man.Top <= (Time.Top + 360) And Man.Left >= (Time.Left - 360) And Man.Left <= (Time.Left + 360) Then
    Time.Visible = False
    BombMove.Interval = 200
End If
'Bonus Items (FOOD)
If Man_Prot = False And Food.Visible = True And Man.Top >= (Food.Top - 360) And Man.Top <= (Food.Top + 360) And Man.Left >= (Food.Left - 360) And Man.Left <= (Food.Left + 360) Then
    Food.Visible = False
    Score = Score + 800
End If

    'Banana Finishes
If Level >= 1 And Banana.Item(0).Visible = False And Banana.Item(1).Visible = False And Banana.Item(2).Visible = False And Banana.Item(3).Visible = False And Banana.Item(4).Visible = False And Banana.Item(5).Visible = False And Banana.Item(6).Visible = False And Banana.Item(7).Visible = False And Lives >= 1 Then
    Call TimeStop
    Call HideAll
    Randomizer.Interval = 1
    ANS = MsgBox("Proceed to next level...", vbOKOnly, "Level Cleared")
    Randomizer.Interval = 0
    Call New_Level
End If
End Sub

Private Sub BombMove_Timer()

'Bomb Movement and Bounce direction.....................
For X = 0 To (Bomb_Num - 1)
    If Bomb_Dir(X) = 1 Then
        Bomb.Item(X).Top = Bomb.Item(X).Top + Bomb_Speed
        Bomb.Item(X).Left = Bomb.Item(X).Left - Bomb_Speed
    End If
    
    If Bomb_Dir(X) = 2 Then
        Bomb.Item(X).Top = Bomb.Item(X).Top + Bomb_Speed
    End If
    
    If Bomb_Dir(X) = 3 Then
        Bomb.Item(X).Top = Bomb.Item(X).Top + Bomb_Speed
        Bomb.Item(X).Left = Bomb.Item(X).Left + Bomb_Speed
    End If
    
    If Bomb_Dir(X) = 4 Then
        Bomb.Item(X).Left = Bomb.Item(X).Left - Bomb_Speed
    End If
    
    If Bomb_Dir(X) = 6 Then
        Bomb.Item(X).Left = Bomb.Item(X).Left + Bomb_Speed
    End If
    
    If Bomb_Dir(X) = 7 Then
        Bomb.Item(X).Top = Bomb.Item(X).Top - Bomb_Speed
        Bomb.Item(X).Left = Bomb.Item(X).Left - Bomb_Speed
    End If
    
    If Bomb_Dir(X) = 8 Then
        Bomb.Item(X).Top = Bomb.Item(X).Top - Bomb_Speed
    End If
    
    If Bomb_Dir(X) = 9 Then
        Bomb.Item(X).Top = Bomb.Item(X).Top - Bomb_Speed
        Bomb.Item(X).Left = Bomb.Item(X).Left + Bomb_Speed
    End If

'Bounce Direction...
    If Bomb_Dir(X) = 1 Then
        If Bomb.Item(X).Top >= 3360 Then
           Bomb.Item(X).Top = 3360
           Bomb_Dir(X) = 7
        End If
        If Bomb.Item(X).Left <= 0 Then
           Bomb.Item(X).Left = 0
           Bomb_Dir(X) = 3
        End If
    End If

    If Bomb_Dir(X) = 2 Then
        If Bomb.Item(X).Top >= 3360 Then
           Bomb.Item(X).Top = 3360
           Bomb_Dir(X) = 8
        End If
    End If

    If Bomb_Dir(X) = 3 Then
        If Bomb.Item(X).Top >= 3360 Then
           Bomb.Item(X).Top = 3360
           Bomb_Dir(X) = 9
        End If
        If Bomb.Item(X).Left >= 4200 Then
           Bomb.Item(X).Left = 4200
           Bomb_Dir(X) = 1
        End If
    End If

    If Bomb_Dir(X) = 4 Then
        If Bomb.Item(X).Left <= 0 Then
           Bomb.Item(X).Left = 0
           Bomb_Dir(X) = 6
        End If
    End If

    If Bomb_Dir(X) = 6 Then
        If Bomb.Item(X).Left >= 4200 Then
           Bomb.Item(X).Left = 4200
           Bomb_Dir(X) = 4
        End If
    End If

    If Bomb_Dir(X) = 7 Then
        If Bomb.Item(X).Top <= 0 Then
           Bomb.Item(X).Top = 0
           Bomb_Dir(X) = 1
        End If
        If Bomb.Item(X).Left <= 0 Then
           Bomb.Item(X).Left = 0
           Bomb_Dir(X) = 9
        End If
    End If

    If Bomb_Dir(X) = 8 Then
        If Bomb.Item(X).Top <= 0 Then
           Bomb.Item(X).Top = 0
           Bomb_Dir(X) = 2
        End If
    End If

    If Bomb_Dir(X) = 9 Then
        If Bomb.Item(X).Top <= 0 Then
           Bomb.Item(X).Top = 0
           Bomb_Dir(X) = 3
        End If
        If Bomb.Item(X).Left >= 4200 Then
           Bomb.Item(X).Left = 4200
           Bomb_Dir(X) = 7
        End If
    End If

'CHECKS Whether Bomb Hits Man ..................
'...............................................

If Man_Prot = False And Bomb.Item(X).Visible = True And Man.Top >= (Bomb.Item(X).Top - 360) And Man.Top <= (Bomb.Item(X).Top + 360) And Man.Left >= (Bomb.Item(X).Left - 360) And Man.Left <= (Bomb.Item(X).Left + 360) Then
    Lives = Lives - 1

    Call TimeStop
    Man.Visible = False
    Smoke.Top = Man.Top
    Smoke.Left = Man.Left
    Smoke.Visible = True
    ANS = MsgBox("You've been Hit!!" & vbNewLine & "Current lives:" & Val(ShowLives.Caption), vbOKOnly, "BooM!")
    Smoke.Visible = False
    Man.Visible = True
    Call TimeStart

    Man_Prot = True
    Man_Dir_S = False
    Man_Dir_W = False
    Man_Dir_E = False
    Man_Dir_N = False
End If

Next X

'If no more lives
If Lives <= 0 Then Call End_Game

End Sub

Private Sub ShowAll()      'Shows all items
Man.Visible = True

For X = 0 To 7              'Show banana
    Banana.Item(X).Visible = True
Next X

For X = 0 To (Bomb_Num - 1) 'Show Bomb
    Bomb.Item(X).Visible = True
Next X

If Bonus(0) >= 4 Then Food.Visible = True
If Bonus(1) >= 7 Then Time.Visible = True
If Bonus(2) = 9 Then Heart.Visible = True

End Sub

Private Sub HideAll()   'Hides all items
Man.Visible = False
For X = 0 To 7
    Banana.Item(X).Visible = False
Next X
For X = 0 To (Bomb_Num - 1)
    Bomb.Item(X).Visible = False
Next X
Heart.Visible = False
Food.Visible = False
Time.Visible = False
End Sub

Private Sub TimeStart() 'Resumes time...
ManMove.Enabled = True
BombMove.Enabled = True
CheckEat.Enabled = True
End Sub

Private Sub TimeStop()  'Stops Time...
ManMove.Enabled = False
BombMove.Enabled = False
CheckEat.Enabled = False
End Sub

Private Sub End_Game()  'Man no more lives
Call TimeStop
Call HideAll

Man_Prot = False

ANS = MsgBox("No more lives ... Try again!", vbOKOnly, "GAME OVER")

If Score > Val(ShowHiScore.Caption) Then
    Man_Happy.Visible = True
    ShowHiScore.Caption = Str(Score)
    ANS = MsgBox("You've made a New Hi-Score!", vbOKOnly, "Yes!!!")
Else:
    Man_Sad.Visible = True
End If

lblMyname.Visible = True    'Show my name
Randomizer.Interval = 10    'Start placing things again
End Sub

Private Sub Randomizer_Timer()
For X = 0 To 7              'Randomly places Banana
    Banana.Item(X).Top = (Rnd * 3360)
    Banana.Item(X).Left = (Rnd * 4200)
Next X

For X = 0 To 23 'Randomly Places Bomb
    Bomb.Item(X).Top = (Rnd * 3360)
    Bomb.Item(X).Left = (Rnd * 4200)
    Bomb_Dir(X) = ((Rnd * 8) + 1)
        If Bomb_Dir(X) = 5 Then Bomb_Dir(X) = 7
Next X

For X = 0 To 2  'Availability of BOnus!!
    Bonus(X) = (Rnd * 9)
    Heart.Top = (Rnd * 3360)
    Heart.Left = (Rnd * 4200)
    Food.Top = (Rnd * 3360)
    Food.Left = (Rnd * 4200)
    Time.Top = (Rnd * 3360)
    Time.Left = (Rnd * 4200)
Next X

End Sub

Private Sub ScoreMove_Timer()   'Updates Stats
ShowLives.Caption = Str(Lives)
ShowScore.Caption = Str(Score)
ShowLevel.Caption = Str(Level)
ShowBanana.Caption = Str(Banana_Left)
ShowBomb.Caption = Str(Bomb_Left)
End Sub

Private Sub Timer1_Timer()
If frmMenu.Visible = 0 Then
frmMenu.Left = BHunter.Left + 1000
frmMenu.Top = BHunter.Top + 1800
frmMenu.Show
Timer1.Enabled = False
Else
Unload frmMenu
Exit Sub
End If

End Sub
