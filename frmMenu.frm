VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2628
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   2352
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   219
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   196
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   5080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   5360
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order form"
      Height          =   192
      Left            =   600
      TabIndex        =   6
      Top             =   1680
      Width           =   756
   End
   Begin VB.Image Image5 
      Height          =   384
      Left            =   120
      Picture         =   "frmMenu.frx":0000
      Top             =   1560
      Width           =   384
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hint and Strategy"
      Height          =   192
      Left            =   600
      TabIndex        =   5
      Top             =   1200
      Width           =   1224
   End
   Begin VB.Image Image4 
      Height          =   384
      Left            =   120
      Picture         =   "frmMenu.frx":0442
      Top             =   1080
      Width           =   384
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "About the Game"
      Height          =   192
      Left            =   600
      TabIndex        =   4
      Top             =   720
      Width           =   1152
   End
   Begin VB.Image Image3 
      Height          =   384
      Left            =   120
      Picture         =   "frmMenu.frx":0884
      Top             =   600
      Width           =   384
   End
   Begin VB.Image Image2 
      Height          =   384
      Left            =   120
      Picture         =   "frmMenu.frx":0CC6
      Top             =   2160
      Width           =   384
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   120
      Picture         =   "frmMenu.frx":1108
      Top             =   120
      Width           =   384
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   184
      Y1              =   170
      Y2              =   170
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000007&
      X1              =   -10
      X2              =   174
      Y1              =   170
      Y2              =   170
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Close menu"
      Height          =   192
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   492
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comments and Suggestions"
      Height          =   384
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1176
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************
'© 2001 by Alexander Anikin
'for more my code samples visit:
'http://www.hotmix.kiev.ua
'*******************************

Private Sub Form_Activate()
For e = 30 To 2790 Step 3
Width = e
DoEvents
Next e


End Sub

Private Sub Form_Deactivate()
Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = 0
Label2.ForeColor = 0
Label3.ForeColor = 0
Label4.ForeColor = 0
Label5.ForeColor = 0
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = 0
Label2.ForeColor = 0
Label3.ForeColor = 0
Label4.ForeColor = 0
Label5.ForeColor = 0
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = 0
Label2.ForeColor = 0
Label3.ForeColor = 0
Label4.ForeColor = 0
Label5.ForeColor = 0
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = 0
Label2.ForeColor = 0
Label3.ForeColor = vbWhite
Label4.ForeColor = 0
Label5.ForeColor = 0
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = 0
Label2.ForeColor = 0
Label3.ForeColor = 0
Label4.ForeColor = vbWhite
Label5.ForeColor = 0
End Sub

Private Sub Label1_Click()
frmUser.Show
frmMenu.Hide
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbWhite
Label2.ForeColor = 0
Label3.ForeColor = 0
Label4.ForeColor = 0
Label5.ForeColor = 0
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbWhite
Label1.ForeColor = 0
Label3.ForeColor = 0
Label4.ForeColor = 0
Label5.ForeColor = 0
End Sub

Private Sub Label3_Click()
frmAbout.Show
frmMenu.Hide
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = 0
Label2.ForeColor = 0
Label3.ForeColor = vbWhite
Label4.ForeColor = 0
Label5.ForeColor = 0
End Sub

Private Sub Label4_Click()
frmMenu.Hide
ANS = MsgBox("     ---===Hint and Strategy===---     " & vbNewLine & "     ------------------------------------------------------------------     " & vbNewLine & "     1. Use Hotkeys, constantly pause the game     " & vbNewLine & "     2. Use Hotkeys, constantly switch the speed of the game.     ", vbOKOnly, "Hints")
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = 0
Label2.ForeColor = 0
Label3.ForeColor = 0
Label4.ForeColor = vbWhite
Label5.ForeColor = 0
End Sub

Private Sub Label5_Click()
frmOrder.Show
frmMenu.Hide
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = 0
Label2.ForeColor = 0
Label3.ForeColor = 0
Label4.ForeColor = 0
Label5.ForeColor = vbWhite
End Sub
