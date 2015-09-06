VERSION 5.00
Begin VB.Form frmHow 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "How to play..."
   ClientHeight    =   6864
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4824
   LinkTopic       =   "Form1"
   ScaleHeight     =   6864
   ScaleWidth      =   4824
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Show keys"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   1812
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "keys"
      ForeColor       =   &H00FFFFFF&
      Height          =   2892
      Left            =   360
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   3972
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Click on the links for more detail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   252
         Left            =   600
         TabIndex        =   11
         Top             =   2520
         Width           =   2892
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "<< Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   372
         Left            =   2520
         MouseIcon       =   "frmHow.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1920
         Width           =   1092
      End
      Begin VB.Image lblhow2 
         Height          =   384
         Index           =   3
         Left            =   2040
         Picture         =   "frmHow.frx":0442
         Top             =   1920
         Width           =   384
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "<< Life"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   372
         Left            =   2520
         MouseIcon       =   "frmHow.frx":0884
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   1200
         Width           =   1092
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "<< Banana"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   372
         Left            =   2520
         MouseIcon       =   "frmHow.frx":0CC6
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   480
         Width           =   1332
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "<< Bomb"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   372
         Left            =   840
         MouseIcon       =   "frmHow.frx":1108
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1920
         Width           =   1092
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "<< Food"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   372
         Left            =   840
         MouseIcon       =   "frmHow.frx":154A
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   1200
         Width           =   1092
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "<< Player"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   372
         Left            =   840
         MouseIcon       =   "frmHow.frx":198C
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   480
         Width           =   1092
      End
      Begin VB.Image lblhow2 
         Height          =   360
         Index           =   2
         Left            =   2040
         Picture         =   "frmHow.frx":1DCE
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   360
      End
      Begin VB.Image lblhow2 
         Height          =   360
         Index           =   4
         Left            =   2040
         Picture         =   "frmHow.frx":2210
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Image lblhow2 
         Height          =   360
         Index           =   0
         Left            =   360
         Picture         =   "frmHow.frx":2652
         Stretch         =   -1  'True
         Top             =   480
         Width           =   360
      End
      Begin VB.Image lblhow2 
         Height          =   360
         Index           =   1
         Left            =   360
         Picture         =   "frmHow.frx":2A94
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   360
      End
      Begin VB.Image lblhow2 
         Height          =   360
         Index           =   5
         Left            =   360
         Picture         =   "frmHow.frx":2ED6
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   360
      End
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OK, Back to the Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   252
      Left            =   1200
      MouseIcon       =   "frmHow.frx":3318
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   6480
      Width           =   2052
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHow.frx":375A
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   2172
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   4452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "How to play??"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   19.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   612
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2172
   End
End
Attribute VB_Name = "frmHow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
Frame1.Visible = True
Else
If Check1.Value = False Then
Frame1.Visible = False
End If
End If

End Sub

Private Sub Form_DblClick()
frmMenu.Show
End Sub

Private Sub Form_Load()
frmHow.Top = (Screen.Height - frmHow.Height) / 2
frmHow.Left = (Screen.Width - frmHow.Width) / 2
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MousePointer = 15
  MousePointer = 1
End Sub
  


Private Sub Label3_Click()
ANS = MsgBox("This is the man!" & vbNewLine & "Use directional keys to move around", vbExclamation, "This is YOU!")
End Sub

Private Sub Label4_Click()

ANS = MsgBox("Food, worth 800 points each", vbExclamation, "food")
End Sub

Private Sub Label5_Click()

ANS = MsgBox("After a bomb hit you," & vbNewLine & "you will get protection if you don't move", vbExclamation, "Bomb")
End Sub

Private Sub Label6_Click()

ANS = MsgBox("You need to collect these bananas." & vbNewLine & "Each one worth 100 points!", vbExclamation, "banana")
End Sub

Private Sub Label7_Click()

ANS = MsgBox("One Extra Life!", vbExclamation, "Lives")
End Sub

Private Sub Label8_Click()

ANS = MsgBox("Very useful! They slow down the speed of the bomb!", vbExclamation, "Time is critical")
End Sub

Private Sub Label9_Click()
frmHow.Hide
BHunter.Show
End Sub
