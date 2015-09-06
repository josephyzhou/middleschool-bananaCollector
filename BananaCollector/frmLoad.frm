VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   1890
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer22 
      Interval        =   9999
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer21 
      Interval        =   9500
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer20 
      Interval        =   9000
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer19 
      Interval        =   8500
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer18 
      Interval        =   8000
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer17 
      Interval        =   7500
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer16 
      Interval        =   7000
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer15 
      Interval        =   6500
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer14 
      Interval        =   6000
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer13 
      Interval        =   5500
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer12 
      Interval        =   5000
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer11 
      Interval        =   4500
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer10 
      Interval        =   4000
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer9 
      Interval        =   3500
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer8 
      Interval        =   3000
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer7 
      Interval        =   2500
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer6 
      Interval        =   2000
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer5 
      Interval        =   1500
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   2400
      Top             =   720
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   1920
      Top             =   1440
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   240
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   840
      Top             =   1440
   End
   Begin VB.Label lblSeconds 
      BackColor       =   &H00000000&
      Caption         =   "Seconds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   3720
      TabIndex        =   3
      Top             =   1200
      Width           =   1212
   End
   Begin VB.Label lblTime 
      BackColor       =   &H00000000&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   3240
      TabIndex        =   2
      Top             =   1200
      Width           =   372
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Time Remaining:"
      BeginProperty Font 
         Name            =   "Haettenschweiler"
         Size            =   16.5
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   1080
      TabIndex        =   1
      Top             =   1200
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   600
      MousePointer    =   4  'Icon
      TabIndex        =   0
      Top             =   480
      Width           =   4572
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmLoad.Left = (Screen.Width - frmLoad.Width) / 2
frmLoad.Top = (Screen.Height - frmLoad.Height) / 2

End Sub


Private Sub Timer1_Timer()
frmLoad.Hide
BHunter.Show
End Sub

Private Sub Timer10_Timer()
Label1.Caption = "Loading G"
Timer9.Enabled = False
End Sub

Private Sub Timer11_Timer()
Label1.Caption = "Loading Ga"
Timer10.Enabled = False
End Sub

Private Sub Timer12_Timer()
Label1.Caption = "Loading Gam"
Timer11.Enabled = False
End Sub

Private Sub Timer13_Timer()
Label1.Caption = "Loading Game"
Timer12.Enabled = False
End Sub

Private Sub Timer14_Timer()
Label1.Caption = "Loading Game ."
Timer13.Enabled = False
End Sub

Private Sub Timer15_Timer()
Label1.Caption = "Loading Game .."
Timer14.Enabled = False
End Sub

Private Sub Timer16_Timer()
Label1.Caption = "Loading Game ..."
Timer15.Enabled = False
End Sub

Private Sub Timer17_Timer()
Label1.Caption = "L"
Timer16.Enabled = False
End Sub

Private Sub Timer18_Timer()
Label1.Caption = "Lo"
Timer17.Enabled = False
End Sub

Private Sub Timer19_Timer()
Label1.Caption = "Loa"
Timer18.Enabled = False
End Sub

Private Sub Timer2_Timer()
lblTime.Caption = Val(lblTime.Caption) - 1
If Val(lblTime.Caption) = 0 Then
lblTime.Caption = Val(lblTime.Caption) - 0
End If
End Sub

Private Sub Timer20_Timer()
Label1.Caption = "Load"
Timer19.Enabled = False
End Sub

Private Sub Timer21_Timer()
Label1.Caption = "Loadi"
Timer20.Enabled = False
End Sub

Private Sub Timer22_Timer()
Label1.Caption = "Loadin"
Timer21.Enabled = False
End Sub

Private Sub Timer3_Timer()
Label1.Caption = "Lo"
End Sub

Private Sub Timer4_Timer()
Label1.Caption = "Loa"
Timer3.Enabled = False

End Sub

Private Sub Timer5_Timer()
Label1.Caption = "Load"
Timer4.Enabled = False
End Sub

Private Sub Timer6_Timer()
Label1.Caption = "Loadi"
Timer5.Enabled = False
End Sub

Private Sub Timer7_Timer()
Label1.Caption = "Loadin"
Timer6.Enabled = False
End Sub

Private Sub Timer8_Timer()
Label1.Caption = "Loading"
Timer7.Enabled = False
End Sub

Private Sub Timer9_Timer()
Label1.Caption = "Loading "
Timer8.Enabled = False
End Sub
