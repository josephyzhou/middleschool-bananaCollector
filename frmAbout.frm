VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   2556
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4476
   LinkTopic       =   "Form1"
   ScaleHeight     =   2556
   ScaleWidth      =   4476
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label4 
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
      MouseIcon       =   "frmAbout.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2160
      Width           =   2052
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "yufeizhou88@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   840
      MouseIcon       =   "frmAbout.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1440
      Width           =   2772
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Banana Collector: Yufei Zhou"
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
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   3252
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   600
      X2              =   3840
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   480
      Y1              =   480
      Y2              =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "ABOUT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1332
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "USER32" () As Long
Private Declare Function SetWindowRgn Lib "USER32" _
(ByVal hWnd As Long, ByVal hRgn As Long, _
ByVal bRedraw As Boolean) As Long
Private Const SW_SHOWMAXIMIZED = 3
Private Sub Navigate(frm As Form, ByVal NavTo As String)
hBrowse = ShellExecute(frm.hWnd, "open", NavTo, "", "", SW_SHOWMAXIMIZED)
End Sub

Private Sub Form_Load()
frmAbout.Top = (Screen.Height - frmAbout.Height) / 2
frmAbout.Left = (Screen.Width - frmAbout.Width) / 2

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = &H80FF&
Label3.ForeColor = vbWhite

End Sub

Private Sub Label4_Click()
frmAbout.Hide
BHunter.Show
End Sub
Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbWhite
End Sub
Private Sub Label3_Click()
MousePointer = 11
Navigate Me, "mailto: yufeizhou88@hotmail.com"
MousePointer = 0

End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.ForeColor = &H80FF&

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  MousePointer = 15
  Call ReleaseCapture
  Call SendMessage(hWnd, &HA1, 2, 0&)
  MousePointer = 1
End Sub
  

