VERSION 5.00
Begin VB.Form frmUser 
   BackColor       =   &H00FFFFFF&
   Caption         =   "User Profile"
   ClientHeight    =   4332
   ClientLeft      =   48
   ClientTop       =   420
   ClientWidth     =   5004
   LinkTopic       =   "Form1"
   ScaleHeight     =   4332
   ScaleWidth      =   5004
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Send your comments"
      Height          =   372
      Left            =   1440
      TabIndex        =   6
      Top             =   3720
      Width           =   1932
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1644
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1800
      Width           =   4572
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   3372
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2412
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1332
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   852
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   852
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim username As String
Dim Email As String
Dim suggestion As String

Private Sub Command1_Click()
username = Text1.Text
Email = Text2.Text


ANS = MsgBox("Thank you for your suggestion, " & username & "." & vbNewLine & "Your email address is: " & Email & "." & vbNewLine & "I will contact with you as soon as" & vbNewLine & "I recieve your comments/suggestions/questions." & vbNewLine & "Yufei Zhou", vbOKOnly, "Thank you for your reply!")
Unload Me

End Sub
