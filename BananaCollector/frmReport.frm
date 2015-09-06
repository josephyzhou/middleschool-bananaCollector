VERSION 5.00
Begin VB.Form frmReport 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Ticket Report"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   5205
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Confirm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   18
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Personal Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtLastName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtFirstName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtAdress 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox txtCity 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtState 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtZip 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Last Name:"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   14.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "First Name:"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   14.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Adress:"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   14.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "City:"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   14.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "State:"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   14.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Zip Code:"
         BeginProperty Font 
            Name            =   "Haettenschweiler"
            Size            =   14.25
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   2520
         Width           =   855
      End
   End
   Begin VB.Label lblFinal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   17
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label lblAmount 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   4800
      Width           =   4695
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "Total: $"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label lblPayment 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   4080
      Width           =   4335
   End
   Begin VB.Label lblPrice 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   3495
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox ("Thanks for your order!")
Unload Me
End Sub

Private Sub Form_Load()
txtFirstName.Text = strFirstName
txtLastName.Text = strLastName
txtAdress = strAdress
txtCity.Text = strCity
txtState.Text = strState
txtZip.Text = strZip

lblPrice.Caption = "Ticket Price: $" & intPri

lblPayment.Caption = "You choose to pay by " & strPayment
lblAmount.Caption = "Numbers of Tickets you order: " & intTic
lblFinal.Caption = intPri * intTic
End Sub

Private Sub Label7_Click()

End Sub
