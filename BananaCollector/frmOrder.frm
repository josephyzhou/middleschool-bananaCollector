VERSION 5.00
Begin VB.Form frmOrder 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Order form"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Reset"
      Height          =   375
      Left            =   360
      TabIndex        =   26
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   375
      Left            =   360
      TabIndex        =   25
      Top             =   3720
      Width           =   1095
   End
   Begin VB.HScrollBar hsbYears 
      Height          =   375
      LargeChange     =   5
      Left            =   4560
      Max             =   30
      Min             =   1
      TabIndex        =   10
      Top             =   3360
      Value           =   1
      Width           =   3375
   End
   Begin VB.CommandButton cmdPrice 
      Caption         =   "Price of the Product"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   3960
      Width           =   3375
   End
   Begin VB.OptionButton optCard 
      Caption         =   "Credit Card"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
   End
   Begin VB.OptionButton optMail 
      Caption         =   "Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.OptionButton optVisa 
      Caption         =   "Visa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OptionButton optMaster 
      Caption         =   "Master"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OptionButton optCheck 
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Payment"
      Height          =   2055
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   2772
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Product Name"
      Height          =   2895
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4215
      Begin VB.TextBox txtLastName 
         Height          =   375
         Left            =   1440
         TabIndex        =   18
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtFirstName 
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtAdress 
         Height          =   375
         Left            =   1080
         TabIndex        =   16
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtCity 
         Height          =   375
         Left            =   1080
         TabIndex        =   15
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox txtState 
         Height          =   375
         Left            =   1080
         TabIndex        =   14
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txtZip 
         Height          =   375
         Left            =   2640
         TabIndex        =   13
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label10 
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
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label9 
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
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   21
         Top             =   1800
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
         Left            =   120
         TabIndex        =   20
         Top             =   2280
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
         Left            =   1680
         TabIndex        =   19
         Top             =   2280
         Width           =   855
      End
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Numbers of Product"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "yufeizhou88@hotmail.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Order The full virsion of this program!!!"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intTicket As Integer
Dim intPrice As Integer

Private Sub cmdClear_Click()
    Label8.Caption = ""
    hsbYears.Value = hsbYears.Min
    txtFirstName.Text = ""
    txtLastName.Text = ""
    txtAdress.Text = ""
    txtState.Text = ""
    txtZip.Text = ""
    txtCity.Text = ""
    cmdPrice.Caption = "Price of the Tickets"
End Sub

Private Sub cmdPrice_Click()
If intTicket <= 5 Then
cmdPrice.Caption = "Price: $30"
Else
If intTicket >= 6 And intTicket <= 20 Then
cmdPrice.Caption = "Price: $28"
Else
If intTicket >= 21 And intTicket <= 30 Then
cmdPrice.Caption = "Price: $25"
End If
End If
End If
End Sub

Private Sub cmdSubmit_Click()
strFirstName = txtFirstName.Text
strLastName = txtLastName.Text
strAdress = txtAdress
strCity = txtCity.Text
strState = txtState.Text
strZip = txtZip.Text
intTic = Val(Label8.Caption)


If intTicket <= 5 Then
intPri = 30
Else
If intTicket >= 6 And intTicket <= 20 Then
intPri = 28
Else
If intTicket >= 21 And intTicket <= 30 Then
intPri = 25
End If
End If
End If

If optMail.Value = True Then
strPayment = optMail.Caption
Else
If optVisa.Value = True Then
strPayment = optVisa.Caption
Else
If optCheck.Value = True Then
strPayment = optCheck.Caption
Else
If optMaster.Value = True Then
strPayment = optMaster.Caption
Else
End If
End If
End If
End If
frmOrder.Hide
frmReport.Show
End Sub

Private Sub hsbYears_Change()
Label8.Caption = hsbYears.Value
intTicket = Val(Label8.Caption)

End Sub

Private Sub optCard_Click()
optMail.Visible = False
optCard.Visible = False
optVisa.Visible = True
optMaster.Visible = True
optCheck.Visible = True

End Sub
