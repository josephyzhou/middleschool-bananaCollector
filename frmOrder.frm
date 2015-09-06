VERSION 5.00
Begin VB.Form frmOrder 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Order form"
   ClientHeight    =   5028
   ClientLeft      =   48
   ClientTop       =   420
   ClientWidth     =   6312
   LinkTopic       =   "Form1"
   ScaleHeight     =   5028
   ScaleWidth      =   6312
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Payment"
      Height          =   1932
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   2772
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Product Name"
      Height          =   1932
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2772
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
      Height          =   252
      Left            =   2160
      TabIndex        =   3
      Top             =   4680
      Width           =   2172
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Order other product made by Yufei Zhou"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   372
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   3972
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
