VERSION 5.00
Begin VB.Form frmEMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E-Mail Settings..."
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   Icon            =   "frmEMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Text            =   "Sergey"
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "One of your visitors has a question..."
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox txtMail 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Text            =   "serechenka@icqmail.com"
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Administrator's Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Subject of the message:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Administrator's E-Mail:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1530
   End
End
Attribute VB_Name = "frmEMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
LoadEMail
End Sub

Private Sub cmdSave_Click()
Me.Hide
SaveEMail
End Sub
