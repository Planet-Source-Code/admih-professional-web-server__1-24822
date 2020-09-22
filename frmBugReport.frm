VERSION 5.00
Begin VB.Form frmBugReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bug Report..."
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "frmBugReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send Report"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "How we can contact you"
      Height          =   1095
      Left            =   1680
      TabIndex        =   2
      Top             =   3720
      Width           =   3255
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtEMAIl 
         Height          =   285
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bug"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtBug 
         Height          =   3135
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   240
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmBugReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSend_Click()
If Me.txtBug = "" Then Exit Sub
If Me.txtEMAIl = "" Then Exit Sub
If Me.txtName = "" Then Exit Sub
WebServer.SendMail "mail.icqmail.com", "serechenka@icqmail.com", "Sergey", Me.txtEMAIl, Me.txtName, Me.txtBug, "Bug Report(PWS)"
Unload Me
End Sub
