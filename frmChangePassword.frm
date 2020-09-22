VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                     Change Password"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "frmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "New Username"
      Height          =   1095
      Left            =   2400
      TabIndex        =   7
      Top             =   720
      Width           =   2295
      Begin VB.TextBox txtnewUserNameConf 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtnewUserName 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.TextBox txtOldUserName 
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "          New Password           "
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   2295
      Begin VB.TextBox txtPasswordConf 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtNewPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox txtOldPassword 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Old Username"
      Height          =   195
      Left            =   3000
      TabIndex        =   9
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Old Password"
      Height          =   195
      Left            =   720
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange_Click()
If WebServer.CheckPassword(Me.txtOldPassword) <> True Then
    MsgBox "Either your password or username is not correct", vbCritical, "ERROR"
    Exit Sub
End If

If WebServer.CheckName(Me.txtOldUserName) <> True Then
    MsgBox "Either your password or username is not correct", vbCritical, "ERROR"
    Exit Sub
End If

If Me.txtNewPassword <> Me.txtPasswordConf Then
    MsgBox "New passwords do not match", vbCritical, "ERROR"
    Exit Sub
End If

If Me.txtnewUserName <> Me.txtnewUserNameConf Then
    MsgBox "New usernames do not match", vbCritical, "ERROR"
    Exit Sub
End If
WebServer.ChangeAdministrator Me.txtnewUserName, Me.txtNewPassword

Unload Me
End Sub
