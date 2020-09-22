VERSION 5.00
Begin VB.Form frmPassword 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2400
   ClientLeft      =   4110
   ClientTop       =   3975
   ClientWidth     =   4965
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2400
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   135
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   4575
   End
   Begin VB.TextBox txtPassowrd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      Default         =   -1  'True
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   20
      TabIndex        =   8
      Top             =   40
      Width           =   4935
      Begin VB.Label Label6 
         Caption         =   "This area of application requres password authorization.  Please enter you password."
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Left            =   720
         TabIndex        =   10
         Top             =   1320
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Username"
         Height          =   195
         Left            =   720
         TabIndex        =   9
         Top             =   960
         Width           =   720
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "This options contains important informations, and to enter you must have password."
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   4815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   720
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Username:"
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   960
      Width           =   765
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
WebServer.SSTab1.Tab = 0
Unload Me
End Sub

Private Sub cmdEnter_Click()

If WebServer.CheckPassword(Me.txtPassowrd) = True Then
    If WebServer.CheckName(Me.txtUserName) = True Then
        WebServer.SSTab1.Tab = 4
        WebServer.mnuViewSecurity.Visible = True
        WebServer.mnuViewSecurity.Checked = True
        WebServer.mnuViewOptions.Checked = False
        WebServer.mnuViewStatistics.Checked = False
        WebServer.mnuViewVisitors.Checked = False
        WebServer.mnuViewLog.Checked = False
        Unload Me
        Exit Sub
    Else
        MsgBox "Please check your password and username", vbCritical, "ERROR"
        Exit Sub
    End If
Else
    MsgBox "Please check your password and username", vbCritical, "ERROR"
End If


Unload Me

End Sub
