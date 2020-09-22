VERSION 5.00
Begin VB.Form frmShowFolder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folder"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "frmShowFolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmShowFolder.frx":000C
      Left            =   2160
      List            =   "frmShowFolder.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtDescription 
      Height          =   3495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   960
      Width           =   5535
   End
   Begin VB.TextBox txtFolderName 
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Folder Access:"
      Height          =   195
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Folder's Name:"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   1050
   End
End
Attribute VB_Name = "frmShowFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
For i = 1 To frmSecuredFolder.lstFolders.ListItems.Count
    If frmSecuredFolder.lstFolders.ListItems(i).Text = Me.txtFolderName Then
        frmSecuredFolder.lstFolders.ListItems(i).SubItems(1) = Me.Combo1
        frmSecuredFolder.lstFolders.ListItems(i).SubItems(2) = Me.txtDescription
    End If
Next i
End Sub

Function ShowFolder()
Me.cmdCancel.Value = False
Me.cmdSave.Value = False
Me.cmdOK.Value = True
End Function

Function SetFolder()
Me.cmdCancel.Value = True
Me.cmdSave.Value = True
Me.cmdOK.Value = False

End Function
