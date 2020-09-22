VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSecuredFolder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folders Options..."
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frmSecuredFolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Folder Settings"
      Height          =   3135
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Width           =   6375
      Begin VB.TextBox txtDescription 
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   10
         Top             =   600
         Width           =   6135
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmSecuredFolder.frx":000C
         Left            =   4200
         List            =   "frmSecuredFolder.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtSecuredFolder 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "Secured"
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton cmdBRW 
         Caption         =   "..."
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear List"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   3840
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   $"frmSecuredFolder.frx":0043
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   7335
      Begin MSComctlLib.ListView lstFolders 
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   3836
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Folder Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Access"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   5520
      Width           =   855
   End
End
Attribute VB_Name = "frmSecuredFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Dim itmX As ListItem

If Me.Combo1.Text = "" Then
    MsgBox "Please Select Type of access to the folder"
    Exit Sub
End If

If Me.txtSecuredFolder = "" Then
    Me.txtSecuredFolder.SetFocus
    MsgBox "Please specify folder!", vbCritical, "ERROR"
    Exit Sub
End If

Set itmX = lstFolders.ListItems.Add(, , Me.txtSecuredFolder)
itmX.SubItems(1) = Me.Combo1.Text
itmX.SubItems(2) = Me.txtDescription
Me.txtSecuredFolder = ""

End Sub

Private Sub cmdBRW_Click()
Dim strFolder
strFolder = BrowseForFolder(App.Path)
Do While InStr(strFolder, "\") <> 0
    strFolder = Mid(strFolder, InStr(strFolder, "\") + 1)
Loop
Me.txtSecuredFolder = strFolder
End Sub

Private Sub cmdClear_Click()
For i = 1 To lstFolders.ListItems.Count
    lstFolders.ListItems.Remove (i)
Next i
End Sub

Private Sub cmdOK_Click()
Me.Hide
End Sub

Private Sub cmdRemove_Click()
lstFolders.ListItems.Remove (lstFolders.SelectedItem.Index)
End Sub

Private Sub Form_Load()
Me.Combo1.Text = Me.Combo1.List(0)
End Sub


Private Sub lstFolders_DblClick()
If Me.lstFolders.SelectedItem Is Nothing Then Exit Sub
If Me.lstFolders.SelectedItem.Text <> "" Then
    Load frmShowFolder
    frmShowFolder.txtDescription = Me.lstFolders.SelectedItem.SubItems(2)
    frmShowFolder.txtFolderName = Me.lstFolders.SelectedItem.Text
    frmShowFolder.Combo1.Text = Me.lstFolders.SelectedItem.SubItems(1)
    'frmShowFolder.ShowFolder
    frmShowFolder.Show
End If
End Sub
