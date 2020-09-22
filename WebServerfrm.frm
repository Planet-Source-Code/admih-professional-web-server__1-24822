VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form WebServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PWS - Professional Web Server"
   ClientHeight    =   6405
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7455
   Icon            =   "WebServerfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrSendMessages 
      Interval        =   60000
      Left            =   4320
      Top             =   6000
   End
   Begin PWS.TrayArea TrayArea 
      Left            =   3720
      Top             =   6000
      _ExtentX        =   900
      _ExtentY        =   900
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start"
      Height          =   375
      Left            =   6000
      TabIndex        =   16
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6000
      TabIndex        =   17
      Top             =   6000
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Options"
      TabPicture(0)   =   "WebServerfrm.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblEMail"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdPath"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtWebFolder"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtServerIP"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdCopyIP"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Statistics"
      TabPicture(1)   =   "WebServerfrm.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtTotalVisitors"
      Tab(1).Control(1)=   "MSChart"
      Tab(1).Control(2)=   "Label5"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Visitors"
      TabPicture(2)   =   "WebServerfrm.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "ListView1"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Log"
      TabPicture(3)   =   "WebServerfrm.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(1)=   "txtLOG"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Security"
      TabPicture(4)   =   "WebServerfrm.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picSettings"
      Tab(4).Control(1)=   "picLogged"
      Tab(4).Control(2)=   "picUsers"
      Tab(4).Control(3)=   "picIntro"
      Tab(4).Control(4)=   "trvSecurity"
      Tab(4).ControlCount=   5
      Begin VB.PictureBox picSettings 
         Height          =   5175
         Left            =   -72720
         ScaleHeight     =   5115
         ScaleWidth      =   4995
         TabIndex        =   49
         Top             =   480
         Visible         =   0   'False
         Width           =   5055
         Begin VB.CommandButton cmdSecuredFolder 
            Caption         =   "Secured Folders"
            Height          =   375
            Left            =   3360
            TabIndex        =   53
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdPassword 
            Caption         =   "Password"
            Height          =   375
            Left            =   1800
            TabIndex        =   52
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdEMail 
            Caption         =   "E-Mail Settings"
            Height          =   375
            Left            =   240
            TabIndex        =   51
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   $"WebServerfrm.frx":0396
            Height          =   615
            Left            =   240
            TabIndex        =   54
            Top             =   1440
            Width           =   4455
         End
         Begin VB.Label Label9 
            Caption         =   "Settings"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1680
            TabIndex        =   50
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Incoming E-Mails"
         Height          =   3855
         Left            =   3000
         TabIndex        =   47
         Top             =   1560
         Width           =   4335
         Begin MSComctlLib.ListView lstEMail 
            Height          =   3495
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   6165
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Sender's Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Sender's E-Mail"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Question"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.PictureBox picLogged 
         Height          =   5175
         Left            =   -72720
         ScaleHeight     =   5115
         ScaleWidth      =   4995
         TabIndex        =   42
         Top             =   480
         Visible         =   0   'False
         Width           =   5055
         Begin VB.Timer tmrUserLogOut 
            Enabled         =   0   'False
            Index           =   0
            Interval        =   60000
            Left            =   4080
            Top             =   1920
         End
         Begin VB.CommandButton cmdClearLoggedUsers 
            Caption         =   "Clear All"
            Height          =   255
            Left            =   3840
            TabIndex        =   46
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton cmdRemoveLoggedUser 
            Caption         =   "Remove"
            Height          =   255
            Left            =   3840
            TabIndex        =   45
            Top             =   840
            Width           =   975
         End
         Begin MSComctlLib.ListView lsvLOGGEDUSERS 
            Height          =   4455
            Left            =   90
            TabIndex        =   44
            Top             =   600
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   7858
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "User Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Minutes Left"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Logged Users"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1200
            TabIndex        =   43
            Top             =   120
            Width           =   2475
         End
      End
      Begin VB.PictureBox picUsers 
         Height          =   5175
         Left            =   -72720
         ScaleHeight     =   5115
         ScaleWidth      =   4995
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   5055
         Begin VB.CommandButton cmdWriteEMail 
            Caption         =   "Write e-mail"
            Height          =   255
            Left            =   3840
            TabIndex        =   38
            Top             =   1560
            Width           =   975
         End
         Begin VB.CommandButton cmdChangeUser 
            Caption         =   "Change..."
            Height          =   255
            Left            =   3840
            TabIndex        =   37
            Top             =   1200
            Width           =   975
         End
         Begin VB.CommandButton cmdAddUser 
            Caption         =   "Add User"
            Height          =   255
            Left            =   3840
            TabIndex        =   36
            Top             =   840
            Width           =   975
         End
         Begin MSComctlLib.ListView lsvUsers 
            Height          =   4455
            Left            =   90
            TabIndex        =   35
            Top             =   600
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   7858
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "User Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Password"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Users"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1920
            TabIndex        =   34
            Top             =   120
            Width           =   1035
         End
      End
      Begin VB.PictureBox picIntro 
         Height          =   5175
         Left            =   -72720
         ScaleHeight     =   5115
         ScaleWidth      =   4875
         TabIndex        =   39
         Top             =   480
         Width           =   4935
         Begin VB.CommandButton cmdLogOut 
            Caption         =   "LogOut"
            Height          =   255
            Left            =   1800
            TabIndex        =   41
            Top             =   4680
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   $"WebServerfrm.frx":0440
            Height          =   735
            Left            =   240
            TabIndex        =   40
            Top             =   120
            Width           =   4335
         End
      End
      Begin VB.TextBox txtTotalVisitors 
         Height          =   285
         Left            =   -71400
         TabIndex        =   31
         Top             =   5400
         Width           =   1455
      End
      Begin MSComctlLib.TreeView trvSecurity 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   30
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   9128
         _Version        =   393217
         Indentation     =   295
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.CommandButton cmdCopyIP 
         Caption         =   "Copy"
         Height          =   255
         Left            =   5160
         TabIndex        =   29
         Top             =   1200
         Width           =   615
      End
      Begin MSChart20Lib.MSChart MSChart 
         Height          =   5055
         Left            =   -74880
         OleObjectBlob   =   "WebServerfrm.frx":04DF
         TabIndex        =   27
         Top             =   360
         Width           =   7215
      End
      Begin VB.Frame Frame4 
         Caption         =   $"WebServerfrm.frx":301C
         Height          =   855
         Left            =   -74880
         TabIndex        =   23
         Top             =   4920
         Width           =   7215
         Begin VB.CommandButton cmdVIClear 
            Caption         =   "Clear"
            Height          =   495
            Left            =   3000
            TabIndex        =   26
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdVISave 
            Caption         =   "Save to File"
            Height          =   495
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdVICopy 
            Caption         =   "Copy to Clipboard"
            Height          =   495
            Left            =   5760
            TabIndex        =   24
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   $"WebServerfrm.frx":30B2
         Height          =   855
         Left            =   -74880
         TabIndex        =   19
         Top             =   4920
         Width           =   7215
         Begin VB.CommandButton cmdLICopy 
            Caption         =   "Copy to Clipboard"
            Height          =   495
            Left            =   5760
            TabIndex        =   22
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdLISave 
            Caption         =   "Save to File"
            Height          =   495
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton cmdLIClear 
            Caption         =   "Clear"
            Height          =   495
            Left            =   3000
            TabIndex        =   20
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.TextBox txtLOG 
         Height          =   4575
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   18
         Top             =   360
         Width           =   7215
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "User IP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Last Time Visited"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Times requested file"
            Object.Width           =   3069
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "User Name"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   "IP Restrictins"
         Height          =   3855
         Left            =   240
         TabIndex        =   8
         Top             =   1560
         Width           =   2655
         Begin VB.CommandButton cmdAddIP 
            Caption         =   "Add"
            Enabled         =   0   'False
            Height          =   315
            Left            =   1680
            TabIndex        =   13
            Top             =   3480
            Width           =   855
         End
         Begin VB.CommandButton cmdRemoveIP 
            Caption         =   "Remove"
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            TabIndex        =   12
            Top             =   3480
            Width           =   855
         End
         Begin VB.Frame Frame2 
            Caption         =   "Restricted IPs:"
            Height          =   2775
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   2415
            Begin VB.ListBox lstRestricedIP 
               Enabled         =   0   'False
               Height          =   2400
               Left            =   120
               TabIndex        =   11
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.CheckBox chkResrictIP 
            Caption         =   "Enable IP Restrictions"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.TextBox txtServerIP 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtWebFolder 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Top             =   840
         Width           =   3015
      End
      Begin VB.CommandButton cmdPath 
         Caption         =   "..."
         Height          =   255
         Left            =   5160
         TabIndex        =   3
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total Visitors:"
         Height          =   195
         Left            =   -72480
         TabIndex        =   32
         Top             =   5400
         Width           =   945
      End
      Begin VB.Label lblEMail 
         Caption         =   "serechenka@icqmail.com"
         Height          =   255
         Left            =   4800
         TabIndex        =   28
         Top             =   5520
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "If you have any question or suggestion, please E-Mail to me."
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   5520
         Width           =   4245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Location of Web Pages:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1725
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Server IP:"
         Height          =   195
         Left            =   1200
         TabIndex        =   6
         Top             =   1200
         Width           =   705
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   2160
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2640
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wskMail 
      Left            =   3240
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "mail.icqmail.com"
      RemotePort      =   25
      LocalPort       =   6000
   End
   Begin VB.Label conlab 
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Connections:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   930
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "E&dit"
      Visible         =   0   'False
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatistics 
         Caption         =   "Statistics"
      End
      Begin VB.Menu mnuViewVisitors 
         Caption         =   "&Visitors"
      End
      Begin VB.Menu mnuViewLog 
         Caption         =   "&Log"
      End
      Begin VB.Menu mnuViewSecurity 
         Caption         =   "&Security"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpGettingStarted 
         Caption         =   "Getting Started"
      End
      Begin VB.Menu mnuHelpBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents..."
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "Index"
      End
      Begin VB.Menu mnuHelpBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpBugReport 
         Caption         =   "Bug Report"
      End
      Begin VB.Menu mnuHelpBreak3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About PWS"
      End
   End
   Begin VB.Menu mnuIcon 
      Caption         =   "Icon's Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuIconShowPws 
         Caption         =   "Show PWS"
      End
      Begin VB.Menu mnuIconBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIconAbout 
         Caption         =   "About PWS"
      End
      Begin VB.Menu mnuIconBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIconExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuStatistics 
      Caption         =   "Statistics"
      Visible         =   0   'False
      Begin VB.Menu mnuStatisticsChangeStyle 
         Caption         =   "Change Style"
         Begin VB.Menu mnuStatisticsChangeStyleTo 
            Caption         =   "3D Bar"
            Index           =   0
         End
         Begin VB.Menu mnuStatisticsChangeStyleTo 
            Caption         =   "2D Bar"
            Index           =   1
         End
         Begin VB.Menu mnuStatisticsChangeStyleTo 
            Caption         =   "3D Line"
            Index           =   2
         End
         Begin VB.Menu mnuStatisticsChangeStyleTo 
            Caption         =   "2D Line"
            Index           =   3
         End
         Begin VB.Menu mnuStatisticsChangeStyleTo 
            Caption         =   "3D Area"
            Index           =   4
         End
         Begin VB.Menu mnuStatisticsChangeStyleTo 
            Caption         =   "2D Area"
            Index           =   5
         End
         Begin VB.Menu mnuStatisticsChangeStyleTo 
            Caption         =   "3D Step"
            Index           =   6
         End
         Begin VB.Menu mnuStatisticsChangeStyleTo 
            Caption         =   "2D Step"
            Checked         =   -1  'True
            Index           =   7
         End
         Begin VB.Menu mnuStatisticsChangeStyleTo 
            Caption         =   "3D Combination"
            Index           =   8
         End
         Begin VB.Menu mnuStatisticsChangeStyleTo 
            Caption         =   "2D Combination"
            Index           =   9
         End
      End
      Begin VB.Menu mnuStatisticsChangeMax 
         Caption         =   "Change Max Value"
      End
      Begin VB.Menu mnuStatisticsBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatisticsReset 
         Caption         =   "Reset"
      End
   End
End
Attribute VB_Name = "WebServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







'Date of Creation       |    10:48 PM 28/06/2001
'------------------------------------------------------------------------------
'Product Name           |    PWS(Professional Web Server)
'------------------------------------------------------------------------------
'Creator                |    Sergey Kloubkov
'------------------------------------------------------------------------------
'Contact infromation    |    serechenka@icqmail.com
'------------------------------------------------------------------------------
'Purpose of Application |    Ability of mid-user to create Home Based Web Server


'Help Informaiton:
'This is my first publication of any software.  So you might find mistakes or anything like that
'so if you do, just e-mail them to me.  Thanks in advance
'If you will be using this software please e-mail me =)














'I'll try to put as much infromation about purpose of button or certain control as possible.

'This software has a major bug, that I could not fix. You cann't store very large files
'it takes too long to open for application to go thought the file, create string and
'send it to the client. If you find solution to it, PLEASE send it to me.


'Here we just tell application that there are some strings that have to be memorized,
'and will be used(or not) by application.
Dim Connections As Integer
Dim strStatus
Dim Green_Light As Boolean
Dim WithEvents adoUsersRS As Recordset
Attribute adoUsersRS.VB_VarHelpID = -1
Dim WithEvents adoAdminRS As Recordset
Attribute adoAdminRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOW = 5
'/////////////////////////////////////////////////////////////////////////////////////

'Here is a first code of the object
Private Sub chkResrictIP_Click()
If Me.chkResrictIP.Value = 1 Then 'check if Restriced Ip has been activated
    lstRestricedIP.Enabled = True 'disable the list of Restriced IPs
    cmdRemoveIP.Enabled = True 'disable "Remove IP" button
    cmdAddIP.Enabled = True 'disable "Add IP" button
Else 'and it has not been activated it, then activate all the necessary controls
    lstRestricedIP.Enabled = False 'activate list of Restriced Ips
    cmdRemoveIP.Enabled = False 'activate "Remove IP" button
    cmdAddIP.Enabled = False 'actiave "Add IP" button
End If 'and thats it
End Sub
Private Sub cmdAddIP_Click()
'Here we Dim our string inside the code of control, this tells application that
'this will be used only inside of this code
Dim strIP
'Set strIP to whatever user enters in the Input Box
strIP = InputBox$("Please enter IP, that you want to restric from visiting your Web Site")
'Check strIP has anything iside(mmm...like IP), if not, then exit
If strIP = "" Then Exit Sub
'run throught the List Of Restricted IPs
For i = 0 To Me.lstRestricedIP.ListCount
    'check if that one is alrady in
    If Me.lstRestricedIP.List(i) = strIP Then
        'and tell user that... well, you can read it
        MsgBox "This IP is alrady in the list", vbCritical, "Input ERROR"
        Exit Sub
    End If
Next i
'othervise just add IP to the list
lstRestricedIP.AddItem strIP
End Sub
Private Sub cmdAddUser_Click()
'show form
frmUsers.Show
'and tell form that we want to add user
frmUsers.AddUser True
End Sub
Private Sub cmdChangeUser_Click()
'show form
If Me.lsvUsers.SelectedItem Is Nothing Then Exit Sub
frmUsers.Show
'and tell form that we want to change the users settings, also tell what number of the user
frmUsers.moveTo Me.lsvUsers.SelectedItem.Index - 1
End Sub
Private Sub cmdClearLoggedUsers_Click()
'run thought the list of logged users
For i = 1 To Me.lsvLOGGEDUSERS.ListItems.Count
    'disable the counter
    Me.tmrUserLogOut(i - 1).Enabled = False
    'and remove user
    Me.lsvLOGGEDUSERS.ListItems.Remove (i)
Next i
End Sub
Private Sub cmdCopyIP_Click()
'copy to the clipboard IP of server
Clipboard.SetText Me.txtServerIP
End Sub
Private Sub cmdEMail_Click()
'show form
frmEMail.Show
End Sub
Private Sub cmdLIClear_Click()
'set all the text(txtLOG) to nothing, basically clear it
Me.txtLOG = ""
End Sub
Private Sub cmdLICopy_Click()
'set clipboard to log
Clipboard.SetText Me.txtLOG
End Sub
Private Sub cmdLISave_Click()
'show dialog(the one, where you choose file name, folder)
Me.CommonDialog1.ShowOpen
'set Title of the Dialog Box
Me.CommonDialog1.DialogTitle = "Please select a file name"
'Open file that has been selected in Dialog Box
Open Me.CommonDialog1.FileName For Output As #1
    'Save to the file LOG
    Print #1, Me.txtLOG
'Close the file
Close #1
End Sub
Private Sub cmdLogOut_Click()
'Change Selected Tab
Me.SSTab1.Tab = 0
End Sub
Private Sub cmdPassword_Click()
'show form, so user will not be able to work with application until form is unloaded
frmChangePassword.Show 1
End Sub
Private Sub cmdPath_Click()
'show the dialog box(the one, where you coose folder),and change path of Web Folder
Me.txtWebFolder = BrowseForFolder(App.Path) & "\"
'Write to ini file, so next time, when you start application, path will be restored
WriteSetting "WebFolder", "Options", Me.txtWebFolder
End Sub
Private Sub cmdRemoveIP_Click()
'check if user have selected IP to remove
If Me.lstRestricedIP.Text = "" Then
    'if there has been no IP selected, then show message
    MsgBox "Please select IP, you want to remove from the list of restriced IP", vbCritical, "ERRPR"
    'and exit from sub
    Exit Sub
End If
'run thought the IP list
For i = 0 To Me.lstRestricedIP.ListCount
    'check if item is selected
    If Me.lstRestricedIP.List(i) = Me.lstRestricedIP.Text Then
        'if it is, then delete it
        Me.lstRestricedIP.RemoveItem (i)
        'and exit sub
        Exit Sub
    End If
Next i
End Sub
Private Sub cmdRemoveLoggedUser_Click()
'Disable timer
If Me.lsvLOGGEDUSERS.SelectedItem Is Nothing Then Exit Sub
Me.tmrUserLogOut(Me.lsvLOGGEDUSERS.SelectedItem.Index - 1).Enabled = False
'and erase user from the list
Me.lsvLOGGEDUSERS.ListItems.Remove (Me.lsvLOGGEDUSERS.SelectedItem.Index)
End Sub
Private Sub cmdSecuredFolder_Click()
'show form
frmSecuredFolder.Show
End Sub
Private Sub cmdVIClear_Click()
'clear the list
Me.ListView1.ListItems.Clear
End Sub
Private Sub cmdVICopy_Click()
'set clipboard to nothing
Clipboard.SetText ""
'run throught the list of users
For i = 1 To Me.ListView1.ListItems.Count
    'check if there is anything in the list
    If Clipboard.GetText <> "" Then
        'if there is something then just add to it text
        Clipboard.SetText Clipboard.GetText & Chr(13) & Chr(9) & Me.ListView1.ListItems.Item(i) & Chr(9) & Me.ListView1.ListItems.Item(i).SubItems(1) & Chr(9) & Me.ListView1.ListItems.Item(i).SubItems(2)
    Else
        'if there is nothing, then set text
        Clipboard.SetText Me.ListView1.ListItems.Item(i) & Chr(9) & Me.ListView1.ListItems.Item(i).SubItems(1) & Chr(9) & Me.ListView1.ListItems.Item(i).SubItems(2)
    End If
Next i
End Sub
Private Sub cmdVISave_Click()
'show Dialog Box, where you slected File to save
Me.CommonDialog1.ShowOpen
'change title of the Dialog Box
Me.CommonDialog1.DialogTitle = "Please select a file name"
'Open file, that has been selected in the Dialog Box
Open Me.CommonDialog1.FileName For Output As #1
    'Run thought the list
    For i = 1 To Me.ListView1.ListItems.Count
        'write to the file
        Print #1, Me.ListView1.ListItems.Item(i) & Chr(9) & Me.ListView1.ListItems.Item(i).SubItems(1) & Chr(9) & Me.ListView1.ListItems.Item(i).SubItems(2)
    Next i
'clsoe the file
Close #1
End Sub
Private Sub cmdWriteEMail_Click()
If Me.lsvUsers.SelectedItem Is Nothing Then Exit Sub
'by using Shell, we ask computer to write e-mail, Why this is good?
'well, because computer will open default Email client
ShellExecute hwnd, "open", "mailto:" & Me.lsvUsers.SelectedItem.Tag, vbNullString, vbNullString, SW_SHOW
End Sub
Private Sub Command1_Click()
'sorry for not chaning name =)
'Set connections to 1
Connections = 1
'If application has been working bofore, or Winsock(control) is being used,
'then just close it
Me.Winsock1(0).Close
'Set local port of the Winsock, its not necessary, if application has been running before,
'but it is necessary to set it
Me.Winsock1(0).LocalPort = 80
'and open Winsock again, ant listen Port(80)
Me.Winsock1(0).Listen
'check if there is any information in the LOG
If Len(Me.txtLOG) = 0 Then
    'if there is none, then set a line
    Me.txtLOG = "Server started" & Time
Else
    'but if there is something, we have to add a line
    Me.txtLOG = Me.txtLOG & Chr(13) & Chr(10) & "Server started  " & Time
End If
'set visibility of the Start button to false(so user will not be able to click on it again)
Command1.Visible = False
'and show Close button, so user will be able to close server
Command2.Visible = True
End Sub
Private Sub Ip(GetD, Index, ConnectD)

If ConnectD = "Connect" Then
    Me.txtLOG = Me.txtLOG & Chr(13) & Chr(10) & ConnectD & " " & Time & " with IP " & Winsock1(Index).RemoteHostIP
Else
    Me.txtLOG = Me.txtLOG & Chr(13) & Chr(10) & "" & ConnectD & Time & " " & Winsock1(Index).RemoteHostIP
End If
End Sub
Private Sub Command2_Click()
'set visiblity of the Close button to false
Command2.Visible = False
'and show Start button
Command1.Visible = True
'close winsock, so nobody can connect to server
Me.Winsock1(0).Close
'and add tot he LOG that server has been closed
Me.txtLOG = Me.txtLOG & Chr(13) & Chr(10) & "Server closed at " & Time
End Sub
Private Sub Form_Load()

Dim itmX As ListItem
Dim db As Connection
Me.Show
Me.SSTab1.Tab = 0
Me.txtWebFolder = App.Path & "\"
Me.txtServerIP = Me.Winsock1(0).LocalIP
Me.trvSecurity.Nodes.Add , , "Main", "Main"
Me.trvSecurity.Nodes.Add "Main", tvwChild, , "Users"
Me.trvSecurity.Nodes.Add "Main", tvwChild, , "Logged Users"
Me.trvSecurity.Nodes.Add "Main", tvwChild, , "Settings"
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & App.Path & "\db.mdb;"
Set adoUsersRS = New Recordset
adoUsersRS.Open "select * from Users Order by UserName", db, adOpenStatic, adLockOptimistic
If Not adoUsersRS.EOF Then
    adoUsersRS.MoveFirst
End If
Do While Not adoUsersRS.EOF
    Set itmX = Me.lsvUsers.ListItems.Add(, , adoUsersRS.Fields(0))
    itmX.Tag = adoUsersRS.Fields(4)
    itmX.SubItems(1) = adoUsersRS.Fields(1)
    adoUsersRS.MoveNext
Loop
db.Close
Set db = Nothing
Set adoUsersRS = Nothing
Set db = New Connection
db.CursorLocation = adUseClient
db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & App.Path & "\db.mdb;"
Set adoAdminRS = New Recordset
adoAdminRS.Open "select * from Security", db, adOpenStatic, adLockOptimistic
Me.MSChart.RowCount = 25
Me.MSChart.Column = 2
Me.MSChart.Row = 1
Me.MSChart.Data = 10
Me.MSChart.Column = 1
For i = 1 To 25
    Me.MSChart.Row = i
    Me.MSChart.Data = 0
    Me.MSChart.RowLabel = i - 1
Next i
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Set TrayArea.Icon = Me.Icon
TrayArea.Visible = True
Me.Hide
End Sub
Private Sub lblEMail_Click()
ShellExecute hwnd, "open", "mailto:serechenka@icqmail.com", vbNullString, vbNullString, SW_SHOW
End Sub
Private Sub lblEMail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Me.lblEMail.Font.Underline = False Then Me.lblEMail.Font.Underline = True
Me.lblEMail.ForeColor = vbBlue
End Sub
Private Sub lstEMail_DblClick()
MsgBox Me.lstEMail.SelectedItem.SubItems(2)
End Sub
Private Sub lsvLOGGEDUSERS_DblClick()
MsgBox lsvLOGGEDUSERS.SelectedItem.Tag
End Sub
Private Sub mnuFileExit_Click()
End
End Sub
Private Sub mnuHelpAbout_Click()
mnuIconAbout_Click
End Sub
Private Sub mnuHelpBugReport_Click()
frmBugReport.Show
End Sub

Public Sub mnuIconAbout_Click()
MsgBox Chr(9) & "PWS - Professional Web Server" & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "This program has been designed by Sergey kloubkov" & Chr(13) & Chr(10) & "If you have any quiestion or suggestion feel free " & Chr(13) & Chr(10) & "to use Bug Report form." & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Chr(9) & Chr(9) & "Thank you"
End Sub
Private Sub mnuIconExit_Click()
End
End Sub
Private Sub mnuIconShowPws_Click()
WebServer.Show
TrayArea.Visible = False
End Sub
Private Sub mnuStatisticsChangeMax_Click()
Dim lngNumber As Long
lngNumber = InputBox$("Please Enter the Maximum value of the chart", "Enter ...")
Me.MSChart.Column = 2
Me.MSChart.Row = 1
Me.MSChart.Data = lngNumber
End Sub
Private Sub mnuStatisticsChangeStyleTo_Click(Index As Integer)
For i = 0 To Me.mnuStatisticsChangeStyleTo.UBound
Me.mnuStatisticsChangeStyleTo(i).Checked = False
Next i
Me.mnuStatisticsChangeStyleTo(Index).Checked = True
Me.MSChart.chartType = Index
Me.MSChart.Refresh
End Sub
Private Sub mnuStatisticsReset_Click()
For i = 1 To 24
Me.MSChart.Column = 1
Me.MSChart.Row = i
Me.MSChart.Data = 0
Next i
End Sub
Private Sub mnuViewLog_Click()
mnuViewOptions.Checked = False
mnuViewStatistics.Checked = False
mnuViewVisitors.Checked = False
mnuViewLog.Checked = True
mnuViewSecurity.Visible = False
mnuViewSecurity.Checked = False
Me.SSTab1.Tab = 3
End Sub
Private Sub mnuViewOptions_Click()
mnuViewOptions.Checked = True
mnuViewStatistics.Checked = False
mnuViewVisitors.Checked = False
mnuViewLog.Checked = False
mnuViewSecurity.Visible = False
mnuViewSecurity.Checked = False
Me.SSTab1.Tab = 0
End Sub
Private Sub mnuViewSecurity_Click()
Me.SSTab1.Tab = 4
End Sub
Private Sub mnuViewStatistics_Click()
mnuViewOptions.Checked = False
mnuViewStatistics.Checked = True
mnuViewVisitors.Checked = False
mnuViewLog.Checked = False
mnuViewSecurity.Visible = False
mnuViewSecurity.Checked = False
Me.SSTab1.Tab = 1
End Sub
Private Sub mnuViewVisitors_Click()
mnuViewOptions.Checked = False
mnuViewStatistics.Checked = False
mnuViewVisitors.Checked = True
mnuViewLog.Checked = False
mnuViewSecurity.Visible = False
mnuViewSecurity.Checked = False
Me.SSTab1.Tab = 2
End Sub
Private Sub MSChart_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
PopupMenu mnuStatistics
End If
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
If Me.SSTab1.Tab = 4 Then
    If frmPassword.Visible = False Then
        Me.SSTab1.Tab = 3
        frmPassword.Show 1
    End If
End If
If Me.SSTab1.Tab = 0 Then
    mnuViewOptions.Checked = True
    mnuViewStatistics.Checked = False
    mnuViewVisitors.Checked = False
    mnuViewLog.Checked = False
End If
If Me.SSTab1.Tab = 1 Then
    mnuViewOptions.Checked = False
    mnuViewStatistics.Checked = True
    mnuViewVisitors.Checked = False
    mnuViewLog.Checked = False
End If
If Me.SSTab1.Tab = 2 Then
    mnuViewOptions.Checked = False
    mnuViewStatistics.Checked = False
    mnuViewVisitors.Checked = True
    mnuViewLog.Checked = False
End If
If Me.SSTab1.Tab = 3 Then
    mnuViewOptions.Checked = False
    mnuViewStatistics.Checked = False
    mnuViewVisitors.Checked = False
    mnuViewLog.Checked = True
End If
End Sub
Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.lblEMail.Font.Underline = False
Me.lblEMail.ForeColor = vbBlack
End Sub
Private Sub tmrSendMessages_Timer()
If lstEMail.ListItems.Count = 0 Then Exit Sub
For i = 1 To lstEMail.ListItems.Count
SendMail "mail.icqmail.com", frmEMail.txtMail, frmEMail.txtName, lstEMail.ListItems(i).SubItems(1), Me.lstEMail.ListItems(i).Text, Me.lstEMail.ListItems(i).SubItems(2), frmEMail.txtSubject
lstEMail.ListItems.Remove (i)
Next i
End Sub
Private Sub tmrUserLogOut_Timer(Index As Integer)
Dim itmX As ListItem
For i = 1 To Me.lsvLOGGEDUSERS.ListItems.Count
If Me.lsvLOGGEDUSERS.ListItems(i).Text = Me.tmrUserLogOut(Index).Tag Then
Set itmX = Me.lsvLOGGEDUSERS.ListItems(i)
itmX.SubItems(1) = Val(itmX.SubItems(1)) - 1
End If
Next i
If itmX.SubItems(1) <= 0 Then
Me.lsvLOGGEDUSERS.ListItems.Remove (i - 1)
Me.tmrUserLogOut.Item(Index).Enabled = False
End If
End Sub

Private Sub TrayArea_MouseDown(Button As Integer)
If Button = 2 Then
    PopupMenu mnuIcon
End If
End Sub

Private Sub trvSecurity_DblClick()
If Me.trvSecurity.SelectedItem.Text = "Users" Then
Me.picUsers.Visible = True
Me.picIntro.Visible = False
Me.picLogged.Visible = False
picSettings.Visible = False
End If
If Me.trvSecurity.SelectedItem.Text = "Main" Then
Me.picUsers.Visible = False
Me.picIntro.Visible = True
Me.picLogged.Visible = False
picSettings.Visible = False
End If
If Me.trvSecurity.SelectedItem.Text = "Logged Users" Then
Me.picIntro.Visible = False
Me.picUsers.Visible = False
Me.picLogged.Visible = True
picSettings.Visible = False
End If
If Me.trvSecurity.SelectedItem.Text = "Settings" Then
Me.picIntro.Visible = False
Me.picUsers.Visible = False
Me.picLogged.Visible = False
picSettings.Visible = True
End If
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Me.txtTotalVisitors = Val(Me.txtTotalVisitors) + 1
WriteSetting "Visitors", "Options", Me.txtTotalVisitors
Ip strData$, Index, "Connect"
Dim itmX As ListItem
If Index = 0 Then
    Connections = Connections + 1
    conlab = conlab + 1
    Load Winsock1(Connections)
    Winsock1(Connections).LocalPort = 0
    Winsock1(Connections).Accept requestID
    For i = 1 To Me.ListView1.ListItems.Count
        If Me.ListView1.ListItems.Item(i).Text = Me.Winsock1(Connections).RemoteHostIP Then
            Me.ListView1.ListItems.Item(i).SubItems(1) = Time
            Me.ListView1.ListItems.Item(i).SubItems(2) = Val(Me.ListView1.ListItems.Item(i).SubItems(2)) + 1
            Exit Sub
        End If
    Next i
    Set itmX = Me.ListView1.ListItems.Add(, , Me.Winsock1(Connections).RemoteHostIP)
    itmX.SubItems(1) = Time
    itmX.SubItems(2) = 1
End If
End Sub
Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
On Error GoTo ErrHandler
Winsock1(Index).GetData strData$
If Mid$(strData$, 1, 3) = "GET" Then
    findget = InStr(strData$, "GET ")
    spc2 = InStr(findget + 5, strData$, " ")
    Page = Mid$(strData$, findget + 5, spc2 - (findget + 4))
    If Page = " " Then
        Ip strData$, Index, "Requested index.html "
    ElseIf InStr(Page, "asp") <> 0 Then
        Ip strData$, Index, "Requested ASP file"
    Else
    Ip strData$, Index, "Requested " & Page
    End If
    If chkResrictIP.Value = 1 Then
        For i = 0 To Me.lstRestricedIP.ListCount
            If Me.lstRestricedIP.List(i) = Winsock1(Index).RemoteHostIP Then
                Winsock1(Connections).Close
                Exit Sub
            End If
        Next i
    End If
    SendPage Page, Index
End If
Exit Sub

ErrHandler:
Exit Sub
End Sub
Private Sub Winsock1_SendComplete(Index As Integer)
Winsock1(Index).Close
conlab = conlab - 1
End Sub
Public Sub SendPage(Page, Index)
On Error GoTo ErrHandler
Dim strMessage
Dim strUserName
Dim strPassword
Dim itmX As ListItem
If Page = " " Then Page = "index.html"
    Me.MSChart.Column = 1
    Me.MSChart.Row = Val(Format$(Time$, "hh")) + 1
    Me.MSChart.Data = Val(Me.MSChart.Data) + 1
    If UserLogged(Me.Winsock1(Index).RemoteHostIP) = True Then
        ResetTimer (strUserName)
    End If
    If Left(Page, Len("login.asp")) = "login.asp" Then
        Page = LoginASP(Page, Index)
    End If
    If Left(Page, Len("sendmail.asp")) = "sendmail.asp" Then
        Page = DecodeMail(Page)
    End If
    If Right(Page, 1) = " " Then Page = Left(Page, Len(Page) - 1)
    If InStr(Page, "/") = 0 And InStr(Page, ".") = 0 Then Page = Page & "/"
    If Mid(Page, InStr(Page, "/") + 1) = "" Then
        strFolder = Left(Page, InStr(Page, "/") - 1)
        strFolder = strFolder & "/index.html"
        Page = strFolder
    End If
    If InStr(Page, "/") <> 0 Then
        strFolder = Left(Page, InStr(Page, "/") - 1)
        For i = 1 To frmSecuredFolder.lstFolders.ListItems.Count
            If strFolder = frmSecuredFolder.lstFolders.ListItems.Item(i).Text Then
                If frmSecuredFolder.lstFolders.ListItems.Item(i).SubItems(1) = "Members Only" Then
                    If UserLogged(Me.Winsock1(Index).RemoteHostIP) = True Then
                        ResetTimer (strUserName)
                        Page = Replace(Page, "/", "\")
                    Else
                        Page = "rdrct_login_err.html"
                    End If
                ElseIf frmSecuredFolder.lstFolders.ListItems.Item(i).SubItems(1) = "Free Access" Then
                    If Mid(Page, InStr(Page, "/") + 1) = " " Then
                        strFolder = Left(Page, InStr(Page, "/") - 1)
                        strFolder = strFolder & "/index.html"
                        Page = strFolder
                    End If
                Else
                    Page = "index.html"
                End If
            End If
        Next i
    End If
Nr = FreeFile
Tx$ = " "
Lg = FileLen(txtWebFolder & Page)
Open txtWebFolder & Page For Binary As Nr
    tx1$ = ""
    For M = 1 To Lg
        Get #Nr, , Tx$
        tx1$ = tx1$ + Tx$
    Next M
Close Nr
Winsock1(Index).SendData tx1$
Exit Sub

ErrHandler:
MsgBox Page, , Err.Description
Exit Sub
End Sub
Function LoginASP(Page, Index)
Dim itmX As ListItem
Page = Mid(Page, InStr(Page, "?") + 1)
Page = Mid(Page, InStr(Page, "=") + 1)
strUserName = Left(Page, InStr(Page, "&") - 1)
Page = Mid(Page, InStr(Page, "&") + 1)
Page = Mid(Page, InStr(Page, "=") + 1)
strPassword = Left(Page, InStr(Page, "&") - 1)
For i = 1 To Me.lsvUsers.ListItems.Count
    If Me.lsvUsers.ListItems(i).Text = strUserName Then
        If Me.lsvUsers.ListItems(i).SubItems(1) = strPassword Then
            LoginASP = "rdrct_login_ok.html"
        Exit For
        End If
    End If
    LoginASP = "loginerr.html"
Next i
Set itmX = Me.lsvLOGGEDUSERS.ListItems.Add(, , strUserName)
itmX.SubItems(1) = 5
itmX.Tag = Me.Winsock1(Index).RemoteHostIP
Load Me.tmrUserLogOut(Me.tmrUserLogOut.UBound + 1)
Me.tmrUserLogOut(Me.tmrUserLogOut.UBound).Enabled = True
Me.tmrUserLogOut(Me.tmrUserLogOut.UBound).Tag = strUserName
For i = 1 To ListView1.ListItems.Count
    If ListView1.ListItems.Item(i).Text = Winsock1.Item(Index).RemoteHostIP Then
        ListView1.ListItems.Item(i).SubItems(3) = strUserName
    End If
Next i
End Function
Function UserLogged(strIP) As Boolean
For i = 1 To Me.lsvLOGGEDUSERS.ListItems.Count
    If Me.lsvLOGGEDUSERS.ListItems(i).Tag = strIP Then
        UserLogged = True
        Exit Function
    End If
Next i
UserLogged = False
End Function
Function ResetTimer(strUserName)
For i = 1 To Me.tmrUserLogOut.UBound
    If Me.tmrUserLogOut(i).Tag = strUserName Then
        Me.tmrUserLogOut(i).Enabled = False
        Me.tmrUserLogOut(i).Enabled = True
        Me.lsvLOGGEDUSERS.ListItems(i - 1).SubItems(1) = 5
    End If
Next i
End Function
Function DecodeMail(Page)
Dim strName
Dim strMail
Dim strBody
Dim strAddr
Dim itmX As ListItem
strAddr = Page
strAddr = Mid(strAddr, InStr(strAddr, "?") + 1)
strAddr = Mid(strAddr, InStr(strAddr, "=") + 1)
strName = Left(strAddr, InStr(strAddr, "&") - 1)
strAddr = Mid(strAddr, InStr(strAddr, "=") + 1)
strMail = Left(strAddr, InStr(strAddr, "&") - 1)
strAddr = Mid(strAddr, InStr(strAddr, "=") + 1)
If InStr(strAddr, "&") <> 0 Then
    strBody = Left(strAddr, InStr(strAddr, "&") - 1)
Else
    strBody = strAddr
End If
strBody = Replace(strBody, "%0D", Chr(13))
strBody = Replace(strBody, "%0A", Chr(10))
Set itmX = Me.lstEMail.ListItems.Add(, , strName)
itmX.SubItems(1) = strMail
itmX.SubItems(2) = strBody
DecodeMail = "rdrct_mail.html"
End Function

Function SendMail(strSMTPHost, strRecieversMail, strRecieversName, strSendersMail, strSendersName, strBody, strSubject)
wskMail.Close
wskMail.Connect strSMTPHost, "25"

Do While wskMail.State <> sckConnected
    DoEvents
Loop

Do While Green_Light = False
    DoEvents
Loop
wskMail.SendData "MAIL FROM: " & strSendersName & Chr$(13) & Chr$(10)
Do While strStatus <> 1
    DoEvents
Loop
wskMail.SendData "RCPT TO: " & strRecieversMail & Chr$(13) & Chr$(10)
Do While strStatus <> 2
    DoEvents
Loop
wskMail.SendData "DATA" & Chr$(13) & Chr$(10)
Do While strStatus <> 3
    DoEvents
Loop
wskMail.SendData "FROM: " & strSendersName & " <" & strSendersMail & ">" & Chr$(13) & Chr$(10)
wskMail.SendData "TO: " & strRecieversName & " <" & strRecieversMail & ">" & Chr$(13) & Chr$(10)
wskMail.SendData "SUBJECT: " & strSubject & Chr$(13) & Chr$(10)
wskMail.SendData Chr$(13) & Chr$(10)
wskMail.SendData strBody & Chr$(13) & Chr$(10)
wskMail.SendData Chr$(13) & Chr$(10) & "." & Chr$(13) & Chr$(10)
Do While strStatus <> 4
    DoEvents
Loop
wskMail.SendData "QUIT" & Chr$(13) & Chr$(10)
wskMail.Close
End Function
Private Sub wskMail_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
wskMail.GetData strData
Reply = Mid(strData, 1, 3)
If Reply = 250 Or Reply = 354 Then
    strStatus = strStatus + 1
End If
If Reply = 220 Then
    Green_Light = True
End If
End Sub
Function CheckPassword(strUserPassword) As Boolean
If strUserPassword <> adoAdminRS.Fields(1) Then
    CheckPassword = False
    Exit Function
End If
CheckPassword = True
End Function
Function CheckName(strUserName) As Boolean
If strUserName <> adoAdminRS.Fields(0) Then
    CheckName = False
    Exit Function
End If
CheckName = True
End Function
Function ChangeAdministrator(strUserName, strUserPassword)
adoAdminRS.Delete
adoAdminRS.AddNew
adoAdminRS.Fields(0) = strUserName
adoAdminRS.Fields(1) = strUserPassword
adoAdminRS.Update
End Function
