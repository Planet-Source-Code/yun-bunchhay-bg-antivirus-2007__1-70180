VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BG Antivirus 2007 beta"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9420
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":0CCA
   ScaleHeight     =   8805
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAutoRefresh 
      Enabled         =   0   'False
      Left            =   840
      Top             =   7680
   End
   Begin BGAntiVirus.cSysTray cSysTray1 
      Left            =   1200
      Top             =   6600
      _ExtentX        =   900
      _ExtentY        =   900
      InTray          =   0   'False
      TrayIcon        =   "frmMain.frx":6CDF
      TrayTip         =   "BG Antivirus 2007 Beta"
   End
   Begin MSComctlLib.ImageList img16x16 
      Left            =   480
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":79B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7CD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F84
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8224
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8461
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8705
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frameLicense 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Last Update :"
      Height          =   6975
      Left            =   2160
      TabIndex        =   42
      Top             =   1800
      Width           =   7335
      Begin VB.TextBox txtLicense 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   6735
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   43
         Text            =   "frmMain.frx":89AD
         Top             =   165
         Width           =   6495
      End
   End
   Begin VB.Frame frameTool 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Last Update :"
      Height          =   6975
      Left            =   2160
      TabIndex        =   27
      Top             =   1800
      Width           =   7335
      Begin VB.CommandButton cmdTool_ProcessMan 
         Caption         =   "&Processes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   29
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdTool_Startup 
         Caption         =   "S&tartup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   30
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdScanInvalidReg 
         Caption         =   "&Scan Invalid Registry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdEnableReg 
         Caption         =   "&Registry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   31
         Top             =   360
         Width           =   1455
      End
      Begin VB.Frame frameTool_Process 
         BackColor       =   &H00E0E0E0&
         Height          =   5775
         Left            =   240
         TabIndex        =   49
         Top             =   960
         Width           =   6735
         Begin VB.CommandButton cmdProcessEnd 
            Caption         =   "&End Process"
            Height          =   375
            Left            =   2040
            TabIndex        =   54
            Top             =   300
            Width           =   1575
         End
         Begin VB.CommandButton cmdProcessRefresh 
            Caption         =   "&Refresh"
            Height          =   375
            Left            =   285
            TabIndex        =   53
            Top             =   300
            Width           =   1575
         End
         Begin MSComctlLib.ListView lvProcess 
            Height          =   2655
            Left            =   240
            TabIndex        =   50
            Top             =   840
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   4683
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "img16x16"
            SmallIcons      =   "img16x16"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "File Name"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "File Path"
               Object.Width           =   7849
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ProID(ToKill)"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lvProcessDetail 
            Height          =   1935
            Left            =   240
            TabIndex        =   52
            Top             =   3600
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Window Caption"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Window Class"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Top level parent caption"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Top level parent class"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Process"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Handle"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label lblTotalProcess 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   5865
            TabIndex        =   85
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Process Total:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   375
            Left            =   4320
            TabIndex        =   84
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame frameTool_Startup 
         BackColor       =   &H00E0E0E0&
         Height          =   5775
         Left            =   240
         TabIndex        =   34
         Top             =   960
         Width           =   6735
         Begin VB.CommandButton cmdStartUp_Del 
            Caption         =   "&Delete Selected"
            Height          =   375
            Left            =   240
            TabIndex        =   51
            Top             =   360
            Width           =   1455
         End
         Begin MSComctlLib.ListView lvStartUp 
            Height          =   4575
            Left            =   240
            TabIndex        =   48
            Top             =   960
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   8070
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Name"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "File"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Type"
               Object.Width           =   2646
            EndProperty
         End
      End
      Begin VB.Frame frameTool_EnableReg 
         BackColor       =   &H00E0E0E0&
         Height          =   5775
         Left            =   240
         TabIndex        =   33
         Top             =   960
         Width           =   6735
         Begin VB.CommandButton cmdClearAutorun 
            Caption         =   "&Delete Autorun.inf"
            Height          =   495
            Left            =   4320
            TabIndex        =   80
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CheckBox chkLMRegTool 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Disable Registry Tools"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   435
            TabIndex        =   64
            Top             =   3585
            Width           =   2535
         End
         Begin VB.CheckBox chkLMNoSR 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Disable System Restore"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   435
            TabIndex        =   70
            Top             =   4425
            Width           =   2535
         End
         Begin VB.CheckBox chkLmLimitSR 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Limit System Restore Check Point"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   450
            TabIndex        =   69
            Top             =   4965
            Width           =   3375
         End
         Begin VB.CheckBox chkLMNoMSI 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Disable MSI"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   450
            TabIndex        =   68
            Top             =   5235
            Width           =   1455
         End
         Begin VB.CommandButton cmdCleanReg 
            Caption         =   "&Clean Reg"
            Height          =   495
            Left            =   4320
            TabIndex        =   67
            Top             =   360
            Width           =   1935
         End
         Begin VB.CheckBox chkLMNoSRConfig 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Disable System Restore Configuration"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   435
            TabIndex        =   66
            Top             =   4695
            Width           =   3855
         End
         Begin VB.CheckBox chkLMNoTaskmgr 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Disable Task Manager"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   435
            TabIndex        =   65
            Top             =   3855
            Width           =   2415
         End
         Begin VB.CheckBox chkLMNoFolderOption 
            BackColor       =   &H00E0E0E0&
            Caption         =   "No Folder Option"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   435
            TabIndex        =   63
            Top             =   4140
            Width           =   1935
         End
         Begin VB.CheckBox chkCUNoFolderOption 
            BackColor       =   &H00E0E0E0&
            Caption         =   "No Folder Option"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   420
            TabIndex        =   62
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CheckBox chkCUNoCmd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Disable Command Prompt"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   435
            TabIndex        =   61
            Top             =   2340
            Width           =   2775
         End
         Begin VB.CheckBox chkCURegTool 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Disable Registry Tools"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   420
            TabIndex        =   60
            Top             =   645
            Width           =   2415
         End
         Begin VB.CheckBox chkCUNoChangePwd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Disable Change Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   435
            TabIndex        =   59
            Top             =   2625
            Width           =   2775
         End
         Begin VB.CheckBox chkCUNoLock 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Disable Lock Computer"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   435
            TabIndex        =   58
            Top             =   2055
            Width           =   2415
         End
         Begin VB.CheckBox chkCUNoCLose 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Disable Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   435
            TabIndex        =   57
            Top             =   1470
            Width           =   1695
         End
         Begin VB.CheckBox chkCUNoLogoff 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Disable Log off"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   435
            TabIndex        =   56
            Top             =   1755
            Width           =   1695
         End
         Begin VB.CheckBox chkCUTaskmgr 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Disable Task Manager"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   420
            TabIndex        =   55
            Top             =   915
            Width           =   2535
         End
         Begin VB.Label Label12 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Local Machine"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   375
            Left            =   390
            TabIndex        =   72
            Top             =   3225
            Width           =   2655
         End
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Current User"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   375
            Left            =   360
            TabIndex        =   71
            Top             =   300
            Width           =   2655
         End
      End
      Begin VB.Frame frameTool_ScanReg 
         BackColor       =   &H00E0E0E0&
         Height          =   5775
         Left            =   240
         TabIndex        =   32
         Top             =   960
         Width           =   6735
         Begin VB.TextBox txtCurKey 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   735
            Left            =   360
            MultiLine       =   -1  'True
            TabIndex        =   81
            Top             =   1080
            Width           =   6015
         End
         Begin VB.CommandButton cmdDeleteInvalidKey 
            Caption         =   "&Delete Selected"
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
            Left            =   3840
            TabIndex        =   41
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton cmdStartStop 
            Caption         =   "&Start"
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
            Left            =   2640
            TabIndex        =   35
            Top             =   360
            Width           =   915
         End
         Begin MSComctlLib.ListView lvErrorRegKey 
            Height          =   3255
            Left            =   240
            TabIndex        =   39
            Top             =   2280
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   5741
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "img16x16"
            SmallIcons      =   "img16x16"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Found"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "RootKey"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "SubKey"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Value"
               Object.Width           =   2646
            EndProperty
         End
         Begin VB.Label lblScanRegError 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1800
            TabIndex        =   38
            Top             =   1920
            Width           =   675
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Errors Found :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Left            =   360
            TabIndex        =   37
            Top             =   1920
            Width           =   1395
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblScanRegStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Scan :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Left            =   360
            TabIndex        =   36
            Top             =   600
            Width           =   1995
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.Frame frameVirusSig 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Last Update :"
      Height          =   6975
      Left            =   2280
      TabIndex        =   15
      Top             =   1800
      Width           =   7335
      Begin MSComctlLib.ListView lvVirusList 
         Height          =   3975
         Left            =   360
         TabIndex        =   40
         Top             =   1560
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   7011
         View            =   2
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         Icons           =   "img16x16"
         SmallIcons      =   "img16x16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdOnlineUp 
         Caption         =   "Online Update"
         Height          =   375
         Left            =   4800
         TabIndex        =   26
         Top             =   6360
         Width           =   1935
      End
      Begin VB.CommandButton cmdOfflineUp 
         Caption         =   "Offline Update"
         Height          =   375
         Left            =   4800
         TabIndex        =   25
         Top             =   5760
         Width           =   1935
      End
      Begin VB.CommandButton cmdCheckCRC 
         Caption         =   "&Browse Virus"
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtCRC32 
         Enabled         =   0   'False
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtVirusName 
         Height          =   375
         Left            =   3480
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdAddToDef 
         Caption         =   "&Add New Virus Sig."
         Height          =   375
         Left            =   5280
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3120
         Top             =   6120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblVirusCount 
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label lblLastUpdate 
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2160
         TabIndex        =   23
         Top             =   6015
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Update :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   375
         Left            =   480
         TabIndex        =   22
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Virus Count :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   375
         Left            =   480
         TabIndex        =   21
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Virus List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   375
         Left            =   360
         TabIndex        =   20
         Top             =   1200
         Width           =   3255
      End
   End
   Begin VB.Frame frameScan 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   2160
      TabIndex        =   0
      Top             =   1800
      Width           =   7335
      Begin MSComctlLib.ListView lvVirusFound 
         Height          =   3375
         Left            =   240
         TabIndex        =   79
         Top             =   3360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img16x16"
         SmallIcons      =   "img16x16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Virus Name"
            Object.Width           =   3034
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "File Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtPath 
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   600
         Width           =   5895
      End
      Begin VB.CommandButton cmdPath 
         Caption         =   ". . ."
         Height          =   375
         Left            =   6240
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "&Scan"
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "S&top"
         Height          =   615
         Left            =   1680
         TabIndex        =   2
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   5685
         TabIndex        =   3
         Text            =   "512"
         Top             =   2385
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanning : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   6855
      End
      Begin VB.Label lblPath 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000040C0&
         Height          =   615
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   6975
      End
      Begin VB.Label lblCount 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5160
         TabIndex        =   11
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label lblFound 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4380
         TabIndex        =   10
         Top             =   2760
         Width           =   600
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Limit Size of Files to be scanned :           ( KB)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   2400
         Width           =   4215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanning Path:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label lblCleaned 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4380
         TabIndex        =   6
         Top             =   3120
         Width           =   1200
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Virus Found:                of"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   2880
         TabIndex        =   14
         Top             =   2760
         Width           =   3615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Virus Cleaned:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   3120
         Width           =   3615
      End
   End
   Begin VB.Frame frameSetting 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Last Update :"
      Height          =   6975
      Left            =   2160
      TabIndex        =   45
      Top             =   1800
      Width           =   7335
      Begin VB.CommandButton cmdConfigAppBlock 
         Caption         =   "E&dit Application Blocker"
         Height          =   375
         Left            =   4215
         TabIndex        =   91
         Top             =   3300
         Width           =   2655
      End
      Begin VB.CheckBox chkControlAll 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Control all applications (Advanced User only)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Left            =   825
         TabIndex        =   92
         Top             =   3645
         Width           =   4455
      End
      Begin VB.TextBox txtRefreshRate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2085
         TabIndex        =   88
         Text            =   "10"
         Top             =   2865
         Width           =   600
      End
      Begin VB.CheckBox chkAutoRefeshProList 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Auto Refresh &Process List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Left            =   465
         TabIndex        =   87
         Top             =   2385
         Width           =   2775
      End
      Begin VB.CheckBox chkBlockRisk 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Auto Block &Risk Process (Run only when auto refresh is on)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Left            =   480
         TabIndex        =   86
         Top             =   1980
         Width           =   5655
      End
      Begin VB.CommandButton cmdRestoreDefault 
         Caption         =   "&Restore Default"
         Height          =   375
         Left            =   5400
         TabIndex        =   82
         Top             =   5040
         Width           =   1455
      End
      Begin VB.TextBox txtScanRegExt 
         Height          =   1215
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   78
         Top             =   5520
         Width           =   6495
      End
      Begin VB.OptionButton optScanExt 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Scan &Specified Extensions Only (Recommanded)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   360
         TabIndex        =   77
         Top             =   5040
         Width           =   4815
      End
      Begin VB.OptionButton optScanAll 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Scan &All Registry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   360
         TabIndex        =   76
         Top             =   4680
         Width           =   1935
      End
      Begin VB.CheckBox chkAutoScan 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Enable Auto Scan (Under Construction) /  Invisible"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Left            =   480
         TabIndex        =   75
         Top             =   1560
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.CheckBox chkStartMin 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Start &Minimized"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   480
         TabIndex        =   74
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CheckBox chkAutoStart 
         BackColor       =   &H00E0E0E0&
         Caption         =   "A&uto Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Left            =   480
         TabIndex        =   73
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkExeBlock 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Enable Application &Blocker"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Left            =   465
         TabIndex        =   90
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Refresh Rate                   seconds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   375
         Left            =   720
         TabIndex        =   89
         Top             =   2910
         Width           =   3495
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Scan Registry Setting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   480
         TabIndex        =   83
         Top             =   4200
         Width           =   3975
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Application Setting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   480
         TabIndex        =   46
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame frameAbout 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Last Update :"
      Height          =   6975
      Left            =   2160
      TabIndex        =   44
      Top             =   1800
      Width           =   7335
      Begin VB.Timer tmrAbout 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   3600
         Top             =   5760
      End
      Begin VB.TextBox txtAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   8175
         Left            =   360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   47
         Text            =   "frmMain.frx":8BFF
         Top             =   120
         Width           =   6615
      End
   End
   Begin VB.Image imgAbout 
      Height          =   690
      Left            =   390
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":8CB8
      Top             =   5400
      Width           =   1710
   End
   Begin VB.Image imgLicense 
      Height          =   690
      Left            =   390
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":90E0
      Top             =   4740
      Width           =   1710
   End
   Begin VB.Image imgTools 
      Height          =   690
      Left            =   390
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":952A
      Top             =   4065
      Width           =   1710
   End
   Begin VB.Image imgVirusDef 
      Height          =   690
      Left            =   390
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":991D
      Top             =   3390
      Width           =   1710
   End
   Begin VB.Image imgSetting 
      Height          =   690
      Left            =   390
      Picture         =   "frmMain.frx":9DC6
      Top             =   2715
      Width           =   1710
   End
   Begin VB.Image imgScan 
      Height          =   690
      Left            =   390
      Picture         =   "frmMain.frx":A22D
      Top             =   2040
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Image Image6 
      Height          =   690
      Left            =   390
      Picture         =   "frmMain.frx":A60D
      Top             =   5415
      Width           =   1710
   End
   Begin VB.Image Image5 
      Height          =   690
      Left            =   390
      Picture         =   "frmMain.frx":AA35
      Top             =   4740
      Width           =   1710
   End
   Begin VB.Image Image4 
      Height          =   690
      Left            =   390
      Picture         =   "frmMain.frx":AE7F
      Top             =   4065
      Width           =   1710
   End
   Begin VB.Image Image2 
      Height          =   690
      Left            =   390
      Picture         =   "frmMain.frx":B272
      Top             =   3390
      Width           =   1710
   End
   Begin VB.Image Image3 
      Height          =   690
      Left            =   390
      Picture         =   "frmMain.frx":B71B
      Top             =   2715
      Width           =   1710
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   390
      Picture         =   "frmMain.frx":BB82
      Top             =   2040
      Width           =   1710
   End
   Begin VB.Menu mnuop 
      Caption         =   "Op"
      Visible         =   0   'False
      Begin VB.Menu mnuOpSelAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuopUnSelAll 
         Caption         =   "Unselect All"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "tray"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu mnuProMenu 
      Caption         =   "pro_menu"
      Visible         =   0   'False
      Begin VB.Menu mnuProMenu_Ban 
         Caption         =   "&This Process is NOT Safe"
      End
      Begin VB.Menu mnuProMenu_Safe 
         Caption         =   "&This Process Is Safe"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'scan registry variable with EVENTS
Dim WithEvents cReg As cRegSearch
Attribute cReg.VB_VarHelpID = -1


Private Sub cmdClearAutorun_Click()
    Dim fso As New FileSystemObject
    Dim drv As Drive
    Dim drvs As Drives
    On Error Resume Next    'in case not found, and on cd
    Set drvs = fso.Drives
    For Each drv In drvs
        DoEvents
        Kill drv.DriveLetter & ":\autorun.inf"
    Next
    Call ShowTrayMessage("Delete Autorun", "Autorun.inf files from all drives were removed.")
    Set fso = Nothing
    Set drv = Nothing
    Set drvs = Nothing
End Sub



'SysTray Events
Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)
    Me.WindowState = vbNormal
    Show
    Me.cSysTray1.InTray = False
    'Me.SetFocus
End Sub

Private Sub cSysTray1_MouseDown(Button As Integer, Id As Long)
'    Dim MSG As Long
'    'unknown statement
'    MSG = X / Screen.TwipsPerPixelX
'    Select Case MSG
'        Case WM_RBUTTONDOWN     'right click
'            PopupMenu mnuTray
'        Case WM_LBUTTONDBLCLK   'double click
'            mnuOpen_Click
'    End Select
'   MsgBox Button
    If Button = 2 Then   'right click
        PopupMenu mnuTray
    End If
End Sub

'========================================'
' FORM EVENTS                            '
'========================================'
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Hide
        Me.cSysTray1.InTray = True
    End If
End Sub

Private Sub Form_Load()
    
    'check if application is already loaded
    If App.PrevInstance = True Then
        MsgBox "Application is already running."
        'Call ShowWindow(app., 1)
        End
    End If
    Me.Show
    'loading scan page
    '-----------------
    
    'set windows to topmost
    'SetTopMostWindow Me.hwnd, True
    
    'frmSplash.lblLoading.Caption = "Loading Scanning Tool ..."
    'frmSplash.Refresh
    'set default to 512 kb
    FileSize = 524288
    Call imgScan_Click

    'loading setting page
    '-----------------------
    Call LoadSetting
    
    'loading virus sig. page
    '-----------------------
    Call RefreshDefList
    
    
    'loading tool page
    '-----------------------
    'scan reg content
    Set cReg = New cRegSearch
    Call cmdScanInvalidReg_Click
    
    'process content
    Call GetProcessList(Me.lvProcess)
    Me.lblTotalProcess.Caption = Me.lvProcess.ListItems.Count
    Call CheckProcess
    
    'registry content
    Call LoadRegistry
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Me.Tag = "" Then Exit Sub
'    Dim MSG As Long
'    'unknown statement
'    MSG = X / Screen.TwipsPerPixelX
'    Select Case MSG
'        Case WM_RBUTTONDOWN     'right click
'            PopupMenu mnuTray
'        Case WM_LBUTTONDBLCLK   'double click
'            mnuOpen_Click
'    End Select
End Sub

'========================================'
' POPUP MENU                             '
'========================================'

'promenu popup
Private Sub mnuProMenu_Ban_Click()
    Dim f As Long
    f = FreeFile
    Open App.Path & "\AttPro.bin" For Append As f
    Print #f, vbCrLf & Me.lvProcess.ListItems(Me.lvProcess.SelectedItem.Index).Text
    Close f
    Call cmdProcessRefresh_Click
End Sub

Private Sub mnuProMenu_Safe_Click()
    Dim f As Long
    Dim strTemp As String, Temp1 As String, Temp2 As String
    f = FreeFile
    Open App.Path & "\AttPro.bin" For Binary As f
    strTemp = Input$(LOF(f), 1)
    Close f
    Temp1 = Left$(strTemp, InStr(1, strTemp, Me.lvProcess.ListItems(Me.lvProcess.SelectedItem.Index).Text) - 2)
    Temp2 = Right$(strTemp, Len(strTemp) - InStr(1, strTemp, Me.lvProcess.ListItems(Me.lvProcess.SelectedItem.Index).Text) - Len(Me.lvProcess.ListItems(Me.lvProcess.SelectedItem.Index).Text) - 1)
    strTemp = Temp1 & Temp2
    strTemp = Replace(strTemp, vbCrLf & vbCrLf, vbCrLf)
    Open App.Path & "\AttPro.bin" For Output As f
    Print #f, strTemp
    Close f
    Call cmdProcessRefresh_Click
End Sub

'systray popup
Private Sub mnuClose_Click()
'    On Error Resume Next
'    Dim twnd As Long
'    Dim pos As Long
'    Dim traydata As NOTIFYICONDATA
'
'    Dim ticon As Long
'    'find the position of the separator #
'    pos = InStr(1, Me.Tag, "#", vbTextCompare)
'    'get the window handle and icon
'    twnd = CLng(Mid(Me.Tag, 1, pos - 1))
'    ticon = CLng(Mid(Me.Tag, InStr(pos + 1, Me.Tag, "#", vbTextCompare) + 1))
'
'    closewindow twnd
'    'form's window handle
'    traydata.cbSize = Len(traydata)
'    traydata.hwnd = CLng(Mid(Me.Tag, pos + 1, InStr(pos + 1, Me.Tag, "#", vbTextCompare) - pos - 1))
'    traydata.hIcon = ticon
'    traydata.uFlags = NIF_ICON
'
'    'Shell_NotifyIcon with NIM_DELETE
'    'doesn't work anyone can tell me why?
'    Shell_NotifyIcon NIM_DELETE, traydata
'
'    DestroyIcon (ticon)
'    DestroyWindow (traydata.hwnd)
'    UpdateTrayWindow
    'cSysTray1.InTray = False
    'Call Form_Unload(False)
    If blnScan = True Then
        If MsgBox("Scanning is in progress. Are you sure to exit?", vbYesNoCancel, "Abort Scanning") = vbYes Then
            Exit Sub
        End If
    End If

    cReg.StopSearch
    Me.cSysTray1.InTray = False
    End
End Sub


Private Sub mnuOpen_Click()
'
'    On Error Resume Next
'    Dim twnd As Long
'    Dim pos As Long
'    Dim traydata As NOTIFYICONDATA
'
'    Dim ticon As Long
'    'find the position of the separator #
'    pos = InStr(1, Me.Tag, "#", vbTextCompare)
'    'get the window handle and icon
'    twnd = CLng(Mid(Me.Tag, 1, pos - 1))
'    ticon = CLng(Mid(Me.Tag, InStr(pos + 1, Me.Tag, "#", vbTextCompare) + 1))
'
'    ShowWindow twnd, 5
'
'    'form's window handle
'    traydata.cbSize = Len(traydata)
'    traydata.hwnd = CLng(Mid(Me.Tag, pos + 1, InStr(pos + 1, Me.Tag, "#", vbTextCompare) - pos - 1))
'    traydata.hIcon = ticon
'    traydata.uFlags = NIF_ICON
'
'    'Shell_NotifyIcon with NIM_DELETE
'    'doesn't work anyone can tell me why?
'    Shell_NotifyIcon NIM_DELETE, traydata
'
'    DestroyIcon (ticon)
'    DestroyWindow (traydata.hwnd)
'    UpdateTrayWindow
    Me.WindowState = vbNormal
    Me.Show
    cSysTray1.InTray = False
    
End Sub

'========================================'
' SCAN                                   '
'========================================'

Private Sub cmdPath_Click()
    Me.txtPath.Text = BrowseForFolder("Select Folder to Scan")
End Sub

Private Sub cmdScan_Click()
    'check scan status
    If blnScan = False Then 'not scanning
        If Me.txtPath.Text <> "" Then   'scan path is set
            lvVirusFound.ListItems.Clear
            Dim ST As Variant, ET As Variant
            'get start time
            ST = Time
            'reset all statistic
            blnScan = True
            Me.Text1.Enabled = False
            Me.lblCount.Caption = 0
            Me.lblFound.Caption = 0
            Me.lblCleaned.Caption = 0
            Me.cmdStop.SetFocus
            ' strScanDetail = "<Font Name='Verdana' Size=3 Color=Blue>Scanning starts at : " & Time$ & "<br>Scanning Path : <i>" & Me.txtPath.Text & "</i><br>-----------------------------<br></font>"
            ' Call UpdateDetail(strScanDetail, WebBrowser1)
            Dim X As ListItem
            Set X = Me.lvVirusFound.ListItems.Add(, , "Starts at :", 1, 1)
            X.SubItems(1) = Time$
            Set X = Nothing
            Set X = Me.lvVirusFound.ListItems.Add(, , "Scanning Path :", 1, 1)
            X.SubItems(1) = Me.txtPath.Text
            Set X = Nothing
            'start scanning
            Call ScanFile(Me.txtPath.Text)
            
            'finish scanning
            'get end time
            ET = Time
            'calculate time
            Dim x1 As String
            x1 = CalculateTime(ET - ST)
            ' strScanDetail = strScanDetail & "<font size=3 color=BLUE>-----------------------------<br>"
            If blnScan = True Then
                ' strScanDetail = strScanDetail & "Scanning finishes at :" & Time$ & "</font>"
                Set X = Me.lvVirusFound.ListItems.Add(, , "Finishes at :", 1, 1)
                X.SubItems(1) = Time$
                'tray message
                Call ShowTrayMessage("BG Antivirus 2007 beta", "Scanning is completed." & vbCrLf & "Scanning time is " & x1)
            Else
                ' strScanDetail = strScanDetail & "Scanning was cancelled at :" & Time$ & "</font>"
                Set X = Me.lvVirusFound.ListItems.Add(, , "Cancelled at :", 1, 1)
                X.SubItems(1) = Time$
                'MsgBox "Scanning is cancelled." & vbCrLf & "Scanning time is " & x1
                'tray message
                Call ShowTrayMessage("BG Antivirus 2007 beta", "Scanning is cancelled." & vbCrLf & "Scanning time is " & x1)
            End If
            Set X = Nothing
            ' Call UpdateDetail(strScanDetail, WebBrowser1)
            Me.Text1.Enabled = True
            blnScan = False
            Me.lblPath.Caption = ""
        End If
    End If
End Sub

Private Sub cmdStop_Click()
    If blnScan = True Then
        If MsgBox("Are you sure to stop scanning?", vbYesNo + vbDefaultButton2, "Abort Scanning") = vbYes Then
            blnScan = False
        End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
        'check scanning
    If blnScan = True Then
        If MsgBox("Scanning is in progress. Are you sure to exit?", vbYesNoCancel, "Abort Scanning") = vbYes Then
            cReg.StopSearch
            Me.cSysTray1.InTray = False
            End
        Else
            Cancel = True
        End If
    End If
    
    'check reg scanning

    
    'stop application
'    On Error Resume Next
'    'show any window that was hidden in the tray
'    Dim twnd As Long
'    twnd = CLng(Mid(Me.Tag, 1, (InStr(1, Me.Tag, "#", vbTextCompare) - 1)))
'    ShowWindow twnd, 5
'    ShowAllWindows

End Sub

Private Sub lvVirusFound_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub


'textbox determine file size
Private Sub Text1_Change()
    
    If Me.Text1.Text <> "" Then
        If Me.Text1.Text = 0 Then
            FileSize = 52428800
        Else
            FileSize = Int(Me.Text1.Text) * 1024
        End If
    Else
        FileSize = 1024
        Me.Text1.Text = 1
    End If
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    Dim num As String
    num = "0123456789"
    If InStr(1, num, Chr(KeyAscii)) < 1 Then
        KeyAscii = 0
    End If
    
End Sub


'========================================'
' MENU                                   '
'========================================'

Private Sub imgAbout_Click()
    
    'change menu
    Me.imgScan.Visible = True
    Me.imgSetting.Visible = True
    Me.imgVirusDef.Visible = True
    Me.imgTools.Visible = True
    Me.imgLicense.Visible = True
    Me.imgAbout.Visible = False
    
    'change content
    Me.frameScan.Visible = False
    Me.frameSetting.Visible = False
    Me.frameVirusSig.Visible = False
    Me.frameTool.Visible = False
    Me.frameLicense.Visible = False
    Me.frameAbout.Visible = True
    
    'enable scroll
    Me.tmrAbout.Enabled = True
    
End Sub

Private Sub imgLicense_Click()
    
    'change menu
    Me.imgScan.Visible = True
    Me.imgSetting.Visible = True
    Me.imgVirusDef.Visible = True
    Me.imgTools.Visible = True
    Me.imgLicense.Visible = False
    Me.imgAbout.Visible = True
    
    'change content
    Me.frameScan.Visible = False
    Me.frameSetting.Visible = False
    Me.frameVirusSig.Visible = False
    Me.frameTool.Visible = False
    Me.frameLicense.Visible = True
    Me.frameAbout.Visible = False
    
    'disanable scroll
    Me.tmrAbout.Enabled = False
End Sub

Private Sub imgScan_Click()
    
    'change menu
    Me.imgScan.Visible = False
    Me.imgSetting.Visible = True
    Me.imgVirusDef.Visible = True
    Me.imgTools.Visible = True
    Me.imgLicense.Visible = True
    Me.imgAbout.Visible = True
    
    'change content
    Me.frameScan.Visible = True
    Me.frameSetting.Visible = False
    Me.frameVirusSig.Visible = False
    Me.frameTool.Visible = False
    Me.frameLicense.Visible = False
    Me.frameAbout.Visible = False
    
    'disanable scroll
    Me.tmrAbout.Enabled = False
End Sub

Private Sub imgSetting_Click()
    
    'change menu
    Me.imgScan.Visible = True
    Me.imgSetting.Visible = False
    Me.imgVirusDef.Visible = True
    Me.imgTools.Visible = True
    Me.imgLicense.Visible = True
    Me.imgAbout.Visible = True
    
    'change content
    Me.frameScan.Visible = False
    Me.frameSetting.Visible = True
    Me.frameVirusSig.Visible = False
    Me.frameTool.Visible = False
    Me.frameLicense.Visible = False
    Me.frameAbout.Visible = False
    
    'disanable scroll
    Me.tmrAbout.Enabled = False
End Sub

Private Sub imgTools_Click()
    
    'change menu
    Me.imgScan.Visible = True
    Me.imgSetting.Visible = True
    Me.imgVirusDef.Visible = True
    Me.imgTools.Visible = False
    Me.imgLicense.Visible = True
    Me.imgAbout.Visible = True
    
    'change content
    Me.frameScan.Visible = False
    Me.frameSetting.Visible = False
    Me.frameVirusSig.Visible = False
    Me.frameTool.Visible = True
    Me.frameLicense.Visible = False
    Me.frameAbout.Visible = False
    
    'disanable scroll
    Me.tmrAbout.Enabled = False
End Sub

Private Sub imgVirusDef_Click()
    
    'change menu
    Me.imgScan.Visible = True
    Me.imgSetting.Visible = True
    Me.imgVirusDef.Visible = False
    Me.imgTools.Visible = True
    Me.imgLicense.Visible = True
    Me.imgAbout.Visible = True
    
    'change content
    Me.frameScan.Visible = False
    Me.frameSetting.Visible = False
    Me.frameVirusSig.Visible = True
    Me.frameTool.Visible = False
    Me.frameLicense.Visible = False
    Me.frameAbout.Visible = False
    
    'disanable scroll
    Me.tmrAbout.Enabled = False
End Sub

'========================================'
' VIRUS SIGNATURE                        '
'========================================'

Private Sub cmdAddToDef_Click()

    If Me.txtVirusName.Text <> "" Then

        Dim v As VirusSig
        v.Name = Me.txtVirusName.Text
        v.Type = "CRC"
        v.Value = Me.txtCRC32.Text
        Call WriteSig(v)
        'clear textboxes
        Me.txtVirusName.Text = ""
        Me.txtCRC32.Text = ""
        'refresh virus list
        Call RefreshDefList
        
    End If
    
End Sub

Private Sub cmdCheckCRC_Click()
    
    Me.CommonDialog1.DialogTitle = "Open to check CRC signature"
    'Me.CommonDialog1.Flags
    Me.CommonDialog1.ShowOpen
    If Me.CommonDialog1.FileName <> "" Then
        Me.txtCRC32.Text = CRC.GetCRC(Me.CommonDialog1.FileName)
    End If
    
End Sub

Private Sub lvVirusList_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Sub RefreshDefList()
    
    Dim i As Long
    Me.lvVirusList.ListItems.Clear
    For i = 0 To UBound(VSig)
        Me.lvVirusList.ListItems.Add , , VSig(i).Name, 4, 4
    Next i
    Me.lblVirusCount.Caption = VSInfo.VirusCount
    Me.lblLastUpdate.Caption = Format(VSInfo.LastUpdate, "dd mmmm yyyy")
    
End Sub



Private Sub txtVirusName_KeyPress(KeyAscii As Integer)
    
    If Chr(KeyAscii) = "," Then
        KeyAscii = 0
        MsgBox "Disallowed character for name.", vbCritical
    End If
    
End Sub

'========================================'
' SETTING                                '
'========================================'

Private Sub txtScanRegExt_Change()
    Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", 1, "RegExt", Me.txtScanRegExt.Text)
End Sub

Private Sub chkAutoScan_Click()
    'CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanRegOption", 1
End Sub

Private Sub chkAutoStart_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AutoStart", Val(Me.chkAutoStart.Value)
    'add to startup
    If Me.chkAutoStart.Value = 1 Then
        Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Run", 1, "BGAntivirus", App.Path & "\" & App.EXEName & ".exe")
        'CreateStringValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", 1, "BGAntivirus", App.Path & "\" & App.EXEName & ".exe"
    Else
        DeleteStartup GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Run", "BGAntiVirus"
        Call cmdTool_Startup_Click
    End If
End Sub

Private Sub chkStartMin_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "StartMin", Val(Me.chkStartMin.Value)
End Sub

Private Sub cmdRestoreDefault_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "RefreshRate", 10
    Me.optScanExt.Value = True
    Me.txtScanRegExt.Text = "OCX, DLL, EXE, VBS, SYS, VXD"
    CreateStringValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", 1, "RegExt", "OCX, DLL, EXE, VBS, SYS, VXD"
    Call LoadSetting
End Sub

Private Sub optScanAll_Click()
    intSettingRegOption = 1
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanRegOption", 1
End Sub

Private Sub optScanExt_Click()
    intSettingRegOption = 0
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanRegOption", 0
End Sub

Public Sub LoadSetting()
    'check general setting
    Me.chkAutoStart.Value = Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AutoStart"))
    Me.chkStartMin.Value = Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "StartMin"))
    
    'EXE File Block
    Me.chkExeBlock.Value = Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AppBlock"))
    Me.chkControlAll.Value = Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ControlAll"))
    'check refresh rate
    txtRefreshRate.Text = GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "RefreshRate")
    Me.tmrAutoRefresh.interval = txtRefreshRate.Text * 1000
    'check auto refresh status
    Me.chkAutoRefeshProList.Value = Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AutoRefresh"))
'    If chkAutoRefeshProList.Value = 1 Then
'        Me.tmrAutoRefresh.Enabled = True
'    Else
'        Me.tmrAutoRefresh.Enabled = False
'    End If
    
    'process block
    Me.chkBlockRisk.Value = Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "BlockRisk"))
    'start minimized
    If Me.chkStartMin.Value = 1 Then
        'Call HideToTray
        Hide
        Me.cSysTray1.InTray = True
    End If
    'Me.chkAutoScan.Value = ""
    'check options
    intSettingRegOption = Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanRegOption"))
    If intSettingRegOption = 1 Then 'scan all
        Me.optScanAll.Value = True
    Else    'scan specific file extension
        Me.optScanExt.Value = True
    End If
    Me.txtScanRegExt.Text = GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "RegExt")
    'get scan reg ext to variable for scanning
    strScanRegExt = GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "RegExt")
End Sub

Private Sub chkAutoRefeshProList_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AutoRefresh", Val(Me.chkAutoRefeshProList.Value)
    If chkAutoRefeshProList.Value = 1 Then
        Me.tmrAutoRefresh.Enabled = True
    Else
        Me.tmrAutoRefresh.Enabled = False
    End If
End Sub

Private Sub chkBlockRisk_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "BlockRisk", Val(Me.chkBlockRisk.Value)
End Sub

Private Sub chkExeBlock_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AppBlock", Val(Me.chkExeBlock.Value)
    'edit in EXEFile/Shell/Open/Command
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    If Me.chkExeBlock.Value = 1 Then
        ' CreateStringValue GetClassKey("HKEY_CLASS_ROOT"), "\exefile\shell\open\command\", 1, , App.Path & "\" & App.EXEName & ".EXE %1 %*"
        sh.regwrite "HKCR\exefile\shell\open\command\original", Chr$(34) + "%1" + Chr$(34) + " %*"
        sh.regwrite "HKCR\exefile\shell\open\command\", App.Path & "\AppBlock.EXE %1 %*"
        chkControlAll.Enabled = True
    Else
        ' CreateStringValue GetClassKey("HKEY_CLASS_ROOT"), "\exefile\shell\open\command\", 1, , """%1"" %*"
        sh.regwrite "HKCR\exefile\shell\open\command\", Chr$(34) + "%1" + Chr$(34) + " %*"
        chkControlAll.Enabled = False
    End If
End Sub

Private Sub chkControlAll_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ControlAll", Val(Me.chkControlAll.Value)
End Sub

Private Sub cmdConfigAppBlock_Click()
    frmEditAppBlock.Show vbModal
End Sub
Private Sub txtRefreshRate_KeyUp(KeyCode As Integer, Shift As Integer)
    'Dim strTemp As String
    'strTemp = "1234567890" & vbBack
    'MsgBox KeyCode
    'If InStr(1, strTemp, Chr$(KeyCode)) > 0 Then
    If Len(Me.txtRefreshRate.Text) >= 4 Then Me.txtRefreshRate.Text = Left$(Me.txtRefreshRate.Text, 3)
    If KeyCode = 8 Or (KeyCode >= 96 And KeyCode <= 105) Then
        If Len(Me.txtRefreshRate.Text) = 0 Then Me.txtRefreshRate.Text = 0
        CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "RefreshRate", Me.txtRefreshRate.Text
        Me.tmrAutoRefresh.interval = Me.txtRefreshRate.Text * 1000
    End If
End Sub
'========================================'
' TOOLS                                  '
'========================================'

'Sub Menu in Tools
'-----------------
Private Sub cmdEnableReg_Click()
    Me.frameTool_ScanReg.Visible = False
    Me.frameTool_Process.Visible = False
    Me.frameTool_Startup.Visible = False
    Me.frameTool_EnableReg.Visible = True
    
    Call LoadRegistry
End Sub

Private Sub cmdScanInvalidReg_Click()
    Me.frameTool_ScanReg.Visible = True
    Me.frameTool_Process.Visible = False
    Me.frameTool_Startup.Visible = False
    Me.frameTool_EnableReg.Visible = False
End Sub

Private Sub cmdTool_Startup_Click()
    Me.frameTool_ScanReg.Visible = False
    Me.frameTool_Process.Visible = False
    Me.frameTool_Startup.Visible = True
    Me.frameTool_EnableReg.Visible = False
    'retrieve all startup reg
    Call GetAllRun
End Sub

Private Sub cmdTool_ProcessMan_Click()
    Me.frameTool_ScanReg.Visible = False
    Me.frameTool_Process.Visible = True
    Me.frameTool_Startup.Visible = False
    Me.frameTool_EnableReg.Visible = False
End Sub

'Scan Registry
'-------------

Private Sub lvErrorRegKey_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Item.Checked = False
    Else
        Item.Checked = True
    End If
End Sub

Private Sub lvErrorRegKey_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lvErrorRegKey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'right click
    If Button = 2 Then
        'select item first
        'Call lvErrorRegKey_ItemClick
        PopupMenu mnuop, , Me.lvErrorRegKey.Left + Me.frameTool.Left + Me.frameTool_ScanReg.Left + X, Me.lvErrorRegKey.Top + Me.frameTool.Top + Me.frameTool_ScanReg.Top + Y
    End If
End Sub

Private Sub mnuOpSelAll_Click()
    Dim i As Long
    With Me.lvErrorRegKey
    For i = 1 To .ListItems.Count
        'checked
        .ListItems.Item(i).Checked = True
    Next
    End With
End Sub

Private Sub mnuopUnSelAll_Click()
    Dim i As Long
    With Me.lvErrorRegKey
    For i = 1 To .ListItems.Count
        'unchecked
        .ListItems.Item(i).Checked = False
    Next
    End With
End Sub


Private Sub cmdStartStop_Click()
    
    'button CAPTION
    If cmdStartStop.Caption = "&Start" Then 'start => stop
        cmdStartStop.Caption = "&Stop"
    Else
        cmdStartStop.Caption = "&Start"     'stop => start
        cReg.StopSearch
        txtCurKey.Text = ""
        Exit Sub
    End If
    
    'clear items
    Me.lvErrorRegKey.ListItems.Clear
    txtCurKey.Text = ""
    lblScanRegStatus.Caption = "Scanning :"
    lblScanRegError.Caption = 0
    
    'SEARCH START
    '============
    '0=HKEY_ALL
    cReg.RootKey = 0
    'Don't search in any specific subkey (Search in all subkeys)
    cReg.SubKey = ""
    'Only find errors in value names and value values
    cReg.SearchFlags = KEY_NAME * 0 + VALUE_NAME * 1 + VALUE_VALUE * 1 + WHOLE_STRING * 0
    'Search for registry values with the suffix "C:\"
    cReg.SearchString = "C:\"
    'Start searching for invalid registry values
    cReg.DoSearch
    '=============
    'SEARCH FINISH
    
    txtCurKey.Text = ""
End Sub

Private Sub cmdDeleteInvalidKey_Click()

    Dim removed As Long, i As Integer
    'I don't think this is necessary, but if the registry backup takes a while, this program tells the user to wait.
    txtCurKey.FontSize = 12
    txtCurKey.FontBold = True
    'txtCurKey.Text = "Creating Registry Backup..."
    BackupReg
    'change status
    'txtCurKey.Text = "Registry Backup completed. Cleaning Errors..."
        
    'Loop through every item in lvwRegErrors
    With Me.lvErrorRegKey
    For i = 1 To .ListItems.Count
        'checked to be deleted
        If .ListItems.Item(i).Checked = True Then
            'Delete the registry error and mark the item as removed
            DeleteRegKey GetClassKey(.ListItems.Item(i).SubItems(1)), .ListItems.Item(i).SubItems(2), .ListItems.Item(i).SubItems(3)
            '.ListItems.Item(i).Text = "Cleaned"
            .ListItems.Item(i).Icon = 2
            .ListItems.Item(i).SmallIcon = 2
            removed = removed + 1
        End If
    Next
    End With
    'no deletion
    If removed = 0 Then GoTo endSub
    'change last status
    txtCurKey.Text = "Cleaning Errors completed."
    Call ShowTrayMessage("BG Antivirus 2007 Beta", "Cleaning Registry Errors completed and backup. Cleaned " & removed & " of " & Me.lvErrorRegKey.ListItems.Count & " .")
endSub:
    txtCurKey.FontSize = 8
    txtCurKey.FontBold = False
    txtCurKey.Text = ""
    
End Sub

'Create a backup of the registry, using the "regedit.exe /e" command takes too long.
Public Sub BackupReg()

    Dim i As Integer
    Dim TheKey As String
    Dim TheValue As String
    Dim DefaultValue As Boolean
    Dim BackupFilename As String
    Dim f As Long
    
    'check folder backup
    If FileorFolderExists(App.Path & "\RegBak") = False Then MkDir App.Path & "\RegBak"
    
    BackupFilename = App.Path & "\RegBak\Backup_" & Format(Date, "dd-mm-yyyy") & "_" & Format(Time, "hh-nn-ss") & ".reg"
    'MsgBox BackupFilename
    
    'open file to write
    f = FreeFile
    Open BackupFilename For Output As #f
    Print #f, "REGEDIT4" & vbCrLf
    'Loops through all the checked items and saves the values reg file
    With lvErrorRegKey
    For i = 1 To .ListItems.Count
        If .ListItems.Item(i).Checked = True Then
        
            TheKey = ReverseString(.ListItems.Item(i).SubItems(1) & "\" & .ListItems.Item(i).SubItems(2))
            'the value might ends with a "\", then it's the default value for that key
            If Right$(TheKey, 1) = "\" Then DefaultValue = True: TheKey = Mid(TheKey, 2)
            TheValue = Chr(34) & Replace(ReverseString(Mid(TheKey, 1, InStr(1, TheKey, "\") - 1)), "\", "\\") & Chr(34)
            TheKey = ReverseString(Mid(TheKey, InStr(1, TheKey, "\") + 1))
            If DefaultValue = True Then TheValue = "@"
            'add key to .reg file
            Print #f, "[" & TheKey & "]" '& vbCrLf
            Print #f, TheValue & "=" & Chr(34) & .ListItems.Item(i).SubItems(3) & Chr(34) '& vbCrLf
            
        End If
    Next
    Close #f
    End With
    
End Sub


'class cRegSearch event
Private Sub cReg_SearchFound(ByVal sRootKey As String, ByVal sKey As String, ByVal sValue As Variant, ByVal lFound As FOUND_WHERE)
    
    Dim KN As String    'KeyName
    Dim FileorPath As String  'File Path
    Dim X As ListItem
    
    'WHERE
    Select Case lFound
    Case FOUND_IN_KEY_NAME
        KN = "KEY_NAME"
    Case FOUND_IN_VALUE_NAME
        KN = "VALUE NAME"
    Case FOUND_IN_VALUE_VALUE
        KN = "VALUE VALUE"
    End Select

    FileorPath = sValue
    
    'Condition !
    'If Right$(FileorPath, 4) = ".EXE" Or Right$(FileorPath, 4) = ".exe" Or Right$(FileorPath, 4) = ".DLL" Or Right$(FileorPath, 4) = ".dll" Or Right$(FileorPath, 4) = ".OCX" Or Right$(FileorPath, 4) = ".ocx" Or Right$(FileorPath, 4) = ".SYS" Or Right$(FileorPath, 4) = ".sys" Or Right$(FileorPath, 4) = ".VXD" Or Right$(FileorPath, 4) = ".vxd" Or Right$(FileorPath, 3) = ".AX" Or Right$(FileorPath, 3) = ".ax" Then
    
    'check if actual file exist as in registry
    If FileorFolderExists(FormatValue(FileorPath)) = False Then 'not exist => invalid key
        
        If intSettingRegOption = 1 Then 'scan all
            'add to list for any key
            With Me.lvErrorRegKey
                Set X = .ListItems.Add(, , KN, 5, 5)
                X.SubItems(1) = sRootKey
                X.SubItems(2) = sKey
                X.SubItems(3) = sValue
            End With
            Set X = Nothing
            'add to counter
            Me.lblScanRegError.Caption = Int(Me.lblScanRegError.Caption) + 1
        Else    'scan specific extension
            'MsgBox FileorPath
            'MsgBox Mid(FileorPath, InStr(1, FileorPath, ".") + 1, 3)
            'MsgBox InStr(1, LCase(Mid(FileorPath, InStr(1, FileorPath, ".") + 1, 3)), LCase(strScanRegExt))
            'If InStr(1, Right$(FileorPath, 3), strScanRegExt, vbTextCompare) > 0 Then    'found in extension
            If InStr(1, strScanRegExt, Mid(FileorPath, InStr(1, FileorPath, ".") + 1, 3), vbTextCompare) > 0 Then    'found in extension
                With Me.lvErrorRegKey
                    Set X = .ListItems.Add(, , KN, 5, 5)
                    X.SubItems(1) = sRootKey
                    X.SubItems(2) = sKey
                    X.SubItems(3) = sValue
                End With
                Set X = Nothing
                'add to counter
                Me.lblScanRegError.Caption = Int(Me.lblScanRegError.Caption) + 1
            End If
        End If
    End If
    
End Sub

'class cRegSearch event
Private Sub cReg_SearchFinished(ByVal lReason As Long)
    
    If lReason = 0 Then
        Me.lblScanRegStatus.Caption = "Scan Completed"
    ElseIf lReason = 1 Then
        Me.lblScanRegStatus.Caption = "Scan Cancelled"
    Else
        Me.lblScanRegStatus.Caption = "Scan Error"
    End If
    cmdStartStop.Caption = "&Start"
End Sub

'class cRegSearch event, when change key to search
Private Sub cReg_SearchKeyChanged(ByVal sFullKeyName As String)
    txtCurKey.Text = sFullKeyName
End Sub

'========================================'
' Process                                '
'========================================'

'check all process
Sub CheckProcess()
        
    Dim f As Long, strAttPro As String
    f = FreeFile
    Open App.Path & "\AttPro.bin" For Binary As f
    strAttPro = Input$(LOF(f), 1)
    Close f
    'MsgBox strAttPro
    
    Dim intBlockRisk As Integer, haveRisk As Boolean
    intBlockRisk = Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "BlockRisk"))
    
    Dim i As Long
    With Me.lvProcess
        For i = 1 To .ListItems.Count
            If .ListItems(i).SubItems(1) = "System Process" Then
                'MsgBox .ListItems(i).Text & " SYSTEM"
                .ListItems(i).SmallIcon = 7
            Else
                If InStr(1, strAttPro, .ListItems(i).Text, vbTextCompare) > 0 Then 'check in virus list
                    .ListItems(i).SmallIcon = 5 'mark risk
                    haveRisk = True 'mark variable
                Else
                    .ListItems(i).SmallIcon = 6
                End If
            End If
        Next
        'run only when Enable Auto Block, and Have Risk too
        If haveRisk = True And intBlockRisk = 1 Then
            For i = 1 To .ListItems.Count
                If .ListItems(i).SmallIcon = 5 Then  'risk item
                    Process_Kill .ListItems(i).SubItems(2)  'kill process
                End If
            Next
            'after kill all risk, refresh list
            Call GetProcessList(Me.lvProcess)
        End If
    End With
    
End Sub

'process clicked => get detail
Private Sub lvProcess_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Me.lvProcessDetail.ListItems.Clear
    pname = ""
    GetProcessesPids Trim(LCase(Item.Text)), procpids
    'all application instances
    'e.g multiple internet explorer windows
    Dim i As Integer
    i = 1
    While procpids(i) <> -1
        PID = procpids(i)
        GetWindowList Me.lvProcessDetail
        i = i + 1
    Wend
End Sub

'right click for dropdown menu

Private Sub lvProcess_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Me.lvProcess.ListItems(Me.lvProcess.SelectedItem.Index).SmallIcon = 7 Then 'check on icon instead (can check on subitem(1)) 'system process
            mnuProMenu_Ban.Enabled = False
            mnuProMenu_Safe.Enabled = False
            PopupMenu mnuProMenu
        Else
            If Me.lvProcess.ListItems(Me.lvProcess.SelectedItem.Index).SmallIcon = 5 Then   'attention process
                mnuProMenu_Ban.Enabled = False
                mnuProMenu_Safe.Enabled = True
                PopupMenu mnuProMenu
            Else    'safe process
                mnuProMenu_Ban.Enabled = True
                mnuProMenu_Safe.Enabled = False
                PopupMenu mnuProMenu
            End If
        End If
    End If

End Sub

Private Sub cmdProcessRefresh_Click()
    'clear items from all listviews
    Me.lvProcess.ListItems.Clear
    Me.lvProcessDetail.ListItems.Clear
    'refresh data
    Call GetProcessList(Me.lvProcess)
    Call CheckProcess
End Sub

Private Sub tmrAutoRefresh_Timer()
    Call cmdProcessRefresh_Click
    If Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "BlockRisk")) = 1 Then
        Call CheckProcess
    End If
End Sub

Private Sub cmdProcessEnd_Click()
    If MsgBox("Do you really want to end this process?", vbYesNo + vbQuestion, "End Process : " & Me.lvProcess.SelectedItem.Text) = vbYes Then
        Dim Pro_ID As Long
        Pro_ID = Me.lvProcess.SelectedItem.SubItems(2)
        Process_Kill Pro_ID
        Call cmdProcessRefresh_Click
    End If
End Sub

'========================================'
' STARTUP                                '
'========================================'

'Enumerate from all RUN
Private Sub GetAllRun()
    On Error Resume Next
    Dim X As ListItem, hKey As Long, lCount As Long, i As Long
    lvStartUp.ListItems.Clear
    'Enumerate from HKEY_LOCAL_MACHINE , Run
    hKey = OpenKey(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Run")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        Set X = lvStartUp.ListItems.Add(, , EnumValue(hKey, i))
        X.SubItems(1) = GetKeyValue(hKey, EnumValue(hKey, i))
        X.SubItems(2) = "HKEY_LOCAL_MACHINE"
        Set X = Nothing
    Next i
    
    'Enumerate from HKEY_LOCAL_MACHINE , RunServices
    hKey = OpenKey(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\RunServices")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        Set X = lvStartUp.ListItems.Add(, , EnumValue(hKey, i))
        X.SubItems(1) = GetKeyValue(hKey, EnumValue(hKey, i))
        X.SubItems(2) = "HKEY_LOCAL_MACHINE (Service)"
        Set X = Nothing
    Next i
    
    'Enumerate from HKEY_CURRENT_USER , Run
    hKey = OpenKey(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Run")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        Set X = lvStartUp.ListItems.Add(, , EnumValue(hKey, i))
        X.SubItems(1) = GetKeyValue(hKey, EnumValue(hKey, i))
        X.SubItems(2) = "HKEY_CURRENT_USER"
        Set X = Nothing
    Next i
    
    'get startup from tasks
    Dim fso As New FileSystemObject
    Dim sFolder As Folder
    Dim sFiles As Files
    Dim sFile As File
    Set sFolder = fso.GetFolder("C:\Windows\Tasks")
    Set sFiles = sFolder.Files
    If sFiles.Count > 0 Then
        For Each sFile In sFiles
            Set X = Me.lvStartUp.ListItems.Add(, , sFile.Name)
            X.SubItems(1) = sFile.Path
            X.SubItems(2) = "Tasks"
            Set X = Nothing
        Next
    End If
    'get startup from current user startup
    Dim strUserProfile As String
    strUserProfile = Environ$("UserProfile") & "\Start Menu\Programs\Startup"
    'Set sFolder = fso.GetFolder("%userprofile%\Start Menu\Programs\Startup")
    Set sFolder = fso.GetFolder(strUserProfile)
    'MsgBox sFolder.Path
    Set sFiles = sFolder.Files
    If sFiles.Count > 0 Then
        For Each sFile In sFiles
            'Set X = Me.lvStartUp.ListItems.Add(, , "User Startup")
            Set X = Me.lvStartUp.ListItems.Add(, , sFile.Name)
            X.SubItems(1) = sFile.Path
            X.SubItems(2) = "User Startup"
            Set X = Nothing
        Next
    End If
    'get startup from current all user startup
    Set sFolder = fso.GetFolder("C:\Documents and Settings\All Users\Start Menu\Programs\Startup")
    Set sFiles = sFolder.Files
    If sFiles.Count > 0 Then
        For Each sFile In sFiles
            Set X = Me.lvStartUp.ListItems.Add(, , sFile.Name)
            X.SubItems(1) = sFile.Path
            X.SubItems(2) = "All User Startup"
            Set X = Nothing
        Next
    End If
End Sub

'show Startup Folder /TASK?
Sub StartUpFolder()
'    txtCmdLine.Text = ""
'    txtName.Text = ""
'    If optStartMenu.Value = True Then
'        optStartMenu.FontBold = True
'        optRunServices.FontBold = False
'        optRun.FontBold = False
'        optRun2.FontBold = False
'        optWinINI.FontBold = False
'        Option1.FontBold = False
'    End If
'    ShellExecute 0, "open", CheckFolderID(StartUp), "", CheckFolderID(StartUp), 1
End Sub

'show Win.ini file
Sub ShowWinINIFile()
'    txtCmdLine.Text = ""
'    txtName.Text = ""
'    ShellExecute 0, "open", "notepad.exe", WinDir & "\win.ini", "", 1
'    If optWinINI.Value = True Then
'        optWinINI.FontBold = True
'        optRunServices.FontBold = False
'        optRun.FontBold = False
'        optStartMenu.FontBold = False
'        optRun2.FontBold = False
'        Option1.FontBold = False
'    End If
End Sub

'show System.ini file
Sub ShowSystemINIFile()
'    txtCmdLine.Text = ""
'    txtName.Text = ""
'    If Option1.Value = True Then
'        Option1.FontBold = True
'        optRunServices.FontBold = False
'        optRun.FontBold = False
'        optWinINI.FontBold = False
'        optStartMenu.FontBold = False
'        optRun2.FontBold = False
'    End If
'
'    MsgBox "Please note that the line n3 must be something like this" & vbCrLf & " [shell=Explorer.exe]  if you have something else" & vbCrLf & "Its possible that the system is loading in abnormal way", vbInformation, "Warning do not edit this if you dont know"
'    ShellExecute 0, "open", "notepad.exe", WinDir & "\system.ini", "", 1
End Sub

Private Sub cmdStartUp_Del_Click()
    With Me.lvStartUp
    Dim i As Long
    Dim tmp As Long
    Dim fso As New FileSystemObject
    For i = 1 To .ListItems.Count
        'checked to be deleted
        If .ListItems.Item(i).Checked = True Then
            'Delete startup
            'If .ListItems.Item(i).SubItems(2) = "ScheduledTask" Then    'Schedule Task   'not added yet
            'DeleteStartup GetClassKey("HKEY_LOCAL_MACHINE"), "Software/Microsoft/Windows/CurrentVersion/Policies/Explorer/Run", TASK
            If .ListItems.Item(i).SubItems(2) = "HKEY_LOCAL_MACHINE (Service)" Then 'run service startup
                DeleteStartup GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\RunServices", .ListItems.Item(i).Text
            ElseIf .ListItems.Item(i).SubItems(2) = "HKEY_LOCAL_MACHINE" Or .ListItems.Item(i).SubItems(2) = "HKEY_CURRENT_USER" Then 'normal startup
                If .ListItems.Item(i).Text = "BGAntivirus" Then
                    Me.chkAutoStart.Value = 0    'delete autostart
                    Exit For
                End If
                DeleteStartup GetClassKey(.ListItems.Item(i).SubItems(2)), "Software\Microsoft\Windows\CurrentVersion\Run", .ListItems.Item(i).Text
            Else    'If .ListItems.Item(i).SubItems(2) = "Startup" Then  'startup folder, else all are file
                'Kill .ListItems.Item(i).SubItems(1)  'kill filepath
                fso.DeleteFile .ListItems.Item(i).SubItems(1), True 'file system need to be deleted by force
            End If
        End If
    Next
    Set fso = Nothing
    End With
    'refresh startup run
    Call GetAllRun
End Sub

'========================================'
' REGISTRY                               '
'========================================'

Sub LoadRegistry()
    'Check in registry
    'Current User
    'Me.chkCUTaskmgr.Value = Val(ReadValue(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr"))
'    Me.chkCUNoLogoff.Value = Val(ReadValue(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\Explorer", "NoLogoff"))
'    Me.chkCUNoCLose.Value = Val(ReadValue(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\Explorer", "NoClose"))
'    Me.chkCUNoLock.Value = Val(ReadValue(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableLockWorkstation"))
'    Me.chkCUNoChangePwd.Value = Val(ReadValue(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableChangePassword"))
'    Me.chkCURegTool.Value = Val(ReadValue(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"))
'    Me.chkCUNoCmd.Value = Val(ReadValue(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD"))
'    Me.chkCUNoFolderOption.Value = Val(ReadValue(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"))
    Me.chkCUTaskmgr.Value = Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr"))
    Me.chkCUNoLogoff.Value = Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\Explorer", "NoLogoff"))
    Me.chkCUNoCLose.Value = Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\Explorer", "NoClose"))
    Me.chkCUNoLock.Value = Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableLockWorkstation"))
    Me.chkCUNoChangePwd.Value = Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableChangePassword"))
    Me.chkCURegTool.Value = Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"))
    Me.chkCUNoCmd.Value = Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD"))
    Me.chkCUNoFolderOption.Value = Val(GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"))

    'Local Machine
'    Me.chkLMNoFolderOption.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"))
'    Me.chkLMRegTool.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"))
'    Me.chkLMNoTaskmgr.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr"))
'    Me.chkLMNoSRConfig.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig"))
'    Me.chkLMNoSR.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR"))
'    Me.chkLmLimitSR.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing"))
'    Me.chkLMNoMSI.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI"))
    Me.chkLMNoFolderOption.Value = Val(GetString(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"))
    Me.chkLMRegTool.Value = Val(GetString(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"))
    Me.chkLMNoTaskmgr.Value = Val(GetString(GetClassKey("HKEY_LOCAL_MACHINE"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr"))
    Me.chkLMNoSRConfig.Value = Val(GetString(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig"))
    Me.chkLMNoSR.Value = Val(GetString(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR"))
    Me.chkLmLimitSR.Value = Val(GetString(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing"))
    Me.chkLMNoMSI.Value = Val(GetString(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI"))

End Sub

Private Sub chkCUNoChangePwd_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableChangePassword", Val(Me.chkCUNoChangePwd.Value)
End Sub

Private Sub chkCUNoCLose_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\Explorer", "NoClose", Val(Me.chkCUNoCLose.Value)
End Sub

Private Sub chkCUNoCmd_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD", Val(Me.chkCUNoCmd.Value)
End Sub

Private Sub chkCUNoFolderOption_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", Val(Me.chkCUNoFolderOption.Value)
End Sub

Private Sub chkCUNoLock_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableLockWorkstation", Val(Me.chkCUNoLock.Value)
End Sub

Private Sub chkCUNoLogoff_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\Explorer", "NoLogoff", Val(Me.chkCUNoLogoff.Value)
End Sub

Private Sub chkCURegTool_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", Val(Me.chkCURegTool.Value)
End Sub

Private Sub chkCUTaskmgr_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr", Val(Me.chkCUTaskmgr.Value)
End Sub

Private Sub chkLmLimitSR_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing", Val(Me.chkLmLimitSR.Value)
End Sub

Private Sub chkLMNoFolderOption_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", Val(Me.chkLMNoFolderOption.Value)
End Sub

Private Sub chkLMNoMSI_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI", Val(Me.chkLMNoMSI.Value)
End Sub

Private Sub chkLMNoSR_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR", Val(Me.chkLMNoSR.Value)
End Sub

Private Sub chkLMNoSRConfig_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig", Val(Me.chkLMNoSRConfig.Value)
End Sub

Private Sub chkLMNoTaskmgr_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr", Val(Me.chkLMNoTaskmgr.Value)
End Sub

Private Sub chkLMRegTool_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", Val(Me.chkLMRegTool.Value)
End Sub

Private Sub cmdCleanReg_Click()
    Call CleanReg
    Me.chkCUNoCLose.Value = 0
    Me.chkCUNoLogoff.Value = 0
    Me.chkCUNoLock.Value = 0
    Me.chkCUNoChangePwd.Value = 0
    Call LoadRegistry
End Sub

'About
'=============================
Private Sub tmrAbout_Timer()
    If txtAbout.Top <= -8175 Then txtAbout.Top = 8175
    txtAbout.Top = txtAbout.Top - 15
End Sub
