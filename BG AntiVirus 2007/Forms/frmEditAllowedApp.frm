VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditAppBlock 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Application Blocker"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   Icon            =   "frmEditAllowedApp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAllow 
      Caption         =   "Allow"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeny 
      Caption         =   "Deny"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   6360
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvAppList 
      Height          =   5775
      Left            =   315
      TabIndex        =   0
      Top             =   360
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   10186
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Path"
         Object.Width           =   9613
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   2390
      EndProperty
   End
End
Attribute VB_Name = "frmEditAppBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============
' FORM LOAD
'===============
Private Sub Form_Load()
    Me.lvAppList.ListItems.Clear
    Call LoadBanList
    Call LoadAllowList
End Sub

'===================
'FUNCTIONS
'===================
Sub LoadBanList()
    Dim bCount As Long
    Dim i As Long
    If GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", "Count") <> "" Then
        bCount = GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", "Count")
    Else
        bCount = 0
    End If
    Dim X As ListItem
    For i = 1 To bCount
        Set X = Me.lvAppList.ListItems.Add(, , GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", CStr(i)))
        X.SubItems(1) = "Deny"
        Set X = Nothing
    Next
End Sub

Sub LoadAllowList()
    Dim aCount As Long
    Dim i As Long
    If GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", "Count") <> "" Then
        aCount = GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", "Count")
    Else
        aCount = 0
    End If
    Dim X As ListItem
    For i = 1 To aCount
        Set X = Me.lvAppList.ListItems.Add(, , GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", CStr(i)))
        X.SubItems(1) = "Allow"
        Set X = Nothing
    Next
End Sub

'=====================
' EVENTS
'=====================

Private Sub cmdAdd_Click()
    Me.CommonDialog1.Filter = "*.exe | Executable File"
    Me.CommonDialog1.DialogTitle = "Browse for EXE File"
    Me.CommonDialog1.ShowOpen
    'file selected
    If Me.CommonDialog1.FileName <> "" Then
        'ask for action
        Dim X As ListItem
        If MsgBox("Allow This Application To Run", vbYesNo + vbDefaultButton1, "Add New Application") = vbYes Then
            'add to list
            Set X = Me.lvAppList.ListItems.Add(, , Me.CommonDialog1.FileName)
            X.SubItems(1) = "Allow"
            Set X = Nothing
        Else
            'add to list
            Set X = Me.lvAppList.ListItems.Add(, , Me.CommonDialog1.FileName)
            X.SubItems(1) = "Deny"
            Set X = Nothing
        End If
    End If
End Sub

Private Sub cmdAllow_Click()
    Me.lvAppList.SelectedItem.SubItems(1) = "Allow"
End Sub

Private Sub cmdDeny_Click()
    Me.lvAppList.SelectedItem.SubItems(1) = "Deny"
End Sub

Private Sub cmdRemove_Click()
    If MsgBox("Are you sure to remove the selected item?", vbYesNoCancel + vbDefaultButton3, "Remove Application") = vbYes Then
        Me.lvAppList.ListItems.Remove Me.lvAppList.SelectedItem.Index
    End If
End Sub

Private Sub cmdSave_Click()
    Dim aCount As Long, bCount As Long
    Dim i As Long
    
    'clear previous list in registry
    'from allow list
    If GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", "Count") <> "" Then
        aCount = GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", "Count")
    Else
        aCount = 0
    End If
    For i = 1 To aCount
        'use delete registry (function in startup)
        DeleteStartup GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", CStr(i)
    Next
    DeleteStartup GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", "Count"
    aCount = 0
    
    If GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", "Count") <> "" Then
        bCount = GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", "Count")
    Else
        bCount = 0
    End If
    For i = 1 To bCount
        'use delete registry (function in startup)
        DeleteStartup GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", CStr(i)
    Next
    DeleteStartup GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", "Count"
    bCount = 0
    
    'add from list to registry
    For i = 1 To Me.lvAppList.ListItems.Count
        If Me.lvAppList.ListItems(i).SubItems(1) = "Allow" Then
            'update count
            aCount = aCount + 1
            'update in registry / count
            Call CreateDwordValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", "Count", aCount)
            'add app to allow list
            Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", 1, CStr(aCount), Me.lvAppList.ListItems(i).Text)
        Else    'deny
            'update count
            bCount = bCount + 1
            'update in registry / count
            Call CreateDwordValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", "Count", bCount)
            'add app to allow list
            Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", 1, CStr(bCount), Me.lvAppList.ListItems(i).Text)
        End If
    Next
    MsgBox "Save Complete", vbInformation, "Save Application List"
End Sub
