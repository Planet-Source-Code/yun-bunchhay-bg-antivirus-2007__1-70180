VERSION 5.00
Begin VB.Form frmAlert 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BG Antivirus / Application Blocker"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8400
   Icon            =   "frmAlert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPDeny 
      Caption         =   "Al&ways Deny"
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdDeny 
      Caption         =   "&Deny This &Time"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdAllow 
      Caption         =   "&Allow This Time"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdPAllow 
      Caption         =   "A&lways Allow"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox txtFilePath 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   2700
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   5445
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Prompt for Action"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Status : "
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
      Left            =   2040
      TabIndex        =   1
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   240
      Picture         =   "frmAlert.frx":0CCA
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File :"
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
      Left            =   2040
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim fso As New FileSystemObject
'Dim fn As File

Private Sub Form_Load()
    'Dim fso As New FileSystemObject
    'MsgBox fso.GetBaseName(Command$)
    'MsgBox fso.GetFileName(Command$)
    
    'avoid open it own file
    If Command$ = "" Then End
    
    ' check setting
    If ReadRegShell("HKCU\Software\BGAntivirus\AppBlock") = 0 Then
        'not enable => allow all file
        Call cmdAllow_Click
    End If
    
    'checked => manage file
    '======================
    'manage all files
    '======================
    Dim bCount As Long
    Dim i As Long
    
    If Val(ReadRegShell("HKCU\Software\BGAntivirus\ControlAll")) = 1 Then
        Me.txtFilePath.Text = Command$
        ' check setting
        ' allow
        Dim aCount As Long
        aCount = ReadRegShell("HKCU\Software\BGAntivirus\Allow\aCount")
        
        For i = 1 To aCount
            'in allow list
            'MsgBox GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", CStr(i))
            If Command$ = ReadRegShell("HKEY_CURRENT_USER\Software\BGAntivirus\Allow\" & CStr(i)) Then
                Call cmdAllow_Click
                'automatically END
            End If
        Next
        
        ' deny
        
        'Dim i As Long
        bCount = ReadRegShell("HKCU\Software\BGAntivirus\Ban\bCount")
        
        For i = 1 To bCount
            'in ban list
            'MsgBox GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", CStr(i))
            'Set fn = fso.GetFile(ReadRegShell("HKEY_CURRENT_USER\Software\BGAntivirus\Ban\" & CStr(i)))
            'MsgBox fn.ShortPath
            If Command$ = ReadRegShell("HKEY_CURRENT_USER\Software\BGAntivirus\Ban\" & CStr(i)) Then
                'Call cmdDeny_Click
                Me.lblStatus.Caption = "Application Blocked"
                Me.cmdPDeny.Enabled = False
                Me.cmdDeny.Enabled = False
                Me.cmdPAllow.Enabled = False
            End If
        Next
    Else
        'manage only BLOCKED Files
        bCount = ReadRegShell("HKCU\Software\BGAntivirus\Ban\bCount")
        
        For i = 1 To bCount
            'in ban list
            If Command$ = ReadRegShell("HKEY_CURRENT_USER\Software\BGAntivirus\Ban\" & CStr(i)) Then
                'Call cmdDeny_Click
                Me.lblStatus.Caption = "Application Blocked"
                Me.cmdPDeny.Enabled = False
                Me.cmdDeny.Enabled = False
                Me.cmdPAllow.Enabled = False
                Exit For
            End If
        Next
        'check if blocked or not
        If Me.lblStatus.Caption <> "Application Blocked" Then
            'not found, not block => allow
            Call cmdAllow_Click
        End If
    End If
    'New app run normal
    
End Sub

Private Sub cmdAllow_Click()
    Shell Command$, vbNormalFocus
    'MsgBox "alloe. end"
    End
End Sub

Private Sub cmdDeny_Click()
    End
End Sub

Private Sub cmdPAllow_Click()
    Dim aCount As Long
    aCount = Val(ReadRegShell("HKCU\Software\BGAntivirus\Allow\aCount"))
        
    'update count
    aCount = aCount + 1
    
    'update in registry
    Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", 1, "aCount", CStr(aCount))
    'Call CreateDwordValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", "aCount", aCount)
    'WriteRegShell "HKEY_CURRENT_USER\Software\BGAntivirus\Allow\aCount", CStr(aCount)
    
    'add new app to allow list
    Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", 1, CStr(aCount), Command$)
    Call cmdAllow_Click
    'End
End Sub

Private Sub cmdPDeny_Click()
    Dim bCount As Long
    bCount = Val(ReadRegShell("HKCU\Software\BGAntivirus\Ban\bCount"))
    
    'update count
    bCount = bCount + 1
    
    'update in registry
    Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", 1, "bCount", CStr(bCount))
    'Call CreateDwordValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", "bCount", bCount)
    'WriteRegShell "HKEY_CURRENT_USER\Software\BGAntivirus\Ban\bCount", CStr(bCount)
    
    'add new app to allow list
    Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", 1, CStr(bCount), Command$)
    Call cmdDeny_Click
    'End
End Sub
