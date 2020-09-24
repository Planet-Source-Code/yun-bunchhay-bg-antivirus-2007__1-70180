VERSION 5.00
Begin VB.Form frmTrayMsg 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Close"
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Timer TimerUnload1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2760
      Top             =   2040
   End
   Begin VB.Timer TimerUnloadTime 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3360
      Top             =   2040
   End
   Begin VB.Timer TimerLoad 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3960
      Top             =   2040
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "frmTrayMsg.frx":0000
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
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
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmTrayMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Long
Dim j As Long

Private Sub cmdClose_Click()
    Me.TimerUnload1.Enabled = True
End Sub

Private Sub Form_Load()
    SetTopMostWindow Me.hwnd, True
    Call MakeTransparent(Me, 230)
    Me.Left = Screen.Width - 4605
    i = 0
    j = 0
    Me.TimerLoad.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Me.TimerLoad.Enabled = False
'    Me.TimerUnload1.Enabled = False
'    Me.TimerUnloadTime.Enabled = False
End Sub

Private Sub TimerLoad_Timer()
    If i < 2715 Then
        Me.Top = Screen.Height - i '- 500    '500 for taskbar
        i = i + 20
    Else
        Me.TimerLoad.Enabled = False
        Me.TimerUnloadTime.Enabled = True
    End If
End Sub

Private Sub TimerUnload1_Timer()
    If j < 230 Then
        'Me.Top = Screen.Height - j '- 500    '500 for taskbar
        'fade out
        Call MakeTransparent(Me, 230 - j)
        j = j + 5
    Else
        Me.TimerUnload1.Enabled = False
        Unload Me
    End If
End Sub

Private Sub TimerUnloadTime_Timer()
    Me.TimerUnload1.Enabled = True
    Me.TimerUnloadTime.Enabled = False
End Sub
