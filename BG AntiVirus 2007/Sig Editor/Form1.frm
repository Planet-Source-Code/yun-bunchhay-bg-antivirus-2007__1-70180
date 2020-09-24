VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BG Antivirus Sig Editor"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5850
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Manua l  l   y          Add"
      Height          =   3735
      Left            =   5520
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox txtLastUpdate 
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox txtCount 
      Enabled         =   0   'False
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   4320
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   4320
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6588
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
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Value"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'new DataType for Virus Signature
Private Type VirusSig

    Name As String
    Type As String
    Value As String
    Action As String
    ActtionVal As String
    
End Type

'new DataType for Virus Signature Info
Private Type VS_Info
    
    VirusCount As Long
    LastUpdate As Date
    
End Type

    'declare Virus Def & info
    Dim VSig() As VirusSig
    Dim VSInfo As VS_Info

Public fName As String

Private Sub Command1_Click()
    On Error GoTo err
    CommonDialog1.ShowOpen
    fName = Me.CommonDialog1.FileName
    Call ReadSig(fName)
err:
End Sub

Public Sub ReadSig(ByVal fn As String)
    
    'declare Virus Def & info
    Dim VSig() As VirusSig
    Dim VSInfo As VS_Info
    ListView1.ListItems.Clear
    Dim f As Long
    On Error GoTo Trap_Error
    f = FreeFile
    Dim X As ListItem
    Open fn For Binary Access Read As #f
        Get #f, , VSInfo
        ReDim VSig(VSInfo.VirusCount - 1) As VirusSig
        Dim i As Integer
        For i = 0 To VSInfo.VirusCount - 1
            Get #f, , VSig(i)
            Set X = ListView1.ListItems.Add(, , VSig(i).Name)
            X.SubItems(1) = VSig(i).Type
            X.SubItems(2) = VSig(i).Value
            X.Checked = True
            Set X = Nothing
        Next
    Close #f
    Me.txtCount.Text = VSInfo.VirusCount
    Me.txtLastUpdate.Text = Format(VSInfo.LastUpdate, "dd/mmmm/yyyy")
    Exit Sub
Trap_Error:
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure GetData of Form frmBinaAccess"
    
End Sub

Public Sub WriteSig()
    
    Dim f As Long, intCounter As Long
    Dim i As Long
    On Error GoTo Trap_Error
    f = FreeFile
    
    For i = 1 To ListView1.ListItems.Count
        'MsgBox ListView1.ListItems(i).Text
        If ListView1.ListItems(i).Checked Then
            intCounter = intCounter + 1
        End If
    Next
    
    ReDim VSig(intCounter - 1) As VirusSig
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked Then
            'ReDim Preserve VSig(UBound(VSig) + 1) As VirusSig
            VSig(i - 1).Name = ListView1.ListItems(i).Text
            VSig(i - 1).Type = ListView1.ListItems(i).SubItems(1)
            VSig(i - 1).Value = ListView1.ListItems(i).SubItems(2)
        End If
    Next
    
    'add 1 for count
    VSInfo.VirusCount = intCounter        'UBound(VSig)
    VSInfo.LastUpdate = Format(Me.txtLastUpdate.Text, "dd/mmmm/yyyy")
    
    'change virus last update
    'VSInfo.LastUpdate = Format("07 June 2007", "Short Date")
    
    Open fName For Binary Access Write As #f
        Put #f, , VSInfo
        For i = 0 To UBound(VSig)
            Put #f, , VSig(i)
        Next
    Close #f

   'On Error GoTo 0
   Exit Sub

Trap_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure"
End Sub

Private Sub Command2_Click()
    Call WriteSig
    Call ReadSig(fName)
End Sub

Private Sub Command3_Click()
    Dim n As String, X As ListItem, vp As String
    Dim CRC As New clsCRC
    CRC.BuildTable
    n = InputBox("Insert virus name:", "Virus Name")
    vp = InputBox("Insert virus path:", "Virus Path")
    Set X = Me.ListView1.ListItems.Add(, , n)
    X.SubItems(1) = "CRC"
    X.SubItems(2) = CRC.GetCRC(vp)
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'MsgBox Item.Checked
    'MsgBox Item.Text
End Sub

