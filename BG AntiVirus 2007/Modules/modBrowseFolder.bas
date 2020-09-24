Attribute VB_Name = "modBrowseFolder"
Option Explicit

Private Type BrowseInfo
lngHwnd As Long
pIDLRoot As Long
pszDisplayName As Long
lpszTitle As Long
ulFlags As Long
lpfnCallback As Long
lParam As Long
iImage As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem _
    As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "Shell32" (lpbi _
    As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "Shell32" _
    (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function BrowseForFolder(ByVal strPrompt As String) As _
    String
On Error GoTo ehBrowseForFolder
Dim intNull As Integer
Dim lngIDList As Long, lngResult As Long
Dim strPath As String
Dim udtBI As BrowseInfo
With udtBI
.lngHwnd = 0
.lpszTitle = lstrcat(strPrompt, "")
.ulFlags = BIF_RETURNONLYFSDIRS
End With
lngIDList = SHBrowseForFolder(udtBI)
If lngIDList <> 0 Then
strPath = String(MAX_PATH, 0)
lngResult = SHGetPathFromIDList(lngIDList, strPath)
Call CoTaskMemFree(lngIDList)
intNull = InStr(strPath, vbNullChar)
If intNull > 0 Then
strPath = Left(strPath, intNull - 1)
End If
End If
BrowseForFolder = strPath
Exit Function
ehBrowseForFolder:
BrowseForFolder = Empty
End Function
