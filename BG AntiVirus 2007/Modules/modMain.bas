Attribute VB_Name = "modMain"
Option Explicit

'CRC32 variable
Public CRC As New clsCRC
'file size to be scanned virus
Public FileSize As Long
'declare Virus Def & info
Public VSig() As VirusSig
Public VSInfo As VS_Info
'declare variable for scan reg extensions
Public intSettingRegOption As Integer
Public strScanRegExt As String
'for faster DoEvents
Declare Function GetInputState Lib "user32.dll" () As Long
'declaration for Classes of Registry
Public Const HKEY_ALL = &H0&
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003

'new ACCESS KEY for delete startup registry
Private Const KEY_ALL_ACCESS = &H3F '((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

'new DataType for Virus Signature
Public Type VirusSig

    Name As String
    Type As String
    Value As String
    Action As String
    ActtionVal As String
    
End Type

'new DataType for Virus Signature Info
Public Type VS_Info
    
    VirusCount As Long
    LastUpdate As Date
    
End Type

' Transparency Constants
'Const LWA_COLORKEY = &H3
Const LWA_ALPHA = &H3
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
'Private Const HWND_TOPMOST = -1
'Private Const SWP_SHOWWINDOW = &H40
'Private Const SWP_NOOWNERZORDER = &H200
Dim Ret As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'startup sub
Sub Main()

    'If Year(Now()) = 2007 Then 'can be used only in 2007 if enable
        Dim FirstStart As String
        FirstStart = GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AppPath")
        'check if it is ever started
        If FirstStart = "" Then
            'CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AppPath", App.Path
            CreateStringValue HKEY_CURRENT_USER, "Software\BGAntivirus", 1, "AppPath", App.Path & "\" & App.EXEName
            'default app setting    'default setting cmd
            CreateStringValue HKEY_CURRENT_USER, "Software\BGAntivirus", 1, "RegExt", "OCX, DLL, EXE, VBS, SYS, VXD"
            
            CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "RefreshRate", 10
            CreateStringValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", 1, "RegExt", "OCX, DLL, EXE, VBS, SYS, VXD"
        Else
            'check if previous open is the same as the current
            If FirstStart <> App.Path & "\" & App.EXEName Then
                'change key if different
                CreateStringValue HKEY_CURRENT_USER, "Software\BGAntivirus", 1, "AppPath", App.Path & "\" & App.EXEName
            End If
        End If
        Call ReadSig
        frmSplash.Show vbModal
        frmMain.Show
    'Else
    '    MsgBox "Beta version is expired. Please find an official release of this software.", vbInformation, "BG Antivirus 2007 Beta Expire"
    'End If
    
End Sub

Function CalculateTime(ByVal interval As Single) As String

    Dim sec As Long, mn As Long, hr As Long
    hr = Int(CSng(interval * 24))
    mn = Int(CSng(interval * 24 * 60))
    sec = Int(CSng(interval * 24 * 60 * 60))
    CalculateTime = hr & " hr " & (mn - (hr * 60)) & " mn " & (sec - (mn * 60)) & " sec."
    
End Function

'Reverse a string
Public Function ReverseString(TheString As String) As String
    Dim i As Integer
    For i = 1 To Len(TheString)
        ReverseString = ReverseString & Mid(Right$(TheString, i), 1, 1)
    Next
End Function

'Returns the long value of the string entered as ROOT_KEYS
Public Function GetClassKey(cls As String) As Variant
    Select Case cls
    Case "HKEY_ALL"
        GetClassKey = HKEY_ALL
    Case "HKEY_CLASSES_ROOT"
        GetClassKey = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
        GetClassKey = HKEY_CURRENT_USER
    Case "HKEY_LOCAL_MACHINE"
        GetClassKey = HKEY_LOCAL_MACHINE
    Case "HKEY_USERS"
        GetClassKey = HKEY_USERS
    Case "HKEY_PERFORMANCE_DATA"
        GetClassKey = HKEY_PERFORMANCE_DATA
    Case "HKEY_CURRENT_CONFIG"
        GetClassKey = HKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA"
        GetClassKey = HKEY_DYN_DATA
    End Select
End Function

'Another Delete Registry Key function
Public Sub DeleteRegKey(ROOTKEYS As ROOT_KEYS, Path As String, sKey As String)
    
    Dim ValKey As String
    Dim SecKey As String, SlashPos As Single
    SlashPos = InStrRev(Path, "\", compare:=vbTextCompare)
    SecKey = Left(Path, SlashPos - 1)    'This will retreive the section key that I need
    ValKey = Right(Path, Len(Path) - SlashPos)    'This will retreive the ValueKey that I need to delete
    DeleteRegKey2 ROOTKEYS, SecKey, ValKey

End Sub
    
'Another Delete Registry Key function
Public Sub DeleteRegKey2(hKey As ROOT_KEYS, strPath As String, strValue As String)
    Dim Ret
    RegCreateKey hKey, strPath, Ret
    RegDeleteValue Ret, strValue
    RegCloseKey Ret
End Sub

Public Function DeleteStartup(lPredefinedKey As Long, sKeyName As String, sValueName As String)

       Dim lRetVal As Long      'result of the SetValueEx function
       Dim hKey As Long         'handle of open key

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = RegDeleteValue(hKey, sValueName)
       RegCloseKey (hKey)
       
End Function

Sub HideToTray()

    'F10 button minimize any window to tray
        Dim pt As POINTAPI
        Dim i As Long, tmp As Long, fhwnd As Long
        Dim traydata As NOTIFYICONDATA
        Dim wndtext As String * 256
        Dim wndlen As String
        Dim clslen As Long
        Dim clsname As String * 260
        Dim clsinfo As WNDCLASS
        Dim tpid As Long, hProc As Long
        Dim n As Long
        Dim Icon As Long
        
'        fhwnd = GetForegroundWindow
'        If fhwnd = 0 Then
'                'Timer1.Enabled = True
'            Exit Sub
'        End If
        
        'get cursor position
        GetCursorPos pt
        tmp = WindowFromPoint(pt.X, pt.Y)
        'is there a window or not
        If tmp = 0 Then Exit Sub
        fhwnd = GetTopLevelWindow(tmp)
        
        Dim hform As Form
        Set hform = New frmMain
        
        'setup attributes for the tray icon
        'extract the icon of the exe
        GetWindowThreadProcessId fhwnd, tpid
        'icon = ExtractIcon(App.hInstance, GetProcessFullPath(tpid), 0)
        Icon = frmMain.Icon
        traydata.hIcon = Icon
        traydata.cbSize = Len(traydata)
        traydata.uID = vbNull
        
        'send the messages to our window
        traydata.hwnd = hform.hwnd
        'we need to have the handle of the window that is in tray
        hform.Tag = CStr(fhwnd) & "#" & CStr(hform.hwnd) & "#" & CStr(Icon)
        traydata.uFlags = NIF_MESSAGE Or NIF_TIP Or NIF_ICON
        'To know what window is when user clicks on the
        'icon we setup the message identifier to the handle
        'of the window
        traydata.uCallbackMessage = WM_MOUSEMOVE
        wndlen = GetWindowText(fhwnd, wndtext, 256)
        traydata.szTip = Mid(wndtext, 1, wndlen) & vbNullChar
        'add to tray menu
        Shell_NotifyIcon NIM_ADD, traydata
        'hide windows
        ShowWindow fhwnd, 0
End Sub

'Transparent Making area
Public Sub MakeTransparent(ByRef frm As Form, ByVal alpha As Long)
    Ret = GetWindowLong(frm.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong frm.hwnd, GWL_EXSTYLE, Ret
    'change ipAlpha for transparency
    SetLayeredWindowAttributes frm.hwnd, 0, alpha, LWA_ALPHA
End Sub

'Tray Message
Public Sub ShowTrayMessage(ByVal sTitle As String, ByVal sMsg As String)
    Dim frm As New frmTrayMsg
    frm.lblTitle.Caption = sTitle
    frm.lblMessage.Caption = sMsg
    'Load frm
    frm.Show
End Sub

