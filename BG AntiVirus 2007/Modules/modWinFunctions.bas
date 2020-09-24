Attribute VB_Name = "modWinFunctions"
'//////////////////////////////////////////////////////////////////////////////'
' This code were explicitly developed for PSC(Planet Source Code) Users,
' as Open Source Project. This code are property of their author.
' The code is provided "as is" WITHOUT any warranty.

' You may use any of this code in you're own application(s).

' Code by Blagoj Janevski
' Please vote for me on planet-source-code.com
' e-mail: blagoj_bl@yahoo.com for comments,help or anything else.
' (c) XbXan 2006

' module for SysTray and Process
'//////////////////////////////////////////////////////////////////////////////'

Option Explicit
'hot-keys
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'windows functions
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hwnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
'kill process api, already called
'Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Boolean
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Boolean
'Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Boolean
Public Const PROCESS_QUERY_INFORMATION As Long = &H400
Public Const PROCESS_TERMINATE As Long = &H1

Public Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type
Public Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Const GWL_HINSTANCE = (-6)

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'for top most window
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

'process handling
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Long, lpBuffer As Long, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_ALL_ACCESS = STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF


'for systray
Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
'
'Public Declare Function Shell_NotifyIcon Lib "Shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
'Public Const NIM_ADD = &H0
'Public Const NIM_DELETE = &H2
'Public Const NIF_TIP = &H4
'Public Const NIF_MESSAGE = &H1
'Public Const NIF_ICON = &H2
'Public Const WM_MOUSEMOVE = &H200
'Public Const WM_RBUTTONDOWN = &H204
'Public Const WM_LBUTTONDBLCLK = &H203

Public k As Long
Public pname As String  'used for all windows
Public PID As Long  'pid of the specified process
Public procpids(1 To 1000) As Long 'all process instances
Public thwnd(1 To 5000) As Long 'custom array of windows
Private lboxwnd As ListView 'window list is put here
Public iehidden As Long
' ToolHelp 32-bit VB implementation
'--------------------------------------------
'Note: These are not all the functions from ToolHelp.

Private Const MAX_PATH = 260
Private Const TH32CS_SNAPPROCESS = &H2

'Describe a process
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

'Functions
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal dwProcessId As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hObject As Long, p As PROCESSENTRY32) As Boolean
Private Declare Function Process32Next Lib "kernel32" (ByVal hObject As Long, p As PROCESSENTRY32) As Boolean
'Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' This function is not from ToolHelp but you need it to destroy a snapshot
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Function GetProcessesPids(ByVal procname As String, pids() As Long) As Long
Dim hp As Boolean
Dim hsnapshot As Long
Dim pinfo As PROCESSENTRY32
Dim i As Integer
i = 1

'Take the snapshot
hsnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
'First we must setup the dwSize parameter to len(pinfo)
'in order Process32First to work
pinfo.dwSize = Len(pinfo)
hp = Process32First(hsnapshot, pinfo)

While hp
'check the process name
If Mid(LCase(pinfo.szExeFile), 1, Len(procname)) = procname Then
    pids(i) = pinfo.th32ProcessID
    i = i + 1
End If

'Next process
hp = Process32Next(hsnapshot, pinfo)

Wend

'mark the end
pids(i) = -1
'destroy the snapshot
CloseHandle (hsnapshot)
End Function
Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long

Dim procid As Long


If pname <> "ALL" Then
    'find the process who owns the window
    GetWindowThreadProcessId hwnd, procid  'get the process PID
    'hide/show window
    If procid = PID Then ShowWindow hwnd, iehidden
Else
    'hide/show all windows in the system
    ShowWindow hwnd, iehidden
End If
'1 to proceed with enumeration, 0 to stop
EnumWindowsProc = 1
End Function
Public Function GetProcessList(ByVal lbox As ListView) As Boolean

Dim hp As Boolean
Dim hsnapshot As Long
Dim pinfo As PROCESSENTRY32

'Take the snapshot
hsnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
'First we must setup the dwSize parameter to len(pinfo)
'in order Process32First to work
pinfo.dwSize = Len(pinfo)
hp = Process32First(hsnapshot, pinfo)


While hp
    Dim X As ListItem
    Set X = lbox.ListItems.Add(, , pinfo.szExeFile)
    X.SubItems(1) = GetProcessFullPath(pinfo.th32ProcessID)
    X.SubItems(2) = pinfo.th32ProcessID
    Set X = Nothing
    'Next process
    hp = Process32Next(hsnapshot, pinfo)
Wend

'destroy the snapshot
CloseHandle (hsnapshot)

'remove the first item ([System Process])
lbox.ListItems.Remove 1

End Function

Public Function GetProcName(ByVal tpid As Long) As String
Dim hp As Boolean
Dim hsnapshot As Long
Dim pinfo As PROCESSENTRY32

'Take the snapshot
hsnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
'First we must setup the dwSize parameter to len(pinfo)
'in order Process32First to work
pinfo.dwSize = Len(pinfo)
hp = Process32First(hsnapshot, pinfo)


While hp
    If tpid = pinfo.th32ProcessID Then
        GetProcName = pinfo.szExeFile
        'destroy the snapshot
        CloseHandle (hsnapshot)
        Exit Function
    End If
    'Next process
    hp = Process32Next(hsnapshot, pinfo)
Wend


'destroy the snapshot
CloseHandle (hsnapshot)
End Function

Public Function GetWindowList(ByVal lbox As ListView, Optional ByVal labelinfo As Label) As Boolean
    Set lboxwnd = lbox
    EnumWindows AddressOf EnumWindowsProc2, 5
    'labelinfo.Caption = "Window list count " & lbox.ListItems.Count
End Function

Public Function EnumWindowsProc2(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim lclass As Long
Dim clsname As String * 100
Dim wndcaption As String * 256
Dim wndparentclass As String * 100
Dim wndparentcaption As String * 256
Dim lenwnd As Long
Dim tpid As Long
Dim procname As String
Dim procid As Long
Dim li As ListItem
Dim tmp As Long

If pname <> "ALL" Then
    'specified process windows
    GetWindowThreadProcessId hwnd, procid
    If PID = procid Then
        'get the window caption
        lenwnd = GetWindowText(hwnd, wndcaption, 256)
        wndcaption = Mid(wndcaption, 1, lenwnd)
        Set li = lboxwnd.ListItems.Add(, , Trim(wndcaption))
        'get the window class
        lclass = GetClassName(hwnd, clsname, 100)
        clsname = Mid(clsname, 1, lclass)
        'Get top most parent window
        tmp = GetTopLevelWindow(hwnd)
        'Get top most parent window caption
        lenwnd = GetWindowTextLength(tmp)
        GetWindowText tmp, wndparentcaption, 256
        wndparentcaption = Mid(wndparentcaption, 1, lenwnd)
        'Get top most parent window class
        lclass = GetClassName(tmp, wndparentclass, 100)
        wndparentclass = Mid(wndparentclass, 1, lclass)
        'get the process who owns the window
        GetWindowThreadProcessId tmp, tpid
        procname = GetProcName(tpid)

        li.ListSubItems.Add , , Trim(clsname)
        li.ListSubItems.Add , , Trim(wndparentcaption)
        li.ListSubItems.Add , , Trim(wndparentclass)
        li.ListSubItems.Add , , Trim(procname)
        li.ListSubItems.Add , , hwnd
    End If
Else
        'all windows in the system
         'get the window caption
        lenwnd = GetWindowText(hwnd, wndcaption, 256)
        wndcaption = Mid(wndcaption, 1, lenwnd)
        Set li = lboxwnd.ListItems.Add(, , Trim(wndcaption))
        'get the window class
        lclass = GetClassName(hwnd, clsname, 100)
        clsname = Mid(clsname, 1, lclass)
        'Get top most parent window
        tmp = GetTopLevelWindow(hwnd)
        'Get top most parent window caption
        lenwnd = GetWindowTextLength(tmp)
        GetWindowText tmp, wndparentcaption, 256
        wndparentcaption = Mid(wndparentcaption, 1, lenwnd)
        'Get top most parent window class
        lclass = GetClassName(tmp, wndparentclass, 100)
        wndparentclass = Mid(wndparentclass, 1, lclass)
        'get the process who owns the window
        GetWindowThreadProcessId tmp, tpid
        procname = GetProcName(tpid)
        li.ListSubItems.Add , , Trim(clsname)
        li.ListSubItems.Add , , Trim(wndparentcaption)
        li.ListSubItems.Add , , Trim(wndparentclass)
        li.ListSubItems.Add , , Trim(procname)
        li.ListSubItems.Add , , hwnd
End If

'1 to proceed with enumeration, 0 to stop
Set li = Nothing
EnumWindowsProc2 = 1
End Function

Public Function SetTopMostWindow(ByVal thwnd As Long, ByVal b As Boolean) As Boolean
If b Then
    If SetWindowPos(thwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE) <> 0 Then SetTopMostWindow = True Else SetTopMostWindow = False
Else
    If SetWindowPos(thwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE) <> 0 Then SetTopMostWindow = True Else SetTopMostWindow = False
End If

End Function

Public Function GetTopLevelWindow(ByVal ghwnd As Long) As Long
    Dim tophwnd As Long, tmp As Long
    tophwnd = ghwnd
    
    'loop for finding the toplevel window
    While tophwnd <> 0
        tmp = tophwnd
        tophwnd = GetParent(tophwnd)
        If tophwnd = 0 Then GetTopLevelWindow = tmp: Exit Function
    Wend

End Function

Public Function GetProcessFullPath(ByVal tpid As Long) As String

'Thanks to Bor0 for showing me how to get the full path
'of another process

    'in every process at &H2003C there is a pointer
    'to string that contains full path with name of that process
    Dim hProc As Long
    Dim pathproc() As Byte
    Dim pointstr As Long
    Dim strPath As String
    Dim oldprotect As Long
    Dim tstr As String
    
    pathproc = StrConv(String$(128, 0), vbFromUnicode)
    
    hProc = OpenProcess(PROCESS_ALL_ACCESS, 0, tpid)
    If hProc = 0 Then
        GetProcessFullPath = "System Process"   'added by myself to avoid blank path
        Exit Function
    End If
    
    'get the pointer to the string that contains the filepath
    ReadProcessMemory hProc, ByVal &H2003C, pointstr, 4, 0
    
    'read the string
    ReadProcessMemory hProc, ByVal pointstr, pathproc(0), 128, 0
    
    CloseHandle hProc
    
    strPath = StrConv(pathproc, vbUnicode)
    tstr = vbNullChar & vbNullChar & vbNullChar
    
    'get rid of the nulls
    GetProcessFullPath = ClearNulls(strPath, InStr(1, strPath, tstr))
    
End Function


Private Function ClearNulls(ByVal tstr As String, ByVal tpos As Integer) As String
Dim tmp As String, tmp1 As String, tmp2 As String
Dim i As Integer
tmp = vbNullChar

For i = 1 To tpos
    If tmp = Mid(tstr, i, 1) Then
        tmp1 = Mid(tstr, 1, i - 1)
        tmp2 = Mid(tstr, i + 1, tpos - (i + 1))
        tstr = tmp1 & tmp2
    End If
Next
ClearNulls = tstr
End Function
Public Sub UpdateTrayWindow()

Dim hwnd As Long
Dim hwnd2 As Long

hwnd = FindWindow("Shell_TrayWnd", "")
If hwnd <> 0 Then
    hwnd2 = FindWindowEx(hwnd, 0&, "TrayNotifyWnd", "")
    hwnd2 = FindWindowEx(hwnd2, 0&, "SysPager", "")
    hwnd2 = FindWindowEx(hwnd2, 0&, "ToolbarWindow32", "Notification Area")
    UpdateWindow (hwnd2)
End If
    
End Sub
Public Sub closewindow(ByVal ghwnd As Long)
'set that window to foreground window
ShowWindow ghwnd, 5
SetForegroundWindow ghwnd

'send alt+f4 to that window
SendKeys "%{F4}"
End Sub

'Kill process functions
Public Sub Process_Kill(P_ID As Long)
    '// Kill the wanted process
    On Error Resume Next
    Dim hProcess As Long
    Dim lExitCode As Long
    Dim res As Boolean
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, P_ID)
    res = GetExitCodeProcess(hProcess, lExitCode)
    res = TerminateProcess(hProcess, lExitCode)
    CloseHandle (hProcess)
End Sub

