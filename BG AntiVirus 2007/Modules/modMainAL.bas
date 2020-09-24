Attribute VB_Name = "modMainAL"
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

Public Function ReadRegShell(ByVal sKey As String)
    On Error GoTo trapErr
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    ReadRegShell = sh.regread(sKey)
    Exit Function
trapErr:
    ReadRegShell = 0
End Function

Public Sub WriteRegShell(ByVal sKey As String, ByVal sValue As String)
    On Error Resume Next
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    sh.regwrite sKey, sValue
End Sub
