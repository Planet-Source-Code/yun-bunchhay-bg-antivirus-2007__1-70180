Attribute VB_Name = "modRegStartup"
'Registry Key/Value Enumeration Functions

'By Max Raskin 29 August 2000

Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor As Long, lpftLastWriteTime As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Const KEY_QUERY_VALUE = &H1
Private Const MAX_PATH = 260

Enum RegDataTypes
    REG_SZ = 1                         ' Unicode nul terminated string
    REG_EXPAND_SZ = 2                  ' Unicode nul terminated string
    REG_DWORD = 4                      ' 32-bit number
End Enum

Enum RegistryKeys
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum

Enum ValKey
    Values = 0
    Keys = 1
End Enum

Private Type ByteArray
  FirstByte As Byte
  ByteBuffer(255) As Byte
End Type

Dim baData As ByteArray

Public Function OpenKey(RegistryKey As RegistryKeys, Optional SubKey As String) As Long
    If OpenKey <> 0 Then RegCloseKey (OpenKey)
    RegOpenKeyEx RegistryKey, SubKey, 0, KEY_QUERY_VALUE, OpenKey
End Function

Public Function GetCount(RegisteryKeyHandle As Long, ValuesOrKeys As ValKey) As Long
    If ValuesOrKeys = Keys Then RegQueryInfoKey RegisteryKeyHandle, "", 0, 0, GetCount, 0, 0, 0, 0, 0, 0, 0
    If ValuesOrKeys = Values Then RegQueryInfoKey RegisteryKeyHandle, "", 0, 0, 0, 0, 0, GetCount, 0, MAX_PATH + 1, 0, 0
End Function

Public Function EnumKey(RegisteryKeyHandle As Long, KeyIndex As Long) As String
    EnumKey = Space(MAX_PATH + 1)
    RegEnumKey RegisteryKeyHandle, KeyIndex, EnumKey, MAX_PATH + 1
    EnumKey = Trim(EnumKey)
End Function

Public Function EnumValue(RegisteryKeyHandle As Long, KeyIndex As Long) As String
    Dim lBufferLen As Long, i As Integer
    For i = 0 To 255
      baData.ByteBuffer(i) = 0
    Next
    lBufferLen = 255
    EnumValue = Space(MAX_PATH + 1)
    RegQueryInfoKey RegisteryKeyHandle, "", 0, 0, 0, 0, 0, 0, lValNameLen, lValLen, 0, 0
    RegEnumValue RegisteryKeyHandle, KeyIndex, EnumValue, MAX_PATH + 1, 0, 0, baData.FirstByte, lBufferLen
    EnumValue = Trim(EnumValue)
End Function

Public Function DeleteValue(RegisteryKeyHandle As Long, KeyName As String) As Long
    DeleteValue = RegDeleteValue(RegisteryKeyHandle, KeyName)
End Function

Public Function SetValue(RegisteryKeyHandle As RegistryKeys, SubRegistryKey As String, KeyName As String, NewValue As String, Optional DataType As RegDataTypes)
    Dim lRetVal As Long
    lRetVal = OpenKey(RegisteryKeyHandle, SubRegistryKey)
    If DataType = 0 Then DataType = REG_SZ
    RegSetValueEx lRetVal, KeyName, 0, DataType, NewValue, LenB(StrConv(SubKeyValue, vbFromUnicode))
End Function

Public Function GetKeyValue(hKey As Long, KeyName As String) As String
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, KeyName, 0, _
                         lKeyValType, tmpVal, KeyValSize)
    GetKeyValue = Trim(tmpVal)
End Function

