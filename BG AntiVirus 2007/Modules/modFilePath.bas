Attribute VB_Name = "modFilePath"
Option Explicit

'Check if a path or file exists
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

'Checks if a folder or file exists
Public Function FileorFolderExists(FolderOrFilename As String) As Boolean
    If PathFileExists(FolderOrFilename) = 1 Then
        FileorFolderExists = True
    ElseIf PathFileExists(FolderOrFilename) = 0 Then
        FileorFolderExists = False
    End If
End Function

'Format the values by removing everything except for the filename or path
Public Function FormatValue(sValue As String) As String

    Dim FileorPath As String
    'Fix up the file or path so that it's compatible with the FileorFolderExists function
    FileorPath = sValue

    'Find the start of the path or filename (Example:"h6j65ej(C:\Test)")
    If InStr(1, FileorPath, "C:\") Then FileorPath = Mid(FileorPath, InStr(1, FileorPath, "C:\"))
    If InStr(1, FileorPath, "c:\") Then FileorPath = Mid(FileorPath, InStr(1, FileorPath, "c:\"))

    'Remove everything after the path. This definitely doesn't work for all values.
    '(Example:"C:\blablablablablablablabla?5784846\84585")
    If InStr(1, FileorPath, "/") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "/") - 1)
    If InStr(1, FileorPath, "*") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "*") - 1)
    If InStr(1, FileorPath, "?") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "?") - 1)
    If InStr(1, FileorPath, Chr(34)) > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, Chr(34)) - 1)
    If InStr(1, FileorPath, "<") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "<") - 1)
    If InStr(1, FileorPath, ">") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ">") - 1)
    If InStr(1, FileorPath, "|") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "|") - 1)
    If InStr(1, FileorPath, ",") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ",") - 1)
    If InStr(1, FileorPath, "(") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "(") - 1)
    If InStr(1, FileorPath, ";") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ";") - 1)

    'some registry values somehow didn't contain "C:\"
    If InStr(1, FileorPath, "C:\") = 0 Then FileorPath = "C:\"

    'Remove everything before the path or file. The same as the other one except this is for specific extensions
    '(Example:"C:\lalalalalalalala\idfjb.dll\50")
    'I guess I could just do If InStr(1, FileorPath, ".exe") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".exe") + 2), If InStr(1, FileorPath, ".dll") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".dll") + 2), etc, but I'm not sure how that would work out

    If InStr(1, FileorPath, ".exe:") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".exe:") + 3)
    If InStr(1, FileorPath, ".EXE:") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".EXE:") + 3)
    If InStr(1, FileorPath, ".EXE ") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".EXE ") + 3)
    If InStr(1, FileorPath, ".exe ") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".exe ") + 3)
    If InStr(1, FileorPath, ".SYS ") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".SYS ") + 3)
    If InStr(1, FileorPath, ".sys ") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".sys ") + 3)
    If InStr(1, FileorPath, ".EXE\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".EXE\") + 3)
    If InStr(1, FileorPath, ".exe\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".exe\") + 3)
    If InStr(1, FileorPath, ".DLL\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".DLL\") + 3)
    If InStr(1, FileorPath, ".dll\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".dll\") + 3)
    If InStr(1, FileorPath, ".OCX\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".OCX\") + 3)
    If InStr(1, FileorPath, ".ocx\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".ocx\") + 3)
    If InStr(1, FileorPath, "*") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "*") - 1)
    If InStr(3, FileorPath, ":") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ":") - 1)

    '%1 is used for file associations
    '(Example:"C:\WINDOWS\NOTEPAD.EXE %1")
    FormatValue = Replace(FileorPath, " %1", "")
End Function

