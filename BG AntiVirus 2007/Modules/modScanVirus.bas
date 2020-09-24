Attribute VB_Name = "modScanVirus"
Option Explicit

Public blnScan As Boolean   'check scan or stop
Public strScanDetail As String  'Scan Detail HTML text

Sub ScanFile(ByVal sPath As String)

    On Error Resume Next
    
    Dim fso As New FileSystemObject
    Dim X As ListItem
    'main folder
    Dim mFolder As Folder
    'files and folders collections
    Dim sFolders As Folders
    Dim sFiles As Files
    'for loop variables
    Dim sFolder As Folder
    Dim sFile As File
    
    'get main folder
    Set mFolder = fso.GetFolder(sPath)
    'get subfolders in main folder
    Set sFolders = mFolder.SubFolders
    'get files in main folder
    Set sFiles = mFolder.Files
    
    'scan files
    For Each sFile In sFiles
        
        'limit file size
        If sFile.Size > FileSize Then GoTo endFor 'can't use Continue For
        
        'If GetInputState() <> 0 Then DoEvents  'faster but look stuck
        DoEvents
        
        'check if it is stopped
        If blnScan = False Then Exit For
        'show scanned file
        frmMain.lblPath.Caption = sFile
        'add 1 to counter scanned
        frmMain.lblCount.Caption = Int(frmMain.lblCount.Caption) + 1
        'scan virus
        Dim sCRC As String
        sCRC = CRC.GetCRC(sFile)
        Dim i As Long
        
        'scan algorithm 1
        '----------------
        
'        If InStr(1, strVirusDef, sCRC, vbBinaryCompare) > 0 Then
'            'virus found
'            For i = 0 To UBound(VirusDef)
'                If GetInputState() <> 0 Then DoEvents
'                If sCRC = VirusDef(i)(2) Then   'start cleaning
'                    'add to log
'                    strScanDetail = strScanDetail & "Virus Found: <Font Size=3 Color=RED>" & VirusDef(i)(0) & "</font><br>"
'                    strScanDetail = strScanDetail & " File: <Font Size=3 Color=ORANGE><i>" & sFile & "</i></font><br>"
'                    Call UpdateDetail(strScanDetail, frmMain.WebBrowser1)
'
'                    'add 1 to counter
'                    frmMain.lblFound.Caption = Int(frmMain.lblFound.Caption) + 1
'
'                    On Error GoTo errKill
'                    'get filename, after kill, sFile is null?
'                    Dim tempFN As String
'                    tempFN = sFile.Path ' & "\" & sFile.Name
'
'                    'remove file with force
'                    'Kill sFile
'                    sFile.Delete True
'                    'add to log after kill process
'                    'frmMain.txtLog.Text = frmMain.txtLog.Text & "  Removed " & tempFN & vbCrLf
'                    frmMain.lblCleaned.Caption = Int(frmMain.lblCleaned.Caption) + 1
'                    strScanDetail = strScanDetail & " Virus Cleaned<br>"
'                    Call UpdateDetail(strScanDetail, frmMain.WebBrowser1)
'                    GoTo endFor
'errKill:
'                    'add to log after kill error
'                    'frmMain.txtLog.Text = frmMain.txtLog.Text & "  Cannot removed " & tempFN & vbCrLf
'                    strScanDetail = strScanDetail & "<font Size=3 Color=YELLOW><i> Virus Cannot Be Cleaned</i></font><br>"
'                    Call UpdateDetail(strScanDetail, frmMain.WebBrowser1)
'                    Exit For
'                End If
'            Next i
'        End If
'endFor:
'    Next


        'scan algorithm 2
        '----------------
        
        'compare with database
        For i = 0 To UBound(VSig)
            If GetInputState() <> 0 Then DoEvents
            If sCRC = VSig(i).Value Then   'start cleaning
                'add to log
                ' strScanDetail = strScanDetail & "Virus Found: <Font Size=3 Color=RED>" & VSig(i).Name & "</font><br>"
                ' strScanDetail = strScanDetail & " File: <Font Size=3 Color=ORANGE><i>" & sFile & "</i></font><br>"

                ' Call UpdateDetail(strScanDetail, frmMain.WebBrowser1)

                'add 1 to counter
                frmMain.lblFound.Caption = Int(frmMain.lblFound.Caption) + 1

                'On Error GoTo errKill
                'get filename, after kill, sFile is null?
                Dim tempFN As String
                tempFN = sFile.Path ' & "\" & sFile.Name

                'remove file with force
                sFile.Delete True
                'check whether the virus is cleaned or not; if not, go to errKill to show Error Cleaning
                If FileorFolderExists(tempFN) = True Then GoTo errKill
                'Kill sFile.Path
                'add to log after kill process
                'frmMain.txtLog.Text = frmMain.txtLog.Text & "  Removed " & tempFN & vbCrLf
                frmMain.lblCleaned.Caption = Int(frmMain.lblCleaned.Caption) + 1
                ' strScanDetail = strScanDetail & " Virus Cleaned<br>"
                ' Call UpdateDetail(strScanDetail, frmMain.WebBrowser1)
                Set X = frmMain.lvVirusFound.ListItems.Add(, , VSig(i).Name, 2, 2)
                X.SubItems(1) = tempFN
                X.SubItems(2) = "Cleaned"
                Set X = Nothing
                GoTo endFor
errKill:
                'add to log after kill error
                'frmMain.txtLog.Text = frmMain.txtLog.Text & "  Cannot removed " & tempFN & vbCrLf
                ' strScanDetail = strScanDetail & "<font Size=3 Color=YELLOW><i> Virus Cannot Be Cleaned</i></font><br>"
                ' Call UpdateDetail(strScanDetail, frmMain.WebBrowser1)
                Set X = frmMain.lvVirusFound.ListItems.Add(, , VSig(i).Name, 3, 3)
                X.SubItems(1) = tempFN
                X.SubItems(2) = "Clean Failed"
                Set X = Nothing
                Exit For
            End If
        Next i
endFor:
    Next
    
    'scan subfolders
    For Each sFolder In sFolders
        DoEvents
        If blnScan = False Then Exit For
        ScanFile (sFolder)
    Next
    
    'clear variables
    Set fso = Nothing
    Set mFolder = Nothing
    Set sFolders = Nothing
    Set sFiles = Nothing
    Set sFolder = Nothing
    Set sFile = Nothing
    
End Sub
