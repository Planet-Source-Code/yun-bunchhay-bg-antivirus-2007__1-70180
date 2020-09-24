Attribute VB_Name = "modBinaryFile"
Option Explicit

Public Sub ReadSig()

    Dim f As Long
    On Error GoTo Trap_Error
    f = FreeFile
    Open App.Path & "\WDAV.sig" For Binary Access Read As #f
        Get #f, , VSInfo
        ReDim VSig(VSInfo.VirusCount - 1) As VirusSig
        Dim i As Integer
        For i = 0 To VSInfo.VirusCount - 1
            Get #f, , VSig(i)
        Next
    Close #f

   On Error GoTo 0
   Exit Sub

Trap_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetData of Form frmBinaAccess"
End Sub

Public Sub WriteSig(ByRef vs As VirusSig)
    
    Dim f As Long
    On Error GoTo Trap_Error
    f = FreeFile
    
    Dim i As Long
    
    'add 1 item into array
    ReDim Preserve VSig(UBound(VSig) + 1) As VirusSig
    VSig(UBound(VSig)).Name = vs.Name
    VSig(UBound(VSig)).Type = vs.Type
    VSig(UBound(VSig)).Value = vs.Value
    
    'add 1 for count
    VSInfo.VirusCount = UBound(VSig) + 1
    VSInfo.LastUpdate = Format(Date, "dd/mmmm/yyyy")
    
    'change virus last update
    'VSInfo.LastUpdate = Format("07 June 2007", "Short Date")
    
    Open App.Path & "\WDAV.sig" For Binary Access Write As #f
        Put #f, , VSInfo
        For i = 0 To UBound(VSig)
            Put #f, , VSig(i)
        Next
    Close #f

   On Error GoTo 0
   Exit Sub

Trap_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PutData of Form frmBinaAccess"
End Sub

