VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private crcTable(0 To 255) As Long

Private Function CRC32(ByRef bArrayIn() As Byte, ByVal lLen As _
    Long, Optional ByVal lcrc As Long = 0) As Long

    Dim lCurPos As Long
    Dim lTemp As Long
    If lLen = 0 Then Exit Function
    lTemp = lcrc Xor &HFFFFFFFF
    For lCurPos = 0 To lLen
        If GetInputState() <> 0 Then DoEvents
        lTemp = (((lTemp And &HFFFFFF00) \ &H100) And &HFFFFFF) _
            Xor (crcTable((lTemp And 255) Xor bArrayIn(lCurPos)))
    Next lCurPos
    CRC32 = lTemp Xor &HFFFFFFFF

End Function

Public Sub BuildTable()
    
    Dim i As Long, X As Long, CRC As Long
    Const Limit = &HEDB88320
    For i = 0 To 255
        If GetInputState() <> 0 Then DoEvents
        CRC = i
        For X = 0 To 7
            If CRC And 1 Then
                CRC = (((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor Limit
            Else
                CRC = ((CRC And &HFFFFFFFE) \ 2) And &H7FFFFFFF
            End If
        Next X
        crcTable(i) = CRC
    Next i

End Sub

Public Function GetCRC(ByVal SPathFile As String) As String

    Dim b() As Byte, lcrc As Long
    On Error Resume Next
    If Len(SPathFile) = 0 Then Exit Function
    Open SPathFile For Binary Access Read As #1
    ReDim b(FileLen(SPathFile) - 1)
    Get #1, , b
    Close #1
    lcrc = UBound(b())
    lcrc = CRC32(b(), lcrc)
    GetCRC = Hex(lcrc)

End Function

Private Sub Class_initialize()

    Call BuildTable

End Sub



