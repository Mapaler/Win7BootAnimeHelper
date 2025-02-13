Attribute VB_Name = "OpenUnicode"
Public Function TextFile_Read(FileName As String) As String
Dim i As Integer, s As String, BB() As Byte
If Dir(FileName) = "" Then Exit Function
i = FreeFile
If FileLen(FileName) - 1 < 0 Then
    ReDim BB(0)
Else
    ReDim BB(FileLen(FileName) - 1)
End If
Open FileName For Binary As #i
Get #i, , BB
Close #i
s = StrConv(BB, vbUnicode)
TextFile_Read = s
End Function
 
Public Sub TextFile_Write(ByVal FileName As String, _
ByVal vVar As String)
'  On Error Resume Next
  Open FileName For Output As #1
  Print #1, vVar
  Close #1
End Sub



