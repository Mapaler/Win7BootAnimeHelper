Attribute VB_Name = "处理字符串"
Option Explicit

'将 16进制数转换为10进制数
Public Function x16_to_x10(ByVal x16str As String) As Long
Dim x16num As String
Dim x10num As Long
    If IsOK(x16str, "^([a-fA-F\d]+)$") Then
        x16num = x16str
    ElseIf IsOK(x16str, "^(0?x|#|&H)([a-fA-F\d]+)$") Then
        x16num = ReplaceText(x16str, "^(0?x|#|&H)([a-fA-F\d]+)$", "$2")
    Else
        x16num = OnlyRegExp(x16str, "^[a-fA-F\d]$", 0)
    End If
    x10num = CLng("&H" & x16num)
    x16_to_x10 = x10num
End Function

'将 10进制数转换为16进制数
Public Function x10_to_x16(ByVal x10num As Long, Optional ByVal Length As Byte = 1) As String
Dim i%
Dim x16num As String
Dim FormatTemp As String
    FormatTemp = ""
    '补0个数
    For i = 1 To Length
        FormatTemp = FormatTemp & "@"
    Next
    
    x16num = Replace(format$(Hex(x10num), FormatTemp), " ", "0")
    x10_to_x16 = x16num
End Function

'将RGB转换成BGR(字符型)
Public Function RGB_To_BGRstr(ByVal Color As String, Optional ByVal hex16 As Boolean = False) As String
    Dim x16temp As String
    If hex16 Then '进行16进制数转换
        x16temp = Color
        x16temp = Mid$(x16temp, 5, 2) + Mid$(x16temp, 3, 2) + Mid$(x16temp, 1, 2)
        RGB_To_BGRstr = x16temp
    Else '进行10进制数转换
        x16temp = x10_to_x16(Color, 6)
        '交换RGB
        x16temp = Mid$(x16temp, 5, 2) + Mid$(x16temp, 3, 2) + Mid$(x16temp, 1, 2)
        RGB_To_BGRstr = x16_to_x10(x16temp)
    End If
    
End Function
'将RGB转换成BGR(数字型)
Public Function RGB_To_BGR(ByVal xGBRnum As Long) As Long
    Dim R As Integer, G As Integer, B As Integer
    Dim Output As Long
    R = (xGBRnum And &HFF) Mod &H100
    G = ((xGBRnum And &HFF00) \ &H100) Mod &H100
    B = ((xGBRnum And &HFF0000) \ &H10000) Mod &H100
    Output = RGB(B, G, R)
    RGB_To_BGR = Output
End Function

'将16位色变为VB的GBR24位色
Public Function x16b_to_xGBR(ByVal x16bnum As Long) As Long
    Dim R As Integer, G As Integer, B As Integer
    Dim Output As Long, Colortemp As Long
    Colortemp = x16bnum - 32768
    R = (x16bnum Mod &H20) * 8
    G = ((x16bnum \ &H20) Mod &H20) * 8
    B = ((x16bnum \ &H400) Mod &H20) * 8
    'MsgBox R & "," & G & "," & B
    Output = RGB(R, G, B)
    x16b_to_xGBR = Output
End Function

'将VB的GBR24位色变为16位色
Public Function xGBR_to_x16b(ByVal xGBRnum As Long) As String
    Dim R As Integer, G As Integer, B As Integer
    Dim Output As Long
    R = xGBRnum Mod &H100
    G = (xGBRnum \ &H100) Mod &H100
    B = (xGBRnum \ &H10000) Mod &H100
    'MsgBox R & "," & G & "," & B
    Output = 32768 + (R \ 8) + (G \ 8) * &H20 + (B \ 8) * &H400
    xGBR_to_x16b = Output
End Function

'数字不低于最小值
Public Function NotLess(ByVal Num As Long, Optional ByVal than As Long = 0) As Long
    If Num >= than Then
        NotLess = Num
    Else
        NotLess = than
    End If
End Function
'数字不大于最大值
Public Function NotGreater(ByVal Num As Long, ByVal than As Long) As Long
    If Num <= than Then
        NotGreater = Num
    Else
        NotGreater = than
    End If
End Function
'向上取整
Public Function Fixb(ByVal Num As Double) As Long
    If Fix(Num) < Num Then
        Fixb = Fix(Num) + 1
    Else
        Fixb = Fix(Num)
    End If
End Function
