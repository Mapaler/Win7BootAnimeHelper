Attribute VB_Name = "正则表达式函数"
Option Explicit
'存储正则表达式搜索用
Public Type PatternValue
    FirstIndex As Long
    AllValue As String
    InValue() As String
End Type
Public Function SearchText(ByVal text As String, ByVal patrn As String, ByRef Save() As PatternValue, Optional ByVal replStr As String = "$1") As String
'正则表达式搜索
Dim regEx As RegExp
Dim Match As Match
Dim Matchs As MatchCollection
Dim replStrPart() As String

Dim i As Long, j As Long
Set regEx = New RegExp '建立正则表达式
regEx.Pattern = patrn '设置表达式
regEx.IgnoreCase = True 'true则不判断大小写
regEx.Global = True 'false只搜索第一个,true就是全部
Set Matchs = regEx.Execute(text)

replStrPart = Split(replStr, Chr(0)) '使用Chr(0)分开需要搜索的部分的代码
i = 0
For Each Match In Matchs '遍历，并存入实参数组
    ReDim Preserve Save(i)
    With Save(i)
        .FirstIndex = Match.FirstIndex
        .AllValue = Match.value
        ReDim Preserve .InValue(UBound(replStrPart))
        For j = 0 To UBound(replStrPart)
        .InValue(j) = regEx.Replace(Match.value, replStrPart(j))
        Next
    End With
    i = i + 1
Next

SearchText = Matchs.count '显示共有几个？
End Function

Public Function ReplaceText(ByVal text As String, ByVal patrn As String, ByVal replStr As String) As String
'正则表达式替换
Dim regEx '建立变量
Set regEx = New RegExp '建立正则表达式
regEx.Pattern = patrn '设置表达式
regEx.IgnoreCase = True 'true则不判断大小写
regEx.Global = True 'false只搜索第一个,true就是全部
ReplaceText = regEx.Replace(text, replStr) '替换命令
End Function

Public Function IsOK(ByVal text As String, ByVal patrn As String) As Boolean
'正则表达式判断是否符合
Dim regEx
IsOK = False
Set regEx = New RegExp
regEx.IgnoreCase = True 'true则不判断大小写
regEx.Pattern = patrn
IsOK = regEx.Test(text)
End Function
'只保留需要的格式，其他的全丢弃
Public Function OnlyRegExp(myString As String, myPattern As String, Optional ByVal default As String = "", Optional ByVal replStr As String = "") As String
    'Create objects.
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim RetStr As String
    
    '建立正则表达式
    Set objRegExp = New RegExp
    
    '设置表达式
    objRegExp.Pattern = myPattern
    
    'true则不判断大小写
    objRegExp.IgnoreCase = True
    
    'false只搜索第一个,true就是全部
    objRegExp.Global = True
    
    '先检查是否有复合的地方
    If (objRegExp.Test(myString) = True) Then

   ''Get the matches.
        Set colMatches = objRegExp.Execute(myString)   ' Execute search.
        
        For Each objMatch In colMatches   ' Iterate Matches collection.
            If replStr = "" Then
                RetStr = RetStr & objMatch.value
            Else
                RetStr = objRegExp.Replace(objMatch.value, replStr)
            End If
        Next
    Else
        RetStr = default
    End If
    OnlyRegExp = RetStr
End Function

'将环境变量转换为一般地址
Public Function Path_to_RealUrl(myString As String) As String
    'Create objects.
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim RetStr As String
    
    '建立正则表达式
    Set objRegExp = New RegExp
    
    '设置表达式
    objRegExp.Pattern = "%(\w+)%"
    
    'true则不判断大小写
    objRegExp.IgnoreCase = True
    
    'false只搜索第一个,true就是全部
    objRegExp.Global = True
    
    '先设置为原地址
    RetStr = myString
    
    '先检查是否有环境函数的地方
    If (objRegExp.Test(myString) = True) Then
        
        'Get the matches.
        Set colMatches = objRegExp.Execute(myString)   '' Execute search.
        
        For Each objMatch In colMatches   '' Iterate Matches collection.
            Dim Real_Url_Temp As String
            Dim Path_text As String
            '正则表达式替换
            Path_text = objRegExp.Replace(objMatch.value, "$1")
            
'            If IsOK(Path_text, "^HomePath$") Then
'                Real_Url_Temp = Environ("SYSTEMDRIVE") & Environ("HOMEPATH")
'            Else
            If IsOK(Path_text, "^AppTitle$") Then
                Real_Url_Temp = App.Title
            ElseIf IsOK(Path_text, "^AppPath$") Then
                Real_Url_Temp = App.Path
            ElseIf IsOK(Path_text, "^AppEXEName$") Then
                Real_Url_Temp = App.EXEName
            ElseIf IsOK(Path_text, "^ResourceDir$") Then
                Real_Url_Temp = Environ("Windir") & "\Resources"
            Else
                Real_Url_Temp = Environ(Path_text)
            End If
            
            '右方补"\"
            If Right$(Real_Url_Temp, 1) <> "\" Then Real_Url_Temp = Real_Url_Temp & "\"
            '将原文里的环境变量替换为真实地址
            RetStr = Replace(RetStr, objMatch.value, Real_Url_Temp)
        Next
    End If
    
    If Not RetStr = "" Then
        If IsOK(RetStr, FileURL_Parten) = False Then
            RetStr = App.Path & "\" & RetStr
        End If
    End If
    '处理URL中的/\
    RetStr = ReplaceText(RetStr, "/", "\") '将/替换成\
    RetStr = ReplaceText(RetStr, "([^\\/])\\{2,}", "$1\") '将多个/或替换成一个/
    RetStr = ReplaceText(RetStr, "^\\{3,}", "\\") '将多个\\开头的局域网地址替换成两个\
    
    Path_to_RealUrl = RetStr
End Function

Public Function Path_to_String(myString As String) As String
    'Create objects.
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim RetStr As String
    
    Set objRegExp = New RegExp
    objRegExp.Pattern = "%s\(([\w\.]+?)\)"
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    
    '先设置为原地址
    RetStr = myString
    
    '先检查是否有环境函数的地方
    If (objRegExp.Test(myString) = True) Then
        
        'Get the matches.
        Set colMatches = objRegExp.Execute(myString)   '' Execute search.
        
        For Each objMatch In colMatches   '' Iterate Matches collection.
            Dim Real_Text_Temp As String
            Dim Path_text As String
            Real_Text_Temp = ""
            
            '正则表达式替换
            Path_text = objRegExp.Replace(objMatch.value, "$1")
            If IsOK(Path_text, "^AppTitle$") Then
                Real_Text_Temp = App.Title
            ElseIf IsOK(Path_text, "^AppPath$") Then
                Real_Text_Temp = App.Path
            ElseIf IsOK(Path_text, "^AppEXEName$") Then
                Real_Text_Temp = App.EXEName
            ElseIf IsOK(Path_text, "^AppMajor$") Then
                Real_Text_Temp = App.Major
            ElseIf IsOK(Path_text, "^AppMinor$") Then
                Real_Text_Temp = App.Minor
            ElseIf IsOK(Path_text, "^AppRevision$") Then
                Real_Text_Temp = App.Revision
            ElseIf IsOK(Path_text, "^AppBeta$") Then
                Real_Text_Temp = App_Beta
            ElseIf IsOK(Path_text, "^ResourceDir$") Then
                Real_Text_Temp = Environ("Windir") & "\Resources"
            Else '系统环境变量
                If Environ(Path_text) = "" Then
                    Real_Text_Temp = objRegExp.Replace(objMatch.value, "$1")
                Else
                    Real_Text_Temp = Environ(Path_text)
                End If
            End If
            '将原文里的环境变量替换为真实地址
            RetStr = Replace(RetStr, objMatch.value, Real_Text_Temp)
        Next
    End If
    
    Path_to_String = RetStr
End Function
