Attribute VB_Name = "������ʽ����"
Option Explicit
'�洢������ʽ������
Public Type PatternValue
    FirstIndex As Long
    AllValue As String
    InValue() As String
End Type
Public Function SearchText(ByVal text As String, ByVal patrn As String, ByRef Save() As PatternValue, Optional ByVal replStr As String = "$1") As String
'������ʽ����
Dim regEx As RegExp
Dim Match As Match
Dim Matchs As MatchCollection
Dim replStrPart() As String

Dim i As Long, j As Long
Set regEx = New RegExp '����������ʽ
regEx.Pattern = patrn '���ñ��ʽ
regEx.IgnoreCase = True 'true���жϴ�Сд
regEx.Global = True 'falseֻ������һ��,true����ȫ��
Set Matchs = regEx.Execute(text)

replStrPart = Split(replStr, Chr(0)) 'ʹ��Chr(0)�ֿ���Ҫ�����Ĳ��ֵĴ���
i = 0
For Each Match In Matchs '������������ʵ������
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

SearchText = Matchs.count '��ʾ���м�����
End Function

Public Function ReplaceText(ByVal text As String, ByVal patrn As String, ByVal replStr As String) As String
'������ʽ�滻
Dim regEx '��������
Set regEx = New RegExp '����������ʽ
regEx.Pattern = patrn '���ñ��ʽ
regEx.IgnoreCase = True 'true���жϴ�Сд
regEx.Global = True 'falseֻ������һ��,true����ȫ��
ReplaceText = regEx.Replace(text, replStr) '�滻����
End Function

Public Function IsOK(ByVal text As String, ByVal patrn As String) As Boolean
'������ʽ�ж��Ƿ����
Dim regEx
IsOK = False
Set regEx = New RegExp
regEx.IgnoreCase = True 'true���жϴ�Сд
regEx.Pattern = patrn
IsOK = regEx.Test(text)
End Function
'ֻ������Ҫ�ĸ�ʽ��������ȫ����
Public Function OnlyRegExp(myString As String, myPattern As String, Optional ByVal default As String = "", Optional ByVal replStr As String = "") As String
    'Create objects.
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim RetStr As String
    
    '����������ʽ
    Set objRegExp = New RegExp
    
    '���ñ��ʽ
    objRegExp.Pattern = myPattern
    
    'true���жϴ�Сд
    objRegExp.IgnoreCase = True
    
    'falseֻ������һ��,true����ȫ��
    objRegExp.Global = True
    
    '�ȼ���Ƿ��и��ϵĵط�
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

'����������ת��Ϊһ���ַ
Public Function Path_to_RealUrl(myString As String) As String
    'Create objects.
    Dim objRegExp As RegExp
    Dim objMatch As Match
    Dim colMatches   As MatchCollection
    Dim RetStr As String
    
    '����������ʽ
    Set objRegExp = New RegExp
    
    '���ñ��ʽ
    objRegExp.Pattern = "%(\w+)%"
    
    'true���жϴ�Сд
    objRegExp.IgnoreCase = True
    
    'falseֻ������һ��,true����ȫ��
    objRegExp.Global = True
    
    '������Ϊԭ��ַ
    RetStr = myString
    
    '�ȼ���Ƿ��л��������ĵط�
    If (objRegExp.Test(myString) = True) Then
        
        'Get the matches.
        Set colMatches = objRegExp.Execute(myString)   '' Execute search.
        
        For Each objMatch In colMatches   '' Iterate Matches collection.
            Dim Real_Url_Temp As String
            Dim Path_text As String
            '������ʽ�滻
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
            
            '�ҷ���"\"
            If Right$(Real_Url_Temp, 1) <> "\" Then Real_Url_Temp = Real_Url_Temp & "\"
            '��ԭ����Ļ��������滻Ϊ��ʵ��ַ
            RetStr = Replace(RetStr, objMatch.value, Real_Url_Temp)
        Next
    End If
    
    If Not RetStr = "" Then
        If IsOK(RetStr, FileURL_Parten) = False Then
            RetStr = App.Path & "\" & RetStr
        End If
    End If
    '����URL�е�/\
    RetStr = ReplaceText(RetStr, "/", "\") '��/�滻��\
    RetStr = ReplaceText(RetStr, "([^\\/])\\{2,}", "$1\") '�����/���滻��һ��/
    RetStr = ReplaceText(RetStr, "^\\{3,}", "\\") '�����\\��ͷ�ľ�������ַ�滻������\
    
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
    
    '������Ϊԭ��ַ
    RetStr = myString
    
    '�ȼ���Ƿ��л��������ĵط�
    If (objRegExp.Test(myString) = True) Then
        
        'Get the matches.
        Set colMatches = objRegExp.Execute(myString)   '' Execute search.
        
        For Each objMatch In colMatches   '' Iterate Matches collection.
            Dim Real_Text_Temp As String
            Dim Path_text As String
            Real_Text_Temp = ""
            
            '������ʽ�滻
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
            Else 'ϵͳ��������
                If Environ(Path_text) = "" Then
                    Real_Text_Temp = objRegExp.Replace(objMatch.value, "$1")
                Else
                    Real_Text_Temp = Environ(Path_text)
                End If
            End If
            '��ԭ����Ļ��������滻Ϊ��ʵ��ַ
            RetStr = Replace(RetStr, objMatch.value, Real_Text_Temp)
        Next
    End If
    
    Path_to_String = RetStr
End Function
