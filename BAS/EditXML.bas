Attribute VB_Name = "编辑XML"
Option Explicit

' 返回各个节点的值

Public Function GetNodeText(ByVal start_at_node As DOMDocument, ByVal node_name As String, Optional ByVal default_value As String = "", Optional ByVal Index As Long = 0, Optional ByVal Max) As String
    Dim Value_Temp As String
    Dim value_node As IXMLDOMNodeList
    Set value_node = start_at_node.selectNodes(node_name)
    If value_node(Index) Is Nothing Then '如果没有就使用默认值
        Value_Temp = default_value
    Else
        Value_Temp = value_node(Index).text
        If Not IsMissing(Max) Then  '如果设置了最大数字则当做数字处理
            Value_Temp = CLng(OnlyRegExp(Value_Temp, "[-\d]", default_value))
            
            If Value_Temp > CLng(Max) Then '如果超过最大值则等于最大值，需要先将Max转变回数字才能识别
                Value_Temp = CLng(Max)
            End If

        End If
    End If
    GetNodeText = Value_Temp
End Function

' 返回节点属性
Public Function GetNodeAttribute(ByVal start_at_node As DOMDocument, ByVal node_name As String, ByVal AttributeName As String, Optional ByVal default_value As String = "", Optional ByVal Index As Long = 0, Optional ByVal Max) As String
    Dim Value_Temp As String
    Dim value_node As IXMLDOMNodeList
    Dim value_node_Attribute As IXMLDOMNode
    Set value_node = start_at_node.selectNodes(node_name)
    If value_node(Index) Is Nothing Then '如果节点不存在
        Value_Temp = default_value
    Else
        Set value_node_Attribute = value_node(Index).Attributes.getNamedItem(AttributeName)
        If value_node_Attribute Is Nothing Then '如果节点属性不存在
            Value_Temp = default_value
        Else
            Value_Temp = value_node_Attribute.text
            If Not IsMissing(Max) Then  '如果设置了最大数字则当做数字处理
                Value_Temp = CLng(OnlyRegExp(Value_Temp, "[-\d]", default_value))
                
                If Value_Temp > CLng(Max) Then '如果超过最大值则等于最大值，需要先将Max转变回数字才能识别
                    Value_Temp = CLng(Max)
                End If
    
            End If
        End If
    End If
    GetNodeAttribute = Value_Temp
End Function

'从XML读取同名结点个数
Public Function GetAllNode_Lenth(ByVal start_at_node As DOMDocument, ByVal node_name As String) As Long
    Dim i%
    Dim Value_Temp As String
    Dim value_node As IXMLDOMNodeList
    Set value_node = start_at_node.selectNodes(node_name)
    If value_node Is Nothing Then '如果节点不存在
        Value_Temp = 0
    Else
        Value_Temp = value_node.Length
    End If
    GetAllNode_Lenth = Value_Temp
End Function

Public Sub CreateNode(ByVal start_at_node As DOMDocument, ByVal father_node_name As String, ByVal node_name As String, ByVal node_value As String, Optional ByVal Index As Long = 0)

    Dim value_node As IXMLDOMNodeList
    Set value_node = start_at_node.selectNodes(father_node_name)
    
    Dim new_node As IXMLDOMNode
    Set new_node = value_node(Index)
    '如果节点不存在则创建
    If new_node Is Nothing Then
        Call CreateNode(start_at_node, Mid(father_node_name, 1, InStrRev(father_node_name, "/") - 1), Mid(father_node_name, InStrRev(father_node_name, "/") + 1), "")
		Set value_node = start_at_node.selectNodes(father_node_name)
    End If
    Set new_node = value_node(Index).ownerDocument.createElement(node_name)
    new_node.text = node_value
    value_node(Index).appendChild new_node
End Sub

Public Sub DelNode(ByVal start_at_node As DOMDocument, ByVal node_name As String, ByVal node_value As String, Optional ByVal Index As Long = 0)

    Dim value_node As IXMLDOMNode
    Set value_node = start_at_node.selectSingleNode(node_name)
    Dim new_node As IXMLDOMNodeList
    Set new_node = value_node.selectNodes(node_value)

    value_node.removeChild new_node(Index)
End Sub

Public Sub CreateAttribute(ByVal start_at_node As DOMDocument, ByVal node_name As String, ByVal Attribute_name As String, ByVal Attribute_value As String, Optional ByVal Index As Long = 0)

    Dim value_node As IXMLDOMNodeList
    Set value_node = start_at_node.selectNodes(node_name)

    
    Dim new_node As IXMLDOMNode
    Set new_node = value_node(Index).ownerDocument.CreateAttribute(Attribute_name)
    new_node.text = Attribute_value
    value_node(Index).Attributes.setNamedItem new_node
End Sub

Public Sub SaveAttribute(ByVal start_at_node As DOMDocument, ByVal node_name As String, ByVal Attribute_name As String, ByVal Attribute_value As String, Optional ByVal Index As Long = 0)
    Dim value_node As IXMLDOMNodeList
    Dim value_node_Attribute As IXMLDOMNode
    
    Set value_node = start_at_node.selectNodes(node_name)
    
    Dim new_node_f As IXMLDOMNode
    Set new_node_f = value_node(Index)
    If new_node_f Is Nothing Then
        Call CreateNode(start_at_node, Mid(node_name, 1, InStrRev(node_name, "/") - 1), Mid(node_name, InStrRev(node_name, "/") + 1), "")
        Set value_node = start_at_node.selectNodes(node_name)
        Set new_node_f = value_node(Index)
    End If

    Set value_node_Attribute = value_node(Index).Attributes.getNamedItem(Attribute_name)
    
    If value_node_Attribute Is Nothing Then '如果节点属性不存在
        Call CreateAttribute(start_at_node, node_name, Attribute_name, Attribute_value, Index)
    Else
        value_node_Attribute.text = Attribute_value
    End If
End Sub

'从XML读取obj的所有设置
Public Sub Set_All_Attribute_To_Obj(ByVal start_at_node As DOMDocument, ByVal node_name As String, ByRef obj)
    Dim i%
    Dim Value_Temp As String
    Dim value_node As IXMLDOMNode
    Set value_node = start_at_node.selectSingleNode(node_name)
    If value_node Is Nothing Then '如果节点不存在
        Exit Sub
    Else
        For i = 1 To value_node.Attributes.Length
            Dim ValueName As String, newvalue As String
            ValueName = value_node.Attributes(i - 1).baseName
            newvalue = value_node.Attributes(i - 1).text
            
            If IsOK(ValueName, "^Width$") Or IsOK(ValueName, "^Height$") Or IsOK(ValueName, "^Left$") Or IsOK(ValueName, "^Top$") _
            Or IsOK(ValueName, "^ScaleWidth$") Or IsOK(ValueName, "^ScaleHeight$") Or IsOK(ValueName, "^ScaleLeft$") Or IsOK(ValueName, "^ScaleTop$") Then
                newvalue = OnlyRegExp(newvalue, "\d")
                If VarType(obj) = 9 Then '窗体
                    newvalue = newvalue * 15
                End If
                On Error Resume Next '出错则执行下一句
                Call CallByName(obj, ValueName, VbLet, newvalue)
            ElseIf IsOK(ValueName, "Color$") Then
                newvalue = RGB_To_BGR(x16_to_x10(newvalue))
                On Error Resume Next '出错则执行下一句
                Call CallByName(obj, ValueName, VbLet, newvalue)
            ElseIf IsOK(ValueName, "^Picture$") Then
                On Error Resume Next '出错则执行下一句
                obj.Picture = LoadPicture(Path_to_RealUrl(newvalue))
            ElseIf IsOK(ValueName, "^PicURL$") Then
                On Error Resume Next '出错则执行下一句
                '如果是图片按钮则直接改变URL
                If IsOK(TypeName(obj), "^CommandButtonEx$") Then
                    Call CallByName(obj, ValueName, VbLet, newvalue)
                    Call CallByName(obj, "PicbX", VbLet, GetNodeAttribute(start_at_node, node_name, "PicX", 0, , 2147483647))
                    Call CallByName(obj, "PicbY", VbLet, GetNodeAttribute(start_at_node, node_name, "PicY", 0, , 2147483647))
                Else
                    newvalue = Path_to_RealUrl(newvalue)
                    Dim DrawStyle As PaintStyle
                    Dim mX As Long, mY As Long, mWidth As Long, mHeight As Long
                    Dim xAlign As My_xAlign, yAlign As My_yAlign
                    Dim mTop As Long, mRight As Long, mButtom As Long, mLeft As Long
                    
                    DrawStyle = GetNodeAttribute(start_at_node, node_name, "PicDrawStyle", 0, , 4)
                    mX = GetNodeAttribute(start_at_node, node_name, "PicX", 0, , 2147483647)
                    mY = GetNodeAttribute(start_at_node, node_name, "PicY", 0, , 2147483647)
                    mWidth = GetNodeAttribute(start_at_node, node_name, "PicWidth", 0, , 2147483647)
                    mHeight = GetNodeAttribute(start_at_node, node_name, "PicHeight", 0, , 2147483647)
                    xAlign = GetNodeAttribute(start_at_node, node_name, "PicxAlign", 0, , 2)
                    yAlign = GetNodeAttribute(start_at_node, node_name, "PicyAlign", 0, , 2)
                    mTop = GetNodeAttribute(start_at_node, node_name, "PicTop", 0, , 2147483647)
                    mRight = GetNodeAttribute(start_at_node, node_name, "PicRight", 0, , 2147483647)
                    mButtom = GetNodeAttribute(start_at_node, node_name, "PicButtom", 0, , 2147483647)
                    mLeft = GetNodeAttribute(start_at_node, node_name, "PicLeft", 0, , 2147483647)
                    
                    obj.Cls
                    Call PaintPng(newvalue, obj, DrawStyle, mX, mY, mWidth, mHeight, xAlign, yAlign, mTop, mRight, mButtom, mLeft)
                    obj.Refresh
                End If
            ElseIf IsOK(ValueName, "^PicDrawStyle$") Then
                '如果是图片按钮才允许改变图片画画设置
                If IsOK(TypeName(obj), "^CommandButtonEx$") Then
                    On Error Resume Next '出错则执行下一句
                    Call CallByName(obj, ValueName, VbLet, newvalue)
                End If
            ElseIf IsOK(ValueName, "^Caption$") Or IsOK(ValueName, "^Text$") Then
                On Error Resume Next '出错则执行下一句
                Call CallByName(obj, ValueName, VbLet, Path_to_String(newvalue))
            Else
                On Error Resume Next '出错则执行下一句
                Call CallByName(obj, ValueName, VbLet, newvalue)
            End If
            
        Next
    End If
End Sub
'格式化XML

Sub formatDoc(ByRef oDoc, sFilename)
    On Error Resume Next
    Dim oSAXWriter, oSAXReader
    Dim sErrMsg As String

    Set oSAXWriter = CreateObject("Msxml2.MXXMLWriter.6.0")
    Set oSAXReader = CreateObject("Msxml2.SAXXMLReader.6.0")
    With oSAXWriter
        .encoding = "UTF-8"
        .byteOrderMark = True
        .standalone = True
        .omitXMLDeclaration = False
        .indent = True
    End With
    With oSAXReader
        Set .contentHandler = oSAXWriter
        Set .dtdHandler = oSAXWriter
        Set .errorHandler = oSAXWriter
        .putProperty "http://xml.org/sax/properties/lexical-handler", oSAXWriter
        .putProperty "http://xml.org/sax/properties/declaration-handler", oSAXWriter
        .parse oDoc
    End With
    With oDoc
        .loadXML oSAXWriter.Output
        If .parseError.errorCode <> 0 Then
            sErrMsg = .parseError.errorCode & "|" & .parseError.srcText & "|" & .parseError.reason
            On Error GoTo a
            Err.Raise 30000, "formatDoc", sErrMsg
            Exit Sub
a:
        End If
        .Save sFilename
    End With
    Set oSAXWriter = Nothing
    Set oSAXReader = Nothing
End Sub
