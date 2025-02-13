VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WBH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"aFrame"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'保持属性值的局部变量
Private mvaraName As String '局部复制
Private mvaraWidth As Integer '局部复制
Private mvaraHeight As Integer '局部复制
Private mvaraChanged As Boolean '局部复制
Private mvaraFilePath As String '局部复制
Private mvaraFileFolder As String '局部复制
Private mvaraBackground As aImage '局部复制
Private mvaraFrame As aFrame
Private Const dataXML_Node = "/WBA"

Public Property Let aFileFolder(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.aFileFolder = 5
    mvaraFileFolder = vData
End Property


Public Property Get aFileFolder() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.aFileFolder
    aFileFolder = mvaraFileFolder
End Property

Public Sub Export(ByVal BGURL As String)
End Sub

Public Property Let aFilePath(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.aFilePath = 5
    mvaraFilePath = vData
End Property


Public Property Get aFilePath() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.aFilePath
    aFilePath = mvaraFilePath
End Property



Public Property Let aChanged(ByVal vData As Boolean)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.aChanged = 5
    mvaraChanged = vData
End Property


Public Property Get aChanged() As Boolean
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.aChanged
    aChanged = mvaraChanged
End Property



Public Sub AddNew()
    
    mvaraName = "Win7全屏开机动画"
    mvaraWidth = "1024"
    mvaraHeight = "768"
    
    Set mvaraBackground = New aImage
    mvaraBackground.aURL = "NewWBA\Background.jpg"
    mvaraBackground.aX = 0
    mvaraBackground.aY = 0
    
    Set mvaraFrame = New aFrame
    For i = 0 To 104
        mvaraFrame.Add "NewWBA\" & format(i, "000") & ".png", 0, 0
    Next
    
    mvaraChanged = False
End Sub

Public Sub OpenFile(ByVal url As String)
    Dim i As Long

    Dim savefile_XML As DOMDocument
    Set savefile_XML = New DOMDocument
    savefile_XML.Load url '读取设置文件
    If savefile_XML.documentElement Is Nothing Then
        MsgBox "工程文件读取失败", vbCritical
        Exit Sub
    End If
    If savefile_XML.selectSingleNode("/WBA/Config") Is Nothing Then
        MsgBox "该文件不是WBAH工程文件", vbCritical
        Exit Sub
    End If
    mvaraFilePath = url
    If InStrRev(url, "\") > 1 Then
        mvaraFileFolder = Left(url, InStrRev(url, "\") - 1)
    End If
    
    mvaraName = GetNodeAttribute(savefile_XML, "" & dataXML_Node & "/Config", "Name")
    mvaraWidth = GetNodeAttribute(savefile_XML, "" & dataXML_Node & "/Config", "Width", 1024, , 65535)
    mvaraHeight = GetNodeAttribute(savefile_XML, "" & dataXML_Node & "/Config", "Height", 1024, , 65535)
    
    Set mvaraBackground = New aImage
    mvaraBackground.aURL = GetNodeAttribute(savefile_XML, "" & dataXML_Node & "/Config/Background", "Url")
    mvaraBackground.aX = GetNodeAttribute(savefile_XML, "" & dataXML_Node & "/Config/Background", "X", 0, , 65535)
    mvaraBackground.aY = GetNodeAttribute(savefile_XML, "" & dataXML_Node & "/Config/Background", "Y", 0, , 65535)
    
    Set mvaraFrame = New aFrame
    For i = 0 To GetAllNode_Lenth(savefile_XML, "" & dataXML_Node & "/Config/Frames/Image") - 1
        mvaraFrame.Add GetNodeAttribute(savefile_XML, "" & dataXML_Node & "/Config/Frames/Image", "Url", , i), GetNodeAttribute(savefile_XML, "" & dataXML_Node & "/Config/Frames/Image", "X", 0, i, 65535), GetNodeAttribute(savefile_XML, "" & dataXML_Node & "/Config/Frames/Image", "Y", 0, i, 65535)
    Next
    
    mvaraChanged = False
End Sub

Public Property Set aBackground(ByVal vData As aImage)
'向属性指派对象时使用，位于 Set 语句的左边。
'Syntax: Set x.aBackground = Form1
    Set mvaraBackground = vData
End Property


Public Property Get aBackground() As aImage
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.aBackground
    Set aBackground = mvaraBackground
End Property





Public Property Get aFrame() As aFrame
    If mvaraFrame Is Nothing Then
        Set mvaraFrame = New aFrame
    End If


    Set aFrame = mvaraFrame
End Property


Public Property Set aFrame(vData As aFrame)
    Set mvaraFrame = vData
End Property
Private Sub Class_Terminate()
    Set mvaraFrame = Nothing
End Sub



Public Property Let aHeight(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.aHeight = 5
    mvaraHeight = vData
End Property


Public Property Get aHeight() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.aHeight
    aHeight = mvaraHeight
End Property



Public Property Let aWidth(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.aWidth = 5
    mvaraWidth = vData
End Property


Public Property Get aWidth() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.aWidth
    aWidth = mvaraWidth
End Property



Public Property Let aName(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.aName = 5
    mvaraName = vData
End Property


Public Property Get aName() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.aName
    aName = mvaraName
End Property



