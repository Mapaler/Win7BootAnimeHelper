VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "aImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvaraURL As String '局部复制
Private mvaraX As Integer '局部复制
Private mvaraY As Integer '局部复制
Public Property Let aY(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.aY = 5
    mvaraY = vData
End Property


Public Property Get aY() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.aY
    aY = mvaraY
End Property



Public Property Let aX(ByVal vData As Integer)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.aX = 5
    mvaraX = vData
End Property


Public Property Get aX() As Integer
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.aX
    aX = mvaraX
End Property



Public Property Let aURL(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.aURL = 5
    mvaraURL = vData
End Property


Public Property Get aURL() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.aURL
    aURL = mvaraURL
End Property



