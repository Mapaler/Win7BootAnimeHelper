VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'保持属性值的局部变量
Private mvarTypeName As String '局部复制
'保持属性值的局部变量
Private mvarExtensions As String '局部复制
Public Property Let Extensions(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.Extensions = 5
    mvarExtensions = vData
End Property


Public Property Get Extensions() As String
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.Extensions
    Extensions = mvarExtensions
End Property





Public Property Let TypeName(ByVal vData As String)
'向属性指派值时使用，位于赋值语句的左边。
'Syntax: X.TypeName = 5
    mvarTypeName = vData
End Property


Public Property Get TypeName() As String
Attribute TypeName.VB_UserMemId = 0
'检索属性值时使用，位于赋值语句的右边。
'Syntax: Debug.Print X.TypeName
    TypeName = mvarTypeName
End Property



