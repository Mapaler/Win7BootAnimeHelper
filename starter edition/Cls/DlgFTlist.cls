VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DlgFTlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"FileType"
Attribute VB_Ext_KEY = "Member0" ,"FileType"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'局部变量，保存集合
Private mCol As Collection
Public Function Txt() As String
Attribute Txt.VB_UserMemId = 0
    Dim i As Byte, j As Byte, text As String
    Dim Jgf As String, eTemp() As String
    Jgf = Chr$(0)
    text = ""
    For i = 1 To mCol.Count
        text = text & mCol(i).TypeName
        text = text & Jgf
        eTemp = Split(mCol(i).Extensions, ",")
        For j = 0 To UBound(eTemp)
            text = text & "*." & eTemp(j)
            If j < UBound(eTemp) Then text = text & ";"
        Next j
        If i < mCol.Count Then text = text & Jgf
    Next i
    Txt = text
End Function


Public Function Add(TypeName As String, Extensions As String, Optional sKey As String) As FileType
    '创建新对象
    Dim objNewMember As FileType
    Set objNewMember = New FileType


    '设置传入方法的属性
    objNewMember.TypeName = TypeName
    objNewMember.Extensions = Extensions
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '返回已创建的对象
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As FileType
    '引用集合中的一个元素时使用。
    'vntIndexKey 包含集合的索引或关键字，
    '这是为什么要声明为 Variant 的原因
    '语法：Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    '检索集合中的元素数时使用。语法：Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    '删除集合中的元素时使用。
    'vntIndexKey 包含索引或关键字，这是为什么要声明为 Variant 的原因
    '语法：x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '本属性允许用 For...Each 语法枚举该集合。
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '创建类后创建集合
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    '类终止后破坏集合
    Set mCol = Nothing
End Sub

