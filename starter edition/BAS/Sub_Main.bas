Attribute VB_Name = "Sub_Main"
Option Explicit
Sub Main()
'初始化设置
    Call Load_Option
'加载主窗口
    frmEditer.Show
End Sub

'关闭软件时写入设置
Public Sub EndSoft()
    Call cfgWindow_XML.Save(cfgWindow_Url)
    Call cfgConfig_XML.Save(cfgConfig_Url)
    End
End Sub

'判断是否是开发环境
Public Function IsDebugMode() As Boolean
    IsDebugMode = False
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then
        IsDebugMode = True
    End If
End Function
