Attribute VB_Name = "Sub_Main"
Option Explicit
Sub Main()
'��ʼ������
    Call Load_Option
'����������
    frmEditer.Show
End Sub

'�ر����ʱд������
Public Sub EndSoft()
    Call cfgWindow_XML.Save(cfgWindow_Url)
    Call cfgConfig_XML.Save(cfgConfig_Url)
    End
End Sub

'�ж��Ƿ��ǿ�������
Public Function IsDebugMode() As Boolean
    IsDebugMode = False
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then
        IsDebugMode = True
    End If
End Function
