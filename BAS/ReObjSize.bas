Attribute VB_Name = "�ػ�����"
'�ж��Ƿ�Ϊ�ؼ�����
Public Function IsControlArray(tForm As Form, ControlName As String) As Boolean
    If VarType(CallByName(tForm, ControlName, VbGet)) = vbObject Then
        IsControlArray = True
    Else
        IsControlArray = False
    End If
End Function
'��XML��ȡ�����ػ�����
Public Sub ReDrawWindow_FromXML(tForm As Form)
Dim ob As control, tNodeUrl As String

    tNodeUrl = cfgWindows_XMLnode & "/" & tForm.name '����·��
    Call Set_All_Attribute_To_Obj(cfgWindow_XML, tNodeUrl, tForm)
'    Debug.Print tNodeUrl 'Debug���
    
    For Each ob In tForm.Controls
    
        tNodeUrl = cfgWindows_XMLnode & "/" & tForm.name & "/" & ob.name '�ؼ�·��
        Call Set_All_Attribute_To_Obj(cfgWindow_XML, tNodeUrl, ob)
'        Debug.Print tNodeUrl 'Debug���
        
        If IsControlArray(tForm, ob.name) Then '����ǿؼ����飬����֮ǰ��ȡ�Ĺ�ͨ�����⣬�ٶ�ȡ����������
        
            tNodeUrl = cfgWindows_XMLnode & "/" & tForm.name & "/" & ob.name & "/Index_" & ob.Index '�ؼ�����·��
            Call Set_All_Attribute_To_Obj(cfgWindow_XML, tNodeUrl, ob)
'            Debug.Print tNodeUrl 'Debug���
            
        End If
    Next
End Sub
'���洰���С������
Public Sub SaveWindowToXML(tForm As Form)
    Call SaveAttribute(cfgWindow_XML, cfgWindows_XMLnode & "/" & tForm.name, "Width", tForm.Width / 15)
    Call SaveAttribute(cfgWindow_XML, cfgWindows_XMLnode & "/" & tForm.name, "Height", tForm.Height / 15)
End Sub
