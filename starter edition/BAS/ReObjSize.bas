Attribute VB_Name = "重画窗口"
'判断是否为控件数组
Public Function IsControlArray(tForm As Form, ControlName As String) As Boolean
    If VarType(CallByName(tForm, ControlName, VbGet)) = vbObject Then
        IsControlArray = True
    Else
        IsControlArray = False
    End If
End Function
'从XML读取设置重画窗体
Public Sub ReDrawWindow_FromXML(tForm As Form)
Dim ob As control, tNodeUrl As String

    tNodeUrl = cfgWindows_XMLnode & "/" & tForm.name '窗体路径
    Call Set_All_Attribute_To_Obj(cfgWindow_XML, tNodeUrl, tForm)
'    Debug.Print tNodeUrl 'Debug输出
    
    For Each ob In tForm.Controls
    
        tNodeUrl = cfgWindows_XMLnode & "/" & tForm.name & "/" & ob.name '控件路径
        Call Set_All_Attribute_To_Obj(cfgWindow_XML, tNodeUrl, ob)
'        Debug.Print tNodeUrl 'Debug输出
        
        If IsControlArray(tForm, ob.name) Then '如果是控件数组，除了之前读取的共通属性外，再读取单个的属性
        
            tNodeUrl = cfgWindows_XMLnode & "/" & tForm.name & "/" & ob.name & "/Index_" & ob.Index '控件数组路径
            Call Set_All_Attribute_To_Obj(cfgWindow_XML, tNodeUrl, ob)
'            Debug.Print tNodeUrl 'Debug输出
            
        End If
    Next
End Sub
'保存窗体大小到设置
Public Sub SaveWindowToXML(tForm As Form)
    Call SaveAttribute(cfgWindow_XML, cfgWindows_XMLnode & "/" & tForm.name, "Width", tForm.Width / 15)
    Call SaveAttribute(cfgWindow_XML, cfgWindows_XMLnode & "/" & tForm.name, "Height", tForm.Height / 15)
End Sub
