Attribute VB_Name = "���������ļ�"
Option Explicit
'���������ʼ����
Public Sub Load_Option()
Dim i As Long
    If IsOK(App.EXEName, "test") Then '�ǲ��ǲ���ģʽ
        testver = True
    Else
        testver = False
    End If
    
    cfgConfig_Url = App.Path & "\" & cfgConfig_FileName '�����ļ���ַ
    Set cfgConfig_XML = New DOMDocument
    cfgConfig_XML.Load cfgConfig_Url '��ȡ�����ļ�
    If cfgConfig_XML.documentElement Is Nothing Then
        MsgBox "����XML�����ļ�ʧ�ܣ����� " & cfgConfig_FileName & " �Ƿ�������ȷ" & vbCrLf & "Load XML config error,please check " & cfgConfig_FileName & vbCrLf & Now(), vbCritical
        End
    End If
    
    cfgWindow_Url = App.Path & "\" & cfgWindow_FileName '���������ļ���ַ
    Set cfgWindow_XML = New DOMDocument
    cfgWindow_XML.Load cfgWindow_Url '��ȡ�����ļ�
    If cfgWindow_XML.documentElement Is Nothing Then
    End If
    
'    cfgLanguage_Url = App.Path & "\" & cfgLanguage_FileName '���������ļ���ַ
'    Set cfgLanguage_XML = New DOMDocument
'    cfgLanguage_XML.Load cfgLanguage_Url '��ȡ�����ļ�
'    If cfgLanguage_XML.documentElement Is Nothing Then
'    End If
    
    ListDistance = GetNodeAttribute(cfgWindow_XML, cfgWindows_XMLnode, "ListDistance", 10, , 2147483647)
    
    FormatProject = GetNodeAttribute(cfgConfig_XML, cfgConfig_XMLnode, "FormatProject", 0, , 1)
    FrameFocus = GetNodeAttribute(cfgConfig_XML, cfgConfig_XMLnode, "FrameFocus", 0, , 1)
    
    Set F_Proj = New DlgFTlist
    F_Proj.Add "WBAH�����ļ�(*.wba)", "wba"
    F_Proj.Add "XML�ļ�(*.xml)", "xml"
    Set F_IpG = New DlgFTlist
    F_IpG.Add "ͼƬ�ļ�(*.jpg;*.png;*.bmp;*.gif,*.tif)", "jpg,jpeg,png,bmp,gif,tif"
    F_IpG.Add "JPEG�ļ�(*.jpg)", "jpg,jpeg"
    F_IpG.Add "PNG�ļ�(*.png)", "png"
    F_IpG.Add "BMP�ļ�(*.bmp)", "bmp"
    F_IpG.Add "GIF�ļ�(*.gif)", "gif"
    F_IpG.Add "TIFF�ļ�(*.tif)", "tif,tiff"
    Set F_OpG = New DlgFTlist
    F_OpG.Add "PNG�ļ�(*.png)", "png"
    F_OpG.Add "BMP�ļ�(*.bmp)", "bmp"
    F_OpG.Add "JPEG�ļ�(*.jpg)", "jpg"
    F_OpG.Add "GIF�ļ�(*.gif)", "gif"
    F_OpG.Add "TIFF�ļ�(*.tif)", "tif"

End Sub
