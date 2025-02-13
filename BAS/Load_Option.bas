Attribute VB_Name = "加载设置文件"
Option Explicit
'加载软件初始设置
Public Sub Load_Option()
Dim i As Long
    If IsOK(App.EXEName, "test") Then '是不是测试模式
        testver = True
    Else
        testver = False
    End If
    
    cfgConfig_Url = App.Path & "\" & cfgConfig_FileName '设置文件地址
    Set cfgConfig_XML = New DOMDocument
    cfgConfig_XML.Load cfgConfig_Url '读取设置文件
    If cfgConfig_XML.documentElement Is Nothing Then
        MsgBox "加载XML设置文件失败，请检查 " & cfgConfig_FileName & " 是否配置正确" & vbCrLf & "Load XML config error,please check " & cfgConfig_FileName & vbCrLf & Now(), vbCritical
        End
    End If
    
    cfgWindow_Url = App.Path & "\" & cfgWindow_FileName '窗体设置文件地址
    Set cfgWindow_XML = New DOMDocument
    cfgWindow_XML.Load cfgWindow_Url '读取设置文件
    If cfgWindow_XML.documentElement Is Nothing Then
    End If
    
'    cfgLanguage_Url = App.Path & "\" & cfgLanguage_FileName '语言设置文件地址
'    Set cfgLanguage_XML = New DOMDocument
'    cfgLanguage_XML.Load cfgLanguage_Url '读取设置文件
'    If cfgLanguage_XML.documentElement Is Nothing Then
'    End If
    
    ListDistance = GetNodeAttribute(cfgWindow_XML, cfgWindows_XMLnode, "ListDistance", 10, , 2147483647)
    
    FormatProject = GetNodeAttribute(cfgConfig_XML, cfgConfig_XMLnode, "FormatProject", 0, , 1)
    FrameFocus = GetNodeAttribute(cfgConfig_XML, cfgConfig_XMLnode, "FrameFocus", 0, , 1)
    
    Set F_Proj = New DlgFTlist
    F_Proj.Add "WBAH工程文件(*.wba)", "wba"
    F_Proj.Add "XML文件(*.xml)", "xml"
    Set F_IpG = New DlgFTlist
    F_IpG.Add "图片文件(*.jpg;*.png;*.bmp;*.gif,*.tif)", "jpg,jpeg,png,bmp,gif,tif"
    F_IpG.Add "JPEG文件(*.jpg)", "jpg,jpeg"
    F_IpG.Add "PNG文件(*.png)", "png"
    F_IpG.Add "BMP文件(*.bmp)", "bmp"
    F_IpG.Add "GIF文件(*.gif)", "gif"
    F_IpG.Add "TIFF文件(*.tif)", "tif,tiff"
    Set F_OpG = New DlgFTlist
    F_OpG.Add "PNG文件(*.png)", "png"
    F_OpG.Add "BMP文件(*.bmp)", "bmp"
    F_OpG.Add "JPEG文件(*.jpg)", "jpg"
    F_OpG.Add "GIF文件(*.gif)", "gif"
    F_OpG.Add "TIFF文件(*.tif)", "tif"

End Sub
