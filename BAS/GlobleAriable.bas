Attribute VB_Name = "全局常量变量"
    Option Explicit
'----------------------------------------------------
'全局变量类型说明
'----------------------------------------------------
'Public Type FormatCfg
'    name As String
'    Extension() As String
'End Type
'----------------------------------------------------
'全局变量说明
'----------------------------------------------------
Public Const App_Beta = "" '本程序的当前Beta等版本名
Public testver As Boolean

'一般设置
Public cfgConfig_Url As String '设置文件地址
Public Const cfgConfig_FileName = "WBAHConfig.xml" '设置文件名
Public cfgConfig_XML As DOMDocument '保存设置XML变量
Public Const cfgConfig_XMLnode = "/Soft/Config"  '保存窗体设置XML路径
'窗体样式设置
Public cfgWindow_Url As String '
Public Const cfgWindow_FileName = "WBAHWindow.xml"
Public cfgWindow_XML As DOMDocument
Public Const cfgWindows_XMLnode = "/Soft/Window"  '保存窗体设置XML路径
''语言设置
'Public cfgLanguage_Url As String
'Public Const cfgLanguage_FileName = "WBAHConfig.xml"
'Public cfgLanguage_XML As DOMDocument
'Public Const cfgLanguage_XMLnode = "/Soft/Language"

Public Const UpdataURL = "http://www.mapaler.com/tools/checkver/index.php?soft=Win7BootAnimeHelper&mod=xmlverinfo&ver=newest" '升级监测网址
Public Const SoftPage = "http://www.mapaler.com/tools/Win7BootAnimeHelper"

Public Auto_Update As Boolean '是否自动升级

Public Changed As Boolean '是否被修改过
Public FormatProject As Boolean '保存时是否格式化
Public FrameFocus As Boolean '是否到转到帧焦点

Public ListDistance As Long '帧列表间距

Public aWBH As WBH
Public BG_URL As String '储存完整的背景地址
Public Frame_URL() As String '储存完整的帧地址
'----------------------------------------------------
'文件打开方式说明
'----------------------------------------------------
Public F_Proj As DlgFTlist '工程文件扩展名
Public F_IpG As DlgFTlist '导入图片
Public F_OpG As DlgFTlist '导出图片

'----------------------------------------------------
'部分通用正则表达式说明
'----------------------------------------------------
'正则表达式
Public Const FileURL_Parten = "(?:[A-Za-z]:|[\\/])[\\/][^:\*\?""<>\|]*" '真实文件路径正则表达式

Public Const CommandPath_Parten = "(?:(?:^|[^""])(?:[A-Za-z]:|[\\/])[\\/][^:\*\?""<>\| ]*|(?:"")((?:[A-Za-z]:|[\\/])[\\/][^:\*\?""<>\|]*)(?:""))" 'Command里的文件路径正则表达式
