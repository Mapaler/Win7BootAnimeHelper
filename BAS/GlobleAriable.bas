Attribute VB_Name = "ȫ�ֳ�������"
    Option Explicit
'----------------------------------------------------
'ȫ�ֱ�������˵��
'----------------------------------------------------
'Public Type FormatCfg
'    name As String
'    Extension() As String
'End Type
'----------------------------------------------------
'ȫ�ֱ���˵��
'----------------------------------------------------
Public Const App_Beta = "" '������ĵ�ǰBeta�Ȱ汾��
Public testver As Boolean

'һ������
Public cfgConfig_Url As String '�����ļ���ַ
Public Const cfgConfig_FileName = "WBAHConfig.xml" '�����ļ���
Public cfgConfig_XML As DOMDocument '��������XML����
Public Const cfgConfig_XMLnode = "/Soft/Config"  '���洰������XML·��
'������ʽ����
Public cfgWindow_Url As String '
Public Const cfgWindow_FileName = "WBAHWindow.xml"
Public cfgWindow_XML As DOMDocument
Public Const cfgWindows_XMLnode = "/Soft/Window"  '���洰������XML·��
''��������
'Public cfgLanguage_Url As String
'Public Const cfgLanguage_FileName = "WBAHConfig.xml"
'Public cfgLanguage_XML As DOMDocument
'Public Const cfgLanguage_XMLnode = "/Soft/Language"

Public Const UpdataURL = "http://www.mapaler.com/tools/checkver/index.php?soft=Win7BootAnimeHelper&mod=xmlverinfo&ver=newest" '���������ַ
Public Const SoftPage = "http://www.mapaler.com/tools/Win7BootAnimeHelper"

Public Auto_Update As Boolean '�Ƿ��Զ�����

Public Changed As Boolean '�Ƿ��޸Ĺ�
Public FormatProject As Boolean '����ʱ�Ƿ��ʽ��
Public FrameFocus As Boolean '�Ƿ�ת��֡����

Public ListDistance As Long '֡�б���

Public aWBH As WBH
Public BG_URL As String '���������ı�����ַ
Public Frame_URL() As String '����������֡��ַ
'----------------------------------------------------
'�ļ��򿪷�ʽ˵��
'----------------------------------------------------
Public F_Proj As DlgFTlist '�����ļ���չ��
Public F_IpG As DlgFTlist '����ͼƬ
Public F_OpG As DlgFTlist '����ͼƬ

'----------------------------------------------------
'����ͨ��������ʽ˵��
'----------------------------------------------------
'������ʽ
Public Const FileURL_Parten = "(?:[A-Za-z]:|[\\/])[\\/][^:\*\?""<>\|]*" '��ʵ�ļ�·��������ʽ

Public Const CommandPath_Parten = "(?:(?:^|[^""])(?:[A-Za-z]:|[\\/])[\\/][^:\*\?""<>\| ]*|(?:"")((?:[A-Za-z]:|[\\/])[\\/][^:\*\?""<>\|]*)(?:""))" 'Command����ļ�·��������ʽ
