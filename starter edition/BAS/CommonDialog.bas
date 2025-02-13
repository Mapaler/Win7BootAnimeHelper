Attribute VB_Name = "�򿪱���Ի���"
Option Explicit
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer

'����OPENFILENAME��������������ֵ
Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

'�����ʱ�ĸ����
Public Const OFN_READONLY = &H1
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_SHOWHELP = &H10
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Public Const OFN_EXPLORER = &H80000                         '  new look commdlg
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules

Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0

'�򿪶���ļ���
Type DlgFileInfo
 iCount As Long
 sPath As String
 sFile() As String
End Type

'���ܣ� ����CommonDialog��ѡ����ļ�������·�����ļ���
'����˵��: strFileNameΪCommonDialog��Filename����
'��������: DlgFileInfo������һ���Զ������ͣ�����iCount������ѡ���ļ��ĸ�����sPath������ѡ
'���ļ���·����sFile()������ѡ���ļ����ļ�����������·����
'ע������: �ú���Ӧ��CommonDialog.ShowOpen����������ʹ�ã����⵱ǰ·��������
Public Function GetDlgFileInfo(strFilename As String) As DlgFileInfo
 
 Dim sPath, tmpStr As String
 Dim sFile() As String
 Dim iCount As Integer
 Dim i As Integer
On Error GoTo ErrHandle
    strFilename = ReplaceText(strFilename, Chr(0) & "+$", "")
 sPath = CurDir()
 tmpStr = Right$(strFilename, Len(strFilename) - Len(sPath)) '���ļ�����·������
 
 If Left$(tmpStr, 1) = Chr$(0) Then
 'ѡ���˶���ļ�(������һ���ַ�ΪChr$(0))
 For i = 1 To Len(tmpStr)
 If Mid$(tmpStr, i, 1) = Chr$(0) Then
 iCount = iCount + 1
 ReDim Preserve sFile(iCount)
 Else
 sFile(iCount) = sFile(iCount) & Mid$(tmpStr, i, 1)
 End If
 Next i
 Else
 'ֻѡ����һ���ļ�(ע�⣺��Ŀ¼�µ��ļ�����ȥ·�������û��"\"��
 iCount = 1
 ReDim Preserve sFile(iCount)
 If Left$(tmpStr, 1) = "\" Then tmpStr = Right$(tmpStr, Len(tmpStr) - 1)
 sFile(iCount) = tmpStr
 End If
 
 GetDlgFileInfo.iCount = iCount
 ReDim GetDlgFileInfo.sFile(iCount)
 
 If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
 GetDlgFileInfo.sPath = sPath
 
 For i = 1 To iCount
 GetDlgFileInfo.sFile(i) = sFile(i)
 Next i
 
 Exit Function

ErrHandle:
' MsgBox "GetDlgFileInfo����ִ�д��󣡣����ļ���ʱ��ȡ���ˣ�", vbOKOnly + vbCritical, "�Զ��庯������"

End Function


