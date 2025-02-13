Attribute VB_Name = "打开保存对话框"
Option Explicit
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer

'设置OPENFILENAME类所包含的属性值
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

'定义打开时的各项常数
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

'打开多个文件用
Type DlgFileInfo
 iCount As Long
 sPath As String
 sFile() As String
End Type

'功能： 返回CommonDialog所选择的文件数量、路径和文件名
'参数说明: strFileName为CommonDialog的Filename属性
'函数类型: DlgFileInfo。这是一个自定义类型，其中iCount返回所选择文件的个数，sPath返回所选
'择文件的路径，sFile()返回所选择文件的文件名（不包括路径）
'注意事项: 该函数应在CommonDialog.ShowOpen方法后立即使用，以免当前路径被更改
Public Function GetDlgFileInfo(strFilename As String) As DlgFileInfo
 
 Dim sPath, tmpStr As String
 Dim sFile() As String
 Dim iCount As Integer
 Dim i As Integer
On Error GoTo ErrHandle
    strFilename = ReplaceText(strFilename, Chr(0) & "+$", "")
 sPath = CurDir()
 tmpStr = Right$(strFilename, Len(strFilename) - Len(sPath)) '将文件名与路径分离
 
 If Left$(tmpStr, 1) = Chr$(0) Then
 '选择了多个文件(分离后第一个字符为Chr$(0))
 For i = 1 To Len(tmpStr)
 If Mid$(tmpStr, i, 1) = Chr$(0) Then
 iCount = iCount + 1
 ReDim Preserve sFile(iCount)
 Else
 sFile(iCount) = sFile(iCount) & Mid$(tmpStr, i, 1)
 End If
 Next i
 Else
 '只选择了一个文件(注意：根目录下的文件名除去路径后左边没有"\"）
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
' MsgBox "GetDlgFileInfo函数执行错误！（无文件的时候取消了）", vbOKOnly + vbCritical, "自定义函数错误"

End Function


