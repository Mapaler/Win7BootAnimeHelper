VERSION 5.00
Begin VB.Form frmEditer 
   AutoRedraw      =   -1  'True
   Caption         =   "Win7�������������������װ�"
   ClientHeight    =   5310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9945
   Icon            =   "frmEditer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   663
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdExportLongPic 
      Caption         =   "����ԭʼ��ͼ"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Timer Timer_Update 
      Left            =   600
      Top             =   4920
   End
   Begin VB.CommandButton cmdProject_Back 
      Caption         =   "��ԭ����(&R)"
      Height          =   375
      Left            =   1680
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton cmdProject_Browse 
      Caption         =   "�򿪹���(&O)"
      Height          =   375
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Timer timPlay 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   1440
      Top             =   4920
   End
   Begin VB.TextBox txtBackground 
      Height          =   375
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   4
      Top             =   1680
      Width           =   3015
   End
   Begin VB.CommandButton cmdBackground_Browse 
      Caption         =   "���(&B)"
      Height          =   375
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ComboBox comScreenScale 
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton cmdBackground_Setting 
      Caption         =   "����"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdFullscreen 
      Caption         =   "ȫ��(&F)"
      Height          =   615
      Left            =   1680
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "����(E)"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   3015
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "����(&P)"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.PictureBox picBG 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5085
      Left            =   3240
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   335
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   437
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.PictureBox picA 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3000
         Index           =   1
         Left            =   120
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   200
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   200
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.PictureBox picA 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3000
         Index           =   0
         Left            =   1680
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   200
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   200
         TabIndex        =   11
         Top             =   960
         Width           =   3000
      End
   End
   Begin VB.Label lblCurrentFrame 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0/105"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label lbltProject 
      BackStyle       =   0  'Transparent
      Caption         =   "����·��"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblBackground 
      BackStyle       =   0  'Transparent
      Caption         =   "����ͼƬ"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblScreenScale 
      BackStyle       =   0  'Transparent
      Caption         =   "��Ļ����"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "frmEditer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SizeTime_X As Double, SizeTime_Y As Double
Private WindowCaption As String '�洢����ԭʼ����
Private ProjectFileName As String '�洢�����ļ���

Private LoadFinish As Boolean '�����Ƿ�������

Private CurrentFrame As Long '��ǰ֡
Private MaxFrame As Long '���֡
Private Const ChengHao = "*" '"��"


Private Sub cmdBackground_Browse_Click()
    Dim file As OPENFILENAME, lResult As Long
    Dim DlgInfo As DlgFileInfo
    
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
    file.lpstrFile = String$(32767, 0)   '����Ĭ��Ҫ�򿪵��ļ���
    file.nMaxFile = 255 '��ʾ�ļ����ĳ���
    file.lpstrFileTitle = String$(255, 0) '�򿪶Ի���ı���
    file.nMaxFileTitle = 255 '�򿪶Ի���ı���ĳ���
    file.lpstrInitialDir = String$(255, 0)  '���ó�ʼ·��
    file.lpstrFilter = F_IpG.Txt '"SkyDrive��ҳ�����ļ�" & Chr$(0) & "*.*" '�򿪵��ļ�����"
    file.nFilterIndex = 1
    file.lpstrTitle = "ѡ��ͼƬ�ļ�"
    lResult = GetOpenFileName(file) 'ȡ���ļ���
    
    Dim url_t As String
    If lResult <> 0 Then
        DlgInfo = GetDlgFileInfo(file.lpstrFile)
        url_t = DlgInfo.sPath & DlgInfo.sFile(1)
    Else
        Exit Sub
    End If
    
    If Len(aWBH.aFileFolder) > 0 And InStr(1, url_t, aWBH.aFileFolder, vbTextCompare) = 1 Then
        txtBackground.text = Right(url_t, Len(url_t) - Len(aWBH.aFileFolder) - 1)
    Else
        txtBackground.text = url_t
    End If
End Sub

Private Sub cmdBackground_Browse_OLEDragDrop(data As DataObject, effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim extension As String
    If data.Files.Count > 0 Then '��������ļ�
        If data.Files(1) <> "" And Dir(data.Files(1)) <> "" Then '����ļ�����
            extension = Mid(data.Files(1), InStrRev(data.Files(1), ".") + 1)
            If extension = "jpg" Or extension = "jpeg" Or extension = "bmp" Or extension = "png" Or extension = "gif" Then
                If Len(aWBH.aFileFolder) > 0 And InStr(1, data.Files(1), aWBH.aFileFolder, vbTextCompare) = 1 Then
                    txtBackground.text = Right(data.Files(1), Len(data.Files(1)) - Len(aWBH.aFileFolder) - 1)
                Else
                    txtBackground.text = data.Files(1)
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdBackground_Setting_Click()
    frmSetting.PicIndex = 0
    frmSetting.Show 1
End Sub

Private Sub cmdExport_Click()
    Dim i As Long, extension As String
    Dim file As OPENFILENAME, lResult As Long
    Dim DlgInfo As DlgFileInfo
    
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
    file.lpstrFile = aWBH.aName & String$(255, 0)   '����Ĭ��Ҫ�򿪵��ļ���
    file.nMaxFile = 255 '��ʾ�ļ����ĳ���
    file.lpstrFileTitle = String$(255, 0) '�ļ���(��·��)
    file.nMaxFileTitle = 255 '�򿪶Ի���ı���ĳ���
    file.lpstrInitialDir = String$(255, 0)  '���ó�ʼ·��
    file.lpstrFilter = F_OpG.Txt
    file.nFilterIndex = 0
    file.lpstrTitle = "ѡ�񹤳̵���λ��"
    lResult = GetSaveFileName(file) 'ȡ���ļ���
    
'    Select Case file.nFilterIndex
'        Case 1
'            extension = ".png"
'        Case 2
'            extension = ".bmp"
'        Case 3
'            extension = ".jpg"
'        Case 4
'            extension = ".gif"
'    End Select
    extension = "." & F_OpG.Item(file.nFilterIndex).Extensions
    
    Dim url_t As String
    If lResult <> 0 Then
        DlgInfo = GetDlgFileInfo(file.lpstrFile)
        url_t = DlgInfo.sPath & DlgInfo.sFile(1)
        'If Right(url_t, 4) <> Extension Then url_t = url_t & Extension
        If InStrRev(url_t, extension) = 0 Then
            url_t = url_t & extension
        End If
    Else
        Exit Sub
    End If
    
    Dim SizeTime_X_out As Double, SizeTime_Y_out As Double, folderp As String, foldern As String, folderfull As String
    SizeTime_X_out = 1024 / aWBH.aWidth
    SizeTime_Y_out = 768 / aWBH.aHeight
    
    '��ȡ·��
    folderp = DlgInfo.sPath
'    If InStrRev(url_t, "\") > 1 Then
'        folderp = Left(url_t, InStrRev(url_t, "\") - 1)
'    End If
    '��ȡ�ļ���
    If InStrRev(DlgInfo.sFile(1), extension) > 1 Then
        foldern = Left(DlgInfo.sFile(1), InStrRev(DlgInfo.sFile(1), extension) - 1)
    Else
        foldern = DlgInfo.sFile(1)
    End If

    Dim Bitmap_BGout As Long, Bitmap_BGt As Long, Bitmap_Fout As Long, Bitmap_Ft As Long, Graphics As Long
    Dim bmW_BG As Long, bmH_BG As Long, bmW_F As Long, bmH_F As Long
    InitGDIPlus
    
    '���ļ�����Bitmap
    GdipCreateBitmapFromFile StrPtr(BG_URL), Bitmap_BGt
    GdipGetImageWidth Bitmap_BGt, bmW_BG
    GdipGetImageHeight Bitmap_BGt, bmH_BG

    CreateBitmapWithGraphics Bitmap_BGout, Graphics, 1024, 768 '�ؼ�������һ��Image��Graphics����

    GdipDrawImageRectI Graphics, Bitmap_BGt, aWBH.aBackground.aX * SizeTime_X_out, aWBH.aBackground.aY * SizeTime_Y_out, Fixb(bmW_BG * SizeTime_X_out), Fixb(bmH_BG * SizeTime_Y_out)
    
    '���������ڻ���Bitmap��
    Select Case extension
        Case ".png"
            SaveImageToPNG Bitmap_BGout, url_t
        Case ".bmp"
            SaveImageToBMP Bitmap_BGout, url_t
        Case ".jpg"
            SaveImageToJPG Bitmap_BGout, url_t
        Case ".gif"
            SaveImageToGIF Bitmap_BGout, url_t
        Case ".tif"
            SaveImageToTIF Bitmap_BGout, url_t
    End Select
    
    'û���ļ���ʱ
    folderfull = folderp & foldern & "_Frame"
    If Dir(folderfull, vbDirectory) = "" Then
        SHCreateDirectoryEx Me.hWnd, folderfull, ByVal 0&
    End If
    
    For i = 1 To aWBH.aFrame.Count    '���ļ�����Bitmap
        GdipCreateBitmapFromFile StrPtr(Frame_URL(i)), Bitmap_Ft
        GdipGetImageWidth Bitmap_Ft, bmW_F
        GdipGetImageHeight Bitmap_Ft, bmH_F

        CreateBitmapWithGraphics Bitmap_Fout, Graphics, 200, 200 '�ؼ�������һ��Image��Graphics����

        GdipDrawImageRectI Graphics, Bitmap_BGt, aWBH.aBackground.aX * SizeTime_X_out - 412, aWBH.aBackground.aY * SizeTime_Y_out - 284, Fixb(bmW_BG * SizeTime_X_out), Fixb(bmH_BG * SizeTime_Y_out)
        GdipDrawImageRectI Graphics, Bitmap_Ft, aWBH.aFrame(i).aX * SizeTime_X_out, aWBH.aFrame(i).aY * SizeTime_Y_out, Fixb(bmW_F * SizeTime_X_out), Fixb(bmH_F * SizeTime_Y_out)

        '���������ڻ���Bitmap��
        Select Case extension
            Case ".png"
                SaveImageToPNG Bitmap_Fout, folderfull & "\" & format(i, "000") & extension
            Case ".bmp"
                SaveImageToBMP Bitmap_Fout, folderfull & "\" & format(i, "000") & extension
            Case ".jpg"
                SaveImageToJPG Bitmap_Fout, folderfull & "\" & format(i, "000") & extension
            Case ".gif"
                SaveImageToGIF Bitmap_Fout, folderfull & "\" & format(i, "000") & extension
            Case ".tif"
                SaveImageToTIF Bitmap_Fout, folderfull & "\" & format(i, "000") & extension
        End Select
    Next


    'ɨ�ع���
    GdipDeleteGraphics Graphics
    
    GdipDisposeImage Bitmap_BGt
    GdipDisposeImage Bitmap_Ft
    GdipDisposeImage Bitmap_BGout
    GdipDisposeImage Bitmap_Fout
    
    TerminateGDIPlus
    
    MsgBox "����ͼƬ�ѱ��浽��" & url_t & "����" & vbCrLf & "֡�����ļ�������" & foldern & "_Frame��", 64, "�������"
End Sub

Private Sub cmdExportLongPic_Click()
    Dim i As Long, extension As String
    Dim file As OPENFILENAME, lResult As Long
    Dim DlgInfo As DlgFileInfo
    
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
    file.lpstrFile = "Activity_" & aWBH.aName & String$(255, 0)   '����Ĭ��Ҫ�򿪵��ļ���
    file.nMaxFile = 255 '��ʾ�ļ����ĳ���
    file.lpstrFileTitle = String$(255, 0) '�ļ���(��·��)
    file.nMaxFileTitle = 255 '�򿪶Ի���ı���ĳ���
    file.lpstrInitialDir = String$(255, 0)  '���ó�ʼ·��
    file.lpstrFilter = F_OpG.Txt
    file.nFilterIndex = 2 'bmp
    file.lpstrTitle = "ѡ�񱣴�ĳ�ͼ�ļ���"
    lResult = GetSaveFileName(file) 'ȡ���ļ���
    
    Dim url_t As String
    If lResult <> 0 Then
                extension = "." & F_OpG.Item(file.nFilterIndex).Extensions
        DlgInfo = GetDlgFileInfo(file.lpstrFile)
        url_t = DlgInfo.sPath & DlgInfo.sFile(1)
        '��ȡ����չ���ļ���
        If InStrRev(url_t, extension) = 0 Then
            url_t = url_t & extension
        End If
    Else
        Exit Sub
    End If
    
    Dim bgAlpha As Boolean, LongT As String
    LongT = MsgBox("�Ƿ�ȥ��ȫ������", vbYesNoCancel Or vbQuestion)
    If LongT = vbCancel Then
        Exit Sub
    ElseIf LongT = vbYes Then
        bgAlpha = True
    ElseIf LongT = vbNo Then
        bgAlpha = False
    End If
    
    Dim SizeTime_X_out As Double, SizeTime_Y_out As Double, folderp As String, foldern As String, folderfull As String
    SizeTime_X_out = 1024 / aWBH.aWidth
    SizeTime_Y_out = 768 / aWBH.aHeight

    Dim Bitmap_BGt As Long, Bitmap_Ft As Long, Bitmap_Fout As Long, Graphics As Long, Bitmap_FoutTemp As Long, GraphicsTemp As Long
    Dim bmW_BG As Long, bmH_BG As Long, bmW_F As Long, bmH_F As Long
    
    InitGDIPlus
    
    '���������͸�����ȡ����
    If bgAlpha = False Then
        '���ļ����뱳��
        GdipCreateBitmapFromFile StrPtr(BG_URL), Bitmap_BGt
        GdipGetImageWidth Bitmap_BGt, bmW_BG
        GdipGetImageHeight Bitmap_BGt, bmH_BG
    End If
    
    CreateBitmapWithGraphics Bitmap_Fout, Graphics, 200, 21000 '������ͼ
    
    Dim i0 As Integer
    For i = 1 To aWBH.aFrame.Count '��ѭ��֡
        
        GdipCreateBitmapFromFile StrPtr(Frame_URL(i)), Bitmap_Ft
        GdipGetImageWidth Bitmap_Ft, bmW_F
        GdipGetImageHeight Bitmap_Ft, bmH_F

        CreateBitmapWithGraphics Bitmap_FoutTemp, GraphicsTemp, 200, 200     '����Сͼ

        If bgAlpha = False Then
            GdipDrawImageRectI GraphicsTemp, Bitmap_BGt, aWBH.aBackground.aX * SizeTime_X_out - 412, aWBH.aBackground.aY * SizeTime_Y_out - 284, Fixb(bmW_BG * SizeTime_X_out), Fixb(bmH_BG * SizeTime_Y_out)
        End If
        GdipDrawImageRectI GraphicsTemp, Bitmap_Ft, aWBH.aFrame(i).aX * SizeTime_X_out, aWBH.aFrame(i).aY * SizeTime_Y_out, Fixb(bmW_F * SizeTime_X_out), Fixb(bmH_F * SizeTime_Y_out)
        
        GdipDrawImageRectI Graphics, Bitmap_FoutTemp, 0, (i - 1) * 200, 200, 200
    Next
    
    Select Case extension
        Case ".png"
            SaveImageToPNG Bitmap_Fout, url_t
        Case ".bmp"
            SaveImageToBMP Bitmap_Fout, url_t
        Case ".jpg"
            SaveImageToJPG Bitmap_Fout, url_t
        Case ".gif"
            SaveImageToGIF Bitmap_Fout, url_t
        Case ".tif"
            SaveImageToTIF Bitmap_Fout, url_t
    End Select


    'ɨ�ع���
    GdipDeleteGraphics Graphics
    GdipDeleteGraphics GraphicsTemp
    
    GdipDisposeImage Bitmap_BGt
    GdipDisposeImage Bitmap_Ft
    GdipDisposeImage Bitmap_Fout
    GdipDisposeImage Bitmap_FoutTemp
    
    TerminateGDIPlus
    
    MsgBox "������ͼ������ɣ������ʹ�á�ħ���������Ӧ�õ�ϵͳ��", 64, "�������"
End Sub

Private Sub cmdFullscreen_Click()
    If timPlay.Enabled = True Then
        Call cmdPlay_Click
        frmFullscreen.timPlay.Enabled = True
    End If
    frmFullscreen.nFrame = CurrentFrame
    frmFullscreen.Show 1
End Sub

Private Sub cmdPlay_Click()
    If timPlay.Enabled = False Then
        timPlay.Enabled = True
        cmdPlay.Caption = "ֹͣ"
    Else
        timPlay.Enabled = False
        CurrentFrame = 1
        PicRedraw (1)
        cmdPlay.Caption = "��ʼ"
    End If
End Sub

Private Sub cmdProject_Back_Click()
    Dim x As Integer
    If aWBH.aFilePath <> "" And Dir(aWBH.aFilePath) <> "" Then
        x = MsgBox("��ԭΪ���һ�α����" & ProjectFileName & "�汾��", vbQuestion Or vbYesNo)
        If x = vbYes Then '��
            Dim url_t As String
            url_t = aWBH.aFilePath
            Set aWBH = New WBH
            aWBH.OpenFile url_t
            
            Call ReadWBH
            NeiCun_Timer '�����ڴ�
        End If
        If x = vbNo Then '��
            Exit Sub
        End If
    Else
        MsgBox ("�����ļ�����λ�ò�����")
    End If
End Sub

Private Sub cmdProject_Browse_Click()
    Dim file As OPENFILENAME, lResult As Long
    Dim DlgInfo As DlgFileInfo
    
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
    file.lpstrFile = String$(32767, 0)   '����Ĭ��Ҫ�򿪵��ļ���
    file.nMaxFile = 255 '��ʾ�ļ����ĳ���
    file.lpstrFileTitle = String$(255, 0) '�򿪶Ի���ı���
    file.nMaxFileTitle = 255 '�򿪶Ի���ı���ĳ���
    file.lpstrInitialDir = String$(255, 0)  '���ó�ʼ·��
    file.lpstrFilter = F_Proj.Txt '"�ļ�" & Chr$(0) & "*.*" '�򿪵��ļ�����"
    file.nFilterIndex = 1
    file.lpstrTitle = "�򿪹����ļ�"
    lResult = GetOpenFileName(file) 'ȡ���ļ���
    
    Dim url_t As String
    If lResult <> 0 Then
        DlgInfo = GetDlgFileInfo(file.lpstrFile)
        url_t = DlgInfo.sPath & DlgInfo.sFile(1)
    Else
        Exit Sub
    End If
    
    OpenWBH (url_t)
End Sub

Private Sub comScreenScale_Change()
    Call comScreenScale_Click
End Sub

'�����µ���Ļ����
Private Sub comScreenScale_Click()
    If LoadFinish = True Then
        Dim regt() As PatternValue, regnum As Byte
        If IsOK(comScreenScale.text, "\d+\" & ChengHao & "\d+") Then
            regnum = SearchText(comScreenScale.text, "(\d+)\" & ChengHao & "(\d+)", regt(), "$1" & Chr(0) & "$2")
            If regnum > 0 Then
                aWBH.aWidth = regt(0).InValue(0)
                aWBH.aHeight = regt(0).InValue(1)
            End If
        End If
        Call Form_Resize '�ػ�����
    End If
End Sub

Private Sub comScreenScale_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 42 Then
        KeyAscii = 0
    End If
End Sub

'�������
Private Sub Form_Load()
    Dim i As Integer

    Auto_Update = True
    
    Timer_Update.Interval = 1000
    Timer_Update.Enabled = True
    
    '��ȡ��Ļ����
    For i = 0 To GetAllNode_Lenth(cfgConfig_XML, "/Soft/ScreenScale/Scale") - 1
        comScreenScale.AddItem GetNodeAttribute(cfgConfig_XML, "/Soft/ScreenScale/Scale", "Width", , i) _
        & ChengHao & GetNodeAttribute(cfgConfig_XML, "/Soft/ScreenScale/Scale", "Height", , i)
    Next

    Call NewWBAH '����½�
    
    Call ReDrawWindow_FromXML(Me) '�ػ�����
    
    WindowCaption = Me.Caption
    
    If Len(Command) > 0 Then
        Dim filepath() As PatternValue, pathn As Byte
        pathn = SearchText(Command, CommandPath_Parten, filepath, "$1" & Chr(0) & "$2")
        If pathn > 0 Then
            '��һ�����Լ���
            If Len(filepath(0).InValue(0)) > 0 Then
            '��ð��
                OpenWBH filepath(0).InValue(0)
            Else
            '��ð��
                OpenWBH filepath(0).AllValue
            End If
            '�����������Ĵ�
            For i = 1 To pathn - 1
                Shell App.EXEName & " " & filepath(i).AllValue
            Next
        End If
    End If
    
    NeiCun_Timer '�����ڴ�
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWindowToXML(Me)
    Call EndSoft
End Sub
'����ı��С
Private Sub Form_Resize()
    Dim meW As Long, meH As Long
    Dim Border As Integer, sldRealWidth As Integer
    Border = 10
    
    meW = Me.ScaleWidth
    meH = Me.ScaleHeight

    picBG.Height = NotLess(meH - picBG.Top - Border)
    picBG.Width = NotLess(picBG.Height * (aWBH.aWidth / aWBH.aHeight))
    picBG.Left = meW - picBG.Width - Border
    
    lbltProject.Width = NotLess(picBG.Left - lbltProject.Left - Border)
    cmdProject_Browse.Width = NotLess(picBG.Left - cmdProject_Browse.Left - Border * 2) / 2
    cmdProject_Back.Width = cmdProject_Browse.Width
    cmdProject_Back.Left = cmdProject_Browse.Width + cmdProject_Browse.Left + Border
        
    lblScreenScale.Width = NotLess(picBG.Left - lblScreenScale.Left - Border)
    comScreenScale.Width = NotLess(picBG.Left - comScreenScale.Left - Border)
    
    lblBackground.Width = NotLess(picBG.Left - lblBackground.Left - Border)
    txtBackground.Width = NotLess(picBG.Left - txtBackground.Left - Border)
    
    cmdBackground_Browse.Width = NotLess(picBG.Left - cmdBackground_Browse.Left - Border * 2) / 2
    cmdBackground_Setting.Width = cmdBackground_Browse.Width
    cmdBackground_Setting.Left = cmdBackground_Browse.Width + cmdBackground_Browse.Left + Border
    
    cmdPlay.Width = NotLess(picBG.Left - cmdPlay.Left - Border * 2) / 2
    cmdFullscreen.Width = cmdPlay.Width
    cmdFullscreen.Left = cmdPlay.Width + cmdPlay.Left + Border
    
    cmdExport.Width = NotLess(picBG.Left - cmdExport.Left - Border)
    cmdExport.Top = meH - cmdExport.Height - Border
    lblCurrentFrame.Width = NotLess(picBG.Left - lblCurrentFrame.Left - Border)
    lblCurrentFrame.Top = cmdExport.Top - lblCurrentFrame.Height - Border
    cmdExportLongPic.Width = NotLess(picBG.Left - cmdExportLongPic.Left - Border)
    cmdExportLongPic.Top = lblCurrentFrame.Top - cmdExportLongPic.Height - Border
    
    Call PicResize(picBG.Width, picBG.Height)
    SizeTime_X = picBG.Width / aWBH.aWidth
    SizeTime_Y = picBG.Height / aWBH.aHeight
    
    '��ͼ
    PicRedraw (0)
    
End Sub
Public Sub PicRedraw(Mode As Byte)
    Select Case Mode
        Case 0
            '��ͼ
            If IsOK(BG_URL, FileURL_Parten) And Dir(BG_URL) <> "" Then
                picBG.Cls
                Call PaintPng(BG_URL, picBG, 0, aWBH.aBackground.aX * SizeTime_X, aWBH.aBackground.aY * SizeTime_Y, , , , , , , , , SizeTime_X, SizeTime_Y)
                picBG.Refresh
                If IsOK(BG_URL, FileURL_Parten) And Dir(BG_URL) <> "" Then '������תͼƬ���ӿ������ٶ�
                    picA(1).Cls
                    Call PaintPng(BG_URL, picA(1), 0, aWBH.aBackground.aX * SizeTime_X - picA(0).Left, aWBH.aBackground.aY * SizeTime_Y - picA(0).Top, , , , , , , , , SizeTime_X, SizeTime_Y)
                    picA(1).Refresh
                End If
                PicRedraw (1)
            End If
        Case 1
            '��ͼ
            If IsOK(Frame_URL(CurrentFrame), FileURL_Parten) And Dir(Frame_URL(CurrentFrame)) <> "" Then
                'If IsOK(BG_URL, FileURL_Parten) And Dir(BG_URL) <> "" Then
                '    Call PaintPng(BG_URL, picA(0), 0, aWBH.aBackground.aX - picA(0).Left, aWBH.aBackground.aY - picA(0).Top, , , , , , , , , SizeTime_X)
                'End If
                picA(0).PaintPicture picA(1).Image, 0, 0
                Call PaintPng(Frame_URL(CurrentFrame), picA(0), 0, aWBH.aFrame(CurrentFrame).aX * SizeTime_X, aWBH.aFrame(CurrentFrame).aY * SizeTime_Y, , , , , , , , , SizeTime_X, SizeTime_Y)
                picA(0).Refresh
                lblCurrentFrame.Caption = CurrentFrame & "/" & MaxFrame
            End If
    End Select
End Sub
'Ԥ�����ڵ��ػ�
Private Sub PicResize(wW As Long, wH As Long)
    Dim PicWc As Long, PicHc As Long
    PicWc = picBG.Width - picBG.ScaleWidth
    PicHc = picBG.Height - picBG.ScaleHeight

    picBG.Width = wW + PicWc
    picBG.Height = wH + PicHc
    picA(0).Left = picBG.Width * (412 / 1024)
    picA(0).Top = picBG.Height * (284 / 768)
    picA(0).Width = picBG.Width * (200 / 1024)
    picA(0).Height = picBG.Height * (200 / 768)
    picA(1).Width = picA(0).Width
    picA(1).Height = picA(0).Height
End Sub

Private Sub NewWBAH()

    Set aWBH = New WBH
    aWBH.AddNew
    
    Call ReadWBH
    NeiCun_Timer '�����ڴ�
End Sub

Private Function OpenWBH(url As String) As Boolean
    Dim savefile_XML As DOMDocument
    Set savefile_XML = New DOMDocument
    savefile_XML.Load url '��ȡ�����ļ�
    If savefile_XML.documentElement Is Nothing Then
        MsgBox "�����ļ���ȡʧ��", vbCritical
        OpenWBH = False
        Exit Function
    End If
    If savefile_XML.selectSingleNode("/WBA/Config") Is Nothing Then
        MsgBox "��XML�ļ�����WBAH�����ļ�", vbCritical
        OpenWBH = False
        Exit Function
    End If
    
    Set aWBH = New WBH
    aWBH.OpenFile url
    
    Call ReadWBH
    NeiCun_Timer '�����ڴ�
    OpenWBH = True
End Function
'��WBH���ȡ��������
Private Sub ReadWBH()
    LoadFinish = False
    Dim i As Integer, AddNew As Boolean
    
    AddNew = True
    For i = 0 To comScreenScale.ListCount - 1
        If comScreenScale.list(i) = aWBH.aWidth & ChengHao & aWBH.aHeight Then
            AddNew = False
            Exit For
        End If
    Next
    If AddNew = True Then _
    comScreenScale.AddItem aWBH.aWidth & ChengHao & aWBH.aHeight
    
    comScreenScale.text = aWBH.aWidth & ChengHao & aWBH.aHeight
    txtBackground.text = aWBH.aBackground.aURL

    If aWBH.aFrame.Count >= 1 Then MaxFrame = aWBH.aFrame.Count '�������ֵ
    
    Dim URLt As String
    If IsOK(aWBH.aBackground.aURL, FileURL_Parten) Then
        URLt = aWBH.aBackground.aURL
    Else
        URLt = aWBH.aFileFolder & "\" & aWBH.aBackground.aURL
    End If
    BG_URL = URLt
    
    If aWBH.aFrame.Count > 0 Then
        ReDim Frame_URL(0 To aWBH.aFrame.Count)
        For i = 1 To aWBH.aFrame.Count
            If IsOK(aWBH.aFrame(i).aURL, FileURL_Parten) Then
                URLt = aWBH.aFrame(i).aURL
            Else
                URLt = aWBH.aFileFolder & "\" & aWBH.aFrame(i).aURL
            End If
            Frame_URL(i) = URLt
        Next
    End If
    
    picBG.Cls '���ͼ��
    picA(0).Cls
    picA(1).Cls
    
    '��ȡ�����ļ���
    If Len(aWBH.aFilePath) - Len(aWBH.aFileFolder) > 0 Then ProjectFileName = Right(aWBH.aFilePath, Len(aWBH.aFilePath) - Len(aWBH.aFileFolder) - 1) Else ProjectFileName = ""
    If ProjectFileName = "" Then ProjectFileName = "NewWBA.wba"
    
    Call dataChanged
    Call Form_Resize '�ػ�����
    
    CurrentFrame = 1 '�ص���һ֡
    LoadFinish = True
End Sub


Public Sub dataChanged()
    Me.Caption = WindowCaption & " " & aWBH.aName
End Sub

Private Sub picA_OLEDragDrop(Index As Integer, data As DataObject, effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If data.Files.Count > 0 Then '��������ļ�
    
        If data.Files(1) <> "" And Dir(data.Files(1)) <> "" Then '����ļ�����
    
           
            OpenWBH (data.Files(1))
        End If
    End If
End Sub

Private Sub picBG_OLEDragDrop(data As DataObject, effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If data.Files.Count > 0 Then '��������ļ�
    
        If data.Files(1) <> "" And Dir(data.Files(1)) <> "" Then '����ļ�����
            
            OpenWBH (data.Files(1))
        End If
    End If
End Sub

Private Sub timPlay_Timer()
    If timPlay.Enabled = True Then
        If CurrentFrame < MaxFrame Then
            CurrentFrame = CurrentFrame + 1
        Else
            CurrentFrame = NotLess(MaxFrame * (61 / 105))
        End If
    End If
    PicRedraw (1)
End Sub

Private Sub Timer_Update_Timer()
    If Timer_Update.Enabled = True Then
        Call CheckVer(UpdataURL, Auto_Update, Me)
    End If
    Auto_Update = False
    Me.Timer_Update.Enabled = False
End Sub

Private Sub txtBackground_Change()
    If LoadFinish = True Then
        aWBH.aBackground.aURL = txtBackground.text
        
        Dim URLt As String
        If IsOK(aWBH.aBackground.aURL, FileURL_Parten) Then
            URLt = aWBH.aBackground.aURL
        Else
            URLt = aWBH.aFileFolder & "\" & aWBH.aBackground.aURL
        End If
        BG_URL = URLt
        
        PicRedraw (0)
    End If
End Sub

Private Sub txtBackground_OLEDragDrop(data As DataObject, effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim extension As String
    If data.Files.Count > 0 Then '��������ļ�
        If data.Files(1) <> "" And Dir(data.Files(1)) <> "" Then '����ļ�����
            extension = Mid(data.Files(1), InStrRev(data.Files(1), ".") + 1)
            If extension = "jpg" Or extension = "jpeg" Or extension = "bmp" Or extension = "png" Or extension = "gif" Then
                If Len(aWBH.aFileFolder) > 0 And InStr(1, data.Files(1), aWBH.aFileFolder, vbTextCompare) = 1 Then
                    txtBackground.text = Right(data.Files(1), Len(data.Files(1)) - Len(aWBH.aFileFolder) - 1)
                Else
                    txtBackground.text = data.Files(1)
                End If
            End If
        End If
    End If
End Sub
