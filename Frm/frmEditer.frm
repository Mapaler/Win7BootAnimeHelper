VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEditer 
   AutoRedraw      =   -1  'True
   Caption         =   "Win7����������������"
   ClientHeight    =   6735
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14310
   Icon            =   "frmEditer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   954
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.Slider sldFrame 
      Height          =   495
      Left            =   5400
      TabIndex        =   23
      Top             =   5760
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   873
      _Version        =   393216
      Min             =   1
      Max             =   105
      SelStart        =   1
      TickStyle       =   1
      Value           =   1
   End
   Begin VB.Timer timPlay 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   2760
      Top             =   0
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      Height          =   2175
      Left            =   120
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   15
      Top             =   2160
      Width           =   4815
      Begin VB.PictureBox picFrame_In 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   0
         ScaleHeight     =   129
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   289
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   4335
         Begin VB.TextBox txtFrame 
            Height          =   375
            Index           =   0
            Left            =   360
            OLEDropMode     =   1  'Manual
            TabIndex        =   20
            Text            =   "frame/0.png"
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton cmdFrame_Browse 
            Caption         =   "���"
            Height          =   375
            Index           =   0
            Left            =   1920
            OLEDropMode     =   1  'Manual
            TabIndex        =   19
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.CommandButton cmdFrame_Setting 
            Caption         =   "����"
            Height          =   375
            Index           =   0
            Left            =   3000
            TabIndex        =   18
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblFrame 
            BackStyle       =   0  'Transparent
            Caption         =   "000"
            Height          =   225
            Index           =   0
            Left            =   45
            TabIndex        =   21
            Top             =   180
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Shape shaFrame 
            BackColor       =   &H00FFC0C0&
            BackStyle       =   1  'Opaque
            Height          =   510
            Left            =   330
            Top             =   60
            Width           =   3690
         End
      End
      Begin VB.VScrollBar vscFrame 
         Height          =   1815
         Left            =   4440
         TabIndex        =   16
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.TextBox txtBackground 
      Height          =   375
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdBackground_Browse 
      Caption         =   "���"
      Height          =   375
      Left            =   2160
      OLEDropMode     =   1  'Manual
      TabIndex        =   10
      Top             =   1680
      Width           =   975
   End
   Begin VB.ComboBox comScreenScale 
      Height          =   300
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton cmdBackground_Setting 
      Caption         =   "����"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtProjectName 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton cmdFullscreen 
      Caption         =   "ȫ��"
      Height          =   375
      Left            =   12840
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "����"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5520
      Width           =   5055
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "����"
      Height          =   375
      Left            =   12840
      TabIndex        =   2
      Top             =   5520
      Width           =   1335
   End
   Begin VB.PictureBox picBG 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   5085
      Left            =   5880
      OLEDropMode     =   1  'Manual
      ScaleHeight     =   335
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   437
      TabIndex        =   0
      Top             =   0
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
         TabIndex        =   22
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
         TabIndex        =   1
         Top             =   1080
         Width           =   3000
      End
   End
   Begin VB.Timer Timer_Update 
      Left            =   1440
      Top             =   0
   End
   Begin VB.Label lblProjectName 
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lblBackground 
      BackStyle       =   0  'Transparent
      Caption         =   "����ͼƬ"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblScreenScale 
      BackStyle       =   0  'Transparent
      Caption         =   "��Ļ����"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label lblLoop 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      Caption         =   "ѭ����"
      Height          =   255
      Left            =   9600
      TabIndex        =   5
      Top             =   5400
      Width           =   3015
   End
   Begin VB.Label lblStraight 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "ֱ�ӵ�"
      Height          =   255
      Left            =   5520
      TabIndex        =   4
      Top             =   5400
      Width           =   4095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuNew 
         Caption         =   "�½�����(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "�򿪹���(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "���湤��(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveNew 
         Caption         =   "��湤��Ϊ(&A)"
      End
      Begin VB.Menu mnuBack 
         Caption         =   "��ԭ(&V)"
         Shortcut        =   ^B
      End
      Begin VB.Menu heng1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "��������(&E)"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuExAniFrame 
         Caption         =   "��������֡����(&F)"
         Shortcut        =   ^L
      End
      Begin VB.Menu heng2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImportLongPic 
         Caption         =   "���볤ͼ����"
      End
      Begin VB.Menu mnuExportLongPic 
         Caption         =   "������ͼ����"
      End
      Begin VB.Menu heng3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�(&E)"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "ѡ��(&O)"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuAbout 
         Caption         =   "����(&A)"
      End
      Begin VB.Menu mnuCheckUpdates 
         Caption         =   "������(&N)"
      End
   End
End
Attribute VB_Name = "frmEditer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit
Private SizeTime_X As Double, SizeTime_Y As Double
Private EditMode As Byte
Private ListDistance_True As Long, LinkListNumTemp_U As Long, LinkListNumTemp_L As Long
Private WindowCaption As String '�洢����ԭʼ����
Private ProjectFileName As String '�洢�����ļ���

Private LoadFinish As Boolean '�����Ƿ�������
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
            If extension = "jpg" Or extension = "jpeg" Or extension = "bmp" Or extension = "png" Or extension = "gif" Or extension = "tif" Or extension = "tiff" Then
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
    Call mnuExport_Click '����
End Sub

Private Sub cmdFrame_Browse_Click(Index As Integer)
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
        txtFrame(Index).text = Right(url_t, Len(url_t) - Len(aWBH.aFileFolder) - 1)
    Else
        txtFrame(Index).text = url_t
    End If
End Sub

Private Sub cmdFrame_Browse_OLEDragDrop(Index As Integer, data As DataObject, effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim extension As String
    If data.Files.Count > 0 Then '��������ļ�
        If data.Files(1) <> "" And Dir(data.Files(1)) <> "" Then '����ļ�����
            extension = Mid(data.Files(1), InStrRev(data.Files(1), ".") + 1)
            If extension = "jpg" Or extension = "jpeg" Or extension = "bmp" Or extension = "png" Or extension = "gif" Or extension = "tif" Or extension = "tiff" Then
                If Len(aWBH.aFileFolder) > 0 And InStr(1, data.Files(1), aWBH.aFileFolder, vbTextCompare) = 1 Then
                    txtFrame(Index).text = Right(data.Files(1), Len(data.Files(1)) - Len(aWBH.aFileFolder) - 1)
                Else
                    txtFrame(Index).text = data.Files(1)
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdFrame_Setting_Click(Index As Integer)
    frmSetting.PicIndex = Index
    frmSetting.Show 1
End Sub

Private Sub cmdFullscreen_Click()
    If timPlay.Enabled = True Then
        Call cmdPlay_Click
        frmFullscreen.timPlay.Enabled = True
    End If
    frmFullscreen.nFrame = sldFrame.value
    frmFullscreen.Show 1
End Sub

'���֡�б�����
Private Sub NewList_AddList(ByVal CopyNum As Long)
    Dim i As Long
    picFrame_In.Visible = True
    For i = 1 To CopyNum
        Load lblFrame(i)
        Load txtFrame(i)
        Load cmdFrame_Browse(i)
        Load cmdFrame_Setting(i)
    Next
    Call NewList_BoxResize
    Call Form_Resize
End Sub
'ɾ��֡�б�����
Private Sub NewList_DelList()
    Dim i As Long
    
    picFrame_In.Visible = False
    For i = 1 To lblFrame.UBound
        Unload lblFrame(i)
        Unload txtFrame(i)
        Unload cmdFrame_Browse(i)
        Unload cmdFrame_Setting(i)
    Next
    Call NewList_BoxResize
    Call Form_Resize
End Sub
'����б��������Ų��б�
Private Sub NewList_BoxResize()
    Dim i As Long
    Dim mScaleWidth As Long
    Dim Border As Integer
    Border = 5
    
    picFrame_In.Width = NotLess(picFrame.ScaleWidth - vscFrame.Width - picFrame_In.Left * 2)
    mScaleWidth = picFrame_In.ScaleWidth
    For i = 1 To lblFrame.UBound
        cmdFrame_Setting(i).Left = mScaleWidth - cmdFrame_Setting(i).Width - Border
        cmdFrame_Browse(i).Left = cmdFrame_Setting(i).Left - cmdFrame_Browse(i).Width - Border
        txtFrame(i).Width = NotLess(cmdFrame_Browse(i).Left - txtFrame(i).Left - Border)
        shaFrame.Width = NotLess(mScaleWidth - shaFrame.Left - 2)
        'λ��  ��Ϊ���������Ϳ����ˣ����ﲻ����
        lblFrame(i).Top = lblFrame(0).Top + ListDistance_True * (i - 1)
        txtFrame(i).Top = txtFrame(0).Top + ListDistance_True * (i - 1)
        cmdFrame_Browse(i).Top = cmdFrame_Browse(0).Top + ListDistance_True * (i - 1)
        cmdFrame_Setting(i).Top = cmdFrame_Setting(0).Top + ListDistance_True * (i - 1)
    Next
    picFrame_In.Height = ListDistance_True * (lblFrame.UBound) + lblFrame(0).Top
    If picFrame_In.Height <= picFrame.ScaleHeight Then
        picFrame_In.Height = picFrame.ScaleHeight
        vscFrame.Enabled = False
    Else
        vscFrame.Enabled = True
        vscFrame.LargeChange = vscFrame.Max / (NotLess(picFrame_In.Height - picFrame.ScaleHeight, 0) / 1000 + 1)
        vscFrame.SmallChange = vscFrame.Max / (NotLess(picFrame_In.Height - picFrame.ScaleHeight, 0) / 100 + 10)
    End If
    Call vscFrame_Change
End Sub

Private Sub cmdPlay_Click()
    If timPlay.Enabled = False Then
        timPlay.Enabled = True
        cmdPlay.Caption = "ֹͣ"
    Else
        timPlay.Enabled = False
        cmdPlay.Caption = "��ʼ"
    End If
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
        Call dataChanged
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
    If IsDebugMode Then '���Ի���
    End If
    EditMode = 1
    Auto_Update = True
    
    Timer_Update.Interval = 1000
    Timer_Update.Enabled = True
    
    '��ȡ��Ļ����
    For i = 0 To GetAllNode_Lenth(cfgConfig_XML, "/Soft/ScreenScale/Scale") - 1
        comScreenScale.AddItem GetNodeAttribute(cfgConfig_XML, "/Soft/ScreenScale/Scale", "Width", , i) _
        & ChengHao & GetNodeAttribute(cfgConfig_XML, "/Soft/ScreenScale/Scale", "Height", , i)
    Next

    Call mnuNew_Click '����½�
    
    Call ReDrawWindow_FromXML(Me) '�ػ�����
    
    WindowCaption = Me.Caption
    '��ȡ��ʵ����
    ListDistance_True = GetListDistance_True
    shaFrame.Height = LinkListNumTemp_U - LinkListNumTemp_L + 8
    
    Call dataChanged(False)
    
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

'���������б�洢��һ�е�λ�ã�����
Private Function GetListDistance_True()
    'Ϊ��ȷ���б��¿ؼ���λ��
    'ȷ����Сֵ
    LinkListNumTemp_L = lblFrame(0).Top
    If LinkListNumTemp_L > txtFrame(0).Top Then
        LinkListNumTemp_L = txtFrame(0).Top
    End If
    If LinkListNumTemp_L > cmdFrame_Browse(0).Top Then
        LinkListNumTemp_L = cmdFrame_Browse(0).Top
    End If
    If LinkListNumTemp_L > cmdFrame_Setting(0).Top Then
        LinkListNumTemp_L = cmdFrame_Setting(0).Top
    End If
    'ȷ�����ֵ
    LinkListNumTemp_U = lblFrame(0).Top + lblFrame(0).Height
    If LinkListNumTemp_U < txtFrame(0).Top + txtFrame(0).Height Then
        LinkListNumTemp_U = txtFrame(0).Top + txtFrame(0).Height
    End If
    If LinkListNumTemp_U < cmdFrame_Browse(0).Top + cmdFrame_Browse(0).Height Then
        LinkListNumTemp_U = cmdFrame_Browse(0).Top + cmdFrame_Browse(0).Height
    End If
    If LinkListNumTemp_U < cmdFrame_Setting(0).Top + cmdFrame_Setting(0).Height Then
        LinkListNumTemp_U = cmdFrame_Setting(0).Top + cmdFrame_Setting(0).Height
    End If
    GetListDistance_True = LinkListNumTemp_U + ListDistance - LinkListNumTemp_L
End Function

Private Sub Form_Unload(Cancel As Integer)

    If FormatProject = True Then
        Call SaveAttribute(cfgConfig_XML, cfgConfig_XMLnode, "FormatProject", 1)
    Else
        Call SaveAttribute(cfgConfig_XML, cfgConfig_XMLnode, "FormatProject", 0)
    End If
    If FrameFocus = True Then
        Call SaveAttribute(cfgConfig_XML, cfgConfig_XMLnode, "FrameFocus", 1)
    Else
        Call SaveAttribute(cfgConfig_XML, cfgConfig_XMLnode, "FrameFocus", 0)
    End If
    
    Call SaveWindowToXML(Me)
    Dim x As Integer
    If aWBH.aChanged = True Then
        x = MsgBox("����ļ��޸���û�б��棬�Ƿ񱣴�", 547, "��ʾ")
        If x = 6 Then '��
            Call mnuSave_Click
            Call EndSoft
        End If
        If x = 7 Then '��
            Call EndSoft
        End If
        If x = 2 Then 'ȡ��
            Cancel = True
        End If
    Else
        Call EndSoft
    End If
End Sub
'����ı��С
Private Sub Form_Resize()
    Dim meW As Long, meH As Long
    Dim Border As Integer, sldRealWidth As Integer
    Border = 5
    
    meW = Me.ScaleWidth
    meH = Me.ScaleHeight

    '��ť
    cmdFullscreen.Left = meW - cmdFullscreen.Width - Border
    cmdPlay.Left = meW - cmdPlay.Width - Border
    cmdFullscreen.Top = meH - cmdFullscreen.Height - Border
    cmdPlay.Top = cmdFullscreen.Top - cmdPlay.Height - Border

    lblStraight.Top = cmdPlay.Top 'sldFrame.Top - lblStraight.Height
    lblLoop.Top = cmdPlay.Top 'sldFrame.Top - lblLoop.Height
    sldFrame.Top = lblStraight.Top + lblStraight.Height + Border
    
    picBG.Height = NotLess(lblStraight.Top - picBG.Top - Border * 2)
    picBG.Width = NotLess(picBG.Height * (aWBH.aWidth / aWBH.aHeight))
    picBG.Left = meW - picBG.Width - Border * 2
    
    sldFrame.Width = NotLess(cmdPlay.Left - picBG.Left)
    sldFrame.Left = cmdPlay.Left - sldFrame.Width
    
    sldRealWidth = sldFrame.Width - 13 * 2 '������ʵ�ʳ���
    lblLoop.Width = NotLess(sldRealWidth * ((104 - 60) / 104))
    lblLoop.Left = (sldFrame.Left + sldFrame.Width) - lblLoop.Width - 13
    lblStraight.Width = NotLess(sldRealWidth * (60 / 104))
    lblStraight.Left = lblLoop.Left - lblStraight.Width
    
    
    lblProjectName.Width = NotLess(picBG.Left - lblProjectName.Left - Border)
    txtProjectName.Width = NotLess(picBG.Left - txtProjectName.Left - Border)
    lblScreenScale.Width = NotLess(picBG.Left - lblScreenScale.Left - Border)
    comScreenScale.Width = NotLess(picBG.Left - comScreenScale.Left - Border)
    lblBackground.Width = NotLess(picBG.Left - lblBackground.Left - Border)
    cmdBackground_Setting.Left = picBG.Left - cmdBackground_Setting.Width - Border
    cmdBackground_Browse.Left = cmdBackground_Setting.Left - cmdBackground_Browse.Width - Border
    txtBackground.Width = NotLess(cmdBackground_Browse.Left - txtBackground.Left - Border)
    
    cmdExport.Width = NotLess(picBG.Left - cmdExport.Left - Border)
    cmdExport.Top = meH - cmdExport.Height - Border
    picFrame.Width = NotLess(picBG.Left - picFrame.Left - Border)
    picFrame.Height = NotLess(cmdExport.Top - picFrame.Top - Border)
    vscFrame.Left = picFrame.ScaleWidth - vscFrame.Width
    vscFrame.Height = picFrame.ScaleHeight - vscFrame.Top * 2
    
    Call PicResize(picBG.Width, picBG.Height)
    SizeTime_X = picBG.Width / aWBH.aWidth
    SizeTime_Y = picBG.Height / aWBH.aHeight
    
    '��ͼ
    PicRedraw (0)
    
    Call NewList_BoxResize
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
                Call sldFrame_Change '�ڻ�ͼ
            End If
        Case 1
            '��ͼ
            If IsOK(Frame_URL(sldFrame.value), FileURL_Parten) And Dir(Frame_URL(sldFrame.value)) <> "" Then
                'If IsOK(BG_URL, FileURL_Parten) And Dir(BG_URL) <> "" Then
                '    Call PaintPng(BG_URL, picA(0), 0, aWBH.aBackground.aX - picA(0).Left, aWBH.aBackground.aY - picA(0).Top, , , , , , , , , SizeTime_X)
                'End If
                picA(0).PaintPicture picA(1).Image, 0, 0
                Call PaintPng(Frame_URL(sldFrame.value), picA(0), 0, aWBH.aFrame(sldFrame.value).aX * SizeTime_X, aWBH.aFrame(sldFrame.value).aY * SizeTime_Y, , , , , , , , , SizeTime_X, SizeTime_Y)
                picA(0).Refresh
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

Private Sub lblFrame_Click(Index As Integer)
    sldFrame.value = Index
End Sub

Private Sub mnuAbout_Click()
    If newVer.tWebsite <> "" Then
        ShellExecute Me.hWnd, vbNullString, newVer.tWebsite, vbNullString, vbNullString, SW_SHOWNORMAL
    Else
        ShellExecute Me.hWnd, vbNullString, SoftPage, vbNullString, vbNullString, SW_SHOWNORMAL
    End If
End Sub

Private Sub mnuBack_Click()
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

Private Sub mnuCheckUpdates_Click()
    Me.Timer_Update.Enabled = True
End Sub

Private Sub mnuExit_Click()
    Call Form_Unload(1)
End Sub

Private Sub mnuExport_Click()
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
    
    Dim url_t As String
    If lResult <> 0 Then
                extension = "." & F_OpG.Item(file.nFilterIndex).Extensions
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

Private Sub mnuExAniFrame_Click()
    Dim i As Long, extension As String
    Dim file As OPENFILENAME, lResult As Long
    Dim DlgInfo As DlgFileInfo
    
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
    file.lpstrFile = aWBH.aName & "_����" & String$(255, 0)   '����Ĭ��Ҫ�򿪵��ļ���
    file.nMaxFile = 255 '��ʾ�ļ����ĳ���
    file.lpstrFileTitle = String$(255, 0) '�ļ���(��·��)
    file.nMaxFileTitle = 255 '�򿪶Ի���ı���ĳ���
    file.lpstrInitialDir = String$(255, 0)  '���ó�ʼ·��
    file.lpstrFilter = F_OpG.Txt
    file.nFilterIndex = 0
    file.lpstrTitle = "�����붯������ǰ׺"
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
'        Case 4
'            extension = ".tif"
'    End Select
    
    Dim url_t As String
    If lResult <> 0 Then
        extension = "." & F_OpG.Item(file.nFilterIndex).Extensions
        DlgInfo = GetDlgFileInfo(file.lpstrFile)
        url_t = DlgInfo.sPath & DlgInfo.sFile(1)
        '��ȡ����չ���ļ���
        If InStrRev(url_t, extension) > 1 Then
            url_t = Left(url_t, InStrRev(url_t, extension) - 1)
        End If
    Else
        Exit Sub
    End If
    
    Dim loopNum As Byte, TextT As String
    TextT = OnlyRegExp(InputBox("��61-105֡ѭ��������Ρ����1�Σ����255�Ρ�������������ʹ��Ĭ��ֵ1��", "������ѭ������", 2), "[\d]")
    If Len(TextT) > 0 Then
        loopNum = NotGreater(NotLess(CLng(TextT), 1), 255)
    Else
        loopNum = 1
    End If
    
    Dim bgAlpha As Boolean, LongT As String
    LongT = MsgBox("�Ƿ񱣳�ͼƬ͸��ͨ��", vbYesNoCancel Or vbQuestion)
    If LongT = vbCancel Then
        Exit Sub
    ElseIf LongT = vbYes Then
        bgAlpha = True
    ElseIf LongT = vbNo Then
        bgAlpha = False
    End If
    
    '�м䶯��BOX��С��λ
    Dim boxWidth As Integer, boxHeight As Integer, boLeft As Integer, boxTop As Integer
    boxWidth = aWBH.aWidth * (200 / 1024)
    boxHeight = aWBH.aHeight * (200 / 768)
    boLeft = aWBH.aWidth * (412 / 1024)
    boxTop = aWBH.aHeight * (284 / 768)

    '֡������
    Dim Fnum_once As Integer, Fnum_loop As Integer, Fnum_long As Byte
    Fnum_once = aWBH.aFrame.Count * (60 / 105)
    Fnum_loop = aWBH.aFrame.Count * (45 / 105)
    Fnum_long = Len(CStr(Fnum_once + Fnum_loop * loopNum))
    
    Dim Bitmap_BGt As Long, Bitmap_Fout As Long, Bitmap_Ft As Long, Graphics As Long
    Dim bmW_BG As Long, bmH_BG As Long, bmW_F As Long, bmH_F As Long
    
    InitGDIPlus
    
    '���������͸�����ȡ����
    If bgAlpha = False Then
        '���ļ�����Bitmap
        GdipCreateBitmapFromFile StrPtr(BG_URL), Bitmap_BGt
        GdipGetImageWidth Bitmap_BGt, bmW_BG
        GdipGetImageHeight Bitmap_BGt, bmH_BG
    End If
    
'    Dim a As Long
'    a = Fnum_once + Fnum_loop * loopNum
    Dim i0 As Integer
    For i = 1 To Fnum_once + Fnum_loop * loopNum '��ѭ��֡
        'i0����ʵ֡�������ظ���ѭ����������Ӧ��֡��
        If i <= Fnum_once + Fnum_loop Then
            i0 = i
        Else
            i0 = i - (Fixb((i - Fnum_once) / Fnum_loop) - 1) * Fnum_loop
        End If
        
        GdipCreateBitmapFromFile StrPtr(Frame_URL(i0)), Bitmap_Ft
        GdipGetImageWidth Bitmap_Ft, bmW_F
        GdipGetImageHeight Bitmap_Ft, bmH_F

        CreateBitmapWithGraphics Bitmap_Fout, Graphics, boxWidth, boxHeight '�ؼ�������һ��Image��Graphics����

        If bgAlpha = False Then
            GdipDrawImageRectI Graphics, Bitmap_BGt, aWBH.aBackground.aX - boLeft, aWBH.aBackground.aY - boxTop, bmW_BG, bmH_BG
        End If

        GdipDrawImageRectI Graphics, Bitmap_Ft, aWBH.aFrame(i0).aX, aWBH.aFrame(i0).aY, bmW_F, bmH_F

        '���������ڻ���Bitmap��
        Select Case extension
            Case ".png"
                SaveImageToPNG Bitmap_Fout, url_t & format(i, String(Fnum_long, "0")) & extension
            Case ".bmp"
                SaveImageToBMP Bitmap_Fout, url_t & format(i, String(Fnum_long, "0")) & extension
            Case ".jpg"
                SaveImageToJPG Bitmap_Fout, url_t & format(i, String(Fnum_long, "0")) & extension
            Case ".gif"
                SaveImageToGIF Bitmap_Fout, url_t & format(i, String(Fnum_long, "0")) & extension
            Case ".tif"
                SaveImageToTIF Bitmap_Fout, url_t & format(i, String(Fnum_long, "0")) & extension
        End Select
    Next


    'ɨ�ع���
    GdipDeleteGraphics Graphics
    
    GdipDisposeImage Bitmap_BGt
    GdipDisposeImage Bitmap_Ft
    GdipDisposeImage Bitmap_Fout
    
    TerminateGDIPlus
    
    MsgBox "֡�����ѵ����������ʹ��VirtualDub��Ulead GIF Animator���������ת��ΪGIF��", 64, "�������"
End Sub

Private Sub mnuExportLongPic_Click()
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

Private Sub mnuImportLongPic_Click()
    Dim file As OPENFILENAME, lResult As Long
    Dim DlgInfo As DlgFileInfo
    
    Dim x As Integer
    If aWBH.aChanged = True Then
        x = MsgBox("����ļ��޸���û�б��棬�Ƿ񱣴�", 547, "��ʾ")
        If x = 6 Then '��
            Call mnuSave_Click
        End If
        If x = 7 Then '��
        End If
        If x = 2 Then 'ȡ��
            Exit Sub
        End If
    End If
    
    file.lStructSize = Len(file)
    file.hwndOwner = Me.hWnd
    file.flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
    file.lpstrFile = String$(32767, 0)   '����Ĭ��Ҫ�򿪵��ļ���
    file.nMaxFile = 255 '��ʾ�ļ����ĳ���
    file.lpstrFileTitle = String$(255, 0) '�򿪶Ի���ı���
    file.nMaxFileTitle = 255 '�򿪶Ի���ı���ĳ���
    file.lpstrInitialDir = String$(255, 0)  '���ó�ʼ·��
    file.lpstrFilter = F_IpG.Txt '"�ļ�" & Chr$(0) & "*.*" '�򿪵��ļ�����"
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
    
    Dim filepath As String
    If Len(aWBH.aFileFolder) > 0 And InStr(url_t, aWBH.aFileFolder) = 1 Then
        filepath = Right(url_t, Len(url_t) - Len(aWBH.aFileFolder) - 1)
    Else
        filepath = url_t
    End If
    
    
    Dim OffsetXs As String, offsetX As Integer, OffsetXdefault As Integer
    OffsetXdefault = ((aWBH.aWidth / 1024) * 200 - 200) / 2
    OffsetXs = InputBox("����������Xƫ�ƣ�����������", "���붯����ͼ", OffsetXdefault)
    If IsOK(OffsetXs, "^-?\d+$") Then
        offsetX = CInt(OffsetXs)
    Else
        '�����ȡ������ʾ
        If StrPtr(OffsetXs) <> 0 Then MsgBox "ƫ�Ʋ��������Ϲ���"
        Exit Sub
    End If
    
    Dim i As Integer
    For i = 1 To (lblFrame.Count - 1)
        txtFrame(i).text = filepath
        aWBH.aFrame(i).aX = offsetX
        aWBH.aFrame(i).aY = (i - 1) * -200
    Next
End Sub

Private Sub mnuNew_Click()
    Dim i As Integer
    
    Dim x As Integer
    If Not aWBH Is Nothing Then
        If aWBH.aChanged = True Then
            x = MsgBox("����ļ��޸���û�б��棬�Ƿ񱣴�", 547, "��ʾ")
            If x = vbYes Then '��
                Call mnuSave_Click
            End If
            If x = vbNo Then '��
            End If
            If x = vbCancel Then 'ȡ��
                Exit Sub
            End If
        End If
    End If
    
    Set aWBH = New WBH
    aWBH.AddNew
    
    Call ReadWBH
    NeiCun_Timer '�����ڴ�
End Sub

Private Sub mnuOpen_Click()
    Dim file As OPENFILENAME, lResult As Long
    Dim DlgInfo As DlgFileInfo
    
    Dim x As Integer
    If aWBH.aChanged = True Then
        x = MsgBox("����ļ��޸���û�б��棬�Ƿ񱣴�", 547, "��ʾ")
        If x = vbYes Then '��
            Call mnuSave_Click
        End If
        If x = vbNo Then '��
        End If
        If x = vbCancel Then 'ȡ��
            Exit Sub
        End If
    End If
    
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
    
    txtProjectName.text = aWBH.aName
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
    
    Call NewList_DelList
    Call NewList_AddList(aWBH.aFrame.Count)
    For i = 1 To aWBH.aFrame.Count
        lblFrame(i).Caption = format(i, "000")
        txtFrame(i).text = aWBH.aFrame(i).aURL
        lblFrame(i).Visible = True
        txtFrame(i).Visible = True
        cmdFrame_Browse(i).Visible = True
        cmdFrame_Setting(i).Visible = True
    Next i
    If aWBH.aFrame.Count >= 1 Then sldFrame.Max = aWBH.aFrame.Count '�������ֵ
    
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
    
    Call Form_Resize '�ػ�����
    vscFrame.value = 1 '�ص���һ֡
    LoadFinish = True
    Call dataChanged(False)
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show 1
End Sub

Private Sub mnuSave_Click()
    If aWBH.aFilePath <> "" Then
        aWBH.SaveFile aWBH.aFilePath
    Else
        Call mnuSaveNew_Click '������Ϊ
    End If
    dataChanged (False)
End Sub

Private Sub mnuSaveNew_Click()
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
    file.lpstrFilter = F_Proj.Txt '"SkyDrive��ҳ�����ļ�" & Chr$(0) & "*.*" '�򿪵��ļ�����"
    file.nFilterIndex = 1
    file.lpstrTitle = "ѡ�񹤳̱���λ��"
    lResult = GetSaveFileName(file) 'ȡ���ļ���
    
    Dim url_t As String
    If lResult <> 0 Then
        DlgInfo = GetDlgFileInfo(file.lpstrFile)
        url_t = DlgInfo.sPath & DlgInfo.sFile(1)
        Select Case file.nFilterIndex
            Case 2
                If Right(url_t, 4) <> ".xml" Then url_t = url_t & ".xml"
            Case Else
                If Right(url_t, 4) <> ".wba" Then url_t = url_t & ".wba"
        End Select
    Else
        Exit Sub
    End If

    Dim aWBHn As New WBH
    Set aWBHn = aWBH
    
    Dim i As Integer
    aWBHn.aFilePath = url_t
'    If aWBHn.aFileFolder <> Left(DlgInfo.sPath, Len(DlgInfo.sPath) - 1) Then
'        '�����µ�·��
'        aWBHn.aFileFolder = Left(DlgInfo.sPath, Len(DlgInfo.sPath) - 1)
'        If Not IsOK(aWBHn.aBackground.aURL, FileURL_Parten) Then
'            aWBHn.aBackground.aURL = DlgInfo.sPath & aWBHn.aBackground.aURL
'        End If
'        For i = 1 To aWBH.aFrame.Count
'            If Not IsOK(aWBH.aFrame(i).aURL, FileURL_Parten) Then
'                aWBHn.aBackground.aURL = DlgInfo.sPath & aWBHn.aFrame(i).aURL
'            End If
'        Next
'    End If
    
    aWBHn.SaveFile url_t
    
    OpenWBH (url_t)
    'Call mnuSave_Click '�������
End Sub

Public Sub dataChanged(Optional ByVal Changed As Boolean = True)
    If Changed = True Then
        Me.Caption = WindowCaption & " " & ProjectFileName & "*"
        aWBH.aChanged = True
    Else
        Me.Caption = WindowCaption & " " & ProjectFileName
        aWBH.aChanged = False
    End If
End Sub

Private Sub picA_OLEDragDrop(Index As Integer, data As DataObject, effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If data.Files.Count > 0 Then '��������ļ�
    
        If data.Files(1) <> "" And Dir(data.Files(1)) <> "" Then '����ļ�����
    
            Dim m As Integer
            If aWBH.aChanged = True Then
                m = MsgBox("����ļ��޸���û�б��棬�Ƿ񱣴�", 547, "��ʾ")
                If m = vbYes Then '��
                    Call mnuSave_Click
                End If
                If m = vbNo Then '��
                End If
                If m = vbCancel Then 'ȡ��
                    Exit Sub
                End If
            End If
            
            OpenWBH (data.Files(1))
        End If
    End If
End Sub

Private Sub picBG_OLEDragDrop(data As DataObject, effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    If data.Files.Count > 0 Then '��������ļ�
    
        If data.Files(1) <> "" And Dir(data.Files(1)) <> "" Then '����ļ�����
    
            Dim m As Integer
            If aWBH.aChanged = True Then
                m = MsgBox("����ļ��޸���û�б��棬�Ƿ񱣴�", 547, "��ʾ")
                If m = vbYes Then '��
                    Call mnuSave_Click
                End If
                If m = vbNo Then '��
                End If
                If m = vbCancel Then 'ȡ��
                    Exit Sub
                End If
            End If
            
            OpenWBH (data.Files(1))
        End If
    End If
End Sub

Private Sub sldFrame_Change()
    Dim cha As Long
'    cha = -NotLess(picFrame_In.Height - picFrame.ScaleHeight, 0) * (vscFrame.value / vscFrame.Max)
    
    shaFrame.Top = LinkListNumTemp_L - 4 + ListDistance_True * (sldFrame.value - 1)
    If FrameFocus = True Then 'Ѱ��֡����
        vscFrame.value = ((sldFrame.value - 1) / (sldFrame.Max - 1)) * vscFrame.Max
    End If
'    shaFrame.Top = LinkListNumTemp_L - 4 + ListDistance_True * (sldFrame.value - 1) + cha

    '��ͼ
    PicRedraw (1)
    
End Sub

Private Sub sldFrame_Scroll()
    Call sldFrame_Change
End Sub

Private Sub timPlay_Timer()
    If timPlay.Enabled = True Then
        If sldFrame.value < sldFrame.Max Then
            sldFrame.value = sldFrame.value + 1
        Else
            sldFrame.value = NotLess(sldFrame.Max * (61 / 105))
        End If
    End If
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
        Call dataChanged
    End If
End Sub

Private Sub txtBackground_OLEDragDrop(data As DataObject, effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim extension As String
    If data.Files.Count > 0 Then '��������ļ�
        If data.Files(1) <> "" And Dir(data.Files(1)) <> "" Then '����ļ�����
            extension = Mid(data.Files(1), InStrRev(data.Files(1), ".") + 1)
            If extension = "jpg" Or extension = "jpeg" Or extension = "bmp" Or extension = "png" Or extension = "gif" Or extension = "tif" Or extension = "tiff" Then
                If Len(aWBH.aFileFolder) > 0 And InStr(1, data.Files(1), aWBH.aFileFolder, vbTextCompare) = 1 Then
                    txtBackground.text = Right(data.Files(1), Len(data.Files(1)) - Len(aWBH.aFileFolder) - 1)
                Else
                    txtBackground.text = data.Files(1)
                End If
            End If
        End If
    End If
End Sub

Private Sub txtFrame_Change(Index As Integer)
    If LoadFinish = True Then
        aWBH.aFrame(Index).aURL = txtFrame(Index).text
        
        Dim URLt As String
    
        If IsOK(aWBH.aFrame(Index).aURL, FileURL_Parten) Then
            URLt = aWBH.aFrame(Index).aURL
        Else
            URLt = aWBH.aFileFolder & "\" & aWBH.aFrame(Index).aURL
        End If
        Frame_URL(Index) = URLt
    
        PicRedraw (1)
        Call dataChanged
    End If
End Sub

Private Sub txtFrame_OLEDragDrop(Index As Integer, data As DataObject, effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim extension As String
    If data.Files.Count > 0 Then '��������ļ�
        If data.Files(1) <> "" And Dir(data.Files(1)) <> "" Then '����ļ�����
            extension = Mid(data.Files(1), InStrRev(data.Files(1), ".") + 1)
            If extension = "jpg" Or extension = "jpeg" Or extension = "bmp" Or extension = "png" Or extension = "gif" Or extension = "tif" Or extension = "tiff" Then
                If Len(aWBH.aFileFolder) > 0 And InStr(1, data.Files(1), aWBH.aFileFolder, vbTextCompare) = 1 Then
                    txtFrame(Index).text = Right(data.Files(1), Len(data.Files(1)) - Len(aWBH.aFileFolder) - 1)
                Else
                    txtFrame(Index).text = data.Files(1)
                End If
            End If
        End If
    End If
End Sub

Private Sub txtProjectName_Change()
    If LoadFinish = True Then
        aWBH.aName = txtProjectName.text
        Call dataChanged
    End If
End Sub

Private Sub vscFrame_Change()
'������
    Dim i As Long
    Dim cha As Long
    '��ǰ��Picture_LinkList_In.Top
'    cha = -NotLess(picFrame_In.Height - picFrame.ScaleHeight, 0) * (vscFrame.value / vscFrame.Max)
'    shaFrame.Top = LinkListNumTemp_L - 4 + ListDistance_True * (sldFrame.value - 1) + cha
    
    'λ��
    picFrame_In.Top = -(picFrame_In.Height - picFrame.Height) * (vscFrame.value / vscFrame.Max)
    
'    For i = 1 To lblFrame.UBound
'        lblFrame(i).Top = lblFrame(0).Top + ListDistance_True * (i - 1) + cha
'        txtFrame(i).Top = txtFrame(0).Top + ListDistance_True * (i - 1) + cha
'        cmdFrame_Browse(i).Top = cmdFrame_Browse(0).Top + ListDistance_True * (i - 1) + cha
'        cmdFrame_Setting(i).Top = cmdFrame_Setting(0).Top + ListDistance_True * (i - 1) + cha
'    Next
End Sub

Private Sub vscFrame_Scroll()
    Call vscFrame_Change
End Sub
