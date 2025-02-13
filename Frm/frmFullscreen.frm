VERSION 5.00
Begin VB.Form frmFullscreen 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   582
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer timPlay 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   7200
      Top             =   240
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   375
      Left            =   -1500
      TabIndex        =   1
      Top             =   -1500
      Width           =   495
   End
   Begin VB.PictureBox picBG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5085
      Left            =   240
      ScaleHeight     =   339
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   441
      TabIndex        =   0
      Top             =   240
      Width           =   6615
      Begin VB.PictureBox picA 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3000
         Index           =   0
         Left            =   2280
         ScaleHeight     =   200
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   200
         TabIndex        =   4
         Top             =   1920
         Width           =   3000
      End
      Begin VB.PictureBox picA 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3000
         Index           =   1
         Left            =   720
         ScaleHeight     =   200
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   200
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "信息"
         ForeColor       =   &H8000000E&
         Height          =   3750
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   3750
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "信息"
         ForeColor       =   &H80000013&
         Height          =   3750
         Index           =   1
         Left            =   135
         TabIndex        =   5
         Top             =   135
         Width           =   3750
      End
   End
End
Attribute VB_Name = "frmFullscreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nFrame As Integer
Private SizeTime_X As Double, SizeTime_Y As Double

Private Sub cmdExit_Click()
    timPlay.Enabled = False
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If timPlay.Enabled = False Then
            timPlay.Enabled = True
        Else
            timPlay.Enabled = False
        End If
    End If
    If KeyCode = vbKeyLeft Then
        If nFrame > 1 Then nFrame = nFrame - 1
        Call PicRedraw(1)
        Call Info
    End If
    If KeyCode = vbKeyRight Then
        If nFrame < aWBH.aFrame.Count Then nFrame = nFrame + 1
        Call PicRedraw(1)
        Call Info
    End If
    If KeyCode = vbKeyHome Then
        nFrame = 1
        Call PicRedraw(1)
        Call Info
    End If
    If KeyCode = vbKeyEnd Then
        nFrame = aWBH.aFrame.Count
        Call PicRedraw(1)
        Call Info
    End If
    If KeyCode = vbKeyEscape Then
        Call cmdExit_Click '退出
    End If
End Sub

Private Sub Form_Load()
    Call ReDrawWindow_FromXML(Me) '重画窗体
    If nFrame < 1 Then nFrame = 1
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    Call Info
    Call meResize
End Sub

Private Sub Form_Resize()
    Call meResize
End Sub

Private Sub meResize()
    Dim mW As Long, mH As Long
    mW = Me.ScaleWidth
    mH = Me.ScaleHeight
    Call PicResize(mW, mH)
End Sub

Public Sub PicResize(wW As Long, wH As Long)
    picBG.Width = wW
    picBG.Left = 0
    picBG.Height = wH
    picBG.Top = 0
    picA(0).Left = picBG.Width * (412 / 1024)
    picA(0).Top = picBG.Height * (284 / 768)
    picA(0).Width = picBG.Width * (200 / 1024)
    picA(0).Height = picBG.Height * (200 / 768)
    picA(1).Width = picA(0).Width
    picA(1).Height = picA(0).Height
    SizeTime_X = picBG.Width / aWBH.aWidth
    SizeTime_Y = picBG.Height / aWBH.aHeight
    
    PicRedraw (0)
End Sub

Private Sub Info()
    Dim Txt As String
    Txt = Txt & "Esc 退出" & vbCrLf
    Txt = Txt & "Space 播放 / 暂停" & vbCrLf
    Txt = Txt & "Left/Right  后退/前进 1 帧" & vbCrLf
    Txt = Txt & "Home/End  最前/最后帧" & vbCrLf
    Txt = Txt & "当前帧 " & nFrame
    lblInfo(0).Caption = Txt
    lblInfo(1).Caption = Txt
End Sub

Private Sub picA_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub picBG_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub

Public Sub PicRedraw(Mode As Byte)
    Select Case Mode
        Case 0
            '画图
            If IsOK(BG_URL, FileURL_Parten) And Dir(BG_URL) <> "" Then
                picBG.Cls
                Call PaintPng(BG_URL, picBG, 0, aWBH.aBackground.aX * SizeTime_X, aWBH.aBackground.aY * SizeTime_Y, , , , , , , , , SizeTime_X, SizeTime_Y)
                picBG.Refresh
                If IsOK(BG_URL, FileURL_Parten) And Dir(BG_URL) <> "" Then '生成中转图片，加快运行速度
                    picA(1).Cls
                    Call PaintPng(BG_URL, picA(1), 0, aWBH.aBackground.aX * SizeTime_X - picA(0).Left, aWBH.aBackground.aY * SizeTime_Y - picA(0).Top, , , , , , , , , SizeTime_X, SizeTime_Y)
                    picA(1).Refresh
                End If
                Call PicRedraw(1) '内画图
            End If
        Case 1
            '画图
            If IsOK(Frame_URL(nFrame), FileURL_Parten) And Dir(Frame_URL(nFrame)) <> "" Then
                'If IsOK(BG_URL, FileURL_Parten) And Dir(BG_URL) <> "" Then
                '    Call PaintPng(BG_URL, picA(0), 0, aWBH.aBackground.aX - picA(0).Left, aWBH.aBackground.aY - picA(0).Top, , , , , , , , , SizeTime)
                'End If
                picA(0).PaintPicture picA(1).Image, 0, 0
                Call PaintPng(Frame_URL(nFrame), picA(0), 0, aWBH.aFrame(nFrame).aX * SizeTime_X, aWBH.aFrame(nFrame).aY * SizeTime_Y, , , , , , , , , SizeTime_X, SizeTime_Y)
                picA(0).Refresh
            End If
    End Select
End Sub

Private Sub timPlay_Timer()
    If timPlay.Enabled = True Then
        If nFrame < aWBH.aFrame.Count Then
            nFrame = nFrame + 1
        Else
            nFrame = NotLess(aWBH.aFrame.Count * (61 / 105))
        End If
    End If
    PicRedraw (1)
    Call Info
End Sub
