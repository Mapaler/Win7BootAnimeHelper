VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选项"
   ClientHeight    =   3075
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   3450
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chkFormatProject 
      Caption         =   "格式化输出工程文件"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3135
   End
   Begin VB.TextBox txtFrameTime_F 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox txtFrameTime_W 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   3135
   End
   Begin VB.CheckBox chkFrameFocus 
      Caption         =   "拖动滑条，帧列表自动跳转到帧焦点"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label lblFrameTime_F 
      BackStyle       =   0  'Transparent
      Caption         =   "全屏动画每帧时间 (毫秒)"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label lblFrameTime_W 
      BackStyle       =   0  'Transparent
      Caption         =   "窗口动画每帧时间 (毫秒)"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3255
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If IsOK(txtFrameTime_W.text, "^[0-9]{0,5}$") Then
        If txtFrameTime_W.text > 65535 Or txtFrameTime_W.text < 0 Then
            MsgBox "只允许输入0-65535", 48, "窗体动画每帧时间输入错误"
            Exit Sub
        Else
            frmEditer.timPlay.Interval = txtFrameTime_W.text
        End If
    Else
        MsgBox "只允许输入整数", 48, "窗体动画每帧时间输入错误"
        Exit Sub
    End If
    
    If IsOK(txtFrameTime_F.text, "^[0-9]{0,5}$") Then
        If txtFrameTime_F.text > 65535 Or txtFrameTime_F.text < 0 Then
            MsgBox "只允许输入0-65535", 48, "全屏动画每帧时间输入错误"
            Exit Sub
        Else
            frmFullscreen.timPlay.Interval = txtFrameTime_F.text
        End If
    Else
        MsgBox "只允许输入整数", 48, "全屏动画每帧时间输入错误"
        Exit Sub
    End If
    
    
    On Error Resume Next '出错则执行下一句
    Call SaveAttribute(cfgWindow_XML, cfgWindows_XMLnode & "/" & frmEditer.name & "/" & frmEditer.timPlay.name, "Interval", frmEditer.timPlay.Interval)
    Call SaveAttribute(cfgWindow_XML, cfgWindows_XMLnode & "/" & frmFullscreen.name & "/" & frmFullscreen.timPlay.name, "Interval", frmFullscreen.timPlay.Interval)
    
    FrameFocus = chkFrameFocus.value
    FormatProject = chkFormatProject.value
    
    Unload Me
End Sub

Private Sub Form_Load()
    '置中窗体
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
    If FormatProject = True Then
        chkFormatProject.value = 1
    Else
        chkFormatProject.value = 0
    End If
    
    If FrameFocus = True Then
        chkFrameFocus.value = 1
    Else
        chkFrameFocus.value = 0
    End If
    
    txtFrameTime_W.text = frmEditer.timPlay.Interval
    txtFrameTime_F.text = frmFullscreen.timPlay.Interval
End Sub

Private Sub txtFrameTime_W_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtFrameTime_F_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
