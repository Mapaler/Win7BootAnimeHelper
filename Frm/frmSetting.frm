VERSION 5.00
Begin VB.Form frmSetting 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "图片设置"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   98
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   179
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtY 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblY 
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblX 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PicIndex As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If PicIndex = 0 Then
        aWBH.aBackground.aX = NotGreater(txtX.text, 30000)
        aWBH.aBackground.aY = NotGreater(txtY.text, 30000)
        frmEditer.PicRedraw (0)
    ElseIf PicIndex > 0 Then
        aWBH.aFrame(PicIndex).aX = NotGreater(txtX.text, 30000)
        aWBH.aFrame(PicIndex).aY = NotGreater(txtY.text, 30000)
        frmEditer.PicRedraw (1)
    End If
    frmEditer.dataChanged
    Unload Me
End Sub

Private Sub Form_Load()
    Call ReDrawWindow_FromXML(Me) '重画窗体
    If PicIndex = 0 Then
        txtX.text = aWBH.aBackground.aX
        txtY.text = aWBH.aBackground.aY
    ElseIf PicIndex > 0 Then
        txtX.text = aWBH.aFrame(PicIndex).aX
        txtY.text = aWBH.aFrame(PicIndex).aY
    End If
End Sub

Private Sub txtX_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtY_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub
