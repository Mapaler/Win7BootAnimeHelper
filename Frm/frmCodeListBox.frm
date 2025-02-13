VERSION 5.00
Begin VB.Form frmColorBox 
   AutoRedraw      =   -1  'True
   Caption         =   "颜色转换"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4815
   Icon            =   "frmCodeListBox.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   208
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picColorShow16Bit 
      Height          =   975
      Left            =   2520
      ScaleHeight     =   915
      ScaleWidth      =   2115
      TabIndex        =   9
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmd16to24 
      Caption         =   "←转换为24位色"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton cmd24to16 
      Caption         =   "转换为16位色→"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2175
   End
   Begin VB.PictureBox picColorShow24Bit 
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   2115
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txt16Bit 
      Height          =   375
      Left            =   2520
      MaxLength       =   6
      TabIndex        =   2
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox txt24Bit 
      Height          =   375
      Left            =   120
      MaxLength       =   8
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command_Ok 
      Cancel          =   -1  'True
      Caption         =   " 关闭"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lbl16ColorShow 
      BackStyle       =   0  'Transparent
      Caption         =   "16位色颜色预览"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lbl24ColorShow 
      BackStyle       =   0  'Transparent
      Caption         =   "24位色颜色预览"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lbl16Bit 
      BackStyle       =   0  'Transparent
      Caption         =   "16位色"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lbl24Bit 
      BackStyle       =   0  'Transparent
      Caption         =   "24位色"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmColorBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command_Ok_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call ReDrawWindow_FromXML(Me) '重画窗体
    Me.Icon = frmEditer.Icon
End Sub

Private Sub Form_Resize()
    Dim meW As Long, meH As Long
    Dim Border As Integer
    Border = 5
    meW = Me.ScaleWidth
    meH = Me.ScaleHeight
    '关闭按钮
    Command_Ok.Left = meW - Command_Ok.Width - Border
    Command_Ok.Top = meH - Command_Ok.Height - Border
    '24位宽
    lbl24Bit.Width = NotLess((meW / 2 - lbl24Bit.Left * 2), 0)
    txt24Bit.Width = lbl24Bit.Width
    cmd24to16.Width = lbl24Bit.Width
    lbl24ColorShow.Width = lbl24Bit.Width
    picColorShow24Bit.Width = lbl24Bit.Width
    '16位宽
    lbl16Bit.Width = lbl24Bit.Width
    txt16Bit.Width = lbl16Bit.Width
    cmd16to24.Width = lbl16Bit.Width
    lbl16ColorShow.Width = lbl16Bit.Width
    picColorShow16Bit.Width = lbl16Bit.Width
    '16位左
    lbl16Bit.Left = lbl24Bit.Left * 3 + lbl24Bit.Width
    txt16Bit.Left = txt24Bit.Left * 3 + txt24Bit.Width
    cmd16to24.Left = cmd24to16.Left * 3 + cmd24to16.Width
    lbl16ColorShow.Left = lbl24ColorShow.Left * 3 + lbl24ColorShow.Width
    picColorShow16Bit.Left = picColorShow24Bit.Left * 3 + picColorShow24Bit.Width
    '颜色预览
    picColorShow24Bit.Height = NotLess(Command_Ok.Top - Border - picColorShow24Bit.Top, 25)
    picColorShow16Bit.Height = NotLess(Command_Ok.Top - Border - picColorShow16Bit.Top, 25)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWindowToXML(Me)
End Sub


Private Sub cmd24to16_Click()
    Dim Color24b As Long, Color16b As Long
    Color24b = RGB_To_BGR(x16_to_x10(txt24Bit))
    Color16b = xGBR_to_x16b(Color24b)
    txt16Bit.text = "0x" & x10_to_x16(Color16b, 4)
    picColorShow24Bit.BackColor = Color24b
    picColorShow16Bit.BackColor = x16b_to_xGBR(Color16b)
End Sub
Private Sub txt24Bit_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("F")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("f")) Or KeyAscii = Asc("X") Or KeyAscii = Asc("x") Or KeyAscii = 8 Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub cmd16to24_Click()
    Dim Color24b As Long, Color16b As Long
    Color16b = x16_to_x10(txt16Bit)
    Color24b = x16b_to_xGBR(Color16b)
    txt24Bit.text = "0x" & x10_to_x16(RGB_To_BGR(Color24b), 6)
    picColorShow24Bit.BackColor = Color24b
    picColorShow16Bit.BackColor = x16b_to_xGBR(Color16b)
End Sub
Private Sub txt16Bit_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("F")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("f")) Or KeyAscii = Asc("X") Or KeyAscii = Asc("x") Or KeyAscii = 8 Or KeyAscii = 1 Then
    Else
        KeyAscii = 0
    End If
End Sub
