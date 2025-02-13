Attribute VB_Name = "��ʾͼƬ"
Option Explicit
'*************************************************************************
'**ģ �� ����ModPaintPNG
'**˵    ������ʾPNGͼƬ��ģ��
'**�� �� �ˣ��໽е�����
'**��    �ڣ�2008��11��13��
'**��    ����V1.0
'**��    ע������GDI��ʾPNGͼƬ.PNG�����ʵ�ְ�͸��,�Ƚ�ʡ��Դ.
'**          ��ˮ�������� ��Ȩ����
'**          ����ģ��/��ģ���������վ:  http://www.m5home.com
'*************************************************************************

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type
Private Enum GpStatus
    Ok = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
End Enum
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As GpStatus
Private Declare Function GdiplusShutdown Lib "gdiplus" (ByVal token As Long) As GpStatus
Private Declare Function GdipDrawImage Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal x As Single, ByVal y As Single) As GpStatus
Private Declare Function GdipDrawImageRect Lib "gdiplus" (ByVal Graphics As Long, ByVal Image As Long, ByVal x As Single, ByVal y As Single, ByVal Width As Single, ByVal Height As Single) As GpStatus
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hDC As Long, Graphics As Long) As GpStatus
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As GpStatus
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As String, Image As Long) As GpStatus
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As GpStatus
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As GpStatus
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As GpStatus

Dim gdip_Token As Long, gdip_pngImage As Long, gdip_Graphics As Long, Picname As String

Public Enum PaintStyle
    pDrawNormal = 0
    pDrawAllBox = 1
    pDrawInbox = 2
    pSupportBox = 3
    pShowPart = 4
End Enum
Public Enum My_xAlign
    mxCenter = 0
    mxLeft = 1
    mxRight = 2
End Enum
Public Enum My_yAlign
    myCenter = 0
    myTop = 1
    myButtom = 2
End Enum

Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

'��ʾͼƬ��ԭʼ����
Public Sub PaintPng_AllSetting(ByVal sFilename As String, ByVal mhdc As Long, Optional ByVal mX As Long = 0, Optional ByVal mY As Long = 0, Optional ByVal lngWidth As Long = 0, Optional ByVal lngHeight As Long = 0)
'��ʾPNGͼƬ��ָ����DC����
    Call GDI_Initialize
    
    If GdipCreateFromHDC(mhdc, gdip_Graphics) <> Ok Then
        GdiplusShutdown gdip_Token
    Else
        Call GdipLoadImageFromFile(StrConv(GetShortName(sFilename), vbUnicode), gdip_pngImage)
        If lngWidth = 0 Then
            Call GdipGetImageWidth(gdip_pngImage, lngWidth)
        End If
        If lngHeight = 0 Then
            Call GdipGetImageHeight(gdip_pngImage, lngHeight)
        End If
        Call GdipDrawImageRect(gdip_Graphics, gdip_pngImage, mX, mY, lngWidth, lngHeight)
    End If
    
    Call GDI_Terminate
End Sub
'��ʾͼƬ����ǿ����
Public Sub PaintPng(ByVal sFilename As String, ByRef mPicBox, Optional ByVal mMode As PaintStyle = pDrawNormal, _
Optional ByVal mX As Long = 0, Optional ByVal mY As Long = 0, Optional ByVal mWidth As Long = 0, Optional ByVal mHeight As Long = 0, _
Optional ByVal xAlign As My_xAlign = mxCenter, Optional ByVal yAlign As My_yAlign = myCenter, _
Optional ByVal mTop As Long = 0, Optional ByVal mRight As Long = 0, Optional ByVal mButtom As Long = 0, Optional ByVal mLeft As Long = 0, _
Optional ByVal mSizeTime_X As Double = 1, Optional ByVal mSizeTime_Y As Double = 1)
'mModeΪ��ʾģʽ
    Dim phdc As Long
    phdc = mPicBox.hDC

'���ԭͼ����
    Dim lngHeight As Long, lngWidth As Long
    Call GDI_Initialize
    
    If GdipCreateFromHDC(phdc, gdip_Graphics) <> Ok Then
        GdiplusShutdown gdip_Token
    Else
        Call GdipLoadImageFromFile(StrConv(GetShortName(sFilename), vbUnicode), gdip_pngImage)
        Call GdipGetImageHeight(gdip_pngImage, lngHeight) '���ԭͼ����
        Call GdipGetImageWidth(gdip_pngImage, lngWidth)
    End If
    
    Call GDI_Terminate
'��û�������
    Dim boxWidth As Long, boxHeight As Long
    boxWidth = mPicBox.Width
    boxHeight = mPicBox.Height
'���������ͼƬ�ĳ������
    Dim mWidth_Scale As Single, mHeight_Scale As Single
    '��ֹ����������Ϊ��
    If lngWidth = 0 Then
        lngWidth = 1
        Debug.Print "ͼƬ���Ϊ0"
    End If
    If lngHeight = 0 Then
        lngHeight = 1
        Debug.Print "ͼƬ�߶�Ϊ0"
    End If
    
    mWidth_Scale = boxWidth / lngWidth
    mHeight_Scale = boxHeight / lngHeight

    Dim showHeight As Long, showWidth As Long

'��ʼѡ��ģʽ
    Select Case mMode
        Case pDrawNormal '���������ʾ����
            If (mSizeTime_X <> 1 Or mSizeTime_Y <> 1) And (mWidth = 0 And mHeight = 0) Then
                Call PaintPng_AllSetting(sFilename, phdc, mX, mY, Fixb(lngWidth * mSizeTime_X), Fixb(lngHeight * mSizeTime_Y))
            Else
                Call PaintPng_AllSetting(sFilename, phdc, mX, mY, mWidth, mHeight)
            End If
        Case pDrawAllBox '��������Box
            showHeight = boxHeight - mTop - mButtom
            showWidth = boxWidth - mLeft - mRight
            Call PaintPng_AllSetting(sFilename, phdc, mX + mLeft, mY + mTop, mWidth + showWidth, mHeight + showHeight)
        Case pDrawInbox 'ֻ��Box����ʾ
            '�������û�������
            boxWidth = mPicBox.Width - mLeft - mRight
            boxHeight = mPicBox.Height - mTop - mButtom
            
            If lngWidth = 0 Then lngWidth = boxWidth '�����޷���ȷ��ȡ��
            If lngHeight = 0 Then lngHeight = boxHeight '�����޷���ȷ��ȡ��
            '�������ó������
            mWidth_Scale = boxWidth / lngWidth
            mHeight_Scale = boxHeight / lngHeight

            If mWidth_Scale < mHeight_Scale Then
                showWidth = boxWidth
                showHeight = CLng(lngHeight * mWidth_Scale)
            Else
                showWidth = CLng(lngWidth * mHeight_Scale)
                showHeight = boxHeight
            End If
            'mWidth_Scale < mHeight_Scale�ȸ����ͼƬ
            If mWidth_Scale < mHeight_Scale Then
                mX = mLeft
                Select Case yAlign
                    Case myCenter
                        mY = (boxHeight - showHeight) / 2 + mTop
                    Case myTop
                        mY = mTop
                    Case myButtom
                        mY = boxHeight - showHeight + mTop
                End Select
            Else
                Select Case xAlign
                    Case mxCenter
                        mY = (boxWidth - showWidth) / 2 + mLeft
                    Case mxLeft
                        mX = mLeft
                    Case mxRight
                        mY = boxWidth - showWidth + mLeft
                End Select
                mY = mTop
            End If
            Call PaintPng_AllSetting(sFilename, phdc, mX, mY, showWidth, showHeight)
        Case pSupportBox '�Ŵ�Box
            If VarType(mPicBox) = 9 Then '����
                mPicBox.Height = (lngHeight + mTop + mButtom) * 15
                mPicBox.Width = (lngWidth + mLeft + mRight) * 15
            Else
                mPicBox.Height = lngHeight + mTop + mButtom
                mPicBox.Width = lngWidth + mLeft + mRight
            End If
            phdc = mPicBox.hDC
            Call PaintPng_AllSetting(sFilename, phdc, mX + mLeft, mY + mTop, mWidth, mHeight)
        Case pShowPart 'ֻ��ʾһ����
            If lngWidth = 0 Then lngWidth = boxWidth '�����޷���ȷ��ȡ��
            If lngHeight = 0 Then lngHeight = boxHeight '�����޷���ȷ��ȡ��
            '����Ƚ����⣬ֻ��ʾͼƬ���֣�������Ҫ���ĳ������
            mWidth_Scale = boxWidth / mWidth
            mHeight_Scale = boxHeight / mHeight
            
            If mWidth_Scale < mHeight_Scale Then
                showWidth = CLng(lngWidth * mWidth_Scale)
                showHeight = CLng(lngHeight * mWidth_Scale)
                mLeft = -(mX * mWidth_Scale)
                mRight = -(mY * mWidth_Scale)
            Else
                showWidth = CLng(lngWidth * mHeight_Scale)
                showHeight = CLng(lngHeight * mHeight_Scale)
                mLeft = -(mX * mHeight_Scale)
                mRight = -(mY * mHeight_Scale)
            End If
            Call GdipDrawImageRect(gdip_Graphics, gdip_pngImage, mLeft, mRight, showWidth, showHeight)
    End Select
End Sub

Private Sub GDI_Initialize()
    Dim GpInput As GdiplusStartupInput
    
    GpInput.GdiplusVersion = 1
    gdip_Graphics = 0
    gdip_pngImage = 0
    If GdiplusStartup(gdip_Token, GpInput) <> Ok Then
        Debug.Print "GDI��ʼʧ�ܣ�"
'        MsgBox "GDI��ʼʧ�ܣ�"
    End If
End Sub

Private Sub GDI_Terminate()
    GdipDisposeImage gdip_pngImage
    GdipDeleteGraphics gdip_Graphics
    GdiplusShutdown gdip_Token
End Sub

Private Function GetShortName(ByVal sLongFileName As String) As String
    Dim lRetVal&, sShortPathName$
    sShortPathName = Space(255)
    Call GetShortPathName(sLongFileName, sShortPathName, 255)
    If InStr(sShortPathName, Chr(0)) > 0 Then
        GetShortName = Trim(Mid(sShortPathName, 1, InStr(sShortPathName, Chr(0)) - 1))
    Else
        GetShortName = Trim(sShortPathName)
    End If
End Function


