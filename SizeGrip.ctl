VERSION 5.00
Begin VB.UserControl SizeGrip 
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MousePointer    =   8  'Size NW SE
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ToolboxBitmap   =   "SizeGrip.ctx":0000
   Begin VB.Image imgDesignMode 
      Height          =   228
      Left            =   1116
      Picture         =   "SizeGrip.ctx":0312
      Top             =   2052
      Width           =   228
   End
   Begin VB.Image img21pix 
      Height          =   252
      Left            =   720
      Picture         =   "SizeGrip.ctx":07C8
      Top             =   2052
      Width           =   252
   End
   Begin VB.Image img17pix 
      Height          =   204
      Left            =   432
      Picture         =   "SizeGrip.ctx":0D4A
      Top             =   2052
      Width           =   204
   End
   Begin VB.Image img14pix 
      Height          =   168
      Left            =   180
      Picture         =   "SizeGrip.ctx":1100
      Top             =   2052
      Width           =   168
   End
End
Attribute VB_Name = "SizeGrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Type COLORS_RGB
    R As Long
    G As Long
    b As Long
End Type

Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, ByRef pccolorref As Long) As Long

Private Const SW_HIDE = 0
Private Const SW_SHOW = 5

Private Const WM_SYSCOMMAND = &H112&
Private Const SC_SIZE_SE  As Long = &HF008&

Private WithEvents mForm As Form
Attribute mForm.VB_VarHelpID = -1

Private mBackColor As Long
Private mWidth As Long
Private mHeight As Long
Private mSettingImage As Boolean
Private mBackColorIsFromAmbient As Boolean

Private Sub mForm_Resize()
    If mForm.WindowState = vbNormal Then
        UserControl.Width = ScaleX(mWidth, vbPixels, vbTwips)
        UserControl.Height = ScaleY(mHeight, vbPixels, vbTwips)
        PositionOnForm
        ShowWindow UserControl.hwnd, SW_SHOW
    Else
        If IsWindowVisible(UserControl.hwnd) <> 0 Then
            ShowWindow UserControl.hwnd, SW_HIDE
        End If
    End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "BackColor" Then
        If mBackColorIsFromAmbient Then
            BackColor = Ambient.BackColor
        Else
            mBackColorIsFromAmbient = (mBackColor = Ambient.BackColor)
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    mWidth = 19
    mHeight = 19
    mBackColor = -1
End Sub

Private Sub UserControl_InitProperties()
    BackColor = Ambient.BackColor
    SetGrip
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mForm Is Nothing Then Exit Sub
    ReleaseCapture
    PostMessage mForm.hwnd, WM_SYSCOMMAND, SC_SIZE_SE, 0&
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
    SetGrip
End Sub

Private Sub UserControl_Resize()
    If Not mSettingImage Then
        If Ambient.UserMode Then
            UserControl.Width = ScaleX(mWidth, vbPixels, vbTwips)
            UserControl.Height = ScaleY(mHeight, vbPixels, vbTwips)
        Else
            UserControl.Width = ScaleX(19, vbPixels, vbTwips)
            UserControl.Height = ScaleY(19, vbPixels, vbTwips)
        End If
        PositionOnForm
    End If
End Sub

Private Sub UserControl_Show()
    If Not Ambient.UserMode Then
        PositionOnForm
    End If
End Sub

Private Sub PositionOnForm()
    On Error GoTo EExit
    Extender.Left = Extender.Container.ScaleWidth - Extender.Container.ScaleX(mWidth, vbPixels, Extender.Container.ScaleMode)
    Extender.Top = Extender.Container.ScaleHeight - Extender.Container.ScaleY(mHeight, vbPixels, Extender.Container.ScaleMode)
    Extender.ZOrder
EExit:
End Sub

Private Sub SetGrip()
    SetImage
    If Ambient.UserMode Then
        Set mForm = Parent
    End If
End Sub

Public Property Let BackColor(nValue As OLE_COLOR)
    If nValue <> mBackColor Then
        mBackColor = nValue
        mBackColorIsFromAmbient = (mBackColor = Ambient.BackColor)
        PropertyChanged "BackColor"
        SetGrip
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mBackColor
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", BackColor, Ambient.BackColor
End Sub

Private Function RGBColor(nColor As Long) As COLORS_RGB
    If nColor < 0 Then Exit Function
    RGBColor.R = nColor And 255
    RGBColor.G = (nColor \ 256) And 255
    RGBColor.b = (nColor \ 65536) And 255
End Function

Private Sub SetImage()
    Dim Y As Long
    Dim X As Long
    Dim iDarkColor As Long
    Dim iLightColor As Long
    Dim iBackColorRGB As COLORS_RGB
    Dim iAuxRGBColor As COLORS_RGB
    Dim iColor As Long
    Dim iBackColor As Long
    Dim iAmbientUserMode As Boolean
    Dim iPxColor As Long
    Dim iTx As Single
    Dim iPic As StdPicture
    Dim iCC As Byte
    Dim iCB As Long
    
    mSettingImage = True
    iAmbientUserMode = Ambient.UserMode
    TranslateColor mBackColor, 0, iBackColor
    If iAmbientUserMode Then
        iTx = Screen.TwipsPerPixelX
        If iTx >= 15 Then
            Set iPic = img14pix.Picture
            mWidth = 14
            mHeight = 14
        ElseIf iTx >= 12 Then
            Set iPic = img17pix.Picture
            mWidth = 17
            mHeight = 17
        ElseIf iTx >= 10 Then
            Set iPic = img21pix.Picture
            mWidth = 21
            mHeight = 21
        ElseIf iTx >= 7 Then ' 192 DPI
            Set iPic = StretchPicNN(img14pix.Picture, 2)
            mWidth = 28
            mHeight = 28
        ElseIf iTx >= 6 Then
            Set iPic = StretchPicNN(img17pix.Picture, 2)
            mWidth = 34
            mHeight = 34
        ElseIf iTx >= 5 Then
            Set iPic = StretchPicNN(img21pix.Picture, 2)
            mWidth = 42
            mHeight = 42
        ElseIf iTx >= 4 Then  ' 289 to 360 DPI
            Set iPic = StretchPicNN(img17pix.Picture, 3)
            mWidth = 51
            mHeight = 51
        ElseIf iTx >= 3 Then   ' 361 to 480 DPI
            Set iPic = StretchPicNN(img21pix.Picture, 3)
            mWidth = 63
            mHeight = 63
        ElseIf iTx >= 2 Then   ' 481 to 720 DPI
            Set iPic = StretchPicNN(img21pix.Picture, 5)
            mWidth = 105
            mHeight = 105
        Else ' greater than 720 DPI
            Set iPic = StretchPicNN(img21pix.Picture, 10)
            mWidth = 210
            mHeight = 210
        End If
    Else
        mWidth = 19
        mHeight = 19
        Set iPic = imgDesignMode.Picture
    End If
    UserControl.AutoRedraw = True
    UserControl.BackColor = iBackColor
    Set UserControl.Picture = Nothing
    UserControl.Cls
    
    iBackColorRGB = RGBColor(iBackColor)
    iCB = GetColorBrightness(iBackColor)
    If iCB > 230 Then
        iCC = 100
    ElseIf iCB > 180 Then
        iCC = 52
    ElseIf iCB < 100 Then
        iCC = (230 - iCB) * 0.8
    Else
        iCC = 100
    End If
    
    iAuxRGBColor.R = iBackColorRGB.R - iCC
    iAuxRGBColor.G = iBackColorRGB.G - iCC
    iAuxRGBColor.b = iBackColorRGB.b - iCC
    If iAuxRGBColor.R < 0 Then iAuxRGBColor.R = 0
    If iAuxRGBColor.G < 0 Then iAuxRGBColor.G = 0
    If iAuxRGBColor.b < 0 Then iAuxRGBColor.b = 0
    iDarkColor = RGB(iAuxRGBColor.R, iAuxRGBColor.G, iAuxRGBColor.b)
    
    iAuxRGBColor.R = iBackColorRGB.R + iCC
    iAuxRGBColor.G = iBackColorRGB.G + iCC
    iAuxRGBColor.b = iBackColorRGB.b + iCC
    If iAuxRGBColor.R > 255 Then iAuxRGBColor.R = 255
    If iAuxRGBColor.G > 255 Then iAuxRGBColor.G = 255
    If iAuxRGBColor.b > 255 Then iAuxRGBColor.b = 255
    iLightColor = RGB(iAuxRGBColor.R, iAuxRGBColor.G, iAuxRGBColor.b)
    
    UserControl.Width = ScaleX(mWidth * 2, vbPixels, vbTwips)
    UserControl.Height = ScaleY(mHeight, vbPixels, vbTwips)
    UserControl.PaintPicture iPic, mWidth, 0
    
    For Y = 0 To UserControl.ScaleHeight - 1
        For X = mWidth To UserControl.ScaleWidth - 1
            iPxColor = GetPixel(UserControl.hdc, X, Y)
            Select Case iPxColor
                Case 14215660
                    iColor = iBackColor
                Case 10597816, 10728632
                    iColor = IIf(iCB < 160, iLightColor, iDarkColor)
                Case 16777215
                    iColor = IIf(iCB < 160, iDarkColor, iLightColor)
                Case Else
                    iColor = iPxColor
            End Select
            SetPixel UserControl.hdc, X - mWidth, Y, iColor
        Next X
    Next Y
    UserControl.Width = ScaleX(mWidth, vbPixels, vbTwips)
    Set iPic = UserControl.Image
    UserControl.Cls
    UserControl.PaintPicture iPic, 0, 0
    Set iPic = UserControl.Image
    UserControl.Cls
    Set UserControl.Picture = iPic
    mSettingImage = False
End Sub

Public Sub Refresh()
    mForm_Resize
End Sub

Private Function StretchPicNN(nPic As StdPicture, nFactor As Long) As StdPicture
    Dim iPicInfo As BITMAP
    Dim PicSizeW As Long
    Dim PicSizeH As Long
    Dim iW As Long
    Dim iH As Long
    
    iW = UserControl.Width
    iH = UserControl.Height
    
    GetObjectAPI nPic.Handle, Len(iPicInfo), iPicInfo
    PicSizeW = iPicInfo.bmWidth
    PicSizeH = iPicInfo.bmHeight
    
    UserControl.Width = PicSizeW * nFactor * Screen.TwipsPerPixelX
    UserControl.Height = PicSizeH * nFactor * Screen.TwipsPerPixelY
    
    UserControl.PaintPicture nPic, 0, 0, PicSizeW * nFactor, PicSizeH * nFactor
    
    Set StretchPicNN = UserControl.Image
    UserControl.Cls

    UserControl.Width = iW
    UserControl.Height = iH

End Function

Private Function GetColorBrightness(ByVal nColor As Long) As Long
    Dim iR As Byte
    Dim iG As Byte
    Dim iB As Byte
    
    OleTranslateColor nColor, 0, nColor
    
    iR = nColor And 255
    iG = (nColor \ 256) And 255
    iB = (nColor \ 65536) And 255
    
    GetColorBrightness = (0.2125 * iR + 0.7154 * iG + 0.0721 * iB)
End Function
