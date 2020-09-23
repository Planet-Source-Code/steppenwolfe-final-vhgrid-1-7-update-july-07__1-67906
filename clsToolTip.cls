VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsToolTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements GXISubclass

Private Const BM_TRANSPARENT                    As Long = &H1

Private Const CLR_NONE                          As Long = -1

Private Const DT_LEFT                           As Long = &H0&
Private Const DT_CENTER                         As Long = &H1&
Private Const DT_VCENTER                        As Long = &H4&
Private Const DT_WORDBREAK                      As Long = &H10&
Private Const DT_SINGLELINE                     As Long = &H20&
Private Const DT_CALCRECT                       As Long = &H400&
Private Const DT_END_ELLIPSIS                   As Long = &H8000&
Private Const DT_RTLREADING                     As Long = &H20000

Private Const FW_NORMAL                         As Long = 400
Private Const FW_BOLD                           As Long = 700

Private Const ILD_NORMAL                        As Long = &H0
Private Const ILD_TRANSPARENT                   As Long = &H1

Private Const LF_ANTIALIASED_QUALITY            As Long = 4

Private Const LOGPIXELSY                        As Long = 90

Private Const SS_OWNERDRAW                      As Long = &HD

Private Const SWP_SHOWWINDOW                    As Long = &H40
Private Const SWP_NOACTIVATE                    As Long = &H10

Private Const WM_TIMER                          As Long = &H113&

Private Const WS_OVERLAPPED                     As Long = &H0
Private Const WS_POPUP                          As Long = &H80000000
Private Const WS_CHILD                          As Long = &H40000000
Private Const WS_MINIMIZE                       As Long = &H20000000
Private Const WS_CLIPSIBLINGS                   As Long = &H4000000
Private Const WS_CAPTION                        As Long = &HC00000
Private Const WS_BORDER                         As Long = &H800000
Private Const WS_SYSMENU                        As Long = &H80000
Private Const WS_THICKFRAME                     As Long = &H40000
Private Const WS_MINIMIZEBOX                    As Long = &H20000
Private Const WS_MAXIMIZEBOX                    As Long = &H10000
Private Const WS_OVERLAPPEDWINDOW               As Long = _
(WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)

Private Const WS_EX_TOOLWINDOW                  As Long = &H80


Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Public Enum ETPToolTipPosition
    RightBottom = 0&
    RightCenter = 1&
    RightTop = 2&
    LeftBottom = 3&
    LeftCenter = 4&
    LeftTop = 5&
End Enum

Private Enum ImageLockMode
    ImageLockModeRead = &H1
    ImageLockModeWrite = &H2
    ImageLockModeUserInputBuf = &H4
End Enum

Private Enum PixelFormat
    PixelFormatIndexed = &H10000
    PixelFormatGDI = &H20000
    PixelFormatAlpha = &H40000
    PixelFormatPAlpha = &H80000
    PixelFormatExtended = &H100000
    PixelFormatCanonical = &H200000
    PixelFormatUndefined = 0
    PixelFormatDontCare = 0
    PixelFormat1bppIndexed = &H30101
    PixelFormat4bppIndexed = &H30402
    PixelFormat8bppIndexed = &H30803
    PixelFormat16bppGreyScale = &H101004
    PixelFormat16bppRGB555 = &H21005
    PixelFormat16bppRGB565 = &H21006
    PixelFormat16bppARGB1555 = &H61007
    PixelFormat24bppRGB = &H21808
    PixelFormat32bppRGB = &H22009
    PixelFormat32bppARGB = &H26200A
    PixelFormat32bppPARGB = &HE200B
    PixelFormat48bppRGB = &H10300C
    PixelFormat64bppARGB = &H34400D
    PixelFormat64bppPARGB = &H1C400E
    PixelFormatMax = 15
End Enum


Private Type GdiplusStartupInput
   GdiplusVersion                               As Long
   DebugEventCallback                           As Long
   SuppressBackgroundThread                     As Long
   SuppressExternalCodecs                       As Long
End Type

Private Type BITMAPDATA
    Width                                       As Long
    Height                                      As Long
    stride                                      As Long
    PixelFormat                                 As Long
    scan0                                       As Long
    Reserved                                    As Long
End Type

Private Type ICONINFO
    fIcon                                       As Long
    xHotspot                                    As Long
    yHotspot                                    As Long
    hBmMask                                     As Long
    hbmColor                                    As Long
End Type

Private Type RECT
    left                                        As Long
    top                                         As Long
    Right                                       As Long
    Bottom                                      As Long
End Type

Private Type POINTAPI
    x                                           As Long
    y                                           As Long
End Type

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                                      As Long
    dwFlags                                     As TRACKMOUSEEVENT_FLAGS
    hwndTrack                                   As Long
    dwHoverTime                                 As Long
End Type

Private Type VERSIONINFO
    dwOSVersionInfoSize                         As Long
    dwMajorVersion                              As Long
    dwMinorVersion                              As Long
    dwBuildNumber                               As Long
    dwPlatformId                                As Long
    szCSDVersion                                As String * 128
End Type

Private Type LOGFONT
    lfHeight                                    As Long
    lfWidth                                     As Long
    lfEscapement                                As Long
    lfOrientation                               As Long
    lfWeight                                    As Long
    lfItalic                                    As Byte
    lfUnderline                                 As Byte
    lfStrikeOut                                 As Byte
    lfCharSet                                   As Byte
    lfOutPrecision                              As Byte
    lfClipPrecision                             As Byte
    lfQuality                                   As Byte
    lfPitchAndFamily                            As Byte
    lfFaceName(32)                              As Byte
End Type


Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function TrackMouseEventComCtl Lib "comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersion As VERSIONINFO) As Long

Private Declare Function CreateWindowExA Lib "user32" (ByVal dwExStyle As Long, _
                                                       ByVal lpClassName As String, _
                                                       ByVal lpWindowName As String, _
                                                       ByVal dwStyle As Long, _
                                                       ByVal x As Long, _
                                                       ByVal y As Long, _
                                                       ByVal nWidth As Long, _
                                                       ByVal nHeight As Long, _
                                                       ByVal hWndParent As Long, _
                                                       ByVal hMenu As Long, _
                                                       ByVal hInstance As Long, _
                                                       lpParam As Any) As Long

Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, _
                                                       ByVal lpClassName As Long, _
                                                       ByVal lpWindowName As Long, _
                                                       ByVal dwStyle As Long, _
                                                       ByVal x As Long, _
                                                       ByVal y As Long, _
                                                       ByVal nWidth As Long, _
                                                       ByVal nHeight As Long, _
                                                       ByVal hWndParent As Long, _
                                                       ByVal hMenu As Long, _
                                                       ByVal hInstance As Long, _
                                                       lpParam As Any) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                               ByVal hWndInsertAfter As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               ByVal cx As Long, _
                                               ByVal cy As Long, _
                                               ByVal wFlags As Long)

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
                                                      lpPoint As POINTAPI) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, _
                                                 lpRect As RECT, _
                                                 ByVal hBrush As Long) As Long

Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, _
                                                lpSourceRect As RECT) As Long

Private Declare Function DrawTextA Lib "user32" (ByVal hdc As Long, _
                                                 ByVal lpStr As String, _
                                                 ByVal nCount As Long, _
                                                 lpRect As RECT, _
                                                 ByVal wFormat As Long) As Long

Private Declare Function DrawTextW Lib "user32" (ByVal hdc As Long, _
                                                 ByVal lpStr As Long, _
                                                 ByVal nCount As Long, _
                                                 lpRect As RECT, _
                                                 ByVal wFormat As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function ImageList_GetIconSize Lib "comctl32" (ByVal hIml As Long, _
                                                               cx As Long, _
                                                               cy As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal nBkMode As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function CreateDc Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                 lpDeviceName As Any, _
                                                                 lpOutput As Any, _
                                                                 lpInitData As Any) As Long

Private Declare Function CreateFontIndirectA Lib "gdi32" (lpLogFont As LOGFONT) As Long

Private Declare Function CreateFontIndirectW Lib "gdi32" (lpLogFont As LOGFONT) As Long

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, _
                                                ByVal nNumerator As Long, _
                                                ByVal nDenominator As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nIndex As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal hIml As Long, _
                                                            ByVal I As Long, _
                                                            ByVal hdcDst As Long, _
                                                            ByVal x As Long, _
                                                            ByVal y As Long, _
                                                            ByVal fStyle As Long) As Long

Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long, _
                                               lpPoint As POINTAPI) As Long

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, _
                                             ByVal x As Long, _
                                             ByVal y As Long) As Long

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, _
                                                ByVal nIDEvent As Long, _
                                                ByVal uElapse As Long, _
                                                ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal nIDEvent As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long

Private Declare Function ExcludeClipRect Lib "gdi32" (ByVal hdc As Long, _
                                                      ByVal X1 As Long, _
                                                      ByVal y1 As Long, _
                                                      ByVal x2 As Long, _
                                                      ByVal y2 As Long) As Long

Private Declare Function ImageList_GetIcon Lib "comctl32" (ByVal hImageList As Long, _
                                                           ByVal ImgIndex As Long, _
                                                           ByVal fuFlags As Long) As Long

Private Declare Function DestroyIcon Lib "user32" (ByVal hicon As Long) As Long

Private Declare Function ImageList_DrawEx Lib "comctl32" (ByVal hIml As Long, _
                                                          ByVal I As Long, _
                                                          ByVal hdcDst As Long, _
                                                          ByVal x As Long, _
                                                          ByVal y As Long, _
                                                          ByVal dx As Long, _
                                                          ByVal dy As Long, _
                                                          ByVal rgbBk As Long, _
                                                          ByVal rgbFg As Long, _
                                                          ByVal fStyle As Long) As Long

Private Declare Function GetIconInfo Lib "user32" (ByVal hicon As Long, _
                                                   piconinfo As ICONINFO) As Long

Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, _
                                                       inputbuf As GdiplusStartupInput, _
                                                       Optional ByVal outputbuf As Long = 0) As Long

Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)

Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, _
                                                          graphics As Long) As Long

Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, _
                                                          Width As Long) As Long

Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, _
                                                           Height As Long) As Long

Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long

Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long

Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, _
                                                                    ByVal hpal As Long, _
                                                                    bmap As Long) As Long

Private Declare Function GdipDrawImage Lib "gdiplus" (ByVal graphics As Long, _
                                                      ByVal image As Long, _
                                                      ByVal x As Single, _
                                                      ByVal y As Single) As Long

Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal lImage As Long, _
                                                                lFormat As Long) As Long

Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal BITMAP As Long, _
                                                           rc As RECT, _
                                                           ByVal flags As ImageLockMode, _
                                                           ByVal PixelFormat As Long, _
                                                           lockedBitmapData As BITMAPDATA) As Long

Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, _
                                                                  ByVal Height As Long, _
                                                                  ByVal stride As Long, _
                                                                  ByVal PixelFormat As Long, _
                                                                  scan0 As Any, BITMAP As Long) As Long

Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal BITMAP As Long, _
                                                             lockedBitmapData As BITMAPDATA) As Long

Private Declare Function GdipBitmapGetPixel Lib "gdiplus" (ByVal BITMAP As Long, _
                                                           ByVal x As Long, _
                                                           ByVal y As Long, _
                                                           color As Long) As Long

Private m_bIsNt                                 As Boolean
Private m_bRestrictToClient                     As Boolean
Private m_bIsActive                             As Boolean
Private m_bMultiline                            As Boolean
Private m_bGradient                             As Boolean
Private m_bXPColors                             As Boolean
Private m_bAutoTips                             As Boolean
Private m_bTimerActive                          As Boolean
Private m_bTrackUser32                          As Boolean
Private m_bShowing                              As Boolean
Private m_bVisiting                             As Boolean
Private m_bUseUnicode                           As Boolean
Private m_bFontRightLeading                     As Boolean
Private m_bGdiPlusLoaded                        As Boolean
Private m_bUseGdiPlus                           As Boolean
Private m_lGdiPlusToken                         As Long
Private m_lLastX                                As Long
Private m_lLastY                                As Long
Private m_lSafeTimer                            As Long
Private m_lColorOffset                          As Long
Private m_lTransparency                         As Long
Private m_lhTitleFnt                            As Long
Private m_lhTextFnt                             As Long
Private m_lToolTipPosition                      As Long
Private m_lToolHwnd                             As Long
Private m_lBackColor                            As Long
Private m_lDelayTime                            As Long
Private m_lForeColor                            As Long
Private m_lIcon                                 As Long
Private m_lVisibleTime                          As Long
Private m_lParentHwnd                           As Long
Private m_lCtrlHwnd                             As Long
Private m_lWidth                                As Long
Private m_lHeight                               As Long
Private m_lImlHwnd                              As Long
Private m_sText                                 As String
Private m_sTitle                                As String
Private m_tRWnd                                 As RECT
Private m_tRTxt                                 As RECT
Private m_tRTtl                                 As RECT
Private m_tRIcn                                 As RECT
Private m_oFont                                 As StdFont
Private m_oTitleFont                            As StdFont
Private m_oTextFont                             As StdFont
Private m_cToolDc                               As clsStoreDc
Private m_cRender                               As clsRender
Private m_TTSubclass                            As GXMSubclass


Private Sub Class_Initialize()

    m_bIsNt = CompatabilityCheck
    Set m_TTSubclass = New GXMSubclass
    Set m_cRender = New clsRender
    Set m_cToolDc = New clsStoreDc
    m_bTrackUser32 = FunctionExported("TrackMouseEvent", "User32")
    m_lWidth = 150
    m_lBackColor = GetSysColor(&H80000018 And &H1F)
    m_lForeColor = &H343434
    m_lIcon = -1
    m_lTransparency = -1
    m_lColorOffset = &HCACACA
    m_lDelayTime = 1
    m_lVisibleTime = 2
    m_lForeColor = &H333333

End Sub

Private Function LibraryExists(ByVal sModule As String) As Boolean
'/* test for library support

Dim lhMod   As Long
Dim bLoad   As Boolean

    lhMod = GetModuleHandleA(sModule)
    If (lhMod = 0) Then
        lhMod = LoadLibraryA(sModule)
        bLoad = True
    End If
    If Not (lhMod = 0) Then
        LibraryExists = True
        If bLoad Then
            FreeLibrary lhMod
        End If
    End If

End Function


'**********************************************************************
'*                              GDIPLUS
'**********************************************************************

Public Property Get GdiPlusLoaded() As Boolean
    GdiPlusLoaded = m_bGdiPlusLoaded
End Property

Public Property Get UseGdiPlus() As Boolean
    UseGdiPlus = m_bUseGdiPlus
End Property

Public Property Let UseGdiPlus(ByVal PropVal As Boolean)
    If PropVal Then
        If LibraryExists("gdiplus") Then
            LoadGdiPlus
        End If
    Else
        GdiUnload
    End If
    m_bUseGdiPlus = PropVal
End Property

Private Function LoadGdiPlus() As Boolean
'/* load gdi library

Dim tGPInput As GdiplusStartupInput

    GdiUnload
    tGPInput.GdiplusVersion = 1
    If (GdiplusStartup(m_lGdiPlusToken, tGPInput) = 0) Then
        LoadGdiPlus = True
        m_bGdiPlusLoaded = True
    End If

End Function

Private Sub GdiUnload()
'/* unload gdi library

    If m_bGdiPlusLoaded Then
        If Not (m_lGdiPlusToken = 0) Then
            GdiplusShutdown m_lGdiPlusToken
            m_lGdiPlusToken = 0
            m_bGdiPlusLoaded = False
        End If
    End If
    
End Sub

Private Sub DrawAlphaIcon(ByVal lIndex As Long, _
                          ByVal lDestDc As Long, _
                          ByVal lX As Long, _
                          ByVal lY As Long)

'/* could use variation of this to extract and
'/* draw 32b alpha bitmaps/png also..

Dim lGraphics   As Long
Dim lpBitmap    As Long
Dim lhIcon      As Long
Dim lHeight     As Long
Dim lWidth      As Long
Dim lpBmpNew    As Long
Dim lFormat     As Long
Dim tBmpData    As BITMAPDATA
Dim tIcnInfo    As ICONINFO
Dim tRect       As RECT

    '/* fetch icon handle
    lhIcon = ImageList_GetIcon(m_lImlHwnd, lIndex, 1&)
    '/* icon data structure
    GetIconInfo lhIcon, tIcnInfo
    '/* load graphics
    GdipCreateFromHDC lDestDc, lGraphics
    '/* copy icon bmp to a new image
    GdipCreateBitmapFromHBITMAP tIcnInfo.hbmColor, 0&, lpBitmap
    '/* dispose of resource
    DeleteObject (tIcnInfo.hbmColor)
    DeleteObject (tIcnInfo.hBmMask)
    '/* get the format
    GdipGetImagePixelFormat lpBitmap, lFormat
    '/* not alpha
    If Not (lFormat < PixelFormat32bppRGB) Then
        '/* image dimensions
        GdipGetImageHeight lpBitmap, lHeight
        GdipGetImageWidth lpBitmap, lWidth
        With tRect
            .Bottom = lHeight
            .left = 0
            .Right = lWidth
            .top = 0
        End With
        '/* create a new 32b bmp
        GdipCreateBitmapFromScan0 lWidth, lHeight, 0&, PixelFormat32bppARGB, ByVal 0&, lpBmpNew
        '/* copy 32b data structure
        GdipBitmapLockBits lpBitmap, tRect, ImageLockModeRead, lFormat, tBmpData
        GdipBitmapLockBits lpBmpNew, tRect, ImageLockModeWrite Or ImageLockModeUserInputBuf, PixelFormat32bppARGB, tBmpData
        '/* unlock
        GdipBitmapUnlockBits lpBmpNew, tBmpData
        GdipBitmapUnlockBits lpBitmap, tBmpData
        '/* test for alpha channel
        If IsAlphaBitmap(lpBmpNew, lWidth, lHeight) Then
            '/* draw the alpha image
            GdipDrawImage lGraphics, lpBmpNew, lX, lY
        Else
            '/* draw flat image
            ImageList_DrawEx m_lImlHwnd, lIndex, lDestDc, lX, lY, 0&, 0&, CLR_NONE, 0&, ILD_TRANSPARENT
        End If
    Else
        '/* draw flat image
        ImageList_DrawEx m_lImlHwnd, lIndex, lDestDc, lX, lY, 0&, 0&, CLR_NONE, 0&, ILD_TRANSPARENT
    End If

    '/* cleanup
    DestroyIcon lhIcon
    GdipDisposeImage lpBitmap
    GdipDisposeImage lpBmpNew
    GdipDeleteGraphics lGraphics
   
End Sub

Private Function IsAlphaBitmap(ByVal lBitmap As Long, _
                               ByVal lWidth As Long, _
                               ByVal lHeight As Long) As Boolean

'/* if there is an alpha channel
'/* colors are right shifted

Dim lY      As Long
Dim lX      As Long
Dim lArgb   As Long

    For lY = 0 To lHeight - 1
        For lX = 0 To lWidth - 1
            GdipBitmapGetPixel lBitmap, lX, lY, lArgb
            If (lArgb > &HFF000000) Then
                If (lArgb < &HFFFFFFFF) Then
                    IsAlphaBitmap = True
                    Exit For
                End If
            End If
        Next lX
        If IsAlphaBitmap Then
            Exit For
        End If
    Next lY

End Function


'**********************************************************************
'*                              PROPERTIES
'**********************************************************************

Public Property Get AutoTips() As Boolean
    AutoTips = m_bAutoTips
End Property

Public Property Let AutoTips(ByVal PropVal As Boolean)
    m_bAutoTips = PropVal
End Property

Public Property Get BackColor() As Long
    BackColor = m_lBackColor
End Property

Public Property Let BackColor(ByVal PropVal As Long)
    If m_bXPColors Then
        m_lBackColor = m_cRender.XPShift(PropVal)
    Else
        m_lBackColor = PropVal
    End If
End Property

Public Property Get ColorOffset() As Long
    ColorOffset = m_lColorOffset
End Property

Public Property Let ColorOffset(ByVal PropVal As Long)
    If m_bXPColors Then
        m_lColorOffset = m_cRender.XPShift(PropVal)
    Else
        m_lColorOffset = PropVal
    End If
End Property

Public Property Get CtrlHwnd() As Long
    CtrlHwnd = m_lCtrlHwnd
End Property

Public Property Let CtrlHwnd(ByVal PropVal As Long)
    m_lCtrlHwnd = PropVal
End Property

Public Property Get DelayTime() As Long
    DelayTime = m_lDelayTime
End Property

Public Property Let DelayTime(ByVal PropVal As Long)
    m_lDelayTime = PropVal
End Property

Public Property Get Font() As StdFont
    Set Font = m_oFont
End Property

Public Property Get FontRightLeading() As Boolean
'/* [get] right align fonts
    FontRightLeading = m_bFontRightLeading
End Property

Public Property Let FontRightLeading(ByVal PropvVal As Boolean)
'/* [let] right align fonts
    m_bFontRightLeading = PropvVal
End Property

Public Property Set Font(ByVal PropVal As StdFont)
    Set m_oFont = PropVal
End Property

Public Property Get ForeColor() As Long
    ForeColor = m_lForeColor
End Property

Public Property Let ForeColor(ByVal PropVal As Long)
    m_lForeColor = PropVal
End Property

Public Property Get Gradient() As Boolean
    Gradient = m_bGradient
End Property

Public Property Let Gradient(ByVal PropVal As Boolean)
    m_bGradient = PropVal
End Property

Public Property Get Height() As Long
    Height = m_lHeight
End Property

Public Property Let Height(ByVal PropVal As Long)
    m_lHeight = PropVal
End Property

Public Property Get Icon() As Long
    Icon = m_lIcon
End Property

Public Property Let Icon(ByVal PropVal As Long)
    m_lIcon = PropVal
End Property

Public Property Get ImlHwnd() As Long
    ImlHwnd = m_lImlHwnd
End Property

Public Property Let ImlHwnd(ByVal PropVal As Long)
    m_lImlHwnd = PropVal
End Property

Public Property Get Multiline() As Boolean
    Multiline = m_bMultiline
End Property

Public Property Let Multiline(ByVal PropVal As Boolean)
    m_bMultiline = PropVal
End Property

Public Property Get RestrictToClient() As Boolean
    RestrictToClient = m_bRestrictToClient
End Property

Public Property Let RestrictToClient(ByVal PropVal As Boolean)
    m_bRestrictToClient = PropVal
End Property

Public Property Get Showing() As Boolean
    Showing = m_bShowing
End Property

Public Property Get Text() As String
    Text = m_sText
End Property

Public Property Let Text(ByVal PropVal As String)
    m_sText = PropVal
End Property

Public Property Get Title() As String
    Title = m_sTitle
End Property

Public Property Let Title(ByVal PropVal As String)
    m_sTitle = PropVal
End Property

Public Property Get ToolTipPosition() As ETPToolTipPosition
    ToolTipPosition = m_lToolTipPosition
End Property

Public Property Let ToolTipPosition(ByVal PropVal As ETPToolTipPosition)
    m_lToolTipPosition = PropVal
End Property

Public Property Get Transparency() As Long
    Transparency = m_lTransparency
End Property

Public Property Let Transparency(ByVal PropVal As Long)
    m_lTransparency = PropVal
End Property

Public Property Get UseUnicode() As Boolean
    UseUnicode = m_bUseUnicode
End Property

Public Property Let UseUnicode(ByVal PropVal As Boolean)
    m_bUseUnicode = PropVal
End Property

Public Property Get VisibleTime() As Long
    VisibleTime = m_lVisibleTime
End Property

Public Property Let VisibleTime(ByVal PropVal As Long)
    m_lVisibleTime = PropVal
End Property

Public Property Get Width() As Long
    Width = m_lWidth
End Property

Public Property Let Width(ByVal PropVal As Long)
    m_lWidth = PropVal
End Property

Public Property Get XPColors() As Boolean
    XPColors = m_bXPColors
End Property

Public Property Let XPColors(ByVal PropVal As Boolean)
    m_bXPColors = PropVal
End Property

'> Create
'>>>>>>>>>>>>>>>>
Private Sub CalcTextRect(ByVal lHdc As Long, _
                         ByVal lWidth As Long, _
                         ByVal sText As String, _
                         ByVal lFlags As Long, _
                         ByVal lFont As Long, _
                         ByRef tRect As RECT)

Dim lFntOld As Long

    tRect.Right = lWidth
    lFntOld = SelectObject(lHdc, lFont)
    If m_bIsNt Then
        DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags
    Else
        DrawTextA lHdc, sText, -1, tRect, lFlags
    End If
    SelectObject lHdc, lFntOld
    
End Sub

Private Sub CalculateOffsets()

Dim lHdc    As Long
Dim lCx     As Long
Dim lCy     As Long
Dim lTHgt   As Long
Dim lMnHt   As Long
Dim lFlags  As Long
Dim tPnt    As POINTAPI

    '/* icon size
    If Not (m_lImlHwnd = 0) Then
        If Not (m_lIcon = -1) Then
            ImageList_GetIconSize m_lImlHwnd, lCx, lCy
            With m_tRIcn
                .left = 8
                .top = 8
                .Right = (lCx + 8)
                .Bottom = (lCy + 8)
                lMnHt = (.Bottom + 8)
                m_tRTtl.left = (.Right + 8)
            End With
        Else
            m_tRTtl.left = 8
        End If
    Else
        m_tRTtl.left = 8
    End If
    
    '/* create fonts
    CreateFonts
    lHdc = GetDC(m_lToolHwnd)
    
    If m_bMultiline Then
        lFlags = DT_WORDBREAK Or DT_CALCRECT
        '/* calculate title height
        With m_tRTtl
            .top = 8
            CalcTextRect lHdc, m_lWidth, m_sTitle, lFlags, m_lhTitleFnt, m_tRTtl
            lTHgt = .Bottom + 4
        End With
        '/* calculate text
        With m_tRTxt
            .top = lTHgt
            .left = m_tRTtl.left
            CalcTextRect lHdc, m_lWidth, m_sText, lFlags, m_lhTextFnt, m_tRTxt
            lTHgt = .Bottom + 10
            If (.Right > m_tRTtl.Right) Then
                m_lWidth = .Right + 10
            Else
                m_lWidth = m_tRTtl.Right + 10
            End If
        End With
    Else
        lFlags = DT_LEFT Or DT_CALCRECT
        m_tRTtl.top = 8
        CalcTextRect lHdc, 0, m_sTitle, lFlags, m_lhTitleFnt, m_tRTtl
        With m_tRTtl
            m_lWidth = .Right + 6
        End With
        With m_tRTxt
            .top = 8
            .left = m_lWidth
            CalcTextRect lHdc, 0, m_sText, lFlags, m_lhTextFnt, m_tRTxt
            lTHgt = .Bottom + 8
            m_lWidth = .Right + 4
        End With
    End If
    '/* test minimum
    If (lTHgt > lMnHt) Then
        m_lHeight = lTHgt
    Else
        m_lHeight = lMnHt
    End If
    
    ReleaseDC m_lToolHwnd, lHdc
    GetCursorPos tPnt
    ScreenToClient m_lParentHwnd, tPnt
    
    '/* window size
    With m_tRWnd
        '/* shadow offset
        .Right = m_lWidth + 4
        .Bottom = m_lHeight + 4
        Select Case m_lToolTipPosition
        Case 0
            .top = tPnt.y + 4
            .left = tPnt.x + 4
        Case 1
            .top = (tPnt.y + 4) + (m_lHeight / 2)
            .left = tPnt.x + 4
        Case 2
            .top = (tPnt.y + 4) + m_lHeight
            .left = tPnt.x + 4
        Case 3
            .top = tPnt.y + 4
            .left = tPnt.x - 4
        Case 4
            .top = (tPnt.y + 4) + (m_lHeight / 2)
            .left = tPnt.x - 4
        Case 5
            .top = (tPnt.y + 4) + m_lHeight
            .left = tPnt.x - 4
        End Select
    End With
    
End Sub

Private Function CompatabilityCheck() As Boolean

Dim tVer  As VERSIONINFO

    tVer.dwOSVersionInfoSize = Len(tVer)
    GetVersionEx tVer
    If tVer.dwMajorVersion >= 5 Then
        CompatabilityCheck = True
    End If

End Function

Private Sub CreateToolTip()

Dim bOnDesktop  As Boolean
Dim lTTStyle    As Long

    If m_bRestrictToClient Then
        m_lParentHwnd = GetParent(m_lCtrlHwnd)
        If m_lParentHwnd = 0 Then
            m_lParentHwnd = m_lCtrlHwnd
        End If
        bOnDesktop = (m_lParentHwnd = GetDesktopWindow())
        If bOnDesktop Then
            m_bRestrictToClient = False
        End If
    Else
        m_lParentHwnd = GetDesktopWindow()
        bOnDesktop = True
    End If
    lTTStyle = SS_OWNERDRAW Or WS_CHILD Or WS_CLIPSIBLINGS Or WS_OVERLAPPED
    
    '/* create window
    If m_bIsNt Then
        m_lToolHwnd = CreateWindowExW(-bOnDesktop * WS_EX_TOOLWINDOW, StrPtr("Static"), StrPtr(""), lTTStyle, _
            0&, 0&, 0&, 0&, m_lParentHwnd, 0, App.hInstance, ByVal 0&)
    Else
        m_lToolHwnd = CreateWindowExA(-bOnDesktop * WS_EX_TOOLWINDOW, "Static", "", lTTStyle, _
            0&, 0&, 0&, 0&, m_lParentHwnd, 0, App.hInstance, ByVal 0&)
    End If

End Sub

Private Sub SetPosition()
    
    If Not (m_lToolHwnd = 0) Then
        With m_tRWnd
            If m_bRestrictToClient Then
                SetWindowPos m_lToolHwnd, 0&, .left, .top, .Right, .Bottom, SWP_SHOWWINDOW Or SWP_NOACTIVATE
            Else
                SetWindowPos m_lToolHwnd, -1&, .left, .top, .Right, .Bottom, SWP_SHOWWINDOW Or SWP_NOACTIVATE
            End If
        End With
    End If

End Sub

Public Sub ShowTipManual(ByVal lHwnd As Long, _
                         ByVal sText As String, _
                         ByVal sTitle As String, _
                         ByVal lIcon As Long, _
                         ByVal lWidth As Long)

    If Not (lHwnd = 0) Then
        m_lCtrlHwnd = lHwnd
        Text = sText
        Title = sTitle
        Icon = lIcon
        Width = lWidth
        DrawTip
    End If

End Sub

Public Sub Start()
    
    If Not (m_lCtrlHwnd = 0) Then
        If Not m_bIsActive Then
            MessageAttach
        End If
    End If
    
End Sub

'> Render
'>>>>>>>>>>>>>>>>
Public Sub DrawTip()

    CreateToolTip
    CalculateOffsets
    SetPosition
    RenderTip
            
End Sub

Private Sub RenderTip()

Dim lWdc    As Long
Dim lHdc    As Long
Dim lFntOld As Long
Dim tRect   As RECT

    GetClientRect m_lToolHwnd, tRect
    '/* draw dc
    With m_cToolDc
        .Width = tRect.Right
        .Height = tRect.Bottom
        lHdc = .hdc
    End With
    '/* gradient
    If m_bGradient Then
        With tRect
            m_cRender.Gradient lHdc, 0, .Right, 0, .Bottom, m_lColorOffset, m_lBackColor, Fill_Vertical
        End With
    Else
        PaintTip lHdc, tRect
    End If

    '/* shadow
    FrameTip lHdc, tRect
    '/* text
    lFntOld = SelectObject(lHdc, m_lhTitleFnt)
    DrawText lHdc, m_sTitle, m_tRTtl
    SelectObject lHdc, lFntOld
    lFntOld = SelectObject(lHdc, m_lhTextFnt)
    DrawText lHdc, m_sText, m_tRTxt
    SelectObject lHdc, lFntOld
    '/* icon
    DrawIcon lHdc, m_lIcon, m_tRIcn
    
    lWdc = GetDC(m_lToolHwnd)

    '/* shadow inset
    With tRect
        ExcludeClipRect lWdc, .Right - 3, .top, .Right, .top + 4
        ExcludeClipRect lWdc, .left, .Bottom, .left + 4, .Bottom - 3
    End With
    '/* draw to window
    With tRect
        If Not (m_lTransparency = -1) Then
            m_cRender.AlphaBlit lWdc, 0, 0, .Right, .Bottom, lHdc, 0, 0, .Right, .Bottom, m_lTransparency
        Else
            m_cRender.Blit lWdc, 0, 0, .Right, .Bottom, lHdc, 0, 0, SRCCOPY
        End If
    End With
    
    '/* cleanup
    SelectObject lHdc, lFntOld
    ReleaseDC m_lToolHwnd, lWdc
    
End Sub

Private Sub DrawText(ByVal lHdc As Long, _
                     ByVal sText As String, _
                     ByRef tRect As RECT)

Dim lFlags As Long

    SetBkMode lHdc, BM_TRANSPARENT
    SetTextColor lHdc, m_lForeColor
    If m_bMultiline Then
        lFlags = DT_WORDBREAK Or DT_LEFT
    Else
        lFlags = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS
    End If
    sText = sText & Chr$(0)
    If m_bIsNt Then
        If m_bFontRightLeading Then
            lFlags = lFlags Or DT_RTLREADING
        End If
        DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags
    Else
        DrawTextA lHdc, sText, -1, tRect, lFlags
    End If
    
End Sub

Private Sub DrawIcon(ByVal lHdc As Long, _
                     ByVal lIcon As Long, _
                     ByRef tRect As RECT)

    If Not (m_lImlHwnd = 0) Then
        If (lIcon > -1) Then
            With tRect
                If m_bGdiPlusLoaded Then
                    DrawAlphaIcon lIcon, lHdc, .left, .top
                Else
                    ImageList_DrawEx m_lImlHwnd, lIcon, lHdc, .left, .top, 0&, 0&, CLR_NONE, 0&, ILD_TRANSPARENT
                End If
            End With
        End If
    End If

End Sub

Private Sub FrameTip(ByVal lHdc As Long, _
                     ByRef tRect As RECT)

Dim lhBrush     As Long
Dim lhPen       As Long
Dim lhPenOld    As Long
Dim tPnt        As POINTAPI
Dim tRcpy       As RECT
    
    CopyRect tRcpy, tRect
    '/* draw 'shadow' lines
    With tRcpy
        '/* outer
        lhPen = CreatePen(0, 1, &H999999)
        lhPenOld = SelectObject(lHdc, lhPen)
        MoveToEx lHdc, .left, .Bottom - 1, tPnt
        LineTo lHdc, .Right - 1, .Bottom - 1
        LineTo lHdc, .Right - 1, .top
        SelectObject lHdc, lhPenOld
        DeleteObject lhPen
        '/* mid
        lhPen = CreatePen(0, 1, &H555555)
        lhPenOld = SelectObject(lHdc, lhPen)
        MoveToEx lHdc, .left, .Bottom - 2, tPnt
        LineTo lHdc, .Right - 2, .Bottom - 2
        LineTo lHdc, .Right - 2, .top
        SelectObject lHdc, lhPenOld
        DeleteObject lhPen
        '/* inner
        lhPen = CreatePen(0, 1, &H444444)
        lhPenOld = SelectObject(lHdc, lhPen)
        MoveToEx lHdc, .left, .Bottom - 3, tPnt
        LineTo lHdc, .Right - 3, .Bottom - 3
        LineTo lHdc, .Right - 3, .top
        SelectObject lHdc, lhPenOld
        DeleteObject lhPen
        .Right = .Right - 3
        .Bottom = .Bottom - 3
    End With
    '/* draw the frame
    lhBrush = CreateSolidBrush(&H808080)
    FrameRect lHdc, tRcpy, lhBrush
    DeleteObject lhBrush

End Sub

Private Sub PaintTip(ByVal lHdc As Long, _
                     ByRef tRect As RECT)

Dim lhBrush As Long

    lhBrush = CreateSolidBrush(m_lBackColor)
    FillRect lHdc, tRect, lhBrush
    DeleteObject lhBrush

End Sub


'> Fonts
'>>>>>>>>>>>>>>>>
Private Sub CreateFonts()

    DestroyFonts
    If (Font Is Nothing) Then
        DefaultFont m_oTitleFont, True
        DefaultFont m_oTextFont
    Else
        Set m_oTitleFont = m_oFont
        Set m_oTextFont = m_oFont
    End If
    m_lhTitleFnt = CreateFont(m_oTitleFont, True)
    m_lhTextFnt = CreateFont(m_oTextFont)
    
End Sub

Private Function CreateFont(ByVal oFont As StdFont, _
                            Optional ByVal bTitle As Boolean) As Long
'*/ change list font

Dim lChar       As Long
Dim lHdc        As Long
Dim tLF         As LOGFONT
Dim bteFont()   As Byte

    If Not (oFont Is Nothing) Then
        lHdc = CreateDc("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
        With tLF
            bteFont = StrConv(oFont.Name, vbFromUnicode)
            For lChar = 0 To UBound(bteFont)
                .lfFaceName(lChar) = bteFont(lChar)
            Next lChar
            If bTitle Then
                oFont.Bold = True
                oFont.Size = 9
            End If
            .lfHeight = -MulDiv(oFont.Size, GetDeviceCaps(lHdc, LOGPIXELSY), 72)
            .lfItalic = oFont.Italic
            .lfWeight = IIf(oFont.Bold, FW_BOLD, FW_NORMAL)
            .lfUnderline = oFont.Underline
            .lfStrikeOut = oFont.Strikethrough
            .lfCharSet = 4
            .lfQuality = LF_ANTIALIASED_QUALITY
        End With
        DeleteDC lHdc
        If m_bUseUnicode Then
            CreateFont = CreateFontIndirectW(tLF)
        Else
            CreateFont = CreateFontIndirectA(tLF)
        End If
    End If

End Function

Private Sub DefaultFont(ByRef oFont As StdFont, _
                        Optional ByVal bTitle As Boolean)

    Set oFont = New StdFont
    With oFont
        .Charset = 3
        .Name = "MS Sans Serif"
        .Weight = 400
        .Size = 8
        If bTitle Then
            .Bold = True
        End If
    End With

End Sub

Private Sub DestroyFonts()

    If Not (m_lhTitleFnt = 0) Then
        DeleteObject m_lhTitleFnt
        m_lhTitleFnt = 0
    End If
    If Not (m_oTitleFont Is Nothing) Then
        Set m_oTitleFont = Nothing
    End If
    
    If Not (m_lhTextFnt = 0) Then
        DeleteObject m_lhTextFnt
        m_lhTextFnt = 0
    End If
    If Not (m_oTextFont Is Nothing) Then
        Set m_oTextFont = Nothing
    End If

End Sub

Private Function FunctionExported(ByVal sFunction As String, _
                                  ByVal sModule As String) As Boolean

Dim lhMod       As Long
Dim bLoaded     As Boolean

    lhMod = GetModuleHandleA(sModule)
    If (lhMod = 0) Then
        lhMod = LoadLibraryA(sModule)
        If lhMod Then
            bLoaded = True
        End If
    End If
    If lhMod Then
        If (GetProcAddress(lhMod, sFunction)) Then
            FunctionExported = True
        End If
    End If
    If bLoaded Then
        FreeLibrary lhMod
    End If

End Function

Private Sub TrackMouse(ByVal lHwnd As Long)

Dim tMe As TRACKMOUSEEVENT_STRUCT

    With tMe
        .cbSize = Len(tMe)
        .dwFlags = TME_LEAVE
        .hwndTrack = lHwnd
    End With

    If m_bTrackUser32 Then
        TrackMouseEvent tMe
    Else
        TrackMouseEventComCtl tMe
    End If

End Sub

'> Subclass
'>>>>>>>>>>>>>>>>

Public Function StartTimer() As Boolean
'/* start display timer

    If Not m_bTimerActive Then
        m_TTSubclass.AddMessage m_lCtrlHwnd, WM_TIMER, MSG_BEFORE
        SetTimer m_lCtrlHwnd, 1&, 100&, 0&
        m_bTimerActive = True
    End If

End Function

Public Function StopTimer() As Boolean
'/* stop display timer

    If m_bTimerActive Then
        KillTimer m_lCtrlHwnd, 1&
        m_TTSubclass.DeleteMessage m_lCtrlHwnd, WM_TIMER, MSG_BEFORE
        DestroyToolTip
        m_bTimerActive = False
        m_lSafeTimer = 0
    End If

End Function


Private Sub MessageAttach()

    If Not m_TTSubclass Is Nothing Then
        If Not m_bIsActive Then
            With m_TTSubclass
                .Subclass m_lCtrlHwnd, Me
                If m_bAutoTips Then
                    .AddMessage m_lCtrlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                    .AddMessage m_lCtrlHwnd, WM_MOUSEMOVE, MSG_BEFORE
                    .AddMessage m_lCtrlHwnd, WM_MOUSELEAVE, MSG_BEFORE
                End If
            End With
            m_bIsActive = True
        End If
    End If
    
End Sub

Private Sub MessageDetach()

    If Not m_TTSubclass Is Nothing Then
        If Not (m_lCtrlHwnd = 0) Then
            If m_bIsActive Then
                With m_TTSubclass
                    If m_bAutoTips Then
                        .DeleteMessage m_lCtrlHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                        .DeleteMessage m_lCtrlHwnd, WM_MOUSEMOVE, MSG_BEFORE
                        .DeleteMessage m_lCtrlHwnd, WM_MOUSELEAVE, MSG_BEFORE
                    End If
                    .UnSubclass m_lCtrlHwnd
                End With
                m_bIsActive = False
            End If
        End If
    End If
    
End Sub

Private Sub GXISubclass_WndProc(ByVal bBefore As Boolean, _
                                bHandled As Boolean, _
                                lReturn As Long, _
                                ByVal lHwnd As Long, _
                                ByVal uMsg As eMsg, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long, _
                                lParamUser As Long)
    
    Select Case uMsg
    Case WM_LBUTTONDOWN
        StopTimer

    Case WM_MOUSEMOVE
        If Not m_bVisiting Then
            StartTimer
            m_bVisiting = True
        End If
        TrackMouse lHwnd
    
    Case WM_MOUSELEAVE
        m_bVisiting = False
        StopTimer
    
    Case WM_TIMER
        ToolTipTrack
    End Select
    
End Sub

Private Sub ToolTipTrack()

Dim tPnt As POINTAPI

    '/* stationary cursor
    GetCursorPos tPnt
    With tPnt
        If (.x = m_lLastX) Then
            If (.y = m_lLastY) Then
                m_lSafeTimer = m_lSafeTimer + 1
            Else
                m_lSafeTimer = 0
            End If
        Else
            m_lSafeTimer = 0
        End If
        m_lLastX = .x
        m_lLastY = .y
    End With
    '/* show tip
    If (m_lSafeTimer > (m_lDelayTime * 10)) Then
        If Not m_bShowing Then
            DrawTip
            m_bShowing = True
        End If
    End If
    '/* destroy tip
    If (m_lSafeTimer > ((m_lDelayTime + m_lVisibleTime) * 10)) Then
        If m_bShowing Then
            DestroyToolTip
            StopTimer
            m_bShowing = False
        End If
    End If
    
End Sub

'> Cleanup
'>>>>>>>>>>>>>>>>
Public Sub DestroyToolTip()

    If Not (m_lToolHwnd = 0) Then
        DestroyWindow m_lToolHwnd
        m_lToolHwnd = 0
        m_bShowing = False
    End If

End Sub

Public Sub Destroy()

    If Not (m_lToolHwnd = 0) Then
        StopTimer
        If m_bIsActive Then
            MessageDetach
        End If
        DestroyToolTip
        DestroyFonts
        If Not Font Is Nothing Then Set Font = Nothing
        If Not m_TTSubclass Is Nothing Then Set m_TTSubclass = Nothing
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        If Not m_cToolDc Is Nothing Then Set m_cToolDc = Nothing
    End If

End Sub

Private Sub Class_Terminate()
    GdiUnload
    Destroy
End Sub

