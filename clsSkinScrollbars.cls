VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSkinScrollbars"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'/* what a pain in the ass this was ;o(
'/* Thanks to Carles PV for the idea..
'/* http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=63527&lngWId=1

Implements GXISubclass


Private Const GWL_STYLE                         As Long = -16
Private Const GWL_EXSTYLE                       As Long = (-20)

Private Const OBJID_HSCROLL                     As Long = &HFFFFFFFA
Private Const OBJID_VSCROLL                     As Long = &HFFFFFFFB
Private Const OBJID_SIZEGRIP                    As Long = &HFFFFFFF9


Private Const SB_LINEDOWN                       As Long = 1
Private Const SB_LINELEFT                       As Long = 0
Private Const SB_LINERIGHT                      As Long = 1
Private Const SB_LINEUP                         As Long = 0

Private Const SBN_ENDSCROLL                     As Long = &H8

Private Const SS_OWNERDRAW                      As Long = &HD

Private Const SW_HIDE                           As Long = &H0
Private Const SW_NORMAL                         As Long = &H1

Private Const SWP_NOSIZE                        As Long = &H1
Private Const SWP_NOMOVE                        As Long = &H2
Private Const SWP_NOZORDER                      As Long = &H4
Private Const SWP_NOREDRAW                      As Long = &H8
Private Const SWP_NOACTIVATE                    As Long = &H10
Private Const SWP_FRAMECHANGED                  As Long = &H20
Private Const SWP_SHOWWINDOW                    As Long = &H40
Private Const SWP_HIDEWINDOW                    As Long = &H80
Private Const SWP_NOCOPYBITS                    As Long = &H100
Private Const SWP_NOOWNERZORDER                 As Long = &H200
Private Const SWP_NOSENDCHANGING                As Long = &H400

Private Const VK_LBUTTON                        As Long = &H1
Private Const VK_PAGEUP                         As Long = &H21
Private Const VK_PAGEDOWN                       As Long = &H22
Private Const VK_PGEND                          As Long = &H23
Private Const VK_PGHOME                         As Long = &H24


Private Const WS_EX_TOOLWINDOW                  As Long = &H80&
Private Const WS_EX_LEFTSCROLLBAR               As Long = &H4000

Private Const WS_CHILD                          As Long = &H40000000
Private Const WS_HSCROLL                        As Long = &H100000
Private Const WS_VSCROLL                        As Long = &H200000
Private Const WM_VSCROLL                        As Long = &H115
Private Const WM_HSCROLL                        As Long = &H114


'/* system metrics
Private Enum SYSTEM_METRICS
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
    SM_CXVSCROLL = 2
    SM_CYHSCROLL = 3
    SM_CYCAPTION = 4
    SM_CXBORDER = 5
    SM_CYBORDER = 6
    SM_CYVTHUMB = 9
    SM_CXHTHUMB = 10
    SM_CXICON = 11
    SM_CYICON = 12
    SM_CXCURSOR = 13
    SM_CYCURSOR = 14
    SM_CYMENU = 15
    SM_CXFULLSCREEN = 16
    SM_CYFULLSCREEN = 17
    SM_CYKANJIWINDOW = 18
    SM_MOUSEPRESENT = 19
    SM_CYVSCROLL = 20
    SM_CXHSCROLL = 21
    SM_CXMIN = 28
    SM_CYMIN = 29
    SM_CXSIZE = 30
    SM_CYSIZE = 31
    SM_CXFRAME = 32
    SM_CYFRAME = 33
    SM_CXMINTRACK = 34
    SM_CYMINTRACK = 35
    SM_CYSMCAPTION = 51
    SM_CXMINIMIZED = 57
    SM_CYMINIMIZED = 58
    SM_CXMAXTRACK = 59
    SM_CYMAXTRACK = 60
    SM_CXMAXIMIZED = 61
    SM_CYMAXIMIZED = 62
End Enum

Private Enum eSHitTest
    eshNone = 0&
    eshTop = 1&
    eshVThumb = 2&
    eshBottom = 3&
    eshLeft = 4&
    eshHThumb = 5&
    eshRight = 6&
End Enum

Public Enum EFSScrollDirection
    efsNone = 0&
    efsHorizontal = 1&
    efsVertical = 2&
End Enum

Public Enum EFSOrient
    efoHoriz = OBJID_HSCROLL
    efoVert = OBJID_VSCROLL
End Enum


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

Private Type BITMAP
    bmType                                      As Long
    bmWidth                                     As Long
    bmHeight                                    As Long
    bmWidthBytes                                As Long
    bmPlanes                                    As Integer
    bmBitsPixel                                 As Integer
    bmBits                                      As Long
End Type

Private Type SCROLLBARINFO
    cbSize                                      As Long
    rcScrollBar                                 As RECT
    dxyLineButton                               As Long
    xyThumbTop                                  As Long
    xyThumbBottom                               As Long
    Reserved                                    As Long
    rgstate(0 To 5)                             As Long
End Type

Private Type PAINTSTRUCT
    hdc                                         As Long
    fErase                                      As Boolean
    rcPaint                                     As RECT
    fRestore                                    As Boolean
    fIncUpdate                                  As Boolean
    rgbReserved(32)                             As Byte
End Type


Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
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

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                   ByVal nCount As Long, _
                                                                   lpObject As Any) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long

Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, _
                                                lpSourceRect As RECT) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function GetScrollBarInfo Lib "user32" (ByVal hwnd As Long, _
                                                        ByVal idObject As Long, _
                                                        psbi As SCROLLBARINFO) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

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

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSource As Any, _
                                                                     ByVal Length As Long)

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal nCmdShow As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

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

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
                                                ByVal ptX As Long, _
                                                ByVal ptY As Long) As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, _
                                                  lpPaint As PAINTSTRUCT) As Long

Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, _
                                                lpPaint As PAINTSTRUCT) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, _
                                                    ByVal y1 As Long, _
                                                    ByVal x2 As Long, _
                                                    ByVal y2 As Long) As Long

Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal hRgn As Long) As Long

Private Declare Function EraseRect Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, _
                                                                        lpRect As RECT, _
                                                                        ByVal bErase As Long) As Long

Private m_bInitialised                          As Boolean
Private m_bSkinScrollBar                        As Boolean
Private m_bUseScrollBarTheme                    As Boolean
Private m_bScrollBarActive                      As Boolean
Private m_bVisible                              As Boolean
Private m_bSelect                               As Boolean
Private m_bEnabled                              As Boolean
Private m_bPainting                             As Boolean
Private m_bRefresh                              As Boolean
Private m_bByPassHitTest                        As Boolean
Private m_lStyle                                As Long
Private m_lParentHwnd                           As Long
Private m_lScrollLuminence                      As Long
Private m_lThemeColor                           As Long
Private m_lScrollSkinStyle                      As Long
Private m_lHBtnWidth                            As Long
Private m_lVBtnWidth                            As Long
Private m_hWndParent                            As Long
Private m_lVXScroll                             As Long
Private m_lVYScroll                             As Long
Private m_lHXScroll                             As Long
Private m_lHYScroll                             As Long
Private m_lVYThumb                              As Long
Private m_lHXThumb                              As Long
Private m_lCtrlHnd                              As Long
Private m_lVScrollHnd                           As Long
Private m_lHScrollHnd                           As Long
Private m_lSizerHnd                             As Long
Private m_sngLuminence                          As Single
Private m_eScrollDirection                      As EFSScrollDirection
Private m_eHitTest                              As eSHitTest
Private m_pVTrack                               As StdPicture
Private m_pVThumb                               As StdPicture
Private m_pVBtDwn                               As StdPicture
Private m_pVBtUp                                As StdPicture
Private m_cVTrackDc                             As clsStoreDc
Private m_cVThumbDc                             As clsStoreDc
Private m_cBtDwnDc                              As clsStoreDc
Private m_cBtUpDc                               As clsStoreDc
Private m_pHTrack                               As StdPicture
Private m_pHThumb                               As StdPicture
Private m_pHBtLft                               As StdPicture
Private m_pHBtRgt                               As StdPicture
Private m_pSizer                                As StdPicture
Private m_cHTrackDc                             As clsStoreDc
Private m_cHThumbDc                             As clsStoreDc
Private m_cHBtLftDc                             As clsStoreDc
Private m_cHBtRgtDc                             As clsStoreDc
Private m_cSzrDc                                As clsStoreDc
Private m_cRender                               As clsRender
Private m_GXScrollbar                           As GXMSubclass

Private Sub Class_Initialize()

    m_bEnabled = True

End Sub

Friend Property Get ScrollBarActive() As Boolean
    ScrollBarActive = m_bScrollBarActive
End Property

Friend Property Let ScrollBarActive(ByVal PropVal As Boolean)
    m_bScrollBarActive = PropVal
End Property

Friend Property Get ByPassHitTest() As Boolean
    ByPassHitTest = m_bByPassHitTest
End Property

Friend Property Let ByPassHitTest(ByVal PropVal As Boolean)
    m_bByPassHitTest = PropVal
End Property

Friend Property Get ScrollDirection() As EFSScrollDirection
    ScrollDirection = m_eScrollDirection
End Property

Friend Property Get Enabled() As Boolean
    Enabled = m_bEnabled
End Property

Friend Property Let Enabled(ByVal PropVal As Boolean)
    m_bEnabled = PropVal
End Property

Private Property Get IHBtLft() As StdPicture
'/* scrollbar horizontal button left
    Set IHBtLft = m_pHBtLft
End Property

Private Property Set IHBtLft(PropVal As StdPicture)
    Set m_pHBtLft = PropVal
End Property

Private Property Get IHBtRgt() As StdPicture
'/* scrollbar horizontal button right
    Set IHBtRgt = m_pHBtRgt
End Property

Private Property Set IHBtRgt(PropVal As StdPicture)
    Set m_pHBtRgt = PropVal
End Property

Private Property Get IHThumb() As StdPicture
'/* scrollbar horizontal thumb
    Set IHThumb = m_pHThumb
End Property

Private Property Set IHThumb(PropVal As StdPicture)
    Set m_pHThumb = PropVal
End Property

Private Property Get IHTrack() As StdPicture
'/* scrollbar horizontal track
    Set IHTrack = m_pHTrack
End Property

Private Property Set IHTrack(PropVal As StdPicture)
    Set m_pHTrack = PropVal
End Property

Private Property Get ISizer() As StdPicture
'/* scrollbar sizer
    Set ISizer = m_pSizer
End Property

Private Property Set ISizer(PropVal As StdPicture)
    Set m_pSizer = PropVal
End Property

Private Property Get IVBtDwn() As StdPicture
'/* scrollbar vertical button down
    Set IVBtDwn = m_pVBtDwn
End Property

Private Property Set IVBtDwn(PropVal As StdPicture)
    Set m_pVBtDwn = PropVal
End Property

Private Property Get IVBtUp() As StdPicture
'/* scrollbar vertical button up
    Set IVBtUp = m_pVBtUp
End Property

Private Property Set IVBtUp(PropVal As StdPicture)
    Set m_pVBtUp = PropVal
End Property

Private Property Get IVThumb() As StdPicture
'/* scrollbar vertical thumb
    Set IVThumb = m_pVThumb
End Property

Private Property Set IVThumb(PropVal As StdPicture)
    Set m_pVThumb = PropVal
End Property

Private Property Get IVTrack() As StdPicture
'/* scrollbar vertical track
    Set IVTrack = m_pVTrack
End Property

Private Property Set IVTrack(PropVal As StdPicture)
    Set m_pVTrack = PropVal
End Property

Friend Property Get ScrollBarSkinStyle() As Long
    ScrollBarSkinStyle = m_lScrollSkinStyle
End Property

Friend Property Let ScrollBarSkinStyle(ByVal PropVal As Long)
    m_lScrollSkinStyle = PropVal
End Property

Friend Property Get ScrollLuminence() As Long
    ScrollLuminence = m_lScrollLuminence
End Property

Friend Property Let ScrollLuminence(PropVal As Long)
    m_lScrollLuminence = PropVal
End Property

Friend Property Get ScrollThemeColor() As Long
    ScrollThemeColor = m_lThemeColor
End Property

Friend Property Let ScrollThemeColor(ByVal PropVal As Long)
    m_lThemeColor = PropVal
End Property

Friend Property Get SkinScrollBar() As Boolean
    SkinScrollBar = m_bSkinScrollBar
End Property

Friend Property Let SkinScrollBar(ByVal PropVal As Boolean)
    m_bSkinScrollBar = PropVal
End Property

Friend Property Get UseScrollBarTheme() As Boolean
    UseScrollBarTheme = m_bUseScrollBarTheme
End Property

Friend Property Let UseScrollBarTheme(ByVal PropVal As Boolean)
    m_bUseScrollBarTheme = PropVal
End Property


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'> Functions
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Friend Sub ResetScrollBarSkin()

    m_bScrollBarActive = False
    CleanupSkin
    
End Sub

Private Function BuildImageList() As Boolean

On Error GoTo Handler

    Select Case m_lScrollSkinStyle
    '*/ azure
    Case 0
        '* / horz
        Set IHBtLft = LoadResPicture("AZURE-SCROLLBUTTONLEFT", vbResBitmap)
        Set IHBtRgt = LoadResPicture("AZURE-SCROLLBUTTONRIGHT", vbResBitmap)
        Set IHThumb = LoadResPicture("AZURE-SCROLLHORZTHUMB", vbResBitmap)
        Set IHTrack = LoadResPicture("AZURE-SCROLLHORZSHAFT", vbResBitmap)
        '/* vert
        Set IVBtDwn = LoadResPicture("AZURE-SCROLLBUTTONBOTTOM", vbResBitmap)
        Set IVBtUp = LoadResPicture("AZURE-SCROLLBUTTONTOP", vbResBitmap)
        Set IVThumb = LoadResPicture("AZURE-SCROLLVERTTHUMB", vbResBitmap)
        Set IVTrack = LoadResPicture("AZURE-SCROLLVERTSHAFT", vbResBitmap)
        '/* sizer
        Set ISizer = LoadResPicture("AZURE-SCROLLSIZER", vbResBitmap)
    '/* classic
    Case 1
        Set IHBtLft = LoadResPicture("CLASSIC-SCROLLBUTTONLEFT", vbResBitmap)
        Set IHBtRgt = LoadResPicture("CLASSIC-SCROLLBUTTONRIGHT", vbResBitmap)
        Set IHThumb = LoadResPicture("CLASSIC-SCROLLHORZTHUMB", vbResBitmap)
        Set IHTrack = LoadResPicture("CLASSIC-SCROLLHORZSHAFT", vbResBitmap)
        Set IVBtDwn = LoadResPicture("CLASSIC-SCROLLBUTTONBOTTOM", vbResBitmap)
        Set IVBtUp = LoadResPicture("CLASSIC-SCROLLBUTTONTOP", vbResBitmap)
        Set IVThumb = LoadResPicture("CLASSIC-SCROLLVERTTHUMB", vbResBitmap)
        Set IVTrack = LoadResPicture("CLASSIC-SCROLLVERTSHAFT", vbResBitmap)
        Set ISizer = LoadResPicture("CLASSIC-SCROLLSIZER", vbResBitmap)
    
    '/* gloss
    Case 2
        Set IHBtLft = LoadResPicture("GLOSS-SCROLLBUTTONLEFT", vbResBitmap)
        Set IHBtRgt = LoadResPicture("GLOSS-SCROLLBUTTONRIGHT", vbResBitmap)
        Set IHThumb = LoadResPicture("GLOSS-SCROLLHORZTHUMB", vbResBitmap)
        Set IHTrack = LoadResPicture("GLOSS-SCROLLHORZSHAFT", vbResBitmap)
        Set IVBtDwn = LoadResPicture("GLOSS-SCROLLBUTTONBOTTOM", vbResBitmap)
        Set IVBtUp = LoadResPicture("GLOSS-SCROLLBUTTONTOP", vbResBitmap)
        Set IVThumb = LoadResPicture("GLOSS-SCROLLVERTTHUMB", vbResBitmap)
        Set IVTrack = LoadResPicture("GLOSS-SCROLLVERTSHAFT", vbResBitmap)
        Set ISizer = LoadResPicture("GLOSS-SCROLLSIZER", vbResBitmap)
        
    '/* metal
    Case 3
        Set IHBtLft = LoadResPicture("METAL-SCROLLBUTTONLEFT", vbResBitmap)
        Set IHBtRgt = LoadResPicture("METAL-SCROLLBUTTONRIGHT", vbResBitmap)
        Set IHThumb = LoadResPicture("METAL-SCROLLHORZTHUMB", vbResBitmap)
        Set IHTrack = LoadResPicture("METAL-SCROLLHORZSHAFT", vbResBitmap)
        Set IVBtDwn = LoadResPicture("METAL-SCROLLBUTTONBOTTOM", vbResBitmap)
        Set IVBtUp = LoadResPicture("METAL-SCROLLBUTTONTOP", vbResBitmap)
        Set IVThumb = LoadResPicture("METAL-SCROLLVERTTHUMB", vbResBitmap)
        Set IVTrack = LoadResPicture("METAL-SCROLLVERTSHAFT", vbResBitmap)
        Set ISizer = LoadResPicture("METAL-SCROLLSIZER", vbResBitmap)
        
    '/* xp silver
    Case 4
        Set IHBtLft = LoadResPicture("XPS-SCROLLBUTTONLEFT", vbResBitmap)
        Set IHBtRgt = LoadResPicture("XPS-SCROLLBUTTONRIGHT", vbResBitmap)
        Set IHThumb = LoadResPicture("XPS-SCROLLHORZTHUMB", vbResBitmap)
        Set IHTrack = LoadResPicture("XPS-SCROLLHORZSHAFT", vbResBitmap)
        Set IVBtDwn = LoadResPicture("XPS-SCROLLBUTTONBOTTOM", vbResBitmap)
        Set IVBtUp = LoadResPicture("XPS-SCROLLBUTTONTOP", vbResBitmap)
        Set IVThumb = LoadResPicture("XPS-SCROLLVERTTHUMB", vbResBitmap)
        Set IVTrack = LoadResPicture("XPS-SCROLLVERTSHAFT", vbResBitmap)
        Set ISizer = LoadResPicture("XPS-SCROLLSIZER", vbResBitmap)
    
    '/* xp blue
    Case 5
        Set IHBtLft = LoadResPicture("XPB-SCROLLBUTTONLEFT", vbResBitmap)
        Set IHBtRgt = LoadResPicture("XPB-SCROLLBUTTONRIGHT", vbResBitmap)
        Set IHThumb = LoadResPicture("XPB-SCROLLHORZTHUMB", vbResBitmap)
        Set IHTrack = LoadResPicture("XPB-SCROLLHORZSHAFT", vbResBitmap)
        Set IVBtDwn = LoadResPicture("XPB-SCROLLBUTTONBOTTOM", vbResBitmap)
        Set IVBtUp = LoadResPicture("XPB-SCROLLBUTTONTOP", vbResBitmap)
        Set IVThumb = LoadResPicture("XPB-SCROLLVERTTHUMB", vbResBitmap)
        Set IVTrack = LoadResPicture("XPB-SCROLLVERTSHAFT", vbResBitmap)
        Set ISizer = LoadResPicture("XPB-SCROLLSIZER", vbResBitmap)
        
    '/* xp green
    Case 6
        Set IHBtLft = LoadResPicture("XPG-SCROLLBUTTONLEFT", vbResBitmap)
        Set IHBtRgt = LoadResPicture("XPG-SCROLLBUTTONRIGHT", vbResBitmap)
        Set IHThumb = LoadResPicture("XPG-SCROLLHORZTHUMB", vbResBitmap)
        Set IHTrack = LoadResPicture("XPG-SCROLLHORZSHAFT", vbResBitmap)
        Set IVBtDwn = LoadResPicture("XPG-SCROLLBUTTONBOTTOM", vbResBitmap)
        Set IVBtUp = LoadResPicture("XPG-SCROLLBUTTONTOP", vbResBitmap)
        Set IVThumb = LoadResPicture("XPG-SCROLLVERTTHUMB", vbResBitmap)
        Set IVTrack = LoadResPicture("XPG-SCROLLVERTSHAFT", vbResBitmap)
        Set ISizer = LoadResPicture("XPG-SCROLLSIZER", vbResBitmap)
        
    '/* vista arrow
    Case 7
        Set IHBtLft = LoadResPicture("VISTA-SCROLLBUTTONLEFT", vbResBitmap)
        Set IHBtRgt = LoadResPicture("VISTA-SCROLLBUTTONRIGHT", vbResBitmap)
        Set IHThumb = LoadResPicture("VISTA-SCROLLHORZTHUMB", vbResBitmap)
        Set IHTrack = LoadResPicture("VISTA-SCROLLHORZSHAFT", vbResBitmap)
        Set IVBtDwn = LoadResPicture("VISTA-SCROLLBUTTONBOTTOM", vbResBitmap)
        Set IVBtUp = LoadResPicture("VISTA-SCROLLBUTTONTOP", vbResBitmap)
        Set IVThumb = LoadResPicture("VISTA-SCROLLVERTTHUMB", vbResBitmap)
        Set IVTrack = LoadResPicture("VISTA-SCROLLVERTSHAFT", vbResBitmap)
        Set ISizer = LoadResPicture("VISTA-SCROLLSIZER", vbResBitmap)
        
    '/* silver
    Case 8
        Set IHBtLft = LoadResPicture("SILVER-SCROLLBUTTONLEFT", vbResBitmap)
        Set IHBtRgt = LoadResPicture("SILVER-SCROLLBUTTONRIGHT", vbResBitmap)
        Set IHThumb = LoadResPicture("SILVER-SCROLLHORZTHUMB", vbResBitmap)
        Set IHTrack = LoadResPicture("SILVER-SCROLLHORZSHAFT", vbResBitmap)
        Set IVBtDwn = LoadResPicture("SILVER-SCROLLBUTTONBOTTOM", vbResBitmap)
        Set IVBtUp = LoadResPicture("SILVER-SCROLLBUTTONTOP", vbResBitmap)
        Set IVThumb = LoadResPicture("SILVER-SCROLLVERTTHUMB", vbResBitmap)
        Set IVTrack = LoadResPicture("SILVER-SCROLLVERTSHAFT", vbResBitmap)
        Set ISizer = LoadResPicture("SILVER-SCROLLSIZER", vbResBitmap)
    End Select
    
    If m_bUseScrollBarTheme Then
        ScrollThemeSettings
    End If
    BuildImageList = True

Handler:
    On Error GoTo 0
    
End Function

Private Sub ScrollThemeSettings()

    Select Case m_lScrollLuminence
    Case 0
        m_sngLuminence = 0.2
    Case 1
        m_sngLuminence = 0.5
    Case 2
        m_sngLuminence = 1
    End Select
    
End Sub

Friend Sub LoadSkin(ByVal lCtrlHnd As Long, ByVal lParHnd As Long)

    m_lCtrlHnd = lCtrlHnd
    m_lParentHwnd = lParHnd
    If m_bSkinScrollBar Then
        If BuildImageList Then
            ScrollbarAttach
        End If
    End If
    
End Sub

Private Sub CreateScrollBars()

Dim bOnDesktop As Boolean

    m_hWndParent = GetParent(m_lParentHwnd)
    bOnDesktop = (m_hWndParent = GetDesktopWindow())

    If (m_lHScrollHnd = 0) Then
        m_lHScrollHnd = CreateWindowEx(-bOnDesktop * WS_EX_TOOLWINDOW, "Static", vbNullString, WS_CHILD Or SS_OWNERDRAW, 0&, 0&, 0&, 0&, m_hWndParent, 0&, App.hInstance, ByVal 0&)
        MaskShow m_lHScrollHnd, False
    End If
    If (m_lVScrollHnd = 0) Then
        m_lVScrollHnd = CreateWindowEx(-bOnDesktop * WS_EX_TOOLWINDOW, "Static", vbNullString, WS_CHILD Or SS_OWNERDRAW, 0&, 0&, 0&, 0&, m_hWndParent, 0&, App.hInstance, ByVal 0&)
        MaskShow m_lVScrollHnd, False
    End If
    If (m_lSizerHnd = 0) Then
        m_lSizerHnd = CreateWindowEx(-bOnDesktop * WS_EX_TOOLWINDOW, "Static", vbNullString, WS_CHILD Or SS_OWNERDRAW, 0&, 0&, 0&, 0&, m_hWndParent, 0&, App.hInstance, ByVal 0&)
        MaskShow m_lSizerHnd, False
    End If

End Sub

Private Sub DestroyScrollBars()

    If Not (m_lHScrollHnd = 0) Then
        DestroyWindow m_lHScrollHnd
        m_lHScrollHnd = 0
    End If
    If Not (m_lVScrollHnd = 0) Then
        DestroyWindow m_lVScrollHnd
        m_lVScrollHnd = 0
    End If
    If Not (m_lSizerHnd = 0) Then
        DestroyWindow m_lSizerHnd
        m_lSizerHnd = 0
    End If
    
End Sub

Private Sub GetScrollBarRect(ByVal eOrient As EFSOrient, _
                             ByRef tRect As RECT)

Dim tScBar As SCROLLBARINFO

    If Not (m_lCtrlHnd = 0) Then
        With tScBar
            '/* get bar params
            .cbSize = Len(tScBar)
            GetScrollBarInfo m_lCtrlHnd, eOrient, tScBar
            '/* copy size to working rect
            CopyRect tRect, .rcScrollBar
        End With
    End If

End Sub

Private Sub GetScrollBarThumb(ByVal eOrient As EFSOrient, _
                              ByRef tPnt As POINTAPI)

Dim tScBar As SCROLLBARINFO

    If Not (m_lCtrlHnd = 0) Then
        With tScBar
            '/* get bar params
            .cbSize = Len(tScBar)
            GetScrollBarInfo m_lCtrlHnd, eOrient, tScBar
            tPnt.y = .xyThumbTop
            tPnt.x = .xyThumbBottom
        End With
    End If

End Sub

Private Sub InitHorzImages()
'/* init hz image classes

Dim tBmp            As BITMAP

    '/* image size
    GetObject IHBtLft.handle, Len(tBmp), tBmp
    With tBmp
        m_lHBtnWidth = (.bmWidth / 2)
    End With

    '/* create the image classes
    Set m_cHTrackDc = New clsStoreDc
    With m_cHTrackDc
        .CreateFromPicture IHTrack
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
    Set m_cHThumbDc = New clsStoreDc
    With m_cHThumbDc
        .CreateFromPicture IHThumb
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
    Set m_cHBtLftDc = New clsStoreDc
    With m_cHBtLftDc
        .CreateFromPicture IHBtLft
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
    Set m_cHBtRgtDc = New clsStoreDc
    With m_cHBtRgtDc
        .CreateFromPicture IHBtRgt
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
End Sub

Private Sub InitSizerImage()
'/* sizer
    
    Set m_cSzrDc = New clsStoreDc
    With m_cSzrDc
        .CreateFromPicture ISizer
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
End Sub

Private Sub InitVertImages()
'/* init vt image classes

Dim tBmp As BITMAP

    GetObject IVBtDwn.handle, Len(tBmp), tBmp
    With tBmp
        m_lVBtnWidth = (.bmWidth / 2)
    End With

    '/* create the image classes
    Set m_cVTrackDc = New clsStoreDc
    With m_cVTrackDc
        .CreateFromPicture IVTrack
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
    Set m_cVThumbDc = New clsStoreDc
    With m_cVThumbDc
        .CreateFromPicture IVThumb
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
    Set m_cBtDwnDc = New clsStoreDc
    With m_cBtDwnDc
        .CreateFromPicture IVBtDwn
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
    Set m_cBtUpDc = New clsStoreDc
    With m_cBtUpDc
        .CreateFromPicture IVBtUp
        If m_bUseScrollBarTheme Then
            .ColorizeImage m_lThemeColor, m_sngLuminence
        End If
    End With
    
End Sub

Private Function LeftAligned() As Boolean

Dim lExStyle As Long

    lExStyle = GetWindowLong(m_lCtrlHnd, GWL_EXSTYLE)
    If ((lExStyle And WS_EX_LEFTSCROLLBAR) = WS_EX_LEFTSCROLLBAR) Then
        LeftAligned = True
    End If
    
End Function

Private Function LeftKeyState() As Boolean
'/* left button pressed state

    If ((GetKeyState(VK_LBUTTON) And &H80) > 1) Then
        LeftKeyState = True
    End If

End Function

Private Sub MaskShow(ByVal lHwnd As Long, _
                      ByVal bVisible As Boolean)
    
    If Not (lHwnd = 0) Then
        If bVisible Then
            ShowWindow lHwnd, SW_NORMAL
        Else
            ShowWindow lHwnd, SW_HIDE
        End If
    End If
    
End Sub

Private Sub ResetSCPos()

Dim lPos    As Long
Dim uPt     As POINTAPI
Dim tRect   As RECT

    If HasVertical Then
        If Not (m_lVScrollHnd = 0) Then
            GetScrollBarRect efoVert, tRect
            CopyMemory uPt, tRect, Len(uPt)
            ScreenToClient m_hWndParent, uPt
            With tRect
                OffsetRect tRect, -.left, -.top
                OffsetRect tRect, uPt.x, uPt.y
                SetWindowPos m_lVScrollHnd, 0&, .left, .top, (.Right - .left), (.Bottom - .top), SWP_SHOWWINDOW
            End With
        End If
    Else
        MaskShow m_lVScrollHnd, False
    End If
    
    If HasHorizontal Then
        If Not (m_lHScrollHnd = 0) Then
            GetScrollBarRect efoHoriz, tRect
            CopyMemory uPt, tRect, Len(uPt)
            ScreenToClient m_hWndParent, uPt
            With tRect
                OffsetRect tRect, -.left, -.top
                OffsetRect tRect, uPt.x, uPt.y
                SetWindowPos m_lHScrollHnd, 0&, .left, .top, (.Right - .left), (.Bottom - .top), SWP_SHOWWINDOW
            End With
        End If
    Else
        MaskShow m_lHScrollHnd, False
    End If
    
    If HasSizer Then
        If Not (m_lSizerHnd = 0) Then
            GetScrollBarRect efoHoriz, tRect
            CopyMemory uPt, tRect, Len(uPt)
            ScreenToClient m_hWndParent, uPt
            With tRect
                OffsetRect tRect, -.left, -.top
                OffsetRect tRect, uPt.x, uPt.y
                If LeftAligned Then
                    lPos = .left - m_lHXScroll
                Else
                    lPos = .Right
                End If
                If ((uPt.x + m_lHXScroll) > .Right) Then
                    SetWindowPos m_lSizerHnd, 0&, 0&, 0&, 0&, 0&, SWP_HIDEWINDOW
                Else
                    SetWindowPos m_lSizerHnd, 0&, lPos, .top, m_lHXScroll, m_lHYScroll, SWP_SHOWWINDOW
                End If
            End With
        End If
    Else
        MaskShow m_lSizerHnd, False
    End If

End Sub

Public Sub ScrollbarAttach()

    If Not (m_lCtrlHnd = 0) Then
        '/* instance render class
        Set m_cRender = New clsRender
        '/* store system sizes
        ScrollbarMetrics
        '/* create scrollbars
        CreateScrollBars
        '/* load image classes
        InitVertImages
        InitHorzImages
        InitSizerImage
        '/* destroy temp images
        DestroyImages
        m_bScrollBarActive = True
        '/* attach messages
        SkinAttach
    End If
    
End Sub

Private Function ScrollBarHitTest() As eSHitTest

Dim tPnt    As POINTAPI
Dim tPTb    As POINTAPI
Dim tRect   As RECT

    ScrollBarHitTest = eshNone
    If LeftKeyState Then
        '/* vertical
        GetScrollBarRect efoVert, tRect
        GetCursorPos tPnt
        With tPnt
            If Not (PtInRect(tRect, .x, .y) = 0) Then
                If .y >= (tRect.Bottom - m_lHXScroll) Then
                    ScrollBarHitTest = eshBottom
                ElseIf .y <= (tRect.top + m_lHXScroll) Then
                    ScrollBarHitTest = eshTop
                Else
                    GetScrollBarThumb efoVert, tPTb
                    With tRect
                        .top = .top + tPTb.y
                        .Bottom = .top + tPTb.x
                    End With
                    If Not (PtInRect(tRect, .x, .y) = 0) Then
                        ScrollBarHitTest = eshVThumb
                    End If
                End If
            Else
                If (m_eScrollDirection = efsVertical) Then
                    If m_eHitTest = eshVThumb Then
                        ScrollBarHitTest = eshVThumb
                    End If
                End If
            End If
        End With
        If (ScrollBarHitTest = eshNone) Then
            '/* horizontal
            GetScrollBarRect efoHoriz, tRect
            With tPnt
                If Not (PtInRect(tRect, .x, .y) = 0) Then
                    If .x >= (tRect.Right - m_lVYScroll) Then
                        ScrollBarHitTest = eshRight
                    ElseIf .x <= (tRect.left + m_lVYScroll) Then
                        ScrollBarHitTest = eshLeft
                    Else
                        GetScrollBarThumb efoHoriz, tPTb
                        With tRect
                            .left = .left + tPTb.y
                            .Right = .left + tPTb.x
                        End With
                        If Not (PtInRect(tRect, .x, .y) = 0) Then
                            ScrollBarHitTest = eshHThumb
                        End If
                    End If
                Else
                    If (m_eScrollDirection = efsHorizontal) Then
                        If m_eHitTest = eshHThumb Then
                            ScrollBarHitTest = eshHThumb
                        End If
                    End If
                End If
            End With
        End If
    End If
    If Not (ScrollBarHitTest = eshNone) Then
        m_eHitTest = ScrollBarHitTest
    End If
    
End Function

Private Sub ScrollbarMetrics()
'/* store scrollbar metrics

    m_lVXScroll = GetSystemMetrics(SM_CXVSCROLL)
    m_lVYScroll = GetSystemMetrics(SM_CYVSCROLL)
    m_lHXScroll = GetSystemMetrics(SM_CXHSCROLL)
    m_lHYScroll = GetSystemMetrics(SM_CYHSCROLL)
    m_lVYThumb = GetSystemMetrics(SM_CYVTHUMB)
    m_lHXThumb = GetSystemMetrics(SM_CXHTHUMB)
    
End Sub

Private Sub HorzBarPaint()

'/* paint horz scrollbar
Dim lHdc        As Long
Dim lDiff       As Long
Dim lHOffset    As Long
Dim lLOffset    As Long
Dim lXOffset    As Long
Dim lSclHdc     As Long
Dim lSzrDc      As Long
Dim lBmp        As Long
Dim lBmpOld     As Long
Dim lTmpDc      As Long
Dim lRgn        As Long
Dim tRect       As RECT
Dim tSzr        As RECT
Dim tPTb        As POINTAPI

On Error GoTo Handler

    GetClientRect m_lHScrollHnd, tRect
    '/* thumb size
    GetScrollBarThumb efoHoriz, tPTb
    With tPTb
        lDiff = .x - .y
    End With
    '/* backbuffer the image
    lTmpDc = m_cHTrackDc.hdc
    lHdc = CreateCompatibleDC(lTmpDc)
    With tRect
        lBmp = CreateCompatibleBitmap(lTmpDc, .Right, .Bottom)
    End With
    lBmpOld = SelectObject(lHdc, lBmp)
    lSclHdc = GetDC(m_lHScrollHnd)
    
    If Not m_bRefresh Then
        Select Case ScrollBarHitTest
        Case eshLeft
            lLOffset = 1
        Case eshRight
            lHOffset = 1
        Case eshHThumb
            If m_bEnabled Then
                lXOffset = 1
            End If
        End Select
    Else
        m_bRefresh = False
    End If
    If m_bByPassHitTest Then
        lXOffset = 0
    End If
    With tRect
        '/* build to temp dc
        With m_cHTrackDc
            m_cRender.Stretch lHdc, 0, 0, tRect.Right, m_lHXScroll, .hdc, 0, 0, .Width, .Height, SRCCOPY
        End With
        With m_cHBtLftDc
            m_cRender.Stretch lHdc, 0, 0, m_lHXScroll, m_lHYScroll, .hdc, (m_lHBtnWidth * lLOffset), 0, (.Width / 2), .Height, SRCCOPY
        End With
        With m_cHBtRgtDc
            m_cRender.Stretch lHdc, (tRect.Right - m_lHYScroll), tRect.top, m_lHXScroll, m_lHYScroll, .hdc, (m_lHBtnWidth * lHOffset), 0, (.Width / 2), .Height, SRCCOPY
        End With
        With m_cHThumbDc
            If (lDiff > .Width) Then
                m_cRender.Stretch lHdc, tPTb.y, 0, 6, m_lHYScroll, .hdc, ((.Width / 2) * lXOffset), 0, 6, .Height, SRCCOPY
                m_cRender.Stretch lHdc, (tPTb.y + 6), 0, lDiff - 8, m_lHYScroll, .hdc, (((.Width / 2) * lXOffset) + 6), 0, ((.Width / 2) - 12), .Height, SRCCOPY
                m_cRender.Stretch lHdc, (tPTb.x - 6), 0, 6, m_lHYScroll, .hdc, (((.Width / 2) * (lXOffset + 1)) - 6), 0, 6, .Height, SRCCOPY
            Else
                m_cRender.Stretch lHdc, tPTb.y, 0, lDiff, m_lHYScroll, .hdc, ((.Width / 2) * lXOffset), 0, (.Width / 2), .Height, SRCCOPY
            End If
            If (lDiff > 20) Then
                DrawGlyph False, lHdc, lXOffset, tPTb
            Else
                FrameScrollButton lHdc, False, tPTb
            End If
        End With
        '/* copy to dest
        m_cRender.Blit lSclHdc, 0, 0, .Right, .Bottom, lHdc, 0, 0, SRCCOPY
    End With
    
    '/* sizer
    If HasSizer Then
        GetWindowRect m_lSizerHnd, tRect
        GetWindowRect m_lParentHwnd, tSzr
        If (tRect.left > tSzr.left) Then
            lSzrDc = GetDC(m_lSizerHnd)
            With m_cSzrDc
                m_cRender.Stretch lSzrDc, 0, 0, m_lHYScroll, m_lVXScroll, .hdc, 0, 0, .Width, .Height, SRCCOPY
            End With
            ReleaseDC m_lSizerHnd, lSzrDc
        Else
            MaskShow m_lSizerHnd, False
        End If
    End If
    
    '/* cleanup
    SelectObject lHdc, lBmpOld
    DeleteObject lBmp
    DeleteDC lHdc
    ReleaseDC m_lVScrollHnd, lSclHdc
    
Handler:
    On Error GoTo 0

End Sub

Private Sub VertBarPaint()

'/* paint vert scrollbar
Dim lHdc        As Long
Dim lDiff       As Long
Dim lHOffset    As Long
Dim lLOffset    As Long
Dim lXOffset    As Long
Dim lSclHdc     As Long
Dim lBmp        As Long
Dim lBmpOld     As Long
Dim lTmpDc      As Long
Dim tRect       As RECT
Dim tPTb        As POINTAPI

On Error GoTo Handler

    GetClientRect m_lVScrollHnd, tRect
    '/* thumb size
    GetScrollBarThumb efoVert, tPTb
    With tPTb
        lDiff = .x - .y
    End With
    
    '/* backbuffer the image
    lTmpDc = m_cVTrackDc.hdc
    lHdc = CreateCompatibleDC(lTmpDc)
    
    With tRect
        lBmp = CreateCompatibleBitmap(lTmpDc, .Right, .Bottom)
    End With
    lBmpOld = SelectObject(lHdc, lBmp)
    lSclHdc = GetDC(m_lVScrollHnd)
    
    If Not m_bRefresh Then
        Select Case ScrollBarHitTest
        Case eshTop
            lLOffset = 1
        Case eshBottom
            lHOffset = 1
        Case eshVThumb
            If m_bEnabled Then
                lXOffset = 1
            End If
        End Select
    Else
        m_bRefresh = False
    End If
    If m_bByPassHitTest Then
        lXOffset = 0
    End If

    With tRect
        '/* build to temp dc
        With m_cVTrackDc
            m_cRender.Stretch lHdc, 0, 0, m_lVXScroll, tRect.Bottom, .hdc, 0, 0, .Width, .Height, SRCCOPY
        End With
        With m_cBtUpDc
            m_cRender.Stretch lHdc, 0, 0, m_lVXScroll, m_lVYScroll, .hdc, (m_lVBtnWidth * lLOffset), 0, (.Width / 2), .Height, SRCCOPY
        End With
        With m_cBtDwnDc
            m_cRender.Stretch lHdc, 0, (tRect.Bottom - m_lVYScroll), m_lVXScroll, m_lVYScroll, .hdc, (m_lVBtnWidth * lHOffset), 0, (.Width / 2), .Height, SRCCOPY
        End With
        With m_cVThumbDc
            If (lDiff > .Height) Then
                '/* thumb: top
                m_cRender.Stretch lHdc, 0, (tPTb.x - lDiff), m_lVXScroll, 6, .hdc, ((.Width / 2) * lXOffset), 0, (.Width / 2), 6, SRCCOPY
                '/* center
                m_cRender.Stretch lHdc, 0, (tPTb.x - (lDiff - 6)), m_lVXScroll, (lDiff - 12), .hdc, ((.Width / 2) * lXOffset), 6, (.Width / 2), (.Height - 12), SRCCOPY
                '/* bottom
                m_cRender.Stretch lHdc, 0, (tPTb.x - 6), m_lVXScroll, 6, .hdc, ((.Width / 2) * lXOffset), (.Height - 6), (.Width / 2), 6, SRCCOPY
            Else
                m_cRender.Stretch lHdc, 0, (tPTb.x - lDiff), m_lVXScroll, lDiff, .hdc, ((.Width / 2) * lXOffset), 0, (.Width / 2), .Height, SRCCOPY
            End If
            If (lDiff > 20) Then
                DrawGlyph True, lHdc, lXOffset, tPTb
            Else
                FrameScrollButton lHdc, True, tPTb
            End If
        End With
        '/* copy to dest
        m_cRender.Blit lSclHdc, 0, 0, .Right, .Bottom, lHdc, 0, 0, SRCCOPY
    End With
    
    '/* cleanup
    SelectObject lHdc, lBmpOld
    DeleteObject lBmp
    DeleteDC lHdc
    ReleaseDC m_lVScrollHnd, lSclHdc

Handler:
    On Error GoTo 0
    
End Sub

Private Sub FrameScrollButton(ByVal lHdc As Long, _
                              ByVal bVertical As Boolean, _
                              ByRef tPTb As POINTAPI)
'/* frame bottom/right of small scroll button

Dim lhPen       As Long
Dim lhPenOld    As Long
Dim lBaseClr    As Long
Dim tPnt        As POINTAPI

    If bVertical Then
        lBaseClr = GetPixel(m_cVThumbDc.hdc, 2&, 0&)
        lhPen = CreatePen(0&, 1&, lBaseClr)
        lhPenOld = SelectObject(lHdc, lhPen)
        MoveToEx lHdc, 0&, tPTb.x - 1, tPnt
        LineTo lHdc, m_lVXScroll, tPTb.x - 1
        MoveToEx lHdc, 0&, tPTb.y - 1, tPnt
        LineTo lHdc, m_lVXScroll, tPTb.y - 1
    Else
        lBaseClr = GetPixel(m_cVThumbDc.hdc, 0&, 2&)
        lhPen = CreatePen(0&, 1&, lBaseClr)
        lhPenOld = SelectObject(lHdc, lhPen)
        MoveToEx lHdc, tPTb.x - 1, 1, tPnt
        LineTo lHdc, tPTb.x - 1, m_lVYScroll - 1
        MoveToEx lHdc, tPTb.y - 1, 1, tPnt
        LineTo lHdc, tPTb.y - 1, m_lVYScroll - 1
    End If
    SelectObject lHdc, lhPenOld
    DeleteObject lhPen
    lhPen = 0

End Sub

Private Sub DrawGlyph(ByVal bVertical As Boolean, _
                      ByVal lHdc As Long, _
                      ByVal lOffset As Long, _
                      ByRef tPTb As POINTAPI)

Dim lBaseClr    As Long
Dim lCenter     As Long
Dim lhPen       As Long
Dim lhPenOld    As Long
Dim tPnt        As POINTAPI

    If bVertical Then
        lBaseClr = GetPixel(m_cVThumbDc.hdc, 3&, 3&)
        If (lOffset > 0) Then
            lBaseClr = m_cRender.BlendColor(lBaseClr, &HFDFDFD)
        Else
            lBaseClr = m_cRender.BlendColor(lBaseClr, &H0)
        End If
        lhPen = CreatePen(0&, 1&, lBaseClr)
        lhPenOld = SelectObject(lHdc, lhPen)
        '/* draw glyph lines
        With tPTb
            lCenter = .y + ((.x - .y) / 2)
        End With
        MoveToEx lHdc, 8&, (lCenter - 6), tPnt
        LineTo lHdc, (m_lVXScroll - 8), (lCenter - 6)
        MoveToEx lHdc, 6&, (lCenter - 3), tPnt
        LineTo lHdc, (m_lVXScroll - 6), (lCenter - 3)
        MoveToEx lHdc, 4&, lCenter, tPnt
        LineTo lHdc, (m_lVXScroll - 4), lCenter
        MoveToEx lHdc, 6&, (lCenter + 3), tPnt
        LineTo lHdc, (m_lVXScroll - 6), (lCenter + 3)
        MoveToEx lHdc, 8&, (lCenter + 6), tPnt
        LineTo lHdc, (m_lVXScroll - 8), (lCenter + 6)
        SelectObject lHdc, lhPenOld
        DeleteObject lhPen
        
    Else
        lBaseClr = GetPixel(m_cHThumbDc.hdc, 3&, 3&)
        If (lOffset > 0) Then
            lBaseClr = m_cRender.BlendColor(lBaseClr, &HFDFDFD)
        Else
            lBaseClr = m_cRender.BlendColor(lBaseClr, &H0)
        End If
        lhPen = CreatePen(0&, 1&, lBaseClr)
        lhPenOld = SelectObject(lHdc, lhPen)
        '/* draw glyph lines
        With tPTb
            lCenter = .y + ((.x - .y) / 2)
        End With
        MoveToEx lHdc, (lCenter - 6), 8&, tPnt
        LineTo lHdc, (lCenter - 6), (m_lVXScroll - 8)
        MoveToEx lHdc, (lCenter - 3), 6&, tPnt
        LineTo lHdc, (lCenter - 3), (m_lVXScroll - 6)
        MoveToEx lHdc, lCenter, 4&, tPnt
        LineTo lHdc, lCenter, (m_lVXScroll - 4)
        MoveToEx lHdc, (lCenter + 3), 6&, tPnt
        LineTo lHdc, (lCenter + 3), (m_lVXScroll - 6)
        MoveToEx lHdc, (lCenter + 6), 8&, tPnt
        LineTo lHdc, (lCenter + 6), (m_lVXScroll - 8)
        SelectObject lHdc, lhPenOld
        DeleteObject lhPen
    End If
    
End Sub

Public Sub Resize()
    ResetSCPos
End Sub

Friend Sub ScrollVertical(ByVal bDown As Boolean)

    If bDown Then
        SendMessageLong m_lCtrlHnd, WM_VSCROLL, SB_LINEDOWN, 0&
    Else
        SendMessageLong m_lCtrlHnd, WM_VSCROLL, SB_LINEUP, 0&
    End If
    
End Sub

Friend Sub ScrollHorizontal(ByVal bRight As Boolean)

    If bRight Then
        SendMessageLong m_lCtrlHnd, WM_HSCROLL, SB_LINERIGHT, 0&
    Else
        SendMessageLong m_lCtrlHnd, WM_HSCROLL, SB_LINELEFT, 0&
    End If
    
End Sub

Friend Function HasHorizontal() As Boolean

Dim lStyle As Long

    lStyle = GetWindowLong(m_lCtrlHnd, GWL_STYLE)
    If Not ((lStyle And WS_HSCROLL) = 0) Then
        HasHorizontal = True
    End If

End Function

Friend Function HasVertical() As Boolean

Dim lStyle As Long

    lStyle = GetWindowLong(m_lCtrlHnd, GWL_STYLE)
    If Not ((lStyle And WS_VSCROLL) = 0) Then
        HasVertical = True
    End If

End Function

Private Function HasSizer() As Boolean

Dim lStyle As Long

    lStyle = GetWindowLong(m_lCtrlHnd, GWL_STYLE)
    If Not ((lStyle And WS_HSCROLL) = 0) Then
        If Not ((lStyle And WS_VSCROLL) = 0) Then
            HasSizer = True
        End If
    End If
    
End Function

Friend Sub Refresh()

    m_bRefresh = True
    EraseSc m_lHScrollHnd
    EraseSc m_lVScrollHnd
    EraseSc m_lSizerHnd
    
End Sub

Public Sub EraseSc(ByVal lHwnd As Long)

Dim tRect As RECT

    If Not (lHwnd = 0) Then
        GetClientRect lHwnd, tRect
        EraseRect lHwnd, tRect, 0&
    End If

End Sub

Private Sub GridRefresh()

Dim tRect As RECT

    If Not (m_lParentHwnd = 0) Then
        GetClientRect m_lParentHwnd, tRect
        EraseRect m_lParentHwnd, tRect, 0&
    End If

End Sub

Private Sub ScrollbarStatus()

    If (IsWindowVisible(m_lCtrlHnd) = 0) Then
        MaskShow m_lHScrollHnd, False
        MaskShow m_lVScrollHnd, False
        MaskShow m_lSizerHnd, False
    Else
        MaskShow m_lHScrollHnd, HasHorizontal
        MaskShow m_lVScrollHnd, HasVertical
        MaskShow m_lSizerHnd, HasSizer
    End If

End Sub

Public Property Get Visible() As Boolean
    Visible = m_bVisible
End Property

Public Property Let Visible(PropVal As Boolean)

    If PropVal Then
        If Not m_bVisible Then
            ScrollbarStatus
            m_bScrollBarActive = True
            SendMessageLong m_lCtrlHnd, WM_PAINT, 0&, 0&
        End If
    Else
        MaskShow m_lHScrollHnd, False
        MaskShow m_lVScrollHnd, False
        MaskShow m_lSizerHnd, False
        m_bScrollBarActive = False
    End If
    m_bVisible = PropVal

End Property

Private Sub SkinAttach()

    If m_GXScrollbar Is Nothing Then
        Set m_GXScrollbar = New GXMSubclass
        With m_GXScrollbar
            If Not (m_lVScrollHnd = 0) Then
                .Subclass m_lVScrollHnd, Me
                .AddMessage m_lVScrollHnd, WM_PAINT, MSG_BEFORE
                .AddMessage m_lVScrollHnd, WM_ERASEBKGND, MSG_BEFORE
            End If
            If Not (m_lHScrollHnd = 0) Then
                .Subclass m_lHScrollHnd, Me
                .AddMessage m_lHScrollHnd, WM_PAINT, MSG_BEFORE
                .AddMessage m_lHScrollHnd, WM_ERASEBKGND, MSG_BEFORE
            End If
            If Not (m_lSizerHnd = 0) Then
                .Subclass m_lSizerHnd, Me
                .AddMessage m_lSizerHnd, WM_PAINT, MSG_BEFORE
                .AddMessage m_lSizerHnd, WM_ERASEBKGND, MSG_BEFORE
            End If
            .Subclass m_lCtrlHnd, Me
            .AddMessage m_lCtrlHnd, WM_STYLECHANGED, MSG_BEFORE
            .AddMessage m_lCtrlHnd, WM_WINDOWPOSCHANGED, MSG_BEFORE
            .AddMessage m_lCtrlHnd, WM_VSCROLL, MSG_BEFORE
            .AddMessage m_lCtrlHnd, WM_HSCROLL, MSG_BEFORE
            .AddMessage m_lCtrlHnd, WM_SIZE, MSG_BEFORE
            .AddMessage m_lCtrlHnd, WM_KEYUP, MSG_BEFORE
            .AddMessage m_lCtrlHnd, WM_KEYDOWN, MSG_BEFORE
            .AddMessage m_lCtrlHnd, WM_MOUSEWHEEL, MSG_AFTER
        End With
    End If
    
End Sub

Private Sub SkinDetach()

    If Not m_GXScrollbar Is Nothing Then
        With m_GXScrollbar
            If Not (m_lVScrollHnd = 0) Then
                .DeleteMessage m_lVScrollHnd, WM_PAINT, MSG_BEFORE
                .DeleteMessage m_lVScrollHnd, WM_ERASEBKGND, MSG_BEFORE
                .UnSubclass m_lVScrollHnd
            End If
            If Not (m_lHScrollHnd = 0) Then
                .DeleteMessage m_lHScrollHnd, WM_PAINT, MSG_BEFORE
                .DeleteMessage m_lHScrollHnd, WM_ERASEBKGND, MSG_BEFORE
                .UnSubclass m_lHScrollHnd
            End If
            If Not (m_lSizerHnd = 0) Then
                .DeleteMessage m_lSizerHnd, WM_PAINT, MSG_BEFORE
                .DeleteMessage m_lSizerHnd, WM_ERASEBKGND, MSG_BEFORE
                .UnSubclass m_lSizerHnd
            End If
            .DeleteMessage m_lCtrlHnd, WM_STYLECHANGED, MSG_BEFORE
            .DeleteMessage m_lCtrlHnd, WM_WINDOWPOSCHANGED, MSG_BEFORE
            .DeleteMessage m_lCtrlHnd, WM_VSCROLL, MSG_BEFORE
            .DeleteMessage m_lCtrlHnd, WM_HSCROLL, MSG_BEFORE
            .DeleteMessage m_lCtrlHnd, WM_SIZE, MSG_BEFORE
            .DeleteMessage m_lCtrlHnd, WM_KEYUP, MSG_BEFORE
            .DeleteMessage m_lCtrlHnd, WM_KEYDOWN, MSG_BEFORE
            .DeleteMessage m_lCtrlHnd, WM_MOUSEWHEEL, MSG_AFTER
            .UnSubclass m_lCtrlHnd
        End With
        Set m_GXScrollbar = Nothing
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

Dim lStyle  As Long
Dim tPaint  As PAINTSTRUCT

    ScrollbarStatus
    If Not m_bScrollBarActive Then
        Exit Sub
    End If

    Select Case uMsg
    '/* client paint
    Case WM_PAINT
    If Not m_bPainting Then
        m_bPainting = True
        BeginPaint lHwnd, tPaint
        If Not m_bSelect Then
            If HasHorizontal Then
                HorzBarPaint
            End If
            If HasVertical Then
                VertBarPaint
            End If
        Else
            m_bSelect = False
        End If
        EndPaint lHwnd, tPaint
        m_bPainting = False
        bHandled = True
    Else
        lReturn = m_GXScrollbar.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
    End If

    Case WM_KEYUP
        Select Case wParam
        Case VK_PGHOME, VK_PGEND
            If HasVertical Then
                VertBarPaint
            End If
        End Select
    
    Case WM_KEYDOWN
        Select Case wParam
        Case VK_PAGEDOWN, VK_PAGEUP
            If HasVertical Then
                VertBarPaint
            End If
        End Select

    '/* horizontal scroll
    Case WM_HSCROLL
        lReturn = m_GXScrollbar.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        If Not (wParam = SBN_ENDSCROLL) Then
            m_eScrollDirection = efsHorizontal
        Else
            m_eScrollDirection = efsNone
        End If
        m_bSelect = True
        HorzBarPaint
        bHandled = True
    
    '/* vertical Scroll
    Case WM_VSCROLL
        lReturn = m_GXScrollbar.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        If Not (wParam = SBN_ENDSCROLL) Then
            m_eScrollDirection = efsVertical
        Else
            m_eScrollDirection = efsNone
        End If
        m_bSelect = True
        VertBarPaint
        bHandled = True
    
    Case WM_MOUSEWHEEL
        If HasVertical Then
            VertBarPaint
        End If
        lReturn = m_GXScrollbar.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        bHandled = True
        
    Case WM_WINDOWPOSCHANGED, WM_STYLECHANGED
        ScrollbarStatus
    
    Case WM_SIZE
        lStyle = GetWindowLong(lHwnd, GWL_STYLE)
        If Not (lStyle = m_lStyle) Then
            ResetSCPos
            m_lStyle = lStyle
        End If

    Case WM_ERASEBKGND
        If HasHorizontal Then
            HorzBarPaint
        End If
        If HasVertical Then
            VertBarPaint
        End If
        lReturn = m_GXScrollbar.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        bHandled = True
    End Select

End Sub

Private Sub CleanUp()

    If Not (m_lCtrlHnd = 0) Then
        If m_bSkinScrollBar Then
            CleanupSkin
        End If
        m_lCtrlHnd = 0
        m_bInitialised = False
    End If
    
End Sub

Private Sub DestroyImages()

    If Not m_pVTrack Is Nothing Then Set m_pVTrack = Nothing
    If Not m_pVThumb Is Nothing Then Set m_pVThumb = Nothing
    If Not m_pVBtDwn Is Nothing Then Set m_pVBtDwn = Nothing
    If Not m_pVBtUp Is Nothing Then Set m_pVBtUp = Nothing
    If Not m_pHTrack Is Nothing Then Set m_pHTrack = Nothing
    If Not m_pHThumb Is Nothing Then Set m_pHThumb = Nothing
    If Not m_pHBtLft Is Nothing Then Set m_pHBtLft = Nothing
    If Not m_pHBtRgt Is Nothing Then Set m_pHBtRgt = Nothing
    If Not m_pSizer Is Nothing Then Set m_pSizer = Nothing
    
End Sub

Private Sub CleanupSkin()

    If Not (m_lCtrlHnd = 0) Then
        SkinDetach
        DestroyScrollBars
        DestroyImages
        If Not m_cVTrackDc Is Nothing Then Set m_cVTrackDc = Nothing
        If Not m_cVThumbDc Is Nothing Then Set m_cVThumbDc = Nothing
        If Not m_cBtDwnDc Is Nothing Then Set m_cBtDwnDc = Nothing
        If Not m_cBtUpDc Is Nothing Then Set m_cBtUpDc = Nothing
        If Not m_cHTrackDc Is Nothing Then Set m_cHTrackDc = Nothing
        If Not m_cHThumbDc Is Nothing Then Set m_cHThumbDc = Nothing
        If Not m_cHBtLftDc Is Nothing Then Set m_cHBtLftDc = Nothing
        If Not m_cHBtRgtDc Is Nothing Then Set m_cHBtRgtDc = Nothing
        If Not m_cSzrDc Is Nothing Then Set m_cSzrDc = Nothing
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        m_lCtrlHnd = 0
    End If

End Sub

Private Sub Class_Terminate()
    CleanUp
End Sub
