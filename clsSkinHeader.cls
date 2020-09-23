VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSkinHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements GXISubclass


Private Const FW_NORMAL                         As Long = 400
Private Const FW_BOLD                           As Long = 700

Private Const HD_CALGN                          As Long = &H3

Private Const HDI_WIDTH                         As Long = &H1
Private Const HDI_TEXT                          As Long = &H2
Private Const HDI_FORMAT                        As Long = &H4
Private Const HDI_IMAGE                         As Long = &H20
Private Const HDI_ORDER                         As Long = &H80

Private Const HDF_LEFT                          As Long = &H0
Private Const HDF_RIGHT                         As Long = &H1
Private Const HDF_CENTER                        As Long = &H2
Private Const HDF_RTLREADING                    As Long = &H4
Private Const HDF_IMAGE                         As Long = &H800
Private Const HDF_STRING                        As Long = &H4000
Private Const HDF_BITMAP_ON_RIGHT               As Long = &H1000

Private Const HDM_FIRST                         As Long = &H1200
Private Const HDM_GETITEMCOUNT                  As Long = (HDM_FIRST + 0)
Private Const HDM_GETITEMA                      As Long = (HDM_FIRST + 3)
Private Const HDM_SETITEMA                      As Long = (HDM_FIRST + 4)
Private Const HDM_HITTEST                       As Long = (HDM_FIRST + 6)
Private Const HDM_GETITEMRECT                   As Long = (HDM_FIRST + 7)
Private Const HDM_GETITEMW                      As Long = (HDM_FIRST + 11)
Private Const HDM_SETITEMW                      As Long = (HDM_FIRST + 12)

Private Const HHT_NOWHERE                       As Long = &H1
Private Const HHT_ONHEADER                      As Long = &H2
Private Const HHT_ONDIVIDER                     As Long = &H4
Private Const HHT_ONDIVOPEN                     As Long = &H8

Private Const ILD_TRANSPARENT                   As Long = &H1

Private Const LF_ANTIALIASED_QUALITY            As Long = 4
Private Const LF_CLEARTYPE_QUALITY              As Long = 5

Private Const LVM_FIRST                         As Long = &H1000
Private Const LVM_GETHEADER                     As Long = (LVM_FIRST + 31)

Private Const LOGPIXELSY                        As Long = 90

Private Const VER_PLATFORM_WIN32_NT             As Long = 2

Private Const VK_LBUTTON                        As Long = &H1


Private Enum CURSOR_RESOURCE
    OCR_NORMAL = 32512&
    OCR_IBEAM = 32513&
    OCR_WAIT = 32514&
    OCR_CROSS = 32515&
    OCR_UP = 32516&
    OCR_SIZE = 32640&
    OCR_ICON = 32641&
    OCR_SIZENWSE = 32642&
    OCR_SIZENESW = 32643&
    OCR_SIZEWE = 32644&
    OCR_SIZENS = 32645&
    OCR_SIZEALL = 32646&
    OCR_ICOCUR = 32647&
    OCR_NO = 32648&
    OCR_HAND = 32649&
    OCR_APPSTARTING = 32650&
End Enum

Public Enum EHdrHitTest
    HdrOffColumn = 0&
    HdrHitColumn = 1&
    HdrHitHorzSizer = 2&
    HdrHitVertSizer = 3&
    HdrHitFilter = 4&
End Enum

Public Enum EHdrTextAlign
    HdrTextALignLeft = HDF_LEFT
    HdrTextALignCentre = HDF_CENTER
    HdrTextALignRight = HDF_RIGHT
    HdrTextAlignRTLReading = HDF_RTLREADING
End Enum

Private Enum EColAlignFlags
    DT_TOP = &H0&
    DT_LEFT = &H0&
    DT_CENTER = &H1&
    DT_RIGHT = &H2&
    DT_VCENTER = &H4&
    DT_BOTTOM = &H8&
    DT_WORDBREAK = &H10&
    DT_SINGLELINE = &H20&
    DT_EXPANDTABS = &H40&
    DT_TABSTOP = &H80&
    DT_NOCLIP = &H100&
    DT_EXTERNALLEADING = &H200&
    DT_CALCRECT = &H400&
    DT_NOPREFIX = &H800&
    DT_INTERNAL = &H1000&
    DT_EDITCONTROL = &H2000&
    DT_PATH_ELLIPSIS = &H4000&
    DT_END_ELLIPSIS = &H8000&
    DT_MODIFYSTRING = &H10000
    DT_RTLREADING = &H20000
    DT_WORD_ELLIPSIS = &H40000
End Enum

Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1&
    TME_LEAVE = &H2&
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
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

Private Type HDITEMA
    Mask                                        As Long
    cxy                                         As Long
    pszText                                     As String
    hbm                                         As Long
    cchTextMax                                  As Long
    fmt                                         As Long
    lParam                                      As Long
    iImage                                      As Long
    iOrder                                      As Long
    type                                        As Long
    pvFilter                                    As Long
End Type

Private Type HDITEMW
    Mask                                        As Long
    cxy                                         As Long
    pszText                                     As Long
    hbm                                         As Long
    cchTextMax                                  As Long
    fmt                                         As Long
    lParam                                      As Long
    iImage                                      As Long
    iOrder                                      As Long
    type                                        As Long
    pvFilter                                    As Long
End Type

Private Type PAINTSTRUCT
    hdc                                         As Long
    fErase                                      As Boolean
    rcPaint                                     As RECT
    fRestore                                    As Boolean
    fIncUpdate                                  As Boolean
    rgbReserved(32)                             As Byte
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

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                                      As Long
    dwFlags                                     As TRACKMOUSEEVENT_FLAGS
    hwndTrack                                   As Long
    dwHoverTime                                 As Long
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

Private Type OSVERSIONINFO
    dwVersionInfoSize                           As Long
    dwMajorVersion                              As Long
    dwMinorVersion                              As Long
    dwBuildNumber                               As Long
    dwPlatformId                                As Long
    szCSDVersion(0 To 127)                      As Byte
End Type

Private Type HD_HITTESTINFO
    pt                                          As POINTAPI
    flags                                       As Long
    iItem                                       As Long
End Type

Private Declare Function SendMessageLongA Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                             ByVal wMsg As Long, _
                                                                             ByVal wParam As Long, _
                                                                             ByVal lParam As Long) As Long

Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    lParam As Any) As Long

Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    lParam As Any) As Long

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Function CreateFontIndirectA Lib "gdi32" (lpLogFont As LOGFONT) As Long

Private Declare Function CreateFontIndirectW Lib "gdi32" (lpLogFont As LOGFONT) As Long

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

Private Declare Function GetTextExtentPoint32A Lib "gdi32" (ByVal hdc As Long, _
                                                            ByVal lpsz As String, _
                                                            ByVal cbString As Long, _
                                                            lpSize As POINTAPI) As Long

Private Declare Function GetTextExtentPoint32W Lib "gdi32" (ByVal hdc As Long, _
                                                            ByVal lpsz As Long, _
                                                            ByVal cbString As Long, _
                                                            lpSize As POINTAPI) As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long

Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Long, _
                                                  ByVal lpString2 As Long) As Long


Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
                                                      lpPoint As POINTAPI) As Long

Private Declare Function CreateDc Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                 lpDeviceName As Any, _
                                                                 lpOutput As Any, _
                                                                 lpInitData As Any) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, _
                                                ByVal nNumerator As Long, _
                                                ByVal nDenominator As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nIndex As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal nBkMode As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Declare Function ImageList_Draw Lib "comctl32" (ByVal hIml As Long, _
                                                        ByVal I As Long, _
                                                        ByVal hdcDst As Long, _
                                                        ByVal x As Long, _
                                                        ByVal y As Long, _
                                                        ByVal fStyle As Long) As Long


Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, _
                                                    ByVal y1 As Long, _
                                                    ByVal x2 As Long, _
                                                    ByVal y2 As Long) As Long

Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal hRgn As Long) As Long

Private Declare Function BeginPaint Lib "user32" (ByVal hwnd As Long, _
                                                  lpPaint As PAINTSTRUCT) As Long

Private Declare Function EndPaint Lib "user32" (ByVal hwnd As Long, _
                                                lpPaint As PAINTSTRUCT) As Long

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                   ByVal nCount As Long, _
                                                                   lpObject As Any) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, _
                                                lpSourceRect As RECT) As Long

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function TrackMouseEventComCtl Lib "comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
                                                ByVal ptX As Long, _
                                                ByVal ptY As Long) As Long

Private Declare Function ImageList_GetIconSize Lib "comctl32" (ByVal hIml As Long, _
                                                               cx As Long, _
                                                               cy As Long) As Long

Private Declare Function UpdateWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, _
                                                      lpRect As Long, _
                                                      ByVal bErase As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long

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

Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Private Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function InflateRect Lib "user32" (lpRect As RECT, _
                                                   ByVal x As Long, _
                                                   ByVal y As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, _
                                                                      ByVal lpCursorId As CURSOR_RESOURCE) As Long

Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long

Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, _
                                                 lpRect As RECT, _
                                                 ByVal hBrush As Long) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
                                                     
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal x As Long, _
                                                ByVal y As Long, _
                                                ByVal crColor As Long) As Long

Private Declare Function EraseRect Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, _
                                                                        lpRect As RECT, _
                                                                        ByVal bErase As Long) As Long

Private m_bPainting                             As Boolean
Private m_bTrackUser32                          As Boolean
Private m_bSkinLoaded                           As Boolean
Private m_bUseHeaderTheme                       As Boolean
Private m_bHeaderActive                         As Boolean
Private m_bUseUnicode                           As Boolean
Private m_bDragState                            As Boolean
Private m_bIsNt                                 As Boolean
Private m_bIsXp                                 As Boolean
Private m_bColumnVerticalText                   As Boolean
Private m_bSortDescending                       As Boolean
Private m_bToolTips                             As Boolean
Private m_bTipGradient                          As Boolean
Private m_bTipXPColors                          As Boolean
Private m_bTipMultiline                         As Boolean
Private m_bHeaderSizeable                       As Boolean
Private m_bFilterLoaded                         As Boolean
Private m_bCustomCursors                        As Boolean
Private m_bHeaderFlat                           As Boolean
Private m_bHeaderFixedWidth                     As Boolean
Private m_bColumnFilters                        As Boolean
Private m_bClickInColumnSizer                   As Boolean
Private m_bClientClicked                        As Boolean
Private m_bFontRightLeading                     As Boolean
Private m_lTipOnColumn                          As Long
Private m_lhDragCursor                          As Long
Private m_lhNSSizeCursor                        As Long
Private m_lhNormalCursor                        As Long
Private m_lhWESizeCursor                        As Long
Private m_lhLockedCursor                        As Long
Private m_lColumnCountChange                    As Long
Private m_lTipVisibleTime                       As Long
Private m_lTipDelayTime                         As Long
Private m_lTipPosition                          As Long
Private m_lTipColor                             As Long
Private m_lTipOffsetColor                       As Long
Private m_lTipTransparency                      As Long
Private m_lTipForeColor                         As Long
Private m_lSelectedColumn                       As Long
Private m_lHeaderTextEffect                     As Long
Private m_lColumnSpace                          As Long
Private m_lColumnSorted                         As Long
Private m_lvFntDc                               As Long
Private m_lHeaderDc                             As Long
Private m_lHeaderLuminence                      As Long
Private m_lThemeColor                           As Long
Private m_lHeaderSkinStyle                      As Long
Private m_lTextForeColor                        As Long
Private m_lTextHighLite                         As Long
Private m_lTextPressed                          As Long
Private m_lHGHwnd                               As Long
Private m_lHdrHwnd                              As Long
Private m_lHeaderBmpWidth                       As Long
Private m_lHeaderBmpHeight                      As Long
Private m_lHeaderHeight                         As Long
Private m_lImageWidth                           As Long
Private m_lHdc()                                As Long
Private m_lBmp()                                As Long
Private m_lBmpOld()                             As Long
Private m_lCurrState                            As Long
Private m_lhFnt                                 As Long
Private m_lImlHdHndl                            As Long
Private m_sngLuminence                          As Single
Private m_bColumnFiltered()                     As Boolean
Private m_bColumnLocked()                       As Boolean
Private m_sToolTipHint()                        As String
Private m_eHeaderLastHitState                   As EHdrHitTest
Private m_oFont                                 As StdFont
Private m_oVtFont                               As StdFont
Private m_oTipFont                              As StdFont
Private m_pHeader                               As StdPicture
Private m_pArrowCursor                          As StdPicture
Private m_pLockedCursor                         As StdPicture
Private m_pNSSizeCursor                         As StdPicture
Private m_pWESizeCursor                         As StdPicture
Private m_pDragCursor                           As StdPicture
Private m_cRender                               As clsRender
Private m_cHeaderDc                             As clsStoreDc
Private m_cColumnToolTip                        As clsToolTip
Private m_GXHeader                              As GXMSubclass


Private Sub Class_Initialize()
    
    InitCommonControls
    VersionCheck
    m_bTrackUser32 = FunctionExported("TrackMouseEvent", "User32")
    '/* default font colors
    m_lTextForeColor = &H111111
    m_lTextHighLite = &H676767
    m_lTextPressed = &HDEDEDE
    '/* tool tip defaults
    m_lTipColor = GetSysColor(&H80000018 And &H1F)
    m_lTipDelayTime = 2
    m_bTipMultiline = True
    m_lTipTransparency = 180
    m_lTipVisibleTime = 3
    m_lColumnSorted = -1
    Set m_cRender = New clsRender
    CreateCursors
    ReDim m_bColumnFiltered(0)
    ReDim m_sToolTipHint(0)
    
End Sub

Private Function VersionCheck() As Boolean

Dim tVer    As OSVERSIONINFO

    tVer.dwVersionInfoSize = Len(tVer)
    GetVersionEx tVer
    m_bIsNt = ((tVer.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
    If (tVer.dwMajorVersion >= 5) Then
        m_bIsXp = True
    End If
    If Not m_bIsNt Then
        m_bUseUnicode = False
    End If
    VersionCheck = m_bIsNt

End Function


'> Properties
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Property Get ColumnAlign(ByVal lColumn As Long) As EHdrTextAlign
'*/ retieve a columns text alignment

Dim tHI As HDITEMA
Dim tHW As HDITEMW

    If Not (m_lHdrHwnd = 0) Then
        If m_bIsNt Then
            With tHW
                .Mask = HDI_FORMAT
                SendMessageW m_lHdrHwnd, HDM_GETITEMW, lColumn, tHW
                ColumnAlign = (HD_CALGN And .fmt)
            End With
        Else
            With tHI
                .Mask = HDI_FORMAT
                SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, tHI
                ColumnAlign = (HD_CALGN And .fmt)
            End With
        End If
    End If

End Property

Friend Property Get ColumnAtIndex(ByVal lIndex As Long) As Long

Dim lCt As Long

    For lCt = 0 To ColumnCount - 1
        If (ColumnIndex(lCt) = lIndex) Then
            ColumnAtIndex = lCt
            Exit For
        End If
    Next lCt

End Property

Private Property Get ColumnCount() As Long
'*/ retieve column count
    If Not (m_lHdrHwnd = 0) Then
        ColumnCount = SendMessageLongA(m_lHdrHwnd, HDM_GETITEMCOUNT, 0&, 0&)
    End If
End Property

Private Property Get ColumnIcon(ByVal lColumn As Long) As Long
'*/ retieve header icon index

Dim tHI As HDITEMA
Dim tHW As HDITEMW

    If Not (m_lImlHdHndl = 0) Then
        ColumnIcon = -1
        If m_bIsNt Then
            With tHW
                .Mask = HDI_FORMAT
                SendMessageW m_lHdrHwnd, HDM_GETITEMW, lColumn, tHW
                If (.fmt And HDF_IMAGE) = HDF_IMAGE Then
                    .Mask = HDI_IMAGE
                    SendMessageW m_lHdrHwnd, HDM_GETITEMW, lColumn, tHW
                    ColumnIcon = .iImage
                End If
            End With
        Else
            With tHI
                .Mask = HDI_FORMAT
                SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, tHI
                If (.fmt And HDF_IMAGE) = HDF_IMAGE Then
                    .Mask = HDI_IMAGE
                    SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, tHI
                    ColumnIcon = .iImage
                End If
            End With
        End If
    End If

End Property

Private Property Let ColumnIcon(ByVal lColumn As Long, _
                                ByVal lIcon As Long)
'*/ change header icon

Dim lAlign  As Long
Dim uHDI    As HDITEMA

    If Not (m_lImlHdHndl = 0) Then
        With uHDI
            .Mask = HDI_FORMAT
            If m_bIsNt Then
                SendMessageW m_lHdrHwnd, HDM_GETITEMW, lColumn, uHDI
            Else
                SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, uHDI
            End If
            lAlign = HD_CALGN And .fmt
            .iImage = lIcon
            .fmt = HDF_STRING Or lAlign Or HDF_IMAGE * -(lIcon > -1 And m_lImlHdHndl <> 0) Or HDF_BITMAP_ON_RIGHT
            .Mask = HDI_IMAGE * -(lIcon > -1) Or HDI_FORMAT
        End With
        If m_bIsNt Then
            SendMessageW m_lHdrHwnd, HDM_SETITEMW, lColumn, uHDI
        Else
            SendMessageA m_lHdrHwnd, HDM_SETITEMA, lColumn, uHDI
        End If
    End If

End Property

Friend Property Get ColumnIndex(ByVal lColumn As Long) As Long

Dim tHI     As HDITEMA
Dim tHW     As HDITEMW

    If m_bIsNt Then
        With tHW
            .Mask = HDI_ORDER
            If Not (SendMessageW(m_lHdrHwnd, HDM_GETITEMW, lColumn, tHW) = 0) Then
                ColumnIndex = .iOrder
            End If
        End With
    Else
        With tHI
            .Mask = HDI_ORDER
            If Not (SendMessageA(m_lHdrHwnd, HDM_GETITEMA, lColumn, tHI) = 0) Then
                ColumnIndex = .iOrder
            End If
        End With
    End If

End Property

Friend Property Get ColumnCountChange() As Long
    ColumnCountChange = m_lColumnCountChange
End Property

Friend Property Let ColumnCountChange(ByVal PropVal As Long)
    
    ReDim Preserve m_bColumnLocked(0 To PropVal)
    ReDim Preserve m_sToolTipHint(0 To PropVal)
    ReDim Preserve m_bColumnFiltered(0 To PropVal)
    m_lColumnCountChange = PropVal

End Property

Friend Property Get ColumnFilters() As Boolean
    ColumnFilters = m_bColumnFilters
End Property

Friend Property Let ColumnFilters(ByVal PropVal As Boolean)
    m_bColumnFilters = PropVal
End Property

Friend Property Get ColumnLocked(ByVal lColumn As Long) As Boolean

On Error GoTo Handler

    If Not (lColumn = -1) Then
        ColumnLocked = m_bColumnLocked(lColumn)
    End If

Handler:
    On Error GoTo 0

End Property

Friend Property Let ColumnLocked(ByVal lColumn As Long, _
                                 ByVal PropVal As Boolean)
    m_bColumnLocked(lColumn) = PropVal
End Property

Friend Property Get ColumnSortDescending() As Boolean
    ColumnSortDescending = m_bSortDescending
End Property

Friend Property Let ColumnSortDescending(ByVal PropVal As Boolean)
    m_bSortDescending = PropVal
End Property

Friend Property Get ColumnSorted() As Long
    ColumnSorted = m_lColumnSorted
End Property

Friend Property Let ColumnSorted(ByVal PropVal As Long)
    m_lColumnSorted = PropVal
End Property

Private Property Get ColumnText(ByVal lColumn As Long) As String
'*/ get a columns heading

Dim aText(261)  As Byte
Dim lLen        As Long
Dim sTemp       As String
Dim tHI         As HDITEMA
Dim tHW         As HDITEMW

    If Not (m_lHdrHwnd = 0) Then
        If m_bIsNt Then
            With tHW
                .pszText = VarPtr(aText(0))
                .cchTextMax = UBound(aText) + 1
                .Mask = HDI_TEXT
            End With
            SendMessageW m_lHdrHwnd, HDM_GETITEMW, lColumn, tHW
            ColumnText = PointerToString(tHW.pszText)
        Else
            With tHI
                sTemp = String(260, Chr$(0))
                .pszText = sTemp
                .cchTextMax = 261
                .Mask = HDI_TEXT
            End With
            SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, tHI
            ColumnText = tHI.pszText
            lLen = InStr(ColumnText, vbNullChar)
            If lLen Then
                ColumnText = left$(ColumnText, lLen - 1)
            End If
        End If
    End If

End Property

Friend Property Get ColumnVerticalText() As Boolean
    ColumnVerticalText = m_bColumnVerticalText
End Property

Friend Property Let ColumnVerticalText(ByVal PropVal As Boolean)

    If m_bHeaderActive Then
        If PropVal Then
            SetVerticalFont
        Else
            DestroyVericalFont
        End If
    End If
    m_bColumnVerticalText = PropVal
    
End Property

Private Property Get ColumnWidth(ByVal lColumn As Long) As Long
'*/ retrieve a columns length

Dim tHI As HDITEMA

    If Not (m_lHdrHwnd = 0) Then
        With tHI
            .Mask = HDI_WIDTH
            SendMessageA m_lHdrHwnd, HDM_GETITEMA, lColumn, tHI
            ColumnWidth = .cxy
        End With
    End If

End Property

Friend Property Get CustomCursors() As Boolean
    CustomCursors = m_bCustomCursors
End Property

Friend Property Let CustomCursors(ByVal PropVal As Boolean)
    m_bCustomCursors = PropVal
    CreateCursors
End Property

Friend Property Get DragState() As Boolean
    DragState = m_bDragState
End Property

Friend Property Let DragState(ByVal PropVal As Boolean)
    m_bDragState = PropVal
End Property

Friend Property Get ColumnFiltered(ByVal lColumn As Long) As Boolean

On Error GoTo Handler

    ColumnFiltered = m_bColumnFiltered(lColumn)

Handler:
    On Error GoTo 0

End Property

Friend Property Let ColumnFiltered(ByVal lColumn As Long, _
                                   ByVal PropVal As Boolean)

On Error GoTo Handler

    If (lColumn > UBound(m_bColumnFiltered)) Then
        ReDim m_bColumnFiltered(0 To (ColumnCount - 1))
    End If
    m_bColumnFiltered(lColumn) = PropVal

Handler:
    On Error GoTo 0

End Property

Friend Property Get ClickInColumnSizer() As Boolean
'*/ [get]
    ClickInColumnSizer = m_bClickInColumnSizer
End Property

Friend Property Let ClientClicked(ByVal PropVal As Boolean)
'*/ [get]
    m_bClientClicked = PropVal
End Property

Friend Property Get FilterLoaded() As Boolean
'*/ [get]
    FilterLoaded = m_bFilterLoaded
End Property

Friend Property Let FilterLoaded(ByVal PropVal As Boolean)
'*/ [let]
    m_bFilterLoaded = PropVal
End Property

Friend Property Get FontRightLeading() As Boolean
'/* [get] right align fonts
    FontRightLeading = m_bFontRightLeading
End Property

Friend Property Let FontRightLeading(ByVal PropvVal As Boolean)
'/* [let] right align fonts
    m_bFontRightLeading = PropvVal
End Property

Friend Property Get HeaderActive() As Boolean
    HeaderActive = m_bHeaderActive
End Property

Friend Property Let HeaderActive(ByVal PropVal As Boolean)
    m_bHeaderActive = PropVal
End Property

Friend Property Get HeaderFlat() As Boolean
    HeaderFlat = m_bHeaderFlat
End Property

Friend Property Let HeaderFlat(ByVal PropVal As Boolean)
    m_bHeaderFlat = PropVal
End Property

Friend Property Get HeaderFixedWidth() As Boolean
    HeaderFixedWidth = m_bHeaderFixedWidth
End Property

Friend Property Let HeaderFixedWidth(ByVal PropVal As Boolean)
    m_bHeaderFixedWidth = PropVal
End Property

Friend Property Get HeaderForeColor() As OLE_COLOR
    HeaderForeColor = m_lTextForeColor
End Property

Friend Property Let HeaderForeColor(ByVal PropVal As OLE_COLOR)
    m_lTextForeColor = PropVal
End Property

Friend Property Get HeaderHighLite() As OLE_COLOR
    HeaderHighLite = m_lTextHighLite
End Property

Friend Property Let HeaderHighLite(ByVal PropVal As OLE_COLOR)
    m_lTextHighLite = PropVal
End Property

Friend Property Get HeaderIml() As Long
    HeaderIml = m_lImlHdHndl
End Property

Friend Property Let HeaderIml(ByVal PropVal As Long)
    m_lImlHdHndl = PropVal
End Property

Friend Property Get HeaderLastHitState() As EHdrHitTest
    HeaderLastHitState = m_eHeaderLastHitState
End Property

Friend Property Let HeaderLastHitState(ByVal PropVal As EHdrHitTest)
    m_eHeaderLastHitState = PropVal
End Property

Friend Property Get HeaderLuminence() As Long
    HeaderLuminence = m_lHeaderLuminence
End Property

Friend Property Let HeaderLuminence(ByVal PropVal As Long)
    m_lHeaderLuminence = PropVal
End Property

Friend Property Get HeaderPressed() As OLE_COLOR
    HeaderPressed = m_lTextPressed
End Property

Friend Property Let HeaderPressed(ByVal PropVal As OLE_COLOR)
    m_lTextPressed = PropVal
End Property

Friend Property Get HeaderSizeable() As Boolean
    HeaderSizeable = m_bHeaderSizeable
End Property

Friend Property Let HeaderSizeable(ByVal PropVal As Boolean)
    m_bHeaderSizeable = PropVal
End Property

Friend Property Get HeaderSkinStyle() As Long
    HeaderSkinStyle = m_lHeaderSkinStyle
End Property

Friend Property Let HeaderSkinStyle(ByVal PropVal As Long)
    m_lHeaderSkinStyle = PropVal
End Property

Friend Property Get HeaderTextEffect() As Long
    HeaderTextEffect = m_lHeaderTextEffect
End Property

Friend Property Let HeaderTextEffect(ByVal PropVal As Long)
    m_lHeaderTextEffect = PropVal
End Property

Friend Property Get HeaderThemeColor() As Long
    HeaderThemeColor = m_lThemeColor
End Property

Friend Property Let HeaderThemeColor(ByVal PropVal As Long)
    m_lThemeColor = PropVal
End Property

Private Property Get IHeader() As StdPicture
'/* header image
    Set IHeader = m_pHeader
End Property

Private Property Set IHeader(ByVal PropVal As StdPicture)
    Set m_pHeader = PropVal
End Property

Friend Property Get ParentHwnd() As Long
    ParentHwnd = m_lHGHwnd
End Property

Friend Property Let ParentHwnd(ByVal PropVal As Long)
    m_lHGHwnd = PropVal
    If Not (m_lHGHwnd = 0) Then
        m_lHdrHwnd = HeaderHwnd
    End If
End Property

Friend Property Get TipColor() As Long
    TipColor = m_lTipColor
End Property

Friend Property Let TipColor(ByVal PropVal As Long)
    m_lTipColor = PropVal
End Property

Friend Property Get TipForeColor() As Long
    TipForeColor = m_lTipForeColor
End Property

Friend Property Let TipForeColor(ByVal PropVal As Long)
    m_lTipForeColor = PropVal
End Property

Friend Property Get TipDelayTime() As Long
    TipDelayTime = m_lTipDelayTime
End Property

Friend Property Let TipDelayTime(ByVal PropVal As Long)
    m_lTipDelayTime = PropVal
End Property

Friend Property Get TipFont() As StdFont
    TipFont = m_oTipFont
End Property

Friend Property Set TipFont(ByVal PropVal As StdFont)
    m_oTipFont = PropVal
End Property

Friend Property Get TipGradient() As Boolean
    TipGradient = m_bTipGradient
End Property

Friend Property Let TipGradient(ByVal PropVal As Boolean)
    m_bTipGradient = PropVal
End Property

Friend Property Get TipMultiline() As Boolean
    TipMultiline = m_bTipMultiline
End Property

Friend Property Let TipMultiline(ByVal PropVal As Boolean)
    m_bTipMultiline = PropVal
End Property

Friend Property Get TipOffsetColor() As Long
    TipOffsetColor = m_lTipOffsetColor
End Property

Friend Property Let TipOffsetColor(ByVal PropVal As Long)
    m_lTipOffsetColor = PropVal
End Property

Friend Property Get TipLoaded() As Boolean
    If Not (m_cColumnToolTip Is Nothing) Then
        TipLoaded = m_cColumnToolTip.Showing
    End If
End Property

Friend Property Get TipPosition() As Long
    TipPosition = m_lTipPosition
End Property

Friend Property Let TipPosition(ByVal PropVal As Long)
    m_lTipPosition = PropVal
End Property

Friend Property Get TipTransparency() As Long
    TipTransparency = m_lTipTransparency
End Property

Friend Property Let TipTransparency(ByVal PropVal As Long)
    m_lTipTransparency = PropVal
End Property

Friend Property Get TipVisibleTime() As Long
    TipVisibleTime = m_lTipVisibleTime
End Property

Friend Property Let TipVisibleTime(ByVal PropVal As Long)
    m_lTipVisibleTime = PropVal
End Property

Friend Property Get TipXPColors() As Boolean
    TipXPColors = m_bTipXPColors
End Property

Friend Property Let TipXPColors(ByVal PropVal As Boolean)
    m_bTipXPColors = PropVal
End Property

Friend Property Get TipHint(ByVal lColumn As Long) As String

On Error GoTo Handler

    TipHint = m_sToolTipHint(lColumn)

Handler:
    On Error GoTo 0

End Property

Friend Property Let TipHint(ByVal lColumn As Long, _
                            ByVal sHint As String)

On Error GoTo Handler

    If (lColumn > UBound(m_sToolTipHint)) Then
        ReDim m_sToolTipHint(0 To ColumnCount - 1)
    End If
    m_sToolTipHint(lColumn) = sHint

Handler:
    On Error GoTo 0

End Property

Friend Property Get ToolTips() As Boolean
    ToolTips = m_bToolTips
End Property

Friend Property Let ToolTips(ByVal PropVal As Boolean)
    
    If PropVal Then
        If (m_cColumnToolTip Is Nothing) Then
            Set m_cColumnToolTip = New clsToolTip
        End If
    Else
        If Not (m_cColumnToolTip Is Nothing) Then
            Set m_cColumnToolTip = Nothing
        End If
    End If
    m_bToolTips = PropVal
    
End Property

Friend Property Get UseHeaderTheme() As Boolean
    UseHeaderTheme = m_bUseHeaderTheme
End Property

Friend Property Let UseHeaderTheme(ByVal PropVal As Boolean)
    m_bUseHeaderTheme = PropVal
End Property

Friend Property Get UseUnicode() As Boolean
    UseUnicode = m_bUseUnicode
End Property

Friend Property Let UseUnicode(ByVal PropVal As Boolean)
    m_bUseUnicode = PropVal
End Property


'> Procedures
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub BackbufferDestroy()
'/* destroy backbuffer

Dim lCt As Long

On Error GoTo Handler

    For lCt = 0 To 3
        SelectObject m_lHdc(lCt), m_lBmpOld(lCt)
        DeleteObject m_lBmp(lCt)
        DeleteDC m_lHdc(lCt)
        m_lBmpOld(lCt) = 0
        m_lHdc(lCt) = 0
        m_lBmp(lCt) = 0
    Next lCt

Handler:

End Sub

Private Sub BackbufferImages()
'/* store state images

Dim lCt    As Long
Dim lTmpDc As Long

On Error GoTo Handler

    ReDim m_lHdc(3)
    ReDim m_lBmp(3)
    ReDim m_lBmpOld(3)

    lTmpDc = m_cHeaderDc.hdc
    For lCt = 0 To 3
        m_lHdc(lCt) = CreateCompatibleDC(lTmpDc)
        m_lBmp(lCt) = CreateCompatibleBitmap(lTmpDc, m_lHeaderBmpWidth, m_lHeaderBmpHeight)
        m_lBmpOld(lCt) = SelectObject(m_lHdc(lCt), m_lBmp(lCt))
        m_cRender.Blit m_lHdc(lCt), 0, 0, m_lHeaderBmpWidth, m_lHeaderBmpHeight, m_cHeaderDc.hdc, (m_lHeaderBmpWidth * lCt), 0, SRCCOPY
    Next lCt

Handler:

End Sub

Private Sub ColumnRender(ByVal lOffset As Long, _
                         ByVal lColumn As Long)

'/* render columns

Dim lCt         As Long
Dim lDrawDc     As Long
Dim lBmp        As Long
Dim lBmpOld     As Long
Dim lX          As Long
Dim lHdc        As Long
Dim lHdrClt     As Long
Dim tTemp       As RECT
Dim tRect       As RECT

    '/* get the dc
    lHdc = GetDC(m_lHdrHwnd)
    lDrawDc = CreateCompatibleDC(lHdc)
    '/* get coordinates
    GetClientRect m_lHdrHwnd, tRect

    '/* create the bmp
    With tRect
        lBmp = CreateCompatibleBitmap(lHdc, .Right, tRect.Bottom)
    End With
    lBmpOld = SelectObject(lDrawDc, lBmp)

    lX = 0
    For lCt = 0 To (ColumnCount - 1)
        If m_bIsNt Then
            SendMessageW m_lHdrHwnd, HDM_GETITEMRECT, lCt, tTemp
        Else
            SendMessageA m_lHdrHwnd, HDM_GETITEMRECT, lCt, tTemp
        End If
        With tTemp
            If (lCt = lColumn) Then
                lX = lOffset
            Else
                lX = 0
            End If
            If m_bDragState Then
                lX = 0
            End If
            m_lHeaderHeight = (.Bottom - .top)
            '/* left side
            m_cRender.Stretch lDrawDc, .left, 3, 3, (.Bottom - 6), m_lHdc(lX), 0, 3, 3, (m_lHeaderBmpHeight - 6), SRCCOPY
            '/* right side
            m_cRender.Stretch lDrawDc, (.Right - 3), (.top + 3), 3, (.Bottom - 6), m_lHdc(lX), (m_lHeaderBmpWidth - 3), 3, 3, (m_lHeaderBmpHeight - 6), SRCCOPY
            '/* top left corner
            m_cRender.Stretch lDrawDc, .left, 0, 3, 3, m_lHdc(lX), 0, 0, 3, 3, SRCCOPY
            '/* top
            m_cRender.Stretch lDrawDc, (.left + 3), 0, ((.Right - .left) - 3), 3, m_lHdc(lX), 3, 0, (m_lHeaderBmpWidth - 3), 3, SRCCOPY
            '/* bottom
            m_cRender.Stretch lDrawDc, (.left + 3), (.Bottom - 3), ((.Right - .left) - 3), 3, m_lHdc(lX), 3, (m_lHeaderBmpHeight - 3), (m_lHeaderBmpWidth - 3), 3, SRCCOPY
            '/* bottom left corner
            m_cRender.Stretch lDrawDc, .left, (.Bottom - 3), 3, 3, m_lHdc(lX), 0, (m_lHeaderBmpHeight - 3), 3, 3, SRCCOPY
            '/* center
            m_cRender.Stretch lDrawDc, (.left + 3), 3, ((.Right - .left) - 6), (.Bottom - 6), m_lHdc(lX), 3, 3, (m_lHeaderBmpWidth - 6), (m_lHeaderBmpHeight - 6), SRCCOPY
        End With
        DrawColumnIcon lCt, lDrawDc, tTemp
        DrawColumnText lDrawDc, lCt, lX, tTemp
    Next lCt

    '/* draw header end piece
    If m_bIsNt Then
        SendMessageW m_lHdrHwnd, HDM_GETITEMRECT, ColumnAtIndex(ColumnCount - 1), tTemp
    Else
        SendMessageA m_lHdrHwnd, HDM_GETITEMRECT, ColumnAtIndex(ColumnCount - 1), tTemp
    End If
    lHdrClt = tTemp.Right
    
    With tRect
        '/* left side
        m_cRender.Stretch lDrawDc, lHdrClt, 3, 3, (.Bottom - 6), m_lHdc(3), 0, 3, 3, (m_lHeaderBmpHeight - 6), SRCCOPY
        '/* right side
        m_cRender.Stretch lDrawDc, (.Right - 3), (.top + 3), 3, (.Bottom - 6), m_lHdc(3), (m_lHeaderBmpWidth - 3), 3, 3, (m_lHeaderBmpHeight - 6), SRCCOPY
        '/* top left corner
        m_cRender.Stretch lDrawDc, lHdrClt, 0, 3, 3, m_lHdc(0), 0, 0, 3, 3, SRCCOPY
        '/* top
        m_cRender.Stretch lDrawDc, (lHdrClt + 3), 0, (.Right - 3), 3, m_lHdc(3), 3, 0, (m_lHeaderBmpWidth - 3), 3, SRCCOPY
        '/* bottom
        m_cRender.Stretch lDrawDc, (lHdrClt + 3), (.Bottom - 3), (.Right - 3), 3, m_lHdc(3), 3, (m_lHeaderBmpHeight - 3), (m_lHeaderBmpWidth - 3), 3, SRCCOPY
        '/* bottom left corner
        m_cRender.Stretch lDrawDc, lHdrClt, (.Bottom - 3), 3, 3, m_lHdc(3), 0, (m_lHeaderBmpHeight - 3), 3, 3, SRCCOPY
        '/* center
        m_cRender.Stretch lDrawDc, (lHdrClt + 3), 3, (.Right - 6), (.Bottom - 6), m_lHdc(3), 3, 3, (m_lHeaderBmpWidth - 6), (m_lHeaderBmpHeight - 6), SRCCOPY
        'FloodFill lHdc, 0&, 0&, vbWhite
        '/* draw to screen
        'm_cRender.AlphaBlit lHdc, .Left, .Top, .Right, .bottom, lDrawDc, 0, 0, .Right, .bottom, 150
        m_cRender.Blit lHdc, .left, .top, .Right, .Bottom, lDrawDc, 0, 0, SRCCOPY
    End With

    SelectObject lDrawDc, lBmpOld
    DeleteObject lBmp
    lBmpOld = 0
    DeleteDC lDrawDc
    ReleaseDC m_lHdrHwnd, lHdc

End Sub

Private Sub CreateCursors()

On Error GoTo Handler

    DestroyCursors
    If m_bCustomCursors Then
        '/* why does this work in c++, but not in vb? poor vb6 res implementation..
        'm_lhNormalCursor = LoadCursorEx(App.hInstance, "CURSOR-ARROW")
        Set m_pArrowCursor = LoadResPicture("CURSOR-ARROW", vbResCursor)
        m_lhNormalCursor = m_pArrowCursor.handle
        Set m_pLockedCursor = LoadResPicture("CURSOR-LOCKED", vbResCursor)
        m_lhLockedCursor = m_pLockedCursor.handle
        Set m_pNSSizeCursor = LoadResPicture("CURSOR-SIZENS", vbResCursor)
        m_lhNSSizeCursor = m_pNSSizeCursor.handle
        Set m_pWESizeCursor = LoadResPicture("CURSOR-SIZEWE", vbResCursor)
        m_lhWESizeCursor = m_pWESizeCursor.handle
        Set m_pDragCursor = LoadResPicture("CURSOR-DRAG", vbResCursor)
        m_lhDragCursor = m_pDragCursor.handle
    Else
        m_lhNormalCursor = LoadCursor(0&, OCR_NORMAL)
        m_lhLockedCursor = LoadCursor(0&, OCR_NO)
        m_lhNSSizeCursor = LoadCursor(0&, OCR_SIZENS)
        m_lhWESizeCursor = LoadCursor(0&, OCR_SIZEWE)
    End If

On Error GoTo 0
Exit Sub

Handler:
    m_bCustomCursors = False

End Sub

Private Sub ColumnToolTip(ByVal lColumn As Long)

    If Not (lColumn = -1) Then
        If Not (LenB(m_sToolTipHint(lColumn)) = 0) Then
            With m_cColumnToolTip
                .CtrlHwnd = m_lHdrHwnd
                .ImlHwnd = m_lImlHdHndl
                .XPColors = m_bTipXPColors
                .BackColor = m_lTipColor
                .ForeColor = m_lTipForeColor
                .ColorOffset = m_lTipOffsetColor
                .DelayTime = m_lTipDelayTime
                .Gradient = m_bTipGradient
                .Icon = ColumnIcon(lColumn)
                .Multiline = m_bTipMultiline
                .Text = m_sToolTipHint(lColumn)
                .Title = ColumnText(lColumn)
                .Transparency = m_lTipTransparency
                .VisibleTime = m_lTipVisibleTime
                .ToolTipPosition = m_lTipPosition
                .Start
                .StartTimer
            End With
        End If
    End If

End Sub

Private Function CreateFont(ByVal oFont As StdFont, _
                            Optional ByVal bVertical As Boolean) As Long
'*/ change list font

Dim lChar       As Long
Dim lHdc        As Long
Dim uLF         As LOGFONT
Dim bteFont()   As Byte

On Error GoTo Handler

    If (oFont Is Nothing) Then
        Set oFont = New StdFont
        With oFont
            .Charset = 3
            .Name = "MS Sans Serif"
            .Weight = 400
            .Size = 9
        End With
    End If
    
    lHdc = CreateDc("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    With uLF
        bteFont = StrConv(oFont.Name, vbFromUnicode)
        For lChar = 0 To UBound(bteFont)
                .lfFaceName(lChar) = bteFont(lChar)
           Next lChar
         .lfHeight = -MulDiv(oFont.Size, GetDeviceCaps(lHdc, LOGPIXELSY), 72)
         .lfItalic = oFont.Italic
         .lfWeight = IIf(oFont.Bold, FW_BOLD, FW_NORMAL)
         .lfUnderline = oFont.Underline
         .lfStrikeOut = oFont.Strikethrough
         .lfCharSet = 4
         If m_bIsXp Then
            .lfQuality = LF_CLEARTYPE_QUALITY
         Else
            .lfQuality = LF_ANTIALIASED_QUALITY
         End If
         If bVertical Then
            .lfEscapement = 900
        End If
    End With
    DeleteDC lHdc
    
    If m_bUseUnicode Then
        CreateFont = CreateFontIndirectW(uLF)
    Else
        CreateFont = CreateFontIndirectA(uLF)
    End If

Handler:
    On Error GoTo 0

End Function

Private Sub DestroyCursors()

    If Not (m_lhNSSizeCursor = 0) Then
        DestroyCursor m_lhNSSizeCursor
        m_lhNSSizeCursor = 0
    End If
    If Not (m_lhNormalCursor = 0) Then
        DestroyCursor m_lhNormalCursor
        m_lhNormalCursor = 0
    End If
    If Not (m_lhWESizeCursor = 0) Then
        DestroyCursor m_lhWESizeCursor
        m_lhWESizeCursor = 0
    End If
    If Not (m_lhLockedCursor = 0) Then
        DestroyCursor m_lhLockedCursor
        m_lhLockedCursor = 0
    End If
    If Not (m_lhDragCursor = 0) Then
        DestroyCursor m_lhDragCursor
        m_lhDragCursor = 0
    End If
    
    If Not m_pArrowCursor Is Nothing Then Set m_pArrowCursor = Nothing
    If Not m_pLockedCursor Is Nothing Then Set m_pLockedCursor = Nothing
    If Not m_pNSSizeCursor Is Nothing Then Set m_pNSSizeCursor = Nothing
    If Not m_pWESizeCursor Is Nothing Then Set m_pWESizeCursor = Nothing
    If Not m_pDragCursor Is Nothing Then Set m_pDragCursor = Nothing

End Sub

Private Sub DestroyFont()

    If Not (m_lhFnt = 0) Then
        DeleteObject m_lhFnt
        m_lhFnt = 0
        If Not (m_oFont Is Nothing) Then
            Set m_oFont = Nothing
        End If
    End If

End Sub

Private Sub DestroyVericalFont()

    If Not (m_lvFntDc = 0) Then
        DeleteObject m_lvFntDc
        m_lvFntDc = 0
        If Not (m_oVtFont Is Nothing) Then
            Set m_oVtFont = Nothing
        End If
    End If

End Sub

Private Function DrawColumnIcon(ByVal lColumn As Long, _
                                ByVal lHdc As Long, _
                                ByRef tRect As RECT) As Boolean
'/* draw header icon

Dim lIndex  As Long
Dim lRgn    As Long

On Error GoTo Handler

    If Not (lColumn = -1) Then
        If Not m_bHeaderFlat Then
            If (lColumn = m_lColumnSorted) Then
                DrawSortArrow lHdc, tRect
            End If
        End If
        If m_bColumnFilters Then
            If m_bColumnFiltered(lColumn) Then
                DrawFilterButton lHdc, lColumn, tRect
            End If
        End If
        If Not (m_lImlHdHndl = 0) Then
            lIndex = ColumnIcon(lColumn)
            If (lIndex > -1) Then
                With tRect
                    lRgn = CreateRectRgn(.left, .top, .Right, .Bottom)
                    SelectClipRgn lHdc, lRgn
                    ImageList_Draw m_lImlHdHndl, lIndex, lHdc, (.left + 4), (.top + ((.Bottom - .top - m_lImageWidth) \ 2)), ILD_TRANSPARENT
                    SelectClipRgn lHdc, 0&
                    DeleteObject lRgn
                End With
                DrawColumnIcon = True
            End If
        End If
    End If

Handler:
    On Error GoTo 0
    
End Function

Private Sub DrawColumnText(ByVal lHdc As Long, _
                           ByVal lColumn As Long, _
                           ByVal lState As Long, _
                           ByRef tRect As RECT)

'/* draw column caption
Dim bIcon       As Boolean
Dim lFontOld    As Long
Dim lAlign      As Long
Dim lWidth      As Long
Dim lOffset     As Long
Dim lFlags      As Long
Dim sText       As String
Dim tPnt        As POINTAPI
Dim tRText      As RECT
Dim tShft       As RECT

On Error GoTo Handler

    sText = ColumnText(lColumn)
    If LenB(sText) = 0 Then Exit Sub
    lFontOld = SelectObject(lHdc, m_lhFnt)
    SetBkMode lHdc, 1

    '/* pressed
    If (lState = 1) Then
        SetTextColor lHdc, m_lTextPressed
        '/* over
    ElseIf lState = 2 Then
        SetTextColor lHdc, m_lTextHighLite
    Else
        SetTextColor lHdc, m_lTextForeColor
    End If

    '/* calculate text ellipses
    CopyRect tShft, tRect
    '/ calculate text size
    If m_bIsNt Then
        GetTextExtentPoint32W lHdc, StrPtr(sText), lstrlenW(StrPtr(sText)), tPnt
    Else
        GetTextExtentPoint32A lHdc, sText, Len(sText), tPnt
    End If
    
    '/* has icon
    bIcon = (ColumnIcon(lColumn) > -1)
    '/* text alignment
    lAlign = ColumnAlign(lColumn)

    LSet tRText = tRect
    lWidth = ColumnWidth(lColumn)
        
    If bIcon Then
        lOffset = (m_lImageWidth + 2)
    End If

    Select Case lAlign
    '/* left
    Case 0
        lOffset = (lOffset + 4)
        With tRText
            .left = (.left + lOffset)
        End With
    '/* right
    Case 1
        With tRText
            .left = (.left + (lWidth - tPnt.x) - 4)
        End With
    '/* center
    Case 2
        With tRText
            .left = (.left + (lWidth - tPnt.x) / 2)
        End With
    '/* left
    Case 3
        lOffset = lOffset + 4
        With tRText
            .left = (.left + lOffset)
        End With
    End Select

    With tRText
        '/* minimum x right align
        If (lAlign = 1) Then
            If (.left + lOffset) < (tRect.left + lOffset + tPnt.x) Then
                .left = (tRect.left + lOffset + 8)
            End If
        '/* minimum x center align
        ElseIf (lAlign = 2) Then
            If (.left + lOffset) < (tRect.left + lOffset + tPnt.x) Then
                .left = (tRect.left + lOffset) + 8
            End If
        '/* minimum x left align
        Else
            If ((.left - 2) < tRect.left) Then
                .left = (tRect.left)
            End If
        End If
    End With
    
    '/* draw elipses
    If m_bColumnVerticalText Then
        If ((tPnt.x + lOffset + 14) > lWidth) Then
            If ((tPnt.x + lOffset + 14) < m_lHeaderHeight) Then
                SelectObject lHdc, lFontOld
                lFontOld = 0
                lFontOld = SelectObject(lHdc, m_lvFntDc)
                lFlags = DT_LEFT Or DT_BOTTOM Or DT_SINGLELINE
                OffsetRect tRText, -3, 5
            Else
                tRText.Right = tRText.Right - 4
                lFlags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS
            End If
        Else
            lFlags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
        End If
    Else
        With tRText
            '/* test min size
            If ((tPnt.x + lOffset + 14) > lWidth) Then
                .Right = .Right - 4
                lFlags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS
            Else
                lFlags = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
            End If
        End With
    End If
    
    m_lColumnSpace = tRText.left + tPnt.x + lOffset

    If m_lHeaderTextEffect = 1 Then
        TextEmbossed lHdc, sText, lFlags, tRText
    ElseIf m_lHeaderTextEffect = 2 Then
        TextEngraved lHdc, sText, lFlags, tRText
    Else
        '/* draw normal text
        If m_bIsNt Then
            If m_bFontRightLeading Then
                lFlags = lFlags Or DT_RTLREADING
            End If
            DrawTextW lHdc, StrPtr(sText), -1, tRText, lFlags
        Else
            DrawTextA lHdc, sText, -1, tRText, lFlags
        End If
    End If
    
On Error GoTo 0

Handler:
    '/* delete font
    SelectObject lHdc, lFontOld

End Sub

Private Sub DrawFilterButton(ByVal lHdc As Long, _
                             ByVal lColumn As Long, _
                             ByRef tRect As RECT)

Dim bPressed    As Boolean
Dim lhPen       As Long
Dim lhPenOld    As Long
Dim lArwClr     As Long
Dim lRgn        As Long
Dim lBkClr      As Long
Dim lClrOst     As Long
Dim lHzOffset   As Long
Dim lhBrush     As Long
Dim tPnt        As POINTAPI
Dim tRFlt       As RECT

On Error GoTo Handler

    lHzOffset = ((tRect.Bottom \ 2) - 7)

    With tRFlt
        .left = (tRect.left + (tRect.Right - tRect.left) - 22)
        .Right = (.left + 14)
        .top = lHzOffset
        .Bottom = (lHzOffset + 14)
    End With

    '/* fill brush
    lBkClr = GetPixelColor(m_lHdc(0), 4, 4)
    lClrOst = m_cRender.BlendColor(&H999999, lBkClr)
    lArwClr = m_cRender.BlendColor(vbBlue, lBkClr)
    
    With tRFlt
        lRgn = CreateRectRgn((tRect.left + 2), (.top - 1), (.Right + 1), (.Bottom + 1))
        SelectClipRgn lHdc, lRgn
        If LeftKeyState Then
            If (HeaderHitTest = HdrHitFilter) Then
                If (lColumn = m_lSelectedColumn) Then
                    bPressed = True
                End If
            End If
        End If
        m_cRender.Gradient lHdc, .left, .Right - .left, .top, .Bottom - .top, lBkClr, lClrOst, Fill_Vertical, bPressed
        
        '/* draw outline
        lhPen = CreatePen(0&, 2&, lArwClr)
        lhPenOld = SelectObject(lHdc, lhPen)
        MoveToEx lHdc, (.left + 7), (.top + 2), tPnt
        LineTo lHdc, (.left + 7), (.top + 6)
        SelectObject lHdc, lhPenOld
        DeleteObject lhPen
        
        lhPen = CreatePen(0&, 1&, lArwClr)
        lhPenOld = SelectObject(lHdc, lhPen)
        MoveToEx lHdc, (.left + 3), (.top + 6), tPnt
        LineTo lHdc, (.Right - 3), (.top + 6)
        MoveToEx lHdc, (.left + 4), (.top + 8), tPnt
        LineTo lHdc, (.Right - 4), (.top + 8)
        MoveToEx lHdc, (.left + 5), (.top + 10), tPnt
        LineTo lHdc, (.Right - 5), (.top + 10)
    End With
    
    InflateRect tRFlt, 1, 1
    lClrOst = m_cRender.BlendColor(&H0, lBkClr)
    lhBrush = CreateSolidBrush(lClrOst)
    FrameRect lHdc, tRFlt, lhBrush
    
    '/* adjust rect
    With tRect
        .Right = (.Right - 22)
    End With
    
    '/* cleanup
    DeleteObject lhBrush
    SelectClipRgn lHdc, 0&
    DeleteObject lRgn
    SelectObject lHdc, lhPenOld
    DeleteObject lhPen

Handler:
    On Error GoTo 0

End Sub

Private Sub DrawSortArrow(ByVal lHdc As Long, _
                          ByRef tRect As RECT)

Dim lhPen       As Long
Dim lhPenOld    As Long
Dim lRgn        As Long
Dim lArwClr     As Long
Dim lVtCtr      As Long
Dim lhBrush     As Long
Dim lOldBrush   As Long
Dim tPnt        As POINTAPI

    '/* vertical center
    lVtCtr = (tRect.Bottom - tRect.top) / 2
    '/* arrow fill color
    lArwClr = GetPixelColor(lHdc, 0, 0)
    '/* outline pen
    lhPen = CreatePen(0&, 1&, &H808080)
    lhPenOld = SelectObject(lHdc, lhPen)
    '/* fill brush
    lhBrush = CreateSolidBrush(lArwClr)
    lOldBrush = SelectObject(lHdc, lhBrush)
    '/* track fill path
    BeginPath lHdc
    '/* draw outline
    With tRect
        lRgn = CreateRectRgn(.left, .top, .Right, .Bottom)
        SelectClipRgn lHdc, lRgn
        If Not m_bSortDescending Then
            MoveToEx lHdc, (.Right - 22), (lVtCtr + 3), tPnt
            LineTo lHdc, (.Right - 15), (lVtCtr - 4)
            LineTo lHdc, (.Right - 8), (lVtCtr + 3)
            LineTo lHdc, (.Right - 22), (lVtCtr + 3)
        Else
            MoveToEx lHdc, (.Right - 22), (lVtCtr - 4), tPnt
            LineTo lHdc, (.Right - 15), (lVtCtr + 3)
            LineTo lHdc, (.Right - 8), (lVtCtr - 4)
            LineTo lHdc, (.Right - 22), (lVtCtr - 4)
        End If
    End With
    '/* finished tracking
    EndPath lHdc
    '/* fill the shape
    StrokeAndFillPath lHdc
    '/* adjust rect
    With tRect
        .Right = (.Right - 22)
    End With
    '/* cleanup
    SelectClipRgn lHdc, 0&
    DeleteObject lRgn
    SelectObject lHdc, lOldBrush
    DeleteObject lhBrush
    SelectObject lHdc, lhPenOld
    DeleteObject lhPen

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

Private Function GetPixelColor(ByVal lHdc As Long, _
                               ByVal lX As Long, _
                               ByVal lY As Long) As Long

    GetPixelColor = GetPixel(lHdc, lX, lY)

End Function

Friend Property Get HeaderDc() As Long

    If Not (m_lHdrHwnd = 0) Then
        m_lHeaderDc = GetDC(m_lHdrHwnd)
        ReleaseDC m_lHdrHwnd, m_lHeaderDc
        HeaderDc = m_lHeaderDc
    End If
    
End Property

Private Function HeaderHwnd() As Long
'*/ return the column header handle

    If Not (m_lHGHwnd = 0) Then
        HeaderHwnd = SendMessageLongA(m_lHGHwnd, LVM_GETHEADER, 0&, 0&)
    End If
    
End Function

Friend Function HeaderHitTest() As EHdrHitTest

Dim lColumn     As Long
Dim lHzOffset   As Long
Dim tPnt        As POINTAPI
Dim tRect       As RECT
Dim tRcpy       As RECT
Dim tRBtn       As RECT
Dim tHdHit      As HD_HITTESTINFO

    GetCursorPos tPnt
    ScreenToClient m_lHdrHwnd, tPnt
    With tHdHit
        .pt = tPnt
        SendMessageA m_lHdrHwnd, HDM_HITTEST, 0&, tHdHit
        lColumn = .iItem
        If Not (lColumn = -1) Then
            m_lSelectedColumn = lColumn
            Select Case .flags
            Case HHT_NOWHERE
                HeaderHitTest = HdrOffColumn
            Case HHT_ONDIVIDER, HHT_ONDIVOPEN
                HeaderHitTest = HdrHitHorzSizer
            Case HHT_ONHEADER
                If m_bIsNt Then
                    SendMessageW m_lHdrHwnd, HDM_GETITEMRECT, lColumn, tRect
                Else
                    SendMessageA m_lHdrHwnd, HDM_GETITEMRECT, lColumn, tRect
                End If
                CopyRect tRcpy, tRect
                With tRect
                    .top = .Bottom - 8
                    .Bottom = .Bottom + 8
                End With
                If Not (PtInRect(tRect, tPnt.x, tPnt.y) = 0) Then
                    HeaderHitTest = HdrHitVertSizer
                Else
                    If m_bColumnFiltered(lColumn) Then
                        lHzOffset = ((tRcpy.Bottom \ 2) - 7)
                        With tRBtn
                            If (lColumn = m_lColumnSorted) Then
                                .left = tRcpy.left + ((tRcpy.Right - tRcpy.left) - 44)
                                .Right = (tRcpy.Right - 30)
                            Else
                                .left = ((tRcpy.left + (tRcpy.Right - tRcpy.left)) - 20)
                                .Right = (tRcpy.Right - 4)
                            End If
                            .top = lHzOffset
                            .Bottom = (lHzOffset + 14)
                        End With
                        With tPnt
                            If Not (PtInRect(tRBtn, .x, .y) = 0) Then
                                HeaderHitTest = HdrHitFilter
                            Else
                                HeaderHitTest = HdrHitColumn
                            End If
                        End With
                    Else
                        HeaderHitTest = HdrHitColumn
                    End If
                End If
            End Select
        Else
            HeaderHitTest = HdrOffColumn
        End If
    End With

End Function

Private Sub HeaderThemeSettings()

    Select Case m_lHeaderLuminence
    Case 0
        m_sngLuminence = 0.3
    Case 1
        m_sngLuminence = 0.6
    Case 2
        m_sngLuminence = 1
    End Select

End Sub

Private Function LoadHeaderImage() As Boolean

On Error GoTo Handler

    Select Case m_lHeaderSkinStyle
    '/* azure
    Case 0
        Set IHeader = LoadResPicture("AZURE-HEADER", vbResBitmap)
    '/* classic
    Case 1
        Set IHeader = LoadResPicture("CLASSIC-HEADER", vbResBitmap)
    '/* gloss
    Case 2
        Set IHeader = LoadResPicture("GLOSS-HEADER", vbResBitmap)
    '/* metal
    Case 3
        Set IHeader = LoadResPicture("METAL-HEADER", vbResBitmap)
    '/* xp silver
    Case 4, 5, 6
        Set IHeader = LoadResPicture("XP-HEADER", vbResBitmap)
    '/* vista arrow
    Case 7
        Set IHeader = LoadResPicture("VISTA-HEADER", vbResBitmap)
    '/* silver
    Case 8
        Set IHeader = LoadResPicture("SILVER-HEADER", vbResBitmap)
    End Select
    If m_bUseHeaderTheme Then
        HeaderThemeSettings
    End If
    LoadHeaderImage = True

Handler:
On Error GoTo 0

End Function

Private Function LeftKeyState() As Boolean

    If ((GetKeyState(VK_LBUTTON) And &H80) > 1) Then
        LeftKeyState = True
    End If

End Function

Friend Sub LoadSkin()

Dim lX  As Long
Dim lY  As Long

    SkinHeaders
    '/* icon size
    If Not (m_lImlHdHndl = 0) Then
        ImageList_GetIconSize m_lImlHdHndl, lX, lY
        If (lX = 0) Then
            m_lImageWidth = 16
        End If
        m_lImageWidth = lX
    Else
        m_lImageWidth = 16
    End If
    If m_bColumnVerticalText Then
        SetVerticalFont
    Else
        DestroyVericalFont
    End If
    m_bHeaderActive = True
    
End Sub

Private Function PointerToString(ByVal lpString As Long) As String
'/* get string from pointer

Dim lLen As Long

    If Not (lpString = 0) Then
        lLen = lstrlenW(ByVal lpString)
        If (lLen > 0) Then
            '/* allocate string with nLen chars
            PointerToString = String$(lLen, 0)
            lstrcpyW StrPtr(PointerToString), lpString
        End If
    End If

End Function

Friend Sub Refresh(ByVal lColumn As Long)

Dim lCt As Long

    If (lColumn = -1) Then
        For lCt = 0 To (ColumnCount - 1)
            InvalidateRect m_lHdrHwnd, ByVal 0&, 1&
            UpdateWindow m_lHdrHwnd
        Next lCt
    Else
        m_lCurrState = lColumn
        SendMessageLongA m_lHdrHwnd, WM_PAINT, 0&, 0&
    End If

End Sub

Friend Function ResetHeaderSkin() As Boolean

    If Not m_cHeaderDc Is Nothing Then Set m_cHeaderDc = Nothing
    If Not IHeader Is Nothing Then Set IHeader = Nothing
    If Not m_oFont Is Nothing Then Set m_oFont = Nothing
    If Not m_oVtFont Is Nothing Then Set m_oVtFont = Nothing
    m_bColumnVerticalText = False
    
    If m_bHeaderActive Then
        BackbufferDestroy
        m_bHeaderActive = False
    End If

End Function

Private Sub SkinHeaders()

Dim tBmp As BITMAP

    If Not (m_lHdrHwnd = 0) Then
        If m_bSkinLoaded Then
            If LoadHeaderImage Then
                Set m_cHeaderDc = New clsStoreDc
                With m_cHeaderDc
                    .CreateFromPicture IHeader
                    If m_bUseHeaderTheme Then
                        .ColorizeImage m_lThemeColor, m_sngLuminence
                    End If
                End With
                '/* default sizes
                GetObject IHeader.handle, Len(tBmp), tBmp
                With tBmp
                    m_lHeaderBmpWidth = (.bmWidth / 4)
                    m_lHeaderBmpHeight = .bmHeight
                End With
                BackbufferImages
            End If
        Else
            Set m_GXHeader = New GXMSubclass
            If LoadHeaderImage Then
                Set m_cHeaderDc = New clsStoreDc
                With m_cHeaderDc
                    .CreateFromPicture IHeader
                    If m_bUseHeaderTheme Then
                        .ColorizeImage m_lThemeColor, m_sngLuminence
                    End If
                End With
                '/* default sizes
                GetObject IHeader.handle, Len(tBmp), tBmp
                With tBmp
                    m_lHeaderBmpWidth = (.bmWidth / 4)
                    m_lHeaderBmpHeight = .bmHeight
                End With
                BackbufferImages
                SkinAttach
                m_bSkinLoaded = True
            End If
        End If
        DestroyHeaderImage
    End If

End Sub

Friend Sub SetNSSizingCursor()
    SetCursor m_lhNSSizeCursor
End Sub

Friend Sub SetNormalCursor()
    SetCursor m_lhNormalCursor
End Sub

Friend Sub SetWESizingCursor()
    SetCursor m_lhWESizeCursor
End Sub

Friend Sub SetLockedCursor()
    SetCursor m_lhLockedCursor
End Sub

Friend Sub SetDragCursor()
    SetCursor m_lhDragCursor
End Sub

Friend Sub SetFont(ByVal oFont As StdFont)

    DestroyFont
    m_lhFnt = CreateFont(oFont)
    
End Sub

Private Sub SetVerticalFont()

    DestroyVericalFont
    Set m_oVtFont = New StdFont
    m_oVtFont.Name = "Arial"
    m_oVtFont.Size = 9
    m_lvFntDc = CreateFont(m_oVtFont, True)

End Sub

Private Sub TextEmbossed(ByVal lHdc As Long, _
                         ByVal sText As String, _
                         ByVal lFlags As Long, _
                         ByRef tRect As RECT)
'/* engraved text

Dim lFrClr As Long

    If (m_lCurrState = 1) Then
        lFrClr = m_lTextPressed
    ElseIf (m_lCurrState = 2) Then
        lFrClr = m_lTextHighLite
    Else
        lFrClr = m_lTextForeColor
    End If
    
    If m_bIsNt Then
        SetTextColor lHdc, vbWhite
        DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags
        OffsetRect tRect, 1, 1
        SetTextColor lHdc, &H808080
        DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags
        OffsetRect tRect, 1, 0
        SetTextColor lHdc, lFrClr
        DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags
    Else
        SetTextColor lHdc, &HCCCCCC
        DrawTextA lHdc, sText, -1, tRect, lFlags
        OffsetRect tRect, 1, 1
        SetTextColor lHdc, &H808080
        DrawTextA lHdc, sText, -1, tRect, lFlags
        OffsetRect tRect, 1, 0
        SetTextColor lHdc, lFrClr
        DrawTextA lHdc, sText, -1, tRect, lFlags
    End If

End Sub

Private Sub TextEngraved(ByVal lHdc As Long, _
                         ByVal sText As String, _
                         ByVal lFlags As Long, _
                         ByRef tRect As RECT)
'/* embossed text

Dim lFrClr As Long

    If (m_lCurrState = 1) Then
        lFrClr = m_lTextPressed
    ElseIf (m_lCurrState = 2) Then
        lFrClr = m_lTextHighLite
    Else
        lFrClr = m_lTextForeColor
    End If
    
    If m_bIsNt Then
        SetTextColor lHdc, &HEFEFEF
        DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags
        OffsetRect tRect, 1, 1
        SetTextColor lHdc, &H808080
        DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags
        InflateRect tRect, 1, 1
        SetTextColor lHdc, lFrClr
        DrawTextW lHdc, StrPtr(sText), -1, tRect, lFlags
    Else
        SetTextColor lHdc, vbWhite
        DrawTextA lHdc, sText, -1, tRect, lFlags
        OffsetRect tRect, 1, 1
        SetTextColor lHdc, &H808080
        DrawTextA lHdc, sText, -1, tRect, lFlags
        InflateRect tRect, 1, 1
        SetTextColor lHdc, lFrClr
        DrawTextA lHdc, sText, -1, tRect, lFlags
    End If

End Sub

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
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Sub SkinAttach()

    If Not (m_lHdrHwnd = 0) Then
        If Not m_GXHeader Is Nothing Then
            If Not m_bSkinLoaded Then
                With m_GXHeader
                    .Subclass m_lHdrHwnd, Me
                    .AddMessage m_lHdrHwnd, WM_PAINT, MSG_BEFORE
                    .AddMessage m_lHdrHwnd, WM_MOUSEMOVE, MSG_BEFORE
                    .AddMessage m_lHdrHwnd, WM_MOUSELEAVE, MSG_BEFORE
                    .AddMessage m_lHdrHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                    .AddMessage m_lHdrHwnd, WM_LBUTTONUP, MSG_BEFORE
                    .AddMessage m_lHdrHwnd, WM_ERASEBKGND, MSG_BEFORE
                    .AddMessage m_lHdrHwnd, WM_SETCURSOR, MSG_BEFORE
                    .AddMessage m_lHdrHwnd, WM_SIZE, MSG_BEFORE
                    m_bSkinLoaded = True
                End With
            End If
        End If
    End If

End Sub

Private Sub SkinDetach()

    If Not (m_lHdrHwnd = 0) Then
        If Not m_GXHeader Is Nothing Then
            If m_bSkinLoaded Then
                With m_GXHeader
                    .DeleteMessage m_lHdrHwnd, WM_PAINT, MSG_BEFORE
                    .DeleteMessage m_lHdrHwnd, WM_MOUSEMOVE, MSG_BEFORE
                    .DeleteMessage m_lHdrHwnd, WM_MOUSELEAVE, MSG_BEFORE
                    .DeleteMessage m_lHdrHwnd, WM_LBUTTONDOWN, MSG_BEFORE
                    .DeleteMessage m_lHdrHwnd, WM_LBUTTONUP, MSG_BEFORE
                    .DeleteMessage m_lHdrHwnd, WM_ERASEBKGND, MSG_BEFORE
                    .DeleteMessage m_lHdrHwnd, WM_SETCURSOR, MSG_BEFORE
                    .DeleteMessage m_lHdrHwnd, WM_SIZE, MSG_BEFORE
                    .UnSubclass m_lHdrHwnd
                    m_bSkinLoaded = False
                End With
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

Dim tPaint As PAINTSTRUCT

    If Not m_bHeaderActive Then Exit Sub

    Select Case uMsg
    Case WM_PAINT
        If Not m_bPainting Then
            m_bPainting = True
            BeginPaint lHwnd, tPaint
            If m_bHeaderFlat Then
                m_lCurrState = 0
            End If
            Select Case m_lCurrState
            Case 1
                ColumnRender 1, m_lSelectedColumn
            Case 2
                ColumnRender 2, m_lSelectedColumn
            Case Else
                ColumnRender 0, 0
            End Select
            EndPaint lHwnd, tPaint
            m_bPainting = False
            bHandled = True
        Else
            lReturn = m_GXHeader.CallOldWndProc(lHwnd, uMsg, wParam, lParam)
        End If

    Case WM_LBUTTONDOWN
        m_eHeaderLastHitState = HeaderHitTest
        If m_bToolTips Then
            m_cColumnToolTip.StopTimer
        End If
        Select Case m_eHeaderLastHitState
        Case HdrHitHorzSizer
            If m_bColumnLocked(m_lSelectedColumn) Then
                bHandled = True
            End If
            m_lCurrState = 0
        Case HdrHitVertSizer
            m_bClickInColumnSizer = True
            m_lCurrState = 0
        Case HdrHitFilter
            m_lCurrState = 0
        Case HdrHitColumn
            If m_bDragState Then
                m_lCurrState = 0
            Else
                m_lCurrState = 1
            End If
        End Select
        SendMessageLongA lHwnd, WM_PAINT, 0&, 0&

    Case WM_LBUTTONUP
        m_lCurrState = 0
        SendMessageLongA lHwnd, WM_PAINT, 0&, 0&
        m_bClickInColumnSizer = False
        m_bClientClicked = False
        
    Case WM_MOUSEMOVE
        If LeftKeyState Then
            If m_bToolTips Then
                m_cColumnToolTip.StopTimer
            End If
            Select Case m_eHeaderLastHitState
            Case HdrHitColumn
                m_lCurrState = 1
            Case Else
                m_lCurrState = 0
            End Select
        Else
            Select Case HeaderHitTest
            Case HdrHitHorzSizer, HdrHitVertSizer, HdrHitFilter
                If m_bToolTips Then
                    m_cColumnToolTip.StopTimer
                End If
                m_lCurrState = 0
            Case HdrHitColumn
                If m_bFilterLoaded Then
                    If m_bToolTips Then
                        m_cColumnToolTip.StopTimer
                    End If
                    m_lCurrState = 0
                Else
                    If m_bDragState Then
                        m_lCurrState = 0
                        If m_bToolTips Then
                            m_cColumnToolTip.StopTimer
                        End If
                    Else
                        m_lCurrState = 2
                        If m_bToolTips Then
                            If Not (m_lTipOnColumn = m_lSelectedColumn) Then
                                ColumnToolTip m_lSelectedColumn
                                m_lTipOnColumn = m_lSelectedColumn
                            End If
                        End If
                    End If
                End If
            End Select
            If m_bDragState Then
                SetDragCursor
            Else
                SendMessageLongA lHwnd, WM_PAINT, 0&, 0&
            End If
        End If
        TrackMouse lHwnd

    Case WM_MOUSELEAVE
        If m_bToolTips Then
            m_cColumnToolTip.StopTimer
        End If
        m_lCurrState = 0
        SendMessageLongA lHwnd, WM_PAINT, 0&, 0&
        
    Case WM_SETCURSOR
        If LeftKeyState And (m_eHeaderLastHitState = HdrOffColumn) Then
            SetNormalCursor
            bHandled = True
        Else
            Select Case HeaderHitTest
            Case HdrHitVertSizer
                If m_bHeaderSizeable Then
                    SetNSSizingCursor
                Else
                    SetNormalCursor
                End If
            Case HdrHitHorzSizer
                If ColumnLocked(m_lSelectedColumn) Then
                    SetLockedCursor
                ElseIf m_bHeaderFixedWidth Then
                    SetNormalCursor
                Else
                    SetWESizingCursor
                End If
            Case Else
                SetNormalCursor
            End Select
            bHandled = True
        End If
        
    Case WM_SIZE
        SendMessageLongA lHwnd, WM_PAINT, 0&, 0&
    
    Case WM_ERASEBKGND
        lReturn = 1
        bHandled = True
    End Select

End Sub


'> Cleanup
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Sub DestroyHeaderImage()
    If Not m_pHeader Is Nothing Then Set m_pHeader = Nothing
End Sub

Private Sub CleanUp()

    If Not (m_lHdrHwnd = 0) Then
        DestroyCursors
        SkinDetach
        DestroyFont
        DestroyVericalFont
        If m_bHeaderActive Then
            BackbufferDestroy
        End If
        DestroyHeaderImage
        If Not m_oFont Is Nothing Then Set m_oFont = Nothing
        If Not m_oVtFont Is Nothing Then Set m_oVtFont = Nothing
        If Not m_oTipFont Is Nothing Then Set m_oTipFont = Nothing
        If Not m_cHeaderDc Is Nothing Then Set m_cHeaderDc = Nothing
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        If Not m_cColumnToolTip Is Nothing Then Set m_cColumnToolTip = Nothing
        If Not m_GXHeader Is Nothing Then Set m_GXHeader = Nothing
        Erase m_lHdc
        Erase m_lBmp
        Erase m_lBmpOld
        m_lHdrHwnd = 0
        m_lImlHdHndl = 0
        m_lHGHwnd = 0
        m_lhFnt = 0
    End If

End Sub

Private Sub Class_Terminate()
    CleanUp
End Sub


