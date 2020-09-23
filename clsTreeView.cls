VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTreeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'/* Rewrite of Carles P.V. api treeview
'/* http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57047&lngWId=1
'/* Complete rewrite: changed subclasser, removed IPAO, added skinned scrollbars, unicode support
'/* custom checkboxes, and removed ole (will be moved to grid in 1.5)

Implements GXISubclass


Private Const ALLOCATE_SIZE                     As Long = 100

Private Const CCM_FIRST                         As Long = &H2000
Private Const CCM_SETUNICODEFORMAT              As Long = (CCM_FIRST + 5)
Private Const CCM_GETUNICODEFORMAT              As Long = (CCM_FIRST + 6)

Private Const CDDS_PREPAINT                     As Long = &H1
Private Const CDDS_POSTPAINT                    As Long = &H2
Private Const CDDS_PREERASE                     As Long = &H3
Private Const CDDS_POSTERASE                    As Long = &H4
Private Const CDDS_ITEM                         As Long = &H10000
Private Const CDDS_ITEMPREPAINT                 As Long = CDDS_ITEM Or CDDS_PREPAINT
Private Const CDDS_ITEMPOSTPAINT                As Long = CDDS_ITEM Or CDDS_POSTPAINT
Private Const CDDS_SUBITEM                      As Long = &H20000

Private Const CDIS_SELECTED                     As Long = &H1
Private Const CDIS_GRAYED                       As Long = &H2
Private Const CDIS_DISABLED                     As Long = &H4
Private Const CDIS_CHECKED                      As Long = &H8
Private Const CDIS_FOCUS                        As Long = &H10
Private Const CDIS_DEFAULT                      As Long = &H20
Private Const CDIS_HOT                          As Long = &H40
Private Const CDIS_MARKED                       As Long = &H80
Private Const CDIS_INDETERMINATE                As Long = &H100

Private Const CDRF_DODEFAULT                    As Long = &H0
Private Const CDRF_NOTIFYITEMDRAW               As Long = &H20
Private Const CDRF_NOTIFYSUBITEMDRAW            As Long = &H20

Private Const CLR_NONE                          As Long = &HFFFFFFFF

Private Const COLOR_WINDOW                      As Long = 5
Private Const COLOR_WINDOWTEXT                  As Long = 8

Private Const FW_NORMAL                         As Long = 400
Private Const FW_BOLD                           As Long = 700

Private Const GW_HWNDNEXT                       As Long = &H2

Private Const GWL_STYLE                         As Long = (-16)
Private Const GWL_EXSTYLE                       As Long = (-20)

Private Const ILC_MASK                          As Long = &H1
Private Const ILC_COLORDDB                      As Long = &HFE
Private Const ILC_COLOR32                       As Long = &H20

Private Const LF_ANTIALIASED_QUALITY            As Long = 4
Private Const LF_CLEARTYPE_QUALITY              As Long = 5

Private Const LOGPIXELSY                        As Long = 90

Private Const MAX_PATH                          As Long = 260

Private Const MOUSEEVENTF_MOVE                  As Long = &H1
Private Const MOUSEEVENTF_LEFTDOWN              As Long = &H2
Private Const MOUSEEVENTF_LEFTUP                As Long = &H4

Private Const H_MAX                             As Long = &HFFFF + 1
Private Const NM_FIRST                          As Long = H_MAX
Private Const NM_CLICK                          As Long = (NM_FIRST - 2)
Private Const NM_DBLCLK                         As Long = (NM_FIRST - 3)
Private Const NM_RCLICK                         As Long = (NM_FIRST - 5)
Private Const NM_RDBLCLK                        As Long = (NM_FIRST - 6)
Private Const NM_SETFOCUS                       As Long = (NM_FIRST - 7)
Private Const NM_CUSTOMDRAW                     As Long = (NM_FIRST - 12)

Private Const OLE_FORMAT_ID                     As Long = &HFFFFB044

Private Const PATH_SEPARATOR                    As String = "\"

Private Const SII_SWAPMASK                      As Long = 3
Private Const SII_CHECKED                       As Long = 2
Private Const SII_UNCHECKED                     As Long = 1

Private Const SW_SHOW                           As Long = 5

Private Const SWP_NOSIZE                        As Long = &H1
Private Const SWP_NOMOVE                        As Long = &H2
Private Const SWP_NOZORDER                      As Long = &H4
Private Const SWP_NOACTIVATE                    As Long = &H10
Private Const SWP_FRAMECHANGED                  As Long = &H20
Private Const SWP_NOOWNERZORDER                 As Long = &H200

Private Const T_EXPAND_DELAY                    As Long = 1000

Private Const TVE_COLLAPSE                      As Long = &H1
Private Const TVE_EXPAND                        As Long = &H2
Private Const TVE_TOGGLE                        As Long = &H3

Private Const TVGN_ROOT                         As Long = &H0
Private Const TVGN_NEXT                         As Long = &H1
Private Const TVGN_PREVIOUS                     As Long = &H2
Private Const TVGN_PARENT                       As Long = &H3
Private Const TVGN_CHILD                        As Long = &H4
Private Const TVGN_DROPHILITE                   As Long = &H8
Private Const TVGN_CARET                        As Long = &H9
Private Const TVGN_FIRSTVISIBLE                 As Long = &H5
Private Const TVGN_NEXTVISIBLE                  As Long = &H6
Private Const TVGN_PREVIOUSVISIBLE              As Long = &H7
Private Const TVGN_LASTVISIBLE                  As Long = &HA

Private Const TVI_ROOT                          As Long = &HFFFF0000
Private Const TVI_FIRST                         As Long = &HFFFF0001
Private Const TVI_LAST                          As Long = &HFFFF0002
Private Const TVI_SORT                          As Long = &HFFFF0003

Private Const TVIF_TEXT                         As Long = &H1
Private Const TVIF_IMAGE                        As Long = &H2
Private Const TVIF_PARAM                        As Long = &H4
Private Const TVIF_SELECTEDIMAGE                As Long = &H20
Private Const TVIF_all                          As Long = TVIF_TEXT Or TVIF_IMAGE Or TVIF_PARAM Or TVIF_SELECTEDIMAGE
Private Const TVIF_CHILDREN                     As Long = &H40
Private Const TVIF_STATE                        As Long = &H8
Private Const TVIF_HANDLE                       As Long = &H10

Private Const TVIS_SELECTED                     As Long = &H2
Private Const TVIS_CUT                          As Long = &H4
Private Const TVIS_DROPHILITED                  As Long = &H8
Private Const TVIS_BOLD                         As Long = &H10
Private Const TVIS_EXPANDED                     As Long = &H20
Private Const TVIS_EXPANDEDONCE                 As Long = &H40
Private Const TVIS_STATEIMAGEMASK               As Long = &HF000
Private Const TVHT_NOWHERE                      As Long = &H1
Private Const TVHT_ONITEMICON                   As Long = &H2
Private Const TVHT_ONITEMLABEL                  As Long = &H4
Private Const TVHT_ONITEMINDENT                 As Long = &H8
Private Const TVHT_ONITEMBUTTON                 As Long = &H10
Private Const TVHT_ONITEMRIGHT                  As Long = &H20
Private Const TVHT_ONITEMSTATEICON              As Long = &H40
Private Const TVHT_ABOVE                        As Long = &H100
Private Const TVHT_BELOW                        As Long = &H200
Private Const TVHT_TORIGHT                      As Long = &H400
Private Const TVHT_TOLEFT                       As Long = &H800
Private Const TVHT_ONITEM                       As Long = TVHT_ONITEMICON Or TVHT_ONITEMLABEL Or TVHT_ONITEMSTATEICON

Private Const TV_FIRST                          As Long = &H1100
Private Const TVM_INSERTITEMA                   As Long = (TV_FIRST + 0)
Private Const TVM_DELETEITEM                    As Long = (TV_FIRST + 1)
Private Const TVM_EXPAND                        As Long = (TV_FIRST + 2)
Private Const TVM_GETITEMRECT                   As Long = (TV_FIRST + 4)
Private Const TVM_GETCOUNT                      As Long = (TV_FIRST + 5)
Private Const TVM_GETINDENT                     As Long = (TV_FIRST + 6)
Private Const TVM_SETINDENT                     As Long = (TV_FIRST + 7)
Private Const TVM_GETIMAGELIST                  As Long = (TV_FIRST + 8)
Private Const TVM_SETIMAGELIST                  As Long = (TV_FIRST + 9)
Private Const TVM_GETNEXTITEM                   As Long = (TV_FIRST + 10)
Private Const TVM_SELECTITEM                    As Long = (TV_FIRST + 11)
Private Const TVM_GETITEMA                      As Long = (TV_FIRST + 12)
Private Const TVM_SETITEMA                      As Long = (TV_FIRST + 13)
Private Const TVM_EDITLABELA                    As Long = (TV_FIRST + 14)
Private Const TVM_GETEDITCONTROL                As Long = (TV_FIRST + 15)
Private Const TVM_HITTEST                       As Long = (TV_FIRST + 17)
Private Const TVM_CREATEDRAGIMAGE               As Long = (TV_FIRST + 18)
Private Const TVM_SORTCHILDREN                  As Long = (TV_FIRST + 19)
Private Const TVM_ENSUREVISIBLE                 As Long = (TV_FIRST + 20)
Private Const TVM_ENDEDITLABELNOW               As Long = (TV_FIRST + 22)
Private Const TVM_SETINSERTMARK                 As Long = (TV_FIRST + 26)
Private Const TVM_SETITEMHEIGHT                 As Long = (TV_FIRST + 27)
Private Const TVM_GETITEMHEIGHT                 As Long = (TV_FIRST + 28)
Private Const TVM_SETBKCOLOR                    As Long = (TV_FIRST + 29)
Private Const TVM_SETTEXTCOLOR                  As Long = (TV_FIRST + 30)
Private Const TVM_GETBKCOLOR                    As Long = (TV_FIRST + 31)
Private Const TVM_GETTEXTCOLOR                  As Long = (TV_FIRST + 32)
Private Const TVM_SETINSERTMARKCOLOR            As Long = (TV_FIRST + 37)
Private Const TVM_GETINSERTMARKCOLOR            As Long = (TV_FIRST + 38)
Private Const TVM_SETLINECOLOR                  As Long = (TV_FIRST + 40)
Private Const TVM_GETLINECOLOR                  As Long = (TV_FIRST + 41)

Private Const TVM_INSERTITEMW                   As Long = (TV_FIRST + 50)
Private Const TVM_GETITEMW                      As Long = (TV_FIRST + 62)
Private Const TVM_SETITEMW                      As Long = (TV_FIRST + 63)
Private Const TVM_EDITLABELW                    As Long = (TV_FIRST + 65)


Private Const TVN_FIRST                         As Long = -400
Private Const TVN_SELCHANGINGA                  As Long = (TVN_FIRST - 1)
Private Const TVN_SELCHANGEDA                   As Long = (TVN_FIRST - 2)
Private Const TVN_ITEMEXPANDINGA                As Long = (TVN_FIRST - 5)
Private Const TVN_ITEMEXPANDEDA                 As Long = (TVN_FIRST - 6)
Private Const TVN_BEGINDRAGA                    As Long = (TVN_FIRST - 7)
Private Const TVN_DELETEITEMA                   As Long = (TVN_FIRST - 9)
Private Const TVN_BEGINLABELEDITA               As Long = (TVN_FIRST - 10)
Private Const TVN_ENDLABELEDITA                 As Long = (TVN_FIRST - 11)

Private Const TVN_SELCHANGINGW                  As Long = (TVN_FIRST - 50)
Private Const TVN_SELCHANGEDW                   As Long = (TVN_FIRST - 51)
Private Const TVN_ITEMEXPANDINGW                As Long = (TVN_FIRST - 54)
Private Const TVN_ITEMEXPANDEDW                 As Long = (TVN_FIRST - 55)
Private Const TVN_BEGINDRAGW                    As Long = (TVN_FIRST - 56)
Private Const TVN_DELETEITEMW                   As Long = (TVN_FIRST - 58)
Private Const TVN_BEGINLABELEDITW               As Long = (TVN_FIRST - 59)
Private Const TVN_ENDLABELEDITW                 As Long = (TVN_FIRST - 60)


Private Const TVS_HASBUTTONS                    As Long = &H1
Private Const TVS_HASLINES                      As Long = &H2
Private Const TVS_LINESATROOT                   As Long = &H4
Private Const TVS_EDITLABELS                    As Long = &H8
Private Const TVS_DISABLEDRAGDROP               As Long = &H10
Private Const TVS_SHOWSELALWAYS                 As Long = &H20
Private Const TVS_CHECKBOXES                    As Long = &H100
Private Const TVS_TRACKSELECT                   As Long = &H200
Private Const TVS_SINGLEEXPAND                  As Long = &H400
Private Const TVS_FULLROWSELECT                 As Long = &H1000

Private Const TVSIL_NORMAL                      As Long = &H0
Private Const TVSIL_STATE                       As Long = &H2

Private Const VER_PLATFORM_WIN32_NT             As Long = 2

Private Const VK_LBUTTON                        As Long = &H1
Private Const VK_TAB                            As Long = &H9

Private Const WC_TREEVIEW                       As String = "SysTreeView32"

Private Const WM_SIZE                           As Long = &H5
Private Const WM_SETFOCUS                       As Long = &H7
Private Const WM_KILLFOCUS                      As Long = &H8
Private Const WM_SETREDRAW                      As Long = &HB
Private Const WM_SETFONT                        As Long = &H30
Private Const WM_MOUSEACTIVATE                  As Long = &H21
Private Const WM_NOTIFY                         As Long = &H4E
Private Const WM_KEYDOWN                        As Long = &H100
Private Const WM_KEYUP                          As Long = &H101
Private Const WM_CHAR                           As Long = &H102
Private Const WM_TIMER                          As Long = &H113
Private Const WM_HSCROLL                        As Long = &H114
Private Const WM_VSCROLL                        As Long = &H115
Private Const WM_MOUSEMOVE                      As Long = &H200
Private Const WM_LBUTTONUP                      As Long = &H202
Private Const WM_LBUTTONDOWN                    As Long = &H201
Private Const WM_RBUTTONDOWN                    As Long = &H204
Private Const WM_RBUTTONUP                      As Long = &H205
Private Const WM_MBUTTONDOWN                    As Long = &H207
Private Const WM_MBUTTONUP                      As Long = &H208
Private Const WM_MOUSELEAVE                     As Long = &H2A3

Private Const WS_TABSTOP                        As Long = &H10000
Private Const WS_BORDER                         As Long = &H800000
Private Const WS_CHILD                          As Long = &H40000000

Private Const WS_EX_CLIENTEDGE                  As Long = &H200
Private Const WS_EX_RTLREADING                  As Long = &H2000
Private Const WS_EX_LEFTSCROLLBAR               As Long = &H4000

Private Const SB_LINEUP                         As Long = 0
Private Const SB_LINELEFT                       As Long = 0
Private Const SB_LINEDOWN                       As Long = 1
Private Const SB_LINERIGHT                      As Long = 1
Private Const SB_PAGEUP                         As Long = 2
Private Const SB_PAGELEFT                       As Long = 2
Private Const SB_PAGEDOWN                       As Long = 3
Private Const SB_PAGERIGHT                      As Long = 3
Private Const SB_TOP                            As Long = 6
Private Const SB_LEFT                           As Long = 6
Private Const SB_BOTTOM                         As Long = 7
Private Const SB_RIGHT                          As Long = 7
Private Const SB_ENDSCROLL                      As Long = 8


Private Enum TRACKMOUSEEVENT_FLAGS
    TME_HOVER = &H1
    TME_LEAVE = &H2
    TME_QUERY = &H40000000
    TME_CANCEL = &H80000000
End Enum

Public Enum ETVThemeLuminence
    estThemeSoft = 0&
    estThemePastel = 1&
    estThemeHard = 2&
End Enum

Public Enum ETVThemeStyle
    etvAzure = 0&
    etvClassic = 1&
    etvGloss = 2&
    etvMetallic = 3&
    etvXPSilver = 4&
    etvXPBlue = 5&
    etvXPGreen = 6&
    etvVista = 7&
End Enum

Public Enum ETVBorderStyle
    tbsNone = 0&
    tbsFixedSingle = 1&
End Enum

Public Enum ETVScrollConstants
    tvsHome = 0&
    tvsPageUp = 1&
    tvsUp = 2&
    tvsDown = 3&
    tvsPageDown = 4&
    tvsEnd = 5&
    tvsLeft = 6&
    tvsPageLeft = 7&
    tvsLineLeft = 8&
    tvsLineRight = 9&
    tvsPageRight = 10&
    tvsRight = 11&
End Enum

Private Enum ETVStateConstants
    tvcStateCut = TVIS_CUT
    tvcStateDropHilited = TVIS_DROPHILITED
    tvcStateBold = TVIS_BOLD
    tvcStateExpanded = TVIS_EXPANDED
    tvcStateExpandedOnce = TVIS_EXPANDEDONCE
End Enum

Private Enum SYSTEM_METRICS
    SM_CXSCREEN = 0&
    SM_CYSCREEN = 1&
    SM_CXVSCROLL = 2&
    SM_CYHSCROLL = 3&
    SM_CYCAPTION = 4&
    SM_CXBORDER = 5&
    SM_CYBORDER = 6&
    SM_CYVTHUMB = 9&
    SM_CXHTHUMB = 10&
    SM_CXICON = 11&
    SM_CYICON = 12&
    SM_CXCURSOR = 13&
    SM_CYCURSOR = 14&
    SM_CYMENU = 15&
    SM_CXFULLSCREEN = 16&
    SM_CYFULLSCREEN = 17&
    SM_CYKANJIWINDOW = 18&
    SM_MOUSEPRESENT = 19&
    SM_CYVSCROLL = 20&
    SM_CXHSCROLL = 21&
    SM_CXMIN = 28&
    SM_CYMIN = 29&
    SM_CXSIZE = 30&
    SM_CYSIZE = 31&
    SM_CXFRAME = 32&
    SM_CYFRAME = 33&
    SM_CXMINTRACK = 34&
    SM_CYMINTRACK = 35&
    SM_CXSMICON = 49&
    SM_CYSMICON = 50&
    SM_CYSMCAPTION = 51&
    SM_CXMINIMIZED = 57&
    SM_CYMINIMIZED = 58&
    SM_CXMAXTRACK = 59&
    SM_CYMAXTRACK = 60&
    SM_CXMAXIMIZED = 61&
    SM_CYMAXIMIZED = 62&
End Enum


Private Type TVITEMA
    Mask                                        As Long
    hItem                                       As Long
    State                                       As Long
    stateMask                                   As Long
    pszText                                     As String
    cchTextMax                                  As Long
    iImage                                      As Long
    iSelectedImage                              As Long
    cChildren                                   As Long
    lParam                                      As Long
End Type

Private Type TVITEMW
    Mask                                        As Long
    hItem                                       As Long
    State                                       As Long
    stateMask                                   As Long
    pszText                                     As Long
    cchTextMax                                  As Long
    iImage                                      As Long
    iSelectedImage                              As Long
    cChildren                                   As Long
    lParam                                      As Long
End Type

Private Type TVITEMEXA
    Mask                                        As Long
    hItem                                       As Long
    State                                       As Long
    stateMask                                   As Long
    pszText                                     As String
    cchTextMax                                  As Long
    iImage                                      As Long
    iSelectedImage                              As Long
    cChildren                                   As Long
    lParam                                      As Long
    iIntegral                                   As Long
End Type

Private Type TVINSERTSTRUCTA
    hParent                                     As Long
    hInsertAfter                                As Long
    Item                                        As TVITEMEXA
End Type

Private Type TVITEMEXW
    Mask                                        As Long
    hItem                                       As Long
    State                                       As Long
    stateMask                                   As Long
    pszText                                     As Long
    cchTextMax                                  As Long
    iImage                                      As Long
    iSelectedImage                              As Long
    cChildren                                   As Long
    lParam                                      As Long
    iIntegral                                   As Long
End Type

Private Type TVINSERTSTRUCTW
    hParent                                     As Long
    hInsertAfter                                As Long
    Item                                        As TVITEMEXW
End Type

Private Type NMHDR
    hwndFrom                                    As Long
    idfrom                                      As Long
    code                                        As Long
End Type

Private Type POINTAPI
    x                                           As Long
    y                                           As Long
End Type

Private Type NMTREEVIEW
    hdr                                         As NMHDR
    action                                      As Long
    itemOld                                     As TVITEMW
    itemNew                                     As TVITEMW
    ptDrag                                      As POINTAPI
End Type

Private Type NMTVDISPINFO
    hdr                                         As NMHDR
    Item                                        As TVITEMW
End Type

Private Type TVHITTESTINFO
    pt                                          As POINTAPI
    flags                                       As Long
    hItem                                       As Long
End Type

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                                      As Long
    dwFlags                                     As TRACKMOUSEEVENT_FLAGS
    hwndTrack                                   As Long
    dwHoverTime                                 As Long
End Type

Private Type RECT
    left                                        As Long
    top                                         As Long
    Right                                       As Long
    Bottom                                      As Long
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

Private Type tSubData
    hwnd                                        As Long
    nAddrSub                                    As Long
    nAddrOrig                                   As Long
    nMsgCntA                                    As Long
    nMsgCntB                                    As Long
    aMsgTblA()                                  As Long
    aMsgTblB()                                  As Long
    sCode                                       As String
End Type

Private Type NMCUSTOMDRAW
    hdr                                         As NMHDR
    dwDrawStage                                 As Long
    hdc                                         As Long
    rc                                          As RECT
    dwItemSpec                                  As Long
    uItemState                                  As Long
    lItemlParam                                 As Long
End Type

Private Type NMTVCUSTOMDRAW
    nmcmd                                       As NMCUSTOMDRAW
    clrText                                     As Long
    clrTextBk                                   As Long
    iLevel                                      As Long
End Type

Private Type NODE_DATA
    hNode                                       As Long
    sKey                                        As String
    sTag                                        As String
End Type

Private Type OSVERSIONINFO
    dwVersionInfoSize                           As Long
    dwMajorVersion                              As Long
    dwMinorVersion                              As Long
    dwBuildNumber                               As Long
    dwPlatformId                                As Long
    szCSDVersion(0 To 127)                      As Byte
End Type


Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long

Private Declare Function ImageList_Create Lib "comctl32" (ByVal MinCx As Long, _
                                                          ByVal MinCy As Long, _
                                                          ByVal flags As Long, _
                                                          ByVal cInitial As Long, _
                                                          ByVal cGrow As Long) As Long

Private Declare Function ImageList_Add Lib "comctl32" (ByVal hImageList As Long, _
                                                       ByVal hBitmap As Long, _
                                                       ByVal hBitmapMask As Long) As Long

Private Declare Function ImageList_AddMasked Lib "comctl32" (ByVal hImageList As Long, _
                                                             ByVal hbmImage As Long, _
                                                             ByVal crMask As Long) As Long

Private Declare Function ImageList_AddIcon Lib "comctl32" (ByVal hImageList As Long, _
                                                           ByVal hicon As Long) As Long

Private Declare Function ImageList_GetIconSize Lib "comctl32" (ByVal hIml As Long, _
                                                               cx As Long, _
                                                               cy As Long) As Long

Private Declare Function ImageList_Destroy Lib "comctl32" (ByVal hImageList As Long) As Long

Private Declare Function ImageList_GetImageCount Lib "comctl32" (ByVal hImageList As Long) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function TrackMouseEventComCtl Lib "comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, _
                                                      lpPoint As POINTAPI) As Long

Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, _
                                                      lpPoint As POINTAPI) As Long

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, _
                                                     lpSrc1Rect As RECT, _
                                                     lpSrc2Rect As RECT) As Long

Private Declare Function EqualRect Lib "user32" (lpRect1 As RECT, _
                                                 lpRect2 As RECT) As Long

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
                                                ByVal x As Long, _
                                                ByVal y As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, _
                                                    ByVal nIndex As Long) As Long

Private Declare Function CreateFontIndirectA Lib "gdi32" (lpLogFont As LOGFONT) As Long

Private Declare Function CreateFontIndirectW Lib "gdi32" (lpLogFont As LOGFONT) As Long


Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, _
                                                ByVal nNumerator As Long, _
                                                ByVal nDenominator As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal nCmdShow As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, _
                                                           ByVal HPALETTE As Long, _
                                                           pccolorref As Long) As Long

Private Declare Sub InitCommonControls Lib "comctl32" ()

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

Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long, _
                                                  ByVal nWidth As Long, _
                                                  ByVal nHeight As Long, _
                                                  ByVal bRepaint As Long) As Long

Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

Private Declare Function SetWindowLongW Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long, _
                                                      ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLongA Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long) As Long

Private Declare Function GetWindowLongW Lib "user32" (ByVal hwnd As Long, _
                                                      ByVal nIndex As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, _
                                                                     lpvSource As Any, _
                                                                     ByVal cbCopy As Long)

Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    lParam As Any) As Long

Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal wMsg As Long, _
                                                    ByVal wParam As Long, _
                                                    lParam As Any) As Long

Private Declare Function SendMessageLongA Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long

Private Declare Function SendMessageLongW Lib "user32" Alias "SendMessageW" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long

Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function timeGetTime Lib "winmm" () As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
                                                        ByVal lpProcName As String) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As SYSTEM_METRICS) As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal y As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Long, _
                                                  ByVal lpString2 As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal wCmd As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal fEnable As Long) As Long

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, _
                                              ByVal dx As Long, _
                                              ByVal dy As Long, _
                                              ByVal cButtons As Long, _
                                              ByVal dwExtraInfo As Long)

Private Declare Function GetMessageExtraInfo Lib "user32" () As Long

Private Declare Function ImageList_Remove Lib "comctl32.dll" (ByVal hIml As Long, _
                                                              ByVal I As Long) As Long

Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, _
                                                 ByVal crColor As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal x As Long, _
                                                  ByVal y As Long) As Long

Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hwnd As Long, _
                                                           ByVal pszSubAppName As Long, _
                                                           ByVal pszSubIdList As Long) As Long

Private Declare Function EraseRect Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, _
                                                                        lpRect As RECT, _
                                                                        ByVal bErase As Long) As Long


Public Event Click()
Public Event NodeClick(ByVal hNode As Long)
Public Event NodeCheck(ByVal hNode As Long)
Public Event NodeDblClick(ByVal hNode As Long)
Public Event SelectionChanged()
Public Event BeforeExpand(ByVal hNode As Long, ByVal ExpandedOnce As Boolean)
Public Event AfterExpand(ByVal hNode As Long, ByVal ExpandedOnce As Boolean)
Public Event Collapse(ByVal hNode As Long)
Public Event BeforeLabelEdit(ByVal hNode As Long, Cancel As Long)
Public Event AfterLabelEdit(ByVal hNode As Long, Cancel As Long, NewString As String)
Public Event KeyDown(KeyCode As Long, Shift As Long)
Public Event KeyPress(KeyAscii As Long)
Public Event KeyUp(KeyCode As Long, Shift As Long)
Public Event MouseDown(Button As Long, Shift As Long, x As Long, y As Long)
Public Event MouseMove(Button As Long, Shift As Long, x As Long, y As Long)
Public Event MouseUp(Button As Long, Shift As Long, x As Long, y As Long)
Public Event MouseEnter()
Public Event MouseLeave()
'Public Event Resize()
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Long, Shift As Long, x As Single, y As Single, State As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Long, Shift As Long, x As Single, y As Single)
Public Event OLECompleteDrag(Effect As Long)


Private m_bNodeClick                            As Boolean
Private m_bCheckBoxes                           As Boolean
Private m_bFullRowSelect                        As Boolean
Private m_bHasButtons                           As Boolean
Private m_bHasLines                             As Boolean
Private m_bHasRootLines                         As Boolean
Private m_bHideSelection                        As Boolean
Private m_bLabelEdit                            As Boolean
Private m_bSingleExpand                         As Boolean
Private m_bTrackSelect                          As Boolean
Private m_bOLEDragAutoExpand                    As Boolean
Private m_bInitialized                          As Boolean
Private m_bHoldDeletePostProcess                As Boolean
Private m_bTrack                                As Boolean
Private m_bTrackUser32                          As Boolean
Private m_bInControl                            As Boolean
Private m_bNodeDropInsertAfter                  As Boolean
Private m_bIsNt                                 As Boolean
Private m_bIsXp                                 As Boolean
Private m_bUseUnicode                           As Boolean
Private m_bScrollBarLeftAlign                   As Boolean
Private m_bVisible                              As Boolean
Private m_bEnabled                              As Boolean
Private m_bFontRightLeading                     As Boolean
Private m_lGridHwnd                             As Long
Private m_lSelectedBackColor                    As Long
Private m_lDisabledBackColor                    As Long
Private m_lDisabledForeColor                    As Long
Private m_lFocusBackColor                       As Long
Private m_lFocusForeColor                       As Long
Private m_lBackColor                            As Long
Private m_lForeColor                            As Long
Private m_lHeight                               As Long
Private m_lWidth                                As Long
Private m_hModShell32                           As Long
Private m_lTVHwnd                               As Long
Private m_lImlNodeHndl                          As Long
Private m_lImageListCount                       As Long
Private m_lHFont                                As Long
Private m_lXCoord                               As Long
Private m_lYCoord                               As Long
Private m_lNodeDrag                             As Long
Private m_lNodeDrop                             As Long
Private m_lNodeCount                            As Long
Private m_lExpandCounter                        As Long
Private m_lParentHwnd                           As Long
Private m_lImlStateHndl                         As Long
Private m_lOLEDragMode                          As Long
Private m_eCheckBoxSkinStyle                    As ETVThemeStyle
Private m_eThemeLuminence                       As ETVThemeLuminence
Private m_eScrollBarSkinStyle                   As ETVThemeStyle
Private m_IChecked                              As StdPicture
Private m_IChkDisabled                          As StdPicture
Private m_IUnChecked                            As StdPicture
Private m_cKey                                  As Collection
Private m_uNodeData()                           As NODE_DATA
Private m_cChkCheckDc                           As clsStoreDc
Private m_cRender                               As clsRender
Private m_cSkinScrollBars                       As clsSkinScrollbars
Private WithEvents m_oFont                      As StdFont
Attribute m_oFont.VB_VarHelpID = -1
Private m_cTreeSubclass                         As GXMSubclass



Private Sub m_oFont_FontChanged(ByVal PropertyName As String)
    Set Font = m_oFont
End Sub


'**********************************************************************
'*                              CONSTRUCTORS
'**********************************************************************

Private Sub Class_Initialize()

    m_hModShell32 = LoadLibraryA("shell32.dll")
    InitCommonControls
    m_bEnabled = True
    m_lSelectedBackColor = GetSysColor(vbButtonFace And &H1F&)
    m_lFocusBackColor = GetSysColor(vbHighlight And &H1F&)
    m_lFocusForeColor = GetSysColor(vbHighlightText And &H1F&)
    m_lDisabledForeColor = &H999999
    m_lDisabledBackColor = &HCCCCCC
    
    Set m_oFont = New StdFont
    Set m_cRender = New clsRender
    Set m_cTreeSubclass = New GXMSubclass
    ReDim m_uNodeData(0 To 1)
    Set m_cKey = New Collection

End Sub

Private Function CreateTreeView() As Boolean

Dim lExStyle As Long
Dim lTVStyle As Long

    '/* base styles
    lTVStyle = WS_CHILD Or WS_TABSTOP Or TVS_SHOWSELALWAYS Or TVS_DISABLEDRAGDROP
    '/* extended style
    lExStyle = WS_EX_CLIENTEDGE
    '/* left align scrollbar
    If m_bScrollBarLeftAlign Then
        lExStyle = lExStyle Or WS_EX_LEFTSCROLLBAR
    End If
    If m_bFontRightLeading Then
        lExStyle = lExStyle Or WS_EX_RTLREADING
    End If
    '/* create the treeview
    If m_bIsNt Then
        m_lTVHwnd = CreateWindowExW(lExStyle, StrPtr(WC_TREEVIEW), StrPtr(""), lTVStyle, 0&, 0&, 0&, 0&, m_lParentHwnd, 0&, App.hInstance, ByVal 0&)
    Else
        m_lTVHwnd = CreateWindowExA(lExStyle, WC_TREEVIEW, "", lTVStyle, 0&, 0&, 0&, 0&, m_lParentHwnd, 0&, App.hInstance, ByVal 0&)
    End If
    If m_bIsXp Then
        SetWindowTheme m_lTVHwnd, StrPtr(" "), StrPtr(" ")
    End If

    '/* set defaults
    If Not (m_lTVHwnd = 0) Then
        SendMessageLongA m_lTVHwnd, TVM_SETBKCOLOR, 0&, GetSysColor(COLOR_WINDOW)
        SendMessageLongA m_lTVHwnd, TVM_SETTEXTCOLOR, 0&, GetSysColor(COLOR_WINDOWTEXT)
        'Set m_oFont = Ambient.Font
        m_oFont_FontChanged vbNullString
        '/* show the control
        ShowWindow m_lTVHwnd, SW_SHOW
        CreateTreeView = True
    End If

End Function

Public Function Initialize(ByVal lParentHwnd As Long, _
                           ByVal lHeight As Long, _
                           ByVal lWidth As Long) As Boolean
'/* initialize treeview

    If Not m_bInitialized Then
        '/* owner handle
        m_lHeight = lHeight
        m_lWidth = lWidth
        m_lParentHwnd = lParentHwnd
        VersionCheck
        Initialize = CreateTreeView()
        UseUnicode = True
        If Not (m_lTVHwnd = 0) Then
            TreeAttach
            m_lXCoord = -1
            m_lYCoord = -1
            m_bInitialized = True
        End If
    End If

End Function

Public Function IsUnicode(ByVal sText As String) As Boolean
'/* test for unicode
'/* good link: http://www.unicodeactivex.com/UnicodeTutorialVb.htm

Dim lLen   As Long
Dim bLen   As Long
Dim bmap() As Byte

    If LenB(sText) Then
        bmap = sText
        bLen = UBound(bmap)
        For lLen = 1 To bLen Step 2
            If (bmap(lLen) > 0) Then
                IsUnicode = True
                Exit For
            End If
        Next lLen
    End If

End Function

Private Function PointerToString(ByVal lpString As Long) As String
'/* get string from pointer

Dim lLen As Long

    If Not (lpString = 0) Then
        lLen = lstrlenW(ByVal lpString)
        If Not (lLen = 0) Then
            '/* allocate string with nLen chars
            PointerToString = String$(lLen, Chr$(0))
            lstrcpyW StrPtr(PointerToString), lpString
        End If
    End If

End Function

Private Function SetUnicode(ByVal bEnable As Boolean) As Boolean
'/* enable/disable unicode processing

Dim lRet As Long

    If Not (m_lTVHwnd = 0) Then
        If m_bIsNt Then
            If bEnable Then
                If Not UnicodeState Then
                    lRet = SendMessageLongW(m_lTVHwnd, CCM_SETUNICODEFORMAT, 1&, 0&)
                End If
            Else
                If UnicodeState Then
                    lRet = SendMessageLongW(m_lTVHwnd, CCM_SETUNICODEFORMAT, 0&, 0&)
                End If
            End If
        End If
    Else
        lRet = -1
    End If
    SetUnicode = (lRet = 0)

End Function

Private Function UnicodeState() As Boolean
'/* get control unicode readiness

    If Not (m_lTVHwnd = 0) Then
        UnicodeState = SendMessageLongW(m_lTVHwnd, CCM_GETUNICODEFORMAT, 0&, 0&) <> 0
    End If

End Function

Private Function VersionCheck() As Boolean
'/* operating system check

Dim tVer As OSVERSIONINFO

    With tVer
        .dwVersionInfoSize = LenB(tVer)
        GetVersionEx tVer
        m_bIsNt = ((.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
        If (.dwMajorVersion >= 5) Then
            m_bIsXp = True
        End If
    End With
    If Not m_bIsNt Then
        m_bUseUnicode = False
    End If
    VersionCheck = m_bIsNt

End Function


'**********************************************************************
'*                              PROPERTIES
'**********************************************************************

Public Property Get BackColor() As Long
'/* [get] backcolor
    BackColor = m_lBackColor
End Property

Public Property Let BackColor(ByVal PropVal As Long)
'/* [let] backcolor
    m_lBackColor = PropVal
    If Not (m_lTVHwnd = 0) Then
        OleTranslateColor PropVal, 0&, m_lBackColor
        SendMessageLongA m_lTVHwnd, TVM_SETBKCOLOR, 0&, m_lBackColor
    End If
End Property

Public Property Get BorderStyle() As ETVBorderStyle
'/* [get] border style
    If Not (m_lTVHwnd = 0) Then
        BorderStyle = -((GetWindowLongA(m_lTVHwnd, GWL_EXSTYLE) And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE)
    End If
End Property

Public Property Let BorderStyle(ByVal PropVal As ETVBorderStyle)
'/* [let] border style
    If Not (m_lTVHwnd = 0) Then
        Select Case PropVal
        Case tbsNone
            SetWindowLongA m_lTVHwnd, GWL_EXSTYLE, 0&
        Case tbsFixedSingle
            SetWindowLongA m_lTVHwnd, GWL_EXSTYLE, WS_EX_CLIENTEDGE
        End Select
    End If
End Property

Public Property Get CheckBoxes() As Boolean
'/* [get] checkbox state
    CheckBoxes = m_bCheckBoxes
End Property

Public Property Let CheckBoxes(ByVal PropVal As Boolean)
'/* [let] toggle checkboxes
Dim lNode As Long

    If Not (m_lTVHwnd = 0) Then
        m_bCheckBoxes = PropVal
        If PropVal Then
            WindowStyle m_lTVHwnd, GWL_STYLE, TVS_CHECKBOXES, 0
        Else
            WindowStyle m_lTVHwnd, GWL_STYLE, 0, TVS_CHECKBOXES
            For lNode = 1 To m_lNodeCount
                SetTvStateImage m_uNodeData(lNode).hNode, 0
            Next lNode
            ImageList_Destroy SendMessageLongA(m_lTVHwnd, TVM_GETIMAGELIST, TVSIL_STATE, 0&)
        End If
    End If

End Property

Private Sub DefaultFont(ByRef oFont As StdFont)

    Set oFont = New StdFont
    With oFont
        .Charset = 3
        .Name = "MS Sans Serif"
        .Weight = 400
        .Size = 8
    End With

End Sub

Public Property Let DisabledBackColor(ByVal PropVal As Long)
'/* [let] Disabled BackColor
    m_lDisabledBackColor = PropVal
End Property

Public Property Get DisabledBackColor() As Long
'/* [get] Disabled BackColor
    DisabledBackColor = m_lDisabledBackColor
End Property

Public Property Get DisabledForeColor() As Long
'/* [get] Disabled ForeColor
    DisabledForeColor = m_lDisabledForeColor
End Property

Public Property Let DisabledForeColor(ByVal PropVal As Long)
'/* [let] Disabled ForeColor
    m_lDisabledForeColor = PropVal
End Property

Public Property Get OLEDragMode() As Long
'/* [get] hide selected
    OLEDragMode = m_lOLEDragMode
End Property

Public Property Let OLEDragMode(ByVal PropVal As Long)
'/* [let] hide selected
    If Not (m_lTVHwnd = 0) Then
        If (PropVal = 1) Then
            WindowStyle m_lTVHwnd, GWL_STYLE, 0, TVS_DISABLEDRAGDROP
        Else
            WindowStyle m_lTVHwnd, GWL_STYLE, TVS_DISABLEDRAGDROP, 0
        End If
    End If
    m_lOLEDragMode = PropVal
End Property

Public Property Get FocusBackColor() As Long
'/* [get] Selected ForeColor
    FocusBackColor = m_lFocusBackColor
End Property

Public Property Let FocusBackColor(ByVal PropVal As Long)
'/* [let] Selected ForeColor
    m_lFocusBackColor = PropVal
End Property

Public Property Get FocusForeColor() As Long
'/* [get] Selected ForeColor
    FocusForeColor = m_lFocusForeColor
End Property

Public Property Let FocusForeColor(ByVal PropVal As Long)
'/* [let] Selected ForeColor
    m_lFocusForeColor = PropVal
End Property

Public Property Get FontRightLeading() As Boolean
'/* [get] right align fonts
    FontRightLeading = m_bFontRightLeading
End Property

Public Property Let FontRightLeading(ByVal PropvVal As Boolean)
'/* [let] right align fonts
    m_bFontRightLeading = PropvVal
End Property

Public Property Get SelectedBackColor() As Long
'/* [get] Selected ForeColor
    SelectedBackColor = m_lSelectedBackColor
End Property

Public Property Let SelectedBackColor(ByVal PropVal As Long)
'/* [let] Selected ForeColor
    m_lSelectedBackColor = PropVal
End Property

Public Property Get Enabled() As Boolean
'/* [get] Enabled
    Enabled = m_bEnabled
End Property

Public Property Let Enabled(ByVal PropVal As Boolean)
'/* [let] Enabled
    If Not (m_lTVHwnd = 0) Then
        If PropVal Then
            SendMessageLongA m_lTVHwnd, TVM_SETBKCOLOR, 0&, m_lBackColor
            SendMessageLongA m_lTVHwnd, TVM_SETTEXTCOLOR, 0&, m_lForeColor
            EnableWindow m_lTVHwnd, 1&
        Else
            SendMessageLongA m_lTVHwnd, TVM_SETBKCOLOR, 0&, m_lDisabledBackColor
            SendMessageLongA m_lTVHwnd, TVM_SETTEXTCOLOR, 0&, m_lDisabledForeColor
            EnableWindow m_lTVHwnd, 0&
        End If
    End If
    If Not (m_cSkinScrollBars Is Nothing) Then
        m_cSkinScrollBars.Enabled = PropVal
    End If
    m_bEnabled = PropVal
End Property

Public Property Get Font() As StdFont
'/* [get] retrieve list font
    If Not (m_oFont Is Nothing) Then
        Set Font = m_oFont
    End If
End Property

Public Property Set Font(ByVal oFont As StdFont)
'*/ [set] change list font

Dim lHdc        As Long
Dim lChar       As Long
Dim uLF         As LOGFONT
Dim bteFont()   As Byte

    If oFont Is Nothing Then
        DefaultFont oFont
    End If
    Set m_oFont = oFont
    If Not (m_lTVHwnd = 0) Then
        DestroyFont
        '/* extract properties
        With uLF
            bteFont = StrConv(oFont.Name, vbFromUnicode)
            For lChar = 0 To UBound(bteFont)
                .lfFaceName(lChar) = bteFont(lChar)
            Next lChar
            lHdc = GetDC(m_lTVHwnd)
            .lfHeight = -MulDiv(oFont.Size, GetDeviceCaps(lHdc, LOGPIXELSY), 72)
            .lfItalic = oFont.Italic
            .lfWeight = IIf(oFont.Bold, FW_BOLD, FW_NORMAL)
            .lfUnderline = oFont.Underline
            .lfStrikeOut = oFont.Strikethrough
            If m_bUseUnicode Then
                .lfCharSet = 134
            Else
                .lfCharSet = 3
            End If
            If m_bIsXp Then
                .lfQuality = LF_CLEARTYPE_QUALITY
            Else
                .lfQuality = LF_ANTIALIASED_QUALITY
            End If
        End With
        '/* create the log font
        If m_bUseUnicode Then
            m_lHFont = CreateFontIndirectW(uLF)
            SendMessageLongW m_lTVHwnd, WM_SETFONT, m_lHFont, True
        Else
            m_lHFont = CreateFontIndirectA(uLF)
            SendMessageLongA m_lTVHwnd, WM_SETFONT, m_lHFont, True
        End If
    End If
    ReleaseDC m_lTVHwnd, lHdc
    Set oFont = Nothing

End Property

Public Property Get ForeColor() As OLE_COLOR
'/* [get] forecolor
    ForeColor = m_lForeColor
End Property

Public Property Let ForeColor(ByVal PropVal As OLE_COLOR)
'/* [let] forecolor
    m_lForeColor = PropVal
    If Not (m_lTVHwnd = 0) Then
        OleTranslateColor PropVal, 0, m_lForeColor
        SendMessageLongA m_lTVHwnd, TVM_SETTEXTCOLOR, 0&, m_lForeColor
    End If
End Property

Public Property Get FullRowSelect() As Boolean
'/* [get] node row select
    FullRowSelect = m_bFullRowSelect
End Property

Public Property Let FullRowSelect(ByVal PropVal As Boolean)
'/* [get] node row select
    If Not (m_lTVHwnd = 0) Then
        m_bFullRowSelect = PropVal
        If PropVal Then
            WindowStyle m_lTVHwnd, GWL_STYLE, TVS_FULLROWSELECT, 0
        Else
            WindowStyle m_lTVHwnd, GWL_STYLE, 0, TVS_FULLROWSELECT
        End If
    End If
End Property

Public Property Let GridHwnd(ByVal PropVal As Long)
    m_lGridHwnd = PropVal
End Property

Public Property Get HasButtons() As Boolean
'/* [get] node has buttons
    HasButtons = m_bHasButtons
End Property

Public Property Let HasButtons(ByVal PropVal As Boolean)
'/* [get] node has buttons
    If Not (m_lTVHwnd = 0) Then
        m_bHasButtons = PropVal
        If PropVal Then
            WindowStyle m_lTVHwnd, GWL_STYLE, TVS_HASBUTTONS, 0
        Else
            WindowStyle m_lTVHwnd, GWL_STYLE, 0, TVS_HASBUTTONS
        End If
    End If
End Property

Public Property Get HasLines() As Boolean
'/* [get] nodes have lines
    HasLines = m_bHasLines
End Property

Public Property Let HasLines(ByVal PropVal As Boolean)
'/* [let] nodes have lines
    If Not (m_lTVHwnd = 0) Then
        m_bHasLines = PropVal
        If PropVal Then
            WindowStyle m_lTVHwnd, GWL_STYLE, TVS_HASLINES, 0
        Else
            WindowStyle m_lTVHwnd, GWL_STYLE, 0, TVS_HASLINES
        End If
    End If
End Property

Public Property Get HasRootLines() As Boolean
'/* [get] parent has root lines
    HasRootLines = m_bHasRootLines
End Property

Public Property Let HasRootLines(ByVal PropVal As Boolean)
'/* [let] parent has root lines
    If Not (m_lTVHwnd = 0) Then
        m_bHasRootLines = PropVal
        If m_bHasRootLines Then
            WindowStyle m_lTVHwnd, GWL_STYLE, TVS_LINESATROOT, 0
        Else
            WindowStyle m_lTVHwnd, GWL_STYLE, 0, TVS_LINESATROOT
        End If
    End If
End Property

Public Property Get Height() As Long
'/* [get] Height
    Height = m_lHeight
End Property

Public Property Let Height(ByVal PropVal As Long)
'/* [let] Height
    m_lHeight = PropVal
End Property

Public Property Get HideSelection() As Boolean
'/* [get] hide selected
    HideSelection = m_bHideSelection
End Property

Public Property Let HideSelection(ByVal PropVal As Boolean)
'/* [let] hide selected
    If Not (m_lTVHwnd = 0) Then
        m_bHideSelection = PropVal
        If PropVal Then
            WindowStyle m_lTVHwnd, GWL_STYLE, 0, TVS_SHOWSELALWAYS
        Else
            WindowStyle m_lTVHwnd, GWL_STYLE, TVS_SHOWSELALWAYS, 0
        End If
    End If
End Property

Public Property Get hwnd() As Long
'/* [get] treeview hwnd
    hwnd = m_lTVHwnd
End Property

Public Property Get InsertMarkColor() As OLE_COLOR
'/* [get] insertion mark color
    If Not (m_lTVHwnd = 0) Then
        InsertMarkColor = SendMessageLongA(m_lTVHwnd, TVM_GETINSERTMARKCOLOR, 0&, 0&)
    End If
End Property

Public Property Let InsertMarkColor(ByVal PropVal As OLE_COLOR)
'/* [let] insertion mark color
Dim lColor As Long

    If Not (m_lTVHwnd = 0) Then
        OleTranslateColor PropVal, 0&, lColor
        SendMessageLongA m_lTVHwnd, TVM_SETINSERTMARKCOLOR, 0&, lColor
    End If

End Property

Public Property Get ItemHeight() As Long
'/* [get] node height
    If Not (m_lTVHwnd = 0) Then
        ItemHeight = SendMessageLongA(m_lTVHwnd, TVM_GETITEMHEIGHT, 0&, 0&)
    End If
End Property

Public Property Let ItemHeight(ByVal PropVal As Long)
'/* [let] node height
    If Not (m_lTVHwnd = 0) Then
        SendMessageLongA m_lTVHwnd, TVM_SETITEMHEIGHT, PropVal, 0&
    End If
End Property

Public Property Get ItemIndent() As Long
'/* [get] node indent
    If Not (m_lTVHwnd = 0) Then
        ItemIndent = SendMessageLongA(m_lTVHwnd, TVM_GETINDENT, 0&, 0&)
    End If
End Property

Public Property Let ItemIndent(ByVal PropVal As Long)
'/* [let] node indent
    If Not (m_lTVHwnd = 0) Then
        SendMessageLongA m_lTVHwnd, TVM_SETINDENT, PropVal, 0&
    End If
End Property

Public Property Get LabelEdit() As Boolean
'/* [get] node editable
    LabelEdit = m_bLabelEdit
End Property

Public Property Let LabelEdit(ByVal PropVal As Boolean)
'/* [let] node editable
    If Not (m_lTVHwnd = 0) Then
        m_bLabelEdit = PropVal
        If PropVal Then
            WindowStyle m_lTVHwnd, GWL_STYLE, TVS_EDITLABELS, 0
        Else
            WindowStyle m_lTVHwnd, GWL_STYLE, 0, TVS_EDITLABELS
        End If
    End If
End Property

Public Property Get LineColor() As OLE_COLOR
'/* [get] linecolor
    If Not (m_lTVHwnd = 0) Then
        LineColor = SendMessageLongA(m_lTVHwnd, TVM_GETLINECOLOR, 0&, 0&)
    End If
End Property

Public Property Let LineColor(ByVal PropVal As OLE_COLOR)
'/* [let] linecolor
Dim lColor As Long

    If Not (m_lTVHwnd = 0) Then
        OleTranslateColor PropVal, 0, lColor
        SendMessageLongA m_lTVHwnd, TVM_SETLINECOLOR, 0&, lColor
    End If

End Property

Public Property Get NodeBold(ByVal lNode As Long) As Boolean
'/* [get] node is boldened
    If Not (m_lTVHwnd = 0) Then
        NodeBold = CBool(GetTVState(lNode, tvcStateBold))
    End If
End Property

Public Property Let NodeBold(ByVal lNode As Long, _
                             ByVal PropVal As Boolean)
'/* [let] node is boldened
    If Not (m_lTVHwnd = 0) Then
        SetTVState lNode, tvcStateBold, PropVal
    End If
End Property

Public Property Get NodeChecked(ByVal lNode As Long) As Boolean
'/* [get] node checked
    If (m_lTVHwnd And m_bCheckBoxes) Then
        NodeChecked = (GetTvStateImage(lNode) = SII_CHECKED)
    End If
End Property

Public Property Let NodeChecked(ByVal lNode As Long, _
                                ByVal PropVal As Boolean)
'/* [let] node checked
    If (m_lTVHwnd And m_bCheckBoxes) Then
        If PropVal Then
            SetTvStateImage lNode, SII_CHECKED
        Else
            SetTvStateImage lNode, SII_UNCHECKED
        End If
    End If
End Property

Public Property Get NodeChild(ByVal lNode As Long) As Long
'/* [get] node child
    If Not (m_lTVHwnd = 0) Then
        NodeChild = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_CHILD, lNode)
    End If
End Property

Public Property Get NodeChildren(ByVal lNode As Long) As Long
'/* [get] node child count
Dim lNext As Long

    If Not (m_lTVHwnd = 0) Then
        lNext = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_CHILD, lNode)
        If lNext Then
            Do
                NodeChildren = NodeChildren + 1
                lNext = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_NEXT, lNext)
            Loop Until (lNext = 0)
        End If
    End If

End Property

Public Property Get NodeCount() As Long
'/* [get] node count
    If Not (m_lTVHwnd = 0) Then
        NodeCount = SendMessageLongA(m_lTVHwnd, TVM_GETCOUNT, 0&, 0&)
    End If
End Property

Public Property Get NodeExpanded(ByVal lNode As Long) As Boolean
'/* [get] node expanded
    If Not (m_lTVHwnd = 0) Then
        NodeExpanded = GetTVState(lNode, tvcStateExpanded)
    End If
End Property

Public Property Get NodeExpandedOnce(ByVal lNode As Long) As Boolean
'/* [get] node expanded one level
    If Not (m_lTVHwnd = 0) Then
        NodeExpandedOnce = GetTVState(lNode, tvcStateExpandedOnce)
    End If
End Property

Public Property Get NodeFirstSibling(ByVal lNode As Long) As Long
'/* [get] first node sibling
Dim lPrev As Long

    lPrev = lNode
    Do
        NodeFirstSibling = lPrev
        lPrev = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_PREVIOUS, lPrev)
    Loop Until (lPrev = 0)

End Property

Public Property Get NodeFirstVisible() As Long
'/* [get] first visible node
    If Not (m_lTVHwnd = 0) Then
        If m_bIsNt Then
            NodeFirstVisible = SendMessageLongW(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_FIRSTVISIBLE, 0&)
        Else
            NodeFirstVisible = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_FIRSTVISIBLE, 0&)
        End If
    End If
End Property

Public Property Get NodeFullPath(ByVal lNode As Long, _
                                 Optional ByVal PathSeparator As String = PATH_SEPARATOR)
'/* [get] node full path
Dim lParent As Long

    If Not (m_lTVHwnd = 0) Then
        If lNode Then
            NodeFullPath = GetNodeText(lNode)
            lParent = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_PARENT, lNode)
            Do While lParent
                NodeFullPath = GetNodeText(lParent) & PathSeparator & NodeFullPath
                lParent = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_PARENT, lParent)
            Loop
        End If
    End If

End Property

Public Property Get NodeGhosted(ByVal lNode As Long) As Boolean
'/* [get] node is ghosted
    If Not (m_lTVHwnd = 0) Then
        NodeGhosted = CBool(GetTVState(lNode, tvcStateCut))
    End If
End Property

Public Property Let NodeGhosted(ByVal lNode As Long, _
                                ByVal PropVal As Boolean)
'/* [let] node is ghosted
    If Not (m_lTVHwnd = 0) Then
        SetTVState lNode, tvcStateCut, PropVal
    End If
End Property

Public Property Get NodeHilited(ByVal lNode As Long) As Boolean
'/* [get] node is hilited
    If Not (m_lTVHwnd = 0) Then
        NodeHilited = CBool(GetTVState(lNode, tvcStateDropHilited))
    End If
End Property

Public Property Let NodeHilited(ByVal lNode As Long, _
                                ByVal PropVal As Boolean)
'/* [let] node is hilited
    If Not (m_lTVHwnd = 0) Then
        SetTVState lNode, tvcStateDropHilited, PropVal
    End If
End Property

Public Property Get NodeImage(ByVal lNode As Long) As Long
'/* [get] selected node image
Dim lIndex As Long

    If Not (m_lTVHwnd = 0) Then
        lIndex = GetTvImage(lNode)
        If (lIndex < 0) Then
            NodeImage = -1
        ElseIf (lIndex > (m_lImageListCount - 1)) Then
            NodeImage = -1
        Else
            NodeImage = lIndex
        End If
    End If

End Property

Public Property Let NodeImage(ByVal lNode As Long, _
                              ByVal PropVal As Long)
'/* [let] selected node image
    If Not (m_lTVHwnd = 0) Then
        If (PropVal < 0) Then
            SetTvImage lNode, m_lImageListCount
        ElseIf (PropVal > (m_lImageListCount - 1)) Then
            SetTvImage lNode, m_lImageListCount
        Else
            SetTvImage lNode, PropVal
        End If
    End If
End Property

Public Property Get NodeLastSibling(ByVal lNode As Long) As Long
'/* [get] last node sibling
Dim lNext As Long

    lNext = lNode
    Do
        NodeLastSibling = lNext
        lNext = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_NEXT, lNext)
    Loop Until (lNext = 0)

End Property

Public Property Get NodeLastVisible() As Long
'/* [get] last visible node
Dim lItemHeight As Long
Dim tRect       As RECT
Dim tTVHI       As TVHITTESTINFO

    If Not (m_lTVHwnd = 0) Then
        GetClientRect m_lTVHwnd, tRect
        If m_bUseUnicode Then
            lItemHeight = SendMessageLongW(m_lTVHwnd, TVM_GETITEMHEIGHT, 0&, 0&)
        Else
            lItemHeight = SendMessageLongA(m_lTVHwnd, TVM_GETITEMHEIGHT, 0&, 0&)
        End If
        With tTVHI
            .pt.y = tRect.Bottom - lItemHeight
            Do While (.pt.y > -1)
                SendMessageA m_lTVHwnd, TVM_HITTEST, 0&, tTVHI
                If .hItem Then
                    NodeLastVisible = .hItem
                    Exit Do
                End If
                .pt.y = .pt.y - lItemHeight
            Loop
        End With
    End If

End Property

Public Property Get NodeLevel(ByVal lNode As Long) As Long
'/* [get] node depth level
Dim lParent As Long

    If Not (m_lTVHwnd = 0) Then
        If lNode Then
            NodeLevel = 1
            lParent = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_PARENT, lNode)
            Do While lParent
                NodeLevel = NodeLevel + 1
                lParent = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_PARENT, lParent)
            Loop
        End If
    End If

End Property

Public Property Get NodeNext(ByVal lNode As Long) As Long
'/* [get] next node
    If Not (m_lTVHwnd = 0) Then
        NodeNext = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_NEXTVISIBLE, lNode)
    End If
End Property

Public Property Get NodeNextSibling(ByVal lNode As Long) As Long
'/* [get] next node sibling
    If Not (m_lTVHwnd = 0) Then
        NodeNextSibling = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_NEXT, lNode)
    End If
End Property

Public Property Get NodeParent(ByVal lNode As Long) As Long
'/* [get] node parent
    If Not (m_lTVHwnd = 0) Then
        NodeParent = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_PARENT, lNode)
    End If
End Property

Public Property Get NodePlusMinusButton(ByVal lNode As Long) As Boolean
'/* [get] node expand button
    If Not (m_lTVHwnd = 0) Then
        NodePlusMinusButton = CBool(GetTVChildren(lNode))
    End If
End Property

Public Property Let NodePlusMinusButton(ByVal lNode As Long, _
                                        ByVal PropVal As Boolean)
'/* [let] node expand button
    If Not (m_lTVHwnd = 0) Then
        SetTVChildren lNode, IIf(PropVal, 1, 0)
    End If
End Property

Public Property Get NodePrevious(ByVal lNode As Long) As Long
'/* [get] previous node
    If Not (m_lTVHwnd = 0) Then
        NodePrevious = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_PREVIOUSVISIBLE, lNode)
    End If
End Property

Public Property Get NodePreviousSibling(ByVal lNode As Long) As Long
'/* [get] previous node sibling
    If Not (m_lTVHwnd = 0) Then
        NodePreviousSibling = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_PREVIOUS, lNode)
    End If
End Property

Public Property Get NodeRoot() As Long
'/* [get] node root
    If Not (m_lTVHwnd = 0) Then
        NodeRoot = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0&)
    End If
End Property

Public Property Get NodeSelectedImage(ByVal lNode As Long) As Long
'/* [get] selected node selected image
Dim lIndex As Long

    If Not (m_lTVHwnd = 0) Then
        lIndex = GetTvSelectedImage(lNode)
        If (lIndex < 0) Then
            NodeSelectedImage = -1
        ElseIf (lIndex > (m_lImageListCount - 1)) Then
            NodeSelectedImage = -1
        Else
            NodeSelectedImage = lIndex
        End If
    End If

End Property

Public Property Let NodeSelectedImage(ByVal lNode As Long, _
                                      ByVal PropVal As Long)
'/* [let] selected node selected image
    If Not (m_lTVHwnd = 0) Then
        If (PropVal < 0) Then
            SetTvSelectedImage lNode, m_lImageListCount
        ElseIf (PropVal > (m_lImageListCount - 1)) Then
            SetTvSelectedImage lNode, m_lImageListCount
        Else
            SetTvSelectedImage lNode, PropVal
        End If
    End If
End Property

Public Property Get NodeTag(ByVal lNode As Long) As String
'/* [get] node tag
    If Not (m_lTVHwnd = 0) Then
        NodeTag = m_uNodeData(GetTvParam(lNode)).sTag
    End If
End Property

Public Property Let NodeTag(ByVal lNode As Long, _
                            ByVal PropVal As String)
'/* [let] node tag
    If Not (m_lTVHwnd = 0) Then
        m_uNodeData(GetTvParam(lNode)).sTag = PropVal
    End If
End Property

Public Property Get NodeText(ByVal lNode As Long) As String
'/* [get] selected node text
    If Not (m_lTVHwnd = 0) Then
        NodeText = GetNodeText(lNode)
    End If
End Property

Public Property Let NodeText(ByVal lNode As Long, _
                             ByVal PropVal As String)
'/* [let] selected node text
    If Not (m_lTVHwnd = 0) Then
        SetNodeText lNode, PropVal
    End If
End Property

Public Property Get NodeVisible(ByVal lNode As Long, _
                                Optional ByVal bTextOnly As Boolean = False) As Boolean
'/* [get] node visibility
Dim tRWnd As RECT
Dim tRItm As RECT
Dim tRInt As RECT

    If Not (m_lTVHwnd = 0) Then
        GetClientRect m_lTVHwnd, tRWnd
        tRItm.left = lNode
        SendMessageA m_lTVHwnd, TVM_GETITEMRECT, -bTextOnly, tRItm
        IntersectRect tRInt, tRWnd, tRItm
        NodeVisible = CBool(EqualRect(tRInt, tRItm))
    End If

End Property

Public Property Get ScrollBarLeftAlign() As Boolean
'/* [get] scrollbar orientation
    ScrollBarLeftAlign = m_bScrollBarLeftAlign
End Property

Public Property Let ScrollBarLeftAlign(ByVal PropVal As Boolean)
'/* [let] scrollbar orientation
    If m_bInitialized Then
        If Not (m_lTVHwnd = 0) Then
            If PropVal Then
                WindowStyle m_lTVHwnd, GWL_EXSTYLE, WS_EX_LEFTSCROLLBAR, 0
            Else
                WindowStyle m_lTVHwnd, GWL_EXSTYLE, 0, WS_EX_LEFTSCROLLBAR
            End If
        End If
    End If
    m_bScrollBarLeftAlign = PropVal
End Property

Public Property Get SelectedNode() As Long
'/* [get] selected node
    If Not (m_lTVHwnd = 0) Then
        If m_bUseUnicode Then
            SelectedNode = SendMessageLongW(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_CARET, 0&)
        Else
            SelectedNode = SendMessageLongW(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_CARET, 0&)
        End If
    End If
End Property

Public Property Let SelectedNode(ByVal PropVal As Long)
'/* [let] selected node
    If Not (m_lTVHwnd = 0) Then
        If m_bUseUnicode Then
            SendMessageLongW m_lTVHwnd, TVM_SELECTITEM, TVGN_CARET, PropVal
        Else
            SendMessageLongA m_lTVHwnd, TVM_SELECTITEM, TVGN_CARET, PropVal
        End If
    End If
End Property

Public Property Get SingleExpand() As Boolean
'/* [get] parent expands nodes
    SingleExpand = m_bSingleExpand
End Property

Public Property Let SingleExpand(ByVal PropVal As Boolean)
'/* [let] parent expands nodes
    If Not (m_lTVHwnd = 0) Then
        m_bSingleExpand = PropVal
        If PropVal Then
            WindowStyle m_lTVHwnd, GWL_STYLE, TVS_SINGLEEXPAND, 0
        Else
            WindowStyle m_lTVHwnd, GWL_STYLE, 0, TVS_SINGLEEXPAND
        End If
    End If
End Property

Public Property Get TrackSelect() As Boolean
'/* [get] track selected nodes
    TrackSelect = m_bTrackSelect
End Property

Public Property Let TrackSelect(ByVal PropVal As Boolean)
'/* [let] track selected nodes
    If Not (m_lTVHwnd = 0) Then
        m_bTrackSelect = PropVal
        If PropVal Then
            WindowStyle m_lTVHwnd, GWL_STYLE, TVS_TRACKSELECT, 0
        Else
            WindowStyle m_lTVHwnd, GWL_STYLE, 0, TVS_TRACKSELECT
        End If
    End If
End Property

Public Property Get UseUnicode() As Boolean
'/* [get] unicode state
    UseUnicode = m_bUseUnicode
End Property

Public Property Let UseUnicode(ByVal PropVal As Boolean)
'/* [let] unicode state
    If m_bIsNt Then
        If PropVal Then
            m_bUseUnicode = True
        Else
            m_bUseUnicode = False
        End If
        If Not (m_lTVHwnd = 0) Then
            SetUnicode m_bUseUnicode
        End If
    End If
End Property

Public Property Get Visible() As Boolean
'/* [get] Visible
    Visible = m_bVisible
End Property

Public Property Let Visible(ByVal PropVal As Boolean)
'/* [let] Visible
    If Not (m_lTVHwnd = 0) Then
        If Not (m_cSkinScrollBars Is Nothing) Then
            With m_cSkinScrollBars
                .Visible = PropVal
                If PropVal Then
                    .Refresh
                End If
            End With
        End If
    End If
    m_bVisible = PropVal
End Property

Public Property Get Width() As Long
'/* [get] Width
    Width = m_lWidth
End Property

Public Property Let Width(ByVal PropVal As Long)
'/* [let] Width
    m_lWidth = PropVal
End Property

'**********************************************************************
'*                              GRAPHICS
'**********************************************************************

Private Function GetMaskColor(ByVal lHdc As Long) As Long
'/* get checkbox mask color
    GetMaskColor = GetPixel(lHdc, 0&, 0&)
End Function

Public Sub ImlStateAddBmp(ByVal lBitmap As Long, _
                          Optional ByVal lMaskColor As Long = CLR_NONE)

'*/ add a bitmap to header iml

On Error GoTo Handler

    If Not (m_lImlStateHndl = 0) Then
        If Not (lMaskColor = CLR_NONE) Then
            ImageList_AddMasked m_lImlStateHndl, lBitmap, lMaskColor
        Else
            ImageList_Add m_lImlStateHndl, lBitmap, 0&
        End If
    End If

Handler:
    On Error GoTo 0

End Sub

Public Sub ImlClear(ByVal lImlHnd As Long)

Dim lCt     As Long
Dim lCount  As Long

    If Not (lImlHnd = 0) Then
        lCount = ImageList_GetImageCount(lImlHnd) - 1
        For lCt = lCount To 0 Step -1
            ImlRemoveIcon lImlHnd, lCt
        Next lCt
    End If

End Sub

Public Sub ImlRemoveIcon(ByVal lImlHnd As Long, _
                         ByVal lIndex As Long)

    If Not (lImlHnd = 0) Then
        ImageList_Remove lImlHnd, lIndex
    End If

End Sub

Public Property Get hImlNode() As Long
    hImlNode = m_lImlNodeHndl
End Property

Public Property Let hImlNode(ByVal PropVal As Long)
'/* load external imagelist
    If Not (m_lTVHwnd = 0) Then
        If Not (PropVal = 0) Then
            m_lImlNodeHndl = PropVal
            SendMessageLongA m_lTVHwnd, TVM_SETIMAGELIST, TVSIL_NORMAL, m_lImlNodeHndl
        End If
    End If
    m_lImageListCount = ImageList_GetImageCount(m_lImlNodeHndl)
    
End Property

Public Sub InitImlState()
'*/ initialize header imagelist

    If Not (m_lTVHwnd = 0) Then
        DestroyImlState
        m_lImlStateHndl = ImageList_Create(16&, 16&, ILC_COLOR32 Or ILC_MASK, 0&, 0&)
        If Not (m_lImlStateHndl = 0) Then
            SendMessageLongA m_lTVHwnd, TVM_SETIMAGELIST, TVSIL_STATE, m_lImlStateHndl
        End If
    End If

End Sub

Private Function LoadCheckBoxImages(ByVal eCheckBoxStyle As ETVThemeStyle) As Boolean
'/* load checkbox skin images

    ResetSkinnedCheckboxes

    Select Case eCheckBoxStyle
    '/* azure
    Case 0
        Set m_IChecked = LoadResPicture("AZURE-CHKPUSHED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("AZURE-CHKEMPTY", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("AZURE-CHKDISABLED", vbResBitmap)
    '/* classic
    Case 1
        Set m_IChecked = LoadResPicture("CLASSIC-CHKPUSHED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("CLASSIC-CHKEMPTY", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("CLASSIC-CHKDISABLED", vbResBitmap)
    '/* gloss
    Case 2
        Set m_IChecked = LoadResPicture("GLOSS-CHKPUSHED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("GLOSS-CHKEMPTY", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("GLOSS-CHKDISABLED", vbResBitmap)
    '/* metal
    Case 3
        Set m_IChecked = LoadResPicture("METAL-CHKPUSHED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("METAL-CHKEMPTY", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("METAL-CHKDISABLED", vbResBitmap)
    '/* xp
    Case 4, 5, 6
        Set m_IChecked = LoadResPicture("XP-CHKPUSHED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("XP-CHKEMPTY", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("XP-CHKDISABLED", vbResBitmap)
    '/* vista
    Case 7
        Set m_IChecked = LoadResPicture("VISTA-CHKPUSHED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("VISTA-CHKEMPTY", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("VISTA-CHKDISABLED", vbResBitmap)
    '/* silver
    Case 8
        Set m_IChecked = LoadResPicture("METAL-CHKPUSHED", vbResBitmap)
        Set m_IUnChecked = LoadResPicture("METAL-CHKEMPTY", vbResBitmap)
        Set m_IChkDisabled = LoadResPicture("METAL-CHKDISABLED", vbResBitmap)
    End Select

    '/* success
    LoadCheckBoxImages = True

End Function

Private Sub ResetSkinnedCheckboxes()
'/* reset checkbox images

    ImlClear m_lImlStateHndl
    If Not m_IChecked Is Nothing Then Set m_IChecked = Nothing
    If Not m_cChkCheckDc Is Nothing Then Set m_cChkCheckDc = Nothing
    If Not m_IUnChecked Is Nothing Then Set m_IUnChecked = Nothing
    If Not m_IChkDisabled Is Nothing Then Set m_IChkDisabled = Nothing

End Sub

Public Sub SkinCheckBox(ByVal eCheckBoxStyle As ETVThemeStyle, _
                        Optional ByVal lThemecolor As Long = -1, _
                        Optional ByVal sLuminence As Single = 0.2)

'/* skin the listview checkboxes

Dim lMask As Long

    '/* system image sizes
    InitImlState
    '/* load images
    LoadCheckBoxImages eCheckBoxStyle
    '/* image dc's
    Set m_cChkCheckDc = New clsStoreDc
    '/* create dc's
    If Not (lThemecolor = -1) Then
        With m_cChkCheckDc
            '/* create image dc
            .CreateFromPicture m_IChkDisabled
            '/* colorize
            .ColorizeImage lThemecolor, sLuminence
            '/* new mask color
            lMask = GetMaskColor(.hdc)
            '/* extract bitmap handle
            ImlStateAddBmp .ExtractBitmap, lMask
        End With
        Set m_cChkCheckDc = Nothing
        Set m_cChkCheckDc = New clsStoreDc
        With m_cChkCheckDc
            .CreateFromPicture m_IUnChecked
            .ColorizeImage lThemecolor, sLuminence
            ImlStateAddBmp .ExtractBitmap, lMask
        End With
        Set m_cChkCheckDc = Nothing
        Set m_cChkCheckDc = New clsStoreDc
        With m_cChkCheckDc
            .CreateFromPicture m_IChecked
            .ColorizeImage lThemecolor, sLuminence
            ImlStateAddBmp .ExtractBitmap, lMask
        End With
        Set m_cChkCheckDc = Nothing
    Else
        ImlStateAddBmp m_IChkDisabled.handle, &HFF00FF
        ImlStateAddBmp m_IUnChecked.handle, &HFF00FF
        ImlStateAddBmp m_IChecked.handle, &HFF00FF
    End If

End Sub

Public Sub SkinScrollBars(ByVal eStyle As ETVThemeStyle, _
                          Optional ByVal lThemecolor As Long = -1, _
                          Optional sLuminence As Single = 0)
'/* skin scrollbars

    '/* reset dc class
    If Not (m_cSkinScrollBars Is Nothing) Then
        m_cSkinScrollBars.ResetScrollBarSkin
    Else
        Set m_cSkinScrollBars = New clsSkinScrollbars
    End If
    '/* set properties
    '/* pass though to skinscrollbars class
    With m_cSkinScrollBars
        .Visible = False
        If Not (lThemecolor = -1) Then
            .UseScrollBarTheme = True
        Else
            .UseScrollBarTheme = False
        End If
        .ScrollThemeColor = lThemecolor
        .ScrollLuminence = sLuminence
        .ScrollBarSkinStyle = eStyle
        .SkinScrollBar = True
        .LoadSkin m_lTVHwnd, m_lParentHwnd
    End With

End Sub


'**********************************************************************
'*                              FUNCTIONS
'**********************************************************************

Friend Function AddNode(Optional ByVal lRelative As Long, _
                        Optional ByVal eRelation As Long, _
                        Optional ByVal sKey As String, _
                        Optional ByVal sText As String, _
                        Optional ByVal lImage As Long = -1, _
                        Optional ByVal lSelectedImage As Long = -1, _
                        Optional ByVal bPlusButton As Boolean = False, _
                        Optional ByVal sTag As String = vbNullString) As Long

    If Not (m_lTVHwnd = 0) Then
        AddNode = AddNodeData(m_lNodeCount + 1, lRelative, eRelation, sText, lImage, lSelectedImage, bPlusButton)
        If AddNode Then
            On Error GoTo Handler
            m_cKey.Add AddNode, sKey
            m_lNodeCount = m_lNodeCount + 1
            If (m_lNodeCount Mod ALLOCATE_SIZE = 1) Then
                ReDim Preserve m_uNodeData(0 To m_lNodeCount + ALLOCATE_SIZE)
            End If
            With m_uNodeData(m_lNodeCount)
                .sKey = sKey
                .hNode = AddNode
                .sTag = sTag
            End With
        End If
    End If

Exit Function

Handler:
    SendMessageLongA m_lTVHwnd, TVM_DELETEITEM, 0&, AddNode
    AddNode = 0

End Function

Private Function AddNodeData(ByVal pvlID As Long, _
                             ByVal lRelative As Long, _
                             ByVal eRelation As Long, _
                             ByVal sText As String, _
                             ByVal lImage As Long, _
                             ByVal lSelectedImage As Long, _
                             ByVal bForcePlusButton As Boolean) As Long

Dim lPrevious As Long
Dim uTVISA    As TVINSERTSTRUCTA
Dim uTVISW    As TVINSERTSTRUCTW

    If m_bUseUnicode Then
        With uTVISW
            With .Item
                .Mask = TVIF_all
                .lParam = pvlID
                sText = sText & vbNullChar
                .pszText = StrPtr(sText)
                .cchTextMax = LenB(sText) + 1
                If (lImage > -1) Then
                    If (lImage < m_lImageListCount) Then
                        .iImage = lImage
                    Else
                        .iImage = m_lImageListCount
                    End If
                Else
                    .iImage = m_lImageListCount
                End If
                If (lSelectedImage > -1) Then
                    If (lSelectedImage < m_lImageListCount) Then
                        .iSelectedImage = lSelectedImage
                    Else
                        .iSelectedImage = m_lImageListCount
                    End If
                Else
                    .iSelectedImage = m_lImageListCount
                End If
                If bForcePlusButton Then
                    .cChildren = 1
                    .Mask = .Mask Or TVIF_CHILDREN
                End If
            End With
            If lRelative Then
                .hParent = lRelative
            Else
                .hParent = TVI_ROOT
            End If
            Select Case eRelation
            Case trnFirst
                .hInsertAfter = TVI_FIRST
            Case trnLast
                .hInsertAfter = TVI_LAST
            Case trnSort
                .hInsertAfter = TVI_SORT
            Case trnNext
                .hParent = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_PARENT, lRelative)
                .hInsertAfter = lRelative
            Case trnPrevious
                .hParent = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_PARENT, lRelative)
                lPrevious = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_PREVIOUS, lRelative)
                If lPrevious Then
                    .hInsertAfter = lPrevious
                Else
                    .hInsertAfter = TVI_FIRST
                End If
            Case Else
                .hInsertAfter = TVI_LAST
            End Select
        End With
        AddNodeData = SendMessageA(m_lTVHwnd, TVM_INSERTITEMW, 0&, uTVISW)
    Else
        With uTVISA
            With .Item
                .Mask = TVIF_all
                .lParam = pvlID
                .pszText = sText & vbNullChar
                .cchTextMax = Len(sText) + 1
                If (lImage > -1) Then
                    If (lImage < m_lImageListCount) Then
                        .iImage = lImage
                    Else
                        .iImage = m_lImageListCount
                    End If
                Else
                    .iImage = m_lImageListCount
                End If
                If (lSelectedImage > -1) Then
                    If (lSelectedImage < m_lImageListCount) Then
                        .iSelectedImage = lSelectedImage
                    Else
                        .iSelectedImage = m_lImageListCount
                    End If
                Else
                    .iSelectedImage = m_lImageListCount
                End If
                If bForcePlusButton Then
                    .cChildren = 1
                    .Mask = .Mask Or TVIF_CHILDREN
                End If
            End With
            If lRelative Then
                .hParent = lRelative
            Else
                .hParent = TVI_ROOT
            End If
            Select Case eRelation
            Case trnFirst
                .hInsertAfter = TVI_FIRST
            Case trnLast
                .hInsertAfter = TVI_LAST
            Case trnSort
                .hInsertAfter = TVI_SORT
            Case trnNext
                .hParent = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_PARENT, lRelative)
                .hInsertAfter = lRelative
            Case trnPrevious
                .hParent = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_PARENT, lRelative)
                lPrevious = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_PREVIOUS, lRelative)
                If lPrevious Then
                    .hInsertAfter = lPrevious
                Else
                    .hInsertAfter = TVI_FIRST
                End If
            Case Else
                .hInsertAfter = TVI_LAST
            End Select
        End With
        AddNodeData = SendMessageA(m_lTVHwnd, TVM_INSERTITEMA, 0&, uTVISA)
    End If

End Function

Public Sub CheckChildren(ByVal hNode As Long, _
                         ByVal New_NodeChecked As Boolean)

Dim hNext As Long

    If Not (m_lTVHwnd = 0) Then
        hNext = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_CHILD, hNode)
        Do While hNext
            If New_NodeChecked Then
                SetTvStateImage hNext, SII_CHECKED
            Else
                SetTvStateImage hNext, SII_UNCHECKED
            End If
            CheckChildren hNext, New_NodeChecked
            hNext = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_NEXT, hNext)
        Loop
    End If

End Sub

Public Sub Clear()

    If Not (m_lTVHwnd = 0) Then
        SendMessageLongA m_lTVHwnd, TVM_DELETEITEM, 0, TVI_ROOT
        ReDim m_uNodeData(0 To 1)
        Set m_cKey = Nothing
        Set m_cKey = New Collection
        m_lNodeCount = 0
    End If

End Sub

Public Sub Collapse(ByVal hNode As Long, _
                    Optional ByVal CollapseChildren As Boolean = False)

Dim hNext As Long

    If Not (m_lTVHwnd = 0) Then
        SendMessageLongA m_lTVHwnd, TVM_EXPAND, TVE_COLLAPSE, hNode
        If CollapseChildren Then
            hNext = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_CHILD, hNode)
            Do While hNext
                Collapse hNext, CollapseChildren:=True
                hNext = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_NEXT, hNext)
            Loop
        End If
    End If

End Sub

Public Function DeleteNode(ByVal hNode As Long) As Boolean

    If Not (m_lTVHwnd = 0) Then
        If (SendMessageLongA(m_lTVHwnd, TVM_DELETEITEM, 0, hNode)) Then
            If Not m_bHoldDeletePostProcess Then
                NodeDeletePostProcess
            End If
            DeleteNode = True
        End If
    End If

End Function

Public Sub EnsureVisible(ByVal hNode As Long)

    If Not (m_lTVHwnd = 0) Then
        SendMessageLongA m_lTVHwnd, TVM_ENSUREVISIBLE, 0, hNode
    End If

End Sub

Public Sub Expand(ByVal hNode As Long, _
                  Optional ByVal ExpandChildren As Boolean = False)

Dim hNext As Long

    If Not (m_lTVHwnd = 0) Then
        SendMessageLongA m_lTVHwnd, TVM_EXPAND, TVE_EXPAND, hNode
        If ExpandChildren Then
            hNext = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_CHILD, hNode)
            Do While hNext
                Expand hNext, ExpandChildren:=True
                hNext = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_NEXT, hNext)
            Loop
        End If
    End If

End Sub

Public Function HitTest(Optional ByVal FullRowHit As Boolean = True) As Long

Dim hNode As Long
Dim lfHit As Long

    If Not (m_lTVHwnd = 0) Then
        NodeHitTest hNode, lfHit
        If FullRowHit Then
            HitTest = hNode
        Else
            If (lfHit And TVHT_ONITEM) Then
                HitTest = hNode
            End If
        End If
    End If

End Function

Private Function LeftKeyState() As Boolean
'/* left button pressed state

    If ((GetKeyState(VK_LBUTTON) And &H80) > 1) Then
        LeftKeyState = True
    End If

End Function

Public Sub SetRedrawMode(ByVal Enable As Boolean)

    If Not (m_lTVHwnd = 0) Then
        SendMessageLongA m_lTVHwnd, WM_SETREDRAW, -Enable, 0
    End If

End Sub

Public Sub SortChildren(ByVal hNode As Long, _
                        Optional ByVal SortAllLevels As Boolean = False)

'Don't know why, but fRecurse param. is not working.
'So, sort recursively using recursive call.

Dim hNext As Long

    If Not (m_lTVHwnd = 0) Then
        SendMessageLongA m_lTVHwnd, TVM_SORTCHILDREN, 0, hNode
        If SortAllLevels Then
            hNext = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_CHILD, hNode)
            Do While hNext
                SortChildren hNext, SortAllLevels:=True
                hNext = SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_NEXT, hNext)
            Loop
        End If
    End If

End Sub

Private Sub WindowStyle(ByVal lHwnd As Long, _
                        ByVal lType As Long, _
                        ByVal lStyle As Long, _
                        ByVal lStyleNot As Long)

'/* set window style bits

Dim lNewStyle As Long

    If m_bIsNt Then
        lNewStyle = GetWindowLongW(lHwnd, lType)
    Else
        lNewStyle = GetWindowLongA(lHwnd, lType)
    End If
    lNewStyle = (lNewStyle And Not lStyleNot) Or lStyle
    If m_bIsNt Then
        SetWindowLongW lHwnd, lType, lNewStyle
    Else
        SetWindowLongA lHwnd, lType, lNewStyle
    End If
    SetWindowPos lHwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED

End Sub

Public Sub HoldDeletePostProcess(ByVal Hold As Boolean)

    m_bHoldDeletePostProcess = Hold
    If Not Hold Then
        NodeDeletePostProcess
    End If

End Sub

Public Function IsValidKey(ByVal Key As String) As Boolean

On Error GoTo Handler

    m_cKey.Add 0, Key
    m_cKey.Remove m_cKey.Count
    IsValidKey = True

Handler:

End Function

Public Function GetKeyNode(ByVal Key As String) As Long

    If Not (m_lTVHwnd = 0) Then
        GetKeyNode = m_cKey(Key)
    End If

End Function

Public Function GetNodeKey(ByVal hNode As Long) As String

    If Not (m_lTVHwnd = 0) Then
        GetNodeKey = m_uNodeData(GetTvParam(hNode)).sKey
    End If

End Function

Public Function StartLabelEdit(ByVal hNode As Long) As Boolean

    If m_bUseUnicode Then
        If Not (m_lTVHwnd = 0) Then
            StartLabelEdit = CBool(SendMessageLongW(m_lTVHwnd, TVM_EDITLABELW, 0&, hNode))
        End If
    Else
        If Not (m_lTVHwnd = 0) Then
            StartLabelEdit = CBool(SendMessageLongA(m_lTVHwnd, TVM_EDITLABELA, 0&, hNode))
        End If
    End If

End Function

Public Sub EndLabelEdit(ByVal Cancel As Boolean)

    If Not (m_lTVHwnd = 0) Then
        SendMessageLongA m_lTVHwnd, TVM_ENDEDITLABELNOW, -Cancel, 0
    End If

End Sub

Public Sub Refresh(Optional ByVal bErase As Boolean)
'/* refresh the listview

Dim tRect As RECT

    If Not (m_lTVHwnd = 0) Then
        GetWindowRect m_lTVHwnd, tRect
        OffsetRect tRect, -tRect.left, -tRect.top
        If bErase Then
            EraseRect m_lTVHwnd, tRect, 1&
        Else
            EraseRect m_lTVHwnd, tRect, 0&
        End If
    End If

End Sub

Public Sub RePaint()
'/* repaint grid

    If Not (m_lTVHwnd = 0) Then
        SendMessageLongA m_lTVHwnd, WM_PAINT, 0&, 0&
    End If

End Sub

Public Sub SetInsertionMark(ByVal hNode As Long, _
                            Optional ByVal InsertAfter As Boolean = True)

    If Not (m_lTVHwnd = 0) Then
        SendMessageLongA m_lTVHwnd, TVM_SETINSERTMARK, -InsertAfter, hNode
    End If

End Sub

Public Sub SetHilitedNode(ByVal hNode As Long)

    If Not (m_lTVHwnd = 0) Then
        SendMessageLongA m_lTVHwnd, TVM_SELECTITEM, TVGN_DROPHILITE, hNode
    End If

End Sub

Public Sub Scroll(ByVal Direction As ETVScrollConstants)

    Select Case Direction
    Case tvsHome
        SendMessageLongA m_lTVHwnd, WM_VSCROLL, SB_TOP, 0&
    Case tvsPageUp
        SendMessageLongA m_lTVHwnd, WM_VSCROLL, SB_PAGEUP, 0&
    Case tvsUp
        SendMessageLongA m_lTVHwnd, WM_VSCROLL, SB_LINEUP, 0&
    Case tvsDown
        SendMessageLongA m_lTVHwnd, WM_VSCROLL, SB_LINEDOWN, 0&
    Case tvsPageDown
        SendMessageLongA m_lTVHwnd, WM_VSCROLL, SB_PAGEDOWN, 0&
    Case tvsEnd
        SendMessageLongA m_lTVHwnd, WM_VSCROLL, SB_BOTTOM, 0&
    Case tvsLeft
        SendMessageLongA m_lTVHwnd, WM_HSCROLL, SB_LEFT, 0&
    Case tvsPageLeft
        SendMessageLongA m_lTVHwnd, WM_HSCROLL, SB_PAGELEFT, 0&
    Case tvsLineLeft
        SendMessageLongA m_lTVHwnd, WM_HSCROLL, SB_LINELEFT, 0&
    Case tvsLineRight
        SendMessageLongA m_lTVHwnd, WM_HSCROLL, SB_LINERIGHT, 0&
    Case tvsPageRight
        SendMessageLongA m_lTVHwnd, WM_HSCROLL, SB_PAGERIGHT, 0&
    Case tvsRight
        SendMessageLongA m_lTVHwnd, WM_HSCROLL, SB_RIGHT, 0&
    End Select

End Sub

Private Function GetNodeText(ByVal lNode As Long) As String

Dim lLen    As Long
Dim uTVW   As TVITEMW
Dim aBuf() As Byte

    ReDim aBuf(MAX_PATH) As Byte
    With uTVW
        .hItem = lNode
        .Mask = TVIF_TEXT
        .pszText = VarPtr(aBuf(0))
        .cchTextMax = MAX_PATH + 1
        If m_bUseUnicode Then
            SendMessageW m_lTVHwnd, TVM_GETITEMW, 0&, uTVW
            GetNodeText = PointerToString(.pszText)
        Else
            SendMessageA m_lTVHwnd, TVM_GETITEMA, 0&, uTVW
            GetNodeText = StrConv(aBuf(), vbUnicode)
            lLen = InStr(GetNodeText, Chr$(0))
            If (lLen > 0) Then
                GetNodeText = left$(GetNodeText, lLen - 1)
            End If
        End If
    End With
    

End Function

Private Sub SetNodeText(ByVal lNode As Long, _
                        ByVal sText As String)

Dim uTVA As TVITEMA
Dim uTVW As TVITEMW

    sText = sText & Chr$(0)
    If m_bUseUnicode Then
        With uTVW
            .hItem = lNode
            .Mask = TVIF_TEXT
            .pszText = StrPtr(sText)
            .cchTextMax = LenB(sText)
        End With
        SendMessageW m_lTVHwnd, TVM_SETITEMW, 0&, uTVW
    Else
        With uTVA
            .hItem = lNode
            .Mask = TVIF_TEXT
            .pszText = sText
            .cchTextMax = Len(sText)
        End With
        SendMessageA m_lTVHwnd, TVM_SETITEMA, 0&, uTVA
    End If

End Sub

Private Function GetTvParam(ByVal lNode As Long) As Long

Dim uTVA As TVITEMA
Dim uTVW As TVITEMW

    If m_bUseUnicode Then
        With uTVW
            .hItem = lNode
            .Mask = TVIF_PARAM
            If (SendMessageW(m_lTVHwnd, TVM_GETITEMW, 0&, uTVW)) Then
                GetTvParam = .lParam
            End If
        End With
    Else
        With uTVA
            .hItem = lNode
            .Mask = TVIF_PARAM
            If (SendMessageA(m_lTVHwnd, TVM_GETITEMA, 0&, uTVA)) Then
                GetTvParam = .lParam
            End If
        End With
    End If

End Function

Private Sub SetTvParam(ByVal hNode As Long, _
                          ByVal lParam As Long)

Dim uTVA As TVITEMA
Dim uTVW As TVITEMW

    If m_bUseUnicode Then
        With uTVW
            .hItem = hNode
            .Mask = TVIF_PARAM
            .lParam = lParam
        End With
        SendMessageW m_lTVHwnd, TVM_SETITEMW, 0&, uTVW
    Else
        With uTVA
            .hItem = hNode
            .Mask = TVIF_PARAM
            .lParam = lParam
        End With
        SendMessageA m_lTVHwnd, TVM_SETITEMA, 0&, uTVA
    End If

End Sub

Private Function GetTvImage(ByVal lNode As Long) As Long

Dim uTVA As TVITEMA
Dim uTVW As TVITEMW

    If m_bUseUnicode Then
        With uTVW
            .hItem = lNode
            .Mask = TVIF_IMAGE
            If (SendMessageW(m_lTVHwnd, TVM_GETITEMW, 0&, uTVW)) Then
                GetTvImage = .iImage
            End If
        End With
    Else
        With uTVA
            .hItem = lNode
            .Mask = TVIF_IMAGE
            If (SendMessageA(m_lTVHwnd, TVM_GETITEMA, 0&, uTVA)) Then
                GetTvImage = .iImage
            End If
        End With
    End If

End Function

Private Sub SetTvImage(ByVal lNode As Long, _
                       ByVal lImage As Long)

Dim uTVA As TVITEMA
Dim uTVW As TVITEMW

    If m_bUseUnicode Then
        With uTVW
            .hItem = lNode
            .Mask = TVIF_IMAGE
            .iImage = lImage
        End With
        SendMessageW m_lTVHwnd, TVM_SETITEMW, 0&, uTVW
    Else
        With uTVA
            .hItem = lNode
            .Mask = TVIF_IMAGE
            .iImage = lImage
        End With
        SendMessageA m_lTVHwnd, TVM_SETITEMA, 0&, uTVA
    End If

End Sub

Private Function GetTvSelectedImage(ByVal lNode As Long) As Long

Dim uTVA As TVITEMA
Dim uTVW As TVITEMW

    If m_bUseUnicode Then
        With uTVW
            .hItem = lNode
            .Mask = TVIF_SELECTEDIMAGE
            If (SendMessageW(m_lTVHwnd, TVM_GETITEMW, 0&, uTVW)) Then
                GetTvSelectedImage = .iSelectedImage
            End If
        End With
    Else
        With uTVA
            .hItem = lNode
            .Mask = TVIF_SELECTEDIMAGE
            If (SendMessageA(m_lTVHwnd, TVM_GETITEMA, 0&, uTVA)) Then
                GetTvSelectedImage = .iSelectedImage
            End If
        End With
    End If

End Function

Private Sub SetTvSelectedImage(ByVal lNode As Long, _
                               ByVal lSelectedImage As Long)

Dim uTVA As TVITEMA
Dim uTVW As TVITEMW

    If m_bUseUnicode Then
        With uTVW
            .hItem = lNode
            .Mask = TVIF_SELECTEDIMAGE
            .iSelectedImage = lSelectedImage
        End With
        SendMessageW m_lTVHwnd, TVM_SETITEMW, 0&, uTVW
    Else
        With uTVA
            .hItem = lNode
            .Mask = TVIF_SELECTEDIMAGE
            .iSelectedImage = lSelectedImage
        End With
        SendMessageA m_lTVHwnd, TVM_SETITEMA, 0&, uTVA
    End If

End Sub

Private Function GetTvStateImage(ByVal lNode As Long) As Long

Dim uTVA As TVITEMA
Dim uTVW As TVITEMW

    If m_bUseUnicode Then
        With uTVW
            .hItem = lNode
            .Mask = TVIF_HANDLE Or TVIF_STATE
            .stateMask = TVIS_STATEIMAGEMASK
            If (SendMessageW(m_lTVHwnd, TVM_GETITEMW, 0&, uTVW)) Then
                GetTvStateImage = STATEIMAGEMASKTOINDEX(.State And TVIS_STATEIMAGEMASK)
            End If
        End With
    Else
        With uTVA
            .hItem = lNode
            .Mask = TVIF_HANDLE Or TVIF_STATE
            .stateMask = TVIS_STATEIMAGEMASK
            If (SendMessageA(m_lTVHwnd, TVM_GETITEMA, 0&, uTVA)) Then
                GetTvStateImage = STATEIMAGEMASKTOINDEX(.State And TVIS_STATEIMAGEMASK)
            End If
        End With
    End If
    
End Function

Private Sub SetTvStateImage(ByVal lNode As Long, _
                            ByVal lIndex As Long)

Dim uTVA As TVITEMA
Dim uTVW As TVITEMW

    If m_bUseUnicode Then
        With uTVW
            .hItem = lNode
            .Mask = TVIF_HANDLE Or TVIF_STATE
            .stateMask = TVIS_STATEIMAGEMASK
            .State = INDEXTOSTATEIMAGEMASK(lIndex)
        End With
        SendMessageW m_lTVHwnd, TVM_SETITEMW, 0&, uTVW
    Else
        With uTVA
            .hItem = lNode
            .Mask = TVIF_HANDLE Or TVIF_STATE
            .stateMask = TVIS_STATEIMAGEMASK
            .State = INDEXTOSTATEIMAGEMASK(lIndex)
        End With
        SendMessageA m_lTVHwnd, TVM_SETITEMA, 0&, uTVA
    End If

End Sub

Private Function INDEXTOSTATEIMAGEMASK(ByVal lIndex As Long) As Long
    INDEXTOSTATEIMAGEMASK = lIndex * (2 ^ 12)
End Function

Private Function STATEIMAGEMASKTOINDEX(ByVal lState As Long) As Long
    STATEIMAGEMASKTOINDEX = lState / (2 ^ 12)
End Function

Private Function GetTVState(ByVal lNode As Long, _
                            ByVal lState As ETVStateConstants) As Boolean

Dim uTVA As TVITEMA
Dim uTVW As TVITEMW

    If m_bUseUnicode Then
        With uTVW
            .hItem = lNode
            .Mask = TVIF_HANDLE Or TVIF_STATE
            If (SendMessageW(m_lTVHwnd, TVM_GETITEMW, 0&, uTVW)) Then
                GetTVState = (.State And lState)
            End If
        End With
    Else
        With uTVA
            .hItem = lNode
            .Mask = TVIF_HANDLE Or TVIF_STATE
            If (SendMessageA(m_lTVHwnd, TVM_GETITEMA, 0&, uTVA)) Then
                GetTVState = (.State And lState)
            End If
        End With
    End If

End Function

Private Sub SetTVState(ByVal lNode As Long, _
                       ByVal lState As ETVStateConstants, _
                       ByVal bAdd As Boolean)

Dim uTVA As TVITEMA
Dim uTVW As TVITEMW

    If m_bUseUnicode Then
        With uTVW
            .hItem = lNode
            .Mask = TVIF_HANDLE Or TVIF_STATE
            .stateMask = lState
            .State = bAdd And lState
        End With
        SendMessageW m_lTVHwnd, TVM_SETITEMW, 0&, uTVW
    Else
        With uTVA
            .hItem = lNode
            .Mask = TVIF_HANDLE Or TVIF_STATE
            .stateMask = lState
            .State = bAdd And lState
        End With
        SendMessageA m_lTVHwnd, TVM_SETITEMA, 0&, uTVA
    End If

End Sub

Private Function GetTVChildren(ByVal lNode As Long) As Long

Dim uTVA As TVITEMA
Dim uTVW As TVITEMW

    If m_bUseUnicode Then
        With uTVW
            .hItem = lNode
            .Mask = TVIF_CHILDREN
            If (SendMessageW(m_lTVHwnd, TVM_GETITEMW, 0&, uTVW)) Then
                GetTVChildren = .cChildren
            End If
        End With
    Else
        With uTVA
            .hItem = lNode
            .Mask = TVIF_CHILDREN
            If (SendMessageA(m_lTVHwnd, TVM_GETITEMA, 0&, uTVA)) Then
                GetTVChildren = .cChildren
            End If
        End With
    End If

End Function

Private Sub SetTVChildren(ByVal lNode As Long, _
                          ByVal lChildren As Long)

Dim uTVA As TVITEMA
Dim uTVW As TVITEMW

    If m_bUseUnicode Then
        With uTVW
            .hItem = lNode
            .Mask = TVIF_CHILDREN
            .cChildren = lChildren
        End With
        SendMessageW m_lTVHwnd, TVM_SETITEMW, 0, uTVW
    Else
        With uTVA
            .hItem = lNode
            .Mask = TVIF_CHILDREN
            .cChildren = lChildren
        End With
        SendMessageA m_lTVHwnd, TVM_SETITEMA, 0, uTVA
    End If

End Sub

Private Sub NodeHitTest(lNode As Long, _
                        lfHit As Long)

Dim uTVHI As TVHITTESTINFO

    With uTVHI
        GetCursorPos .pt
        ScreenToClient m_lTVHwnd, .pt
        SendMessageA m_lTVHwnd, TVM_HITTEST, 0, uTVHI
        lfHit = .flags
        lNode = .hItem
    End With

End Sub

Private Function NodeHasChildren(ByVal lNode As Long) As Boolean

    NodeHasChildren = CBool(SendMessageLongA(m_lTVHwnd, TVM_GETNEXTITEM, TVGN_CHILD, lNode))

End Function

Private Sub NodeGetRect(ByVal lNode As Long, _
                        tRect As RECT)

    tRect.left = lNode
    SendMessageA m_lTVHwnd, TVM_GETITEMRECT, 0&, tRect

End Sub

Private Sub NodeDeletePostProcess()

Dim lNode As Long
Dim lLast As Long
Dim uTVA  As TVITEMA
Dim uTVW  As TVITEMW

    For lNode = m_lNodeCount To 1 Step -1
        If (m_uNodeData(lNode).hNode = 0) Then
            m_cKey.Remove m_uNodeData(lNode).sKey
        End If
    Next lNode
    '-- Remove (move) array items
    uTVA.Mask = TVIF_PARAM
    uTVW.Mask = TVIF_PARAM
    lLast = 0
    For lNode = 1 To m_lNodeCount
        If (m_uNodeData(lNode).hNode = 0) Then
            If (lLast = 0) Then
                lLast = lNode
            End If
        Else
            If m_bUseUnicode Then
                If lLast Then
                    m_uNodeData(lLast) = m_uNodeData(lNode)
                    With uTVW
                        .hItem = m_uNodeData(lLast).hNode
                        .lParam = lLast
                    End With
                    SendMessageW m_lTVHwnd, TVM_SETITEMW, 0, uTVW
                    lLast = lLast + 1
                End If
            Else
                If lLast Then
                    m_uNodeData(lLast) = m_uNodeData(lNode)
                    With uTVA
                        .hItem = m_uNodeData(lLast).hNode
                        .lParam = lLast
                    End With
                    SendMessageA m_lTVHwnd, TVM_SETITEMA, 0, uTVA
                    lLast = lLast + 1
                End If
            End If
        End If
    Next lNode
    If lLast Then
        m_lNodeCount = lLast - 1
    End If
    ReDim Preserve m_uNodeData(0 To m_lNodeCount + ((ALLOCATE_SIZE + 1) - m_lNodeCount Mod ALLOCATE_SIZE))

End Sub

Private Function StripNulls(ByVal sString As String) As String

Dim lPos As Long

    lPos = InStr(sString, Chr$(0))
    If lPos = 1 Then
        StripNulls = ""
    ElseIf (lPos > 1) Then
        StripNulls = left$(sString, lPos - 1)
        Exit Function
    End If
    StripNulls = sString

End Function

Private Function StringFromPointer(ByVal pszText As Long, _
                                   ByVal cchTextMax As Long) As String

Dim aBuffer() As Byte

    If lstrlen(pszText) > 1 Then
        ReDim aBuffer(0 To cchTextMax - 1) As Byte
        CopyMemory aBuffer(0), ByVal pszText, cchTextMax - 1
        StringFromPointer = StripNulls(StrConv(aBuffer(), vbUnicode))
    End If

End Function

Private Sub GetClientCursorPos(x As Long, _
                               y As Long)

Dim uPt As POINTAPI

    GetCursorPos uPt
    ScreenToClient m_lTVHwnd, uPt
    With uPt
        x = .x
        y = .y
    End With

End Sub

Private Function ButtonState(ByVal uMsg As Long) As Long

    Select Case uMsg
    Case WM_LBUTTONDOWN, WM_LBUTTONUP
        ButtonState = vbLeftButton
    Case WM_RBUTTONDOWN, WM_RBUTTONUP
        ButtonState = vbRightButton
    Case WM_MBUTTONDOWN, WM_MBUTTONUP
        ButtonState = vbMiddleButton
    Case WM_MOUSEMOVE
        Select Case True
        Case GetAsyncKeyState(vbKeyLButton) < 0
            ButtonState = vbLeftButton
        Case GetAsyncKeyState(vbKeyRButton) < 0
            ButtonState = vbRightButton
        Case GetAsyncKeyState(vbKeyMButton) < 0
            ButtonState = vbMiddleButton
        End Select
    End Select

End Function

Private Function ShiftState() As Long

Dim lS As Long

    If GetAsyncKeyState(vbKeyShift) < 0 Then
        lS = lS Or vbShiftMask
    End If
    If GetAsyncKeyState(vbKeyMenu) < 0 Then
        lS = lS Or vbAltMask
    End If
    If GetAsyncKeyState(vbKeyControl) < 0 Then
        lS = lS Or vbCtrlMask
    End If
    ShiftState = lS

End Function

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

Private Sub TrackMouseLeave(ByVal lHwnd As Long)

Dim uTME As TRACKMOUSEEVENT_STRUCT

    If m_bTrack Then
        With uTME
            .cbSize = Len(uTME)
            .dwFlags = TME_LEAVE
            .hwndTrack = lHwnd
        End With
        If m_bTrackUser32 Then
            TrackMouseEvent uTME
        Else
            TrackMouseEventComCtl uTME
        End If
    End If

End Sub


'**********************************************************************
'*                              SUBCLASSING
'**********************************************************************

Private Sub TreeAttach()

    With m_cTreeSubclass
        If Not (m_lParentHwnd = 0) Then
            .Subclass m_lParentHwnd, Me
            .AddMessage m_lParentHwnd, WM_MOUSEACTIVATE, MSG_AFTER
            .AddMessage m_lParentHwnd, WM_SIZE, MSG_AFTER
            .AddMessage m_lParentHwnd, WM_NOTIFY, MSG_AFTER
            .AddMessage m_lParentHwnd, WM_MOUSEMOVE, MSG_BEFORE
        End If
        If Not (m_lTVHwnd = 0) Then
            .Subclass m_lTVHwnd, Me
            .AddMessage m_lTVHwnd, WM_SIZE, MSG_BEFORE
            .AddMessage m_lTVHwnd, WM_KEYDOWN, MSG_BEFORE
            .AddMessage m_lTVHwnd, WM_CHAR, MSG_BEFORE
            .AddMessage m_lTVHwnd, WM_KEYUP, MSG_BEFORE
            .AddMessage m_lTVHwnd, WM_LBUTTONDOWN, MSG_BEFORE
            .AddMessage m_lTVHwnd, WM_RBUTTONDOWN, MSG_BEFORE
            .AddMessage m_lTVHwnd, WM_MBUTTONDOWN, MSG_BEFORE
            .AddMessage m_lTVHwnd, WM_MOUSEMOVE, MSG_BEFORE
            .AddMessage m_lTVHwnd, WM_LBUTTONUP, MSG_BEFORE
            .AddMessage m_lTVHwnd, WM_RBUTTONUP, MSG_BEFORE
            .AddMessage m_lTVHwnd, WM_MBUTTONUP, MSG_BEFORE
            .AddMessage m_lTVHwnd, WM_MOUSELEAVE, MSG_BEFORE
        End If
    End With
            
End Sub

Private Sub TreeDetach()

    With m_cTreeSubclass
        If Not (m_lParentHwnd = 0) Then
            .DeleteMessage m_lParentHwnd, WM_MOUSEACTIVATE, MSG_AFTER
            .DeleteMessage m_lParentHwnd, WM_SIZE, MSG_AFTER
            .DeleteMessage m_lParentHwnd, WM_NOTIFY, MSG_AFTER
            .DeleteMessage m_lParentHwnd, WM_MOUSEMOVE, MSG_BEFORE
            .UnSubclass m_lParentHwnd
        End If
        If Not (m_lTVHwnd = 0) Then
            .DeleteMessage m_lTVHwnd, WM_SIZE, MSG_BEFORE
            .DeleteMessage m_lTVHwnd, WM_KEYDOWN, MSG_BEFORE
            .DeleteMessage m_lTVHwnd, WM_CHAR, MSG_BEFORE
            .DeleteMessage m_lTVHwnd, WM_KEYUP, MSG_BEFORE
            .DeleteMessage m_lTVHwnd, WM_LBUTTONDOWN, MSG_BEFORE
            .DeleteMessage m_lTVHwnd, WM_RBUTTONDOWN, MSG_BEFORE
            .DeleteMessage m_lTVHwnd, WM_MBUTTONDOWN, MSG_BEFORE
            .DeleteMessage m_lTVHwnd, WM_MOUSEMOVE, MSG_BEFORE
            .DeleteMessage m_lTVHwnd, WM_LBUTTONUP, MSG_BEFORE
            .DeleteMessage m_lTVHwnd, WM_RBUTTONUP, MSG_BEFORE
            .DeleteMessage m_lTVHwnd, WM_MBUTTONUP, MSG_BEFORE
            .DeleteMessage m_lTVHwnd, WM_MOUSELEAVE, MSG_BEFORE
            .UnSubclass m_lTVHwnd
        End If
    End With

End Sub

Private Sub MoveCursor()
'/* a hack to bypass scr hittest

Dim tPcd As POINTAPI

    GetCursorPos tPcd
    ScreenToClient m_lTVHwnd, tPcd
    With tPcd
        mouse_event MOUSEEVENTF_MOVE, (.x + 4), .y, 0&, GetMessageExtraInfo()
    End With
    
End Sub

Private Sub GXISubclass_WndProc(ByVal bBefore As Boolean, _
                                bHandled As Boolean, _
                                lReturn As Long, _
                                ByVal lHwnd As Long, _
                                ByVal uMsg As eMsg, _
                                ByVal wParam As Long, _
                                ByVal lParam As Long, _
                                lParamUser As Long)

Dim lZHnd       As Long
Dim lNode       As Long
Dim lfHit       As Long
Dim lEdit       As Long
Dim lCancel     As Long
Dim lXPos       As Long
Dim lYPos       As Long
Dim sText       As String
Dim uNMH        As NMHDR
Dim uNMTV       As NMTREEVIEW
Dim uNMTVDI     As NMTVDISPINFO

    Select Case uMsg
    Case WM_NOTIFY
        CopyMemory uNMH, ByVal lParam, Len(uNMH)
        If uNMH.hwndFrom = m_lTVHwnd Then
            Select Case uNMH.code
            Case NM_CLICK, NM_RCLICK
                NodeHitTest lNode, lfHit
                If lNode Then
                    If ((lfHit And (TVHT_ONITEMICON Or TVHT_ONITEMLABEL)) Or m_bFullRowSelect) Then
                        RaiseEvent NodeClick(lNode)
                        m_bNodeClick = True
                        GetClientCursorPos lXPos, lYPos
                        RaiseEvent MouseUp((uNMH.code = NM_CLICK) + 2, ShiftState(), lXPos, lYPos)
                        RaiseEvent Click
                    End If
                End If
                    
            Case NM_CUSTOMDRAW
                lReturn = CustomDrawNode(lParam)
                    
            Case NM_DBLCLK, NM_RDBLCLK
                NodeHitTest lNode, lfHit
                If (lfHit And TVHT_ONITEMSTATEICON) Then
                    lReturn = 1
                Else
                    RaiseEvent NodeDblClick(lNode)
                End If
                    
            Case TVN_SELCHANGEDA, TVN_SELCHANGEDW
                If Not m_bNodeClick Then
                    CopyMemory uNMTV, ByVal lParam, Len(uNMTV)
                    RaiseEvent NodeClick(uNMTV.itemNew.hItem)
                End If
                m_bNodeClick = False
                RaiseEvent SelectionChanged
                    
            Case TVN_ITEMEXPANDINGA, TVN_ITEMEXPANDINGW
                CopyMemory uNMTV, ByVal lParam, Len(uNMTV)
                With uNMTV
                    If .action = TVE_EXPAND Then
                        With .itemNew
                            RaiseEvent BeforeExpand(.hItem, CBool(.State And TVIS_EXPANDEDONCE))
                        End With
                    End If
                End With
                    
            Case TVN_ITEMEXPANDEDA, TVN_ITEMEXPANDEDW
                CopyMemory uNMTV, ByVal lParam, Len(uNMTV)
                With uNMTV
                    Select Case .action
                    Case TVE_EXPAND
                        With .itemNew
                            RaiseEvent AfterExpand(.hItem, CBool(.State And TVIS_EXPANDEDONCE))
                        End With
                    Case TVE_COLLAPSE
                        RaiseEvent Collapse(.itemNew.hItem)
                    End Select
                End With
                    
            Case TVN_DELETEITEMA, TVN_DELETEITEMW
                CopyMemory uNMTV, ByVal lParam, Len(uNMTV)
                m_uNodeData(uNMTV.itemOld.lParam).hNode = 0
                
            Case TVN_BEGINLABELEDITA, TVN_BEGINLABELEDITW
                CopyMemory uNMTVDI, ByVal lParam, Len(uNMTVDI)
                RaiseEvent BeforeLabelEdit(uNMTVDI.Item.hItem, lCancel)
                If lCancel Then
                    lEdit = SendMessageLongA(m_lTVHwnd, TVM_GETEDITCONTROL, 0, 0)
                    SendMessageLongA lEdit, WM_KILLFOCUS, 0&, 0&
                End If
                
            Case TVN_ENDLABELEDITA, TVN_ENDLABELEDITW
                CopyMemory uNMTVDI, ByVal lParam, Len(uNMTVDI)
                With uNMTVDI.Item
                    sText = StringFromPointer(.pszText, .cchTextMax)
                    RaiseEvent AfterLabelEdit(.hItem, lCancel, sText)
                    If (lCancel = 0 Xor GetAsyncKeyState(vbKeyEscape)) Then
                        lReturn = 1
                        bHandled = True
                    End If
                End With
            End Select
        End If
        
    Case WM_KEYDOWN
        RaiseEvent KeyDown(wParam And &H7FFF&, ShiftState())
        If wParam = VK_TAB Then
            SetFocus m_lGridHwnd
            bHandled = True
        End If
        
    Case WM_CHAR
        If wParam = VK_TAB Then
            RaiseEvent KeyPress(wParam And &H7FFF&)
            bHandled = True
        Else
            If ((wParam And &H7FFF&) = vbKeySpace) Then
                If m_bCheckBoxes Then
                    RaiseEvent NodeCheck(SelectedNode())
                End If
            End If
            RaiseEvent KeyPress(wParam And &H7FFF&)
        End If
            
    Case WM_KEYUP
        RaiseEvent KeyUp(wParam And &H7FFF&, ShiftState())
        
    Case WM_LBUTTONDOWN, WM_RBUTTONDOWN, WM_MBUTTONDOWN
        GetClientCursorPos lXPos, lYPos
        RaiseEvent MouseDown(ButtonState(uMsg), ShiftState(), lXPos, lYPos)
        If ButtonState(uMsg) = vbLeftButton Then
            NodeHitTest lNode, lfHit
            If (lfHit = 16) Then
                If LeftKeyState Then
                    MoveCursor
                    m_cSkinScrollBars.Refresh
                End If
            End If
            If lNode Then
                If ((lfHit And TVHT_ONITEMSTATEICON) = TVHT_ONITEMSTATEICON) Then
                    SetTvStateImage lNode, SII_SWAPMASK - GetTvStateImage(lNode)
                    RaiseEvent NodeCheck(lNode)
                    bHandled = True
                End If
            End If
        End If
        bHandled = bHandled Or (ButtonState(uMsg) <> vbLeftButton)
            
    Case WM_MOUSEMOVE
        If Not m_bInControl Then
            m_bInControl = True
            TrackMouseLeave lHwnd
        End If
        GetClientCursorPos lXPos, lYPos
        If (lXPos <> m_lXCoord Or lXPos <> m_lYCoord) Then
            m_lXCoord = lXPos
            m_lYCoord = lYPos
        End If
        RaiseEvent MouseMove(ButtonState(uMsg), ShiftState(), lXPos, lYPos)

    Case WM_MOUSELEAVE
        m_bInControl = False
        m_lXCoord = -1
        m_lYCoord = -1
        RaiseEvent MouseLeave
            
    Case WM_LBUTTONUP, WM_RBUTTONUP, WM_MBUTTONUP
        GetClientCursorPos lXPos, lYPos
        RaiseEvent MouseUp(ButtonState(uMsg), ShiftState(), lXPos, lYPos)
        RaiseEvent Click
    End Select

End Sub

Private Function CustomDrawNode(ByVal lParam As Long) As Long

Dim lReturn     As Long
Dim lState      As Long
Dim tNmTvCd     As NMTVCUSTOMDRAW

    CopyMemory tNmTvCd, ByVal lParam, Len(tNmTvCd)
    Select Case tNmTvCd.nmcmd.dwDrawStage
    Case CDDS_PREPAINT
        lReturn = CDRF_NOTIFYITEMDRAW
    Case CDDS_ITEMPREPAINT
        lState = tNmTvCd.nmcmd.uItemState
        If Not m_bEnabled Then
            lState = CDIS_DISABLED
        End If
        If ((lState And CDIS_FOCUS) = CDIS_FOCUS) Then
            SetBkColor tNmTvCd.nmcmd.hdc, m_lFocusBackColor
            SetTextColor tNmTvCd.nmcmd.hdc, m_lFocusForeColor
        ElseIf ((lState And CDIS_DISABLED) = CDIS_DISABLED) Then
            SetBkColor tNmTvCd.nmcmd.hdc, m_lDisabledBackColor
            SetTextColor tNmTvCd.nmcmd.hdc, m_lDisabledForeColor
        ElseIf ((lState And CDIS_HOT) = CDIS_HOT) Then
            SetBkColor tNmTvCd.nmcmd.hdc, m_lFocusBackColor
            SetTextColor tNmTvCd.nmcmd.hdc, m_lFocusForeColor
        ElseIf ((lState And CDIS_SELECTED) = CDIS_SELECTED) Then
            SetBkColor tNmTvCd.nmcmd.hdc, m_lSelectedBackColor
            SetTextColor tNmTvCd.nmcmd.hdc, m_lForeColor
        Else
            SetBkColor tNmTvCd.nmcmd.hdc, m_lBackColor
            SetTextColor tNmTvCd.nmcmd.hdc, m_lForeColor
        End If
    Case CDDS_ITEMPOSTPAINT
        lReturn = CDRF_DODEFAULT
    End Select

    CustomDrawNode = lReturn
            
End Function

'**********************************************************************
'*                              CONTROL
'**********************************************************************

Friend Sub Resize(Optional ByVal lAlign As Long)

Dim tRect As RECT
Dim tRClt As RECT

    If Not (m_cSkinScrollBars Is Nothing) Then
        m_cSkinScrollBars.ByPassHitTest = True
    End If
    If Not (m_lTVHwnd = 0) Then
        GetClientRect m_lParentHwnd, tRect
        Select Case lAlign
        '/* left
        Case 0
            SetWindowPos m_lTVHwnd, 0&, 0&, 0&, m_lWidth - 10, tRect.Bottom, SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOZORDER
        '/* right
        Case 1
            GetWindowRect m_lGridHwnd, tRClt
            OffsetRect tRClt, -tRClt.left, -tRClt.top
            SetWindowPos m_lTVHwnd, 0&, tRClt.Right + 10, 0&, m_lWidth, tRect.Bottom, SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOZORDER
        '/* top
        Case 2
            SetWindowPos m_lTVHwnd, 0&, 0&, 0&, tRect.Right, m_lHeight, SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOZORDER
        '/* bottom
        Case 3
            GetWindowRect m_lGridHwnd, tRClt
            OffsetRect tRClt, -tRClt.left, -tRClt.top
            SetWindowPos m_lTVHwnd, 0&, 0&, tRClt.Bottom + 10, tRect.Right, m_lHeight, SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOZORDER
        End Select
    End If
    If Not (m_cSkinScrollBars Is Nothing) Then
        m_cSkinScrollBars.Resize
        m_cSkinScrollBars.ByPassHitTest = False
        m_cSkinScrollBars.Refresh
    End If

End Sub


'**********************************************************************
'*                              CLEANUP
'**********************************************************************

Private Sub DestroyTreeview()

On Error Resume Next

    If Not (m_lTVHwnd = 0) Then
        TreeDetach
        DestroyFont
        DestroyImlState
        Erase m_uNodeData()
        DestroyTree
        If Not m_cTreeSubclass Is Nothing Then Set m_cTreeSubclass = Nothing
        If Not m_IChecked Is Nothing Then Set m_IChecked = Nothing
        If Not m_IChkDisabled Is Nothing Then Set m_IChkDisabled = Nothing
        If Not m_IUnChecked Is Nothing Then Set m_IUnChecked = Nothing
        If Not m_cChkCheckDc Is Nothing Then Set m_cChkCheckDc = Nothing
        If Not m_cRender Is Nothing Then Set m_cRender = Nothing
        If Not m_cSkinScrollBars Is Nothing Then Set m_cSkinScrollBars = Nothing
        If Not m_oFont Is Nothing Then Set m_oFont = Nothing
        If Not m_cKey Is Nothing Then Set m_cKey = Nothing
        If Not (m_hModShell32 = 0) Then
            FreeLibrary m_hModShell32
            m_hModShell32 = 0
        End If
        m_lParentHwnd = 0
    End If
    
On Error GoTo 0

End Sub

Private Sub DestroyTree()

    If Not (m_lTVHwnd = 0) Then
        DestroyWindow m_lTVHwnd
        m_lTVHwnd = 0
    End If

End Sub

Private Sub DestroyImlState()
'*/ destroy header image list

    If Not (m_lImlStateHndl = 0) Then
        If ImageList_Destroy(m_lImlStateHndl) Then
            m_lImlStateHndl = 0
        End If
    End If

End Sub

Private Sub DestroyFont()

    If Not (m_lHFont = 0) Then
        If DeleteObject(m_lHFont) Then
            m_lHFont = 0
        End If
    End If

End Sub

Private Sub Class_Terminate()
    DestroyTreeview
End Sub
